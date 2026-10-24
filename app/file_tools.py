"""Independent image/PDF engine copied from Orbit Files 3.2. Original project is untouched."""
import sys, os, math, io, tempfile


from pathlib import Path


from PIL import Image, ImageOps


import pymupdf


from PyQt6.QtCore import (Qt, QRectF, QPointF, QThread, pyqtSignal, QUrl,
                          QTimer, QAbstractNativeEventFilter, QSettings)


from PyQt6.QtGui import (QColor, QPainter, QPainterPath, QPen, QFont, QPixmap,
                         QImage, QDesktopServices, QIcon, QFontDatabase, QCursor, QKeySequence)


from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QDialog, QDialogButtonBox, QFormLayout,
    QSpinBox, QSlider, QComboBox, QMessageBox, QProgressBar, QCheckBox,
    QSystemTrayIcon, QMenu, QColorDialog, QKeySequenceEdit)


SUPPORTED = {'.png', '.jpg', '.jpeg', '.webp', '.bmp', '.tif', '.tiff', '.gif', '.pdf'}


FORMATS = {'PNG':'PNG', 'JPG':'JPEG', 'WEBP':'WEBP', 'BMP':'BMP', 'TIFF':'TIFF'}


ACCENT = '#ef743f'


STYLE = """
QWidget {font-family: 'Segoe UI'; font-size: 13px; color: #293547;}
QDialog {background: #f7f8fb;}
QPushButton {background:#fff; border:1px solid #e3e7ed; border-radius:10px; padding:9px 15px;}
QPushButton:hover {background:#fff1e7; border-color:#efab88;}
QPushButton:disabled {color:#a4aab5; background:#f1f2f5;}
QPushButton[primary="true"] {background:#c25121; color:white; border:0; font-weight:600;}
QPushButton[primary="true"]:hover {background:#ac4219;}
QSpinBox, QComboBox {background:white; border:1px solid #dce1e9; border-radius:7px; padding:7px; min-width:100px;}
QProgressBar {border:0; background:#e7ebf1; border-radius:4px; height:6px; text-align:center;}
QProgressBar::chunk {background:#ef743f; border-radius:4px;}
QCheckBox {spacing:8px;}
QSlider::groove:horizontal {height:5px; background:#e0e4eb; border-radius:2px;}
QSlider::handle:horizontal {width:17px; margin:-6px 0; background:#ef743f; border-radius:8px;}
"""


def unique_path(folder, stem, suffix):
    p = Path(folder) / (stem + suffix)
    n = 2
    while p.exists():
        p = Path(folder) / f'{stem} ({n}){suffix}'
        n += 1
    return p


def open_image(path):
    with Image.open(path) as src:
        # Check the lazy header before decoding, transposing, or copying pixels.
        if src.width * src.height > 100_000_000:
            raise ValueError('الصورة كبيرة جدًا. الحد الأقصى 100 مليون بكسل.')
        # Animated files intentionally use their first frame.
        im = ImageOps.exif_transpose(src).copy()
    return im


def flatten(im):
    if im.mode in ('RGBA', 'LA') or 'transparency' in im.info:
        rgba = im.convert('RGBA')
        bg = Image.new('RGB', im.size, 'white')
        bg.paste(rgba, mask=rgba.getchannel('A'))
        return bg
    return im.convert('RGB')


class ConversionCancelled(InterruptedError):
    """The current input is rolled back; previously completed inputs stay saved."""


def check_cancelled(cancelled):
    if cancelled and cancelled():
        raise ConversionCancelled('تم إلغاء العملية.')


def atomic_output(dest, writer, cancelled=None):
    """Publish one complete file without replacing any existing destination.

    The staging file lives on the same volume. Windows rename and POSIX link
    both fail if a competing process has already published the chosen name.
    """
    dest = Path(dest)
    check_cancelled(cancelled)
    fd, tmp_name = tempfile.mkstemp(prefix='.orbit-', suffix='.tmp', dir=str(dest.parent))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        writer(tmp)
        check_cancelled(cancelled)
        candidate = dest
        number = 2
        while True:
            check_cancelled(cancelled)
            try:
                if sys.platform == 'win32':
                    # Unlike os.replace, MoveFile without REPLACE_EXISTING is exclusive.
                    os.rename(tmp, candidate)
                else:
                    os.link(tmp, candidate)
                return candidate
            except FileExistsError:
                candidate = dest.with_name(f'{dest.stem} ({number}){dest.suffix}')
                number += 1
    finally:
        tmp.unlink(missing_ok=True)


def save_image(im, dest, fmt, quality=88, cancelled=None):
    check_cancelled(cancelled)
    if fmt in ('JPEG', 'BMP'):
        im = flatten(im)
    elif fmt == 'PNG' and im.mode not in ('RGB', 'RGBA', 'L', 'LA', 'P'):
        im = im.convert('RGBA')
    elif fmt == 'WEBP':
        im = im.convert('RGBA' if 'A' in im.getbands() else 'RGB')
    opts = {'quality': quality} if fmt in ('JPEG', 'WEBP') else {}
    if fmt == 'PNG': opts.update(optimize=True)
    if fmt == 'TIFF': opts.update(compression='tiff_deflate')
    return atomic_output(dest, lambda tmp: im.save(tmp, format=fmt, **opts), cancelled)


def transform_file(path, action, options, folder, progress=lambda v: None, cancelled=None):
    """Return complete copies, rolling back this input on failure/cancellation."""
    path = Path(path)
    outputs = []
    try:
        check_cancelled(cancelled)
        if path.suffix.lower() == '.pdf':
            if action not in ('convert', 'compress'):
                raise ValueError('هذه الأداة مخصصة للصور. حوّل صفحات PDF إلى صور أولًا.')
            with pymupdf.open(path) as doc:
                if doc.needs_pass: raise ValueError('ملف PDF محمي بكلمة مرور.')
                if not len(doc): raise ValueError('ملف PDF لا يحتوي على صفحات.')
                if action == 'compress' or options.get('format') == 'PDF':
                    data = doc.tobytes(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True)
                    check_cancelled(cancelled)
                    if len(data) >= path.stat().st_size: data = path.read_bytes()
                    dest = unique_path(folder, path.stem + '_optimized', '.pdf')
                    outputs.append(atomic_output(dest, lambda tmp: tmp.write_bytes(data), cancelled))
                else:
                    fmt = FORMATS[options['format']]
                    for i, page in enumerate(doc):
                        check_cancelled(cancelled)
                        # Bound raster dimensions, including unusually huge PDF pages.
                        edge = max(page.rect.width, page.rect.height)
                        if not math.isfinite(edge) or edge <= 0:
                            raise ValueError('أبعاد صفحة PDF غير صالحة.')
                        scale = min(2.0, 6000 / edge)
                        pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
                        check_cancelled(cancelled)
                        with Image.frombytes('RGB', (pix.width, pix.height), pix.samples) as im:
                            dest = unique_path(folder, f'{path.stem}_page_{i+1:03}', '.' + options['format'].lower())
                            outputs.append(save_image(im, dest, fmt, options.get('quality', 88), cancelled))
                        progress(int(100*(i+1)/len(doc)))
        else:
            im = open_image(path)
            try:
                check_cancelled(cancelled)
                fmt_name = options.get('format', 'PNG')
                if action == 'convert' and fmt_name == 'PDF':
                    dest = unique_path(folder, path.stem + '_converted', '.pdf')
                    buf = io.BytesIO()
                    im.convert('RGBA' if 'A' in im.getbands() else 'RGB').save(buf, format='PNG')
                    with pymupdf.open() as doc:
                        page = doc.new_page(width=im.width * .75, height=im.height * .75)
                        page.insert_image(page.rect, stream=buf.getvalue())
                        outputs.append(atomic_output(dest, lambda tmp: doc.save(tmp, garbage=4, deflate=True), cancelled))
                else:
                    if action == 'compress':
                        fmt_name = 'WEBP' if options.get('transparent') else 'JPG'
                    elif action == 'resize':
                        w,h = options['width'], options['height']
                        if w < 1 or h < 1: raise ValueError('العرض والارتفاع يجب أن يكونا أكبر من صفر.')
                        if options.get('keep_ratio', True):
                            scale = min(w/im.width, h/im.height)
                            w,h = max(1,round(im.width*scale)), max(1,round(im.height*scale))
                        if w*h > 100_000_000: raise ValueError('المقاس المطلوب أكبر من الحد المسموح.')
                        im = im.resize((w,h), Image.Resampling.LANCZOS)
                    elif action == 'rotate': im = im.rotate(-options['angle'], expand=True)
                    elif action == 'crop': im = im.crop(options['box'])
                    elif action == 'annotate': im = options['image'].copy()
                    fmt = FORMATS[fmt_name]
                    dest = unique_path(folder, path.stem + '_' + action, '.' + fmt_name.lower())
                    outputs.append(save_image(im, dest, fmt, options.get('quality', 88), cancelled))
            finally:
                im.close()
        check_cancelled(cancelled)
        return outputs
    except BaseException as error:
        # Only paths published by THIS call are owned and safe to remove.
        leftovers = []
        for out in outputs:
            try: out.unlink(missing_ok=True)
            except OSError: leftovers.append(str(out))
        if leftovers:
            raise OSError('تعذر حذف بعض النسخ غير المكتملة:\n' + '\n'.join(leftovers)) from error
        raise


class Worker(QThread):
    progress = pyqtSignal(int)
    completed = pyqtSignal(object, object)
    def __init__(self, files, action, opts, folder):
        super().__init__(); self.files=tuple(files); self.action=action; self.opts=dict(opts); self.folder=folder
        self.results=[]; self.errors=[]; self.cancelled=False
    def run(self):
        for i, path in enumerate(self.files):
            if self.isInterruptionRequested(): self.cancelled=True; break
            try:
                destination=Path(path).absolute().parent if self.folder is None else self.folder
                self.results.extend(transform_file(path, self.action, self.opts, destination,
                    lambda p: self.progress.emit(int((i+p/100)*100/len(self.files))),
                    cancelled=self.isInterruptionRequested))
            except ConversionCancelled:
                self.cancelled=True; break
            except Exception as e:
                self.errors.append(f'{Path(path).name}\n{e}')
            self.progress.emit(int((i+1)*100/len(self.files)))
        if self.isInterruptionRequested(): self.cancelled=True
        self.completed.emit(self.results, self.errors)


class EditCanvas(QWidget):
    def __init__(self,im,mode):
        super().__init__(); self.original=im.convert('RGBA'); self.mode=mode
        self.setMinimumSize(640,420); self.setMouseTracking(True)
        self.start=None; self.end=None; self.strokes=[]; self.color='#ef743f'; self.line_width=5
        self.image_rect=QRectF()
    def paintEvent(self,e):
        p=QPainter(self); p.fillRect(self.rect(),QColor('#e8ecf2')); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        scale=min((self.width()-32)/self.original.width,(self.height()-32)/self.original.height)
        w,h=self.original.width*scale,self.original.height*scale
        self.image_rect=QRectF((self.width()-w)/2,(self.height()-h)/2,w,h)
        data=self.original.tobytes('raw','RGBA'); q=QImage(data,self.original.width,self.original.height,QImage.Format.Format_RGBA8888)
        p.drawImage(self.image_rect,q)
        if self.mode=='crop' and self.start and self.end:
            crop=QRectF(self.to_view(self.start),self.to_view(self.end)).normalized()
            shade=QPainterPath(); shade.addRect(self.image_rect); hole=QPainterPath(); hole.addRect(crop)
            p.fillPath(shade.subtracted(hole),QColor(20,30,45,125))
            p.setPen(QPen(QColor(ACCENT),2)); p.drawRect(crop)
        for color,width,points in self.strokes:
            p.setPen(QPen(QColor(color),width*scale,Qt.PenStyle.SolidLine,Qt.PenCapStyle.RoundCap,Qt.PenJoinStyle.RoundJoin))
            for a,b in zip(points,points[1:]): p.drawLine(self.to_view(a),self.to_view(b))
    def to_image(self,pos):
        r=self.image_rect
        return (max(0,min(self.original.width,(pos.x()-r.x())/r.width()*self.original.width)),
                max(0,min(self.original.height,(pos.y()-r.y())/r.height()*self.original.height)))
    def to_view(self,pt):
        return QPointF(self.image_rect.x()+pt[0]/self.original.width*self.image_rect.width(),self.image_rect.y()+pt[1]/self.original.height*self.image_rect.height())
    def mousePressEvent(self,e):
        if e.button()!=Qt.MouseButton.LeftButton or not self.image_rect.contains(e.position()): return
        self.start=self.to_image(e.position()); self.end=self.start
        if self.mode=='annotate': self.strokes.append((self.color,self.line_width,[self.start,self.start]))
        self.update()
    def mouseMoveEvent(self,e):
        if self.start and e.buttons() & Qt.MouseButton.LeftButton:
            self.end=self.to_image(e.position())
            if self.mode=='annotate' and self.strokes: self.strokes[-1][2].append(self.end)
            self.update()
    def crop_box(self):
        if not self.start or not self.end: return None
        x1,y1=self.start; x2,y2=self.end
        box=(round(min(x1,x2)),round(min(y1,y2)),round(max(x1,x2)),round(max(y1,y2)))
        return box if box[2]-box[0]>=2 and box[3]-box[1]>=2 else None
    def result_image(self):
        from PIL import ImageDraw
        im=self.original.copy(); d=ImageDraw.Draw(im)
        for color,width,pts in self.strokes:
            d.line(pts,fill=color,width=width,joint='curve')
            r=width/2
            for x,y in (pts[0],pts[-1]): d.ellipse((x-r,y-r,x+r,y+r),fill=color)
        return im


class ImageEditor(QDialog):
    def __init__(self,path,mode,parent):
        super().__init__(parent); self.mode=mode
        self.setWindowTitle('قص الصورة' if mode=='crop' else 'الرسم على الصورة'); self.resize(830,650)
        layout=QVBoxLayout(self)
        hint=QLabel('اسحب لتحديد الجزء المطلوب، ثم احفظ نسخة.' if mode=='crop' else 'ارسم بالماوس، ثم احفظ نسخة.'); hint.setAlignment(Qt.AlignmentFlag.AlignRight); layout.addWidget(hint)
        self.canvas=EditCanvas(open_image(path),mode); layout.addWidget(self.canvas,1)
        row=QHBoxLayout()
        if mode=='annotate':
            color=QPushButton('لون القلم'); color.clicked.connect(self.choose_color); row.addWidget(color)
            size=QSpinBox(); size.setRange(1,60); size.setValue(5); size.setSuffix(' px'); size.valueChanged.connect(lambda v:setattr(self.canvas,'line_width',v)); row.addWidget(size)
            undo=QPushButton('تراجع'); undo.clicked.connect(self.undo); row.addWidget(undo)
        else:
            reset=QPushButton('إعادة التحديد'); reset.clicked.connect(self.reset); row.addWidget(reset)
        row.addStretch(); cancel=QPushButton('إلغاء'); cancel.clicked.connect(self.reject); row.addWidget(cancel)
        save=QPushButton('حفظ نسخة'); save.setProperty('primary',True); save.clicked.connect(self.validate); row.addWidget(save); layout.addLayout(row)
    def reset(self): self.canvas.start=None; self.canvas.end=None; self.canvas.update()
    def undo(self):
        if self.canvas.strokes: self.canvas.strokes.pop(); self.canvas.update()
    def choose_color(self):
        c=QColorDialog.getColor(QColor(self.canvas.color),self)
        if c.isValid(): self.canvas.color=c.name()
    def validate(self):
        if self.mode=='crop' and not self.canvas.crop_box():
            QMessageBox.information(self,'تحديد القص','اسحب على الصورة لتحديد جزء بحجم مناسب.'); return
        self.accept()


from PyQt6.QtWidgets import QScrollArea, QFrame, QTextEdit


def settings_dialog(parent,action,path,files=None):
    d=QDialog(parent); d.setWindowTitle({'compress':'ضغط الملف','resize':'تغيير المقاس','rotate':'تدوير الصورة'}[action])
    d.setMinimumWidth(380); layout=QVBoxLayout(d); form=QFormLayout(); form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
    controls={}
    if action=='compress':
        files=files or [path]
        has_pdf=any(Path(p).suffix.lower()=='.pdf' for p in files)
        has_images=any(Path(p).suffix.lower()!='.pdf' for p in files)
        if has_pdf:
            layout.addWidget(QLabel('تحسين بنية PDF دون تغيير جودة الصفحات.\nقد لا يقل الحجم إذا كان الملف مضغوطًا بالفعل.'))
        if has_images:
            q=QSpinBox(); q.setRange(10,100); q.setValue(78); q.setSuffix(' %'); form.addRow('الجودة',q); controls['quality']=q
            transparent=QCheckBox('الحفاظ على الشفافية بصيغة WEBP'); form.addRow(transparent); controls['transparent']=transparent
            note=QLabel('الحفظ بصيغة JPG، أو WEBP عند تفعيل الشفافية.\nالحجم النهائي يعتمد على الصورة والجودة.'); note.setWordWrap(True); layout.addWidget(note)
    elif action=='resize':
        im=open_image(path)
        layout.addWidget(QLabel(f'المقاس الحالي: {im.width} × {im.height} بكسل'))
        for name,label,value in [('width','العرض',im.width),('height','الارتفاع',im.height)]:
            c=QSpinBox(); c.setRange(1,30000); c.setValue(value); form.addRow(label,c); controls[name]=c
        keep=QCheckBox('الحفاظ على النسبة داخل المقاس المحدد'); keep.setChecked(True); form.addRow(keep); controls['keep_ratio']=keep
        layout.addWidget(QLabel('يُطبّق المقاس نفسه على كل الصور المختارة.'))
    else:
        combo=QComboBox(); combo.addItems(['90° مع عقارب الساعة','180°','90° عكس عقارب الساعة']); form.addRow('زاوية التدوير',combo); controls['angle']=combo
    layout.addLayout(form)
    buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
    buttons.button(QDialogButtonBox.StandardButton.Ok).setText('متابعة'); buttons.button(QDialogButtonBox.StandardButton.Cancel).setText('إلغاء')
    buttons.accepted.connect(d.accept); buttons.rejected.connect(d.reject); layout.addWidget(buttons)
    if d.exec()!=QDialog.DialogCode.Accepted: return None
    opts={}
    for key,c in controls.items():
        opts[key]=c.isChecked() if isinstance(c,QCheckBox) else ([90,180,270][c.currentIndex()] if isinstance(c,QComboBox) else c.value())
    return opts


def make_icon():
    pix=QPixmap(128,128); pix.fill(Qt.GlobalColor.transparent); p=QPainter(pix); p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setPen(Qt.PenStyle.NoPen); p.setBrush(QColor('#ef743f')); p.drawRoundedRect(4,4,120,120,30,30)
    p.setPen(QPen(QColor('white'),9)); p.setBrush(Qt.BrushStyle.NoBrush); p.drawEllipse(QPointF(64,64),34,34); p.setPen(Qt.PenStyle.NoPen); p.setBrush(QColor('white')); p.drawEllipse(QPointF(64,64),12,12); p.end()
    return QIcon(pix)


