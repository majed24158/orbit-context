"""Orbit Context: independent Explorer tools, without a wheel or global shortcut."""
import sys,os,argparse
from pathlib import Path
from PyQt6.QtCore import Qt,QTimer,QObject,QSettings,QLockFile,QUrl
from PyQt6.QtGui import QFont,QFontDatabase,QDesktopServices
from PyQt6.QtWidgets import (QApplication,QWidget,QDialog,QVBoxLayout,QHBoxLayout,
    QLabel,QPushButton,QComboBox,QFormLayout,QCheckBox,QFileDialog,QMessageBox,
    QProgressBar,QSystemTrayIcon,QMenu,QTextEdit,QDialogButtonBox,QScrollArea,QFrame)
from file_tools import (SUPPORTED,FORMATS,STYLE,Worker,ImageEditor,settings_dialog,make_icon)
from word_export import WORD_EXTENSIONS,WordWorker
import registration

VERSION='1.2'
IMAGE_PDF_EXTENSIONS=frozenset(SUPPORTED)
SMART_ACTIONS=frozenset({'docx_ai','text_ai','txt_docx','txt_format'})
WORD_SMART_ACTIONS=frozenset({'docx_ai','txt_docx'})
TXT_ACTIONS=frozenset({'txt_docx','txt_format'})
SMART_TITLES={'docx_ai':'تحويل إلى Word — Codex','text_ai':'استخراج TXT — Codex',
              'txt_docx':'تحويل TXT إلى Word — Codex','txt_format':'تنسيق TXT — Codex'}
SUPPORTED=SUPPORTED | WORD_EXTENSIONS | {'.txt'}
def preferences(): return QSettings('OrbitContextTools','OrbitContextTools')

def app_executable():
    if not getattr(sys,'frozen',False):
        raise RuntimeError('للتثبيت استخدم OrbitContext.exe المرفق، أو ابنِ ملف التشغيل من المصدر.')
    return Path(sys.executable)

class SmartOptionsDialog(QDialog):
    """Per-operation choices; cloud permission is managed explicitly in Manager."""
    def __init__(self,parent,source,action,settings):
        super().__init__(parent); self.action=action; self.settings=settings
        self.setWindowTitle(SMART_TITLES[action]); self.resize(560,510)
        layout=QVBoxLayout(self)
        heading=QLabel(SMART_TITLES[action]+'\n'+Path(source).name); heading.setWordWrap(True); layout.addWidget(heading)
        disclosure=QLabel('سيُرسل محتوى الملف إلى Codex عبر الحساب المتصل لمعالجته.\nالمطابقة بنسبة 100% مع الأصل غير مضمونة؛ راجع النص والجداول والتنسيق قبل الاعتماد على النتيجة.')
        disclosure.setWordWrap(True); disclosure.setStyleSheet('background:#fff0e5;color:#693a26;padding:12px;border-radius:8px;'); layout.addWidget(disclosure)
        form=QFormLayout(); self.text_style=QComboBox()
        for label,value in [('نص بسيط','plain'),('الحفاظ على ترتيب الأسطر','lines'),('عناوين وقوائم Markdown','markdown'),('جداول مفصولة بعلامة Tab','tsv'),('تنسيق مخصص حسب تعليماتي','custom')]:
            self.text_style.addItem(label,value)
        self.pref_prefix='smart/'+action+'/'
        self.text_style.setCurrentIndex(max(0,self.text_style.findData(str(settings.value(self.pref_prefix+'textStyle','plain')))))
        if action not in WORD_SMART_ACTIONS: form.addRow('تنسيق النص الناتج',self.text_style)
        layout.addLayout(form)
        label=QLabel('تعليمات التنسيق الخاصة بك (اختياري)' if action in WORD_SMART_ACTIONS else 'تعليمات التنسيق الخاصة بك (اختياري، ومطلوب عند اختيار تنسيق مخصص)'); label.setWordWrap(True); layout.addWidget(label)
        self.instructions=QTextEdit(); self.instructions.setAcceptRichText(False)
        self.instructions.setPlaceholderText('مثال: حافظ على العناوين، واجعل الجداول قابلة للتحرير، وضع كل سؤال في فقرة مستقلة.')
        self.instructions.setPlainText(str(settings.value(self.pref_prefix+'instructions',''))); layout.addWidget(self.instructions,1)
        self.review=QCheckBox('مراجعة التنسيق في الخلفية باستخدام Word')
        self.review.setChecked(settings.value(self.pref_prefix+'review',action=='docx_ai',type=bool))
        if action in WORD_SMART_ACTIONS:
            layout.addWidget(self.review)
            note=QLabel('الناتج مستند Word قابل للتحرير. عند تفعيل المراجعة، يُستخدم Word مخفيًا لفحص شكل الصفحات ومقارنته بالأصل، مثل مواقع العناوين وخلايا الجدول. يحتاج Word مثبّتًا ولا يفتح النتيجة.')
            note.setWordWrap(True); note.setStyleSheet('color:#526174;'); layout.addWidget(note)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Ok|QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText('بدء التحويل'); buttons.button(QDialogButtonBox.StandardButton.Cancel).setText('إلغاء')
        buttons.accepted.connect(self.accept_options); buttons.rejected.connect(self.reject); layout.addWidget(buttons)
    def options(self):
        return {'instructions':self.instructions.toPlainText().strip(),
                'text_style':self.text_style.currentData() if self.action not in WORD_SMART_ACTIONS else 'plain',
                'review':self.review.isChecked() if self.action in WORD_SMART_ACTIONS else False}
    def accept_options(self):
        opts=self.options()
        if opts['text_style']=='custom' and not opts['instructions']:
            QMessageBox.information(self,'تعليمات التنسيق','اكتب تعليماتك للتنسيق المخصص أو اختر تنسيقًا آخر.'); return
        self.settings.setValue(self.pref_prefix+'instructions',opts['instructions'])
        self.settings.setValue(self.pref_prefix+'textStyle',opts['text_style'])
        self.settings.setValue(self.pref_prefix+'review',opts['review']); self.settings.sync(); self.accept()

class Manager(QWidget):
    def __init__(self):
        super().__init__(); self.settings=preferences()
        self.setWindowTitle('Orbit Context '+VERSION+' — أدوات الزر الأيمن'); self.resize(660,700); self.setMinimumSize(480,400)
        outer=QVBoxLayout(self); outer.setContentsMargins(0,0,0,0)
        scroll=QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.Shape.NoFrame)
        content=QWidget(); scroll.setWidget(content); outer.addWidget(scroll)
        layout=QVBoxLayout(content); layout.setContentsMargins(26,22,26,22); layout.setSpacing(13)
        title=QLabel('أدواتك في قائمة الزر الأيمن'); title.setStyleSheet('font-size:24px;font-weight:650;'); layout.addWidget(title)
        subtitle=QLabel('اختر صورة أو PDF أو Word أو TXT بزر الماوس الأيمن،\nثم «أدوات الملفات — Orbit» واختر الأداة.'); subtitle.setWordWrap(True); layout.addWidget(subtitle)
        self.state=QLabel(); self.state.setWordWrap(True); self.state.setStyleSheet('background:#fff0e5;color:#693a26;padding:14px;border-radius:10px;'); layout.addWidget(self.state)
        row=QHBoxLayout(); self.install_button=QPushButton('تثبيت / تحديث قسم الأدوات'); self.install_button.setProperty('primary',True); self.install_button.clicked.connect(self.install_menu); row.addWidget(self.install_button)
        self.remove_button=QPushButton('إزالة قسم الأدوات'); self.remove_button.clicked.connect(self.remove_menu); row.addWidget(self.remove_button); layout.addLayout(row)
        form=QFormLayout(); self.destination=QComboBox()
        for text,value in [('نفس مجلد الملف الأصلي','original'),('اسألني كل مرة','ask'),('آخر مجلد حفظ ناجح','last')]: self.destination.addItem(text,value)
        self.destination.setCurrentIndex(max(0,self.destination.findData(str(self.settings.value('outputMode','original')))))
        form.addRow('حفظ النتائج',self.destination); layout.addLayout(form)
        self.notices=QCheckBox('إظهار إشعار نتيجة مختصر'); self.notices.setChecked(self.settings.value('notify',True,type=bool)); layout.addWidget(self.notices)
        self.codex_allowed=QCheckBox('السماح بإرسال محتوى الملفات التي أحولها إلى Codex')
        self.codex_allowed.setChecked(self.settings.value('codexAllowed',False,type=bool))
        layout.addWidget(self.codex_allowed)
        cloud_note=QLabel('التحويل الذكي يرسل محتوى الملف عبر حساب Codex المتصل. يكون متوقفًا حتى تفعّل السماح وتحفظ الإعدادات. لا تحتاج الأدوات المحلية إلى هذا السماح.')
        cloud_note.setWordWrap(True); cloud_note.setStyleSheet('color:#526174;'); layout.addWidget(cloud_note)
        save=QPushButton('حفظ الإعدادات'); save.clicked.connect(self.save_preferences); layout.addWidget(save)
        note=QLabel('التحويل الذكي: صور وPDF إلى Word قابل للتحرير أو TXT.\nTXT: تحويل إلى Word أو إعادة تنسيق النص حسب تعليماتك.\nالمطابقة الكاملة غير مضمونة، وتبقى مراجعة النتيجة ضرورية.\nWord إلى PDF: بوورد المثبّت، مخفيًا وللقراءة فقط؛ DOC / DOCX / DOCM / RTF.\nالصور: تحويل، ضغط، قص، مقاس، تدوير ورسم.\nPDF محليًا: تحويل الصفحات إلى صور وتحسين الضغط.\nالنتائج نسخ جديدة؛ القسم مخصص لتحديد ملف واحد.'); note.setWordWrap(True); note.setStyleSheet('color:#526174;'); layout.addWidget(note)
        win11=QLabel('في Windows 11 قد يظهر القسم داخل «إظهار المزيد من الخيارات».\nالتثبيت لهذا المستخدم فقط، دون تغيير برنامج فتح الملفات الافتراضي.'); win11.setWordWrap(True); win11.setStyleSheet('color:#526174;'); layout.addWidget(win11)
        layout.addStretch(); self.path=QLabel(str(registration.installed_executable())); self.path.setWordWrap(True); self.path.setStyleSheet('font-size:11px;color:#526174;'); layout.addWidget(self.path)
        self.refresh()
    def refresh(self):
        status,count=registration.status()
        if status=='installed' and registration.installed_executable().is_file(): message=f'القسم مثبّت لـ{count} أنواع ملفات. يمكنك إغلاق هذه النافذة.'
        elif status=='conflict': message='يوجد تعارض تسجيل. لم تُغيّر أي مفاتيح غير مملوكة للبرنامج.'
        elif status=='absent': message='القسم غير مثبّت. اضغط «تثبيت / تحديث قسم الأدوات» مرة واحدة.'
        else: message='تحتاج القائمة تحديثًا. اضغط «تثبيت / تحديث قسم الأدوات» لإضافة جميع الأدوات.'
        self.state.setText(message); self.remove_button.setEnabled(status!='absent')
    def save_preferences(self):
        self.settings.setValue('outputMode',self.destination.currentData()); self.settings.setValue('notify',self.notices.isChecked())
        self.settings.setValue('codexAllowed',self.codex_allowed.isChecked()); self.settings.sync()
        self.state.setText('حُفظت إعدادات هذه النسخة المستقلة.')
    def installation_lock(self):
        parent=registration.installed_executable().parent; parent.mkdir(parents=True,exist_ok=True)
        lock=QLockFile(str(parent/'registration.lock')); lock.setStaleLockTime(0)
        if not lock.tryLock(1000): raise RuntimeError('توجد عملية تثبيت أو إزالة جارية. أعد المحاولة بعد اكتمالها.')
        return lock
    def install_menu(self):
        lock=None
        try:
            lock=self.installation_lock()
            target=registration.stage_executable(app_executable()); registration.install(target)
            self.refresh()
        except Exception as error: QMessageBox.warning(self,'تعذر تثبيت القسم',str(error))
        finally:
            if lock: lock.unlock()
    def remove_menu(self):
        lock=None
        try:
            lock=self.installation_lock(); registration.uninstall(); self.refresh()
        except Exception as error: QMessageBox.warning(self,'تعذر إزالة القسم',str(error))
        finally:
            if lock: lock.unlock()

class ProgressWindow(QDialog):
    def __init__(self,session):
        super().__init__(); self.session=session; self.setWindowTitle('Orbit — جارٍ تجهيز الملف'); self.setMinimumWidth(390)
        layout=QVBoxLayout(self); self.caption=QLabel(); self.caption.setWordWrap(True); layout.addWidget(self.caption)
        self.bar=QProgressBar(); self.bar.setRange(0,100); layout.addWidget(self.bar)
        self.cancel=QPushButton('إلغاء العملية'); self.cancel.clicked.connect(session.cancel); layout.addWidget(self.cancel)
    def closeEvent(self,event): self.session.cancel(); event.ignore()
    def reject(self): self.session.cancel()

class ToolSession(QObject):
    def __init__(self,action,files):
        super().__init__(); self.action=action; self.files=list(files); self.settings=preferences(); self.worker=None; self.outputs=[]; self.errors=[]
        self.progress=ProgressWindow(self); self.tray=None; self.folder=None; self.cancelling=False
    def run(self):
        try:
            if not self.files or not all(Path(p).is_file() and Path(p).suffix.lower() in SUPPORTED for p in self.files):
                raise ValueError('اختر صورة أو PDF أو مستند Word أو TXT مدعومًا موجودًا على الجهاز.')
            if len(self.files)!=1: raise ValueError('قسم الزر الأيمن مخصص لملف واحد في كل مرة.')
            source=Path(self.files[0]); is_pdf=source.suffix.lower()=='.pdf'; is_word=source.suffix.lower() in WORD_EXTENSIONS; is_txt=source.suffix.lower()=='.txt'
            if is_word and self.action!='word_pdf': raise ValueError('أداة مستندات Word هي «تحويل Word إلى PDF».')
            if not is_word and self.action=='word_pdf': raise ValueError('هذه الأداة مخصصة لمستندات Word فقط.')
            if is_txt and self.action not in TXT_ACTIONS: raise ValueError('أدوات TXT هي التحويل إلى Word أو تنسيق النص عبر Codex.')
            if not is_txt and self.action in TXT_ACTIONS: raise ValueError('هذه الأداة مخصصة لملفات TXT فقط.')
            if self.action in ('docx_ai','text_ai') and source.suffix.lower() not in IMAGE_PDF_EXTENSIONS:
                raise ValueError('هذه الأداة مخصصة للصور وملفات PDF فقط.')
            if is_pdf and self.action not in ('PNG','JPG','WEBP','BMP','TIFF','compress','docx_ai','text_ai'):
                raise ValueError('أدوات PDF هي التحويل إلى صور وتحسين الضغط أو التحويل الذكي إلى Word وTXT.')
            action=self.action; opts={}
            if action in SMART_ACTIONS:
                if not self.settings.value('codexAllowed',False,type=bool):
                    QMessageBox.information(None,'Orbit — السماح بالتحويل عبر Codex',
                        'التحويل الذكي متوقف لأن السماح بإرسال المحتوى غير مفعّل.\nافتح Orbit Context دون اختيار أداة، ثم فعّل «السماح بإرسال محتوى الملفات التي أحولها إلى Codex» واضغط «حفظ الإعدادات».')
                    self.finish_app(); return
                dialog=SmartOptionsDialog(None,source,action,self.settings)
                if dialog.exec()!=QDialog.DialogCode.Accepted: self.finish_app(); return
                opts=dialog.options()
            elif action=='word_pdf': pass
            elif action in FORMATS or action=='PDF': action='convert'; opts={'format':self.action}
            elif action in ('crop','annotate'):
                editor=ImageEditor(str(source),action,None)
                if editor.exec()!=QDialog.DialogCode.Accepted: self.finish_app(); return
                opts={'box':editor.canvas.crop_box()} if action=='crop' else {'image':editor.canvas.result_image()}
            else:
                opts=settings_dialog(None,action,str(source),files=self.files)
                if opts is None: self.finish_app(); return
            mode=str(self.settings.value('outputMode','original'))
            if mode=='original': folder=None
            elif mode=='last' and Path(str(self.settings.value('lastFolder',''))).is_dir() and self.settings.value('lastFolder',''):
                folder=str(self.settings.value('lastFolder'))
            else:
                folder=QFileDialog.getExistingDirectory(None,'اختر مجلد حفظ النسخة الجديدة',str(source.parent))
                if not folder: self.finish_app(); return
            if action in SMART_ACTIONS:
                try: from document_conversion import DocumentWorker
                except ImportError as error: raise RuntimeError('محرك التحويل الذكي غير متاح في هذه النسخة. استخدم الإصدار الكامل المرفق.') from error
                self.worker=DocumentWorker(self.files,action,opts,folder)
                self.worker.stage.connect(self.update_stage)
            else: self.worker=WordWorker(self.files,folder) if is_word else Worker(self.files,action,opts,folder)
            self.progress.caption.setText(('جارٍ التحويل عبر Codex في الخلفية: ' if action in SMART_ACTIONS else 'جارٍ تصدير PDF عبر Word في الخلفية: ' if is_word else 'جارٍ تجهيز: ')+source.name)
            self.worker.progress.connect(self.progress.bar.setValue); self.worker.finished.connect(self.completed)
            self.worker.start(); QTimer.singleShot(350,self.show_progress)
        except Exception as error:
            QMessageBox.warning(None,'Orbit — تعذر تنفيذ الأداة',str(error)); self.finish_app()
    def show_progress(self):
        if self.worker and self.worker.isRunning(): self.progress.show()
    def update_stage(self,text):
        if not self.cancelling: self.progress.caption.setText(str(text))
    def cancel(self):
        if self.worker and self.worker.isRunning():
            self.cancelling=True; self.worker.requestInterruption(); self.progress.cancel.setEnabled(False)
            text=('جارٍ إيقاف التحويل عبر Codex؛ تُلغى المهمة الجارية وتبقى النسخ التي اكتمل حفظها.' if self.action in SMART_ACTIONS else
                  'جارٍ الإلغاء؛ قد ننتظر انتهاء خطوة Word الحالية قبل إغلاقه.' if isinstance(self.worker,WordWorker) else
                  'جارٍ الإلغاء؛ تبقى النسخ التي اكتمل حفظها.')
            self.progress.caption.setText(text)
    def completed(self):
        worker=self.worker; self.worker=None
        self.outputs=list(worker.results); self.errors=list(worker.errors); cancelled=worker.cancelled; worker.deleteLater(); self.progress.hide()
        if self.outputs:
            self.folder=str(Path(self.outputs[0]).parent); self.settings.setValue('lastFolder',self.folder)
        if self.errors:
            QMessageBox.warning(None,'Orbit — تعذر حفظ بعض الملفات','\n\n'.join(self.errors)); self.finish_app(); return
        text=(f'حُفظ {len(self.outputs)} ملف جديد' if self.outputs else 'لم يُحفظ أي ملف')+(' · أُلغيت العملية' if cancelled else '')
        if not self.settings.value('notify',True,type=bool): self.finish_app(); return
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray=QSystemTrayIcon(QApplication.windowIcon(),self); menu=QMenu()
            if self.folder: menu.addAction('فتح مجلد النتائج',self.open_folder)
            menu.addAction('إغلاق',self.finish_app); self.tray.setContextMenu(menu); self.tray.setToolTip('Orbit Context · '+text)
            self.tray.messageClicked.connect(self.open_folder); self.tray.show(); self.tray.showMessage('Orbit Context',text,QSystemTrayIcon.MessageIcon.Information,3500)
            QTimer.singleShot(4500,self.finish_app)
        else:
            QMessageBox.information(None,'Orbit Context',text); self.finish_app()
    def open_folder(self):
        if self.folder: QDesktopServices.openUrl(QUrl.fromLocalFile(self.folder))
    def finish_app(self):
        if self.tray: self.tray.hide()
        QApplication.instance().quit()

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--action',choices=list(FORMATS)+['PDF','compress','crop','resize','rotate','annotate','word_pdf']+sorted(SMART_ACTIONS))
    parser.add_argument('--file',action='append',default=[])
    parser.add_argument('--export-menu-plan',metavar='PATH')
    parser.add_argument('--self-test',metavar='PATH')
    parser.add_argument('--word-self-test',metavar='PATH')
    parser.add_argument('--conversion-self-test',metavar='PATH')
    args=parser.parse_args()
    if args.export_menu_plan:
        import json
        Path(args.export_menu_plan).write_text(json.dumps(registration.build_plan(Path(sys.executable)),ensure_ascii=False,indent=2),encoding='utf-8'); return
    app=QApplication(sys.argv); app.setStyle('Fusion'); app.setStyleSheet(STYLE); app.setApplicationName('Orbit Context'); app.setApplicationVersion(VERSION); app.setWindowIcon(make_icon())
    for font in ('segoeui.ttf','segoeuib.ttf','seguisym.ttf'):
        path=Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'/font
        if path.exists(): QFontDatabase.addApplicationFont(str(path))
    app.setFont(QFont('Segoe UI',10))
    if args.conversion_self_test:
        try:
            from verify_conversion import run_checks
            run_checks(args.conversion_self_test); return
        except Exception:
            import json,traceback
            Path(args.conversion_self_test).write_text(json.dumps({'ok':False,'traceback':traceback.format_exc()},ensure_ascii=False,indent=2),encoding='utf-8')
            raise SystemExit(1)
    if args.word_self_test:
        try:
            from verify_word import run_checks
            run_checks(args.word_self_test); return
        except Exception:
            import json,traceback
            Path(args.word_self_test).write_text(json.dumps({'ok':False,'traceback':traceback.format_exc()},ensure_ascii=False,indent=2),encoding='utf-8')
            raise SystemExit(1)
    if args.self_test:
        import tempfile,json,hashlib
        from PIL import Image
        from file_tools import transform_file
        report=Path(args.self_test); report.parent.mkdir(parents=True,exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='orbit-context-check-',dir=report.parent) as td:
            folder=Path(td); image=folder/'sample.png'; Image.new('RGBA',(180,120),(230,80,20,150)).save(image)
            original=hashlib.sha256(image.read_bytes()).hexdigest()
            worker=Worker([str(image)],'convert',{'format':'PDF'},None); worker.start()
            assert worker.wait(10000) and not worker.errors and len(worker.results)==1
            pdf=worker.results[0]; pages=transform_file(pdf,'convert',{'format':'PNG'},folder)
            assert len(pages)==1 and Image.open(pages[0]).size==(270,180)
            assert hashlib.sha256(image.read_bytes()).hexdigest()==original
            result={'ok':True,'version':VERSION,'checks':['Qt runtime loaded','worker conversion without overwriting source','PDF image embedding and page rendering','same-folder destination','context command plan'],'registry_modified':False}
            assert len(registration.build_plan(Path(sys.executable)))==len(registration.root_names())
            report.write_text(json.dumps(result,indent=2),encoding='utf-8')
        return
    if args.action:
        app.setQuitOnLastWindowClosed(False); session=ToolSession(args.action,args.file); app._session=session; QTimer.singleShot(0,session.run)
    else:
        window=Manager(); window.show(); app._manager=window
    sys.exit(app.exec())

if __name__=='__main__': main()
