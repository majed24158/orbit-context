"""Per-user static Explorer cascades. No default file associations are changed."""
import os, sys, json, shutil, tempfile
from pathlib import Path
import winreg
from word_export import WORD_EXTENSIONS

BASE=r'Software\Classes'
OWNER='{D061F6F6-9AE1-48E6-978A-EC1062BF585E}'
OWNER_VALUE='OrbitContextOwner'
EXTENSIONS=('.png','.jpg','.jpeg','.webp','.bmp','.tif','.tiff','.gif','.pdf')+tuple(sorted(WORD_EXTENSIONS))+('.txt',)
IMAGE_VERBS='OrbitContextTools.Verbs.Image'
PDF_VERBS='OrbitContextTools.Verbs.PDF'
WORD_VERBS='OrbitContextTools.Verbs.Word'
TXT_VERBS='OrbitContextTools.Verbs.TXT'
IMAGE_FORMATS='OrbitContextTools.Formats.Image'
PDF_FORMATS='OrbitContextTools.Formats.PDF'
MENU_NAME='OrbitContextTools'

def root_names():
    return [IMAGE_VERBS,PDF_VERBS,WORD_VERBS,TXT_VERBS,IMAGE_FORMATS,PDF_FORMATS]+[
        rf'SystemFileAssociations\{ext}\shell\{MENU_NAME}' for ext in EXTENSIONS]

def command(exe,action):
    exe=str(Path(exe).absolute())
    if '"' in exe or any(c in exe for c in '\r\n'):
        raise ValueError('مسار البرنامج غير صالح.')
    allowed={'PNG','JPG','WEBP','BMP','TIFF','PDF','compress','crop','resize','rotate','annotate','word_pdf',
             'docx_ai','text_ai','txt_docx','txt_format'}
    if action not in allowed: raise ValueError('أداة غير مدعومة.')
    # Explorer calls the executable directly; no cmd.exe or PowerShell expansion.
    return f'"{exe}" --action {action} --file "%1"'

def build_plan(exe):
    """Map owned root -> relative key -> value name -> registry string."""
    roots={name:{'':{OWNER_VALUE:OWNER}} for name in root_names()}
    icon=f'"{Path(exe).absolute()}",0'
    def verb(root,name,label,action=None,cascade=None):
        key='shell\\'+name
        values={'MUIVerb':label,'MultiSelectModel':'Single','Icon':icon}
        if cascade: values['ExtendedSubCommandsKey']=cascade
        roots[root][key]=values
        if action: roots[root][key+'\\command']={'':command(exe,action)}
    for root,pdf in ((IMAGE_VERBS,False),(PDF_VERBS,True)):
        verb(root,'01.convert','تحويل صفحات PDF إلى صور' if pdf else 'تحويل الصيغة',cascade=PDF_FORMATS if pdf else IMAGE_FORMATS)
        verb(root,'02.compress','تحسين ضغط PDF' if pdf else 'ضغط الصورة',action='compress')
        verb(root,'03.docx_ai','تحويل إلى Word — Codex',action='docx_ai')
        verb(root,'04.text_ai','استخراج TXT — Codex',action='text_ai')
        if not pdf:
            for key,label in [('crop','قص الصورة'),('resize','تغيير المقاس'),('rotate','تدوير الصورة'),('annotate','الرسم على الصورة')]:
                verb(root,f'{len(roots[root]):02}.{key}',label,action=key)
    for root,formats in ((IMAGE_FORMATS,('PNG','JPG','WEBP','BMP','TIFF','PDF')),(PDF_FORMATS,('PNG','JPG','WEBP','BMP','TIFF'))):
        for i,fmt in enumerate(formats): verb(root,f'{i+1:02}.{fmt}',fmt,action=fmt)
    verb(WORD_VERBS,'01.word_pdf','تحويل Word إلى PDF',action='word_pdf')
    verb(TXT_VERBS,'01.txt_docx','تحويل TXT إلى Word — Codex',action='txt_docx')
    verb(TXT_VERBS,'02.txt_format','تنسيق TXT — Codex',action='txt_format')
    for ext in EXTENSIONS:
        name=rf'SystemFileAssociations\{ext}\shell\{MENU_NAME}'
        cascade=TXT_VERBS if ext=='.txt' else WORD_VERBS if ext in WORD_EXTENSIONS else PDF_VERBS if ext=='.pdf' else IMAGE_VERBS
        roots[name][''].update({'MUIVerb':'أدوات الملفات — Orbit','ExtendedSubCommandsKey':cascade,'MultiSelectModel':'Single','Icon':icon})
    return roots

def _path(base,root): return base+'\\'+root

def _snapshot(path):
    result={}
    def visit(relative):
        full=path+('\\'+relative if relative else '')
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,full,0,winreg.KEY_READ) as key:
            values={}; i=0
            while True:
                try: name,value,kind=winreg.EnumValue(key,i)
                except OSError: break
                values[name]=(value,kind); i+=1
            result[relative]=values; children=[]; i=0
            while True:
                try: children.append(winreg.EnumKey(key,i)); i+=1
                except OSError: break
        for child in children: visit(relative+'\\'+child if relative else child)
    try: visit('')
    except FileNotFoundError: return None
    return result

def _delete(path):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,path,0,winreg.KEY_READ|winreg.KEY_WRITE) as key:
            children=[]; i=0
            while True:
                try: children.append(winreg.EnumKey(key,i)); i+=1
                except OSError: break
        for child in children: _delete(path+'\\'+child)
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,path)
    except FileNotFoundError: pass

def _restore(path,snapshot):
    for relative,values in snapshot.items():
        full=path+('\\'+relative if relative else '')
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER,full,0,winreg.KEY_WRITE) as key:
            for name,(value,kind) in values.items(): winreg.SetValueEx(key,name,0,kind,value)

def _owned_snapshot(path):
    tree=_snapshot(path)
    if tree is not None and tree[''].get(OWNER_VALUE,(None,None))[0]!=OWNER:
        raise RuntimeError('يوجد قسم آخر بالاسم نفسه؛ لم يُجر أي تعديل عليه: '+path)
    return tree

def install(exe,base=BASE,notify=True):
    exe=Path(exe).absolute()
    if not exe.is_file(): raise FileNotFoundError('ملف تشغيل البرنامج غير موجود.')
    plan=build_plan(exe)
    # Preflight every root before changing any registration.
    old={root:_owned_snapshot(_path(base,root)) for root in plan}
    touched=[]
    try:
        for root,keys in plan.items():
            path=_path(base,root); touched.append(root); _delete(path)
            for relative,values in keys.items():
                target=path+('\\'+relative if relative else '')
                with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER,target,0,winreg.KEY_WRITE) as key:
                    for name,value in values.items(): winreg.SetValueEx(key,name,0,winreg.REG_SZ,value)
    except BaseException:
        for root in reversed(touched):
            path=_path(base,root); _delete(path)
            if old[root] is not None: _restore(path,old[root])
        raise
    if notify: notify_shell()
    return len(EXTENSIONS)

def uninstall(base=BASE,notify=True):
    # Only roots with the exact ownership marker are eligible for removal.
    roots={root:_owned_snapshot(_path(base,root)) for root in root_names()}
    for root,tree in roots.items():
        if tree is not None: _delete(_path(base,root))
    if notify: notify_shell()

def status(base=BASE):
    found=0
    for ext in EXTENSIONS:
        path=_path(base,rf'SystemFileAssociations\{ext}\shell\{MENU_NAME}')
        try:
            tree=_owned_snapshot(path)
            expected=TXT_VERBS if ext=='.txt' else WORD_VERBS if ext in WORD_EXTENSIONS else PDF_VERBS if ext=='.pdf' else IMAGE_VERBS
            if tree and tree[''].get('ExtendedSubCommandsKey',(None,None))[0]==expected: found+=1
        except RuntimeError: return 'conflict',found
    return ('installed' if found==len(EXTENSIONS) else 'partial' if found else 'absent'),found

def notify_shell():
    import ctypes
    ctypes.windll.shell32.SHChangeNotify(0x08000000,0,None,None)

def installed_executable():
    return Path(os.environ.get('LOCALAPPDATA',str(Path.home()/'AppData'/'Local')))/'OrbitContextTools'/'OrbitContext.exe'

def stage_executable(source,target=None):
    source=Path(source).absolute(); target=Path(target or installed_executable()).absolute()
    if not source.is_file(): raise FileNotFoundError(source)
    if source==target: return target
    target.parent.mkdir(parents=True,exist_ok=True)
    fd,temp=tempfile.mkstemp(prefix='.orbit-context-',suffix='.tmp',dir=target.parent); os.close(fd)
    try:
        shutil.copy2(source,temp); os.replace(temp,target)
    except PermissionError as error:
        raise RuntimeError('أغلق كل نوافذ أدوات النسخة المثبتة ثم أعد محاولة التثبيت.') from error
    finally: Path(temp).unlink(missing_ok=True)
    return target
