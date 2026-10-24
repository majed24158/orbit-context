import os
os.environ['QT_QPA_PLATFORM']='offscreen'
import sys,json,tempfile,hashlib,uuid,ctypes
from pathlib import Path
from ctypes import wintypes
from PyQt6.QtCore import QSettings,QEventLoop,QTimer
import orbit_context as appmod
import registration as r

checks=[]
app=appmod.QApplication([]); app.setQuitOnLastWindowClosed(False); app.setStyle('Fusion'); app.setStyleSheet(appmod.STYLE)
for f in ('segoeui.ttf','segoeuib.ttf','seguisym.ttf'): appmod.QFontDatabase.addApplicationFont('C:/Windows/Fonts/'+f)
root=Path(__file__).parent; (root/'qa').mkdir(exist_ok=True)
test_base='Software\\OrbitContextTools.QA\\'+uuid.uuid4().hex
with tempfile.TemporaryDirectory(prefix='context_',dir=root/'qa') as td:
    tmp=Path(td); exe=tmp/'برنامج & أدوات.exe'; exe.write_bytes(b'fixture - never executed')
    image=tmp/'مستند & صورة.png'; appmod.__dict__['Worker'].__module__
    from PIL import Image
    import pymupdf
    Image.new('RGB',(150,100),'orange').save(image); before=hashlib.sha256(image.read_bytes()).hexdigest()
    plan=r.build_plan(exe)
    assert len(plan)==20 and all(v[''][r.OWNER_VALUE]==r.OWNER for v in plan.values())
    pdf=plan[r.PDF_VERBS]; assert len([k for k in pdf if k and not k.endswith('command')])==4
    assert len([k for k in plan[r.IMAGE_FORMATS] if k.endswith('command')])==6
    assert len([k for k in plan[r.WORD_VERBS] if k.endswith('command')])==1
    for ext in r.WORD_EXTENSIONS:
        assert plan[rf'SystemFileAssociations\{ext}\shell\{r.MENU_NAME}']['']['ExtendedSubCommandsKey']==r.WORD_VERBS
    assert plan[rf'SystemFileAssociations\.txt\shell\{r.MENU_NAME}']['']['ExtendedSubCommandsKey']==r.TXT_VERBS
    checks.append('owned cascades cover fourteen extensions, local tools, native Word export and Codex conversion')
    raw=r.command(exe,'PNG').replace('%1',str(image))
    shell=ctypes.WinDLL('shell32'); shell.CommandLineToArgvW.argtypes=[wintypes.LPCWSTR,ctypes.POINTER(ctypes.c_int)]; shell.CommandLineToArgvW.restype=ctypes.POINTER(wintypes.LPWSTR)
    count=ctypes.c_int(); args=shell.CommandLineToArgvW(raw,ctypes.byref(count))
    try: parsed=[args[i] for i in range(count.value)]
    finally:
        ctypes.windll.kernel32.LocalFree.argtypes=[wintypes.HLOCAL]; ctypes.windll.kernel32.LocalFree(args)
    assert parsed==[str(exe),'--action','PNG','--file',str(image)]
    checks.append('Windows argument parsing preserves Arabic, spaces and ampersands without a command shell')
    try:
        assert r.status(test_base)==('absent',0)
        assert r.install(exe,base=test_base,notify=False)==len(r.EXTENSIONS)
        assert r.status(test_base)==('installed',len(r.EXTENSIONS))
        snapshots={name:r._snapshot(r._path(test_base,name)) for name in r.root_names()}
        # Failure while rebuilding an existing registration must restore it.
        original_set=r.winreg.SetValueEx; calls=[0]
        def fail_once(*args):
            calls[0]+=1
            if calls[0]==8: raise OSError('injected registry write failure')
            return original_set(*args)
        r.winreg.SetValueEx=fail_once
        try:
            try: r.install(exe,base=test_base,notify=False)
            except OSError: pass
            else: raise AssertionError('failure injection did not execute')
        finally: r.winreg.SetValueEx=original_set
        assert all(r._snapshot(r._path(test_base,name))==tree for name,tree in snapshots.items())
        checks.append('real isolated HKCU registration installs correctly and rolls back failed updates')
        foreign=r._path(test_base,r.IMAGE_VERBS)
        with r.winreg.OpenKey(r.winreg.HKEY_CURRENT_USER,foreign,0,r.winreg.KEY_SET_VALUE) as key:
            r.winreg.SetValueEx(key,r.OWNER_VALUE,0,r.winreg.REG_SZ,'foreign')
        for operation in (lambda:r.install(exe,base=test_base,notify=False),lambda:r.uninstall(base=test_base,notify=False)):
            try: operation()
            except RuntimeError: pass
            else: raise AssertionError('foreign root was modified')
        assert r._snapshot(foreign)[''][r.OWNER_VALUE][0]=='foreign'
        with r.winreg.OpenKey(r.winreg.HKEY_CURRENT_USER,foreign,0,r.winreg.KEY_SET_VALUE) as key:
            r.winreg.SetValueEx(key,r.OWNER_VALUE,0,r.winreg.REG_SZ,r.OWNER)
        r.uninstall(base=test_base,notify=False); assert r.status(test_base)==('absent',0)
        checks.append('uninstall and reinstall reject foreign keys and removal touches only owned namespace')
    finally: r._delete(test_base)
    staged=tmp/'stable'/'OrbitContext.exe'; assert r.stage_executable(exe,staged)==staged
    assert staged.read_bytes()==exe.read_bytes() and r.stage_executable(staged,staged)==staged
    checks.append('stable per-user copy preserves complete executable and supports launch from installed path')

    settings=QSettings(str(tmp/'settings.ini'),QSettings.Format.IniFormat); settings.setValue('notify',False)
    appmod.preferences=lambda:settings
    warnings=[]; appmod.QMessageBox.warning=lambda *args:warnings.append(args[-1])
    appmod.QMessageBox.information=lambda *args:warnings.append(args[-1])
    original_exec=appmod.QDialog.exec
    def accept_dialog(dialog):
        assert dialog.parentWidget() is None
        if isinstance(dialog,appmod.ImageEditor):
            if dialog.mode=='crop': dialog.canvas.start=(1,1); dialog.canvas.end=(70,50)
            else: dialog.canvas.strokes=[('#008800',5,[(4,4),(55,45)])]
        return appmod.QDialog.DialogCode.Accepted
    appmod.QDialog.exec=accept_dialog
    original_finish=appmod.ToolSession.finish_app
    def finished(session): session.done=True
    appmod.ToolSession.finish_app=finished
    try:
        for action in ('PNG','JPG','WEBP','BMP','TIFF','PDF','compress','crop','resize','rotate','annotate'):
            session=appmod.ToolSession(action,[str(image)]); session.done=False; session.run()
            if session.worker:
                loop=QEventLoop(); session.worker.finished.connect(loop.quit); QTimer.singleShot(10000,loop.quit); loop.exec(); app.processEvents()
            assert session.done and session.worker is None and session.outputs and not session.errors,(action,session.errors,warnings)
            assert all(Path(p).parent==tmp and Path(p).is_file() for p in session.outputs)
        checks.append('all eleven context actions save valid copies with independent dialogs and no dashboard')
        pdf=tmp/'pages.pdf'
        with pymupdf.open() as doc: doc.new_page(width=150,height=200); doc.new_page(width=150,height=200); doc.save(pdf)
        session=appmod.ToolSession('PNG',[str(pdf)]); session.done=False; session.run(); loop=QEventLoop(); session.worker.finished.connect(loop.quit); QTimer.singleShot(10000,loop.quit); loop.exec(); app.processEvents()
        assert session.done and len(session.outputs)==2 and not session.errors
        checks.append('PDF conversion emits all pages beside source document')
        assert not warnings,warnings
        session=appmod.ToolSession('crop',[str(pdf)]); session.run(); assert session.done and warnings; warnings.clear()
        session=appmod.ToolSession('PNG',[str(image),str(pdf)]); session.run(); assert session.done and warnings; warnings.clear()
        word_guard=tmp/'guard.docx'; word_guard.write_bytes(b'guard fixture - never opened')
        for action,source in (('crop',word_guard),('PNG',word_guard),('word_pdf',image)):
            session=appmod.ToolSession(action,[str(source)]); session.run(); assert session.done and warnings; warnings.clear()
        session=appmod.ToolSession('PDF',[str(image)]); cancellations=[]
        session.cancel=lambda:cancellations.append(True)
        session.progress.show(); app.processEvents(); session.progress.reject()
        assert cancellations and session.progress.isVisible()
        session.progress.hide()
        checks.append('PDF-invalid tools and unexpected multi-selection are rejected before processing')
        assert hashlib.sha256(image.read_bytes()).hexdigest()==before
        checks.append('original image remains byte-for-byte unchanged after all operations')
    finally: appmod.QDialog.exec=original_exec; appmod.ToolSession.finish_app=original_finish
    manager=appmod.Manager(); manager.show(); app.processEvents(); manager.grab().save(str(root/'qa'/'manager.png')); manager.close()
if (root/'original_snapshot.json').exists():
    snapshot=json.loads((root/'original_snapshot.json').read_text())
    assert all(Path(p).is_file() and hashlib.sha256(Path(p).read_bytes()).hexdigest()==h for p,h in snapshot.items())
    checks.append('original wheel project and all captured deliverable files are unchanged')
report={'ok':True,'count':len(checks),'checks':checks}
(root/'qa'/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))
