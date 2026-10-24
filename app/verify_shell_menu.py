"""Ask Windows Shell to construct a menu for an isolated test-only extension; never show/invoke it."""
from pathlib import Path
import sys,uuid,json,ctypes,tempfile
from ctypes import wintypes
import registration as r

root=Path(__file__).parent; (root/'qa').mkdir(exist_ok=True)
token=uuid.uuid4().hex; extension='.orbitcontextqa'+token
word_mode='--word' in sys.argv
if word_mode: r.WORD_EXTENSIONS=frozenset({extension})
r.EXTENSIONS=(extension,)
r.IMAGE_VERBS='OrbitContextQA.'+token+'.Image'; r.PDF_VERBS='OrbitContextQA.'+token+'.PDF'
r.WORD_VERBS='OrbitContextQA.'+token+'.Word'
r.TXT_VERBS='OrbitContextQA.'+token+'.TXT'
r.IMAGE_FORMATS='OrbitContextQA.'+token+'.ImageFormats'; r.PDF_FORMATS='OrbitContextQA.'+token+'.PDFFormats'
r.MENU_NAME='OrbitContextQA.'+token
user=ctypes.WinDLL('user32'); user.GetMenuStringW.argtypes=[wintypes.HMENU,ctypes.c_uint,wintypes.LPWSTR,ctypes.c_int,ctypes.c_uint]; user.GetMenuStringW.restype=ctypes.c_int
user.GetSubMenu.argtypes=[wintypes.HMENU,ctypes.c_int]; user.GetSubMenu.restype=wintypes.HMENU
user.GetMenuItemCount.argtypes=[wintypes.HMENU]; user.GetMenuItemCount.restype=ctypes.c_int
user.CreatePopupMenu.restype=wintypes.HMENU
user.DestroyMenu.argtypes=[wintypes.HMENU]
ole=ctypes.OleDLL('ole32'); sh=ctypes.OleDLL('shell32')
ole.CoInitializeEx.argtypes=[ctypes.c_void_p,wintypes.DWORD]
ole.CoTaskMemFree.argtypes=[ctypes.c_void_p]
def guid(text): return (ctypes.c_ubyte*16).from_buffer_copy(uuid.UUID(text).bytes_le)
def method(obj,index,result,*args):
    table=ctypes.cast(obj,ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    return ctypes.WINFUNCTYPE(result,ctypes.c_void_p,*args)(table[index])
def release(obj):
    if obj: method(obj,2,wintypes.ULONG)(obj)
sh.SHParseDisplayName.argtypes=[wintypes.LPCWSTR,ctypes.c_void_p,ctypes.POINTER(ctypes.c_void_p),wintypes.ULONG,ctypes.POINTER(wintypes.ULONG)]
sh.SHBindToParent.argtypes=[ctypes.c_void_p,ctypes.c_void_p,ctypes.POINTER(ctypes.c_void_p),ctypes.POINTER(ctypes.c_void_p)]
ole.CoInitializeEx(None,2)
def menu_items(menu,context,depth=0):
    result=[]
    for i in range(user.GetMenuItemCount(menu)):
        text=ctypes.create_unicode_buffer(1024); user.GetMenuStringW(menu,i,text,1024,0x400)
        sub=user.GetSubMenu(menu,i)
        children=[]
        if sub and depth<3 and (depth or 'أدوات الملفات' in text.value):
            # Populate lazy static cascades exactly as a menu host does, without displaying them.
            method(context,6,ctypes.c_long,wintypes.UINT,wintypes.WPARAM,wintypes.LPARAM)(context,0x117,sub,i)
            children=menu_items(sub,context,depth+1)
        result.append({'label':text.value,'children':children})
    return result
with tempfile.TemporaryDirectory(prefix='shell_',dir=root/'qa') as td:
    path=Path(td)/('test'+extension); path.write_text('isolated shell fixture')
    menu=None; parent=ctypes.c_void_p(); context=ctypes.c_void_p(); pidl=ctypes.c_void_p()
    try:
        r.install(Path(sys.executable),notify=True)
        attrs=wintypes.ULONG(); sh.SHParseDisplayName(str(path),None,ctypes.byref(pidl),0,ctypes.byref(attrs))
        child=ctypes.c_void_p(); iid_folder=guid('000214E6-0000-0000-C000-000000000046')
        sh.SHBindToParent(pidl,ctypes.byref(iid_folder),ctypes.byref(parent),ctypes.byref(child))
        iid_context=guid('000214E4-0000-0000-C000-000000000046'); children=(ctypes.c_void_p*1)(child.value)
        hr=method(parent,10,ctypes.c_long,wintypes.HWND,wintypes.UINT,ctypes.c_void_p,ctypes.c_void_p,ctypes.c_void_p,ctypes.POINTER(ctypes.c_void_p))(parent,None,1,children,ctypes.byref(iid_context),None,ctypes.byref(context))
        assert hr>=0,hex(hr&0xffffffff)
        extended=ctypes.c_void_p(); iid_extended=guid('000214F4-0000-0000-C000-000000000046')
        hr=method(context,0,ctypes.c_long,ctypes.c_void_p,ctypes.POINTER(ctypes.c_void_p))(context,ctypes.byref(iid_extended),ctypes.byref(extended))
        assert hr>=0,hex(hr&0xffffffff)
        release(context); context=extended
        menu=user.CreatePopupMenu()
        hr=method(context,3,ctypes.c_long,wintypes.HMENU,wintypes.UINT,wintypes.UINT,wintypes.UINT,wintypes.UINT)(context,menu,0,1,0x7FFF,0)
        assert hr>=0,hex(hr&0xffffffff)
        items=menu_items(menu,context)
        ours=next((x for x in items if 'أدوات الملفات' in x['label']),None)
        assert ours is not None,'Windows Shell did not build the owned menu'
        if word_mode:
            assert len(ours['children'])==1 and 'Word' in ours['children'][0]['label'],ours
        else:
            assert len(ours['children'])==8,ours
            convert=next(x for x in ours['children'] if 'تحويل' in x['label'])
            assert len(convert['children'])==6,convert
        report={'ok':True,'scope':'Windows Shell IContextMenu metadata on a temporary unique extension; no menu was displayed and no verb invoked','menu':ours}
        (root/'qa'/('shell-word-menu.json' if word_mode else 'shell-menu.json')).write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(report,ensure_ascii=False,indent=2))
    finally:
        if menu: user.DestroyMenu(menu)
        release(context); release(parent)
        if pidl: ole.CoTaskMemFree(pidl)
        r.uninstall(notify=True)
        # Only remove empty ancestors uniquely created for this fixture extension.
        for relative in (f'SystemFileAssociations\\{extension}\\shell',f'SystemFileAssociations\\{extension}'):
            try: r.winreg.DeleteKey(r.winreg.HKEY_CURRENT_USER,r.BASE+'\\'+relative)
            except OSError: pass
ole.CoUninitialize()
