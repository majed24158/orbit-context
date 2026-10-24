"""Word's native PDF renderer in an owned, invisible COM instance."""
from pathlib import Path
import shutil
import tempfile
import secrets
import ctypes
import pythoncom  # Initialize pywin32 on the main thread before QApplication.
from PyQt6.QtCore import QThread, pyqtSignal
from file_tools import atomic_output, check_cancelled, ConversionCancelled

WORD_EXTENSIONS = frozenset({'.doc', '.docx', '.docm', '.rtf'})

_ole32 = ctypes.WinDLL('ole32')
_ole32.CoInitializeEx.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
_ole32.CoInitializeEx.restype = ctypes.c_long
_ole32.CoUninitialize.argtypes = []
_ole32.CoUninitialize.restype = None


def initialize_sta():
    # Use the native reference count: pywin32's initializer-thread shortcut
    # does not increment it for nested calls, so its uninitializer can detach
    # another owner's COM proxies. Only a successful native call is balanced.
    result = _ole32.CoInitializeEx(None, 2)
    if result < 0:
        raise RuntimeError('تعذر تهيئة اتصال Word في خيط العمل: '+hex(result & 0xffffffff))


def uninitialize_com():
    _ole32.CoUninitialize()


def export_word_pdf(source, folder=None, progress=lambda value: None, cancelled=None):
    """Open read-only in a separate Word instance; never attach to running Word.

    Cancellation is cooperative at COM-call boundaries. A completed export is
    published only after our document and application have been closed.
    """
    import pywintypes
    from win32com.client import DispatchEx

    source = Path(source).absolute()
    if not source.is_file() or source.suffix.lower() not in WORD_EXTENSIONS:
        raise ValueError('اختر مستند Word بصيغة DOC أو DOCX أو DOCM أو RTF.')
    destination = source.parent if folder is None else Path(folder).absolute()
    if not destination.is_dir():
        raise ValueError('مجلد الحفظ غير موجود.')
    check_cancelled(cancelled)
    initialize_sta()
    try:
        def write_export(target):
            word = document = None
            owned = False
            old_security = old_links = None
            with tempfile.TemporaryDirectory(prefix='.orbit-word-', dir=destination) as staging:
                pdf = Path(staging) / 'export.pdf'
                try:
                    try:
                        word = DispatchEx('Word.Application')
                    except pywintypes.com_error as error:
                        raise RuntimeError('تعذر تشغيل Microsoft Word. تأكد من تثبيت نسخة سطح المكتب وتفعيلها على الجهاز.') from error
                    if word.Documents.Count or word.UserControl or word.Visible:
                        # Refuse an existing/user-controlled instance before any
                        # mutation, including cleanup. DispatchEx should be new.
                        word = None
                        raise RuntimeError('تعذر إنشاء جلسة Word مستقلة مخفية. لم تُغيّر جلسة Word المفتوحة.')
                    owned = True
                    word.Visible = False
                    word.DisplayAlerts = 0
                    word.ScreenUpdating = False
                    old_security = word.AutomationSecurity
                    word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                    old_links = word.Options.UpdateLinksAtOpen
                    word.Options.UpdateLinksAtOpen = False
                    progress(15)
                    check_cancelled(cancelled)
                    # An explicit random password fails protected inputs instead of
                    # leaving an invisible password dialog waiting for user input.
                    password = secrets.token_hex(6)
                    document = word.Documents.Open(
                        FileName=str(source), ConfirmConversions=False, ReadOnly=True,
                        AddToRecentFiles=False, PasswordDocument=password,
                        PasswordTemplate=password, WritePasswordDocument=password,
                        WritePasswordTemplate=password, Visible=False,
                        OpenAndRepair=False, NoEncodingDialog=True)
                    if document.HasPassword:
                        raise RuntimeError('المستند محمي بكلمة مرور. احفظ نسخة غير محمية للتحويل في الخلفية.')
                    progress(45)
                    check_cancelled(cancelled)
                    document.ExportAsFixedFormat(
                        OutputFileName=str(pdf), ExportFormat=17, OpenAfterExport=False,
                        OptimizeFor=0, Range=0, Item=0, IncludeDocProps=True,
                        KeepIRM=True, CreateBookmarks=1, DocStructureTags=True,
                        BitmapMissingFonts=True, UseISO19005_1=False)
                    progress(85)
                finally:
                    # Quit only the DispatchEx instance we own. Never use
                    # GetActiveObject or terminate an existing WINWORD process.
                    try:
                        if document is not None:
                            document.Close(SaveChanges=0)
                    finally:
                        document = None
                        if word is not None and owned:
                            try:
                                if old_links is not None: word.Options.UpdateLinksAtOpen = old_links
                                if old_security is not None: word.AutomationSecurity = old_security
                            finally:
                                try:
                                    if not word.Documents.Count and not word.UserControl and not word.Visible:
                                        word.Quit(SaveChanges=0)
                                finally: word = None
                check_cancelled(cancelled)
                with pdf.open('rb') as handle:
                    if handle.read(5) != b'%PDF-':
                        raise RuntimeError('لم ينتج Word ملف PDF صالحًا.')
                shutil.copyfile(pdf, target)
        result = atomic_output(destination / (source.stem + '.pdf'), write_export, cancelled)
        progress(100)
        return result

    except pywintypes.com_error as error:
        detail = (error.excepinfo[2] if error.excepinfo and error.excepinfo[2] else '')
        raise RuntimeError('تعذر تصدير المستند عبر Word. تأكد أنه صالح وغير محمي بكلمة مرور، وأن Word مفعّل.'
                           + ('\n' + detail.strip() if detail else '')) from error
    finally:
        uninitialize_com()


class WordWorker(QThread):
    progress = pyqtSignal(int)

    def __init__(self, files, folder):
        super().__init__()
        self.files = tuple(files)
        self.folder = folder
        self.results = []
        self.errors = []
        self.cancelled = False

    def run(self):
        try:
            for source in self.files:
                self.results.append(export_word_pdf(source, self.folder,
                    self.progress.emit, self.isInterruptionRequested))
        except ConversionCancelled:
            self.cancelled = True
        except Exception as error:
            self.errors.append(Path(source).name + '\n' + str(error))
