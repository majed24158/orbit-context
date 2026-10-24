"""Integration checks using locally installed Word, with scratch fixtures only."""
from pathlib import Path
import hashlib
import json
import tempfile
import pythoncom
import pymupdf
from win32com.client import DispatchEx
from word_export import export_word_pdf, WordWorker, initialize_sta, uninitialize_com
from file_tools import ConversionCancelled


def run_checks(report):
    report = Path(report).absolute()
    report.parent.mkdir(parents=True, exist_ok=True)
    checks = []
    initialize_sta()
    host = document = None
    scratch = tempfile.TemporaryDirectory(prefix='word-check-', dir=report.parent, ignore_cleanup_errors=True)
    try:
        with scratch as td:
            root = Path(td)
            host = DispatchEx('Word.Application')
            host.Visible = False
            host.DisplayAlerts = 0
            host.AutomationSecurity = 3
            word_version = str(host.Version)
            document = host.Documents.Add(Visible=False)
            document.Content.Text = 'اختبار تحويل المستند إلى PDF\rOrbit native Word export\r'
            document.Content.Font.Name = 'Arial'
            document.Content.Font.Size = 16
            table_range = document.Range(document.Content.End-1, document.Content.End-1)
            table = document.Tables.Add(table_range, 3, 2)
            table.Borders.Enable = True
            table.Cell(1, 1).Range.Text = 'اسم المستند'
            table.Cell(1, 2).Range.Text = 'تجربة عربية'
            table.Cell(2, 1).Range.Text = 'الحالة'
            table.Cell(2, 2).Range.Text = 'يحفظ التنسيق'
            table = table_range = None
            tail = document.Range(document.Content.End-1, document.Content.End-1)
            tail.InsertParagraphAfter()
            tail = document.Range(document.Content.End-1, document.Content.End-1)
            tail.InsertBreak(7)
            tail.InsertAfter('الصفحة الثانية — اختبار الهوامش وتعدد الصفحات')
            tail = None
            document.Sections(1).Footers(1).Range.Text = 'Orbit — اختبار التذييل'
            sources = []
            for extension, format_id in (('.docx',16),('.doc',0),('.docm',13),('.rtf',6)):
                path = root / ('مستند & اختبار' + extension)
                document.SaveAs2(FileName=str(path), FileFormat=format_id, AddToRecentFiles=False)
                sources.append(path)
            document.Close(SaveChanges=0)
            document = None
            # Keep a document with unsaved content open in another Word instance.
            document = host.Documents.Add(Visible=False)
            document.Content.Text = 'Unsaved sentinel — must remain open.'
            sentinel_text = document.Content.Text
            sentinel_saved = document.Saved
            hashes = {p:hashlib.sha256(p.read_bytes()).hexdigest() for p in sources}
            for source in sources:
                print('Export fixture: ' + source.suffix, flush=True)
                from unittest.mock import patch
                owned = []
                def tracked_dispatch(progid):
                    instance = DispatchEx(progid)
                    owned.append(instance)
                    return instance
                def inspect_hidden(value):
                    if value == 45:
                        assert not owned[0].Visible
                        assert owned[0].Documents(1).ReadOnly
                        assert not owned[0].Documents(1).Windows(1).Visible
                        assert owned[0].AutomationSecurity == 3
                try:
                    with patch('win32com.client.DispatchEx', side_effect=tracked_dispatch):
                        result = export_word_pdf(source, progress=inspect_hidden)
                finally:
                    owned.clear()
                assert result.parent == root and result.suffix == '.pdf'
                with pymupdf.open(result) as pdf:
                    assert len(pdf) == 2 and 'Orbit native Word export' in pdf[0].get_text()
                assert host.Documents.Count == 1
                assert document.Content.Text == sentinel_text and document.Saved == sentinel_saved
            checks.append('actual Word exports DOCX, DOC, DOCM and RTF as two-page PDFs with Arabic table and footer')
            checks.append('Word and document windows are invisible, input is read-only and macros are disabled during export')
            checks.append('a separate Word instance with unsaved content remains open and unchanged')
            # Compare the native export from the on-disk DOCX to our renderer.
            reference_doc = host.Documents.Open(FileName=str(sources[0]), ReadOnly=True,
                                                AddToRecentFiles=False, Visible=False)
            reference = root / 'reference.pdf'
            try:
                reference_doc.ExportAsFixedFormat(OutputFileName=str(reference), ExportFormat=17,
                    OpenAfterExport=False, OptimizeFor=0, Range=0, Item=0,
                    IncludeDocProps=True, KeepIRM=True, CreateBookmarks=1,
                    DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False)
            finally:
                reference_doc.Close(SaveChanges=0)
                reference_doc = None
            candidate = export_word_pdf(sources[0])
            with pymupdf.open(reference) as a, pymupdf.open(candidate) as b:
                assert len(a) == len(b)
                assert all(pa.rect == pb.rect and pa.get_pixmap().samples == pb.get_pixmap().samples
                           for pa, pb in zip(a, b))
            checks.append('rendered PDF pages match Word native export pixel-for-pixel for the fixture')
            # Existing PDFs must not be replaced; multiple formats share a stem.
            pdfs = sorted(root.glob('مستند*.pdf'))
            assert len(pdfs) == 5
            protected = root / 'protected.docx'
            encrypted = host.Documents.Add(Visible=False)
            try:
                encrypted.Content.Text = 'Protected test fixture'
                encrypted.Password = 'OrbitTest123'
                encrypted.SaveAs2(FileName=str(protected), FileFormat=16,
                                  Password='OrbitTest123', AddToRecentFiles=False)
                assert encrypted.HasPassword
            finally:
                encrypted.Close(SaveChanges=0)
                encrypted = None
            import zipfile
            assert not zipfile.is_zipfile(protected), 'fixture was saved without encryption'
            print('Check protected fixture', flush=True)
            before = set(root.iterdir())
            try:
                export_word_pdf(protected)
            except RuntimeError: pass
            else: raise AssertionError('password-protected file was accepted')
            assert set(root.iterdir()) == before
            checks.append('password-protected input fails without a password dialog or a partial PDF')
            cancel = [False]
            before = set(root.iterdir())
            try:
                export_word_pdf(sources[0], progress=lambda value: cancel.__setitem__(0,value>=85),
                                cancelled=lambda:cancel[0])
            except ConversionCancelled: pass
            else: raise AssertionError('cancelled export was published')
            assert set(root.iterdir()) == before
            checks.append('cancellation after Word export closes our instance and publishes no result')
            import pywintypes
            before = set(root.iterdir())
            with patch('win32com.client.DispatchEx', side_effect=pywintypes.com_error(-2147221164,'Not registered',None,None)):
                try: export_word_pdf(sources[0])
                except RuntimeError as error: assert 'Microsoft Word' in str(error)
                else: raise AssertionError('missing Word was not reported')
            assert set(root.iterdir()) == before
            checks.append('unavailable Word reports a clear error and removes all staging files')
            worker = WordWorker([str(sources[0])], str(root))
            worker.start()
            assert worker.wait(60000), 'Word worker did not finish'
            assert not worker.errors and not worker.cancelled and len(worker.results)==1, worker.errors
            assert all(hashlib.sha256(p.read_bytes()).hexdigest()==h for p,h in hashes.items())
            assert not list(root.glob('.orbit-*'))
            checks.append('Qt worker exports successfully, original documents stay byte-identical and temporary files are removed')
            document.Close(SaveChanges=0)
            document = None
    finally:
        try:
            if document is not None: document.Close(SaveChanges=0)
        finally:
            if host is not None: host.Quit(SaveChanges=0)
            host = document = None
            uninitialize_com()
            scratch.cleanup()
    result = {'ok':True,'checks':checks,'count':len(checks),'word_version':word_version,
              'scope':'real local Word COM integration with disposable fixtures; no production menu changes'}
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    return result


if __name__ == '__main__':
    import sys
    from PyQt6.QtWidgets import QApplication
    app = QApplication([])
    print(json.dumps(run_checks(sys.argv[1] if len(sys.argv)>1 else Path(__file__).parent/'qa'/'word-check.json'),
                     ensure_ascii=False,indent=2))
