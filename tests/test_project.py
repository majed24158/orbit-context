import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
import sys, tempfile, unittest, zipfile
from pathlib import Path
APP = Path(__file__).resolve().parents[1] / 'app'
sys.path.insert(0, str(APP))
import registration
from file_tools import atomic_output
from document_conversion import read_text
from editable_docx import build_docx, validate_pages


class ProjectChecks(unittest.TestCase):
    def test_registration_plan_is_data_only(self):
        plan = registration.build_plan(Path('OrbitContext.exe').absolute())
        self.assertEqual(len(plan), 20)
        self.assertEqual(len(registration.EXTENSIONS), 14)
        self.assertIn(registration.TXT_VERBS, plan)
        commands = [v[''] for tree in plan.values() for k, v in tree.items() if k.endswith('command')]
        self.assertTrue(any('--action docx_ai' in command for command in commands))
        self.assertTrue(all(' --file "%1"' in command for command in commands))

    def test_atomic_output_keeps_existing_file(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'result.txt'; path.write_bytes(b'original')
            output = atomic_output(path, lambda stage: stage.write_bytes(b'new'))
            self.assertNotEqual(output, path)
            self.assertEqual(path.read_bytes(), b'original')
            self.assertEqual(output.read_bytes(), b'new')

    def test_text_limit(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / 'large.txt'
            with source.open('wb') as handle: handle.truncate(1_000_001)
            with self.assertRaises(ValueError): read_text(source)

    def test_editable_document_has_native_text(self):
        block = {'kind':'text', 'bbox':[.1,.1,.8,.1], 'text':'نص قابل للتعديل',
                 'font_size':14, 'bold':False, 'italic':False, 'color':'#000000',
                 'align':'right', 'rtl':True, 'rows':[], 'column_widths':[],
                 'header':False, 'merges':[]}
        pages = [{'width_pt':595, 'height_pt':842, 'blocks':[block]}]
        validated = validate_pages(pages)
        self.assertIsNot(validated, pages)
        with tempfile.TemporaryDirectory() as temporary:
            path = build_docx(pages, Path(temporary) / 'editable.docx')
            with zipfile.ZipFile(path) as package:
                xml = package.read('word/document.xml').decode('utf-8')
                self.assertIn('txbxContent', xml)
                self.assertIn('نص قابل للتعديل', xml)
                self.assertFalse(any(n.startswith('word/media/') for n in package.namelist()))


if __name__ == '__main__': unittest.main()
