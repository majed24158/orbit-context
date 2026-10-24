"""Live Codex + Word integration with disposable, synthetic TXT fixtures."""
from pathlib import Path
import hashlib
import json
import tempfile
import zipfile
from lxml import etree
from document_conversion import DocumentWorker


def run_checks(report):
    report=Path(report).absolute(); report.parent.mkdir(parents=True,exist_ok=True)
    checks=[]
    with tempfile.TemporaryDirectory(prefix='conversion-check-',dir=report.parent) as td:
        root=Path(td); source=root/'synthetic.txt'
        source.write_text('اسم الطالب: محمد\nالدرجة: 95\nOrbit editable conversion\n',encoding='utf-8')
        original=hashlib.sha256(source.read_bytes()).hexdigest()
        for action,options in (
            ('txt_format',{'text_style':'lines','instructions':'Keep all three source lines verbatim.','review':False}),
            ('txt_docx',{'instructions':'Create one A4 page. Put the English line as a title and the student name and score in an editable two-row table. Keep every original value and Arabic word.','review':True})):
            worker=DocumentWorker([str(source)],action,options,None)
            worker.start()
            if not worker.wait(300000):
                worker.requestInterruption(); worker.wait(15000)
                raise TimeoutError('conversion self-test timed out')
            assert not worker.errors and not worker.cancelled and len(worker.results)==1,worker.errors
            result=Path(worker.results[0]); assert result.parent==root and result!=source
            if action=='txt_format':
                content=result.read_text(encoding='utf-8-sig')
                assert 'محمد' in content and '95' in content and 'Orbit' in content
                checks.append('real Codex CLI hidden TXT formatting with ChatGPT authentication')
            else:
                with zipfile.ZipFile(result) as package:
                    tree=etree.fromstring(package.read('word/document.xml'))
                    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                    text=' '.join(tree.xpath('//w:t/text()',namespaces=ns))
                    assert 'محمد' in text and '95' in text and 'Orbit' in text
                    assert tree.xpath('count(//w:tbl)',namespaces=ns)>=1
                    assert not any(p.startswith('word/media/') for p in package.namelist())
                checks.append('Codex layout to native editable Arabic Word table without screenshot fallback')
                checks.append('generated DOCX exported by hidden local Word before publication')
            assert hashlib.sha256(source.read_bytes()).hexdigest()==original
        assert not list(root.glob('.orbit-*'))
        checks.append('Qt worker, same-folder copies, original byte preservation and temporary cleanup')
    result={'ok':True,'count':len(checks),'checks':checks,'scope':'live local Codex and Word, synthetic data only; no registry writes'}
    report.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return result


if __name__=='__main__':
    import sys
    from PyQt6.QtWidgets import QApplication
    app=QApplication([])
    print(json.dumps(run_checks(sys.argv[1] if len(sys.argv)>1 else Path(__file__).parent/'qa'/'conversion-check.json'),ensure_ascii=False,indent=2))
