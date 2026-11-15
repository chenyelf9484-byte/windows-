"""打包后的 EXE 自检，仅使用临时样本，不访问用户文档。"""
import json
import tempfile
import traceback
from pathlib import Path


def run(report_path, with_ocr=False):
    report = {'passed': False, 'checks': []}
    app = None
    try:
        from .gui import App
        from .rules import Matcher, Rule
        from .engine import scan_file, save_file
        app = App()
        app.withdraw()
        app.update()
        assert len(app.tabs.tabs()) == 3
        report['checks'].append('Tk window and resources loaded')
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'中文 sample.txt'
            path.write_bytes('旧公司\r\n'.encode('utf-8'))
            result = scan_file(path, Matcher([Rule('旧公司', '新企业')]))
            assert not result.error and len(result.units) == 1
            saved = save_file(result, 'direct')
            assert saved.status == '成功', saved.detail
            assert path.read_bytes() == '新企业\r\n'.encode('utf-8')
        report['checks'].append('Chinese path, scan and direct save passed')
        if with_ocr:
            from .ocr import OCR
            OCR()
            report['checks'].append('Bundled OCR runtime and models loaded')
        report['passed'] = True
    except Exception:
        report['error'] = traceback.format_exc()
    finally:
        if app is not None:
            app.destroy()
        Path(report_path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0 if report['passed'] else 1
