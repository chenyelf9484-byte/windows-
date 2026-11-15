"""在有桌面会话的机器上进行 Tk 窗口冒烟验证。"""
import sys
import tempfile
import time
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from docreplace.gui import App

app = App()
app.update()
assert len(app.tabs.tabs()) == 3
assert not app.ocr.get()
assert app.mode.get() == 'copy'
app.source.set('旧公司')
app.target.set('新企业')
app.add_rule()
assert app.matcher().replace('旧公司') == '新企业'
with tempfile.TemporaryDirectory() as folder:
    sample = Path(folder)/'界面测试.txt'
    sample.write_text('旧公司\n', encoding='utf-8')
    app.add_path(str(sample))
    def wait_task():
        deadline = time.monotonic()+15
        while app.busy and time.monotonic() < deadline:
            app.update()
            time.sleep(.02)
        assert not app.busy, 'GUI background task timed out'
    with patch('docreplace.gui.messagebox.showerror', side_effect=RuntimeError('GUI task failed')):
        app.start_scan()
        wait_task()
        assert len(app.results) == 1
        assert len(app.results[0].units) == 1
        app.mode.set('direct')
        with patch('docreplace.gui.messagebox.askyesno', return_value=True):
            app.start_save()
            wait_task()
        assert sample.read_text(encoding='utf-8') == '新企业\n'
        assert app.reports[0]['status'] == '成功'
app.destroy()
print('GUI smoke check passed')
