"""Windows Office/WPS 自动化桥。仅临时副本被办公程序打开。"""
import os
from pathlib import Path

KINDS = {'.doc': ('word', '.docx'), '.wps': ('word', '.docx'),
         '.xls': ('excel', '.xlsx'), '.et': ('excel', '.xlsx'),
         '.ppt': ('ppt', '.pptx'), '.dps': ('ppt', '.pptx')}
IDS = {'word': ['Word.Application', 'kwps.Application'],
       'excel': ['Excel.Application', 'ket.Application'],
       'ppt': ['PowerPoint.Application', 'kwpp.Application']}


def convert(source, destination, kind, to_modern=True):
    if not to_modern and Path(destination).suffix.lower() in ('.wps', '.et', '.dps'):
        raise RuntimeError('WPS 专用格式原格式写回尚未验证，暂不保存')
    if os.name != 'nt':
        raise RuntimeError('旧版 Office/WPS 专用格式需在 Windows 上调用办公软件')
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError('未安装 Office/WPS 兼容组件 pywin32') from exc
    pythoncom.CoInitialize()
    app = doc = None
    errors = []
    try:
        for progid in IDS[kind]:
            # WPS 专用格式只使用 WPS，不借用 Word/Excel 进行猜测转换。
            if Path(source).suffix.lower() in ('.wps', '.et', '.dps') and progid == IDS[kind][0]:
                continue
            try:
                app = win32com.client.DispatchEx(progid)
                # 无法禁用自动宏的程序不打开文件。
                app.AutomationSecurity = 3
                break
            except Exception as exc:
                errors.append(str(exc))
                if app:
                    try:
                        app.Quit()
                    except Exception:
                        pass
                app = None
        if app is None:
            raise RuntimeError('没有可用的 Office/WPS 自动化接口：' + '; '.join(errors))
        source, destination = str(Path(source).resolve()), str(Path(destination).resolve())
        if kind == 'word':
            app.Visible = False
            app.DisplayAlerts = 0
            doc = app.Documents.Open(source, ReadOnly=True, AddToRecentFiles=False)
            fmt = 12 if to_modern else 0
            doc.SaveAs2(destination, FileFormat=fmt)
        elif kind == 'excel':
            app.Visible = False
            app.DisplayAlerts = False
            app.AskToUpdateLinks = False
            doc = app.Workbooks.Open(source, UpdateLinks=0, ReadOnly=True, AddToMru=False)
            # WPS 专用回写枚举未经验证，明确拒绝，避免写错格式。
            if not to_modern and Path(destination).suffix.lower() == '.et':
                raise RuntimeError('ET 原格式回写尚未验证，暂不保存')
            doc.SaveAs(destination, FileFormat=51 if to_modern else 56)
        else:
            doc = app.Presentations.Open(source, ReadOnly=True, Untitled=False, WithWindow=False)
            if not to_modern and Path(destination).suffix.lower() == '.dps':
                raise RuntimeError('DPS 原格式回写尚未验证，暂不保存')
            doc.SaveAs(destination, FileFormat=24 if to_modern else 1)
    finally:
        if doc is not None:
            try:
                if kind == 'excel':
                    doc.Close(SaveChanges=False)
                elif kind == 'word':
                    doc.Close(SaveChanges=0)
                else:
                    doc.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
