# 文档替换工具 0.1.1

Windows 10 x64 便携 EXE 项目。直接扫描磁盘、文件夹或文件，预览后批量替换。
已提供源码、自动测试和 Windows 自动构建。上一版已通过 Windows Server 2022 打包及 EXE 自检；Windows 10 与 Office/WPS 人工排版验收仍待完成。

## 已实现

- 多路径递归扫描，格式/大小/修改时间筛选，目录排除、去重、暂停/继续/取消。
- DOCX、XLSX、PPTX 和 TXT/MD/LOG/CSV/TSV 的关键词替换。
- 多组规则、规则 JSON 保存/读取、双击编辑、可选忽略大小写、较长词优先、不连锁替换。
- 扫描结果逐项勾选、上下文查看、调用默认程序打开原文件。
- 另存副本、备份后修改、直接修改原文件三种模式；保存前检查原文件内容变化。
- 可选离线 OCR：默认关闭，文档内图片和独立图片分开选择。
- 图片整行选区重绘、字体/颜色设置、整图替换、前后预览后确认。
- 扫描及执行报告 JSON 导出，保留失败原因和备份位置。
- DOC/XLS/PPT 的 Windows Office/WPS COM 转换桥（尚未实机验证）。

WPS/ET/DPS 专用格式当前显示未支持原因，不修改原文件。PDF、宏文档等不在当前处理范围。
完整边界见 [兼容范围](docs/compatibility.md)。

## 开发运行

Python 3.9+，需要带 Tk 的 Python。核心功能只使用标准库。

```sh
python main.py
python -m unittest discover -s tests -v
```

可选 OCR：

```sh
python -m pip install -r requirements-ocr.txt
```

Windows 旧版文档转换还需要 `pywin32` 和可用 Office/WPS 自动化接口。
办公程序只打开临时副本；不能禁用自动宏的接口拒绝打开。传统 Excel 4.0 宏等特殊文件未纳入验收。

## 在 Windows 上生成 EXE

打包电脑安装 Python 3.11 x64，并包含 Tcl/Tk 和 pip。最终使用者无需安装 Python。
在项目目录运行 PowerShell：

```powershell
# 核心轻量版
powershell -ExecutionPolicy Bypass -File scripts/build_windows.ps1

# 含离线 OCR 模型版本
powershell -ExecutionPolicy Bypass -File scripts/build_windows.ps1 -WithOCR
```

脚本创建独立构建虚拟环境、安装依赖、运行测试和窗口冒烟验证，然后生成：

- `release/DocReplace-win-x64.zip`，解压后运行 `DocReplace/DocReplace.exe`。
- `release/DocReplace-OCR-win-x64.zip`，解压后运行 `DocReplace-OCR/DocReplace-OCR.exe`。

不要只复制 EXE，配套 `_internal` 文件夹也必须保留。字体使用目标电脑的微软雅黑等系统字体，不随包分发。
仓库同时提供手动触发的 GitHub Actions Windows 构建配置；尚未上传或触发运行。
PyInstaller 需要在 Windows 上构建 Windows EXE，参见[官方说明](https://pyinstaller.org/en/stable/)。

## 使用步骤

1. 扫描范围：添加磁盘根目录/文件夹/文件，设置筛选及排除项。
2. 替换规则：添加原关键词和替换文字，或读取示例规则 JSON。替换文字可留空用于删除。
3. 图片识别默认关闭。使用 OCR 包时可以开启并勾选文档内图片/独立图片。
4. 点击开始扫描，预览结果。双击文字命中切换勾选；图片需进入图片编辑窗口预览确认。
5. 选择三种保存方式之一。副本模式必须指定输出目录；同名输出不覆盖。
6. 点击执行，核对本批文件数量和保存模式。完成后查看并导出任务报告。

“直接修改”不备份；“备份后修改”的备份位于原文件目录下 `.docreplace-backups/唯一批次目录/原文件名`。
保存使用临时文件校验后替换，但仍需通过真实 Office/WPS 验收文档格式和排版。
修改原文件后再次执行同一结果会提示文件已变化，请重新扫描。

## 版本状态

源码开发版，已做 macOS 桌面扫描保存流程、中文样图 OCR 和文档往返测试；不宣称已通过 Windows/Office/WPS 实机验收。
现代文档核心测试用于验证结构和内容，不等价于主流办公软件的视觉兼容认证。
