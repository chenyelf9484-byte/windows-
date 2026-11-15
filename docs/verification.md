# 验证记录 · 2026-11-15

开发环境：macOS，本机 Python 3.9 与工具提供的 Python 3.12。

## 已执行

- 21 项 unittest 测试通过，包括规则、XML 局部修改、目录扫描、三种保存、内容变化拒绝覆盖及异常输入。
- 使用 python-docx、openpyxl、python-pptx 创建真实文档，修改后分别重新打开验证内容、样式、公式及结构。
- 使用 rapidocr-onnxruntime 1.4.4 对本机字体生成的中文样图识别，验证命中、新文字重绘与图片可读。
- 桌面冒烟测试：创建实际 Tk 窗口、添加规则、扫描临时 TXT、通过后台队列获取结果、直接保存并核对文件内容，自动关闭。
- Python 源码语法编译通过。

OCR 测试依赖只安装在临时目录，不修改系统 Python。Windows 构建会使用独立虚拟环境。

## 未执行

- Windows EXE 构建及 Windows 10 运行；当前环境没有 Windows 主机或 Wine。
- 实际 Microsoft Office/WPS 打开、编辑、保存及视觉排版验收。
- DOC/XLS/PPT COM 转换桥的 Windows 实测。
- WPS/ET/DPS 原格式写回（当前拒绝修改）。
- OCR 的真实业务图片准确率、Windows OCR 分发包离线运行。

因此本次交付属于源码开发版，不是已通过 Windows 验收的正式 EXE。
