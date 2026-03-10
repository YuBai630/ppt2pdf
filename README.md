# PPT to PDF Converter

[中文版](#中文版) | [English](#english)

一个 cross-platform GUI tool for converting PowerPoint presentations to PDF with 2-up layout.

一个跨平台的 PPT 转 PDF 图形化工具，支持 2-up 布局输出。

---

## Features | 功能特点

- Cross-platform support (Windows & macOS) | 跨平台支持（Windows 和 macOS）
- Graphical user interface | 图形化界面
- Batch conversion | 批量转换
- 2-up vertical layout (2 slides per page) | 2-up 纵向布局（每页2张幻灯片）
- Auto-sync output directory | 输出目录自动同步
- File validation | 文件校验
- Error retry | 错误重试
- Auto-compress to ZIP | 自动压缩为 ZIP

---

## Quick Start | 快速开始

### Prerequisites | 前置要求

**Python 3.8+** must be installed.

必须安装 Python 3.8 或更高版本。

**Software dependencies:**

| Platform | Required Software |
|----------|------------------|
| Windows | Microsoft PowerPoint |
| macOS | LibreOffice (recommended) or Microsoft PowerPoint |

### Installation | 安装

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pypdf python-pptx pywin32  # Windows only
pip install pypdf python-pptx           # macOS
```

### Run | 运行

```bash
python ppt_to_pdf_gui.py
```

---

## macOS: Install LibreOffice | macOS: 安装 LibreOffice

```bash
brew install --cask libreoffice
```

Or download from: https://www.libreoffice.org/download/

---

## Usage | 使用方法

1. Click **选择文件** or **选择文件夹** to add PPT files | 点击 **选择文件** 或 **选择文件夹** 添加 PPT 文件
2. Output directory auto-syncs to file location (can be manually changed) | 输出目录自动同步到文件所在目录（可手动修改）
3. Click **开始转换** | 点击 **开始转换**
4. PDF and ZIP files are saved to the output directory | PDF 和 ZIP 文件保存到输出目录

---

## Output Format | 输出格式

The PDF uses **2-up vertical layout**:

输出的 PDF 采用 **2-up 纵向布局**：

```
+------------------+
|    Slide 1       |  ← Top / 上方
+------------------+
+------------------+
|    Slide 2       |  ← Bottom / 下方
+------------------+
        ↓
      Next page
      下一页
```

---

## File Structure | 文件结构

```
PPT2PDF_GUI/
├── ppt_to_pdf_gui.py    # Main program / 主程序
├── requirements.txt     # Dependencies / 依赖列表
└── README.md            # Documentation / 说明文档
```

---

## FAQ | 常见问题

**Q: "pywin32 is required" error? | 提示"需要安装 pywin32"？**

A: Windows users run:
Windows 用户运行：
```bash
pip install pywin32
```

**Q: Conversion fails on macOS? | macOS 转换失败？**

A: Install LibreOffice:
请安装 LibreOffice：
```bash
brew install --cask libreoffice
```

**Q: File shows "Error"? | 文件显示"错误"？**

A: The PPT file may be corrupted. Open it in PowerPoint to check and repair.
该 PPT 文件可能已损坏，请用 PowerPoint 打开检查并修复。

---

## License | 许可证

MIT License

---

## Author | 作者

Generated with Claude Code

