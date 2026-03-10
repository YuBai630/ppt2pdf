# -*- coding: utf-8 -*-
"""
PPT转PDF图形化管理工具 - 跨平台版
支持 Windows 和 macOS
"""

import os
import sys
import time
import tempfile
import shutil
import threading
import platform
from pathlib import Path
from datetime import datetime

# GUI
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# PPT validation
try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# PDF manipulation
from pypdf import PdfReader, PdfWriter, Transformation

# Platform detection
IS_WINDOWS = platform.system() == 'Windows'
IS_MACOS = platform.system() == 'Darwin'

# Platform-specific imports
if IS_WINDOWS:
    try:
        import win32com.client
        HAS_WIN32COM = True
    except ImportError:
        HAS_WIN32COM = False
else:
    HAS_WIN32COM = False

if IS_MACOS:
    try:
        from AppKit import NSWorkspace, NSURL
        from Foundation import NSThread
        HAS_APPKIT = True
    except ImportError:
        HAS_APPKIT = False
else:
    HAS_APPKIT = False


class PPTConverter:
    """PPT转PDF转换器 - 跨平台支持"""

    def __init__(self, log_callback=None):
        self.log = log_callback or print

    def convert(self, ppt_path, output_dir=None):
        """转换PPT为PDF

        Args:
            ppt_path: PPT文件路径
            output_dir: 输出目录（可选，默认为PPT所在目录）
        """
        ppt_path = Path(ppt_path)
        if output_dir is None:
            output_dir = ppt_path.parent
        else:
            output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        if IS_WINDOWS:
            return self._convert_windows(ppt_path, output_dir)
        elif IS_MACOS:
            return self._convert_macos(ppt_path, output_dir)
        else:
            raise RuntimeError(f"不支持的操作系统: {platform.system()}")

    def _convert_windows(self, ppt_path, output_dir):
        """Windows平台转换"""
        if not HAS_WIN32COM:
            raise RuntimeError("需要安装 pywin32: pip install pywin32")

        temp_dir = tempfile.mkdtemp(prefix="ppt_pdf_")
        temp_pdf = Path(temp_dir) / f"{ppt_path.stem}.pdf"
        final_pdf = output_dir / f"{ppt_path.stem}-handout.pdf"

        try:
            # 启动PowerPoint
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")

            try:
                # 打开并转换
                deck = powerpoint.Presentations.Open(str(ppt_path.absolute()), ReadOnly=True, WithWindow=False)
                deck.SaveAs(str(temp_pdf.absolute()), 32)  # ppSaveAsPDF = 32
                deck.Close()
            finally:
                powerpoint.Quit()
                time.sleep(0.5)

            # 合并为2-up布局
            self._merge_pdf_2up(temp_pdf, final_pdf)

            self.log(f"[OK] {ppt_path.name}")
            return str(final_pdf)

        except Exception as e:
            self.log(f"[ERR] {ppt_path.name}: {e}")
            raise
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _convert_macos(self, ppt_path, output_dir):
        """macOS平台转换 - 使用LibreOffice或手动转换"""
        final_pdf = output_dir / f"{ppt_path.stem}-handout.pdf"

        # 方法1: 尝试使用LibreOffice (更可靠)
        try:
            result = self._convert_with_libreoffice(ppt_path, output_dir)
            if result:
                # 合并为2-up布局
                temp_pdf = output_dir / f"{ppt_path.stem}.pdf"
                self._merge_pdf_2up(temp_pdf, final_pdf)
                self.log(f"[OK] {ppt_path.name}")
                return str(final_pdf)
        except Exception as e:
            self.log(f"LibreOffice转换失败: {e}")

        # 方法2: 使用AppleScript调用PowerPoint/Keynote
        try:
            result = self._convert_with_applescript(ppt_path, output_dir)
            if result:
                temp_pdf = output_dir / f"{ppt_path.stem}.pdf"
                self._merge_pdf_2up(temp_pdf, final_pdf)
                self.log(f"[OK] {ppt_path.name}")
                return str(final_pdf)
        except Exception as e:
            self.log(f"AppleScript转换失败: {e}")

        raise RuntimeError("无法在macOS上转换PPT，请安装LibreOffice或Microsoft PowerPoint")

    def _convert_with_libreoffice(self, ppt_path, output_dir):
        """使用LibreOffice转换"""
        import subprocess

        # 常见LibreOffice路径
        libreoffice_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/usr/bin/libreoffice",
            "/usr/local/bin/libreoffice",
        ]

        soffice = None
        for path in libreoffice_paths:
            if os.path.exists(path):
                soffice = path
                break

        if not soffice:
            raise RuntimeError("未找到LibreOffice，请从 https://www.libreoffice.org 下载安装")

        # 执行转换
        cmd = [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(ppt_path)
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice转换失败: {result.stderr}")

        return True

    def _convert_with_applescript(self, ppt_path, output_dir):
        """使用AppleScript调用PowerPoint转换"""
        import subprocess
        import shlex

        # 使用 shlex.quote 正确处理路径中的特殊字符
        ppt_path_str = shlex.quote(str(ppt_path.absolute()))
        output_path_str = shlex.quote(str(output_dir / f"{ppt_path.stem}.pdf"))

        applescript = f'''
        tell application "Microsoft PowerPoint"
            open POSIX file {ppt_path_str}
            set thePresentation to active presentation
            save thePresentation in POSIX file {output_path_str} as save as PDF
            close thePresentation
        end tell
        '''

        result = subprocess.run(
            ["osascript", "-e", applescript],
            capture_output=True, text=True, timeout=60
        )

        if result.returncode != 0:
            raise RuntimeError(f"AppleScript执行失败: {result.stderr}")

        return True

    def _merge_pdf_2up(self, input_pdf, output_pdf, add_border=True, scale=0.9):
        """将PDF每2页合并为1页（上下排列），可选添加边框和缩放

        Args:
            input_pdf: 输入PDF路径
            output_pdf: 输出PDF路径
            add_border: 是否添加边框
            scale: 缩放比例，默认0.9（90%），使幻灯片周围有10%留白
        """
        reader = PdfReader(input_pdf)
        writer = PdfWriter()

        total_pages = len(reader.pages)

        for i in range(0, total_pages, 2):
            page1 = reader.pages[i]
            page2 = reader.pages[i + 1] if i + 1 < total_pages else None

            width = float(page1.mediabox.width)
            height = float(page1.mediabox.height)
            new_page = writer.add_blank_page(width=width, height=height * 2)

            # 计算缩放后的尺寸和居中偏移
            scaled_width = width * scale
            scaled_height = height * scale
            offset_x = (width - scaled_width) / 2
            offset_y = (height - scaled_height) / 2

            # 合并页面（缩放并居中）
            # 上方幻灯片
            new_page.merge_transformed_page(
                page1,
                Transformation().scale(scale).translate(offset_x, height + offset_y)
            )
            # 下方幻灯片
            if page2:
                new_page.merge_transformed_page(
                    page2,
                    Transformation().scale(scale).translate(offset_x, offset_y)
                )

            # 添加边框（匹配缩放后的尺寸）
            if add_border:
                border_page = self._create_border_page(scaled_width, scaled_height)
                # 上方边框
                new_page.merge_transformed_page(
                    border_page,
                    Transformation().translate(offset_x, height + offset_y)
                )
                # 下方边框
                if page2:
                    new_page.merge_transformed_page(
                        border_page,
                        Transformation().translate(offset_x, offset_y)
                    )

        with open(output_pdf, 'wb') as f:
            writer.write(f)

    def _create_border_page(self, width, height, border_width=2):
        """创建带边框的透明页面"""
        from reportlab.pdfgen import canvas
        from reportlab.lib.colors import black
        import io

        # 使用 reportlab 创建边框
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))
        c.setStrokeColor(black)
        c.setLineWidth(border_width)
        c.rect(0, 0, width, height)
        c.save()
        packet.seek(0)

        return PdfReader(packet).pages[0]

    def compress_to_zip(self, pdf_files, zip_path):
        """将PDF文件压缩到zip"""
        import zipfile

        zip_path = Path(zip_path)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for pdf_file in pdf_files:
                pdf_file = Path(pdf_file)
                if pdf_file.exists():
                    zf.write(pdf_file, pdf_file.name)
                    self.log(f"  添加: {pdf_file.name}")

        self.log(f"[ZIP] 创建: {zip_path.name}")
        return str(zip_path)


class PPTConverterGUI:
    """PPT转PDF图形化管理工具主界面"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PPT to PDF 转换工具")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        # 文件列表 {filepath: {'status': str, 'progress': int}}
        self.files = {}
        self.output_dir = tk.StringVar()
        self.converting = False
        self.cancel_flag = False

        # 转换器
        self.converter = PPTConverter(log_callback=self.log)

        self._setup_ui()
        self._check_platform()

    def _check_platform(self):
        """检查平台兼容性"""
        if not IS_WINDOWS and not IS_MACOS:
            self.log(f"警告: 当前操作系统 {platform.system()} 可能不完全支持")

        if IS_WINDOWS and not HAS_WIN32COM:
            self.log("警告: 需要安装 pywin32: pip install pywin32")

        if IS_MACOS:
            self.log("macOS模式: 需要安装 LibreOffice 或 Microsoft PowerPoint")

    def _setup_ui(self):
        """初始化界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(main_frame, text="PPT to PDF 转换工具",
                                font=("Microsoft YaHei", 16, "bold"))
        title_label.pack(pady=(0, 10))

        # 平台信息
        platform_text = f"平台: {platform.system()} ({platform.machine()})"
        platform_label = ttk.Label(main_frame, text=platform_text,
                                   font=("Microsoft YaHei", 9), foreground="gray")
        platform_label.pack(pady=(0, 5))

        # 拖拽区域
        self._setup_drop_zone(main_frame)

        # 输出目录
        self._setup_output_section(main_frame)

        # 文件列表
        self._setup_file_list(main_frame)

        # 进度条
        self._setup_progress_section(main_frame)

        # 控制按钮
        self._setup_control_buttons(main_frame)

        # 日志区域
        self._setup_log_section(main_frame)

    def _setup_drop_zone(self, parent):
        """设置拖拽区域"""
        drop_frame = ttk.LabelFrame(parent, text="添加文件", padding="10")
        drop_frame.pack(fill=tk.X, pady=(0, 10))

        # 拖拽提示
        self.drop_label = ttk.Label(drop_frame,
                                    text="拖拽PPT文件或文件夹到这里\n或点击下方按钮选择",
                                    font=("Microsoft YaHei", 10),
                                    anchor=tk.CENTER)
        self.drop_label.pack(fill=tk.X, pady=5)

        # 按钮区域
        btn_frame = ttk.Frame(drop_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        ttk.Button(btn_frame, text="选择文件", command=self.select_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="选择文件夹", command=self.select_folder).pack(side=tk.LEFT, padx=5)

    def _setup_output_section(self, parent):
        """设置输出目录区域"""
        output_frame = ttk.Frame(parent)
        output_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(output_frame, text="输出目录:").pack(side=tk.LEFT)
        ttk.Entry(output_frame, textvariable=self.output_dir, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="选择", command=self.select_output).pack(side=tk.LEFT)

    def _setup_file_list(self, parent):
        """设置文件列表区域"""
        list_frame = ttk.LabelFrame(parent, text="文件列表", padding="5")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Treeview
        columns = ("filename", "status", "progress")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)

        self.tree.heading("filename", text="文件名")
        self.tree.heading("status", text="状态")
        self.tree.heading("progress", text="进度")

        self.tree.column("filename", width=400)
        self.tree.column("status", width=100)
        self.tree.column("progress", width=80)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 列表控制按钮
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Button(btn_frame, text="全选", command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="全不选", command=self.deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除选中", command=self.remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重试错误", command=self.retry_errors).pack(side=tk.LEFT, padx=5)

    def _setup_progress_section(self, parent):
        """设置进度条区域"""
        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(progress_frame, text="总进度:").pack(side=tk.LEFT)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                            maximum=100, length=500, mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.progress_label = ttk.Label(progress_frame, text="0%")
        self.progress_label.pack(side=tk.LEFT)

    def _setup_control_buttons(self, parent):
        """设置控制按钮"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_btn = ttk.Button(btn_frame, text="开始转换", command=self.start_conversion)
        self.start_btn.pack(side=tk.LEFT, padx=10)

        self.cancel_btn = ttk.Button(btn_frame, text="取消", command=self.cancel_conversion, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=10)

        ttk.Button(btn_frame, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=10)

    def _setup_log_section(self, parent):
        """设置日志区域"""
        log_frame = ttk.LabelFrame(parent, text="转换日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # ==================== 文件操作 ====================

    def select_files(self):
        """选择文件"""
        files = filedialog.askopenfilenames(
            title="选择PPT文件",
            filetypes=[("PowerPoint文件", "*.ppt;*.pptx"), ("所有文件", "*.*")]
        )
        if files:
            self.add_files(files)

    def select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory(title="选择包含PPT文件的文件夹")
        if folder:
            self.add_folder(folder)

    def select_output(self):
        """选择输出目录"""
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.output_dir.set(folder)

    def add_files(self, filepaths):
        """添加文件到列表"""
        first_dir = None  # 记录第一个文件的目录
        for filepath in filepaths:
            filepath = Path(filepath)
            if filepath.suffix.lower() in ['.ppt', '.pptx']:
                if str(filepath) not in self.files:
                    # 记录第一个有效文件的目录
                    if first_dir is None:
                        first_dir = str(filepath.parent)
                    # 添加到列表
                    self.files[str(filepath)] = {
                        'status': '待校验',
                        'progress': 0,
                        'selected': True
                    }
                    self.tree.insert("", tk.END, iid=str(filepath),
                                    values=(filepath.name, '待校验', '0%'))

        # 自动同步输出目录为第一个文件所在目录
        if first_dir:
            self.output_dir.set(first_dir)

        # 自动校验新添加的文件
        self._validate_pending_files()

    def add_folder(self, folder):
        """添加文件夹中的PPT文件"""
        folder = Path(folder)
        ppt_files = list(folder.glob("*.ppt")) + list(folder.glob("*.pptx"))
        ppt_files = [str(f) for f in ppt_files]
        if ppt_files:
            self.add_files(ppt_files)
            # 输出目录已在 add_files 中自动设置
        else:
            messagebox.showinfo("提示", "该文件夹中没有PPT文件")

    def select_all(self):
        """全选"""
        for item in self.tree.get_children():
            self.tree.selection_add(item)

    def deselect_all(self):
        """全不选"""
        for item in self.tree.selection():
            self.tree.selection_remove(item)

    def remove_selected(self):
        """删除选中项"""
        selected = self.tree.selection()
        for item in selected:
            if item in self.files:
                del self.files[item]
            self.tree.delete(item)

    def clear_list(self):
        """清空列表"""
        if self.converting:
            messagebox.showwarning("警告", "转换进行中，无法清空列表")
            return

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.files.clear()
        self.log("列表已清空")

    # ==================== 文件校验 ====================

    def _validate_pending_files(self):
        """校验待校验的文件"""
        if not HAS_PPTX:
            self.log("警告: python-pptx未安装，跳过文件校验")
            for filepath in self.files:
                if self.files[filepath]['status'] == '待校验':
                    self.files[filepath]['status'] = '就绪'
                    self._update_tree_item(filepath, '就绪', '0%')
            return

        for filepath in list(self.files.keys()):
            if self.files[filepath]['status'] == '待校验':
                valid, error = self.validate_file(filepath)
                if valid:
                    self.files[filepath]['status'] = '就绪'
                    self._update_tree_item(filepath, '就绪', '0%')
                else:
                    self.files[filepath]['status'] = f'错误: {error}'
                    self._update_tree_item(filepath, f'错误: {error}', '-')

    def validate_file(self, filepath):
        """校验单个PPTX文件是否有效"""
        try:
            prs = Presentation(filepath)
            # 检查是否有幻灯片
            if len(prs.slides) == 0:
                return False, "无幻灯片"
            return True, None
        except Exception as e:
            return False, str(e)[:30]

    def _update_tree_item(self, filepath, status, progress):
        """更新树形列表项"""
        try:
            self.tree.item(filepath, values=(Path(filepath).name, status, progress))
        except:
            pass

    # ==================== 转换功能 ====================

    def start_conversion(self):
        """开始转换"""
        if self.converting:
            return

        if not self.files:
            messagebox.showwarning("警告", "请先添加PPT文件")
            return

        self.converting = True
        self.cancel_flag = False
        self.start_btn.configure(state=tk.DISABLED)
        self.cancel_btn.configure(state=tk.NORMAL)

        # 在新线程中执行转换
        thread = threading.Thread(target=self._conversion_thread)
        thread.daemon = True
        thread.start()

    def cancel_conversion(self):
        """取消转换"""
        self.cancel_flag = True
        self.log("正在取消转换...")

    def _conversion_thread(self):
        """转换线程"""
        try:
            self._do_conversion()
        except Exception as e:
            self.log(f"转换错误: {e}")
        finally:
            self.root.after(0, self._conversion_done)

    def _do_conversion(self):
        """执行转换"""
        total_files = len(self.files)
        completed = 0
        success_count = 0
        error_count = 0

        # 获取输出目录
        output_dir = self.output_dir.get()
        self.log(f"开始转换 {total_files} 个文件...")
        self.log(f"输出目录: {output_dir}")

        # 用于存储所有成功转换的PDF（用于压缩）
        success_pdfs = []

        for filepath in list(self.files.keys()):
            if self.cancel_flag:
                self.log("转换已取消")
                break

            file_info = self.files[filepath]
            status = file_info['status']

            if status == '成功':
                completed += 1
                continue

            self.root.after(0, lambda p=filepath: self._update_tree_item(p, '处理中', '...'))

            try:
                # 使用用户指定的输出目录
                result = self.converter.convert(filepath, output_dir)

                if result:
                    success_count += 1
                    self.files[filepath]['status'] = '成功'
                    self.root.after(0, lambda p=filepath: self._update_tree_item(p, '成功', '100%'))
                    success_pdfs.append(result)

            except Exception as e:
                error_count += 1
                self.files[filepath]['status'] = '失败'
                self.root.after(0, lambda p=filepath: self._update_tree_item(p, '失败', '-'))
                self.log(f"[ERR] {Path(filepath).name}: {e}")

            completed += 1
            progress = (completed / total_files) * 100
            self.root.after(0, lambda v=progress: self._update_progress(v))

        # 转换完成后压缩（ZIP保存在输出目录）
        if success_pdfs and not self.cancel_flag:
            self.log("正在压缩...")
            output_path = Path(output_dir)
            dir_name = output_path.name
            zip_path = output_path / f"{dir_name}.zip"
            self.converter.compress_to_zip(success_pdfs, zip_path)

        self.log(f"转换完成: 成功 {success_count}, 失败 {error_count}")

    def _update_progress(self, value):
        """更新进度条"""
        self.progress_var.set(value)
        self.progress_label.configure(text=f"{value:.1f}%")

    def _conversion_done(self):
        """转换完成"""
        self.converting = False
        self.start_btn.configure(state=tk.NORMAL)
        self.cancel_btn.configure(state=tk.DISABLED)

    def retry_errors(self):
        """重试错误文件"""
        if self.converting:
            messagebox.showwarning("警告", "转换进行中，请等待完成")
            return

        error_files = [f for f in self.files if self.files[f]['status'].startswith('错误') or self.files[f]['status'] == '失败']

        if not error_files:
            messagebox.showinfo("提示", "没有需要重试的错误文件")
            return

        self.log(f"重试 {len(error_files)} 个错误文件...")

        for filepath in error_files:
            self.files[filepath]['status'] = '就绪'
            self._update_tree_item(filepath, '就绪', '0%')

        # 重新校验
        self._validate_pending_files()

    # ==================== 日志功能 ====================

    def log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"

        def append_log():
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END)

        try:
            self.root.after(0, append_log)
        except:
            print(log_entry, end='')

    # ==================== 主程序 ====================

    def run(self):
        """运行主程序"""
        # 设置默认输出目录
        default_output = Path.home() / "Desktop" / "PDF输出"
        self.output_dir.set(str(default_output))

        self.log("PPT to PDF 转换工具已启动")
        self.log(f"运行平台: {platform.system()} {platform.release()}")

        if not HAS_PPTX:
            self.log("提示: 安装 python-pptx 可启用文件校验功能 (pip install python-pptx)")

        self.root.mainloop()


def main():
    """主程序入口"""
    app = PPTConverterGUI()
    app.run()


if __name__ == "__main__":
    main()
