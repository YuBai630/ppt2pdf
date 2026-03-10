# -*- coding: utf-8 -*-
"""
PPT转PDF工具 - 2-up纵向布局（上下排列）
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path

# Windows COM support
import win32com.client

# PDF manipulation
from pypdf import PdfReader, PdfWriter, Transformation


def ppt_to_pdf_slides(ppt_path, pdf_path, powerpoint):
    """将PPT转换为单页PDF（每张幻灯片一页）"""
    ppt_path = os.path.abspath(str(ppt_path))
    pdf_path = os.path.abspath(str(pdf_path))

    deck = powerpoint.Presentations.Open(ppt_path, ReadOnly=True, WithWindow=False)
    try:
        deck.SaveAs(pdf_path, 32)  # ppSaveAsPDF = 32
    finally:
        deck.Close()


def merge_pdf_2up_vertical(input_pdf, output_pdf):
    """将PDF每2页合并为1页（上下排列）"""
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    total_pages = len(reader.pages)

    for i in range(0, total_pages, 2):
        page1 = reader.pages[i]
        page2 = reader.pages[i + 1] if i + 1 < total_pages else None

        # 获取页面尺寸
        width = float(page1.mediabox.width)
        height = float(page1.mediabox.height)

        # 创建新页面（宽度不变，高度x2）
        new_height = height * 2
        new_page = writer.add_blank_page(width=width, height=new_height)

        # 放置第1页（上方）- 需要向上平移height
        new_page.merge_transformed_page(
            page1,
            Transformation().translate(0, height)
        )

        # 放置第2页（下方）- 位置不变
        if page2:
            new_page.merge_transformed_page(
                page2,
                Transformation().translate(0, 0)
            )

    with open(output_pdf, 'wb') as f:
        writer.write(f)


def convert_all_ppts(input_dir, output_dir):
    """转换所有PPT为2-up PDF"""
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    # 创建输出目录
    output_dir.mkdir(parents=True, exist_ok=True)

    # 创建临时目录
    temp_dir = tempfile.mkdtemp(prefix="ppt_pdf_")
    print(f"Temp dir: {temp_dir}")

    # 启动PowerPoint
    print("Starting PowerPoint...")
    powerpoint = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")

    success_count = 0
    error_count = 0

    try:
        # 获取所有子文件夹并排序
        folders = sorted([f for f in input_dir.iterdir() if f.is_dir() and not f.name.startswith('.')])

        for folder in folders:
            print(f"\n{'='*60}")
            print(f"Folder: {folder.name}")
            print('='*60)

            # 获取该文件夹下所有PPT文件
            ppt_files = sorted([f for f in folder.glob("*")
                               if f.suffix.lower() in ['.ppt', '.pptx']])

            for ppt_file in ppt_files:
                try:
                    print(f"\n  Converting: {ppt_file.name}")

                    # 第1步：PPT转单页PDF
                    temp_pdf = Path(temp_dir) / f"{ppt_file.stem}.pdf"
                    ppt_to_pdf_slides(ppt_file, temp_pdf, powerpoint)

                    # 第2步：合并为2-up PDF
                    final_pdf = output_dir / f"{ppt_file.stem}.pdf"
                    merge_pdf_2up_vertical(temp_pdf, final_pdf)

                    print(f"  [OK] Done: {final_pdf.name}")
                    success_count += 1

                    # 清理临时PDF
                    temp_pdf.unlink(missing_ok=True)

                except Exception as e:
                    print(f"  [ERR] Error [{ppt_file.name}]: {e}")
                    error_count += 1

    finally:
        print("\nClosing PowerPoint...")
        try:
            powerpoint.Quit()
        except:
            pass

        # 清理临时目录
        print("Cleaning temp files...")
        shutil.rmtree(temp_dir, ignore_errors=True)

    print(f"\n{'='*60}")
    print(f"Conversion Complete!")
    print(f"  Success: {success_count}")
    print(f"  Failed: {error_count}")
    print(f"  Output: {output_dir}")
    print('='*60)


if __name__ == "__main__":
    input_directory = r"C:\Users\yqccc\Desktop\临时文件夹\数据结构PPT修改\2026数据结构"
    output_directory = r"C:\Users\yqccc\Desktop\临时文件夹\数据结构PPT修改\2026数据结构\PDF输出"

    print("PPT to PDF Tool - 2-up Vertical Layout")
    print(f"Input: {input_directory}")
    print(f"Output: {output_directory}")
    print()

    convert_all_ppts(input_directory, output_directory)
