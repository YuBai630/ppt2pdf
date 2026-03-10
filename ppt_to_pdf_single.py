# -*- coding: utf-8 -*-
"""
PPT转PDF - 单文件处理版
一次只处理一个PPTX文件，避免PowerPoint COM不稳定问题
"""

import os
import sys
import time
import tempfile
import shutil
from pathlib import Path

import win32com.client
from pypdf import PdfReader, PdfWriter, Transformation


def convert_single_pptx(pptx_path, output_dir):
    """转换单个PPTX文件为2-up PDF"""
    pptx_path = Path(pptx_path).absolute()
    output_dir = Path(output_dir).absolute()
    output_dir.mkdir(parents=True, exist_ok=True)

    temp_dir = tempfile.mkdtemp(prefix="ppt_single_")
    temp_pdf = Path(temp_dir) / f"{pptx_path.stem}.pdf"
    final_pdf = output_dir / f"{pptx_path.stem}.pdf"

    print(f"Input:  {pptx_path}")
    print(f"Output: {final_pdf}")

    # Step 1: PPT -> PDF (使用独立PowerPoint实例)
    print("\nStep 1: Converting PPT to PDF...")
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")

    try:
        deck = powerpoint.Presentations.Open(str(pptx_path), ReadOnly=True, WithWindow=False)
        deck.SaveAs(str(temp_pdf), 32)  # ppSaveAsPDF = 32
        deck.Close()
        print("  Done!")
    finally:
        powerpoint.Quit()
        time.sleep(1)  # 等待PowerPoint完全退出

    # Step 2: Merge to 2-up
    print("\nStep 2: Creating 2-up layout...")
    reader = PdfReader(temp_pdf)
    writer = PdfWriter()

    total_pages = len(reader.pages)
    print(f"  Total slides: {total_pages}")

    for i in range(0, total_pages, 2):
        page1 = reader.pages[i]
        page2 = reader.pages[i + 1] if i + 1 < total_pages else None

        width = float(page1.mediabox.width)
        height = float(page1.mediabox.height)
        new_page = writer.add_blank_page(width=width, height=height * 2)

        new_page.merge_transformed_page(page1, Transformation().translate(0, height))

        if page2:
            new_page.merge_transformed_page(page2, Transformation().translate(0, 0))

    with open(final_pdf, 'wb') as f:
        writer.write(f)

    print("  Done!")

    # Cleanup
    shutil.rmtree(temp_dir, ignore_errors=True)

    print(f"\n[SUCCESS] {final_pdf}")
    return str(final_pdf)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="PPT转PDF - 单文件处理版")
    parser.add_argument("input", nargs="?", help="输入PPTX文件路径")
    parser.add_argument("-o", "--output", help="输出目录")
    parser.add_argument("--failed", action="store_true", help="处理失败文件列表")
    args = parser.parse_args()

    base_dir = Path(r"C:\Users\yqccc\Desktop\临时文件夹\数据结构PPT修改\2026数据结构")
    output_folder = args.output or str(base_dir / "PDF输出")

    # 失败文件列表（已修正文件名）
    failed_files = [
        r"7. Priority queues\7.03.d-ary_heaps-d叉堆.pptx",
        r"7. Priority queues\7.04.Leftist_heaps-左倾堆.pptx",
        r"9. Hash functions and hash tables\9.01.Hash_table_introduction-哈希表简介.pptx",
        r"9. Hash functions and hash tables\9.06.Open_addressing-开放定址法.pptx",
        r"9. Hash functions and hash tables\9.08.Quadratic_probing-二次探测.pptx",
        r"9. Hash functions and hash tables\9.09.Double_hashing-双重哈希.pptx",
    ]

    print("=" * 60)
    print("PPT to PDF - Single File Mode")
    print("=" * 60)

    if args.failed:
        # 批量处理失败文件
        print(f"\nProcessing {len(failed_files)} failed files...\n")
        success = 0
        failed = 0
        for rel_path in failed_files:
            pptx_file = base_dir / rel_path
            print(f"\n{'='*60}")
            print(f"[{success + failed + 1}/{len(failed_files)}] {rel_path}")
            print("=" * 60)
            if pptx_file.exists():
                try:
                    convert_single_pptx(pptx_file, output_folder)
                    success += 1
                except Exception as e:
                    print(f"[ERROR] {e}")
                    failed += 1
            else:
                print(f"[ERROR] File not found: {pptx_file}")
                failed += 1
        print(f"\n{'='*60}")
        print(f"Batch Complete: {success} success, {failed} failed")
        print("=" * 60)
    elif args.input:
        pptx_file = Path(args.input)
        if not pptx_file.exists():
            print(f"[ERROR] File not found: {pptx_file}")
            sys.exit(1)
        convert_single_pptx(pptx_file, output_folder)
    else:
        # 默认：处理第一个失败文件
        pptx_file = base_dir / failed_files[0]
        if not pptx_file.exists():
            print(f"[ERROR] File not found: {pptx_file}")
            sys.exit(1)
        convert_single_pptx(pptx_file, output_folder)
