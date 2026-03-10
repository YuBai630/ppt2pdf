# -*- coding: utf-8 -*-
"""
PPT转PDF校验脚本
检查：1. 文件完整性  2. 页码匹配（2-up布局）
"""

import math
from pathlib import Path
from pypdf import PdfReader
from pptx import Presentation


def verify_conversion(pptx_dir, pdf_dir):
    """校验PPT到PDF的转换结果"""
    pptx_dir = Path(pptx_dir)
    pdf_dir = Path(pdf_dir)

    # 收集所有PPTX文件
    pptx_files = []
    for folder in sorted(pptx_dir.iterdir()):
        if folder.is_dir() and not folder.name.startswith('.'):
            for f in sorted(folder.glob("*.pptx")):
                pptx_files.append(f)

    print("=" * 70)
    print("PPT to PDF Verification Report")
    print("=" * 70)
    print(f"PPTX directory: {pptx_dir}")
    print(f"PDF directory:  {pdf_dir}")
    print(f"Total PPTX files: {len(pptx_files)}")
    print("=" * 70)

    missing = []      # 缺失的PDF
    page_mismatch = []  # 页码不匹配
    success = 0

    for pptx in pptx_files:
        pdf_file = pdf_dir / f"{pptx.stem}.pdf"

        # 检查1: 文件是否存在
        if not pdf_file.exists():
            missing.append(pptx)
            continue

        # 检查2: 页码匹配
        try:
            # 获取PPT幻灯片数
            prs = Presentation(str(pptx))
            slide_count = len(prs.slides)

            # 获取PDF页数
            reader = PdfReader(str(pdf_file))
            pdf_pages = len(reader.pages)

            # 2-up布局: expected = ceil(slide_count / 2)
            # PowerPoint导出PDF时可能跳过空白/隐藏幻灯片，允许1-2页偏差
            expected_pages = math.ceil(slide_count / 2)
            min_acceptable = math.ceil((slide_count - 4) / 2)  # 允许最多4张幻灯片被跳过
            max_acceptable = expected_pages

            if min_acceptable <= pdf_pages <= max_acceptable:
                success += 1
            else:
                page_mismatch.append({
                    'file': pptx,
                    'slides': slide_count,
                    'expected': expected_pages,
                    'actual': pdf_pages
                })
        except Exception as e:
            page_mismatch.append({
                'file': pptx,
                'error': str(e)
            })

    # 输出报告
    print(f"\n{'='*70}")
    print("RESULTS")
    print(f"{'='*70}")
    print(f"  Success:        {success}")
    print(f"  Missing PDF:    {len(missing)}")
    print(f"  Page mismatch:  {len(page_mismatch)}")

    if missing:
        print(f"\n--- Missing PDF files ({len(missing)}) ---")
        for f in missing:
            print(f"  [MISSING] {f.relative_to(pptx_dir)}")

    if page_mismatch:
        print(f"\n--- Page count mismatches ({len(page_mismatch)}) ---")
        for item in page_mismatch:
            if 'error' in item:
                print(f"  [ERROR]   {item['file'].relative_to(pptx_dir)}: {item['error']}")
            else:
                print(f"  [MISMATCH] {item['file'].relative_to(pptx_dir)}")
                print(f"             Slides: {item['slides']}, Expected PDF pages: {item['expected']}, Actual: {item['actual']}")

    print(f"\n{'='*70}")
    if not missing and not page_mismatch:
        print("[ALL PASSED] All conversions verified successfully!")
    else:
        print(f"[ISSUES FOUND] Please review the errors above")
    print("=" * 70)

    return len(missing) == 0 and len(page_mismatch) == 0


if __name__ == "__main__":
    base_dir = Path(r"C:\Users\yqccc\Desktop\临时文件夹\数据结构PPT修改\2026数据结构")
    pptx_dir = base_dir
    pdf_dir = base_dir / "PDF输出"

    verify_conversion(pptx_dir, pdf_dir)
