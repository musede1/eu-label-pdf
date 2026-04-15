"""CLI 入口:手工跑一次生成。

    python generate.py <店铺简称> <product_name> <model> <batch_number> [--output-dir 路径]
"""
from __future__ import annotations

import argparse
import sys
import uuid
from pathlib import Path

from core import fill_and_export_pdf, find_template, crop_pdf_file

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

DEFAULT_OUTPUT_DIR = Path(r"C:\Users\admin\Desktop\代码\输出")


def main() -> int:
    parser = argparse.ArgumentParser(description="欧代标签 PDF 生成")
    parser.add_argument("shop_code")
    parser.add_argument("product_name")
    parser.add_argument("model")
    parser.add_argument("batch_number")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    args = parser.parse_args()

    try:
        template = find_template(args.shop_code)
        args.output_dir.mkdir(parents=True, exist_ok=True)
        base = f"{args.shop_code}_{args.batch_number}_{uuid.uuid4().hex[:8]}"
        raw_pdf = args.output_dir / f"{base}.pdf"

        print(f"[1/3] 模板: {template}")
        fill_and_export_pdf(template, args.product_name, args.model,
                            args.batch_number, raw_pdf)
        print(f"[2/3] 原始 PDF: {raw_pdf}")
        cropped = crop_pdf_file(raw_pdf)
        raw_pdf.unlink(missing_ok=True)
        final = cropped.with_name(f"{base}.pdf")
        if final.exists():
            final.unlink()
        cropped.rename(final)
        print(f"[3/3] ✅ 输出: {final}")
    except Exception as e:
        print(f"❌ 失败: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
