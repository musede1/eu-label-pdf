"""对所有模板跑一遍,共享一个 Excel 实例提速。"""
import sys
import uuid
import re
from pathlib import Path

import xlwings as xw
from PyPDF2 import PdfReader, PdfWriter

from core import TEMPLATE_DIR

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
OUTPUT_DIR = Path(r"C:\Users\admin\Desktop\代码\输出\batch_test")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

CROP_LEFT_MM, CROP_TOP_MM = 35.2, 81.3
CROP_RIGHT_MM, CROP_BOTTOM_MM = 37.1, 80.2

XL_V_CENTER = -4108

# 测试数据
PRODUCT_NAME = "Test Product Name"
MODEL = "MODEL-001"
BATCH_NUMBER = "BATCH20260415"


def crop_pdf(pdf_path: Path) -> Path:
    mm_to_pt = 72 / 25.4
    left = CROP_LEFT_MM * mm_to_pt
    top = CROP_TOP_MM * mm_to_pt
    right = CROP_RIGHT_MM * mm_to_pt
    bottom = CROP_BOTTOM_MM * mm_to_pt

    output = pdf_path.with_name(pdf_path.stem + "_cropped" + pdf_path.suffix)
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    for page in reader.pages:
        box = page.mediabox
        llx = float(box.lower_left[0]) + left
        lly = float(box.lower_left[1]) + bottom
        urx = float(box.upper_right[0]) - right
        ury = float(box.upper_right[1]) - top
        page.mediabox.lower_left = (llx, lly)
        page.mediabox.upper_right = (urx, ury)
        writer.add_page(page)
    with open(output, "wb") as f:
        writer.write(f)
    return output


def process(app, template: Path, shop_code: str, shop_name: str) -> tuple[bool, str]:
    raw_pdf = OUTPUT_DIR / f"{shop_code}_{shop_name}_{uuid.uuid4().hex[:6]}.pdf"
    try:
        wb = app.books.open(str(template))
        try:
            sht = wb.sheets[0]
            heights = {addr: sht.range(addr).row_height for addr in ("B1", "B2", "B3")}

            for addr, val in (("B1", PRODUCT_NAME), ("B2", MODEL), ("B3", BATCH_NUMBER)):
                cell = sht.range(addr)
                try:
                    cell.api.VerticalAlignment = XL_V_CENTER
                    cell.api.WrapText = False
                except Exception:
                    pass
                cell.value = val

            for addr, h in heights.items():
                sht.range(addr).row_height = h

            sht.to_pdf(str(raw_pdf))
        finally:
            wb.close()
    except Exception as e:
        return False, f"导出失败: {e}"

    try:
        cropped = crop_pdf(raw_pdf)
        raw_pdf.unlink(missing_ok=True)
        final = cropped.with_name(f"{shop_code}_{shop_name}.pdf")
        if final.exists():
            final.unlink()
        cropped.rename(final)
        return True, str(final)
    except Exception as e:
        return False, f"裁剪失败: {e}"


def main():
    templates = sorted(TEMPLATE_DIR.glob("欧代标签-*.xlsx"))
    print(f"共 {len(templates)} 个模板\n")

    rows = []
    with xw.App(visible=False) as app:
        app.display_alerts = False
        for i, template in enumerate(templates, 1):
            m = re.match(r"欧代标签-(.+)-(\d+)\.xlsx", template.name)
            if not m:
                print(f"[{i}/{len(templates)}] ⚠️  跳过(名称不匹配): {template.name}")
                continue
            shop_name, shop_code = m.group(1), m.group(2)
            print(f"[{i}/{len(templates)}] {shop_code} ({shop_name}) ...", end=" ", flush=True)
            ok, msg = process(app, template, shop_code, shop_name)
            print("✅" if ok else "❌", msg if not ok else "")
            rows.append((shop_code, shop_name, ok, msg))

    print("\n=== 汇总 ===")
    success = sum(1 for r in rows if r[2])
    print(f"成功 {success}/{len(rows)},失败 {len(rows) - success}")
    for shop_code, shop_name, ok, msg in rows:
        if not ok:
            print(f"  ❌ {shop_code} {shop_name}: {msg}")
    print(f"\n输出目录: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
