"""探测模板结构:单元格对齐、合并、行高、字体。"""
import sys
from pathlib import Path
import xlwings as xw

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

TEMPLATE = Path(r"C:\Users\admin\Desktop\代码\模板\欧代标签-爱斯海-103.xlsx")

# xlVAlign 枚举
V_ALIGN = {-4160: "xlTop", -4108: "xlCenter", -4107: "xlBottom",
           -4130: "xlJustify", -4117: "xlDistributed"}
H_ALIGN = {1: "xlGeneral", -4131: "xlLeft", -4108: "xlCenter",
           -4152: "xlRight", -4130: "xlJustify", 5: "xlFill",
           -4117: "xlDistributed", 7: "xlCenterAcrossSelection"}

with xw.App(visible=False) as app:
    app.display_alerts = False
    wb = app.books.open(str(TEMPLATE))
    try:
        sht = wb.sheets[0]
        print(f"Sheet 名称: {sht.name}")
        print(f"Sheet 总数: {len(wb.sheets)}, 名称: {[s.name for s in wb.sheets]}")
        print(f"UsedRange: {sht.used_range.address}")
        print()

        for addr in ("A1", "B1", "A2", "B2", "A3", "B3"):
            r = sht.range(addr)
            try:
                api = r.api
                print(f"=== {addr} ===")
                print(f"  值: {repr(r.value)}")
                print(f"  行高: {r.row_height}, 列宽: {r.column_width}")
                print(f"  合并: {api.MergeCells}")
                if api.MergeCells:
                    print(f"  合并区域: {api.MergeArea.Address}")
                print(f"  水平对齐: {H_ALIGN.get(api.HorizontalAlignment, api.HorizontalAlignment)}")
                print(f"  垂直对齐: {V_ALIGN.get(api.VerticalAlignment, api.VerticalAlignment)}")
                print(f"  自动换行 WrapText: {api.WrapText}")
                print(f"  缩小以适应 ShrinkToFit: {api.ShrinkToFit}")
                f = api.Font
                print(f"  字体: {f.Name} / 字号 {f.Size} / 加粗 {f.Bold}")
                print()
            except Exception as e:
                print(f"  读取出错: {e}")
    finally:
        wb.close()
