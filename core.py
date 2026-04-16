"""核心逻辑:填模板 + 导 PDF + 裁剪。CLI 和 HTTP 服务共用。"""
from __future__ import annotations

import atexit
import io
import logging
import threading
from pathlib import Path
from typing import Optional

import xlwings as xw
from PyPDF2 import PdfReader, PdfWriter

log = logging.getLogger(__name__)

import os
_DEFAULT_TEMPLATE_DIR = Path(__file__).resolve().parent / "templates"
TEMPLATE_DIR = Path(os.environ.get("TEMPLATE_DIR", str(_DEFAULT_TEMPLATE_DIR)))

# 裁剪参数(mm)
CROP_LEFT_MM = 35.2
CROP_TOP_MM = 81.3
CROP_RIGHT_MM = 37.1
CROP_BOTTOM_MM = 80.2

XL_V_CENTER = -4108  # xlVAlign.xlCenter

# 统一字体/字号,避免模板间差异(如 PMingLiU-ExtB 渲染英文不全)
UNIFIED_FONT_NAME = "宋体"
UNIFIED_FONT_SIZE = 11

# 常驻 WPS/Excel 实例,避免每次请求冷启动(省 1-3 秒)
_APP: Optional[xw.App] = None
_APP_LOCK = threading.Lock()
# 每处理 N 个文档重建一次实例,防 COM 泄漏/僵死
_MAX_USES = 200
_USE_COUNT = 0


def _get_app() -> xw.App:
    """取常驻 Excel 实例,失效则重建。调用方需持 _APP_LOCK。"""
    global _APP, _USE_COUNT
    if _APP is not None:
        try:
            # 试探活性:访问一个属性如果 COM 已死会抛异常
            _APP.api.Version
            if _USE_COUNT < _MAX_USES:
                return _APP
            log.info(f"达到最大使用次数 {_MAX_USES},重建 Excel 实例")
        except Exception as e:
            log.warning(f"Excel 实例已失效 ({e}),重建")
        _close_app()

    log.info("启动新 Excel/WPS 实例")
    _APP = xw.App(visible=False, add_book=False)
    _APP.display_alerts = False
    _APP.screen_updating = False
    _USE_COUNT = 0
    return _APP


def _close_app() -> None:
    global _APP, _USE_COUNT
    if _APP is not None:
        try:
            _APP.quit()
        except Exception:
            pass
        _APP = None
        _USE_COUNT = 0


@atexit.register
def _cleanup_on_exit() -> None:
    _close_app()


def find_template(shop_code: str, template_dir: Optional[Path] = None) -> Path:
    """按店铺简称找模板文件(欧代标签-*-{shop_code}.xlsx)。"""
    d = template_dir or TEMPLATE_DIR
    pattern = f"欧代标签-*-{shop_code}.xlsx"
    matches = list(d.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"未找到模板: {d}\\{pattern}")
    if len(matches) > 1:
        raise RuntimeError(
            f"店铺简称 {shop_code} 匹配到多个模板: {[p.name for p in matches]}"
        )
    return matches[0]


def fill_and_export_pdf(
    template: Path,
    product_name: str,
    model: str,
    batch_number: str,
    pdf_path: Path,
) -> None:
    """打开模板、填 B1/B2/B3、只导出第一张 sheet 的 PDF。

    复用常驻 Excel/WPS 实例(_APP),请求之间不重新启动 Excel。
    _APP_LOCK 保证同一时刻只有一个请求在操作 COM 对象。
    """
    global _USE_COUNT
    with _APP_LOCK:
        app = _get_app()
        wb = app.books.open(str(template))
        try:
            sht = wb.sheets[0]
            heights = {addr: sht.range(addr).row_height
                       for addr in ("B1", "B2", "B3")}

            # B1 允许自动换行(长产品名分两行显示,模板预留 27pt 行高)
            # B2/B3 关换行(单行内容,避免意外换行)
            wrap_config = (("B1", product_name, True),
                           ("B2", model, False),
                           ("B3", batch_number, False))
            for addr, val, wrap in wrap_config:
                cell = sht.range(addr)
                # 先写值(写值会重置一些格式)
                cell.value = val
                # 再统一设置:对齐、换行、字体。分开 try,方便排错
                try:
                    cell.api.VerticalAlignment = XL_V_CENTER
                except Exception as e:
                    log.warning(f"{addr} set VerticalAlignment failed: {e}")
                try:
                    cell.api.WrapText = wrap
                except Exception as e:
                    log.warning(f"{addr} set WrapText failed: {e}")
                try:
                    cell.api.Font.Name = UNIFIED_FONT_NAME
                    cell.api.Font.Size = UNIFIED_FONT_SIZE
                except Exception as e:
                    log.warning(f"{addr} set Font failed: {e}")

            for addr, h in heights.items():
                sht.range(addr).row_height = h

            sht.to_pdf(str(pdf_path))
        finally:
            try:
                wb.close()
            except Exception as e:
                log.warning(f"关闭工作簿失败,实例可能已损坏: {e}")
                _close_app()
                raise
        _USE_COUNT += 1


def crop_pdf_bytes(pdf_path: Path) -> bytes:
    """按固定四边裁剪,直接返回裁剪后 PDF 的字节流。"""
    mm_to_pt = 72 / 25.4
    left = CROP_LEFT_MM * mm_to_pt
    top = CROP_TOP_MM * mm_to_pt
    right = CROP_RIGHT_MM * mm_to_pt
    bottom = CROP_BOTTOM_MM * mm_to_pt

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

    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def crop_pdf_file(pdf_path: Path) -> Path:
    """按固定四边裁剪,写入 *_cropped.pdf 返回路径。"""
    data = crop_pdf_bytes(pdf_path)
    output = pdf_path.with_name(pdf_path.stem + "_cropped" + pdf_path.suffix)
    output.write_bytes(data)
    return output
