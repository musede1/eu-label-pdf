"""欧代标签 PDF 生成服务 (FastAPI)。

监听 0.0.0.0:5690,提供:
- POST /generate   —— 生成 PDF,二进制返回
- GET  /health     —— 健康检查

启动:
    venv\\Scripts\\python -m uvicorn server:app --host 0.0.0.0 --port 5690
或:
    venv\\Scripts\\python server.py
"""
from __future__ import annotations

import asyncio
import logging
import sys
import tempfile
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel, Field

from core import fill_and_export_pdf, find_template, crop_pdf_bytes, _get_app, _APP_LOCK, _close_app

# 所有 COM / Excel 操作必须固定在同一个线程
# (Windows COM 单线程公寓模型,跨线程用已创建的 COM 对象会抛 CoInitialize 错)
COM_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="com")

# Windows GBK 控制台 utf-8
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("eu-label-pdf")

def _run_in_com_thread(fn, *args, **kwargs):
    """把阻塞任务发到 COM 专属线程执行。"""
    loop = asyncio.get_event_loop()
    return loop.run_in_executor(COM_EXECUTOR, lambda: fn(*args, **kwargs))


@asynccontextmanager
async def lifespan(app: FastAPI):
    # 启动时在 COM 线程里预热 Excel/WPS 实例
    log.info("预热 Excel/WPS 实例...")
    try:
        await _run_in_com_thread(_warmup)
        log.info("预热完成")
    except Exception as e:
        log.warning(f"预热失败(将在首次请求时重试): {e}")
    yield
    # 关闭时也在 COM 线程里退出实例
    log.info("关闭 Excel/WPS 实例")
    try:
        await _run_in_com_thread(_close_app)
    except Exception:
        pass
    COM_EXECUTOR.shutdown(wait=False)


def _warmup() -> None:
    with _APP_LOCK:
        _get_app()


app = FastAPI(title="EU Label PDF", version="0.1.0", lifespan=lifespan)

# Excel COM 不能并发调用,全局串行
LOCK = asyncio.Lock()

TMP_DIR = Path(tempfile.gettempdir()) / "eu-label-pdf"
TMP_DIR.mkdir(parents=True, exist_ok=True)


class GenerateRequest(BaseModel):
    shop_code: str = Field(..., description="店铺简称,如 103")
    product_name: str
    model: str
    batch_number: str


@app.get("/health")
async def health():
    return {"ok": True, "service": "eu-label-pdf"}


@app.post("/generate")
async def generate(req: GenerateRequest):
    log.info(f"generate shop_code={req.shop_code} batch={req.batch_number}")
    try:
        template = find_template(req.shop_code)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    except RuntimeError as e:
        raise HTTPException(409, str(e))

    raw_pdf = TMP_DIR / f"{uuid.uuid4().hex}.pdf"
    try:
        async with LOCK:
            # COM 操作必须固定在 COM_EXECUTOR 的那个线程
            await _run_in_com_thread(
                fill_and_export_pdf,
                template,
                req.product_name,
                req.model,
                req.batch_number,
                raw_pdf,
            )
        # 裁剪是纯 Python,不涉及 COM,放默认线程池即可
        pdf_bytes = await asyncio.to_thread(crop_pdf_bytes, raw_pdf)
    except Exception as e:
        log.exception("generate failed")
        raise HTTPException(500, f"生成失败: {e}")
    finally:
        raw_pdf.unlink(missing_ok=True)

    filename = f"{req.shop_code}_{req.batch_number}.pdf"
    log.info(f"generate ok shop_code={req.shop_code} bytes={len(pdf_bytes)}")
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("server:app", host="0.0.0.0", port=5690, log_level="info")
