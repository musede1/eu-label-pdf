# 欧代标签 PDF 生成

填模板 → WPS 导 PDF → 裁剪成正方形标签。

## 环境

- Windows + 已安装 WPS(或 MS Excel)
- Python 3.10+
- 模板目录: `./templates/`(随仓库走,文件名格式 `欧代标签-*-{店铺简称}.xlsx`)
  换路径可设环境变量 `TEMPLATE_DIR` 覆盖

## 安装

```bash
cd eu-label-pdf
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## 用法

```bash
venv\Scripts\python generate.py <店铺简称> <product_name> <model> <batch_number> [--output-dir 路径]
```

示例:

```bash
venv\Scripts\python generate.py 103 "Wooden Vase" "VS-001" "20260415A"
```

默认输出到 `C:\Users\admin\Desktop\代码\输出\`。

## 流程

1. 按 `店铺简称` 在模板目录里找到 `欧代标签-*-{店铺简称}.xlsx`
2. 打开模板,填 `B1 = product_name`,`B2 = model`,`B3 = batch_number`
3. 用 xlwings 驱动 WPS 导出 PDF(`wb.to_pdf()`)
4. 用 PyPDF2 按固定四边边距裁剪:
   - 左 35.2mm / 上 81.3mm / 右 37.1mm / 下 80.2mm
5. 删除中间文件,留最终 `{shop_code}_{batch}_xxxx.pdf`

## 参数

裁剪边距在 `generate.py` 顶部常量里,需要调整直接改:

```python
CROP_LEFT_MM = 35.2
CROP_TOP_MM = 81.3
CROP_RIGHT_MM = 37.1
CROP_BOTTOM_MM = 80.2
```

## 后续

先按脚本跑通,之后计划包成 HTTP 服务,给 Fastify 业务调用。
