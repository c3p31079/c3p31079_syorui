import json
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from io import BytesIO
import os

# 座標マップの読み込み
with open("backend/coord_map.json", "r", encoding="utf-8") as f:
    COORD_MAP = json.load(f)
with open("backend/check_coord_map.json", "r", encoding="utf-8") as f:
    CHECK_COORD_MAP = json.load(f)

# アイコン画像パス
ICON_PATHS = {
    "△": "backend/icons/triangle.png",
    "×": "backend/icons/cross.png",
    "○": "backend/icons/circle.png",
    "✓": "backend/icons/check.png"
}

TEMPLATE_PATH = "backend/template.xlsx"
GENERATED_DIR = "generated"
os.makedirs(GENERATED_DIR, exist_ok=True)

# Excel生成関数
def generate_excel(data):
    """
    入力データをもとにExcelを生成し、保存パスを返す
    data: {
        "rows": [
            {"part": "...", "item": "...", "subItem": "...", "shape": "△"},
            ...
        ],
        "checks": ["整備班で対応予定", ...]
    }
    """
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # 図形配置
    for row in data.get("rows", []):
        key = f"{row['part']}:{row['item']}"
        if key in COORD_MAP:
            coord = COORD_MAP[key]
            shape_type = row.get("shape")
            if shape_type in ICON_PATHS:
                img = Image(ICON_PATHS[shape_type])
                img.anchor = f"A1"  # 初期位置は左上。後で座標計算
                img.width = 20
                img.height = 20
                # openpyxlではセル座標にピクセル指定できないので、行列で近似
                ws.add_image(img, f"{_coord_to_cell(coord['x'], coord['y'])}")

    # チェック項目配置
    for chk in data.get("checks", []):
        if chk in CHECK_COORD_MAP:
            coord = CHECK_COORD_MAP[chk]
            img = Image(ICON_PATHS["✓"])
            img.anchor = f"A1"
            img.width = 20
            img.height = 20
            ws.add_image(img, f"{_coord_to_cell(coord['x'], coord['y'])}")

    output_path = os.path.join(GENERATED_DIR, "output.xlsx")
    wb.save(output_path)
    return output_path

# PDF生成関数
def generate_pdf(data):
    """
    ExcelをPNG経由でPDFに変換
    """
    excel_path = generate_excel(data)
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    import pandas as pd
    import matplotlib.pyplot as plt

    wb = load_workbook(excel_path)
    ws = wb.active

    # Excel内容をPandasで取得
    data_list = []
    for row in ws.iter_rows(values_only=True):
        data_list.append(list(row))

    df = pd.DataFrame(data_list)
    fig, ax = plt.subplots(figsize=(8, 10))
    ax.axis('off')
    table = ax.table(cellText=df.values, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.5)

    pdf_path = os.path.join(GENERATED_DIR, "output.pdf")
    plt.savefig(pdf_path, bbox_inches='tight')
    plt.close()
    return pdf_path

def _coord_to_cell(x, y):
    """
    x, yピクセル座標から近似セルに変換
    """
    col = chr(min(72, int(x / 64) + 65))  # A〜H固定
    row = min(37, int(y / 20) + 1)
    return f"{col}{row}"
