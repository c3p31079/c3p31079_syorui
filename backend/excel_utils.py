from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import column_index_from_string
import os

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
ICON_DIR = os.path.join(BASE_DIR, "icons")

def create_excel_template():
    """
    Excelテンプレートを読み込み、ワークブックとシートを返す
    """
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    ws.title = "点検チェックシート"
    return wb, ws

def apply_items(ws, items):
    """
    アイテムリストをワークシートに適用
    item = {
        "cell": "B2",
        "type": "text" / "icon",
        "text": "...",       # textの場合
        "icon": "ok.png",    # iconの場合
        "dx": 0,             # 任意
        "dy": 0              # 任意
    }
    """
    for item in items:
        if "cell" not in item or "type" not in item:
            continue

        cell = item["cell"]
        dx = item.get("dx", 0)
        dy = item.get("dy", 0)

        # ======================
        # テキスト挿入
        # ======================
        if item["type"] == "text":
            ws[cell].value = str(item.get("text", ""))
            continue

        # ======================
        # アイコン挿入
        # ======================
        if item["type"] == "icon":
            icon_file = item.get("icon")
            if not icon_file:
                continue

            icon_path = os.path.join(ICON_DIR, icon_file)
            if not os.path.exists(icon_path):
                print(f"[WARN] icon not found: {icon_path}")
                continue

            _insert_image(ws, cell, icon_path, dx, dy)

def _insert_image(ws, cell, icon_path, dx, dy):
    """
    openpyxlで画像を指定セルに挿入（dx, dyで微調整）
    """
    img = Image(icon_path)

    # 列と行番号の取得
    col_letter = ''.join(filter(str.isalpha, cell))
    row_number = int(''.join(filter(str.isdigit, cell)))
    col_index = column_index_from_string(col_letter)

    # anchorにセルを指定
    img.anchor = cell

    # オフセット(px)設定（3.1系）
    img.offset = (dx, dy)

    # ワークシートに画像追加
    ws.add_image(img)
