from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
ICON_DIR = os.path.join(BASE_DIR, "icons")

def create_excel_template():
    """Excelテンプレートを読み込む"""
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
        "text": "...",
        "icon": "ok.png",
        "dx": 0,
        "dy": 0
    }
    """
    for item in items:
        if "cell" not in item or "type" not in item:
            continue

        cell = item["cell"]

        # テキスト挿入
        if item["type"] == "text":
            ws[cell].value = str(item.get("text", ""))
            continue

        # アイコン挿入
        if item["type"] == "icon":
            icon_file = item.get("icon")

            icon_path = os.path.join(ICON_DIR, icon_file)
            if not os.path.exists(icon_path):
                print(f"[WARN] icon not found: {icon_path}")
                continue

            _insert_image(ws, cell, icon_path)

def _insert_image(ws, cell, icon_path):
    """
    openpyxlで画像をセルに挿入（anchorのみ）
    """
    img = Image(icon_path)
    img.anchor = cell
    ws.add_image(img)
