from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
import os

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
ICON_DIR = os.path.join(BASE_DIR, "icons")

def create_excel_template():
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    ws.title = "点検チェックシート"
    return wb, ws

def apply_items(ws, items):
    for item in items:
        if "cell" not in item or "type" not in item:
            continue

        cell = item["cell"]

        # テキスト
        if item["type"] == "text":
            ws[cell].value = str(item.get("text", ""))
            continue

        # アイコン
        if item["type"] == "icon":
            icon_file = item.get("icon")
            icon_path = os.path.join(ICON_DIR, icon_file)

            if not os.path.exists(icon_path):
                print(f"[WARN] icon not found: {icon_path}")
                continue

            _insert_image(
                ws,
                cell,
                icon_path,
                dx=item.get("dx", 0),
                dy=item.get("dy", 0)
            )

def _insert_image(ws, cell, icon_path, dx=0, dy=0):
    """
    dx, dy を考慮して画像をセル内に配置（openpyxl 3.1.5 対応）
    """
    img = Image(icon_path)

    col_letter = ''.join(filter(str.isalpha, cell))
    row = int(''.join(filter(str.isdigit, cell)))

    col = ws[col_letter + "1"].column - 1
    row = row - 1

    marker = AnchorMarker(
        col=col,
        colOff=dx * 9525,   # px → EMU
        row=row,
        rowOff=dy * 9525
    )

    anchor = OneCellAnchor(_from=marker, ext=None)
    img.anchor = anchor
    ws.add_image(img)
