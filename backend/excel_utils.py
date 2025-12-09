from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import coordinate_to_tuple
import os
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.utils import column_index_from_string
from openpyxl.drawing.xdr import XDRPositiveSize2D



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
        print("IMAGE ITEM", item)

        if "cell" not in item or "type" not in item:
            continue

        cell = item["cell"]
        dx = item.get("dx", 0)
        dy = item.get("dy", 0)

        # ======================
        # テキスト
        # ======================
        if item["type"] == "text":
            ws[cell].value = str(item.get("text", ""))
            continue

        # ======================
        # PNGアイコン
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
    img = Image(icon_path)

    col_letter = ''.join(filter(str.isalpha, cell))
    row = int(''.join(filter(str.isdigit, cell)))
    col = column_index_from_string(col_letter) - 1

    marker = AnchorMarker(
        col=col,
        colOff=dx * 9525,  # px → EMU
        row=row - 1,
        rowOff=dy * 9525
    )

    size = XDRPositiveSize2D(img.width * 9525, img.height * 9525)
    img.anchor = OneCellAnchor(_from=marker, ext=size)

    ws.add_image(img)
