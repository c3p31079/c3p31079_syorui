from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import coordinate_to_tuple
import os

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")


def create_excel_template():
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    ws.title = "点検チェックシート"

    # テンプレ最小構成（実際は既存を想定）
    ws["C2"] = ""
    ws["H2"] = ""
    ws["H3"] = ""

    return wb, ws


def apply_items(ws, items):
    for item in items:

        cell = item["cell"]
        x = item.get("x_offset", 0)
        y = item.get("y_offset", 0)

        target_cell = ws[cell]

        if item["type"] == "text":
            target_cell.value = item["text"]

        elif item["type"] in ("check", "circle"):
            # 図形の代わりに記号で表現（今はこれでOK）
            mark = "✓" if item["type"] == "check" else "○"
            target_cell.value = mark



def _insert_image(ws, cell, icon_path, dx, dy):
    img = Image(icon_path)

    row, col = coordinate_to_tuple(cell)
    anchor = ws.cell(row=row, column=col).coordinate

    img.anchor = anchor
    img.left = dx
    img.top = dy

    ws.add_image(img)


def _insert_text(ws, cell, text, dx, dy):
    """
    セルを基準に文字を描画的に微調整
    """
    base_cell = ws[cell]
    if base_cell.value:
        base_cell.value = str(base_cell.value) + " " + text
    else:
        base_cell.value = text
