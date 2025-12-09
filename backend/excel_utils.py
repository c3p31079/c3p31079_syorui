from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import coordinate_to_tuple
import os

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
ICON_DIR = os.path.join(BASE_DIR, "icons")


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

        # 文字入れ
        if not item or "type" not in item:
            continue

        if item["type"] == "text":
            ws[cell].value = item["text"]

        # PNGアイコン（○△×✓）
        if item["type"] in ("circle", "check", "triangle", "cross"):
            icon_file = {
                "circle": "circle.png",
                "triangle": "triangle.png",
                "cross": "cross.png",
                "check": "check.png"
            }.get(item["type"])

            if not icon_file:
                continue

            icon_path = os.path.join(ICON_DIR, icon_file)
            if not os.path.exists(icon_path):
                continue

            img = Image(icon_path)
            img.anchor = cell  # ★ まずはセル左上に固定
            ws.add_image(img)



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
