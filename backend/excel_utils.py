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

    return wb, ws


def apply_items(ws, items):
    for item in items:
        if not item or "type" not in item:
            continue

        cell = item["cell"]

        if item["type"] == "text":
            _insert_text(
                ws,
                cell,
                item.get("text", ""),
                item.get("dx", 0),
                item.get("dy", 0)
            )

        elif item["type"] in ("circle", "check", "triangle", "cross"):
            icon_file = {
                "circle": "circle.png",
                "check": "check.png",
                "triangle": "triangle.png",
                "cross": "cross.png"
            }[item["type"]]

            icon_path = os.path.join(ICON_DIR, icon_file)
            if os.path.exists(icon_path):
                _insert_image(
                    ws,
                    cell,
                    icon_path,
                    item.get("dx", 0),
                    item.get("dy", 0)
                )


def _insert_image(ws, cell, icon_path, dx, dy):
    img = Image(icon_path)

    row, col = coordinate_to_tuple(cell)
    anchor = ws.cell(row=row, column=col).coordinate

    img.anchor = anchor
    img.left = dx
    img.top = dy

    ws.add_image(img)


def _insert_text(ws, cell, text, dx, dy):
    base_cell = ws[cell]

    safe = str(text).replace("\r", "").replace("\n", " ")

    if base_cell.value:
        base_cell.value = f"{base_cell.value} {safe}"
    else:
        base_cell.value = safe
