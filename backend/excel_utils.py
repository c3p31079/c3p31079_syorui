from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import coordinate_to_tuple
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_DIR = os.path.join(BASE_DIR, "static", "icons")


def create_excel_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "点検チェックシート"

    # テンプレ最小構成（実際は既存を想定）
    ws["C2"] = ""
    ws["H2"] = ""
    ws["H3"] = ""

    return wb, ws


def apply_items(ws, items):
    """
    items = [
      {
        type: "check" | "circle" | "text",
        cell: "F6",
        dx: 2,
        dy: 4,
        icon: "check.png",
        text: "任意"
      }
    ]
    """

    for item in items:
        item_type = item.get("type")
        cell = item.get("cell")
        dx = item.get("dx", 0)
        dy = item.get("dy", 0)

        if not cell:
            continue

        if item_type in ["check", "circle"]:
            icon_name = item.get("icon")
            if not icon_name:
                continue

            icon_path = os.path.join(ICON_DIR, icon_name)
            if not os.path.exists(icon_path):
                print(f"⚠ アイコン未存在: {icon_path}")
                continue

            _insert_image(ws, cell, icon_path, dx, dy)

        elif item_type == "text":
            text = item.get("text", "")
            _insert_text(ws, cell, text, dx, dy)


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
