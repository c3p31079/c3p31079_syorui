from openpyxl import load_workbook
from openpyxl.drawing.shapes import Shape
from openpyxl.drawing.image import ShapeProperties
from openpyxl.styles.colors import Color
from openpyxl.drawing.line import LineProperties

import uuid

BASE_FILE = "templates/base.xlsx"

def create_shape(shape_type, x, y, size, color):
    """
    shape_type: circle / triangle / cross / check
    """

    sp = ShapeProperties()
    sp.ln = LineProperties()
    sp.ln.solidFill = Color(color)

    shape = Shape(
        shape_id=str(uuid.uuid4()),
        name=shape_type,
        prst=shape_type,     # "ellipse" "triangle" など
        x=x,
        y=y,
        cx=size,
        cy=size,
        sp=sp
    )
    return shape


def generate_excel(part, item, score, check, coords):
    wb = load_workbook(BASE_FILE)
    ws = wb.active

    # 文字書き換え例
    ws["B2"] = part
    ws["B3"] = item

    # 図形を置く位置（座標が辞書で来る）
    if coords:
        x = coords["x"]
        y = coords["y"]

        # スコア判定
        if score >= 0.5:
            shape_type = "triangle"
            color = "FF0000"
        elif score >= 0.2:
            shape_type = "cross"     # ×
            color = "FF0000"
        else:
            shape_type = "ellipse"   # ○
            color = "FF0000"

        shape = create_shape(shape_type, x, y, 800000, color)
        ws._drawing.shapes.append(shape)

        # チェックボックス → ✓ マーク
        if check:
            check_shape = create_shape("check", x+200000, y-200000, 600000, "00AA00")
            ws._drawing.shapes.append(check_shape)

    out_path = f"/tmp/result_{uuid.uuid4()}.xlsx"
    wb.save(out_path)
    return out_path