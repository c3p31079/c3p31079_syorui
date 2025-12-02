from openpyxl import load_workbook
from openpyxl.drawing.shape import Shape
from mapping import POSITION_MAP

# 図形描画（○ △ ×）
def add_shape(ws, shape_type, x, y, size=40):
    shape = Shape()
    shape.width = size
    shape.height = size
    shape.left = x
    shape.top = y

    if shape_type == "circle":
        shape.prst = "ellipse"
        shape.outline.solidFill = "FF0000"
    elif shape_type == "triangle":
        shape.prst = "triangle"
        shape.outline.solidFill = "FF0000"
    elif shape_type == "cross":
        shape.prst = "lineInv"  # ×
        shape.outline.solidFill = "FF0000"

    ws._shapes.append(shape)


# score → 図形判定
def decide_shape(score, part, item):
    if score == 0.5:
        return "triangle"
    if score == 0.2:
        return "cross"

    # ○は score でなく文字条件
    if "共通" in POSITION_MAP and item in POSITION_MAP["共通"]:
        if POSITION_MAP["共通"][item].get("shape") == "circle":
            return "circle"

    return None


# 部位 × 項目 → 座標
def decide_position(part, item):
    # 部位 → 項目 の2段階条件
    if part in POSITION_MAP:
        if item in POSITION_MAP[part]:
            return POSITION_MAP[part][item]["x"], POSITION_MAP[part][item]["y"]

    # ○用（共通）
    if "共通" in POSITION_MAP and item in POSITION_MAP["共通"]:
        return POSITION_MAP["共通"][item]["x"], POSITION_MAP["共通"][item]["y"]

    return None, None


# メイン処理
def process_excel(part, item, score, output_name="output.xlsx"):
    wb = load_workbook("template.xlsx")
    ws = wb.active

    shape_type = decide_shape(score, part, item)
    x, y = decide_position(part, item)

    if shape_type and x:
        add_shape(ws, shape_type, x, y, size=45)

    wb.save(output_name)
    return output_name
