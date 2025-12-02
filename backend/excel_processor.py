from openpyxl import load_workbook
from openpyxl.drawing.text import RichText, Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.drawing.shape import Shape
from mapping import POSITION_MAP


# ○ △ × ✔ の図形を追加
def add_shape(ws, shape_type, x, y, size=40):
    shape = Shape()
    shape.width = size
    shape.height = size
    shape.left = x
    shape.top = y

    # 図形の種類
    if shape_type == "circle":
        shape.prst = "ellipse"
        shape.outline.solidFill = "FF0000"

    elif shape_type == "triangle":
        shape.prst = "triangle"
        shape.outline.solidFill = "FF0000"

    elif shape_type == "cross":
        shape.prst = "lineInv"
        shape.outline.solidFill = "FF0000"

    elif shape_type == "check":  # ✔ の描画
        shape.prst = "rect"
        shape.outline.solidFill = "FF0000"

        # 図形の中央にテキスト ✔ を入れる
        text = RichText()
        cp = CharacterProperties()
        cp.b = True
        rt = Paragraph(
            pPr=ParagraphProperties(),
            r=[Paragraph(rPr=cp, t="✔")]
        )
        shape.text = "✔"

    ws._shapes.append(shape)


# score → 図形判定
def decide_shape(score, part, item):
    # △
    if score == 0.5:
        return "triangle"

    # ×
    if score == 0.2:
        return "cross"

    # ○ は item で決める
    if "共通" in POSITION_MAP and item in POSITION_MAP["共通"]:
        if POSITION_MAP["共通"][item].get("shape") == "circle":
            return "circle"

    return None


# チェックボックスの形状（レ点）
def checkbox_shape(is_checked):
    if is_checked:
        return "check"
    return None


# 部位 × 項目 → 座標
def decide_position(part, item):
    # 部位 → 項目 の条件
    if part in POSITION_MAP:
        if item in POSITION_MAP[part]:
            return POSITION_MAP[part][item]["x"], POSITION_MAP[part][item]["y"]

    # ○（共通）
    if "共通" in POSITION_MAP and item in POSITION_MAP["共通"]:
        return POSITION_MAP["共通"][item]["x"], POSITION_MAP["共通"][item]["y"]

    return None, None


# メイン処理
def process_excel(part, item, score, is_checked, output_name="output.xlsx"):
    wb = load_workbook("template.xlsx")
    ws = wb.active

    # score + 文字条件の図形
    shape_type = decide_shape(score, part, item)
    x, y = decide_position(part, item)

    if shape_type and x is not None:
        add_shape(ws, shape_type, x, y, size=45)

    # レ点（チェックボックス）
    check_shape = checkbox_shape(is_checked)
    if check_shape and x is not None:
        add_shape(ws, check_shape, x + 50, y, size=45)  # 少し右に描く

    wb.save(output_name)
    return output_name
