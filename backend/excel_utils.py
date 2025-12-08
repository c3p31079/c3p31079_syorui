import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
import os

# 画像ディレクトリとファイル
ICON_DIR = "static/img/icons"
CHECK_ICON = os.path.join(ICON_DIR, "check.png")
CIRCLE_ICON = os.path.join(ICON_DIR, "circle.png")

def create_excel_template():
    """Excelテンプレートを生成"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "点検チェックシート"

    # 公園名・点検年度・設置年度
    ws["C2"] = ""  # 公園名
    ws["H2"] = ""  # 点検年度
    ws["H3"] = ""  # 設置年度

    # 点検部位・項目（サンプル）
    ws["B6"], ws["C6"] = "柱・梁（本体）", ""
    ws["D6"], ws["E6"] = "ぐらつき、破損、変形、腐食、接合部の緩み", ""
    ws["B7"], ws["C7"] = "接合部（継ぎ手）", ""
    ws["D7"], ws["E7"] = "破損、変形、腐食、ボルトの緩み、欠落", ""
    ws["B8"], ws["C8"] = "吊金具", ""
    ws["D8"], ws["E8"] = "破損、変形、腐食、異音、ずれ、摩耗、ボルト緩み、欠落", ""
    ws["B9"], ws["C9"] = "揺動部（チェーン・ロープ）", ""
    ws["D9"], ws["E9"] = "ねじれ、変形、破損、ほつれ、断線、摩耗", ""
    ws["B11"], ws["C11"] = "揺動部（座板・座面）", ""
    ws["D11"], ws["E11"] = "ヒビ、割れ、湾曲、破損、腐朽、金具摩耗、ボルト緩み、欠落", ""
    ws["B12"], ws["C12"] = "安全柵", ""
    ws["D12"], ws["E12"] = "ぐらつき、破損、変形、腐食、ボルト緩み、欠落", ""
    ws["B13"], ws["C13"] = "その他", ""
    ws["D13"], ws["E13"] = "異物、落書き", ""
    ws["B14"], ws["C14"] = "基礎", ""
    ws["D14"], ws["E14"] = "基礎露出、亀裂、破損", ""
    ws["B15"], ws["C15"] = "地表部・安全柵内", ""
    ws["D15"], ws["E15"] = "大きな凹凸、石や根の露出、異物、マットのめくれ、破損、枝", ""

    # 措置・総合結果・対応方針・備考などのエリア（F6:H15）
    # 実施措置: F6:G9
    # 所見: F10:G12
    # 総合結果: F13:G15
    # 対応方針: H6:H10
    # 本格的使用禁止: H11
    # 備考: H12:H15

    return wb, ws

def write_text_at(ws, text, cell, x_offset=0, y_offset=0, font_size=11, bold=False):
    """セル指定+オフセットで文字を書き込み"""
    col = openpyxl.utils.column_index_from_string(cell[0]) + x_offset
    row = int(cell[1:]) + y_offset
    col_letter = get_column_letter(col)
    ws[f"{col_letter}{row}"].value = text
    ws[f"{col_letter}{row}"].font = Font(size=font_size, bold=bold)
    ws[f"{col_letter}{row}"].alignment = Alignment(horizontal="center", vertical="center")

def insert_image_at(ws, img_path, cell, x_offset=0, y_offset=0, width=None, height=None):
    """セル指定+オフセットで画像挿入"""
    if not os.path.exists(img_path):
        print(f"⚠️ 画像が存在しません: {img_path}")
        return
    img = Image(img_path)
    if width: img.width = width
    if height: img.height = height
    col = openpyxl.utils.column_index_from_string(cell[0]) + x_offset
    row = int(cell[1:]) + y_offset
    col_letter = get_column_letter(col)
    ws.add_image(img, f"{col_letter}{row}")

def fill_check_icon(ws, cell, x_offset=0, y_offset=0):
    insert_image_at(ws, CHECK_ICON, cell, x_offset, y_offset, width=16, height=16)

def fill_circle_icon(ws, cell, x_offset=0, y_offset=0):
    insert_image_at(ws, CIRCLE_ICON, cell, x_offset, y_offset, width=16, height=16)

def apply_items(ws, items):
    """
    JSONで受け取った各点検部位・措置・総合結果を一括反映
    items = [
        {
            "part": "柱・梁",
            "cell": "D6",
            "type": "text" / "check" / "circle",
            "text": "ぐらつき、破損",
            "x_offset": 0,
            "y_offset": 0
        },
        ...
    ]
    """
    for item in items:
        typ = item.get("type", "text")
        cell = item.get("cell")
        x = item.get("x_offset", 0)
        y = item.get("y_offset", 0)
        if typ == "text":
            write_text_at(ws, item.get("text", ""), cell, x, y)
        elif typ == "check":
            fill_check_icon(ws, cell, x, y)
        elif typ == "circle":
            fill_circle_icon(ws, cell, x, y)
