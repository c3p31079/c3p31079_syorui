# backend/excel_utils.py
import os
import io
import uuid
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

BASE_DIR = os.path.dirname(__file__)
ICON_DIR = os.path.join(BASE_DIR, "icons")
GENERATED_DIR = os.path.join(BASE_DIR, "generated")
os.makedirs(ICON_DIR, exist_ok=True)
os.makedirs(GENERATED_DIR, exist_ok=True)

# アイコンファイル名（ユーザーが用意すること）
ICON_NAMES = {
    "triangle": "triangle.png",
    "none": "none.png",   # ×
    "circle": "circle.png",
    "check": "check.png"
}

def ensure_icons_exist():
    """
    icons 内が空なら簡易アイコンを自動生成しておく（フォールバック）。
    ただし、可能ならユーザーが自作の PNG を置いてください。
    """
    for key, fname in ICON_NAMES.items():
        path = os.path.join(ICON_DIR, fname)
        if not os.path.exists(path):
            # 簡易自動生成（32x32, 透過背景）
            im = Image.new("RGBA", (32, 32), (0,0,0,0))
            d = ImageDraw.Draw(im)
            if key == "triangle":
                d.polygon([(16,2),(2,30),(30,30)], outline=(200,50,50), width=3)
            elif key == "none":
                d.line([(4,4),(28,28)], fill=(200,50,50), width=3)
                d.line([(28,4),(4,28)], fill=(200,50,50), width=3)
            elif key == "circle":
                d.ellipse([(4,4),(28,28)], outline=(200,50,50), width=3)
            elif key == "check":
                d.line([(4,18),(12,26),(28,6)], fill=(20,150,20), width=3)
            try:
                im.save(path)
            except Exception:
                # 念のためフォルダが存在しない場合は作る
                os.makedirs(ICON_DIR, exist_ok=True)
                im.save(path)

def get_icon_path(icon_type):
    fname = ICON_NAMES.get(icon_type)
    if not fname:
        return None
    path = os.path.join(ICON_DIR, fname)
    return path if os.path.exists(path) else None

def render_range_to_image(template_path, sheet_name="Sheet1", cell_range="A1:H37"):
    """
    A1:H37 のような範囲を簡易的に PNG にレンダリングして BytesIO を返す。
    実用上は openpyxl でセル値を読み取り、Pillow で表を描画します。
    （本格的な Excel レンダリングは外部ツールが必要）
    """
    wb = load_workbook(template_path, data_only=True)
    ws = wb[sheet_name]

    # 固定セルサイズで描画（調整可）
    start_col = 1; end_col = 8   # A .. H
    start_row = 1; end_row = 37

    cell_w = 100
    cell_h = 20
    width = cell_w * (end_col - start_col + 1)
    height = cell_h * (end_row - start_row + 1)

    im = Image.new("RGB", (width, height), (255,255,255))
    draw = ImageDraw.Draw(im)

    # 用意できればフォント指定（なければデフォルト）
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 12)
    except Exception:
        font = ImageFont.load_default()

    # グリッドとテキスト描画
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            x0 = (c - start_col) * cell_w
            y0 = (r - start_row) * cell_h
            x1 = x0 + cell_w
            y1 = y0 + cell_h
            draw.rectangle([x0, y0, x1, y1], outline=(200,200,200))
            val = ws.cell(row=r, column=c).value
            if val is not None:
                txt = str(val)
                draw.text((x0+4, y0+2), txt, fill=(0,0,0), font=font)

    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf

def px_to_cell(x_px, y_px, cell_w=100, cell_h=20, start_col=1, start_row=1):
    """
    簡易：px 座標 -> Excel の列行（例: (120,40) -> 'B3'）
    """
    col_index = int(x_px // cell_w) + start_col
    row_index = int(y_px // cell_h) + start_row
    col_letter = get_column_letter(col_index)
    return f"{col_letter}{row_index}"

def generate_xlsx_with_shapes(template_path, shapes, sheet_name="Sheet1", cell_range="A1:H37"):
    """
    shapes: [ {type: "triangle"/"none"/"circle"/"check", x: px, y: px}, ... ]
    px -> セルに変換してアイコンを貼る（簡易）
    """
    wb = load_workbook(template_path)
    ws = wb[sheet_name]

    for s in shapes:
        typ = s.get("type")
        x = s.get("x")
        y = s.get("y")
        icon_path = get_icon_path(typ)
        if not icon_path:
            continue
        # px -> セル
        cell_addr = px_to_cell(x, y)
        img = XLImage(icon_path)
        img.width = 32
        img.height = 32
        # anchor に cell address を使う
        ws.add_image(img, cell_addr)

    out_name = f"output_{uuid.uuid4().hex}.xlsx"
    out_path = os.path.join(GENERATED_DIR, out_name)
    wb.save(out_path)
    return out_path
