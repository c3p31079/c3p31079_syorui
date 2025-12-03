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

# 論理記号タイプのマッピング
ICON_FILENAMES = {
    "triangle": "triangle.png",
    "cross": "none.png",   # none.png にマッピング
    "none": "none.png",
    "circle": "circle.png",
    "check": "check.png"
}

def ensure_icons_exist():
    """
    If icons are missing, create simple fallback PNGs (so Render won't crash).
    Prefer to place your own PNGs into backend/icons/ (32x32, transparent).
    """
    for key, fname in ICON_FILENAMES.items():
        path = os.path.join(ICON_DIR, fname)
        if not os.path.exists(path):
            # ファイル名として一度だけ作成
            try:
                im = Image.new("RGBA", (32, 32), (0,0,0,0))
                d = ImageDraw.Draw(im)
                if "tri" in key:
                    d.polygon([(16,2),(2,30),(30,30)], outline=(200,50,50), width=3)
                elif key in ("cross","none"):
                    d.line([(4,4),(28,28)], fill=(200,50,50), width=3)
                    d.line([(28,4),(4,28)], fill=(200,50,50), width=3)
                elif "circle" in key:
                    d.ellipse([(4,4),(28,28)], outline=(200,50,50), width=3)
                elif key == "check":
                    d.line([(4,18),(12,26),(28,6)], fill=(20,150,20), width=3)
                im.save(path)
            except Exception:
                os.makedirs(ICON_DIR, exist_ok=True)
                im.save(path)

def get_icon_path(symbol_type):
    fname = ICON_FILENAMES.get(symbol_type)
    if not fname:
        return None
    p = os.path.join(ICON_DIR, fname)
    return p if os.path.exists(p) else None

def render_range_to_image(template_path, sheet_name="Sheet1", cell_range="A1:H37"):
    """
    Render Excel A1:H37 to PNG (simple grid + cell values).
    This is a simple renderer using Pillow for preview only.
    """
    wb = load_workbook(template_path, data_only=True)
    ws = wb[sheet_name]

    # セル範囲をA1:H37！
    start_col = 1; end_col = 8
    start_row = 1; end_row = 37

    cell_w = 100
    cell_h = 20
    width = cell_w * (end_col - start_col + 1)
    height = cell_h * (end_row - start_row + 1)

    im = Image.new("RGB", (width, height), (255,255,255))
    draw = ImageDraw.Draw(im)
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 12)
    except Exception:
        font = ImageFont.load_default()

    # グリッドとテキストを描画
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            x0 = (c - start_col) * cell_w
            y0 = (r - start_row) * cell_h
            x1 = x0 + cell_w
            y1 = y0 + cell_h
            draw.rectangle([x0, y0, x1, y1], outline=(200,200,200))
            val = ws.cell(row=r, column=c).value
            if val is not None:
                draw.text((x0 + 4, y0 + 2), str(val), fill=(0,0,0), font=font)

    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf

def px_to_cell(x_px, y_px, cell_w=100, cell_h=20, start_col=1, start_row=1):
    """
    Convert pixel coordinate (x_px,y_px) used by front-end into approximate Excel cell (like 'B3').
    This is a best-effort mapping (openpyxl does not place images by pixel).
    """
    col_index = int(x_px // cell_w) + start_col
    row_index = int(y_px // cell_h) + start_row
    col_letter = get_column_letter(col_index)
    return f"{col_letter}{row_index}"

def generate_xlsx_with_shapes(template_path, shapes, sheet_name="Sheet1", cell_range="A1:H37"):
    """
    shapes: [{type: 'triangle'|'cross'|'circle'|'check', x: px, y: px}, ...]
    Place images on corresponding approximate cell addresses and save a new workbook.
    """
    wb = load_workbook(template_path)
    ws = wb[sheet_name]

    for s in shapes:
        typ = s.get("type")
        x = s.get("x")
        y = s.get("y")
        # 正規化タイプの許可部分
        if typ == "none":
            typ = "cross"
        icon_path = get_icon_path(typ)
        if not icon_path:
            continue
        cell_addr = px_to_cell(x, y)
        img = XLImage(icon_path)
        img.width = 32
        img.height = 32
        try:
            ws.add_image(img, cell_addr)
        except Exception:
            #アンカーが失敗した場合はA1に配置するよ
            ws.add_image(img, "A1")

    out_name = f"output_{uuid.uuid4().hex}.xlsx"
    out_path = os.path.join(GENERATED_DIR, out_name)
    wb.save(out_path)
    return out_path
