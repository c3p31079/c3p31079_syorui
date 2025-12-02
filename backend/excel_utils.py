import os
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw, ImageFont
import uuid

ICON_DIR = "backend/icons"
os.makedirs(ICON_DIR, exist_ok=True)

def ensure_icons_exist():
    # 図形アイコン生成
    shapes = {
        "triangle": lambda: draw_triangle_icon(),
        "cross": lambda: draw_cross_icon(),
        "circle": lambda: draw_circle_icon(),
        "check": lambda: draw_check_icon()
    }
    for name, func in shapes.items():
        path = os.path.join(ICON_DIR, f"{name}.png")
        if not os.path.exists(path):
            func().save(path)

def draw_triangle_icon(size=50):
    im = Image.new("RGBA", (size, size), (255,255,255,0))
    draw = ImageDraw.Draw(im)
    draw.polygon([(size/2,0),(0,size),(size,size)], outline="red", width=4)
    return im

def draw_cross_icon(size=50):
    im = Image.new("RGBA", (size, size), (255,255,255,0))
    draw = ImageDraw.Draw(im)
    draw.line([(0,0),(size,size)], fill="red", width=4)
    draw.line([(0,size),(size,0)], fill="red", width=4)
    return im

def draw_circle_icon(size=50):
    im = Image.new("RGBA", (size, size), (255,255,255,0))
    draw = ImageDraw.Draw(im)
    draw.ellipse([(0,0),(size,size)], outline="red", width=4)
    return im

def draw_check_icon(size=50):
    im = Image.new("RGBA", (size, size), (255,255,255,0))
    draw = ImageDraw.Draw(im)
    draw.line([(0,size/2),(size/3*1,size)], fill="green", width=4)
    draw.line([(size/3*1,size),(size,0)], fill="green", width=4)
    return im

def render_range_to_image(template_path, sheet_name="Sheet1", cell_range="A1:H37"):
    wb = load_workbook(template_path, data_only=True)
    ws = wb[sheet_name]

    # セルの範囲
    start_col = 1; end_col = 8
    start_row = 1; end_row = 37

    # 画像サイズ計算
    cell_w = 100; cell_h = 20
    width = cell_w * (end_col - start_col + 1)
    height = cell_h * (end_row - start_row + 1)
    im = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(im)
    try:
        font = ImageFont.truetype("arial.ttf", 14)
    except:
        font = ImageFont.load_default()

    # セル描画
    for r in range(start_row, end_row+1):
        for c in range(start_col, end_col+1):
            x0 = (c-start_col)*cell_w
            y0 = (r-start_row)*cell_h
            x1 = x0 + cell_w
            y1 = y0 + cell_h
            draw.rectangle([x0,y0,x1,y1], outline=(200,200,200))
            val = ws.cell(r,c).value
            if val is not None:
                draw.text((x0+4,y0+2), str(val), fill="black", font=font)

    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf

def generate_xlsx_with_shapes(template_path, shapes, sheet_name="Sheet1", cell_range="A1:H37", save_markers=False):
    wb = load_workbook(template_path)
    ws = wb[sheet_name]

    for s in shapes:
        icon_file = os.path.join(ICON_DIR, f"{s['type']}.png")
        if os.path.exists(icon_file):
            img = XLImage(icon_file)
            img.width, img.height = 30, 30
            # Excel上で座標を簡単に計算
            col = int(s['x']/100)+1
            row = int(s['y']/20)+1
            img.anchor = f"{get_column_letter(col)}{row}"
            ws.add_image(img)

    out_path = os.path.join("backend", f"output_{uuid.uuid4().hex}.xlsx")
    wb.save(out_path)
    return out_path
