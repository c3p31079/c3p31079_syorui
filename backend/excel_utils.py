import os
import io
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw

ICON_DIR = "backend/icons"
os.makedirs(ICON_DIR, exist_ok=True)

def ensure_icons_exist():
    # 単純な赤線で図形を作成
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
    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    
    # 画像化する範囲を単純に描画（Pillowでテーブル描画）
    width, height = 850, 650
    im = Image.new("RGB", (width, height), (255,255,255))
    draw = ImageDraw.Draw(im)
    # 枠線などは簡易描画
    for i in range(0, width, int(width/8)):
        draw.line([(i,0),(i,height)], fill=(200,200,200))
    for j in range(0, height, int(height/37)):
        draw.line([(0,j),(width,j)], fill=(200,200,200))
    # PIL Image to BytesIO
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
            img.anchor = f"{s['x']}:{s['y']}"  # Excel座標は簡易
            ws.add_image(img)
    out_path = os.path.join("backend", f"output_{uuid.uuid4().hex}.xlsx")
    wb.save(out_path)
    return out_path
