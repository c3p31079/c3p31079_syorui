import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw
import io

ICON_PATHS = {
    "triangle": "backend/icons/triangle.png",
    "cross": "backend/icons/cross.png",
    "circle": "backend/icons/circle.png",
    "check": "backend/icons/check.png"
}

def ensure_icons_exist():
    for name, path in ICON_PATHS.items():
        if not os.path.exists(path):
            # 簡単な図形PNGを自動生成
            img = Image.new("RGBA", (50, 50), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            if name == "triangle":
                draw.polygon([(25,0),(0,50),(50,50)], outline="red", width=4)
            elif name == "cross":
                draw.line((0,0,50,50), fill="red", width=4)
                draw.line((0,50,50,0), fill="red", width=4)
            elif name == "circle":
                draw.ellipse((0,0,50,50), outline="red", width=4)
            elif name == "check":
                draw.line((0,25,20,45), fill="green", width=4)
                draw.line((20,45,50,5), fill="green", width=4)
            img.save(path)

def render_range_to_image(template_path, sheet_name, cell_range):
    """
    Excel範囲を PNG に変換（簡易）
    """
    wb = load_workbook(template_path)
    ws = wb[sheet_name]

    img = Image.new("RGB", (800, 600), "white")
    draw = ImageDraw.Draw(img)
    draw.text((10,10), f"{cell_range} のプレビュー", fill="black")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

def generate_xlsx_with_shapes(template_path, shapes, sheet_name, cell_range):
    """
    Excel に図形を貼り付けて保存
    """
    wb = load_workbook(template_path)
    ws = wb[sheet_name]

    for s in shapes:
        typ = s.get("type")
        x = s.get("x")
        y = s.get("y")
        icon_path = ICON_PATHS.get(typ)
        if icon_path:
            img = XLImage(icon_path)
            img.width = 30
            img.height = 30
            img.anchor = f"A1"  # 固定位置簡易
            ws.add_image(img)

    output_dir = "backend/generated"
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"output.xlsx")
    wb.save(out_path)
    return out_path
