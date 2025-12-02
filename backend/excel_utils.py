import io
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage, ImageDraw

ICON_DIR = "backend/icons"
os.makedirs(ICON_DIR, exist_ok=True)

def ensure_icons_exist():
    # △/×/〇/チェックを PNG で生成
    size = 50
    # △
    img = PILImage.new("RGBA", (size, size), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    draw.polygon([(size/2,0),(0,size),(size,size)], outline="red", width=4)
    img.save(os.path.join(ICON_DIR,"triangle.png"))
    # ×
    img = PILImage.new("RGBA", (size, size), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    draw.line([(0,0),(size,size)], fill="red", width=4)
    draw.line([(0,size),(size,0)], fill="red", width=4)
    img.save(os.path.join(ICON_DIR,"cross.png"))
    # ○
    img = PILImage.new("RGBA", (size, size), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([(0,0),(size,size)], outline="red", width=4)
    img.save(os.path.join(ICON_DIR,"circle.png"))
    # ✓
    img = PILImage.new("RGBA", (size, size), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    draw.line([(size*0.1,size*0.5),(size*0.4,size*0.8),(size*0.9,size*0.2)], fill="green", width=4)
    img.save(os.path.join(ICON_DIR,"check.png"))

def render_range_to_image(template_path, sheet_name, cell_range):
    # openpyxl + PIL でレンダリング
    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    img = PILImage.new("RGB", (800,600), "white")
    draw = ImageDraw.Draw(img)
    # 簡易描画: セル範囲の境界線などを描画
    # （本格的には excel2img なども可）
    draw.rectangle([0,0,799,599], outline="black")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

def generate_xlsx_with_shapes(template_path, shapes, sheet_name, cell_range, save_markers=False):
    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    for s in shapes:
        icon_path = os.path.join(ICON_DIR,f"{s['type']}.png")
        if os.path.exists(icon_path):
            img = Image(icon_path)
            img.width = img.height = 50
            img.anchor = f"A1"  # 仮、座標に変換する場合は調整
            ws.add_image(img)
    out_path = os.path.join("backend","output.xlsx")
    wb.save(out_path)
    return out_path
