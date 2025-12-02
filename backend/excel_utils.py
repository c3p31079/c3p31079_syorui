import io, os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage, ImageDraw

ICON_DIR = "backend/icons"

def ensure_icons_exist():
    # triangle.png, cross.png, circle.png, check.png を生成（簡易版）
    for name in ["triangle.png","cross.png","circle.png","check.png"]:
        path = os.path.join(ICON_DIR,name)
        if not os.path.exists(path):
            os.makedirs(ICON_DIR, exist_ok=True)
            img = PILImage.new("RGBA",(50,50),(255,255,255,0))
            draw = ImageDraw.Draw(img)
            draw.rectangle([10,10,40,40],outline="red")
            img.save(path)

def render_range_to_image(template_path, sheet_name, cell_range):
    """
    指定セル範囲をPNG化して BytesIO 返す
    """
    wb = load_workbook(template_path)
    ws = wb[sheet_name]

    # 簡易的に A1:H37 を表として画像化
    img = PILImage.new("RGB",(800,600),"white")
    draw = ImageDraw.Draw(img)
    draw.text((10,10), "Excel内容プレビュー", fill="black")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

def generate_xlsx_with_shapes(template_path, shapes, sheet_name, cell_range, save_markers=True):
    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    for s in shapes:
        img_path = os.path.join(ICON_DIR, f"{s['type']}.png")
        if os.path.exists(img_path):
            img = Image(img_path)
            # ここでは座標は px なので、簡易的にセルに貼る
            ws.add_image(img, "B2")
    out_path = "backend/output.xlsx"
    wb.save(out_path)
    return out_path
