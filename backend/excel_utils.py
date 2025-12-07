import openpyxl
from openpyxl.drawing.image import Image
import os

def generate_excel_with_shapes(template_path, data, icon_dir):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 部位・項目・評価
    start_row = 2
    for i, part in enumerate(data.get("parts", []), start=start_row):
        ws[f"A{i}"] = part.get("part", "")
        ws[f"B{i}"] = part.get("item", "")
        ws[f"C{i}"] = part.get("evaluation", "")

        # 評価アイコン
        symbol = part.get("symbol")
        if symbol:
            img_path = os.path.join(icon_dir, f"{symbol}.png")
            if os.path.exists(img_path):
                img = Image(img_path)
                img.width = 15
                img.height = 15
                ws.add_image(img, f"D{i}")

    # チェック項目
    check_start_row = 10
    for j, check in enumerate(data.get("checks", [])):
        ws[f"E{check_start_row + j}"] = check
        img_path = os.path.join(icon_dir, "check.png")
        if os.path.exists(img_path):
            img = Image(img_path)
            img.width = 15
            img.height = 15
            ws.add_image(img, f"F{check_start_row + j}")

    # 自由記述欄（指定セルに追記）
    for cu in data.get("cell_updates", []):
        start_cell = cu.get("start")
        text = cu.get("text", "")
        if start_cell and text:
            existing = ws[start_cell].value or ""
            if existing:
                ws[start_cell].value = f"{existing}\n{text}"
            else:
                ws[start_cell].value = text

    # 保存
    output_path = os.path.join(os.path.dirname(template_path), "inspection.xlsx")
    wb.save(output_path)
    return output_path
