import openpyxl
import os
from openpyxl.drawing.image import Image

def generate_excel_with_shapes(template_path, data, icon_dir):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 部位・評価をテーブルに反映
    for i, part in enumerate(data.get("parts", []), start=2):
        ws[f"A{i}"] = part.get("part", "")
        ws[f"B{i}"] = ""  # 項目名固定ならテンプレにあるので空でもOK
        grade = part.get("grade", "")
        if grade:
            icon_file = os.path.join(icon_dir, f"{grade.lower()}.png")
            if os.path.exists(icon_file):
                img = Image(icon_file)
                ws.add_image(img, f"C{i}")  # Excel上の列は適宜調整

    # チェックボックスに応じた check.png を任意セルに貼る
    for i, check in enumerate(data.get("checks", []), start=2):
        icon_file = os.path.join(icon_dir, "check.png")
        if os.path.exists(icon_file):
            img = Image(icon_file)
            ws.add_image(img, f"D{i}")  # 任意の列に配置

    # 自由記述欄を追加
    for remark in data.get("remarks", []):
        cell = ws[remark["name"]]  # テンプレに事前にセル指定しておく
        if cell.value:
            cell.value += f"\n{remark['value']}"
        else:
            cell.value = remark["value"]

    output_path = os.path.join(os.path.dirname(template_path), "inspection.xlsx")
    wb.save(output_path)
    return output_path
