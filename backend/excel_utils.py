import openpyxl
import os
from openpyxl.drawing.image import Image

def generate_excel_with_shapes(template_path, data):
    """
    template_path: Excelテンプレート
    data: JSON {parts: [], checks: [], remarks: "文字列"}
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 点検部位・チェック項目書き込み
    start_row = 2
    for i, part in enumerate(data.get("parts", []), start=start_row):
        ws[f"A{i}"] = part.get("part_name", "")
        ws[f"B{i}"] = part.get("item", "")
        ws[f"C{i}"] = part.get("grade", "")  # A/B/C
        ws[f"D{i}"] = part.get("comment", "")

    # チェック項目を check.png で描画
    for chk in data.get("checks", []):
        img_path = os.path.join(os.path.dirname(__file__), "../static/img/check.png")
        img = Image(img_path)
        row = chk.get("row", start_row)
        col = chk.get("col", 5)  # E列など指定
        ws.add_image(img, f"{openpyxl.utils.get_column_letter(col)}{row}")

    # 備考欄書き込み（既存文字の改行追加）
    remarks_text = data.get("remarks", "")
    remarks_cell = ws["H12"]  # 例: H12
    if remarks_cell.value:
        remarks_cell.value += f"\n{remarks_text}"
    else:
        remarks_cell.value = remarks_text

    # 保存
    output_path = os.path.join(os.path.dirname(template_path), "inspection.xlsx")
    wb.save(output_path)
    return output_path
