import openpyxl
import os

def generate_excel_with_shapes(template_path, data, coord_map, check_coord_map):
    """
    Excel生成関数
    template_path: テンプレートExcel
    data: フロントからのJSON
    coord_map: 部位・項目の座標
    check_coord_map: チェック項目の座標
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 部位をA1に書く
    ws["A1"] = str(data.get("parts", ""))

    # チェック項目をB列以降に書く
    for i, check in enumerate(data.get("checks", []), start=2):
        ws[f"B{i}"] = check

    # 保存先
    output_path = os.path.join(os.path.dirname(template_path), "inspection.xlsx")
    wb.save(output_path)
    return output_path
