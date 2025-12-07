import openpyxl
import os

def generate_excel_with_shapes(template_path, data, coord_map, check_coord_map):
    """
    Excelを生成して保存し、ファイルパスを返す。
    template_path: テンプレートExcel
    data: フロントから送られたJSONデータ
    coord_map, check_coord_map: Excel上の座標マッピング
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 例: 部位名をA1に書く
    ws["A1"] = str(data.get("parts", ""))

    # 例: チェック項目をB列に書く
    for i, check in enumerate(data.get("checks", []), start=2):
        ws[f"B{i}"] = check

    output_path = os.path.join(os.path.dirname(template_path), "inspection.xlsx")
    wb.save(output_path)
    return output_path
