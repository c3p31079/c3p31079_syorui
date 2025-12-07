import openpyxl, os

def generate_excel_with_shapes(template_path, data, coord_map, check_coord_map):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 部位・項目・評価をテーブル形式で書き込む
    parts = data.get("parts", [])
    for row_idx, part in enumerate(parts, start=2):
        ws[f"A{row_idx}"] = part.get("part", "")
        ws[f"B{row_idx}"] = part.get("item", "")
        ws[f"C{row_idx}"] = part.get("evaluation", "")

    # チェック項目を横軸に書き込む
    checks = data.get("checks", [])
    for idx, check in enumerate(checks):
        coord = check_coord_map.get(check, {"x": 0, "y": 0})
        col = coord["x"] // 10 + 1  # 適当にセル換算
        ws.cell(row=1, column=col, value=check)

    output_path = os.path.join(os.path.dirname(template_path), "inspection.xlsx")
    wb.save(output_path)
    return output_path
