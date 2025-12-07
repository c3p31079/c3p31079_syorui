import openpyxl
from openpyxl.drawing.image import Image
import os

ICON_DIR = os.path.join(os.path.dirname(__file__), 'icons')

# 部位ごとの自由座標（Excel上のセル名または座標）
PART_POSITIONS = {
    "chain": {"B":"F5","C":"G5"},
    "joint": {"B":"F6","C":"G6"},
    "pole": {"B":"F7","C":"G7"},
    "seat": {"B":"F8","C":"G8"},
}

# 措置の座標
ACTION_POSITIONS = {
    "grease":"B10",
    "bolt":"B11",
    "hanger":"B12",
    "chain":"B13",
    "seat":"B14",
    "removal":"B15",
    "other":"B16"
}

# 対応方針の座標
PLAN_POSITIONS = {
    "maintenance":"C20",
    "repair":"C21",
    "improvement":"C22",
    "precision":"C23",
    "removal":"C24",
    "other":"C25"
}

# 上旬・中旬・下旬座標
PERIOD_POSITIONS = {
    "early":"D30",
    "mid":"E30",
    "late":"F30"
}

# 総合結果座標
OVERALL_POSITIONS = {
    "A":"G5",
    "B":"G6",
    "C":"G7",
    "D":"G8"
}

def add_image(ws, img_name, cell):
    img_path = os.path.join(ICON_DIR, img_name)
    if os.path.exists(img_path):
        img = Image(img_path)
        ws.add_image(img, cell)

def write_to_excel(data, template_path):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 公園名、点検日、設置年度
    ws["B1"] = data.get("park_name", "")
    ws["C1"] = data.get("inspection_date", "")
    ws["D1"] = data.get("install_year", "")

    # 点検部位
    for part in data.get("parts", []):
        name = part["part"]
        grade = part["grade"]
        if grade in ["B","C"]:
            icon = "triangle.png" if grade=="B" else "none.png"
            pos = PART_POSITIONS.get(name, {}).get(grade)
            if pos:
                add_image(ws, icon, pos)

    # 措置
    actions = data.get("actions", {})
    for action, value in actions.items():
        if isinstance(value, bool) and value:
            pos = ACTION_POSITIONS.get(action)
            if pos:
                add_image(ws, "check.png", pos)
        elif isinstance(value, int):
            pos = ACTION_POSITIONS.get(action)
            if pos:
                ws[pos] = value

    # 対応方針
    plans = data.get("plans", {})
    for plan, value in plans.items():
        if isinstance(value, bool) and value:
            pos = PLAN_POSITIONS.get(plan)
            if pos:
                add_image(ws, "check.png", pos)
        elif isinstance(value, str):
            pos = PLAN_POSITIONS.get(plan)
            if pos:
                ws[pos] = value

    # 対応予定時期
    month = data.get("response_month")
    period = data.get("period")
    if month:
        ws["E40"] = f"{month}月"
    if period:
        pos = PERIOD_POSITIONS.get(period)
        if pos:
            add_image(ws, "circle.png", pos)

    # 総合結果
    overall = data.get("overall")
    if overall:
        pos = OVERALL_POSITIONS.get(overall)
        if pos:
            add_image(ws, "check.png", pos)
        # D の詳細欄
        if overall == "D":
            ws["H8"] = data.get("overall_d_detail","")

    # 備考・所見
    ws["B50"] = data.get("remarks","")
    ws["B51"] = data.get("observations","")

    return wb
