from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os
import json

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = "backend/template.xlsx"
OUTPUT_DIR = "backend/generated"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# JSON 読み込み関数
def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)

# Excel ファイル生成関数
def create_excel(inspector, entries):
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    ws["B2"] = inspector  # 点検者名

    # JSON データ読み込み
    coord_map = load_json("backend/coord_map.json")
    check_coord_map = load_json("backend/check_coord_map.json")
    shapes = {
        "triangle": "backend/icons/triangle.png",
        "cross": "backend/icons/none.png",
        "circle": "backend/icons/circle.png",
        "check": "backend/icons/check.png"
    }

    # 入力内容ごとに Excel に書き込み
    for entry in entries:
        row = ws.max_row + 1
        部位 = entry.get("部位")
        項目 = entry.get("項目")
        評価 = entry.get("評価")
        ws.cell(row=row, column=1, value=部位)
        ws.cell(row=row, column=2, value=項目)
        ws.cell(row=row, column=3, value=評価)

        # 図形配置
        if 部位 in coord_map:
            base_x = coord_map[部位]["x"]
            base_y = coord_map[部位]["y"]
            img_path = None
            if 評価 == "△":
                img_path = shapes["triangle"]
            elif 評価 == "×":
                img_path = shapes["cross"]
            elif 評価 == "○":
                img_path = shapes["circle"]
            elif 評価 == "✓":
                img_path = shapes["check"]
            if img_path and os.path.exists(img_path):
                img = Image(img_path)
                # openpyxl はセル座標単位で配置
                img.anchor = f"A{row+1}"
                ws.add_image(img)

        # チェック項目
        if 部位 == "チェック項目" and 評価 == "✓":
            if 項目 in check_coord_map:
                x = check_coord_map[項目]["x"]
                y = check_coord_map[項目]["y"]
                img = Image(shapes["check"])
                img.anchor = f"A{row+1}"
                ws.add_image(img)

    output_path = os.path.join(OUTPUT_DIR, "inspection.xlsx")
    wb.save(output_path)
    wb.close()
    return output_path

# Excel ダウンロード API
@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    try:
        data = request.get_json()
        inspector = data.get("inspector", "")
        entries = data.get("entries", [])
        path = create_excel(inspector, entries)
        return send_file(path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# JSON API
@app.route("/api/coord_map", methods=["GET"])
def api_coord_map():
    return jsonify(load_json("backend/coord_map.json"))

@app.route("/api/check_coord_map", methods=["GET"])
def api_check_coord_map():
    return jsonify(load_json("backend/check_coord_map.json"))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
