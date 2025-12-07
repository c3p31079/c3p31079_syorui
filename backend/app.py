from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import json
import io
import os

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
ICON_DIR = os.path.join(BASE_DIR, "icon")

# JSON読み込み
with open(os.path.join(BASE_DIR, "coord_map.json"), "r", encoding="utf-8") as f:
    coord_map = json.load(f)

with open(os.path.join(BASE_DIR, "check_coord_map.json"), "r", encoding="utf-8") as f:
    check_map = json.load(f)

@app.route("/")
def home():
    return "Backend running"

@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    try:
        data = request.get_json()
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # 点検部位・項目を配置（△や×など）
        for row in data.get("items", []):
            part = row.get("part")
            item = row.get("item")
            shape_type = row.get("shape")  # triangle / cross / circle
            if part in coord_map and item in coord_map[part]:
                coords = coord_map[part][item].get(shape_type)
                if coords:
                    img_path = os.path.join(ICON_DIR, f"{shape_type}.png")
                    if os.path.exists(img_path):
                        img = XLImage(img_path)
                        img.anchor = 'A1'
                        img.width = 20
                        img.height = 20
                        img.drawing.top = coords['y']
                        img.drawing.left = coords['x']
                        ws.add_image(img)

        # チェック項目を配置（✓）
        for check_label in data.get("checks", []):
            if check_label in check_map:
                coords = check_map[check_label]
                img_path = os.path.join(ICON_DIR, "check.png")
                if os.path.exists(img_path):
                    img = XLImage(img_path)
                    img.anchor = 'A1'
                    img.width = 20
                    img.height = 20
                    img.drawing.top = coords['y']
                    img.drawing.left = coords['x']
                    ws.add_image(img)

        # Excelを返す
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name="inspection.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
