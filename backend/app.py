import json
import io
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = "backend/template.xlsx"
ICON_PATHS = {
    "triangle": "backend/icon/triangle.png",
    "cross": "backend/icon/cross.png",
    "check": "backend/icon/check.png"
}

# JSON読み込み
with open("backend/coord_map.json", encoding="utf-8") as f:
    coord_map = json.load(f)

with open("backend/check_coord_map.json", encoding="utf-8") as f:
    check_coord_map = json.load(f)

@app.route("/")
def home():
    return "Backend running"

@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    try:
        data = request.get_json()
        items = data.get("items", [])
        checks = data.get("checks", [])

        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # 点検項目の図形配置
        for item in items:
            part = item.get("part")
            subitem = item.get("item")
            mark = item.get("mark")
            pos = coord_map.get(part, {}).get(subitem, {}).get(mark)
            if pos:
                img = Image(ICON_PATHS.get(mark))
                img.anchor = f"{pos['x']},{pos['y']}"
                ws.add_image(img)

        # チェック項目のチェックマーク配置
        for check in checks:
            pos = check_coord_map.get(check)
            if pos:
                img = Image(ICON_PATHS.get("check"))
                img.anchor = f"{pos['x']},{pos['y']}"
                ws.add_image(img)

        # ファイルをバイトで返す
        file_stream = io.BytesIO()
        wb.save(file_stream)
        file_stream.seek(0)
        return send_file(
            file_stream,
            as_attachment=True,
            download_name="inspection.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        print("Error:", e)
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
