import os
import io
import json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image

app = Flask(__name__)
CORS(app)

# JSON ファイルパス
COORD_MAP_PATH = "backend/coord_map.json"
CHECK_COORD_MAP_PATH = "backend/check_coord_map.json"
TEMPLATE_PATH = "backend/template.xlsx"
ICON_PATHS = {
    "triangle": "backend/icon/triangle.png",
    "cross": "backend/icon/cross.png",
    "check": "backend/icon/check.png"
}

# JSON 読み込み
with open(COORD_MAP_PATH, encoding="utf-8") as f:
    coord_map = json.load(f)

with open(CHECK_COORD_MAP_PATH, encoding="utf-8") as f:
    check_coord_map = json.load(f)

@app.route("/")
def home():
    return "Backend running"

@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    try:
        data = request.json
        rows = data.get("rows", [])
        checks = data.get("checks", [])

        # Excel テンプレート読み込み
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # 行単位で図形を追加
        for row in rows:
            part = row.get("part")
            item = row.get("item")
            eval_type = row.get("evaluation")  # "triangle" or "cross"

            if part in coord_map and item in coord_map[part]:
                coords = coord_map[part][item].get(eval_type)
                if coords:
                    img = XLImage(ICON_PATHS[eval_type])
                    img.anchor = f"A1"  # 仮に左上にアンカー。後で座標で調整可能
                    img.width = 24
                    img.height = 24
                    ws.add_image(img)
                    img.drawing.top = coords["y"]
                    img.drawing.left = coords["x"]

        # チェックボックス対応
        for check in checks:
            coords = check_coord_map.get(check)
            if coords:
                img = XLImage(ICON_PATHS["check"])
                img.anchor = f"A1"
                img.width = 24
                img.height = 24
                ws.add_image(img)
                img.drawing.top = coords["y"]
                img.drawing.left = coords["x"]

        # Excel バイト出力
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
        print(e)
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
