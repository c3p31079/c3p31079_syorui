import os
import json
import io
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
COORD_MAP_PATH = os.path.join(BASE_DIR, "coord_map.json")
CHECK_COORD_MAP_PATH = os.path.join(BASE_DIR, "check_coord_map.json")
ICON_DIR = os.path.join(BASE_DIR, "icons")

ICON_MAP = {
    "check": os.path.join(ICON_DIR, "check.png"),
    "triangle": os.path.join(ICON_DIR, "triangle.png"),
    "cross": os.path.join(ICON_DIR, "cross.png"),
}

app = Flask(__name__)
CORS(app)

# Excel に画像を xy 座標で貼る
def paste_icon(ws, icon_path, x, y, size=20):
    img = Image(icon_path)
    img.width = size
    img.height = size
    img.anchor = f"A1"
    img._from.colOff = int(x * 9525)   # px → EMU
    img._from.rowOff = int(y * 9525)
    ws.add_image(img)

@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    try:
        data = request.json

        part = data.get("part")             # 点検部位
        item = data.get("item")             # 点検項目
        evaluation = data.get("evaluation") # △ or ×
        checks = data.get("checks", [])     # チェックボックス配列

        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # JSON 読み込み
        with open(COORD_MAP_PATH, encoding="utf-8") as f:
            coord_map = json.load(f)

        with open(CHECK_COORD_MAP_PATH, encoding="utf-8") as f:
            check_coord_map = json.load(f)

        # △ / × を貼る
        if part and item and evaluation:
            mark = "triangle" if evaluation == "△" else "cross"

            if part in coord_map and item in coord_map[part]:
                mark_info = coord_map[part][item].get(mark)
                if mark_info:
                    paste_icon(
                        ws,
                        ICON_MAP[mark],
                        mark_info["x"],
                        mark_info["y"]
                    )

        # ✓ を貼る
        for label in checks:
            if label in check_coord_map:
                info = check_coord_map[label]
                paste_icon(
                    ws,
                    ICON_MAP["check"],
                    info["x"],
                    info["y"]
                )

        # Excel 出力
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="inspection_result.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("❌ ERROR:", e)
        return jsonify({"error": str(e)}), 500


@app.route("/")
def index():
    return "Backend running ✅"

if __name__ == "__main__":
    app.run(debug=True)
