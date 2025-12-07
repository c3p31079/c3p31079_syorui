import io
import json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
ICON_DIR = os.path.join(BASE_DIR, "icon")

# JSON ファイルパス
MAP_JSON = os.path.join(BASE_DIR, "coord_map.json")
CHECK_JSON = os.path.join(BASE_DIR, "check_coord_map.json")

# アイコン読み込み
ICON_FILES = {
    "triangle": os.path.join(ICON_DIR, "triangle.png"),
    "cross": os.path.join(ICON_DIR, "cross.png"),
    "check": os.path.join(ICON_DIR, "check.png")
}

# JSON 読み込み
with open(MAP_JSON, encoding="utf-8") as f:
    coord_map = json.load(f)

with open(CHECK_JSON, encoding="utf-8") as f:
    check_map = json.load(f)

@app.route("/")
def home():
    return "Backend is running"

@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    try:
        data = request.get_json()
        items = data.get("items", [])
        checks = data.get("checks", [])

        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # 点検部位・項目の図形配置
        for it in items:
            part = it.get("part")
            item = it.get("item")
            mark = it.get("mark")  # triangle または cross

            # 座標取得
            try:
                xy = coord_map[part][item][mark]
                x, y = xy["x"], xy["y"]
            except KeyError:
                continue  # 見つからなければスキップ

            img_path = ICON_FILES.get(mark)
            if img_path and os.path.exists(img_path):
                img = Image(img_path)
                img.anchor = f'A1'  # Excel上のアンカーはA1固定
                # XY座標を px → EMU に変換
                img.width, img.height = 20, 20  # 適度なサイズ
                img.drawing.top = y
                img.drawing.left = x
                ws.add_image(img)

        # チェック項目の✓
        for c in checks:
            try:
                xy = check_map[c]
                x, y = xy["x"], xy["y"]
            except KeyError:
                continue
            img_path = ICON_FILES.get("check")
            if img_path and os.path.exists(img_path):
                img = Image(img_path)
                img.anchor = "A1"
                img.width, img.height = 20, 20
                img.drawing.top = y
                img.drawing.left = x
                ws.add_image(img)

        # バッファに書き込み
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
