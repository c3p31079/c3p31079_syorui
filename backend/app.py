import os
import io
import json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage

app = Flask(__name__)
CORS(app)

# テンプレート Excel
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.xlsx")

# JSON 読み込み
with open("coord_map.json", encoding="utf-8") as f:
    COORD_MAP = json.load(f)

with open("check_coord_map.json", encoding="utf-8") as f:
    CHECK_COORD_MAP = json.load(f)

# アイコン画像パス
ICONS = {
    "triangle": os.path.join(os.path.dirname(__file__), "triangle.png"),
    "cross": os.path.join(os.path.dirname(__file__), "cross.png"),
    "check": os.path.join(os.path.dirname(__file__), "check.png")
}

@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    try:
        data = request.get_json()
        items = data.get("items", [])
        checks = data.get("checks", [])

        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # 点検部位・項目に応じた図形配置
        for item in items:
            part = item["part"]
            subitem = item["item"]
            mark = item["mark"]  # triangle / cross
            coord = COORD_MAP.get(part, {}).get(subitem, {}).get(mark)
            if coord:
                img = Image(ICONS[mark])
                img.anchor = f"A1"  # まず A1 基準
                ws.add_image(img)
                # openpyxl の xy 指定は限定的 → 別の方法で正確配置可能
                # 基本はセル指定＋座標補正
                img.drawing.left = coord["x"]
                img.drawing.top = coord["y"]

        # チェック項目に応じた✓配置
        for check_name in checks:
            coord = CHECK_COORD_MAP.get(check_name)
            if coord:
                img = Image(ICONS["check"])
                img.anchor = "A1"
                ws.add_image(img)
                img.drawing.left = coord["x"]
                img.drawing.top = coord["y"]

        # バイト出力
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

@app.route("/")
def home():
    return "Backend running"

if __name__ == "__main__":
    app.run(debug=True)
