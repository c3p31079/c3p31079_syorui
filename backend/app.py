from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import json
import os
from excel_generator import create_preview_image, generate_excel_file

app = Flask(__name__)
CORS(app)

#マップの読み込み！
with open("coord_map.json", encoding="utf-8") as f:
    COORD_MAP = json.load(f)
with open("check_coord_map.json", encoding="utf-8") as f:
    CHECK_MAP = json.load(f)

@app.route("/preview", methods=["POST"])
def preview():
    data = request.json
    part = data.get("part", "")
    item = data.get("item", "")
    score = float(data.get("score", 0))
    checked_list = data.get("checked", [])  # 文字列の配列

    png_path = create_preview_image(part, item, score, checked_list)
    return send_file(png_path, mimetype="image/png")

@app.route("/generate", methods=["POST"])
def generate():
    data = request.json
    part = data.get("part", "")
    item = data.get("item", "")
    score = float(data.get("score", 0))
    checked_list = data.get("checked", [])  # 文字列の配列
    key = f"{part}:{item}"
    coords = COORD_MAP.get(key)

    xlsx_path = generate_excel_file(part, item, score, checked_list, coords)
    return send_file(xlsx_path,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name="inspection.xlsx")

@app.route("/", methods=["GET"])
def home():
    return jsonify({"status":"ok"})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)