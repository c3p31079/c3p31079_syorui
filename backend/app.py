# Excel に図形を配置し、Excel と PNG を返す API
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import json

from excel_render import render_range_to_png
from shape_utils import place_shapes_on_excel

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = "backend/template.xlsx"
GENERATED_DIR = "backend/generated/"
COORD_PATH = "backend/coord_map.json"
CHECK_COORD_PATH = "backend/check_coord_map.json"

# 生成ファイルを置くフォルダ
os.makedirs(GENERATED_DIR, exist_ok=True)


@app.route("/")
def home():
    return "Backend running"


# ① Excel に図形を置き、PNG を返す
@app.route("/api/generate", methods=["POST"])
def generate():
    """
    入力されたデータ（点検部位/項目/記号など）をもとに、
    Excel に図形を配置 → PNG 作成 → PNG(Base64) を返す。
    """
    data = request.json
    items = data.get("items", [])  # 複数行入力

    coord_map = json.load(open(COORD_PATH, encoding="utf-8"))
    check_coord_map = json.load(open(CHECK_COORD_PATH, encoding="utf-8"))

    excel_path = GENERATED_DIR + "output.xlsx"

    # Excel に図形を配置
    place_shapes_on_excel(
        template_path=TEMPLATE_PATH,
        output_path=excel_path,
        items=items,
        coord_map=coord_map,
        check_coord_map=check_coord_map
    )

    # PNG 生成
    png_path = GENERATED_DIR + "preview.png"
    render_range_to_png(
        excel_path,
        "Sheet1",
        "A1:H37",
        png_path
    )

    # PNG をフロントに返す（URL 形式）
    return jsonify({"preview": "/generated/preview.png"})


# ② Excel のダウンロード
@app.route("/api/download")
def download():
    excel_path = GENERATED_DIR + "output.xlsx"
    return send_file(excel_path, as_attachment=True)


# ③ Render 用の静的ファイルルーティング
@app.route("/generated/<path:path>")
def serve_generated(path):
    return send_file(GENERATED_DIR + path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
