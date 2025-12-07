import os
import json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

app = Flask(__name__)
CORS(app)

# パス設定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
GENERATED_DIR = os.path.join(BASE_DIR, "generated")
os.makedirs(GENERATED_DIR, exist_ok=True)

# アイコンディレクトリ
ICON_DIR = os.path.join(BASE_DIR, "icons")

# 起動確認
@app.route("/", methods=["GET"])
def home():
    return "Backend running"

# Excel ダウンロード API
@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    """
    フロントから受け取った情報をもとに
    template.xlsx に図形（画像）を貼り付け、
    Excel を生成して返す
    """

    # リクエスト取得
    data = request.get_json(silent=True)
    if data is None:
        return jsonify({"error": "JSON が不正です"}), 400

    shapes = data.get("shapes", [])
    checks = data.get("checks", [])

    # Excel 読み込み
    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({"error": "template.xlsx が見つかりません"}), 500

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # 図形配置（xy指定）
    def add_icon(icon_name, x, y):
        icon_path = os.path.join(ICON_DIR, icon_name)
        if not os.path.exists(icon_path):
            raise FileNotFoundError(icon_path)

        img = XLImage(icon_path)
        img.anchor = f"A1"
        img.anchor._from.colOff = int(x * 9525)
        img.anchor._from.rowOff = int(y * 9525)
        ws.add_image(img)

    # 点検部位＋項目（△ / ×）
    for s in shapes:
        add_icon(
            s["icon"],   # triangle.png / none.png
            s["x"],
            s["y"]
        )

    # チェック項目（✓）
    for c in checks:
        add_icon(
            "check.png",
            c["x"],
            c["y"]
        )

    # 保存 
    output_path = os.path.join(GENERATED_DIR, "inspection_result.xlsx")
    wb.save(output_path)

    # 返却
    return send_file(
        output_path,
        as_attachment=True,
        download_name="inspection_result.xlsx"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
