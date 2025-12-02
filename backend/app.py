# backend/app.py
import os
import json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_utils import render_range_to_image, generate_xlsx_with_shapes, ensure_icons_exist

app = Flask(__name__)
# /api/* に対して任意オリジンからのアクセスを許可（Render + GitHub Pages 対応）
CORS(app, resources={r"/api/*": {"origins": "*"}}, supports_credentials=True)

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
COORDS_PATH = os.path.join(os.path.dirname(BASE_DIR), "docs", "coords.json")  # docs/coords.json

# アイコン存在確認（無ければ簡易生成しておく：Render 起動で失敗しないため）
ensure_icons_exist()

@app.route("/")
def home():
    return "Backend running"

@app.route("/api/sheet-image", methods=["POST"])
def sheet_image():
    """
    Request JSON: { "sheet":"Sheet1", "range":"A1:H37" }
    Returns PNG (BytesIO) of the range for preview.
    """
    data = request.get_json() or {}
    sheet = data.get("sheet", "Sheet1")
    cell_range = data.get("range", "A1:H37")
    try:
        img_buf = render_range_to_image(TEMPLATE_PATH, sheet, cell_range)
        return send_file(img_buf, mimetype="image/png", as_attachment=False, download_name="sheet.png")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/generate-xlsx", methods=["POST"])
def generate_xlsx():
    """
    Request JSON:
    {
      "shapes": [{"type":"triangle","x":120,"y":80}, ...],
      "sheet":"Sheet1",
      "range":"A1:H37"
    }
    Returns generated xlsx file.
    """
    data = request.get_json() or {}
    shapes = data.get("shapes", [])
    sheet = data.get("sheet", "Sheet1")
    cell_range = data.get("range", "A1:H37")

    try:
        out_path = generate_xlsx_with_shapes(TEMPLATE_PATH, shapes, sheet, cell_range)
        return send_file(out_path,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True,
                         download_name=os.path.basename(out_path))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/coords", methods=["GET"])
def api_coords():
    # docs/coords.json を返す（フロントはこれを使ってプルダウン・座標取得）
    try:
        with open(COORDS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # Render expects 0.0.0.0 and PORT env var
    app.run(host="0.0.0.0", port=port)
