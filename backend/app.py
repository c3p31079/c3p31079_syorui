import os
import io
import json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_utils import render_range_to_image, generate_xlsx_with_shapes, ensure_icons_exist

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = "backend/template.xlsx"

# 図形アイコン確認
ensure_icons_exist()

@app.route("/")
def home():
    return "Backend running"

@app.route("/api/sheet-image", methods=["POST"])
def sheet_image():
    """
    JSON: { "sheet": "Sheet1", "range": "A1:H37" }
    PNG を返す
    """
    data = request.get_json() or {}
    sheet = data.get("sheet", "Sheet1")
    cell_range = data.get("range", "A1:H37")

    img_buf = render_range_to_image(TEMPLATE_PATH, sheet, cell_range)
    return send_file(img_buf, mimetype="image/png", as_attachment=False, download_name="sheet.png")

@app.route("/api/generate-xlsx", methods=["POST"])
def generate_xlsx():
    """
    JSON:
    {
      "shapes": [{"type":"triangle","x":120,"y":80}, ...],
      "sheet":"Sheet1",
      "range":"A1:H37"
    }
    """
    data = request.get_json() or {}
    shapes = data.get("shapes", [])
    sheet = data.get("sheet", "Sheet1")
    rng = data.get("range", "A1:H37")

    out_path = generate_xlsx_with_shapes(TEMPLATE_PATH, shapes, sheet, rng)
    return send_file(out_path,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=os.path.basename(out_path))

@app.route("/api/coords", methods=["GET"])
def get_coords():
    # 点検部位・項目・座標マッピング
    coords_path = "docs/coords.json"
    with open(coords_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return jsonify(data)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
