import os
import io
import json
import uuid
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_utils import render_range_to_image, generate_xlsx_with_shapes, ensure_icons_exist

app = Flask(__name__)
CORS(app)

# tテンプレートパス
TEMPLATE_PATH = "backend/template.xlsx"

# 三角形／×／丸／チェック
ensure_icons_exist()

@app.route("/api/sheet-image", methods=["POST"])
def sheet_image():
    """
    Expect JSON: { "sheet": "Sheet1", "range": "A1:H37" }
    Returns PNG binary.
    """
    data = request.get_json() or {}
    sheet = data.get("sheet", "Sheet1")
    cell_range = data.get("range", "A1:H37")
    # render to in-memory PNG
    img_buf = render_range_to_image(TEMPLATE_PATH, sheet, cell_range)
    return send_file(img_buf, mimetype="image/png", as_attachment=False, download_name="sheet.png")

@app.route("/api/generate-xlsx", methods=["POST"])
def generate_xlsx():
    """
    Expect JSON:
    {
      "shapes": [ {"type":"triangle","x":120,"y":80}, ... ],
      "sheet":"Sheet1",
      "range":"A1:H37",
      "save_markers": false (optional)
    }
    Returns generated xlsx.
    """
    data = request.get_json() or {}
    shapes = data.get("shapes", [])
    sheet = data.get("sheet", "Sheet1")
    rng = data.get("range", "A1:H37")
    save_markers = data.get("save_markers", False)

    out_path = generate_xlsx_with_shapes(TEMPLATE_PATH, shapes, sheet, rng, save_markers=save_markers)
    return send_file(out_path,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=os.path.basename(out_path))

@app.route("/api/icons", methods=["GET"])
def list_icons():
    # マーカーアイコンとそのURLを返す
    icons = ["triangle.png", "cross.png", "circle.png", "check.png"]
    return jsonify({"icons": icons})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
