from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import json
import openpyxl
from openpyxl.drawing.image import Image
import io

app = Flask(__name__)
CORS(app)

TEMPLATE = "backend/template.xlsx"
COORD_MAP = "backend/coord_map.json"

@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    template_path = os.path.join(BASE_DIR, "template.xlsx")
    output_dir = os.path.join(BASE_DIR, "generated")
    os.makedirs(output_dir, exist_ok=True)

    output_path = os.path.join(output_dir, "output.xlsx")

    wb = load_workbook(template_path)
    wb.save(output_path)

    return send_file(output_path, as_attachment=True)


# @app.route("/api/download-excel", methods=["POST"])
# def download_excel():
#     data = request.json

#     part = data["part"]
#     item = data["item"]
#     checks = data["checks"]

#     wb = openpyxl.load_workbook(TEMPLATE)
#     ws = wb.active

#     with open(COORD_MAP, encoding="utf-8") as f:
#         coord_map = json.load(f)

#     # △ ×
#     if part in coord_map and item in coord_map[part]:
#         for icon, pos in coord_map[part][item].items():
#             img = Image(f"backend/icons/{icon}.png")
#             img.anchor = pos["cell"] if "cell" in pos else None
#             img.left = pos["x"]
#             img.top = pos["y"]
#             ws.add_image(img)

#     # ✓
#     for chk in checks:
#         pos = coord_map["checks"][chk]
#         img = Image("backend/icons/check.png")
#         img.left = pos["x"]
#         img.top = pos["y"]
#         ws.add_image(img)

#     output = io.BytesIO()
#     wb.save(output)
#     output.seek(0)

#     return send_file(
#         output,
#         download_name="inspection.xlsx",
#         as_attachment=True
#     )


@app.route("/")
def health():
    return "Backend OK"
