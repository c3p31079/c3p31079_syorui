from flask import Flask, request, send_file, jsonify, send_from_directory
from flask_cors import CORS
from excel_utils import generate_excel_with_shapes
import os
import json

app = Flask(__name__, static_folder="../docs", static_url_path="/")
CORS(app)

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
COORD_MAP_PATH = os.path.join(BASE_DIR, "coord_map.json")
CHECK_COORD_MAP_PATH = os.path.join(BASE_DIR, "check_coord_map.json")

# JSON読み込み
with open(COORD_MAP_PATH, encoding="utf-8") as f:
    coord_map = json.load(f)

with open(CHECK_COORD_MAP_PATH, encoding="utf-8") as f:
    check_coord_map = json.load(f)

@app.route("/")
def serve_index():
    return send_from_directory(app.static_folder, "index.html")

@app.route("/api/generate-excel", methods=["POST"])
def generate_excel():
    data = request.json
    try:
        output_path = generate_excel_with_shapes(TEMPLATE_PATH, data, coord_map, check_coord_map)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
