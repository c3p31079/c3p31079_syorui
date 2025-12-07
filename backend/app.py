import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import json
from excel_utils import generate_excel_with_shapes  # 実装済み関数を前提

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
COORD_MAP_PATH = os.path.join(BASE_DIR, "coord_map.json")
CHECK_COORD_MAP_PATH = os.path.join(BASE_DIR, "check_coord_map.json")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")

with open(COORD_MAP_PATH, encoding="utf-8") as f:
    coord_map = json.load(f)
with open(CHECK_COORD_MAP_PATH, encoding="utf-8") as f:
    check_coord_map = json.load(f)

@app.route("/api/coord-map")
def get_coord_map():
    return jsonify(coord_map)

@app.route("/api/check-coord-map")
def get_check_coord_map():
    return jsonify(check_coord_map)

@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    try:
        data = request.json
        file_path = generate_excel_with_shapes(TEMPLATE_PATH, data, coord_map, check_coord_map)
        return send_file(file_path, as_attachment=True, download_name="inspection.xlsx")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/generate-excel", methods=["POST"])
def generate_excel():
    data = request.json  # JSONデータを受け取る
    try:
        output_path = generate_excel_with_shapes(TEMPLATE_PATH, data, coord_map={}, check_coord_map={})
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
