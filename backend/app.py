# -*- coding: utf-8 -*-
"""
Flask サーバー
・フロントから受け取った情報で Excel を生成して返す
"""

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_generator import generate_excel_file
import io

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "Backend Running"

# ---------------------------
# Excel 生成 API
# ---------------------------
@app.route("/api/generate-excel", methods=["POST"])
def api_generate_excel():
    try:
        data = request.get_json()

        rows = data.get("rows", [])
        checks = data.get("checks", [])

        # Excel バイナリ生成
        output_stream = generate_excel_file(rows, checks)

        return send_file(
            output_stream,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="inspection.xlsx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
