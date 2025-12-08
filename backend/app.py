import os
import io
import json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_utils import generate_inspection_excel  # 自作モジュール

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "Backend running"

@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():
    """
    フロントから送信されたJSONデータを受け取り、Excelを生成して返す
    """
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data received"}), 400

    # Excel生成
    excel_bytes = generate_inspection_excel(data)
    
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="点検チェックシート.xlsx"
    )

if __name__ == "__main__":
    # 開発環境用デバッグモード
    app.run(debug=True)
