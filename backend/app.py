import os
import io
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_utils import generate_excel

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "Backend running"

# ここに入れる: excel_utils.generate_excel をインポート済み
@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400
    
    # Excel生成（自由座標対応）
    excel_bytes = generate_excel(data)

    # ダウンロード用レスポンス
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='inspection.xlsx'
    )
    

if __name__ == "__main__":
    app.run(debug=True, port=5000)
