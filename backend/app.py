from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_utils import generate_excel, generate_pdf
import os

app = Flask(__name__)
CORS(app)

# 生成したファイルを保存するフォルダ
os.makedirs("generated", exist_ok=True)

@app.route("/")
def home():
    return "Backend running"

# Excelダウンロード用
@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    data = request.get_json()
    output_path = generate_excel(data)
    return send_file(output_path, as_attachment=True)

# PDFダウンロード用
@app.route("/api/download-pdf", methods=["POST"])
def download_pdf():
    data = request.get_json()
    output_path = generate_pdf(data)
    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
