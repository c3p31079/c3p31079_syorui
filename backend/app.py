from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
from excel_utils import generate_excel_with_shapes
import os

app = Flask(__name__, template_folder="../docs")  
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.xlsx")
ICON_DIR = os.path.join(os.path.dirname(__file__), "icons")

@app.route("/")
def home():
    # docs内のHTMLをレンダリング
    return render_template("CheckSheet.html")

@app.route("/api/generate-excel", methods=["POST"])
def generate_excel():
    data = request.json
    try:
        output_path = generate_excel_with_shapes(TEMPLATE_PATH, data, ICON_DIR)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
