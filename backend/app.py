from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from excel_utils import generate_excel_with_shapes
import os

app = Flask(__name__, template_folder="../docs", static_folder="../static")
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "../template.xlsx")

# ページルート
@app.route("/")
def home():
    return render_template("CheckSheet.html")

@app.route("/CheckSheet")
def check_sheet():
    return render_template("CheckSheet.html")

# Excel生成API
@app.route("/api/generate-excel", methods=["POST"])
def generate_excel():
    data = request.json
    try:
        output_path = generate_excel_with_shapes(
            TEMPLATE_PATH, data
        )
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/inspection/<int:inspection_id>/results")
def get_results(inspection_id):
    # 例：JSONを返す
    return jsonify({
        "parts": [
            {"part": "chain", "grade": "A", "confidence": 0.9, "is_ai_predicted": True},
            {"part": "joint", "grade": "B", "confidence": 0.7, "is_ai_predicted": False}
        ]
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
