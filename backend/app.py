from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_utils import write_to_excel
import io
import os

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.xlsx")

@app.route("/")
def home():
    return "Backend running"

@app.route("/api/save_sheet", methods=["POST"])
def save_sheet():
    data = request.json
    wb = write_to_excel(data, TEMPLATE_PATH)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output,
                     download_name="CheckSheet.xlsx",
                     as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True)
