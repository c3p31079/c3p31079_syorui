from flask import Flask, request, send_file
from flask_cors import CORS
from io import BytesIO
from excel_utils import create_excel_template, apply_items

app = Flask(__name__)
CORS(app)

@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():

    data = request.json
    items = data.get("items", [])

    wb, ws = create_excel_template()
    apply_items(ws, items)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="点検チェックシート.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/")
def home():
    return "Backend running"

if __name__ == "__main__":
    app.run(debug=True)
