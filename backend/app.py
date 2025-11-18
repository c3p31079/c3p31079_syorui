from flask import Flask, request, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    inspector = data.get("inspector", "")

    template_path = "template.xlsx"
    output_dir = "generated"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "generated.xlsx")

    wb = load_workbook(template_path)
    ws = wb.active

    ws["B2"] = inspector   # Excel の必要なセルに書き込む

    wb.save(output_path)
    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

