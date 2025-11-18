from flask import Flask, request, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate():
    data = request.json
    inspector = data["inspector"]
    date = data["date"]
    target = data["target"]
    score = data["score"]

    # Load template
    wb = load_workbook("template.xlsx")
    ws = wb.active

    # Replace placeholders
    mapping = {
        "【点検者名】": inspector,
        "【日付】": date,
        "【対象物】": target,
        "【評価】": score
    }

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value in mapping:
                cell.value = mapping[cell.value]

    # Save output
    output_path = "output.xlsx"
    wb.save(output_path)

    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    # Render 用
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
