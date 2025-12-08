from flask import Flask, request, send_file
from flask_cors import CORS
from io import BytesIO
from excel_utils import create_excel_template, apply_items

app = Flask(__name__)
CORS(app)

@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():
    data = request.get_json()

    wb, ws = create_excel_template()

    # 基本情報
    ws["C2"] = data.get("search_park", "")
    ws["H2"] = data.get("inspection_year", "")
    ws["H3"] = data.get("install_year_num", "")

    # 座標JSON一括反映
    items = data.get("items", [])
    apply_items(ws, items)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="点検チェックシート.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)
