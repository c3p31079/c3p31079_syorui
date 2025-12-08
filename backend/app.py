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

    # 公園名・点検年度・設置年度
    ws["C2"] = data.get("search_park", "")
    ws["H2"] = data.get("inspection_year", "")
    ws["H3"] = data.get("install_year_num", "")

    # JSONで受け取った座標・文字・アイコンを一括反映
    items = data.get("items", [])
    apply_items(ws, items)

    # Excelをバイナリで返す
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name="点検チェックシート.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
