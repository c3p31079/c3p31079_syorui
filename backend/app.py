from flask import Flask, request, send_file
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import column_index_from_string
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
import os
import io

app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
ICON_DIR = os.path.join(BASE_DIR, "icons")

def _insert_image(ws, cell, icon_path, dx=0, dy=0):
    img = Image(icon_path)
    col_letter = ''.join(filter(str.isalpha, cell))
    row = int(''.join(filter(str.isdigit, cell)))
    col = column_index_from_string(col_letter) - 1

    marker = AnchorMarker(col=col, colOff=dx*9525, row=row-1, rowOff=dy*9525)
    size = XDRPositiveSize2D(img.width*9525, img.height*9525)
    img.anchor = OneCellAnchor(_from=marker, ext=size)
    ws.add_image(img)

@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():
    data = request.get_json()

    # テンプレート読み込み
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # 公園名や年なども反映（例）
    ws["B2"].value = data.get("search_park", "")
    ws["B3"].value = data.get("inspection_year", "")

    # アイテムを順番に処理
    for item in data.get("items", []):
        if item["type"] == "icon":
            icon_file = item.get("icon")
            icon_path = os.path.join(ICON_DIR, icon_file)
            if os.path.exists(icon_path):
                _insert_image(ws, item["cell"], icon_path, item.get("dx",0), item.get("dy",0))
        elif item["type"] == "text":
            ws[item["cell"]].value = item.get("text", "")

    # バイトIOに保存して返す
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     download_name="点検チェックシート.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
