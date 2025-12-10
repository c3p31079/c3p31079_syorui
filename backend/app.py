from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.utils import column_index_from_string
from flask_cors import CORS
import io
import os

EMU = 9525
ICON_PX = 32

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
ICON_DIR = os.path.join(BASE_DIR, "icons")

app = Flask(__name__)
CORS(app)

def insert_icon(ws, cell, icon_file, dx=0, dy=0):
    """Excelの指定セルにアイコン画像を配置"""
    icon_path = os.path.join(ICON_DIR, icon_file)
    if not os.path.exists(icon_path):
        print(f"[WARN] icon not found: {icon_path}")
        return

    img = Image(icon_path)
    img.width = ICON_PX
    img.height = ICON_PX

    col_letter = ''.join(filter(str.isalpha, cell))
    row_number = int(''.join(filter(str.isdigit, cell)))

    col = column_index_from_string(col_letter) - 1
    row = row_number - 1

    marker = AnchorMarker(col=col, colOff=dx*EMU, row=row, rowOff=dy*EMU)
    anchor = OneCellAnchor(_from=marker, ext=(ICON_PX*EMU, ICON_PX*EMU))
    img.anchor = anchor
    ws.add_image(img)

@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():
    data = request.get_json(silent=True)
    if data is None:
        return jsonify({"error": "JSONが正しく送信されていません"}), 400

    print("[INFO] Received data:", data)

    # Excel読み込み
    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({"error": "テンプレートファイルが見つかりません"}), 500

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # -------------------------------
    # JSONの構造に合わせてアイコンを配置
    # -------------------------------
    # data = {
    #   "inspection_sections": [
    #       {
    #           "items": [
    #               { "type": "icon", "cell": "F6", "dx": 2, "dy": 4, "icon": "check.png" },
    #               { "type": "text", "cell": "F8", "dx": 35, "dy": 0, "text": "2" },
    #               ...
    #           ]
    #       }
    #   ]
    # }
    sections = data.get("inspection_sections", [])
    for section in sections:
        for item in section.get("items", []):
            if item.get("type") != "icon":
                continue
            cell = item.get("cell")
            icon_file = item.get("icon")
            dx = item.get("dx", 0)
            dy = item.get("dy", 0)
            if cell and icon_file:
                insert_icon(ws, cell, icon_file, dx, dy)

    # ExcelをBytesIOに保存
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    # ファイル送信
    return send_file(
        stream,
        as_attachment=True,
        download_name="点検チェックシート.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)
