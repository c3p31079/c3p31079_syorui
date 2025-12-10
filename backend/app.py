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

    # ラジオボタン情報を反映
    # data は { "radio_buttons": { "D6": "B", "D7": "C", ... } } の形式を想定
    radio_buttons = data.get("radio_buttons", {})
    for cell, value in radio_buttons.items():
        if value == "A":
            continue

        # 結合セルの左上セルに修正
        for merged_range in ws.merged_cells.ranges:
            if cell in merged_range:
                cell = merged_range.start_cell.coordinate
                break

        if value == "B":
            insert_icon(ws, cell, "triangle.png")
        elif value == "C":
            insert_icon(ws, cell, "none.png")

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
    # 外部アクセスしたい場合は host='0.0.0.0' に変更
    app.run(debug=True)
