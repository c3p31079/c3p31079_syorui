from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D  # 追加
from openpyxl.utils import column_index_from_string
from flask_cors import CORS
import io
import os
from openpyxl.utils import get_column_letter


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
    # 修正：tupleではなくXDRPositiveSize2Dを使用
    anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(ICON_PX*EMU, ICON_PX*EMU))
    img.anchor = anchor
    ws.add_image(img)


def insert_text(ws, cell, text, dx=0, dy=0):
    """
    結合セルでもテキストを書き込めるように修正
    """
    col_letter = ''.join(filter(str.isalpha, cell))
    row_number = int(''.join(filter(str.isdigit, cell)))
    col = column_index_from_string(col_letter) - 1
    row = row_number - 1

    # 結合セルの左上を取得
    for merged_range in ws.merged_cells.ranges:
        if cell in merged_range:
            cell = merged_range.start_cell.coordinate
            col_letter = ''.join(filter(str.isalpha, cell))
            row_number = int(''.join(filter(str.isdigit, cell)))
            col = column_index_from_string(col_letter) - 1
            row = row_number - 1
            break

    # 後で対処
    ws.cell(row=row+1, column=col+1, value=text)


@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():
    data = request.get_json(silent=True)
    if data is None:
        return jsonify({"error": "JSONが正しく送信されていません"}), 400

    print("[INFO] Received data:", data)

    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({"error": "テンプレートファイルが見つかりません"}), 500

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # ============================
    # 1. ラジオボタン結果を反映
    # ============================
    radio_buttons = data.get("radio_buttons", {})
    for cell, value in radio_buttons.items():
        if value == "A":
            continue
        # 結合セル対応
        for merged_range in ws.merged_cells.ranges:
            if cell in merged_range:
                cell = merged_range.start_cell.coordinate
                break
        if value == "B":
            insert_icon(ws, cell, "triangle.png")
        elif value == "C":
            insert_icon(ws, cell, "none.png")

    # ============================
    # 2. items 配列を反映
    # ============================
    items = data.get("items", [])
    for item in items:
        cell = item.get("cell")
        if not cell:
            continue

        if item.get("type") == "icon" and item.get("icon"):
            dx = item.get("dx", 0)
            dy = item.get("dy", 0)
            insert_icon(ws, cell, item["icon"], dx=dx, dy=dy)
        elif item.get("type") == "text" and item.get("text") is not None:
            dx = item.get("dx", 0)
            dy = item.get("dy", 0)
            insert_text(ws, cell, item["text"], dx=dx, dy=dy)

    # ExcelをBytesIOに保存
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(
        stream,
        as_attachment=True,
        download_name="点検チェックシート.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)
