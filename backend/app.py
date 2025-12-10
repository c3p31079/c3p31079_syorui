import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import io

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = "backend/template.xlsx"
ICON_DIR = "backend/icons"  # 画像アイコンはここに置く

# --- ユーティリティ関数 ---
def insert_icon(ws, cell, icon_filename, dx=0, dy=0):
    icon_path = os.path.join(ICON_DIR, icon_filename)
    if not os.path.exists(icon_path):
        print(f"アイコンが存在しません: {icon_path}")
        return
    img = Image(icon_path)
    img.anchor = cell
    img.anchor._from.colOff = dx * 9525  # OpenPyXLの単位に変換
    img.anchor._from.rowOff = dy * 9525
    ws.add_image(img)

def insert_text(ws, cell, text, dx=0, dy=0):
    ws[cell] = text
    # 位置調整（簡易）
    # OpenPyXLではセル位置自体の微調整は難しいため、必要に応じてセルをずらす

# --- ルート ---
@app.route("/")
def home():
    return "Backend running"

# --- Excel生成API ---
@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "JSONが正しく送信されていません"}), 400

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    items = data.get("items", [])
    for item in items:
        cell = item.get("cell")
        if not cell:
            continue
        dx = item.get("dx", 0)
        dy = item.get("dy", 0)
        item_type = item.get("type")
        # アイコン系
        if item_type in ("icon", "checkbox", "radio") and item.get("icon"):
            insert_icon(ws, cell, os.path.basename(item["icon"]), dx, dy)
        # テキスト系
        elif item_type in ("text", "number") and item.get("value") is not None:
            text = item.get("text") or str(item.get("value"))
            insert_text(ws, cell, text, dx, dy)

    # 出力用バッファに保存
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, download_name="inspection_result.xlsx", as_attachment=True)

# --- 実行 ---
if __name__ == "__main__":
    app.run(debug=True)
