# Flask API：フロントから rows / checked_items を受け取り Excel を返す
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
import os, io, json
from excel_utils import place_image_xy

app = Flask(__name__, static_folder="../docs", static_url_path="/")
CORS(app)

TEMPLATE = "backend/template.xlsx"
OUTPUT_DIR = "backend/generated"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# アイコンパス（icon key -> filename）
ICON_FILES = {
    "triangle": "backend/icons/triangle.png",
    "cross":    "backend/icons/cross.png",
    "circle":   "backend/icons/circle.png",
    "check":    "backend/icons/check.png"
}

# JSON 読み込みユーティリティ
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# API: map（dummy: フロント向け）
@app.route("/api/map", methods=["GET"])
def api_map():
    # docs/map.json をそのまま返す（フロントは /api/map を参照）
    return send_file("docs/map.json")

# API: check map
@app.route("/api/check-map", methods=["GET"])
def api_check_map():
    return send_file("backend/check_coord_map.json")

# API: coord_map（部位:項目 -> x,y,icon）
@app.route("/api/coord_map", methods=["GET"])
def api_coord_map():
    return send_file("backend/coord_map.json")

# Excel 生成エンドポイント
@app.route("/api/download-excel", methods=["POST"])
def download_excel():
    """
    期待 JSON:
    {
      "rows":[ {"part":"チェーン","item":"腐食","symbol":"triangle"}, ... ],
      "checked_items": ["ボルトの緩み", ...]
    }
    """
    try:
        data = request.get_json() or {}
        rows = data.get("rows", [])
        checked = data.get("checked_items", [])

        # テンプレートをロード
        wb = load_workbook(TEMPLATE)
        ws = wb.active

        # coord_map: 部位:項目 -> {x:, y:, icon:}
        coord_map = load_json("backend/coord_map.json")
        check_map = load_json("backend/check_coord_map.json")

        # 1) 部位×項目 の入力に基づいてアイコンを配置（XY）
        for r in rows:
            part = r.get("part")
            item = r.get("item")
            symbol = r.get("symbol")  # triangle / cross / circle
            key = f"{part}:{item}"
            if key in coord_map:
                entry = coord_map[key]
                x = entry.get("x")
                y = entry.get("y")
                # シンボル選択が優先。もし coord_map に icon 指定があれば優先してもよい
                icon_key = symbol if symbol else entry.get("icon", "triangle")
                icon_path = ICON_FILES.get(icon_key)
                if icon_path and os.path.exists(icon_path):
                    place_image_xy(ws, icon_path, x, y)

        # 2) チェック項目の ✓ を配置
        for chk in checked:
            if chk in check_map:
                pos = check_map[chk]
                x = pos.get("x")
                y = pos.get("y")
                icon_path = ICON_FILES.get("check")
                if icon_path and os.path.exists(icon_path):
                    place_image_xy(ws, icon_path, x, y)

        # 3) 追加で明細も表形式で記録（任意: template に合わせて調整可）
        #    ここでは template の下に書き出す（例: A40から）
        start_row = 40
        col_map = {"part":1, "item":2, "symbol":3}
        for i, r in enumerate(rows):
            row_no = start_row + i
            ws.cell(row=row_no, column=col_map["part"], value=r.get("part"))
            ws.cell(row=row_no, column=col_map["item"], value=r.get("item"))
            ws.cell(row=row_no, column=col_map["symbol"], value=r.get("symbol"))

        # 保存はメモリで行い、send_file 用の BytesIO を返す
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)

        return send_file(
            bio,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="inspection.xlsx"
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# 静的ファイル（docs）配信
@app.route("/", defaults={"path":""})
@app.route("/<path:path>")
def serve_docs(path):
    # docs の静的ファイルを返す（ローカル確認用）
    if path == "" or path == "/":
        return app.send_static_file("index.html")
    return app.send_static_file(path)

if __name__ == "__main__":
    # デバッグ時は port を 5000 にする（必要に応じて変更）
    app.run(host="0.0.0.0", port=5000, debug=True)
