from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_generator import generate_excel_file
import io

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "Backend Running"

# Excel 生成 API
@app.route("/api/generate-excel", methods=["POST"])
def api_generate_excel():
    try:
        data = request.get_json()

        rows = data.get("rows", [])
        checks = data.get("checks", [])

        # Excel バイナリ生成
        output_stream = generate_excel_file(rows, checks)

        return send_file(
            output_stream,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="inspection.xlsx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/generate-excel", methods=["POST"])
def generate_excel():
    try:
        data = request.get_json()
        part = data.get("part")
        item = data.get("item")
        checks = data.get("checks", [])

        # Excel 作成
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image as XLImage

        template_path = "backend/template.xlsx"
        wb = load_workbook(template_path)
        ws = wb["Sheet1"]

        # 座標：配置後で自由に編集できる構造）
        import json
        with open("backend/coord_map.json", "r", encoding="utf-8") as f:
            coord_map = json.load(f)

        # 点検部位 ＆ 項目 → ○ or × を貼る
        key = f"{part}::{item}"
        if key in coord_map["parts"]:
            pos = coord_map["parts"][key]
            img = XLImage("backend/icons/triangle.png")  # △ など
            img.width = 20
            img.height = 20
            img.anchor = f"CELL:{pos['x']},{pos['y']}"
            ws.add_image(img)

        # チェック項目 → ✓ を貼る
        for check in checks:
            if check in coord_map["checks"]:
                pos = coord_map["checks"][check]
                img = XLImage("backend/icons/check.png")
                img.width = 20
                img.height = 20
                img.anchor = f"CELL:{pos['x']},{pos['y']}"
                ws.add_image(img)

        # 保存
        output_path = "backend/output.xlsx"
        wb.save(output_path)

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
