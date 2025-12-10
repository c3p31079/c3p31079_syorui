from flask import Flask, request, send_file
from flask_cors import CORS
import io
from excel_utils import create_excel_template, apply_items

app = Flask(__name__)
CORS(app)

@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():
    """
    POSTで受け取ったJSONデータ(items)を元に
    Excelテンプレートに書き込み、ファイルを返却する
    """
    data = request.json
    items = data.get("items", [])

    # Excelテンプレート生成
    wb, ws = create_excel_template()

    # アイテム適用（テキスト・アイコン）
    apply_items(ws, items)

    # メモリ上に保存して返却
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="点検チェックシート.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/")
def home():
    return "Backend running"

if __name__ == "__main__":
    app.run(debug=True)
