# from flask import Flask, request, send_file
# from flask_cors import CORS
# import io
# from excel_utils import create_excel_template, apply_items

# app = Flask(__name__)
# CORS(app)

# @app.route("/api/generate_excel", methods=["POST"])
# def generate_excel():
#     data = request.json

#     print("=== フロントから受信した items ===")
#     print(data.get("items"))

#     items = data.get("items", [])

#     wb, ws = create_excel_template()
#     apply_items(ws, items)

#     output = io.BytesIO()
#     wb.save(output)
#     output.seek(0)

#     return send_file(
#         output,
#         as_attachment=True,
#         download_name="点検チェックシート.xlsx",
#         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#     )
    

# @app.route("/")
# def home():
#     return "Backend running"

# if __name__ == "__main__":
#     # 他端末からアクセス可能にする場合 host=0.0.0.0 に変更
#     app.run(host="127.0.0.1", port=5000, debug=True)


from flask import Flask, request, send_file
from flask_cors import CORS
import io
from excel_utils import create_excel_template, apply_items

app = Flask(__name__)
CORS(app)  # 全オリジンからのアクセスを許可

@app.route("/api/generate_excel", methods=["POST"])
def generate_excel():
    data = request.json
    print("=== items received ===")
    print(data.get("items"))
    
    items = data.get("items", [])

    wb, ws = create_excel_template()
    apply_items(ws, items)

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
    app.run(host="127.0.0.1", port=5000, debug=True)
