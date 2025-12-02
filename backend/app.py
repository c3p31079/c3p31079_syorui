from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from excel_processor import process_excel

app = Flask(__name__)
CORS(app)

@app.route("/generate", methods=["POST"])
def generate():
    data = request.json

    part  = data.get("part")
    item  = data.get("item")
    score = float(data.get("score", 0))

    output_path = "generated.xlsx"
    process_excel(part, item, score, output_name=output_path)

    return send_file(output_path, as_attachment=True)


@app.route("/", methods=["GET"])
def home():
    return "API is working", 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
