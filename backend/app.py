from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import json
from excel_generator import generate_excel

app = Flask(__name__)
CORS(app)

# 座標マッピング読込
with open("coord_map.json", encoding="utf-8") as f:
    COORD_MAP = json.load(f)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.json

    part = data.get("part")
    item = data.get("item")
    score = float(data.get("score"))
    check = bool(data.get("check"))

    key = f"{part}:{item}"
    coords = COORD_MAP.get(key, None)

    output_path = generate_excel(
        part=part,
        item=item,
        score=score,
        check=check,
        coords=coords
    )

    return send_file(output_path,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name="inspection.xlsx")


@app.route("/", methods=["GET"])
def home():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)