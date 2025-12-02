const canvas = document.getElementById("previewCanvas");
const ctx = canvas.getContext("2d");

// 図形の座標マッピング（例）
const coordMap = {
    "チェーン:腐食": { x: 120, y: 80 },
    "チェーン:摩耗": { x: 220, y: 160 },
    "ロープ:亀裂": { x: 340, y: 200 }
};

// ◯ / △ / × の描画関数
function drawCircle(x, y, size = 30) {
    ctx.beginPath();
    ctx.lineWidth = 4;
    ctx.strokeStyle = "red";
    ctx.arc(x, y, size, 0, Math.PI * 2);
    ctx.stroke();
}

function drawTriangle(x, y, size = 30) {
    ctx.beginPath();
    ctx.lineWidth = 4;
    ctx.strokeStyle = "red";
    ctx.moveTo(x, y - size);
    ctx.lineTo(x - size, y + size);
    ctx.lineTo(x + size, y + size);
    ctx.closePath();
    ctx.stroke();
}

function drawCross(x, y, size = 25) {
    ctx.beginPath();
    ctx.strokeStyle = "red";
    ctx.lineWidth = 4;
    ctx.moveTo(x - size, y - size);
    ctx.lineTo(x + size, y + size);
    ctx.moveTo(x + size, y - size);
    ctx.lineTo(x - size, y + size);
    ctx.stroke();
}

function drawCheck(x, y, size = 25) {
    ctx.beginPath();
    ctx.strokeStyle = "green";
    ctx.lineWidth = 4;
    ctx.moveTo(x - size, y);
    ctx.lineTo(x - size/3, y + size);
    ctx.lineTo(x + size, y - size/2);
    ctx.stroke();
}

// プレビュー更新処理
function updatePreview() {
    const part = document.getElementById("part").value;
    const item = document.getElementById("item").value;
    const score = Number(document.getElementById("score").value);
    const checked = document.getElementById("check").checked;

    // キャンバスクリア
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const key = `${part}:${item}`;
    if (!(key in coordMap)) return; // 対応なし

    const { x, y } = coordMap[key];

    // スコアに応じた図形
    if (score >= 0.5) {
        drawTriangle(x, y);
    } else if (score >= 0.2) {
        drawCross(x, y);
    } else {
        drawCircle(x, y); // 例：スコアが小さい場合は○
    }

    // チェックボックス → レ点追加
    if (checked) {
        drawCheck(x + 70, y - 10);
    }
}

// 入力変更イベントでリアルタイム更新
["part", "item", "score", "check"].forEach(id => {
    document.getElementById(id).addEventListener("input", updatePreview);
    document.getElementById(id).addEventListener("change", updatePreview);
});

// 初期状態
updatePreview();
