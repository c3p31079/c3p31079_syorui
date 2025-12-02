const canvas = document.getElementById("previewCanvas");
const ctx = canvas.getContext("2d");

// A4 サイズに合わせた大きさ
canvas.width = 900;
canvas.height = 1280;

// 点検部位＋項目 → 座標マッピング！必要に応じて追加していくよ！
const coordMap = {
    "チェーン:腐食": { x: 200, y: 180 },
    "チェーン:摩耗": { x: 300, y: 260 },
    "ロープ:亀裂":   { x: 450, y: 320 }
};

// チェック専用の図形座標
const checkCoords = {
    "整備班で対応予定": { x: 120, y: 780 },
    "修繕工事で対応予定": { x: 120, y: 830 },
    "本格的な使用禁止": { x: 120, y: 880 }
};

// 図形描画関数
function drawCircle(x, y, size = 35) {
    ctx.beginPath();
    ctx.lineWidth = 4;
    ctx.strokeStyle = "red";
    ctx.arc(x, y, size, 0, Math.PI * 2);
    ctx.stroke();
}

function drawTriangle(x, y, size = 35) {
    ctx.beginPath();
    ctx.lineWidth = 4;
    ctx.strokeStyle = "red";
    ctx.moveTo(x, y - size);
    ctx.lineTo(x - size, y + size);
    ctx.lineTo(x + size, y + size);
    ctx.closePath();
    ctx.stroke();
}

function drawCross(x, y, size = 30) {
    ctx.beginPath();
    ctx.strokeStyle = "red";
    ctx.lineWidth = 4;
    ctx.moveTo(x - size, y - size);
    ctx.lineTo(x + size, y + size);
    ctx.moveTo(x + size, y - size);
    ctx.lineTo(x - size, y + size);
    ctx.stroke();
}

function drawCheck(x, y, size = 30) {
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

    // チェック項目（複数ある）
    const checkboxes = document.querySelectorAll(".fix-check");

    // プレビュークリア（白背景）
    ctx.fillStyle = "white";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // 点検部位＋項目の図形
    const key = `${part}:${item}`;
    if (coordMap[key]) {
        const { x, y } = coordMap[key];

        if (score >= 0.5) {
            drawTriangle(x, y);
        } else if (score >= 0.2) {
            drawCross(x, y);
        } else {
            drawCircle(x, y);
        }
    }

    // チェックボックスのレ点描画
    checkboxes.forEach(cb => {
        if (cb.checked) {
            const coord = checkCoords[cb.value];
            if (coord) drawCheck(coord.x, coord.y);
        }
    });
}

// 入力で更新
["part", "item", "score"].forEach(id => {
    document.getElementById(id).addEventListener("input", updatePreview);
    document.getElementById(id).addEventListener("change", updatePreview);
});

// チェックボックス
document.querySelectorAll(".fix-check").forEach(cb => {
    cb.addEventListener("change", updatePreview);
});

// 初期描画
updatePreview();
