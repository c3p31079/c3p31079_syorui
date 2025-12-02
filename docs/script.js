// 入力値を sessionStorage へ保存
const previewBtn = document.getElementById("previewBtn");
if (previewBtn) {
    previewBtn.addEventListener("click", () => {
        const data = {
            part: document.getElementById("part").value,
            item: document.getElementById("item").value,
            score: Number(document.getElementById("score").value),
            checked: document.getElementById("checked").checked
        };

        sessionStorage.setItem("inspectData", JSON.stringify(data));
        location.href = "preview.html";
    });
}

// Excel 生成（xlsx）
const excelBtn = document.getElementById("excelBtn");
if (excelBtn) {
    excelBtn.addEventListener("click", () => {
        const part = document.getElementById("part").value;
        const item = document.getElementById("item").value;
        const score = Number(document.getElementById("score").value);
        const checked = document.getElementById("checked").checked;

        const wb = XLSX.utils.book_new();
        const wsData = [
            ["部位", "項目", "スコア", "点検済み"],
            [part, item, score, checked ? "✔" : ""]
        ];

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, "点検データ");
        XLSX.writeFile(wb, "inspection.xlsx");
    });
}

// プレビュー画面：キャンバス表示
const canvas = document.getElementById("previewCanvas");
if (canvas) {
    const ctx = canvas.getContext("2d");

    const coordMap = {
        "チェーン:腐食": { x: 120, y: 80 },
        "チェーン:摩耗": { x: 220, y: 160 },
        "ロープ:亀裂": { x: 340, y: 200 }
    };

    function drawCircle(x, y, size=30) {
        ctx.beginPath();
        ctx.lineWidth = 4;
        ctx.strokeStyle = "red";
        ctx.arc(x, y, size, 0, Math.PI * 2);
        ctx.stroke();
    }

    function drawTriangle(x, y, size=30) {
        ctx.beginPath();
        ctx.lineWidth = 4;
        ctx.strokeStyle = "red";
        ctx.moveTo(x, y-size);
        ctx.lineTo(x-size, y+size);
        ctx.lineTo(x+size, y+size);
        ctx.closePath();
        ctx.stroke();
    }

    function drawCross(x, y, size=25) {
        ctx.beginPath();
        ctx.strokeStyle = "red";
        ctx.lineWidth = 4;
        ctx.moveTo(x-size, y-size);
        ctx.lineTo(x+size, y+size);
        ctx.moveTo(x+size, y-size);
        ctx.lineTo(x-size, y+size);
        ctx.stroke();
    }

    function drawCheck(x, y, size=25) {
        ctx.beginPath();
        ctx.strokeStyle = "green";
        ctx.lineWidth = 4;
        ctx.moveTo(x-size, y);
        ctx.lineTo(x-size/3, y+size);
        ctx.lineTo(x+size, y-size/2);
        ctx.stroke();
    }

    function renderPreview() {
        const raw = sessionStorage.getItem("inspectData");
        if (!raw) return;

        const { part, item, score, checked } = JSON.parse(raw);

        ctx.clearRect(0, 0, canvas.width, canvas.height);

        const key = `${part}:${item}`;
        if (!(key in coordMap)) return;

        const { x, y } = coordMap[key];

        if (score >= 0.5) drawTriangle(x, y);
        else if (score >= 0.2) drawCross(x, y);
        else drawCircle(x, y);

        if (checked) drawCheck(x + 70, y - 10);
    }

    renderPreview();
}