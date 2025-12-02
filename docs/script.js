const BACKEND_BASE = "https://c3p31079-syorui.onrender.com";
let coords = {};

// 初期化：点検部位・項目プルダウン
async function initForm() {
    const res = await fetch(`${BACKEND_BASE}/api/coords`);
    coords = await res.json();

    const partEl = document.getElementById("part");
    const itemEl = document.getElementById("item");
    partEl.innerHTML = "";
    itemEl.innerHTML = "";

    const parts = [...new Set(Object.keys(coords).map(k => k.split(":")[0]))];
    parts.forEach(p => partEl.add(new Option(p, p)));

    function updateItems() {
        const selected = partEl.value;
        itemEl.innerHTML = "";
        Object.keys(coords)
            .filter(k => k.startsWith(selected + ":"))
            .map(k => k.split(":")[1])
            .forEach(i => itemEl.add(new Option(i, i)));
    }

    partEl.addEventListener("change", updateItems);
    updateItems();
}

document.addEventListener("DOMContentLoaded", initForm);

// プレビュー処理
document.getElementById("doPreview").addEventListener("click", async () => {
    const part = document.getElementById("part").value;
    const item = document.getElementById("item").value;
    const score = document.getElementById("score").value;
    const checksRaw = document.getElementById("checks").value;
    const checks = checksRaw.split(",").map(s => s.trim()).filter(Boolean);

    // Backend から A1:H37 PNG を取得
    const res = await fetch(`${BACKEND_BASE}/api/sheet-image`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sheet: "Sheet1", range: "A1:H37" })
    });

    if (!res.ok) {
        alert("プレビュー画像の生成に失敗しました");
        return;
    }

    const blob = await res.blob();
    const reader = new FileReader();
    reader.onload = function () {
        const imgDataUrl = reader.result;
        const payload = { part, item, score, checks, imgDataUrl };
        sessionStorage.setItem("previewPayload", JSON.stringify(payload));
        location.href = "preview.html";
    };
    reader.readAsDataURL(blob);
});

// preview.html 側の描画と Excel 生成 
document.addEventListener("DOMContentLoaded", async () => {
    const canvas = document.getElementById("previewCanvas");
    if (!canvas) return;

    const ctx = canvas.getContext("2d");
    const payload = JSON.parse(sessionStorage.getItem("previewPayload") || "{}");

    if (!payload.imgDataUrl) return;

    // PNG 背景描画
    const img = new Image();
    img.onload = () => {
        canvas.width = img.width;
        canvas.height = img.height;
        ctx.drawImage(img, 0, 0);

        // 図形描画（座標は coords.json から）
        const part = payload.part;
        const item = payload.item;
        const key = `${part}:${item}`;
        let positions = coords[key];

        if (!positions) return;
        if (!Array.isArray(positions)) positions = [positions];

        positions.forEach(pos => {
            const { x, y } = pos;
            ctx.save();
            ctx.strokeStyle = payload.score === "triangle" ? "orange" :
                              payload.score === "cross" ? "red" : "green";
            ctx.lineWidth = 3;
            if (payload.score === "triangle") {
                ctx.beginPath();
                ctx.moveTo(x, y - 10);
                ctx.lineTo(x - 10, y + 10);
                ctx.lineTo(x + 10, y + 10);
                ctx.closePath();
                ctx.stroke();
            } else if (payload.score === "cross") {
                ctx.beginPath();
                ctx.moveTo(x - 10, y - 10);
                ctx.lineTo(x + 10, y + 10);
                ctx.moveTo(x + 10, y - 10);
                ctx.lineTo(x - 10, y + 10);
                ctx.stroke();
            } else if (payload.score === "circle") {
                ctx.beginPath();
                ctx.arc(x, y, 10, 0, Math.PI * 2);
                ctx.stroke();
            }
            ctx.restore();
        });
    };
    img.src = payload.imgDataUrl;

    // Excel ダウンロード
    const downloadBtn = document.getElementById("downloadExcel");
    downloadBtn.addEventListener("click", async () => {
        // Backend に図形情報送信して生成
        const shapes = positions.map(pos => ({ type: payload.score, x: pos.x, y: pos.y }));
        const res = await fetch(`${BACKEND_BASE}/api/generate-xlsx`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ shapes, sheet: "Sheet1", range: "A1:H37" })
        });

        if (!res.ok) {
            alert("Excel 生成に失敗しました");
            return;
        }

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "output.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    });

    // 戻るボタン
    document.getElementById("backBtn").addEventListener("click", () => {
        location.href = "index.html";
    });
});
