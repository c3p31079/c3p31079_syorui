// 共通フロント処理：index.html と preview.html 両方で読み込む
const BACKEND_BASE = window.BACKEND_BASE || "https://c3p31079-syorui.onrender.com";
let coordsMap = {};

// バックエンドからcoords.jsonを取得
async function loadCoords() {
  const res = await fetch(`${BACKEND_BASE}/api/coords`);
  if (!res.ok) throw new Error("coords fetch failed: " + res.status);
  const data = await res.json();
  coordsMap = data;
  return data;
}

/* index.html 側 */
if (document.getElementById("doPreview")) {
  document.addEventListener("DOMContentLoaded", async () => {
    try {
      await loadCoords();
      initForm();
    } catch (e) {
      console.error(e);
      document.getElementById("status").textContent = "座標データの読み込みに失敗しました";
    }
  });

  function initForm() {
    const partEl = document.getElementById("part");
    const itemEl = document.getElementById("item");
    // 部位リスト（coordsMap のキーから抽出）
    const parts = [...new Set(Object.keys(coordsMap).map(k => k.split(":")[0]))];
    partEl.innerHTML = "";
    parts.forEach(p => partEl.add(new Option(p, p)));

    function updateItems() {
      const part = partEl.value;
      const items = [...new Set(Object.keys(coordsMap)
        .filter(k => k.startsWith(part + ":"))
        .map(k => k.split(":")[1]))];
      itemEl.innerHTML = "";
      items.forEach(it => itemEl.add(new Option(it, it)));
    }

    partEl.addEventListener("change", updateItems);
    updateItems();

    document.getElementById("doPreview").addEventListener("click", async () => {
      const part = partEl.value;
      const item = itemEl.value;
      const symbol = document.getElementById("symbol").value;
      const checks = document.getElementById("checks").value.split(",").map(s => s.trim()).filter(Boolean);

      document.getElementById("status").textContent = "生成中…";

      try {
        const res = await fetch(`${BACKEND_BASE}/api/sheet-image`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ sheet: "Sheet1", range: "A1:H37" })
        });
        if (!res.ok) throw new Error("sheet-image failed: " + res.status);
        const blob = await res.blob();
        const reader = new FileReader();
        reader.onload = function () {
          const imgDataUrl = reader.result;
          const payload = { part, item, symbol, checks, imgDataUrl };
          sessionStorage.setItem("previewPayload", JSON.stringify(payload));
          location.href = "preview.html";
        };
        reader.readAsDataURL(blob);
      } catch (e) {
        console.error(e);
        alert("プレビュー生成に失敗しました: " + e.message);
        document.getElementById("status").textContent = "";
      }
    });
  }
}

/* preview.html 側 */
if (document.getElementById("previewCanvas")) {
  document.addEventListener("DOMContentLoaded", async () => {
    try {
      if (!Object.keys(coordsMap).length) await loadCoords();
      renderPreview();
    } catch (e) {
      console.error(e);
      document.getElementById("status").textContent = "プレビュー読み込み失敗";
    }
  });

  function drawShapeOnCtx(ctx, type, x, y) {
    ctx.lineWidth = 3;
    if (type === "triangle") {
      ctx.strokeStyle = "orange";
      ctx.beginPath(); ctx.moveTo(x, y-12); ctx.lineTo(x-12, y+12); ctx.lineTo(x+12, y+12); ctx.closePath(); ctx.stroke();
    } else if (type === "none") {
      ctx.strokeStyle = "red";
      ctx.beginPath(); ctx.moveTo(x-10,y-10); ctx.lineTo(x+10,y+10); ctx.moveTo(x+10,y-10); ctx.lineTo(x-10,y+10); ctx.stroke();
    } else if (type === "circle") {
      ctx.strokeStyle = "red";
      ctx.beginPath(); ctx.arc(x,y,12,0,Math.PI*2); ctx.stroke();
    } else if (type === "check") {
      ctx.strokeStyle = "green";
      ctx.beginPath(); ctx.moveTo(x-8,y); ctx.lineTo(x-3,y+12); ctx.lineTo(x+12,y-10); ctx.stroke();
    }
  }

  async function renderPreview() {
    const payload = JSON.parse(sessionStorage.getItem("previewPayload") || "{}");
    if (!payload || !payload.imgDataUrl) { alert("プレビュー情報がありません"); return; }

    const canvas = document.getElementById("previewCanvas");
    const ctx = canvas.getContext("2d");
    const img = new Image();
    img.onload = () => {
      // キャンバスのピクセルサイズを画像の本来のサイズに設定
      canvas.width = img.width;
      canvas.height = img.height;
      ctx.clearRect(0,0,canvas.width,canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

      // 図形描画：coordsMap のキー "部位:項目" -> 座標配列
      const key = `${payload.part}:${payload.item}`;
      let poses = coordsMap[key] || [];
      if (!Array.isArray(poses)) poses = [poses];

      poses.forEach(pos => {
        drawShapeOnCtx(ctx, payload.symbol, pos.x, pos.y);
      });

      // checks はページ右などに列挙して表示しておく
      if (payload.checks && payload.checks.length) {
        ctx.font = "14px sans-serif";
        ctx.fillStyle = "#333";
        let y = 20;
        payload.checks.forEach(ch => {
          ctx.fillText("・ " + ch, canvas.width - 260, y);
          y += 18;
        });
      }
    };
    img.src = payload.imgDataUrl;

    // ダウンロード（Excel 生成）
    document.getElementById("downloadExcel").addEventListener("click", async () => {
      const key = `${payload.part}:${payload.item}`;
      let poses = coordsMap[key] || [];
      if (!Array.isArray(poses)) poses = [poses];
      const shapes = poses.map(p => ({ type: payload.symbol, x: p.x, y: p.y }));
      try {
        const res = await fetch(`${BACKEND_BASE}/api/generate-xlsx`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ shapes, sheet: "Sheet1", range: "A1:H37" })
        });
        if (!res.ok) throw new Error("generate-xlsx failed: " + res.status);
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "inspection.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      } catch (e) {
        console.error(e);
        alert("Excel生成に失敗しました: " + e.message);
      }
    });

    document.getElementById("backBtn").addEventListener("click", () => {
      history.back();
    });
  }
}
