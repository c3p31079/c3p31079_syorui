const canvas = document.getElementById("previewCanvas");
const ctx = canvas.getContext("2d");

// 仮の座標マップ
const coordMap = {
    "チェーン:腐食": { x: 120, y: 80 },
    "チェーン:摩耗": { x: 220, y: 160 },
    "ロープ:亀裂": { x: 340, y: 200 }
};

// 図形描画関数
function drawCircle(x, y, size=30){
    ctx.beginPath(); ctx.lineWidth=4; ctx.strokeStyle="red";
    ctx.arc(x,y,size,0,Math.PI*2); ctx.stroke();
}
function drawTriangle(x, y, size=30){
    ctx.beginPath(); ctx.lineWidth=4; ctx.strokeStyle="red";
    ctx.moveTo(x,y-size); ctx.lineTo(x-size,y+size); ctx.lineTo(x+size,y+size); ctx.closePath(); ctx.stroke();
}
function drawCross(x,y,size=25){ ctx.beginPath(); ctx.strokeStyle="red"; ctx.lineWidth=4;
    ctx.moveTo(x-size,y-size); ctx.lineTo(x+size,y+size);
    ctx.moveTo(x+size,y-size); ctx.lineTo(x-size,y+size);
    ctx.stroke();
}
function drawCheck(x,y,size=25){ ctx.beginPath(); ctx.strokeStyle="green"; ctx.lineWidth=4;
    ctx.moveTo(x-size,y); ctx.lineTo(x-size/3,y+size); ctx.lineTo(x+size,y-size/2); ctx.stroke();
}

// プレビュー更新
function updatePreview(){
    const part = localStorage.getItem("part")||"チェーン";
    const item = localStorage.getItem("item")||"腐食";
    const score = Number(localStorage.getItem("score")||0.5);
    const checked = localStorage.getItem("check") === "true";

    ctx.clearRect(0,0,canvas.width,canvas.height);

    const key = `${part}:${item}`;
    if (!(key in coordMap)) return;

    const { x, y } = coordMap[key];

    if(score>=0.5) drawTriangle(x,y);
    else if(score>=0.2) drawCross(x,y);
    else drawCircle(x,y);

    if(checked) drawCheck(x+70,y-10);
}

// 戻るボタン
document.getElementById("backBtn").addEventListener("click",()=>window.history.back());

// Excel生成ボタン
document.getElementById("downloadBtn").addEventListener("click",async ()=>{
    const shapes = [
        {type:"triangle", x:120, y:80}
    ]; // 実際は動的に
    const resp = await fetch("http://localhost:5000/api/generate-xlsx",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body: JSON.stringify({shapes:shapes})
    });
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download="output.xlsx"; a.click();
});

updatePreview();
