const canvas = document.getElementById("previewCanvas");
const ctx = canvas.getContext("2d");

function drawCircle(x,y,size=30){ctx.beginPath();ctx.lineWidth=4;ctx.strokeStyle="red";ctx.arc(x,y,size,0,2*Math.PI);ctx.stroke();}
function drawTriangle(x,y,size=30){ctx.beginPath();ctx.lineWidth=4;ctx.strokeStyle="red";ctx.moveTo(x,y-size);ctx.lineTo(x-size,y+size);ctx.lineTo(x+size,y+size);ctx.closePath();ctx.stroke();}
function drawCross(x,y,size=25){ctx.beginPath();ctx.strokeStyle="red";ctx.lineWidth=4;ctx.moveTo(x-size,y-size);ctx.lineTo(x+size,y+size);ctx.moveTo(x+size,y-size);ctx.lineTo(x-size,y+size);ctx.stroke();}
function drawCheck(x,y,size=25){ctx.beginPath();ctx.strokeStyle="green";ctx.lineWidth=4;ctx.moveTo(x-size,y);ctx.lineTo(x-size/3,y+size);ctx.lineTo(x+size,y-size/2);ctx.stroke();}

const BACKEND_BASE = "http://localhost:5000";

function updatePreview(){
    const payload = JSON.parse(sessionStorage.getItem("previewPayload")||"{}");
    if(!payload.imgDataUrl) return;

    const img = new Image();
    img.onload = function(){
        ctx.clearRect(0,0,canvas.width,canvas.height);
        ctx.drawImage(img,0,0,canvas.width,canvas.height);

        // 図形（仮に1つだけ例）
        const {part,item,score,checks} = payload;
        const coordMap = {
            "チェーン:腐食": {x:120,y:80},
            "チェーン:摩耗": {x:220,y:160},
            "ロープ:亀裂": {x:340,y:200}
        };
        const key = `${part}:${item}`;
        if(key in coordMap){
            const {x,y}=coordMap[key];
            if(score>=0.5) drawTriangle(x,y);
            else if(score>=0.2) drawCross(x,y);
            else drawCircle(x,y);
        }

        // チェック
        if(checks && checks.length>0){
            checks.forEach((c,i)=>drawCheck(400,50+i*30));
        }
    };
    img.src = payload.imgDataUrl;
}

document.getElementById("backBtn").addEventListener("click",()=>window.history.back());
document.getElementById("downloadBtn").addEventListener("click",async ()=>{
    const shapes = [
        {type:"triangle", x:120, y:80}
    ]; // 実際は動的に
    const resp = await fetch(`${BACKEND_BASE}/api/generate-xlsx`,{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body: JSON.stringify({shapes:shapes})
    });
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download="output.xlsx";
    a.click();
});

updatePreview();
