const canvas=document.getElementById("previewCanvas");
const ctx=canvas.getContext("2d");

function drawTriangle(x,y,size=30){
  ctx.beginPath(); ctx.lineWidth=4; ctx.strokeStyle="red";
  ctx.moveTo(x,y-size); ctx.lineTo(x-size,y+size); ctx.lineTo(x+size,y+size);
  ctx.closePath(); ctx.stroke();
}
function drawCross(x,y,size=25){
  ctx.beginPath(); ctx.strokeStyle="red"; ctx.lineWidth=4;
  ctx.moveTo(x-size,y-size); ctx.lineTo(x+size,y+size);
  ctx.moveTo(x+size,y-size); ctx.lineTo(x-size,y+size);
  ctx.stroke();
}
function drawCheck(x,y,size=25){
  ctx.beginPath(); ctx.strokeStyle="green"; ctx.lineWidth=4;
  ctx.moveTo(x-size,y); ctx.lineTo(x-size/3,y+size); ctx.lineTo(x+size,y-size/2);
  ctx.stroke();
}

// 座標マップ
const coordMap={
  "チェーン:腐食":{x:120,y:80},
  "チェーン:摩耗":{x:220,y:160},
  "安全柵:緩み":{x:300,y:200}
};

const payload=JSON.parse(sessionStorage.getItem("previewPayload")||"{}");
const img=new Image();
img.onload=function(){
  ctx.clearRect(0,0,canvas.width,canvas.height);
  ctx.drawImage(img,0,0,canvas.width,canvas.height);

  payload.subItem?.forEach(si=>{
    const key=`${payload.part}:${payload.item}`;
    if(coordMap[key]){
      const {x,y}=coordMap[key];
      if(si.includes("×") || si.includes("使用禁止")) drawCross(x,y);
      else if(si.includes("△")) drawTriangle(x,y);
      else drawCheck(x+50,y);
    }
  });
};
img.src=payload.imgDataUrl;

// 戻る
document.getElementById("goBack").addEventListener("click",()=>{location.href="index.html";});

// Excel生成
document.getElementById("downloadExcel").addEventListener("click",async ()=>{
  const shapes = [];
  payload.subItem?.forEach(si=>{
    const key=`${payload.part}:${payload.item}`;
    if(coordMap[key]){
      const {x,y}=coordMap[key];
      let type="check";
      if(si.includes("×")||si.includes("使用禁止")) type="cross";
      else if(si.includes("△")) type="triangle";
      else type="check";
      shapes.push({type,x,y});
    }
  });

  const res=await fetch("https://c3p31079-syorui.onrender.com/api/generate-xlsx",{
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify({shapes,sheet:"Sheet1",range:"A1:H37"})
  });
  if(!res.ok){alert("Excel生成失敗"); return;}
  const blob=await res.blob();
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="inspection.xlsx";
  a.click();
});
