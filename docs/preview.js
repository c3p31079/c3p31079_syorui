const canvas = document.getElementById("previewCanvas");
const ctx = canvas.getContext("2d");

function drawShape(type,x,y){
  ctx.lineWidth = 4;
  ctx.strokeStyle = type==="check"?"green":"red";
  if(type==="triangle"){
    ctx.beginPath();
    ctx.moveTo(x,y-25);
    ctx.lineTo(x-25,y+25);
    ctx.lineTo(x+25,y+25);
    ctx.closePath();
    ctx.stroke();
  }else if(type==="cross"){
    ctx.beginPath();
    ctx.moveTo(x-25,y-25);
    ctx.lineTo(x+25,y+25);
    ctx.moveTo(x+25,y-25);
    ctx.lineTo(x-25,y+25);
    ctx.stroke();
  }else if(type==="circle"){
    ctx.beginPath();
    ctx.arc(x,y,25,0,2*Math.PI);
    ctx.stroke();
  }else if(type==="check"){
    ctx.beginPath();
    ctx.moveTo(x-12,y);
    ctx.lineTo(x-8,y+25);
    ctx.lineTo(x+25,y-12);
    ctx.stroke();
  }
}

document.addEventListener("DOMContentLoaded",async()=>{
  const payload = JSON.parse(sessionStorage.getItem("previewPayload"));
  const img = new Image();
  img.onload = ()=>{
    canvas.width = img.width;
    canvas.height = img.height;
    ctx.drawImage(img,0,0);
    const {part,item,subItems} = payload;
    const keys = Array.isArray(subItems)?subItems:[];
    keys.forEach(sub=>{
      const key = [part,item,sub].filter(Boolean).join(":");
      if(coordMap[key]){
        drawShape("triangle",coordMap[key].x,coordMap[key].y);
      }
    });
  };
  img.src = payload.imgDataUrl;
});
