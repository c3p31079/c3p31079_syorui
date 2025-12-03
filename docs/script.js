const formContainer = document.getElementById("formContainer");
const checkContainer = document.getElementById("checkContainer");
const addRowBtn = document.getElementById("addRow");
const downloadExcelBtn = document.getElementById("downloadExcel");
const downloadPDFBtn = document.getElementById("downloadPDF");
const statusEl = document.getElementById("status");

// チェックボックス生成
fetch("https://c3p31079-syorui.onrender.com/backend/check_coord_map.json")
  .then(r => r.json())
  .then(data => {
    Object.keys(data).forEach(k => {
      const label = document.createElement("label");
      const input = document.createElement("input");
      input.type = "checkbox";
      input.value = k;
      label.appendChild(input);
      label.appendChild(document.createTextNode(k));
      checkContainer.appendChild(label);
      checkContainer.appendChild(document.createElement("br"));
    });
  });

// フォーム行追加
function addRow() {
  const div = document.createElement("div");
  div.className = "formRow";

  // 点検部位
  const partSelect = document.createElement("select");
  partSelect.innerHTML = Object.keys(itemMap).map(p => `<option>${p}</option>`).join("");
  div.appendChild(partSelect);

  // 項目
  const itemSelect = document.createElement("select");
  div.appendChild(itemSelect);

  // サブ項目
  const subSelect = document.createElement("select");
  div.appendChild(subSelect);

  // 記号
  const shapeSelect = document.createElement("select");
  ["△","×","○","✓"].forEach(s => {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    shapeSelect.appendChild(opt);
  });
  div.appendChild(shapeSelect);

  // 項目変更時にサブ項目更新
  partSelect.addEventListener("change", () => updateItem(partSelect, itemSelect));
  itemSelect.addEventListener("change", () => updateSub(itemSelect, subSelect));

  updateItem(partSelect, itemSelect);
  updateSub(itemSelect, subSelect);

  formContainer.appendChild(div);
}

function updateItem(partSelect, itemSelect){
  const part = partSelect.value;
  itemSelect.innerHTML = itemMap[part].map(i => `<option>${i}</option>`).join("");
}

function updateSub(itemSelect, subSelect){
  const item = itemSelect.value;
  if(subItemMap[item]){
    subSelect.innerHTML = subItemMap[item].map(s => `<option>${s}</option>`).join("");
  } else {
    subSelect.innerHTML = "";
  }
}

addRowBtn.addEventListener("click", addRow);
addRow(); // 初期行

async function download(type){
  const rows = [];
  document.querySelectorAll(".formRow").forEach(div => {
    rows.push({
      part: div.children[0].value,
      item: div.children[1].value,
      subItem: div.children[2].value,
      shape: div.children[3].value
    });
  });
  const checks = Array.from(checkContainer.querySelectorAll("input:checked")).map(i=>i.value);
  const payload = { rows, checks };

  statusEl.textContent = `${type.toUpperCase()}生成中…`;
  try{
    const res = await fetch(`https://c3p31079-syorui.onrender.com/api/download-${type}`, {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body: JSON.stringify(payload)
    });
    if(!res.ok) throw new Error(res.statusText);
    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = type==="excel"?"output.xlsx":"output.pdf";
    document.body.appendChild(a);
    a.click();
    a.remove();
    statusEl.textContent = `${type.toUpperCase()}をダウンロードしました`;
  } catch(err){
    statusEl.textContent = `エラー: ${err.message}`;
  }
}

downloadExcelBtn.addEventListener("click", ()=>download("excel"));
downloadPDFBtn.addEventListener("click", ()=>download("pdf"));
