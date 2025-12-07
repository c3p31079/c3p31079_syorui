// JSON読み込み
let coordMap = {};  // 点検部位・項目用
let checkMap = [];  // チェック項目用

fetch("coord_map.json")
  .then(res => res.json())
  .then(data => coordMap = data)
  .catch(err => console.error(err));

fetch("check_coord_map.json")
  .then(res => res.json())
  .then(data => checkMap = Object.keys(data))
  .catch(err => console.error(err));

const itemsTable = document.getElementById("itemsTable");
const addRowBtn = document.getElementById("addRowBtn");
const checksContainer = document.getElementById("checksContainer");
const downloadBtn = document.getElementById("downloadBtn");

// チェックボックス生成
function createCheckBoxes() {
  checksContainer.innerHTML = "";
  checkMap.forEach(label => {
    const div = document.createElement("div");
    div.className = "check-item";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = label;
    input.id = `check-${label}`;
    const labelEl = document.createElement("label");
    labelEl.htmlFor = `check-${label}`;
    labelEl.textContent = label;
    div.appendChild(input);
    div.appendChild(labelEl);
    checksContainer.appendChild(div);
  });
}

// 新しい行追加
function addRow(part="", item="", shape="triangle") {
  const row = document.createElement("tr");

  // 点検部位
  const partTd = document.createElement("td");
  const partSelect = document.createElement("select");
  partSelect.className = "partSelect";
  Object.keys(coordMap).forEach(key => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = key;
    if(key === part) opt.selected = true;
    partSelect.appendChild(opt);
  });
  partTd.appendChild(partSelect);
  row.appendChild(partTd);

  // 項目
  const itemTd = document.createElement("td");
  const itemSelect = document.createElement("select");
  itemSelect.className = "itemSelect";
  function updateItemOptions() {
    itemSelect.innerHTML = "";
    const selectedPart = partSelect.value;
    if(coordMap[selectedPart]) {
      Object.keys(coordMap[selectedPart]).forEach(i => {
        const opt = document.createElement("option");
        opt.value = i;
        opt.textContent = i;
        if(i === item) opt.selected = true;
        itemSelect.appendChild(opt);
      });
    }
  }
  updateItemOptions();
  partSelect.addEventListener("change", updateItemOptions);
  itemTd.appendChild(itemSelect);
  row.appendChild(itemTd);

  // 評価（図形）
  const shapeTd = document.createElement("td");
  const shapeSelect = document.createElement("select");
  ["triangle","cross","circle"].forEach(s => {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    if(s===shape) opt.selected=true;
    shapeSelect.appendChild(opt);
  });
  shapeTd.appendChild(shapeSelect);
  row.appendChild(shapeTd);

  itemsTable.querySelector("tbody").appendChild(row);
}

// 初期化
addRow();       // 最低1行
createCheckBoxes();

// 行追加ボタン
addRowBtn.addEventListener("click", ()=> addRow());

// Excelダウンロード
downloadBtn.addEventListener("click", async ()=>{
  try {
    const rows = itemsTable.querySelectorAll("tbody tr");
    const items = Array.from(rows).map(tr=>{
      return {
        part: tr.querySelector(".partSelect").value,
        item: tr.querySelector(".itemSelect").value,
        shape: tr.querySelector("select:nth-child(1)").value
      };
    });
    const checks = Array.from(checksContainer.querySelectorAll("input[type=checkbox]:checked"))
      .map(cb => cb.value);

    const res = await fetch("/api/download-excel", {
      method: "POST",
      headers: { "Content-Type":"application/json" },
      body: JSON.stringify({ items, checks })
    });

    if(!res.ok) throw new Error(res.status);

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "inspection.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

  } catch(e) {
    alert("Excel ダウンロードに失敗しました: " + e);
    console.error(e);
  }
});
