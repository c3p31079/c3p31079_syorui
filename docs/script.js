const rowsContainer = document.getElementById("rows-container");
const addRowBtn = document.getElementById("add-row");
const downloadBtn = document.getElementById("download-btn");
const checkContainer = document.getElementById("check-container");

let coordMap = {};
let checkCoordMap = {};
let itemMap = {};
let subItemMap = {};

// JSON 読み込み
async function loadJSON() {
  coordMap = await (await fetch("coord_map.json")).json();
  checkCoordMap = await (await fetch("check_coord_map.json")).json();
  itemMap = await (await fetch("map.json")).then(res => res.json());
  subItemMap = await (await fetch("subitem_map.json")).then(res => res.json());
}

// プルダウン生成
function createRow() {
  const div = document.createElement("div");
  div.className = "row-item";

  const partSelect = document.createElement("select");
  partSelect.innerHTML = `<option value="">--点検部位--</option>`;
  Object.keys(coordMap).forEach(p => {
    partSelect.innerHTML += `<option value="${p}">${p}</option>`;
  });

  const itemSelect = document.createElement("select");
  itemSelect.innerHTML = `<option value="">--項目--</option>`;

  const evalSelect = document.createElement("select");
  evalSelect.innerHTML = `<option value="triangle">△</option><option value="cross">×</option>`;

  partSelect.onchange = () => {
    itemSelect.innerHTML = `<option value="">--項目--</option>`;
    if (coordMap[partSelect.value]) {
      Object.keys(coordMap[partSelect.value]).forEach(i => {
        itemSelect.innerHTML += `<option value="${i}">${i}</option>`;
      });
    }
  };

  div.appendChild(partSelect);
  div.appendChild(itemSelect);
  div.appendChild(evalSelect);
  rowsContainer.appendChild(div);
}

// チェックボックス生成
function createCheckBoxes() {
  checkContainer.innerHTML = "";
  Object.keys(checkCoordMap).forEach(c => {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = c;
    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(c));
    checkContainer.appendChild(label);
  });
}

// ダウンロード
downloadBtn.onclick = async () => {
  const rows = [];
  document.querySelectorAll(".row-item").forEach(row => {
    const selects = row.querySelectorAll("select");
    rows.push({
      part: selects[0].value,
      item: selects[1].value,
      evaluation: selects[2].value
    });
  });

  const checks = [];
  document.querySelectorAll("#check-container input:checked").forEach(cb => {
    checks.push(cb.value);
  });

  try {
    const res = await fetch("/api/download-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows, checks })
    });

    if (!res.ok) throw new Error(`Error ${res.status}`);
    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "inspection.xlsx";
    a.click();
  } catch (err) {
    alert("Excel ダウンロードに失敗しました: " + err);
  }
};

// 初期化
loadJSON().then(() => {
  createRow();
  createCheckBoxes();
});

addRowBtn.onclick = createRow;
