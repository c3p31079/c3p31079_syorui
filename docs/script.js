const BACKEND_URL = "https://c3p31079-syorui.onrender.com"; 

const partSelect = document.getElementById("partSelect");
const itemSelect = document.getElementById("itemSelect");
const checkList = document.getElementById("checkList");
const downloadBtn = document.getElementById("downloadBtn");

let partItemMap = {};
let checkCoordMap = {};

/* 初期化 */
async function init() {
  // 点検部位・項目
  const res1 = await fetch("map.json");
  partItemMap = await res1.json();

  // チェック項目
  const res2 = await fetch("check_coord_map.json");
  checkCoordMap = await res2.json();

  initPartSelect();
  initCheckList();
}

/* プルダウン */
function initPartSelect() {
  partSelect.innerHTML = "";
  Object.keys(partItemMap).forEach(part => {
    const opt = document.createElement("option");
    opt.value = part;
    opt.textContent = part;
    partSelect.appendChild(opt);
  });
  updateItemSelect();
}

function updateItemSelect() {
  const part = partSelect.value;
  itemSelect.innerHTML = "";

  Object.keys(partItemMap[part]).forEach(item => {
    const opt = document.createElement("option");
    opt.value = item;
    opt.textContent = item;
    itemSelect.appendChild(opt);
  });
}

partSelect.onchange = updateItemSelect;

/* チェックボックス */
function initCheckList() {
  checkList.innerHTML = "";
  Object.keys(checkCoordMap).forEach(key => {
    const label = document.createElement("label");
    label.innerHTML = `
      <input type="checkbox" value="${key}">
      ${key}
    `;
    checkList.appendChild(label);
  });
}

/* ダウンロード */
downloadBtn.onclick = async () => {
  const part = partSelect.value;
  const item = itemSelect.value;

  const checks = [...checkList.querySelectorAll("input:checked")]
    .map(c => c.value);

  const payload = {
    part,
    item,
    checks
  };

  const res = await fetch(`${BACKEND_URL}/api/download-excel`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  if (!res.ok) {
    alert("Excel ダウンロードに失敗しました: " + res.status);
    return;
  }

  const blob = await res.blob();
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "inspection.xlsx";
  a.click();

  URL.revokeObjectURL(url);
};

init();
