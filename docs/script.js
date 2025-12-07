const inspectBody = document.getElementById("inspectBody");
const addRowBtn = document.getElementById("addRowBtn");
const checkListDiv = document.getElementById("checkList");
const downloadBtn = document.getElementById("downloadBtn");
const statusEl = document.getElementById("status");

let partItemMap = {};      // 点検部位 → 項目[]
let checkItemMap = {};     // チェック項目（キーのみ使用）

/* 初期ロード */
async function init() {
  await loadMaps();
  addRow();
  renderCheckList();
}
init();

/* JSON読込 */
async function loadMaps() {
  // 点検部位・項目
  const res1 = await fetch("/api/map");
  partItemMap = await res1.json();

  // チェック項目
  const res2 = await fetch("/api/check-map");
  checkItemMap = await res2.json();
}

/* 行追加 */
function addRow() {
  const tr = document.createElement("tr");

  // 点検部位
  const tdPart = document.createElement("td");
  const selPart = document.createElement("select");
  selPart.innerHTML = `<option value="">選択</option>`;
  Object.keys(partItemMap).forEach(p => {
    selPart.innerHTML += `<option value="${p}">${p}</option>`;
  });
  tdPart.appendChild(selPart);

  // 項目
  const tdItem = document.createElement("td");
  const selItem = document.createElement("select");
  selItem.innerHTML = `<option value="">選択</option>`;
  tdItem.appendChild(selItem);

  selPart.addEventListener("change", () => {
    selItem.innerHTML = `<option value="">選択</option>`;
    const items = partItemMap[selPart.value] || [];
    items.forEach(i => {
      selItem.innerHTML += `<option value="${i}">${i}</option>`;
    });
  });

  // 評価
  const tdEval = document.createElement("td");
  const selEval = document.createElement("select");
  selEval.innerHTML = `
    <option value="">選択</option>
    <option value="triangle">△</option>
    <option value="cross">×</option>
    <option value="circle">○</option>
  `;
  tdEval.appendChild(selEval);

  // 削除
  const tdDel = document.createElement("td");
  const delBtn = document.createElement("button");
  delBtn.textContent = "削除";
  delBtn.onclick = () => tr.remove();
  tdDel.appendChild(delBtn);

  tr.append(tdPart, tdItem, tdEval, tdDel);
  inspectBody.appendChild(tr);
}

addRowBtn.onclick = addRow;

/* チェック項目 */
function renderCheckList() {
  checkListDiv.innerHTML = "";
  Object.keys(checkItemMap).forEach(label => {
    const el = document.createElement("label");
    el.innerHTML = `
      <input type="checkbox" value="${label}">
      ${label}
    `;
    checkListDiv.appendChild(el);
  });
}

/* Excel ダウンロード */
downloadBtn.onclick = async () => {
  statusEl.textContent = "Excel生成中…";

  const rows = [];
  inspectBody.querySelectorAll("tr").forEach(tr => {
    const sels = tr.querySelectorAll("select");
    if (sels.length === 3 && sels[0].value && sels[1].value && sels[2].value) {
      rows.push({
        part: sels[0].value,
        item: sels[1].value,
        symbol: sels[2].value
      });
    }
  });

  const checked_items = [...document.querySelectorAll("#checkList input:checked")]
    .map(cb => cb.value);

  try {
    const res = await fetch("/api/download-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows, checked_items })
    });

    if (!res.ok) throw new Error(res.status);

    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "inspection.xlsx";
    a.click();
    URL.revokeObjectURL(a.href);

    statusEl.textContent = "ダウンロード完了";
  } catch (e) {
    statusEl.textContent = "Excel ダウンロードに失敗しました";
    console.error(e);
  }
};
