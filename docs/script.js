// JSON データ読み込み
const mapData = JSON.parse(document.getElementById("map-data").textContent);
const checkData = JSON.parse(document.getElementById("check-data").textContent);

const itemsBody = document.getElementById("items-body");
const addRowBtn = document.getElementById("add-row-btn");
const checksContainer = document.getElementById("checks-container");
const downloadBtn = document.getElementById("download-btn");

// 初期化: チェック項目を表示
checkData.forEach(item => {
  const label = document.createElement("label");
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.value = item;
  label.appendChild(checkbox);
  label.appendChild(document.createTextNode(item));
  checksContainer.appendChild(label);
});

// 行追加関数
function addRow() {
  const tr = document.createElement("tr");

  const partTd = document.createElement("td");
  const partSelect = document.createElement("select");
  partSelect.innerHTML = Object.keys(mapData).map(p => `<option value="${p}">${p}</option>`).join("");
  partTd.appendChild(partSelect);

  const itemTd = document.createElement("td");
  const itemSelect = document.createElement("select");
  const updateItems = () => {
    const part = partSelect.value;
    itemSelect.innerHTML = Object.keys(mapData[part]).map(i => `<option value="${i}">${i}</option>`).join("");
  };
  partSelect.addEventListener("change", updateItems);
  updateItems();
  itemTd.appendChild(itemSelect);

  const markTd = document.createElement("td");
  const markSelect = document.createElement("select");
  markSelect.innerHTML = `<option value="triangle">△</option><option value="cross">×</option>`;
  markTd.appendChild(markSelect);

  tr.appendChild(partTd);
  tr.appendChild(itemTd);
  tr.appendChild(markTd);
  itemsBody.appendChild(tr);
}

addRowBtn.addEventListener("click", addRow);
addRow(); // 最初の行

// ダウンロードボタン
downloadBtn.addEventListener("click", async () => {
  const items = [];
  document.querySelectorAll("#items-body tr").forEach(tr => {
    const part = tr.querySelector("td:nth-child(1) select").value;
    const item = tr.querySelector("td:nth-child(2) select").value;
    const mark = tr.querySelector("td:nth-child(3) select").value;
    items.push({ part, item, mark });
  });

  const checks = Array.from(document.querySelectorAll("#checks-container input:checked")).map(c => c.value);

  try {
    const res = await fetch("/api/download-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ items, checks })
    });
    if (!res.ok) throw new Error(res.status);
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "inspection.xlsx";
    a.click();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    alert(`Excel ダウンロードに失敗しました: ${err}`);
  }
});
