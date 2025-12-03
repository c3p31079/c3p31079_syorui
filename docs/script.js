const inspectorInput = document.getElementById("inspector");
const entryContainer = document.getElementById("entryContainer");
const addEntryBtn = document.getElementById("addEntryBtn");
const checkContainer = document.getElementById("checkContainer");
const downloadExcelBtn = document.getElementById("downloadExcelBtn");

let coordMap = {};
let checkCoordMap = {};

// 初期データ取得
async function init() {
  coordMap = await (await fetch("/api/coord_map")).json();
  checkCoordMap = await (await fetch("/api/check_coord_map")).json();
  renderCheckItems();
}
init();

// チェック項目をチェックボックスで表示
function renderCheckItems() {
  for (const key of Object.keys(checkCoordMap)) {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = key;
    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(key));
    checkContainer.appendChild(label);
    checkContainer.appendChild(document.createElement("br"));
  }
}

// 項目フォーム追加
addEntryBtn.addEventListener("click", () => {
  const div = document.createElement("div");
  
  const 部位 = document.createElement("select");
  部位.innerHTML = Object.keys(coordMap).map(k => `<option>${k}</option>`).join("");
  const 項目 = document.createElement("input");
  項目.type = "text";
  const 評価 = document.createElement("input");
  評価.type = "text";
  
  div.appendChild(部位);
  div.appendChild(項目);
  div.appendChild(評価);
  entryContainer.appendChild(div);
});

// Excel ダウンロード
downloadExcelBtn.addEventListener("click", async () => {
  const entries = [];
  
  for (const div of entryContainer.children) {
    const 部位 = div.children[0].value;
    const 項目 = div.children[1].value;
    const 評価 = div.children[2].value;
    entries.push({ 部位, 項目, 評価 });
  }

  // チェック項目も評価として追加
  for (const input of checkContainer.querySelectorAll("input[type=checkbox]")) {
    if (input.checked) entries.push({ 部位: "チェック項目", 項目: input.value, 評価: "✓" });
  }

  try {
    const response = await fetch("/api/download-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ inspector: inspectorInput.value.trim(), entries })
    });
    if (!response.ok) throw new Error(`サーバーエラー: ${response.status}`);
    
    const blob = await response.blob();
    const a = document.createElement("a");
    a.href = window.URL.createObjectURL(blob);
    a.download = "inspection.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(a.href);
  } catch (err) {
    alert(`Excel ダウンロードに失敗しました: ${err.message}`);
  }
});
