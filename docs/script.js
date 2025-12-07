const API_URL = "https://c3p31079-syorui.onrender.com/api/download-excel";

let mapData = {};
let checkItems = [];

// 初期ロード
async function loadData() {
  mapData = await fetch("map.json").then(res => res.json());
  checkItems = await fetch("check_coord_map.json").then(res => res.json());

  const partSelect = document.getElementById("partSelect");
  partSelect.innerHTML = '<option value="">選択</option>';

  Object.keys(mapData).forEach(part => {
    const opt = document.createElement("option");
    opt.value = part;
    opt.textContent = part;
    partSelect.appendChild(opt);
  });

  renderChecks();
}

// 点検項目連動
document.getElementById("partSelect").onchange = e => {
  const itemSelect = document.getElementById("itemSelect");
  itemSelect.innerHTML = '<option value="">選択</option>';

  const part = e.target.value;
  if (!part) return;

  Object.keys(mapData[part]).forEach(item => {
    const opt = document.createElement("option");
    opt.value = item;
    opt.textContent = item;
    itemSelect.appendChild(opt);
  });
};

// チェック項目描画
function renderChecks() {
  const container = document.getElementById("checkContainer");
  container.innerHTML = "";

  checkItems.forEach(label => {
    const div = document.createElement("div");

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = label;

    const span = document.createElement("span");
    span.textContent = label;

    div.appendChild(cb);
    div.appendChild(span);
    container.appendChild(div);
  });
}

// Excel ダウンロード
document.getElementById("downloadBtn").onclick = async () => {
  const part = document.getElementById("partSelect").value;
  const item = document.getElementById("itemSelect").value;
  const evaluation = document.getElementById("evalSelect").value;

  const checks = Array.from(
    document.querySelectorAll("#checkContainer input:checked")
  ).map(cb => cb.value);

  const payload = {
    part,
    item,
    evaluation,
    checks
  };

  const res = await fetch(API_URL, {
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
  a.download = "inspection_result.xlsx";
  a.click();
  URL.revokeObjectURL(url);
};

loadData();
