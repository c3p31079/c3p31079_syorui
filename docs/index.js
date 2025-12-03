// 点検部位・項目・サブ項目の入力、チェック項目、プレビュー・Excel生成
const BACKEND_BASE = "https://c3p31079-syorui.onrender.com";

// 初期行追加
addRow();

// 「＋ 行追加」ボタン
document.getElementById("addRow").addEventListener("click", addRow);

// 「プレビューを見る」ボタン
document.getElementById("doPreview").addEventListener("click", async () => {
  const { entries, checks } = collectFormData();
  try {
    const res = await fetch(`${BACKEND_BASE}/api/sheet-image`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ entries, checks })
    });
    if (!res.ok) { alert("プレビュー生成に失敗しました"); return; }

    const blob = await res.blob();
    const reader = new FileReader();
    reader.onload = () => {
      sessionStorage.setItem("previewImg", reader.result);
      location.href = "preview.html";
    };
    reader.readAsDataURL(blob);
  } catch (err) {
    console.error(err);
    alert("プレビュー生成中にエラーが発生しました");
  }
});

// 「Excelをダウンロード」ボタン
document.getElementById("downloadExcel").addEventListener("click", async () => {
  const { entries, checks } = collectFormData();
  try {
    const res = await fetch(`${BACKEND_BASE}/api/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ entries, checks })
    });
    if (!res.ok) { alert("Excel生成に失敗しました"); return; }

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "inspection.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    console.error(err);
    alert("Excel生成中にエラーが発生しました");
  }
});

// フォーム行を追加する関数
function addRow() {
  const container = document.getElementById("form-container");
  const div = document.createElement("div");
  div.classList.add("form-row-group");

  // 点検部位セレクト
  const partSel = document.createElement("select");
  partSel.classList.add("part");
  Object.keys(itemMap).forEach(p => {
    const opt = document.createElement("option");
    opt.value = p; opt.textContent = p;
    partSel.appendChild(opt);
  });

  // 項目セレクト
  const itemSel = document.createElement("select");
  itemSel.classList.add("item");
  updateItems(partSel.value, itemSel);

  // サブ項目セレクト
  const subSel = document.createElement("select");
  subSel.classList.add("subItem");
  updateSubItems(itemSel.value, subSel);

  // 部位変更時
  partSel.addEventListener("change", () => updateItems(partSel.value, itemSel));
  // 項目変更時
  itemSel.addEventListener("change", () => updateSubItems(itemSel.value, subSel));

  div.appendChild(partSel);
  div.appendChild(itemSel);
  div.appendChild(subSel);
  container.appendChild(div);
}

// 項目セレクト更新
function updateItems(part, itemSel) {
  itemSel.innerHTML = "";
  itemMap[part].forEach(i => {
    const opt = document.createElement("option");
    opt.value = i; opt.textContent = i;
    itemSel.appendChild(opt);
  });
  updateSubItems(itemSel.value, itemSel.nextElementSibling);
}

// サブ項目セレクト更新
function updateSubItems(item, subSel) {
  subSel.innerHTML = "";
  if (subItemMap[item]) {
    subItemMap[item].forEach(s => {
      const opt = document.createElement("option");
      opt.value = s; opt.textContent = s;
      subSel.appendChild(opt);
    });
    subSel.style.display = "inline-block";
  } else {
    subSel.style.display = "none";
  }
}

// チェック項目生成
fetch(`${BACKEND_BASE}/api/coords`).then(r => r.json()).then(data => {
  const container = document.getElementById("checks-container");
  Object.keys(data).forEach(key => {
    const label = document.createElement("label");
    const cb = document.createElement("input");
    cb.type = "checkbox"; cb.value = key; cb.classList.add("check-item");
    label.appendChild(cb);
    label.appendChild(document.createTextNode(key));
    container.appendChild(label);
    container.appendChild(document.createElement("br"));
  });
});

// フォームデータ収集関数
function collectFormData() {
  const rows = Array.from(document.querySelectorAll(".form-row-group"));
  const entries = rows.map(r => {
    const part = r.querySelector(".part").value;
    const item = r.querySelector(".item").value;
    const sub = r.querySelector(".subItem") ? r.querySelector(".subItem").value : "";
    return { part, item, sub };
  });

  const checks = Array.from(document.querySelectorAll(".check-item"))
    .filter(cb => cb.checked)
    .map(cb => cb.value);

  return { entries, checks };
}
