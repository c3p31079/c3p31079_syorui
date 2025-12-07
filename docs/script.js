document.addEventListener("DOMContentLoaded", async () => {
  const inspectionTable = document.querySelector("#inspectionTable tbody");
  const addRowBtn = document.getElementById("addRow");
  const downloadBtn = document.getElementById("downloadBtn");
  const checkItemsDiv = document.getElementById("checkItems");

  // JSON を fetch で読み込み
  const itemMapRes = await fetch("map.json");
  const itemMap = await itemMapRes.json();

  const checkCoordRes = await fetch("check_coord_map.json");
  const checkCoordMap = await checkCoordRes.json();
  const checkItems = Object.keys(checkCoordMap);

  const partOptions = Object.keys(itemMap);

  function createPartSelect() {
    const select = document.createElement("select");
    partOptions.forEach(part => {
      const option = document.createElement("option");
      option.value = part;
      option.textContent = part;
      select.appendChild(option);
    });
    return select;
  }

  function createItemSelect(part) {
    const select = document.createElement("select");
    (itemMap[part] || []).forEach(item => {
      const option = document.createElement("option");
      option.value = item;
      option.textContent = item;
      select.appendChild(option);
    });
    return select;
  }

  function createMarkSelect() {
    const select = document.createElement("select");
    ["triangle", "cross"].forEach(mark => {
      const option = document.createElement("option");
      option.value = mark;
      option.textContent = mark;
      select.appendChild(option);
    });
    return select;
  }

  function addRow(partVal = partOptions[0], itemVal = itemMap[partOptions[0]][0], markVal="triangle") {
    const tr = document.createElement("tr");

    const tdPart = document.createElement("td");
    const partSelect = createPartSelect();
    partSelect.value = partVal;
    tdPart.appendChild(partSelect);

    const tdItem = document.createElement("td");
    const itemSelect = createItemSelect(partVal);
    itemSelect.value = itemVal;
    tdItem.appendChild(itemSelect);

    const tdMark = document.createElement("td");
    const markSelect = createMarkSelect();
    markSelect.value = markVal;
    tdMark.appendChild(markSelect);

    const tdRemove = document.createElement("td");
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "削除";
    removeBtn.onclick = () => tr.remove();
    tdRemove.appendChild(removeBtn);

    tr.appendChild(tdPart);
    tr.appendChild(tdItem);
    tr.appendChild(tdMark);
    tr.appendChild(tdRemove);
    inspectionTable.appendChild(tr);

    // 部位変更で項目更新
    partSelect.addEventListener("change", () => {
      const newItemSelect = createItemSelect(partSelect.value);
      tdItem.replaceChild(newItemSelect, itemSelect);
    });
  }

  function createCheckItems() {
    checkItemsDiv.innerHTML = "";
    checkItems.forEach(item => {
      const label = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = item;
      label.appendChild(checkbox);
      label.appendChild(document.createTextNode(item));
      checkItemsDiv.appendChild(label);
    });
  }

  // 行追加ボタン
  addRowBtn.addEventListener("click", () => addRow());

  // 初期化
  addRow();
  createCheckItems();

  // ダウンロード
  downloadBtn.addEventListener("click", async () => {
    const rows = Array.from(inspectionTable.querySelectorAll("tr"));
    const items = rows.map(tr => {
      const selects = tr.querySelectorAll("select");
      return {
        part: selects[0].value,
        item: selects[1].value,
        mark: selects[2].value
      };
    });

    const checks = Array.from(checkItemsDiv.querySelectorAll("input[type=checkbox]:checked")).map(cb => cb.value);

    const data = { items, checks };

    try {
      const res = await fetch("/api/download-excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data)
      });
      if (!res.ok) throw new Error(res.status);
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
      alert("Excel ダウンロードに失敗しました: " + err);
    }
  });

});
