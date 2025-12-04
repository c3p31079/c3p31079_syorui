//  データ読み込み（map.json / check_coord_map.json）
let mapData = {};
let checkMapData = {};

fetch("map.json")
    .then(res => res.json())
    .then(data => {
        mapData = data;
    });

fetch("check_coord_map.json")
    .then(res => res.json())
    .then(data => {
        checkMapData = data;
        generateCheckBoxes();
    });


//  行追加：部位・項目のプルダウン
document.getElementById("addRowBtn").addEventListener("click", addRow);

function addRow() {
    const tbody = document.getElementById("itemBody");

    const tr = document.createElement("tr");

    // 点検部位
    const td1 = document.createElement("td");
    const selectPart = document.createElement("select");
    selectPart.className = "partSelect";
    selectPart.innerHTML = `<option value="">選択</option>`;
    for (const part in mapData) {
        selectPart.innerHTML += `<option value="${part}">${part}</option>`;
    }
    td1.appendChild(selectPart);

    // 項目
    const td2 = document.createElement("td");
    const selectItem = document.createElement("select");
    selectItem.className = "itemSelect";
    selectItem.innerHTML = `<option value="">選択</option>`;
    td2.appendChild(selectItem);

    // 部位変更時に項目を更新
    selectPart.addEventListener("change", () => {
        const val = selectPart.value;
        selectItem.innerHTML = `<option value="">選択</option>`;
        if (val && mapData[val]) {
            mapData[val].forEach(it => {
                selectItem.innerHTML += `<option value="${it}">${it}</option>`;
            });
        }
    });

    // 記号プルダウン
    const td3 = document.createElement("td");
    const selectSymbol = document.createElement("select");
    selectSymbol.className = "symbolSelect";
    selectSymbol.innerHTML = `
        <option value="">選択</option>
        <option value="triangle">△</option>
        <option value="cross">×</option>
        <option value="circle">○</option>
    `;
    td3.appendChild(selectSymbol);

    // 削除
    const td4 = document.createElement("td");
    const del = document.createElement("button");
    del.textContent = "削除";
    del.onclick = () => tr.remove();
    td4.appendChild(del);

    tr.appendChild(td1);
    tr.appendChild(td2);
    tr.appendChild(td3);
    tr.appendChild(td4);

    tbody.appendChild(tr);
}

// チェックボックス生成
function generateCheckBoxes() {
    const area = document.getElementById("checkboxArea");
    area.innerHTML = "";

    for (const key in checkMapData) {
        const div = document.createElement("div");
        div.className = "checkbox-item";

        div.innerHTML = `
            <label>
                <input type="checkbox" class="checkItem" value="${key}">
                ${key}
            </label>
        `;

        area.appendChild(div);
    }
}


// Excel ダウンロード（POST）
document.getElementById("downloadExcelBtn").addEventListener("click", downloadExcel);

function downloadExcel() {
    const status = document.getElementById("status");
    status.textContent = "生成中...（数秒かかる場合があります）";

    // 行データ取得
    const rows = [];
    document.querySelectorAll("#itemBody tr").forEach(tr => {
        const part = tr.querySelector(".partSelect").value;
        const item = tr.querySelector(".itemSelect").value;
        const symbol = tr.querySelector(".symbolSelect").value;
        if (part && item && symbol) {
            rows.push({ part, item, symbol });
        }
    });

    // チェックボックス取得
    const checks = [];
    document.querySelectorAll(".checkItem:checked").forEach(c => checks.push(c.value));

    const payload = {
        rows: rows,
        checks: checks
    };

    fetch("/api/generate-excel", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify(payload)
    })
        .then(res => {
            if (!res.ok) throw new Error("サーバーエラー");
            return res.blob();
        })
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = "inspection.xlsx";
            a.click();
            status.textContent = "Excelダウンロード成功！";
        })
        .catch(err => {
            status.textContent = "Excel ダウンロード失敗… " + err;
        });
}
