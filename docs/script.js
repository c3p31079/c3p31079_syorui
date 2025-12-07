let coordMap = {};
let checkCoordMap = {};

async function loadJSON(url) {
    const res = await fetch(url);
    return await res.json();
}

async function init() {
    coordMap = await loadJSON("map.json");
    checkCoordMap = await loadJSON("check_coord_map.json");

    // 点検部位セレクト
    const partSelect = document.getElementById("partSelect");
    Object.keys(coordMap).forEach(part => {
        const opt = document.createElement("option");
        opt.value = part;
        opt.textContent = part;
        partSelect.appendChild(opt);
    });

    // チェック項目
    const checkContainer = document.getElementById("checkContainer");
    Object.keys(checkCoordMap).forEach(check => {
        const label = document.createElement("label");
        const input = document.createElement("input");
        input.type = "checkbox";
        input.value = check;
        label.appendChild(input);
        label.appendChild(document.createTextNode(check));
        checkContainer.appendChild(label);
    });
}

document.getElementById("downloadBtn").addEventListener("click", async () => {
    const part = document.getElementById("partSelect").value;
    const checks = Array.from(document.querySelectorAll("#checkContainer input[type=checkbox]:checked"))
                        .map(el => el.value);

    // フロントから Render の API へ
    const response = await fetch("https://c3p31079-syorui.onrender.com/api/generate-excel", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({
            parts: [
                { part: part, item: "サンプル項目", evaluation: "良" }
            ],
            checks: checks
        })
    });

    if (!response.ok) {
        alert("Excel生成に失敗しました");
        return;
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "inspection.xlsx";
    a.click();
    window.URL.revokeObjectURL(url);
});

window.addEventListener("DOMContentLoaded", init);
