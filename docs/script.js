const coordMapUrl = "./map.json";
const checkCoordMapUrl = "./check_coord_map.json";
let coordMap = {};
let checkCoordMap = {};

async function loadJSON(url) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`${url} が取得できません`);
    return res.json();
}

async function init() {
    coordMap = await loadJSON(coordMapUrl);
    checkCoordMap = await loadJSON(checkCoordMapUrl);

    // 部位選択
    const partSelect = document.getElementById("partSelect");
    Object.keys(coordMap).forEach(part => {
        const option = document.createElement("option");
        option.value = part;
        option.text = part;
        partSelect.appendChild(option);
    });

    // チェック項目表示
    const checkContainer = document.getElementById("checkContainer");
    checkCoordMap.forEach(key => {
        const label = document.createElement("label");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = key;
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(key));
        checkContainer.appendChild(label);
    });
}

document.getElementById("downloadBtn").onclick = async () => {
    const selectedPart = document.getElementById("partSelect").value;
    const selectedChecks = Array.from(document.querySelectorAll("#checkContainer input:checked"))
                                .map(c => c.value);
    const payload = { parts: selectedPart, items: [], checks: selectedChecks };

    try {
        const res = await fetch("https://YOUR_FLASK_SERVER/api/download-excel", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(await res.text());
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "inspection.xlsx";
        a.click();
    } catch (e) {
        alert("Excel ダウンロードに失敗しました:\n" + e);
    }
};

window.onload = init;
