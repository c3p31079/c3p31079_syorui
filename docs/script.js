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

    // 部位プルダウン生成
    const partSelect = document.getElementById("partSelect");
    Object.keys(coordMap).forEach(part => {
        const option = document.createElement("option");
        option.value = part;
        option.text = part;
        partSelect.appendChild(option);
    });

    // チェック項目生成
    const checkContainer = document.getElementById("checkContainer");
    Object.keys(checkCoordMap).forEach(key => {
        const label = document.createElement("label");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = key;
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(key));
        checkContainer.appendChild(label);
    });
}

document.getElementById("downloadBtn").addEventListener("click", async () => {
    const part = document.getElementById("partSelect").value;
    const checks = Array.from(document.querySelectorAll("#checkContainer input[type=checkbox]:checked"))
                        .map(el => el.value);

    try {
        const response = await fetch("https://https://c3p31079-syorui.onrender.com/api/generate-excel", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify({parts: part, checks: checks})
        });

        if (!response.ok) {
            const err = await response.json();
            alert("Excel生成に失敗しました\n" + (err.error || ""));
            return;
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "inspection.xlsx";
        a.click();
        window.URL.revokeObjectURL(url);
    } catch (e) {
        alert("Excel生成エラー:\n" + e);
    }
});

window.onload = init;
