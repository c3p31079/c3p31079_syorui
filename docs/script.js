let mapData = {};
let checkMapData = {};

async function loadJSON() {
    mapData = await fetch("map.json").then(res => res.json());
    checkMapData = await fetch("check_coord_map.json").then(res => res.json());
}

document.addEventListener("DOMContentLoaded", async () => {
    await loadJSON();

    const checkContainer = document.getElementById("checkContainer");
    checkContainer.innerHTML = "";
    Object.keys(checkMapData).forEach((key, idx) => {
        const label = document.createElement("label");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = key;
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(key));
        checkContainer.appendChild(label);
        checkContainer.appendChild(document.createElement("br"));
    });

    document.getElementById("downloadBtn").addEventListener("click", async () => {
        const part = document.getElementById("partSelect").value;
        const checks = Array.from(checkContainer.querySelectorAll("input[type=checkbox]:checked"))
                            .map(el => el.value);

        try {
            const response = await fetch("/api/generate-excel", {
                method: "POST",
                headers: {"Content-Type": "application/json"},
                body: JSON.stringify({parts: part, checks: checks})
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
        } catch (err) {
            alert("通信に失敗しました: " + err);
        }
    });
});
