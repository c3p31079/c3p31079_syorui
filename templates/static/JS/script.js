document.addEventListener('DOMContentLoaded', function() {
    console.log('CheckSheet ページ読み込み');
});

document.getElementById("downloadBtn").addEventListener("click", async () => {

    // 部位・項目・評価
    const parts = Array.from(document.querySelectorAll("table.CheckSheet_list tbody tr")).map((tr) => {
        const partName = tr.dataset.part;
        const item = tr.cells[1].textContent.trim();
        const grade = Array.from(tr.querySelectorAll("input[type=radio]")).find(r => r.checked)?.value || "";
        let symbol = "";
        if (grade === "A") symbol = "circle";
        else if (grade === "B") symbol = "triangle";
        else if (grade === "C") symbol = "cross";
        return { part: partName, item, evaluation: grade, symbol };
    });

    // チェック項目
    const checks = Array.from(document.querySelectorAll(".CheckSheet_measure_area input[type=checkbox]:checked"))
        .map(cb => {
            const label = cb.nextElementSibling?.textContent.trim() || cb.value;
            const countInput = cb.parentElement.querySelector("input[type=number]");
            return countInput ? `${label} ${countInput.value}箇所` : label;
        });

    // 自由記述欄
    const cell_updates = [
        { start: "H12", text: document.getElementById("observations")?.value || "" },
        { start: "H13", text: document.getElementById("remarks")?.value || "" },
        { start: "H14", text: document.getElementById("action_other_detail")?.value || "" },
        { start: "H15", text: document.getElementById("overall_d_detail")?.value || "" },
        { start: "H16", text: document.getElementById("plan_other_detail")?.value || "" }
    ];

    // POSTしてExcel生成
    const response = await fetch("https://c3p31079-syorui.onrender.com/api/generate-excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ parts, checks, cell_updates })
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
