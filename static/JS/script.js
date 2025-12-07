document.getElementById("downloadBtn").addEventListener("click", async () => {
    const parts = [];  // 部位・項目・評価
    const rows = document.querySelectorAll("tbody tr");
    rows.forEach(row => {
        const partName = row.dataset.part;
        const grade = row.querySelector(`input[type=radio]:checked`)?.value || "";
        const comment = "";  // 自由記述は後で取得
        const item = row.children[1]?.textContent || "";
        parts.push({part_name: partName, item: item, grade: grade, comment: comment});
    });

    const checks = Array.from(document.querySelectorAll("input[type=checkbox]:checked"))
                        .map(el => {
                            return {row: el.dataset.row || 2, col: el.dataset.col || 5};
                        });

    const remarks = document.getElementById("remarks").value;

    const response = await fetch("/api/generate-excel", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({parts: parts, checks: checks, remarks: remarks})
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
