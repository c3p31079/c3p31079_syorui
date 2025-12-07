document.addEventListener("DOMContentLoaded", () => {
    console.log("📄 CheckSheet ページ読み込み");

    const downloadBtn = document.getElementById("downloadBtn");
    if (downloadBtn) {
        downloadBtn.addEventListener("click", async () => {
            const partsData = Array.from(document.querySelectorAll("table.CheckSheet_list tbody tr"))
                .map(row => {
                    const part = row.dataset.part;
                    const grade = row.querySelector(`input[type=radio]:checked`)?.value || "";
                    return { part, grade };
                });

            const checksData = Array.from(document.querySelectorAll(".CheckSheet_measure_area input[type=checkbox]:checked"))
                .map(el => ({ name: el.name, value: el.value }));

            // 自由記述欄
            const remarksData = Array.from(document.querySelectorAll("textarea"))
                .map(el => ({ name: el.id, value: el.value }));

            const payload = { parts: partsData, checks: checksData, remarks: remarksData };

            try {
                const response = await fetch("/api/generate-excel", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) throw new Error("Excel生成に失敗");

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = "inspection.xlsx";
                a.click();
                window.URL.revokeObjectURL(url);
            } catch (err) {
                alert(err.message);
            }
        });
    }
});
