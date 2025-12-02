const generateBtn = document.getElementById("generateBtn");
const statusEl = document.getElementById("status");

generateBtn.addEventListener("click", async () => {
    generateBtn.disabled = true;
    statusEl.textContent = "生成中…";
    statusEl.className = "";

    const payload = {
        part: document.getElementById("part").value,
        item: document.getElementById("item").value,
        score: Number(document.getElementById("score").value),
        check: document.getElementById("check").checked
    };

    try {
        const response = await fetch("https://c3p31079-syorui.onrender.com/generate", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (!response.ok) throw new Error("Excel 生成失敗");

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = url;
        a.download = "inspection.xlsx";
        a.click();

        statusEl.textContent = "生成完了！";
        statusEl.className = "success";

    } catch (err) {
        statusEl.textContent = "エラー: " + err.message;
        statusEl.className = "error";

    } finally {
        generateBtn.disabled = false;
    }
});
