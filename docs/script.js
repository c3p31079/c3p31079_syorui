function generate() {
    document.getElementById("status").innerText = "生成中...";

    const part = document.getElementById("part").value;
    const item = document.getElementById("item").value;
    const score = document.getElementById("score").value;
    const checked = document.getElementById("checked").checked;

    fetch("https://c3p31079-syorui.onrender.com/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ part, item, score, checked })
    })
    .then(response => response.blob())
    .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "result.xlsx";
        a.click();
        document.getElementById("status").innerText = "生成完了!";
    })
    .catch(err => {
        console.error(err);
        document.getElementById("status").innerText = "エラーが発生しました";
    });
}
