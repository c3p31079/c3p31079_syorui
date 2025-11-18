const statusEl = document.getElementById("status");
const btn = document.getElementById("generateBtn");

btn.addEventListener("click", async () => {
  const inspectorName = document.getElementById("inspector").value.trim();

  if (!inspectorName) {
    alert("点検者名を入力してください！");
    return;
  }

  // 生成中の表示
  statusEl.textContent = "Excelを生成中…";
  btn.disabled = true;

  try {
    const response = await fetch("https://c3p31079-syorui.onrender.com/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ inspector: inspectorName })
    });

    if (!response.ok) {
      throw new Error(`サーバーエラー: ${response.status}`);
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "generated.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    statusEl.textContent = "Excelをダウンロードしました！";
  } catch (err) {
    console.error(err);
    statusEl.textContent = `ファイル生成に失敗しました: ${err.message}`;
  } finally {
    btn.disabled = false;
  }
});
