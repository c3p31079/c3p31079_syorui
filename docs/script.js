document.getElementById("generateBtn").addEventListener("click", async () => {
  const inspectorName = document.getElementById("inspector").value.trim();

  if (!inspectorName) {
    alert("点検者名を入力してください！");
    return;
  }

  try {
    const response = await fetch("https://c3p31079-syorui.onrender.com/generate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ inspector: inspectorName })
    });

    if (!response.ok) {
      throw new Error(`サーバーエラー: ${response.status}`);
    }

    // Blobとして受け取る
    const blob = await response.blob();

    // ダウンロードリンク作成
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "generated.xlsx"; // ファイル名
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    alert("Excelをダウンロードしました！");
  } catch (err) {
    console.error(err);
    alert(`ファイル生成に失敗しました: ${err.message}`);
  }
});
