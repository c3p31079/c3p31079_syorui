async function generate() {
    const inspector = document.getElementById("inspector").value;
    const status = document.getElementById("status");

    if (!inspector) {
        status.textContent = "点検者名を入力してください。";
        return;
    }

    status.textContent = "生成中...";

    const url = "https://c3p31079-syorui.onrender.com";

    try {
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ inspector })
        });

        if (!response.ok) {
            status.textContent = "エラー：ファイル生成に失敗しました。";
            return;
        }

        // バイナリデータとして受け取る
        const blob = await response.blob();

        // ダウンロード
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "generated.xlsx";
        link.click();

        status.textContent = "生成成功 ✔  ダウンロードしました。";

    } catch (error) {
        status.textContent = "通信エラー：Render の API に接続できません。";
        console.error(error);
    }
}
