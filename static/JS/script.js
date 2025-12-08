// ============================
// Excel ダウンロード処理
// ============================

document.getElementById("downloadExcelBtn").addEventListener("click", async function () {

    // ============================
    // HTML から値を取得
    // ============================

    const data = {
        search_park: document.getElementById("search_park")?.value || "",
        inspection_year: document.getElementById("inspection_year")?.value || "",
        install_year_num: document.getElementById("install_year_num")?.value || "",

        // Excel に反映する項目（仮：既存ロジック維持）
        items: window.excelItems = [
            // ==============================
            // 実施措置（F6:G9）
            // ==============================
            {
                "type": "check",
                "cell": "F6",
                "dx": 2,
                "dy": 4,
                "icon": "check.png"
            },
            {
                "type": "text",
                "cell": "F8",
                "dx": 35,
                "dy": 0,
                "text": "2"   // 吊金具交換 箇所数
            },

            // ==============================
            // 所見（F10:G12）
            // ==============================
            {
                "type": "text",
                "cell": "F10",
                "dx": 4,
                "dy": 18,
                "text": "吊金具に摩耗が見られる"
            },

            // ==============================
            // 総合結果（F13:G15）
            // ==============================
            {
                "type": "circle",
                "cell": "F14",
                "dx": 2,
                "dy": 3,
                "icon": "circle.png"   // B:経過観察
            },
            {
                "type": "text",
                "cell": "F15",
                "dx": 22,
                "dy": 0,
                "text": "落下防止のため使用注意"
            },

            // ==============================
            // 対応方針（H6:H10）
            // ==============================
            {
                "type": "check",
                "cell": "H7",
                "dx": 1,
                "dy": 3,
                "icon": "check.png"
            },
            {
                "type": "text",
                "cell": "H10",
                "dx": 14,
                "dy": 0,
                "text": "部品調達後対応"
            },

            // ==============================
            // 対応予定時期（H10）
            // ==============================
            {
                "type": "text",
                "cell": "H11",
                "dx": 8,
                "dy": 0,
                "text": "6"
            },
            {
                "type": "circle",
                "cell": "H11",
                "dx": 30,
                "dy": 3,
                "icon": "circle.png"   // 上旬
            },

            // ==============================
            // 本格的使用禁止（H11）
            // ==============================
            {
                "type": "circle",
                "cell": "H11",
                "dx": 55,
                "dy": 3,
                "icon": "circle.png"   // 実施予定
            },

            // ==============================
            // 備考（H12:H15）
            // ==============================
            {
                "type": "text",
                "cell": "H12",
                "dx": 2,
                "dy": 18,
                "text": "次回点検時に重点確認"
            }
        ]
    };

    // ============================
    // Flask API に POST
    // ============================

    try {
        const response = await fetch("http://127.0.0.1:5000/api/generate_excel", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(data)
        });

        if (!response.ok) {
            throw new Error("Excel生成に失敗しました");
        }

        // ============================
        // Blob としてダウンロード
        // ============================

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = url;
        a.download = "点検チェックシート.xlsx";
        document.body.appendChild(a);
        a.click();

        a.remove();
        window.URL.revokeObjectURL(url);

    } catch (error) {
        alert(error.message);
        console.error(error);
    }
});
