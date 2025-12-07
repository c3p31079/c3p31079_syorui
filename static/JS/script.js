// Excel ダウンロードボタン
document.getElementById('downloadExcelBtn').addEventListener('click', function() {
    const inspectionId = 1; // 選択中の点検IDを動的に設定
    const url = `https://c3p31079-syorui.onrender.com/api/download_excel?inspection_id=${inspectionId}`;
    window.open(url, '_blank');  // 新しいタブで Excel をダウンロード
});

