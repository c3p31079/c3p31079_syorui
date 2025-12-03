// sessionStorage に保存したプレビュー画像を表示
window.addEventListener("DOMContentLoaded", () => {
  const container = document.getElementById("preview-container");
  const src = sessionStorage.getItem("previewImg");

  const status = document.createElement("p");
  status.textContent = "プレビューを表示中…";
  container.appendChild(status);

  if (src) {
    const img = document.createElement("img");
    img.src = src;
    img.style.width = "100%";
    container.appendChild(img);
    status.textContent = "プレビュー表示完了";
  } else {
    status.textContent = "プレビュー画像がありません";
  }
});
