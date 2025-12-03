function getParam(name) {
  return new URLSearchParams(window.location.search).get(name);
}

const imgPath = getParam("img");

const img = document.getElementById("previewImg");
const loading = document.getElementById("loading");

img.src = "https://c3p31079-syorui.onrender.com" + imgPath;

img.onload = () => {
  loading.style.display = "none";
  img.style.display = "block";
};

function downloadExcel() {
  window.location.href = "https://c3p31079-syorui.onrender.com/api/download";
}
