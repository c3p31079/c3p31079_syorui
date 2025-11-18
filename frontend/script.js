async function generate() {
  const data = {
    inspector: document.getElementById("inspector").value,
    date: document.getElementById("date").value,
    target: document.getElementById("target").value,
    score: document.getElementById("score").value
  };

  // ←← デプロイ後のあなたの Render API URL に書き換えて！
  const apiUrl = "https://your-backend.onrender.com/generate";

  const res = await fetch(apiUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data)
  });

  const blob = await res.blob();
  const url = window.URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "output.xlsx";
  a.click();
}
