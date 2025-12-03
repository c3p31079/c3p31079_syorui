// フォーム行を追加し、入力データをバックエンドへ送信
const rows = document.getElementById("rows");
const addBtn = document.getElementById("addRowBtn");
const genBtn = document.getElementById("generateBtn");
const statusEl = document.getElementById("status");

function addRow() {
  const tr = document.createElement("tr");

  tr.innerHTML = `
    <td><input type="text" class="part"></td>
    <td><input type="text" class="category"></td>
    <td>
      <select class="mark">
        <option>△</option>
        <option>×</option>
        <option>○</option>
        <option>✓</option>
      </select>
    </td>
    <td>
      <label><input type="checkbox" value="整備班で対応予定">整備班で対応予定</label>
      <label><input type="checkbox" value="修繕・修繕工事で対応予定">修繕工事</label>
      <label><input type="checkbox" value="本格的な使用禁止措置">使用禁止</label>
    </td>
    <td><button onclick="this.parentNode.parentNode.remove()">－</button></td>
  `;

  rows.appendChild(tr);
}

addBtn.onclick = addRow;
addRow(); // 初期行


genBtn.onclick = async () => {

  const items = [];
  for (const tr of rows.children) {
    const part = tr.querySelector(".part").value;
    const category = tr.querySelector(".category").value;
    const mark = tr.querySelector(".mark").value;

    const checks = [...tr.querySelectorAll("input[type=checkbox]:checked")]
      .map(ch => ch.value);

    items.push({ part, category, mark, checks });
  }

  statusEl.textContent = "プレビュー画像を生成中…";

  const res = await fetch("https://c3p31079-syorui.onrender.com/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ items })
  });

  const json = await res.json();
  window.location.href = "preview.html?img=" + encodeURIComponent(json.preview);
};
