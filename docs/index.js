document.addEventListener("DOMContentLoaded",()=>{
  const partSelect = document.getElementById("part");
  const itemSelect = document.getElementById("item");
  const subItemSelect = document.getElementById("subItem");

  // 点検部位プルダウン
  Object.keys(itemMap).forEach(p=>{
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;
    partSelect.appendChild(opt);
  });

  partSelect.addEventListener("change",()=>{
    const part = partSelect.value;
    itemSelect.innerHTML = "";
    (itemMap[part]||[]).forEach(it=>{
      const opt = document.createElement("option");
      opt.value = it;
      opt.textContent = it;
      itemSelect.appendChild(opt);
    });
    itemSelect.dispatchEvent(new Event("change"));
  });

  itemSelect.addEventListener("change",()=>{
    const part = partSelect.value;
    const item = itemSelect.value;
    subItemSelect.innerHTML = "";
    const subs = subItemMap[item]||[];
    subs.forEach(s=>{
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      subItemSelect.appendChild(opt);
    });
  });

  partSelect.dispatchEvent(new Event("change"));
});
