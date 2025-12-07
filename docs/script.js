import { itemMap } from './map.json';
import { checkList } from './check_coord_map.json';

const tbody = document.getElementById('inspection-tbody');
const addRowBtn = document.getElementById('add-row');
const downloadBtn = document.getElementById('download-excel');
const checksContainer = document.getElementById('checks-container');

// 初期プルダウン設定
function populatePartSelect(select) {
    select.innerHTML = '<option value="">選択してください</option>';
    for (let part in itemMap) {
        const opt = document.createElement('option');
        opt.value = part;
        opt.textContent = part;
        select.appendChild(opt);
    }
}

// 項目プルダウン
function populateItemSelect(select, part) {
    select.innerHTML = '<option value="">選択してください</option>';
    if (part && itemMap[part]) {
        itemMap[part].forEach(item => {
            const opt = document.createElement('option');
            opt.value = item;
            opt.textContent = item;
            select.appendChild(opt);
        });
    }
}

// チェックボックス生成
function renderChecks() {
    checksContainer.innerHTML = '';
    checkList.forEach(check => {
        const label = document.createElement('label');
        label.className = 'check-item';
        const input = document.createElement('input');
        input.type = 'checkbox';
        input.value = check;
        label.appendChild(input);
        label.append(check);
        checksContainer.appendChild(label);
    });
}

// 行追加
addRowBtn.onclick = () => {
    const tr = document.createElement('tr');

    const tdPart = document.createElement('td');
    const partSelect = document.createElement('select');
    partSelect.className = 'part-select';
    populatePartSelect(partSelect);
    tdPart.appendChild(partSelect);

    const tdItem = document.createElement('td');
    const itemSelect = document.createElement('select');
    itemSelect.className = 'item-select';
    tdItem.appendChild(itemSelect);

    partSelect.onchange = () => populateItemSelect(itemSelect, partSelect.value);

    const tdMark = document.createElement('td');
    const markSelect = document.createElement('select');
    markSelect.className = 'mark-select';
    markSelect.innerHTML = `
        <option value="triangle">△</option>
        <option value="cross">×</option>
    `;
    tdMark.appendChild(markSelect);

    const tdRemove = document.createElement('td');
    const removeBtn = document.createElement('button');
    removeBtn.textContent = '-';
    removeBtn.onclick = () => tr.remove();
    tdRemove.appendChild(removeBtn);

    tr.appendChild(tdPart);
    tr.appendChild(tdItem);
    tr.appendChild(tdMark);
    tr.appendChild(tdRemove);
    tbody.appendChild(tr);
};

// 初期設定
document.querySelectorAll('.part-select').forEach(select => populatePartSelect(select));

// チェックボックス描画
renderChecks();

// ダウンロード
downloadBtn.onclick = async () => {
    const items = [];
    tbody.querySelectorAll('tr').forEach(tr => {
        const part = tr.querySelector('.part-select').value;
        const item = tr.querySelector('.item-select').value;
        const mark = tr.querySelector('.mark-select').value;
        if (part && item) items.push({ part, item, mark });
    });

    const checks = [];
    checksContainer.querySelectorAll('input[type=checkbox]:checked').forEach(chk => {
        checks.push(chk.value);
    });

    try {
        const res = await fetch('/api/download-excel', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ items, checks })
        });
        if (!res.ok) throw new Error(res.statusText);
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'inspection.xlsx';
        a.click();
        window.URL.revokeObjectURL(url);
    } catch (err) {
        alert('Excel ダウンロードに失敗しました: ' + err);
    }
};
