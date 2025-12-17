let dataList = [];
const months = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];
const currentYear = new Date().getFullYear();

window.onload = () => {
    initAllSelectors();
    setDefaultDate();
    setupTimeInputs();
};

function setupTimeInputs() {
    const formatTime = (e) => {
        let val = e.target.value.replace(/\D/g, '');
        if (val.length > 4) val = val.slice(0, 4);
        if (val.length >= 3) {
            val = val.slice(0, 2) + ':' + val.slice(2);
        }
        e.target.value = val;
    };
    document.getElementById('timeStart').addEventListener('input', formatTime);
    document.getElementById('timeEnd').addEventListener('input', formatTime);
}

function setDefaultDate() {
    const today = new Date();
    document.getElementById('entryDay').value = today.getDate();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    document.getElementById('entryMonth').value = m;
    document.getElementById('entryYear').value = today.getFullYear();
}

// --- SEÇİCİLERİ DOLDURMA ---
function initAllSelectors() {
    const fillRange = (id, start, end, def) => {
        const sel = document.getElementById(id);
        sel.innerHTML = '';
        for (let i = start; i <= end; i++) {
            const opt = document.createElement('option');
            opt.value = i;
            opt.innerText = i;
            if (i === def) opt.selected = true;
            sel.appendChild(opt);
        }
    };
    const fillMonthsName = (id, def) => {
        const sel = document.getElementById(id);
        months.forEach(m => {
            const opt = document.createElement('option');
            opt.value = m;
            opt.innerText = m;
            if (m === def) opt.selected = true;
            sel.appendChild(opt);
        });
    };
    const fillMonthsNum = (id) => {
        const sel = document.getElementById(id);
        months.forEach((m, idx) => {
            const val = String(idx + 1).padStart(2, '0');
            const opt = document.createElement('option');
            opt.value = val;
            opt.innerText = `${val} - ${m}`;
            sel.appendChild(opt);
        });
    };
    const fillYears = (id) => {
        const sel = document.getElementById(id);
        for (let i = currentYear - 1; i <= currentYear + 2; i++) {
            const opt = document.createElement('option');
            opt.value = i;
            opt.innerText = i;
            if (i === currentYear) opt.selected = true;
            sel.appendChild(opt);
        }
    };

    fillRange('startDay', 1, 31, 20);
    fillMonthsName('startMonth', 'Kasım');
    fillRange('endDay', 1, 31, 20);
    fillMonthsName('endMonth', 'Aralık');

    fillRange('entryDay', 1, 31, 1);
    fillMonthsNum('entryMonth');
    fillYears('entryYear');
}

// --- DOSYA YÜKLEME ---
document.getElementById('uploadExcel').addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
        dataList = [];
        let readMode = false;

        json.forEach(row => {
            const c0 = (row[0] || '').toString();

            if (c0.includes("Dönem")) {
                const txt = row[1] || "";
                try {
                    const parts = txt.split(' - ');
                    if (parts.length === 2) {
                        const s = parts[0].trim().split(' ');
                        if (s.length >= 2) { document.getElementById('startDay').value = s[0]; document.getElementById('startMonth').value = s[1]; }
                        const e = parts[1].trim().split(' ');
                        if (e.length >= 2) { document.getElementById('endDay').value = e[0]; document.getElementById('endMonth').value = e[1]; }
                    }
                } catch (e) { }
            }

            if (c0.includes("Adı Soyadı")) document.getElementById('studentName').value = row[1] || "";
            if (c0.includes("Tarih")) { readMode = true; return; }
            if (c0.includes("TOPLAM")) { readMode = false; return; }

            if (readMode && row[0]) {
                let d = row[0];
                if (typeof d === 'number') {
                    const dojb = XLSX.SSF.parse_date_code(d);
                    d = `${String(dojb.d).padStart(2, '0')}/${String(dojb.m).padStart(2, '0')}/${dojb.y}`;
                } else if (typeof d === 'string') {
                    d = d.replace(/\./g, '/').replace(/-/g, '/');
                }
                let t = (row[1] || "").toString();
                dataList.push({ date: d, time: t, desc: row[3] || "" });
            }
        });
        sortDataByDate();
    };
    reader.readAsArrayBuffer(file);
});

// --- EKLEME / GÜNCELLEME ---
document.getElementById('entryForm').addEventListener('submit', (e) => {
    e.preventDefault();

    const d = String(document.getElementById('entryDay').value).padStart(2, '0');
    const m = document.getElementById('entryMonth').value;
    const y = document.getElementById('entryYear').value;
    const dateF = `${d}/${m}/${y}`; // Kesinlikle DD/MM/YYYY

    const timeF = `${document.getElementById('timeStart').value} - ${document.getElementById('timeEnd').value}`;

    // TEK METİN GİRİŞİ
    const descF = document.getElementById('descIn').value;

    const item = { date: dateF, time: timeF, desc: descF };
    const idx = parseInt(document.getElementById('editIndex').value);
    if (idx > -1) {
        dataList[idx] = item;
        renderTable();
    } else {
        dataList.push(item);
        sortDataByDate();
    }

    resetForm();
});

// --- SIRALAMA ---
function sortDataByDate() {
    dataList.sort((a, b) => {
        const [d1, m1, y1] = a.date.split('/').map(Number);
        const [d2, m2, y2] = b.date.split('/').map(Number);

        if (y1 !== y2) return y1 - y2;
        if (m1 !== m2) return m1 - m2;
        if (d1 !== d2) return d1 - d2;

        try {
            const getMin = (t) => {
                let clean = t.split('-')[0].replace(/\s/g, '').replace(/\./g, ':');
                let [h, m] = clean.split(':').map(Number);
                return (h * 60) + (m || 0);
            }
            return getMin(a.time) - getMin(b.time);
        } catch (e) { return 0; }
    });
    renderTable();
}

function moveRow(index, direction) {
    if (direction === 'up') {
        if (index === 0) return;
        [dataList[index], dataList[index - 1]] = [dataList[index - 1], dataList[index]];
    } else {
        if (index === dataList.length - 1) return;
        [dataList[index], dataList[index + 1]] = [dataList[index + 1], dataList[index]];
    }
    renderTable();
}

function renderTable() {
    const tbody = document.querySelector('#mainTable tbody');
    tbody.innerHTML = '';
    let totalMin = 0;

    dataList.forEach((item, i) => {
        try {
            let clean = item.time.toString().replace(/\s/g, '').replace(/\./g, ':');
            let parts = clean.split('-');
            if (parts.length === 2) {
                const [h1, m1] = parts[0].split(':').map(Number);
                const [h2, m2] = parts[1].split(':').map(Number);
                if (!isNaN(h1) && !isNaN(h2)) {
                    let diff = (h2 * 60 + (m2 || 0)) - (h1 * 60 + (m1 || 0));
                    if (diff < 0) diff += 1440;
                    totalMin += diff;
                }
            }
        } catch (e) { }

        tbody.innerHTML += `
                <tr>
                    <td class="col-a">${item.date}</td>
                    <td class="col-b">${item.time}</td>
                    <td class="col-c"></td>
                    <td class="col-d">${item.desc}</td>
                    <td style="text-align:center; border:none; background:#fff; white-space:nowrap;">
                        <button class="btn btn-outline-secondary btn-move" onclick="moveRow(${i}, 'up')" ${i === 0 ? 'disabled' : ''}>⬆️</button>
                        <button class="btn btn-outline-secondary btn-move" onclick="moveRow(${i}, 'down')" ${i === dataList.length - 1 ? 'disabled' : ''}>⬇️</button>
                        <button class="btn btn-sm btn-outline-warning ms-1" onclick="editItem(${i})">✏️</button>
                        <button class="btn btn-sm btn-outline-danger" onclick="deleteItem(${i})">🗑️</button>
                    </td>
                </tr>`;
    });
    document.getElementById('totalDisplay').innerText = (totalMin / 60).toLocaleString('tr-TR', { minimumFractionDigits: 1 });
}

function editItem(i) {
    const item = dataList[i];
    try {
        const parts = item.date.split('/');
        if (parts.length === 3) {
            document.getElementById('entryDay').value = parseInt(parts[0]);
            document.getElementById('entryMonth').value = parts[1];
            document.getElementById('entryYear').value = parts[2];
        }
    } catch (e) { }

    try {
        const clean = item.time.toString().replace(/\s/g, '').replace(/\./g, ':');
        const tParts = clean.split('-');
        if (tParts.length === 2) {
            const formatH = (t) => {
                let [h, m] = t.split(':');
                return `${h.padStart(2, '0')}:${(m || '00').padStart(2, '0')}`;
            }
            document.getElementById('timeStart').value = formatH(tParts[0]);
            document.getElementById('timeEnd').value = formatH(tParts[1]);
        }
    } catch (e) { }

    // Sadece metni koy
    document.getElementById('descIn').value = item.desc;
    document.getElementById('editIndex').value = i;
    document.getElementById('addBtn').innerText = "GÜNCELLE";
    document.getElementById('addBtn').className = "btn btn-warning btn-action w-100";
    document.getElementById('cancelBtn').style.display = 'block';

    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function deleteItem(i) {
    if (confirm("Silinsin mi?")) { dataList.splice(i, 1); renderTable(); }
}

function resetForm() {
    document.getElementById('entryForm').reset();
    document.getElementById('editIndex').value = -1;
    document.getElementById('addBtn').innerText = "Listeye Ekle";
    document.getElementById('addBtn').className = "btn btn-dark btn-action w-100";
    document.getElementById('cancelBtn').style.display = 'none';
    setDefaultDate();
}

// --- EXCEL EXPORT ---
async function downloadExactExcel() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Calisma Raporu');

    ws.getColumn(1).width = 24.9;
    ws.getColumn(2).width = 20.5;
    ws.getColumn(3).width = 12.8;
    ws.getColumn(4).width = 29.2;

    const baseStyle = {
        font: { name: 'Calibri', size: 12, bold: true },
        border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } },
        alignment: { vertical: 'middle', wrapText: true }
    };

    const ROW_HEIGHT = 19.5;
    ws.addRow([]);

    const pTxt = `${document.getElementById('startDay').value} ${document.getElementById('startMonth').value} - ${document.getElementById('endDay').value} ${document.getElementById('endMonth').value}`;

    const r2 = ws.addRow(['Birim:', 'Etkinlik Koordinatörlüğü', '', '']);
    r2.height = ROW_HEIGHT;
    ws.mergeCells('B2:D2');
    r2.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
    r2.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };

    const r3 = ws.addRow(['Çalışma Dönemi', pTxt, '', '']);
    r3.height = ROW_HEIGHT;
    ws.mergeCells('B3:D3');
    r3.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
    r3.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };

    const r4 = ws.addRow(['Öğrencinin Adı Soyadı:', document.getElementById('studentName').value, '', '']);
    r4.height = ROW_HEIGHT;
    ws.mergeCells('B4:D4');
    r4.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
    r4.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };

    const r5 = ws.addRow(['Tarih', 'Çalıştığı Saat Aralığı', 'İmza', 'Yapılan İşin Tanımı']);
    r5.height = ROW_HEIGHT;
    r5.eachCell(cell => cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true });

    dataList.forEach(item => {
        const r = ws.addRow([item.date, item.time, '', item.desc]);
        r.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
        r.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };
        r.getCell(3).alignment = { vertical: 'middle', horizontal: 'center' };
        r.getCell(4).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    });

    const totalVal = document.getElementById('totalDisplay').innerText;
    const lastRowNum = ws.rowCount + 1;
    const lastRow = ws.addRow(['TOPLAM ÇALIŞMA SAATİ', totalVal, '', '']);
    lastRow.height = ROW_HEIGHT;

    ws.mergeCells(`B${lastRowNum}:D${lastRowNum}`);

    lastRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'right' };
    lastRow.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };

    for (let i = 2; i <= ws.rowCount; i++) {
        const row = ws.getRow(i);
        for (let j = 1; j <= 4; j++) {
            const cell = row.getCell(j);
            const currentAlign = cell.alignment || baseStyle.alignment;
            cell.style = { ...baseStyle, alignment: currentAlign };
        }
    }

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'Rapor_Final.xlsx');

    setTimeout(() => {
        alert("✅ Rapor başarıyla indirildi!\n\nLütfen oluşturulan Excel dosyasını açıp içeriğini kontrol etmeyi unutmayın.");
    }, 500);
}