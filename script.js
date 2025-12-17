// --- Configuration ---
// Mapping Task ID กับ Row Number ใน Excel (อิงตามไฟล์ที่ส่งมา)
// Row 1 ใน ExcelJS คือ Row 1 ใน Excel จริงๆ (1-based index)
const TASKS = {
    '2.1': { row: 14, min: 40, max: 70, time: 2, default: false }, // จ่ายยา OPD
    '2.2': { row: 15, min: 40, max: 70, time: 2.5, default: false }, // Check ยา
    '2.3': { row: 16, min: 1, max: 5, time: 5, default: true },    // ประสานงาน
    '2.4': { row: 17, min: 1, max: 5, time: 15, default: true },   // ADR
    '2.5': { row: 18, min: 0, max: 4, time: 3, default: true },    // DIS
    '2.6': { row: 19, min: 0, max: 5, time: 15, default: true },   // Counsel New
    '2.7': { row: 20, min: 2, max: 6, time: 8, default: true },    // Counsel Tech
    '2.8': { row: 21, min: 2, max: 6, time: 10, default: true },   // Warfarin
    '2.9': { row: 22, min: 1, max: 3, time: 10, default: false },  // TB (Thu)
    '2.10': { row: 23, min: 1, max: 5, time: 10, default: true },  // CDCU
    '2.11': { row: 24, min: 7, max: 10, time: 10, default: false },// Asthma (Thu)
    '2.12': { row: 25, min: 16, max: 20, time: 10, default: false },// HIV (Wed)
    '2.13': { row: 26, min: 0, max: 1, time: 20, default: true },  // Tertiary
    '2.14': { row: 27, min: 0, max: 1, time: 150, default: true }, // Primary (Minimal)
    // Academic (3.x)
    '3.1': { row: 31, val: 1, time: 120 }, // SOAP
    '3.2': { row: 32, val: 1, time: 120 }, // JC
    '3.3': { row: 33, val: 1, time: 60 },  // แนะนำ
    '3.4': { row: 34, val: 1, time: 60 },  // สอน
    '3.5': { row: 35, val: 1, time: 60 }   // สรุป
};

const MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];

// --- Initialization ---
window.onload = function() {
    const today = new Date();
    const mSelect = document.getElementById('monthSelect');
    const ySelect = document.getElementById('yearSelect');

    MONTHS.forEach((m, i) => {
        let opt = document.createElement('option');
        opt.value = i;
        opt.text = m;
        if (i === today.getMonth()) opt.selected = true;
        mSelect.appendChild(opt);
    });

    let currentYear = today.getFullYear() + 543;
    for (let y = currentYear - 1; y <= currentYear + 1; y++) {
        let opt = document.createElement('option');
        opt.value = y - 543; // AD
        opt.text = y; // BE
        if (y === currentYear) opt.selected = true;
        ySelect.appendChild(opt);
    }

    mSelect.addEventListener('change', generateTable);
    ySelect.addEventListener('change', generateTable);
    generateTable();
};

// --- UI Logic ---
function generateTable() {
    const m = parseInt(document.getElementById('monthSelect').value);
    const y = parseInt(document.getElementById('yearSelect').value);
    const tbody = document.getElementById('dayTableBody');
    tbody.innerHTML = '';

    const daysInMonth = new Date(y, m + 1, 0).getDate();

    for (let d = 1; d <= daysInMonth; d++) {
        const dateObj = new Date(y, m, d);
        const dayOfWeek = dateObj.getDay(); // 0=Sun
        const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

        let row = document.createElement('tr');
        row.className = `border-b transition ${isWeekend ? 'row-disabled' : 'row-hover bg-white'}`;
        
        row.innerHTML = `
            <td class="p-4 text-center">
                <div class="font-bold text-gray-700 text-lg">${d}</div>
                <div class="text-xs text-gray-400">${['อา','จ','อ','พ','พฤ','ศ','ส'][dayOfWeek]}</div>
            </td>
            <td class="p-4">
                <select class="status-select input-pastel p-2 text-sm" onchange="toggleRow(this)">
                    <option value="work" ${!isWeekend ? 'selected' : ''}>ทำงานปกติ</option>
                    <option value="off" ${isWeekend ? 'selected' : ''}>วันหยุด</option>
                    <option value="half">ลาครึ่งวัน</option>
                </select>
            </td>
            <td class="p-4">
                <select class="location-select input-pastel p-2 text-sm">
                    <option value="normal">เวรปกติ (ห้องยา)</option>
                    <option value="floor2">เวรชั้น 2</option>
                    <option value="bigDispense">จ่ายยาใหญ่</option>
                    <option value="helper">ช่วยห้องยา</option>
                </select>
            </td>
            <td class="p-4">
                <div class="flex flex-wrap gap-3">
                    <label class="inline-flex items-center cursor-pointer bg-white border rounded-lg px-2 py-1 hover:bg-pink-50"><input type="checkbox" class="aca-soap mr-2 accent-pink-500" data-d="${d}"> <span class="text-xs">SOAP</span></label>
                    <label class="inline-flex items-center cursor-pointer bg-white border rounded-lg px-2 py-1 hover:bg-pink-50"><input type="checkbox" class="aca-jc mr-2 accent-pink-500" data-d="${d}"> <span class="text-xs">JC</span></label>
                    <label class="inline-flex items-center cursor-pointer bg-white border rounded-lg px-2 py-1 hover:bg-pink-50"><input type="checkbox" class="aca-teach mr-2 accent-pink-500" data-d="${d}"> <span class="text-xs">สอน</span></label>
                    <label class="inline-flex items-center cursor-pointer bg-white border rounded-lg px-2 py-1 hover:bg-pink-50"><input type="checkbox" class="aca-intro mr-2 accent-pink-500" data-d="${d}"> <span class="text-xs">แนะนำ</span></label>
                    <label class="inline-flex items-center cursor-pointer bg-white border rounded-lg px-2 py-1 hover:bg-pink-50"><input type="checkbox" class="aca-sum mr-2 accent-pink-500" data-d="${d}"> <span class="text-xs">สรุป</span></label>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    }
}

function toggleRow(select) {
    const row = select.closest('tr');
    const inputs = row.querySelectorAll('select:not(.status-select), input');
    if(select.value === 'off') {
        row.classList.add('row-disabled');
        row.classList.remove('bg-white', 'row-hover');
        inputs.forEach(el => el.disabled = true);
    } else {
        row.classList.remove('row-disabled');
        row.classList.add('bg-white', 'row-hover');
        inputs.forEach(el => el.disabled = false);
    }
}

// --- Calculation Core ---
async function processAndDownload() {
    const btn = document.querySelector('button[onclick="processAndDownload()"]');
    const msg = document.getElementById('statusMessage');
    
    try {
        btn.disabled = true;
        msg.classList.remove('hidden');
        msg.innerText = "กำลังอ่านไฟล์ Template...";

        // 1. Gather User Inputs
        const inputs = gatherInputs();
        
        // 2. Generate Data (Logic)
        const computedData = generateWorkData(inputs);

        // 3. Load & Modify Excel Template
        msg.innerText = "กำลังเขียนข้อมูลลง Excel...";
        const workbook = new ExcelJS.Workbook();
        
        // Fetch template.xlsx from the same directory
        const response = await fetch('./template.xlsx');
        if (!response.ok) throw new Error("ไม่พบไฟล์ template.xlsx กรุณาตรวจสอบว่ามีไฟล์นี้อยู่ในโฟลเดอร์");
        
        const buffer = await response.arrayBuffer();
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.getWorksheet(1); // Get first sheet

        // 4. Update Excel Cells
        applyDataToWorksheet(worksheet, computedData, inputs);

        // 5. Download
        msg.innerText = "กำลังดาวน์โหลด...";
        const outBuffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([outBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        saveAs(blob, `P4P_${inputs.monthName}_${inputs.yearBE}.xlsx`);

    } catch (err) {
        alert("เกิดข้อผิดพลาด: " + err.message);
        console.error(err);
    } finally {
        btn.disabled = false;
        msg.classList.add('hidden');
    }
}

function gatherInputs() {
    const m = parseInt(document.getElementById('monthSelect').value);
    const y = parseInt(document.getElementById('yearSelect').value);
    
    // Gather days
    let days = [];
    const rows = document.querySelectorAll('#dayTableBody tr');
    rows.forEach((row, idx) => {
        const day = idx + 1;
        const status = row.querySelector('.status-select').value;
        if(status === 'off') return;

        days.push({
            day: day,
            dow: new Date(y, m, day).getDay(),
            loc: row.querySelector('.location-select').value,
            status: status,
            soap: row.querySelector('.aca-soap').checked,
            jc: row.querySelector('.aca-jc').checked,
            teach: row.querySelector('.aca-teach').checked,
            intro: row.querySelector('.aca-intro').checked,
            sum: row.querySelector('.aca-sum').checked,
            data: {}
        });
    });

    let target = document.getElementById('targetScore').value;
    // Calculate simple weight: Full day = 1, Half = 0.5
    let totalWorkDays = days.reduce((acc, cur) => acc + (cur.status === 'half' ? 0.5 : 1), 0);
    target = target ? parseInt(target) : Math.floor(totalWorkDays * 95);

    return {
        userName: document.getElementById('userName').value,
        monthName: MONTHS[m],
        yearBE: y + 543,
        days: days,
        target: target
    };
}

function generateWorkData(inputs) {
    let days = inputs.days;
    
    // 1. Initial Fill & Rules
    days.forEach(d => {
        // Defaults
        for(let [tid, conf] of Object.entries(TASKS)) {
            if(conf.default) d.data[tid] = getRandom(conf.min, conf.max);
            else d.data[tid] = 0;
        }

        // Location specific
        if(d.loc === 'floor2') { d.data['2.1'] = getRandom(40,70); d.data['2.2'] = getRandom(40,70); }
        if(d.loc === 'bigDispense') { d.data['2.1'] = getRandom(40,70); }
        if(d.loc === 'helper') { d.data['2.2'] = getRandom(40,70); }

        // Day specific
        if(d.dow === 3) d.data['2.12'] = getRandom(16,20); // Wed HIV
        if(d.dow === 4) { d.data['2.9'] = getRandom(1,3); d.data['2.11'] = getRandom(7,10); } // Thu TB/Asthma

        // Academic
        if(d.soap) d.data['3.1'] = 1;
        if(d.jc) d.data['3.2'] = 1;
        if(d.teach) d.data['3.4'] = 1;
        if(d.intro) d.data['3.3'] = 1;
        if(d.sum) d.data['3.5'] = 1;

        enforceTimeLimit(d);
    });

    // 2. Balance Score
    balanceScore(days, inputs.target);
    return days;
}

function enforceTimeLimit(dayObj) {
    let limit = 420;
    let safety = 0;
    while(calcMins(dayObj.data) > limit && safety < 1000) {
        let keys = Object.keys(dayObj.data).filter(k => dayObj.data[k] > 0);
        if(keys.length) {
            let k = keys[Math.floor(Math.random() * keys.length)];
            dayObj.data[k]--;
        }
        safety++;
    }
}

function balanceScore(days, target) {
    let current = calcTotalScore(days);
    let loop = 0;
    while(Math.abs(current - target) > 15 && loop < 3000) {
        let d = days[Math.floor(Math.random() * days.length)];
        let adjustable = ['2.3', '2.4', '2.8', '2.10'];
        let k = adjustable[Math.floor(Math.random() * adjustable.length)];

        if(current < target) {
            if(calcMins(d.data) < 415) d.data[k]++;
        } else {
            if(d.data[k] > 0) d.data[k]--;
        }
        current = calcTotalScore(days);
        loop++;
    }
}

function calcMins(data) {
    return Object.entries(data).reduce((acc, [tid, val]) => acc + (val * TASKS[tid].time), 0);
}

function calcTotalScore(days) {
    let totalMins = days.reduce((acc, d) => acc + calcMins(d.data), 0);
    // Factor derived from CSV (Weight avg ~1.5 * 0.171)
    return totalMins * 0.25; 
}

function getRandom(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

function applyDataToWorksheet(ws, days, inputs) {
    // 1. Update Header Info (Example positions, adjust if needed)
    // Assuming "Month Year" is in a specific cell, e.g., A1 or merged title
    // But usually, we just fill the numbers.
    // Let's try to find the cell with "ประจำเดือน" and replace part of it if possible,
    // OR just leave the header to the user.
    // Better: Update the User Name if found.
    
    // 2. Fill Numbers
    // Excel Column E is index 5.
    days.forEach(d => {
        let colIndex = 4 + d.day; // E=5, so 4+1 = 5.
        
        for(let [tid, val] of Object.entries(d.data)) {
            if(val > 0) {
                let row = TASKS[tid].row;
                let cell = ws.getRow(row).getCell(colIndex);
                cell.value = val;
            }
        }
    });

    // 3. Update Header Name/Month (Optional - requires knowing exact cell)
    // Example: Replace cell B2 contents
    let headerCell = ws.getCell('A1'); 
    if(headerCell.value && typeof headerCell.value === 'string') {
        headerCell.value = `แบบบันทึกภาระงาน P4P รายบุคคล ประจำเดือน ${inputs.monthName} ${inputs.yearBE}`;
    }
    
    let nameCell = ws.getCell('B2'); // Adjust based on your CSV
    if(nameCell.value && typeof nameCell.value === 'string') {
        nameCell.value = `ชื่อ-นามสกุล ${inputs.userName}    ตำแหน่ง เภสัชกร`;
    }
}

