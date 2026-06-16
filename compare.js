// === МОДУЛЬ КОМБІНОВАНОЇ ЗВІРКИ ТА КРОС-АНАЛІТИКИ (compare.js) ===
let lastCompareData = null; // Глобальне сховище для експорту в Excel

document.addEventListener('DOMContentLoaded', () => {
    const tabCompare = document.getElementById('tabCompare');
    const containerCompare = document.getElementById('tableContainerCompare');
    const compareStartDateInput = document.getElementById('compareStartDate');
    const compareEndDateInput = document.getElementById('compareEndDate');

    if (!tabCompare || !containerCompare) return;

    // Установка сьогоднішньої дати за замовчуванням
    const today = new Date();
    compareStartDateInput.value = today.toISOString().split('T')[0];
    compareEndDateInput.value = today.toISOString().split('T')[0];

    tabCompare.addEventListener('click', () => {
        switchTab(tabCompare, containerCompare);
        
        const role = sessionStorage.getItem('kamagonAuthRole');
        const select = document.getElementById('compareYardSelect');
        
        if (role === 'Адмін' && select && select.options.length === 0) {
            const uniqueYards = Object.keys(yardDictionary).map(k => yardDictionary[k].yard).filter((v, i, a) => v && a.indexOf(v) === i).sort();
            uniqueYards.forEach(y => {
                const opt = document.createElement('option');
                opt.value = opt.textContent = y;
                select.appendChild(opt);
            });
        }
        
        const userYard = sessionStorage.getItem('kamagonAuthYard');
        if (role === 'РДУ' && select) {
            select.value = userYard;
            select.disabled = true;
        }
    });

    document.getElementById('loadCompareDataBtn').addEventListener('click', loadUnifiedCompareData);
});

function clearCompareDashboard() {
    document.getElementById('compareTableWrapper').innerHTML = "<p class='disabled'>Натисніть 'Завантажити дані' для побудови звітів.</p>";
    if (window.myCombinedCompareCharts) {
        window.myCombinedCompareCharts.forEach(c => c.destroy());
        window.myCombinedCompareCharts = null;
    }
}

async function loadUnifiedCompareData() {
    const role = sessionStorage.getItem('kamagonAuthRole');
    const select = document.getElementById('compareYardSelect');
    const yard = role === 'Адмін' ? (select ? select.value : "") : sessionStorage.getItem('kamagonAuthYard');
    
    const startVal = document.getElementById('compareStartDate').value;
    const endVal = document.getElementById('compareEndDate').value;

    if (!startVal || !endVal || !yard) return alert("Оберіть період дат та автодвір!");

    let datesList = [];
    let curr = new Date(startVal);
    let endD = new Date(endVal);
    
    if (curr > endD) return alert("Дата початку не може бути більшою за дату кінця!");

    while (curr <= endD) {
        let dd = String(curr.getDate()).padStart(2, '0');
        let mm = String(curr.getMonth() + 1).padStart(2, '0');
        let yyyy = curr.getFullYear();
        datesList.push(`${dd}.${mm}.${yyyy}`);
        curr.setDate(curr.getDate() + 1);
    }

    const btn = document.getElementById('loadCompareDataBtn');
    btn.innerText = "⏳ Збір даних..."; btn.disabled = true;

    try {
        const [resPlan, resAutoFact, resRdu] = await Promise.all([
            fetch(`${RESULTS_SCRIPT_URL}?action=getAggregatedData&yard=${encodeURIComponent(yard)}`).then(r => r.json()),
            fetch(`${RESULTS_SCRIPT_URL}?action=getFactAggregatedData&yard=${encodeURIComponent(yard)}`).then(r => r.json()),
            fetch(`${RESULTS_SCRIPT_URL}?action=getRduAggregatedData&yard=${encodeURIComponent(yard)}`).then(r => r.json())
        ]);

        const availK = (typeof fleetDictionary !== 'undefined' && fleetDictionary[yard]) ? fleetDictionary[yard].kamag : 0;
        const availM = (typeof fleetDictionary !== 'undefined' && fleetDictionary[yard]) ? fleetDictionary[yard].man : 0;

        let maxK = availK;
        let maxM = availM;

        const findMaxFleetBounds = (rows) => {
            if (!rows) return;
            rows.forEach(row => {
                const countStr = String(row[3] || "");
                const parts = countStr.split('|');
                if (parts[0]) maxK = Math.max(maxK, parts[0].split(',').length);
                if (parts[1]) maxM = Math.max(maxM, parts[1].split(',').length);
            });
        };

        findMaxFleetBounds(resPlan.savedRows);
        findMaxFleetBounds(resRdu.savedRows);

        let yardNorms = { k: 12, m: 6 };
        if (typeof yardDictionary !== 'undefined') {
            for (let node in yardDictionary) {
                if (yardDictionary[node].yard === yard) {
                    yardNorms.k = yardDictionary[node].normKamag || 12;
                    yardNorms.m = yardDictionary[node].normMan || 6;
                    break;
                }
            }
        }

        const vehicleRows = [];
        for (let i = 1; i <= maxK; i++) vehicleRows.push(`Kamag ${i}`);
        for (let i = 1; i <= maxM; i++) vehicleRows.push(`Маневровий ${i}`);

        const planFleet = {}, rduFleet = {};
        const planOps = {}, autoFactOps = {};
        const planCap = {}, autoFactCap = {};

        datesList.forEach(dStr => {
            planOps[dStr] = Array(24).fill(0);
            autoFactOps[dStr] = Array(24).fill(0);
            planCap[dStr] = Array(24).fill(0);
            autoFactCap[dStr] = Array(24).fill(0);
        });

        vehicleRows.forEach(v => {
            planFleet[v] = {};
            rduFleet[v] = {};
            datesList.forEach(dStr => {
                planFleet[v][dStr] = Array(24).fill(0);
                rduFleet[v][dStr] = Array(24).fill(0);
            });
        });

        const normalizeDay = (val) => {
            let s = String(val);
            if (s.includes('T')) {
                const d = new Date(s);
                return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
            }
            return s;
        };

        if (resPlan.savedRows) {
            resPlan.savedRows.forEach(row => {
                const dStr = normalizeDay(row[1]);
                if (!datesList.includes(dStr)) return;
                const h = parseInt(row[2], 10);
                planOps[dStr][h] = parseInt(row[4], 10) || 0;

                const [kStr, mStr] = String(row[3]).split('|');
                if (kStr) {
                    kStr.split(',').forEach((bit, idx) => {
                        if (idx < maxK) planFleet[`Kamag ${idx+1}`][dStr][h] = parseInt(bit, 10) || 0;
                    });
                }
                if (mStr) {
                    mStr.split(',').forEach((bit, idx) => {
                        if (idx < maxM) planFleet[`Маневровий ${idx+1}`][dStr][h] = parseInt(bit, 10) || 0;
                    });
                }
            });
        }

        if (resAutoFact.savedRows) {
            resAutoFact.savedRows.forEach(row => {
                const dStr = normalizeDay(row[1]);
                if (!datesList.includes(dStr)) return;
                const h = parseInt(row[2], 10);
                autoFactOps[dStr][h] = parseInt(row[4], 10) || 0;

                const [kStr, mStr] = String(row[3]).split('|');
                let activeK = 0, activeM = 0;
                if (kStr) activeK = kStr.split(',').map(Number).filter(Boolean).length;
                if (mStr) activeM = mStr.split(',').map(Number).filter(Boolean).length;

                autoFactCap[dStr][h] = (activeK * yardNorms.k) + (activeM * yardNorms.m);
            });
        }

        if (resRdu.savedRows) {
            resRdu.savedRows.forEach(row => {
                const dStr = normalizeDay(row[1]);
                if (!datesList.includes(dStr)) return;
                const h = parseInt(row[2], 10);

                const [kStr, mStr] = String(row[3]).split('|');
                if (kStr) {
                    kStr.split(',').forEach((bit, idx) => {
                        if (idx < maxK) rduFleet[`Kamag ${idx+1}`][dStr][h] = parseInt(bit, 10) || 0;
                    });
                }
                if (mStr) {
                    mStr.split(',').forEach((bit, idx) => {
                        if (idx < maxM) rduFleet[`Маневровий ${idx+1}`][dStr][h] = parseInt(bit, 10) || 0;
                    });
                }
            });
        }

        datesList.forEach(dStr => {
            for (let h = 0; h < 24; h++) {
                let activeK = 0, activeM = 0;
                for (let i = 1; i <= maxK; i++) if (planFleet[`Kamag ${i}`][dStr][h] === 1) activeK++;
                for (let i = 1; i <= maxM; i++) if (planFleet[`Маневровий ${i}`][dStr][h] === 1) activeM++;
                planCap[dStr][h] = (activeK * yardNorms.k) + (activeM * yardNorms.m);
            }
        });

        lastCompareData = { yard, datesList, vehicleRows, planFleet, rduFleet, planOps, autoFactOps, planCap, autoFactCap, maxK, maxM };

        buildCompareTableHTML(vehicleRows, planFleet, rduFleet, datesList, planOps, autoFactOps);
        buildCombinedChart(datesList, planOps, autoFactOps, planCap, autoFactCap);

    } catch (e) {
        console.error(e);
        alert("Помилка при зборі аналітики!");
    } finally {
        btn.innerText = "Завантажити дані"; btn.disabled = false;
    }
}

function buildCompareTableHTML(vehicleRows, planFleet, rduFleet, datesList, planOps, autoFactOps) {
    const container = document.getElementById('compareTableWrapper');
    
    // Вычисляем общее количество колонок для красивой строки-разделителя
    const totalCols = 1 + (datesList.length * 26) + 2; 

    let html = `<h3 style="margin: 5px 0; color: #334155; border-left: 4px solid #ea580c; padding-left: 10px;">Порівняльна матриця роботи ТЗ</h3>
    <table><thead><tr><th class="sticky-col" style="min-width: 150px;" rowspan="2">ТЗ / Стан</th>`;
    
    datesList.forEach(dStr => {
        html += `<th colspan="26" style="text-align: center; font-weight: bold; background-color: #e9ecef; border-left: 2px solid #6c757d; border-right: 2px solid #6c757d; padding: 4px 0;">${dStr}</th>`;
    });
    html += `<th rowspan="2" style="background-color: #cbd5e1; font-weight:bold; min-width:55px; vertical-align: middle;">Разом План</th>
             <th rowspan="2" style="background-color: #cbd5e1; font-weight:bold; min-width:45px; vertical-align: middle;">Разом РДУ</th></tr><tr>`;

    datesList.forEach(dStr => {
        for (let h = 0; h < 24; h++) {
            html += `<th class="kamag-header-vertical" style="height:30px; ${h === 0 ? 'border-left: 2px solid #6c757d;' : ''}">${h}:00</th>`;
        }
        html += `<th style="background-color: #dee2e6; font-weight:bold; font-size:10px;">Σ П</th>`;
        html += `<th style="background-color: #dee2e6; font-weight:bold; font-size:10px; border-right: 2px solid #6c757d;">Σ Р</th>`;
    });
    html += `</tr></thead><tbody>`;

    // Рендерим строки машин
    vehicleRows.forEach(v => {
        html += `<tr><td class="sticky-col" style="font-weight:bold; background-color:#fff;">${v}</td>`;
        let grandPlanHours = 0;
        let grandRduHours = 0;

        datesList.forEach(dStr => {
            let dayPlanHours = 0;
            let dayRduHours = 0;

            for (let h = 0; h < 24; h++) {
                const pBit = planFleet[v][dStr][h];
                const rBit = rduFleet[v][dStr][h];
                
                let bgStyle = "";
                let displayVal = "";
                let borderStyle = h === 0 ? "border-left: 2px solid #6c757d;" : "";

                if (pBit === 1 && rBit === 1) {
                    bgStyle = "background-color: #ea580c; color: #fff; font-weight: bold;"; 
                    displayVal = "1";
                    dayPlanHours++; dayRduHours++;
                } else if (pBit === 0 && rBit === 1) {
                    bgStyle = "background-color: #be123c; color: #fff; font-weight: bold;"; 
                    displayVal = "1";
                    dayRduHours++;
                } else if (pBit === 1 && rBit === 0) {
                    bgStyle = "background-color: #ffedd5; color: #9a3412; font-weight: bold; border: 1px solid #fed7aa;"; 
                    displayVal = "0";
                    dayPlanHours++;
                } else {
                    bgStyle = "background-color: #ffffff;";
                    displayVal = "";
                }

                html += `<td class="kamag-cell" style="${bgStyle} ${borderStyle}">${displayVal}</td>`;
            }
            
            grandPlanHours += dayPlanHours;
            grandRduHours += dayRduHours;

            html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5;">${dayPlanHours || ''}</td>`;
            html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right: 2px solid #6c757d;">${dayRduHours || ''}</td>`;
        });
        
        html += `<td style="text-align:center; font-weight:bold; background-color:#cbd5e1;">${grandPlanHours || ''}</td>
                 <td style="text-align:center; font-weight:bold; background-color:#cbd5e1;">${grandRduHours || ''}</td></tr>`;
    });

    // Строка с графиками внутри основной таблицы
    html += `<tr><td class="sticky-col" style="font-weight:bold; background-color:#fff; vertical-align: middle;">Графік роботи</td>`;
    datesList.forEach((dStr, index) => {
        html += `<td colspan="26" style="padding: 0; background: #fff; vertical-align: bottom; border-left: 2px solid #6c757d; border-right: 2px solid #6c757d;">
            <div style="height: 220px; width: 100%;">
                <canvas id="combinedCompareChart_${index}"></canvas>
            </div>
        </td>`;
    });
    html += `<td style="background-color:#cbd5e1;"></td><td style="background-color:#cbd5e1;"></td></tr>`;
    
    // --- ВСТАВКА ОПЕРАЦИЙ В ТУ ЖЕ ТАБЛИЦУ (БЕЗ ЗАКРЫТИЯ СТРУКТУРЫ) ---
    // Строка-заголовок секции
    html += `<tr><td colspan="${totalCols}" style="padding: 20px 0 5px 0; background-color: #fff; border: none; pointer-events: none;">
        <h3 style="margin: 0; color: #334155; border-left: 4px solid #10b981; padding-left: 10px; text-align: left;">Сумарні дані по операціям за період</h3>
    </td></tr>`;
    
    // Дублирующие подзаголовки часов (чтобы перед глазами была сетка)
    html += `<tr style="background-color: #f8fafc;"><th class="sticky-col" style="min-width: 150px; text-align: left;" rowspan="2">Параметр / Години</th>`;
    datesList.forEach(dStr => {
        html += `<th colspan="26" style="text-align: center; font-weight: bold; background-color: #f8fafc; border-left: 2px solid #6c757d; border-right: 2px solid #6c757d; padding: 4px 0;">${dStr}</th>`;
    });
    html += `<th rowspan="2" colspan="2" style="background-color: #cbd5e1; font-weight:bold; min-width:65px; vertical-align: middle;">Всього</th></tr><tr style="background-color: #f8fafc;">`;
    
    datesList.forEach(dStr => {
        for (let h = 0; h < 24; h++) {
            html += `<th class="kamag-header-vertical" style="height:30px; ${h === 0 ? 'border-left: 2px solid #6c757d;' : ''}">${h}:00</th>`;
        }
        html += `<th colspan="2" style="background-color: #dee2e6; font-weight:bold; font-size:10px; border-right: 2px solid #6c757d; text-align:center;">Σ</th>`;
    });
    html += `</tr>`;
    
    // 1. Строка: План
    html += `<tr><td class="sticky-col" style="font-weight:bold; background-color:#fff;">Операції (План)</td>`;
    let grandPlanOps = 0;
    datesList.forEach(dStr => {
        let daySum = 0;
        for(let h=0; h<24; h++) {
            let val = planOps[dStr][h];
            let border = h === 0 ? "border-left: 2px solid #6c757d;" : "";
            html += `<td class="kamag-cell" style="${border} background-color: #fff9c4; font-weight:bold;">${val || ''}</td>`;
            daySum += val;
        }
        grandPlanOps += daySum;
        html += `<td colspan="2" style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right: 2px solid #6c757d;">${daySum || '0'}</td>`;
    });
    html += `<td colspan="2" style="text-align:center; font-weight:bold; background-color:#cbd5e1;">${grandPlanOps || '0'}</td></tr>`;
    
    // 2. Строка: Факт
    html += `<tr><td class="sticky-col" style="font-weight:bold; background-color:#fff;">Операції (Факт)</td>`;
    let grandFactOps = 0;
    datesList.forEach(dStr => {
        let daySum = 0;
        for(let h=0; h<24; h++) {
            let val = autoFactOps[dStr][h];
            let border = h === 0 ? "border-left: 2px solid #6c757d;" : "";
            html += `<td class="kamag-cell" style="${border} background-color: #d1fae5; font-weight:bold;">${val || ''}</td>`;
            daySum += val;
        }
        grandFactOps += daySum;
        html += `<td colspan="2" style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right: 2px solid #6c757d;">${daySum || '0'}</td>`;
    });
    html += `<td colspan="2" style="text-align:center; font-weight:bold; background-color:#cbd5e1;">${grandFactOps || '0'}</td></tr>`;
    
    // 3. Строка: Разница
    html += `<tr><td class="sticky-col" style="font-weight:bold; background-color:#fff;">Різниця (Ф - П)</td>`;
    let grandDiffOps = 0;
    datesList.forEach(dStr => {
        let daySum = 0;
        for(let h=0; h<24; h++) {
            let diff = autoFactOps[dStr][h] - planOps[dStr][h];
            let border = h === 0 ? "border-left: 2px solid #6c757d;" : "";
            let diffStyle = diff < 0 ? "color:#dc2626; background:#fef2f2;" : (diff > 0 ? "color:#16a34a; background:#f0fdf4;" : "");
            html += `<td class="kamag-cell" style="${border} font-weight:bold; ${diffStyle}">${diff || '0'}</td>`;
            daySum += diff;
        }
        grandDiffOps += daySum;
        let dayDiffStyle = daySum < 0 ? "color:#dc2626;" : (daySum > 0 ? "color:#16a34a;" : "");
        html += `<td colspan="2" style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right: 2px solid #6c757d; ${dayDiffStyle}">${daySum}</td>`;
    });
    let grandDiffStyle = grandDiffOps < 0 ? "color:#dc2626;" : (grandDiffOps > 0 ? "color:#16a34a;" : "");
    html += `<td colspan="2" style="text-align:center; font-weight:bold; background-color:#cbd5e1; ${grandDiffStyle}">${grandDiffOps}</td></tr>`;
    
    // Закрываем общую таблицу
    html += `</tbody></table>`;
    
    container.innerHTML = html;
}

function buildCombinedChart(datesList, planOps, autoFactOps, planCap, autoFactCap) {
    if (window.myCombinedCompareCharts) {
        window.myCombinedCompareCharts.forEach(c => c.destroy());
    }
    window.myCombinedCompareCharts = [];

    datesList.forEach((dateStr, index) => {
        const ctx = document.getElementById(`combinedCompareChart_${index}`);
        if (!ctx) return;

        const parentDiv = ctx.parentElement;
        ctx.width = parentDiv.clientWidth;
        ctx.height = 220;

        const labels = [];
        for (let h = 0; h < 24; h++) labels.push(`${h}:00`);

        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    { type: 'bar', label: 'Операції (План)', data: planOps[dateStr], backgroundColor: 'rgba(255, 170, 0, 0.6)', borderColor: '#ffaa00', borderWidth: 1, borderRadius: 2, order: 3 },
                    { type: 'bar', label: 'Операції (Факт)', data: autoFactOps[dateStr], backgroundColor: 'rgba(16, 185, 129, 0.6)', borderColor: '#10b981', borderWidth: 1, borderRadius: 2, order: 4 },
                    { type: 'line', label: 'Транспорт (План)', data: planCap[dateStr], borderColor: '#2563eb', backgroundColor: '#2563eb', borderWidth: 2.5, tension: 0.2, pointRadius: 2, order: 1 },
                    { type: 'line', label: 'Транспорт (Факт)', data: autoFactCap[dateStr], borderColor: '#dc2626', backgroundColor: '#dc2626', borderWidth: 2.5, tension: 0.2, pointRadius: 2, order: 2 }
                ]
            },
            options: {
                animation: false,
                responsive: false, 
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    x: { display: false, grid: { display: false } },
                    y: { 
                        type: 'linear', 
                        beginAtZero: true, 
                        display: index === 0, 
                        title: { display: true, text: 'Шт / Год', font: { size: 10, weight: 'bold' } } 
                    }
                },
                plugins: {
                    legend: { 
                        display: index === 0, 
                        position: 'top', 
                        labels: { boxWidth: 12, font: { size: 10 } } 
                    },
                    tooltip: { callbacks: { title: (items) => `День: ${dateStr} о ${items[0].label}` } }
                }
            }
        });
        
        window.myCombinedCompareCharts.push(chart);
    });
}

// СИНХРОНІЗАЦІЯ ВИПРАВЛЕНА: Об'єднуємо підсумки на 2 комірки в Excel (sheet.mergeCells)
window.exportCompareToExcel = async function(workbook) {
    if (!lastCompareData) return alert("Спочатку завантажте дані звірки на екран!");
    const { yard, datesList, vehicleRows, planFleet, rduFleet, planOps, autoFactOps, planCap, autoFactCap, maxK, maxM } = lastCompareData;
    
    const sheet = workbook.addWorksheet('Звірка');
    const alignCenter = { vertical: 'middle', horizontal: 'center' };
    const fillHeader = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } };
    const fillMatch = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEA580C' } }; 
    const fillExcess = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBE123C' } }; 
    const fillIdle = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEDD5' } }; 
    const fillSum = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F3F5' } };
    const borderThin = { style: 'thin', color: { argb: 'FFCCCCCC' } };
    const borderMedium = { style: 'medium', color: { argb: 'FF6C757D' } };
    
    const getBorders = (isLeftEdge, isRightEdge) => ({
        top: borderThin, bottom: borderThin,
        left: isLeftEdge ? borderMedium : borderThin,
        right: isRightEdge ? borderMedium : borderThin
    });

    sheet.getColumn(1).width = 20; 
    sheet.getColumn(2).width = 4;  
    
    let colCounter = 3; 
    datesList.forEach(() => {
        for(let h=0; h<24; h++) { sheet.getColumn(colCounter++).width = 4; }
        sheet.getColumn(colCounter++).width = 6; 
        sheet.getColumn(colCounter++).width = 6; 
    });

    sheet.addRow([`Звіт звірки ТЗ та операцій по автодвору: ${yard}`]).font = { bold: true, size: 14 };
    sheet.addRow([]);

    const rowDays = sheet.addRow();
    rowDays.getCell(1).value = "ТЗ / Стан";
    rowDays.getCell(1).font = { bold: true };
    rowDays.getCell(1).alignment = alignCenter;

    datesList.forEach((dStr, dIdx) => {
        const startCol = 3 + dIdx * 26; 
        const endCol = startCol + 25; 
        sheet.mergeCells(3, startCol, 3, endCol);
        const cell = sheet.getCell(3, startCol);
        cell.value = dStr;
        cell.alignment = alignCenter;
        cell.font = { bold: true };
        cell.fill = fillHeader;
    });
    
    const totalPlanColIdx = 3 + datesList.length * 26;
    sheet.getCell(3, totalPlanColIdx).value = "Разов План";
    sheet.getCell(3, totalPlanColIdx).font = { bold: true };
    sheet.getCell(3, totalPlanColIdx + 1).value = "Разом РДУ";
    sheet.getCell(3, totalPlanColIdx + 1).font = { bold: true };

    const rowHours = sheet.addRow();
    let currentC = 3; 
    datesList.forEach(() => {
        for (let h = 0; h < 24; h++) {
            const cell = rowHours.getCell(currentC++);
            cell.value = h;
            cell.alignment = alignCenter;
            cell.font = { size: 9 };
        }
        const cellP = rowHours.getCell(currentC++); cellP.value = "Σ П"; cellP.font = { bold: true, size: 9 }; cellP.fill = fillHeader; cellP.alignment = alignCenter;
        const cellR = rowHours.getCell(currentC++); cellR.value = "Σ Р"; cellR.font = { bold: true, size: 9 }; cellR.fill = fillHeader; cellR.alignment = alignCenter;
    });

    vehicleRows.forEach(v => {
        const row = sheet.addRow();
        row.getCell(1).value = v;
        row.getCell(1).font = { bold: true };
        
        let grandPlanHours = 0;
        let grandRduHours = 0;
        let cCol = 3; 

        datesList.forEach(dStr => {
            let dayPlanHours = 0;
            let dayRduHours = 0;

            for (let h = 0; h < 24; h++) {
                const pBit = planFleet[v][dStr][h];
                const rBit = rduFleet[v][dStr][h];
                const cell = row.getCell(cCol);
                cell.alignment = alignCenter;
                cell.border = getBorders(h === 0, false);

                if (pBit === 1 && rBit === 1) {
                    cell.value = 1; cell.fill = fillMatch; cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
                    dayPlanHours++; dayRduHours++;
                } else if (pBit === 0 && rBit === 1) {
                    cell.value = 1; cell.fill = fillExcess; cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
                    dayRduHours++;
                } else if (pBit === 1 && rBit === 0) {
                    cell.value = 0; cell.fill = fillIdle; cell.font = { color: { argb: 'FF9A3412' }, bold: true };
                    dayPlanHours++;
                }
                cCol++;
            }
            grandPlanHours += dayPlanHours;
            grandRduHours += dayRduHours;

            const cellSumP = row.getCell(cCol++); cellSumP.value = dayPlanHours || ""; cellSumP.fill = fillSum; cellSumP.font = { bold: true }; cellSumP.alignment = alignCenter;
            const cellSumR = row.getCell(cCol++); cellSumR.value = dayRduHours || ""; cellSumR.fill = fillSum; cellSumR.font = { bold: true }; cellSumR.alignment = alignCenter; cellSumR.border = getBorders(false, true);
        });

        const cellGrandP = row.getCell(cCol++); cellGrandP.value = grandPlanHours || ""; cellGrandP.font = { bold: true }; cellGrandP.fill = fillHeader; cellGrandP.alignment = alignCenter;
        const cellGrandR = row.getCell(cCol++); cellGrandR.value = grandRduHours || ""; cellGrandR.font = { bold: true }; cellGrandR.fill = fillHeader; cellGrandR.alignment = alignCenter;
    });

    sheet.addRow([]);
    const imgStartRow = sheet.rowCount + 1; 
    sheet.addRow(["Графіки суміщеної роботи (Транспорт та Операції):"]).font = { bold: true };

    for (let r = imgStartRow + 1; r <= imgStartRow + 12; r++) {
        sheet.getRow(r).height = 18; 
    }

    datesList.forEach((dStr, dIdx) => {
        const canvas = document.createElement('canvas');
        canvas.width = 640; 
        canvas.height = 200;
        const ctx = canvas.getContext('2d');

        const tempChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: Array(24).fill('').map((_, h) => `${h}:00`),
                datasets: [
                    { type: 'bar', label: 'Оп. План', data: planOps[dStr], backgroundColor: 'rgba(255, 170, 0, 0.6)' },
                    { type: 'bar', label: 'Оп. Факт', data: autoFactOps[dStr], backgroundColor: 'rgba(16, 185, 129, 0.6)' },
                    { type: 'line', label: 'Тр. План', data: planCap[dStr], borderColor: '#2563eb', borderWidth: 2, pointRadius: 1, tension: 0.1 },
                    { type: 'line', label: 'Тр. Факт', data: autoFactCap[dStr], borderColor: '#dc2626', borderWidth: 2, pointRadius: 1, tension: 0.1 }
                ]
            },
            options: {
                animation: false,
                responsive: false,
                scales: { 
                    x: { ticks: { font: { size: 8 } } }, 
                    y: { beginAtZero: true, ticks: { font: { size: 8 } } } 
                },
                plugins: { legend: { display: dIdx === 0, labels: { font: { size: 8 } } } } 
            }
        });

        const base64 = canvas.toDataURL('image/png');
        tempChart.destroy(); 

        const startColumnZeroIndexed = 2 + dIdx * 26; 
        const endColumnZeroIndexed = startColumnZeroIndexed + 24; 

        sheet.addImage(workbook.addImage({ base64: base64.split(',')[1], extension: 'png' }), {
            tl: { col: startColumnZeroIndexed, row: imgStartRow + 1 },
            br: { col: endColumnZeroIndexed, row: imgStartRow + 12 },
            editAs: 'twoCell' 
        });
    });

    //for (let i = 0; i < 10; i++) sheet.addRow([]);

    sheet.addRow([]);
    sheet.addRow(["Сумарні дані по операціям"]).font = { bold: true, size: 12 };
    sheet.addRow([]);

    // Шапка годин під нижню таблицю операцій (Синхронно мержимо стовпчик Σ на 2 комірки)
    const rowHoursOps = sheet.addRow();
    let currentCOps = 3;
    datesList.forEach(() => {
        for (let h = 0; h < 24; h++) {
            const cell = rowHoursOps.getCell(currentCOps++);
            cell.value = h;
            cell.alignment = alignCenter;
            cell.font = { size: 9 };
        }
        const lblCell = rowHoursOps.getCell(currentCOps);
        lblCell.value = "Σ"; lblCell.font = { bold: true, size: 9 }; lblCell.fill = fillHeader; lblCell.alignment = alignCenter;
        sheet.mergeCells(rowHoursOps.number, currentCOps, rowHoursOps.number, currentCOps + 1);
        currentCOps += 2;
    });

    const opsConfig = [
        { label: "Операції (План)", data: planOps, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } } },
        { label: "Операції (Факт)", data: autoFactOps, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } } },
        { label: "Різниця (Ф - П)", isDiff: true }
    ];

    opsConfig.forEach(cfg => {
        const row = sheet.addRow();
        row.getCell(1).value = cfg.label;
        row.getCell(1).font = { bold: true };

        let grandTotal = 0;
        let cCol = 3; 

        datesList.forEach(dStr => {
            let daySum = 0;
            for (let h = 0; h < 24; h++) {
                const cell = row.getCell(cCol);
                cell.alignment = alignCenter;
                cell.border = getBorders(h === 0, false);

                let val = 0;
                if (cfg.isDiff) {
                    val = autoFactOps[dStr][h] - planOps[dStr][h];
                    cell.value = val;
                    cell.font = { bold: true, color: { argb: val < 0 ? 'FFDC2626' : (val > 0 ? 'FF16A34A' : 'FF000000') } };
                } else {
                    val = cfg.data[dStr][h];
                    if (val > 0) {
                        cell.value = val;
                        cell.fill = cfg.fill;
                        cell.font = { bold: true };
                    }
                }
                daySum += val;
                cCol++;
            }
            grandTotal += daySum;
            
            // Запис підсумку з очищенням та злиттям 2-х фінальних стовпчиків
            const cellDaySum = row.getCell(cCol);
            cellDaySum.value = daySum;
            cellDaySum.fill = fillSum;
            cellDaySum.font = { bold: true };
            cellDaySum.alignment = alignCenter;
            if (cfg.isDiff) {
                cellDaySum.font = { bold: true, color: { argb: daySum < 0 ? 'FFDC2626' : (daySum > 0 ? 'FF16A34A' : 'FF000000') } };
            }
            
            const cellDaySum2 = row.getCell(cCol + 1);
            cellDaySum2.fill = fillSum;
            
            sheet.mergeCells(row.number, cCol, row.number, cCol + 1);
            cellDaySum2.border = getBorders(false, true);
            
            cCol += 2;
        });

        const cellGrand = row.getCell(cCol++);
        cellGrand.value = grandTotal;
        cellGrand.fill = fillHeader;
        cellGrand.font = { bold: true };
        cellGrand.alignment = alignCenter;
        if (cfg.isDiff) {
            cellGrand.font = { bold: true, color: { argb: grandTotal < 0 ? 'FFDC2626' : (grandTotal > 0 ? 'FF16A34A' : 'FF000000') } };
        }
    });
};