// === МОДУЛЬ КОМБІНОВАНОЇ ЗВІРКИ ТА КРОС-АНАЛІТИКИ (compare.js) ===
document.addEventListener('DOMContentLoaded', () => {
    const tabCompare = document.getElementById('tabCompare');
    const containerCompare = document.getElementById('tableContainerCompare');
    const compareStartDateInput = document.getElementById('compareStartDate');
    const compareEndDateInput = document.getElementById('compareEndDate');

    if (!tabCompare || !containerCompare) return;

    // Встановлення сьогоднішньої дати за замовчуванням для старту і кінця періоду
    const today = new Date();
    compareStartDateInput.value = today.toISOString().split('T')[0];
    compareEndDateInput.value = today.toISOString().split('T')[0];

    tabCompare.addEventListener('click', () => {
        switchTab(tabCompare, containerCompare);
        
        // Налаштування селектора для адміна
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
        
        clearCompareDashboard();
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

    // Генеруємо масив строк дат всередині обраного періоду
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
        // Паралельний запит до трьох баз даних (План, Авто-Факт, Ручний Факт РДУ)
        const [resPlan, resAutoFact, resRdu] = await Promise.all([
            fetch(`${RESULTS_SCRIPT_URL}?action=getAggregatedData&yard=${encodeURIComponent(yard)}`).then(r => r.json()),
            fetch(`${RESULTS_SCRIPT_URL}?action=getFactAggregatedData&yard=${encodeURIComponent(yard)}`).then(r => r.json()),
            fetch(`${RESULTS_SCRIPT_URL}?action=getRduAggregatedData&yard=${encodeURIComponent(yard)}`).then(r => r.json())
        ]);

        // Базові ліміти флоту з довідника
        const availK = (typeof fleetDictionary !== 'undefined' && fleetDictionary[yard]) ? fleetDictionary[yard].kamag : 0;
        const availM = (typeof fleetDictionary !== 'undefined' && fleetDictionary[yard]) ? fleetDictionary[yard].man : 0;

        // Динамічно скануємо максимальні межі рядків (враховуємо додані руками машини в РДУ)
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

        // Формуємо фінальний масив рядків транспорту
        const vehicleRows = [];
        for (let i = 1; i <= maxK; i++) vehicleRows.push(`Kamag ${i}`);
        for (let i = 1; i <= maxM; i++) vehicleRows.push(`Маневровий ${i}`);

        // Багатомірні структури під збірку матриць за днями
        const planFleet = {}, rduFleet = {};
        const planOps = {}, autoFactOps = {};
        const planCap = {}, autoFactCap = {};

        // Ініціалізація порожніх масивів під кожну дату
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

        // 1. ПАРСИНГ ПЛАНУ
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

        // 2. ПАРСИНГ АВТОМАТИЧНОГО ФАКТУ (ЗІ СКАНУВАНЬ)
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

        // 3. ПАРСИНГ РУЧНОГО ФАКТУ РДУ
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

        // Розрахунок планової пропускної потужності
        datesList.forEach(dStr => {
            for (let h = 0; h < 24; h++) {
                let activeK = 0, activeM = 0;
                for (let i = 1; i <= maxK; i++) if (planFleet[`Kamag ${i}`][dStr][h] === 1) activeK++;
                for (let i = 1; i <= maxM; i++) if (planFleet[`Маневровий ${i}`][dStr][h] === 1) activeM++;
                planCap[dStr][h] = (activeK * yardNorms.k) + (activeM * yardNorms.m);
            }
        });

        // Збірка HTML матриці
        buildCompareTableHTML(vehicleRows, planFleet, rduFleet, datesList);

        // Отрисовка вбудованих графіків
        buildCombinedChart(datesList, planOps, autoFactOps, planCap, autoFactCap);

    } catch (e) {
        console.error(e);
        alert("Помилка при зборі аналітики!");
    } finally {
        btn.innerText = "Завантажити дані"; btn.disabled = false;
    }
}

function buildCompareTableHTML(vehicleRows, planFleet, rduFleet, datesList) {
    const container = document.getElementById('compareTableWrapper');
    
    let html = `<h3 style="margin: 5px 0; color: #334155; border-left: 4px solid #ea580c; padding-left: 10px;">Порівняльна матриця роботи ТЗ</h3>
    <table><thead><tr><th style="min-width: 150px;" rowspan="2">ТЗ / Стан</th>`;
    
    // Верхній ярус шапки — дати
    datesList.forEach(dStr => {
        html += `<th colspan="26" style="text-align: center; font-weight: bold; background-color: #e9ecef; border-left: 2px solid #6c757d; border-right: 2px solid #6c757d; padding: 4px 0;">${dStr}</th>`;
    });
    html += `<th rowspan="2" style="background-color: #cbd5e1; font-weight:bold; min-width:55px; vertical-align: middle;">Разом План</th>
             <th rowspan="2" style="background-color: #cbd5e1; font-weight:bold; min-width:45px; vertical-align: middle;">Разом РДУ</th></tr><tr>`;

    // Нижній ярус шапки — години та посуточні підсумки
    datesList.forEach(dStr => {
        for (let h = 0; h < 24; h++) {
            html += `<th class="kamag-header-vertical" style="height:30px; ${h === 0 ? 'border-left: 2px solid #6c757d;' : ''}">${h}:00</th>`;
        }
        html += `<th style="background-color: #dee2e6; font-weight:bold; font-size:10px;">Σ П</th>`;
        html += `<th style="background-color: #dee2e6; font-weight:bold; font-size:10px; border-right: 2px solid #6c757d;">Σ Р</th>`;
    });
    html += `</tr></thead><tbody>`;

    // Заповнення рядків транспорту
    vehicleRows.forEach(v => {
        html += `<tr><td style="font-weight:bold; background-color:#fff;">${v}</td>`;
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
                    bgStyle = "background-color: #ea580c; color: #fff; font-weight: bold;"; // СОВПАДЕНИЕ
                    displayVal = "1";
                    dayPlanHours++; dayRduHours++;
                } else if (pBit === 0 && rBit === 1) {
                    bgStyle = "background-color: #dc2626; color: #fff; font-weight: bold;"; // СВЕРХ СМЕНЫ (КРАСНЫЙ)
                    displayVal = "1";
                    dayRduHours++;
                } else if (pBit === 1 && rBit === 0) {
                    bgStyle = "background-color: #ffedd5; color: #9a3412; font-weight: bold; border: 1px solid #fed7aa;"; // ПРОСТОЙ (БЛЕДНЫЙ)
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
        
        // Кінцеві глобальні підсумки за період
        html += `<td style="text-align:center; font-weight:bold; background-color:#cbd5e1;">${grandPlanHours || ''}</td>
                 <td style="text-align:center; font-weight:bold; background-color:#cbd5e1;">${grandRduHours || ''}</td></tr>`;
    });

    // Вставка рядка з графіками точнісінько під годинами кожного дня
    html += `<tr><td style="font-weight:bold; background-color:#fff; vertical-align: middle;">Графік роботи</td>`;
    
    datesList.forEach((dStr, index) => {
        html += `<td colspan="26" style="padding: 0; background: #fff; vertical-align: bottom; border-left: 2px solid #6c757d; border-right: 2px solid #6c757d;">
            <div style="height: 220px; width: 100%;">
                <canvas id="combinedCompareChart_${index}"></canvas>
            </div>
        </td>`;
    });

    html += `<td style="background-color:#cbd5e1;"></td><td style="background-color:#cbd5e1;"></td></tr>`;
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
                responsive: false, // Вимикаємо адаптивність, щоб утримати піксельну сітку th стовпчиків
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    x: { display: false, grid: { display: false } }, // Ховаємо підписи th, бо вони дублюють шапку
                    y: { 
                        type: 'linear', 
                        beginAtZero: true, 
                        display: index === 0, // Показуємо шкалу значень тільки для першого (лівого) графіка
                        title: { display: true, text: 'Шт / Год', font: { size: 10, weight: 'bold' } } 
                    }
                },
                plugins: {
                    legend: { 
                        display: index === 0, // Легенду виводимо також лише один раз на лівому графіку
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