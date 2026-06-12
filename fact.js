// --- МОДУЛЬ РОБОТИ З ФАКТИЧНИМИ ДАНИМИ (fact.js) ---
let actualFlightsData = [];
let factCalculatedEvents = [];
let factOpsMatrix = {}; 

document.addEventListener('DOMContentLoaded', () => {
    const tabFact = document.getElementById('tabFact');
    const containerFact = document.getElementById('tableContainerFact');
    const factFileInput = document.getElementById('factFileInput');
    const factYardSelect = document.getElementById('factYardSelect');
    const factStartDate = document.getElementById('factStartDate');
    const factEndDate = document.getElementById('factEndDate');

    if (!tabFact || !containerFact) return;

    // 1. Переключення вкладки
    tabFact.addEventListener('click', () => {
        if (typeof switchTab === 'function') {
            switchTab(tabFact, containerFact);
        } else {
            document.querySelectorAll('.tabs .tab-btn').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.main-content .table-container').forEach(cont => cont.style.display = 'none');
            tabFact.classList.add('active');
            containerFact.style.display = 'flex';
        }
        updateFactYardsDropdown();
        loadSavedFactYardsList();
        renderFactDashboard();
    });

    const loadFactBtn = document.getElementById('loadFactGoogleYardBtn');
    if (loadFactBtn) loadFactBtn.addEventListener('click', loadFactFromGoogle);

    const saveFactBtn = document.getElementById('saveFactGoogleBtn');
    if (saveFactBtn) saveFactBtn.addEventListener('click', saveFactToGoogle);

    const saveAllFactBtn = document.getElementById('saveAllFactGoogleBtn');
    if (saveAllFactBtn) saveAllFactBtn.addEventListener('click', saveAllFactToGoogle);

    // 2. Обробник файлу CSV
    if (factFileInput) {
        factFileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            if (sessionStorage.getItem('kamagonAuth') !== 'true') {
                alert("Завантаження доступне лише для авторизованих користувачів!");
                return;
            }

            const reader = new FileReader();
            const statusEl = document.getElementById('factFileStatus');
            statusEl.innerText = "⏳ Обробка та розрахунок...";

            reader.onload = function(event) {
                processFactData(event.target.result);
                statusEl.innerText = `✅ Oбролено рейсів: ${actualFlightsData.length}`;
                updateFactYardsDropdown();
                renderFactDashboard();
            };
            reader.readAsText(file, 'windows-1251');
        });
    }

    // 3. Слухачі фільтрів
    if (factYardSelect) factYardSelect.addEventListener('change', renderFactDashboard);
    if (factStartDate) factStartDate.addEventListener('change', renderFactDashboard);
    if (factEndDate) factEndDate.addEventListener('change', renderFactDashboard);
});

// Робота з часом
function parseCSVDateTime(str) {
    if (!str || !str.includes(' ')) return null;
    const [datePart, timePart] = str.split(' ');
    const [dd, mm, yyyy] = datePart.split('.').map(Number);
    const [hh, min, sec] = timePart.split(':').map(Number);
    return new Date(yyyy, mm - 1, dd, hh, min, sec || 0);
}

function modifyMinutes(dateObj, minutes) {
    if (!dateObj) return null;
    return new Date(dateObj.getTime() + minutes * 60000);
}

function formatTimeOnly(dateObj) {
    if (!dateObj) return '—';
    return dateObj.toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}

function formatDateOnly(dateObj) {
    if (!dateObj) return '—';
    return dateObj.toLocaleDateString('uk-UA', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

// Хелпер отримання списку відфільтрованих дат
function getFactFilteredDates(yard) {
    const startStr = document.getElementById('factStartDate').value;
    const endStr = document.getElementById('factEndDate').value;
    
    const startDate = startStr ? new Date(startStr).setHours(0,0,0,0) : null;
    const endDate = endStr ? new Date(endStr).setHours(23,59,59,999) : null;

    return Object.keys(factOpsMatrix[yard] || {}).sort((a, b) => {
        const [d1, m1, y1] = a.split('.').map(Number);
        const [d2, m2, y2] = b.split('.').map(Number);
        return new Date(y1, m1 - 1, d1) - new Date(y2, m2 - 1, d2);
    }).filter(dateStr => {
        const [dd, mm, yyyy] = dateStr.split('.').map(Number);
        const currentTimestamp = new Date(yyyy, mm - 1, dd).getTime();
        if (startDate && currentTimestamp < startDate) return false;
        if (endDate && currentTimestamp > endDate) return false;
        return true;
    });
}

// Допоміжна функція єдиного розрахунку флоту (для таблиці та для збереження)
function getFactFleetRequirements(yard, dates) {
    const availK = (typeof fleetDictionary !== 'undefined' && fleetDictionary[yard]) ? fleetDictionary[yard].kamag : 0;
    const availM = (typeof fleetDictionary !== 'undefined' && fleetDictionary[yard]) ? fleetDictionary[yard].man : 0;

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

    const virtualFleetRadio = document.querySelector('input[name="virtualFleetType"]:checked');
    const yardVirtualType = virtualFleetRadio ? virtualFleetRadio.value : (availK >= availM ? 'kamag' : 'man');

    const hourlyRequirements = {};
    let maxExtraK = 0;
    let maxExtraM = 0;

    dates.forEach(d => {
        hourlyRequirements[d] = Array(24).fill(null).map(() => ({ kamag: 0, man: 0, ops: 0 }));
        for (let h = 0; h < 24; h++) {
            let ops = (factOpsMatrix[yard] && factOpsMatrix[yard][d] && factOpsMatrix[yard][d][h]) ? factOpsMatrix[yard][d][h].total : 0;
            hourlyRequirements[d][h].ops = ops;

            let neededOps = ops;
            let assignedK = 0;
            while (neededOps > 0 && assignedK < availK) {
                assignedK++;
                neededOps -= yardNorms.k;
            }
            let assignedM = 0;
            while (neededOps > 0 && assignedM < availM) {
                assignedM++;
                neededOps -= yardNorms.m;
            }

            let extraK = 0;
            let extraM = 0;
            if (neededOps > 0) {
                if (yardVirtualType === 'kamag') {
                    extraK = Math.ceil(neededOps / yardNorms.k);
                    if (extraK > maxExtraK) maxExtraK = extraK;
                } else {
                    extraM = Math.ceil(neededOps / yardNorms.m);
                    if (extraM > maxExtraM) maxExtraM = extraM;
                }
            }

            hourlyRequirements[d][h].kamag = assignedK + extraK;
            hourlyRequirements[d][h].man = assignedM + extraM;
        }
    });

    return {
        hourlyRequirements,
        totalK: availK + maxExtraK,
        totalM: availM + maxExtraM,
        yardNorms
    };
}

// Парсинг та розрахунок 4-х операцій
function processFactData(text) {
    actualFlightsData = [];
    factCalculatedEvents = [];
    factOpsMatrix = {};

    const lines = text.split(/\r?\n/);
    if (lines.length < 2) return;

    const rawRows = [];
    for (let i = 1; i < lines.length; i++) {
        const currentLine = lines[i].trim();
        if (!currentLine) continue;

        const cols = currentLine.split(';').map(c => c.trim().replace(/^"|"$/g, ''));
        if (cols.length < 10) continue;

        rawRows.push({
            flight: cols[0], route: cols[1], statement: cols[2],
            startLoadStr: cols[3], endLoadStr: cols[4], departureStr: cols[5], arrivalStr: cols[6],
            nodeA: cols[7], nodeB: cols[8], vehicle: cols[9]
        });
    }

    const groupedByFlight = {};
    rawRows.forEach(row => {
        if (!groupedByFlight[row.flight]) groupedByFlight[row.flight] = [];
        groupedByFlight[row.flight].push(row);
    });

    for (const flightId in groupedByFlight) {
        const group = groupedByFlight[flightId];
        
        group.sort((a, b) => {
            const dA = parseCSVDateTime(a.startLoadStr);
            const dB = parseCSVDateTime(b.startLoadStr);
            return (dA ? dA.getTime() : 0) - (dB ? dB.getTime() : 0);
        });

        group.forEach((row, index) => {
            const containerOrder = index + 1; 

            const dStartLoad = parseCSVDateTime(row.startLoadStr);
            const dEndLoad = parseCSVDateTime(row.endLoadStr);
            const dDeparture = parseCSVDateTime(row.departureStr);
            const dArrival = parseCSVDateTime(row.arrivalStr);

            const yardAConfig = (typeof yardDictionary !== 'undefined') ? yardDictionary[row.nodeA] : null;
            const yardBConfig = (typeof yardDictionary !== 'undefined') ? yardDictionary[row.nodeB] : null;

            row.yardA = yardAConfig ? yardAConfig.yard : "Невідомий автодвір";
            row.yardB = yardBConfig ? yardBConfig.yard : "Невідомий автодвір";
            row.containerOrder = containerOrder;

            let normPlacementBufferA = yardAConfig ? (yardAConfig.factPlacementBuffer || 0) : 0;
            let normLeaveBufferA = yardAConfig ? (yardAConfig.factLeaveBuffer || 0) : 0;

            const dPlacementA = dStartLoad ? modifyMinutes(dStartLoad, -normPlacementBufferA) : null;
            const dRampLeaveA = dDeparture ? modifyMinutes(dDeparture, -normLeaveBufferA) : null;

            row.calculatedPlacement = dPlacementA;
            row.calculatedRampLeave = dRampLeaveA;

            let normPrepB = 0;
            let normUnloadB = 0;
            if (yardBConfig) {
                normPrepB = (containerOrder === 1) ? (yardBConfig.first || 0) : (yardBConfig.second || 0);
                normUnloadB = yardBConfig.unload || 0;
            }

            const dPlacementB = dArrival ? modifyMinutes(dArrival, normPrepB) : null;
            const dRampLeaveB = dPlacementB ? modifyMinutes(dPlacementB, normUnloadB) : null;

            row.calculatedUnloadStart = dPlacementB;
            row.calculatedUnloadEnd = dRampLeaveB;

            actualFlightsData.push(row);

            if (row.yardA !== "Невідомий автодвір") {
                if (dPlacementA) {
                    factCalculatedEvents.push({ yard: row.yardA, flight: row.flight, vehicle: row.vehicle, eventType: "1. Постановка", dateTime: dPlacementA });
                    recordMatrixOp(row.yardA, dPlacementA, "op1");
                }
                if (dRampLeaveA) {
                    factCalculatedEvents.push({ yard: row.yardA, flight: row.flight, vehicle: row.vehicle, eventType: "2. Забір", dateTime: dRampLeaveA });
                    recordMatrixOp(row.yardA, dRampLeaveA, "op2");
                }
            }

            if (row.yardB !== "Невідомий автодвір") {
                if (dPlacementB) {
                    factCalculatedEvents.push({ yard: row.yardB, flight: row.flight, vehicle: row.vehicle, eventType: "3. Постановка", dateTime: dPlacementB });
                    recordMatrixOp(row.yardB, dPlacementB, "op3");
                }
                if (dRampLeaveB) {
                    factCalculatedEvents.push({ yard: row.yardB, flight: row.flight, vehicle: row.vehicle, eventType: "4. Забір", dateTime: dRampLeaveB });
                    recordMatrixOp(row.yardB, dRampLeaveB, "op4");
                }
            }
        });
    }

    factCalculatedEvents.sort((a, b) => a.dateTime.getTime() - b.dateTime.getTime());
}

function recordMatrixOp(yard, dateObj, opKey) {
    const dateStr = formatDateOnly(dateObj);
    const hour = dateObj.getHours();

    if (!factOpsMatrix[yard]) factOpsMatrix[yard] = {};
    if (!factOpsMatrix[yard][dateStr]) {
        factOpsMatrix[yard][dateStr] = Array(24).fill(null).map(() => ({ total: 0, op1: 0, op2: 0, op3: 0, op4: 0 }));
    }
    factOpsMatrix[yard][dateStr][hour].total++;
    factOpsMatrix[yard][dateStr][hour][opKey]++;
}

function updateFactYardsDropdown() {
    const select = document.getElementById('factYardSelect');
    if (!select) return;

    const currentVal = select.value;
    select.innerHTML = '<option value="" disabled selected>-- Оберіть автодвір --</option>';

    const yardsSet = new Set();
    actualFlightsData.forEach(f => {
        if (f.yardA !== "Невідомий автодвір") yardsSet.add(f.yardA);
        if (f.yardB !== "Невідомий автодвір") yardsSet.add(f.yardB);
    });

    const sortedYards = Array.from(yardsSet).sort();
    sortedYards.forEach(yard => {
        const opt = document.createElement('option');
        opt.value = opt.textContent = yard;
        select.appendChild(opt);
    });

    if (currentVal && yardsSet.has(currentVal)) {
        select.value = currentVal;
    } else if (sortedYards.length > 0) {
        select.value = sortedYards[0];
    }
}

function renderFactDashboard() {
    const wrapper = document.getElementById('factTableWrapper');
    const selectedYard = document.getElementById('factYardSelect').value;

    if (!wrapper) return;

    // ПРОВЕРКА ИСПРАВЛЕНА: Проверяем наличие либо сырых CSV рейсов, либо загруженной из БД матрицы
    const hasDbMatrix = factOpsMatrix[selectedYard] && Object.keys(factOpsMatrix[selectedYard]).length > 0;

    if (actualFlightsData.length === 0 && !hasDbMatrix) {
        wrapper.innerHTML = "<p class='disabled'>Будь ласка, завантажте CSV-файл з фактичними даними або натисніть 'Завантажити з бази'.</p>";
        return;
    }

    if (!selectedYard) {
        wrapper.innerHTML = "<p class='disabled'>Оберіть автодвір зі списку для відображення аналітики.</p>";
        return;
    }

    // Рендерим аккордеоны. Если загружено из базы (нет детальных рейсов), автоматически сворачиваем блоки 2 и 3
    wrapper.innerHTML = `
        <div class="fact-accordion" id="accBlockMatrix">
            <div class="fact-accordion-header">Розрахунок</div>
            <div class="fact-accordion-content" id="factContentMatrix"></div>
        </div>

        <div class="fact-accordion ${actualFlightsData.length === 0 ? 'collapsed' : ''}" id="accBlockEvents">
            <div class="fact-accordion-header">Події ${actualFlightsData.length === 0 ? '(доступно тільки при завантаженні CSV)' : ''}</div>
            <div class="fact-accordion-content" id="factContentEvents"></div>
        </div>

        <div class="fact-accordion ${actualFlightsData.length === 0 ? 'collapsed' : ''}" id="accBlockFlights">
            <div class="fact-accordion-header">Рейси ${actualFlightsData.length === 0 ? '(доступно тільки при завантаженні CSV)' : ''}</div>
            <div class="fact-accordion-content" id="factContentFlights"></div>
        </div>
    `;

    document.querySelectorAll('.fact-accordion-header').forEach(header => {
        header.addEventListener('click', () => {
            header.parentElement.classList.toggle('collapsed');
        });
    });

    fillCalculatedFleetMatrix(selectedYard);
    fillEventsContent(selectedYard);
    fillFlightsContent(selectedYard);
}

// Побудова відфільтрованої горизонтальної матриці + графіки
function fillCalculatedFleetMatrix(yard) {
    const container = document.getElementById('factContentMatrix');
    const dates = getFactFilteredDates(yard);

    if (dates.length === 0) {
        container.innerHTML = "<p class='disabled'>Немає розрахованих операцій за обраний період дат.</p>";
        return;
    }

    const availK = (typeof fleetDictionary !== 'undefined' && fleetDictionary[yard]) ? fleetDictionary[yard].kamag : 0;
    const availM = (typeof fleetDictionary !== 'undefined' && fleetDictionary[yard]) ? fleetDictionary[yard].man : 0;

    const { hourlyRequirements, totalK, totalM, yardNorms } = getFactFleetRequirements(yard, dates);

    let html = `<table class="kamag-table"><thead><tr><th style="min-width: 140px;">ТЗ / Години</th>`;
    dates.forEach(d => {
        html += `<th colspan="25" style="text-align:center; font-weight:bold; background-color:#e9ecef; border-left:2px solid #6c757d; border-right:2px solid #6c757d;">${d}</th>`;
    });
    html += `<th style="text-align:center;">Всього, год</th></tr><tr><th></th>`;
    dates.forEach(() => {
        for (let h = 0; h < 24; h++) {
            html += `<th class="kamag-header-vertical" style="${h === 0 ? 'border-left: 2px solid #6c757d;' : ''}">${h}:00</th>`;
        }
        html += `<th style="text-align:center; font-weight:bold; background-color:#dee2e6; border-right:2px solid #6c757d;">Σ</th>`;
    });
    html += `<th></th></tr></thead><tbody>`;

    // КАМАГи
    for (let k = 1; k <= totalK; k++) {
        const isVirtual = k > availK;
        html += `<tr><td style="font-weight:bold;">Kamag ${k}${isVirtual ? ' (дод.)' : ''}</td>`;
        let totalRowSum = 0;
        dates.forEach(d => {
            let dailySum = 0;
            for (let h = 0; h < 24; h++) {
                const activeUnits = hourlyRequirements[d][h].kamag;
                const isActive = k <= activeUnits;
                const border = h === 0 ? "style='border-left: 2px solid #6c757d;'" : "";
                
                if (isActive) {
                    let cellClass = isVirtual ? "kamag-cell kamag-active-virtual" : "kamag-cell kamag-active";
                    html += `<td class="${cellClass}" ${border}>1</td>`;
                    dailySum++; totalRowSum++;
                } else {
                    html += `<td class="kamag-cell" ${border}></td>`;
                }
            }
            html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right:2px solid #6c757d;">${dailySum || ''}</td>`;
        });
        html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalRowSum}</td></tr>`;
    }

    // МАНи
    for (let m = 1; m <= totalM; m++) {
        const isVirtual = m > availM;
        html += `<tr><td style="font-weight:bold;">Маневровий ${m}${isVirtual ? ' (дод.)' : ''}</td>`;
        let totalRowSum = 0;
        dates.forEach(d => {
            let dailySum = 0;
            for (let h = 0; h < 24; h++) {
                const activeUnits = hourlyRequirements[d][h].man;
                const isActive = m <= activeUnits;
                const border = h === 0 ? "style='border-left: 2px solid #6c757d;'" : "";
                
                if (isActive) {
                    let cellClass = isVirtual ? "kamag-cell kamag-active-virtual" : "kamag-cell kamag-active";
                    html += `<td class="${cellClass}" ${border}>1</td>`;
                    dailySum++; totalRowSum++;
                } else {
                    html += `<td class="kamag-cell" ${border}></td>`;
                }
            }
            html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right:2px solid #6c757d;">${dailySum || ''}</td>`;
        });
        html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalRowSum}</td></tr>`;
    }

    // Задіяно ТЗ підсумки
    const summaryRows = [
        { label: "Задіяно фіз. КАМАГ", type: 'physK' },
        { label: "Задіяно вірт. КАМАГ", type: 'virtK' },
        { label: "Задіяно фіз. МАН", type: 'physM' },
        { label: "Задіяно вірт. МАН", type: 'virtM' }
    ];

    summaryRows.forEach(row => {
        let rowClass = "class='fleet-summary-row'";
        if (row.type === 'physK') rowClass = "class='fleet-summary-top'";
        if (row.type === 'virtM') rowClass = "class='fleet-summary-bottom'";

        html += `<tr ${rowClass}><td style="font-weight:bold;">${row.label}</td>`;
        let totalRowSum = 0;

        dates.forEach(d => {
            let dailySum = 0;
            for (let h = 0; h < 24; h++) {
                const req = hourlyRequirements[d][h];
                let val = 0;

                if (row.type === 'physK') val = Math.min(req.kamag, availK);
                if (row.type === 'virtK') val = Math.max(0, req.kamag - availK);
                if (row.type === 'physM') val = Math.min(req.man, availM);
                if (row.type === 'virtM') val = Math.max(0, req.man - availM);

                const border = h === 0 ? "border-left: 2px solid #6c757d;" : "";
                html += `<td class="kamag-cell" style="${border} background-color:#fff9c4; font-weight:bold;">${val || ''}</td>`;
                dailySum += val; totalRowSum += val;
            }
            html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right:2px solid #6c757d;">${dailySum || ''}</td>`;
        });
        html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalRowSum}</td></tr>`;
    });

    html += `<tr><td style="font-weight:bold; height:15px; background-color:#f8fafc;" colspan="${1 + dates.length * 25}"></td></tr>`;
    
    // Всього операцій
    html += `<tr><td style="font-weight:bold;">Всього операцій</td>`;
    let totalGlobalOps = 0;
    dates.forEach(d => {
        let dailySum = 0;
        for (let h = 0; h < 24; h++) {
            const val = hourlyRequirements[d][h].ops;
            const border = h === 0 ? "style='border-left: 2px solid #6c757d;'" : "";
            html += `<td class="kamag-cell" style="background-color:#fff9c4; font-weight:bold;" ${border}>${val || ''}</td>`;
            dailySum += val; totalGlobalOps += val;
        }
        html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right:2px solid #6c757d;">${dailySum || ''}</td>`;
    });
    html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalGlobalOps}</td></tr>`;

    // Непокриті фіз. флот
    html += `<tr><td style="font-weight:bold;">Непокриті (фіз. флот)</td>`;
    let totalUncoveredPhys = 0;
    dates.forEach(d => {
        let dailySum = 0;
        for (let h = 0; h < 24; h++) {
            const req = hourlyRequirements[d][h];
            const physCap = (Math.min(req.kamag, availK) * yardNorms.k) + (Math.min(req.man, availM) * yardNorms.m);
            const val = Math.max(0, req.ops - physCap);
            const border = h === 0 ? "style='border-left: 2px solid #6c757d;'" : "";
            
            if (val > 0) {
                html += `<td class="kamag-cell" ${border}><span class="uncovered-alert">${val}</span></td>`;
                dailySum += val; totalUncoveredPhys += val;
            } else {
                html += `<td class="kamag-cell" ${border}></td>`;
            }
        }
        html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right:2px solid #6c757d;">${dailySum || ''}</td>`;
    });
    html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalUncoveredPhys}</td></tr>`;

    // Непокриті залишок
    html += `<tr><td style="font-weight:bold;">Непокриті (залишок)</td>`;
    let totalUncoveredAbs = 0;
    dates.forEach(d => {
        let dailySum = 0;
        for (let h = 0; h < 24; h++) {
            const req = hourlyRequirements[d][h];
            const totalCap = (req.kamag * yardNorms.k) + (req.man * yardNorms.m);
            const val = Math.max(0, req.ops - totalCap);
            const border = h === 0 ? "style='border-left: 2px solid #6c757d;'" : "";
            
            if (val > 0) {
                html += `<td class="kamag-cell" ${border}><span class="uncovered-alert">${val}</span></td>`;
                dailySum += val; totalUncoveredAbs += val;
            } else {
                html += `<td class="kamag-cell" ${border}></td>
`;
            }
        }
        html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right:2px solid #6c757d;">${dailySum || ''}</td>`;
    });
    html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalUncoveredAbs}</td></tr>`;

    // Рендер Графіків
    html += `<tr><td class="fact-chart-row-label">Графік</td>`;
    dates.forEach((d, idx) => {
        html += `<th colspan="24" class="fact-chart-td"><div class="fact-chart-wrapper"><canvas id="fact_chart_${idx}"></canvas></div></th><td class="fact-chart-td-space"></td>`;
    });
    html += `<td class="fact-chart-td-end"></td></tr>`;

    html += `</tbody></table>`;
    container.innerHTML = html;

    if (window.myFactDayCharts) window.myFactDayCharts.forEach(c => c.destroy());
    window.myFactDayCharts = [];

    dates.forEach((d, index) => {
        const ctx = document.getElementById(`fact_chart_${index}`);
        if (!ctx) return;
        const parentDiv = ctx.parentElement;
        
        ctx.width = parentDiv.clientWidth; ctx.height = 135; 

        const chartLabels = [], opsData = [], capacityData = [];
        for (let h = 0; h < 24; h++) {
            chartLabels.push(`${h}:00`);
            opsData.push(hourlyRequirements[d][h].ops);
            
            const req = hourlyRequirements[d][h];
            const cap = (req.kamag * yardNorms.k) + (req.man * yardNorms.m);
            capacityData.push(cap);
        }

        window.myFactDayCharts.push(new Chart(ctx, {
            type: 'bar',
            data: { labels: chartLabels, datasets: [
                { type: 'line', label: 'Потужність', data: capacityData, borderColor: '#0d47a1', backgroundColor: '#0d47a1', borderWidth: 2, tension: 0.3, pointRadius: 2, yAxisID: 'y' },
                { type: 'bar', label: 'Операції', data: opsData, backgroundColor: 'rgba(255, 193, 7, 0.7)', borderColor: '#ffaa00', borderWidth: 1, borderRadius: 2, yAxisID: 'y' }
            ]},
            options: { 
                animation: false, responsive: false, maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false }, 
                scales: { x: { display: false }, y: { type: 'linear', display: false, beginAtZero: true } }, 
                plugins: { 
                    legend: { display: true, position: 'bottom', align: 'center', labels: { boxWidth: 12, font: { size: 10 }, padding: 4 } }, 
                    tooltip: { callbacks: { title: (items) => `${d} ${items[0].label}`, label: (item) => `${item.dataset.label}: ${item.raw}` } } 
                }, 
                layout: { padding: 0 } 
            }
        }));
    });
}

// Наповнення Блоку 2: Стрічка подій
function fillEventsContent(yard) {
    const container = document.getElementById('factContentEvents');
    const allowedDates = getFactFilteredDates(yard);
    const filteredEvents = factCalculatedEvents.filter(ev => ev.yard === yard && allowedDates.includes(formatDateOnly(ev.dateTime)));

    if (filteredEvents.length === 0) {
        container.innerHTML = "<p class='disabled'>Подій за обраний період не знайдено.</p>";
        return;
    }

    let html = `<table><thead><tr>
        <th>День</th><th>Час події</th><th>Рейс</th><th>Номер ТЗ</th><th>Назва операції</th>
    </tr></thead><tbody>`;

    filteredEvents.forEach(ev => {
        let opColor = (ev.eventType.startsWith("1") || ev.eventType.startsWith("3")) ? "color:#15803d;" : "color:#b45309;";

        html += `<tr>
            <td style="font-weight:bold;">${formatDateOnly(ev.dateTime)}</td>
            <td>${formatTimeOnly(ev.dateTime)}</td>
            <td>${ev.flight}</td>
            <td style="font-weight:bold; color:#0369a1;">${ev.vehicle}</td>
            <td style="font-weight:bold; ${opColor}">${ev.eventType}</td>
        </tr>`;
    });
    html += `</tbody></table>`;
    container.innerHTML = html;
}

// Наповнення Блоку 3: Реєстр рейсів
function fillFlightsContent(yard) {
    const container = document.getElementById('factContentFlights');
    const allowedDates = getFactFilteredDates(yard);
    
    const filtered = actualFlightsData.filter(f => {
        const isMatchYard = f.yardA === yard || f.yardB === yard;
        const flightDateStr = formatDateOnly(f.calculatedPlacement);
        return isMatchYard && allowedDates.includes(flightDateStr);
    });

    if (filtered.length === 0) {
        container.innerHTML = "<p class='disabled'>Рейсів за обраний період не знайдено.</p>";
        return;
    }

    let html = `<table><thead><tr>
        <th>Рейс</th><th>Порядок</th><th>Маршрут</th><th>Відомість</th><th>ТЗ</th>
        <th>Автодвір А</th><th>Вузол А</th><th>1. Постановка (Факт)</th><th>Початок скан.</th><th>Кінець скан.</th><th>2. Забір (Факт)</th><th>Виїзд</th>
        <th>Автодвір Б</th><th>Вузол Б</th><th>Приїзд</th><th>3. Постановка (Факт)</th><th>4. Забір (Факт)</th>
    </tr></thead><tbody>`;

    filtered.forEach(f => {
        html += `<tr>
            <td>${f.flight}</td>
            <td style="text-align:center; font-weight:bold; color:#64748b;">${f.containerOrder}</td>
            <td class="col-route">${f.route}</td>
            <td>${f.statement}</td>
            <td style="font-weight:bold; color:#0369a1;">${f.vehicle}</td>
            <td>${f.yardA}</td><td>${f.nodeA}</td>
            <td style="color:#15803d; font-weight:bold;">${formatTimeOnly(f.calculatedPlacement)}</td>
            <td>${f.startLoadStr ? f.startLoadStr.split(' ')[1] : '—'}</td>
            <td>${f.endLoadStr ? f.endLoadStr.split(' ')[1] : '—'}</td>
            <td style="color:#b45309; font-weight:bold;">${formatTimeOnly(f.calculatedRampLeave)}</td>
            <td>${f.departureStr ? f.departureStr.split(' ')[1] : '—'}</td>
            <td>${f.yardB}</td><td>${f.nodeB}</td>
            <td>${f.arrivalStr ? f.arrivalStr.split(' ')[1] : '—'}</td>
            <td style="color:#15803d; font-weight:bold;">${formatTimeOnly(f.calculatedUnloadStart)}</td>
            <td style="color:#b45309; font-weight:bold;">${formatTimeOnly(f.calculatedUnloadEnd)}</td>
        </tr>`;
    });
    html += `</tbody></table>`;
    container.innerHTML = html;
}

async function loadSavedFactYardsList() {
    try {
        const response = await fetch(RESULTS_SCRIPT_URL + '?action=getFactYards');
        const data = await response.json();
        
        if (data.yards && data.yards.length > 0) {
            const select = document.getElementById('factYardSelect');
            const currentVal = select.value;
            const existingYards = new Set(Array.from(select.options).map(o => o.value));
            
            data.yards.forEach(y => {
                if (!existingYards.has(y)) {
                    const opt = document.createElement('option');
                    opt.value = opt.textContent = y;
                    select.appendChild(opt);
                }
            });
            if (currentVal) select.value = currentVal;
            document.getElementById('factFileStatus').innerText = "Список збережених автодворів (Факт) підвантажено.";
        }
    } catch (e) {
        console.error("Помилка завантаження списку автодворів Факт:", e);
    }
}

async function loadFactFromGoogle() {
    const yard = document.getElementById('factYardSelect').value;
    if (!yard) return alert("Оберіть автодвір!");
    const btn = document.getElementById('loadFactGoogleYardBtn');
    const originalText = btn.innerText;
    btn.innerText = "⏳...";

    try {
        const response = await fetch(`${RESULTS_SCRIPT_URL}?action=getFactAggregatedData&yard=${encodeURIComponent(yard)}`);
        const data = await response.json();

        if (data.savedRows && data.savedRows.length > 0) {
            factCalculatedEvents = []; 
            actualFlightsData = []; 
            factOpsMatrix[yard] = {};

            data.savedRows.forEach(row => {
                let [y, day, hour, fleetCountStr, ops] = row;
                
                let dayStr = String(day);
                if (dayStr.includes('T') && dayStr.includes('Z')) {
                    const d = new Date(dayStr);
                    dayStr = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
                }

                if (!factOpsMatrix[yard][dayStr]) {
                    factOpsMatrix[yard][dayStr] = Array(24).fill(null).map(() => ({ total: 0, op1: 0, op2: 0, op3: 0, op4: 0 }));
                }

                factOpsMatrix[yard][dayStr][hour] = {
                    total: parseInt(ops, 10) || 0,
                    op1: 0, op2: 0, op3: 0, op4: 0
                };
            });

            renderFactDashboard();
            document.getElementById('factFileStatus').innerText = `Дані (Факт) ${yard} завантажені!`;
        } else {
            alert("Даних для цього автодвору не знайдено в базі.");
        }
    } catch (e) {
        alert("Помилка завантаження з бази");
        console.error(e);
    } finally {
        btn.innerText = originalText;
    }
}

async function saveFactToGoogle() {
    const yard = document.getElementById('factYardSelect').value;
    if (!yard) return alert("Оберіть автодвір!");

    const btn = document.getElementById('saveFactGoogleBtn');
    btn.innerText = "⏳ Збереження...";

    const days = getFactFilteredDates(yard);
    if (days.length === 0) {
        alert("Немає даних у вибраному періоді для збереження!");
        btn.innerText = "Зберегти (поточний)";
        return;
    }

    const { hourlyRequirements, totalK, totalM } = getFactFleetRequirements(yard, days);
    const aggregatedRows = [];

    days.forEach(day => {
        for (let h = 0; h < 24; h++) {
            const opsCount = (factOpsMatrix[yard] && factOpsMatrix[yard][day]) ? factOpsMatrix[yard][day][h].total : 0;
            const req = hourlyRequirements[day][h];
            
            const kArr = Array(totalK).fill(0).map((_, i) => i < req.kamag ? 1 : 0);
            const mArr = Array(totalM).fill(0).map((_, i) => i < req.man ? 1 : 0);
            const stateString = `${kArr.join(',')}|${mArr.join(',')}`;

            if (opsCount > 0 || req.kamag > 0 || req.man > 0) {
                aggregatedRows.push([yard, day, h, stateString, opsCount]);
            }
        }
    });

    try {
        await fetch(RESULTS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'saveFactAggregated', yard: yard, rows: aggregatedRows, dates: days })
        });
        btn.innerText = "✅ Збережено!";
    } catch (e) {
        btn.innerText = "❌ Помилка";
    }
    setTimeout(() => btn.innerText = "Зберегти (поточний)", 3000);
}

async function saveAllFactToGoogle() {
    const yards = Object.keys(factOpsMatrix);
    if (yards.length === 0) return alert("Немає розрахованих даних для збереження!");

    const btn = document.getElementById('saveAllFactGoogleBtn');
    btn.innerText = "⏳ Збереження...";
    btn.disabled = true;

    const aggregatedRows = [];
    const allFilteredDays = new Set();

    yards.forEach(yard => {
        const days = getFactFilteredDates(yard);
        
        if (days.length > 0) {
            const { hourlyRequirements, totalK, totalM } = getFactFleetRequirements(yard, days);
            
            days.forEach(day => {
                allFilteredDays.add(day);
                for (let h = 0; h < 24; h++) {
                    const opsCount = (factOpsMatrix[yard] && factOpsMatrix[yard][day]) ? factOpsMatrix[yard][day][h].total : 0;
                    const req = hourlyRequirements[day][h];

                    const kArr = Array(totalK).fill(0).map((_, i) => i < req.kamag ? 1 : 0);
                    const mArr = Array(totalM).fill(0).map((_, i) => i < req.man ? 1 : 0);
                    const stateString = `${kArr.join(',')}|${mArr.join(',')}`;

                    if (opsCount > 0 || req.kamag > 0 || req.man > 0) {
                        aggregatedRows.push([yard, day, h, stateString, opsCount]);
                    }
                }
            });
        }
    });

    if (aggregatedRows.length === 0) {
        alert("Немає розрахованих даних у вибраному періоді дат!");
        btn.innerText = "Зберегти ВСІ";
        btn.disabled = false;
        return;
    }

    try {
        await fetch(RESULTS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'saveAllFactAggregated', yards: yards, rows: aggregatedRows, dates: Array.from(allFilteredDays) })
        });
        btn.innerText = "✅ Всі збережено!";
    } catch (e) {
        console.error(e);
        btn.innerText = "❌ Помилка";
    }
    setTimeout(() => {
        btn.innerText = "Зберегти ВСІ";
        btn.disabled = false;
    }, 3000);
}