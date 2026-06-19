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
    });

    // А список дворов загружаем один раз при старте страницы
    setTimeout(() => {
        updateFactYardsDropdown();
        loadSavedFactYardsList();
    }, 1000);

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

// Допоміжна функція єдиного розрахунку флоту
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

// ОБНОВЛЕННЫЙ ПАРСИНГ: Извлечение причины создания рейса
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
        if (cols.length < 11) continue;

        rawRows.push({
            flight: cols[0], 
            reason: cols[1] || '—', 
            route: cols[2], 
            statement: cols[3],
            startLoadStr: cols[4], 
            endLoadStr: cols[5], 
            departureStr: cols[6], 
            arrivalStr: cols[7],
            nodeA: cols[8], 
            nodeB: cols[9], 
            vehicle: cols[10],
            container: cols[11] || '—'
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

            // Считываем настройки нормативов и буферов
            let normPlacementBufferA = yardAConfig ? (yardAConfig.factPlacementBuffer || 0) : 0;
            let normFirstContainerBufferA = yardAConfig ? (yardAConfig.factFirstContainerBuffer || 0) : 0;

            // 1. Постановка на Автодворе А (всегда от Почала сканування назад)
            const dPlacementA = dStartLoad ? modifyMinutes(dStartLoad, -normPlacementBufferA) : null;
            
            // 2. Забор на Автодворе А (Универсальная логика для всех контейнеров)
            // Теперь строго: Конец сканирования конкретного контейнера + норматив из колонки P
            const dRampLeaveA = dEndLoad ? modifyMinutes(dEndLoad, normFirstContainerBufferA) : null;

            row.calculatedPlacement = dPlacementA;
            row.calculatedRampLeave = dRampLeaveA;

            // Расчет нормативов для Автодвора Б (Прибытие)
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

            // Логирование событий и накопление матрицы операций
            if (row.yardA !== "Невідомий автодвір") {
                if (dPlacementA) {
                    factCalculatedEvents.push({ yard: row.yardA, flight: row.flight, reason: row.reason, vehicle: row.vehicle, container: row.container, eventType: "1. Постановка", dateTime: dPlacementA });
                    recordMatrixOp(row.yardA, dPlacementA, "op1");
                }
                if (dRampLeaveA) {
                    factCalculatedEvents.push({ yard: row.yardA, flight: row.flight, reason: row.reason, vehicle: row.vehicle, container: row.container, eventType: "2. Забір", dateTime: dRampLeaveA });
                    recordMatrixOp(row.yardA, dRampLeaveA, "op2");
                }
            }

            if (row.yardB !== "Невідомий автодвір") {
                if (dPlacementB) {
                    factCalculatedEvents.push({ yard: row.yardB, flight: row.flight, reason: row.reason, vehicle: row.vehicle, container: row.container, eventType: "3. Постановка", dateTime: dPlacementB });
                    recordMatrixOp(row.yardB, dPlacementB, "op3");
                }
                if (dRampLeaveB) {
                    factCalculatedEvents.push({ yard: row.yardB, flight: row.flight, reason: row.reason, vehicle: row.vehicle, container: row.container, eventType: "4. Забір", dateTime: dRampLeaveB });
                    recordMatrixOp(row.yardB, dRampLeaveB, "op4");
                }
            }
        });
    }

    // Финальная сортировка массива событий по времени
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

    const hasDbMatrix = factOpsMatrix[selectedYard] && Object.keys(factOpsMatrix[selectedYard]).length > 0;

    if (actualFlightsData.length === 0 && !hasDbMatrix) {
        wrapper.innerHTML = "<p class='disabled'>Будь ласка, завантажте CSV-файл з фактичними даними або натисніть 'Завантажити з бази'.</p>";
        return;
    }

    if (!selectedYard) {
        wrapper.innerHTML = "<p class='disabled'>Оберіть автодвір зі списку для відображення аналітики.</p>";
        return;
    }

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

    // СЛОВАРЬ ДНЕЙ НЕДЕЛИ
    const dayNamesShort = ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];

    let html = `<table class="kamag-table"><thead><tr><th style="min-width: 140px;">ТЗ / Години</th>`;
    
    // ДОБАВЛЯЕМ ДЕНЬ НЕДЕЛИ В ЦИКЛ ВЫВОДА ДАТ
    dates.forEach(d => {
        const [dd, mm, yyyy] = d.split('.').map(Number);
        const dateObj = new Date(yyyy, mm - 1, dd);
        const dayName = dayNamesShort[dateObj.getDay()];

        html += `<th colspan="25" style="text-align:center; font-weight:bold; background-color:#e9ecef; border-left:2px solid #6c757d; border-right:2px solid #6c757d; padding: 4px 0;">
            ${d}<br><span style="font-size: 11px; font-weight: normal; color: #6c757d;">${dayName}</span>
        </th>`;
    });
    html += `<th style="text-align:center;">Всього, год</th></tr><tr><th></th>`;
    dates.forEach(() => {
        for (let h = 0; h < 24; h++) {
            html += `<th class="kamag-header-vertical" style="${h === 0 ? 'border-left: 2px solid #6c757d;' : ''}">${h}:00</th>`;
        }
        html += `<th style="text-align:center; font-weight:bold; background-color:#dee2e6; border-right:2px solid #6c757d;">Σ</th>`;
    });
    html += `<th></th></tr></thead><tbody>`;

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
            html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right: 2px solid #6c757d;">${dailySum || ''}</td>`;
        });
        html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalRowSum}</td></tr>`;
    });

    html += `<tr><td style="font-weight:bold; height:15px; background-color:#f8fafc;" colspan="${1 + dates.length * 25}"></td></tr>`;
    
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
        html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right: 2px solid #6c757d;">${dailySum || ''}</td>`;
    });
    html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalGlobalOps}</td></tr>`;

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
        html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right: 2px solid #6c757d;">${dailySum || ''}</td>`;
    });
    html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalUncoveredPhys}</td></tr>`;

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
                html += `<td class="kamag-cell" ${border}></td>`;
            }
        }
        html += `<td style="text-align:center; font-weight:bold; background-color:#f1f3f5; border-right: 2px solid #6c757d;">${dailySum || ''}</td>`;
    });
    html += `<td style="text-align:center; font-weight:bold; background-color:#e9ecef;">${totalUncoveredAbs}</td></tr>`;

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
            capacityData.push((req.kamag * yardNorms.k) + (req.man * yardNorms.m));
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

// ОБНОВЛЕННЫЙ БЛОК 2: Добавлена колонка Контейнер (data-col="4")
function fillEventsContent(yard) {
    const container = document.getElementById('factContentEvents');
    const allowedDates = getFactFilteredDates(yard);
    const filteredEvents = factCalculatedEvents.filter(ev => ev.yard === yard && allowedDates.includes(formatDateOnly(ev.dateTime)));

    if (filteredEvents.length === 0) {
        container.innerHTML = "<p class='disabled'>Подій за обраний період не знайдено.</p>";
        return;
    }

    let html = `<style>
        #factEventsTable th, #factEventsTable td { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; padding: 4px 6px; }
    </style>
    <table id="factEventsTable" style="table-layout: fixed; width: 1000px;"><thead><tr>
        <th style="width:130px;">Автодвір<br><input type="text" class="fact-col-filter filter-input" data-col="0"></th>
        <th style="width:90px;">День<br><input type="text" class="fact-col-filter filter-input" data-col="1"></th>
        <th style="width:80px;">Час події<br><input type="text" class="fact-col-filter filter-input" data-col="2"></th>
        <th style="width:80px;">Рейс<br><input type="text" class="fact-col-filter filter-input" data-col="3"></th>
        <th style="width:150px;">Причина створення<br><input type="text" class="fact-col-filter filter-input" data-col="4"></th>
        <th style="width:100px;">Номер ТЗ<br><input type="text" class="fact-col-filter filter-input" data-col="5"></th>
        <th style="width:120px;">Контейнер<br><input type="text" class="fact-col-filter filter-input" data-col="6"></th>
        <th style="width:150px;">Назва операції<br><input type="text" class="fact-col-filter filter-input" data-col="7"></th>
    </tr></thead><tbody>`;

    filteredEvents.forEach(ev => {
        let opColor = (ev.eventType.startsWith("1") || ev.eventType.startsWith("3")) ? "color:#15803d;" : "color:#b45309;";
        html += `<tr>
            <td style="font-weight:bold; color:#475569;" title="${ev.yard}">${ev.yard}</td>
            <td style="font-weight:bold;" title="${formatDateOnly(ev.dateTime)}">${formatDateOnly(ev.dateTime)}</td>
            <td title="${formatTimeOnly(ev.dateTime)}">${formatTimeOnly(ev.dateTime)}</td>
            <td title="${ev.flight}">${ev.flight}</td>
            <td style="color:#64748b; font-weight:500;" title="${ev.reason || '—'}">${ev.reason || '—'}</td>
            <td style="font-weight:bold; color:#0369a1;" title="${ev.vehicle}">${ev.vehicle}</td>
            <td style="font-weight:bold; color:#475569;" title="${ev.container || '—'}">${ev.container || '—'}</td>
            <td style="font-weight:bold; ${opColor}" title="${ev.eventType}">${ev.eventType}</td>
        </tr>`;
    });
    html += `</tbody></table>`;
    container.innerHTML = html;
    
    attachFactLiveFilters();
}

// ОБНОВЛЕННЫЙ БЛОК 3: Добавлена колонка Причина створення (data-col="1")
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

    let html = `<style>
        #factFlightsTable th, #factFlightsTable td { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; padding: 4px 6px; }
    </style>
    <table id="factFlightsTable" style="table-layout: fixed; width: 1880px;"><thead><tr>
        <th style="width:60px;">Рейс<br><input type="text" class="fact-col-filter filter-input" data-col="0"></th>
        <th style="width:110px;">Дата та час<br><input type="text" class="fact-col-filter filter-input" data-col="1"></th>
        <th style="width:120px;">Причина ств.<br><input type="text" class="fact-col-filter filter-input" data-col="2"></th>
        <th style="width:50px;">Пор.<br><input type="text" class="fact-col-filter filter-input" data-col="3"></th>
        <th style="width:140px;">Маршрут<br><input type="text" class="fact-col-filter filter-input" data-col="4"></th>
        <th style="width:80px;">Відомість<br><input type="text" class="fact-col-filter filter-input" data-col="5"></th>
        <th style="width:80px;">ТЗ<br><input type="text" class="fact-col-filter filter-input" data-col="6"></th>
        <th style="width:90px;">Контейнер<br><input type="text" class="fact-col-filter filter-input" data-col="7"></th>
        <th style="width:100px;">Автодвір А<br><input type="text" class="fact-col-filter filter-input" data-col="8"></th>
        <th style="width:100px;">Вузол А<br><input type="text" class="fact-col-filter filter-input" data-col="9"></th>
        <th style="width:65px;">1. Пост.<br><input type="text" class="fact-col-filter filter-input" data-col="10"></th>
        <th style="width:65px;">Поч. скан.<br><input type="text" class="fact-col-filter filter-input" data-col="11"></th>
        <th style="width:65px;">Кін. скан.<br><input type="text" class="fact-col-filter filter-input" data-col="12"></th>
        <th style="width:65px;">2. Забір<br><input type="text" class="fact-col-filter filter-input" data-col="13"></th>
        <th style="width:65px;">Виїзд<br><input type="text" class="fact-col-filter filter-input" data-col="14"></th>
        <th style="width:100px;">Автодвір Б<br><input type="text" class="fact-col-filter filter-input" data-col="15"></th>
        <th style="width:100px;">Вузол Б<br><input type="text" class="fact-col-filter filter-input" data-col="16"></th>
        <th style="width:65px;">Приїзд<br><input type="text" class="fact-col-filter filter-input" data-col="17"></th>
        <th style="width:65px;">3. Пост.<br><input type="text" class="fact-col-filter filter-input" data-col="18"></th>
        <th style="width:65px;">4. Забір<br><input type="text" class="fact-col-filter filter-input" data-col="19"></th>
    </tr></thead><tbody>`;

    filtered.forEach(f => {
        // Формуємо об'єднану дату і час для рейсу
        let dtObj = f.calculatedPlacement || f.calculatedUnloadStart;
        let flightDateTimeStr = dtObj ? `${formatDateOnly(dtObj)} ${formatTimeOnly(dtObj)}` : '—';

        html += `<tr>
            <td title="${f.flight}">${f.flight}</td>
            <td style="font-weight:bold; color:#0f172a;" title="${flightDateTimeStr}">${flightDateTimeStr}</td>
            <td style="color:#64748b; font-weight:500;" title="${f.reason}">${f.reason}</td>
            <td style="text-align:center; font-weight:bold; color:#64748b;" title="${f.containerOrder}">${f.containerOrder}</td>
            <td class="col-route" title="${f.route}">${f.route}</td>
            <td title="${f.statement}">${f.statement}</td>
            <td style="font-weight:bold; color:#0369a1;" title="${f.vehicle}">${f.vehicle}</td>
            <td style="font-weight:bold; color:#475569;" title="${f.container || '—'}">${f.container || '—'}</td>
            <td title="${f.yardA}">${f.yardA}</td><td title="${f.nodeA}">${f.nodeA}</td>
            <td style="color:#15803d; font-weight:bold;" title="${formatTimeOnly(f.calculatedPlacement)}">${formatTimeOnly(f.calculatedPlacement)}</td>
            <td title="${f.startLoadStr ? f.startLoadStr.split(' ')[1] : '—'}">${f.startLoadStr ? f.startLoadStr.split(' ')[1] : '—'}</td>
            <td title="${f.endLoadStr ? f.endLoadStr.split(' ')[1] : '—'}">${f.endLoadStr ? f.endLoadStr.split(' ')[1] : '—'}</td>
            <td style="color:#b45309; font-weight:bold;" title="${formatTimeOnly(f.calculatedRampLeave)}">${formatTimeOnly(f.calculatedRampLeave)}</td>
            <td title="${f.departureStr ? f.departureStr.split(' ')[1] : '—'}">${f.departureStr ? f.departureStr.split(' ')[1] : '—'}</td>
            <td title="${f.yardB}">${f.yardB}</td><td title="${f.nodeB}">${f.nodeB}</td>
            <td title="${f.arrivalStr ? f.arrivalStr.split(' ')[1] : '—'}">${f.arrivalStr ? f.arrivalStr.split(' ')[1] : '—'}</td>
            <td style="color:#15803d; font-weight:bold;" title="${formatTimeOnly(f.calculatedUnloadStart)}">${formatTimeOnly(f.calculatedUnloadStart)}</td>
            <td style="color:#b45309; font-weight:bold;" title="${formatTimeOnly(f.calculatedUnloadEnd)}">${formatTimeOnly(f.calculatedUnloadEnd)}</td>
        </tr>`;
    });
    html += `</tbody></table>`;
    container.innerHTML = html;
    
    attachFactLiveFilters();
}

// === ПОВНІСТЮ ЗАМІНИТИ ФУНКЦІЮ attachFactLiveFilters У fact.js ===
function attachFactLiveFilters() {
    const wrapper = document.getElementById('factTableWrapper');
    if (!wrapper || wrapper.dataset.filtersBound === 'true') return;
    
    wrapper.dataset.filtersBound = 'true';
    
    wrapper.addEventListener('input', (e) => {
        if (!e.target.classList.contains('fact-col-filter')) return;
        
        const eventsTable = document.getElementById('factEventsTable');
        const flightsTable = document.getElementById('factFlightsTable');

        if (eventsTable && flightsTable) {
            const evFlight = eventsTable.querySelector('.fact-col-filter[data-col="3"]'); 
            const evReason = eventsTable.querySelector('.fact-col-filter[data-col="4"]'); 
            const evVehicle = eventsTable.querySelector('.fact-col-filter[data-col="5"]'); 
            
            const flFlight = flightsTable.querySelector('.fact-col-filter[data-col="0"]'); 
            const flReason = flightsTable.querySelector('.fact-col-filter[data-col="2"]'); // Зсунуто через стовпець "Дата та час"
            const flVehicle = flightsTable.querySelector('.fact-col-filter[data-col="6"]'); // Зсунуто через стовпець "Дата та час"

            if (e.target === evFlight && flFlight) flFlight.value = evFlight.value;
            if (e.target === flFlight && evFlight) evFlight.value = flFlight.value;

            if (e.target === evReason && flReason) flReason.value = evReason.value;
            if (e.target === flReason && evReason) evReason.value = flReason.value;

            if (e.target === evVehicle && flVehicle) flVehicle.value = evVehicle.value;
            if (e.target === flVehicle && evVehicle) evVehicle.value = flVehicle.value;
        }

        if (eventsTable) {
            const evInputs = eventsTable.querySelectorAll('.fact-col-filter');
            eventsTable.querySelectorAll('tbody tr').forEach(row => {
                let keep = true;
                evInputs.forEach(input => {
                    const colIdx = parseInt(input.getAttribute('data-col'), 10);
                    const val = input.value.trim().toLowerCase();
                    if (val && !row.children[colIdx].textContent.toLowerCase().includes(val)) {
                        keep = false;
                    }
                });
                row.style.display = keep ? '' : 'none';
            });
        }

        if (flightsTable) {
            const flInputs = flightsTable.querySelectorAll('.fact-col-filter');
            flightsTable.querySelectorAll('tbody tr').forEach(row => {
                let keep = true;
                flInputs.forEach(input => {
                    const colIdx = parseInt(input.getAttribute('data-col'), 10);
                    const val = input.value.trim().toLowerCase();
                    if (val && !row.children[colIdx].textContent.toLowerCase().includes(val)) {
                        keep = false;
                    }
                });
                row.style.display = keep ? '' : 'none';
            });
        }
    });
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
    } catch (e) { console.error("Помилка списку автодворів Факт:", e); }
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
            factCalculatedEvents = []; actualFlightsData = []; factOpsMatrix[yard] = {};
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
                factOpsMatrix[yard][dayStr][hour] = { total: parseInt(ops, 10) || 0, op1: 0, op2: 0, op3: 0, op4: 0 };
            });
            renderFactDashboard();
            document.getElementById('factFileStatus').innerText = `Дані (Факт) ${yard} завантажені!`;
        } else { alert("Даних для цього автодвору не знайдено в базі."); }
    } catch (e) { alert("Помилка завантаження"); console.error(e); } finally { btn.innerText = originalText; }
}

async function saveFactToGoogle() {
    const yard = document.getElementById('factYardSelect').value;
    if (!yard) return alert("Оберіть автодвір!");
    const btn = document.getElementById('saveFactGoogleBtn');
    btn.innerText = "⏳ Збереження...";
    const days = getFactFilteredDates(yard);
    if (days.length === 0) { alert("Немає даних у вибраному періоді!"); btn.innerText = "Зберегти (поточний)"; return; }
    const { hourlyRequirements, totalK, totalM } = getFactFleetRequirements(yard, days);
    const aggregatedRows = [];
    days.forEach(day => {
        for (let h = 0; h < 24; h++) {
            const opsCount = (factOpsMatrix[yard] && factOpsMatrix[yard][day]) ? factOpsMatrix[yard][day][h].total : 0;
            const req = hourlyRequirements[day][h];
            const kArr = Array(totalK).fill(0).map((_, i) => i < req.kamag ? 1 : 0);
            const mArr = Array(totalM).fill(0).map((_, i) => i < req.man ? 1 : 0);
            const stateString = `${kArr.join(',')}|${mArr.join(',')}`;
            if (opsCount > 0 || req.kamag > 0 || req.man > 0) { aggregatedRows.push([yard, day, h, stateString, opsCount]); }
        }
    });
    try {
        await fetch(RESULTS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'saveFactAggregated', yard: yard, rows: aggregatedRows, dates: days })
        });
        btn.innerText = "✅ Збережено!";
    } catch (e) { btn.innerText = "❌ Помилка"; }
    setTimeout(() => btn.innerText = "Зберегти (поточний)", 3000);
}

async function saveAllFactToGoogle() {
    const yards = Object.keys(factOpsMatrix);
    if (yards.length === 0) return alert("Немає розрахованих даних для збереження!");
    const btn = document.getElementById('saveAllFactGoogleBtn');
    btn.innerText = "⏳ Збереження..."; btn.disabled = true;
    const aggregatedRows = []; const allFilteredDays = new Set();
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
                    if (opsCount > 0 || req.kamag > 0 || req.man > 0) { aggregatedRows.push([yard, day, h, stateString, opsCount]); }
                }
            });
        }
    });
    if (aggregatedRows.length === 0) { alert("Немає розрахованих даних!"); btn.innerText = "Зберегти ВСІ"; btn.disabled = false; return; }
    try {
        await fetch(RESULTS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'saveAllFactAggregated', yards: yards, rows: aggregatedRows, dates: Array.from(allFilteredDays) })
        });
        btn.innerText = "✅ Всі збережено!";
    } catch (e) { console.error(e); btn.innerText = "❌ Помилка"; }
    setTimeout(() => { btn.innerText = "Зберегти ВСІ"; btn.disabled = false; }, 3000);
}


window.exportFactToExcel = async function(workbook) {
    const exportMode = document.getElementById('exportModeSelect').value;
    let yardsToExport = [];

    if (exportMode === 'current') {
        const currentYard = document.getElementById('factYardSelect').value;
        if (currentYard) yardsToExport.push(currentYard);
    } else {
        yardsToExport = Object.keys(factOpsMatrix).sort();
    }

    if (yardsToExport.length === 0) {
        alert("Немає автодворів для експорту!");
        return false;
    }

    let hasData = false;
    for (const yard of yardsToExport) {
        const allowedDates = getFactFilteredDates(yard);
        if (allowedDates.length === 0) continue;

        const hasFlights = actualFlightsData.some(f => 
            (f.yardA === yard || f.yardB === yard) && allowedDates.includes(formatDateOnly(f.calculatedPlacement))
        );
        const hasEvents = factCalculatedEvents.some(ev => 
            ev.yard === yard && allowedDates.includes(formatDateOnly(ev.dateTime))
        );

        if (hasFlights || hasEvents) { hasData = true; break; }
    }

    if (!hasData) {
        alert("Немає розрахованих даних (Факт) для експорту за обраний період!");
        return false;
    }

    const fillHeader = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } };
    const alignCenter = { vertical: 'middle', horizontal: 'center' };

    // Створюємо рівно два майстер-листи для накопичення даних
    const sheetFlights = workbook.addWorksheet('Рейси (Факт)');
    const flightHeaders = [
        "Автодвір (Аналітика)", "Рейс", "Дата та час", "Причина створення", "Порядок", "Маршрут", "Відомість", "ТЗ", "Контейнер",
        "Автодвір А", "Вузол А", "1. Постановка", "Початок скан.", "Кінець скан.", "2. Забір", "Виїзд",
        "Автодвір Б", "Вузол Б", "Приїзд", "3. Постановка", "4. Забір"
    ];
    const headerRowF = sheetFlights.addRow(flightHeaders);
    headerRowF.font = { bold: true };
    headerRowF.eachCell(cell => { cell.fill = fillHeader; cell.alignment = alignCenter; });

    const sheetEvents = workbook.addWorksheet('Події (Факт)');
    const eventHeaders = ["Автодвір (Аналітика)", "День", "Час події", "Рейс", "Номер ТЗ", "Контейнер", "Назва операції"];
    const headerRowE = sheetEvents.addRow(eventHeaders);
    headerRowE.font = { bold: true };
    headerRowE.eachCell(cell => { cell.fill = fillHeader; cell.alignment = alignCenter; });

    for (const yard of yardsToExport) {
        const allowedDates = getFactFilteredDates(yard);
        if (allowedDates.length === 0) continue;

        const filteredFlights = actualFlightsData.filter(f => {
            const isMatchYard = f.yardA === yard || f.yardB === yard;
            const flightDateStr = formatDateOnly(f.calculatedPlacement);
            return isMatchYard && allowedDates.includes(flightDateStr);
        });

        const filteredEvents = factCalculatedEvents.filter(ev => 
            ev.yard === yard && allowedDates.includes(formatDateOnly(ev.dateTime))
        );

        filteredFlights.forEach(f => {
            let dtObj = f.calculatedPlacement || f.calculatedUnloadStart;
            let flightDateTimeStr = dtObj ? `${formatDateOnly(dtObj)} ${formatTimeOnly(dtObj)}` : '—';

            sheetFlights.addRow([
                yard, 
                f.flight,
                flightDateTimeStr,
                f.reason,
                f.containerOrder,
                f.route,
                f.statement,
                f.vehicle,
                f.container || '—',
                f.yardA,
                f.nodeA,
                formatTimeOnly(f.calculatedPlacement),
                f.startLoadStr ? f.startLoadStr.split(' ')[1] : '—',
                f.endLoadStr ? f.endLoadStr.split(' ')[1] : '—',
                formatTimeOnly(f.calculatedRampLeave),
                f.departureStr ? f.departureStr.split(' ')[1] : '—',
                f.yardB,
                f.nodeB,
                f.arrivalStr ? f.arrivalStr.split(' ')[1] : '—',
                formatTimeOnly(f.calculatedUnloadStart),
                formatTimeOnly(f.calculatedUnloadEnd)
            ]);
        });

        filteredEvents.forEach(ev => {
            sheetEvents.addRow([
                yard, 
                formatDateOnly(ev.dateTime),
                formatTimeOnly(ev.dateTime),
                ev.flight,
                ev.reason || '—',
                ev.vehicle,
                ev.container || '—',
                ev.eventType
            ]);
        });
    }

    sheetFlights.columns.forEach(col => col.width = 16);
    sheetEvents.columns.forEach(col => col.width = 16);

    return true;
};