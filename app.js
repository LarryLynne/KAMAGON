// Константа с индексами (числа со скрина МИНУС 1)
const colIdx = {
    route: 4, deadline: 7,
    days: [9, 10, 11, 12, 13, 14, 15],
    points: [16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27], // От Старта до Конца
    timings: [
        { arr: 28, dep: 29 }, // Старт
        { arr: 35, dep: 36 }, // П.Т.1
        { arr: 49, dep: 50 }, // П.Т.2
        { arr: 63, dep: 64 }, // П.Т.3
        { arr: 77, dep: 78 }, // П.Т.4
        { arr: 91, dep: 92 }, // П.Т.5
        { arr: 105, dep: 106 },// П.Т.6
        { arr: 119, dep: 120 },// П.Т.7
        { arr: 133, dep: 134 },// П.Т.8
        { arr: 147, dep: 148 },// П.Т.9
        { arr: 161, dep: 162 } // П.Т.10
    ],
    endTimings: { arr: 175, rel: 176 }, // Конец: Приїзд и Вивільнення
    meta: {
        delivery: 188, vehicle: 189, format: 190, code: 191, move: 196
    }
};

// Глобальный словарь для схем маршрутов
let routeDictionary = {};
let yardDictionary = {};

// Блокируем инпут до загрузки справочника
const fileInput = document.getElementById('fileInput');
fileInput.disabled = true; 

async function loadRouteSchemas() {
    const label = document.getElementById('fileInputLabel');
    const fileInput = document.getElementById('fileInput');
    
    label.classList.add('disabled');
    fileInput.disabled = true;

    const scriptUrl = 'https://script.google.com/macros/s/AKfycbxT4cGlFO8YcDzdeLaqSpThqgYbTbmhDoT8LSaB4FDNsLy0cGgsCa_V-zMINs3WhpcIEA/exec'; 

    try {
        const response = await fetch(scriptUrl);
        const data = await response.json(); 
        
        if (data.routes && data.yards) {
            routeDictionary = data.routes;
            yardDictionary = data.yards;
            console.log(`Довідники завантажено! Маршрутів: ${Object.keys(routeDictionary).length}`);
        } else {
            routeDictionary = data; 
        }
        
        if (allSchedules.length === 0) {
            document.getElementById('fileStatus').innerText = "Довідник завантажено. Можна обирати файли.";
        }
        
        label.classList.remove('disabled');
        fileInput.disabled = false;
        
    } catch (e) {
        console.error("Помилка:", e);
        document.getElementById('fileStatus').innerText = "Помилка завантаження довідника!";
    }
}

window.addEventListener('DOMContentLoaded', loadRouteSchemas);

class Schedule {
    constructor(row) {
        this.route = row[colIdx.route] || "";
        this.deadline = row[colIdx.deadline] || "";
        this.days = colIdx.days.map(i => row[i] === 1 || row[i] === "1");
        
        this.pointNames = colIdx.points.map(i => row[i] || "");
        
        this.allTimes = [];
        colIdx.timings.forEach(t => {
            this.allTimes.push(formatTime(row[t.arr])); 
            this.allTimes.push(formatTime(row[t.dep])); 
        });
        
        this.allTimes.push(formatTime(row[colIdx.endTimings.arr])); 
        this.allTimes.push(formatTime(row[colIdx.endTimings.rel]));

        this.deliveryType = row[colIdx.meta.delivery] || "";
        this.vehicleType = (row[colIdx.meta.vehicle] || "").toString().trim(); 
        this.loadFormat = row[colIdx.meta.format] || "";
        this.code = row[colIdx.meta.code] || "";
        this.moveType = row[colIdx.meta.move] || "";

        const dictKey = this.route.trim() + "|" + this.deliveryType.trim();
        this.schema = routeDictionary[dictKey] || "Схема не знайдена";
    }
}

// --- ГЛОБАЛЬНІ МАСИВИ (Оригінали та відфільтровані копії) ---
let allSchedules = [];
let filteredAllSchedules = [];

let detailedSchedules = [];
let filteredDetailedSchedules = [];

let yardEvents = [];
let filteredYardEvents = [];

let renderedCount = 0;         
let detailedRenderedCount = 0; 
let eventsRenderedCount = 0;

const CHUNK_SIZE = 200;        
const DETAILED_CHUNK_SIZE = 200; 
const EVENTS_CHUNK_SIZE = 300;

const workerCode = `
    importScripts('https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js');
    self.onmessage = function(e) {
        try {
            const data = new Uint8Array(e.data);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            self.postMessage({ success: true, rows: rows });
        } catch (error) {
            self.postMessage({ success: false, error: error.message });
        }
    };
`;

document.getElementById('fileInput').addEventListener('change', async (e) => {
    const files = e.target.files;
    if (!files.length) return;
    
    allSchedules = [];
    filteredAllSchedules = [];
    renderedCount = 0; 
    
    const statusText = document.getElementById('fileStatus');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');

    progressContainer.style.display = 'block';

    for (let i = 0; i < files.length; i++) {
        await processFile(files[i], i + 1, files.length, statusText, progressBar);
    }

    statusText.innerText = `Готово! Прочитано рейсів: ${allSchedules.length}`;
    progressContainer.style.display = 'none'; 
    
    filteredAllSchedules = [...allSchedules]; // Ініціалізуємо фільтрований масив
    initTable();
});

function processFile(file, currentFileNum, totalFilesNum, statusText, progressBar) {
    return new Promise((resolve) => {
        const reader = new FileReader();

        statusText.innerText = `Файл ${currentFileNum} из ${totalFilesNum}: Чтение...`;
        progressBar.style.width = '5%';
        progressBar.classList.remove('progress-animated');

        reader.onload = (e) => {
            statusText.innerText = `Файл ${currentFileNum} из ${totalFilesNum}: Распаковка...`;
            progressBar.classList.add('progress-animated');

            let currentProgress = 5;
            const progressInterval = setInterval(() => {
                currentProgress += (85 - currentProgress) * 0.05;
                progressBar.style.width = `${currentProgress}%`;
            }, 300);

            const blob = new Blob([workerCode], { type: 'application/javascript' });
            const workerUrl = URL.createObjectURL(blob);
            const worker = new Worker(workerUrl);

            worker.onmessage = (msgEvent) => {
                clearInterval(progressInterval);
                progressBar.classList.remove('progress-animated');
                
                const result = msgEvent.data;
                if (!result.success) {
                    console.error("Помилка:", result.error);
                    resolve(); 
                    return;
                }

                const rows = result.rows;
                statusText.innerText = `Файл ${currentFileNum} из ${totalFilesNum}: Сборка...`;
                
                let r = 2; 
                let startProgress = currentProgress; 

                function processRowsChunk() {
                    let end = Math.min(r + 2000, rows.length);
                    for (; r < end; r++) {
                        const row = rows[r];
                        if (row[colIdx.route] && row[colIdx.meta.vehicle] !== undefined && String(row[colIdx.meta.vehicle]).trim() !== "") {
                            allSchedules.push(new Schedule(row));
                        }
                    }

                    let progressPercent = startProgress + ((r / rows.length) * (100 - startProgress));
                    progressBar.style.width = `${progressPercent}%`;

                    if (r < rows.length) {
                        setTimeout(processRowsChunk, 0);
                    } else {
                        worker.terminate(); 
                        URL.revokeObjectURL(workerUrl);
                        resolve(); 
                    }
                }
                processRowsChunk();
            };

            worker.postMessage(e.target.result);
        };
        reader.readAsArrayBuffer(file);
    });
}

// ВАЖЛИВО: Оновлено генерацію шапки таблиці (додано інпути)
function initTable() {
    const wrapper = document.getElementById('rawTableWrapper'); 
    const container = document.getElementById('tableContainerRaw');
    
    if (filteredAllSchedules.length === 0) {
        wrapper.innerHTML = "";
        return;
    }

    container.classList.add('hide-pt');

    let html = `<table><thead><tr>`;
    let c = 0; 
    
    html += `<th class="col-route">Маршрут<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Дедлайн<br><input type="text" size = 1 class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    
    ['Пн','Вт','Ср','Чт','Пт','Сб','Нд'].forEach(d => html += `<th class="col-day">${d}<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`);
    
    html += `<th>Початкова<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Приїзд<br><input type="text" size = 1 class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Виїзд<br><input type="text" size = 1 class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;

    for(let i=1; i<=10; i++) {
        html += `<th class="pt-col">П.Т. №${i}<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
        html += `<th class="pt-col">Приїзд<br><input type="text" size = 1 class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
        html += `<th class="pt-col">Виїзд<br><input type="text" size = 1 class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    }

    html += `<th>Кінцева<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Приїзд<br><input type="text" size = 1 class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Вивільнення<br><input type="text" size = 1 class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Тип доставки<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Тип ТЗ<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Схема БДФ<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Формат<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Код<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th>Тип переміщення<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;

    html += `</tr></thead><tbody id="tableBody"></tbody></table>`;
    wrapper.innerHTML = html;
    
    renderChunk();
}

function renderChunk() {
    // Малюємо з filteredAllSchedules
    if (renderedCount >= filteredAllSchedules.length) return;

    const tbody = document.getElementById('tableBody');
    if (!tbody) return;

    let html = "";
    let end = Math.min(renderedCount + CHUNK_SIZE, filteredAllSchedules.length);

    for (let i = renderedCount; i < end; i++) {
        const item = filteredAllSchedules[i];
        html += `<tr>`;
        html += `<td class="col-route">${item.route}</td>`;
        html += `<td>${item.deadline}</td>`;
        item.days.forEach(d => html += `<td class="col-day ${d ? 'day-on' : 'day-off'}">${d ? '1' : '0'}</td>`);
        html += `<td>${item.pointNames[0]}</td><td>${item.allTimes[0] || ""}</td><td>${item.allTimes[1] || ""}</td>`;
        for(let j=1; j <= 10; j++) {
            html += `<td class="pt-col">${item.pointNames[j]}</td><td class="pt-col">${item.allTimes[j*2] || ""}</td><td class="pt-col">${item.allTimes[j*2 + 1] || ""}</td>`;
        }
        html += `<td>${item.pointNames[11]}</td><td>${item.allTimes[22] || ""}</td><td>${item.allTimes[23] || ""}</td>`;
        html += `<td>${item.deliveryType}</td><td>${item.vehicleType}</td><td><strong>${item.schema}</strong></td><td>${item.loadFormat}</td><td>${item.code}</td><td>${item.moveType}</td>`;
        html += `</tr>`;
    }
    tbody.insertAdjacentHTML('beforeend', html);
    renderedCount = end;
}

function formatTime(val) {
    if (val === undefined || val === null || val === "") return "—";
    if (typeof val === "string" && val.includes(":")) return val.substring(0, 5); 
    let num = parseFloat(val);
    if (!isNaN(num)) {
        let fraction = num - Math.floor(num); 
        let totalSeconds = Math.round(fraction * 86400); 
        let hours = Math.floor(totalSeconds / 3600);
        let minutes = Math.floor((totalSeconds % 3600) / 60);
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
    }
    return "—"; 
}

document.getElementById('togglePtBtn').addEventListener('click', function() {
    const container = document.getElementById('tableContainerRaw');
    if (container) {
        container.classList.toggle('hide-pt');
        this.classList.toggle('btn-active');
    }
});

// --- Модалка невідомих маршрутів ---
const unknownModal = document.getElementById('unknownRoutesModal');
const closeBtn = document.querySelector('.close-btn');
const copyBtn = document.getElementById('copyUnknownBtn');
let currentUnknownRoutesText = ""; 

document.getElementById('unknownRoutesBtn').addEventListener('click', () => {
    const container = document.getElementById('unknownRoutesTableContainer');
    const unknownSet = new Set();
    
    // Перевіряємо по оригінальному масиву (щоб фільтри не заважали пошуку помилок)
    allSchedules.forEach(item => {
        if (item.schema === "Схема не знайдена") {
            unknownSet.add(`${item.route} (Доставка: ${item.deliveryType})`);
        }
    });

    if (unknownSet.size === 0) {
        container.innerHTML = "<p style='padding: 10px; color: green;'>Всі маршрути мають схему в довіднику!</p>";
        copyBtn.style.display = 'none'; 
        currentUnknownRoutesText = "";
    } else {
        copyBtn.style.display = 'inline-block'; 
        currentUnknownRoutesText = Array.from(unknownSet).join('\n');
        let html = `<table><thead><tr><th>Невідомий маршрут</th></tr></thead><tbody>`;
        unknownSet.forEach(route => html += `<tr><td>${route}</td></tr>`);
        html += `</tbody></table>`;
        container.innerHTML = html;
    }
    unknownModal.style.display = 'block';
});

copyBtn.addEventListener('click', () => {
    if (!currentUnknownRoutesText) return;
    navigator.clipboard.writeText(currentUnknownRoutesText).then(() => {
        const originalText = copyBtn.innerText;
        copyBtn.innerText = "✅ Скопійовано!";
        copyBtn.style.backgroundColor = "#4caf50"; 
        copyBtn.style.color = "white";
        setTimeout(() => {
            copyBtn.innerText = originalText;
            copyBtn.style.backgroundColor = "";
            copyBtn.style.color = "";
        }, 2000);
    }).catch(err => alert("Помилка копіювання"));
});

closeBtn.addEventListener('click', () => unknownModal.style.display = 'none');
window.addEventListener('click', (event) => { if (event.target === unknownModal) unknownModal.style.display = 'none'; });

// --- Оновлення довідника ---
document.getElementById('updateDictBtn').addEventListener('click', async function() {
    const btn = this;
    const originalText = btn.innerText;
    btn.innerText = "⏳ Оновлення...";
    btn.disabled = true;

    await loadRouteSchemas();

    if (allSchedules.length > 0) {
        allSchedules.forEach(item => {
            const dictKey = item.route.trim() + "|" + item.deliveryType.trim();
            item.schema = routeDictionary[dictKey] || "Схема не знайдена";
        });

        const tbody = document.getElementById('tableBody');
        if (tbody) {
            tbody.innerHTML = ""; 
            let tempRendered = renderedCount; 
            renderedCount = 0; 
            while (renderedCount < tempRendered) renderChunk();
        }
    }
    btn.innerText = "✅ Оновлено!";
    setTimeout(() => { btn.innerText = originalText; btn.disabled = false; }, 2000);
});

// --- Вкладки ---
const tabRaw = document.getElementById('tabRaw');
const tabDetailed = document.getElementById('tabDetailed');
const tabEvents = document.getElementById('tabEvents');
const tabKamag = document.getElementById('tabKamag'); 

const containerRaw = document.getElementById('tableContainerRaw');
const containerDetailed = document.getElementById('tableContainerDetailed');
const containerEvents = document.getElementById('tableContainerEvents');
const containerKamag = document.getElementById('tableContainerKamag'); 

function switchTab(activeTabBtn, activeContainer) {
    [tabRaw, tabDetailed, tabEvents, tabKamag].forEach(btn => btn.classList.remove('active'));
    [containerRaw, containerDetailed, containerEvents, containerKamag].forEach(cont => cont.style.display = 'none');
    
    activeTabBtn.classList.add('active');
    activeContainer.style.display = activeContainer === containerKamag ? 'flex' : 'block';
}

tabRaw.addEventListener('click', () => switchTab(tabRaw, containerRaw));
tabDetailed.addEventListener('click', () => { switchTab(tabDetailed, containerDetailed); });
tabEvents.addEventListener('click', () => { switchTab(tabEvents, containerEvents); });
tabKamag.addEventListener('click', () => { switchTab(tabKamag, containerKamag); });

// --- Генерація ---
document.getElementById('generateDetailedBtn').addEventListener('click', () => {
    if (allSchedules.length === 0) return alert("Спочатку завантажте вихідні графіки!");

    const btn = document.getElementById('generateDetailedBtn');
    const originalText = btn.innerText;
    btn.innerText = "⏳ Генерація...";
    btn.disabled = true;

    setTimeout(() => {
        generateDetailedSchedules();
        calculateRampTimes();
        calculateUnloadingTimes();
        
        filteredDetailedSchedules = [...detailedSchedules]; // Оновлюємо фільтр
        initDetailedTable();

        generateYardEvents();
        filteredYardEvents = [...yardEvents]; // Оновлюємо фільтр
        initEventsTable();

        calculateKamagRequirements();
        
        tabDetailed.click();
        btn.innerText = originalText;
        btn.disabled = false;
    }, 50);
});

const dayNames = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];

function generateDetailedSchedules() {
    detailedSchedules = [];
    allSchedules.forEach(item => {
        if (!item.schema || item.schema === "ХЗ" || item.schema === "Схема не знайдена" || item.schema === "—") return;

        // 1. СИСТЕМА ДИНАМІЧНИХ ТОЧОК
        const activeNodes = [];
        for (let i = 0; i < 12; i++) {
            if (item.pointNames[i] && item.pointNames[i].toString().trim() !== "") {
                activeNodes.push({
                    name: item.pointNames[i].toString().trim(),
                    timeArr: item.allTimes[i * 2],
                    timeDep: item.allTimes[i * 2 + 1]
                });
            }
        }

        // 2. Очищаємо схему від можливих пробілів (щоб "12П 13П" стало "12П13П")
        const cleanSchema = item.schema.toString().replace(/\s+/g, '');
        const miniSchemas = [];
        for (let i = 0; i < cleanSchema.length; i += 3) {
            miniSchemas.push(cleanSchema.substring(i, i + 3));
        }

        item.days.forEach((isActiveDay, dayIndex) => {
            if (!isActiveDay) return; 

            miniSchemas.forEach(mini => {
                if (mini.length < 3) return; 

                const startIndex = parseInt(mini[0], 10) - 1; 
                const endIndex = parseInt(mini[1], 10) - 1;
                const containerType = mini.substring(2);

                // 3. СУПЕР-ЗАХИСТ: Додано перевірку isNaN (чи є це взагалі числом)
                if (isNaN(startIndex) || isNaN(endIndex) || startIndex < 0 || endIndex < 0 || startIndex >= activeNodes.length || endIndex >= activeNodes.length) return; 

                const nodeA = activeNodes[startIndex];
                const nodeB = activeNodes[endIndex];

                // Фінальний запобіжник на випадок чорної магії JS
                if (!nodeA || !nodeB) return;

                const yardDataA = yardDictionary[nodeA.name];
                const yardDataB = yardDictionary[nodeB.name];

                let arrMins = getAbsoluteMinutes(dayNames[dayIndex], nodeB.timeArr);
                let finalArrivalB = formatAbsoluteMinutes(arrMins);

                detailedSchedules.push({
                    originalRoute: item.route,
                    originalCode: item.code, 
                    day: dayNames[dayIndex],
                    miniSchema: mini,
                    containerType: containerType,
                    yardA: yardDataA ? yardDataA.yard : "—", 
                    nodeA: nodeA.name,
                    timePlacementA: "—", 
                    timeDepartureA: nodeA.timeDep || "—",
                    yardB: yardDataB ? yardDataB.yard : "—",
                    nodeB: nodeB.name,
                    timeArrivalB: finalArrivalB,
                    timeUnloadStart: "—", 
                    timeUnloadEnd: "—",   
                    vehicle: item.vehicleType,
                    moveType: item.moveType // <--- ДОБАВЛЕНО СЮДА
                });
            });
        });
    });
}

function initDetailedTable() {
    const container = document.getElementById('tableContainerDetailed');
    if (filteredDetailedSchedules.length === 0) {
        container.innerHTML = "<p style='padding:20px;'>Немає даних.</p>";
        return;
    }

    let c = 0;
    let html = `<table><thead><tr>
        <th class="col-day">№<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="originalRoute" class="sortable">Маршрут<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th>Код<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="day" class="sortable col-day">День<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="miniSchema" class="sortable">Схема<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="containerType" class="sortable">Тип<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="yardA" class="sortable" style="background-color: #fff3cd;">Автодвір А<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="timePlacementA" class="sortable" style="background-color: #d4edda;">Постановка<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="nodeA" class="sortable">Точка А<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="timeDepartureA" class="sortable">Виїзд<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="yardB" class="sortable" style="background-color: #fff3cd;">Автодвір Б<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="nodeB" class="sortable">Точка Б<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="timeArrivalB" class="sortable">Приїзд<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="timeUnloadStart" class="sortable" style="background-color: #cce5ff;">Постановка (вивант.)<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="timeUnloadEnd" class="sortable" style="background-color: #cce5ff;">Кінець вивант.<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="vehicle" class="sortable">Тип ТЗ<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
    </tr></thead><tbody id="detailedTableBody"></tbody></table>`;
    
    container.innerHTML = html;
    detailedRenderedCount = 0;
    renderDetailedChunk();
}

function renderDetailedChunk() {
    if (detailedRenderedCount >= filteredDetailedSchedules.length) return;
    const tbody = document.getElementById('detailedTableBody');
    if (!tbody) return;

    let html = "";
    let end = Math.min(detailedRenderedCount + DETAILED_CHUNK_SIZE, filteredDetailedSchedules.length);

    for (let i = detailedRenderedCount; i < end; i++) {
        const item = filteredDetailedSchedules[i];
        html += `<tr>
            <td class="col-day" style="color: #999;">${i + 1}</td>
            <td class="col-route">${item.originalRoute}</td>
            <td>${item.originalCode}</td>
            <td class="col-day" style="font-weight: bold;">${item.day}</td>
            <td style="text-align: center;">${item.miniSchema}</td>
            <td style="text-align: center;">${item.containerType}</td>
            <td style="font-weight: bold;">${item.yardA}</td>
            <td style="text-align: center; color: #d32f2f; font-weight: bold;">${item.timePlacementA || "—"}</td>
            <td>${item.nodeA}</td>
            <td style="text-align: center;">${item.timeDepartureA}</td>
            <td style="font-weight: bold;">${item.yardB}</td>
            <td>${item.nodeB}</td>
            <td style="text-align: center;">${item.timeArrivalB}</td>
            <td style="text-align: center; color: #0056b3; font-weight: bold;">${item.timeUnloadStart}</td>
            <td style="text-align: center; color: #0056b3; font-weight: bold;">${item.timeUnloadEnd}</td>
            <td>${item.vehicle}</td>
        </tr>`;
    }
    tbody.insertAdjacentHTML('beforeend', html);
    detailedRenderedCount = end;
}

[containerRaw, containerDetailed, containerEvents].forEach(container => {
    container.addEventListener('scroll', () => {
        if (container.scrollTop + container.clientHeight >= container.scrollHeight - 500) {
            if (container === containerRaw) renderChunk();
            else if (container === containerDetailed) renderDetailedChunk();
            else if (container === containerEvents) renderEventsChunk();
        }
    });
});

const dayMap = { 'Пн': 0, 'Вт': 1, 'Ср': 2, 'Чт': 3, 'Пт': 4, 'Сб': 5, 'Нд': 6 };
const reverseDayMap = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];

function getAbsoluteMinutes(dayStr, timeStr) {
    if (!timeStr || typeof timeStr !== 'string' || !timeStr.includes(':')) return Infinity; 
    
    // Якщо формат уже містить день (напр. "Пн 12:40"), розбиваємо його
    const parts = timeStr.trim().split(' ');
    const targetDay = parts.length === 2 ? parts[0] : dayStr;
    const targetTime = parts.length === 2 ? parts[1] : timeStr;

    const [hh, mm] = targetTime.split(':').map(Number);
    return dayMap[targetDay] * 1440 + hh * 60 + mm;
}

function formatAbsoluteMinutes(mins) {
    if (mins === Infinity) return "—";
    while (mins < 0) mins += 10080; 
    mins = mins % 10080;
    const d = Math.floor(mins / 1440);
    const h = Math.floor((mins % 1440) / 60);
    const m = mins % 60;
    return `${reverseDayMap[d]} ${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function calculateRampTimes() {
    const interval = parseInt(document.getElementById('rampInterval').value, 10) || 10;
    const groups = {};
    detailedSchedules.forEach(item => {
        if (item.timeDepartureA === "—") return;
        const key = `${item.nodeA}_${item.nodeB}`;
        if (!groups[key]) groups[key] = [];
        item.absDep = getAbsoluteMinutes(item.day, item.timeDepartureA);
        item.timeDepartureA = formatAbsoluteMinutes(item.absDep);
        groups[key].push(item);
    });
    
    for (const key in groups) {
        const group = groups[key];
        // Сортуємо за часом виїзду, а якщо він однаковий — за назвою маршруту (щоб контейнери одного рейсу стояли поруч)
        group.sort((a, b) => a.absDep - b.absDep || a.originalRoute.localeCompare(b.originalRoute));
        if (group.length === 0) continue;
        
        const lastOfGroup = group[group.length - 1];
        let prevAbsDep = lastOfGroup.absDep - 10080; 
        
        for (let i = 0; i < group.length; i++) {
            let item1 = group[i];
            let item2 = (i + 1 < group.length) ? group[i + 1] : null;
            
            // Перевіряємо, чи це два контейнери одного рейсу
            let isTwin = item2 && 
                         item1.originalRoute === item2.originalRoute && 
                         item1.day === item2.day && 
                         item1.absDep === item2.absDep;

            let proposedPlacement = prevAbsDep + interval;
            let maxPlacementTime = item1.absDep - (23 * 60); 
            let finalPlacement = Math.max(proposedPlacement, maxPlacementTime);
            
            if (isTwin) {
                let totalLoadTime = item1.absDep - finalPlacement;
                
                if (totalLoadTime <= 120) {
                    // ПАРАЛЕЛЬНА ПОСТАНОВКА (<= 2 годин)
                    item1.timePlacementA = formatAbsoluteMinutes(finalPlacement);
                    item2.timePlacementA = formatAbsoluteMinutes(finalPlacement);
                    prevAbsDep = item1.absDep;
                } else {
                    // ПОСЛІДОВНА ПОСТАНОВКА (> 2 годин)
                    let tHalf = Math.floor((totalLoadTime - 10) / 2);
                    
                    // Перший контейнер (ставиться раніше, виїжджає з рампи раніше)
                    item1.timePlacementA = formatAbsoluteMinutes(finalPlacement);
                    item1.absDep = finalPlacement + tHalf; // Змінюємо час виїзду для 1-го
                    item1.timeDepartureA = formatAbsoluteMinutes(item1.absDep);
                    
                    // Другий контейнер (через 10 хв після виїзду 1-го)
                    item2.timePlacementA = formatAbsoluteMinutes(finalPlacement + tHalf + 10);
                    prevAbsDep = item2.absDep;
                }
                i++; // Пропускаємо другий контейнер у циклі, бо ми його вже обробили
            } else {
                // Одинарний контейнер
                item1.timePlacementA = formatAbsoluteMinutes(finalPlacement);
                prevAbsDep = item1.absDep;
            }
        }
    }
}

function calculateUnloadingTimes() {
    const containerCounts = {};
    detailedSchedules.forEach(item => {
        item.timeUnloadStart = "—";
        item.timeUnloadEnd = "—";
        
        // ДОБАВЛЕНО: Якщо порожній, не рахуємо час вивантаження взагалі
        if (item.moveType && item.moveType.toLowerCase().includes("порожній")) return;

        if (item.vehicle !== "Шасі BDF" || item.timeArrivalB === "—") return;
        const yardDataB = yardDictionary[item.nodeB];
        if (!yardDataB) return; 

        const trackerKey = `${item.day}_${item.originalCode}`;
        if (!containerCounts[trackerKey]) containerCounts[trackerKey] = 0;
        containerCounts[trackerKey]++;

        const isFirst = containerCounts[trackerKey] === 1;
        const prepTimeMins = isFirst ? yardDataB.first : yardDataB.second;
        const unloadTimeMins = yardDataB.unload;
        const arrivalMins = getAbsoluteMinutes(item.day, item.timeArrivalB);
        
        if (arrivalMins !== Infinity) {
            item.timeUnloadStart = formatAbsoluteMinutes(arrivalMins + prepTimeMins);
            item.timeUnloadEnd = formatAbsoluteMinutes(arrivalMins + prepTimeMins + unloadTimeMins);
        }
    });
}

// --- Сортування ---
let sortState = { key: null, asc: true };

document.getElementById('tableContainerDetailed').addEventListener('click', function(e) {
    if (e.target.tagName === 'TH' && e.target.hasAttribute('data-sort')) {
        sortDetailedSchedules(e.target.getAttribute('data-sort'), e.target);
    }
});

function sortDetailedSchedules(key, thElement) {
    if (sortState.key === key) sortState.asc = !sortState.asc;
    else { sortState.key = key; sortState.asc = true; }

    const asc = sortState.asc;
    const getValue = (item, k) => {
        // Універсальна логіка для всіх колонок з часом
        const timeCols = ['timePlacementA', 'timeDepartureA', 'timeArrivalB', 'timeUnloadStart', 'timeUnloadEnd'];
        if (timeCols.includes(k)) {
            return getAbsoluteMinutes(item.day, item[k]);
        }
        
        // Логіка для сортування за днем тижня
        if (k === 'day') return dayMap[item.day] !== undefined ? dayMap[item.day] : Infinity;
        
        // Для всього іншого (текст)
        return item[k] !== undefined && item[k] !== null ? item[k] : "";
    };

    // Сортуємо оригінальний масив
    detailedSchedules.sort((a, b) => {
        const valA = getValue(a, key);
        const valB = getValue(b, key);
        if (typeof valA === 'number' && typeof valB === 'number') return asc ? valA - valB : valB - valA;
        return asc ? String(valA).localeCompare(String(valB)) : String(valB).localeCompare(String(valA));
    });

    const tr = thElement.parentElement;
    Array.from(tr.children).forEach(th => th.classList.remove('sort-asc', 'sort-desc'));
    thElement.classList.add(asc ? 'sort-asc' : 'sort-desc');

    // Перезастосовуємо фільтри (вони оновлять екран автоматично)
    applyFiltersDetailed();
    document.getElementById('tableContainerDetailed').scrollTop = 0; 
}

function generateYardEvents() {
    yardEvents = [];
    const addEvent = (yard, eventName, absMins, code) => {
        if (absMins === Infinity || isNaN(absMins)) return;
        const formatted = formatAbsoluteMinutes(absMins); 
        const parts = formatted.split(' '); 
        yardEvents.push({ yard: yard, code: code, event: eventName, day: parts[0], time: parts[1], absMins: absMins });
    };

    detailedSchedules.forEach(item => {
        // --- ПОДІЇ ДЛЯ ТОЧКИ А (ЗАВАНТАЖЕННЯ) ---
        if (item.yardA && item.yardA !== "—") {
            if (item.timePlacementA && item.timePlacementA !== "—") {
                const parts = item.timePlacementA.split(' ');
                addEvent(item.yardA, "Постановка на завантаження", getAbsoluteMinutes(parts[0], parts[1]), item.originalCode);
            }
            if (item.timeDepartureA && item.timeDepartureA !== "—") {
                addEvent(item.yardA, "Забір контейнера з-під завантаження", item.absDep - 15, item.originalCode);
            }
        }
        
        // --- ПОДІЇ ДЛЯ ТОЧКИ Б (ВИВАНТАЖЕННЯ) ---
        if (item.yardB && item.yardB !== "—") {
            
            // ПЕРЕВІРКА НА "ПОРОЖНІЙ"
            if (item.moveType && item.moveType.toLowerCase().includes("порожній")) {
                // Тільки одна операція рівно по приїзду
                let arrMins = getAbsoluteMinutes(item.day, item.timeArrivalB);
                addEvent(item.yardB, "Забір з-під вивантаження", arrMins, item.originalCode);
            } else {
                // СТАНДАРТНА ЛОГІКА
                if (item.timeUnloadStart && item.timeUnloadStart !== "—") {
                    const parts = item.timeUnloadStart.split(' ');
                    addEvent(item.yardB, "Постановка на вивантаження", getAbsoluteMinutes(parts[0], parts[1]), item.originalCode);
                }
                
                if (item.timeUnloadEnd && item.timeUnloadEnd !== "—") {
                    const parts = item.timeUnloadEnd.split(' ');
                    const endMins = getAbsoluteMinutes(parts[0], parts[1]);
                    
                    addEvent(item.yardB, "Забір з-під вивантаження", endMins, item.originalCode);
                    
                    /*let moveTimeMins = (yardDictionary[item.nodeB] && yardDictionary[item.nodeB].move !== undefined) ? yardDictionary[item.nodeB].move : 5;
                    addEvent(item.yardB, "Переїзд на зону пустих контейнерів", endMins + moveTimeMins, item.originalCode);*/
                }
            }
        }
    });
    yardEvents.sort((a, b) => a.absMins - b.absMins);
}

function initEventsTable() {
    const container = document.getElementById('tableContainerEvents');
    if (filteredYardEvents.length === 0) {
        container.innerHTML = "<p style='padding:20px;'>Немає подій.</p>";
        return;
    }

    let c = 0;
    let html = `<table><thead><tr>
        <th data-sort-event="yard" class="sortable">Автодвір<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th>Код<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort-event="event" class="sortable">Подія<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort-event="day" class="sortable col-day">День<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort-event="time" class="sortable">Час<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
    </tr></thead><tbody id="eventsTableBody"></tbody></table>`;
    
    container.innerHTML = html;
    eventsRenderedCount = 0;
    renderEventsChunk();
}

function renderEventsChunk() {
    if (eventsRenderedCount >= filteredYardEvents.length) return;
    const tbody = document.getElementById('eventsTableBody');
    if (!tbody) return;

    let html = "";
    let end = Math.min(eventsRenderedCount + EVENTS_CHUNK_SIZE, filteredYardEvents.length);

    for (let i = eventsRenderedCount; i < end; i++) {
        const ev = filteredYardEvents[i];
        html += `<tr>
            <td style="font-weight: bold;">${ev.yard}</td>
            <td>${ev.code}</td>
            <td>${ev.event}</td>
            <td class="col-day" style="font-weight: bold;">${ev.day}</td>
            <td style="text-align: center;">${ev.time}</td>
        </tr>`;
    }
    tbody.insertAdjacentHTML('beforeend', html);
    eventsRenderedCount = end;
}

// --- РОЗРАХУНОК КАМАГІВ ---
let kamagData = {}; 
let totalOpsData = {}; // Тепер тут лише загальна кількість операцій

function calculateKamagRequirements() {
    kamagData = {};
    totalOpsData = {}; 

    yardEvents.forEach(ev => {
        // Ініціалізуємо структури для підрахунку загальної кількості операцій
        if (!totalOpsData[ev.yard]) totalOpsData[ev.yard] = {};
        if (!totalOpsData[ev.yard][ev.day]) totalOpsData[ev.yard][ev.day] = Array(24).fill(0);

        const hour = parseInt(ev.time.split(':')[0], 10);
        if (!isNaN(hour)) {
            totalOpsData[ev.yard][ev.day][hour]++;
        }
    });

    // Розраховуємо потребу в КАМАГах (1 машина на 12 операцій)
    for (let yard in totalOpsData) {
        kamagData[yard] = {};
        for (let day in totalOpsData[yard]) {
            kamagData[yard][day] = totalOpsData[yard][day].map(count => Math.ceil(count / 12));
        }
    }

    // Оновлюємо список автодворів у селекті
    const yardSelect = document.getElementById('kamagYardSelect');
    const currentVal = yardSelect.value;
    yardSelect.innerHTML = "";
    Object.keys(totalOpsData).sort().forEach(yard => {
        const option = document.createElement('option');
        option.value = yard;
        option.textContent = yard;
        yardSelect.appendChild(option);
    });
    if (currentVal && totalOpsData[currentVal]) yardSelect.value = currentVal;

    renderKamagTable();
}

function renderKamagTable() {
    const yard = document.getElementById('kamagYardSelect').value;
    const wrapper = document.getElementById('kamagTableWrapper');

    if (!yard || !kamagData[yard]) {
        wrapper.innerHTML = "<p style='padding:20px;'>Немає даних для цього автодвору.</p>";
        return;
    }

    const daysOfWeek = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];

    // Додано 4-й параметр includeCharts (за замовчуванням false)
    function generateMatrixHTML(title, rowHeaders, dataProvider, includeCharts = false) {
        let html = `<h3 style="margin: 20px 0 10px 0; color: #334155; border-left: 4px solid #ffaa00; padding-left: 10px;">${title}</h3>`;
        html += `<table><thead>`;
        
        html += `<tr><th style="min-width: 120px;"></th>`;
        daysOfWeek.forEach(d => {
            html += `<th colspan="25" style="text-align: center; font-weight: bold; background-color: #e9ecef; border-left: 2px solid #6c757d; border-right: 2px solid #6c757d;">${d}</th>`;
        });
        html += `<th style="text-align: center; line-height: 1.2;">Всього</th></tr>`;

        html += `<tr><th style="font-size: 10px;">Рядок / Година</th>`;
        daysOfWeek.forEach(d => {
            for (let i = 0; i < 24; i++) {
                let borderStyle = i === 0 ? "border-left: 2px solid #6c757d;" : "";
                html += `<th class="kamag-header-vertical" style="${borderStyle}">${i}:00</th>`;
            }
            html += `<th style="text-align: center; font-weight: bold; font-size: 10px; background-color: #dee2e6; border-right: 2px solid #6c757d; min-width: 35px;">Σ</th>`;
        });
        html += `<th></th></tr></thead><tbody>`;

        rowHeaders.forEach(rowName => {
            html += `<tr><td style="font-weight: bold; font-size: 11px; white-space: normal;">${rowName}</td>`;
            let totalRowSum = 0;
            
            daysOfWeek.forEach(d => {
                let dailySum = 0;
                for (let h = 0; h < 24; h++) {
                    let val = dataProvider(rowName, d, h);
                    let borderStyle = h === 0 ? "border-left: 2px solid #6c757d;" : "";
                    
                    if (val > 0) {
                        let cellClass = title.includes("КАМАГ") ? "kamag-cell kamag-active" : "kamag-cell";
                        let cellStyle = !title.includes("КАМАГ") ? "background-color: #fff9c4; font-weight: bold;" : "";
                        html += `<td class="${cellClass}" style="${borderStyle} ${cellStyle}">${val}</td>`;
                        dailySum += val;
                        totalRowSum += val;
                    } else {
                        html += `<td class="kamag-cell" style="${borderStyle}"></td>`;
                    }
                }
                html += `<td style="text-align: center; font-weight: bold; background-color: #f1f3f5; border-right: 2px solid #6c757d;">${dailySum > 0 ? dailySum : ''}</td>`;
            });
            html += `<td style="text-align: center; font-weight: bold; background-color: #e9ecef;">${totalRowSum}</td></tr>`;
        });

        // ЯКЩО ПОТРІБНІ ГРАФІКИ (для таблиці операцій)
        if (includeCharts) {
            html += `<tr><td style="font-weight: bold; font-size: 11px;">Графік</td>`;
            daysOfWeek.forEach((d, index) => {
                // Кожна діаграма розтягується рівно на 24 колонки годин
                html += `<td colspan="24" style="border-left: 2px solid #6c757d; vertical-align: bottom; padding: 5px; background: #fff;">
                            <div style="height: 60px; width: 100%; position: relative;">
                                <canvas id="chart_${index}"></canvas>
                            </div>
                         </td>
                         <td style="border-right: 2px solid #6c757d; background-color: #dee2e6;"></td>`;
            });
            html += `<td style="background-color: #e9ecef;"></td></tr>`;
        }

        html += `</tbody></table>`;
        return html;
    }

    // 1. Таблиця КАМАГів (машин)
    let maxK = 0;
    daysOfWeek.forEach(d => { if(kamagData[yard][d]) maxK = Math.max(maxK, ...kamagData[yard][d]); });
    const kamagRows = Array.from({length: maxK}, (_, i) => `Камаг ${i+1}`);
    const kamagHTML = generateMatrixHTML(`Розрахунок КАМАГів (машин)`, kamagRows, (row, day, hour) => {
        const kNum = parseInt(row.split(' ')[1]);
        return (kamagData[yard][day] && kamagData[yard][day][hour] >= kNum) ? 1 : 0;
    });

    // 2. Таблиця Операцій + ДІАГРАМИ (передаємо includeCharts = true)
    const opsHTML = generateMatrixHTML(`Кількість операцій (всього)`, ["Всього операцій"], (row, day, hour) => {
        return (totalOpsData[yard] && totalOpsData[yard][day]) ? totalOpsData[yard][day][hour] : 0;
    }, true);

    wrapper.innerHTML = kamagHTML + "<div style='height: 40px;'></div>" + opsHTML;

    // 3. Ініціалізуємо 7 ОРЕМИХ ДІАГРАМ
    if (window.myDayCharts) {
        window.myDayCharts.forEach(c => c.destroy()); // Чистимо старі при перемиканні
    }
    window.myDayCharts = [];

    daysOfWeek.forEach((d, index) => {
        const ctx = document.getElementById(`chart_${index}`);
        if (!ctx) return;

        const chartLabels = [];
        const chartData = [];
        for (let h = 0; h < 24; h++) {
            chartLabels.push(`${h}:00`);
            const val = (totalOpsData[yard] && totalOpsData[yard][d]) ? totalOpsData[yard][d][h] : 0;
            chartData.push(val);
        }

        const newChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: chartLabels,
                datasets: [{
                    data: chartData,
                    backgroundColor: '#ffc107',
                    borderColor: '#ffaa00',
                    borderWidth: 1,
                    borderRadius: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: { display: false, beginAtZero: true }, // Вісь Y прихована для компактності
                    x: { display: false } // Вісь X прихована (години вже є в шапці таблиці)
                },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            title: (items) => `${d} ${items[0].label}`,
                            label: (item) => `Операцій: ${item.raw}`
                        }
                    }
                },
                layout: { padding: 0 }
            }
        });
        window.myDayCharts.push(newChart);
    });
}

document.getElementById('kamagYardSelect').addEventListener('change', renderKamagTable);
document.getElementById('kamagDaySelect').addEventListener('change', renderKamagTable);

// =========================================
// ФІЛЬТРАЦІЯ ТА ЕКСПОРТ (НОВЕ)
// =========================================

// Допоміжні функції для перетворення об'єктів у масиви рядків (для фільтра і Excel)
function getRawValues(item) {
    const vals = [
        item.route, item.deadline,
        ...(item.days.map(d => d ? "1" : "0")),
        item.pointNames[0], item.allTimes[0] || "", item.allTimes[1] || ""
    ];
    for(let j=1; j<=10; j++) {
        vals.push(item.pointNames[j], item.allTimes[j*2] || "", item.allTimes[j*2 + 1] || "");
    }
    vals.push(
        item.pointNames[11], item.allTimes[22] || "", item.allTimes[23] || "",
        item.deliveryType, item.vehicleType, item.schema, item.loadFormat, item.code, item.moveType
    );
    return vals;
}

function getDetailedValues(item) {
    return [
        "", // Місце під індекс (заповниться при рендері/експорті)
        item.originalRoute, 
        item.originalCode, // Додано код
        item.day, 
        item.miniSchema, 
        item.containerType,
        item.yardA, 
        item.timePlacementA || "—", 
        item.nodeA, 
        item.timeDepartureA,
        item.yardB, 
        item.nodeB, 
        item.timeArrivalB,
        item.timeUnloadStart, 
        item.timeUnloadEnd, 
        item.vehicle
    ];
}

function getEventsValues(ev) {
    return [ev.yard, ev.code, ev.event, ev.day, ev.time];
}

// Універсальний рушій фільтрації
// Універсальний "Розумний" рушій фільтрації
function filterDataArray(containerId, dataArray, valuesExtractor) {
    const inputs = document.querySelectorAll(`#${containerId} .filter-input`);
    const filters = [];
    
    inputs.forEach(input => {
        const val = input.value.trim().toLowerCase();
        if (val) {
            filters.push({ col: parseInt(input.getAttribute('data-col')), val: val });
        }
    });

    if (filters.length === 0) return [...dataArray];

    return dataArray.filter(item => {
        const rowVals = valuesExtractor(item);
        
        // Кожна заповнена колонка повинна збігатися (логіка AND між стовпцями)
        return filters.every(f => {
            const cellStr = String(rowVals[f.col] || "").toLowerCase();
            
            // Розбиваємо введений текст по комі (логіка OR всередині стовпця)
            // Наприклад: "київ, львів" -> ['київ', 'львів']
            const searchTerms = f.val.split(',').map(s => s.trim()).filter(Boolean);
            
            // Якщо хоча б один шматочок тексту є в комірці — рядок підходить!
            return searchTerms.some(term => cellStr.includes(term));
        });
    });
}

// Повертаємо слухач подій на 'input' з затримкою (debounce), щоб можна було друкувати
let filterTimeout;
document.addEventListener('input', function(e) {
    if (e.target.classList.contains('filter-input')) {
        clearTimeout(filterTimeout);
        filterTimeout = setTimeout(() => {
            const container = e.target.closest('.table-container');
            if (container.id === 'tableContainerRaw') applyFiltersRaw();
            else if (container.id === 'tableContainerDetailed') applyFiltersDetailed();
            else if (container.id === 'tableContainerEvents') applyFiltersEvents();
        }, 300); // 300мс затримки, щоб інтерфейс не зависав під час швидкого друку
    }
});
// Слухач подій для списків (використовуємо 'change' замість 'input')
document.addEventListener('change', function(e) {
    if (e.target.classList.contains('filter-input')) {
        const container = e.target.closest('.table-container');
        if (container.id === 'tableContainerRaw') applyFiltersRaw();
        else if (container.id === 'tableContainerDetailed') applyFiltersDetailed();
        else if (container.id === 'tableContainerEvents') applyFiltersEvents();
    }
});

function applyFiltersRaw() {
    filteredAllSchedules = filterDataArray('tableContainerRaw', allSchedules, getRawValues);
    renderedCount = 0;
    const tbody = document.getElementById('tableBody');
    if(tbody) tbody.innerHTML = "";
    renderChunk();
}

function applyFiltersDetailed() {
    filteredDetailedSchedules = filterDataArray('tableContainerDetailed', detailedSchedules, getDetailedValues);
    detailedRenderedCount = 0;
    const tbody = document.getElementById('detailedTableBody');
    if(tbody) tbody.innerHTML = "";
    renderDetailedChunk();
}

function applyFiltersEvents() {
    filteredYardEvents = filterDataArray('tableContainerEvents', yardEvents, getEventsValues);
    eventsRenderedCount = 0;
    const tbody = document.getElementById('eventsTableBody');
    if(tbody) tbody.innerHTML = "";
    renderEventsChunk();
}

document.addEventListener('input', function(e) {
    if (e.target.classList.contains('filter-input')) {
        clearTimeout(filterTimeout);
        filterTimeout = setTimeout(() => {
            const container = e.target.closest('.table-container');
            if (container.id === 'tableContainerRaw') applyFiltersRaw();
            else if (container.id === 'tableContainerDetailed') applyFiltersDetailed();
            else if (container.id === 'tableContainerEvents') applyFiltersEvents();
        }, 300); // 300мс затримки, щоб не гальмувати при швидкому наборі
    }
});

// Експорт в Excel
document.getElementById('exportExcelBtn').addEventListener('click', async () => {
    const btn = document.getElementById('exportExcelBtn');
    const originalText = btn.innerText;
    btn.innerText = "⏳ Формування...";
    btn.disabled = true;

    try {
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Звіт');
        
        if (tabRaw.classList.contains('active')) {
            // --- ЭКСПОРТ ИСХОДНЫХ ГРАФИКОВ ---
            const headers = ["Маршрут", "Дедлайн", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд", "Початкова", "Приїзд", "Виїзд"];
            for(let i=1; i<=10; i++) headers.push(`П.Т. №${i}`, "Приїзд", "Виїзд");
            headers.push("Кінцева", "Приїзд", "Вивільнення", "Тип доставки", "Тип ТЗ", "Схема БДФ", "Формат", "Код", "Тип переміщення");
            
            sheet.addRow(headers);
            filteredAllSchedules.forEach(item => sheet.addRow(getRawValues(item)));

        } else if (tabDetailed.classList.contains('active')) {
            // --- ЭКСПОРТ ДЕТАЛИЗИРОВАННЫХ РЕЙСОВ ---
            const headers = ["№", "Маршрут", "Код", "День", "Схема", "Тип", "Автодвір А", "Постановка", "Точка А", "Виїзд", "Автодвір Б", "Точка Б", "Приїзд", "Постановка (вивант.)", "Кінець вивант.", "Тип ТЗ"];
            sheet.addRow(headers);
            filteredDetailedSchedules.forEach((item, index) => {
                let vals = getDetailedValues(item);
                vals[0] = index + 1;
                sheet.addRow(vals);
            });

        } else if (tabKamag.classList.contains('active')) {
            // --- ЭКСПОРТ КАМАГОВ И ДИАГРАММ С ФОРМАТИРОВАНИЕМ ---
            const yard = document.getElementById('kamagYardSelect').value;
            const days = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];
            
            // 0. Налаштування ширини колонок
            sheet.getColumn(1).width = 18; // Рядок / Година
            for (let i = 2; i <= 1 + 25 * 7; i++) {
                sheet.getColumn(i).width = 3.5; // Години (робимо вузькими)
                if ((i - 1) % 25 === 0) sheet.getColumn(i).width = 5; // Стовпці сум (трохи ширші)
            }
            sheet.getColumn(2 + 25 * 7).width = 8; // Загальна сума "Разом"

            // Словник стилів
            const alignCenter = { vertical: 'middle', horizontal: 'center' };
            const fillHeader = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } };
            const fillSum = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F3F5' } };
            const fillActive = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC107' } }; 
            const fillOps = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };

            const borderThin = { style: 'thin', color: { argb: 'FFCCCCCC' } };
            const borderMedium = { style: 'medium', color: { argb: 'FF6C757D' } };

            const getBorders = (isLeftEdge, isRightEdge) => ({
                top: borderThin, bottom: borderThin,
                left: isLeftEdge ? borderMedium : borderThin,
                right: isRightEdge ? borderMedium : borderThin
            });

            // 1. ЗАГОЛОВОК
            sheet.addRow([`Звіт по КАМАГам: ${yard}`]).font = { bold: true, size: 14 };
            sheet.addRow([]);

            // 2. ШАПКА ТАБЛИЦІ (Дні тижня)
            const rowDays = sheet.addRow(["День"]);
            rowDays.getCell(1).font = { bold: true };
            rowDays.getCell(1).alignment = alignCenter;

            days.forEach((d, index) => {
                const startCol = 2 + index * 25;
                const endCol = startCol + 24;
                sheet.mergeCells(3, startCol, 3, endCol); // Рядок 3 - це дні
                
                const cell = sheet.getCell(3, startCol);
                cell.value = d;
                cell.alignment = alignCenter;
                cell.font = { bold: true };
                cell.fill = fillHeader;
                cell.border = getBorders(true, true);
            });
            sheet.getCell(3, 2 + 25 * 7).value = "Всього";
            sheet.getCell(3, 2 + 25 * 7).font = { bold: true };

            // 3. ШАПКА ТАБЛИЦІ (Години)
            const rowHours = sheet.addRow(["Рядок / Година"]);
            rowHours.getCell(1).font = { size: 10 };
            
            let currentCol = 2;
            days.forEach((d, index) => {
                for (let h = 0; h < 24; h++) {
                    const cell = rowHours.getCell(currentCol);
                    cell.value = h; // Тільки число для економії місця
                    cell.alignment = alignCenter;
                    cell.font = { size: 9 };
                    cell.border = getBorders(h === 0, false);
                    currentCol++;
                }
                const sumCell = rowHours.getCell(currentCol);
                sumCell.value = "Σ";
                sumCell.alignment = alignCenter;
                sumCell.font = { bold: true, size: 9 };
                sumCell.fill = fillHeader;
                sumCell.border = getBorders(false, true);
                currentCol++;
            });

            // 4. ДАНІ КАМАГІВ
            let maxK = 0;
            days.forEach(d => { if(kamagData[yard][d]) maxK = Math.max(maxK, ...kamagData[yard][d]); });
            
            for (let k = 1; k <= maxK; k++) {
                const row = sheet.addRow([`Камаг ${k}`]);
                row.getCell(1).font = { bold: true, size: 10 };
                let weekTotal = 0;
                let cCol = 2;

                days.forEach(d => {
                    let daySum = 0;
                    for (let h = 0; h < 24; h++) {
                        let val = (kamagData[yard][d] && kamagData[yard][d][h] >= k) ? 1 : "";
                        const cell = row.getCell(cCol);
                        cell.value = val;
                        cell.alignment = alignCenter;
                        cell.border = getBorders(h === 0, false);
                        
                        if (val === 1) {
                            cell.fill = fillActive;
                            daySum++; weekTotal++;
                        }
                        cCol++;
                    }
                    // Денна сума
                    const dSumCell = row.getCell(cCol);
                    dSumCell.value = daySum || "";
                    dSumCell.alignment = alignCenter;
                    dSumCell.font = { bold: true };
                    dSumCell.fill = fillSum;
                    dSumCell.border = getBorders(false, true);
                    cCol++;
                });
                
                // Тижнева сума
                const wSumCell = row.getCell(cCol);
                wSumCell.value = weekTotal;
                wSumCell.alignment = alignCenter;
                wSumCell.font = { bold: true };
                wSumCell.fill = fillHeader;
            }

            sheet.addRow([]); // Пробіл

            // 5. ДАНІ ОПЕРАЦІЙ
            const opsRow = sheet.addRow(["Всього операцій"]);
            opsRow.getCell(1).font = { bold: true, size: 10 };
            
            let totalWeekOps = 0;
            let oCol = 2;
            
            days.forEach(d => {
                let daySum = 0;
                for (let h = 0; h < 24; h++) {
                    let val = (totalOpsData[yard] && totalOpsData[yard][d]) ? totalOpsData[yard][d][h] : 0;
                    const cell = opsRow.getCell(oCol);
                    cell.value = val || "";
                    cell.alignment = alignCenter;
                    cell.font = { bold: true, size: 9 };
                    cell.border = getBorders(h === 0, false);
                    
                    if (val > 0) cell.fill = fillOps; // Світло-жовтий фон
                    
                    daySum += val;
                    totalWeekOps += val;
                    oCol++;
                }
                // Денна сума
                const dSumCell = opsRow.getCell(oCol);
                dSumCell.value = daySum || "";
                dSumCell.alignment = alignCenter;
                dSumCell.font = { bold: true };
                dSumCell.fill = fillSum;
                dSumCell.border = getBorders(false, true);
                oCol++;
            });
            
            // Тижнева сума
            const wOpsSumCell = opsRow.getCell(oCol);
            wOpsSumCell.value = totalWeekOps;
            wOpsSumCell.alignment = alignCenter;
            wOpsSumCell.font = { bold: true };
            wOpsSumCell.fill = fillHeader;

            // 6. ВСТАВКА ДІАГРАМ (Горизонтально під кожним днем)
            sheet.addRow([]);
            sheet.addRow(["Графіки:"]).font = { bold: true };
            
            const imgRow = sheet.rowCount; // Рядок, де будуть картинки

            for (let i = 0; i < 7; i++) {
                const canvas = document.getElementById(`chart_${i}`);
                if (canvas) {
                    const base64 = canvas.toDataURL('image/png');
                    const imageId = workbook.addImage({
                        base64: base64,
                        extension: 'png',
                    });
                    
                    // tl: { col, row } вказує верхній лівий кут. 
                    // Колонки в ExcelJS для зображень починаються з 0 (0 = A, 1 = B).
                    // Наші дні починаються з колонки B (індекс 1), і кожен займає 25 колонок.
                    sheet.addImage(imageId, {
                        tl: { col: 1 + i * 25, row: imgRow },
                        ext: { width: 620, height: 100 } // Ширина підігнана під 24 вузькі колонки
                    });
                }
            }
            
            // Робимо кілька порожніх рядків під графіками, щоб вони не перекривали можливий текст нижче
            for(let i=0; i<6; i++) sheet.addRow([]); 
        }

        // Сохранение файла
        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), `Kamagon_Export_${new Date().getTime()}.xlsx`);

    } catch (err) {
        console.error(err);
        alert("Помилка при експорті!");
    } finally {
        btn.innerText = originalText;
        btn.disabled = false;
    }
});
