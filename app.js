// --- АВТОРИЗАЦІЯ З ПРИХОВУВАННЯМ ---
const APP_PASSWORD = "299792458"; 

function checkAuth() {
    return new Promise((resolve) => {
        if (sessionStorage.getItem('kamagonAuth') === 'true') {
            resolve(true);
            return;
        }
        
        const modal = document.getElementById('authModal');
        const input = document.getElementById('authPassInput');
        modal.style.display = 'block';
        input.value = '';
        input.focus();

        document.getElementById('authConfirmBtn').onclick = () => {
            if (input.value === APP_PASSWORD) {
                sessionStorage.setItem('kamagonAuth', 'true');
                modal.style.display = 'none';
                updateAuthVisibility(); // Показываем всё
                //document.getElementById('tabRaw').click();
                resolve(true);
            } else {
                alert("Невірний пароль!");
                input.value = '';
            }
        };

        document.getElementById('authCancelBtn').onclick = () => {
            modal.style.display = 'none';
            resolve(false);
        };
    });
}

function updateAuthVisibility() {
    const isAuth = sessionStorage.getItem('kamagonAuth') === 'true';
    const loginBtn = document.getElementById('loginBtn');
    
    // Скрываем или показываем все защищенные элементы
    document.querySelectorAll('.auth-hidden').forEach(el => {
        if (isAuth) {
            el.classList.remove('auth-hidden');
        }
    });

    // Если авторизован — прячем ключик, он больше не нужен
    if (isAuth && loginBtn) loginBtn.style.display = 'none';
}

// Привязываем вызов модалки к ключику
document.getElementById('loginBtn').addEventListener('click', checkAuth);

// Проверяем статус при загрузке
document.addEventListener('DOMContentLoaded', updateAuthVisibility);

// Константа с индексами (числа со скрина МИНУС 1)
const colIdx = {
    dateStart: 1, dateEnd: 2, // <--- НОВЫЕ СТОЛБЦЫ B и C
    route: 4, deadline: 7,
    days: [9, 10, 11, 12, 13, 14, 15],
    points: [16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27],
    timings: [
        { arr: 28, dep: 29 }, { arr: 35, dep: 36 }, { arr: 49, dep: 50 },
        { arr: 63, dep: 64 }, { arr: 77, dep: 78 }, { arr: 91, dep: 92 },
        { arr: 105, dep: 106 },{ arr: 119, dep: 120 },{ arr: 133, dep: 134 },
        { arr: 147, dep: 148 },{ arr: 161, dep: 162 }
    ],
    endTimings: { arr: 175, rel: 176 },
    meta: {
        delivery: 188, vehicle: 189, format: 190, code: 191, move: 196
    }
};

// Функция для безопасного чтения дат из Excel
function parseExcelDate(val) {
    if (!val) return null;
    if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
}

// Глобальные словари
let routeDictionary = {};
let yardDictionary = {};
let fleetDictionary = {}; // НОВЫЙ СЛОВАРЬ ФЛОТА

const fileInput = document.getElementById('fileInput');
const DICT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxT4cGlFO8YcDzdeLaqSpThqgYbTbmhDoT8LSaB4FDNsLy0cGgsCa_V-zMINs3WhpcIEA/exec';
const RESULTS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzvbyu5rzhhFiezY6_rNN9-51XZ2h0UBFx0RDMxnGif_XRz_LtU7gWOJ28_RDT4STD3vQ/exec';

fileInput.disabled = true; 

async function loadRouteSchemas() {
    const label = document.getElementById('fileInputLabel');
    label.classList.add('disabled');
    fileInput.disabled = true;
    document.getElementById('fileStatus').innerText = "Завантаження даних з хмари...";

    try {
        const response = await fetch(DICT_SCRIPT_URL);
        const data = await response.json();
        routeDictionary = data.routes;
        yardDictionary = data.yards;
        fleetDictionary = data.fleet || {}; // Грузим флот
        console.log("Довідники завантажені");
        
        await loadSavedYardsList();
        
    } catch (e) {
        console.error("Ошибка справочников:", e);
        document.getElementById('fileStatus').innerText = "Помилка завантаження довідників!";
    } finally {
        label.classList.remove('disabled');
        fileInput.disabled = false;
        
        if (document.getElementById('fileStatus').innerText === "Завантаження даних з хмари...") {
            document.getElementById('fileStatus').innerText = "Система готова до роботи.";
        }
    }
}

async function loadSavedYardsList() {
    try {
        const response = await fetch(RESULTS_SCRIPT_URL + '?action=getYards');
        const data = await response.json();
        
        if (data.yards && data.yards.length > 0) {
            const yardSelect = document.getElementById('kamagYardSelect');
            yardSelect.innerHTML = '<option value="" disabled selected>-- Оберіть автодвір --</option>';
            data.yards.forEach(y => {
                const opt = document.createElement('option');
                opt.value = opt.textContent = y;
                yardSelect.appendChild(opt);
            });
            document.getElementById('fileStatus').innerText = "Список збережених автодворів підвантажено.";
        }
    } catch (e) {
        console.error("Помилка завантаження списку автодворів:", e);
    }
}

// Загрузка данных двора (Адаптировано под флот Kamag+МАН)
document.getElementById('loadGoogleYardBtn').addEventListener('click', async () => {
    const yard = document.getElementById('kamagYardSelect').value;
    const btn = document.getElementById('loadGoogleYardBtn');
    btn.innerText = "⏳...";

    try {
        const response = await fetch(`${RESULTS_SCRIPT_URL}?action=getAggregatedData&yard=${encodeURIComponent(yard)}`);
        const data = await response.json();

        if (data.savedRows && data.savedRows.length > 0) {
            fleetActiveState[yard] = {};
            totalOpsData[yard] = {};

            const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;
            const availM = fleetDictionary[yard] ? fleetDictionary[yard].man : 0;

            // 1. Знаходимо максимальну потребу з урахуванням розбитого формату
            let maxLoadedK = availK;
            let maxLoadedM = availM;
            data.savedRows.forEach(row => {
                let countStr = String(row[3]);
                let separator = countStr.includes('|') ? '|' : (countStr.includes('.') ? '.' : ',');
                const counts = countStr.split(separator);
                
                maxLoadedK = Math.max(maxLoadedK, parseInt(counts[0], 10) || 0);
                maxLoadedM = Math.max(maxLoadedM, parseInt(counts[1], 10) || 0);
            });

            // 2. Відновлюємо стан
            data.savedRows.forEach(row => {
                let [y, day, hour, fleetCountStr, ops] = row;
                
                let dayStr = String(day);
                if (dayStr.includes('T') && dayStr.includes('Z')) {
                    const d = new Date(dayStr);
                    dayStr = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
                }

                if (!fleetActiveState[yard][dayStr]) {
                    fleetActiveState[yard][dayStr] = Array(24).fill(null).map(() => ({ 
                        kamag: Array(maxLoadedK).fill(false), 
                        man: Array(maxLoadedM).fill(false) 
                    }));
                    totalOpsData[yard][dayStr] = Array(24).fill(0);
                }

                totalOpsData[yard][dayStr][hour] = parseInt(ops, 10) || 0;
                
                // --- ФІКС МАНІВ: Розумний парсинг ---
                let countStr = String(fleetCountStr);
                let separator = countStr.includes('|') ? '|' : (countStr.includes('.') ? '.' : ',');
                const counts = countStr.split(separator);
                
                const countK = parseInt(counts[0], 10) || 0;
                const countM = parseInt(counts[1], 10) || 0;
                
                for(let k = 0; k < countK; k++) fleetActiveState[yard][dayStr][hour].kamag[k] = true;
                for(let m = 0; m < countM; m++) fleetActiveState[yard][dayStr][hour].man[m] = true;
            });

            renderKamagTable();
            document.getElementById('fileStatus').innerText = `Дані ${yard} завантажені!`;
        } else {
            alert("Даних для цього автодвору не знайдено.");
        }
    } catch (e) {
        alert("Помилка завантаження");
    } finally {
        btn.innerText = "Завантажити з бази";
    }
});

window.addEventListener('DOMContentLoaded', loadRouteSchemas);

class Schedule {
    constructor(row) {
        this.dateStart = parseExcelDate(row[colIdx.dateStart]); // <--- Читаем старт
        this.dateEnd = parseExcelDate(row[colIdx.dateEnd]);     // <--- Читаем финиш
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

// Глобальные массивы
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
    filteredAllSchedules = [...allSchedules];
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

// Модалка неизвестных маршрутов
const unknownModal = document.getElementById('unknownRoutesModal');
const closeBtn = document.querySelector('.close-btn');
const copyBtn = document.getElementById('copyUnknownBtn');
let currentUnknownRoutesText = ""; 

document.getElementById('unknownRoutesBtn').addEventListener('click', () => {
    const container = document.getElementById('unknownRoutesTableContainer');
    const unknownSet = new Set();
    const routesWithBDF = new Set();
    
    allSchedules.forEach(item => {
        if (item.vehicleType === "Шасі BDF") {
            routesWithBDF.add(item.route);
        }
    });

    allSchedules.forEach(item => {
        if (item.schema === "Схема не знайдена") {
            unknownSet.add(`${item.route}|${item.vehicleType}`);
        }
    });

    if (unknownSet.size === 0) {
        container.innerHTML = "<p style='padding: 10px; color: green;'>Всі маршрути мають схему в довіднику!</p>";
        copyBtn.style.display = 'none'; 
        currentUnknownRoutesText = "";
    } else {
        copyBtn.style.display = 'inline-block'; 
        let tableRows = "";
        let copyLines = [];
        
        unknownSet.forEach(entry => {
            const [route, vehicle] = entry.split('|');
            const hasBdf = routesWithBDF.has(route) ? "Так" : "Ні";
            const bdfColor = hasBdf === "Так" ? "#2e7d32" : "#c62828"; 
            
            tableRows += `<tr>
                <td style="text-align: left; padding: 6px 10px;">${route}</td>
                <td style="text-align: center; padding: 6px 10px;">${vehicle}</td>
                <td style="text-align: center; padding: 6px 10px; font-weight: bold; color: ${bdfColor};">${hasBdf}</td>
            </tr>`;
            copyLines.push(`${route}\t${vehicle}\t${hasBdf}`); 
        });
        
        currentUnknownRoutesText = copyLines.join('\n');
        let html = `<table style="width: 100%; table-layout: fixed;">
            <thead>
                <tr>
                    <th style="width: 60%;">Маршрут</th>
                    <th style="width: 20%;">Тип ТЗ</th>
                    <th style="width: 20%;">Буває Шасі BDF?</th>
                </tr>
            </thead>
            <tbody>${tableRows}</tbody>
        </table>`;
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

// Вкладки
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
tabDetailed.addEventListener('click', () => switchTab(tabDetailed, containerDetailed));
tabEvents.addEventListener('click', () => switchTab(tabEvents, containerEvents));
tabKamag.addEventListener('click', () => {
    switchTab(tabKamag, containerKamag);
    // Як тільки вкладка стала видимою (display: flex), перемальовуємо матрицю, 
    // щоб графіки Chart.js могли правильно розрахувати свою ширину
    if (Object.keys(totalOpsData).length > 0) {
        renderKamagTable();
    }
});

// Генерация


const dayNames = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];

const generateModal = document.getElementById('generateModal');
let generatedDatesList = [];

// 1. Открытие модалки по кнопке "Порахуй"
document.getElementById('generateDetailedBtn').addEventListener('click', () => {
    if (allSchedules.length === 0) return alert("Спочатку завантажте вихідні графіки!");

    // Заполняем список автодворов из словаря
    const yardSelect = document.getElementById('genYardSelect');
    yardSelect.innerHTML = '<option value="ALL">Всі автодвори (повний розрахунок)</option>';
    
    const uniqueYards = Object.keys(yardDictionary).map(k => yardDictionary[k].yard).filter((v, i, a) => v && a.indexOf(v) === i).sort();
    uniqueYards.forEach(y => {
        const opt = document.createElement('option');
        opt.value = opt.textContent = y;
        yardSelect.appendChild(opt);
    });

    // Устанавливаем даты по умолчанию (Сегодня -> +6 дней)
    const today = new Date();
    document.getElementById('genDateStart').value = today.toISOString().split('T')[0];
    const nextWeek = new Date(today);
    nextWeek.setDate(today.getDate() + 6);
    document.getElementById('genDateEnd').value = nextWeek.toISOString().split('T')[0];

    generateModal.style.display = 'block';
});

// Закрытие модалки
document.getElementById('closeGenerateModal').addEventListener('click', () => { generateModal.style.display = 'none'; });

// 2. Кнопка расчета ВНУТРИ модалки
document.getElementById('confirmGenerateBtn').addEventListener('click', () => {
    const yardOpt = document.getElementById('genYardSelect').value;
    const dStart = new Date(document.getElementById('genDateStart').value);
    const dEnd = new Date(document.getElementById('genDateEnd').value);
    const useDates = document.getElementById('genUseScheduleDates').checked;

    if (isNaN(dStart) || isNaN(dEnd) || dStart > dEnd) return alert("Некоректний діапазон дат!");

    generateModal.style.display = 'none';

    // Формируем список дат для расчета
    generatedDatesList = [];
    let curr = new Date(dStart);
    while (curr <= dEnd) {
        generatedDatesList.push(new Date(curr));
        curr.setDate(curr.getDate() + 1);
    }

    const btn = document.getElementById('generateDetailedBtn');
    const originalText = btn.innerText;
    btn.innerText = "⏳ Генерація..."; btn.disabled = true;

    setTimeout(() => {
        generateDetailedSchedules(yardOpt, useDates);
        calculateRampTimes();
        calculateUnloadingTimes();
        
        filteredDetailedSchedules = [...detailedSchedules];
        initDetailedTable();

        generateYardEvents();
        filteredYardEvents = [...yardEvents];
        initEventsTable();

        calculateFleetRequirements(); 
        
        tabDetailed.click();
        btn.innerText = originalText;
        btn.disabled = false;
    }, 50);
});

function formatDateToDDMMYYYY(dateObj) {
    return `${String(dateObj.getDate()).padStart(2, '0')}.${String(dateObj.getMonth() + 1).padStart(2, '0')}.${dateObj.getFullYear()}`;
}

// 3. НОВАЯ ЛОГИКА ГЕНЕРАЦИИ ПО ДАТАМ
function generateDetailedSchedules(targetYard, useDates) {
    detailedSchedules = [];
    
    allSchedules.forEach(item => {
        if (!item.schema || item.schema === "ХЗ" || item.schema === "Схема не знайдена" || item.schema === "—") return;

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

        const cleanSchema = item.schema.toString().replace(/\s+/g, '');
        const miniSchemas = [];
        for (let i = 0; i < cleanSchema.length; i += 3) {
            miniSchemas.push(cleanSchema.substring(i, i + 3));
        }

        // --- ГЛАВНАЯ МАГИЯ: ПРОХОДИМ ПО ВЫБРАННЫМ ДАТАМ ---
        generatedDatesList.forEach(targetDate => {
            
            // 1. Проверка по датам графика (столбцы B и C)
            if (useDates) {
                const tDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate()).getTime();
                const sDate = item.dateStart ? new Date(item.dateStart.getFullYear(), item.dateStart.getMonth(), item.dateStart.getDate()).getTime() : 0;
                const eDate = item.dateEnd ? new Date(item.dateEnd.getFullYear(), item.dateEnd.getMonth(), item.dateEnd.getDate()).getTime() : Infinity;
                
                if (tDate < sDate) return; // График еще не начался
                if (item.dateEnd && tDate > eDate) return; // График уже закончился
            }

            // 2. Проверка по дню недели
            // В JS: 0=Вс, 1=Пн. У нас массив days: 0=Пн, 6=Вс.
            const dayOfWeekIdx = (targetDate.getDay() + 6) % 7; 
            if (!item.days[dayOfWeekIdx]) return; // В этот день недели не ездит

            const dateString = formatDateToDDMMYYYY(targetDate);

            miniSchemas.forEach(mini => {
                if (mini.length < 3) return; 

                const startIndex = parseInt(mini[0], 10) - 1; 
                const endIndex = parseInt(mini[1], 10) - 1;
                const containerType = mini.substring(2);

                if (isNaN(startIndex) || isNaN(endIndex) || startIndex < 0 || endIndex < 0 || startIndex >= activeNodes.length || endIndex >= activeNodes.length) return; 

                const nodeA = activeNodes[startIndex];
                const nodeB = activeNodes[endIndex];

                if (!nodeA || !nodeB) return;

                const yardDataA = yardDictionary[nodeA.name];
                const yardDataB = yardDictionary[nodeB.name];
                
                // Фильтр по автодвору
                if (targetYard !== "ALL" && (!yardDataA || yardDataA.yard !== targetYard) && (!yardDataB || yardDataB.yard !== targetYard)) return;

                let arrMins = getAbsoluteMinutes(dateString, nodeB.timeArr);
                let finalArrivalB = formatAbsoluteMinutes(arrMins);

                detailedSchedules.push({
                    originalRoute: item.route,
                    originalCode: item.code, 
                    day: dateString, // <--- ТЕПЕРЬ ТУТ "14.05.2026"
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
                    moveType: item.moveType 
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

function getAbsoluteMinutes(dateStr, timeStr) {
    if (!timeStr || typeof timeStr !== 'string' || !timeStr.includes(':')) return Infinity; 
    
    const parts = timeStr.trim().split(' ');
    const targetDateStr = parts.length === 2 ? parts[0] : dateStr;
    const targetTimeStr = parts.length === 2 ? parts[1] : timeStr;

    // Парсим дату "DD.MM.YYYY"
    const [dd, mm, yyyy] = targetDateStr.split('.');
    const [hh, min] = targetTimeStr.split(':').map(Number);
    
    // Создаем реальный объект даты и возвращаем минуты с 1970 года
    const dateObj = new Date(yyyy, mm - 1, dd, hh, min);
    return Math.floor(dateObj.getTime() / 60000); 
}

function formatAbsoluteMinutes(mins) {
    if (mins === Infinity || isNaN(mins)) return "—";
    
    const d = new Date(mins * 60000);
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    const hh = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    
    return `${dd}.${mm}.${yyyy} ${hh}:${min}`;
}

// НОВАЯ ЛОГИКА "ПЕРШОЇ ПОСТАНОВКИ" (ПО ДНЯМ)
function calculateRampTimes() {
    const interval = parseInt(document.getElementById('rampInterval').value, 10) || 10;
    
    // Группируем по узлу и дню!
    const groups = {};
    detailedSchedules.forEach(item => {
        if (item.timeDepartureA === "—") return;
        const key = `${item.nodeA}_${item.day}`;
        if (!groups[key]) groups[key] = [];
        item.absDep = getAbsoluteMinutes(item.day, item.timeDepartureA);
        item.timeDepartureA = formatAbsoluteMinutes(item.absDep); // Форматируем сразу
        groups[key].push(item);
    });
    
    for (const key in groups) {
        const group = groups[key];
        group.sort((a, b) => a.absDep - b.absDep || a.originalRoute.localeCompare(b.originalRoute));
        if (group.length === 0) continue;
        
        const yardConf = yardDictionary[group[0].nodeA];
        const hasFixedFirst = yardConf && yardConf.firstPlacement && yardConf.firstPlacement !== "0:00";
        const firstAbs = hasFixedFirst ? getAbsoluteMinutes(group[0].day, yardConf.firstPlacement) : 0;

        let prevAbsDep = group[0].absDep - 10080; 
        
        for (let i = 0; i < group.length; i++) {
            let item1 = group[i];
            let item2 = (i + 1 < group.length) ? group[i + 1] : null;
            
            let isTwin = item2 && 
                         item1.originalRoute === item2.originalRoute && 
                         item1.day === item2.day && 
                         item1.absDep === item2.absDep;

            let proposedPlacement;
            // Если это первый контейнер дня и есть фиксированное время
            if (i === 0 && hasFixedFirst) {
                proposedPlacement = firstAbs;
            } else {
                proposedPlacement = prevAbsDep + interval;
            }

            let maxPlacementTime = item1.absDep - (23 * 60); 
            let finalPlacement = Math.max(proposedPlacement, maxPlacementTime);
            
            if (isTwin) {
                let totalLoadTime = item1.absDep - finalPlacement;
                if (totalLoadTime <= 120) {
                    item1.timePlacementA = formatAbsoluteMinutes(finalPlacement);
                    item2.timePlacementA = formatAbsoluteMinutes(finalPlacement);
                    prevAbsDep = item1.absDep;
                } else {
                    let tHalf = Math.floor((totalLoadTime - 10) / 2);
                    item1.timePlacementA = formatAbsoluteMinutes(finalPlacement);
                    item1.absDep = finalPlacement + tHalf; 
                    item1.timeDepartureA = formatAbsoluteMinutes(item1.absDep);
                    item2.timePlacementA = formatAbsoluteMinutes(finalPlacement + tHalf + 10);
                    prevAbsDep = item2.absDep;
                }
                i++; 
            } else {
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

// Сортировка
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
        const timeCols = ['timePlacementA', 'timeDepartureA', 'timeArrivalB', 'timeUnloadStart', 'timeUnloadEnd'];
        if (timeCols.includes(k)) return getAbsoluteMinutes(item.day, item[k]);
        if (k === 'day') {
            if (!item.day || !item.day.includes('.')) return Infinity;
            const [dd, mm, yyyy] = item.day.split('.');
            return new Date(yyyy, mm - 1, dd).getTime();
        }
        return item[k] !== undefined && item[k] !== null ? item[k] : "";
    };

    detailedSchedules.sort((a, b) => {
        const valA = getValue(a, key);
        const valB = getValue(b, key);
        if (typeof valA === 'number' && typeof valB === 'number') return asc ? valA - valB : valB - valA;
        return asc ? String(valA).localeCompare(String(valB)) : String(valB).localeCompare(String(valA));
    });

    const tr = thElement.parentElement;
    Array.from(tr.children).forEach(th => th.classList.remove('sort-asc', 'sort-desc'));
    thElement.classList.add(asc ? 'sort-asc' : 'sort-desc');

    applyFiltersDetailed();
    document.getElementById('tableContainerDetailed').scrollTop = 0; 
}

// ГЕНЕРАЦИЯ СОБЫТИЙ С ПРОВЕРКОЙ НОВЫХ ФЛАГОВ (1, 2, 3, 4)
function generateYardEvents() {
    yardEvents = [];
    
    // 1. Сразу формируем список разрешенных дат (те, что выбрал логист в модалке)
    const allowedDates = generatedDatesList.map(d => {
        return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
    });

    const addEvent = (yard, nodeName, eventIndex, eventName, absMins, code) => {
        if (absMins === Infinity || isNaN(absMins)) return;
        const yardConf = yardDictionary[nodeName];
        let flag = 0;
        
        if (yardConf) {
            if (eventIndex === 1) flag = yardConf.event1;
            else if (eventIndex === 2) flag = yardConf.event2;
            else if (eventIndex === 3) flag = yardConf.event3;
            else if (eventIndex === 4) flag = yardConf.event4;
        }
        
        if (flag === 1) {
            const formatted = formatAbsoluteMinutes(absMins); 
            const parts = formatted.split(' '); 
            
            // --- ГЛАВНЫЙ ФИЛЬТР ---
            // Если дата события перевалила за полночь и не входит в выбранный нами диапазон - отбрасываем!
            if (allowedDates.length > 0 && !allowedDates.includes(parts[0])) return;

            yardEvents.push({ yard: yard, code: code, event: eventName, day: parts[0], time: parts[1], absMins: absMins });
        }
    };

    detailedSchedules.forEach(item => {
        if (item.vehicle !== "Шасі BDF") return;

        if (item.moveType && item.moveType.toLowerCase().includes("порожній")) {
            if (item.yardB && item.yardB !== "—") {
                addEvent(item.yardB, item.nodeB, 4, "4. Забір", getAbsoluteMinutes(item.day, item.timeArrivalB), item.originalCode);
            }
            return; 
        }

        if (item.yardA && item.yardA !== "—") {
            if (item.timePlacementA && item.timePlacementA !== "—") {
                const parts = item.timePlacementA.split(' ');
                addEvent(item.yardA, item.nodeA, 1, "1. Постановка", getAbsoluteMinutes(parts[0], parts[1]), item.originalCode);
            }
            if (item.timeDepartureA && item.timeDepartureA !== "—") {
                addEvent(item.yardA, item.nodeA, 2, "2. Забір", item.absDep - 15, item.originalCode);
            }
        }
        
        if (item.yardB && item.yardB !== "—") {
            if (item.timeUnloadStart && item.timeUnloadStart !== "—") {
                const parts = item.timeUnloadStart.split(' ');
                addEvent(item.yardB, item.nodeB, 3, "3. Постановка", getAbsoluteMinutes(parts[0], parts[1]), item.originalCode);
            }
            
            if (item.timeUnloadEnd && item.timeUnloadEnd !== "—") {
                const parts = item.timeUnloadEnd.split(' ');
                addEvent(item.yardB, item.nodeB, 4, "4. Забір", getAbsoluteMinutes(parts[0], parts[1]), item.originalCode);
            }
        }
    });
    
    yardEvents.sort((a, b) => a.absMins - b.absMins);
    assignKamagsToEvents();
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
    assignKamagsToEvents();
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

// ==========================================
// НОВЫЙ БЛОК РАСЧЕТА ФЛОТА (KamagИ И МАНЫ)
// ==========================================
let totalOpsData = {}; 
let fleetActiveState = {}; 

function calculateFleetRequirements() {
    totalOpsData = {}; 
    fleetActiveState = {}; 

    // Считаем все операции
    yardEvents.forEach(ev => {
        if (!totalOpsData[ev.yard]) totalOpsData[ev.yard] = {};
        if (!totalOpsData[ev.yard][ev.day]) totalOpsData[ev.yard][ev.day] = Array(24).fill(0);
        const hour = parseInt(ev.time.split(':')[0], 10);
        if (!isNaN(hour)) totalOpsData[ev.yard][ev.day][hour]++;
    });

    const yardNorms = {};
    for(let node in yardDictionary) {
        let y = yardDictionary[node].yard;
        if(!yardNorms[y]) yardNorms[y] = { k: yardDictionary[node].normKamag || 12, m: yardDictionary[node].normMan || 6 };
    }

    // Распределяем и вычисляем дефицит
    for (let yard in totalOpsData) {
        fleetActiveState[yard] = {};
        const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;
        const availM = fleetDictionary[yard] ? fleetDictionary[yard].man : 0;
        const normK = yardNorms[yard] ? yardNorms[yard].k : 12;
        const normM = yardNorms[yard] ? yardNorms[yard].m : 6;

        // Узнаем, сколько МАКСИМУМ дополнительных Kamagов понадобится на этой неделе
        let maxExtraK = 0;
        for (let day in totalOpsData[yard]) {
            for (let h = 0; h < 24; h++) {
                let ops = totalOpsData[yard][day][h];
                let cap = (availK * normK) + (availM * normM);
                if (ops > cap) {
                    let extra = Math.ceil((ops - cap) / normK);
                    if (extra > maxExtraK) maxExtraK = extra;
                }
            }
        }

        const totalK = availK + maxExtraK; // Физические + Потребность

        for (let day in totalOpsData[yard]) {
            fleetActiveState[yard][day] = Array(24).fill(null).map(() => ({ 
                kamag: Array(totalK).fill(false), 
                man: Array(availM).fill(false) 
            }));
            
            for (let h = 0; h < 24; h++) {
                let neededOps = totalOpsData[yard][day][h];
                
                // 1. Насыщаем ФИЗИЧЕСКИЕ Kamagи
                let assignedK = 0;
                while (neededOps > 0 && assignedK < availK) {
                    fleetActiveState[yard][day][h].kamag[assignedK] = true;
                    assignedK++;
                    neededOps -= normK;
                }
                
                // 2. Насыщаем ФИЗИЧЕСКИЕ МАНы
                let assignedM = 0;
                while (neededOps > 0 && assignedM < availM) {
                    fleetActiveState[yard][day][h].man[assignedM] = true;
                    assignedM++;
                    neededOps -= normM;
                }

                // 3. Если все еще не хватает - насыщаем ВИРТУАЛЬНЫЕ Kamagи (Потреба)
                while (neededOps > 0 && assignedK < totalK) {
                    fleetActiveState[yard][day][h].kamag[assignedK] = true;
                    assignedK++;
                    neededOps -= normK;
                }
            }
        }
    }

    const yardSelect = document.getElementById('kamagYardSelect');
    const currentVal = yardSelect.value;
    yardSelect.innerHTML = "";
    Object.keys(totalOpsData).sort().forEach(yard => {
        const option = document.createElement('option');
        option.value = option.textContent = yard;
        yardSelect.appendChild(option);
    });
    if (currentVal && totalOpsData[currentVal]) yardSelect.value = currentVal;

    renderKamagTable();
}

function renderKamagTable() {
    const yard = document.getElementById('kamagYardSelect').value;
    const wrapper = document.getElementById('kamagTableWrapper');

    if (!yard || !totalOpsData[yard]) {
        wrapper.innerHTML = "<p style='padding:20px;'>Немає даних для цього автодвору.</p>";
        return;
    }

    const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;
    const availM = fleetDictionary[yard] ? fleetDictionary[yard].man : 0;
    
    // Динамически определяем totalK из стейта
    // Динамически определяем totalK из стейта (по первой доступной дате)
    let totalK = availK;
    const availableDates = Object.keys(fleetActiveState[yard] || {});
    if (availableDates.length > 0) {
        totalK = fleetActiveState[yard][availableDates[0]][0].kamag.length;
    }

    const yardNorms = { k: 12, m: 6 };
    for(let node in yardDictionary) {
        if(yardDictionary[node].yard === yard) {
            yardNorms.k = yardDictionary[node].normKamag || 12;
            yardNorms.m = yardDictionary[node].normMan || 6;
            break;
        }
    }

    // Собираем все уникальные даты, которые есть в данных двора, и сортируем хронологически
    const daysOfWeek = Object.keys(totalOpsData[yard] || {}).sort((a, b) => {
        const [d1, m1, y1] = a.split('.');
        const [d2, m2, y2] = b.split('.');
        return new Date(y1, m1-1, d1) - new Date(y2, m2-1, d2);
    });

    function generateMatrixHTML(title, rowHeaders, dataProvider, includeCharts = false) {
        let html = `<h3 style="margin: 5px 0 5px 0; color: #334155; border-left: 4px solid #ffaa00; padding-left: 10px;">${title}</h3><table><thead><tr><th style="min-width: 120px;"></th>`;
        
        const dayNamesShort = ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
        
        daysOfWeek.forEach(d => {
            const [dd, mm, yyyy] = d.split('.');
            const dateObj = new Date(yyyy, mm - 1, dd);
            const dayName = dayNamesShort[dateObj.getDay()];
            
            html += `<th colspan="25" style="text-align: center; font-weight: bold; background-color: #e9ecef; border-left: 2px solid #6c757d; border-right: 2px solid #6c757d; padding: 4px 0;">
                ${d}<br><span style="font-size: 11px; font-weight: normal; color: #6c757d;">${dayName}</span>
            </th>`;
        });
        
        html += `<th style="text-align: center; line-height: 1.2;">Всього</th></tr><tr><th style="font-size: 10px;">Рядок / Година</th>`;
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
                    
                    if (title.includes("Машини") || title.includes("Флот")) {
                        const isKamag = rowName.startsWith("Kamag");
                        const match = rowName.match(/\d+/);
                        const idx = match ? parseInt(match[0]) - 1 : 0;
                        
                        let isActive = val === 1;
                        let cellClass = "kamag-cell kamag-editable";
                        
                        // Логика раскраски (Синий = Флот, Оранжевый = Дефицит)
                        if (isActive) {
                            if (isKamag && idx >= availK) cellClass += " kamag-active-virtual";
                            else cellClass += " kamag-active";
                        }
                        
                        let dataType = isKamag ? "kamag" : "man";
                        html += `<td class="${cellClass}" style="${borderStyle} cursor:pointer;" data-yard="${yard}" data-day="${d}" data-hour="${h}" data-type="${dataType}" data-index="${idx}">${isActive ? 1 : ''}</td>`;
                        if (isActive) { dailySum++; totalRowSum++; }
                    } else {
                        let cellStyle = "background-color: #fff9c4; font-weight: bold;";
                        if (val !== 0 && val !== "") {
                            html += `<td class="kamag-cell" style="${borderStyle} ${cellStyle}">${val}</td>`;
                            if (typeof val === 'number') { dailySum += val; totalRowSum += val; }
                        } else {
                            html += `<td class="kamag-cell" style="${borderStyle}"></td>`;
                        }
                    }
                }
                html += `<td style="text-align: center; font-weight: bold; background-color: #f1f3f5; border-right: 2px solid #6c757d;">${dailySum > 0 ? dailySum : ''}</td>`;
            });
            html += `<td style="text-align: center; font-weight: bold; background-color: #e9ecef;">${totalRowSum}</td></tr>`;
        });

        if (includeCharts) {
            html += `<tr><td style="font-weight: bold; font-size: 11px;">Графік</td>`;
            daysOfWeek.forEach((d, index) => {
                html += `<td colspan="24" style="border-left: 2px solid #6c757d; vertical-align: bottom; padding: 0; background: #fff;"><div style="height: 60px; width: 100%;"><canvas id="chart_${index}"></canvas></div></td><td style="border-right: 2px solid #6c757d; background-color: #dee2e6;"></td>`;
            });
            html += `<td style="background-color: #e9ecef;"></td></tr>`;
        }
        html += `</tbody></table>`;
        return html;
    }

    // Збираємо рядки для матриці флоту
    const hideVirtual = document.getElementById('hideVirtualFleet').checked;
    const rowHeaders = [];
    
    // 1. Фізичні камаги
    for(let i=1; i<=availK; i++) rowHeaders.push(`Kamag ${i}`);
    
    // 2. Фізичні маневрові
    for(let i=1; i<=availM; i++) rowHeaders.push(`Маневровий ${i}`);
    
    // 3. Віртуальні (додаткові) камаги - показуємо тільки якщо чек-бокс НЕ активний
    if (!hideVirtual) {
        for(let i=availK+1; i<=totalK; i++) rowHeaders.push(`Kamag ${i} (дод.)`);
    }
    
    let fleetHTML = generateMatrixHTML(`Флот`, rowHeaders, (row, day, hour) => {
        if (!fleetActiveState[yard][day] || !fleetActiveState[yard][day][hour]) return 0;
        const isKamag = row.startsWith("Kamag");
        const match = row.match(/\d+/);
        const idx = match ? parseInt(match[0]) - 1 : 0;
        const state = fleetActiveState[yard][day][hour];
        return (isKamag ? state.kamag[idx] : state.man[idx]) ? 1 : 0;
    });

    const opsHTML = generateMatrixHTML(`Операції`, ["Всього операцій", "Непокриті (фіз. флот)", "Непокриті (залишок)"], (row, day, hour) => {
        const totalOps = (totalOpsData[yard] && totalOpsData[yard][day]) ? totalOpsData[yard][day][hour] : 0;
        
        if (row === "Всього операцій") {
            return totalOps;
        }

        let capPhysical = 0;
        let capTotal = 0;

        if (fleetActiveState[yard] && fleetActiveState[yard][day] && fleetActiveState[yard][day][hour]) {
            const st = fleetActiveState[yard][day][hour];
            
            // Вважаємо тільки фізичний флот (сині)
            const activePhysK = st.kamag.slice(0, availK).filter(Boolean).length;
            const activePhysM = st.man.filter(Boolean).length;
            capPhysical = (activePhysK * yardNorms.k) + (activePhysM * yardNorms.m);

            // Вважаємо весь флот (сині + оранжеві)
            const activeTotalK = st.kamag.filter(Boolean).length;
            capTotal = (activeTotalK * yardNorms.k) + (activePhysM * yardNorms.m);
        }

        if (row === "Непокриті (фіз. флот)") {
            const uncoveredPhys = Math.max(0, totalOps - capPhysical);
            // Використовуємо інший ID для кліків
            return `<span id="uncovered_phys_${day}_${hour}" class="${uncoveredPhys > 0 ? 'uncovered-alert' : ''}">${uncoveredPhys > 0 ? uncoveredPhys : ''}</span>`;
        } else {
            const uncoveredAbs = Math.max(0, totalOps - capTotal);
            return `<span id="uncovered_abs_${day}_${hour}" class="${uncoveredAbs > 0 ? 'uncovered-alert' : ''}">${uncoveredAbs > 0 ? uncoveredAbs : ''}</span>`;
        }
    }, true);

    wrapper.innerHTML = fleetHTML + "<div style='height: 10px;'></div>" + opsHTML;

    if (window.myDayCharts) window.myDayCharts.forEach(c => c.destroy());
    window.myDayCharts = [];

    daysOfWeek.forEach((d, index) => {
        const ctx = document.getElementById(`chart_${index}`);
        if (!ctx) return;
        const parentDiv = ctx.parentElement;
        ctx.width = parentDiv.clientWidth; ctx.height = 60;

        const chartLabels = [], opsData = [], capacityData = [];
        for (let h = 0; h < 24; h++) {
            chartLabels.push(`${h}:00`);
            opsData.push((totalOpsData[yard] && totalOpsData[yard][d]) ? totalOpsData[yard][d][h] : 0);
            let cap = 0;
            if (fleetActiveState[yard] && fleetActiveState[yard][d] && fleetActiveState[yard][d][h]) {
                cap += fleetActiveState[yard][d][h].kamag.filter(Boolean).length * yardNorms.k;
                cap += fleetActiveState[yard][d][h].man.filter(Boolean).length * yardNorms.m;
            }
            capacityData.push(cap);
        }

        window.myDayCharts.push(new Chart(ctx, {
            type: 'bar',
            data: { labels: chartLabels, datasets: [
                // ЗМІНЕНО: Прибрали yAxisID: 'y1', тепер лінія використовує ту саму вісь 'y', що й стовпчики
                { type: 'line', label: 'Потужність', data: capacityData, borderColor: '#0d47a1', backgroundColor: '#0d47a1', borderWidth: 2, tension: 0.3, pointRadius: 2, yAxisID: 'y' },
                { type: 'bar', label: 'Операції', data: opsData, backgroundColor: 'rgba(255, 193, 7, 0.7)', borderColor: '#ffaa00', borderWidth: 1, borderRadius: 2, yAxisID: 'y' }
            ]},
            options: { 
                animation: false, 
                responsive: false, 
                maintainAspectRatio: false, 
                interaction: { mode: 'index', intersect: false }, 
                scales: { 
                    x: { display: false }, 
                    // ЗМІНЕНО: Залишили тільки одну вісь Y, прибрали y1 взагалі
                    y: { type: 'linear', display: false, beginAtZero: true } 
                }, 
                plugins: { 
                    legend: { display: false }, 
                    tooltip: { callbacks: { title: (items) => `${d} ${items[0].label}`, label: (item) => `${item.dataset.label}: ${item.raw}` } } 
                }, 
                layout: { padding: 0 } 
            }
        }));
    });
}

document.getElementById('kamagYardSelect').addEventListener('change', renderKamagTable);

// КЛИКИ ПО МАТРИЦЕ ФЛОТА
document.getElementById('kamagTableWrapper').addEventListener('click', function(e) {
    if (e.target.classList.contains('kamag-editable')) {
        if (sessionStorage.getItem('kamagonAuth') !== 'true') return;
        const cell = e.target;
        const yard = cell.getAttribute('data-yard');
        const day = cell.getAttribute('data-day');
        const hour = parseInt(cell.getAttribute('data-hour'));
        const type = cell.getAttribute('data-type'); 
        const idx = parseInt(cell.getAttribute('data-index'));

        const currentState = fleetActiveState[yard][day][hour][type][idx];
        const newState = !currentState;
        fleetActiveState[yard][day][hour][type][idx] = newState;
        
        const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;

        if (newState) {
            if (type === 'kamag' && idx >= availK) cell.classList.add('kamag-active-virtual');
            else cell.classList.add('kamag-active');
            cell.innerText = '1';
        } else {
            cell.classList.remove('kamag-active', 'kamag-active-virtual');
            cell.innerText = '';
        }

        const yardNorms = { k: 12, m: 6 };
        for(let node in yardDictionary) {
            if(yardDictionary[node].yard === yard) {
                yardNorms.k = yardDictionary[node].normKamag || 12;
                yardNorms.m = yardDictionary[node].normMan || 6;
                break;
            }
        }

        const totalOps = totalOpsData[yard][day][hour] || 0;
        
        // Перераховуємо фізичну та загальну потужність
        const capPhysical = fleetActiveState[yard][day][hour].kamag.slice(0, availK).filter(Boolean).length * yardNorms.k +
                            fleetActiveState[yard][day][hour].man.filter(Boolean).length * yardNorms.m;
        const capTotal = fleetActiveState[yard][day][hour].kamag.filter(Boolean).length * yardNorms.k +
                         fleetActiveState[yard][day][hour].man.filter(Boolean).length * yardNorms.m;
        
        // Оновлюємо рядок фізичного дефіциту
        const uncoveredPhys = Math.max(0, totalOps - capPhysical);
        const uncoveredPhysCell = document.getElementById(`uncovered_phys_${day}_${hour}`);
        if (uncoveredPhysCell) {
            uncoveredPhysCell.innerText = uncoveredPhys > 0 ? uncoveredPhys : '';
            uncoveredPhysCell.className = uncoveredPhys > 0 ? 'uncovered-alert' : '';
        }

        // Оновлюємо рядок абсолютного дефіциту
        const uncoveredAbs = Math.max(0, totalOps - capTotal);
        const uncoveredAbsCell = document.getElementById(`uncovered_abs_${day}_${hour}`);
        if (uncoveredAbsCell) {
            uncoveredAbsCell.innerText = uncoveredAbs > 0 ? uncoveredAbs : '';
            uncoveredAbsCell.className = uncoveredAbs > 0 ? 'uncovered-alert' : '';
        }

        // Вычисляем индекс дня динамически на основе реальных дат
        const daysOfYard = Object.keys(totalOpsData[yard] || {}).sort((a, b) => {
            const [d1, m1, y1] = a.split('.');
            const [d2, m2, y2] = b.split('.');
            return new Date(y1, m1-1, d1) - new Date(y2, m2-1, d2);
        });
        const dayIndex = daysOfYard.indexOf(day);
        if (window.myDayCharts && window.myDayCharts[dayIndex]) {
            window.myDayCharts[dayIndex].data.datasets[0].data[hour] = cap;
            window.myDayCharts[dayIndex].update();
        }
    }
});

function assignKamagsToEvents() {
    yardEvents = yardEvents.filter(ev => ev.event !== "Чергування");
    const activeKamagsLog = {}, hourTracker = {};

    yardEvents.forEach(ev => {
        ev.kamag = "—"; 
        const hour = parseInt(ev.time.split(':')[0], 10);
        if (isNaN(hour)) return;

        const key = `${ev.yard}_${ev.day}_${hour}`;
        if (!hourTracker[key]) hourTracker[key] = 0;
        
        const availK = fleetDictionary[ev.yard] ? fleetDictionary[ev.yard].kamag : 0;

        if (fleetActiveState[ev.yard] && fleetActiveState[ev.yard][ev.day] && fleetActiveState[ev.yard][ev.day][hour]) {
            const st = fleetActiveState[ev.yard][ev.day][hour];
            const activeResources = [];
            
            st.kamag.forEach((isActive, idx) => { 
                if(isActive) activeResources.push(idx < availK ? `Kamag ${idx+1}` : `Kamag ${idx+1} (дод.)`); 
            });
            st.man.forEach((isActive, idx) => { if(isActive) activeResources.push(`Маневровий ${idx+1}`); });

            if (activeResources.length > 0) {
                const assignedMachine = activeResources[hourTracker[key] % activeResources.length];
                ev.kamag = assignedMachine;
                activeKamagsLog[`${ev.yard}_${ev.day}_${hour}_${ev.kamag}`] = true;
            } else ev.kamag = "Немає ТЗ!"; 
        }
        hourTracker[key]++;
    });

    for (let y in fleetActiveState) {
        const availK = fleetDictionary[y] ? fleetDictionary[y].kamag : 0;
        for (let d in fleetActiveState[y]) {
            for (let h = 0; h < 24; h++) {
                const st = fleetActiveState[y][d][h];
                if (st) {
                    st.kamag.forEach((isActive, kIndex) => {
                        const name = kIndex < availK ? `Kamag ${kIndex + 1}` : `Kamag ${kIndex + 1} (дод.)`;
                        if (isActive && !activeKamagsLog[`${y}_${d}_${h}_${name}`]) {
                            yardEvents.push({ yard: y, code: "—", event: "Чергування", day: d, time: `${String(h).padStart(2, '0')}:00`, absMins: getAbsoluteMinutes(d, `${String(h).padStart(2, '0')}:00`), kamag: name });
                        }
                    });
                    st.man.forEach((isActive, mIndex) => {
                        const name = `Маневровий ${mIndex + 1}`;
                        if (isActive && !activeKamagsLog[`${y}_${d}_${h}_${name}`]) {
                            yardEvents.push({ yard: y, code: "—", event: "Чергування", day: d, time: `${String(h).padStart(2, '0')}:00`, absMins: getAbsoluteMinutes(d, `${String(h).padStart(2, '0')}:00`), kamag: name });
                        }
                    });
                }
            }
        }
    }
    yardEvents.sort((a, b) => a.absMins - b.absMins);
}
// =========================================
// ФІЛЬТРАЦІЯ ТА ЕКСПОРТ (ВОССТАНОВЛЕНО ПОЛНОСТЬЮ)
// =========================================

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
        "", // Місце під індекс
        item.originalRoute, 
        item.originalCode, 
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
        
        return filters.every(f => {
            const cellStr = String(rowVals[f.col] || "").toLowerCase();
            const searchTerms = f.val.split(',').map(s => s.trim()).filter(Boolean);
            return searchTerms.some(term => cellStr.includes(term));
        });
    });
}

let filterTimeout;
document.addEventListener('input', function(e) {
    if (e.target.classList.contains('filter-input')) {
        clearTimeout(filterTimeout);
        filterTimeout = setTimeout(() => {
            const container = e.target.closest('.table-container');
            if (container.id === 'tableContainerRaw') {
                filteredAllSchedules = filterDataArray('tableContainerRaw', allSchedules, getRawValues);
                renderedCount = 0;
                document.getElementById('tableBody').innerHTML = "";
                renderChunk();
            } else if (container.id === 'tableContainerDetailed') {
                filteredDetailedSchedules = filterDataArray('tableContainerDetailed', detailedSchedules, getDetailedValues);
                detailedRenderedCount = 0;
                document.getElementById('detailedTableBody').innerHTML = "";
                renderDetailedChunk();
            } else if (container.id === 'tableContainerEvents') {
                filteredYardEvents = filterDataArray('tableContainerEvents', yardEvents, getEventsValues);
                eventsRenderedCount = 0;
                document.getElementById('eventsTableBody').innerHTML = "";
                renderEventsChunk();
            }
        }, 300); 
    }
});

// ВОССТАНОВЛЕННЫЙ И ОБНОВЛЕННЫЙ ЭКСПОРТ В EXCEL
// ВОССТАНОВЛЕННЫЙ И ОБНОВЛЕННЫЙ ЭКСПОРТ В EXCEL
document.getElementById('exportExcelBtn').addEventListener('click', async () => {
    const btn = document.getElementById('exportExcelBtn');
    const originalText = btn.innerText;
    btn.innerText = "⏳ Формування...";
    btn.disabled = true;

    try {
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Звіт');
        
        if (tabRaw.classList.contains('active')) {
            const headers = ["Маршрут", "Дедлайн", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд", "Початкова", "Приїзд", "Виїзд"];
            for(let i=1; i<=10; i++) headers.push(`П.Т. №${i}`, "Приїзд", "Виїзд");
            headers.push("Кінцева", "Приїзд", "Вивільнення", "Тип доставки", "Тип ТЗ", "Схема БДФ", "Формат", "Код", "Тип переміщення");
            
            sheet.addRow(headers);
            filteredAllSchedules.forEach(item => sheet.addRow(getRawValues(item)));

        } else if (tabDetailed.classList.contains('active')) {
            const headers = ["№", "Маршрут", "Код", "День", "Схема", "Тип", "Автодвір А", "Постановка", "Точка А", "Виїзд", "Автодвір Б", "Точка Б", "Приїзд", "Постановка (вивант.)", "Кінець вивант.", "Тип ТЗ"];
            sheet.addRow(headers);
            filteredDetailedSchedules.forEach((item, index) => {
                let vals = getDetailedValues(item);
                vals[0] = index + 1;
                sheet.addRow(vals);
            });

        } else if (tabKamag.classList.contains('active')) {
            const yard = document.getElementById('kamagYardSelect').value;
            const days = Object.keys(totalOpsData[yard] || {}).sort((a, b) => {
                const [d1, m1, y1] = a.split('.');
                const [d2, m2, y2] = b.split('.');
                return new Date(y1, m1-1, d1) - new Date(y2, m2-1, d2);
            });
            
            sheet.getColumn(1).width = 18; 
            for (let i = 2; i <= 1 + 25 * 7; i++) {
                sheet.getColumn(i).width = 3.5; 
                if ((i - 1) % 25 === 0) sheet.getColumn(i).width = 5; 
            }
            sheet.getColumn(2 + 25 * 7).width = 8; 

            const alignCenter = { vertical: 'middle', horizontal: 'center' };
            const fillHeader = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } };
            const fillSum = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F3F5' } };
            
            // --- НОВЫЕ ЦВЕТА ---
            const fillActive = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBDEFB' } }; // Синий цвет физического флота
            const fillActiveVirtual = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE0B2' } }; // Оранжевый для требуемого (доп.) флота
            const fillOps = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
            const fillUncovered = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEBEE' } }; // Красный фон для непокрытых операций
            const fontUncovered = { bold: true, size: 9, color: { argb: 'FFD32F2F' } }; // Красный текст

            const borderThin = { style: 'thin', color: { argb: 'FFCCCCCC' } };
            const borderMedium = { style: 'medium', color: { argb: 'FF6C757D' } };

            const getBorders = (isLeftEdge, isRightEdge) => ({
                top: borderThin, bottom: borderThin,
                left: isLeftEdge ? borderMedium : borderThin,
                right: isRightEdge ? borderMedium : borderThin
            });

            sheet.addRow([`Звіт по Флоту: ${yard}`]).font = { bold: true, size: 14 };
            sheet.addRow([]);

            const rowDays = sheet.addRow(["День"]);
            rowDays.getCell(1).font = { bold: true };
            rowDays.getCell(1).alignment = alignCenter;

            const dayNamesShort = ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];

            days.forEach((d, index) => {
                const startCol = 2 + index * 25;
                const endCol = startCol + 24;
                sheet.mergeCells(3, startCol, 3, endCol); 
                
                const [dd, mm, yyyy] = d.split('.');
                const dateObj = new Date(yyyy, mm - 1, dd);
                const dayName = dayNamesShort[dateObj.getDay()];
                
                const cell = sheet.getCell(3, startCol);
                cell.value = `${d} (${dayName})`; // Буде виглядати як "14.05.2026 (Чт)"
                cell.alignment = alignCenter;
                cell.font = { bold: true };
                cell.fill = fillHeader;
                cell.border = getBorders(true, true);
            });
            sheet.getCell(3, 2 + 25 * 7).value = "Всього";
            sheet.getCell(3, 2 + 25 * 7).font = { bold: true };

            const rowHours = sheet.addRow(["Рядок / Година"]);
            rowHours.getCell(1).font = { size: 10 };
            
            let currentCol = 2;
            days.forEach((d, index) => {
                for (let h = 0; h < 24; h++) {
                    const cell = rowHours.getCell(currentCol);
                    cell.value = h; 
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

            // --- ОБНОВЛЕННЫЙ ЭКСПОРТ KamagОВ (с учетом потребности) ---
            const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;
            const availM = fleetDictionary[yard] ? fleetDictionary[yard].man : 0;
            
            // Динамически получаем общее количество Kamagов (включая дополнительные)
            let totalK = availK;
            const availableDates = Object.keys(fleetActiveState[yard] || {});
            if (availableDates.length > 0 && fleetActiveState[yard][availableDates[0]][0]) {
                totalK = fleetActiveState[yard][availableDates[0]][0].kamag.length;
            }

            for (let k = 1; k <= totalK; k++) {
                const rowLabel = k <= availK ? `Kamag ${k}` : `Kamag ${k} (дод.)`;
                const row = sheet.addRow([rowLabel]);
                row.getCell(1).font = { bold: true, size: 10 };
                let weekTotal = 0;
                let cCol = 2;

                days.forEach(d => {
                    let daySum = 0;
                    for (let h = 0; h < 24; h++) {
                        let val = (fleetActiveState[yard][d] && fleetActiveState[yard][d][h].kamag[k-1]) ? 1 : "";
                        const cell = row.getCell(cCol);
                        cell.value = val;
                        cell.alignment = alignCenter;
                        cell.border = getBorders(h === 0, false);
                        
                        if (val === 1) {
                            // Красим физический флот в синий, а дополнительный (потребность) - в оранжевый
                            cell.fill = k <= availK ? fillActive : fillActiveVirtual;
                            daySum++; weekTotal++;
                        }
                        cCol++;
                    }
                    const dSumCell = row.getCell(cCol);
                    dSumCell.value = daySum || "";
                    dSumCell.alignment = alignCenter;
                    dSumCell.font = { bold: true };
                    dSumCell.fill = fillSum;
                    dSumCell.border = getBorders(false, true);
                    cCol++;
                });
                
                const wSumCell = row.getCell(cCol);
                wSumCell.value = weekTotal;
                wSumCell.alignment = alignCenter;
                wSumCell.font = { bold: true };
                wSumCell.fill = fillHeader;
            }

            // ЭКСПОРТ МАНОВ
            for (let m = 1; m <= availM; m++) {
                const row = sheet.addRow([`Маневровий ${m}`]);
                row.getCell(1).font = { bold: true, size: 10 };
                let weekTotal = 0;
                let cCol = 2;

                days.forEach(d => {
                    let daySum = 0;
                    for (let h = 0; h < 24; h++) {
                        let val = (fleetActiveState[yard][d] && fleetActiveState[yard][d][h].man[m-1]) ? 1 : "";
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
                    const dSumCell = row.getCell(cCol);
                    dSumCell.value = daySum || "";
                    dSumCell.alignment = alignCenter;
                    dSumCell.font = { bold: true };
                    dSumCell.fill = fillSum;
                    dSumCell.border = getBorders(false, true);
                    cCol++;
                });
                
                const wSumCell = row.getCell(cCol);
                wSumCell.value = weekTotal;
                wSumCell.alignment = alignCenter;
                wSumCell.font = { bold: true };
                wSumCell.fill = fillHeader;
            }

            sheet.addRow([]);

            // Готовим нормы для расчета "Непокрытых операций"
            const yardNorms = { k: 12, m: 6 };
            for(let node in yardDictionary) {
                if(yardDictionary[node].yard === yard) {
                    yardNorms.k = yardDictionary[node].normKamag || 12;
                    yardNorms.m = yardDictionary[node].normMan || 6;
                    break;
                }
            }

            // --- ЭКСПОРТ ОПЕРАЦИЙ ---
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
                    
                    if (val > 0) cell.fill = fillOps;
                    
                    daySum += val;
                    totalWeekOps += val;
                    oCol++;
                }
                const dSumCell = opsRow.getCell(oCol);
                dSumCell.value = daySum || "";
                dSumCell.alignment = alignCenter;
                dSumCell.font = { bold: true };
                dSumCell.fill = fillSum;
                dSumCell.border = getBorders(false, true);
                oCol++;
            });
            
            const wOpsSumCell = opsRow.getCell(oCol);
            wOpsSumCell.value = totalWeekOps;
            wOpsSumCell.alignment = alignCenter;
            wOpsSumCell.font = { bold: true };
            wOpsSumCell.fill = fillHeader;

            // --- ДОБАВЛЕНО: ЭКСПОРТ НЕПОКРЫТЫХ ОПЕРАЦИЙ (ДЕФИЦИТ) ---
            const uncoveredPhysRow = sheet.addRow(["Непокриті (фіз. флот)"]);
            uncoveredPhysRow.getCell(1).font = { bold: true, size: 10 };
            
            let totalWeekUncoveredPhys = 0;
            let uColPhys = 2;
            
            days.forEach(d => {
                let daySum = 0;
                for (let h = 0; h < 24; h++) {
                    let totalOps = (totalOpsData[yard] && totalOpsData[yard][d]) ? totalOpsData[yard][d][h] : 0;
                    let capPhysical = 0;
                    if (fleetActiveState[yard] && fleetActiveState[yard][d] && fleetActiveState[yard][d][h]) {
                        capPhysical += fleetActiveState[yard][d][h].kamag.slice(0, availK).filter(Boolean).length * yardNorms.k;
                        capPhysical += fleetActiveState[yard][d][h].man.filter(Boolean).length * yardNorms.m;
                    }
                    let uncovered = Math.max(0, totalOps - capPhysical);

                    const cell = uncoveredPhysRow.getCell(uColPhys);
                    cell.value = uncovered || "";
                    cell.alignment = alignCenter;
                    cell.border = getBorders(h === 0, false);
                    if (uncovered > 0) { cell.fill = fillUncovered; cell.font = fontUncovered; }
                    
                    daySum += uncovered;
                    totalWeekUncoveredPhys += uncovered;
                    uColPhys++;
                }
                const dSumCell = uncoveredPhysRow.getCell(uColPhys);
                dSumCell.value = daySum || "";
                dSumCell.alignment = alignCenter; dSumCell.font = { bold: true }; dSumCell.fill = fillSum; dSumCell.border = getBorders(false, true);
                uColPhys++;
            });
            const wUncoveredPhysSumCell = uncoveredPhysRow.getCell(uColPhys);
            wUncoveredPhysSumCell.value = totalWeekUncoveredPhys; wUncoveredPhysSumCell.alignment = alignCenter; wUncoveredPhysSumCell.font = { bold: true }; wUncoveredPhysSumCell.fill = fillHeader;

            // --- ЭКСПОРТ НЕПОКРЫТЫХ ОПЕРАЦИЙ (АБСОЛЮТНЫЙ ЗАЛИШОК) ---
            const uncoveredAbsRow = sheet.addRow(["Непокриті (залишок)"]);
            uncoveredAbsRow.getCell(1).font = { bold: true, size: 10 };
            
            let totalWeekUncoveredAbs = 0;
            let uColAbs = 2;
            
            days.forEach(d => {
                let daySum = 0;
                for (let h = 0; h < 24; h++) {
                    let totalOps = (totalOpsData[yard] && totalOpsData[yard][d]) ? totalOpsData[yard][d][h] : 0;
                    let capTotal = 0;
                    if (fleetActiveState[yard] && fleetActiveState[yard][d] && fleetActiveState[yard][d][h]) {
                        capTotal += fleetActiveState[yard][d][h].kamag.filter(Boolean).length * yardNorms.k;
                        capTotal += fleetActiveState[yard][d][h].man.filter(Boolean).length * yardNorms.m;
                    }
                    let uncovered = Math.max(0, totalOps - capTotal);

                    const cell = uncoveredAbsRow.getCell(uColAbs);
                    cell.value = uncovered || "";
                    cell.alignment = alignCenter;
                    cell.border = getBorders(h === 0, false);
                    if (uncovered > 0) { cell.fill = fillUncovered; cell.font = fontUncovered; }
                    
                    daySum += uncovered;
                    totalWeekUncoveredAbs += uncovered;
                    uColAbs++;
                }
                const dSumCell = uncoveredAbsRow.getCell(uColAbs);
                dSumCell.value = daySum || "";
                dSumCell.alignment = alignCenter; dSumCell.font = { bold: true }; dSumCell.fill = fillSum; dSumCell.border = getBorders(false, true);
                uColAbs++;
            });
            const wUncoveredAbsSumCell = uncoveredAbsRow.getCell(uColAbs);
            wUncoveredAbsSumCell.value = totalWeekUncoveredAbs; wUncoveredAbsSumCell.alignment = alignCenter; wUncoveredAbsSumCell.font = { bold: true }; wUncoveredAbsSumCell.fill = fillHeader;
            wUncoveredSumCell.alignment = alignCenter;
            wUncoveredSumCell.font = { bold: true };
            wUncoveredSumCell.fill = fillHeader;

            sheet.addRow([]);
            sheet.addRow(["Графіки:"]).font = { bold: true };
            
            const imgRow = sheet.rowCount; 

            for (let i = 0; i < 7; i++) {
                const canvas = document.getElementById(`chart_${i}`);
                if (canvas) {
                    const base64 = canvas.toDataURL('image/png');
                    const imageId = workbook.addImage({
                        base64: base64,
                        extension: 'png',
                    });
                    
                    sheet.addImage(imageId, {
                        tl: { col: 1 + i * 25, row: imgRow },
                        ext: { width: 620, height: 100 } 
                    });
                }
            }
            
            for(let i=0; i<6; i++) sheet.addRow([]); 
        }

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


// СОХРАНЕНИЕ
document.getElementById('saveGoogleBtn').addEventListener('click', async () => {
    const yard = document.getElementById('kamagYardSelect').value;
    if (!yard) return alert("Оберіть автодвір!");

    const btn = document.getElementById('saveGoogleBtn');
    btn.innerText = "⏳ Збереження...";

    const aggregatedRows = [];
    const days = Object.keys(totalOpsData[yard] || {});

    days.forEach(day => {
        for (let h = 0; h < 24; h++) {
            const opsCount = (totalOpsData[yard] && totalOpsData[yard][day]) ? totalOpsData[yard][day][h] : 0;
            
            let actK = 0, actM = 0;
            if (fleetActiveState[yard] && fleetActiveState[yard][day] && fleetActiveState[yard][day][h]) {
                actK = fleetActiveState[yard][day][h].kamag.filter(Boolean).length;
                actM = fleetActiveState[yard][day][h].man.filter(Boolean).length;
            }
            
            if (opsCount > 0 || actK > 0 || actM > 0) aggregatedRows.push([yard, day, h, `${actK}|${actM}`, opsCount]);
        }
    });

    try {
        await fetch(RESULTS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'saveAggregated', yard: yard, rows: aggregatedRows })
        });
        btn.innerText = "✅ Збережено!";
    } catch (e) {
        btn.innerText = "❌ Помилка";
    }
    setTimeout(() => btn.innerText = "Зберегти (поточний)", 3000);
});

document.getElementById('saveAllGoogleBtn').addEventListener('click', async () => {
    const yards = Object.keys(totalOpsData);
    if (yards.length === 0) return alert("Немає розрахованих даних для збереження!");

    const btn = document.getElementById('saveAllGoogleBtn');
    btn.innerText = "⏳ Збереження...";
    btn.disabled = true;

    const aggregatedRows = [];
    //const days = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];

    yards.forEach(yard => {
        const days = Object.keys(totalOpsData[yard] || {});
        days.forEach(day => {
            for (let h = 0; h < 24; h++) {
                const opsCount = (totalOpsData[yard] && totalOpsData[yard][day]) ? totalOpsData[yard][day][h] : 0;
                
                let actK = 0, actM = 0;
                if (fleetActiveState[yard] && fleetActiveState[yard][day] && fleetActiveState[yard][day][h]) {
                    actK = fleetActiveState[yard][day][h].kamag.filter(Boolean).length;
                    actM = fleetActiveState[yard][day][h].man.filter(Boolean).length;
                }
                
                if (opsCount > 0 || actK > 0 || actM > 0) aggregatedRows.push([yard, day, h, `${actK}|${actM}`, opsCount]);
            }
        });
    });

    try {
        await fetch(RESULTS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'saveAllAggregated', yards: yards, rows: aggregatedRows })
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
});

// Автоматична активація вкладки при завантаженні
document.addEventListener('DOMContentLoaded', () => {
    tabKamag.click();
});

document.getElementById('hideVirtualFleet').addEventListener('change', renderKamagTable);