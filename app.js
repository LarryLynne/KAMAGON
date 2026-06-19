// --- ГЛОБАЛЬНЫЕ СЛОВАРИ ---
let routeDictionary = {};
let yardDictionary = {};
let fleetDictionary = {}; 
let usersDictionary = {}; // СЛОВАРЬ ПОЛЬЗОВАТЕЛЕЙ
let totalOpsData = {}; 
let fleetActiveState = {}; 
let systemFleetState = {}; // Базовое состояние автоматического расчета системы

function checkAuth() {
    return new Promise((resolve) => {
        if (sessionStorage.getItem('kamagonAuth') === 'true') {
            resolve(true);
            return;
        }
        
        if (Object.keys(usersDictionary).length === 0) {
            alert("Дочекайтеся завантаження даних!");
            resolve(false);
            return;
        }
        
        const modal = document.getElementById('authModal');
        const loginInput = document.getElementById('authLoginInput');
        const passInput = document.getElementById('authPassInput');
        
        modal.style.display = 'block';
        loginInput.value = '';
        passInput.value = '';
        loginInput.focus();

        document.getElementById('authConfirmBtn').onclick = () => {
            const login = loginInput.value.trim();
            const pass = passInput.value.trim();

            const userObj = usersDictionary[login];
            if (userObj && String(userObj.pass) === pass) {
                sessionStorage.setItem('kamagonAuth', 'true');
                sessionStorage.setItem('kamagonAuthUser', login);
                sessionStorage.setItem('kamagonAuthRole', userObj.role); // Сохраняем роль
                sessionStorage.setItem('kamagonAuthYard', userObj.yard); // Сохраняем двор
                modal.style.display = 'none';
                updateAuthVisibility();
                location.reload(); // Перезагружаем для применения жестких ограничений вкладок
                resolve(true);
            } else {
                alert("Невірний логін або пароль!");
                passInput.value = '';
                passInput.focus();
            }
        };

        document.getElementById('authCancelBtn').onclick = () => {
            modal.style.display = 'none';
            resolve(false);
        };
    });
}

// ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ДЛЯ НАМЕРТВО ФИКСИРОВАНИЯ АВТОДВОРА ПОЛЬЗОВАТЕЛЯ РДУ
function enforceUserYardLock(userYard) {
    if (!userYard) return;
    ['kamagYardSelect', 'factYardSelect', 'compareYardSelect'].forEach(id => {
        const select = document.getElementById(id);
        if (select) {
            // Если в списке вдруг нет такого двора, принудительно создаем его
            let optionExists = Array.from(select.options).some(opt => opt.value === userYard);
            if (!optionExists) {
                const opt = document.createElement('option');
                opt.value = opt.textContent = userYard;
                select.appendChild(opt);
            }
            select.value = userYard;
            select.disabled = true; // Запрещаем менять
            select.dispatchEvent(new Event('change'));
        }
    });
}

// МОДЕРНИЗИРОВАННЫЙ КОНТРОЛЬ ДОСТУПА И ВИДИМОСТИ ВКЛАДОК
// --- Найти и заменить функцию updateAuthVisibility в app.js ---
function updateAuthVisibility() {
    const isAuth = sessionStorage.getItem('kamagonAuth') === 'true';
    const activeUser = sessionStorage.getItem('kamagonAuthUser');
    const role = sessionStorage.getItem('kamagonAuthRole');
    const userYard = sessionStorage.getItem('kamagonAuthYard');
    const loginBtn = document.getElementById('loginBtn');
    
    // 1. Базово переключаем элементы с классом auth-hidden
    document.querySelectorAll('.auth-hidden').forEach(el => {
        if (isAuth) el.classList.remove('auth-hidden');
    });

    // Хватай все элементы интерфейса для разграничения прав
    const tabRaw = document.getElementById('tabRaw');
    const tabDetailed = document.getElementById('tabDetailed');
    const tabEvents = document.getElementById('tabEvents');
    const tabKamag = document.getElementById('tabKamag');
    const tabFact = document.getElementById('tabFact');
    const tabRdu = document.getElementById('tabRdu');

    const fileInputLabel = document.getElementById('fileInputLabel');
    const saveGoogleBtn = document.getElementById('saveGoogleBtn');
    const saveAllGoogleBtn = document.getElementById('saveAllGoogleBtn');
    const tabCompare = document.getElementById('tabCompare');

    if (isAuth) {
        if (isAuth) {
            if (role === 'РДУ') {
                // РДУ видит три вкладки: Розрахунок, Введення РДУ и Звірка
                if (tabRaw) tabRaw.style.display = 'none';
                if (tabDetailed) tabDetailed.style.display = 'none';
                if (tabEvents) tabEvents.style.display = 'none';
                if (tabFact) tabFact.style.display = 'none';
                
                if (tabKamag) tabKamag.style.display = 'block';
                if (tabRdu) tabRdu.style.display = 'block';
                if (tabCompare) tabCompare.style.display = 'block';

                if (fileInputLabel) fileInputLabel.style.display = 'none';
                if (saveGoogleBtn) saveGoogleBtn.style.display = 'none';
                if (saveAllGoogleBtn) saveAllGoogleBtn.style.display = 'none';

                const activeTab = document.querySelector('.tabs .tab-btn.active');
                if (activeTab && (activeTab === tabRaw || activeTab === tabDetailed || activeTab === tabEvents || activeTab === tabFact)) {
                    if (tabKamag) tabKamag.click();
                }
                enforceUserYardLock(userYard);
            } else {
                // Админ видит абсолютно всё
                if (tabRaw) tabRaw.style.display = 'block';
                if (tabDetailed) tabDetailed.style.display = 'block';
                if (tabEvents) tabEvents.style.display = 'block';
                if (tabKamag) tabKamag.style.display = 'block';
                if (tabFact) tabFact.style.display = 'block';
                if (tabRdu) tabRdu.style.display = (role === 'Адмін') ? 'block' : 'none';
                if (tabCompare) tabCompare.style.display = 'block';

                if (fileInputLabel) fileInputLabel.style.display = 'flex';
                if (saveGoogleBtn) saveGoogleBtn.style.display = 'inline-flex';
                if (saveAllGoogleBtn) saveAllGoogleBtn.style.display = 'inline-flex';

                ['kamagYardSelect', 'factYardSelect', 'compareYardSelect'].forEach(id => {
                    const select = document.getElementById(id);
                    if (select) select.disabled = false;
                });
            }
        }
    } else {
        // === ПРАВА ДЛЯ НЕАВТОРИЗОВАННЫХ ПОЛЬЗОВАТЕЛЕЙ (ГОСТЬ) ===
        // Гость видит только вкладку "Розрахунок" (активна по умолчанию)
        if (fileInputLabel) fileInputLabel.style.display = 'none';
        if (saveGoogleBtn) saveGoogleBtn.style.display = 'none';
        if (saveAllGoogleBtn) saveAllGoogleBtn.style.display = 'none';
    }

    // Обновляем плашку профиля в правом углу
    if (isAuth && loginBtn) {
        loginBtn.innerHTML = `${activeUser} (${role})`; 
        loginBtn.title = "Вийти з акаунту";
        
        loginBtn.removeEventListener('click', checkAuth);
        loginBtn.onclick = () => {
            if (confirm("Вийти з акаунту?")) {
                sessionStorage.clear();
                location.reload();
            }
        };
    }
}

document.getElementById('loginBtn').addEventListener('click', checkAuth);
document.addEventListener('DOMContentLoaded', updateAuthVisibility);

// Константа с индексами (числа со скрина МИНУС 1)
const colIdx = {
    dateStart: 1, dateEnd: 2, 
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

function parseExcelDate(val) {
    if (!val) return null;
    if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
}

function getDayOfWeekFromDotStr(dateStr) {
    if (!dateStr || !dateStr.includes('.')) return "";
    const [dd, mm, yyyy] = dateStr.split('.');
    const d = new Date(yyyy, mm - 1, dd);
    const dayNamesShort = ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
    return isNaN(d.getTime()) ? "" : dayNamesShort[d.getDay()];
}

function getFilteredDays(yard) {
    const startStr = document.getElementById('kamagStartDate').value;
    const endStr = document.getElementById('kamagEndDate').value;
    const startDate = startStr ? new Date(startStr).setHours(0,0,0,0) : null;
    const endDate = endStr ? new Date(endStr).setHours(23,59,59,999) : null;

    return Object.keys(totalOpsData[yard] || {}).sort((a, b) => {
        const [d1, m1, y1] = a.split('.');
        const [d2, m2, y2] = b.split('.');
        return new Date(y1, m1-1, d1) - new Date(y2, m2-1, d2);
    }).filter(d => {
        const [dd, mm, yyyy] = d.split('.');
        const currentD = new Date(yyyy, mm-1, dd).getTime();
        if (startDate && currentD < startDate) return false;
        if (endDate && currentD > endDate) return false;
        return true;
    });
}

const fileInput = document.getElementById('fileInput');
const DICT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxT4cGlFO8YcDzdeLaqSpThqgYbTbmhDoT8LSaB4FDNsLy0cGgsCa_V-zMINs3WhpcIEA/exec';
const RESULTS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzvbyu5rzhhFiezY6_rNN9-51XZ2h0UBFx0RDMxnGif_XRz_LtU7gWOJ28_RDT4STD3vQ/exec';

fileInput.disabled = true; 

async function loadRouteSchemas() {
    const label = document.getElementById('fileInputLabel');
    label.classList.add('disabled');
    fileInput.disabled = true;
    document.getElementById('fileStatus').innerText = "Завантаження даних...";

    try {
        const response = await fetch(DICT_SCRIPT_URL);
        const data = await response.json();
        routeDictionary = data.routes;
        yardDictionary = data.yards;
        fleetDictionary = data.fleet || {}; 
        usersDictionary = data.users || {}; 
        console.log("Довідники завантажені");
        
        await loadSavedYardsList();
        
    } catch (e) {
        console.error("Ошибка справочников:", e);
        document.getElementById('fileStatus').innerText = "Помилка завантаження довідників!";
    } finally {
        label.classList.remove('disabled');
        fileInput.disabled = false;
        
        if (document.getElementById('fileStatus').innerText === "Завантаження даних...") {
            document.getElementById('fileStatus').innerText = "Готово.";
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
            document.getElementById('fileStatus').innerText = "Готово.";

            // ЗАЩИТА РДУ ПОСЛЕ ЗАГРУЗКИ СПИСКА ИЗ БАЗЫ
            const isAuth = sessionStorage.getItem('kamagonAuth') === 'true';
            const role = sessionStorage.getItem('kamagonAuthRole');
            const userYard = sessionStorage.getItem('kamagonAuthYard');
            if (isAuth && role === 'РДУ') {
                enforceUserYardLock(userYard);
            }
        }
    } catch (e) {
        console.error("Помилка завантаження списку автодворів:", e);
    }
}

// Загрузка данных двора
document.getElementById('loadGoogleYardBtn').addEventListener('click', async () => {
    const yard = document.getElementById('kamagYardSelect').value;
    const btn = document.getElementById('loadGoogleYardBtn');
    btn.innerText = "⏳...";

    try {
        const response = await fetch(`${RESULTS_SCRIPT_URL}?action=getAggregatedData&yard=${encodeURIComponent(yard)}`);
        const data = await response.json();

        if (data.savedRows && data.savedRows.length > 0) {
            yardEvents = []; 
            fleetActiveState[yard] = {};
            totalOpsData[yard] = {};

            const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;
            const availM = fleetDictionary[yard] ? fleetDictionary[yard].man : 0;

            let maxLoadedK = availK;
            let maxLoadedM = availM;
            data.savedRows.forEach(row => {
                let countStr = String(row[3]);
                let separator = countStr.includes('|') ? '|' : (countStr.includes('.') ? '.' : ',');
                const counts = countStr.split(separator);
                
                let kLen = counts[0] && counts[0].includes(',') ? counts[0].split(',').length : (parseInt(counts[0], 10) || 0);
                let mLen = counts.length > 1 && counts[1] && counts[1].includes(',') ? counts[1].split(',').length : (parseInt(counts[1], 10) || 0);

                maxLoadedK = Math.max(maxLoadedK, kLen);
                maxLoadedM = Math.max(maxLoadedM, mLen);
            });

            data.savedRows.forEach(row => {
                let [y, day, hour, fleetCountStr, ops] = row;
                
                let dayStr = String(day);
                if (dayStr.includes('T') && dayStr.includes('Z')) {
                    const d = new Date(dayStr);
                    dayStr = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
                }

                if (!fleetActiveState[yard][dayStr]) {
                    fleetActiveState[yard][dayStr] = Array(24).fill(null).map(() => ({ 
                        kamag: Array(maxLoadedK).fill(0), 
                        man: Array(maxLoadedM).fill(0) 
                    }));
                    totalOpsData[yard][dayStr] = Array(24).fill(0);
                }

                totalOpsData[yard][dayStr][hour] = parseInt(ops, 10) || 0;
                
                let countStr = String(fleetCountStr);
                let separator = countStr.includes('|') ? '|' : (countStr.includes('.') ? '.' : ',');
                const counts = countStr.split(separator);
                
                const parseStateArr = (str, fallbackCount) => {
                    if (str && str.includes(',')) return str.split(',').map(Number); 
                    const c = parseInt(str, 10) || 0; 
                    return Array(Math.max(c, fallbackCount)).fill(0).map((_, i) => i < c ? 1 : 0);
                };

                const stateK = parseStateArr(counts[0], maxLoadedK);
                const stateM = parseStateArr(counts[1], maxLoadedM);
                
                for(let k = 0; k < stateK.length; k++) {
                    if (k < fleetActiveState[yard][dayStr][hour].kamag.length) {
                        fleetActiveState[yard][dayStr][hour].kamag[k] = stateK[k];
                    }
                }
                for(let m = 0; m < stateM.length; m++) {
                    if (m < fleetActiveState[yard][dayStr][hour].man.length) {
                        fleetActiveState[yard][dayStr][hour].man[m] = stateM[m];
                    }
                }
            });

            systemFleetState[yard] = {};
            for (let day in fleetActiveState[yard]) {
                systemFleetState[yard][day] = fleetActiveState[yard][day].map(hourObj => ({
                    kamag: [...hourObj.kamag],
                    man: [...hourObj.man]
                }));
            }

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
        this.dateStart = parseExcelDate(row[colIdx.dateStart]); 
        this.dateEnd = parseExcelDate(row[colIdx.dateEnd]);     
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
    html += `<th class="col-time">Дедлайн<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    
    ['Пн','Вт','Ср','Чт','Пт','Сб','Нд'].forEach(d => html += `<th class="col-day">${d}<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`);
    
    html += `<th class="col-point">Початкова<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-time">Приїзд<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-time">Виїзд<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;

    for(let i=1; i<=10; i++) {
        html += `<th class="pt-col col-point">П.Т. №${i}<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
        html += `<th class="pt-col col-time">Приїзд<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
        html += `<th class="pt-col col-time">Виїзд<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    }

    html += `<th class="col-point">Кінцева<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-time">Приїзд<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-time">Вивільнення<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-meta">Тип доставки<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-meta">Тип ТЗ<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-meta">Схема БДФ<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-meta">Формат<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-meta">Код<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;
    html += `<th class="col-meta">Тип переміщення<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>`;

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
        html += `<td class="col-route" title="${item.route}">${item.route}</td>`;
        html += `<td class="col-time">${item.deadline}</td>`;
        item.days.forEach(d => html += `<td class="col-day ${d ? 'day-on' : 'day-off'}">${d ? '1' : '0'}</td>`);
        html += `<td class="col-point" title="${item.pointNames[0]}">${item.pointNames[0]}</td><td class="col-time">${item.allTimes[0] || ""}</td><td class="col-time">${item.allTimes[1] || ""}</td>`;
        for(let j=1; j <= 10; j++) {
            html += `<td class="pt-col col-point" title="${item.pointNames[j]}">${item.pointNames[j]}</td><td class="pt-col col-time">${item.allTimes[j*2] || ""}</td><td class="pt-col col-time">${item.allTimes[j*2 + 1] || ""}</td>`;
        }
        html += `<td class="col-point" title="${item.pointNames[11]}">${item.pointNames[11]}</td><td class="col-time">${item.allTimes[22] || ""}</td><td class="col-time">${item.allTimes[23] || ""}</td>`;
        html += `<td class="col-meta" title="${item.deliveryType}">${item.deliveryType}</td><td class="col-meta" title="${item.vehicleType}">${item.vehicleType}</td><td class="col-meta" title="${item.schema}"><strong>${item.schema}</strong></td><td class="col-meta" title="${item.loadFormat}">${item.loadFormat}</td><td class="col-meta" title="${item.code}">${item.code}</td><td class="col-meta" title="${item.moveType}">${item.moveType}</td>`;
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

/*document.getElementById('togglePtBtn').addEventListener('click', function() {
    const container = document.getElementById('tableContainerRaw');
    if (container) {
        container.classList.toggle('hide-pt');
        this.classList.toggle('btn-active');
    }
});*/

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
const tabFact = document.getElementById('tabFact'); 
const tabRdu = document.getElementById('tabRdu'); // Вкладка РДУ

const containerRaw = document.getElementById('tableContainerRaw');
const containerDetailed = document.getElementById('tableContainerDetailed');
const containerEvents = document.getElementById('tableContainerEvents');
const containerKamag = document.getElementById('tableContainerKamag'); 
const containerFact = document.getElementById('tableContainerFact'); 
const containerRdu = document.getElementById('tableContainerRdu'); // Контейнер РДУ

function switchTab(activeTabBtn, activeContainer) {
    const tabCompare = document.getElementById('tabCompare');
    const containerCompare = document.getElementById('tableContainerCompare');
    [tabRaw, tabDetailed, tabEvents, tabKamag, tabFact, tabRdu, tabCompare].forEach(btn => {
        if (btn) btn.classList.remove('active');
    });
    [containerRaw, containerDetailed, containerEvents, containerKamag, containerFact, containerRdu, containerCompare].forEach(cont => {
        if (cont) cont.style.display = 'none';
    });
    
    activeTabBtn.classList.add('active');
    
    if (activeContainer === containerKamag || activeContainer === containerFact || activeContainer === containerRdu) {
        activeContainer.style.display = 'flex';
    } else {
        activeContainer.style.display = 'block';
    }

    const exportModeSelect = document.getElementById('exportModeSelect');
    if (exportModeSelect) {
        // Показуємо вибір режимів на вкладках Розрахунок, Факт та Результати
        if (activeTabBtn === tabKamag || activeTabBtn === document.getElementById('tabFact') || activeTabBtn === tabCompare) {
            exportModeSelect.classList.remove('hidden');
        } else {
            exportModeSelect.classList.add('hidden');
        }
    }
}

tabRaw.addEventListener('click', () => switchTab(tabRaw, containerRaw));
tabDetailed.addEventListener('click', () => switchTab(tabDetailed, containerDetailed));
tabEvents.addEventListener('click', () => switchTab(tabEvents, containerEvents));
tabKamag.addEventListener('click', () => {
    switchTab(tabKamag, containerKamag);
    if (Object.keys(totalOpsData).length > 0) {
        renderKamagTable();
    }
});

// Генерация
const generateModal = document.getElementById('generateModal');
let generatedDatesList = [];

document.getElementById('generateDetailedBtn').addEventListener('click', () => {
    if (allSchedules.length === 0) return alert("Спочатку завантажте вихідні графіки!");

    const yardSelect = document.getElementById('genYardSelect');
    yardSelect.innerHTML = '<option value="ALL">Всі автодвори (повний розрахунок)</option>';
    
    const uniqueYards = Object.keys(yardDictionary).map(k => yardDictionary[k].yard).filter((v, i, a) => v && a.indexOf(v) === i).sort();
    uniqueYards.forEach(y => {
        const opt = document.createElement('option');
        opt.value = opt.textContent = y;
        yardSelect.appendChild(opt);
    });

    const today = new Date();
    document.getElementById('genDateStart').value = today.toISOString().split('T')[0];
    const nextWeek = new Date(today);
    nextWeek.setDate(today.getDate() + 6);
    document.getElementById('genDateEnd').value = nextWeek.toISOString().split('T')[0];

    generateModal.style.display = 'block';
});

document.getElementById('closeGenerateModal').addEventListener('click', () => { generateModal.style.display = 'none'; });

document.getElementById('confirmGenerateBtn').addEventListener('click', () => {
    const yardOpt = document.getElementById('genYardSelect').value;
    const dStart = new Date(document.getElementById('genDateStart').value);
    const dEnd = new Date(document.getElementById('genDateEnd').value);
    const useDates = document.getElementById('genUseScheduleDates').checked;

    if (isNaN(dStart) || isNaN(dEnd) || dStart > dEnd) return alert("Некоректний діапазон дат!");

    generateModal.style.display = 'none';

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

        generatedDatesList.forEach(targetDate => {
            if (useDates) {
                const tDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate()).getTime();
                const sDate = item.dateStart ? new Date(item.dateStart.getFullYear(), item.dateStart.getMonth(), item.dateStart.getDate()).getTime() : 0;
                const eDate = item.dateEnd ? new Date(item.dateEnd.getFullYear(), item.dateEnd.getMonth(), item.dateEnd.getDate()).getTime() : Infinity;
                
                if (tDate < sDate) return; 
                if (item.dateEnd && tDate > eDate) return; 
            }

            const dayOfWeekIdx = (targetDate.getDay() + 6) % 7; 
            if (!item.days[dayOfWeekIdx]) return; 

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
                
                if (targetYard !== "ALL" && (!yardDataA || yardDataA.yard !== targetYard) && (!yardDataB || yardDataB.yard !== targetYard)) return;

                let arrMins = getAbsoluteMinutes(dateString, nodeB.timeArr);
                let finalArrivalB = formatAbsoluteMinutes(arrMins);

                detailedSchedules.push({
                    originalRoute: item.route,
                    originalCode: item.code, 
                    day: dateString, 
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
        <th data-sort="originalRoute" class="sortable">Маршрут<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th>Код<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th data-sort="day" class="sortable col-day">День<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
        <th>День тижня<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
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
            <td class="col-route">${item.originalRoute}</td>
            <td>${item.originalCode}</td>
            <td class="col-day" style="font-weight: bold;">${item.day}</td>
            <td style="text-align: center; font-weight: bold; color: #495057;">${getDayOfWeekFromDotStr(item.day)}</td>
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

    const [dd, mm, yyyy] = targetDateStr.split('.');
    const [hh, min] = targetTimeStr.split(':').map(Number);
    
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

function calculateRampTimes() {
    const interval = parseInt(document.getElementById('rampInterval').value, 10) || 10;
    
    const groups = {};
    detailedSchedules.forEach(item => {
        if (item.timeDepartureA === "—") return;
        
        // Групуємо за вузлом та маршрутом
        const key = `${item.nodeA}_${item.originalRoute}`;
        
        if (!groups[key]) groups[key] = [];
        
        item.absDep = getAbsoluteMinutes(item.day, item.timeDepartureA);
        groups[key].push(item);
    });
    
    for (const key in groups) {
        const group = groups[key];
        
        // Сортуємо хронологічно
        group.sort((a, b) => a.absDep - b.absDep);
        
        if (group.length === 0) continue;
        
        const yardConf = yardDictionary[group[0].nodeA];
        
        let firstAbs = Infinity;
        if (yardConf && yardConf.firstPlacement !== undefined) {
            let fp = String(yardConf.firstPlacement).trim();
            if (fp !== "" && fp !== "—" && fp !== "-") {
                if (!fp.includes(':')) fp += ":00"; 
                firstAbs = getAbsoluteMinutes(group[0].day, fp);
            }
        }
        
        if (isNaN(firstAbs) || firstAbs === Infinity) {
            firstAbs = getAbsoluteMinutes(group[0].day, "00:00");
        }
        if (isNaN(firstAbs) || firstAbs === Infinity) {
            firstAbs = group[0].absDep - 120;
        }

        let prevAbsDep = 0; 
        let isFirstBatch = true;
        
        let i = 0;
        while (i < group.length) {
            let currentAbsDep = group[i].absDep;
            
            // Збираємо ВУЗОЛ: всі контейнери, що виїжджають в одну і ту ж хвилину (двійники, трійники)
            let batch = [];
            while (i < group.length && group[i].absDep === currentAbsDep) {
                batch.push(group[i]);
                i++;
            }

            let proposedPlacement;
            if (isFirstBatch) {
                proposedPlacement = firstAbs;
                isFirstBatch = false;
            } else {
                proposedPlacement = prevAbsDep + interval;
            }

            let maxPlacementTime = currentAbsDep - (23 * 60); 
            let finalPlacement = Math.max(proposedPlacement, maxPlacementTime);
            
            // Застосовуємо однаковий час постановки і ЗБЕРІГАЄМО виїзд для всіх контейнерів у партії
            batch.forEach(item => {
                item.timePlacementA = formatAbsoluteMinutes(finalPlacement);
                item.timeDepartureA = formatAbsoluteMinutes(item.absDep);
            });

            prevAbsDep = currentAbsDep;
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

    filteredDetailedSchedules = filterDataArray('tableContainerDetailed', detailedSchedules, getDetailedValues);
    detailedRenderedCount = 0;
    document.getElementById('detailedTableBody').innerHTML = "";
    renderDetailedChunk();
    document.getElementById('tableContainerDetailed').scrollTop = 0; 
}

function generateYardEvents() {
    yardEvents = [];
    
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
        <th>День тижня<br><input type="text" class="filter-input" data-col="${c++}" onclick="event.stopPropagation()"></th>
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
            <td style="text-align: center; font-weight: bold; color: #495057;">${getDayOfWeekFromDotStr(ev.day)}</td>
            <td style="text-align: center;">${ev.time}</td>
        </tr>`;
    }
    tbody.insertAdjacentHTML('beforeend', html);
    eventsRenderedCount = end;
}

function calculateFleetRequirements() {
    if (yardEvents && yardEvents.length > 0) {
        totalOpsData = {}; 
        yardEvents.forEach(ev => {
            if (!totalOpsData[ev.yard]) totalOpsData[ev.yard] = {};
            if (!totalOpsData[ev.yard][ev.day]) totalOpsData[ev.yard][ev.day] = Array(24).fill(0);
            const hour = parseInt(ev.time.split(':')[0], 10);
            if (!isNaN(hour)) totalOpsData[ev.yard][ev.day][hour]++;
        });
    }

    if (Object.keys(totalOpsData).length === 0) return;

    fleetActiveState = {}; 

    const yardNorms = {};
    for(let node in yardDictionary) {
        let y = yardDictionary[node].yard;
        if(!yardNorms[y]) yardNorms[y] = { k: yardDictionary[node].normKamag || 12, m: yardDictionary[node].normMan || 6 };
    }

    for (let yard in totalOpsData) {
        fleetActiveState[yard] = {};
        const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;
        const availM = fleetDictionary[yard] ? fleetDictionary[yard].man : 0;
        const fontK = yardNorms[yard] ? yardNorms[yard].k : 12;
        const fontM = yardNorms[yard] ? yardNorms[yard].m : 6;

        const virtualFleetRadio = document.querySelector('input[name="virtualFleetType"]:checked');
        let yardVirtualType = virtualFleetRadio ? virtualFleetRadio.value : (availK >= availM ? 'kamag' : 'man');
        let maxExtraK = 0;
        let maxExtraM = 0;

        for (let day in totalOpsData[yard]) {
            for (let h = 0; h < 24; h++) {
                let ops = totalOpsData[yard][day][h];
                let cap = (availK * fontK) + (availM * fontM);
                if (ops > cap) {
                    if (yardVirtualType === 'kamag') {
                        let extra = Math.ceil((ops - cap) / fontK);
                        if (extra > maxExtraK) maxExtraK = extra;
                    } else {
                        let extra = Math.ceil((ops - cap) / fontM);
                        if (extra > maxExtraM) maxExtraM = extra;
                    }
                }
            }
        }

        const totalK = availK + maxExtraK;
        const totalM = availM + maxExtraM;

        for (let day in totalOpsData[yard]) {
            fleetActiveState[yard][day] = Array(24).fill(null).map(() => ({ 
                kamag: Array(totalK).fill(0), 
                man: Array(totalM).fill(0) 
            }));
            
            for (let h = 0; h < 24; h++) {
                let neededOps = totalOpsData[yard][day][h];
                let assignedK = 0;
                while (neededOps > 0 && assignedK < availK) {
                    fleetActiveState[yard][day][h].kamag[assignedK] = 1;
                    assignedK++;
                    neededOps -= fontK;
                }
                
                let assignedM = 0;
                while (neededOps > 0 && assignedM < availM) {
                    fleetActiveState[yard][day][h].man[assignedM] = 1;
                    assignedM++;
                    neededOps -= fontM;
                }

                if (yardVirtualType === 'kamag') {
                    while (neededOps > 0 && assignedK < totalK) {
                        fleetActiveState[yard][day][h].kamag[assignedK] = 1;
                        assignedK++;
                        neededOps -= fontK;
                    }
                } else {
                    while (neededOps > 0 && assignedM < totalM) {
                        fleetActiveState[yard][day][h].man[assignedM] = 1;
                        assignedM++;
                        neededOps -= fontM;
                    }
                }
            }
        }
    }

    for (let y in totalOpsData) {
        systemFleetState[y] = {};
        for (let day in totalOpsData[y]) {
            if (fleetActiveState[y] && fleetActiveState[y][day]) {
                systemFleetState[y][day] = fleetActiveState[y][day].map(hourObj => ({
                    kamag: [...hourObj.kamag],
                    man: [...hourObj.man]
                }));
            }
        }
    }

    const yardSelect = document.getElementById('kamagYardSelect');
    const currentVal = yardSelect.value;
    yardSelect.innerHTML = "";
    
    // Берем все уникальные автодворы из глобального справочника, а не только из текущего расчета
    const uniqueYards = Object.keys(yardDictionary).map(k => yardDictionary[k].yard).filter((v, i, a) => v && a.indexOf(v) === i).sort();
    
    uniqueYards.forEach(yard => {
        const option = document.createElement('option');
        option.value = option.textContent = yard;
        yardSelect.appendChild(option);
    });
    
    if (currentVal && uniqueYards.includes(currentVal)) {
        yardSelect.value = currentVal;
    }

    // СТРАХОВОЧНЫЙ СЛОТ КУПИРОВАНИЯ ДЛЯ РДУ ПОСЛЕ ПЕРЕСЧЕТА МАТРИЦЫ
    const isAuth = sessionStorage.getItem('kamagonAuth') === 'true';
    const role = sessionStorage.getItem('kamagonAuthRole');
    const userYard = sessionStorage.getItem('kamagonAuthYard');
    if (isAuth && role === 'РДУ') {
        enforceUserYardLock(userYard);
    }

    renderKamagTable();
}

function renderKamagTable() {
    const yard = document.getElementById('kamagYardSelect').value;
    const wrapper = document.getElementById('kamagTableWrapper');

    if (!yard || !totalOpsData[yard]) {
        wrapper.innerHTML = "<p style='padding:20px;'>Немає даних для цього автодвору.</p>";
        return;
    }

    const savedScrollLeft = wrapper ? wrapper.scrollLeft : 0;
    const savedScrollTop = wrapper ? wrapper.scrollTop : 0;

    const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;
    const availM = fleetDictionary[yard] ? fleetDictionary[yard].man : 0;
    
    let totalK = availK;
    let totalM = availM;
    const availableDates = Object.keys(fleetActiveState[yard] || {});
    if (availableDates.length > 0 && fleetActiveState[yard][availableDates[0]][0]) {
        totalK = fleetActiveState[yard][availableDates[0]][0].kamag.length;
        totalM = fleetActiveState[yard][availableDates[0]][0].man.length;
    }

    const yardNorms = { k: 12, m: 6 };
    for(let node in yardDictionary) {
        if(yardDictionary[node].yard === yard) {
            yardNorms.k = yardDictionary[node].normKamag || 12;
            yardNorms.m = yardDictionary[node].normMan || 6;
            break;
        }
    }

    const daysOfWeek = getFilteredDays(yard);

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
        
        let isFleet = title.includes("Флот") || title.includes("Машини");
        let unit = isFleet ? "год" : "шт"; 
        html += `<th style="text-align: center; line-height: 1.2;">Всього,<br>${unit}</th></tr><tr><th style="font-size: 10px;">${isFleet ? 'ТЗ / Години' : 'Операції / Години'}</th>`;
        
        daysOfWeek.forEach(d => {
            for (let i = 0; i < 24; i++) {
                html += `<th class="kamag-header-vertical" style="${i === 0 ? 'border-left: 2px solid #6c757d;' : ''}">${i}:00</th>`;
            }
            html += `<th style="text-align: center; font-weight: bold; font-size: 10px; background-color: #dee2e6; border-right: 2px solid #6c757d; min-width: 40px;">Σ</th>`;
        });
        html += `<th></th></tr></thead><tbody>`;

        rowHeaders.forEach(rowName => {
            let isVehicleRow = rowName.startsWith("Kamag") || rowName.startsWith("Маневровий");
            let rowClass = "";
            if (rowName === "Задіяно фіз. КАМАГ") rowClass = ' class="fleet-summary-top"';
            else if (rowName === "Задіяно вірт. МАН") rowClass = ' class="fleet-summary-bottom"';
            else if (rowName.startsWith("Задіяно")) rowClass = ' class="fleet-summary-row"';

            if (isFleet && isVehicleRow) {
                let isKamag = rowName.startsWith("Kamag");
                let dataType = isKamag ? "kamag" : "man";
                let match = rowName.match(/\d+/);
                let idx = match ? parseInt(match[0]) - 1 : 0;
                html += `<tr${rowClass}><td class="fleet-row-header" data-yard="${yard}" data-type="${dataType}" data-index="${idx}" style="font-weight: bold; font-size: 11px; cursor: pointer;">${rowName}</td>`;
            } else {
                html += `<tr${rowClass}><td style="font-weight: bold; font-size: 11px;">${rowName}</td>`;
            }

            let totalRowSum = 0;
            daysOfWeek.forEach(d => {
                let dailySum = 0;
                for (let h = 0; h < 24; h++) {
                    let val = dataProvider(rowName, d, h);
                    let borderStyle = h === 0 ? "border-left: 2px solid #6c757d;" : "";
                    
                    if (isFleet && isVehicleRow) {
                        let isKamag = rowName.startsWith("Kamag");
                        let dataType = isKamag ? "kamag" : "man";
                        let match = rowName.match(/\d+/);
                        let idx = match ? parseInt(match[0]) - 1 : 0;

                        let isActive = val > 0;
                        let isManual = val === 2;
                        let cellClass = "kamag-cell kamag-editable";

                        if (isActive) {
                            if (isManual) {
                                if (isKamag && idx >= availK) cellClass += " kamag-manual-virtual";
                                else if (!isKamag && idx >= availM) cellClass += " kamag-manual-virtual";
                                else cellClass += " kamag-manual-physical";
                            } else {
                                if (isKamag && idx >= availK) cellClass += " kamag-active-virtual";
                                else if (!isKamag && idx >= availM) cellClass += " kamag-active-virtual";
                                else cellClass += " kamag-active";
                            }
                        }
                        html += `<td class="${cellClass}" style="${borderStyle}" data-yard="${yard}" data-day="${d}" data-hour="${h}" data-type="${dataType}" data-index="${idx}">${isActive ? 1 : ''}</td>`;
                        if (isActive) { dailySum++; totalRowSum++; }
                    } else {
                        if (val !== 0 && val !== "") {
                            html += `<td class="kamag-cell" style="${borderStyle} background-color: #fff9c4; font-weight: bold;">${val}</td>`;
                            let numVal = typeof val === 'number' ? val : (parseInt(String(val).replace(/<[^>]*>/g, ''), 10) || 0);
                            dailySum += numVal; totalRowSum += numVal;
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
                html += `<td colspan="24" style="border-left: 2px solid #6c757d; vertical-align: bottom; padding: 0; background: #fff;"><div style="height: 135px; width: 100%;"><canvas id="chart_${index}"></canvas></div></td><td style="border-right: 2px solid #6c757d; background-color: #dee2e6;"></td>`;
            });
            html += `<td style="background-color: #e9ecef;"></td></tr>`;
        }
        html += `</tbody></table>`;
        return html;
    }

    const hideVirtual = document.getElementById('hideVirtualFleet').checked;
    const rowHeaders = [];
    for(let i=1; i<=availK; i++) rowHeaders.push(`Kamag ${i}`);
    for(let i=1; i<=availM; i++) rowHeaders.push(`Маневровий ${i}`);

    if (!hideVirtual) {
        for(let i=availK+1; i<=totalK; i++) rowHeaders.push(`Kamag ${i} (дод.)`);
        for(let i=availM+1; i<=totalM; i++) rowHeaders.push(`Маневровий ${i} (дод.)`);
    }

    rowHeaders.push("Задіяно фіз. КАМАГ", "Задіяно вірт. КАМАГ", "Задіяно фіз. МАН", "Задіяно вірт. МАН");
    
    let fleetHTML = generateMatrixHTML(`Флот`, rowHeaders, (row, day, hour) => {
        if (!fleetActiveState[yard][day] || !fleetActiveState[yard][day][hour]) return 0;
        if (row.startsWith("Задіяно")) {
            const st = fleetActiveState[yard][day][hour];
            if (row === "Задіяно фіз. КАМАГ") return st.kamag.slice(0, availK).filter(Boolean).length || "";
            if (row === "Задіяно вірт. КАМАГ") return st.kamag.slice(availK).filter(Boolean).length || "";
            if (row === "Задіяно фіз. МАН") return st.man.slice(0, availM).filter(Boolean).length || "";
            if (row === "Задіяно вірт. МАН") return st.man.slice(availM).filter(Boolean).length || "";
        }
        const isKamag = row.startsWith("Kamag");
        const match = row.match(/\d+/);
        const idx = match ? parseInt(match[0]) - 1 : 0;
        return isKamag ? fleetActiveState[yard][day][hour].kamag[idx] : fleetActiveState[yard][day][hour].man[idx];
    });

    let opsHTML = generateMatrixHTML(`Операції`, ["Всього операцій", "Непокриті (фіз. флот)", "Непокриті (залишок)"], (row, day, hour) => {
        const totalOps = (totalOpsData[yard] && totalOpsData[yard][day]) ? totalOpsData[yard][day][hour] : 0;
        if (row === "Всього операцій") return totalOps;

        let activePhysK = 0, activeVirtK = 0, activePhysM = 0, activeVirtM = 0;
        if (fleetActiveState[yard] && fleetActiveState[yard][day] && fleetActiveState[yard][day][hour]) {
            const st = fleetActiveState[yard][day][hour];
            activePhysK = st.kamag.slice(0, availK).filter(Boolean).length;
            activeVirtK = st.kamag.slice(availK).filter(Boolean).length;
            activePhysM = st.man.slice(0, availM).filter(Boolean).length;
            activeVirtM = st.man.slice(availM).filter(Boolean).length;
        }

        if (row === "Непокриті (фіз. флот)") {
            const capPhysical = (activePhysK * yardNorms.k) + (activePhysM * yardNorms.m);
            const uncoveredPhys = Math.max(0, totalOps - capPhysical);
            return uncoveredPhys > 0 ? `<span class="uncovered-alert">${uncoveredPhys}</span>` : '';
        } else {
            const activeTotalK = hideVirtual ? activePhysK : (activePhysK + activeVirtK);
            const activeTotalM = hideVirtual ? activePhysM : (activePhysM + activeVirtM);
            const capTotal = (activeTotalK * yardNorms.k) + (activeTotalM * yardNorms.m);
            const uncoveredAbs = Math.max(0, totalOps - capTotal);
            return uncoveredAbs > 0 ? `<span class="uncovered-alert">${uncoveredAbs}</span>` : '';
        }
    }, true);

    wrapper.innerHTML = fleetHTML + "<div style='height: 10px;'></div>" + opsHTML;
    wrapper.scrollLeft = savedScrollLeft;
    wrapper.scrollTop = savedScrollTop;

    if (window.myDayCharts) window.myDayCharts.forEach(c => c.destroy());
    window.myDayCharts = [];

    daysOfWeek.forEach((d, index) => {
        const ctx = document.getElementById(`chart_${index}`);
        if (!ctx) return;
        const parentDiv = ctx.parentElement;
        ctx.width = parentDiv.clientWidth; ctx.height = 135; 

        const chartLabels = [], opsData = [], capacityData = [];
        for (let h = 0; h < 24; h++) {
            chartLabels.push(`${h}:00`);
            opsData.push((totalOpsData[yard] && totalOpsData[yard][d]) ? totalOpsData[yard][d][h] : 0);
            let cap = 0;
            if (fleetActiveState[yard] && fleetActiveState[yard][d] && fleetActiveState[yard][d][h]) {
                const st = fleetActiveState[yard][d][h];
                const activeK = hideVirtual ? st.kamag.slice(0, availK).filter(Boolean).length : st.kamag.filter(Boolean).length;
                const activeM = hideVirtual ? st.man.slice(0, availM).filter(Boolean).length : st.man.filter(Boolean).length;
                cap += activeK * yardNorms.k + activeM * yardNorms.m;
            }
            capacityData.push(cap);
        }

        window.myDayCharts.push(new Chart(ctx, {
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
                    legend: { display: true, position: 'bottom', labels: { boxWidth: 12, font: { size: 10 }, padding: 4 } }, 
                    tooltip: { callbacks: { title: (items) => `${d} ${items[0].label}`, label: (item) => `${item.dataset.label}: ${item.raw}` } } 
                }, 
                layout: { padding: 0 } 
            }
        }));
    });
}

document.getElementById('kamagYardSelect').addEventListener('change', renderKamagTable);

document.getElementById('kamagTableWrapper').addEventListener('click', function(e) {
    if (e.target.classList.contains('fleet-row-header')) {
        if (sessionStorage.getItem('kamagonAuth') !== 'true') return;
        const header = e.target;
        const yard = header.getAttribute('data-yard');
        const type = header.getAttribute('data-type');
        const idx = parseInt(header.getAttribute('data-index'));

        if (!confirm(`Очистити всі одинички для "${header.innerText}" за обраний період відображення?`)) return;

        const days = getFilteredDays(yard);
        days.forEach(day => {
            for (let h = 0; h < 24; h++) {
                if (fleetActiveState[yard] && fleetActiveState[yard][day] && fleetActiveState[yard][day][h]) {
                    fleetActiveState[yard][day][h][type][idx] = 0;
                }
            }
        });
        renderKamagTable();
        return;
    }

    if (e.target.classList.contains('kamag-editable')) {
        if (sessionStorage.getItem('kamagonAuth') !== 'true') return;
        const cell = e.target;
        const yard = cell.getAttribute('data-yard');
        const day = cell.getAttribute('data-day');
        const hour = parseInt(cell.getAttribute('data-hour'));
        const type = cell.getAttribute('data-type'); 
        const idx = parseInt(cell.getAttribute('data-index'));

        const currentState = fleetActiveState[yard][day][hour][type][idx];
        let newState = 0;

        if (currentState > 0) {
            newState = 0; 
        } else {
            const wasSystemActive = systemFleetState[yard] && 
                                    systemFleetState[yard][day] && 
                                    systemFleetState[yard][day][hour] && 
                                    systemFleetState[yard][day][hour][type] && 
                                    systemFleetState[yard][day][hour][type][idx] === 1;
            newState = wasSystemActive ? 1 : 2; 
        }
        
        fleetActiveState[yard][day][hour][type][idx] = newState;
        renderKamagTable();
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
        const availM = fleetDictionary[ev.yard] ? fleetDictionary[ev.yard].man : 0;

        if (fleetActiveState[ev.yard] && fleetActiveState[ev.yard][ev.day] && fleetActiveState[ev.yard][ev.day][hour]) {
            const st = fleetActiveState[ev.yard][ev.day][hour];
            const activeResources = [];
            
            st.kamag.forEach((isActive, idx) => { 
                if(isActive) activeResources.push(idx < availK ? `Kamag ${idx+1}` : `Kamag ${idx+1} (дод.)`); 
            });
            st.man.forEach((isActive, idx) => { 
                if(isActive) activeResources.push(idx < availM ? `Маневровий ${idx+1}` : `Маневровий ${idx+1} (дод.)`); 
            });

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
        const availM = fleetDictionary[y] ? fleetDictionary[y].man : 0;
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
                        const name = mIndex < availM ? `Маневровий ${mIndex + 1}` : `Маневровий ${mIndex + 1} (дод.)`;
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

function getRawValues(item) {
    const vals = [item.route, item.deadline, ...(item.days.map(d => d ? "1" : "0")), item.pointNames[0], item.allTimes[0] || "", item.allTimes[1] || ""];
    for(let j=1; j<=10; j++) vals.push(item.pointNames[j], item.allTimes[j*2] || "", item.allTimes[j*2 + 1] || "");
    vals.push(item.pointNames[11], item.allTimes[22] || "", item.allTimes[23] || "", item.deliveryType, item.vehicleType, item.schema, item.loadFormat, item.code, item.moveType);
    return vals;
}

function getDetailedValues(item) {
    return [item.originalRoute, item.originalCode, item.day, getDayOfWeekFromDotStr(item.day), item.miniSchema, item.containerType, item.yardA, item.timePlacementA || "—", item.nodeA, item.timeDepartureA, item.yardB, item.nodeB, item.timeArrivalB, item.timeUnloadStart, item.timeUnloadEnd, item.vehicle];
}

function getEventsValues(ev) {
    return [ev.yard, ev.code, ev.event, ev.day, getDayOfWeekFromDotStr(ev.day), ev.time];
}

function filterDataArray(containerId, dataArray, valuesExtractor) {
    const inputs = document.querySelectorAll(`#${containerId} .filter-input`);
    const filters = [];
    inputs.forEach(input => {
        const val = input.value.trim().toLowerCase();
        if (val) filters.push({ col: parseInt(input.getAttribute('data-col')), val: val });
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
            if (!container) return; 
            if (container.id === 'tableContainerRaw') {
                filteredAllSchedules = filterDataArray('tableContainerRaw', allSchedules, getRawValues);
                renderedCount = 0; document.getElementById('tableBody').innerHTML = ""; renderChunk();
            } else if (container.id === 'tableContainerDetailed') {
                filteredDetailedSchedules = filterDataArray('tableContainerDetailed', detailedSchedules, getDetailedValues);
                detailedRenderedCount = 0; document.getElementById('detailedTableBody').innerHTML = ""; renderDetailedChunk();
            } else if (container.id === 'tableContainerEvents') {
                filteredYardEvents = filterDataArray('tableContainerEvents', yardEvents, getEventsValues);
                eventsRenderedCount = 0; document.getElementById('eventsTableBody').innerHTML = ""; renderEventsChunk();
            }
        }, 300); 
    }
});

// --- Повністю замінити обробник exportExcelBtn в самому кінці app.js ---
document.getElementById('exportExcelBtn').addEventListener('click', async () => {
    const btn = document.getElementById('exportExcelBtn');
    const originalText = btn.innerText;
    btn.disabled = true;

    try {
        const workbook = new ExcelJS.Workbook();
        
        if (tabCompare.classList.contains('active')) {
            if (typeof window.exportCompareToExcel === 'function') {
                const success = await window.exportCompareToExcel(workbook);
                if (!success) { btn.innerText = originalText; btn.disabled = false; return; }
            } else {
                alert("Модуль експорту звірки не знайдено!");
                btn.innerText = originalText; btn.disabled = false; return;
            }
        }
        else if (document.getElementById('tabFact').classList.contains('active')) {
            if (typeof window.exportFactToExcel === 'function') {
                const success = await window.exportFactToExcel(workbook);
                if (!success) { btn.innerText = originalText; btn.disabled = false; return; }
            } else {
                alert("Модуль експорту Факту не знайдено!");
                btn.innerText = originalText; btn.disabled = false; return;
            }
        }
        else if (tabRaw.classList.contains('active')) {
            const headers = ["Маршрут", "Дедлайн", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд", "Початкова", "Приїзд", "Виїзд"];
            for(let i=1; i<=10; i++) headers.push(`П.Т. №${i}`, "Приїзд", "Виїзд");
            headers.push("Кінцева", "Приїзд", "Вивільнення", "Тип доставки", "Тип ТЗ", "Схема БДФ", "Формат", "Код", "Тип переміщення");
            const sheet = workbook.addWorksheet('Звіт'); sheet.addRow(headers);
            filteredAllSchedules.forEach(item => sheet.addRow(getRawValues(item)));
        } else if (tabDetailed.classList.contains('active')) {
            const headers = ["Маршрут", "Код", "День", "День тижня", "Схема", "Тип", "Автодвір А", "Постановка", "Точка А", "Виїзд", "Автодвір Б", "Точка Б", "Приїзд", "Постановка (вивант.)", "Кінець вивант.", "Тип ТЗ"];
            const sheet = workbook.addWorksheet('Звіт'); sheet.addRow(headers);
            filteredDetailedSchedules.forEach((item) => sheet.addRow(getDetailedValues(item)));
        } else if (tabEvents.classList.contains('active')) {
            const headers = ["Автодвір", "Код", "Подія", "День", "День тижня", "Час"];
            const sheet = workbook.addWorksheet('Звіт'); sheet.addRow(headers);
            filteredYardEvents.forEach((ev) => sheet.addRow(getEventsValues(ev)));
        } else if (tabKamag.classList.contains('active')) {
            const startStr = document.getElementById('kamagStartDate').value;
            const endStr = document.getElementById('kamagEndDate').value;
            const startDate = startStr ? new Date(startStr).setHours(0,0,0,0) : null;
            const endDate = endStr ? new Date(endStr).setHours(23,59,59,999) : null;
            const hideVirtual = document.getElementById('hideVirtualFleet').checked;
            const exportMode = document.getElementById('exportModeSelect').value;
            
            let yardsToExport = [];
            if (exportMode === 'current') {
                const currentYard = document.getElementById('kamagYardSelect').value;
                if (currentYard) yardsToExport.push(currentYard);
            } else yardsToExport = Object.keys(totalOpsData).sort();

            if (yardsToExport.length === 0) {
                alert("Немає розрахованих даних для експорту!");
                btn.innerText = originalText; btn.disabled = false; return;
            }

            const alignCenter = { vertical: 'middle', horizontal: 'center' };
            const fillHeader = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } };
            const fillSum = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F3F5' } };
            const fillActive = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBDEFB' } }; 
            const fillActiveVirtual = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE0B2' } }; 
            const fillOps = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
            const fillUncovered = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEBEE' } }; 
            const fontUncovered = { bold: true, size: 9, color: { argb: 'FFD32F2F' } };
            const borderThin = { style: 'thin', color: { argb: 'FFCCCCCC' } };
            const borderMedium = { style: 'medium', color: { argb: 'FF6C757D' } };
            const getBorders = (isLeftEdge, isRightEdge) => ({ top: borderThin, bottom: borderThin, left: isLeftEdge ? borderMedium : borderThin, right: isRightEdge ? borderMedium : borderThin });

            yardsToExport.forEach(yard => {
                const sheetName = yard.substring(0, 31).replace(/[\\\?\*\[\]\/]/g, "");
                const sheet = workbook.addWorksheet(sheetName);

                const yardNorms = { k: 12, m: 6 };
                for(let node in yardDictionary) {
                    if(yardDictionary[node].yard === yard) {
                        yardNorms.k = yardDictionary[node].normKamag || 12;
                        yardNorms.m = yardDictionary[node].normMan || 6;
                        break;
                    }
                }

                const days = Object.keys(totalOpsData[yard] || {}).sort((a, b) => {
                    const [d1, m1, y1] = a.split('.'); const [d2, m2, y2] = b.split('.');
                    return new Date(y1, m1-1, d1) - new Date(y2, m2-1, d2);
                }).filter(d => {
                    const [dd, mm, yyyy] = d.split('.'); const currentD = new Date(yyyy, mm-1, dd).getTime();
                    if (startDate && currentD < startDate) return false;
                    if (endDate && currentD > endDate) return false;
                    return true;
                });

                if (days.length === 0) return; 
                sheet.getColumn(1).width = 18; 
                for (let i = 2; i <= 1 + 25 * days.length; i++) {
                    sheet.getColumn(i).width = 3.5; if ((i - 1) % 25 === 0) sheet.getColumn(i).width = 5; 
                }

                sheet.addRow([`Звіт по Флоту: ${yard}`]).font = { bold: true, size: 14 }; sheet.addRow([]);
                const rowDays = sheet.addRow(["День"]); rowDays.getCell(1).font = { bold: true }; rowDays.getCell(1).alignment = alignCenter;

                const dayNamesShort = ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
                days.forEach((d, index) => {
                    const startCol = 2 + index * 25; const endCol = startCol + 24; sheet.mergeCells(3, startCol, 3, endCol); 
                    const [dd, mm, yyyy] = d.split('.'); const dateObj = new Date(yyyy, mm - 1, dd);
                    const cell = sheet.getCell(3, startCol); cell.value = `${d} (${dayNamesShort[dateObj.getDay()]})`;
                    cell.alignment = alignCenter; cell.font = { bold: true }; cell.fill = fillHeader; cell.border = getBorders(true, true);
                });
                sheet.getCell(3, 2 + 25 * days.length).value = "Всього"; sheet.getCell(3, 2 + 25 * days.length).font = { bold: true };

                const rowHours = sheet.addRow(["Рядок / Години"]); rowHours.getCell(1).font = { size: 10 };
                let currentCol = 2;
                days.forEach((d, index) => {
                    for (let h = 0; h < 24; h++) {
                        const cell = rowHours.getCell(currentCol); cell.value = h; cell.alignment = alignCenter; cell.font = { size: 9 }; cell.border = getBorders(h === 0, false); currentCol++;
                    }
                    const sumCell = rowHours.getCell(currentCol); sumCell.value = "Σ"; sumCell.alignment = alignCenter; sumCell.font = { bold: true, size: 9 }; sumCell.fill = fillHeader; sumCell.border = getBorders(false, true); currentCol++;
                });

                const availK = fleetDictionary[yard] ? fleetDictionary[yard].kamag : 0;
                const availM = fleetDictionary[yard] ? fleetDictionary[yard].man : 0;
                let totalK = availK, totalM = availM;
                const availableDates = Object.keys(fleetActiveState[yard] || {});
                if (availableDates.length > 0 && fleetActiveState[yard][availableDates[0]][0]) {
                    totalK = fleetActiveState[yard][availableDates[0]][0].kamag.length; totalM = fleetActiveState[yard][availableDates[0]][0].man.length;
                }

                for (let k = 1; k <= availK; k++) {
                    const row = sheet.addRow([`Kamag ${k}`]); row.getCell(1).font = { bold: true, size: 10 };
                    let weekTotal = 0, cCol = 2;
                    days.forEach(d => {
                        let daySum = 0;
                        for (let h = 0; h < 24; h++) {
                            let val = (fleetActiveState[yard][d] && fleetActiveState[yard][d][h].kamag[k-1]) ? 1 : "";
                            const cell = row.getCell(cCol); cell.value = val; cell.alignment = alignCenter; cell.border = getBorders(h === 0, false);
                            if (val === 1) { cell.fill = fillActive; daySum++; weekTotal++; } cCol++;
                        }
                        const dSumCell = row.getCell(cCol); dSumCell.value = daySum || ""; dSumCell.alignment = alignCenter; dSumCell.font = { bold: true }; dSumCell.fill = fillSum; dSumCell.border = getBorders(false, true); cCol++;
                    });
                    const wSumCell = row.getCell(cCol); wSumCell.value = weekTotal; wSumCell.alignment = alignCenter; wSumCell.font = { bold: true }; wSumCell.fill = fillHeader;
                }

                for (let m = 1; m <= availM; m++) {
                    const row = sheet.addRow([`Маневровий ${m}`]); row.getCell(1).font = { bold: true, size: 10 };
                    let weekTotal = 0, cCol = 2;
                    days.forEach(d => {
                        let daySum = 0;
                        for (let h = 0; h < 24; h++) {
                            let val = (fleetActiveState[yard][d] && fleetActiveState[yard][d][h].man[m-1]) ? 1 : "";
                            const cell = row.getCell(cCol); cell.value = val; cell.alignment = alignCenter; cell.border = getBorders(h === 0, false);
                            if (val === 1) { cell.fill = fillActive; daySum++; weekTotal++; } cCol++;
                        }
                        const dSumCell = row.getCell(cCol); dSumCell.value = daySum || ""; dSumCell.alignment = alignCenter; dSumCell.font = { bold: true }; dSumCell.fill = fillSum; dSumCell.border = getBorders(false, true); cCol++;
                    });
                    const wSumCell = row.getCell(cCol); wSumCell.value = weekTotal; wSumCell.alignment = alignCenter; wSumCell.font = { bold: true }; wSumCell.fill = fillHeader;
                }

                if (!hideVirtual) {
                    for (let k = availK + 1; k <= totalK; k++) {
                        const row = sheet.addRow([`Kamag ${k} (дод.)`]); row.getCell(1).font = { bold: true, size: 10 };
                        let weekTotal = 0, cCol = 2;
                        days.forEach(d => {
                            let daySum = 0;
                            for (let h = 0; h < 24; h++) {
                                let val = (fleetActiveState[yard][d] && fleetActiveState[yard][d][h].kamag[k-1]) ? 1 : "";
                                const cell = row.getCell(cCol); cell.value = val; cell.alignment = alignCenter; cell.border = getBorders(h === 0, false);
                                if (val === 1) { cell.fill = fillActiveVirtual; daySum++; weekTotal++; } cCol++;
                            }
                            const dSumCell = row.getCell(cCol); dSumCell.value = daySum || ""; dSumCell.alignment = alignCenter; dSumCell.font = { bold: true }; dSumCell.fill = fillSum; dSumCell.border = getBorders(false, true); cCol++;
                        });
                        const wSumCell = row.getCell(cCol); wSumCell.value = weekTotal; wSumCell.alignment = alignCenter; wSumCell.font = { bold: true }; wSumCell.fill = fillHeader;
                    }
                    for (let m = availM + 1; m <= totalM; m++) {
                        const row = sheet.addRow([`Маневровий ${m} (дод.)`]); row.getCell(1).font = { bold: true, size: 10 };
                        let weekTotal = 0, cCol = 2;
                        days.forEach(d => {
                            let daySum = 0;
                            for (let h = 0; h < 24; h++) {
                                let val = (fleetActiveState[yard][d] && fleetActiveState[yard][d][h].man[m-1]) ? 1 : "";
                                const cell = row.getCell(cCol); cell.value = val; cell.alignment = alignCenter; cell.border = getBorders(h === 0, false);
                                if (val === 1) { cell.fill = fillActiveVirtual; daySum++; weekTotal++; } cCol++;
                            }
                            const dSumCell = row.getCell(cCol); dSumCell.value = daySum || ""; dSumCell.alignment = alignCenter; dSumCell.font = { bold: true }; dSumCell.fill = fillSum; dSumCell.border = getBorders(false, true); cCol++;
                        });
                        const wSumCell = row.getCell(cCol); wSumCell.value = weekTotal; wSumCell.alignment = alignCenter; wSumCell.font = { bold: true }; wSumCell.fill = fillHeader;
                    }
                }
                sheet.addRow([]);

                const excelOpsConfig = ["Всього операцій", "Задіяно фіз. КАМАГ", "Задіяно вірт. ЛЕГ", "Задіяно фіз. МАН", "Задіяно вірт. МАН", "Непокриті (фіз. флот)", "Непокриті (залишок)"];
                excelOpsConfig.forEach(rowName => {
                    const row = sheet.addRow([rowName]); row.getCell(1).font = { bold: true, size: 10 };
                    let weekTotal = 0, oCol = 2;
                    days.forEach(d => {
                        let daySum = 0;
                        for (let h = 0; h < 24; h++) {
                            let totalOps = (totalOpsData[yard] && totalOpsData[yard][d]) ? totalOpsData[yard][d][h] : 0;
                            let activePhysK = 0, activeVirtK = 0, activePhysM = 0, activeVirtM = 0;
                            if (fleetActiveState[yard] && fleetActiveState[yard][d] && fleetActiveState[yard][d][h]) {
                                const st = fleetActiveState[yard][d][h];
                                activePhysK = st.kamag.slice(0, availK).filter(Boolean).length; activeVirtK = st.kamag.slice(availK).filter(Boolean).length;
                                activePhysM = st.man.slice(0, availM).filter(Boolean).length; activeVirtM = st.man.slice(availM).filter(Boolean).length;
                            }
                            let val = 0, isUncovered = false;
                            if (rowName === "Всього операцій") val = totalOps;
                            else if (rowName === "Задіяно фіз. КАМАГ") val = activePhysK;
                            else if (rowName === "Задіяно вірт. КАМАГ") val = activeVirtK;
                            else if (rowName === "Задіяно фіз. МАН") val = activePhysM;
                            else if (rowName === "Задіяно вірт. МАН") val = activeVirtM;
                            else if (rowName === "Непокриті (фіз. флот)") { val = Math.max(0, totalOps - (activePhysK * yardNorms.k + activePhysM * yardNorms.m)); isUncovered = true; }
                            else if (rowName === "Непокриті (залишок)") { val = Math.max(0, totalOps - ((hideVirtual ? activePhysK : activePhysK + activeVirtK) * yardNorms.k + (hideVirtual ? activePhysM : activePhysM + activeVirtM) * yardNorms.m)); isUncovered = true; }

                            const cell = row.getCell(oCol); cell.value = val || ""; cell.alignment = alignCenter; cell.border = getBorders(h === 0, false);
                            if (rowName === "Всього операцій" && val > 0) cell.fill = fillOps;
                            if (isUncovered && val > 0) { cell.fill = fillUncovered; cell.font = fontUncovered; }
                            daySum += val; weekTotal += val; oCol++;
                        }
                        const dSumCell = row.getCell(oCol); dSumCell.value = daySum || ""; dSumCell.alignment = alignCenter; dSumCell.font = { bold: true }; dSumCell.fill = fillSum; dSumCell.border = getBorders(false, true); oCol++;
                    });
                    const wSumCell = row.getCell(oCol); wSumCell.value = weekTotal; wSumCell.alignment = alignCenter; wSumCell.font = { bold: true }; wSumCell.fill = fillHeader;
                });

                sheet.addRow([]); sheet.addRow(["Графіки:"]).font = { bold: true };
                const imgRow = sheet.rowCount; 
                days.forEach((d, index) => {
                    const chartLabels = [], opsData = [], capacityData = [];
                    for (let h = 0; h < 24; h++) {
                        chartLabels.push(`${h}:00`); opsData.push((totalOpsData[yard] && totalOpsData[yard][d]) ? totalOpsData[yard][d][h] : 0);
                        let cap = 0;
                        if (fleetActiveState[yard] && fleetActiveState[yard][d] && fleetActiveState[yard][d][h]) {
                            const st = fleetActiveState[yard][d][h];
                            cap += (hideVirtual ? st.kamag.slice(0, availK).filter(Boolean).length : st.kamag.filter(Boolean).length) * yardNorms.k;
                            cap += (hideVirtual ? st.man.slice(0, availM).filter(Boolean).length : st.man.filter(Boolean).length) * yardNorms.m;
                        }
                        capacityData.push(cap);
                    }
                    const canvas = document.createElement('canvas'); canvas.width = 620; canvas.height = 100; const ctx = canvas.getContext('2d');
                    const tempChart = new Chart(ctx, { type: 'bar', data: { labels: chartLabels, datasets: [{ type: 'line', data: capacityData, borderColor: '#0d47a1', backgroundColor: '#0d47a1', borderWidth: 2 }, { type: 'bar', data: opsData, backgroundColor: 'rgba(255, 193, 7, 0.7)' }]}, options: { animation: false, responsive: false, scales: { x: { display: false }, y: { display: false, beginAtZero: true } }, plugins: { legend: { display: false } } } });
                    const base64 = canvas.toDataURL('image/png'); tempChart.destroy();
                    sheet.addImage(workbook.addImage({ base64: base64.split(',')[1], extension: 'png' }), { tl: { col: 1 + index * 25, row: imgRow }, ext: { width: 620, height: 100 } });
                });
                for(let i=0; i<6; i++) sheet.addRow([]); 
            });
        }
        
        // === ТЕПЕР saveAs ЗНАХОДИТЬСЯ ТУТ (ВИЛИП ІЗ ТАБКАМАГ) І ПРАЦЮЄ ДЛЯ ВСІХ ВКЛАДОК ===
        saveAs(new Blob([await workbook.xlsx.writeBuffer()]), `Kamagon_Export_${new Date().getTime()}.xlsx`);
    } catch (err) { 
        console.error(err); 
        alert("Помилка при експорті!"); 
    } finally { 
        btn.innerText = originalText; 
        btn.disabled = false; 
    }
});

// СОХРАНЕНИЕ ТЕКУЩЕГО ДВОРА
document.getElementById('saveGoogleBtn').addEventListener('click', async () => {
    const yard = document.getElementById('kamagYardSelect').value;
    if (!yard) return alert("Оберіть автодвір!");
    const btn = document.getElementById('saveGoogleBtn'); btn.innerText = "⏳...";

    const aggregatedRows = []; const days = Object.keys(totalOpsData[yard] || {});
    days.forEach(day => {
        for (let h = 0; h < 24; h++) {
            const opsCount = (totalOpsData[yard] && totalOpsData[yard][day]) ? totalOpsData[yard][day][h] : 0;
            let stateString = "0|0", hasActive = false;
            if (fleetActiveState[yard] && fleetActiveState[yard][day] && fleetActiveState[yard][day][h]) {
                const kArr = fleetActiveState[yard][day][h].kamag; const mArr = fleetActiveState[yard][day][h].man;
                stateString = `${kArr.join(',')}|${mArr.join(',')}`; hasActive = kArr.some(v => v > 0) || mArr.some(v => v > 0);
            }
            if (opsCount > 0 || hasActive) aggregatedRows.push([yard, day, h, stateString, opsCount]);
        }
    });

    try {
        await fetch(RESULTS_SCRIPT_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'saveAggregated', yard: yard, rows: aggregatedRows, dates: days }) });
        btn.innerText = "✅ Збережено!";
    } catch (e) { btn.innerText = "❌ Помилка"; }
    setTimeout(() => btn.innerText = "Зберегти (поточний)", 3000);
});

// СОХРАНЕНИЕ ВСЕХ ДВОРОВ
document.getElementById('saveAllGoogleBtn').addEventListener('click', async () => {
    const yards = Object.keys(totalOpsData); if (yards.length === 0) return alert("Немає розрахованих данных для збереження!");
    const btn = document.getElementById('saveAllGoogleBtn'); btn.innerText = "⏳..."; btn.disabled = true;

    const aggregatedRows = [], allDays = new Set();
    yards.forEach(yard => {
        const days = Object.keys(totalOpsData[yard] || {});
        days.forEach(day => {
            allDays.add(day);
            for (let h = 0; h < 24; h++) {
                const opsCount = (totalOpsData[yard] && totalOpsData[yard][day]) ? totalOpsData[yard][day][h] : 0;
                let stateString = "0|0", hasActive = false;
                if (fleetActiveState[yard] && fleetActiveState[yard][day] && fleetActiveState[yard][day][h]) {
                    const kArr = fleetActiveState[yard][day][h].kamag; const mArr = fleetActiveState[yard][day][h].man;
                    stateString = `${kArr.join(',')}|${mArr.join(',')}`; hasActive = kArr.some(v => v > 0) || mArr.some(v => v > 0);
                }
                if (opsCount > 0 || hasActive) aggregatedRows.push([yard, day, h, stateString, opsCount]);
            }
        });
    });

    try {
        await fetch(RESULTS_SCRIPT_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'saveAllAggregated', yards: yards, rows: aggregatedRows, dates: Array.from(allDays) }) });
        btn.innerText = "✅ Всі збережено!";
    } catch (e) { btn.innerText = "❌ Помилка"; }
    setTimeout(() => { btn.innerText = "Зберегти ВСІ"; btn.disabled = false; }, 3000);
});

document.addEventListener('DOMContentLoaded', () => {
    tabKamag.click();
    document.querySelectorAll('input[name="virtualFleetType"]').forEach(radio => {
        radio.addEventListener('change', () => { if (Object.keys(totalOpsData).length > 0) calculateFleetRequirements(); });
    });
});

document.getElementById('hideVirtualFleet').addEventListener('change', renderKamagTable);
document.getElementById('kamagStartDate').addEventListener('change', renderKamagTable);
document.getElementById('kamagEndDate').addEventListener('change', renderKamagTable);