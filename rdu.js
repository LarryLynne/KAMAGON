// === МОДУЛЬ РУЧНОГО ВВОДУ ДЛЯ ДИСПЕТЧЕРІВ РДУ (rdu.js) ===
let rduStateMatrix = {}; // Хранилище: { "07.06.2026": { "Kamag 1": [0,1,1,...], ... } }
let isRduMouseDown = false;
let rduPaintMode = 1; // 1 - закрашивать, 0 - -стирать

let lastPaintedHour = null;
let lastPaintedVehicle = null;

let lockedRduDates = {};

// НОВА ФУНКЦІЯ: Перевіряє статус і блокує/розблоковує кнопку
function updateRduSaveButtonState() {
    const role = sessionStorage.getItem('kamagonAuthRole');
    const userYard = getSelectedRduYard();
    const rawDate = document.getElementById('rduWorkDate').value;
    const btn = document.getElementById('saveRduGoogleBtn');
    
    if (!rawDate || !userYard || !btn) return;

    const [yyyy, mm, dd] = rawDate.split('-');
    const dateStr = `${dd}.${mm}.${yyyy}`;
    const lockKey = `${userYard}_${dateStr}`;

    // Якщо роль РДУ і цей день вже збережений/завантажений
    if (role === 'РДУ' && lockedRduDates[lockKey]) {
        btn.disabled = true;
        btn.innerText = "🔒 Дані вже в базі";
        btn.style.backgroundColor = "#e2e8f0"; // Сірий фон
        btn.style.color = "#64748b";           // Сірий текст
        btn.style.cursor = "not-allowed";
        btn.style.borderColor = "#cbd5e1";
    } else {
        btn.disabled = false;
        btn.innerText = "Зберегти день";
        btn.style.backgroundColor = ""; 
        btn.style.color = "";
        btn.style.cursor = "pointer";
        btn.style.borderColor = "";
    }
}

document.addEventListener('DOMContentLoaded', () => {
    const tabRdu = document.getElementById('tabRdu');
    const containerRdu = document.getElementById('tableContainerRdu');
    const rduDateInput = document.getElementById('rduWorkDate');

    if (!tabRdu || !containerRdu) return;

    const today = new Date();
    rduDateInput.value = today.toISOString().split('T')[0];

    // ЗАМЕНИТЬ обработчик tabRdu.addEventListener:
    tabRdu.addEventListener('click', () => {
        switchTab(tabRdu, containerRdu);
        
        const container = document.getElementById('rduYardContainer');
        
        // Создаем селект один раз, если его еще нет
        if (!document.getElementById('rduYardSelect')) {
            container.innerHTML = `Автодвір: <select id="rduYardSelect" style="height: 28px; padding: 0 5px; font-size: 11px; min-width: 140px; border-color: #adb5bd;"></select>`;
            const select = document.getElementById('rduYardSelect');
            
            const role = sessionStorage.getItem('kamagonAuthRole');
            
            if (role === 'Адмін') {
                // Админ видит все дворы
                const uniqueYards = Object.keys(yardDictionary).map(k => yardDictionary[k].yard).filter((v, i, a) => v && a.indexOf(v) === i).sort();
                uniqueYards.forEach(y => {
                    const opt = document.createElement('option');
                    opt.value = opt.textContent = y;
                    select.appendChild(opt);
                });
            } else {
                // РДУ видит только свои дворы (поддержка нескольких через запятую)
                const userYard = sessionStorage.getItem('kamagonAuthYard') || "";
                const allowedYards = userYard.split(',').map(y => y.trim()).filter(Boolean);
                
                allowedYards.forEach(y => {
                    const opt = document.createElement('option');
                    opt.value = opt.textContent = y;
                    select.appendChild(opt);
                });
                
                // Если двор только один - блокируем выбор. Если больше - можно переключаться!
                if (allowedYards.length <= 1) {
                    select.disabled = true;
                }
            }
            
            select.addEventListener('change', renderRduGrid);
            
            // Сразу делаем его красивым и с поиском
            if (typeof makeSelectSearchable === 'function') {
                makeSelectSearchable('rduYardSelect');
            }
        }
        
        renderRduGrid();
    });

    rduDateInput.addEventListener('change', renderRduGrid);
    document.getElementById('loadRduGoogleBtn').addEventListener('click', loadRduFromGoogle);
    document.getElementById('saveRduGoogleBtn').addEventListener('click', saveRduToGoogle);
    
    // ОБРАБОТЧИК ДЛЯ РУЧНОГО ДОБАВЛЕНИЯ ТРАНСПОРТА
    document.getElementById('rduAddVehicleBtn').addEventListener('click', () => {
        const userYard = getSelectedRduYard();
        const rawDate = rduDateInput.value;
        if (!rawDate || !userYard) return alert("Оберіть дату та автодвір!");

        const [yyyy, mm, dd] = rawDate.split('-');
        const dateStr = `${dd}.${mm}.${yyyy}`;
        
        if (!rduStateMatrix[dateStr]) rduStateMatrix[dateStr] = {};

        const type = document.getElementById('rduVehicleTypeSelect').value;
        const avail = (typeof fleetDictionary !== 'undefined' && fleetDictionary[userYard]) ? 
                      (type === 'Kamag' ? fleetDictionary[userYard].kamag : fleetDictionary[userYard].man) : 0;
        
        let maxNum = avail;
        Object.keys(rduStateMatrix[dateStr]).forEach(v => {
            if (v.startsWith(type)) {
                const match = v.match(/\d+/);
                if (match) maxNum = Math.max(maxNum, parseInt(match[0], 10));
            }
        });

        const newVName = `${type} ${maxNum + 1}${maxNum + 1 > avail ? ' (дод.)' : ''}`;
        rduStateMatrix[dateStr][newVName] = Array(24).fill(0);
        
        renderRduGrid();
    });

    document.addEventListener('mouseup', () => isRduMouseDown = false);
});

function getSelectedRduYard() {
    // Теперь мы всегда читаем значение из селекта, независимо от роли
    const select = document.getElementById('rduYardSelect');
    return select ? select.value : "";
}

function renderRduGrid() {
    const wrapper = document.getElementById('rduTableWrapper');
    const userYard = getSelectedRduYard(); 
    const rawDate = document.getElementById('rduWorkDate').value;

    if (!rawDate) {
        wrapper.innerHTML = "<p class='disabled'>Оберіть день роботи.</p>";
        return;
    }
    if (!userYard) {
        wrapper.innerHTML = "<p class='disabled'>Автодвір не визначено.</p>";
        return;
    }

    const [yyyy, mm, dd] = rawDate.split('-');
    const dateStr = `${dd}.${mm}.${yyyy}`;

    const availK = (typeof fleetDictionary !== 'undefined' && fleetDictionary[userYard]) ? fleetDictionary[userYard].kamag : 0;
    const availM = (typeof fleetDictionary !== 'undefined' && fleetDictionary[userYard]) ? fleetDictionary[userYard].man : 0;

    if (!rduStateMatrix[dateStr]) rduStateMatrix[dateStr] = {};

    // Динамически вычисляем максимальный индекс строк на сетке (база + добавленные вручную)
    let maxK = availK;
    let maxM = availM;
    Object.keys(rduStateMatrix[dateStr]).forEach(v => {
        const match = v.match(/\d+/);
        if (match) {
            const num = parseInt(match[0], 10);
            if (v.startsWith('Kamag')) maxK = Math.max(maxK, num);
            if (v.startsWith('Маневровий')) maxM = Math.max(maxM, num);
        }
    });

    if (maxK === 0 && maxM === 0) {
        wrapper.innerHTML = `<p class='disabled'>У довіднику не знайдено флоту для автодвору ${userYard}.</p>`;
        return;
    }

    const vehicleRows = [];
    // 1. Спочатку весь фізичний (наявний) транспорт
    for (let i = 1; i <= availK; i++) vehicleRows.push(`Kamag ${i}`);
    for (let i = 1; i <= availM; i++) vehicleRows.push(`Маневровий ${i}`);
    
    // 2. Потім весь додатковий (віртуальний) транспорт знизу
    for (let i = availK + 1; i <= maxK; i++) vehicleRows.push(`Kamag ${i} (дод.)`);
    for (let i = availM + 1; i <= maxM; i++) vehicleRows.push(`Маневровий ${i} (дод.)`);

    vehicleRows.forEach(v => {
        if (!rduStateMatrix[dateStr][v]) rduStateMatrix[dateStr][v] = Array(24).fill(0);
    });

    let html = `<table class="kamag-table" style="user-select: none;">
        <thead>
            <tr>
                <th style="min-width: 150px; background-color: #f1f5f9;">ТЗ / Години роботи</th>`;
    
    for (let h = 0; h < 24; h++) {
        html += `<th class="kamag-header-vertical" style="height:30px;">${h}:00</th>`;
    }
    html += `<th style="background-color: #e2e8f0; font-weight:bold;">Σ, год</th></tr></thead><tbody>`;

    vehicleRows.forEach(v => {
        html += `<tr><td style="font-weight:bold; background-color:#fff;">${v}</td>`;
        let dailySum = 0;

        for (let h = 0; h < 24; h++) {
            const isRowActive = rduStateMatrix[dateStr][v][h] === 1;
            let cellClass = "kamag-cell rdu-editable-cell";
            if (isRowActive) {
                cellClass += " kamag-manual-physical";
                dailySum++;
            }
            html += `<td class="${cellClass}" data-vehicle="${v}" data-hour="${h}"></td>`;
        }
        html += `<td class="rdu-row-sum" style="text-align:center; font-weight:bold; background-color:#f1f3f5;">${dailySum || ''}</td></tr>`;
    });

    html += `</tbody></table>`;
    wrapper.innerHTML = html;

    attachRduMouseEvents();
    updateRduSaveButtonState();
}

function attachRduMouseEvents() {
    const cells = document.querySelectorAll('.rdu-editable-cell');
    const rawDate = document.getElementById('rduWorkDate').value;
    if (!rawDate) return;
    const [yyyy, mm, dd] = rawDate.split('-');
    const dateStr = `${dd}.${mm}.${yyyy}`;

    cells.forEach(cell => {
        cell.addEventListener('mousedown', (e) => {
            e.preventDefault(); 
            isRduMouseDown = true;

            const vehicle = cell.getAttribute('data-vehicle');
            const hour = parseInt(cell.getAttribute('data-hour'), 10);

            rduPaintMode = rduStateMatrix[dateStr][vehicle][hour] === 0 ? 1 : 0;
            rduStateMatrix[dateStr][vehicle][hour] = rduPaintMode;
            
            if (rduPaintMode === 1) cell.classList.add('kamag-manual-physical');
            else cell.classList.remove('kamag-manual-physical');
            
            lastPaintedHour = hour;
            lastPaintedVehicle = vehicle;
            updateRowSum(cell.parentElement);
        });

        cell.addEventListener('mouseenter', () => {
            if (!isRduMouseDown) return;

            const vehicle = cell.getAttribute('data-vehicle');
            const hour = parseInt(cell.getAttribute('data-hour'), 10);

            if (vehicle === lastPaintedVehicle && lastPaintedHour !== null) {
                const start = Math.min(lastPaintedHour, hour);
                const end = Math.max(lastPaintedHour, hour);
                
                for (let h = start; h <= end; h++) {
                    rduStateMatrix[dateStr][vehicle][h] = rduPaintMode;
                    const targetCell = cell.parentElement.querySelector(`[data-hour="${h}"]`);
                    if (targetCell) {
                        if (rduPaintMode === 1) targetCell.classList.add('kamag-manual-physical');
                        else targetCell.classList.remove('kamag-manual-physical');
                    }
                }
            } else {
                rduStateMatrix[dateStr][vehicle][hour] = rduPaintMode;
                if (rduPaintMode === 1) cell.classList.add('kamag-manual-physical');
                else cell.classList.remove('kamag-manual-physical');
            }

            lastPaintedHour = hour;
            lastPaintedVehicle = vehicle;
            updateRowSum(cell.parentElement);
        });
    });
}

function updateRowSum(trElement) {
    const activeCells = trElement.querySelectorAll('.rdu-editable-cell.kamag-manual-physical').length;
    trElement.querySelector('.rdu-row-sum').innerText = activeCells > 0 ? activeCells : '';
}

async function saveRduToGoogle() {
    const userYard = getSelectedRduYard();
    const activeUser = sessionStorage.getItem('kamagonAuthUser') || "Невідомий";
    const role = sessionStorage.getItem('kamagonAuthRole');
    const rawDate = document.getElementById('rduWorkDate').value;
    
    if (!rawDate || !userYard) return alert("Оберіть дату та автодвір!");

    const [yyyy, mm, dd] = rawDate.split('-');
    const dateStr = `${dd}.${mm}.${yyyy}`;
    const lockKey = `${userYard}_${dateStr}`;
    const btn = document.getElementById('saveRduGoogleBtn');

    // 1. Локальна перевірка (якщо ми ВЖЕ знаємо, що день заблоковано)
    if (role === 'РДУ' && lockedRduDates[lockKey]) {
        return alert("Ці дані вже були збережені раніше. Для перезапису зверніться до Адміністратора.");
    }

    btn.innerText = "⏳ Перевірка..."; 
    btn.disabled = true;

    // 2. ЗАХИСТ ВІД СЛІПОГО ПЕРЕЗАПИСУ:
    // Якщо це РДУ, і статус цього дня ще невідомий (undefined), робимо швидкий запит до бази
    if (role === 'РДУ' && lockedRduDates[lockKey] === undefined) {
        try {
            const checkRes = await fetch(`${RESULTS_SCRIPT_URL}?action=getRduAggregatedData&yard=${encodeURIComponent(userYard)}`);
            const checkData = await checkRes.json();
            
            let hasExistingData = false;
            if (checkData.savedRows && checkData.savedRows.length > 0) {
                // Шукаємо, чи є в базі хоча б один рядок саме за цю дату
                hasExistingData = checkData.savedRows.some(row => {
                    let dayStr = String(row[1]);
                    if (dayStr.includes('T') && dayStr.includes('Z')) {
                        const d = new Date(dayStr);
                        dayStr = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
                    }
                    return dayStr === dateStr;
                });
            }
            
            if (hasExistingData) {
                // Дані вже є! Блокуємо і скасовуємо збереження
                lockedRduDates[lockKey] = true;
                updateRduSaveButtonState();
                return alert("Виявлено існуючі дані в базі! Ви не можете перезаписати їх наосліп. Зверніться до Адміністратора.");
            } else {
                // База порожня для цього дня, дозволяємо зберігати
                lockedRduDates[lockKey] = false; 
            }
        } catch (e) {
            btn.innerText = "❌ Помилка зв'язку";
            setTimeout(() => updateRduSaveButtonState(), 2000);
            return alert("Не вдалося перевірити базу даних перед збереженням. Спробуйте ще раз.");
        }
    }

    // 3. Якщо перевірка пройдена успішно (або це Адмін) — починаємо збереження
    btn.innerText = "⏳ Збереження..."; 

    const vehicles = Object.keys(rduStateMatrix[dateStr] || {});
    const kamags = vehicles.filter(v => v.startsWith('Kamag')).sort((a,b) => parseInt(a.match(/\d+/)) - parseInt(b.match(/\d+/)));
    const mans = vehicles.filter(v => v.startsWith('Маневровий')).sort((a,b) => parseInt(a.match(/\d+/)) - parseInt(b.match(/\d+/)));

    const aggregatedRows = [];

    for (let h = 0; h < 24; h++) {
        const kStates = kamags.map(k => rduStateMatrix[dateStr][k][h]);
        const mStates = mans.map(m => rduStateMatrix[dateStr][m][h]);
        
        const stateString = `${kStates.join(',')}|${mStates.join(',')}`;
        const hasActive = kStates.some(v => v > 0) || mStates.some(v => v > 0);

        if (hasActive) {
            aggregatedRows.push([userYard, dateStr, h, stateString, 0, activeUser]);
        }
    }

    try {
        await fetch(RESULTS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'saveRduAggregated', yard: userYard, rows: aggregatedRows, dates: [dateStr] })
        });
        
        // Після успішного збереження відразу блокуємо день для РДУ
        lockedRduDates[lockKey] = true;
        btn.innerText = "✅ Збережено!";
    } catch (e) {
        alert("Помилка збереження.");
        btn.innerText = "❌ Помилка";
    }
    
    setTimeout(() => { 
        updateRduSaveButtonState(); 
    }, 2000);
}
async function loadRduFromGoogle() {
    const userYard = getSelectedRduYard();
    const rawDate = document.getElementById('rduWorkDate').value;
    if (!rawDate || !userYard) return alert("Оберіть дату та автодвір!");

    const [yyyy, mm, dd] = rawDate.split('-');
    const dateStr = `${dd}.${mm}.${yyyy}`;

    const btn = document.getElementById('loadRduGoogleBtn');
    btn.innerText = "⏳...";

    const availK = (typeof fleetDictionary !== 'undefined' && fleetDictionary[userYard]) ? fleetDictionary[userYard].kamag : 0;
    const availM = (typeof fleetDictionary !== 'undefined' && fleetDictionary[userYard]) ? fleetDictionary[userYard].man : 0;

    try {
        const response = await fetch(`${RESULTS_SCRIPT_URL}?action=getRduAggregatedData&yard=${encodeURIComponent(userYard)}`);
        const data = await response.json();

        if (rduStateMatrix[dateStr]) {
            Object.keys(rduStateMatrix[dateStr]).forEach(v => rduStateMatrix[dateStr][v].fill(0));
        } else {
            rduStateMatrix[dateStr] = {};
        }

        if (data.savedRows && data.savedRows.length > 0) {
            let hasDataForCurrentDay = false; // Трекаємо чи є дані саме за цей день

            data.savedRows.forEach(row => {
                let [y, day, hour, fleetCountStr] = row;
                let dayStr = String(day);
                if (dayStr.includes('T') && dayStr.includes('Z')) {
                    const d = new Date(dayStr);
                    dayStr = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
                }

                if (dayStr === dateStr) {
                    hasDataForCurrentDay = true; // Знайшли дані!
                    
                    let countStr = String(fleetCountStr);
                    const [kStr, mStr] = countStr.split('|');
                    
                    if (kStr) {
                        const kBits = kStr.split(',').map(Number);
                        kBits.forEach((bit, idx) => {
                            const vName = `Kamag ${idx + 1}${idx + 1 > availK ? ' (дод.)' : ''}`;
                            if (!rduStateMatrix[dateStr][vName]) rduStateMatrix[dateStr][vName] = Array(24).fill(0);
                            rduStateMatrix[dateStr][vName][hour] = bit;
                        });
                    }
                    if (mStr) {
                        const mBits = mStr.split(',').map(Number);
                        mBits.forEach((bit, idx) => {
                            const vName = `Маневровий ${idx + 1}${idx + 1 > availM ? ' (дод.)' : ''}`;
                            if (!rduStateMatrix[dateStr][vName]) rduStateMatrix[dateStr][vName] = Array(24).fill(0);
                            rduStateMatrix[dateStr][vName][hour] = bit;
                        });
                    }
                }
            });

            // Блокуємо, якщо дані знайдені
            const lockKey = `${userYard}_${dateStr}`;
            lockedRduDates[lockKey] = hasDataForCurrentDay;

            document.getElementById('fileStatus').innerText = "Дані РДУ завантажено з бази.";
        } else {
            // Якщо даних немає, знімаємо блокування
            const lockKey = `${userYard}_${dateStr}`;
            lockedRduDates[lockKey] = false;
            alert("Збережених ручних даних РДУ для цього дня не знайдено.");
        }
        renderRduGrid();
    } catch (e) {
        alert("Помилка завантаження з бази");
        console.error(e);
    } finally {
        btn.innerText = "Завантажити з бази";
    }
}