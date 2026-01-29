const STATE_KEY = 'trackingAppStateV3';
// KONFIGURACJA
// Wklej tutaj URL swojej aplikacji Google Apps Script po wdrożeniu:
const GOOGLE_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbztqUP8saUm76ZzeWqKxght5uvjpUSQu5PrLx4EGarL1LgKGt9aU9aaSLli_WYojuVU/exec';

// === GLOBAL STATE ===
let appState = {
  lines: {}, // { uniqueId: { name: 'Stanowisko 1', workerName: '', ... } }
  globalSessions: [], // Unified history of all completed sessions
  globalBreaks: [], // Unified history of all breaks
  activeLineId: null
};

// Helper references to clean up old intervals
let intervalRegistry = {
  // lineId: { work: null, break: null }
};

// Elementy DOM (Globalne / Wspólne)
const el = {
  screen0: document.getElementById('screen0'),
  screen1: document.getElementById('screen1'),
  screen2: document.getElementById('screen2'),
  screen3: document.getElementById('screen3'),
  linesGrid: document.getElementById('linesGrid'),

  // INPUTY EKRAN 1
  workerName: document.getElementById('workerName'),
  machine: document.getElementById('machine'),
  peopleCount: document.getElementById('peopleCount'),
  startTime: document.getElementById('startTime'),
  startTimeDisplay: document.getElementById('startTimeDisplay'),

  // EKRAN 2 (DISPLAY)
  displayWorkerName: document.getElementById('displayWorkerName'),
  displayMachine: document.getElementById('displayMachine'),
  displayPeopleCount: document.getElementById('displayPeopleCount'),
  displayCurrentProduct: document.getElementById('displayCurrentProduct'),
  totalLabeledDisplay: document.getElementById('totalLabeledDisplay'),

  // EKRAN 2 (CONTROLS)
  productSearch: document.getElementById('productSearch'),
  productList: document.getElementById('productList'),

  workTimer: document.getElementById('workTimer'),
  workTimerContainer: document.getElementById('workTimerContainer'),
  breakTimer: document.getElementById('breakTimer'),
  breakTimerContainer: document.getElementById('breakTimerContainer'),

  startBtn: document.getElementById('startBtn'),
  breakBtn: document.getElementById('breakBtn'),
  stopBtn: document.getElementById('stopBtn'),
  resumeBtn: document.getElementById('resumeBtn'),

  // EKRAN 2 (INPUTS CLOSE)
  quantityInput: document.getElementById('quantityInput'),
  quantityValidation: document.getElementById('quantityValidation'),
  labelTypeSelect: document.getElementById('labelTypeSelect'),
  labelTypeValidation: document.getElementById('labelTypeValidation'),
  batchNumberInput: document.getElementById('batchNumberInput'),
  batchValidation: document.getElementById('batchValidation'),
  productionDateInput: document.getElementById('productionDateInput'),
  prodDateValidation: document.getElementById('prodDateValidation'),
  cardNumberInput: document.getElementById('cardNumberInput'),
  cardValidation: document.getElementById('cardValidation'),
  notesInput: document.getElementById('notesInput'),

  // MODALS & TABLES
  breakModal: document.getElementById('breakModal'),
  breakReason: document.getElementById('breakReason'),
  breakDescription: document.getElementById('breakDescription'),
  descriptionRequired: document.getElementById('descriptionRequired'),
  breakValidation: document.getElementById('breakValidation'),

  currentBreaksContainer: document.getElementById('currentBreaksContainer'),
  currentBreaksBody: document.getElementById('currentBreaksBody'),

  labelingHistoryBody: document.getElementById('labelingHistoryBody'),
  breakHistoryBody: document.getElementById('breakHistoryBody'),
  peopleChangeHistoryBody: document.getElementById('peopleChangeHistoryBody'),

  changePeopleModal: document.getElementById('changePeopleModal'),
  newPeopleCount: document.getElementById('newPeopleCount'),

  // EKRAN 3 GLOBAL SUMMARY
  screen3: document.getElementById('screen3'),
  summaryLabelingBody: document.getElementById('summaryLabelingBody'),
  summaryBreakHistoryBody: document.getElementById('summaryBreakHistoryBody'),
  summaryPeopleChangeBody: document.getElementById('summaryPeopleChangeBody'),

  // Stats Elements
  totalSessions: document.getElementById('totalSessions'),
  totalPieces: document.getElementById('totalPieces'),
  totalWorkMinutes: document.getElementById('totalWorkMinutes'),
  totalBreakMinutes: document.getElementById('totalBreakMinutes'),

  // Break Local

};

// --- UTILS ---
function generateId() {
  return 'line_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

function formatHMS(sec) {
  const h = Math.floor(sec / 3600).toString().padStart(2, '0');
  const m = Math.floor((sec % 3600) / 60).toString().padStart(2, '0');
  const s = Math.floor(sec % 60).toString().padStart(2, '0');
  return `${h}:${m}:${s}`;
}

function formatTime(date) {
  if (!date) return '';
  return new Date(date).toLocaleTimeString('pl-PL', { hour: '2-digit', minute: '2-digit' });
}

function getActiveLine() {
  if (!appState.activeLineId) return null;
  return appState.lines[appState.activeLineId];
}

// --- STATE MANAGEMENT ---
function saveState() {
  try {
    // Before save, make sure current UI inputs are synced to active line state
    syncInputsToState();
    localStorage.setItem(STATE_KEY, JSON.stringify(appState));
  } catch (e) {
    console.warn('Save state failed', e);
  }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STATE_KEY);
    if (!raw) return false;
    const state = JSON.parse(raw);
    if (!state || !state.lines) return false;

    appState = state;
    // Restore intervals for running lines
    Object.keys(appState.lines).forEach(lineId => {
      const line = appState.lines[lineId];
      if (line.isWorking) startRealIntervals(lineId);
    });

    // If we were on screen 1 or 2, verify active line exists
    if (appState.activeLineId && !appState.lines[appState.activeLineId]) {
      appState.activeLineId = null;
      currentScreen = 0;
    }

    renderDashboard();
    return true;
  } catch (e) {
    console.warn('Load state failed', e);
    return false;
  }
}

// Capture inputs from Screen 1 & 2 into the Active Line object
function syncInputsToState() {
  if (!appState.activeLineId) return;
  const line = appState.lines[appState.activeLineId];
  if (!line) return;

  // Screen 1 Inputs
  if (currentScreen === 1) {
    line.workerName = el.workerName.value;
    line.machine = el.machine.value;
    line.peopleCount = el.peopleCount.value;
    line.plannedStartTime = el.startTime.value;
  }

  // Screen 2 Inputs (Drafts)
  if (currentScreen === 2) {
    // We don't overwrite configuration if it was already set by "Start Work"
    // But we do save the 'Stop' form drafts
    line.drafts = {
      quantity: el.quantityInput.value,
      labelType: el.labelTypeSelect.value,
      batchNumber: el.batchNumberInput.value,
      productionDate: el.productionDateInput.value,
      cardNumber: el.cardNumberInput.value,
      notes: el.notesInput.value,
      packagingType: document.querySelector('input[name="packagingType"]:checked')?.value || 'Folia'
    };
  }
}

// Populate UI from Line Object
function restoreInputsFromState() {
  if (!appState.activeLineId) {
    // Clear inputs if no line
    el.workerName.value = '';
    el.machine.value = '';
    el.peopleCount.value = '';
    el.startTime.value = '';
    el.startTimeDisplay.value = '';
    return;
  }
  const line = appState.lines[appState.activeLineId];

  // Screen 1
  el.workerName.value = line.workerName || '';
  el.machine.value = line.machine || '';
  el.peopleCount.value = line.peopleCount || '';
  el.startTime.value = line.plannedStartTime || '';
  el.startTimeDisplay.value = line.plannedStartTime || '';

  // Screen 2 Display
  el.displayWorkerName.textContent = line.config?.workerName || line.workerName || '-';
  el.displayMachine.textContent = line.config?.machine || line.machine || '-';
  el.displayPeopleCount.textContent = line.peopleCount || '-';
  el.displayCurrentProduct.textContent = line.currentProduct || '—';

  // Screen 2 Controls
  if (line.drafts) {
    el.quantityInput.value = line.drafts.quantity || '';
    el.labelTypeSelect.value = line.drafts.labelType || '';
    el.batchNumberInput.value = line.drafts.batchNumber || '';
    el.productionDateInput.value = line.drafts.productionDate || '';
    el.cardNumberInput.value = line.drafts.cardNumber || '';
    el.notesInput.value = line.drafts.notes || '';

    if (line.drafts.packagingType) {
      const radio = document.querySelector(`input[name="packagingType"][value="${line.drafts.packagingType}"]`);
      if (radio) radio.checked = true;
    }
  } else {
    // Default packing
    const radio = document.querySelector(`input[name="packagingType"][value="Folia"]`);
    if (radio) radio.checked = true;
  }

  // Update Search
  if (el.productSearch) el.productSearch.value = line.currentProduct || '';

  // Timers Visuals
  updateTimerVisuals(line);
  updateButtonStates(line);

  // Tables (Local History for this Line)
  updateLocalHistoryTables(line);

  // Current Breaks (Local)
  updateCurrentBreaksTable(line);

  // People Change History (Local)
  updatePeopleChangeTable(line);
}

let currentScreen = 0;

function showScreen(n) {
  // Hide all
  el.screen0.classList.remove('active');
  el.screen1.classList.remove('active');
  el.screen2.classList.remove('active');
  el.screen3.classList.remove('active');

  currentScreen = n;

  if (n === 0) {
    appState.activeLineId = null;
    el.screen0.classList.add('active');
    renderDashboard();
  }
  else if (n === 1) {
    restoreInputsFromState();
    el.screen1.classList.add('active');
  }
  else if (n === 2) {
    restoreInputsFromState();
    el.screen2.classList.add('active');
    const line = getActiveLine();
    if (line) updateCurrentBreaksTable(line);
  }
  else if (n === 3) {
    el.screen3.classList.add('active');
    generateGlobalSummary();
  }
  saveState();
}

// --- DASHBOARD LOGIC ---
function addNewLine() {
  const id = generateId();
  // Find the highest existing line number
  let maxNum = 0;
  Object.values(appState.lines).forEach(l => {
    const match = l.name.match(/\d+/);
    if (match) {
      const num = parseInt(match[0], 10);
      if (num > maxNum) maxNum = num;
    }
  });
  const newNum = maxNum + 1;

  appState.lines[id] = {
    id: id,
    name: `Linia ${newNum}`,
    workerName: '',
    machine: '',
    peopleCount: '',
    plannedStartTime: '', // "Godzina rozpoczęcia" (Shift start)

    // Session State
    config: null, // Locked config when "Start Work Day" clicked { worker, machine... }
    currentProduct: '',
    workStartAt: null, // Timestamp when "Start" clicked
    isOnBreak: false,
    breakStartAt: null,
    totalBreakSeconds: 0,
    breaksCurrent: [], // Breaks in THIS active session

    peopleCountHistory: [], // Specific to this line
    drafts: {}
  };
  renderDashboard();
  saveState();
}

function renderDashboard() {
  el.linesGrid.innerHTML = '';
  Object.values(appState.lines).forEach(line => {
    const card = document.createElement('div');
    card.className = 'line-card';

    // Determine status
    let statusClass = '';
    let statusText = 'Oczekiwanie';
    let timerPreview = '--:--:--';

    if (line.isOnBreak) {
      statusClass = 'status-break';
      statusText = 'Przerwa';
      if (line.breakStartAt) {
        const sec = Math.floor((Date.now() - line.breakStartAt) / 1000);
        timerPreview = formatHMS(sec);
      }
    } else if (line.isWorking) {
      statusClass = 'status-active';
      statusText = 'Praca';
      if (line.workStartAt) {
        const sec = Math.floor((Date.now() - line.workStartAt) / 1000);
        timerPreview = formatHMS(sec);
      }
    } else if (!line.config) {
      statusClass = ''; // Gray
      statusText = 'Nie skonfigurowano';
      timerPreview = '';
    } else if (line.currentProduct) {
      statusClass = 'status-stopped';
      statusText = 'Gotowy';
      timerPreview = '00:00:00';
    } else if (line.config) {
      statusClass = 'status-stopped'; // Orange for 'Ready but waiting'
      statusText = 'Oczekiwanie';
      timerPreview = '00:00:00';
    }

    if (statusClass && statusClass.trim() !== '') {
      card.classList.add(statusClass);
    }

    card.innerHTML = `
            <div class="delete-line-btn" title="Usuń linię" onclick="event.stopPropagation(); confirmDeleteLine('${line.id}');">✕</div>
            <span class="status-badge">${statusText}</span>
            <h3>${line.name}</h3>
            <div class="info-row"><span>Pracownik:</span> <strong>${line.config?.workerName || line.workerName || '-'}</strong></div>
            <div class="info-row"><span>Maszyna:</span> <strong>${line.config?.machine || line.machine || '-'}</strong></div>
            <div class="info-row"><span>Produkt:</span> <strong>${line.currentProduct || '-'}</strong></div>
            <div class="timer-preview" id="preview_${line.id}">${timerPreview}</div>
         `;

    // Only show delete button if line is not configured (or whatever user rule)
    if (!line.config) {
      const delBtn = card.querySelector('.delete-line-btn');
      if (delBtn) delBtn.style.display = 'block';
    }

    card.onclick = () => enterLine(line.id);
    el.linesGrid.appendChild(card);
  });

  // Start dashboard updater for live timers
  if (!window.dashboardInterval) {
    window.dashboardInterval = setInterval(updateDashboardTimers, 1000);
  }
}

function confirmDeleteLine(lineId) {
  const line = appState.lines[lineId];
  if (!line) return;

  if (line.isWorking) {
    alert('Nie można usunąć aktywnej linii! Zatrzymaj pracę najpierw.');
    return;
  }

  if (confirm(`Czy na pewno chcesz usunąć "${line.name}"?\nTa operacja jest nieodwracalna.`)) {
    delete appState.lines[lineId];
    saveState();
    renderDashboard();
  }
}

function updateDashboardTimers() {
  if (currentScreen !== 0) return;
  Object.values(appState.lines).forEach(line => {
    const previewEl = document.getElementById(`preview_${line.id}`);
    if (!previewEl) return;

    if (line.isOnBreak && line.breakStartAt) {
      const sec = Math.floor((Date.now() - line.breakStartAt) / 1000);
      previewEl.textContent = formatHMS(sec);
    } else if (line.isWorking && line.workStartAt) {
      const sec = Math.floor((Date.now() - line.workStartAt) / 1000);
      previewEl.textContent = formatHMS(sec);
    }
  });
}

function enterLine(id) {
  appState.activeLineId = id;
  const line = appState.lines[id];

  if (line.config) {
    showScreen(2);
  } else {
    showScreen(1);
  }
}

// --- SCREEN 1 LOGIC ---
function startWorkDay() {
  const line = getActiveLine();
  if (!line) return;

  line.workerName = el.workerName.value.trim();
  line.machine = el.machine.value;
  line.peopleCount = el.peopleCount.value;
  line.plannedStartTime = el.startTime.value;

  if (!line.workerName || !line.machine || !line.peopleCount || !line.plannedStartTime) {
    alert('Wszystkie pola są wymagane!');
    return;
  }

  // Lock config
  line.config = {
    workerName: line.workerName,
    machine: line.machine,
    peopleCount: line.peopleCount,
    plannedStartTime: line.plannedStartTime
  };

  showScreen(2);
  saveState();
}

// --- SCREEN 2 LOGIC (DETAILS) ---
function selectProduct(product) {
  const line = getActiveLine();
  if (!line) return;

  line.currentProduct = product;
  el.displayCurrentProduct.textContent = product;
  el.productSearch.value = product;
  el.productList.style.display = 'none'; // hide

  updateButtonStates(line);
  saveState();
}

function resetLineSession(line) {
  line.workStartAt = null;
  line.isOnBreak = false;
  line.breakStartAt = null;
  line.totalBreakSeconds = 0;
  line.breaksCurrent = [];
  // Dont clear config
  line.drafts = {}; // clear form
  line.currentProduct = '';
  line.isWorking = false;

  // Clear UI inputs
  if (appState.activeLineId === line.id) {
    el.quantityInput.value = '';
    el.notesInput.value = '';
    el.batchNumberInput.value = '';
    el.cardNumberInput.value = '';
    el.productSearch.value = '';
    restoreInputsFromState();
  }
  stopRealIntervals(line.id);
}

function startLabeling() {
  const line = getActiveLine();
  if (!line) return;
  if (!line.currentProduct) {
    alert('Wybierz produkt!');
    return;
  }

  line.isWorking = true;
  line.workStartAt = Date.now();
  startRealIntervals(line.id);

  updateButtonStates(line);
  saveState();
}

function startBreak() {
  const line = getActiveLine();
  if (!line) return;

  line.isOnBreak = true;
  line.breakStartAt = Date.now();

  // UI Handling for Break Modal
  el.breakReason.value = '';
  el.breakDescription.value = '';
  el.breakModal.classList.add('active');

  startRealIntervals(line.id);
  updateButtonStates(line);
  saveState();
}

function handleBreakReasonChange() {
  const val = el.breakReason.value;
  if (val === 'Awaria') {
    el.descriptionRequired.classList.remove('hidden');
  } else {
    el.descriptionRequired.classList.add('hidden');
  }
}

function confirmBreak() {
  const line = getActiveLine();
  if (!line) return;

  const reason = el.breakReason.value;
  const description = (el.breakDescription.value || '').trim();
  if (!reason) { alert('Wybierz przyczynę!'); return; }
  if (reason === 'Awaria' && !description) { alert('Opis wymagany!'); return; }

  line.breakReasonDraft = reason;
  line.breakDescDraft = description;

  el.breakModal.classList.remove('active');
  updateButtonStates(line); // Will show Resume
  saveState();
}

function cancelBreak() {
  const line = getActiveLine();
  if (!line) return;

  // If we "cancel" break in UI, we effectively cancel entering details, but the timer was running.
  // For simplicity: Undo the break entirely.
  line.isOnBreak = false;
  line.breakStartAt = null;

  el.breakModal.classList.remove('active');
  updateButtonStates(line);
  saveState();
}

function resumeWork() {
  const line = getActiveLine();
  if (!line || !line.isOnBreak || !line.breakStartAt) return;

  const now = Date.now();
  const duration = Math.floor((now - line.breakStartAt) / 1000);
  line.totalBreakSeconds += duration;

  line.breaksCurrent.push({
    startTime: line.breakStartAt,
    endTime: now,
    duration: duration,
    reason: line.breakReasonDraft || 'Inny',
    description: line.breakDescDraft || '',
    product: line.currentProduct
  });

  // NEW: Add to Global Break History immediately
  if (!appState.globalBreaks) appState.globalBreaks = [];
  appState.globalBreaks.push({
    lineId: line.id, // For local filtering
    lineName: line.name,
    product: line.currentProduct,
    sessionIndex: appState.globalSessions.filter(s => s.lineId === line.id).length + 1,
    startTime: line.breakStartAt,
    endTime: now,
    duration: duration,
    reason: line.breakReasonDraft || 'Inny',
    description: line.breakDescDraft || ''
  });

  line.isOnBreak = false;
  line.breakStartAt = null;
  line.breakReasonDraft = null;
  line.breakDescDraft = null;

  startRealIntervals(line.id); // Restart timer interval

  // Delay UI update to prevent "Ghost Click" on Stop button (tablet issue)
  setTimeout(() => {
    updateButtonStates(line);
    updateTimerVisuals(line);
  }, 150);

  // FORCE TABLE UPDATE
  const freshLine = appState.lines[line.id];
  updateCurrentBreaksTable(freshLine);

  saveState();
}

function stopLabeling() {
  const line = getActiveLine();
  if (!line) return;

  // Validate Inputs
  if (!validateForm(true)) {
    alert('Uzupełnij wymagane pola!');
    return;
  }

  // Close Break if active
  if (line.isOnBreak && line.breakStartAt) {
    resumeWork();
  }

  const now = Date.now();
  const session = {
    id: Date.now() + '_' + Math.random(),
    lineId: line.id,
    lineName: line.name,
    workerName: line.config.workerName,
    machine: line.config.machine,
    peopleCount: parseInt(line.peopleCount),
    product: line.currentProduct,

    startTime: line.workStartAt,
    endTime: now,
    workDuration: Math.max(1, Math.floor((now - line.workStartAt) / 1000 / 60)),
    breakDuration: Math.ceil(line.totalBreakSeconds / 60),

    quantity: parseInt(el.quantityInput.value),
    labelType: el.labelTypeSelect.value,
    batchNumber: el.batchNumberInput.value,
    productionDate: el.productionDateInput.value,
    cardNumber: el.cardNumberInput.value,
    notes: el.notesInput ? el.notesInput.value : '',
    packagingType: document.querySelector('input[name="packagingType"]:checked')?.value || 'Folia',

    breaks: [...line.breaksCurrent]
  };

  // Add to Global History
  appState.globalSessions.push(session);

  // Auto-send single session
  sendToGoogleSheets(session);

  // Reset Line
  resetLineSession(line);
  updateLocalHistoryTables(line);
  updateTotalLabeled();

  alert('Sesja zapisana i wysłana (w tle)!');
  //showScreen(0);
}

// --- TIMERS ---
function startRealIntervals(lineId) {
  stopRealIntervals(lineId);

  // We only update UI if this line is active
  if (appState.activeLineId !== lineId) return;

  const line = appState.lines[lineId];

  intervalRegistry[lineId] = setInterval(() => {
    // Only update UI if we are still looking at this line
    if (appState.activeLineId !== lineId) return;

    updateTimerVisuals(line);
  }, 1000);
}

function stopRealIntervals(lineId) {
  if (intervalRegistry[lineId]) {
    clearInterval(intervalRegistry[lineId]);
    delete intervalRegistry[lineId];
  }
}

function updateTimerVisuals(line) {
  if (!el.workTimerContainer) return;

  if (line.isOnBreak) {
    el.workTimerContainer.classList.add('hidden');
    el.breakTimerContainer.classList.remove('hidden');
    if (line.breakStartAt) {
      const s = Math.floor((Date.now() - line.breakStartAt) / 1000);
      el.breakTimer.textContent = formatHMS(s);
    }
  } else {
    el.breakTimerContainer.classList.add('hidden');
    el.workTimerContainer.classList.remove('hidden');
    if (line.workStartAt) {
      const s = Math.floor((Date.now() - line.workStartAt) / 1000);
      el.workTimer.textContent = formatHMS(s);
    } else {
      el.workTimer.textContent = '00:00:00';
    }
  }
}

function updateButtonStates(line) {
  if (line.isOnBreak) {
    el.startBtn.disabled = true;
    el.breakBtn.disabled = true;
    el.stopBtn.disabled = true;
    el.resumeBtn.classList.remove('hidden');
  } else if (line.isWorking) {
    el.startBtn.disabled = true;
    el.breakBtn.disabled = false;
    el.stopBtn.disabled = false;
    el.resumeBtn.classList.add('hidden');
  } else if (line.currentProduct) {
    el.startBtn.disabled = false;
    el.breakBtn.disabled = true;
    el.stopBtn.disabled = true;
    el.resumeBtn.classList.add('hidden');
  } else {
    el.startBtn.disabled = true;
    el.breakBtn.disabled = true;
    el.stopBtn.disabled = true;
    el.resumeBtn.classList.add('hidden');
  }
}

// --- LOCAL HISTORY ---
function updateLocalHistoryTables(line) {
  // Filter globalSessions where lineId == line.id
  const lineSessions = appState.globalSessions.filter(s => s.lineId === line.id);
  const tbody = el.labelingHistoryBody;
  tbody.innerHTML = '';

  if (lineSessions.length === 0) {
    tbody.innerHTML = '<tr class="empty-table"><td colspan="15">Brak sesji</td></tr>';
  } else {
    // Sort descending
    [...lineSessions].reverse().forEach((s, idx) => {
      const tr = document.createElement('tr');
      let breakDisplay = s.breakDuration > 0 ? `${s.breakDuration} min` : 'Nie';
      tr.innerHTML = `
                <td>${lineSessions.length - idx}</td> <!-- Real number in list -->
                <td>${s.product}</td>
                <td>${s.batchNumber}</td>
                <td>${s.productionDate}</td>
                <td>${s.peopleCount}</td>
                <td>${s.cardNumber}</td>
                <td>${s.labelType}</td>
                <td>${s.packagingType}</td>
                <td>${formatTime(s.startTime)}</td>
                <td>${formatTime(s.endTime)}</td>
                <td>${s.workDuration}</td>
                <td>${s.quantity}</td>
                <td>${s.notes}</td>
                <td>${breakDisplay}</td>
                <td>${getNormStatusHtml(s)}</td>
              `;
      tbody.appendChild(tr);
    });
  }
  updateTotalLabeled(lineSessions);
}

function updateCurrentBreaksTable(line) {
  const tbody = document.getElementById('breakHistoryBody');
  if (!tbody) {
    console.warn('breakHistoryBody not found');
    return;
  }

  tbody.innerHTML = '';

  // Use GLOBAL HISTORY for this line, instead of transient breaksCurrent
  // This ensures the table persists after Stop (resetLineSession)
  const lineBreaks = (appState.globalBreaks || []).filter(b => b.lineId === line.id);

  if (lineBreaks.length === 0) {
    tbody.innerHTML = '<tr class="empty-table"><td colspan="7">Brak przerw</td></tr>';
    return;
  }

  // Sort descending (newest on top) or ascending? User didn't specify, but history usually desc.
  // Original was breaksCurrent (chrono). Let's keep chronological for "Session History" feel, or Desc?
  // Let's do Reverse Chronological (Newest First) so they see what they just added.
  [...lineBreaks].reverse().forEach((b, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
            <td>${lineBreaks.length - idx}</td>
            <td>${b.sessionIndex || '-'}</td>
            <td>${formatTime(b.startTime)}</td>
            <td>${formatTime(b.endTime)}</td>
            <td>${Math.ceil(b.duration / 60)}</td>
            <td>${b.reason} ${b.description ? '(' + b.description + ')' : ''}</td>
            <td>${b.product || '-'}</td>
         `;
    tbody.appendChild(tr);
  });
}

function updateTotalLabeled(sessions = []) {
  if (!sessions.length && appState.activeLineId) {
    sessions = appState.globalSessions.filter(s => s.lineId === appState.activeLineId);
  }
  const total = sessions.reduce((sum, s) => sum + (parseInt(s.quantity) || 0), 0);
  el.totalLabeledDisplay.textContent = `${total} szt.`;
}

// --- GLOBAL SUMMARY (Screen 3) ---
function generateGlobalSummary() {
  const tbody = el.summaryLabelingBody;
  tbody.innerHTML = '';

  const sessions = appState.globalSessions;
  if (sessions.length === 0) {
    tbody.innerHTML = '<tr class="empty-table"><td colspan="16">Brak danych</td></tr>';
    return;
  }

  // Sort by time descending
  [...sessions].sort((a, b) => b.endTime - a.endTime).forEach((s, idx) => {
    const tr = document.createElement('tr');
    // Determine sync icon
    // Green check or Cloud if synced, Warning/Red Dot if not
    const syncStatus = s.synced ? '<span style="color:green" title="Wysłano">☁️</span>' : '<span style="color:orange" title="Nie wysłano">⚠️</span>';

    let breakDisplay = s.breakDuration > 0 ? `${s.breakDuration} min` : 'Nie';

    tr.innerHTML = `
                <td>
                  ${sessions.length - idx} <br>
                  ${syncStatus}
                </td>
                <td>${s.lineName}</td> 
                <td>${s.product}</td>
                <td>${s.batchNumber}</td>
                <td>${s.productionDate}</td>
                <td>${s.peopleCount}</td>
                <td>${s.cardNumber}</td>
                <td>${s.labelType}</td>
                <td>${s.packagingType}</td>
                <td>${formatTime(s.startTime)}</td>
                <td>${formatTime(s.endTime)}</td>
                <td>${s.workDuration}</td>
                <td>${s.quantity}</td>
                <td>${s.notes}</td>
                <td>${breakDisplay}</td>
                <td>${getNormStatusHtml(s)}</td>
          `;
    tbody.appendChild(tr);
  });

  // Update stats
  el.totalSessions.innerText = sessions.length;
  el.totalPieces.innerText = sessions.reduce((s, x) => s + (x.quantity || 0), 0);
  el.totalWorkMinutes.innerText = sessions.reduce((s, x) => s + (x.workDuration || 0), 0);
  el.totalBreakMinutes.innerText = sessions.reduce((s, x) => s + (x.breakDuration || 0), 0);

  // Update Backup Button Logic
  const unsyncedCount = sessions.filter(s => !s.synced).length;
  const btn = document.querySelector('.btn-primary[onclick="sendToGoogleSheets()"]');
  if (btn) {
    if (unsyncedCount > 0) {
      btn.innerText = `Wyślij niezatwierdzone (${unsyncedCount})`;
      btn.classList.add('pulse-anim'); // Optional: add attention animation class?
      btn.style.backgroundColor = '#e67e22'; // Orange warning color
    } else {
      btn.innerText = 'Wszystko wysłane ✓';
      btn.style.backgroundColor = '#27ae60'; // Green ok color
    }
  }

  // Render Break History
  const breakTbody = el.summaryBreakHistoryBody || document.getElementById('summaryBreakHistoryBody');
  if (breakTbody) {
    breakTbody.innerHTML = '';
    const breaks = appState.globalBreaks || [];
    if (breaks.length === 0) {
      breakTbody.innerHTML = '<tr class="empty-table"><td colspan="9">Brak przerw</td></tr>';
    } else {
      [...breaks].sort((a, b) => b.endTime - a.endTime).forEach((b, idx) => {
        const tr = document.createElement('tr');
        const min = Math.ceil(b.duration / 60);
        tr.innerHTML = `
                 <td>${breaks.length - idx}</td>
                 <td>${b.lineName || '-'}</td>
                 <td>${b.product || '-'}</td>
                 <td>${b.sessionIndex || '-'}</td>
                 <td>${b.reason}</td>
                 <td>${b.startTime ? formatTime(b.startTime) : '-'}</td>
                 <td>${b.endTime ? formatTime(b.endTime) : '-'}</td>
                 <td>${min}</td>
                 <td>${b.description || '-'}</td>
              `;
        breakTbody.appendChild(tr);
      });
    }
  }

  // Render People Change History (Global)
  const peopleTbody = el.summaryPeopleChangeBody || document.getElementById('summaryPeopleChangeBody');
  if (peopleTbody) {
    peopleTbody.innerHTML = '';
    const peopleChanges = [];

    // Aggregate all changes from all lines
    Object.values(appState.lines).forEach(line => {
      if (line.peopleCountHistory && line.peopleCountHistory.length > 0) {
        line.peopleCountHistory.forEach(change => {
          peopleChanges.push({
            ...change,
            lineName: line.name
          });
        });
      }
    });

    if (peopleChanges.length === 0) {
      peopleTbody.innerHTML = '<tr class="empty-table"><td colspan="6">Brak zmian</td></tr>';
    } else {
      // Sort by date desc
      peopleChanges.sort((a, b) => new Date(b.date) - new Date(a.date));
      peopleChanges.forEach((change, idx) => {
        const tr = document.createElement('tr');
        // If we stored sessionIndex in the change objects, display it
        // Or if not, we can assume it triggered a new session, but maybe '-' for now if not tracked.
        const sessId = change.sessionIndex ? change.sessionIndex : '-';

        tr.innerHTML = `
          <td>${peopleChanges.length - idx}</td>
          <td>${change.lineName || '-'}</td>
          <td>${sessId}</td>
          <td>${change.from}</td>
          <td>${change.to}</td>
          <td>${new Date(change.date).toLocaleString('pl-PL')}</td>
        `;
        peopleTbody.appendChild(tr);
      });
    }
  }
}

function confirmStartNewDay() {
  if (confirm('Czy na pewno chcesz wyczyścić WSZYSTKIE dane (wszystkie linie i historie)?')) {
    localStorage.removeItem(STATE_KEY);
    // Reload
    window.location.reload();
  }
}

// === SEND TO GOOGLE SHEETS ===
async function sendToGoogleSheets(singleSession = null) {
  // 1. Get Web App URL
  let webAppUrl = GOOGLE_WEB_APP_URL;
  if (!webAppUrl || webAppUrl.trim() === '') {
    webAppUrl = localStorage.getItem('GOOGLE_WEB_APP_URL_OVERRIDE');
  }

  if (!webAppUrl) {
    const userUrl = prompt('Nie skonfigurowano URL Google Apps Script! Podaj go teraz:');
    if (userUrl) {
      localStorage.setItem('GOOGLE_WEB_APP_URL_OVERRIDE', userUrl);
      webAppUrl = userUrl;
    } else {
      alert('Brak URL - nie można wysłać danych.');
      return;
    }
  }

  // UI Feedback only if bulk send
  const btn = document.querySelector('.btn-primary[onclick="sendToGoogleSheets()"]');
  let originalText = btn ? btn.innerText : '';
  if (!singleSession && btn) {
    btn.disabled = true;
    btn.innerText = 'Wysyłanie...';
  }

  try {
    // If single session provided, try to send it.
    // If NO single session, send ALL that are NOT synced yet.

    let sessionsToSend = [];
    if (singleSession) {
      sessionsToSend = [singleSession];
    } else {
      sessionsToSend = appState.globalSessions.filter(s => !s.synced);
    }

    if (sessionsToSend.length === 0) {
      if (!singleSession) alert('Wszystkie sesje zostały już wysłane!');
      if (btn) {
        btn.innerText = 'Wszystko wysłane ✓';
        btn.disabled = false;
      }
      return;
    }

    const preparedSessions = sessionsToSend.map(s => {
      // Calculate derived stats for Sheet
      let netWorkDuration = s.workDuration - (s.breakDuration || 0);
      if (netWorkDuration <= 0) netWorkDuration = 1;

      // Ensure we have peopleCount for norm calc
      const pCount = s.peopleCount || 1;

      // Try to calculate norm target based on normsMap
      const targetNorm = calculateNormTarget(s.product, s.packagingType, s.machine, pCount);

      let normPercent = '';
      if (targetNorm > 0) {
        // Speed per hour
        const realSpeed = (s.quantity / (netWorkDuration / 60));
        normPercent = Math.round((realSpeed / targetNorm) * 100) + '%';
      }

      return {
        ...s,
        peopleCount: pCount,
        normPercent: normPercent
      };
    });

    await fetch(webAppUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessions: preparedSessions })
    });

    // Mark as synced
    sessionsToSend.forEach(s => s.synced = true);
    saveState(); // Save the fact that we synced them

    if (!singleSession) {
      alert(`Wysłano ${sessionsToSend.length} nowych sesji!`);
      if (btn) btn.innerText = 'Wysłano ✓';
    } else {
      console.log('Single session synced to Sheet', singleSession.id);
    }

  } catch (error) {
    console.error('Send error:', error);
    if (!singleSession) alert('Błąd wysyłania: ' + error.message);

    if (!singleSession && btn) {
      btn.innerText = originalText;
      btn.disabled = false;
    }
  } finally {
    if (!singleSession) {
      setTimeout(() => {
        if (btn && btn.disabled && btn.innerText.includes('Wysłano')) {
          btn.disabled = false;
          btn.innerText = originalText;
        }
      }, 3000);
    }
  }
}

// === CLOCK & DATE ===
function updateClock() {
  const now = new Date();
  const dateStr = now.toLocaleDateString('pl-PL', {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  });
  const timeStr = now.toLocaleTimeString('pl-PL', {
    hour: '2-digit',
    minute: '2-digit'
  });

  const elements = document.querySelectorAll('.headerDateTime');
  elements.forEach(el => {
    el.innerHTML = `<span style="font-size: 1.6em; font-weight: bold;">${timeStr}</span><br><span style="font-size: 1.2em;">${dateStr}</span>`;
  });
}
setInterval(updateClock, 1000);
updateClock();

// === DATA LOADING HELPERS ===
// (Preserved from legacy script)
let allProducts = [];
let productMap = {};
let normsMap = {};

function getLineType(machineName) {
  if (!machineName) return 'MANUAL';
  const lower = machineName.toLowerCase();
  if (lower.includes('kapturkownica') || lower.includes('kapturkownicą')) {
    return 'AUTOMATIC';
  }
  return 'MANUAL';
}

function calculateNormTarget(product, packaging, machine, people) {
  if (!product || !packaging || !machine || !people) return 0;
  const jarType = productMap[product];
  if (!jarType) return 0;
  const normsForJar = normsMap[jarType];
  if (!normsForJar) return 0;
  const lineType = getLineType(machine);
  const configKey = `${lineType}_${packaging}`;
  const normsForConfig = normsForJar[configKey];
  if (!normsForConfig) return 0;
  const peopleInt = parseInt(people);
  if (isNaN(peopleInt) || peopleInt < 1 || peopleInt > 6) return 0;
  return normsForConfig[peopleInt] || 0;
}

function getNormStatusHtml(session) {
  if (!session.quantity || !session.workDuration || session.workDuration <= 0) return '-';
  let netWorkDuration = session.workDuration - (session.breakDuration || 0);
  if (netWorkDuration <= 0) netWorkDuration = 1;
  const realSpeed = (session.quantity / netWorkDuration) * 60;
  const people = session.peopleCount || 1;
  const machineUsed = session.machine;
  const targetNorm = calculateNormTarget(session.product, session.packagingType, machineUsed, people);
  if (targetNorm === 0) return '<span title="Brak normy">-</span>';
  const icon = realSpeed >= targetNorm ? '✅' : '❌';
  const tooltip = `Aktualne: ${Math.round(realSpeed)} szt/h / Norma: ${Math.round(targetNorm)} szt/h`;
  return `<span title="${tooltip}" style="cursor: help;">${icon}</span>`;
}

function validateForm(showErrors = true) {
  const qty = parseInt(el.quantityInput.value);
  const labelType = el.labelTypeSelect.value;
  const batchNumber = el.batchNumberInput.value.trim();
  const productionDate = el.productionDateInput.value;
  const cardNumber = el.cardNumberInput.value.trim();

  let hasError = false;
  if (!qty || qty <= 0) {
    if (showErrors && el.quantityValidation) el.quantityValidation.classList.add('show');
    hasError = true;
  } else {
    if (el.quantityValidation) el.quantityValidation.classList.remove('show');
  }
  if (!labelType) {
    if (showErrors && el.labelTypeValidation) el.labelTypeValidation.classList.add('show');
    hasError = true;
  } else {
    if (el.labelTypeValidation) el.labelTypeValidation.classList.remove('show');
  }
  if (!batchNumber) {
    if (showErrors && el.batchValidation) el.batchValidation.classList.add('show');
    hasError = true;
  } else {
    if (el.batchValidation) el.batchValidation.classList.remove('show');
  }
  if (!productionDate) {
    if (showErrors && el.prodDateValidation) el.prodDateValidation.classList.add('show');
    hasError = true;
  } else {
    if (el.prodDateValidation) el.prodDateValidation.classList.remove('show');
  }
  if (!cardNumber) {
    if (showErrors && el.cardValidation) el.cardValidation.classList.add('show');
    hasError = true;
  } else {
    if (el.cardValidation) el.cardValidation.classList.remove('show');
  }
  if (!cardNumber) {
    if (showErrors && el.cardValidation) el.cardValidation.classList.add('show');
    hasError = true;
  } else {
    if (el.cardValidation) el.cardValidation.classList.remove('show');
  }
  return !hasError;
}

// --- PEOPLE CHANGE LOGIC ---
function openChangePeopleModal() {
  const line = getActiveLine();
  if (!line) return;
  // Set current value
  el.newPeopleCount.value = line.peopleCount;
  el.changePeopleModal.classList.add('active');
  el.changePeopleModal.style.display = 'flex';
}

function cancelChangePeople() {
  el.changePeopleModal.classList.remove('active');
  el.changePeopleModal.style.display = 'none';
}

function confirmChangePeople() {
  const line = getActiveLine();
  if (!line) return;

  const oldVal = line.peopleCount;
  const newVal = el.newPeopleCount.value;

  if (oldVal === newVal) {
    cancelChangePeople();
    return;
  }

  // 1. If timer hasn't started yet, allow simple update without closing session
  if (!line.isWorking && !line.workStartAt) {
    line.peopleCount = newVal;
    el.displayPeopleCount.textContent = newVal;
    if (line.config) line.config.peopleCount = newVal;

    cancelChangePeople();
    saveState();
    return;
  }

  // 2. Validate Form (Must be valid to close current session)
  if (!validateForm(true)) {
    alert('Aby zmienić ilość osób, musisz uzupełnić dane bieżącej sesji (ilość, partia, data itp.), ponieważ zostanie ona zamknięta.');
    return;
  }

  // 2. Close current session (Logic borrowed from stopLabeling)
  const now = Date.now();
  let workDuration = 0;
  if (line.workStartAt) {
    workDuration = Math.floor((now - line.workStartAt) / 1000);
  }
  // Subtract breaks
  if (line.isOnBreak && line.breakStartAt) {
    // If currently on break, we count it as break time until now
    const breakSeg = Math.floor((now - line.breakStartAt) / 1000);
    line.totalBreakSeconds += breakSeg;
    // Add to local breaks list for this session
    line.breaksCurrent.push({
      startTime: line.breakStartAt,
      endTime: now,
      duration: breakSeg,
      reason: line.breakReasonDraft || 'Zmiana osób',
      description: line.breakDescDraft || ''
    });
    // Global Break History too? Yes, ideally close it properly.
    if (!appState.globalBreaks) appState.globalBreaks = [];
    appState.globalBreaks.push({
      lineName: line.name,
      product: line.currentProduct,
      sessionIndex: appState.globalSessions.filter(s => s.lineId === line.id).length + 1,
      startTime: line.breakStartAt,
      endTime: now,
      duration: breakSeg,
      reason: line.breakReasonDraft || 'Zmiana osób',
      description: line.breakDescDraft || ''
    });
  }

  workDuration -= line.totalBreakSeconds;
  if (workDuration < 0) workDuration = 0; // Safety

  const session = {
    id: generateId(),
    lineId: line.id,
    lineName: line.name,
    workerName: line.config.workerName,
    machine: line.config.machine,
    peopleCount: parseInt(line.peopleCount),
    product: line.currentProduct,
    batchNumber: el.batchNumberInput.value,
    productionDate: el.productionDateInput.value,
    cardNumber: el.cardNumberInput.value,
    labelType: el.labelTypeSelect.value,
    packagingType: document.querySelector('input[name="packagingType"]:checked')?.value || 'Folia',
    notes: el.notesInput.value,
    startTime: line.workStartAt ? line.workStartAt : now,
    endTime: now,
    workDuration: Math.ceil(workDuration / 60), // min
    breakDuration: Math.ceil(line.totalBreakSeconds / 60), // min
    quantity: parseInt(el.quantityInput.value) || 0,
    breaks: [...line.breaksCurrent] // Copy
  };

  appState.globalSessions.push(session);
  // Send to Google Sheets (Async)
  sendToGoogleSheets(session);

  // 3. Update Line State (People Count)
  line.peopleCount = newVal;
  // Get index of the session we just closed (it's the last one for this line)
  const justClosedSessionIndex = appState.globalSessions.filter(s => s.lineId === line.id).length;

  // Add to History
  if (!line.peopleCountHistory) line.peopleCountHistory = [];
  line.peopleCountHistory.push({
    from: oldVal,
    to: newVal,
    date: new Date().toISOString(),
    sessionIndex: justClosedSessionIndex // The session that ended with this change
  });

  // 4. Start New Session (Logic borrowed from startLabeling)
  // Reset session-specific counters, but KEEP config (product, worker, etc.)
  line.workStartAt = now;
  line.isOnBreak = false;
  line.breakStartAt = null;
  line.totalBreakSeconds = 0;
  line.breaksCurrent = [];

  // Clear Quantity/Notes for new session? Usually yes.
  // But maybe keep Batch/Card/Date? 
  // User logic: "started a new session with the new number of people".
  // Let's clear Quantity but keep valid fields.
  el.quantityInput.value = '';
  // el.notesInput.value = ''; // Optional: clear notes?

  // Update UI
  el.displayPeopleCount.textContent = newVal;
  updatePeopleChangeTable(line);
  updateLocalHistoryTables(line);
  updateCurrentBreaksTable(line);

  updateTimerVisuals(line);
  updateButtonStates(line);

  cancelChangePeople();
  saveState();
}

function updatePeopleChangeTable(line) {
  const tbody = el.peopleChangeHistoryBody || document.getElementById('peopleChangeHistoryBody');
  if (!tbody) return;

  tbody.innerHTML = '';
  const history = line.peopleCountHistory || [];

  if (history.length === 0) {
    tbody.innerHTML = '<tr class="empty-table"><td colspan="4">Brak zmian</td></tr>';
    return;
  }

  // Sort descending by date
  [...history].reverse().forEach((change, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${history.length - idx}</td>
      <td>${change.from}</td>
      <td>${change.to}</td>
      <td>${new Date(change.date).toLocaleString('pl-PL')}</td>
    `;
    tbody.appendChild(tr);
  });
}

// --- SHEETS LOADERS ---
async function loadMachinesFromSheet() {
  const SHEET_ID = '1JTyATb7I4fMj7rywuByjGZmtMXNDP0uc010i3ZBoq0s';
  const GID = '542333307';
  const COLUMN_INDEX = 3;
  const SKIP_HEADER = true;
  const url = `https://corsproxy.io/?${encodeURIComponent(
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${GID}`
  )}`;

  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const lines = text.trim().split('\n');

    const machineSelect = document.getElementById('machine');
    machineSelect.innerHTML = '<option value="">-- Wybierz maszynę --</option>';
    const uniqueMachines = new Set();

    lines.forEach((line, index) => {
      if (SKIP_HEADER && index === 0) return;
      const columns = line.split(',');
      const machineName = columns[COLUMN_INDEX]?.replace(/"/g, '').trim();
      if (machineName && machineName !== '') {
        uniqueMachines.add(machineName);
      }
    });

    uniqueMachines.forEach((machine) => {
      const option = document.createElement('option');
      option.value = machine;
      option.textContent = machine;
      machineSelect.appendChild(option);
    });
  } catch (error) {
    console.error('❌ Błąd ładowania maszyn:', error);
    document.getElementById('machine').innerHTML = '<option value="">Błąd ładowania</option>';
  }
}

async function loadProductsFromSheet() {
  const SHEET_ID = '1JTyATb7I4fMj7rywuByjGZmtMXNDP0uc010i3ZBoq0s';
  const GID = '542333307';
  const SKIP_HEADER = true;
  const url = `https://corsproxy.io/?${encodeURIComponent(
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${GID}`
  )}`;

  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const lines = text.trim().split('\n');

    allProducts = [];
    productMap = {};

    lines.forEach((line, index) => {
      if (SKIP_HEADER && index === 0) return;
      const columns = line.split(',');
      const productName = columns[0]?.replace(/"/g, '').trim();
      const jarType = columns[1]?.replace(/"/g, '').trim();
      if (productName) {
        allProducts.push(productName);
        if (jarType) {
          productMap[productName] = jarType;
        }
      }
    });
    allProducts.sort();
  } catch (error) {
    console.error('❌ Błąd ładowania produktów:', error);
  }
}

async function loadNormsFromSheet() {
  const SHEET_ID = '1JTyATb7I4fMj7rywuByjGZmtMXNDP0uc010i3ZBoq0s';
  const GID = '520596027';
  const SKIP_HEADER = true;
  const url = `https://corsproxy.io/?${encodeURIComponent(
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${GID}`
  )}`;

  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const lines = text.trim().split('\n');
    normsMap = {};

    lines.forEach((line, index) => {
      if (SKIP_HEADER && index === 0) return;
      const cols = line.split(',');
      const cleanCols = cols.map(c => c ? c.replace(/"/g, '').trim() : '');
      const jarType = cleanCols[0];
      const lineType = cleanCols[1]?.toUpperCase();
      const packaging = cleanCols[2];

      if (jarType && lineType && packaging) {
        if (!normsMap[jarType]) normsMap[jarType] = {};
        const configKey = `${lineType}_${packaging}`;
        const norms = [
          0,
          parseInt(cleanCols[3]) || 0, // 1 os
          parseInt(cleanCols[4]) || 0, // 2 os
          parseInt(cleanCols[5]) || 0, // 3 os
          parseInt(cleanCols[6]) || 0, // 4 os
          parseInt(cleanCols[7]) || 0, // 5 os
          parseInt(cleanCols[8]) || 0  // 6 os
        ];
        normsMap[jarType][configKey] = norms;
      }
    });
  } catch (e) {
    console.error('❌ Błąd ładowania norm:', e);
  }
}

async function loadLabelTypesFromSheet() {
  const SHEET_ID = '1JTyATb7I4fMj7rywuByjGZmtMXNDP0uc010i3ZBoq0s';
  const GID = '542333307';
  const COLUMN_INDEX = 5; // 5 = kolumna F (Rodzaj etykiety)
  const SKIP_HEADER = true;
  const url = `https://corsproxy.io/?${encodeURIComponent(
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${GID}`
  )}`;
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const lines = text.trim().split('\n');
    const select = document.getElementById('labelTypeSelect');
    select.innerHTML = '<option value="">-- Wybierz rodzaj --</option>';
    const unique = new Set();
    lines.forEach((line, index) => {
      if (SKIP_HEADER && index === 0) return;
      const columns = line.split(',');
      const val = columns[COLUMN_INDEX]?.replace(/"/g, '').trim();
      if (val) unique.add(val);
    });
    unique.forEach(val => {
      const opt = document.createElement('option');
      opt.value = val;
      opt.textContent = val;
      select.appendChild(opt);
    });
  } catch (err) {
    console.error('Błąd rodzajów etykiet', err);
  }
}

function filterProducts() {
  const searchInput = document.getElementById('productSearch');
  const productList = document.getElementById('productList');
  const searchText = searchInput.value.toLowerCase().trim();
  if (searchText.length === 0) {
    productList.style.display = 'none';
    return;
  }
  const searchTerms = searchText.split(' ').filter(term => term.length > 0);
  const filtered = allProducts.filter((product) => {
    const pLower = product.toLowerCase();
    return searchTerms.every(term => pLower.includes(term));
  });
  populateProductList(filtered);
  productList.style.display = 'block';
}

function populateProductList(products) {
  const productList = document.getElementById('productList');
  productList.innerHTML = '';
  if (products.length === 0) {
    const item = document.createElement('div');
    item.className = 'product-item';
    item.textContent = 'Brak wyników...';
    item.style.color = '#7f8c8d';
    item.style.cursor = 'default';
    productList.appendChild(item);
    return;
  }
  products.forEach((product) => {
    const item = document.createElement('div');
    item.className = 'product-item';
    item.textContent = product;
    item.onclick = () => selectProduct(product);
    productList.appendChild(item);
  });
}

function showProductDropdown() {
  const searchInput = document.getElementById('productSearch');
  if (allProducts.length > 0 && searchInput.value.trim().length > 0) {
    filterProducts();
  }
}

// === TIME PICKER ===
let selectedHour = '06';
let selectedMinute = '00';

function openTimePicker() {
  try {
    const modal = document.getElementById('timePickerModal');
    if (!modal) return;
    const currentVal = document.getElementById('startTime').value;
    if (currentVal) {
      const [h, m] = currentVal.split(':');
      selectedHour = h;
      selectedMinute = m;
    } else {
      const now = new Date();
      selectedHour = String(now.getHours()).padStart(2, '0');
      selectedMinute = String(now.getMinutes()).padStart(2, '0');
    }
    renderTimeColumns();
    modal.classList.add('active');
    modal.style.display = 'flex';
    setTimeout(() => {
      scrollToSelection('hourColumn', selectedHour);
      scrollToSelection('minuteColumn', selectedMinute);
    }, 10);
  } catch (e) { console.error(e); }
}

function closeTimePicker() {
  const modal = document.getElementById('timePickerModal');
  if (modal) {
    modal.classList.remove('active');
    modal.style.display = 'none';
  }
}

function confirmTime() {
  const startTimeInput = document.getElementById('startTime');
  const startTimeDisplay = document.getElementById('startTimeDisplay');
  const timeStr = `${selectedHour}:${selectedMinute}`;
  startTimeInput.value = timeStr;
  startTimeDisplay.value = timeStr;

  // Sync to state
  syncInputsToState();
  closeTimePicker();
}

function renderTimeColumns() {
  const hourCol = document.getElementById('hourColumn');
  const minCol = document.getElementById('minuteColumn');
  hourCol.innerHTML = '';
  minCol.innerHTML = '';

  for (let i = 0; i < 24; i++) {
    const val = String(i).padStart(2, '0');
    const el = document.createElement('div');
    el.className = 'picker-item';
    if (val === selectedHour) el.classList.add('selected');
    el.textContent = val;
    el.onclick = () => {
      selectedHour = val;
      renderTimeColumns();
    };
    hourCol.appendChild(el);
  }

  for (let i = 0; i < 60; i++) {
    const val = String(i).padStart(2, '0');
    const el = document.createElement('div');
    el.className = 'picker-item';
    if (val === selectedMinute) el.classList.add('selected');
    el.textContent = val;
    el.onclick = () => {
      selectedMinute = val;
      renderTimeColumns();
    };
    minCol.appendChild(el);
  }
}

function scrollToSelection(colId, val) {
  const col = document.getElementById(colId);
  const selected = col.querySelector('.selected');
  if (selected) {
    selected.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }
}

// Click outside close
window.onclick = function (event) {
  const modal = document.getElementById('timePickerModal');
  if (event.target == modal) {
    closeTimePicker();
  }
  if (!event.target.closest('#productSearch') && !event.target.closest('#productList')) {
    const pl = document.getElementById('productList');
    if (pl) pl.style.display = 'none';
  }
}

// --- INIT ---
document.addEventListener('DOMContentLoaded', () => {
  loadState();
  if (Object.keys(appState.lines).length === 0) {
    addNewLine();
  }

  // Resources
  loadMachinesFromSheet();
  setTimeout(loadProductsFromSheet, 500);
  setTimeout(loadNormsFromSheet, 500);
  setTimeout(loadLabelTypesFromSheet, 1000);
});
