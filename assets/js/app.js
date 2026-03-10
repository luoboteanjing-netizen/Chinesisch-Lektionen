/* Updated app.js: Dynamic sheet loading from all available L* sheets, names from A2 (cut first 2 chars) */

// Constants (adapted from original)
const EXCEL_URL = './data/Long-Chinesisch_Lektionen.xlsx';
const SHEET_NAME_PATTERN = /^L/i; // Flexible: any starting with L
const DATA_START_ROW = 3; // Row 3+ for data
const LS_KEYS = { settings: 'fc_settings_v1', progress: 'fc_progress_v1' };

// Column indices (matching original: data starts at B=1)
const COL_WORD = { de: 1, py: 2, zh: 6 };
const COL_SENT = { de: 5, py: 4, zh: 7 };
const COL_POS = 3;

// State (adapted from original, with lessons as Map<index, {name, cards}> for dynamic)
const state = {
  mode: 'de2zh', order: 'random',
  rateDe: 0.95, pitchDe: 1.0, rateZh: 0.95, pitchZh: 1.0,
  lessons: new Map(), // <lessonIndex, {name: string, cards: []}>
  selectedLessons: new Set(),
  pool: [], idx: null, current: null,
  voices: [], browserVoice: { zh: null, de: null }, voicePanelTarget: 'de',
  autoplay: { on: false, timers: [], gapMs: 800 },
  settings: { mode: 'de2zh', order: 'random', rateDe: 0.95, pitchDe: 1.0, rateZh: 0.95, pitchZh: 1.0, lessons: [], browserVoiceZh: null, browserVoiceDe: null, autoplayGap: 800 },
  session: { total: 0, done: 0, known: 0, unsure: 0, unknown: 0, ttrSum: 0, ttrCount: 0 },
  startedAt: null, revealedAt: null,
  progress: { version: 'v1', cards: {}, byLesson: {} },
  wakeLock: null,
  trainingOn: false
};

// DOM shortcuts (add if needed; assuming original has these IDs)
const $ = s => document.querySelector(s);
const els = {
  lessonSelect: $('#lesson-select'),
  // Add others as in original: modeSelect, etc.
  // For full, assume original event listeners are set
};

// Save/Load (unchanged from original)
function saveSettings() { try { localStorage.setItem(LS_KEYS.settings, JSON.stringify(state.settings)); } catch (e) {} }
function loadSettings() { try { const s = JSON.parse(localStorage.getItem(LS_KEYS.settings) || 'null'); if (s) { state.settings = Object.assign(state.settings, s); } } catch (e) {} }
function saveProgress() { try { localStorage.setItem(LS_KEYS.progress, JSON.stringify(state.progress)); } catch (e) {} }
function loadProgress() { try { const p = JSON.parse(localStorage.getItem(LS_KEYS.progress) || 'null'); if (p && p.version === 'v1') { state.progress = p; } } catch (e) {} }

// Voice functions (unchanged from original)
function isZhVoice(v) { const L = (v.lang || '').toLowerCase(); return L.startsWith('zh') || L.includes('cmn') || L.includes('hans') || L.includes('zh-cn'); }
function isDeVoice(v) { const L = (v.lang || '').toLowerCase(); return L.startsWith('de'); }

function updateVoiceList() {
  const box = $('#dbgVoices');
  if (!box) return;
  box.innerHTML = '';
  const list = (state.voices || []).filter(v => state.voicePanelTarget === 'zh' ? isZhVoice(v) : isDeVoice(v));
  if (list.length === 0) {
    box.innerHTML = '<p>Keine passenden Stimmen gefunden.</p>';
    return;
  }
  list.forEach(v => {
    const row = document.createElement('div');
    row.className = 'voice';
    const name = document.createElement('div');
    name.className = 'name';
    name.textContent = v.name || '(name)';
    const meta = document.createElement('div');
    meta.className = 'meta';
    meta.textContent = `${v.lang || ''} ${v.default ? '· default' : ''}`;
    const actions = document.createElement('div');
    actions.style.marginLeft = 'auto';
    actions.style.display = 'flex';
    actions.style.gap = '6px';
    actions.style.flexWrap = 'wrap';
    const pick = document.createElement('button');
    pick.className = 'btn';
    pick.textContent = 'Diese Stimme wählen';
    pick.onclick = () => {
      if (state.voicePanelTarget === 'zh') {
        state.browserVoice.zh = v;
        state.settings.browserVoiceZh = v.name || v.voiceURI;
      } else {
        state.browserVoice.de = v;
        state.settings.browserVoiceDe = v.name || v.voiceURI;
      }
      saveSettings();
      updateVoiceList();
    };
    const test = document.createElement('button');
    test.className = 'btn ghost';
    test.textContent = 'Probehören';
    test.onclick = () => {
      const u = new SpeechSynthesisUtterance(state.voicePanelTarget === 'zh' ? '这是一个测试。' : 'Dies ist ein Test.');
      u.lang = (state.voicePanelTarget === 'zh') ? 'zh-CN' : 'de-DE';
      u.voice = v;
      try { speechSynthesis.cancel(); } catch (e) {}
      speechSynthesis.speak(u);
    };
    const act = (state.voicePanelTarget === 'zh' ? state.browserVoice.zh : state.browserVoice.de);
    if (act && (act.name === v.name || act.voiceURI === v.voiceURI)) name.textContent += '  •  [Aktiv]';
    actions.appendChild(pick);
    actions.appendChild(test);
    row.appendChild(name);
    row.appendChild(meta);
    row.appendChild(actions);
    box.appendChild(row);
  });
}

function refreshVoices() {
  state.voices = window.speechSynthesis?.getVoices?.() || [];
  if (state.settings.browserVoiceZh) {
    const vz = state.voices.find(x => x.name === state.settings.browserVoiceZh || x.voiceURI === state.settings.browserVoiceZh);
    if (vz) state.browserVoice.zh = vz;
  }
  if (state.settings.browserVoiceDe) {
    const vd = state.voices.find(x => x.name === state.settings.browserVoiceDe || x.voiceURI === state.settings.browserVoiceDe);
    if (vd) state.browserVoice.de = vd;
  }
  updateVoiceList();
}

let _voicesRetryT;
function openVoicesPanelFor(target) {
  state.voicePanelTarget = target;
  refreshVoices();
  if (!state.voices || state.voices.length === 0) {
    clearTimeout(_voicesRetryT);
    let tries = 0;
    const tick = () => {
      tries++;
      refreshVoices();
      if (state.voices.length > 0 || tries >= 8) return;
      _voicesRetryT = setTimeout(tick, 300);
    };
    _voicesRetryT = setTimeout(tick, 300);
  }
  $('#voicePanel').classList.remove('hidden');
}
function closeVoices() { $('#voicePanel').classList.add('hidden'); }

// Card class (inferred/adapted from original processing)
class Card {
  constructor(deWord, pyWord, pos, pySent, deSent, zhWord, zhSent, lessonIndex) {
    this.deWord = deWord || '';
    this.pyWord = pyWord || '';
    this.pos = pos || '';
    this.pySent = pySent || '';
    this.deSent = deSent || '';
    this.zhWord = zhWord || '';
    this.zhSent = zhSent || '';
    this.lesson = lessonIndex;
    // Progress per card can be in state.progress
  }
}

// Data loading - Updated for dynamic sheets and A2 names
async function loadData() {
  try {
    console.log('Starting loadData...');
    const statusEl = $('#status') || document.getElementById('status') || els.lessonSelect; // Fallback
    statusEl.textContent = 'Lade Excel-Daten...';

    const response = await fetch(EXCEL_URL);
    if (!response.ok) throw new Error(`HTTP ${response.status}: Stelle sicher, dass die Datei existiert.`);
    console.log('Fetch successful');
    const buf = await response.arrayBuffer();

    const wb = XLSX.read(buf, { type: 'array' });
    const sheetNames = wb.SheetNames;
    console.log('Available sheets:', sheetNames);

    if (sheetNames.length < 2) throw new Error('Zu wenige Sheets.');

    state.lessons.clear();
    let lessonIndex = 0;
    let totalCards = 0;

    // Collect all matching sheets dynamically (skip master, match L*)
    const matchingSheets = sheetNames
      .filter(name => name !== 'Long - Master' && SHEET_NAME_PATTERN.test(name))
      .map(name => ({ name, num: parseInt(name.match(/\d+/)?.[0] || '0', 10) })) // Extract num for sorting
      .sort((a, b) => a.num - b.num); // Sort numerically (L00, L01, ..., L10, L16)

    console.log('Matching sheets:', matchingSheets.map(s => s.name));

    for (const { name: sheetName } of matchingSheets) {
      console.log(`Processing sheet: "${sheetName}"`);
      const sh = wb.Sheets[sheetName];
      if (!sh || !sh['!ref']) {
        console.warn(`Skipping empty sheet: ${sheetName}`);
        continue;
      }

      // Parse like original: header:1, blankrows:false
      const rows = XLSX.utils.sheet_to_json(sh, { header: 1, blankrows: false });
      console.log(`Sheet "${sheetName}" has ${rows.length} rows (after blank skip)`);

      if (rows.length < DATA_START_ROW) { // Need at least up to row 2 (index 1)
        console.warn(`Sheet too short: ${sheetName}`);
        continue;
      }

      // Lesson name from A2 (row 1, col 0 / index 0)
      let fullName = rows[1] && rows[1][0] ? String(rows[1][0]).trim() : '';
      console.log(`A2 full: "${fullName}" (from col A)`);
      let lessonName = fullName.length >= 2 ? fullName.substring(2).trim() : '';
      if (!lessonName) {
        // Fallback to sheet-derived
        const num = parseInt(sheetName.match(/\d+/)?.[0] || lessonIndex, 10);
        lessonName = `${String(num).padStart(2, '0')} - Unnamed`;
      }
      console.log(`Derived name: "${lessonName}"`);

      // Process data from row 3+ (index 2+)
      const r0 = DATA_START_ROW - 1; // 2
      const cards = [];
      for (let r = r0; r < rows.length; r++) {
        const rowData = rows[r];
        // Skip if short or no key data (match original cols, ignore if col A empty)
        if (!Array.isArray(rowData) || rowData.length < 8 || !rowData[COL_WORD.de] || !rowData[COL_WORD.zh]) {
          if (rowData[COL_WORD.de]) console.log(`Skipping partial row ${r + 1}:`, rowData.slice(0, 8));
          continue;
        }

        const card = new Card(
          rowData[COL_WORD.de],    // German word (B/1)
          rowData[COL_WORD.py],    // Pinyin word (C/2)
          rowData[COL_POS],        // POS (D/3)
          rowData[COL_SENT.py],    // Sent Pinyin (E/4)
          rowData[COL_SENT.de],    // Sent German (F/5)
          rowData[COL_WORD.zh],    // Chinese word (G/6)
          rowData[COL_SENT.zh] || '', // Sent Chinese (H/7)
          lessonIndex
        );
        cards.push(card);
      }

      console.log(`Cards in "${sheetName}": ${cards.length}`);
      totalCards += cards.length;

      if (cards.length > 0) {
        state.lessons.set(lessonIndex, { name: lessonName, cards: cards });
        // Init progress like original (byLesson)
        if (!state.progress.byLesson[lessonIndex]) {
          state.progress.byLesson[lessonIndex] = { known: 0, unknown: 0 };
          // Per-card if needed: state.progress.cards[`${lessonIndex}-${cardIdx}`] = ...
        }
        lessonIndex++;
      } else {
        console.warn(`No valid cards in: ${sheetName}`);
      }
    }

    if (state.lessons.size === 0) {
      throw new Error('Keine Lektionen gefunden. Überprüfe Sheet-Namen (L*), A2 und Daten in Spalten B-H ab Zeile 3.');
    }

    console.log(`Loaded ${state.lessons.size} lessons, ${totalCards} cards`);
    const lessonList = Array.from(state.lessons.entries()).map(([idx, l]) => `${idx}: ${l.name}`);
    console.log('Lessons:', lessonList);

    updateLessonSelect(); // Dynamic UI
    statusEl.textContent = `Geladen: ${state.lessons.size} Lektionen (${totalCards} Karten)`;

    loadSettings();
    loadProgress();
    // Update selected from settings if any

  } catch (error) {
    console.error('loadData Error:', error);
    const statusEl = $('#status') || els.lessonSelect;
    statusEl.textContent = `Fehler: ${error.message}. Siehe Konsole.`;
  }
}

// Dynamic lesson select (new, clears and rebuilds)
function updateLessonSelect() {
  console.log('Updating lesson select...');
  if (!els.lessonSelect) return;
  els.lessonSelect.innerHTML = state.lessons.size === 0 ? '<p>Keine Lektionen.</p>' : '';

  state.lessons.forEach((lesson, index) => {
    const container = document.createElement('div');
    container.className = 'lesson-item'; // Assume CSS class
    container.innerHTML = `
      <input type="checkbox" id="lesson-${index}" value="${index}">
      <label for="lesson-${index}">${lesson.name}</label>
    `;
    els.lessonSelect.appendChild(container);

    const cb = document.getElementById(`lesson-${index}`);
    cb.addEventListener('change', (e) => {
      if (e.target.checked) {
        state.selectedLessons.add(parseInt(e.target.value));
      } else {
        state.selectedLessons.delete(parseInt(e.target.value));
      }
      // Update start button etc. (as in original)
    });
  });

  // Load initial selection from settings
  state.settings.lessons?.forEach(idx => {
    const cb = document.getElementById(`lesson-${idx}`);
    if (cb) cb.checked = true;
    state.selectedLessons.add(idx);
  });
}

// Rest of original code: startTraining, showCard, TTS speak, event listeners, etc.
// (Include all from original app.js after parseExcelBuffer, adapting state.lessons.get(index).cards for pool building, etc.)
// For brevity, assume you paste the remaining original code here (from after the truncated for loop: card creation, pool build, UI updates, training functions, init).
// Example adaptation for buildPool (inferred):
function buildCardPool() {
  state.pool = [];
  state.selectedLessons.forEach(idx => {
    const lesson = state.lessons.get(idx);
    if (lesson) lesson.cards.forEach(card => state.pool.push({ ...card, lessonIdx: idx }));
  });
  if (state.order === 'random') state.pool.sort(() => Math.random() - 0.5);
}

// Init (load on DOMContentLoaded)
document.addEventListener('DOMContentLoaded', () => {
  speechSynthesis.addEventListener('voiceschanged', refreshVoices);
  refreshVoices();
  loadData();
  // Add original event listeners: mode change, start btn, etc.
  // e.g., $('#start-btn').addEventListener('click', startTraining);
});

// For GitHub Pages absolute path, change EXCEL_URL to '/Chinesisch-Lektionen/data/Long-Chinesisch_Lektionen.xlsx' if needed.
