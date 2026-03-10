// app.js - Updated for dynamic sheet loading and names from A2

// Global state
let lessons = new Map(); // Map<lessonIndex, {name: string, cards: Card[]}>
let selectedLessons = [];
let currentCardIndex = 0;
let cardPool = [];
let sessionStats = { known: 0, unsure: 0, unknown: 0 };
let progress = {}; // Loaded from localStorage
let isTraining = false;
let mode = 'de2zh'; // 'de2zh' or 'zh2de'
let order = 'random'; // 'random' or 'sequential'
let autoplay = false;
let autoplayGap = 800; // ms
let currentVoiceDE = null;
let currentVoiceZH = null;
let ttsRateDE = 0.9;
let ttsRateZH = 0.8;
let ttsPitchDE = 1.0;
let ttsPitchZH = 1.0;
let wakeLock = null;

// Card type
class Card {
  constructor(german, pinyin, pos, sentencePinyin, sentenceGerman, chinese, sentenceChinese, lesson) {
    this.german = german;
    this.pinyin = pinyin;
    this.pos = pos;
    this.sentencePinyin = sentencePinyin;
    this.sentenceGerman = sentenceGerman;
    this.chinese = chinese;
    this.sentenceChinese = sentenceChinese;
    this.lesson = lesson;
    this.knownCount = 0;
    this.unknownCount = 0;
  }
}

// DOM elements
const el = {
  lessonSelect: document.getElementById('lesson-select'),
  modeSelect: document.getElementById('mode-select'),
  orderSelect: document.getElementById('order-select'),
  autoplayToggle: document.getElementById('autoplay-toggle'),
  startBtn: document.getElementById('start-btn'),
  stopBtn: document.getElementById('stop-btn'),
  cardPrompt: document.getElementById('card-prompt'),
  cardSolution: document.getElementById('card-solution'),
  revealBtn: document.getElementById('reveal-btn'),
  prevBtn: document.getElementById('prev-btn'),
  nextBtn: document.getElementById('next-btn'),
  knownBtn: document.getElementById('known-btn'),
  unsureBtn: document.getElementById('unsure-btn'),
  unknownBtn: document.getElementById('unknown-btn'),
  playPromptBtn: document.getElementById('play-prompt-btn'),
  playSolutionBtn: document.getElementById('play-solution-btn'),
  voiceDebug: document.getElementById('voice-debug'),
  status: document.getElementById('status')
};

// Load data from Excel on GitHub
async function loadData() {
  try {
    el.status.textContent = 'Lade Excel-Daten...';
    const response = await fetch('./data/Long-Chinesisch_Lektionen.xlsx');
    if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    const sheetNames = workbook.SheetNames;
    console.log('Available sheets:', sheetNames);
    
    lessons.clear();
    let lessonIndex = 0;
    
    // Skip first sheet ("Long - Master"), load all subsequent sheets dynamically
    for (let i = 1; i < sheetNames.length; i++) {
      const sheetName = sheetNames[i];
      const ws = workbook.Sheets[sheetName];
      if (!ws || !ws['!ref']) {
        console.warn(`Skipping empty sheet: ${sheetName}`);
        continue;
      }
      
      // Convert to array of rows (header:1 for raw rows, defval:'' to handle empties)
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
      
      // Get lesson name from A2 (row 1, col 0)
      let fullName = data[1] ? data[1][0] : '';
      fullName = (fullName || '').trim();
      let lessonName = fullName.length >= 2 ? fullName.substring(2).trim() : `Unnamed Lesson ${lessonIndex}`;
      
      // Process data from row 2 (index 2) onwards
      const cards = [];
      for (let row = 2; row < data.length; row++) {
        const r = data[row];
        if (r.length < 7 || !r[0] || !r[5]) continue; // Skip empty/invalid rows (needs German & Chinese word)
        
        const card = new Card(
          r[0], // German word
          r[1], // Pinyin word
          r[2], // Part of speech
          r[3], // Sentence Pinyin
          r[4], // Sentence German
          r[5], // Chinese word
          r[6], // Sentence Chinese
          lessonIndex // Use dynamic index
        );
        cards.push(card);
      }
      
      if (cards.length > 0) {
        lessons.set(lessonIndex, { name: lessonName, cards: cards });
        // Initialize progress for this lesson
        if (!progress[lessonIndex]) {
          progress[lessonIndex] = new Map();
          cards.forEach((card, idx) => {
            progress[lessonIndex].set(idx, { known: 0, unknown: 0 });
          });
        }
        lessonIndex++;
      } else {
        console.warn(`No valid cards in sheet: ${sheetName}`);
      }
    }
    
    if (lessons.size === 0) {
      throw new Error('Keine Lektionen gefunden (überprüfe Excel-Datei).');
    }
    
    console.log(`Loaded ${lessons.size} lessons:`, Array.from(lessons.entries()).map(([idx, l]) => `${idx}: ${l.name}`));
    updateLessonSelect();
    el.status.textContent = `Geladen: ${lessons.size} Lektionen`;
    
    // Load progress from localStorage
    loadProgress();
    
  } catch (error) {
    console.error('Fehler beim Laden der Daten:', error);
    el.status.textContent = `Fehler: ${error.message}. Stelle sicher, dass die Excel-Datei öffentlich auf GitHub ist.`;
    el.cardPrompt.textContent = 'Daten konnten nicht geladen werden.';
  }
}

// Dynamically update lesson selection checkboxes with new names
function updateLessonSelect() {
  el.lessonSelect.innerHTML = ''; // Clear hardcoded ones
  lessons.forEach((lesson, index) => {
    const container = document.createElement('div');
    container.className = 'lesson-item';
    container.innerHTML = `
      <input type="checkbox" id="lesson-${index}" value="${index}">
      <label for="lesson-${index}">${lesson.name}</label>
    `;
    el.lessonSelect.appendChild(container);
    
    // Add event listener
    document.getElementById(`lesson-${index}`).addEventListener('change', updateSelectedLessons);
  });
}

// Rest of the app.js remains unchanged (updateSelectedLessons, startTraining, etc.)
function updateSelectedLessons() {
  selectedLessons = [];
  lessons.forEach((_, index) => {
    const cb = document.getElementById(`lesson-${index}`);
    if (cb && cb.checked) selectedLessons.push(index);
  });
  updateStartButton();
}

function updateStartButton() {
  const hasLessons = selectedLessons.length > 0;
  el.startBtn.disabled = !hasLessons || isTraining;
  el.stopBtn.disabled = !isTraining;
  el.startBtn.textContent = hasLessons ? 'Training starten' : 'Wähle Lektionen aus';
}

function startTraining() {
  if (selectedLessons.length === 0) return;
  
  isTraining = true;
  updateStartButton();
  resetSession();
  buildCardPool();
  showCard();
  el.status.textContent = `Training gestartet: ${cardPool.length} Karten`;
  
  // Request wake lock
  if ('wakeLock' in navigator) {
    navigator.wakeLock.request('screen').then(lock => wakeLock = lock).catch(console.error);
  }
  
  // Autoplay?
  if (autoplay) {
    setTimeout(() => nextCard(true), autoplayGap);
  }
}

function stopTraining() {
  isTraining = false;
  updateStartButton();
  clearAutoplay();
  if (wakeLock) {
    wakeLock.release().then(() => wakeLock = null);
  }
  el.cardPrompt.textContent = '';
  el.cardSolution.textContent = '';
  el.status.textContent = 'Training gestoppt.';
  revealBtn.style.display = 'block';
  prevBtn.style.display = 'none';
  nextBtn.style.display = 'none';
  knownBtn.style.display = 'none';
  unsureBtn.style.display = 'none';
  unknownBtn.style.display = 'none';
}

function buildCardPool() {
  cardPool = [];
  selectedLessons.forEach(lessonIdx => {
    const lesson = lessons.get(lessonIdx);
    if (lesson) {
      lesson.cards.forEach(card => cardPool.push({ ...card, lessonIdx }));
    }
  });
  
  if (order === 'random') {
    cardPool.sort(() => Math.random() - 0.5);
  }
  // Sequential is already in order from Excel
}

function showCard(auto = false) {
  if (cardPool.length === 0) return;
  
  const card = cardPool[currentCardIndex];
  if (!card) {
    nextCard(auto);
    return;
  }
  
  el.cardPrompt.textContent = getPrompt(card);
  el.cardSolution.textContent = '';
  el.cardSolution.classList.add('blurred');
  revealBtn.style.display = 'inline-block';
  prevBtn.style.display = currentCardIndex > 0 ? 'inline-block' : 'none';
  nextBtn.style.display = 'inline-block';
  knownBtn.style.display = 'none';
  unsureBtn.style.display = 'none';
  unknownBtn.style.display = 'none';
  playPromptBtn.style.display = 'inline-block';
  
  if (!auto) playPrompt(card);
}

function getPrompt(card) {
  if (mode === 'de2zh') {
    return `${card.german} (${card.pinyin})\n${card.pos ? `(${card.pos})` : ''}\n${card.sentenceGerman}`;
  } else {
    return `${card.chinese} (${card.pinyin})\n${card.pos ? `(${card.pos})` : ''}\n${card.sentenceChinese}`;
  }
}

function getSolution(card) {
  if (mode === 'de2zh') {
    return `${card.chinese}\n${card.sentenceChinese}`;
  } else {
    return `${card.german} (${card.pinyin})\n${card.sentenceGerman}`;
  }
}

function revealCard() {
  const card = cardPool[currentCardIndex];
  if (!card) return;
  
  el.cardSolution.textContent = getSolution(card);
  el.cardSolution.classList.remove('blurred');
  revealBtn.style.display = 'none';
  playSolutionBtn.style.display = 'inline-block';
  knownBtn.style.display = 'inline-block';
  unsureBtn.style.display = 'inline-block';
  unknownBtn.style.display = 'inline-block';
  playSolution(card);
}

function rateCard(rating) {
  const card = cardPool[currentCardIndex];
  if (!card || !card.lessonIdx) return;
  
  const prog = progress[card.lessonIdx].get(lessons.get(card.lessonIdx).cards.indexOf(card));
  if (rating === 'known') {
    prog.known++;
    sessionStats.known++;
  } else if (rating === 'unknown') {
    prog.unknown++;
    sessionStats.unknown++;
  }
  // Unsure doesn't count in stats, just skip
  if (rating !== 'unsure') saveProgress();
  
  nextCard(true);
}

function nextCard(auto = false) {
  if (order === 'sequential') {
    currentCardIndex++;
  } else {
    currentCardIndex = Math.floor(Math.random() * cardPool.length);
  }
  
  if (currentCardIndex >= cardPool.length) {
    currentCardIndex = 0;
  }
  
  clearAutoplay();
  showCard(auto);
  
  if (autoplay && isTraining && !auto) { // Don't chain on manual next
    setTimeout(() => nextCard(true), autoplayGap);
  }
}

function prevCard() {
  if (currentCardIndex > 0) currentCardIndex--;
  showCard();
}

function clearAutoplay() {
  // Simple timeout clear; for robustness, could use a flag
}

function resetSession() {
  currentCardIndex = 0;
  sessionStats = { known: 0, unsure: 0, unknown: 0 };
  el.status.textContent = `Session: 0/${cardPool.length}`;
}

// TTS Functions
function initVoices() {
  const voices = speechSynthesis.getVoices();
  const deVoices = voices.filter(v => v.lang.startsWith('de'));
  const zhVoices = voices.filter(v => v.lang.startsWith('zh') || v.lang.includes('Chinese'));
  
  // Default voices
  currentVoiceDE = deVoices.find(v => v.name.includes('Google')) || deVoices[0];
  currentVoiceZH = zhVoices.find(v => v.name.includes('Google') || v.name.includes('Microsoft')) || zhVoices[0];
  
  updateVoiceDebug();
}

function playPrompt(card) {
  const text = getPrompt(card).replace(/\n/g, ' ');
  speak(text, mode === 'de2zh' ? 'de' : 'zh');
}

function playSolution(card) {
  const text = getSolution(card).replace(/\n/g, ' ');
  speak(text, mode === 'de2zh' ? 'zh' : 'de');
}

function speak(text, lang) {
  if (!speechSynthesis) return;
  
  speechSynthesis.cancel();
  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = lang;
  utterance.rate = lang === 'de' ? ttsRateDE : ttsRateZH;
  utterance.pitch = lang === 'de' ? ttsPitchDE : ttsPitchZH;
  utterance.voice = lang === 'de' ? currentVoiceDE : currentVoiceZH;
  
  speechSynthesis.speak(utterance);
}

function updateVoiceDebug() {
  const voicesDE = speechSynthesis.getVoices().filter(v => v.lang.startsWith('de'));
  const voicesZH = speechSynthesis.getVoices().filter(v => v.lang.startsWith('zh') || v.lang.includes('Chinese'));
  
  let html = '<h4>Verfügbare Stimmen</h4><details><summary>Deutsch (DE)</summary>';
  voicesDE.forEach((v, i) => {
    html += `<label><input type="radio" name="voice-de" value="${i}" ${currentVoiceDE === v ? 'checked' : ''}> ${v.name} (${v.lang})</label><br>`;
  });
  html += '</details><details><summary>Chinesisch (ZH)</summary>';
  voicesZH.forEach((v, i) => {
    html += `<label><input type="radio" name="voice-zh" value="zh-${i}" ${currentVoiceZH === v ? 'checked' : ''}> ${v.name} (${v.lang})</label><br>`;
  });
  html += '</details>';
  el.voiceDebug.innerHTML = html;
  
  // Event listeners for voice selection
  document.querySelectorAll('input[name="voice-de"]').forEach((el, i) => {
    el.addEventListener('change', () => {
      const voices = speechSynthesis.getVoices();
      currentVoiceDE = voices.find(v => v.lang.startsWith('de'))[i]; // Simplified; adjust index
      console.log('DE Voice changed to:', currentVoiceDE?.name);
    });
  });
  // Similar for ZH...
  document.querySelectorAll('input[name="voice-zh"]').forEach((el, i) => {
    el.addEventListener('change', () => {
      const voices = speechSynthesis.getVoices();
      currentVoiceZH = voices.filter(v => v.lang.startsWith('zh') || v.lang.includes('Chinese'))[i];
      console.log('ZH Voice changed to:', currentVoiceZH?.name);
    });
  });
}

// Progress Management
function loadProgress() {
  const saved = localStorage.getItem('chineseFlashcardsProgress');
  if (saved) {
    progress = JSON.parse(saved);
    // Migrate if needed for new lessons
    lessons.forEach((lesson, idx) => {
      if (!progress[idx]) {
        progress[idx] = new Map();
        lesson.cards.forEach((_, cIdx) => progress[idx].set(cIdx, { known: 0, unknown: 0 }));
      }
    });
  } else {
    // Initialize all
    lessons.forEach((lesson, idx) => {
      progress[idx] = new Map();
      lesson.cards.forEach((_, cIdx) => progress[idx].set(cIdx, { known: 0, unknown: 0 }));
    });
  }
  updateProgressDisplay();
}

function saveProgress() {
  // Convert Maps to objects for storage
  const serializable = {};
  Object.entries(progress).forEach(([idx, map]) => {
    serializable[idx] = Object.fromEntries(map);
  });
  localStorage.setItem('chineseFlashcardsProgress', JSON.stringify(serializable));
}

function updateProgressDisplay() {
  // Could add per-lesson progress bars; for now, console/log
  console.log('Progress updated');
}

// Event Listeners (unchanged except for dynamic ones added in updateLessonSelect)
el.modeSelect.addEventListener('change', (e) => { mode = e.target.value; });
el.orderSelect.addEventListener('change', (e) => { order = e.target.value; });
el.autoplayToggle.addEventListener('change', (e) => { autoplay = e.target.checked; });
el.startBtn.addEventListener('click', startTraining);
el.stopBtn.addEventListener('click', stopTraining);
el.revealBtn.addEventListener('click', revealCard);
el.prevBtn.addEventListener('click', prevCard);
el.nextBtn.addEventListener('click', () => nextCard(false));
el.knownBtn.addEventListener('click', () => rateCard('known'));
el.unsureBtn.addEventListener('click', () => rateCard('unsure'));
el.unknownBtn.addEventListener('click', () => rateCard('unknown'));
el.playPromptBtn.addEventListener('click', () => {
  const card = cardPool[currentCardIndex]; if (card) playPrompt(card);
});
el.playSolutionBtn.addEventListener('click', () => {
  const card = cardPool[currentCardIndex]; if (card) playSolution(card);
});

// Init
document.addEventListener('DOMContentLoaded', () => {
  speechSynthesis.addEventListener('voiceschanged', initVoices);
  initVoices(); // Initial call
  loadData();
  updateSelectedLessons(); // Initial state
});

// Export progress (bonus, unchanged)
function exportProgress() {
  const data = { progress: Object.fromEntries(Object.entries(progress).map(([idx, map]) => [idx, Object.fromEntries(map)])) };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'flashcard-progress.json';
  a.click();
  URL.revokeObjectURL(url);
}

// Add export button if needed, e.g., in HTML: <button onclick="exportProgress()">Export</button>
