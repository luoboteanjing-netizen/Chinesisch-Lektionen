/* Flashcards v3: Steuer-Buttons unten, Frage/Antwort ohne Labels, ZH-Frage Satz: Hanzi + Pinyin, TTS(ZH)=Hanzi, zusätzlicher Antwort-Abspiel-Button */
// Excel (GitHub Pages)
const EXCEL_URL = 'https://luoboteanjing-netizen.github.io/Chinesisch-Reader/data/Long-Chinesisch_Lektionen.xlsx';
// Tab-Namen wie L 00 .. L 16
const SHEET_NAME_PATTERN = /^L\s*\d{1,2}$/i;
const MIN_LESSON = 0, MAX_LESSON = 16;
// Daten ab Zeile 3 (1-basiert)
const DATA_START_ROW = 3;
// Spalten (1-basiert) Wort / Satz / Wortart
const COL_WORD = { de:1, py:2, zh:6 };
const COL_SENT = { de:5, py:4, zh:7 };
const COL_POS  = 3; // Wortart

const state = {
  mode:'de2zh',
  order:'random', // 'random' | 'seq'
  rate:0.95,
  voicePref:{ zh:'female', de:'female' },
  voices:[],
  lessons:new Map(), // Map<lesson,[{word:{de,py,zh}, sent:{de,py,zh}, pos:string}]>
  selectedLessons:new Set(),
  pool:[],
  current:null,
  idx:null, // nur für seq
  revealTimer:null,
  countdownTimer:null,
};

const $ = s=>document.querySelector(s);

async function loadExcel(){
  const status = $('#excelStatus');
  try{
    status.textContent = 'Excel wird geladen…';
    const res = await fetch(EXCEL_URL, {cache:'no-store'});
    if(!res.ok) throw new Error('HTTP '+res.status);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});

    state.lessons.clear();
    for(const name of wb.SheetNames){
      if(!SHEET_NAME_PATTERN.test(name)) continue;
      const m=name.match(/\d+/); if(!m) continue; const n=parseInt(m[0],10);
      if(!(n>=MIN_LESSON && n<=MAX_LESSON)) continue;
      const sh = wb.Sheets[name]; if(!sh) continue;
      const rows = XLSX.utils.sheet_to_json(sh, {header:1, blankrows:false});
      if(!rows || !rows.length) continue;
      const r0 = DATA_START_ROW-1;
      const key = String(n);
      if(!state.lessons.has(key)) state.lessons.set(key, []);
      for(let r=r0; r<rows.length; r++){
        const row = rows[r]||[];
        const w = {
          de: String(row[COL_WORD.de-1]||'').trim(),
          py: String(row[COL_WORD.py-1]||'').trim(),
          zh: String(row[COL_WORD.zh-1]||'').trim(),
        };
        const s = {
          de: String(row[COL_SENT.de-1]||'').trim(),
          py: String(row[COL_SENT.py-1]||'').trim(),
          zh: String(row[COL_SENT.zh-1]||'').trim(),
        };
        const pos = String(row[COL_POS-1]||'').trim();
        if(!(w.de||w.zh||s.de||s.zh)) continue;
        state.lessons.get(key).push({word:w, sent:s, pos});
      }
    }
    populateLessonSelect();
    status.textContent = `Excel geladen (${state.lessons.size} Lektion(en)).`;
  }catch(err){
    status.textContent = 'Excel konnte nicht geladen werden: '+err.message;
    console.error(err);
  }
}

function populateLessonSelect(){
  const sel = $('#lessonSelect'); sel.innerHTML='';
  const keys = Array.from(state.lessons.keys()).map(k=>parseInt(k,10)).sort((a,b)=>a-b);
  for(const k of keys){
    const cnt = state.lessons.get(String(k)).length;
    const opt = document.createElement('option');
    opt.value=String(k); opt.textContent=`Lektion ${k} (${cnt})`;
    sel.appendChild(opt);
  }
}

function gatherPool(){
  const out=[];
  for(const k of state.selectedLessons){ const arr=state.lessons.get(k); if(arr) out.push(...arr); }
  state.pool = out; // Excel-Reihenfolge
  state.idx = null; // Reset
}

function clearTimers(){
  if(state.revealTimer){ clearTimeout(state.revealTimer); state.revealTimer=null; }
  if(state.countdownTimer){ clearInterval(state.countdownTimer); state.countdownTimer=null; }
}

function setCard(entry){
  state.current = entry; clearTimers();
  const solBox = $('#solBox'); solBox.classList.add('masked');
  $('#countdown').textContent = '10';

  if(state.mode==='zh2de'){
    // FRAGE: Hanzi (groß), darunter Pinyin + POS, darunter Satz Hanzi + Pinyin
    $('#promptWord').innerHTML = (entry.word.zh||'—');
    $('#promptWordSub').innerHTML = formatPinyinAndPos(entry.word.py, entry.pos);
    $('#promptSent').innerHTML = formatZh(entry.sent.zh, entry.sent.py);

    // ANTWORT: Deutsch (ohne POS)
    $('#solWord').textContent = entry.word.de || '—';
    $('#solSent').textContent = entry.sent.de || '—';
  } else {
    // FRAGE: Deutsch Wort, darunter POS, darunter deutscher Satz
    $('#promptWord').textContent = entry.word.de || '—';
    $('#promptWordSub').textContent = entry.pos ? entry.pos : '';
    $('#promptSent').textContent = entry.sent.de || '—';

    // ANTWORT: Chinesisch Hanzi + Pinyin (ohne POS)
    $('#solWord').innerHTML = formatZh(entry.word.zh, entry.word.py);
    $('#solSent').innerHTML = formatZh(entry.sent.zh, entry.sent.py);
  }

  let remain = 10;
  state.countdownTimer = setInterval(()=>{ remain--; if(remain>=0) $('#countdown').textContent=String(remain); },1000);
  state.revealTimer = setTimeout(()=>{ solBox.classList.remove('masked'); clearTimers(); }, 10000);

  // Buttons aktivieren
  $('#btnNext').disabled = false; $('#btnReveal').disabled = false; $('#btnPlayQ').disabled = false; $('#btnPlayA').disabled = false;
  $('#btnPrev').disabled = (state.order !== 'seq');
}

function nextCard(){
  if(!state.pool.length){ alert('Bitte Lektionen wählen und übernehmen.'); return; }
  if(state.order==='seq'){
    if(state.idx===null) state.idx = 0; else state.idx = (state.idx + 1) % state.pool.length;
    setCard(state.pool[state.idx]);
  } else {
    const entry = state.pool[Math.floor(Math.random()*state.pool.length)];
    setCard(entry);
  }
}

function prevCard(){
  if(state.order!=='seq' || !state.pool.length) return;
  if(state.idx===null) state.idx = 0; else state.idx = (state.idx - 1 + state.pool.length) % state.pool.length;
  setCard(state.pool[state.idx]);
}

function startTraining(){
  if(!state.pool.length){ alert('Bitte zuerst Lektion(en) übernehmen.'); return; }
  if(state.order==='seq'){ state.idx = 0; setCard(state.pool[state.idx]); }
  else { const entry = state.pool[Math.floor(Math.random()*state.pool.length)]; setCard(entry); }
}

function formatZh(hz,py){
  const hanzi = (hz||'').trim(); const pinyin = (py||'').trim();
  return pinyin ? `${hanzi}<br><span class="py">${pinyin}</span>` : (hanzi||'—');
}
function formatPinyinAndPos(py, pos){
  const a = (py||'').trim(); const b = (pos||'').trim();
  if(a && b) return `<span class="py">${a}</span><br><span class=\"prompt small\">${b}</span>`;
  if(a) return `<span class="py">${a}</span>`;
  if(b) return `<span class=\"prompt small\">${b}</span>`;
  return '';
}

// ===== TTS =====
function refreshVoices(){ state.voices = window.speechSynthesis?.getVoices?.() || []; }
function pickVoice(lang, gender){
  if(!state.voices.length) return null;
  const list = state.voices.filter(v => (v.lang||'').toLowerCase().startsWith(lang));
  if(!list.length) return null;
  const want=(gender||'').toLowerCase();
  const isF=s=>/female|weib|女/i.test(s); const isM=s=>/male|männ|男/i.test(s);
  const byName = list.filter(v => want==='female' ? isF(v.name+" "+v.voiceURI) : isM(v.name+" "+v.voiceURI));
  return byName[0] || list.find(v=>v.default) || list[0];
}
function speak(text, lang){
  const u = new SpeechSynthesisUtterance(text||'');
  u.lang=lang; u.rate=state.rate;
  const v = lang.startsWith('zh')? pickVoice('zh', state.voicePref.zh) : pickVoice('de', state.voicePref.de);
  if(v) u.voice=v;
  speechSynthesis.cancel();
  speechSynthesis.speak(u);
}

// Abspielen der FRAGE (Quellsprache)
function playQuestion(){
  if(!state.current) return;
  if(state.mode==='de2zh'){
    const w = state.current.word.de||''; const s = state.current.sent.de||'';
    speak(w,'de-DE'); setTimeout(()=> speak(s,'de-DE'), 700);
  } else {
    // ZH→DE: IMMER Hanzi vorlesen
    const w = state.current.word.zh||''; const s = state.current.sent.zh||'';
    speak(w,'zh-CN'); setTimeout(()=> speak(s,'zh-CN'), 700);
  }
}

// Abspielen der ANTWORT (Zielsprache)
function playAnswer(){
  if(!state.current) return;
  if(state.mode==='de2zh'){
    // Antwort ist Chinesisch: immer Hanzi sprechen
    const w = state.current.word.zh||''; const s = state.current.sent.zh||'';
    speak(w,'zh-CN'); setTimeout(()=> speak(s,'zh-CN'), 700);
  } else {
    // Antwort ist Deutsch
    const w = state.current.word.de||''; const s = state.current.sent.de||'';
    speak(w,'de-DE'); setTimeout(()=> speak(s,'de-DE'), 700);
  }
}

// ===== UI =====
window.addEventListener('DOMContentLoaded', ()=>{
  refreshVoices(); if('speechSynthesis' in window && typeof speechSynthesis.onvoiceschanged!=='undefined') speechSynthesis.onvoiceschanged = refreshVoices;
  loadExcel();

  document.querySelectorAll('input[name="mode"]').forEach(r=> r.addEventListener('change', e=>{ state.mode=e.target.value; if(state.current) setCard(state.current); }));
  document.querySelectorAll('input[name="order"]').forEach(r=> r.addEventListener('change', e=>{ state.order=e.target.value; state.idx=null; $('#btnPrev').disabled = (state.order!=='seq'); }));
  document.querySelectorAll('input[name="zhVoice"]').forEach(r=> r.addEventListener('change', e=> state.voicePref.zh=e.target.value));
  document.querySelectorAll('input[name="deVoice"]').forEach(r=> r.addEventListener('change', e=> state.voicePref.de=e.target.value));
  $('#rateRange').addEventListener('input', e=>{ state.rate=parseFloat(e.target.value); $('#rateVal').textContent=`(${state.rate.toFixed(2)})`; });

  $('#btnUseLessons').addEventListener('click', ()=>{ const sel=$('#lessonSelect'); state.selectedLessons.clear(); for(const o of sel.selectedOptions) state.selectedLessons.add(o.value); gatherPool(); $('#btnStart').disabled = state.pool.length===0; });
  $('#btnClearLessons').addEventListener('click', ()=>{ const sel=$('#lessonSelect'); for(const o of sel.options) o.selected=false; state.selectedLessons.clear(); gatherPool(); $('#btnStart').disabled = true; });

  $('#btnStart').addEventListener('click', startTraining);
  $('#btnNext').addEventListener('click', nextCard);
  $('#btnPrev').addEventListener('click', prevCard);
  $('#btnReveal').addEventListener('click', ()=>{ clearTimers(); $('#solBox').classList.remove('masked'); });
  $('#btnPlayQ').addEventListener('click', playQuestion);
  $('#btnPlayA').addEventListener('click', playAnswer);
});
