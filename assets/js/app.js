/* Flashcards v5.3
 - Stimmenauswahl via Dropdown (pro Sprache) + manuelles Reload der Stimmenliste
 - Kein Countdown, kein Auto-Reveal/-Play (Reveal nur per Button)
 - 1000ms Delay zwischen Wort und Satz beim Abspielen
 - POS=16px, sonst 20px (über CSS)
*/

const EXCEL_URL = 'https://luoboteanjing-netizen.github.io/Chinesisch-Reader/data/Long-Chinesisch_Lektionen.xlsx';
const SHEET_NAME_PATTERN = /^L\s*\d{1,2}$/i;
const MIN_LESSON = 0, MAX_LESSON = 16;
const DATA_START_ROW = 3; // 1-basiert
const COL_WORD = { de:1, py:2, zh:6 };
const COL_SENT = { de:5, py:4, zh:7 };
const COL_POS  = 3; // Wortart

const state = {
  mode:'de2zh', order:'random', rate:0.95,
  voices:[],
  voiceSelected: { zh:null, de:null }, // will store voiceURI
  lessons:new Map(), selectedLessons:new Set(), pool:[],
  current:null, idx:null,
};

const $ = s=>document.querySelector(s);

// ===== Stimmen laden & Dropdowns füllen =====
function refreshVoices(){
  state.voices = window.speechSynthesis?.getVoices?.() || [];
}

function voiceMatchesLang(v, langKey){
  const L = (v.lang||'').toLowerCase();
  if(langKey==='zh') return L.startsWith('zh') || L.includes('cmn') || L.includes('hans') || L.includes('zh-cn');
  if(langKey==='de') return L.startsWith('de');
  return false;
}

function populateVoiceSelects(){
  const zhSel = $('#zhVoiceSelect');
  const deSel = $('#deVoiceSelect');
  if(!zhSel || !deSel) return;
  zhSel.innerHTML=''; deSel.innerHTML='';

  const zhList = state.voices.filter(v=> voiceMatchesLang(v,'zh'));
  const deList = state.voices.filter(v=> voiceMatchesLang(v,'de'));

  function fill(sel, list, currentURI){
    list.forEach((v,i)=>{
      const opt=document.createElement('option');
      opt.value = v.voiceURI || v.name || String(i);
      opt.textContent = `${v.name || 'Voice'} (${v.lang||''})`;
      if(currentURI && (v.voiceURI===currentURI || v.name===currentURI)) opt.selected=true;
      sel.appendChild(opt);
    });
    // Fallback: select first if none
    if(!sel.value && list.length) sel.selectedIndex=0;
  }

  fill(zhSel, zhList, state.voiceSelected.zh);
  fill(deSel, deList, state.voiceSelected.de);

  // store selection
  state.voiceSelected.zh = zhSel.value || (zhList[0]?.voiceURI || zhList[0]?.name || null);
  state.voiceSelected.de = deSel.value || (deList[0]?.voiceURI || deList[0]?.name || null);
}

function getSelectedVoice(langKey){
  const uri = state.voiceSelected[langKey];
  if(!uri) return null;
  // try by voiceURI, then by name
  let v = state.voices.find(v=> (v.voiceURI===uri));
  if(!v) v = state.voices.find(v=> (v.name===uri));
  // final fallback: first matching language
  if(!v) v = state.voices.find(v=> voiceMatchesLang(v, langKey));
  return v || null;
}

// ===== Excel laden =====
async function loadExcel(){
  const status=$('#excelStatus');
  try{
    status.textContent='Excel wird geladen…';
    const res=await fetch(EXCEL_URL,{cache:'no-store'}); if(!res.ok) throw new Error('HTTP '+res.status);
    const buf=await res.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array'});
    state.lessons.clear();
    for(const name of wb.SheetNames){
      if(!SHEET_NAME_PATTERN.test(name)) continue;
      const m=name.match(/\d+/); if(!m) continue; const n=parseInt(m[0],10);
      if(!(n>=MIN_LESSON && n<=MAX_LESSON)) continue;
      const sh=wb.Sheets[name]; if(!sh) continue;
      const rows=XLSX.utils.sheet_to_json(sh,{header:1,blankrows:false}); if(!rows||!rows.length) continue;
      const r0=DATA_START_ROW-1; const key=String(n);
      if(!state.lessons.has(key)) state.lessons.set(key,[]);
      for(let r=r0;r<rows.length;r++){
        const row=rows[r]||[];
        const w={ de:String(row[COL_WORD.de-1]||'').trim(), py:String(row[COL_WORD.py-1]||'').trim(), zh:String(row[COL_WORD.zh-1]||'').trim() };
        const s={ de:String(row[COL_SENT.de-1]||'').trim(), py:String(row[COL_SENT.py-1]||'').trim(), zh:String(row[COL_SENT.zh-1]||'').trim() };
        const pos=String(row[COL_POS-1]||'').trim();
        if(!(w.de||w.zh||s.de||s.zh)) continue;
        state.lessons.get(key).push({word:w, sent:s, pos});
      }
    }
    populateLessonSelect();
    status.textContent=`Excel geladen (${state.lessons.size} Lektion(en)).`;
  }catch(err){ status.textContent='Excel konnte nicht geladen werden: '+err.message; console.error(err); }
}

function populateLessonSelect(){
  const sel=$('#lessonSelect'); sel.innerHTML='';
  const keys=Array.from(state.lessons.keys()).map(k=>parseInt(k,10)).sort((a,b)=>a-b);
  for(const k of keys){ const cnt=state.lessons.get(String(k)).length; const opt=document.createElement('option'); opt.value=String(k); opt.textContent=`Lektion ${k} (${cnt})`; sel.appendChild(opt); }
}

function gatherPool(){ const out=[]; for(const k of state.selectedLessons){ const arr=state.lessons.get(k); if(arr) out.push(...arr);} state.pool=out; state.idx=null; }

function setCard(entry){
  state.current=entry;
  $('#solBox').classList.add('masked');

  if(state.mode==='zh2de'){
    $('#promptWord').innerHTML = (entry.word.zh||'—');
    $('#promptWordSub').innerHTML = formatPinyinAndPos(entry.word.py, entry.pos);
    $('#promptSent').innerHTML = formatZh(entry.sent.zh, entry.sent.py);
    $('#solWord').textContent = entry.word.de || '—';
    $('#solSent').textContent = entry.sent.de || '—';
  } else {
    $('#promptWord').textContent = entry.word.de || '—';
    $('#promptWordSub').textContent = entry.pos ? entry.pos : '';
    $('#promptSent').textContent = entry.sent.de || '—';
    $('#solWord').innerHTML = formatZh(entry.word.zh, entry.word.py);
    $('#solSent').innerHTML = formatZh(entry.sent.zh, entry.sent.py);
  }

  // Buttons aktivieren
  $('#btnNext').disabled=false; $('#btnReveal').disabled=false; $('#btnPlayQ').disabled=false; $('#btnPlayA').disabled=false;
  $('#btnPrev').disabled=(state.order!=='seq');
}

function nextCard(){ if(!state.pool.length){ alert('Bitte Lektionen wählen und übernehmen.'); return; } if(state.order==='seq'){ if(state.idx===null) state.idx=0; else state.idx=(state.idx+1)%state.pool.length; setCard(state.pool[state.idx]); } else { const e=state.pool[Math.floor(Math.random()*state.pool.length)]; setCard(e);} }
function prevCard(){ if(state.order!=='seq'||!state.pool.length) return; if(state.idx===null) state.idx=0; else state.idx=(state.idx-1+state.pool.length)%state.pool.length; setCard(state.pool[state.idx]); }
function startTraining(){ if(!state.pool.length){ alert('Bitte zuerst Lektion(en) übernehmen.'); return; } if(state.order==='seq'){ state.idx=0; setCard(state.pool[state.idx]); } else { const e=state.pool[Math.floor(Math.random()*state.pool.length)]; setCard(e);} }

function doReveal(){ $('#solBox').classList.remove('masked'); }

function formatZh(hz,py){ const hanzi=(hz||'').trim(); const pinyin=(py||'').trim(); return pinyin? `${hanzi}<br><span class="py">${pinyin}</span>` : (hanzi||'—'); }
function formatPinyinAndPos(py,pos){ const a=(py||'').trim(); const b=(pos||'').trim(); if(a && b) return `<span class="py">${a}</span><br><span class=\"prompt small\" style=\"display:inline-block;margin-top:6px;\">${b}</span>`; if(a) return `<span class="py">${a}</span>`; if(b) return `<span class=\"prompt small\" style=\"display:inline-block;margin-top:6px;\">${b}</span>`; return ''; }

// ===== TTS =====
function pickVoiceFor(langKey){
  // use explicit selection if available, otherwise best match
  const v = getSelectedVoice(langKey);
  return v || null;
}

function speak(text,langKey){
  const utter = new SpeechSynthesisUtterance(text||'');
  // Map langKey to lang code for safety
  utter.lang = (langKey==='zh')? 'zh-CN' : 'de-DE';
  utter.rate = state.rate;
  const v = pickVoiceFor(langKey);
  if(v) utter.voice = v;
  try{ speechSynthesis.cancel(); }catch(e){}
  speechSynthesis.speak(utter);
}

const DELAY_MS = 1000;
function playQuestion(){ if(!state.current) return; if(state.mode==='de2zh'){ speak(state.current.word.de,'de'); setTimeout(()=>speak(state.current.sent.de,'de'),DELAY_MS);} else { speak(state.current.word.zh,'zh'); setTimeout(()=>speak(state.current.sent.zh,'zh'),DELAY_MS);} }
function playAnswer(){ if(!state.current) return; if(state.mode==='de2zh'){ speak(state.current.word.zh,'zh'); setTimeout(()=>speak(state.current.sent.zh,'zh'),DELAY_MS);} else { speak(state.current.word.de,'de'); setTimeout(()=>speak(state.current.sent.de,'de'),DELAY_MS);} }

// ===== UI =====
window.addEventListener('DOMContentLoaded',()=>{
  // Stimmen laden (mehrfach, da asynchron)
  refreshVoices();
  if('speechSynthesis' in window){
    if(typeof speechSynthesis.onvoiceschanged !== 'undefined'){
      speechSynthesis.onvoiceschanged = ()=>{ refreshVoices(); populateVoiceSelects(); };
    }
  }
  // Erste Füllung (falls Stimmen schon da sind)
  populateVoiceSelects();

  // Excel laden
  loadExcel();

  // UI Events
  document.querySelectorAll('input[name="mode"]').forEach(r=> r.addEventListener('change', e=>{ state.mode=e.target.value; if(state.current) setCard(state.current); }));
  document.querySelectorAll('input[name="order"]').forEach(r=> r.addEventListener('change', e=>{ state.order=e.target.value; state.idx=null; $('#btnPrev').disabled=(state.order!=='seq'); }));

  $('#rateRange').addEventListener('input', e=>{ state.rate=parseFloat(e.target.value); $('#rateVal').textContent=`(${state.rate.toFixed(2)})`; });

  // Voice selects
  $('#zhVoiceSelect').addEventListener('change', e=>{ state.voiceSelected.zh = e.target.value; });
  $('#deVoiceSelect').addEventListener('change', e=>{ state.voiceSelected.de = e.target.value; });
  $('#btnReloadVoices').addEventListener('click', ()=>{ refreshVoices(); populateVoiceSelects(); });

  // Quick test buttons
  $('#btnTestDE').addEventListener('click', ()=> speak('Das ist ein deutscher Stimmtest.','de'));
  $('#btnTestZH').addEventListener('click', ()=> speak('这是一个中文的语音测试。','zh'));

  // Lessons
  $('#btnUseLessons').addEventListener('click', ()=>{ const sel=$('#lessonSelect'); state.selectedLessons.clear(); for(const o of sel.selectedOptions) state.selectedLessons.add(o.value); gatherPool(); $('#btnStart').disabled=(state.pool.length===0); });
  $('#btnClearLessons').addEventListener('click', ()=>{ const sel=$('#lessonSelect'); for(const o of sel.options) o.selected=false; state.selectedLessons.clear(); gatherPool(); $('#btnStart').disabled=true; });

  // Flow
  $('#btnStart').addEventListener('click', startTraining);
  $('#btnNext').addEventListener('click', nextCard);
  $('#btnPrev').addEventListener('click', prevCard);
  $('#btnReveal').addEventListener('click', doReveal);
  $('#btnPlayQ').addEventListener('click', playQuestion);
  $('#btnPlayA').addEventListener('click', playAnswer);
});
