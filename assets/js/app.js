/* Flashcards v5.4
 - Excel unter ./data standardmäßig; Eingabefeld erlaubt Override
 - Kein Countdown/Autoplay (Reveal nur per Button)
 - Stimmenauswahl via Radio (w/m) mit smarter Heuristik je Sprache + Pitch+Rate
 - 1000ms Delay zwischen Wort und Satz
 - POS=16px, sonst 20px (CSS)
*/

let EXCEL_URL = './data/Long-Chinesisch_Lektionen.xlsx';
const SHEET_NAME_PATTERN = /^L\s*\d{1,2}$/i;
const MIN_LESSON = 0, MAX_LESSON = 16;
const DATA_START_ROW = 3;
const COL_WORD = { de:1, py:2, zh:6 };
const COL_SENT = { de:5, py:4, zh:7 };
const COL_POS  = 3;

const state = {
  mode:'de2zh', order:'random', rate:0.95, pitch:1.0,
  lessons:new Map(), selectedLessons:new Set(), pool:[],
  current:null, idx:null,
  voices:[],
  voicePref:{ zh:'female', de:'female' },
};

const $ = s=>document.querySelector(s);

// ===== Excel laden =====
async function loadExcel(customUrl){
  const status = $('#excelStatus');
  try{
    EXCEL_URL = customUrl || $('#excelPath').value || EXCEL_URL;
    status.textContent = 'Excel wird geladen… ('+EXCEL_URL+')';
    const res = await fetch(EXCEL_URL, {cache:'no-store'});
    if(!res.ok) throw new Error('HTTP '+res.status);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});

    state.lessons.clear();
    let found = [];
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
        const w = { de:String(row[COL_WORD.de-1]||'').trim(), py:String(row[COL_WORD.py-1]||'').trim(), zh:String(row[COL_WORD.zh-1]||'').trim() };
        const s = { de:String(row[COL_SENT.de-1]||'').trim(), py:String(row[COL_SENT.py-1]||'').trim(), zh:String(row[COL_SENT.zh-1]||'').trim() };
        const pos = String(row[COL_POS-1]||'').trim();
        if(!(w.de||w.zh||s.de||s.zh)) continue;
        state.lessons.get(key).push({word:w, sent:s, pos});
      }
      if(state.lessons.get(key).length>0) found.push('L '+key.padStart(2,'0'));
    }
    populateLessonSelect();
    status.textContent = found.length? ('Excel geladen. Gefunden: '+found.join(', ')) : 'Excel geladen, aber keine Lektionsblätter gefunden.';
  }catch(err){ status.textContent = 'Excel konnte nicht geladen werden: '+err.message; console.error(err); }
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
  state.pool = out; state.idx = null;
}

// ===== Karte =====
function setCard(entry){
  state.current = entry;
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

  $('#btnNext').disabled=false; $('#btnReveal').disabled=false; $('#btnPlayQ').disabled=false; $('#btnPlayA').disabled=false;
  $('#btnPrev').disabled=(state.order!=='seq');
}

function nextCard(){ if(!state.pool.length){ alert('Bitte Lektionen wählen und übernehmen.'); return; } if(state.order==='seq'){ if(state.idx===null) state.idx=0; else state.idx=(state.idx+1)%state.pool.length; setCard(state.pool[state.idx]); } else { const e=state.pool[Math.floor(Math.random()*state.pool.length)]; setCard(e);} }
function prevCard(){ if(state.order!=='seq'||!state.pool.length) return; if(state.idx===null) state.idx=0; else state.idx=(state.idx-1+state.pool.length)%state.pool.length; setCard(state.pool[state.idx]); }
function startTraining(){ if(!state.pool.length){ alert('Bitte zuerst Lektion(en) übernehmen.'); return; } if(state.order==='seq'){ state.idx=0; setCard(state.pool[state.idx]); } else { const e=state.pool[Math.floor(Math.random()*state.pool.length)]; setCard(e);} }

function doReveal(){ $('#solBox').classList.remove('masked'); }

function formatZh(hz,py){ const hanzi=(hz||'').trim(); const pinyin=(py||'').trim(); return pinyin? `${hanzi}<br><span class="py">${pinyin}</span>` : (hanzi||'—'); }
function formatPinyinAndPos(py,pos){ const a=(py||'').trim(); const b=(pos||'').trim(); if(a && b) return `<span class="py">${a}</span><br><span class=\"prompt small\" style=\"display:inline-block;margin-top:6px;\">${b}</span>`; if(a) return `<span class="py">${a}</span>`; if(b) return `<span class=\"prompt small\" style=\"display:inline-block;margin-top:6px;\">${b}</span>`; return ''; }

// ===== Stimmen (Web Speech API) =====
function refreshVoices(){ state.voices = window.speechSynthesis?.getVoices?.() || []; }
function matchesLang(v, key){ const L=(v.lang||'').toLowerCase(); if(key==='zh') return L.startsWith('zh')||L.includes('cmn')||L.includes('hans')||L.includes('zh-cn'); if(key==='de') return L.startsWith('de'); return false; }
function genderHint(name){ const s=(name||'').toLowerCase();
  const femaleHint = /(female|weib|女|ting|xiao|mei|ya|li|hui|jing|xin)/i; // heuristisch, z.B. Ting‑Ting, Xiaoxiao, Yating
  const maleHint   = /(male|männ|男|yun|yang|zhiyu|lei|wei|hao)/i; // z.B. Yunyang, Zhiyu
  return femaleHint.test(s)?'female':(maleHint.test(s)?'male':null);
}
function pickVoice(langKey, wantedGender){
  const list = state.voices.filter(v=> matchesLang(v, langKey));
  if(!list.length) return null;
  // 1) exakte Gender-Matches über Heuristik aus name/voiceURI
  const withGender = list.filter(v=> genderHint(v.name+" "+v.voiceURI)===wantedGender);
  if(withGender.length) return withGender[0];
  // 2) wenn gewünschtes Gender nicht verfügbar: nimm Default in Sprache
  const def = list.find(v=> v.default) || null;
  if(def) return def;
  // 3) sonst erstes Element
  return list[0];
}

function speak(text, langKey){
  const u = new SpeechSynthesisUtterance(text||'');
  u.lang = (langKey==='zh')? 'zh-CN' : 'de-DE';
  u.rate = state.rate; u.pitch = state.pitch;
  const v = pickVoice(langKey, state.voicePref[langKey]);
  if(v) u.voice = v;
  try{ speechSynthesis.cancel(); }catch(e){}
  speechSynthesis.speak(u);
}

const DELAY_MS = 1000;
function playQuestion(){ if(!state.current) return; if(state.mode==='de2zh'){ speak(state.current.word.de,'de'); setTimeout(()=>speak(state.current.sent.de,'de'),DELAY_MS);} else { speak(state.current.word.zh,'zh'); setTimeout(()=>speak(state.current.sent.zh,'zh'),DELAY_MS);} }
function playAnswer(){ if(!state.current) return; if(state.mode==='de2zh'){ speak(state.current.word.zh,'zh'); setTimeout(()=>speak(state.current.sent.zh,'zh'),DELAY_MS);} else { speak(state.current.word.de,'de'); setTimeout(()=>speak(state.current.sent.de,'de'),DELAY_MS);} }

// ===== UI =====
window.addEventListener('DOMContentLoaded', ()=>{
  // Excel initial
  $('#excelPath').value = EXCEL_URL;
  $('#btnReloadExcel').addEventListener('click', ()=> loadExcel($('#excelPath').value));
  loadExcel();

  // Modus/Reihenfolge
  document.querySelectorAll('input[name="mode"]').forEach(r=> r.addEventListener('change', e=>{ state.mode=e.target.value; if(state.current) setCard(state.current); }));
  document.querySelectorAll('input[name="zhVoiceGender"]').forEach(r=> r.addEventListener('change', e=> state.voicePref.zh=e.target.value));
  document.querySelectorAll('input[name="deVoiceGender"]').forEach(r=> r.addEventListener('change', e=> state.voicePref.de=e.target.value));

  $('#rateRange').addEventListener('input', e=>{ state.rate=parseFloat(e.target.value); $('#rateVal').textContent=`(${state.rate.toFixed(2)})`; });
  $('#pitchRange').addEventListener('input', e=>{ state.pitch=parseFloat(e.target.value); $('#pitchVal').textContent=`(${state.pitch.toFixed(2)})`; });

  // Stimmen-Handling
  refreshVoices();
  if('speechSynthesis' in window && typeof speechSynthesis.onvoiceschanged!=='undefined'){
    speechSynthesis.onvoiceschanged = ()=>{ refreshVoices(); };
  }
  $('#btnReloadVoices').addEventListener('click', ()=>{ refreshVoices(); console.log('Verfügbare Stimmen:', state.voices); alert('Stimmen neu geladen. Anzahl: '+state.voices.length); });
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
