/* Flashcards v5.5 – Persistenz + Debug-Panel + kleinere Mobile-Überschrift
   - Excel: ./data/Long-Chinesisch_Lektionen.xlsx (UI override möglich)
   - Kein Countdown/Autoplay; Reveal nur per Button
   - Stimmenwahl via Radio (w/m) mit Heuristik + Rate + Pitch; 1s Delay
   - Session-Stats + Selbstbewertung (Gewusst/Nicht gewusst/Unsicher)
   - Persistenz: Einstellungen + Fortschritt in localStorage; Export/Import (JSON)
*/

let EXCEL_URL = './data/Long-Chinesisch_Lektionen.xlsx';
const SHEET_NAME_PATTERN = /^L\s*\d{1,2}$/i;
const MIN_LESSON = 0, MAX_LESSON = 16;
const DATA_START_ROW = 3;
const COL_WORD = { de:1, py:2, zh:6 };
const COL_SENT = { de:5, py:4, zh:7 };
const COL_POS  = 3;

const LS_KEYS = {
  settings:'fc_settings_v1',
  progress:'fc_progress_v1'
};

const state = {
  mode:'de2zh', order:'random', rate:0.95, pitch:1.0,
  lessons:new Map(), selectedLessons:new Set(), pool:[],
  current:null, idx:null,
  voices:[], voicePref:{ zh:'female', de:'female' },
  // Session-Metriken
  session:{ total:0, done:0, known:0, unsure:0, unknown:0, ttrSum:0, ttrCount:0 },
  startedAt:null, revealedAt:null,
  // Persistente Daten
  settings:{ excelPath:EXCEL_URL, mode:'de2zh', zhVoice:'female', deVoice:'female', rate:0.95, pitch:1.0, lessons:[] },
  progress:{ version:'v1', cards:{}, byLesson:{} },
  foundSheets:[],
};

const $ = s=>document.querySelector(s);

function saveSettings(){ try{ localStorage.setItem(LS_KEYS.settings, JSON.stringify(state.settings)); }catch(e){} }
function loadSettings(){ try{ const s=JSON.parse(localStorage.getItem(LS_KEYS.settings)||'null'); if(s){ state.settings = Object.assign(state.settings, s); } }catch(e){} }
function saveProgressDebounced(){ clearTimeout(saveProgressDebounced._t); saveProgressDebounced._t=setTimeout(()=>{ try{ localStorage.setItem(LS_KEYS.progress, JSON.stringify(state.progress)); }catch(e){} }, 400); }
function loadProgress(){ try{ const p=JSON.parse(localStorage.getItem(LS_KEYS.progress)||'null'); if(p && p.version==='v1'){ state.progress=p; } }catch(e){} }

// ===== Excel laden =====
async function loadExcel(customUrl){
  try{
    EXCEL_URL = customUrl || $('#excelPath').value || state.settings.excelPath || EXCEL_URL;
    state.settings.excelPath = EXCEL_URL; saveSettings();
    const res = await fetch(EXCEL_URL, {cache:'no-store'});
    if(!res.ok) throw new Error('HTTP '+res.status);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});

    state.lessons.clear(); state.foundSheets = [];
    for(const name of wb.SheetNames){
      if(!SHEET_NAME_PATTERN.test(name)) continue;
      const m=name.match(/\d+/); if(!m) continue; const n=parseInt(m[0],10);
      if(!(n>=MIN_LESSON && n<=MAX_LESSON)) continue;
      const sh = wb.Sheets[name]; if(!sh) continue;
      const rows = XLSX.utils.sheet_to_json(sh, {header:1, blankrows:false});
      if(!rows || !rows.length) continue;
      const r0 = DATA_START_ROW-1; const key = String(n);
      if(!state.lessons.has(key)) state.lessons.set(key, []);
      for(let r=r0; r<rows.length; r++){
        const row = rows[r]||[];
        const w = { de:String(row[COL_WORD.de-1]||'').trim(), py:String(row[COL_WORD.py-1]||'').trim(), zh:String(row[COL_WORD.zh-1]||'').trim() };
        const s = { de:String(row[COL_SENT.de-1]||'').trim(), py:String(row[COL_SENT.py-1]||'').trim(), zh:String(row[COL_SENT.zh-1]||'').trim() };
        const pos = String(row[COL_POS-1]||'').trim();
        if(!(w.de||w.zh||s.de||s.zh)) continue;
        state.lessons.get(key).push({word:w, sent:s, pos, id:`L${key.padStart(2,'0')}-${r}`});
      }
      if(state.lessons.get(key).length>0) state.foundSheets.push('L '+key.padStart(2,'0'));
    }
    populateLessonSelect();
    updateLessonCoverageBadges();
    // Debug panel data
    updateDebug();
  }catch(err){ console.error('Excel konnte nicht geladen werden:', err); }
}

function populateLessonSelect(){
  const sel = $('#lessonSelect'); sel.innerHTML='';
  const keys = Array.from(state.lessons.keys()).map(k=>parseInt(k,10)).sort((a,b)=>a-b);
  for(const k of keys){
    const total = state.lessons.get(String(k)).length;
    const seen  = (state.progress.byLesson?.[String(k)]?.seen)||0;
    const opt = document.createElement('option');
    opt.value=String(k); opt.textContent=`Lektion ${k} (${total})` + (seen? ` · ${Math.floor(seen*100/total)}% gesehen`: '');
    // Reapply persisted selection
    if(state.settings.lessons?.includes(String(k))) opt.selected = true;
    sel.appendChild(opt);
  }
}

function updateLessonCoverageBadges(){
  const sel = $('#lessonSelect'); if(!sel) return;
  for(const o of sel.options){ const k=o.value; const total = state.lessons.get(String(k))?.length||0; const seen = (state.progress.byLesson?.[k]?.seen)||0; o.textContent = `Lektion ${k} (${total})` + (seen? ` · ${Math.floor(seen*100/total)}% gesehen`: ''); }
}

function gatherPool(){
  const out=[]; for(const k of state.selectedLessons){ const arr=state.lessons.get(k); if(arr) out.push(...arr); }
  state.pool = out; state.idx = null; state.session = { total:out.length, done:0, known:0, unsure:0, unknown:0, ttrSum:0, ttrCount:0 };
  renderSessionStats();
}

// ===== Karte/Flow =====
function setCard(entry){
  state.current = entry;
  $('#solBox').classList.add('masked');
  $('#evalRow').style.display='none';
  state.startedAt = Date.now(); state.revealedAt = null;

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

function nextCard(){
  if(!state.pool.length){ alert('Bitte Lektionen wählen und übernehmen.'); return; }
  if(state.order==='seq'){
    if(state.idx===null) state.idx=0; else state.idx=(state.idx+1)%state.pool.length; setCard(state.pool[state.idx]);
  } else { const e=state.pool[Math.floor(Math.random()*state.pool.length)]; setCard(e); }
}
function prevCard(){ if(state.order!=='seq'||!state.pool.length) return; if(state.idx===null) state.idx=0; else state.idx=(state.idx-1+state.pool.length)%state.pool.length; setCard(state.pool[state.idx]); }
function startTraining(){ if(!state.pool.length){ alert('Bitte zuerst Lektion(en) übernehmen.'); return; } if(state.order==='seq'){ state.idx=0; setCard(state.pool[state.idx]); } else { const e=state.pool[Math.floor(Math.random()*state.pool.length)]; setCard(e);} }

function doReveal(){
  $('#solBox').classList.remove('masked');
  $('#evalRow').style.display='flex';
  state.revealedAt = Date.now();
  const ttr = state.revealedAt - (state.startedAt||state.revealedAt);
  if(ttr>0){ state.session.ttrSum += ttr; state.session.ttrCount += 1; }
  renderSessionStats();
}

function onEvaluate(mark){
  // mark: 'known' | 'unknown' | 'unsure'
  const id = state.current?.id; if(!id) return;
  const lesson = getLessonOfCurrent();
  const rec = state.progress.cards[id] || { lesson, seen:0, success:0, ttrSum:0, ttrCount:0, playsQ:0, playsA:0 };
  rec.seen += 1; if(mark==='known') rec.success += 1; if(state.revealedAt && state.startedAt){ rec.ttrSum += (state.revealedAt - state.startedAt); rec.ttrCount += 1; }
  state.progress.cards[id] = rec;
  // byLesson
  const bl = state.progress.byLesson[lesson] || { seen:0, success:0, total: state.lessons.get(String(lesson))?.length || 0 };
  bl.seen = Object.values(state.progress.cards).filter(x=>x.lesson===lesson).length; // distinct seen
  bl.success = Object.values(state.progress.cards).filter(x=>x.lesson===lesson && x.success>0).length;
  state.progress.byLesson[lesson] = bl;

  // Session
  state.session.done += 1; state.session[mark] += 1;
  renderSessionStats();
  updateLessonCoverageBadges();
  saveProgressDebounced();
  nextCard();
}

function renderSessionStats(){
  const s = state.session; const avg = s.ttrCount? (s.ttrSum/s.ttrCount/1000).toFixed(1) : '—';
  const acc = s.done? Math.round(100*s.known/s.done)+'%' : '—';
  $('#sessionStats').textContent = `Karten: ${s.done}/${s.total} · Korrekt: ${acc} · Ø Aufdeck‑Zeit: ${avg}s`;
}

function getLessonOfCurrent(){
  if(!state.current?.id) return null; const m=state.current.id.match(/^L(\d{2})-/); return m? String(parseInt(m[1],10)) : null;
}

function formatZh(hz,py){ const hanzi=(hz||'').trim(); const pinyin=(py||'').trim(); return pinyin? `${hanzi}<br><span class="py">${pinyin}</span>` : (hanzi||'—'); }
function formatPinyinAndPos(py,pos){ const a=(py||'').trim(); const b=(pos||'').trim(); if(a && b) return `<span class="py">${a}</span><br><span class=\"prompt small\" style=\"display:inline-block;margin-top:6px;\">${b}</span>`; if(a) return `<span class="py">${a}</span>`; if(b) return `<span class=\"prompt small\" style=\"display:inline-block;margin-top:6px;\">${b}</span>`; return ''; }

// ===== Stimmen (Web Speech API) =====
function refreshVoices(){ state.voices = window.speechSynthesis?.getVoices?.() || []; updateDebugVoices(); }
function matchesLang(v, key){ const L=(v.lang||'').toLowerCase(); if(key==='zh') return L.startsWith('zh')||L.includes('cmn')||L.includes('hans')||L.includes('zh-cn'); if(key==='de') return L.startsWith('de'); return false; }
function genderHint(name){ const s=(name||'').toLowerCase(); const femaleHint = /(female|weib|女|ting|xiao|mei|ya|li|hui|jing|xin)/i; const maleHint   = /(male|männ|男|yun|yang|zhiyu|lei|wei|hao)/i; return femaleHint.test(s)?'female':(maleHint.test(s)?'male':null); }
function pickVoice(langKey, wantedGender){ const list = state.voices.filter(v=> matchesLang(v, langKey)); if(!list.length) return null; const withGender = list.filter(v=> genderHint(v.name+" "+v.voiceURI)===wantedGender); if(withGender.length) return withGender[0]; const def = list.find(v=> v.default) || null; if(def) return def; return list[0]; }

function speak(text, langKey){ const u = new SpeechSynthesisUtterance(text||''); u.lang = (langKey==='zh')? 'zh-CN' : 'de-DE'; u.rate = state.rate; u.pitch = state.pitch; const v = pickVoice(langKey, state.voicePref[langKey]); if(v) u.voice = v; try{ speechSynthesis.cancel(); }catch(e){} speechSynthesis.speak(u); }

const DELAY_MS = 1000;
function playQuestion(){ if(!state.current) return; if(state.mode==='de2zh'){ speak(state.current.word.de,'de'); setTimeout(()=>speak(state.current.sent.de,'de'),DELAY_MS);} else { speak(state.current.word.zh,'zh'); setTimeout(()=>speak(state.current.sent.zh,'zh'),DELAY_MS);} }
function playAnswer(){ if(!state.current) return; if(state.mode==='de2zh'){ speak(state.current.word.zh,'zh'); setTimeout(()=>speak(state.current.sent.zh,'zh'),DELAY_MS);} else { speak(state.current.word.de,'de'); setTimeout(()=>speak(state.current.sent.de,'de'),DELAY_MS);} }

// ===== Export/Import & Debug =====
function exportProgress(){ const blob=new Blob([JSON.stringify(state.progress,null,2)],{type:'application/json'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='progress.json'; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href), 1500); }
function importProgress(file){ const r=new FileReader(); r.onload=()=>{ try{ const p=JSON.parse(r.result); if(p && p.version==='v1'){ state.progress=p; saveProgressDebounced(); updateLessonCoverageBadges(); updateDebug(); alert('Fortschritt importiert.'); } else { alert('Ungültiges Format.'); } }catch(e){ alert('Import fehlgeschlagen: '+e.message);} }; r.readAsText(file); }

function updateDebug(){ $('#dbgSettings').textContent = JSON.stringify(state.settings,null,2); $('#dbgLessons').textContent = state.foundSheets.join(', '); $('#dbgProgress').textContent = JSON.stringify({cards:Object.keys(state.progress.cards).length, byLesson:state.progress.byLesson}, null, 2); updateDebugVoices(); }
function updateDebugVoices(){ const box=$('#dbgVoices'); if(!box) return; box.innerHTML=''; (state.voices||[]).forEach(v=>{ const row=document.createElement('div'); row.className='voice'; const name=document.createElement('div'); name.className='name'; name.textContent=v.name||'(name)'; const meta=document.createElement('div'); meta.className='meta'; meta.textContent=`${v.lang||''} · ${v.default?'default':''} · gender? ${genderHint(v.name||'')||'-'}`; const tests=document.createElement('div'); tests.style.marginLeft='auto'; const b1=document.createElement('button'); b1.className='btn ghost'; b1.textContent='Test DE'; b1.onclick=()=>{ const u=new SpeechSynthesisUtterance('Das ist ein deutscher Stimmtest.'); u.lang='de-DE'; u.voice=v; speechSynthesis.cancel(); speechSynthesis.speak(u); }; const b2=document.createElement('button'); b2.className='btn ghost'; b2.textContent='测试中文'; b2.onclick=()=>{ const u=new SpeechSynthesisUtterance('这是一个中文的语音测试。'); u.lang='zh-CN'; u.voice=v; speechSynthesis.cancel(); speechSynthesis.speak(u); }; tests.appendChild(b1); tests.appendChild(b2); row.appendChild(name); row.appendChild(meta); row.appendChild(tests); box.appendChild(row); }); }

// ===== UI Init =====
window.addEventListener('DOMContentLoaded', ()=>{
  // Load persisted
  loadSettings(); loadProgress();
  // Apply settings to UI
  $('#excelPath').value = state.settings.excelPath || EXCEL_URL;
  document.querySelector(`input[name="mode"][value="${state.settings.mode||'de2zh'}"]`)?.click();
  document.querySelector(`input[name="zhVoiceGender"][value="${state.settings.zhVoice||'female'}"]`)?.click();
  document.querySelector(`input[name="deVoiceGender"][value="${state.settings.deVoice||'female'}"]`)?.click();
  $('#rateRange').value = String(state.settings.rate||0.95); state.rate=Number($('#rateRange').value); $('#rateVal').textContent=`(${state.rate.toFixed(2)})`;
  $('#pitchRange').value = String(state.settings.pitch||1.0); state.pitch=Number($('#pitchRange').value); $('#pitchVal').textContent=`(${state.pitch.toFixed(2)})`;

  // Excel
  $('#btnReloadExcel').addEventListener('click', ()=> loadExcel($('#excelPath').value));
  loadExcel(state.settings.excelPath);

  // Mode & voices
  document.querySelectorAll('input[name="mode"]').forEach(r=> r.addEventListener('change', e=>{ state.mode=e.target.value; state.settings.mode=state.mode; saveSettings(); if(state.current) setCard(state.current); }));
  document.querySelectorAll('input[name="zhVoiceGender"]').forEach(r=> r.addEventListener('change', e=>{ state.voicePref.zh=e.target.value; state.settings.zhVoice=e.target.value; saveSettings(); }));
  document.querySelectorAll('input[name="deVoiceGender"]').forEach(r=> r.addEventListener('change', e=>{ state.voicePref.de=e.target.value; state.settings.deVoice=e.target.value; saveSettings(); }));
  $('#rateRange').addEventListener('input', e=>{ state.rate=parseFloat(e.target.value); state.settings.rate=state.rate; $('#rateVal').textContent=`(${state.rate.toFixed(2)})`; saveSettings(); });
  $('#pitchRange').addEventListener('input', e=>{ state.pitch=parseFloat(e.target.value); state.settings.pitch=state.pitch; $('#pitchVal').textContent=`(${state.pitch.toFixed(2)})`; saveSettings(); });

  // Voices
  refreshVoices(); if('speechSynthesis' in window && typeof speechSynthesis.onvoiceschanged!=='undefined') speechSynthesis.onvoiceschanged = ()=>{ refreshVoices(); };
  $('#btnReloadVoices').addEventListener('click', ()=>{ refreshVoices(); alert('Stimmen neu geladen: '+state.voices.length); });
  $('#btnTestDE').addEventListener('click', ()=> speak('Das ist ein deutscher Stimmtest.','de'));
  $('#btnTestZH').addEventListener('click', ()=> speak('这是一个中文的语音测试。','zh'));

  // Lessons persist
  $('#btnUseLessons').addEventListener('click', ()=>{ const sel=$('#lessonSelect'); state.selectedLessons.clear(); const picked=[]; for(const o of sel.selectedOptions){ state.selectedLessons.add(o.value); picked.push(o.value);} state.settings.lessons=picked; saveSettings(); gatherPool(); $('#btnStart').disabled=(state.pool.length===0); });
  $('#btnClearLessons').addEventListener('click', ()=>{ const sel=$('#lessonSelect'); for(const o of sel.options) o.selected=false; state.selectedLessons.clear(); state.settings.lessons=[]; saveSettings(); gatherPool(); $('#btnStart').disabled=true; });

  // Flow
  $('#btnStart').addEventListener('click', startTraining);
  $('#btnNext').addEventListener('click', nextCard);
  $('#btnPrev').addEventListener('click', prevCard);
  $('#btnReveal').addEventListener('click', doReveal);
  $('#btnPlayQ').addEventListener('click', playQuestion);
  $('#btnPlayA').addEventListener('click', playAnswer);

  // Bewertung
  $('#btnKnow').addEventListener('click', ()=> onEvaluate('known'));
  $('#btnDunno').addEventListener('click', ()=> onEvaluate('unknown'));
  $('#btnUnsure').addEventListener('click', ()=> onEvaluate('unsure'));

  // Export/Import
  $('#btnExport').addEventListener('click', exportProgress);
  $('#fileImport').addEventListener('change', e=>{ const f=e.target.files?.[0]; if(f) importProgress(f); e.target.value=''; });

  // Debug
  $('#btnDebug').addEventListener('click', ()=>{ updateDebug(); $('#debugPanel').classList.remove('hidden'); });
  $('#btnCloseDebug').addEventListener('click', ()=> $('#debugPanel').classList.add('hidden'));
});
