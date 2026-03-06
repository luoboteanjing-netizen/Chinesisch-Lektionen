/* v5.5 Browser-optimiert: Stimmen-Panel, ZH-Filter, Persistenz konkreter Browserstimmen, Start-/Between-Delay, ohne Testbuttons im Haupt-UI */

let EXCEL_URL = './data/Long-Chinesisch_Lektionen.xlsx';
const SHEET_NAME_PATTERN = /^L\s*\d{1,2}$/i; const MIN_LESSON=0, MAX_LESSON=16; const DATA_START_ROW=3;
const COL_WORD={de:1,py:2,zh:6}; const COL_SENT={de:5,py:4,zh:7}; const COL_POS=3;

const LS_KEYS={ settings:'fc_settings_v1', progress:'fc_progress_v1' };

const state={
  mode:'de2zh', rate:0.95, pitch:1.0,
  lessons:new Map(), selectedLessons:new Set(), pool:[], idx:null, current:null,
  voices:[], browserVoice:{ zh:null, de:null },
  settings:{ excelPath:'./data/Long-Chinesisch_Lektionen.xlsx', mode:'de2zh', rate:0.95, pitch:1.0, lessons:[], browserVoiceZh:null, browserVoiceDe:null },
  session:{ total:0, done:0, known:0, unsure:0, unknown:0, ttrSum:0, ttrCount:0 },
  startedAt:null, revealedAt:null,
  progress:{ version:'v1', cards:{}, byLesson:{} }
};

const $=s=>document.querySelector(s);

function saveSettings(){ try{ localStorage.setItem(LS_KEYS.settings, JSON.stringify(state.settings)); }catch(e){} }
function loadSettings(){ try{ const s=JSON.parse(localStorage.getItem(LS_KEYS.settings)||'null'); if(s){ state.settings=Object.assign(state.settings,s); } }catch(e){} }
function saveProgressDebounced(){ clearTimeout(saveProgressDebounced._t); saveProgressDebounced._t=setTimeout(()=>{ try{ localStorage.setItem(LS_KEYS.progress, JSON.stringify(state.progress)); }catch(e){} }, 400); }
function loadProgress(){ try{ const p=JSON.parse(localStorage.getItem(LS_KEYS.progress)||'null'); if(p && p.version==='v1'){ state.progress=p; } }catch(e){} }

function isZhVoice(v){ const L=(v.lang||'').toLowerCase(); return L.startsWith('zh')||L.includes('cmn')||L.includes('hans')||L.includes('zh-cn'); }

function updateVoiceList(){ const box=$('#dbgVoices'); if(!box) return; const zhOnly=$('#dbgZhOnly')?.checked||false; box.innerHTML=''; const list=(state.voices||[]).filter(v=>!zhOnly||isZhVoice(v)); if(list.length===0){ box.innerHTML='<div class="meta">Keine Stimmen gefunden. Bitte später erneut öffnen oder Seite neu laden.</div>'; return; }
  list.forEach(v=>{ const row=document.createElement('div'); row.className='voice'; const name=document.createElement('div'); name.className='name'; name.textContent=v.name||'(name)'; const meta=document.createElement('div'); meta.className='meta'; meta.textContent=`${v.lang||''} ${v.default?'· default':''}`; const actions=document.createElement('div'); actions.style.marginLeft='auto'; actions.style.display='flex'; actions.style.gap='6px'; actions.style.flexWrap='wrap';
    const bZh=document.createElement('button'); bZh.className='btn'; bZh.textContent='Als ZH setzen'; bZh.onclick=()=>{ state.browserVoice.zh=v; state.settings.browserVoiceZh=v.name||v.voiceURI; saveSettings(); updateVoiceList(); };
    const bDe=document.createElement('button'); bDe.className='btn'; bDe.textContent='Als DE setzen'; bDe.onclick=()=>{ state.browserVoice.de=v; state.settings.browserVoiceDe=v.name||v.voiceURI; saveSettings(); updateVoiceList(); };
    const tDe=document.createElement('button'); tDe.className='btn ghost'; tDe.textContent='Test DE'; tDe.onclick=()=>{ const u=new SpeechSynthesisUtterance('Das ist ein deutscher Stimmtest.'); u.lang='de-DE'; u.voice=v; try{speechSynthesis.cancel();}catch(e){} speechSynthesis.speak(u); };
    const tZh=document.createElement('button'); tZh.className='btn ghost'; tZh.textContent='测试中文'; tZh.onclick=()=>{ const u=new SpeechSynthesisUtterance('这是一个中文的语音测试。'); u.lang='zh-CN'; u.voice=v; try{speechSynthesis.cancel();}catch(e){} speechSynthesis.speak(u); };
    if(state.browserVoice.zh && (state.browserVoice.zh.name===v.name||state.browserVoice.zh.voiceURI===v.voiceURI)) name.textContent+='  •  [Aktiv: ZH]';
    if(state.browserVoice.de && (state.browserVoice.de.name===v.name||state.browserVoice.de.voiceURI===v.voiceURI)) name.textContent+='  •  [Aktiv: DE]';
    actions.appendChild(bZh); actions.appendChild(bDe); actions.appendChild(tDe); actions.appendChild(tZh);
    row.appendChild(name); row.appendChild(meta); row.appendChild(actions); box.appendChild(row); });
}

function refreshVoices(){ state.voices = window.speechSynthesis?.getVoices?.() || []; if(state.settings.browserVoiceZh){ const vz=state.voices.find(x=>x.name===state.settings.browserVoiceZh||x.voiceURI===state.settings.browserVoiceZh); if(vz) state.browserVoice.zh=vz; } if(state.settings.browserVoiceDe){ const vd=state.voices.find(x=>x.name===state.settings.browserVoiceDe||x.voiceURI===state.settings.browserVoiceDe); if(vd) state.browserVoice.de=vd; } updateVoiceList(); }

let _voicesRetryT; function openVoicesPanel(){ refreshVoices(); if(!state.voices || state.voices.length===0){ clearTimeout(_voicesRetryT); let tries=0; const tick=()=>{ tries++; refreshVoices(); if(state.voices.length>0 || tries>=8) return; _voicesRetryT=setTimeout(tick,300); }; _voicesRetryT=setTimeout(tick,300); } $('#voicePanel').classList.remove('hidden'); }

// Excel laden
async function loadExcel(){ try{ EXCEL_URL = $('#excelPath').value || EXCEL_URL; const res=await fetch(EXCEL_URL,{cache:'no-store'}); const buf=await res.arrayBuffer(); const wb=XLSX.read(buf,{type:'array'}); state.lessons.clear(); for(const name of wb.SheetNames){ if(!SHEET_NAME_PATTERN.test(name)) continue; const m=name.match(/\d+/); if(!m) continue; const n=parseInt(m[0],10); if(!(n>=MIN_LESSON&&n<=MAX_LESSON)) continue; const sh=wb.Sheets[name]; const rows=XLSX.utils.sheet_to_json(sh,{header:1,blankrows:false}); const r0=DATA_START_ROW-1; const key=String(n); if(!state.lessons.has(key)) state.lessons.set(key,[]); for(let r=r0;r<rows.length;r++){ const row=rows[r]||[]; const w={de:String(row[COL_WORD.de-1]||'').trim(), py:String(row[COL_WORD.py-1]||'').trim(), zh:String(row[COL_WORD.zh-1]||'').trim()}; const s={de:String(row[COL_SENT.de-1]||'').trim(), py:String(row[COL_SENT.py-1]||'').trim(), zh:String(row[COL_SENT.zh-1]||'').trim()}; const pos=String(row[COL_POS-1]||'').trim(); if(!(w.de||w.zh||s.de||s.zh)) continue; state.lessons.get(key).push({word:w,sent:s,pos}); } } populateLessonSelect(); }catch(e){ console.error('Excel konnte nicht geladen werden:',e); } }

function populateLessonSelect(){ const sel=$('#lessonSelect'); sel.innerHTML=''; const keys=Array.from(state.lessons.keys()).map(k=>parseInt(k,10)).sort((a,b)=>a-b); for(const k of keys){ const total=state.lessons.get(String(k)).length; const seen=(state.progress.byLesson?.[String(k)]?.seen)||0; const opt=document.createElement('option'); opt.value=String(k); opt.textContent=`Lektion ${k} (${total})`+(seen?` · ${Math.floor(seen*100/total)}% gesehen`: ''); if(state.settings.lessons?.includes(String(k))) opt.selected=true; sel.appendChild(opt); } }

function gatherPool(){ const out=[]; for(const k of state.selectedLessons){ const arr=state.lessons.get(k); if(arr) out.push(...arr); } state.pool=out; state.idx=null; state.session={ total:out.length, done:0, known:0, unsure:0, unknown:0, ttrSum:0, ttrCount:0 }; renderSessionStats(); }

function setCard(entry){ state.current=entry; $('#solBox').classList.add('masked'); state.startedAt=Date.now(); state.revealedAt=null; if(state.mode==='zh2de'){ $('#promptWord').innerHTML=(entry.word.zh||'—'); $('#promptWordSub').innerHTML=formatPinyinAndPos(entry.word.py, entry.pos); $('#promptSent').innerHTML=formatZh(entry.sent.zh, entry.sent.py); $('#solWord').textContent=entry.word.de||'—'; $('#solSent').textContent=entry.sent.de||'—'; } else { $('#promptWord').textContent=entry.word.de||'—'; $('#promptWordSub').textContent=entry.pos?entry.pos:''; $('#promptSent').textContent=entry.sent.de||'—'; $('#solWord').innerHTML=formatZh(entry.word.zh, entry.word.py); $('#solSent').innerHTML=formatZh(entry.sent.zh, entry.sent.py); } $('#btnNext').disabled=false; $('#btnReveal').disabled=false; $('#btnPlayQ').disabled=false; $('#btnPlayA').disabled=false; }

function nextCard(){ if(!state.pool.length) return alert('Bitte Lektionen wählen und übernehmen.'); if(state.idx==null) state.idx=0; else state.idx=(state.idx+1)%state.pool.length; setCard(state.pool[state.idx]); }
function prevCard(){ if(!state.pool.length) return; if(state.idx==null) state.idx=0; else state.idx=(state.idx-1+state.pool.length)%state.pool.length; setCard(state.pool[state.idx]); }
function startTraining(){ const sel=$('#lessonSelect'); state.selectedLessons.clear(); const picked=[]; for(const o of sel.selectedOptions){ state.selectedLessons.add(o.value); picked.push(o.value); } state.settings.lessons=picked; saveSettings(); gatherPool(); if(!state.pool.length){ alert('Bitte zuerst Lektion(en) übernehmen.'); return; } state.idx=0; setCard(state.pool[state.idx]); }

function doReveal(){ $('#solBox').classList.remove('masked'); $('#evalRow')?.style?.setProperty('display','flex'); state.revealedAt=Date.now(); const ttr=state.revealedAt-(state.startedAt||state.revealedAt); if(ttr>0){ state.session.ttrSum+=ttr; state.session.ttrCount+=1; } renderSessionStats(); }

function formatZh(hz,py){ const h=(hz||'').trim(); const p=(py||'').trim(); return p? `${h}<br><span class="py">${p}</span>` : (h||'—'); }
function formatPinyinAndPos(py,pos){ const a=(py||'').trim(); const b=(pos||'').trim(); if(a&&b) return `<span class="py">${a}</span><br><span class="prompt small" style="display:inline-block;margin-top:6px;">${b}</span>`; if(a) return `<span class="py">${a}</span>`; if(b) return `<span class="prompt small" style="display:inline-block;margin-top:6px;">${b}</span>`; return ''; }

const START_DELAY_MS=180; const BETWEEN_DELAY_MS=250; let _ttsPrimed=false; function ttsPrime(cb){ if(_ttsPrimed){ cb(); return; } setTimeout(()=>{ _ttsPrimed=true; cb(); }, START_DELAY_MS); }
function ttsSpeak(text, langKey){ const lang=(langKey==='zh')?'zh-CN':'de-DE'; const u=new SpeechSynthesisUtterance(text||''); u.lang=lang; u.rate=state.rate; u.pitch=state.pitch; const chosen=(langKey==='zh')?state.browserVoice.zh:state.browserVoice.de; if(chosen) u.voice=chosen; else { const L=(langKey==='zh')?'zh':'de'; const cand=(state.voices||[]).filter(v=>(v.lang||'').toLowerCase().startsWith(L)); u.voice=cand.find(v=>v.default)||cand[0]||null; } try{ speechSynthesis.cancel(); }catch(e){} speechSynthesis.speak(u); }
function playQuestion(){ if(!state.current) return; if(state.mode==='de2zh'){ ttsPrime(()=>{ ttsSpeak(state.current.word.de,'de'); setTimeout(()=>ttsSpeak(state.current.sent.de,'de'), BETWEEN_DELAY_MS); }); } else { ttsPrime(()=>{ ttsSpeak(state.current.word.zh,'zh'); setTimeout(()=>ttsSpeak(state.current.sent.zh,'zh'), BETWEEN_DELAY_MS); }); } }
function playAnswer(){ if(!state.current) return; if(state.mode==='de2zh'){ ttsPrime(()=>{ ttsSpeak(state.current.word.zh,'zh'); setTimeout(()=>ttsSpeak(state.current.sent.zh,'zh'), BETWEEN_DELAY_MS); }); } else { ttsPrime(()=>{ ttsSpeak(state.current.word.de,'de'); setTimeout(()=>ttsSpeak(state.current.sent.de,'de'), BETWEEN_DELAY_MS); }); } }

function renderSessionStats(){ const s=state.session; const avg=s.ttrCount? (s.ttrSum/s.ttrCount/1000).toFixed(1) : '—'; const acc=s.done? Math.round(100*s.known/s.done)+'%' : '—'; $('#sessionStats').textContent=`Karten: ${s.done}/${s.total} · Korrekt: ${acc} · Ø Aufdeck‑Zeit: ${avg}s`; }

window.addEventListener('DOMContentLoaded', ()=>{
  loadSettings(); loadProgress();
  $('#excelPath').value = state.settings.excelPath || EXCEL_URL;
  document.querySelector(`input[name="mode"][value="${state.settings.mode||'de2zh'}"]`)?.click();
  $('#rateRange').value=String(state.settings.rate||0.95); state.rate=Number($('#rateRange').value); $('#rateVal').textContent=`(${state.rate.toFixed(2)})`;
  $('#pitchRange').value=String(state.settings.pitch||1.0); state.pitch=Number($('#pitchRange').value); $('#pitchVal').textContent=`(${state.pitch.toFixed(2)})`;

  $('#btnReloadExcel').addEventListener('click', ()=> loadExcel());
  loadExcel();

  document.querySelectorAll('input[name="mode"]').forEach(r=> r.addEventListener('change', e=>{ state.mode=e.target.value; state.settings.mode=state.mode; saveSettings(); if(state.current) setCard(state.current); }));
  $('#rateRange').addEventListener('input', e=>{ state.rate=parseFloat(e.target.value); state.settings.rate=state.rate; $('#rateVal').textContent=`(${state.rate.toFixed(2)})`; saveSettings(); });
  $('#pitchRange').addEventListener('input', e=>{ state.pitch=parseFloat(e.target.value); state.settings.pitch=state.pitch; $('#pitchVal').textContent=`(${state.pitch.toFixed(2)})`; saveSettings(); });

  // Stimmenpanel
  $('#btnVoices')?.addEventListener('click', openVoicesPanel);
  $('#btnCloseVoices')?.addEventListener('click', ()=> $('#voicePanel').classList.add('hidden'));
  $('#dbgZhOnly')?.addEventListener('change', ()=> updateVoiceList());

  // Voices laden + Event abonnieren
  refreshVoices(); if('speechSynthesis' in window && typeof speechSynthesis.onvoiceschanged!=='undefined'){ speechSynthesis.onvoiceschanged=()=>{ refreshVoices(); }; }

  // Flow
  $('#btnStart').addEventListener('click', startTraining);
  $('#btnNext').addEventListener('click', nextCard);
  $('#btnPrev').addEventListener('click', prevCard);
  $('#btnReveal').addEventListener('click', doReveal);
  $('#btnPlayQ').addEventListener('click', playQuestion);
  $('#btnPlayA').addEventListener('click', playAnswer);

  // Export/Import
  $('#btnExport').addEventListener('click', ()=>{ const blob=new Blob([JSON.stringify(state.progress,null,2)],{type:'application/json'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='progress.json'; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1500); });
  $('#fileImport').addEventListener('change', e=>{ const f=e.target.files?.[0]; if(!f) return; const r=new FileReader(); r.onload=()=>{ try{ const p=JSON.parse(r.result); if(p && p.version==='v1'){ state.progress=p; saveProgressDebounced(); populateLessonSelect(); alert('Fortschritt importiert.'); } else alert('Ungültiges Format.'); }catch(err){ alert('Import fehlgeschlagen: '+err.message); } }; r.readAsText(f); e.target.value=''; });
});
