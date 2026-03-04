/* Tabellen-Trainer: zeigt Spalten 1..7 aus Excel je Zeile.
   - Lektion wählbar (Tabs nach Namen L 00..L 16)
   - Klick auf Wort/Satz: TTS (DE oder ZH)
   - Übersetzung ist je nach Richtung versteckt und wird erst nach Klick angezeigt
   - TTS-Stimme (m/w) + Rate einstellbar
   - Excel-Parsing im Browser via SheetJS
*/

// ===== KONFIG =====
const EXCEL_URL = 'https://luoboteanjing-netizen.github.io/Chinesisch-Reader/data/Long-Chinesisch_Lektionen.xlsx';
const SHEET_NAME_PATTERN = /^L\s*\d{1,2}$/i; // L 00 .. L 16
const MIN_LESSON = 0, MAX_LESSON = 16;
const DATA_START_ROW = 3; // 1-basiert

// Spaltenindizes (1-basiert) -> wir zeigen alle 1..7, markieren aber DE/ZH Rollen
const COLS = { c1:1, c2:2, c3:3, c4:4, c5:5, c6:6, c7:7 };
// Bedeutung: Wort: c1=DE, c2=PY, c6=HZ | Satz: c5=DE, c4=PY, c7=HZ

// ===== STATE =====
const state = {
  mode:'de2zh', // de2zh oder zh2de
  recognizing:false, recog:null, rate:0.95,
  voicePref:{ zh:'female', de:'female' }, voices:[],
  lessons:new Map(), selectedLessons:new Set(),
};

const $ = (s)=>document.querySelector(s);

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
      const r0=DATA_START_ROW-1;
      const lessonKey=String(n);
      if(!state.lessons.has(lessonKey)) state.lessons.set(lessonKey,[]);

      for(let r=r0;r<rows.length;r++){
        const row=rows[r]||[];
        // Speichere 1..7, auch wenn leer
        const entry={
          c1: String(row[COLS.c1-1]||'').trim(),
          c2: String(row[COLS.c2-1]||'').trim(),
          c3: String(row[COLS.c3-1]||'').trim(),
          c4: String(row[COLS.c4-1]||'').trim(),
          c5: String(row[COLS.c5-1]||'').trim(),
          c6: String(row[COLS.c6-1]||'').trim(),
          c7: String(row[COLS.c7-1]||'').trim()
        };
        // wenn komplett leer: überspringen
        if(!entry.c1 && !entry.c2 && !entry.c3 && !entry.c4 && !entry.c5 && !entry.c6 && !entry.c7) continue;
        state.lessons.get(lessonKey).push(entry);
      }
    }

    populateLessonSelect();
    status.textContent = `Excel geladen (${state.lessons.size} Lektion(en)).`;
  }catch(err){ status.textContent='Excel konnte nicht geladen werden: '+err.message; console.error(err); }
}

function populateLessonSelect(){
  const sel=$('#lessonSelect'); sel.innerHTML='';
  const keys=Array.from(state.lessons.keys()).map(k=>parseInt(k,10)).sort((a,b)=>a-b);
  for(const k of keys){ const o=document.createElement('option'); o.value=String(k); o.textContent=`Lektion ${k} (${state.lessons.get(String(k)).length})`; sel.appendChild(o); }
}

// ===== Tabelle aufbauen =====
function buildTable(){
  const body=$('#dataBody'); body.innerHTML='';
  const sel=state.selectedLessons; if(!sel.size) return;

  // Welche Spalten sind je Richtung maskiert?
  const maskZH = (state.mode==='de2zh'); // ZH-Spalten maskieren: 2,4,6,7
  const maskDE = (state.mode==='zh2de'); // DE-Spalten maskieren: 1,5

  for(const k of sel){
    const rows = state.lessons.get(k)||[];
    for(const e of rows){
      const tr=document.createElement('tr');
      // Helfer zum Erzeugen einer Zelle
      const mk=(text,colIdx)=>{
        const td=document.createElement('td');
        td.className='cell';
        td.dataset.col = String(colIdx);
        td.textContent = text || '';

        // Sprache bestimmen + Maskierung + Klick‑Verhalten
        const isDE = (colIdx===1 || colIdx===5);
        const isZH = (colIdx===6 || colIdx===7); // Hanzi
        const isPY = (colIdx===2 || colIdx===4);

        // Maskieren abhängig von Richtung
        const shouldMask = (maskDE && isDE) || (maskZH && (isZH || isPY));
        if(shouldMask){ td.classList.add('masked'); }

        // Klickverhalten:
        td.classList.add('clickable');
        td.addEventListener('click', ()=>{
          if(td.classList.contains('masked')){
            td.classList.remove('masked');
            return;
          }
          // Vorlesen
          if(isDE){ speak(text, 'de-DE'); }
          else if(isZH){ speak(text, 'zh-CN'); }
          else if(isPY){
            // Pinyin nicht gut von TTS unterstützt: versuche passendes Hanzi in gleicher Zeile
            const hanzi = (colIdx===2? e.c6 : e.c7) || '';
            if(hanzi) speak(hanzi, 'zh-CN');
          }
        });

        return td;
      };

      tr.appendChild(mk(e.c1,1));
      tr.appendChild(mk(e.c2,2));
      tr.appendChild(mk(e.c3,3));
      tr.appendChild(mk(e.c4,4));
      tr.appendChild(mk(e.c5,5));
      tr.appendChild(mk(e.c6,6));
      tr.appendChild(mk(e.c7,7));
      body.appendChild(tr);
    }
  }
}

// ===== TTS =====
function refreshVoices(){ state.voices = window.speechSynthesis?.getVoices?.() || []; }
function pickVoice(lang, gender){ if(!state.voices.length) return null; const list = state.voices.filter(v=> (v.lang||'').toLowerCase().startsWith(lang)); if(!list.length) return null; const want=(gender||'').toLowerCase(); const isF=s=>/female|weib|女/i.test(s); const isM=s=>/male|männ|男/i.test(s); const byName=list.filter(v=> want==='female'? isF(v.name+" "+v.voiceURI) : isM(v.name+" "+v.voiceURI)); return byName[0] || list.find(v=>v.default) || list[0]; }
function speak(text, lang){ const u=new SpeechSynthesisUtterance(text||''); u.lang=lang; u.rate=state.rate; const v = lang.startsWith('zh')? pickVoice('zh', state.voicePref.zh) : pickVoice('de', state.voicePref.de); if(v) u.voice=v; speechSynthesis.cancel(); speechSynthesis.speak(u); }

// ===== UI =====
window.addEventListener('DOMContentLoaded',()=>{
  refreshVoices(); if('speechSynthesis' in window && typeof speechSynthesis.onvoiceschanged!=='undefined') speechSynthesis.onvoiceschanged = refreshVoices;
  loadExcel().then(()=>{
    // Default: alle Lektionen auswählen
    const sel=$('#lessonSelect'); for(const opt of sel.options) opt.selected=true; commitLessons(); buildTable();
  });

  document.querySelectorAll('input[name="mode"]').forEach(r=> r.addEventListener('change', e=>{ state.mode=e.target.value; buildTable(); }));
  document.querySelectorAll('input[name="zhVoice"]').forEach(r=> r.addEventListener('change', e=> state.voicePref.zh=e.target.value));
  document.querySelectorAll('input[name="deVoice"]').forEach(r=> r.addEventListener('change', e=> state.voicePref.de=e.target.value));
  $('#rateRange').addEventListener('input', e=>{ state.rate=parseFloat(e.target.value); $('#rateVal').textContent = `(${state.rate.toFixed(2)})`; });
  $('#btnUseLessons').addEventListener('click', ()=>{ commitLessons(); buildTable(); });
  $('#btnClearLessons').addEventListener('click', ()=>{ const s=$('#lessonSelect'); for(const o of s.options) o.selected=false; state.selectedLessons.clear(); buildTable(); });
});

function commitLessons(){
  const sel=$('#lessonSelect'); state.selectedLessons.clear();
  for(const o of sel.selectedOptions) state.selectedLessons.add(o.value);
}
