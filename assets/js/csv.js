/* CSV loading, lesson list, card pool */
import { state, $, translate, saveSettings } from "./state.js";

// CSV-Datei dynamisch über URL-Parameter auswählen
const params = new URLSearchParams(location.search);
const csvParam = params.get("csv");

// 🔥 CSV wird jetzt dynamisch zur Laufzeit bestimmt
let CSV_URL = null;

/* ===========================
   CSV RESOLVER (mit Fallback)
   =========================== */
async function resolveCSV() {

    // 1. URL-Parameter hat Priorität
    if (csvParam) {
        const file = `./data/${csvParam}`;
        try {
            const res = await fetch(file, { method: "HEAD" });
            if (res.ok) return file;

            console.warn("CSV aus URL nicht gefunden → Fallback wird verwendet");
        } catch (e) {}
    }

    // 2. Fallback-Reihenfolge
    const candidates = [
        "./data/HSK-Chinesisch_Lektionen.csv",
        "./data/Long-Chinesisch_Lektionen.csv"
    ];

    for (const file of candidates) {
        try {
            const res = await fetch(file, { method: "HEAD" });
            if (res.ok) return file;
        } catch (e) {}
    }
 
    // 3. Harte Fehlerbehandlung
    throw new Error("Keine CSV-Datei gefunden");
}

/* ============================ CSV PARSING ================================= */

function parseCSVLine(line) {
    const result = [];
    let cur = "";
    let quotes = false;

    for (let i = 0; i < line.length; i++) {
        const c = line[i];

        if (c === '"') {
            if (quotes && line[i + 1] === '"') {
                cur += '"';
                i++;
            } else {
                quotes = !quotes;
            }
        } else if (c === ";" && !quotes) {
            result.push(cur);
            cur = "";
        } else {
            cur += c;
        }
    }

    result.push(cur);
    return result;
}

async function loadCSV() {
    try {
		 CSV_URL = await resolveCSV();
        const res = await fetch(CSV_URL);
        const buf = await res.arrayBuffer();
        const text = new TextDecoder("utf-8").decode(buf);

        parseCSV(text);
        // Save the original order so we can restore it without re-fetching
        state.originalLessonOrder = [...state.lessonOrder];
        populateLessonSelect();
    } catch (e) {
alert(translate("csvLoadError"));
        console.error(e);
    }
}

function parseCSV(text) {

    state.lessons.clear();
    state.lessonOrder = [];

    const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
    if (lines.length < 2) return;

    for (let i = 1; i < lines.length; i++) {

        const cols = parseCSVLine(lines[i]);
        if (cols.length < 9) continue;

        // ✅ Skip disabled/commented rows (starting with "*")
        const firstCell = (cols[0] || "").replace(/\uFEFF/g, "").trim();
        if (firstCell.startsWith("*")) continue;

        // ✅ Extract lesson name
        const lesson = (cols[8] || "")
            .replace(/\uFEFF/g, "")
            .trim();

        const entry = {
            word: {
                de: (cols[0] || "").trim(),
                py: (cols[1] || "").trim(),
                zh: (cols[5] || "").trim()
            },
            pos: (cols[2] || "").trim(),
            sent: {
                py: (cols[3] || "").trim(),
                de: (cols[4] || "").trim(),
                zh: (cols[6] || "").trim()
            },
            id: (cols[7] || "").trim(),
            lesson,
            audio: {
                wordFile: (cols[9] || "").trim() || ((cols[7] || "").trim() + '_wrd.mp3'),
                sentFile: (cols[10] || "").trim() || ((cols[7] || "").trim() + '_snt.mp3')
            }
        };

        // ✅ Insert lesson & entry into the map
        if (!state.lessons.has(lesson)) {
            state.lessons.set(lesson, []);
            state.lessonOrder.push(lesson);
        }

        state.lessons.get(lesson).push(entry);
    }
}

/* ============================ LESSON SELECT =============================== */

function populateLessonSelect() {
    const sel = $("#lessonSelect");
    const table = $("#lessonTable");

    sel.innerHTML = "";
    table.innerHTML = "";


const header = `
    <div class="lt-row lt-head">
        <span id="lessonTableHeaderLesson" class="lt-lesson" data-sort="lesson">${translate("lessonTableLesson")}</span>
        <span id="lessonTableHeaderCards" class="lt-total" data-sort="total">${translate("lessonTableCards")}</span>
        <span class="lt-strong"  data-sort="strong">✅</span>    <!-- Box 4+5 -->
        <span class="lt-weak"    data-sort="weak">🤔</span>      <!-- Box 2+3 -->
		<span class="lt-unknown" data-sort="unknown">❌</span>   <!-- Box 1 -->
	</div>
`;

    table.insertAdjacentHTML("beforeend", header);

    for (const k of state.lessonOrder) {

        const cards = state.lessons.get(k) || [];
        const total = cards.length;

        const p = state.progress.byLesson[k] || { known: 0, unknown: 0 };
        const known   = p.known   || 0;
        const unknown = p.unknown || 0;
        const percent = total > 0 ? Math.round((known / total) * 100) : 0;

        // Unter der Haube weiter Optionen befüllen (für Training)
        const opt = document.createElement("option");
        opt.value = k;
        sel.appendChild(opt);

        const row = document.createElement("div");
        row.className = "lt-row";
        row.dataset.lesson = k;

        row.innerHTML = `
            <span class="lt-lesson">${k}</span>
            <span class="lt-total">${total}</span>
			<span class="lt-strong">0</span>
			<span class="lt-weak">0</span>
			<span class="lt-unknown">0</span>
        `;

row.addEventListener("click", () => {
    // ✅ Alle Optionen abwählen
    [...sel.options].forEach(o => o.selected = false);
    document
        .querySelectorAll(".lt-row.selected")
        .forEach(r => r.classList.remove("selected"));

    // ✅ Diese Lektion auswählen
    opt.selected = true;
    row.classList.add("selected");

    // ✅ Genau eine Lektion speichern
    state.settings.lessons = [opt.value];
    saveSettings();

    // ✅ Pool neu bauen
    gatherPoolFromSettings();

    // ✅ Falls Training läuft → direkt neu laden
    if (state.trainingOn) {
        state.idx = null;
        resetSessionStats();
        if (state.pool.length) {
            state.setCard?.(state.pool[0]);
        }
    }
});

        table.appendChild(row);

        // vorauswahl anzeigen
        if (state.settings.lessons.includes(k)) {
            opt.selected = true;
            row.classList.add("selected");
        }
    }
}


let lessonSort = { key: null, asc: true };

/* ============================================================
   SORTIERUNG FÜR LEKTIONSTABELLE
   ============================================================ */
document.addEventListener("click", (ev) => {
    const sortKey = ev.target.dataset.sort;
    if (!sortKey) return;

    // Beim Klick auf "Lektion" → original CSV-Reihenfolge wiederherstellen (ohne Netzwerk-Request)
    if (sortKey === "lesson") {
        state.lessonOrder = [...state.originalLessonOrder];
        lessonSort = { key: null, asc: true };
        populateLessonSelect();
        return;
    }

    // ✅ Alle anderen Spalten sortieren wie bisher
    lessonSort.asc = (lessonSort.key === sortKey) ? !lessonSort.asc : true;
    lessonSort.key = sortKey;

    sortLessons();          // sortiert lessonOrder anhand der gewählten Spalte
    populateLessonSelect(); // Liste neu aufbauen
});

function getLessonStats(lessonName) {
    const cards = state.lessons.get(lessonName) || [];
    let strong = 0;   // Box 4+5, same as the ✅ column
    let weak = 0;     // Box 2+3, same as the 🤔 column
    let unknown = 0;  // Box 1, same as the ❌ column

    for (const card of cards) {
        const box = state.progress.cards[card.id]?.box || 0;
        if (box === 1) unknown++;
        else if (box === 2 || box === 3) weak++;
        else if (box === 4 || box === 5) strong++;
    }

    return {
        lesson: lessonName,
        total: cards.length,
        strong,
        weak,
        unknown
    };
}

/* ============================ POOL HANDLING =============================== */

function resetSessionStats() {
    state.session = {
        total: state.pool.length,
        done: 0,
        known: 0,
        unsure: 0,
        unknown: 0,
        ttrSum: 0,
        ttrCount: 0,
        revealedCount: state.session.revealedCount ?? 0,
        revealedCardIds: state.session.revealedCardIds ?? []
    };
}

function gatherPool() {
    const out = [];
    for (const k of state.selectedLessons) {
        const arr = state.lessons.get(k);
        if (arr) out.push(...arr);
    }
    state.pool = out;
    state.idx = null;
    // resetSessionStats();  // Entfernt - wird jetzt in startTraining() gemacht
}

function gatherPoolFromSettings() {
    state.selectedLessons.clear();
    (state.settings.lessons || []).forEach((x) => state.selectedLessons.add(x));
    gatherPool();
}

function sortLessons() {
    const key = lessonSort.key;
    if (!key) return;

    state.lessonOrder.sort((a, b) => {
        const A = getLessonStats(a);
        const B = getLessonStats(b);

        let vA = A[key];
        let vB = B[key];

        if (typeof vA === "string") vA = vA.toLowerCase();
        if (typeof vB === "string") vB = vB.toLowerCase();

        if (vA < vB) return lessonSort.asc ? -1 : 1;
        if (vA > vB) return lessonSort.asc ? 1 : -1;
        return 0;
    });
}


export {
    resolveCSV,
    parseCSVLine,
    loadCSV,
    parseCSV,
    populateLessonSelect,
    getLessonStats,
    resetSessionStats,
    gatherPool,
    gatherPoolFromSettings
};
