/* Leitner box progress */
import { state } from "./state.js";

// ✅ Fortschritt in der Lektionstabelle live aktualisieren

function updateLessonStatsUI() {
    document.querySelectorAll(".lt-row:not(.lt-head)").forEach(row => {
        const lesson = row.dataset.lesson;
        const cards = state.lessons.get(lesson) ?? [];

        let red = 0;      // Box 1 (falsch/schwach)
        let yellow = 0;   // Box 2 + 3 (unsicher)
        let green = 0;    // Box 4 + 5 (sicher)

        for (const c of cards) {
            const p = state.progress.cards[c.id] ?? { box: 0 };

            if (p.box === 1) red++;
            else if (p.box === 2 || p.box === 3) yellow++;
            else if (p.box === 4 || p.box === 5) green++;
        }

        row.querySelector(".lt-total").textContent   = cards.length;
        row.querySelector(".lt-unknown").textContent = red;
        row.querySelector(".lt-weak").textContent    = yellow;
        row.querySelector(".lt-strong").textContent  = green;
    });
}

function getLeitnerAscii(box) {
    // Box 0–5 → 0–5 gefüllte Kästchen
    const filled = Math.max(0, Math.min(box, 5));
    return "■".repeat(filled) + "□".repeat(5 - filled);
}

// =====================================================
// LEITNER: pro-Karte Status sicherstellen
// =====================================================
function ensureCardProgress(entry) {
    const id = entry.id;
    if (!state.progress.cards[id]) {
        state.progress.cards[id] = {
            box: 0,          // 0 = neu (noch nie gesehen)
            timesCorrect: 0,
            timesWrong: 0,
            lastReview: 0
        };
    }
    return state.progress.cards[id];
}


export {
    updateLessonStatsUI,
    getLeitnerAscii,
    ensureCardProgress
};
