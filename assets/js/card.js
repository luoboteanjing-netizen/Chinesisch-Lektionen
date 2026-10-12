/* Card rendering, navigation, training and browse */
import { state, $, translate, saveSettings, saveProgress } from "./state.js";
import { gatherPool, resetSessionStats } from "./csv.js";
import { ensureCardProgress, getLeitnerAscii, updateLessonStatsUI } from "./leitner.js";
import { playChineseOnReveal } from "./tts.js";

/* ============================ UTILS =============================== */

function hapticFeedback() {
    // Prüft, ob das Gerät Vibration unterstützt
    if ("vibrate" in navigator) {
        navigator.vibrate(70); // Kurzer 40ms Impuls
    }
}

function formatZh(hz, py) {
    hz = (hz || "").trim();
    py = (py || "").trim();
    return py
        ? `${hz}<br><span class="zh-pinyin">${py}</span>`
        : hz || "—";
}

/* ========================================================================== */

function setTheme(theme) {
    const root = document.documentElement;

    // Alle Theme-Klassen entfernen
    root.classList.remove(
        "light-orange",
        "light-warm",
        "light-blue"
    );

    // Dark = keine Klasse
    if (theme !== "dark") {
        root.classList.add(theme);
    }

    // merken
    localStorage.setItem("theme", theme);
    state.settings.theme = theme;
    saveSettings();
	applyTheme(theme);
}

function setTaskbarColor(color) {
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) {
    meta.setAttribute("content", color);
  }
}


function applyTheme(theme) {
 
  switch (theme) {
    case "dark":
      setTaskbarColor("#000000");
      break;

    case "light-orange":
      setTaskbarColor("#ffbb55");
      break;

    case "light-warm":
      setTaskbarColor("#c97c5d");
      break;

    case "light-blue":
      setTaskbarColor("#3b82f6");
      break;

    default:
      setTaskbarColor("#000000");
  }
}


/* ============================ sync & scroll ============================ */

function syncCardHeights() {
    const q = document.querySelector("#promptBox");
    const a = document.querySelector("#solBox");
    if (!q || !a) return;

    q.style.minHeight = "";
    a.style.minHeight = "";
    q.style.maxHeight = "";
    a.style.maxHeight = "";

    if (state.choiceMode) return;

    const h = Math.max(q.offsetHeight, a.offsetHeight);
    q.style.minHeight = h + "px";
    a.style.minHeight = h + "px";
}

function scrollToBottom() {
    requestAnimationFrame(() => {
    window.scrollTo({
        top: document.body.scrollHeight,
        behavior: "smooth"
    });
});
}

function scrollToTop() {
    requestAnimationFrame(() => {
    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });
});
}


/* ============================ CARD RENDERING ============================ */

function renderDisplayToggleUI() {
    const btnHanzi = $("#btnToggleHanzi");
    const btnPinyin = $("#btnTogglePinyin");
    if (!btnHanzi || !btnPinyin) return;

    btnHanzi.classList.toggle("active", state.showHanzi);
    btnPinyin.classList.toggle("active", state.showPinyin);
}

function renderPromptWord(entry) {
    if (state.mode !== "zh2de") {
        $("#promptWord").textContent = entry.word.de || "—";
        $("#promptWordSub").innerHTML = "";
        return;
    }

    const showHanzi = state.showHanzi;
    const showPinyin = state.showPinyin;

    if (showHanzi && showPinyin) {
        $("#promptWord").innerHTML = entry.word.zh || "—";
        $("#promptWordSub").innerHTML = entry.word.py
            ? `<span class="pinyin-word">${entry.word.py}</span>`
            : "";
    } else if (showHanzi) {
        $("#promptWord").innerHTML = entry.word.zh || "—";
        $("#promptWordSub").innerHTML = "";
    } else if (showPinyin) {
        $("#promptWord").innerHTML = entry.word.py
            ? `<span class="pinyin-word">${entry.word.py}</span>`
            : "—";
        $("#promptWordSub").innerHTML = "";
    } else {
        $("#promptWord").innerHTML = entry.word.zh || "—";
        $("#promptWordSub").innerHTML = "";
    }
}

function renderPromptSentence(entry) {
    const showHanzi = state.showHanzi;
    const showPinyin = state.showPinyin;
    const parts = [];

    if (showHanzi && entry.sent.zh) {
        parts.push(entry.sent.zh);
    }
    if (showPinyin && entry.sent.py) {
        parts.push(`<span class="zh-pinyin">${entry.sent.py}</span>`);
    }

    if (parts.length === 0) {
        $("#promptSent").textContent = "";
        return;
    }

    $("#promptSent").innerHTML = parts.join(parts.length > 1 ? "<br>" : " ");
	/* ===scrollToBottom(); === */
}

function renderPromptWordFull(entry) {
    if (state.mode !== "zh2de") {
        $("#promptWord").textContent = entry.word.de || "—";
        $("#promptWordSub").innerHTML = "";
        return;
    }

    $("#promptWord").innerHTML = entry.word.zh || "—";
    $("#promptWordSub").innerHTML = entry.word.py
        ? `<span class="pinyin-word">${entry.word.py}</span>`
        : "";
}

function renderPromptSentenceFull(entry) {
    if (state.mode !== "zh2de") {
        $("#promptSent").textContent = entry.sent.de || "—";
        return;
    }

    const parts = [];
    if (entry.sent.zh) parts.push(entry.sent.zh);
    if (entry.sent.py) parts.push(`<span class="zh-pinyin">${entry.sent.py}</span>`);

    if (parts.length === 0) {
        $("#promptSent").textContent = "";
        return;
    }

    $("#promptSent").innerHTML = parts.join("<br>");
	/* ===scrollToBottom(); === */
}

function setCard(entry, fromHistory = false) {

    /* ---- Timer für verzögerten Satz abbrechen ---- */
    if (state.delayedSentenceTimer) {
        clearTimeout(state.delayedSentenceTimer);
        state.delayedSentenceTimer = null;
    }

    if (!fromHistory) pushToHistory(entry);

    state.current = entry;
    state.startedAt = Date.now();
    state.revealedAt = null;

    /* -------- Titel (Lektion / ID) -------- */
    const cardTitle  = document.querySelector("#cardTitle");
    const cardLesson = document.querySelector("#cardLesson");

    if (cardTitle) {
        const p = ensureCardProgress(entry);
        const ascii = getLeitnerAscii(p.box);

        const revealCount = state.session.revealedCount;
        const cardsInLesson = state.lessons.get(entry.lesson) ?? [];
        const total = cardsInLesson.length;

        cardTitle.innerHTML = `
            <span class="card-title-left">
                ${revealCount} / ${total}
            </span>
            <span class="card-title-right leitner-ascii">
                ${ascii}
            </span>
        `;
    }

    if (cardLesson) {
        cardLesson.textContent = translate("cardLessonTitle", { id: entry.id });
    }

    /* -------- Fortschrittsbalken -------- */
    const stats = document.querySelector("#lessonStats");

    if (stats) {
        const cards = state.lessons.get(entry.lesson) ?? [];
        const total = cards.length;

        let red = 0, yellow = 0, green1 = 0, green2 = 0, green3 = 0, grey = 0;

        for (const c of cards) {
            const p = state.progress.cards[c.id] ?? { box: 0 };
            const box = p.box || 0;

            if (box === 0) grey++;
            else if (box === 1) red++;
            else if (box === 2) yellow++;
            else if (box === 3) green1++;
            else if (box === 4) green2++;
            else if (box === 5) green3++;
        }

        const redPct    = total ? (red    / total) * 100 : 0;
        const yellowPct = total ? (yellow / total) * 100 : 0;
        const green1Pct = total ? (green1 / total) * 100 : 0;
        const green2Pct = total ? (green2 / total) * 100 : 0;
        const green3Pct = total ? (green3 / total) * 100 : 0;
        const greyPct   = total ? (grey   / total) * 100 : 0;

        const leftGreen2 = green3Pct;
        const leftGreen1 = green3Pct + green2Pct;
        const leftYellow = green3Pct + green2Pct + green1Pct;
        const leftRed    = green3Pct + green2Pct + green1Pct + yellowPct;
        const leftGrey   = green3Pct + green2Pct + green1Pct + yellowPct + redPct;

        stats.innerHTML = `
            <div class="lesson-bar-large">
                <div class="lesson-bar-green3" style="left:0%; width:${green3Pct}%"></div>
                <div class="lesson-bar-green2" style="left:${leftGreen2}%; width:${green2Pct}%"></div>
                <div class="lesson-bar-green1" style="left:${leftGreen1}%; width:${green1Pct}%"></div>
                <div class="lesson-bar-yellow" style="left:${leftYellow}%; width:${yellowPct}%"></div>
                <div class="lesson-bar-red" style="left:${leftRed}%; width:${redPct}%"></div>
                <div class="lesson-bar-grey" style="left:${leftGrey}%; width:${greyPct}%"></div>
            </div>
        `;
    }

    /* -------- Karte anzeigen -------- */
	const sol = $("#solBox");

	if (state.browseMode) {
		sol.classList.remove("masked");
	} else {
		sol.classList.add("masked");
	}

    /* ✅ Hilfsfunktionen */
    const hasWordZh = entry.word?.zh && entry.word.zh.trim() !== "";
    const hasWordDe = entry.word?.de && entry.word.de.trim() !== "";

    if (state.mode === "zh2de") {
        /* ---- CH → DE ---- */

        renderPromptWord(entry);
        $("#promptPOS").textContent = entry.pos || "";
        $("#promptSent").innerHTML = "";

        $("#solWord").textContent = entry.word.de;
        $("#solSent").textContent = entry.sent.de;

        /* ✅ KEIN Wort → kein Delay */
        if (!hasWordZh && !hasWordDe) {
            renderPromptSentence(entry);
            syncCardHeights();
        } else {
            state.delayedSentenceTimer = setTimeout(() => {
                renderPromptSentence(entry);
                syncCardHeights();
            }, state.sentenceDelay);
        }

    } else {
        /* ---- DE → CH ---- */

        $("#promptWord").textContent = entry.word.de || "—";
        $("#promptWordSub").innerHTML = "";
        $("#promptPOS").textContent = entry.pos || "";
        $("#promptSent").textContent = "";

        $("#solWord").innerHTML =
            `${entry.word.zh}<br><span class="zh-pinyin">${entry.word.py}</span>`;

        $("#solSent").innerHTML =
            `${entry.sent.zh}<br><span class="zh-pinyin">${entry.sent.py}</span>`;

        /* ✅ KEIN Wort → kein Delay */
        if (!hasWordDe) {
            $("#promptSent").textContent = entry.sent.de || "—";
            syncCardHeights();
        } else {
            state.delayedSentenceTimer = setTimeout(() => {
                $("#promptSent").textContent = entry.sent.de || "—";
                syncCardHeights();
            }, state.sentenceDelay);
        }
    }

    /* -------- Buttons setzen -------- */
    $("#btnReveal").disabled = false;

    hideRatingButtons();
    showNavButtons();
    updateNavButtons();

    syncCardHeights();
	if (state.browseMode) {
    hideRatingButtons();
	}
	
	if (state.browseMode) {

    // Cancel the delayed sentence timer – browse shows everything immediately
    if (state.delayedSentenceTimer) {
        clearTimeout(state.delayedSentenceTimer);
        state.delayedSentenceTimer = null;
    }

    if (state.mode === "zh2de") {

        renderPromptWord(entry);       // respects showHanzi/showPinyin
        renderPromptSentence(entry);   // respects showHanzi/showPinyin

    } else {

        $("#promptSent").textContent = entry.sent.de || "—";
    }

    syncCardHeights();
	}

    renderChoices(entry);
}

function choiceAnswerText(entry) {
    if (!entry) return "—";
    if (state.mode === "zh2de") {
        return (entry.word?.de || "").trim() || (entry.sent?.de || "").trim() || "—";
    }

    const wordZh = (entry.word?.zh || "").trim();
    const wordPy = (entry.word?.py || "").trim();
    const hanzi = wordZh || (entry.sent?.zh || "").trim();
    const pinyin = wordZh ? wordPy : (entry.sent?.py || "").trim();
    const parts = [];
    if (state.showHanzi !== false && hanzi) parts.push(hanzi);
    if (state.showPinyin !== false && pinyin) parts.push(pinyin);
    return parts.join(" ") || hanzi || pinyin || "—";
}

function buildChoiceOptions(entry) {
    const correctText = choiceAnswerText(entry);
    const sameLesson = state.lessons.get(entry.lesson) || [];
    const rest = [];
    for (const cards of state.lessons.values()) {
        if (cards !== sameLesson) rest.push(...cards);
    }
    const candidates = [...sameLesson, ...rest].filter((card) => card.id !== entry.id);
    for (let i = candidates.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [candidates[i], candidates[j]] = [candidates[j], candidates[i]];
    }

    const used = new Set([correctText]);
    const picked = [entry];
    for (const card of candidates) {
        const text = choiceAnswerText(card);
        if (!text || text === "—" || used.has(text)) continue;
        used.add(text);
        picked.push(card);
        if (picked.length === 4) break;
    }

    for (let i = picked.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [picked[i], picked[j]] = [picked[j], picked[i]];
    }
    return picked;
}

function setChoiceLayout(on) {
    document.querySelector("#learnSection")?.classList.toggle("choice-active", !!on);
}

function parkAnswerSpeaker() {
    const speaker = $("#speakerAnswer");
    const home = document.querySelector("#solSent")?.parentElement;
    if (!speaker || !home) return speaker;
    if (speaker.parentElement !== home) home.appendChild(speaker);
    return speaker;
}

function clearChoiceBox() {
    const speaker = parkAnswerSpeaker();
    const box = $("#choiceBox");
    if (box) {
        box.innerHTML = "";
        box.hidden = true;
    }
    if (speaker) speaker.style.visibility = "";
    setChoiceLayout(false);
}

function selectChoice(id) {
    if (!state.choiceMode || state.revealedAt) return;
    state.choiceSelected = id;
    document.querySelectorAll(".choice-btn").forEach((btn) => {
        btn.classList.toggle("selected", btn.dataset.choiceId === id);
    });
}

function renderChoices(entry) {
    const box = $("#choiceBox");
    if (!box) return;
    if (!state.choiceMode || !entry) {
        clearChoiceBox();
        return;
    }

    state.choiceSelected = null;
    setChoiceLayout(true);
    box.hidden = false;
    const speaker = parkAnswerSpeaker();
    box.innerHTML = "";
    buildChoiceOptions(entry).forEach((card) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "choice-btn";
        btn.dataset.choiceId = card.id;
        btn.textContent = choiceAnswerText(card);
        btn.addEventListener("click", () => selectChoice(card.id));
        box.appendChild(btn);
    });

    if (speaker) speaker.style.visibility = "hidden";
    syncCardHeights();
}

function appendZhBlock(el, zh, py) {
    const hz = (zh || "").trim();
    const pinyin = (py || "").trim();
    if (hz) el.appendChild(document.createTextNode(hz));
    if (pinyin) {
        if (hz) el.appendChild(document.createElement("br"));
        const span = document.createElement("span");
        span.className = "zh-pinyin";
        span.textContent = pinyin;
        el.appendChild(span);
    }
}

function choiceRevealNodes(entry) {
    const nodes = [];
    const posText = (entry.pos || "").trim();
    const answer = document.createElement("span");
    answer.className = "choice-answer";
    const sent = document.createElement("span");
    sent.className = "choice-sent";

    if (state.mode === "zh2de") {
        const word = (entry.word?.de || "").trim();
        const sentence = (entry.sent?.de || "").trim();
        if (word) answer.textContent = word;
        if (sentence) sent.textContent = sentence;
    } else {
        appendZhBlock(answer, entry.word?.zh, entry.word?.py);
        appendZhBlock(sent, entry.sent?.zh, entry.sent?.py);
    }

    if (answer.childNodes.length || answer.textContent) nodes.push(answer);
    if (posText) {
        const pos = document.createElement("span");
        pos.className = "choice-pos";
        pos.textContent = posText;
        nodes.push(pos);
    }
    if (sent.childNodes.length || sent.textContent) nodes.push(sent);
    if (!nodes.length) {
        answer.textContent = "—";
        nodes.push(answer);
    }
    return nodes;
}

function scrollChoiceIntoView(btn) {
    const box = document.querySelector("#solBox");
    if (!box || !btn) return;
    const top = btn.offsetTop;
    const bottom = top + btn.offsetHeight;
    const viewBottom = box.scrollTop + box.clientHeight;
    if (bottom > viewBottom - 8) {
        box.scrollTop = bottom - box.clientHeight + 12;
    } else if (top < box.scrollTop) {
        box.scrollTop = Math.max(0, top - 8);
    }
}

function expandCorrectChoice(entry) {
    if (!entry) return;
    const btn = [...document.querySelectorAll(".choice-btn")]
        .find((item) => item.dataset.choiceId === entry.id);
    if (!btn) return;
    btn.classList.add("correct", "expanded");
    btn.disabled = false;
    btn.replaceChildren(...choiceRevealNodes(entry));
    const speaker = $("#speakerAnswer");
    if (speaker) {
        btn.appendChild(speaker);
        speaker.style.visibility = "";
    }
    requestAnimationFrame(() => scrollChoiceIntoView(btn));
}

function gradeChoices() {
    const selected = state.choiceSelected;
    const correct = state.current?.id;
    document.querySelectorAll(".choice-btn").forEach((btn) => {
        btn.disabled = btn.dataset.choiceId !== correct;
        if (selected && btn.dataset.choiceId === selected && selected !== correct) {
            btn.classList.add("wrong");
        }
    });
    expandCorrectChoice(state.current);
}

/* ============================ HISTORY / NAV ============================ */

function pushToHistory(entry) {
    if (state.historyPos < state.history.length - 1) {
        state.history = state.history.slice(0, state.historyPos + 1);
    }
    state.history.push(entry);
    state.historyPos = state.history.length - 1;
}

function updateNavButtons() {
    $("#btnPrev").disabled = state.historyPos <= 0;
    $("#btnNext").disabled = state.pool.length === 0;
}

function nextCard() {
    hapticFeedback();
	state.cardCounter++;
    if (!state.pool.length) return;

    if (state.historyPos < state.history.length - 1) {
        state.historyPos++;
        setCard(state.history[state.historyPos], true);
        syncCardHeights();
        return;
    }
	// 🔥 NEU: Prüfen ob eine Karte wieder erscheinen soll
	const dueIndex = state.reinsertQueue.findIndex(item => item.due <= state.cardCounter);

	if (dueIndex !== -1) {
		const item = state.reinsertQueue.splice(dueIndex, 1)[0];
		setCard(item.card);
		syncCardHeights();
		return;
	}
	
    let next;

    if (state.order === "seq") {
        if (state.idx == null) state.idx = 0;
        else state.idx = (state.idx + 1) % state.pool.length;
        next = state.pool[state.idx];
    } else {
        next = state.pool[Math.floor(Math.random() * state.pool.length)];
    }

    setCard(next);
    syncCardHeights();
}

function prevCard() {
    hapticFeedback();
    if (state.historyPos > 0) {
        state.historyPos--;
        setCard(state.history[state.historyPos], true);
    }
    updateNavButtons();
    syncCardHeights();
}


/* ============================ NAV SHOW/HIDE ============================ */

function hideNavButtons() {
    $("#btnPrev").style.display = "none";
    $("#btnReveal").style.display = "none";
    $("#btnNext").style.display = "none";
}

function showNavButtons() {

    if ((!state.trainingOn && !state.browseMode) || state.autoplay.on) {
        return;
    }

    $("#btnPrev").style.display = "";

	if (state.browseMode) {
		$("#btnReveal").style.display = "";
		$("#btnReveal").style.visibility = "hidden";
		$("#btnReveal").style.pointerEvents = "none";

	} else {

		$("#btnReveal").style.display = "";
		$("#btnReveal").style.visibility = "";
		$("#btnReveal").style.pointerEvents = "";

	}

    $("#btnNext").style.display = "";

    scrollToBottom();
}

/* ============================ REVEAL / RATING ============================ */

function doReveal() {
    hapticFeedback();
    $("#solBox").classList.remove("masked");
    state.revealedAt = Date.now();
	
	// ✅ Kartenzähler hochzählen (nur beim Aufdecken von NEUEN Karten)
	if (state.current && state.current.id && !state.session.revealedCardIds.includes(state.current.id)) {
		state.session.revealedCardIds.push(state.current.id);
		state.session.revealedCount++;
		// Zähler in cardTitle aktualisieren
		const cardTitle = document.querySelector("#cardTitle");
		if (cardTitle) {
			const p = ensureCardProgress(state.current);
			const ascii = getLeitnerAscii(p.box);
			const cardsInLesson = state.lessons.get(state.current.lesson) ?? [];
			const total = cardsInLesson.length;
			cardTitle.innerHTML = `
				<span class="card-title-left">
					${state.session.revealedCount} / ${total}
				</span>
				<span class="card-title-right leitner-ascii">
					${ascii}
				</span>
			`;
		}
	}
	
    // ✅ Beim Aufdecken IMMER Chinesisch abspielen
	playChineseOnReveal(state.current);


    // -----------------------------------------
    // LEITNER: Erste Sichtung → Box 0 → Box 1
    // -----------------------------------------
    const p = ensureCardProgress(state.current);
    if (p.box === 0) {
        p.box = 1;                    // neu → schwach
        p.lastReview = Date.now();
        saveProgress();
        updateLessonStatsUI();
    }

    // -----------------------------------------
    // Timer abbrechen
    // -----------------------------------------
// -----------------------------------------
// Fragekarte VOLLSTÄNDIG anzeigen beim Aufdecken
// -----------------------------------------
if (state.mode === "zh2de") {
    renderPromptWordFull(state.current);
    renderPromptSentenceFull(state.current);
} else {
    // ✅ DE → CH sofort alles anzeigen
    $("#promptSent").textContent = state.current.sent.de || "—";
}

if (state.delayedSentenceTimer) {
    clearTimeout(state.delayedSentenceTimer);
    state.delayedSentenceTimer = null;
}
    // -----------------------------------------
    // Buttons anzeigen
    // -----------------------------------------
    if (state.choiceMode) {
        gradeChoices();
        applyRating(null, { advance: false });
        hideRatingButtons();
        $("#btnPrev").style.display = "";
        $("#btnNext").style.display = "";
        const revealBtn = $("#btnReveal");
        if (revealBtn) {
            revealBtn.disabled = true;
            revealBtn.style.visibility = "hidden";
            revealBtn.style.pointerEvents = "none";
        }
        refreshCardLeitner();
    } else {
        if (!state.autoplay.on) hideNavButtons();
        showRatingButtons();
        enableRating();
    }
    syncCardHeights();
}

function refreshCardLeitner() {
    const asciiEl = document.querySelector("#cardTitle .leitner-ascii");
    if (!asciiEl || !state.current) return;
    asciiEl.textContent = getLeitnerAscii(ensureCardProgress(state.current).box);
}

function showRatingButtons() {
    $("#rateBar").style.display = "flex";
	scrollToBottom();
}

function hideRatingButtons() {
    $("#rateBar").style.display = "none";
}

function enableRating() {
    if (state.choiceMode) {
        const correct = !!state.choiceSelected && state.choiceSelected === state.current?.id;
        $("#btnRateKnown").disabled = !correct;
        $("#btnRateUnknown").disabled = correct;
        return;
    }
    $("#btnRateKnown").disabled = false;
  
    $("#btnRateUnknown").disabled = false;
}

function disableRating() {
    $("#btnRateKnown").disabled = true;
 
    $("#btnRateUnknown").disabled = true;
}

function rate(mark) {
    applyRating(mark, { advance: true });
}

function applyRating(mark, { advance = true } = {}) {
    if (!state.current) return;
    if (state.choiceMode) {
        mark = state.choiceSelected === state.current.id ? "known" : "unknown";
    }
    if (advance) hapticFeedback();

    // -----------------------------------------
    // LEITNER: Bewertung
    // -----------------------------------------
    const p = ensureCardProgress(state.current);

    // Falls Karte gerade erst zum ersten Mal aufgedeckt wurde:
    if (p.box === 0) p.box = 1;

    if (mark === "known") {
        // richtig:
        // - Box 1 → Box 2 (erste korrekte Antwort)
        // - danach normale Leiter hoch
        if (p.box === 1) p.box = 2;
        else p.box = Math.min(p.box + 1, 5);

        p.timesCorrect++;
    }
    else if (mark === "unknown") {
    // falsch → zurück zu 1
    p.box = 1;
    p.timesWrong++;

    // Karte verzögert wieder einplanen (3–8 Karten später)
    const REINSERT_MIN = 3;
    const REINSERT_MAX = 8;
    const delay = Math.floor(Math.random() * (REINSERT_MAX - REINSERT_MIN + 1)) + REINSERT_MIN;

    state.reinsertQueue.push({
        card: state.current,
        due: state.cardCounter + delay
    });
}

    p.lastReview = Date.now();
    saveProgress();
    updateLessonStatsUI();

    // -----------------------------------------
    // DEIN ORIGINALER CODE (unverändert)
    // -----------------------------------------
    state.session.done++;
    if (mark === "known") state.session.known++;
    else if (mark === "unsure") state.session.unsure++;
    else state.session.unknown++;

    const lesson = state.current.lesson;
    if (lesson) {
        if (!state.progress.byLesson[lesson])
            state.progress.byLesson[lesson] = { known: 0, unknown: 0 };

        if (mark === "known")   state.progress.byLesson[lesson].known++;
        if (mark === "unknown") state.progress.byLesson[lesson].unknown++;

        saveProgress();
        updateLessonStatsUI();
    }

    disableRating();
    hideRatingButtons();
    if (!advance) return;
    showNavButtons();
    nextCard();
}

/* ============================ TRAINING ============================ */

let startAsChoice = false;

function startTraining() {
    if (!startAsChoice && state.choiceMode) {
        state.choiceMode = false;
        state.trainingOn = false;
    }

    if (!state.trainingOn) {

        // ----------------------------
        // Training vorbereiten
        // ----------------------------
        state.history = [];
        state.historyPos = -1;
        state.session.revealedCount = 0;
        state.session.revealedCardIds = [];

        state.selectedLessons.clear();
        const sel = $("#lessonSelect");

        const picked = [];
        for (const o of sel.selectedOptions) {
            state.selectedLessons.add(o.value);
            picked.push(o.value);
        }

        state.settings.lessons = picked;
        saveSettings();

        gatherPool();

        if (!state.pool.length) {
            alert(translate("selectLessonAlert"));
            startAsChoice = false;
            state.choiceMode = false;
            return;
        }

        // ----------------------------
        // ✅ Lesson-Variable für beide Abschnitte
        // ----------------------------
        const lesson = state.settings.lessons[0];

        // ----------------------------
        // ✅ Kartenzähler initialisieren (bereits bekannte Karten)
        // ----------------------------
        const cardsInLesson = state.lessons.get(lesson) || [];
        let knownCardsCount = 0;
        const knownCardIds = [];

        for (const card of cardsInLesson) {
            const progress = state.progress.cards[card.id];
            if (progress && progress.box >= 1 && progress.box <= 5) {
                knownCardsCount++;
                knownCardIds.push(card.id);
            }
        }

        state.session.revealedCount = knownCardsCount;
        state.session.revealedCardIds = knownCardIds;

        // ----------------------------
        // ✅ Session-Stats initialisieren (nach bekannte Karten zählen)
        // ----------------------------
        resetSessionStats();

        // ----------------------------
        // ✅ Resume-Index bestimmen
        // ----------------------------
        const resumeIdx = state.settings.resumeIndexByLesson?.[lesson];

        if (typeof resumeIdx === "number" && resumeIdx < state.pool.length) {
            state.idx = resumeIdx;
            // ✅ History mit vorherigen Karten vorladen (1 bis resumeIdx)
            for (let i = 0; i < resumeIdx; i++) {
                state.history.push(state.pool[i]);
            }
            state.historyPos = resumeIdx - 1;
        } else {
            state.idx = 0;
        }

        // ----------------------------
        // ✅ Erste Karte setzen
        // ----------------------------
        state.choiceMode = startAsChoice;
        startAsChoice = false;
        setCard(state.pool[state.idx]);

        // ----------------------------
        // ✅ Training aktivieren
        // ----------------------------
        state.trainingOn = true;
		
		state.browseMode = false;

		$("#btnReveal").style.visibility = "visible";
		$("#btnReveal").style.pointerEvents = "";
		$("#btnReveal").style.opacity = "1";
		
        updateModeButtons();
        showNavButtons();  // Buttons anzeigen beim Training-Start
        scrollToBottom();

    } else {
        startAsChoice = false;
        stopTraining();
    }
}

function startChoice() {
    if (state.choiceMode) {
        stopTraining();
        return;
    }
    state.stopAutoplayOnUserAction?.();
    if (state.browseMode) state.browseMode = false;
    if (state.trainingOn) state.trainingOn = false;
    startAsChoice = true;
    startTraining();
}

function startBrowse() {

    if (!state.browseMode) {

        state.stopAutoplayOnUserAction?.();
        state.trainingOn = false;
        state.choiceMode = false;
        clearChoiceBox();
        state.browseMode = true;

        state.history = [];
        state.historyPos = -1;

        state.selectedLessons.clear();

        const sel = $("#lessonSelect");
        const picked = [];

        for (const o of sel.selectedOptions) {
            state.selectedLessons.add(o.value);
            picked.push(o.value);
        }

        state.settings.lessons = picked;
        saveSettings();

        gatherPool();

        if (!state.pool.length) {
            alert(translate("selectLessonAlert"));
            return;
        }

        state.idx = 0;

        setCard(state.pool[state.idx]);

        updateModeButtons();

        showNavButtons();

    } else {

        stopBrowse();
    }
}


function stopTraining() {
    state.trainingOn = false;
    state.choiceMode = false;
    state.choiceSelected = null;
    clearChoiceBox();
	
	// ✅ Resume-Index der aktuellen Lektion speichern (Training + Autoplay)
if (state.current && state.current.lesson && state.idx !== null) {
    state.settings.resumeIndexByLesson[state.current.lesson] = state.idx;
    saveSettings();
}
	
    updateTrainingBtn();
    hideNavButtons();  // Buttons verstecken beim Training-Stopp

    $("#btnPrev").disabled = true;
    $("#btnReveal").disabled = true;
    $("#btnNext").disabled = true;

    disableRating();
    hideRatingButtons();
	updateLessonStatsUI();
	updateModeButtons();

    $("#solBox").classList.add("masked");
	scrollToTop();
}

function stopBrowse() {

    state.browseMode = false;

    hideNavButtons();
    hideRatingButtons();

    $("#solBox").classList.add("masked");

    updateModeButtons();

    scrollToTop();
}

function updateTrainingBtn() {
    const trainingActive = state.trainingOn && !state.choiceMode;
    $("#btnStart").textContent =
        trainingActive ? translate("trainingStop") : translate("trainingStart");
}

function updateChoiceBtn() {
    const btn = $("#btnChoice");
    if (!btn) return;
    btn.textContent = state.choiceMode ? translate("choiceStop") : translate("choiceStart");
}

function updateBrowseBtn() {
    $("#btnBrowse").textContent =
        state.browseMode ? translate("browseStop") : translate("browseStart");
}

function updateModeButtons() {

    $("#btnStart").classList.remove("active-mode");
    $("#btnBrowse").classList.remove("active-mode");
    $("#btnAutoplay").classList.remove("active-mode");
    $("#btnChoice")?.classList.remove("active-mode");

    if (state.trainingOn && !state.choiceMode) {
        $("#btnStart").classList.add("active-mode");
    }

    if (state.browseMode) {
        $("#btnBrowse").classList.add("active-mode");
    }

    if (state.autoplay.on) {
        $("#btnAutoplay").classList.add("active-mode");
    }

    if (state.choiceMode) {
        $("#btnChoice")?.classList.add("active-mode");
    }

    updateBrowseBtn();
    updateChoiceBtn();
}

state.setCard = setCard;

window.addEventListener("resize", () => {
    if (state.choiceMode) syncCardHeights();
});


export {
    hapticFeedback,
    formatZh,
    setTheme,
    setTaskbarColor,
    applyTheme,
    syncCardHeights,
    scrollToBottom,
    scrollToTop,
    renderDisplayToggleUI,
    renderPromptWord,
    renderPromptSentence,
    renderPromptWordFull,
    renderPromptSentenceFull,
    setCard,
    pushToHistory,
    updateNavButtons,
    nextCard,
    prevCard,
    hideNavButtons,
    showNavButtons,
    doReveal,
    showRatingButtons,
    hideRatingButtons,
    enableRating,
    disableRating,
    rate,
    startTraining,
    startChoice,
    startBrowse,
    stopTraining,
    stopBrowse,
    updateTrainingBtn,
    updateChoiceBtn,
    updateBrowseBtn,
    clearChoiceBox,
    updateModeButtons
};
