/* Search, settings menu, and startup */
import {
    APP_VERSION,
    state,
    $,
    translate,
    TRANSLATIONS,
    saveSettings,
    loadSettings,
    saveProgress,
    loadProgress
} from "./state.js";
import { loadCSV, populateLessonSelect, resetSessionStats } from "./csv.js";
import { updateLessonStatsUI } from "./leitner.js";
import {
    setTheme,
    applyTheme,
    setCard,
    nextCard,
    prevCard,
    doReveal,
    rate,
    startTraining,
    startChoice,
    startBrowse,
    renderDisplayToggleUI,
    updateTrainingBtn,
    updateChoiceBtn,
    updateBrowseBtn,
    updateModeButtons,
    hapticFeedback,
    hideNavButtons
} from "./card.js";
import {
    openVoicesPanelFor,
    closeVoices,
    refreshVoices,
    playQuestion,
    playAnswer
} from "./tts.js";
import {
    toggleAutoplay,
    stopAutoplayOnUserAction,
    updateAutoplayBtn
} from "./autoplay.js";

function openSearchPanel() {
    closeVoices();
    const panel = $("#searchPanel");
    const overlay = $("#searchOverlay");
    if (!panel) return;
    panel.classList.remove("hidden");
    overlay?.classList.add("active");
    setTimeout(() => {
        $("#searchInput")?.focus();
    }, 60);
}

function closeSearchPanel() {
    $("#searchPanel").classList.add("hidden");
    $("#searchOverlay")?.classList.remove("active");
}

function getAllCards() {
    const cards = [];
    for (const lessonCards of state.lessons.values()) {
        cards.push(...lessonCards);
    }
    return cards;
}
function normalizeRemoveDiacritics(s){
    if (!s) return "";
    return s.normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/[0-9]/g,'').replace(/[^\p{L}\s]/gu,'').toLowerCase();
}

function escapeHtml(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function highlightNormalized(original, query) {
    if (!original || !query) return escapeHtml(original || '');

    const originalChars = Array.from(original);
    const normChars = [];
    const mapping = [];

    originalChars.forEach((char, i) => {
        const parts = char.normalize('NFD');
        for (const part of parts) {
            if (!/\p{Diacritic}/u.test(part)) {
                normChars.push(part.toLowerCase());
                mapping.push(i);
            }
        }
    });

    const normText = normChars.join('');
    const normQuery = normalizeRemoveDiacritics(query).replace(/\s+/g, ' ').trim();
    if (!normQuery) return escapeHtml(original);

    let start = 0;
    const ranges = [];
    while (start < normText.length) {
        const idx = normText.indexOf(normQuery, start);
        if (idx === -1) break;
        const endMapIndex = idx + normQuery.length - 1;
        const origStart = mapping[idx];
        const origEnd = mapping[endMapIndex] + 1;
        ranges.push([origStart, origEnd]);
        start = idx + normQuery.length;
    }

    if (!ranges.length) return escapeHtml(original);

    let last = 0;
    const result = [];
    ranges.forEach(([from, to]) => {
        if (from > last) result.push(escapeHtml(original.slice(last, from)));
        result.push(`<mark>${escapeHtml(original.slice(from, to))}</mark>`);
        last = to;
    });
    if (last < original.length) result.push(escapeHtml(original.slice(last)));

    return result.join('');
}

function highlightSimple(original, query) {
    if (!original || !query) return escapeHtml(original || '');
    const escapedQuery = escapeHtml(query);
    const regex = new RegExp(escapedQuery.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&'), 'gi');
    return escapeHtml(original).replace(regex, match => `<mark>${match}</mark>`);
}

function searchCSV(query) {
    query = (query || "").trim();
    const resultsBox = $("#searchResults");
    if (!resultsBox) return;
    if (!query) {
        resultsBox.innerHTML = `<div class="search-hint">${translate('searchHint')}</div>`;
        return;
    }

    // Bestimme gewählte Suche-Language (DE oder ZH)
    const langDeBtn = document.querySelector('#searchLangDe');
    const lang = (langDeBtn && langDeBtn.classList.contains('active')) ? 'de' : 'zh';

    const qRaw = query.toLowerCase();
    const qNorm = normalizeRemoveDiacritics(query).replace(/\s+/g, ' ').trim();
    const qNormNoSpace = qNorm.replace(/\s+/g, '');

    const matchedWord = [];
    const matchedSentence = [];

    getAllCards().forEach((entry) => {
        let isWordMatch = false;
        let isSentenceMatch = false;

        if (lang === 'de') {
            const wordDe = (entry.word.de || '').toLowerCase();
            const sentDe = (entry.sent.de || '').toLowerCase();

            isWordMatch = wordDe.includes(qRaw);
            isSentenceMatch = !isWordMatch && sentDe.includes(qRaw);
        } else {
            const zhWord = (entry.word.zh || '').toLowerCase();
            const sentZh = (entry.sent.zh || '').toLowerCase();
            const pyWordRaw = entry.word.py || '';
            const pySentRaw = entry.sent.py || '';
            const pyWord = normalizeRemoveDiacritics(pyWordRaw).replace(/\s+/g, ' ').trim();
            const pySent = normalizeRemoveDiacritics(pySentRaw).replace(/\s+/g, ' ').trim();
            const pyWordNoSpace = pyWord.replace(/\s+/g, '');
            const pySentNoSpace = pySent.replace(/\s+/g, '');

            isWordMatch = (zhWord && zhWord.includes(qRaw)) ||
                (pyWord && (pyWord.includes(qNorm) || pyWordNoSpace.includes(qNormNoSpace)));

            if (!isWordMatch) {
                const queryTokens = qNorm.split(' ').filter(Boolean);
                const tokenMatch = queryTokens.length > 0 && queryTokens.every(token =>
                    pyWord.includes(token) || pyWordNoSpace.includes(token)
                );
                if (tokenMatch) isWordMatch = true;
            }

            if (!isWordMatch) {
                isSentenceMatch = (sentZh && sentZh.includes(qRaw)) ||
                    (pySent && (pySent.includes(qNorm) || pySentNoSpace.includes(qNormNoSpace)));

                if (!isSentenceMatch) {
                    const queryTokens = qNorm.split(' ').filter(Boolean);
                    const tokenMatch = queryTokens.length > 0 && queryTokens.every(token =>
                        pySent.includes(token) || pySentNoSpace.includes(token)
                    );
                    if (tokenMatch) isSentenceMatch = true;
                }
            }
        }

        if (isWordMatch) {
            matchedWord.push(entry);
        } else if (isSentenceMatch) {
            matchedSentence.push(entry);
        }
    });

    const matched = [...matchedWord, ...matchedSentence];

    if (!matched.length) {
        resultsBox.innerHTML = `<div class="search-empty">${translate('searchEmptyResults')}</div>`;
        return;
    }

    resultsBox.innerHTML = matched.map((entry) => {
        const deWordRaw = entry.word.de || '-';
        const pyWordRaw = entry.word.py || '-';
        const zhWordRaw = entry.word.zh || '-';

        const deSentRaw = entry.sent.de || '-';
        const pySentRaw = entry.sent.py || '-';
        const zhSentRaw = entry.sent.zh || '-';

        const deWord = highlightSimple(deWordRaw, qRaw);
        const deSent = highlightSimple(deSentRaw, qRaw);
        const pyWord = highlightNormalized(pyWordRaw, query);
        const pySent = highlightNormalized(pySentRaw, query);
        const zhWord = highlightSimple(zhWordRaw, qRaw);
        const zhSent = highlightSimple(zhSentRaw, qRaw);

        const posTag = entry.pos ? ` <span style="color: var(--muted); font-size: 0.9em; font-weight: normal;">(${escapeHtml(entry.pos)})</span>` : "";

        const wordLine = lang === 'de' ? `<strong>${deWord}</strong>${posTag}` : `${deWord}${posTag}`;
        const sentLine = lang === 'de' ? `<strong>${deSent}</strong>` : deSent;
        const wordLineZh = lang === 'zh' ? `<strong>${zhWord}</strong>` : zhWord;
        const sentLineZh = lang === 'zh' ? `<strong>${zhSent}</strong>` : zhSent;

        return `
            <div class="search-result" data-entry-id="${escapeHtml(entry.id)}">
                <div class="search-result-header">
                    <span class="search-result-lesson">${escapeHtml(entry.lesson || '–')}</span>
                    <span class="search-result-id">${escapeHtml(entry.id || '–')}</span>
                </div>

                <div class="search-result-main">
                    <div>${wordLine}</div>
                </div>

                <div class="search-result-line">${sentLine}</div>
                <div class="search-result-line">${pyWord}</div>
                <div class="search-result-line">${pySent}</div>
                <div class="search-result-line">${wordLineZh}</div>
                <div class="search-result-line">${sentLineZh}</div>
            </div>`;
    }).join("");

    resultsBox.querySelectorAll(".search-result").forEach((node) => {
        node.addEventListener("click", () => {
            const id = node.dataset.entryId;
            const selected = matched.find((entry) => entry.id === id);
            if (selected) {
                setCard(selected);
                closeSearchPanel();
            }
        });
    });
}

function translateAllUI() {
    document.documentElement.lang = state.settings.lang || "de";
    document.title = `${translate("appTitle")} – v${APP_VERSION}`;

    // 🔧 FIX: Text setzen ohne Kinder zu zerstören
    document.querySelectorAll("[data-i18n]").forEach((node) => {
        const key = node.dataset.i18n;
        if (!key) return;

        const text = translate(key);

        // 👉 Wenn das Element KEINE Kinder hat → normal ersetzen
        if (node.children.length === 0) {
            node.textContent = text;
        } else {
            // 👉 Wenn Kinder existieren → NUR ersten Textknoten ersetzen
            const firstTextNode = [...node.childNodes].find(n => n.nodeType === Node.TEXT_NODE);

            if (firstTextNode) {
                firstTextNode.nodeValue = text + " ";
            } else {
                // Falls kein Textknoten existiert → vorne einfügen
                node.insertBefore(document.createTextNode(text + " "), node.firstChild);
            }
        }
    });

    // Option-Elemente (unverändert ok)
    document.querySelectorAll("option[data-i18n]").forEach((option) => {
        const key = option.dataset.i18n;
        if (!key) return;
        option.textContent = translate(key);
    });

    // Title-Attribute
    document.querySelectorAll("[data-i18n-title]").forEach((node) => {
        const key = node.dataset.i18nTitle;
        if (!key) return;
        node.title = translate(key);
    });

    updateSearchPlaceholder();

    // Table Headers
    const lessonHeader = document.querySelector("#lessonTableHeaderLesson");
    if (lessonHeader) lessonHeader.textContent = translate("lessonTableLesson");

    const cardsHeader = document.querySelector("#lessonTableHeaderCards");
    if (cardsHeader) cardsHeader.textContent = translate("lessonTableCards");

    // Selects synchronisieren
    const uiLangSelect = document.querySelector("#uiLangSelect");
    if (uiLangSelect) uiLangSelect.value = state.settings.lang || "de";

    const themeSelect = document.querySelector("#themeSelect");
    if (themeSelect) themeSelect.value = state.settings.theme || "dark";

    updateTrainingBtn();
    updateAutoplayBtn();
    updateBrowseBtn();
    updateChoiceBtn();
}

function updateSearchPlaceholder() {
    const searchInput = document.querySelector("#searchInput");
    const langDeBtn = document.querySelector('#searchLangDe');
    const activeLang = (langDeBtn && langDeBtn.classList.contains('active')) ? 'de' : 'zh';
    if (!searchInput) return;
    const key = activeLang === 'de' ? 'searchPlaceholderDe' : 'searchPlaceholderZh';
    searchInput.placeholder = translate(key);
}

function setUILanguage(lang) {
    if (!TRANSLATIONS[lang]) return;
    state.settings.lang = lang;
    saveSettings();
    translateAllUI();
}

function renderModeUI() {
    const left  = $("#modeLeft");
    const right = $("#modeRight");

    if (state.mode === "de2zh") {
        left.textContent  = "🇩🇪 DE";
        right.textContent = "🇨🇳 ZH";
    } else {
        left.textContent  = "🇨🇳 ZH";
        right.textContent = "🇩🇪 DE";
    }

    $("#btnOrderToggle").textContent = translate("orderRandom");
    $("#btnOrderToggle").classList.toggle("active", state.order === "random");

    renderDisplayToggleUI();
    updateTrainingBtn();
}

/* ========================================================================== */
/*                                INIT ROUTINE                                */
/* ========================================================================== */

window.addEventListener("DOMContentLoaded", () => {

console.log(`[INIT] Starte Initialisierung … (v${APP_VERSION})`);

    /* ============================================================
       SETTINGS + PROGRESS LADEN + THEME & DELAY INITIALISIEREN
       ============================================================ */
    loadSettings();
	applyTheme(state.settings.theme);
    loadProgress();
    resetSessionStats();  // Session-Stats initialisieren
	
	// ✅ Resume-Fortschritt pro Lektion initialisieren
if (!state.settings.resumeIndexByLesson) {
    state.settings.resumeIndexByLesson = {};
}

// ==========================================
// PWA Install Button
// ==========================================

let deferredPrompt = null;

const installButton = document.getElementById("btnInstall");

if (installButton) {
    // Button zunächst ausblenden
    installButton.style.display = "none";

    // Event abfangen, wenn der Browser die Installation anbietet
    window.addEventListener("beforeinstallprompt", (e) => {
        console.log("✅ beforeinstallprompt ausgelöst");
        e.preventDefault();           // Verhindert den automatischen Prompt
        deferredPrompt = e;           // Speichern für später
        installButton.style.display = "block";   // Button anzeigen
    });

    // Button-Klick → Installation starten
    installButton.addEventListener("click", async () => {
        if (!deferredPrompt) return;

        installButton.style.display = "none";   // Button verstecken

        deferredPrompt.prompt();   // Install-Prompt anzeigen

        const choiceResult = await deferredPrompt.userChoice;
        console.log("Install choice:", choiceResult.outcome);

        deferredPrompt = null;     // Zurücksetzen
    });

    // Optional: Button wieder verstecken, wenn App bereits installiert wurde
    window.addEventListener("appinstalled", () => {
        console.log("✅ App wurde installiert");
        installButton.style.display = "none";
        deferredPrompt = null;
    });
}

    // Theme laden
    const savedTheme = state.settings.theme || localStorage.getItem("theme") || "dark";
    setTheme(savedTheme);

    // UI-Sprache initialisieren
    state.settings.lang = state.settings.lang || "de";
    translateAllUI();

    if (state.settings.sentenceDelay !== undefined) {
        state.sentenceDelay = state.settings.sentenceDelay;
    }

    const delayRange = document.querySelector("#delayRange");
    const delayVal = document.querySelector("#delayVal");
    if (delayRange) {
        delayRange.value = state.sentenceDelay / 1000;
    }
    if (delayVal) {
        delayVal.textContent = `(${(state.sentenceDelay / 1000).toFixed(1)} s)`;
    }

    // MODE, ORDER & DISPLAY TOGGLES
    state.mode      = state.settings.mode  || "de2zh";
    state.order     = state.settings.order || "random";
    state.showHanzi = state.settings.showHanzi !== false;
    state.showPinyin = state.settings.showPinyin !== false;

    // AUTOPLAY GAP
    state.autoplay.gapMs = state.settings.autoplayGap || 800;

    // TTS-Werte übernehmen
    state.rateDe  = state.settings.rateDe;
    state.pitchDe = state.settings.pitchDe;
    state.rateZh  = state.settings.rateZh;
    state.pitchZh = state.settings.pitchZh;

	renderModeUI();

	updateTrainingBtn();
	updateModeButtons();

	hideNavButtons();  // Buttons initial unsichtbar beim App-Start

    /* ============================================================
       CSV LADEN
       ============================================================ */
    loadCSV().then(() => {
    updateLessonStatsUI();
});

    /* ============================================================
       STIMMEN LADEN
       ============================================================ */
    speechSynthesis.onvoiceschanged = () => {
        refreshVoices();
    };
    setTimeout(refreshVoices, 300); // Fallback

 /* ============================================================
   SLIDE-DRAWER (⋮) – Menü öffnen/schließen + Animation
   ============================================================ */
const toggleBtn = document.querySelector("#menuToggle");
const sideMenu  = document.querySelector("#sideMenu");
const overlay   = document.querySelector("#sideOverlay"); // ✅ einzige overlay-Definition

if (toggleBtn && sideMenu) {

    // Menü per Button öffnen/schließen
    toggleBtn.addEventListener("click", () => {
        const isOpen = sideMenu.classList.toggle("open");

        // Für Animation (⋮ → ×)
        document.body.classList.toggle("menu-open", isOpen);
    });
}

// Tap auf Overlay → Menü schließen
if (overlay) {
    overlay.addEventListener("click", () => {
        sideMenu.classList.remove("open");
        document.body.classList.remove("menu-open");
    });
}

/* THEME-SWITCH */
const themeSelect = document.querySelector("#themeSelect");
if (themeSelect) {
    themeSelect.addEventListener("change", (e) => {
        setTheme(e.target.value);
    });
}

const uiLangSelect = document.querySelector("#uiLangSelect");
if (uiLangSelect) {
    uiLangSelect.addEventListener("change", (e) => {
        setUILanguage(e.target.value);
    });
}
  
    /* DELAY RANGE SLIDER */
    if (delayRange) {
        delayRange.addEventListener("input", (e) => {
            const seconds = parseFloat(e.target.value) || 0;
            state.sentenceDelay = seconds * 1000;
            state.settings.sentenceDelay = state.sentenceDelay;
            if (delayVal) delayVal.textContent = `(${seconds.toFixed(1)} s)`;
            saveSettings();
        });
    }

    document.querySelector("#btnVoiceDe")?.addEventListener("click", function() {
        this.classList.add("active");
        setTimeout(() => {
            openVoicesPanelFor("de");
        }, 300);
    });

    document.querySelector("#btnVoiceZh")?.addEventListener("click", function() {
        this.classList.add("active");
        setTimeout(() => {
            openVoicesPanelFor("zh");
        }, 300);
    });

    document.querySelector("#btnCloseVoices")?.addEventListener("click", () => {
        closeVoices();
    });

    document.querySelector("#searchToggle")?.addEventListener("click", () => {
        const panel = $("#searchPanel");
        if (panel && !panel.classList.contains("hidden")) {
            closeSearchPanel();
            return;
        }
        openSearchPanel();
    });

    document.querySelector("#btnCloseSearch")?.addEventListener("click", () => {
        closeSearchPanel();
    });

    document.querySelector("#searchOverlay")?.addEventListener("click", () => {
        closeSearchPanel();
    });

    const searchInput = document.querySelector("#searchInput");
    // Suchsprache Umschalter (DE / ZH)
    const searchLangDeBtn = document.querySelector("#searchLangDe");
    const searchLangZhBtn = document.querySelector("#searchLangZh");
    const setSearchLang = (lang) => {
        if (lang === 'de') {
            searchLangDeBtn?.classList.add('active');
            searchLangZhBtn?.classList.remove('active');
            if (searchInput) searchInput.placeholder = translate('searchPlaceholderDe');
        } else {
            searchLangZhBtn?.classList.add('active');
            searchLangDeBtn?.classList.remove('active');
            if (searchInput) searchInput.placeholder = translate('searchPlaceholderZh');
        }
    };

    // Default bleibt DE
    setSearchLang('de');

    searchLangDeBtn?.addEventListener('click', () => setSearchLang('de'));
    searchLangZhBtn?.addEventListener('click', () => setSearchLang('zh'));

    searchInput?.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
            event.preventDefault();
            searchCSV(searchInput.value);
        }
    });

    document.querySelector("#searchButton")?.addEventListener("click", () => {
        searchCSV(searchInput?.value || "");
    });

    /* ============================================================
       MODUS, REIHENFOLGE, AUTOPLAY
       ============================================================ */
    $("#btnSwapMode").addEventListener("click", function() {
        hapticFeedback();
        const btn = this;
        btn.classList.add("active");

        state.mode = state.mode === "de2zh" ? "zh2de" : "de2zh";
        state.settings.mode = state.mode;
        saveSettings();
        renderModeUI();
        if (state.current) setCard(state.current);

        setTimeout(() => {
            btn.classList.remove("active");
        }, 500);
    });

    $("#btnOrderToggle").addEventListener("click", () => {
        hapticFeedback();
   
        state.order = state.order === "random" ? "seq" : "random";
        state.settings.order = state.order;
        saveSettings();
        renderModeUI();
    });

    $("#btnToggleHanzi").addEventListener("click", () => {
        hapticFeedback();
     
        state.showHanzi = !state.showHanzi;
        state.settings.showHanzi = state.showHanzi;
        saveSettings();
        renderDisplayToggleUI();
        if (state.current) setCard(state.current, true);
    });

    $("#btnTogglePinyin").addEventListener("click", () => {
        hapticFeedback();
   
        state.showPinyin = !state.showPinyin;
        state.settings.showPinyin = state.showPinyin;
        saveSettings();
        renderDisplayToggleUI();
        if (state.current) setCard(state.current, true);
    });

    $("#btnAutoplay").addEventListener("click", () => {
        hapticFeedback();
        toggleAutoplay();
    });

    $("#gapRange").addEventListener("input", (e) => {
       
        const s = parseFloat(e.target.value) || 0.8;
        state.autoplay.gapMs = Math.round(s * 1000);
        state.settings.autoplayGap = state.autoplay.gapMs;
        $("#gapVal").textContent = `(${s.toFixed(1)} s)`;
        saveSettings();
    });

    /* ============================================================
       TRAINING + NAVIGATION
       ============================================================ */

    $("#btnStart").addEventListener("click", () => {
        hapticFeedback();
        stopAutoplayOnUserAction();
        startTraining();
    });
    $("#btnChoice")?.addEventListener("click", () => {
        hapticFeedback();
        stopAutoplayOnUserAction();
        startChoice();
    });
	$("#btnBrowse").addEventListener("click", () => {
        hapticFeedback();
		startBrowse();
	});
    $("#btnNext").addEventListener("click", () => {
       
        nextCard();
    });

    $("#btnPrev").addEventListener("click", () => {
       
        prevCard();
    });

    $("#btnReveal").addEventListener("click", () => {
        
        doReveal();
    });

    /* ============================================================
       AUDIO SPRECHER
       ============================================================ */
    $("#speakerQuestion").addEventListener("click", () => {
    
        playQuestion();
    });

    $("#speakerAnswer").addEventListener("click", (e) => {
        e.stopPropagation();
        playAnswer();
    });

    /* ============================================================
       RATING
       ============================================================ */

    $("#btnRateKnown").addEventListener("click", () => {
      
        rate("known");
    });

    $("#btnRateUnknown").addEventListener("click", () => {
     
        rate("unknown");
    });

  
    /* ============================================================
       IMPORT/EXPORT → jetzt im Seitenmenü
       ============================================================ */

    document.querySelector("#btnMenuExport")?.addEventListener("click", () => {
        const blob = new Blob(
            [JSON.stringify(state.progress, null, 2)],
            { type: "application/json" }
        );
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "progress.json";
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 300);
    });

    document.querySelector("#btnMenuImport")?.addEventListener("click", () => {
        const inp = document.createElement("input");
        inp.type = "file";
        inp.accept = "application/json";

        inp.onchange = (e) => {
            const f = e.target.files?.[0];
            if (!f) return;
            const r = new FileReader();
            r.onload = () => {
                try {
                    const p = JSON.parse(r.result);
                    if (p && p.version === "v1") {
                        state.progress = p;
                        saveProgress();
                        populateLessonSelect();
                        alert(translate("alertImportOk"));
                    } else {
                        alert(translate("alertImportInvalid"));
                    }
                } catch (err) {
                    alert(translate("alertImportError"));
                }
            };
            r.readAsText(f);
        };

        inp.click();
    });

    function getSelectedLessonForReset() {
        const sel = document.querySelector("#lessonSelect");
        if (!sel) return null;

        const selectedOption = sel.selectedOptions[0];
        if (selectedOption) return selectedOption.value;

        return state.settings.lessons[0] || state.lessonOrder[0] || null;
    }

    function resetProgressForLesson(lessonName) {
        if (!lessonName) return;

        const cards = state.lessons.get(lessonName) || [];
        for (const card of cards) {
            delete state.progress.cards[card.id];
        }

        if (state.progress.byLesson?.[lessonName]) {
            delete state.progress.byLesson[lessonName];
        }

        saveProgress();
        populateLessonSelect();
        updateLessonStatsUI();
    }

    function resetProgressAll() {
        state.progress = {
            version: "v1",
            cards: {},
            byLesson: {}
        };

        saveProgress();
        populateLessonSelect();
        updateLessonStatsUI();
    }

    document.querySelector("#btnResetLessonProgress")?.addEventListener("click", () => {
        const lessonName = getSelectedLessonForReset();
        if (!lessonName) {
            alert(translate("noLessonSelected"));
            return;
        }

        if (!confirm(translate("confirmResetLesson", { lesson: lessonName }))) return;

        resetProgressForLesson(lessonName);
        alert(translate("resetLessonDone", { lesson: lessonName }));
    });

    document.querySelector("#btnResetAllProgress")?.addEventListener("click", () => {
        if (!confirm(translate("confirmResetAll"))) return;

        resetProgressAll();
        alert(translate("resetAllDone"));
    });

	// ================================
	// Version im Menü anzeigen
	// ================================
	const verElem = document.querySelector("#appVersion");
	if (verElem) verElem.textContent = APP_VERSION;

/* ============================================================
   DRAG-TO-CLOSE – professionell wie in Mobile-Apps
   ============================================================ */

(function enableDragToClose() {
    const menu = document.querySelector("#sideMenu");
    if (!menu) return;

    let startX = 0;
    let currentX = 0;
    let dragging = false;

    function onStart(e) {
        if (!menu.classList.contains("open")) return;

        dragging = true;
        menu.classList.add("dragging");

        startX = e.touches ? e.touches[0].clientX : e.clientX;
        currentX = startX;
    }

    function onMove(e) {
        if (!dragging) return;

        currentX = e.touches ? e.touches[0].clientX : e.clientX;
        let diff = currentX - startX;

        // ✅ Nur rechts wischen erlaubt diff > 0
        if (diff > 0) {
            // Ziehe das Menü entsprechend nach rechts hinaus
            menu.style.right = `${-diff}px`;
        }
    }

    function onEnd() {
        if (!dragging) return;

        dragging = false;
        menu.classList.remove("dragging");

        let diff = currentX - startX;

        // ✅ Wenn genug nach rechts gewischt → Menü schließen
        if (diff > 40) {
            menu.classList.remove("open");
			document.body.classList.remove("menu-open");
        }

        // Menü resetten
        menu.style.right = "";
    }

    // Touch Events
    menu.addEventListener("touchstart", onStart);
    menu.addEventListener("touchmove", onMove);
    menu.addEventListener("touchend", onEnd);

    // Maus (für Desktop)
    menu.addEventListener("mousedown", onStart);
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onEnd);
})();

/* ============================================================
   KEYBOARD SHORTCUTS (Training & Browse)
   Space = Reveal, Arrow keys = Nav, 1 = Didn't know, 2 = Knew it
   ============================================================ */
document.addEventListener("keydown", (e) => {
    // Don't trigger shortcuts when typing in an input/textarea/select
    const tag = (e.target.tagName || "").toLowerCase();
    if (tag === "input" || tag === "textarea" || tag === "select") return;

    // Only active during training or browse mode, not during autoplay
    if (!state.trainingOn && !state.browseMode) return;
    if (state.autoplay.on) return;

    switch (e.key) {
        case " ":       // Space → Reveal answer
            e.preventDefault();
            if (!$("#btnReveal").disabled && $("#solBox").classList.contains("masked")) {
                doReveal();
            }
            break;

        case "ArrowRight":  // → Next card
            e.preventDefault();
            if (state.pool.length > 0) nextCard();
            break;

        case "ArrowLeft":   // ← Previous card
            e.preventDefault();
            if (state.historyPos > 0) prevCard();
            break;

        case "1":       // 1 → Didn't know
            if (!$("#btnRateUnknown").disabled && $("#rateBar").style.display !== "none") {
                rate("unknown");
            }
            break;

        case "2":       // 2 → Knew it
            if (!$("#btnRateKnown").disabled && $("#rateBar").style.display !== "none") {
                rate("known");
            }
            break;
    }
});

/* ============================================================
   SWIPE-NAVIGATION FÜR KARTEN (Browse & Training)
   ============================================================ */
(function enableSwipeNavigation() {
    const card = document.querySelector("#learnSection");
    if (!card) return;

    let startX = 0;
    let startY = 0;
    const threshold = 60; // Mindestdistanz für einen Swipe

    card.addEventListener("touchstart", (e) => {
        // Nur reagieren, wenn Training oder Browse-Modus aktiv ist
        if (!state.trainingOn && !state.browseMode) return;
        if (state.autoplay.on) return; // Autoplay nicht stören

        startX = e.touches[0].clientX;
        startY = e.touches[0].clientY;
    }, { passive: true });

    card.addEventListener("touchend", (e) => {
        if (!state.trainingOn && !state.browseMode) return;
        if (state.autoplay.on) return;

        const endX = e.changedTouches[0].clientX;
        const endY = e.changedTouches[0].clientY;
        const diffX = endX - startX;
        const diffY = endY - startY;

        // Sicherstellen, dass die horizontale Bewegung dominiert und weit genug ist
        if (Math.abs(diffX) > Math.abs(diffY) && Math.abs(diffX) > threshold) {
            if (diffX > 0) {
                // Swipe nach Rechts -> Vorherige Karte
                if (state.historyPos > 0) prevCard();
            } else {
                // Swipe nach Links -> Nächste Karte
                if (state.pool.length > 0) nextCard();
            }
        }
    }, { passive: true });
})();

/* ============================================
   Overlay tap-to-close
   ============================================ */
// overlay wurde oben im Menüblock definiert

if (overlay) {
    overlay.addEventListener("click", () => {
        sideMenu.classList.remove("open");
        document.body.classList.remove("menu-open");
    });
}

console.log("[INIT] Alles bereit ✅");
});  // ✅ schließt NUR den DOMContentLoaded – korrekt!


if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js")
      .then(reg => console.log("SW registered", reg))
      .catch(err => console.error("SW error", err));
  });
}


export {
    openSearchPanel,
    closeSearchPanel,
    getAllCards,
    normalizeRemoveDiacritics,
    escapeHtml,
    highlightNormalized,
    highlightSimple,
    searchCSV,
    translateAllUI,
    updateSearchPlaceholder,
    setUILanguage,
    renderModeUI
};
