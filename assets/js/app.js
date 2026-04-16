/* ========================================================================== */
/*                           FLASHCARDS – VERSION OPTION A                    */
/*                       Repaired, Stabilized, Fully Functional               */
/* ========================================================================== */

/* -------------------------------------------------------------------------- */
/* TEIL 1 – GLOBAL STATE · SETTINGS · CSV PARSER                              */
/* -------------------------------------------------------------------------- */

/* ===========================
   AUTOMATISCHE VERSIONIERUNG
   =========================== */

/* === Version manuell definieren === */
const APP_VERSION = "6.0";   // beim nächsten Release erhöhen

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

const LS_KEYS = {
    settings: "fc_settings_v1",
    progress: "fc_progress_v1"
};

const TRANSLATIONS = {
    de: {
        appTitle: "Chinesisch Flashcards",
        settingsTitle: "Einstellungen",
        progressTitle: "Fortschritt",
        progressExport: "Fortschritt exportieren",
        progressImport: "Fortschritt importieren",
        progressResetLesson: "Aktuelle Lektion zurücksetzen",
        progressResetAll: "Alle Lektionen zurücksetzen",
        languageTitle: "Sprache",
        uiLanguageLabel: "🌐 UI Sprache",
        uiLanguageSelectDe: "Deutsch",
        uiLanguageSelectEn: "English",
        uiLanguageSelectZh: "中文",
        uiLanguageSelectFr: "Français",
        themeTitle: "Theme",
        themeDark: "Dark Mode",
        themeLightOrange: "Light‑Orange",
        themeLightWarm: "Light‑Warm",
        themeLightBlue: "Light‑Blue",
        delayLabel: "Satz‑Delay (Sekunden):",
        autoplayGapLabel: "Pause zw. Karten im Autoplay:",
        settingsVersion: "Version:",
        modeSwitchTitle: "Richtung umschalten",
        orderRandom: "Zufällig",
        autoPlay: "Autoplay ▶︎",
        autoPlayStop: "Autoplay ■ Stop",
        trainingStart: "Training starten ▶",
        trainingStop: "Training stoppen ■",
        prev: "◀ Zurück",
        reveal: "Aufdecken",
        next: "Nächste ▶",
        rateUnknown: "❌ Nicht gewusst",
        rateKnown: "✅ Gewusst",
        voiceButtonDe: "DE Stimme",
        voiceButtonZh: "ZH Stimme",
        rateDeLabel: "DE Tempo",
        pitchDeLabel: "DE Tonhöhe",
        rateZhLabel: "ZH Tempo",
        pitchZhLabel: "ZH Tonhöhe",
        pinyinButton: "Pīnyīn",
        lessonTableLesson: "Lektion",
        lessonTableCards: "Karten",
        voicePanelTitle: "Stimmen",
        voicePanelClose: "✕",
        voicePanelHintTitle: "Hinweis",
        voicePanelHint: "Wähle eine Stimme für die aktuell geöffnete Sprache.",
        voiceListTitle: "Stimmenliste",
        noVoicesFound: "Keine passenden Stimmen gefunden.",
        namelessVoice: "(namenlos)",
        pickVoice: "Diese Stimme wählen",
        testVoice: "Probehören",
        voiceActiveSuffix: "• [Aktiv]",
        selectLessonAlert: "Bitte zuerst Lektionen auswählen.",
        selectLessonAlert2: "Bitte Lektionen wählen.",
        noLessonSelected: "Keine Lektion ausgewählt.",
        confirmResetLesson: "Fortschritt für die Lektion ‚{lesson}‘ wirklich zurücksetzen?",
        resetLessonDone: "Fortschritt für Lektion ‚{lesson}‘ wurde zurückgesetzt.",
        confirmResetAll: "Fortschritt für alle Lektionen wirklich vollständig zurücksetzen?",
        resetAllDone: "Fortschritt für alle Lektionen wurde zurückgesetzt.",
        alertImportOk: "Fortschritt importiert.",
        alertImportInvalid: "Ungültiges Format.",
        alertImportError: "Fehler beim Import.",
        csvLoadError: "Fehler beim Laden der CSV.",
        cardLessonTitle: "Lektion {id}"
    },
    en: {
        appTitle: "Chinese Flashcards",
        settingsTitle: "Settings",
        progressTitle: "Progress",
        progressExport: "Export progress",
        progressImport: "Import progress",
        progressResetLesson: "Reset current lesson",
        progressResetAll: "Reset all lessons",
        languageTitle: "Language",
        uiLanguageLabel: "🌐 UI language",
        uiLanguageSelectDe: "Deutsch",
        uiLanguageSelectEn: "English",
        uiLanguageSelectZh: "中文",
        uiLanguageSelectFr: "French",
        themeTitle: "Theme",
        themeDark: "Dark mode",
        themeLightOrange: "Light Orange",
        themeLightWarm: "Light Warm",
        themeLightBlue: "Light Blue",
        delayLabel: "Sentence delay (seconds):",
        autoplayGapLabel: "Pause between cards in autoplay:",
        settingsVersion: "Version:",
        modeSwitchTitle: "Switch direction",
        orderRandom: "Random",
        autoPlay: "Autoplay ▶︎",
        autoPlayStop: "Autoplay ■ Stop",
        trainingStart: "Start training ▶",
        trainingStop: "Stop training ■",
        prev: "◀ Back",
        reveal: "Reveal",
        next: "Next ▶",
        rateUnknown: "❌ Didn't know",
        rateKnown: "✅ Knew it",
        voiceButtonDe: "DE Voice",
        voiceButtonZh: "ZH Voice",
        rateDeLabel: "DE rate",
        pitchDeLabel: "DE pitch",
        rateZhLabel: "ZH rate",
        pitchZhLabel: "ZH pitch",
        pinyinButton: "Pīnyīn",
        lessonTableLesson: "Lesson",
        lessonTableCards: "Cards",
        voicePanelTitle: "Voices",
        voicePanelClose: "✕",
        voicePanelHintTitle: "Hint",
        voicePanelHint: "Choose a voice for the currently open language.",
        voiceListTitle: "Voice list",
        noVoicesFound: "No matching voices found.",
        namelessVoice: "(nameless)",
        pickVoice: "Pick this voice",
        testVoice: "Listen",
        voiceActiveSuffix: "• [Active]",
        selectLessonAlert: "Please choose lessons first.",
        selectLessonAlert2: "Please choose lessons.",
        noLessonSelected: "No lesson selected.",
        confirmResetLesson: "Reset progress for lesson '{lesson}'?",
        resetLessonDone: "Progress for lesson '{lesson}' has been reset.",
        confirmResetAll: "Reset progress for all lessons?",
        resetAllDone: "Progress for all lessons has been reset.",
        alertImportOk: "Progress imported.",
        alertImportInvalid: "Invalid format.",
        alertImportError: "Import failed.",
        csvLoadError: "Error loading CSV.",
        cardLessonTitle: "Lesson {id}"
    },
    zh: {
        appTitle: "中文抽认卡",
        settingsTitle: "设置",
        progressTitle: "进度",
        progressExport: "导出进度",
        progressImport: "导入进度",
        progressResetLesson: "重置当前课",
        progressResetAll: "重置所有课程",
        languageTitle: "语言",
        uiLanguageLabel: "🌐 UI 语言",
        uiLanguageSelectDe: "Deutsch",
        uiLanguageSelectEn: "English",
        uiLanguageSelectZh: "中文",
        uiLanguageSelectFr: "Français",
        themeTitle: "主题",
        themeDark: "深色模式",
        themeLightOrange: "浅橙",
        themeLightWarm: "浅暖",
        themeLightBlue: "浅蓝",
        delayLabel: "句子延迟（秒）：",
        autoplayGapLabel: "自动播放卡片之间的暂停：",
        settingsVersion: "版本：",
        modeSwitchTitle: "切换方向",
        orderRandom: "随机",
        autoPlay: "自动播放 ▶︎",
        autoPlayStop: "自动播放 ■ 停止",
        trainingStart: "开始训练 ▶",
        trainingStop: "停止训练 ■",
        prev: "◀ 上一张",
        reveal: "显示",
        next: "下一张 ▶",
        rateUnknown: "❌ 不会",
        rateKnown: "✅ 会了",
        voiceButtonDe: "DE 语音",
        voiceButtonZh: "ZH 语音",
        rateDeLabel: "DE 速率",
        pitchDeLabel: "DE 音调",
        rateZhLabel: "ZH 速率",
        pitchZhLabel: "ZH 音调",
        pinyinButton: "Pīnyīn",
        lessonTableLesson: "课程",
        lessonTableCards: "卡片",
        voicePanelTitle: "语音",
        voicePanelClose: "✕",
        voicePanelHintTitle: "提示",
        voicePanelHint: "为当前语言选择语音。",
        voiceListTitle: "语音列表",
        noVoicesFound: "未找到匹配语音。",
        namelessVoice: "(无名)",
        pickVoice: "选择此语音",
        testVoice: "试听",
        voiceActiveSuffix: "• [已激活]",
        selectLessonAlert: "请先选择课程。",
        selectLessonAlert2: "请选择课程。",
        noLessonSelected: "未选择课程。",
        confirmResetLesson: "确实要重置课程“{lesson}”的进度吗？",
        resetLessonDone: "课程“{lesson}”的进度已重置。",
        confirmResetAll: "确实要重置所有课程的进度吗？",
        resetAllDone: "所有课程的进度已重置。",
        alertImportOk: "进度已导入。",
        alertImportInvalid: "格式无效。",
        alertImportError: "导入失败。",
        csvLoadError: "加载 CSV 时出错。",
        cardLessonTitle: "课程 {id}"
    },
    fr: {
        appTitle: "Cartes Mémoire Chinois",
        settingsTitle: "Paramètres",
        progressTitle: "Progrès",
        progressExport: "Exporter le progrès",
        progressImport: "Importer le progrès",
        progressResetLesson: "Réinitialiser la leçon actuelle",
        progressResetAll: "Réinitialiser toutes les leçons",
        languageTitle: "Langue",
        uiLanguageLabel: "🌐 Langue de l'interface",
        uiLanguageSelectDe: "Deutsch",
        uiLanguageSelectEn: "English",
        uiLanguageSelectZh: "中文",
        uiLanguageSelectFr: "Französisch",
        themeTitle: "Thème",
        themeDark: "Mode sombre",
        themeLightOrange: "Clair Orange",
        themeLightWarm: "Clair Chaud",
        themeLightBlue: "Clair Bleu",
        delayLabel: "Délai de phrase (secondes) :",
        autoplayGapLabel: "Pause entre les cartes en lecture automatique :",
        settingsVersion: "Version :",
        modeSwitchTitle: "Changer de direction",
        orderRandom: "Aléatoire",
        autoPlay: "Lecture automatique ▶︎",
        autoPlayStop: "Lecture automatique ■ Arrêter",
        trainingStart: "Commencer l'entraînement ▶",
        trainingStop: "Arrêter l'entraînement ■",
        prev: "◀ Précédent",
        reveal: "Révéler",
        next: "Suivant ▶",
        rateUnknown: "❌ Ne savait pas",
        rateKnown: "✅ Savait",
        voiceButtonDe: "Voix DE",
        voiceButtonZh: "Voix ZH",
        rateDeLabel: "Vitesse DE",
        pitchDeLabel: "Hauteur DE",
        rateZhLabel: "Vitesse ZH",
        pitchZhLabel: "Hauteur ZH",
        pinyinButton: "Pīnyīn",
        lessonTableLesson: "Leçon",
        lessonTableCards: "Cartes",
        voicePanelTitle: "Voix",
        voicePanelClose: "✕",
        voicePanelHintTitle: "Astuce",
        voicePanelHint: "Choisissez une voix pour la langue actuellement ouverte.",
        voiceListTitle: "Liste des voix",
        noVoicesFound: "Aucune voix correspondante trouvée.",
        namelessVoice: "(sans nom)",
        pickVoice: "Choisir cette voix",
        testVoice: "Écouter",
        voiceActiveSuffix: "• [Actif]",
        selectLessonAlert: "Veuillez d'abord choisir des leçons.",
        selectLessonAlert2: "Veuillez choisir des leçons.",
        noLessonSelected: "Aucune leçon sélectionnée.",
        confirmResetLesson: "Réinitialiser vraiment le progrès de la leçon '{lesson}' ?",
        resetLessonDone: "Le progrès de la leçon '{lesson}' a été réinitialisé.",
        confirmResetAll: "Réinitialiser vraiment le progrès de toutes les leçons ?",
        resetAllDone: "Le progrès de toutes les leçons a été réinitialisé.",
        alertImportOk: "Progrès importé.",
        alertImportInvalid: "Format invalide.",
        alertImportError: "Échec de l'importation.",
        csvLoadError: "Erreur lors du chargement du CSV.",
        cardLessonTitle: "Leçon {id}"
    }
};

const $ = (s) => document.querySelector(s);

/* ============================ GLOBAL STATE =============================== */

const state = {
    mode: "de2zh",
    order: "seq",

    // TTS settings
    rateDe: 0.95,
    pitchDe: 1.0,
    rateZh: 0.95,
    pitchZh: 1.0,

    lessons: new Map(),
    lessonOrder: [],
    selectedLessons: new Set(),

    pool: [],
    idx: null,

    history: [],
    historyPos: -1,

    current: null,
	delayedSentenceTimer: null,
	sentenceDelay: 3000,  // in Millisekunden

    voices: [],
    browserVoice: { zh: null, de: null },
    voicePanelTarget: "de",

    autoplay: {
        on: false,
        timers: [],
        gapMs: 800
    },

    settings: {
        lang: "de",
        mode: "de2zh",
        order: "seq",
        rateDe: 0.95,
        pitchDe: 1.0,
        rateZh: 0.95,
        pitchZh: 1.0,
        showHanzi: true,
        showPinyin: true,
        lessons: [],
        browserVoiceZh: null,
        browserVoiceDe: null,
        autoplayGap: 800,
        resumeIndexByLesson: {}
    },

    session: {
        total: 0,
        done: 0,
        known: 0,
        unsure: 0,
        unknown: 0,
        ttrSum: 0,
        ttrCount: 0,
        revealedCount: 0,
        revealedCardIds: []
    },

    startedAt: null,
    revealedAt: null,

    progress: {
        version: "v1",
        cards: {},
        byLesson: {}
    },

    wakeLock: null,
    trainingOn: false
};

/* ============================ SETTINGS / PROGRESS ========================= */

function saveSettings() {
    try {
        localStorage.setItem(LS_KEYS.settings, JSON.stringify(state.settings));
    } catch (e) {
        console.warn("[Settings Save Error]", e);
    }
}

function loadSettings() {
    try {
        const s = JSON.parse(localStorage.getItem(LS_KEYS.settings) || "null");
        if (s) Object.assign(state.settings, s);
    } catch (e) {
        console.warn("[Settings Load Error]", e);
    }
}

function saveProgress() {
    try {
        localStorage.setItem(LS_KEYS.progress, JSON.stringify(state.progress));
    } catch (e) {
        console.warn("[Progress Save Error]", e);
    }
}

function loadProgress() {
    try {
        const p = JSON.parse(localStorage.getItem(LS_KEYS.progress) || "null");
        if (p && p.version === "v1") state.progress = p;
    } catch (e) {
        console.warn("[Progress Load Error]", e);
    }
}

function translate(key, vars = {}) {
    const lang = state.settings.lang || "de";
    const dictionary = TRANSLATIONS[lang] || TRANSLATIONS.de;
    let text = dictionary[key] ?? TRANSLATIONS.de[key] ?? key;
    Object.entries(vars).forEach(([name, value]) => {
        text = text.replace(new RegExp(`\\{${name}\\}`, "g"), value);
    });
    return text;
}

function translateAllUI() {
    document.documentElement.lang = state.settings.lang || "de";
    document.title = `${translate("appTitle")} – v${APP_VERSION}`;

    document.querySelectorAll("[data-i18n]").forEach((node) => {
        const key = node.dataset.i18n;
        if (!key) return;
        node.textContent = translate(key);
    });

    // Speziell für option elements in selects
    document.querySelectorAll("option[data-i18n]").forEach((option) => {
        const key = option.dataset.i18n;
        if (!key) return;
        option.textContent = translate(key);
    });

    document.querySelectorAll("[data-i18n-title]").forEach((node) => {
        const key = node.dataset.i18nTitle;
        if (!key) return;
        node.title = translate(key);
    });

    // Update lesson table headers
    const lessonHeader = document.querySelector("#lessonTableHeaderLesson");
    if (lessonHeader) lessonHeader.textContent = translate("lessonTableLesson");

    const cardsHeader = document.querySelector("#lessonTableHeaderCards");
    if (cardsHeader) cardsHeader.textContent = translate("lessonTableCards");

    const uiLangSelect = document.querySelector("#uiLangSelect");
    if (uiLangSelect) uiLangSelect.value = state.settings.lang || "de";

    const themeSelect = document.querySelector("#themeSelect");
    if (themeSelect) themeSelect.value = state.settings.theme || "dark";

    updateTrainingBtn();
    updateAutoplayBtn();
}

function setUILanguage(lang) {
    if (!TRANSLATIONS[lang]) return;
    state.settings.lang = lang;
    saveSettings();
    translateAllUI();
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
            lesson
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
            setCard(state.pool[0]);
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

let lessonSort = { key: null, asc: true };

/* ============================================================
   SORTIERUNG FÜR LEKTIONSTABELLE
   ============================================================ */
document.addEventListener("click", (ev) => {
    const sortKey = ev.target.dataset.sort;
    if (!sortKey) return;

    // ✅ Beim Klick auf "Lektion" → original CSV-Reihenfolge wiederherstellen
    if (sortKey === "lesson") {
        loadCSV();      // CSV neu laden = Reihenfolge exakt wie importiert
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
    const total = cards.length;

    const p = state.progress.byLesson[lessonName] || { known: 0, unknown: 0 };
    const known = p.known || 0;
    const unknown = p.unknown || 0;
    const percent = total ? Math.round((known / total) * 100) : 0;

    return {
        lesson: lessonName,
        total,
        known,
        unknown,
        percent
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

/* ============================ UTILS =============================== */

function formatZh(hz, py) {
    hz = (hz || "").trim();
    py = (py || "").trim();
    return py
        ? `${hz}<br><span class="zh-pinyin">${py}</span>`
        : hz || "—";
}

function getLeitnerAscii(box) {
    // Box 0–5 → 0–5 gefüllte Kästchen
    const filled = Math.max(0, Math.min(box, 5));
    return "■".repeat(filled) + "□".repeat(5 - filled);
}

/* ========================================================================== */
/*                                ENDE TEIL 1                                 */
/* ========================================================================== */

/* ========================================================================== */
/*                           TEIL 2 – CARD LOGIC                              */
/*                Rendering · Navigation · Rating · Training                  */
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
}

/* ============================ sync & scroll ============================ */

function syncCardHeights() {
    const q = document.querySelector("#promptBox");
    const a = document.querySelector("#solBox");
    if (!q || !a) return;

    q.style.minHeight = "";
    a.style.minHeight = "";

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
	console.log("SCROLL bottom");
});
}

function scrollToTop() {
    requestAnimationFrame(() => {
    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });
	console.log("SCROLL TOP");
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

    const revealCount = state.session.revealedCount;  // Nur aufgedeckte Karten
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

    if (cardLesson) cardLesson.textContent = translate("cardLessonTitle", { id: entry.id });

// ----------------------------------------------------------
// Fortschrittsbalken (Leitner) – von links nach rechts:
// Grün (4+5) → Gelb (2+3) → Rot (1) → Grau (0)
// ----------------------------------------------------------
const stats = document.querySelector("#lessonStats");
if (stats) {
    const cards = state.lessons.get(entry.lesson) ?? [];
    const total = cards.length;

    let green = 0;   // Box 4 + 5
    let yellow = 0;  // Box 2 + 3
    let red   = 0;   // Box 1
    let grey  = 0;   // Box 0

    for (const c of cards) {
        const p = state.progress.cards[c.id] ?? { box: 0 };

        if (p.box === 0) grey++;
        else if (p.box === 1) red++;
        else if (p.box === 2 || p.box === 3) yellow++;
        else if (p.box === 4 || p.box === 5) green++;
    }

    const greenPct  = total ? (green  / total) * 100 : 0;
    const yellowPct = total ? (yellow / total) * 100 : 0;
    const redPct    = total ? (red    / total) * 100 : 0;
    const greyPct   = total ? (grey   / total) * 100 : 0;

    const leftYellow = greenPct;
    const leftRed    = greenPct + yellowPct;
    const leftGrey   = greenPct + yellowPct + redPct;

    stats.innerHTML = `
        <div class="lesson-bar-large">
            <div class="lesson-bar-green"  style="left:0%;           width:${greenPct}%"></div>
            <div class="lesson-bar-yellow" style="left:${leftYellow}%;width:${yellowPct}%"></div>
            <div class="lesson-bar-red"    style="left:${leftRed}%;   width:${redPct}%"></div>
            <div class="lesson-bar-grey"   style="left:${leftGrey}%;  width:${greyPct}%"></div>
        </div>
    `;
}
    /* -------- Karte anzeigen -------- */
    const sol = $("#solBox");
    sol.classList.add("masked");

    /* Wort, Pinyin & POS werden immer sofort angezeigt */
    if (state.mode === "zh2de") {
        /* ---- CH → DE ---- */

        renderPromptWord(entry);
        $("#promptPOS").textContent = entry.pos || "";

        /* ✅ Satz NICHT sofort anzeigen */
        $("#promptSent").innerHTML = "";

        /* ✅ Lösungskarte sofort setzen */
        $("#solWord").textContent = entry.word.de;
        $("#solSent").textContent = entry.sent.de;

        /* ✅ Verzögertes Einblenden des Satzes */
        state.delayedSentenceTimer = setTimeout(() => {
            renderPromptSentence(entry);
            syncCardHeights();
        }, state.sentenceDelay);

    } else {
        /* ---- DE → CH ---- */

        $("#promptWord").textContent = entry.word.de || "—";
        $("#promptWordSub").innerHTML = "";
        $("#promptPOS").textContent = entry.pos || "";

        /* ✅ Satz NICHT sofort anzeigen */
        $("#promptSent").textContent = "";

        /* ✅ Lösungskarte: CH + Pinyin */
        $("#solWord").innerHTML =
            `${entry.word.zh}<br><span class="zh-pinyin">${entry.word.py}</span>`;
        $("#solSent").innerHTML =
            `${entry.sent.zh}<br><span class="zh-pinyin">${entry.sent.py}</span>`;

        /* ✅ Verzögertes Einblenden des Satzes (Deutsch) */
        state.delayedSentenceTimer = setTimeout(() => {
            $("#promptSent").textContent = entry.sent.de || "—";
            syncCardHeights();
        }, state.sentenceDelay);
    }

    /* -------- Buttons setzen -------- */
    $("#btnReveal").disabled = false;

    hideRatingButtons();
    showNavButtons();
    updateNavButtons();

    syncCardHeights();
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

    if (!state.pool.length) return;

    if (state.historyPos < state.history.length - 1) {
        state.historyPos++;
        setCard(state.history[state.historyPos], true);
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
    if (!state.trainingOn || state.autoplay.on) return;  // Nur im Training UND nicht im Autoplay zeigen!
    $("#btnPrev").style.display = "";
    $("#btnReveal").style.display = "";
    $("#btnNext").style.display = "";
}


/* ============================ REVEAL / RATING ============================ */

function doReveal() {
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
    if (!state.autoplay.on) hideNavButtons();
    showRatingButtons();
    enableRating();
    syncCardHeights();
}

function showRatingButtons() {
    $("#rateBar").style.display = "flex";
}

function hideRatingButtons() {
    $("#rateBar").style.display = "none";
}

function enableRating() {
    $("#btnRateKnown").disabled = false;
  
    $("#btnRateUnknown").disabled = false;
}

function disableRating() {
    $("#btnRateKnown").disabled = true;
 
    $("#btnRateUnknown").disabled = true;
}

function rate(mark) {
    if (!state.current) return;

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
    else if (mark === "unsure") {
        // unsicher → 2 oder 3
        if (p.box < 2)      p.box = 2;   // 1 → 2
        else if (p.box === 2) p.box = 3; // 2 → 3

        p.timesWrong++;
    }
    else if (mark === "unknown") {
        // falsch → zurück zu 1
        p.box = 1;
        p.timesWrong++;
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
    showNavButtons();
    nextCard();
}

/* ============================ TRAINING ============================ */

function startTraining() {

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
        setCard(state.pool[state.idx]);

        // ----------------------------
        // ✅ Training aktivieren
        // ----------------------------
        state.trainingOn = true;
        updateTrainingBtn();
        showNavButtons();  // Buttons anzeigen beim Training-Start
        scrollToBottom();

    } else {
        stopTraining();
    }
}

function stopTraining() {
    state.trainingOn = false;
	
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

    $("#solBox").classList.add("masked");
	scrollToTop();
}

function updateTrainingBtn() {
    $("#btnStart").textContent =
        state.trainingOn ? translate("trainingStop") : translate("trainingStart");
}



/* ========================================================================== */
/*                                ENDE TEIL 2                                 */
/* ========================================================================== */

/* ========================================================================== */
/*                           TEIL 3 – TTS & AUTOPLAY                          */
/* ========================================================================== */


/* ============================ TTS PRIME DELAY ============================ */

function ttsPrime(cb) {
    setTimeout(cb, 120);
}


/* ============================ BUILD UTTERANCE ============================ */

function buildUtterance(text, langKey) {

    const u = new SpeechSynthesisUtterance(text || "");
    u.lang = (langKey === "zh") ? "zh-CN" : "de-DE";

    if (langKey === "zh") {
        u.rate  = state.rateZh;
        u.pitch = state.pitchZh;
    } else {
        u.rate  = state.rateDe;
        u.pitch = state.pitchDe;
    }

    const chosen = (langKey === "zh") ? state.browserVoice.zh : state.browserVoice.de;
    if (chosen) u.voice = chosen;

    return u;
}


/* ============================ DETECT VOICES ============================ */

function isZhVoice(v) {
    const L = (v.lang || "").toLowerCase();
    return (
        L.startsWith("zh") ||
        L.includes("cmn") ||
        L.includes("hans") ||
        L.includes("zh-cn")
    );
}

function isDeVoice(v) {
    const L = (v.lang || "").toLowerCase();
    return L.startsWith("de");
}


/* ============================ REFRESH VOICES ============================ */

function refreshVoices() {

    state.voices = window.speechSynthesis.getVoices() || [];

    if (state.settings.browserVoiceZh) {
        const vz = state.voices.find(v =>
            v.name === state.settings.browserVoiceZh ||
            v.voiceURI === state.settings.browserVoiceZh
        );
        if (vz) state.browserVoice.zh = vz;
    }

    if (state.settings.browserVoiceDe) {
        const vd = state.voices.find(v =>
            v.name === state.settings.browserVoiceDe ||
            v.voiceURI === state.settings.browserVoiceDe
        );
        if (vd) state.browserVoice.de = vd;
    }

    updateVoiceList();
}


/* ============================ VOICE PANEL ============================ */

let voiceRetryTimer = null;

function openVoicesPanelFor(target) {
    state.voicePanelTarget = target;
    refreshVoices();

    if (!state.voices.length) {
        clearTimeout(voiceRetryTimer);
        let tries = 0;

        const attempt = () => {
            tries++;
            refreshVoices();
            if (state.voices.length || tries >= 10) return;
            voiceRetryTimer = setTimeout(attempt, 200);
        };

        voiceRetryTimer = setTimeout(attempt, 200);
    }

    $("#voicePanel").classList.remove("hidden");
}

function closeVoices() {
    $("#voicePanel").classList.add("hidden");
}


/* ============================ UPDATE VOICE LIST ============================ */

function updateVoiceList() {
    const box = $("#dbgVoices");
    if (!box) return;

    box.innerHTML = "";

    const list = (state.voices || []).filter(v =>
        state.voicePanelTarget === "zh" ? isZhVoice(v) : isDeVoice(v)
    );

    if (!list.length) {
        box.innerHTML = `<div>${translate("noVoicesFound")}</div>`;
        return;
    }

    list.forEach(v => {

        const row  = document.createElement("div");
        row.className = "voice";

        const name = document.createElement("div");
        name.className = "name";
        name.textContent = v.name || "(namenlos)";

        const meta = document.createElement("div");
        meta.className = "meta";
        meta.textContent = `${v.lang}${v.default ? " · default" : ""}`;

        const actions = document.createElement("div");
        actions.style.display = "flex";
        actions.style.gap = "6px";
        actions.style.marginLeft = "auto";

        const btnPick = document.createElement("button");
        btnPick.className = "btn";
        btnPick.textContent = translate("pickVoice");

        btnPick.onclick = () => {
            if (state.voicePanelTarget === "zh") {
                state.browserVoice.zh = v;
                state.settings.browserVoiceZh =
                    v.name || v.voiceURI;
            } else {
                state.browserVoice.de = v;
                state.settings.browserVoiceDe =
                    v.name || v.voiceURI;
            }
            saveSettings();
            closeVoices();
        };

        const btnTest = document.createElement("button");
        btnTest.className = "btn ghost";
        btnTest.textContent = translate("testVoice");

        btnTest.onclick = () => {
            const u = new SpeechSynthesisUtterance(
                state.voicePanelTarget === "zh"
                    ? "这是一个测试。"
                    : "Dies ist ein Test."
            );
            u.lang = state.voicePanelTarget === "zh" ? "zh-CN" : "de-DE";
            u.voice = v;
            speechSynthesis.cancel();
            speechSynthesis.speak(u);
        };

        const active = state.voicePanelTarget === "zh"
            ? state.browserVoice.zh
            : state.browserVoice.de;

        if (active &&
            (active.name === v.name || active.voiceURI === v.voiceURI)) {
            name.textContent += ` ${translate("voiceActiveSuffix")}`;
        }

        actions.appendChild(btnPick);
        actions.appendChild(btnTest);

        row.appendChild(name);
        row.appendChild(meta);
        row.appendChild(actions);

        box.appendChild(row);
    });
}


/* ============================ PLAY QUESTION / ANSWER ============================ */

function playQuestion() {
    if (!state.current) return;

    if (state.mode === "de2zh") {
        playSequence(
            state.current.word.de, "de",
            state.current.sent.de, "de"
        );
    } else {
        speechSynthesis.cancel();
        ttsSpeak(state.current.word.zh, "zh");
        setTimeout(() =>
            ttsSpeak(state.current.sent.zh, "zh"),
        600);
    }
}

function playAnswer() {
    if (!state.current) return;

    if (state.mode === "de2zh") {
        ttsSpeak(state.current.word.zh, "zh");
        setTimeout(() =>
            ttsSpeak(state.current.sent.zh, "zh"),
        600);
    } else {
        playSequence(
            state.current.word.de, "de",
            state.current.sent.de, "de"
        );
    }
}


/* ============================ SEQUENCED PLAYBACK ============================ */

function ttsSpeak(text, langKey) {
    const u = buildUtterance(text, langKey);
    speechSynthesis.speak(u);
    return u;
}

function playSequence(a, aLang, b, bLang) {
    ttsPrime(() => {
        speechSynthesis.cancel();
        ttsSpeak(a, aLang);
        setTimeout(() => ttsSpeak(b, bLang), 700);
    });
}


/* ============================ AUTOPLAY ============================ */

function setAutoplay(on) {
 console.log("setAutoplay called:", on);
    state.autoplay.on = on;

    if (!on) {
        speechSynthesis.cancel();
        state.autoplay.timers.forEach(x => clearTimeout(x));
        state.autoplay.timers = [];
        releaseWakeLock();
		scrollToTop();
    }

    updateAutoplayBtn();
}

function updateAutoplayBtn() {
    $("#btnAutoplay").textContent =
        state.autoplay.on ? translate("autoPlayStop") : translate("autoPlay");
}


function ensurePoolForAutoplay() {

    if (state.pool.length) return true;

    if (!state.settings.lessons.length) {

        const sel = $("#lessonSelect");
        const picked = [];

        for (const o of sel.selectedOptions) picked.push(o.value);

        state.settings.lessons = picked;
        saveSettings();
    }

    gatherPoolFromSettings();

    if (!state.pool.length) {
        alert(translate("selectLessonAlert2"));
        return false;
    }

    const lesson = state.settings.lessons[0];
    const resumeIdx = state.settings.resumeIndexByLesson?.[lesson];

    if (typeof resumeIdx === "number" && resumeIdx < state.pool.length) {
        state.idx = resumeIdx;
        setCard(state.pool[resumeIdx]);
    } else {
        state.idx = Math.floor(Math.random() * state.pool.length);
        setCard(state.pool[state.idx]);
    }

    return true;
}


function speakPair(word, sent, langKey, done) {

    if (!state.autoplay.on) return;

    const u1 = buildUtterance(word, langKey);

    u1.onend = () => {

        if (!state.autoplay.on) return;

        const t = setTimeout(() => {

            if (!state.autoplay.on) return;

            const u2 = buildUtterance(sent, langKey);

            u2.onend = () => {
                if (!state.autoplay.on) return;
                done && done();
            };

            speechSynthesis.speak(u2);

        }, 650);

        state.autoplay.timers.push(t);
    };

    speechSynthesis.speak(u1);
}

function playChineseOnReveal(entry) {
    if (!entry) return;

    speechSynthesis.cancel();

    // Wort zuerst
    ttsSpeak(entry.word.zh, "zh");

    // Satz leicht verzögert danach
    setTimeout(() => {
        ttsSpeak(entry.sent.zh, "zh");
    }, 600);
}

function autoplayStep() {

    if (!state.autoplay.on) return;

    if (!ensurePoolForAutoplay()) {
        setAutoplay(false);
        return;
    }

    $("#solBox").classList.add("masked");
    disableRating();

    const qLang = (state.mode === "de2zh") ? "de" : "zh";
    const aLang = (state.mode === "de2zh") ? "zh" : "de";

    ttsPrime(() => {

        speechSynthesis.cancel();

        speakPair(
            state.current.word[qLang],
            state.current.sent[qLang],
            qLang,

            () => {

                if (!state.autoplay.on) return;

                $("#solBox").classList.remove("masked");

                speakPair(
                    state.current.word[aLang],
                    state.current.sent[aLang],
                    aLang,

                    () => {

                        if (!state.autoplay.on) return;

                        const t = setTimeout(() => {

                            if (!state.autoplay.on) return;

                            if (state.order === "seq") {
                                if (state.idx == null) state.idx = 0;
                                else state.idx = (state.idx + 1) % state.pool.length;

                                setCard(state.pool[state.idx]);
                            } else {
                                setCard(
                                    state.pool[Math.floor(Math.random() * state.pool.length)]
                                );
                            }

                            autoplayStep();

                        }, state.autoplay.gapMs);

                        state.autoplay.timers.push(t);
                    }
                );
            }
        );

    });
}


function toggleAutoplay() {

    // === AUTOPLAY START ===
    if (!state.autoplay.on) {

        // Pool laden falls leer
        if (!ensurePoolForAutoplay()) return;

        // Autoplay aktivieren
        setAutoplay(true);
        requestWakeLock();

        // Falls noch keine Karte angezeigt wurde
        if (!state.current) {
            if (state.order === "seq") {
                state.idx = 0;
                setCard(state.pool[state.idx]);
            } else {
                const first = state.pool[Math.floor(Math.random() * state.pool.length)];
                setCard(first);
            }
        }

        scrollToBottom();
        autoplayStep();
        return;
    }

    // === AUTOPLAY STOP ===
    // (WICHTIG: ALLES abbrechen – keine weiteren Timer, kein Speech)
    setAutoplay(false);

    try { speechSynthesis.cancel(); } catch (e) {}

    if (state.autoplay.timers && state.autoplay.timers.length > 0) {
        state.autoplay.timers.forEach(id => clearTimeout(id));
    }

    state.autoplay.timers = [];

    // Verhindert ein sofortiges Wiederstarten
    return;
}


/* ============================ WAKE LOCK ============================ */

async function requestWakeLock() {
    try {
        if ('wakeLock' in navigator && !state.wakeLock) {
            state.wakeLock = await navigator.wakeLock.request("screen");
            state.wakeLock.addEventListener?.("release", () => {
                state.wakeLock = null;
            });
            document.addEventListener("visibilitychange", onVisibilityChange, {
                passive: true
            });
        }
    } catch (e) {}
}

function onVisibilityChange() {
    if (
        document.visibilityState === "visible" &&
        state.autoplay.on &&
        !state.wakeLock
    ) {
        requestWakeLock();
    }
}

function releaseWakeLock() {
    try {
        if (state.wakeLock) state.wakeLock.release?.();
    } catch (e) {}

    state.wakeLock = null;
    document.removeEventListener("visibilitychange", onVisibilityChange);
}


/* ============================ AUTOPLAY SAFETY ============================ */

function stopAutoplayOnUserAction() {
    console.log("autoplay state:", state.autoplay.on);

    if (state.autoplay.on) {
        console.log("stopping autoplay");
        setAutoplay(false);
        speechSynthesis.cancel();
        state.autoplay.timers.forEach(id => clearTimeout(id));
        state.autoplay.timers = [];
    }
    console.log("SCROLL TOP");
    scrollToTop();
}

/* ========================================================================== */
/*                                ENDE TEIL 3                                 */
/* ========================================================================== */

/* ========================================================================== */
/*                           TEIL 4 – INIT & EVENTS                           */
/* ========================================================================== */

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

/* ================================
   Asset-Versionierung aktivieren
   ================================ */
const css = document.querySelector("#cssMain");
const js  = document.querySelector("#jsMain");

if (css) css.href = `assets/css/style.css?v=${APP_VERSION}`;
if (js)  js.src  = `assets/js/app.js?v=${APP_VERSION}`;
    console.log("[INIT] Starte Initialisierung …");

    /* ============================================================
       SETTINGS + PROGRESS LADEN + THEME & DELAY INITIALISIEREN
       ============================================================ */
    loadSettings();
    loadProgress();
    resetSessionStats();  // Session-Stats initialisieren
	
	// ✅ Resume-Fortschritt pro Lektion initialisieren
if (!state.settings.resumeIndexByLesson) {
    state.settings.resumeIndexByLesson = {};
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

    const delayInput = document.querySelector("#delayInput");
    if (delayInput) delayInput.value = state.sentenceDelay / 1000;

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

    /* Slider auf gespeicherte Werte setzen */
    const rateDeRange  = document.querySelector("#rateDeRange");
    const pitchDeRange = document.querySelector("#pitchDeRange");
    const rateZhRange  = document.querySelector("#rateZhRange");
    const pitchZhRange = document.querySelector("#pitchZhRange");

    const rateDeVal  = document.querySelector("#rateDeVal");
    const pitchDeVal = document.querySelector("#pitchDeVal");
    const rateZhVal  = document.querySelector("#rateZhVal");
    const pitchZhVal = document.querySelector("#pitchZhVal");

    if (rateDeRange)  rateDeRange.value  = state.rateDe;
    if (pitchDeRange) pitchDeRange.value = state.pitchDe;
    if (rateZhRange)  rateZhRange.value  = state.rateZh;
    if (pitchZhRange) pitchZhRange.value = state.pitchZh;

    if (rateDeVal)  rateDeVal.textContent  = `(${state.rateDe.toFixed(2)})`;
    if (pitchDeVal) pitchDeVal.textContent = `(${state.pitchDe.toFixed(2)})`;
    if (rateZhVal)  rateZhVal.textContent  = `(${state.rateZh.toFixed(2)})`;
    if (pitchZhVal) pitchZhVal.textContent = `(${state.pitchZh.toFixed(2)})`;

    /* ============================================================
       AUTOPLAY BUTTON In TRAINING-GRUPPE SETZEN
       ============================================================ */
    (function placeAutoplayButton() {
        const trainingBtn = document.querySelector("#btnStart");
        const autoplayBtn = document.querySelector("#btnAutoplay");

        if (!trainingBtn || !autoplayBtn) return;

        const parent = trainingBtn.parentNode;
        let group = parent.querySelector(".training-group");

        if (!group) {
            group = document.createElement("div");
            group.className = "training-group";
            parent.insertBefore(group, trainingBtn);
            group.appendChild(trainingBtn);
        }

        group.appendChild(autoplayBtn);
        autoplayBtn.classList.add("primary");
    })();

 
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
  
    /* DELAY INPUT */
    if (delayInput) {
        delayInput.addEventListener("input", (e) => {
            const seconds = parseFloat(e.target.value) || 0;
            state.sentenceDelay = seconds * 1000;
            state.settings.sentenceDelay = state.sentenceDelay;
            saveSettings();
        });
    }

    /* ============================================================
       STIMMEN-EINSTELLUNG
       ============================================================ */
    rateDeRange?.addEventListener("input", (e) => {
        stopAutoplayOnUserAction();
        state.rateDe = parseFloat(e.target.value);
        state.settings.rateDe = state.rateDe;
        rateDeVal.textContent = `(${state.rateDe.toFixed(2)})`;
        saveSettings();
    });

    pitchDeRange?.addEventListener("input", (e) => {
        stopAutoplayOnUserAction();
        state.pitchDe = parseFloat(e.target.value);
        state.settings.pitchDe = state.pitchDe;
        pitchDeVal.textContent = `(${state.pitchDe.toFixed(2)})`;
        saveSettings();
    });

    rateZhRange?.addEventListener("input", (e) => {
        stopAutoplayOnUserAction();
        state.rateZh = parseFloat(e.target.value);
        state.settings.rateZh = state.rateZh;
        rateZhVal.textContent = `(${state.rateZh.toFixed(2)})`;
        saveSettings();
    });

    pitchZhRange?.addEventListener("input", (e) => {
        stopAutoplayOnUserAction();
        state.pitchZh = parseFloat(e.target.value);
        state.settings.pitchZh = state.pitchZh;
        pitchZhVal.textContent = `(${state.pitchZh.toFixed(2)})`;
        saveSettings();
    });

    document.querySelector("#btnVoiceDe")?.addEventListener("click", () => {
        stopAutoplayOnUserAction();
        openVoicesPanelFor("de");
    });

    document.querySelector("#btnVoiceZh")?.addEventListener("click", () => {
        stopAutoplayOnUserAction();
        openVoicesPanelFor("zh");
    });

    document.querySelector("#btnCloseVoices")?.addEventListener("click", () => {
        closeVoices();
    });

    /* ============================================================
       MODUS, REIHENFOLGE, AUTOPLAY
       ============================================================ */
    $("#btnSwapMode").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        state.mode = state.mode === "de2zh" ? "zh2de" : "de2zh";
        state.settings.mode = state.mode;
        saveSettings();
        renderModeUI();
        if (state.current) setCard(state.current);
    });

    $("#btnOrderToggle").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        state.order = state.order === "random" ? "seq" : "random";
        state.settings.order = state.order;
        saveSettings();
        renderModeUI();
    });

    $("#btnToggleHanzi").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        state.showHanzi = !state.showHanzi;
        state.settings.showHanzi = state.showHanzi;
        saveSettings();
        renderDisplayToggleUI();
        if (state.current) setCard(state.current, true);
    });

    $("#btnTogglePinyin").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        state.showPinyin = !state.showPinyin;
        state.settings.showPinyin = state.showPinyin;
        saveSettings();
        renderDisplayToggleUI();
        if (state.current) setCard(state.current, true);
    });

    $("#btnAutoplay").addEventListener("click", () => {
        toggleAutoplay();
    });

    $("#gapRange").addEventListener("input", (e) => {
        stopAutoplayOnUserAction();
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
        stopAutoplayOnUserAction();
        startTraining();
    });

    $("#btnNext").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        nextCard();
    });

    $("#btnPrev").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        prevCard();
    });

    $("#btnReveal").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        doReveal();
    });

    /* ============================================================
       AUDIO SPRECHER
       ============================================================ */
    $("#speakerQuestion").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        playQuestion();
    });

    $("#speakerAnswer").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        playAnswer();
    });

    /* ============================================================
       RATING
       ============================================================ */

    $("#btnRateKnown").addEventListener("click", () => {
        stopAutoplayOnUserAction();
        rate("known");
    });

    $("#btnRateUnknown").addEventListener("click", () => {
        stopAutoplayOnUserAction();
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

/* ========================================================================== */
/* ENDE TEIL 4 */
/* ========================================================================== */
/* ========================================================================== */
/*                                ENDE TEIL 4                                 */
/* ========================================================================== */