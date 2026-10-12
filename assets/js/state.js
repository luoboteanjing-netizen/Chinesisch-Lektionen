/* Global state, settings storage, translations */
export const APP_VERSION = "7.0.1";
window.APP_VERSION = APP_VERSION;

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
        themeTitle: "🎨 Theme",
        themeDark: "Dark Mode",
        themeLightOrange: "Light‑Orange",
        themeLightWarm: "Light‑Warm",
        themeLightBlue: "Light‑Blue",
        delayLabel: "Pause zw. Wort und Satz:",
        autoplayGapLabel: "Pause zw. Karten im Autoplay:",
        settingsVersion: "Version:",
        modeSwitchTitle: "Richtung umschalten",
        orderRandom: "Zufällig",
		browseStart: "Lernen︎",
		browseStop: "Lernen",
        autoPlay: "Autoplay︎",
        autoPlayStop: "Autoplay",
        trainingStart: "Training︎",
        trainingStop: "Training",
        choiceStart: "Auswahl",
        choiceStop: "Auswahl",
        prev: "◀ Zurück",
        reveal: "Antwort zeigen",
        next: "Nächste ▶",
        rateUnknown: "❌ Nicht gewusst",
        rateKnown: "✅ Gewusst",
        voiceButtonDe: "🇩🇪 DE Stimme",
        voiceButtonZh: "🇨🇳 ZH Stimme",
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
        searchTitle: "Suche",
        searchLabel: "Begriff suchen",
        searchButton: "Suchen",
        searchResultsTitle: "Ergebnisse",
        searchPlaceholderDe: "z.B. Wort oder Satz (Deutsch)",
        searchPlaceholderZh: "z. B. 词或句子 oder cí huò jùzi",
        searchHint: "Tippe einen Begriff ein und drücke Enter.",
        searchEmptyResults: "Keine Treffer gefunden.",
        noVoicesFound: "Keine passenden Stimmen gefunden.",
        namelessVoice: "(namenlos)",
        pickVoice: "✓ Übernehmen",
        testVoice: "▶ Probehören",
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
        cardLessonTitle: "Lektion {id}",

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
        themeTitle: "🎨 Theme",
        themeDark: "Dark mode",
        themeLightOrange: "Light Orange",
        themeLightWarm: "Light Warm",
        themeLightBlue: "Light Blue",
        delayLabel: "Pause between word and sentence:",
        autoplayGapLabel: "Pause between cards in autoplay:",
        settingsVersion: "Version:",
        modeSwitchTitle: "Switch direction",
        orderRandom: "Random",
		browseStart: "Browse Cards",
		browseStop: "Browse Cards",
        autoPlay: "Autoplay︎",
        autoPlayStop: "Autoplay",
        trainingStart: "practice",
        trainingStop: "practice",
        choiceStart: "Choices",
        choiceStop: "Choices",
        prev: "◀ Back",
        reveal: "Show answer",
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
        searchTitle: "Search",
        searchLabel: "Search term",
        searchButton: "Search",
        searchResultsTitle: "Results",
        searchPlaceholderDe: "e.g. word or sentence (German)",
        searchPlaceholderZh: "e.g. 词 or cí huò jùzi (Chinese)",
        searchHint: "Type a term and press Enter.",
        searchEmptyResults: "No results found.",
        noVoicesFound: "No matching voices found.",
        namelessVoice: "(nameless)",
        pickVoice: "✓ accept",
        testVoice: "▶ Listen",
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
        cardLessonTitle: "Lesson {id}",
		
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
        themeTitle: "🎨 主题",
        themeDark: "深色模式",
        themeLightOrange: "浅橙",
        themeLightWarm: "浅暖",
        themeLightBlue: "浅蓝",
        delayLabel: "单词和句子之间要停顿：",
        autoplayGapLabel: "自动播放卡片之间的暂停：",
        settingsVersion: "版本：",
        modeSwitchTitle: "切换方向",
        orderRandom: "随机",
		browseStart: "浏览卡片",
		browseStop: "结束浏览",
        autoPlay: "自动播放",
        autoPlayStop: "自动播放",
        trainingStart: "开始学习",
        trainingStop: "开始学习",
        choiceStart: "选择",
        choiceStop: "选择",
        prev: "◀ 上一张",
        reveal: "显示答案",
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
        searchTitle: "搜索",
        searchLabel: "搜索词",
        searchButton: "搜索",
        searchResultsTitle: "结果",
        searchPlaceholderDe: "例如：单词或句子（德语）",
        searchPlaceholderZh: "例如：词或句子 cí huò jùzi",
        searchHint: "输入一个词并按回车。",
        searchEmptyResults: "未找到结果。",
        noVoicesFound: "未找到匹配语音。",
        namelessVoice: "(无名)",
        pickVoice: "✓ 接受",
        testVoice: "▶ 试听",
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
        cardLessonTitle: "课程 {id}",
		
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
        themeTitle: "🎨 Thème",
        themeDark: "Mode sombre",
        themeLightOrange: "Clair Orange",
        themeLightWarm: "Clair Chaud",
        themeLightBlue: "Clair Bleu",
        delayLabel: "Pause entre le mot et la phrase:",
        autoplayGapLabel: "Pause entre les cartes en lecture automatique :",
        settingsVersion: "Version :",
        modeSwitchTitle: "Changer de direction",
        orderRandom: "Aléatoire",
		browseStart: "apprendre",
		browseStop: "apprendre",
        autoPlay: "lecture auto︎",
        autoPlayStop: "lecture auto",
        trainingStart: "exercices",
        trainingStop: "exercices",
        choiceStart: "Choix",
        choiceStop: "Choix",
        prev: "◀ Précédent",
        reveal: "Afficher la réponse",
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
        searchTitle: "Recherche",
        searchLabel: "Rechercher",
        searchButton: "Rechercher",
        searchResultsTitle: "Résultats",
        searchPlaceholderDe: "ex. mot ou phrase (allemand)",
        searchPlaceholderZh: "ex. 词 ou cí huò jùzi (chinois)",
        searchHint: "Tapez un terme et appuyez sur Entrée.",
        searchEmptyResults: "Aucun résultat trouvé.",
        noVoicesFound: "Aucune voix correspondante trouvée.",
        namelessVoice: "(sans nom)",
        pickVoice: "✓ accepter",
        testVoice: "▶ Écouter",
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
        cardLessonTitle: "Leçon {id}",
		
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
    originalLessonOrder: [],   // preserved CSV import order (never sorted)
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
        resumeIndexByLesson: {},
        githubVoiceZh: null,
        githubSpeedZh: 'slow',
        githubVoiceDe: null,
        githubSpeedDe: 'normal'
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
    trainingOn: false,
	browseMode: false,
    choiceMode: false,
    choiceSelected: null
};

	state.reinsertQueue = [];
	state.cardCounter = 0;

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

export {
    LS_KEYS,
    TRANSLATIONS,
    $,
    state,
    saveSettings,
    loadSettings,
    saveProgress,
    loadProgress,
    translate
};
