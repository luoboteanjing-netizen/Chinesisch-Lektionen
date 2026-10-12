/* Autoplay and screen wake lock */
import { state, $, translate, saveSettings } from "./state.js";
import { gatherPoolFromSettings } from "./csv.js";
import {
    setCard,
    updateModeButtons,
    updateTrainingBtn,
    hideNavButtons,
    hideRatingButtons,
    scrollToTop,
    scrollToBottom,
    disableRating,
    clearChoiceBox
} from "./card.js";
import { buildAudioUrl, buildAudioUrlDe, playAudioResource, buildUtterance, ttsPrime } from "./tts.js";

/* ============================ AUTOPLAY ============================ */

function setAutoplay(on) {
    state.autoplay.on = on;
    if (on) {
        state.choiceMode = false;
        clearChoiceBox();
    }
	updateModeButtons()
    if (!on) {
        speechSynthesis.cancel();
        state.autoplay.timers.forEach(x => clearTimeout(x));
        state.autoplay.timers = [];
        releaseWakeLock();
		scrollToTop();
		
    }
    state.trainingOn = false;
	updateTrainingBtn();
    updateAutoplayBtn();
	hideNavButtons();
	hideRatingButtons();
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

    if (!state.autoplay.on) { done && done(); return; }

    const ghVoice = state.settings.githubVoiceZh;
    const ghSpeed = state.settings.githubSpeedZh || 'slow';

    // ── MP3-Zweig: GitHub-Stimme Chinesisch ──────────────────
    if (ghVoice && langKey === 'zh') {
        (async () => {
            const wUrl = buildAudioUrl(state.current, ghVoice, ghSpeed, 'words');
            const sUrl = buildAudioUrl(state.current, ghVoice, ghSpeed, 'sentences');
            if (wUrl) await playAudioResource(wUrl);
            if (!state.autoplay.on) return;
            if (sUrl) {
                await new Promise(r => setTimeout(r, 400));
                if (!state.autoplay.on) return;
                await playAudioResource(sUrl);
            }
            if (!state.autoplay.on) return;
            done && done();
        })();
        return;
    }

    // ── MP3-Zweig: GitHub-Stimme Deutsch ─────────────────────
    const ghVoiceDe = state.settings.githubVoiceDe;
    const ghSpeedDe = state.settings.githubSpeedDe || 'normal';
    if (ghVoiceDe && langKey === 'de') {
        (async () => {
            const wUrl = buildAudioUrlDe(state.current, ghVoiceDe, ghSpeedDe, 'words');
            const sUrl = buildAudioUrlDe(state.current, ghVoiceDe, ghSpeedDe, 'sentences');
            if (wUrl) await playAudioResource(wUrl);
            if (!state.autoplay.on) return;
            if (sUrl) {
                await new Promise(r => setTimeout(r, 400));
                if (!state.autoplay.on) return;
                await playAudioResource(sUrl);
            }
            if (!state.autoplay.on) return;
            done && done();
        })();
        return;
    }

    // ── Fallback: Browser-TTS ─────────────────────────────────
    const u1 = buildUtterance(word, langKey);
    u1.onend = () => {
        if (!state.autoplay.on) return;
        const t = setTimeout(() => {
            if (!state.autoplay.on) return;
            const u2 = buildUtterance(sent, langKey);
            u2.onend = () => { if (!state.autoplay.on) return; done && done(); };
            speechSynthesis.speak(u2);
        }, 650);
        state.autoplay.timers.push(t);
    };
    speechSynthesis.speak(u1);
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
    if (state.autoplay.on) {
        setAutoplay(false);
        speechSynthesis.cancel();
        state.autoplay.timers.forEach(id => clearTimeout(id));
        state.autoplay.timers = [];
    }
    scrollToTop();
}


state.stopAutoplayOnUserAction = stopAutoplayOnUserAction;


export {
    setAutoplay,
    updateAutoplayBtn,
    ensurePoolForAutoplay,
    speakPair,
    autoplayStep,
    toggleAutoplay,
    requestWakeLock,
    onVisibilityChange,
    releaseWakeLock,
    stopAutoplayOnUserAction
};
