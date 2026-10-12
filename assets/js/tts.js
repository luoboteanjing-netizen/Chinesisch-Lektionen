/* Recorded audio, browser TTS, voice panel */
import { state, $, translate, saveSettings } from "./state.js";

/* ============================ GITHUB AUDIO CONFIG ======================== */

const AUDIO_BASE = './data/audio';
const AUDIO_BASE_DE = './data/audio_de';
const GITHUB_VOICES = ['xiaoxiao', 'yunjian'];
const GITHUB_VOICES_DE = ['katja', 'conrad'];
const GITHUB_SPEEDS = ['slow'];
const GITHUB_SPEEDS_DE = ['normal'];
const CACHE_PREFIX = 'fc-audio-';

const RELEASE_BASE = 'https://pub-368b54c806bb458385aedf5cd96ac804.r2.dev';

/* --- Language-specific audio config lookup --- */
const AUDIO_LANG_CONFIG = {
    zh: { base: AUDIO_BASE, defaultSpeed: 'slow',   cacheKey: v => v },
    de: { base: AUDIO_BASE_DE, defaultSpeed: 'normal', cacheKey: v => 'de-' + v }
};

function buildZipUrl(voice, lang) {
    const speed = AUDIO_LANG_CONFIG[lang]?.defaultSpeed || 'slow';
    return `${RELEASE_BASE}/${voice}_${speed}.zip`;
}

function sanitizeFileName(fn) {
    if (!fn) return null;
    let name = fn.trim();
    name = name.replace(/^\/+/, '');
    if (!/\.mp3$/i.test(name)) name += '.mp3';
    return name;
}

function buildAudioUrl(entry, voice, speed, kind, lang = 'zh') {
    if (!entry || !entry.audio) return null;
    let raw = kind === 'words' ? entry.audio.wordFile : entry.audio.sentFile;
    if (!raw) return null;
    raw = raw.trim();
    let filename = raw;
    if (!/\.mp3$/i.test(raw)) {
        const suffix = kind === 'words' ? '_wrd.mp3' : '_snt.mp3';
        filename = raw + suffix;
    }
    filename = sanitizeFileName(filename);
    if (!filename) return null;
    const base = AUDIO_LANG_CONFIG[lang]?.base || AUDIO_BASE;
    return `${base}/${voice}/${speed}/${encodeURIComponent(filename)}`;
}

/* Backward-compat wrappers (used by existing callers) */
function buildAudioUrlDe(entry, voice, speed, kind) {
    return buildAudioUrl(entry, voice, speed, kind, 'de');
}

function playPreviewAudio(voice, speed, kind, lang = 'zh') {
    const previewEntry = {
        audio: { wordFile: 'L01-1-04_wrd.mp3', sentFile: 'L01-1-04_snt.mp3' }
    };
    const url = buildAudioUrl(previewEntry, voice, speed, kind, lang);
    if (!url) return;
    const a = new Audio(url);
    a.play().catch(e => console.warn('Play preview failed', e));
}

function playPreviewAudioDe(voice, speed, kind) {
    playPreviewAudio(voice, speed, kind, 'de');
}

async function isVoiceCached(voice, lang = 'zh') {
    try {
        const key = CACHE_PREFIX + AUDIO_LANG_CONFIG[lang].cacheKey(voice);
        const cache = await caches.open(key);
        const keys = await cache.keys();
        return keys.length > 0;
    } catch (e) { return false; }
}

async function isVoiceCachedDe(voice) {
    return isVoiceCached(voice, 'de');
}

async function deleteVoiceCache(voice, lang = 'zh') {
    try {
        const key = CACHE_PREFIX + AUDIO_LANG_CONFIG[lang].cacheKey(voice);
        await caches.delete(key);
    } catch (e) { console.warn('Cache löschen fehlgeschlagen', e); }
}

async function deleteVoiceCacheDe(voice) {
    return deleteVoiceCache(voice, 'de');
}

async function downloadVoiceZip(voice, onProgress, lang = 'zh') {
    const cfg = AUDIO_LANG_CONFIG[lang];
    const zipUrl = buildZipUrl(voice, lang);
    onProgress && onProgress(0, 'Verbinde…');
    let response;
    try { response = await fetch(zipUrl); }
    catch (e) { throw new Error('Netzwerkfehler: ' + e.message); }
    if (!response.ok) throw new Error('ZIP nicht gefunden (' + response.status + '): ' + zipUrl);

    const contentLength = response.headers.get('Content-Length');
    const total = contentLength ? parseInt(contentLength, 10) : null;
    let received = 0;
    const chunks = [];
    const reader = response.body.getReader();
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
        received += value.length;
        const mb = Math.round(received / 1024 / 1024 * 10) / 10;
        if (total) onProgress && onProgress(Math.round(received / total * 60), 'Herunterladen… ' + mb + ' MB');
        else onProgress && onProgress(10, 'Herunterladen… ' + mb + ' MB');
    }
    const zipBuffer = new Uint8Array(received);
    let offset = 0;
    for (const chunk of chunks) { zipBuffer.set(chunk, offset); offset += chunk.length; }
    onProgress && onProgress(62, 'Entpacke ZIP…');

    if (!window._fflate) {
        await new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = 'https://cdn.jsdelivr.net/npm/fflate@0.8.2/umd/index.js';
            s.onload = resolve;
            s.onerror = () => reject(new Error('fflate konnte nicht geladen werden'));
            document.head.appendChild(s);
        });
        window._fflate = fflate;
    }
    const unzipped = await new Promise((resolve, reject) => {
        window._fflate.unzip(zipBuffer, (err, result) => { if (err) reject(err); else resolve(result); });
    });
    const cacheKey = CACHE_PREFIX + cfg.cacheKey(voice);
    const cache = await caches.open(cacheKey);
    const files = Object.entries(unzipped).filter(([name]) => /\.mp3$/i.test(name));
    const fileCount = files.length;
    let done = 0;
    for (const [zipPath, data] of files) {
        const filename = zipPath.split('/').pop();
        const cacheUrl = cfg.base + '/' + voice + '/' + cfg.defaultSpeed + '/' + encodeURIComponent(filename);
        const blob = new Blob([data], { type: 'audio/mpeg' });
        await cache.put(cacheUrl, new Response(blob, { headers: { 'Content-Type': 'audio/mpeg' } }));
        done++;
        if (done % 100 === 0 || done === fileCount) {
            onProgress && onProgress(62 + Math.round(done / fileCount * 38), 'Speichere… ' + done + ' / ' + fileCount + ' Dateien');
        }
    }
    onProgress && onProgress(100, 'Fertig! ' + fileCount + ' Dateien gespeichert.');
    return fileCount;
}

async function downloadVoiceZipDe(voice, onProgress) {
    return downloadVoiceZip(voice, onProgress, 'de');
}

async function playAudioResource(url) {
    function playBlob(blob) {
        return new Promise((resolve) => {
            const obj = URL.createObjectURL(blob);
            const a = new Audio(obj);
            a.onended = () => { URL.revokeObjectURL(obj); resolve(); };
            a.onerror = () => { URL.revokeObjectURL(obj); resolve(); };
            a.play().catch(() => resolve());
        });
    }
    try {
        const cached = await caches.match(url);
        if (cached) { await playBlob(await cached.blob()); return; }
        const res = await fetch(url);
        if (res.ok) await playBlob(await res.blob());
    } catch (e) { console.warn('playAudioResource error', e); }
}

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
    document.querySelectorAll(".voice-btn").forEach(btn => btn.classList.remove("active"));
}

/* ============================ UPDATE VOICE LIST ============================ */

function appendBrowserVoiceControls(box) {
    const isZh = state.voicePanelTarget === "zh";
    const wrap = document.createElement("div");
    wrap.className = "voice-tts-config";
    wrap.appendChild(makeVoiceSlider(
        translate(isZh ? "rateZhLabel" : "rateDeLabel"),
        isZh ? state.rateZh : state.rateDe,
        0.6, 1.6, 0.05,
        (value) => {
            if (isZh) {
                state.rateZh = value;
                state.settings.rateZh = value;
            } else {
                state.rateDe = value;
                state.settings.rateDe = value;
            }
            saveSettings();
        }
    ));
    wrap.appendChild(makeVoiceSlider(
        translate(isZh ? "pitchZhLabel" : "pitchDeLabel"),
        isZh ? state.pitchZh : state.pitchDe,
        0.7, 1.5, 0.05,
        (value) => {
            if (isZh) {
                state.pitchZh = value;
                state.settings.pitchZh = value;
            } else {
                state.pitchDe = value;
                state.settings.pitchDe = value;
            }
            saveSettings();
        }
    ));
    box.appendChild(wrap);
}

function makeVoiceSlider(labelText, value, min, max, step, onInput) {
    const block = document.createElement("div");
    block.className = "voice-tts-slider";

    const label = document.createElement("label");
    label.className = "lbl";
    const hint = document.createElement("span");
    hint.className = "hint";
    hint.textContent = `(${Number(value).toFixed(2)})`;
    label.appendChild(document.createTextNode(labelText + " "));
    label.appendChild(hint);

    const input = document.createElement("input");
    input.type = "range";
    input.className = "voice-range";
    input.min = String(min);
    input.max = String(max);
    input.step = String(step);
    input.value = String(value);
    input.addEventListener("input", () => {
        const next = parseFloat(input.value);
        hint.textContent = `(${next.toFixed(2)})`;
        onInput(next);
    });

    block.appendChild(label);
    block.appendChild(input);
    return block;
}

function updateVoiceList() {
    const box = $("#dbgVoices");
    if (!box) return;

    box.innerHTML = "";

    // ── GitHub MP3-Stimmen (DE) ───────────────────────────────
    if (state.voicePanelTarget !== "zh") {
        const deHeader = document.createElement('div');
        deHeader.className = 'dbg-section';
        deHeader.innerHTML = '<div class="h">GitHub Stimmen DE (MP3)</div>';
        box.appendChild(deHeader);

        GITHUB_VOICES_DE.forEach((voiceName) => {
            const row = document.createElement('div');
            row.className = 'voice';
            const name = document.createElement('div');
            name.className = 'name';
            name.textContent = voiceName + ' (MP3)';
            if (state.settings.githubVoiceDe === voiceName) name.textContent += ' ' + translate('voiceActiveSuffix');
            const meta = document.createElement('div');
            meta.className = 'meta';
            meta.textContent = 'Prüfe Cache…';
            isVoiceCachedDe(voiceName).then(cached => {
                meta.textContent = cached ? '✅ Lokal gespeichert' : '☁️ Noch nicht heruntergeladen';
            });
            const progress = document.createElement('div');
            progress.className = 'meta';
            progress.style.display = 'none';
            progress.style.color = 'var(--accent)';
            const actions = document.createElement('div');
            actions.className = 'actions';

            const btnPreview = document.createElement('button');
            btnPreview.className = 'btn ghost';
            btnPreview.textContent = '▶ Probehören';
            btnPreview.onclick = () => playPreviewAudioDe(voiceName, 'normal', 'sentences');

            const btnDownload = document.createElement('button');
            btnDownload.className = 'btn ghost';
            btnDownload.textContent = '⬇ Herunterladen';
            btnDownload.onclick = async () => {
                [btnDownload, btnDelete, btnPick, btnPreview].forEach(b => b.disabled = true);
                progress.style.display = '';
                try {
                    const count = await downloadVoiceZipDe(voiceName, (pct, text) => {
                        progress.textContent = pct + '% – ' + text;
                    });
                    meta.textContent = '✅ Lokal gespeichert (' + count + ' Dateien)';
                    progress.style.display = 'none';
                } catch (e) {
                    progress.textContent = '❌ Fehler: ' + e.message;
                } finally {
                    [btnDownload, btnDelete, btnPick, btnPreview].forEach(b => b.disabled = false);
                }
            };

            const btnDelete = document.createElement('button');
            btnDelete.className = 'btn ghost';
            btnDelete.textContent = '🗑 Cache löschen';
            btnDelete.onclick = async () => {
                await deleteVoiceCacheDe(voiceName);
                meta.textContent = '☁️ Noch nicht heruntergeladen';
            };

            const btnPick = document.createElement('button');
            btnPick.className = 'btn';
            btnPick.textContent = '✓ Übernehmen';
            btnPick.onclick = () => {
                state.settings.githubVoiceDe = voiceName;
                state.settings.githubSpeedDe = 'normal';
                state.browserVoice.de = null;
                state.settings.browserVoiceDe = null;
                saveSettings();
                closeVoices();
            };

            actions.appendChild(btnPreview);
            actions.appendChild(btnDownload);
            actions.appendChild(btnDelete);
            actions.appendChild(btnPick);
            row.appendChild(name);
            row.appendChild(meta);
            row.appendChild(progress);
            row.appendChild(actions);
            box.appendChild(row);
        });

        // Trennlinie vor Browser-Stimmen
        const sep = document.createElement('div');
        sep.className = 'dbg-section';
        sep.innerHTML = '<div class="h">Browser Stimmen DE</div>';
        box.appendChild(sep);
    }

    // ── GitHub MP3-Stimmen (ZH) ───────────────────────────────
    if (state.voicePanelTarget === "zh") {
        const zhHeader = document.createElement('div');
        zhHeader.className = 'dbg-section';
        zhHeader.innerHTML = '<div class="h">GitHub Stimmen ZH (MP3)</div>';
        box.appendChild(zhHeader);

        GITHUB_VOICES.forEach((voiceName) => {
            const row = document.createElement('div');
            row.className = 'voice';
            const name = document.createElement('div');
            name.className = 'name';
            name.textContent = voiceName + ' (MP3)';
            if (state.settings.githubVoiceZh === voiceName) name.textContent += ' ' + translate('voiceActiveSuffix');
            const meta = document.createElement('div');
            meta.className = 'meta';
            meta.textContent = 'Prüfe Cache…';
            isVoiceCached(voiceName).then(cached => {
                meta.textContent = cached ? '✅ Lokal gespeichert' : '☁️ Noch nicht heruntergeladen';
            });
            const progress = document.createElement('div');
            progress.className = 'meta';
            progress.style.display = 'none';
            progress.style.color = 'var(--accent)';
            const actions = document.createElement('div');
            actions.className = 'actions';

            const btnPreview = document.createElement('button');
            btnPreview.className = 'btn ghost';
            btnPreview.textContent = '▶ Probehören';
            btnPreview.onclick = () => playPreviewAudio(voiceName, 'slow', 'sentences');

            const btnDownload = document.createElement('button');
            btnDownload.className = 'btn ghost';
            btnDownload.textContent = '⬇ Herunterladen';
            btnDownload.onclick = async () => {
                [btnDownload, btnDelete, btnPick, btnPreview].forEach(b => b.disabled = true);
                progress.style.display = '';
                try {
                    const count = await downloadVoiceZip(voiceName, (pct, text) => {
                        progress.textContent = pct + '% – ' + text;
                    });
                    meta.textContent = '✅ Lokal gespeichert (' + count + ' Dateien)';
                    progress.style.display = 'none';
                } catch (e) {
                    progress.textContent = '❌ Fehler: ' + e.message;
                } finally {
                    [btnDownload, btnDelete, btnPick, btnPreview].forEach(b => b.disabled = false);
                }
            };

            const btnDelete = document.createElement('button');
            btnDelete.className = 'btn ghost';
            btnDelete.textContent = '🗑 Cache löschen';
            btnDelete.onclick = async () => {
                await deleteVoiceCache(voiceName);
                meta.textContent = '☁️ Noch nicht heruntergeladen';
            };

            const btnPick = document.createElement('button');
            btnPick.className = 'btn';
            btnPick.textContent = '✓ Übernehmen';
            btnPick.onclick = () => {
                state.settings.githubVoiceZh = voiceName;
                state.settings.githubSpeedZh = 'slow';
                state.browserVoice.zh = null;
                state.settings.browserVoiceZh = null;
                saveSettings();
                closeVoices();
            };

            actions.appendChild(btnPreview);
            actions.appendChild(btnDownload);
            actions.appendChild(btnDelete);
            actions.appendChild(btnPick);
            row.appendChild(name);
            row.appendChild(meta);
            row.appendChild(progress);
            row.appendChild(actions);
            box.appendChild(row);
        });

        // Trennlinie vor Browser-Stimmen
        const sep = document.createElement('div');
        sep.className = 'dbg-section';
        sep.innerHTML = '<div class="h">Browser Stimmen ZH</div>';
        box.appendChild(sep);
    }

    // Tempo und Tonhöhe gelten nur für Browser-TTS, nicht für MP3-Stimmen
    appendBrowserVoiceControls(box);

    // ── Browser-Stimmen (wie bisher) ─────────────────────────
    const list = (state.voices || []).filter(v =>
        state.voicePanelTarget === "zh" ? isZhVoice(v) : isDeVoice(v)
    );

    if (!list.length) {
        const noVoices = document.createElement('div');
        noVoices.textContent = translate("noVoicesFound");
        box.appendChild(noVoices);
        return;
    }

    list.forEach(v => {
        const row = document.createElement("div");
        row.className = "voice";

        const name = document.createElement("div");
        name.className = "name";
        name.textContent = v.name || translate("namelessVoice");

        const meta = document.createElement("div");
        meta.className = "meta";
        meta.textContent = `${v.lang}${v.default ? " · default" : ""}`;

        const actions = document.createElement("div");
        actions.className = "actions";

        const btnPick = document.createElement("button");
        btnPick.className = "btn";
        btnPick.textContent = translate("pickVoice");
		btnPick.onclick = () => {
			if (state.voicePanelTarget === "zh") {
				state.browserVoice.zh = v;
				state.settings.browserVoiceZh = v.name || v.voiceURI;
				state.settings.githubVoiceZh = null;   // ← NEU: MP3-Stimme deaktivieren
			} else {
				state.browserVoice.de = v;
				state.settings.browserVoiceDe = v.name || v.voiceURI;
				state.settings.githubVoiceDe = null;   // ← NEU: MP3-Stimme deaktivieren
			}
			saveSettings();
			closeVoices();
		};

        const btnTest = document.createElement("button");
        btnTest.className = "btn ghost";
        btnTest.textContent = translate("testVoice");
        btnTest.onclick = () => {
            const isZh = state.voicePanelTarget === "zh";
            const u = new SpeechSynthesisUtterance(
                isZh ? "我很高兴见到你。" : "Es freut mich sehr dich zu sehen."
            );
            u.lang = isZh ? "zh-CN" : "de-DE";
            u.voice = v;
            u.rate = isZh ? state.rateZh : state.rateDe;
            u.pitch = isZh ? state.pitchZh : state.pitchDe;
            speechSynthesis.cancel();
            speechSynthesis.speak(u);
        };

        const active = state.voicePanelTarget === "zh" ? state.browserVoice.zh : state.browserVoice.de;
        if (active && (active.name === v.name || active.voiceURI === v.voiceURI)) {
            name.textContent += ` ${translate("voiceActiveSuffix")}`;
        }
        actions.appendChild(btnTest);
        actions.appendChild(btnPick);
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
        speechSynthesis.cancel();
        const ghVoiceDe = state.settings.githubVoiceDe;
        const ghSpeedDe = state.settings.githubSpeedDe || 'normal';
        if (ghVoiceDe) {
            const wUrl = buildAudioUrlDe(state.current, ghVoiceDe, ghSpeedDe, 'words');
            const sUrl = buildAudioUrlDe(state.current, ghVoiceDe, ghSpeedDe, 'sentences');
            (async () => {
                if (wUrl) await playAudioResource(wUrl);
                if (sUrl) await new Promise(r => setTimeout(r, 400));
                if (sUrl) await playAudioResource(sUrl);
            })();
            return;
        }
        playSequence(state.current.word.de, "de", state.current.sent.de, "de");
    } else {
        speechSynthesis.cancel();
        const ghVoice = state.settings.githubVoiceZh;
        const ghSpeed = state.settings.githubSpeedZh || 'slow';
        if (ghVoice) {
            const wUrl = buildAudioUrl(state.current, ghVoice, ghSpeed, 'words');
            const sUrl = buildAudioUrl(state.current, ghVoice, ghSpeed, 'sentences');
            (async () => {
                if (wUrl) await playAudioResource(wUrl);
                if (sUrl) await new Promise(r => setTimeout(r, 400));
                if (sUrl) await playAudioResource(sUrl);
            })();
            return;
        }
        ttsSpeak(state.current.word.zh, "zh");
        setTimeout(() => ttsSpeak(state.current.sent.zh, "zh"), 600);
    }
}

function playAnswer() {
    if (!state.current) return;

    if (state.mode === "de2zh") {
        const ghVoice = state.settings.githubVoiceZh;
        const ghSpeed = state.settings.githubSpeedZh || 'slow';
        if (ghVoice) {
            const wUrl = buildAudioUrl(state.current, ghVoice, ghSpeed, 'words');
            const sUrl = buildAudioUrl(state.current, ghVoice, ghSpeed, 'sentences');
            (async () => {
                if (wUrl) await playAudioResource(wUrl);
                if (sUrl) await new Promise(r => setTimeout(r, 400));
                if (sUrl) await playAudioResource(sUrl);
            })();
            return;
        }
        ttsSpeak(state.current.word.zh, "zh");
        setTimeout(() => ttsSpeak(state.current.sent.zh, "zh"), 600);
    } else {
        speechSynthesis.cancel();
        const ghVoiceDe = state.settings.githubVoiceDe;
        const ghSpeedDe = state.settings.githubSpeedDe || 'normal';
        if (ghVoiceDe) {
            const wUrl = buildAudioUrlDe(state.current, ghVoiceDe, ghSpeedDe, 'words');
            const sUrl = buildAudioUrlDe(state.current, ghVoiceDe, ghSpeedDe, 'sentences');
            (async () => {
                if (wUrl) await playAudioResource(wUrl);
                if (sUrl) await new Promise(r => setTimeout(r, 400));
                if (sUrl) await playAudioResource(sUrl);
            })();
            return;
        }
        playSequence(state.current.word.de, "de", state.current.sent.de, "de");
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

// Spielt beim Aufdecken die Antwort-Sprache ab:
// DE→ZH: Chinesisch | ZH→DE: Deutsch
function playOnReveal(entry) {
    if (!entry) return;
    speechSynthesis.cancel();

    if (state.mode === "de2zh") {
        const ghVoice = state.settings.githubVoiceZh;
        const ghSpeed = state.settings.githubSpeedZh || 'slow';
        if (ghVoice) {
            const wUrl = buildAudioUrl(entry, ghVoice, ghSpeed, 'words');
            const sUrl = buildAudioUrl(entry, ghVoice, ghSpeed, 'sentences');
            (async () => {
                if (wUrl) await playAudioResource(wUrl);
                if (sUrl) await new Promise(r => setTimeout(r, 600));
                if (sUrl) await playAudioResource(sUrl);
            })();
            return;
        }
        ttsSpeak(entry.word.zh, "zh");
        setTimeout(() => ttsSpeak(entry.sent.zh, "zh"), 600);
    } else {
        const ghVoiceDe = state.settings.githubVoiceDe;
        const ghSpeedDe = state.settings.githubSpeedDe || 'normal';
        if (ghVoiceDe) {
            const wUrl = buildAudioUrlDe(entry, ghVoiceDe, ghSpeedDe, 'words');
            const sUrl = buildAudioUrlDe(entry, ghVoiceDe, ghSpeedDe, 'sentences');
            (async () => {
                if (wUrl) await playAudioResource(wUrl);
                if (sUrl) await new Promise(r => setTimeout(r, 600));
                if (sUrl) await playAudioResource(sUrl);
            })();
            return;
        }
        ttsSpeak(entry.word.de, "de");
        setTimeout(() => ttsSpeak(entry.sent.de, "de"), 600);
    }
}

function playChineseOnReveal(entry) {
    playOnReveal(entry);
}


export {
    buildZipUrl,
    sanitizeFileName,
    buildAudioUrl,
    buildAudioUrlDe,
    playPreviewAudio,
    playPreviewAudioDe,
    isVoiceCached,
    isVoiceCachedDe,
    deleteVoiceCache,
    deleteVoiceCacheDe,
    downloadVoiceZip,
    downloadVoiceZipDe,
    playAudioResource,
    ttsPrime,
    buildUtterance,
    isZhVoice,
    isDeVoice,
    refreshVoices,
    openVoicesPanelFor,
    closeVoices,
    updateVoiceList,
    playQuestion,
    playAnswer,
    ttsSpeak,
    playSequence,
    playOnReveal,
    playChineseOnReveal
};
