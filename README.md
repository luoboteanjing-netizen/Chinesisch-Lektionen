
# Vokabel-Karteikarten (Deutsch ↔︎ Chinesisch)

Eine kleine, statische Flashcards-App für den Browser. Läuft direkt via GitHub Pages oder lokal (ohne Build-Tooling). CSV-Daten liegen in `./data/`.

## Features

- **CSV-Import** aus `./data/…` (erste Zeile = Header; Zeilen mit `*` in Spalte A werden ignoriert)
- **Lern-UI**: Frage oben, Lösung per Button einblendbar
- **Text-to-Speech** (Web Speech API): Frage/Lösung vorlesen, Stimme/Pitch/Rate konfigurierbar und gespeichert (pro Sprache)
- **Richtungswechsel** per Klick (Deutsch ↔︎ Chinesisch)
- **Lernfortschritt** (pro Karte, lokal gespeichert), Anzeige je Lektion
- **Navigation**: vorherige/nächste Karte, Tastaturkürzel
- **Autoplay** mit einstellbarer Verzögerung
- **Reihenfolge**: sequentiell oder zufällig

## CSV-Spezifikation

- **Pfad**: `./data/`. Standard-Dateiname: `vocab.csv` (anpassbar über Dropdown oder URL `?csv=DATEI.csv`).
- **Header**: Erste Zeile ist ein Header und wird ignoriert.
- **Ignorierte Zeilen**: Wenn die erste Zelle einer Zeile ein `*` enthält, wird die Zeile ignoriert.
- **Spalten** (A → I):
  1. **Wort Deutsch**
  2. **Wort Pinyin**
  3. **Wortart**
  4. **Satz Pinyin**
  5. **Satz Deutsch**
  6. **Wort Hanzi**
  7. **Satz Hanzi**
  8. **ID** (eindeutig, wird für Fortschritt benötigt)
  9. **Lektion** (z. B. `1`, `2`, `3` …)

> **Hinweis:** Die Lösung wird immer in der **Zielsprache** dargestellt. Bei Chinesisch werden **nur Hanzi** für TTS gesprochen (Wort + Satz), wie gefordert.

### Beispiel (CSV-Ausschnitt)

```csv
Wort DE,Wort Pinyin,Wortart,Satz Pinyin,Satz Deutsch,Wort Hanzi,Satz Hanzi,ID,Lektion
Haus,fang1,Subst.,Wo3 ai4 fang1.,Ich liebe Häuser.,房,我爱房子,ID-001,1
*Ignoriert,*,*,*,*,*,*,*,*
gehen,qu4,Verb,Wo3 qu4 xue2xiao4.,Ich gehe zur Schule.,去,我去学校,ID-002,1
```

## Nutzung

1. **Repo clonen** oder ZIP entpacken.
2. Deine CSV-Datei nach `./data/` legen (Standard: `vocab.csv`). Optional: weitere Dateien und per `?csv=DATEI.csv` laden.
3. Öffne `index.html` direkt im Browser **oder** aktiviere GitHub Pages (Branch `main`/`docs` → *Settings → Pages*).

### GitHub Pages

- Push dieses Projekt in ein GitHub-Repository.
- Unter **Settings → Pages** den Branch wählen (z. B. `main`) und den Root-Folder.
- Nach der Veröffentlichung ist die App über `https://<dein-user>.github.io/<repo-name>/` erreichbar.

## Steuerung & Tastenkürzel

- `Leertaste`: Lösung ein-/ausblenden bzw. zur nächsten Karte springen
- `← / →`: vorherige / nächste Karte
- `1` / `2`: als **Gewusst** / **Nicht gewusst** markieren

## Technische Hinweise

- **TTS** nutzt die **Web Speech API** des Browsers. Stimmen variieren je nach Betriebssystem/Browser. Für Chinesisch wird, falls verfügbar, automatisch eine `zh-*`-Stimme vorgeschlagen.
- **Speicherung** erfolgt lokal per `localStorage` (Fortschritt und TTS-Einstellungen je Sprache, CSV-Dateiname, Reihenfolge, Sprachrichtung).
- **Autoplay**: Spielt Frage → (Delay) → zeigt Lösung → (Delay) → nächste Karte; dabei wird jeweils die passende TTS-Stimme genutzt.

## Entwicklung

Keine externen Abhängigkeiten. Alles Vanilla JS/CSS/HTML. CSV-Parser ist integriert und unterstützt Anführungszeichen, Kommas und Zeilenumbrüche.

## Lizenz

MIT – siehe `LICENSE`.
