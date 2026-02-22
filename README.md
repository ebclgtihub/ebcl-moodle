# EBCL-Moodle

Tauri-Desktop-App für die EBCL-Moodle-Verwaltung (React + Rust).

---

## Entwicklung

```bash
npm install     # Abhängigkeiten installieren + Git-Hook aktivieren
npm run tauri dev   # App im Entwicklungsmodus starten
```

---

## Release-Workflow

### 1. Release Notes schreiben

Änderungen in [RELEASE_NOTES.md](RELEASE_NOTES.md) unter `## Unreleased` eintragen:

```markdown
## Unreleased
- Neue Funktion X
- Bug Fix Y
```

### 2. Committen mit Conventional Commits

```bash
git add .
git commit -m "feat: neue Funktion X"
```

Der Git-Hook (`commit-msg`) läuft automatisch und:
- liest den Commit-Typ aus der Message
- bumpt die Versionsnummer in `package.json` und `src-tauri/tauri.conf.json`
- verschiebt den `## Unreleased`-Block in `RELEASE_NOTES.md` zur neuen Version
- staged alle geänderten Dateien in den Commit

**Bump-Regeln:**

| Commit-Präfix | Bump | Beispiel |
|---|---|---|
| `fix:` | patch | 0.3.0 → 0.3.1 |
| `feat:` | minor | 0.3.0 → 0.4.0 |
| `feat!:` / `BREAKING CHANGE` | major | 0.3.0 → 1.0.0 |
| `chore:`, `docs:`, sonstige | — | keine Änderung |

### 3. Pushen

```bash
git push
```

GitHub Actions baut die App für macOS und Windows und erstellt automatisch ein GitHub Release mit den Patch Notes aus `RELEASE_NOTES.md`.

---

## Projektstruktur

```
├── .githooks/
│   └── commit-msg          # Auto-Versionierung bei Commit
├── .github/workflows/
│   └── release.yml         # Build & GitHub Release
├── scripts/
│   ├── bump-version.js     # Bumpt Version in allen Dateien
│   └── extract-notes.js    # Extrahiert Release Notes für CI
├── src/                    # React-Frontend
├── src-tauri/              # Rust-Backend (Tauri)
└── RELEASE_NOTES.md        # Changelog
```

---

## Neuer Rechner / Clone

Nach `npm install` wird der Git-Hook automatisch via `prepare`-Script aktiviert. Kein manueller Schritt nötig.
