# Release Notes

## v2.0.0 - 2026-03-19

### Aktualisieren-Modus
- **Klassen aus Moodle laden** — bestehende Klassen eines Instituts werden direkt aus Moodle geladen (Gruppen + Mitglieder)
- **Kombiniert-Modus „Aktualisieren & Neu anlegen"** — bestehende Klassen aktualisieren und neue Klassen anlegen in einem einzigen Durchlauf
- Drei-Button-Toggle für den Einschreibe-Modus: prominenter Haupt-Button + zwei kleinere Buttons mit Tooltips
- Trainer werden von der Mitgliederanzahl subtrahiert (korrekte Schüler-Zählung im Klassenpool)
- Standardmäßig keine Klassen vorausgewählt — bewusste Auswahl durch Nutzer

### Klassenübersicht-Popup
- Neues „Klassenübersicht"-Modal zeigt pro Klasse: Name, Mitgliederzahl, Auswahlstatus
- Abschnitt „Bereits eingeschrieben": zeigt bestehende Kurs-Einschreibungen (blau = im Pool, grau = nicht im Pool)
- Abschnitt „Neu zugewiesen": zusätzliche Matrix-Zuweisungen in grün

### Matrix-Verbesserungen
- **Auto-Spalten**: Wenn eine Klasse ausgewählt wird, erscheinen ihre bestehenden Kurs-Einschreibungen automatisch als Spalten in der Matrix (RefreshCw-Icon)
- **Auto-Pool-Befüllung**: Beim Laden der Klassen werden fehlende Kurs-Einschreibungen automatisch in freie Pool-Slots eingetragen; `courseSlotCount` wird bei Bedarf erhöht
- Nur ausgewählte Klassen erscheinen in der Matrix (kein Dimmen — vollständiges Ausblenden)
- Klassen-/Kurs-Zähler oben rechts zeigt nur aktive/sichtbare Einträge
- Trennzeile zwischen bestehenden und neuen Klassen im Kombiniert-Modus

### Kurspool & Einstellungen
- **Neuer Tab „Kurse"** in den Einstellungen: Toggle zwischen Excel-Liste (Power Automate) und Moodle direkt
- Excel-Modus: Vorschau aller gecachten Kurse mit Tag, Name, Shorthand
- Moodle-Modus: Katalog-Browser zum gezielten Auswählen welche Kurse in der Matrix erscheinen
- **CORS-Fix**: Kursabruf via Power Automate nutzt jetzt einen Rust-Command (umgeht CORS komplett, funktioniert mit beliebigen URLs)
- Fix: Leere `courseApiUrl` verursachte 404-Fehler auf localhost — Guard eingebaut

### Sicherheit & Persistenz
- Hardcodierte API-Tokens aus Standardkonfiguration entfernt (bestehende Installationen behalten ihre gespeicherten Werte)
- `courseApiUrl` und `sharepointUrl` werden jetzt korrekt gespeichert und beim Neustart wiederhergestellt

### UI-Bereinigung
- Institutübersicht in der Seitenleiste wiederhergestellt (nur wenn Zoho konfiguriert)
- Kursübersicht und Institutübersicht aus der Hauptansicht entfernt (nur noch in Einstellungen)

### Bugfixes
- Fix: Klassen-Offset (`cNum` → `cLabel`) für korrekte Gruppenermittlung in Moodle
- Fix: `classRows`-Initialisierungsreihenfolge in `activeMatrixCourses` behoben
- Fix: Externe API-IDs (Power Automate) vs. numerische Moodle-IDs korrekt unterschieden (RefreshCw-Icon war nie sichtbar)
- Umlaute im Institutsnamen werden automatisch ausgeschrieben (ä→ae, ö→oe, ü→ue, ß→ss)

---

## v1.5.0 - 2026-03-16

- feat: **Neu-anlegen-Modus** — vor der Generierung werden bestehende Moodle-Accounts abgefragt; neue Schüler/Trainer/Klassen werden fortlaufend ab der höchsten existierenden Nummer angelegt (kein Überschreiben)
- feat: **Bleistift-Button** in der Kurs-Matrix zum individuellen Anpassen der Schülerzahl pro Klasse; rotes X zum Zurücksetzen des Custom-Wertes
- feat: **Kurspool** aus dem Backend-Tab in den Anpassungen-Tab verschoben, Beta-Tag entfernt
- fix: Kurs-Anzahl (courseSlotCount) in der Organisation-Sektion wiederhergestellt
- fix: Neue Trainer im Neu-anlegen-Modus werden nur in Gruppen des eigenen Instituts eingeschrieben
- fix: Klassen-Offset wird aus allen aktiven Kursen ermittelt (nicht nur dem ersten)
- fix: Custom-Klassengrößen werden nicht mehr dauerhaft gespeichert (nur Session-gültig)
- chore: Anpassungen-Tab vor Backend-Tab; Reihenfolge im Tab: Kurspool → Tag-Farben → Akzentfarben

## v1.2.0 - 2026-03-16

- feat: Moodle Kursübersicht
- feat: Moodle Kurskatalog-Browser in Einstellungen (Beta)
- fix: Ordner-Name

## v1.0.0 - 2026-03-10

- feat: Zoho CRM Einbindung (Account-Suche, automatische Anlage mit "Institut"-Tag, Abschluss/Deal nach Einschreibung)
- feat: Kurs-Tags mit farbiger Anzeige im Dropdown und Kursübersicht
- feat: Anpassbare Akzentfarben und manuelle Tag-Farb-Zuweisung
- feat: Moodle-Einschreibung mit Fortschrittsbalken und Ergebnis-Popup
- feat: Chronologischer Upload-Ablauf (Moodle → SharePoint → Zoho) mit Abbruch bei Fehler
- feat: Standard-Einschreibedauer in den Einstellungen konfigurierbar
- feat: Bestätigung bei ungewöhnlichen Eingabewerten
- fix: Robustheit und Fehlerbehandlung verbessert

---

## v0.9.0 - 2026-03-09
- feat: Direkter Moodle-Upload

## v0.7.0 - 2026-03-09
- feat: Excel-Export

## v0.6.0 - 2026-03-09
- feat: SharePoint-Export; Power-Automate-Link konfigurierbar; PDF-Verbesserungen

## v0.3.0 - 2026-02-22
- feat: QR-Code für Schüler; automatische Passwortgenerierung; Offline-Cache; Window-State

## v0.2.0 - 2025-12-01
- feat: Auto-Updater; UI-Verbesserungen

## v0.1.6 - 2025-11-01
- Erstes stabiles Release
