# Release Notes

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
- fix: Favoriten-Funktion entfernt (durch CRM-Einbindung ersetzt)

---

## v0.9.0 - 2026-03-09
- feat: direct moodle upload


## v0.7.0 - 2026-03-09
- Feat: Excel Export!


## v0.6.0 - 2026-03-09
- Neue Funktionen: 
    -SharePoint Export
    -PowerAutomate Link einstellbar
- Neue PDF änderungen!


## v0.5.2 - 2026-03-06
- Kein Changelog-Eintrag.
## v0.5.1 - 2026-03-06
- Kein Changelog-Eintrag.
## v0.5.0 - 2026-03-05
- Kein Changelog-Eintrag.
## v0.4.8 - 2026-03-05
- Besserer PDF Export 
- Speicherung der Passworteinstellung

## v0.4.5 - 2026-02-22
Test release 2.0

## v0.4.4 - 2026-02-22
- Kein Changelog-Eintrag.
## v0.4.3 - 2026-02-22
- Test Release

## v0.4.1 - 2026-02-22
- PDF Spalte "Name" in Formfeld umgewandelt!

## v0.3.0 - 2026-02-22
- QR-Code-Anzeige für Schüler hinzugefügt
- Automatische Passwort-Generierung implementiert
- Offline-Cache für Kursdaten
- Window-State wird beim Schließen gespeichert

## v0.2.1 - 2026-01-01
- PDF-Export-Fix
- Daten-Reset beim Programmstart

## v0.2.0 - 2025-12-01
- Auto-Updater integriert
- UI-Verbesserungen

## v0.1.6 - 2025-11-01
- Erstes stabiles Release