# Excel nach Vorlage & Bestelllisten-Import

Stand: 28.09.2026 · Auslöser: Victors Vorlage `Zugangsdaten-Name_2026.xlsx` und Marktplatz-Bestellungen (`order_*.csv`)

## Teil 2 — Excel-Export nach Vorlage (zuerst)

- Dateiname: `Zugangsdaten-{Institut}_{Jahr des Einschreibebeginns}.xlsx`
- Bibliothek: `exceljs` (SheetJS CE schreibt keine Zellformate — [Doku](https://docs.sheetjs.com/docs/getting-started/examples/export/))
- Blatt **Übersicht**: Titel „Zugangsdaten – Übersicht“, Institut, Datum, Freischaltzeitraum, Gesamt-Accounts, Zugang (Link);
  Tabelle Gruppe | Typ | Anzahl Accounts; darunter je Kurs `Kurs 01 | voller Kursname`; Block „ERLÄUTERUNGEN“.
- Blätter **Trainer** / **Klasse-xx**: `Zugang` (Link auf https://world.ebcl.eu/) | `Anmeldename` | `Passwort` | `Name Schüler:in` | `Kurs 01…`
  - Kurszellen: Kurzname mit Link zum Kurs (Gast-Ansicht kam vom aktivierten Gastzugang im Kurs, nicht vom Link — Korrektur 29.09.2026)
  - Kursnummern sind blattübergreifend gleich und passen zur Übersicht
  - Bestehende Konten (Passwort nicht gesetzt): „bestehendes Konto – bisheriges Passwort“
- Format: Schrift 12, Zeilenhöhe 20, dünne Rahmen, zentriert, jede 2. Datenzeile grau (Vorlage: Designfarbe „Hintergrund 2“)
- SharePoint-Zielordner `01 Marktplatz/! Regelbetrieb 2026-27/01 Bestellungen`: Sache des Power-Automate-Flows, nicht der App

## Teil 1 — Bestellliste importieren (danach)

- CSV vom Marktplatz, UTF-8, `;`, Spalten `firstname;lastname;email;skz;bpk;role;grades`
- `grades` → Klasse, `LehrerIn` → Trainer, `bpk` wird verworfen
- Anmeldename = E-Mail (kleingeschrieben); Moodle erlaubt `@ . - _` in Anmeldenamen ([Doku](https://docs.moodle.org/405/en/Site_security_settings))
- Mehrfach vorkommende E-Mail → Zeilen **nicht** anlegen, im Import-Dialog melden
- Konto existiert bereits → übernehmen, Passwort unverändert, in der Excel entsprechend vermerkt
- Abgleich per exaktem Anmeldenamen (`core_user_get_users_by_field`), nicht per `%@institut.com`
