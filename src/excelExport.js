/**
 * Zugangsdaten-Excel nach Victors Vorlage (Zugangsdaten-Name_2026.xlsx).
 *
 * Reine Funktion ohne React-Abhängigkeit — App.jsx bereitet die Daten auf,
 * hier wird nur das Workbook gebaut. exceljs statt SheetJS, weil die
 * Community-Edition von SheetJS keine Zellformate (Füllung, Rahmen) schreibt.
 */
import ExcelJS from 'exceljs';

export const LOGIN_URL = 'https://world.ebcl.eu/';
export const EXISTING_PW_TEXT = 'bestehendes Konto – bisheriges Passwort';

const FONT = { name: 'Aptos', size: 12 };
const ZEBRA = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
const THIN = { style: 'thin' };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const ROW_HEIGHT = 20;

const EXPLANATION_ANONYMOUS =
  'Auf den weiteren Tabellenblättern finden Sie anonymisierte Zugänge für die Lehrer:innen und die Schüler:innen. \n' +
  'Sie können in der Zusatzspalte die Namen der Schüler:innen eintragen. Das erleichtert die Lernerfolgskontrolle.\n' +
  'Alternativ dazu können Sie die Schüler:innen auffordern in der Lernplattform ihre Daten zu personalisieren.';

const EXPLANATION_PERSONAL =
  'Auf den weiteren Tabellenblättern finden Sie die persönlichen Zugänge der Lehrer:innen und Schüler:innen laut Bestellung.\n' +
  'Anmeldename ist jeweils die bei der Bestellung angegebene E-Mail-Adresse.\n' +
  'Bei bereits bestehenden Konten bleibt das bisherige Passwort gültig.';

/** Kursnummer im Format „Kurs 01“. */
export const courseTitle = n => `Kurs ${String(n).padStart(2, '0')}`;

/** Dateiname laut Vorlage: Zugangsdaten-{Institut}_{Jahr}.xlsx */
export function excelFileName(institute, enrolDate) {
  const year = new Date(enrolDate).getFullYear() || new Date().getFullYear();
  const inst = String(institute || '').trim().replace(/\s+/g, '-').replace(/[\\/:*?"<>|]/g, '');
  return `Zugangsdaten-${inst}_${year}.xlsx`;
}

function styleCell(cell, { bold = false, center = true, fill = null } = {}) {
  cell.font = { ...FONT, bold };
  cell.border = BORDER;
  cell.alignment = { vertical: 'middle', horizontal: center ? 'center' : undefined };
  if (fill) cell.fill = fill;
}

function setLink(cell, url) {
  cell.value = { text: url, hyperlink: url };
}

/**
 * Baut ein Konten-Blatt (Trainer oder Klasse) — beide sind gleich aufgebaut.
 * accounts: { user, pw, name, existing, courses:[{id}] }
 * courses:  Kurse dieses Blatts mit blattübergreifender Nummer { id, shorthand, label, num }
 */
function addAccountSheet(wb, name, accounts, courses, nameHeader) {
  // Standardhöhe fürs ganze Blatt: auch Leerzeilen und von Lehrkräften ergänzte Zeilen sind 20 hoch
  const ws = wb.addWorksheet(name, { properties: { defaultRowHeight: ROW_HEIGHT } });
  ws.columns = [
    { width: 20.8 }, { width: 34 }, { width: 18 }, { width: 34 },
    ...courses.map(() => ({ width: 40 })),
  ];
  const header = ['Zugang', 'Anmeldename', 'Passwort', nameHeader, ...courses.map(c => courseTitle(c.num))];
  const headerRow = ws.addRow(header);
  headerRow.height = ROW_HEIGHT;
  headerRow.eachCell(cell => styleCell(cell, { bold: true }));

  accounts.forEach((a, i) => {
    const row = ws.addRow([
      null,
      a.user,
      a.existing ? EXISTING_PW_TEXT : a.pw,
      a.name || '',
      // Kurzname ohne Link — Kurslinks führten ohne Login als Gast zum Produkt
      ...courses.map(c => ((a.courses || []).some(ac => String(ac.id) === String(c.id)) ? (c.shorthand || c.label) : '')),
    ]);
    setLink(row.getCell(1), LOGIN_URL);
    row.height = ROW_HEIGHT;
    const fill = i % 2 === 1 ? ZEBRA : null; // jede zweite Datenzeile
    for (let col = 1; col <= header.length; col++) styleCell(row.getCell(col), { fill });
  });
}

/**
 * @param {object} p
 * @param {string} p.institute
 * @param {string} p.dateStr     Erstellungsdatum (de-DE)
 * @param {string} p.periodStr   Freischaltzeitraum (de-DE)
 * @param {object[]} p.trainers  Konten der Trainer
 * @param {{label:string, accounts:object[]}[]} p.classes
 * @param {boolean} [p.personal] true = Konten aus Bestellliste (Klarnamen)
 * @returns {Promise<ArrayBuffer>}
 */
export async function buildAccessWorkbook({ institute, dateStr, periodStr, trainers = [], classes = [], personal = false }) {
  const wb = new ExcelJS.Workbook();

  // Kurse blattübergreifend nummerieren (Reihenfolge des ersten Auftretens)
  const numbered = new Map();
  const register = accs => accs.forEach(a => (a.courses || []).forEach(c => {
    const key = String(c.id);
    if (!numbered.has(key)) numbered.set(key, { ...c, num: numbered.size + 1 });
  }));
  register(trainers);
  classes.forEach(c => register(c.accounts));
  const coursesOf = accs => {
    const ids = new Set();
    accs.forEach(a => (a.courses || []).forEach(c => ids.add(String(c.id))));
    return [...numbered.values()].filter(c => ids.has(String(c.id)));
  };

  // ─── Übersicht ───
  const ov = wb.addWorksheet('Übersicht', { properties: { defaultRowHeight: ROW_HEIGHT } });
  ov.columns = [{ width: 25.8 }, { width: 60 }, { width: 16.8 }, { width: 50.8 }];
  const total = trainers.length + classes.reduce((s, c) => s + c.accounts.length, 0);
  const put = (values, opts = {}) => {
    const row = ov.addRow(values);
    row.height = ROW_HEIGHT;
    values.forEach((v, i) => { if (v !== null && v !== undefined) styleCell(row.getCell(i + 1), { center: false, ...opts }); });
    return row;
  };
  put(['Zugangsdaten – Übersicht'], { bold: true });
  put(['Institut', institute]);
  put(['Datum:', dateStr]);
  put(['Freischaltzeitraum:', periodStr]);
  put(['Gesamt-Accounts:', total]);
  setLink(put(['Zugang:', LOGIN_URL]).getCell(2), LOGIN_URL);
  ov.addRow([]).height = ROW_HEIGHT;
  put(['Gruppe', 'Typ', 'Anzahl Accounts']);
  if (trainers.length) put(['Trainer', 'Trainer', trainers.length]);
  classes.forEach(c => put([c.label, 'Schüler', c.accounts.length]));
  ov.addRow([]).height = ROW_HEIGHT;
  numbered.forEach(c => put([courseTitle(c.num), c.label]));
  ov.addRow([]).height = ROW_HEIGHT;
  put(['ERLÄUTERUNGEN'], { bold: true });
  const expl = ov.addRow([personal ? EXPLANATION_PERSONAL : EXPLANATION_ANONYMOUS]);
  ov.mergeCells(expl.number, 1, expl.number, 4);
  expl.height = 65;
  styleCell(expl.getCell(1), { center: false });
  expl.getCell(1).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };

  // ─── Konten-Blätter ───
  if (trainers.length) addAccountSheet(wb, 'Trainer', trainers, coursesOf(trainers), 'Name');
  classes.forEach(c => {
    const sheetName = c.label.replace(/[\\/?*[\]:]/g, '').substring(0, 31);
    addAccountSheet(wb, sheetName, c.accounts, coursesOf(c.accounts), 'Name Schüler:in');
  });

  return wb.xlsx.writeBuffer();
}
