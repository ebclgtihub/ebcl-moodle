/**
 * Bestellliste vom Marktplatz Lernapps (order_*.csv) einlesen.
 *
 * Format laut Export: UTF-8, Trennzeichen „;“,
 * Spalten firstname;lastname;email;skz;bpk;role;grades
 *
 * Anmeldename in Moodle = E-Mail. Moodle erlaubt in Anmeldenamen standardmäßig nur
 * Kleinbuchstaben, Ziffern und - _ . @ (Moodle-Doku „Site security settings“).
 * Die bPK wird bewusst nicht übernommen (Datenminimierung).
 */

export const ORDER_COLUMNS = ['firstname', 'lastname', 'email', 'skz', 'bpk', 'role', 'grades'];

const USERNAME_OK = /^[a-z0-9._@-]+$/;
const EMAIL_SHAPE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
const ROLE_MAP = { lehrerin: 'T', 'schülerin': 'S' };

/** Zerlegt eine CSV-Zeile mit „;“ — berücksichtigt Anführungszeichen. */
function splitLine(line) {
  const out = [];
  let cur = '', quoted = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (quoted) {
      if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') quoted = false;
      else cur += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ';') { out.push(cur); cur = ''; }
    else cur += ch;
  }
  out.push(cur);
  return out.map(v => v.trim());
}

/**
 * @param {string} text  Inhalt der CSV
 * @returns {{
 *   persons: {first:string,last:string,email:string,isT:boolean,grade:string,line:number}[],
 *   classes: {grade:string,count:number}[],
 *   teachers: number,
 *   skz: string[],
 *   problems: {line:number,name:string,email:string,reason:string}[],
 * }}
 */
export function parseOrderCsv(text) {
  const lines = String(text).replace(/^﻿/, '').split(/\r?\n/).filter(l => l.trim() !== '');
  if (!lines.length) throw new Error('Die Datei ist leer.');

  const header = splitLine(lines[0]).map(h => h.toLowerCase());
  const missing = ORDER_COLUMNS.filter(c => !header.includes(c));
  if (missing.length) {
    throw new Error(`Das ist keine Marktplatz-Bestellliste — es fehlen die Spalten: ${missing.join(', ')}. Erwartet: ${ORDER_COLUMNS.join(';')}`);
  }
  const col = name => header.indexOf(name);

  const rows = lines.slice(1).map((l, i) => {
    const v = splitLine(l);
    return {
      line: i + 2,
      first: v[col('firstname')] || '',
      last: v[col('lastname')] || '',
      email: (v[col('email')] || '').toLowerCase(),
      skz: v[col('skz')] || '',
      role: (v[col('role')] || '').toLowerCase(),
      grade: v[col('grades')] || '',
    };
  });

  const problems = [];
  const name = r => `${r.first} ${r.last}`.trim();
  const reject = (r, reason) => problems.push({ line: r.line, name: name(r), email: r.email, reason });

  // Mehrfach vorkommende E-Mails: keine der Zeilen anlegen — ein Anmeldename gehört zu genau einem Konto
  const count = new Map();
  rows.forEach(r => { if (r.email) count.set(r.email, (count.get(r.email) || 0) + 1); });

  const persons = [];
  rows.forEach(r => {
    const kind = ROLE_MAP[r.role];
    if (!r.email) return reject(r, 'keine E-Mail-Adresse');
    if (!EMAIL_SHAPE.test(r.email) || !USERNAME_OK.test(r.email)) return reject(r, 'E-Mail enthält Zeichen, die Moodle im Anmeldenamen nicht erlaubt');
    if (count.get(r.email) > 1) return reject(r, `E-Mail kommt ${count.get(r.email)}× vor`);
    if (!r.first || !r.last) return reject(r, 'Vor- oder Nachname fehlt');
    if (!kind) return reject(r, `unbekannte Rolle „${r.role}“`);
    if (kind === 'S' && !r.grade) return reject(r, 'Schüler:in ohne Klasse');
    persons.push({ first: r.first, last: r.last, email: r.email, isT: kind === 'T', grade: kind === 'S' ? r.grade : '', line: r.line });
  });

  const classCount = new Map();
  persons.filter(p => !p.isT).forEach(p => classCount.set(p.grade, (classCount.get(p.grade) || 0) + 1));
  const classes = [...classCount].map(([grade, n]) => ({ grade, count: n })).sort((a, b) => a.grade.localeCompare(b.grade, 'de'));

  return {
    persons,
    classes,
    teachers: persons.filter(p => p.isT).length,
    skz: [...new Set(rows.map(r => r.skz).filter(Boolean))],
    problems,
  };
}
