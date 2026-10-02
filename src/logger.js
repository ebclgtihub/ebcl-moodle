/**
 * Diagnose-Protokoll für Supportfälle.
 *
 * Hält die Einträge im Speicher und gibt sie als Text aus, den Nutzer:innen kopieren
 * oder speichern und an den Support schicken. Mit enablePersistence() überleben die
 * Einträge einen Neustart (14 Tage) — Fehler werden oft erst Tage später gemeldet.
 *
 * Grundsatz: Kein Zugangsgeheimnis darf im Protokoll landen. Deshalb zweistufig —
 *  1. beim Aufzeichnen werden bekannte Muster maskiert (Token, Passwörter,
 *     Flow-URLs, E-Mail-Adressen),
 *  2. beim Export wird zusätzlich jeder aktuell konfigurierte Geheimwert wörtlich
 *     ersetzt, auch wenn er keinem Muster entspricht.
 */

const MAX_ENTRIES = 3000;
const MAX_AGE_DAYS = 14;
const SAVE_DELAY_MS = 3000;
const MAX_TEXT = 1200;

const entries = [];
let secretProvider = () => ({});
let consoleHooked = false;
let persistStore = null;   // Tauri-Store (get/set/save), gesetzt über enablePersistence()
let saveTimer = null;

const SECRET_FIELDS = [
  'wstoken', 'token', 'moodleToken', 'access_token', 'refresh_token', 'id_token',
  'refreshToken', 'zohoRefreshToken', 'client_secret', 'clientSecret', 'zohoClientSecret',
  'password', 'passwort', 'pwd', 'studentPwd', 'trainerPwd', 'secret', 'authorization',
].join('|');

const RULES = [
  // Power-Automate- und Logic-Apps-URLs zuerst: die Signatur steckt in der URL, nur der Host bleibt.
  // Muss vor den Query-Parametern laufen, sonst bleibt deren Ersatz hinter dem Host stehen.
  [/(https?:\/\/[a-z0-9.-]*(?:logic\.azure\.com|powerplatform\.com))(?::\d+)?[^\s"'<>]*/gi, '$1/<entfernt>'],
  // Query-Parameter, die Zugang gewähren
  [/([?&](?:wstoken|token|access_token|refresh_token|client_secret|code|sig|key)=)[^&\s"'<>]+/gi, '$1<entfernt>'],
  // Header-Token — vor den Feldnamen, sonst frisst die Feldregel das Schema-Wort
  [/\b(Bearer|Zoho-oauthtoken)\s+[A-Za-z0-9._-]+/gi, '$1 <entfernt>'],
  // Felder mit Geheimnis-Namen, in JSON wie in Objektschreibweise
  [new RegExp(`(["']?(?:${SECRET_FIELDS})["']?\\s*[:=]\\s*)("[^"]*"|'[^']*'|[^\\s,;}&]+)`, 'gi'), '$1<entfernt>'],
  // Zoho-Token und -Client-IDs (1000.xxxx…)
  [/\b1000\.[A-Za-z0-9]{10,}(?:\.[A-Za-z0-9]{10,})?/g, '<entfernt>'],
  // Lange Zufallsketten ohne Trennzeichen — Moodle-Token sind 32 Hex-Zeichen
  [/\b[a-fA-F0-9]{32,}\b/g, '<entfernt>'],
  [/\b[A-Za-z0-9]{40,}\b/g, '<entfernt>'],
  // E-Mail-Adressen: nur die Domain bleibt
  [/\b[A-Za-z0-9._%+-]+@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,})\b/g, '***@$1'],
];

/** Maskiert Geheimnisse in einem Text. `secrets` sind Werte, die wörtlich ersetzt werden. */
export function redact(input, secrets = []) {
  let s = String(input ?? '');
  for (const v of secrets) {
    if (typeof v === 'string' && v.trim().length >= 6) s = s.split(v).join('<entfernt>');
  }
  for (const [re, rep] of RULES) s = s.replace(re, rep);
  return s;
}

function stringify(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Error) return value.message || value.name;
  if (typeof value === 'string') return value;
  try { return JSON.stringify(value); } catch { return String(value); }
}

function clock(ts) {
  const d = new Date(ts);
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

/** Datum + Uhrzeit — das Protokoll umfasst mehrere Tage. */
function stamp(ts) {
  const d = new Date(ts);
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}. ${clock(ts)}`;
}

/** Speichert gebündelt — nicht bei jedem Eintrag auf die Platte schreiben. */
function scheduleSave() {
  if (!persistStore || saveTimer) return;
  saveTimer = setTimeout(() => { saveTimer = null; flushLog(); }, SAVE_DELAY_MS);
}

/** Schreibt den aktuellen Stand sofort in den Store. Fehler werden geschluckt — das Protokoll darf die App nie stören. */
export async function flushLog() {
  if (!persistStore) return;
  try {
    await persistStore.set('entries', entries);
    await persistStore.save();
  } catch { /* still */ }
}

/**
 * Lädt gespeicherte Einträge (jünger als 14 Tage) vor die aktuellen und speichert ab jetzt mit.
 * Gespeichert wird nur bereits maskierter Text (siehe log()).
 */
export async function enablePersistence(store) {
  if (persistStore || !store) return;
  try {
    const saved = await store.get('entries');
    if (Array.isArray(saved)) {
      const cutoff = Date.now() - MAX_AGE_DAYS * 86400000;
      const keep = saved.filter(e => e && typeof e.text === 'string' && (e.until || e.ts) >= cutoff);
      entries.unshift(...keep);
      if (entries.length > MAX_ENTRIES) entries.splice(0, entries.length - MAX_ENTRIES);
    }
  } catch { /* beschädigte oder fehlende Datei: neu beginnen */ }
  persistStore = store;
  if (typeof window !== 'undefined' && window.addEventListener) {
    window.addEventListener('beforeunload', () => { flushLog(); });
  }
  scheduleSave();
}

/**
 * Schreibt einen Eintrag. Folgt ein identischer Eintrag direkt auf den vorigen,
 * wird nur mitgezählt — sonst verdrängen hunderte gleiche Ablehnungen den Rest.
 */
export function log(level, area, message, detail) {
  let text = [stringify(message), stringify(detail)].filter(Boolean).join(' — ');
  text = redact(text);
  if (text.length > MAX_TEXT) text = `${text.slice(0, MAX_TEXT)} … (+${text.length - MAX_TEXT} Zeichen)`;
  const now = Date.now();
  const last = entries[entries.length - 1];
  if (last && last.level === level && last.area === area && last.text === text) {
    last.count += 1;
    last.until = now;
    scheduleSave();
    return;
  }
  entries.push({ ts: now, level, area, text, count: 1, until: now });
  if (entries.length > MAX_ENTRIES) entries.splice(0, entries.length - MAX_ENTRIES);
  scheduleSave();
}

export const logInfo = (area, message, detail) => log('info', area, message, detail);
export const logWarn = (area, message, detail) => log('warn', area, message, detail);
export const logError = (area, message, detail) => log('error', area, message, detail);

/** Schreibt Warnungen und Fehler der Konsole sowie unbehandelte Fehler mit. Einmalig aufrufen. */
export function captureConsole() {
  if (consoleHooked || typeof console === 'undefined') return;
  consoleHooked = true;
  for (const level of ['warn', 'error']) {
    const original = console[level].bind(console);
    console[level] = (...args) => {
      try { log(level, 'konsole', args.map(stringify).filter(Boolean).join(' ')); } catch { /* Protokoll darf die App nie stören */ }
      original(...args);
    };
  }
  if (typeof window !== 'undefined' && window.addEventListener) {
    window.addEventListener('error', e => log('error', 'laufzeit', e?.message || 'Unbekannter Fehler'));
    window.addEventListener('unhandledrejection', e => log('error', 'laufzeit', 'Unbehandelter Fehler', e?.reason));
  }
}

/** Liefert beim Export die aktuell konfigurierten Geheimwerte, die wörtlich entfernt werden. */
export function setSecretProvider(fn) {
  secretProvider = typeof fn === 'function' ? fn : () => ({});
}

export const entryCount = () => entries.length;
export const clearLog = () => { entries.length = 0; flushLog(); };

/** Baut den exportierbaren Text. `header` sind Schlüssel/Wert-Paare für den Kopf. */
export function buildLogText(header = {}) {
  let secrets = [];
  try { secrets = Object.values(secretProvider() || {}); } catch { secrets = []; }
  const lines = [
    'EBC*L Moodle-Anlage — Diagnose-Protokoll',
    `Erstellt: ${new Date().toLocaleString('de-AT')}`,
    ...Object.entries(header).map(([k, v]) => `${k}: ${v}`),
    `Einträge: ${entries.length} (letzte ${MAX_AGE_DAYS} Tage, höchstens ${MAX_ENTRIES}, älteste zuerst)`,
    'Zugangsdaten, Passwörter und E-Mail-Adressen sind entfernt.',
    '-'.repeat(72),
    ...entries.map(e => {
      const repeat = e.count > 1 ? ` (${e.count}× bis ${clock(e.until)})` : '';
      return `${stamp(e.ts)}  ${e.level.toUpperCase().padEnd(5)} [${e.area}] ${e.text}${repeat}`;
    }),
  ];
  return redact(lines.join('\n'), secrets);
}

/** Kopiert Text in die Zwischenablage — mit Rückfall für Webviews ohne Clipboard-API. */
export async function copyToClipboard(text) {
  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch { /* Rückfall unten */ }
  try {
    const area = document.createElement('textarea');
    area.value = text;
    area.setAttribute('readonly', '');
    area.style.position = 'fixed';
    area.style.opacity = '0';
    document.body.appendChild(area);
    area.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(area);
    return ok;
  } catch {
    return false;
  }
}
