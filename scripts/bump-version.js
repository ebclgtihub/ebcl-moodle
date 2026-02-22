#!/usr/bin/env node
// Bumpt die Version in package.json, package-lock.json, tauri.conf.json
// und verschiebt den "Unreleased"-Abschnitt in RELEASE_NOTES.md
// Aufruf: node scripts/bump-version.js <patch|minor|major>

import { readFileSync, writeFileSync } from 'fs';

const bumpType = process.argv[2];
if (!['patch', 'minor', 'major'].includes(bumpType)) {
  console.error(`Ungültiger Bump-Typ: "${bumpType}". Erlaubt: patch, minor, major`);
  process.exit(1);
}

// --- Version berechnen ---
const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const oldVersion = pkg.version;
let [major, minor, patch] = oldVersion.split('.').map(Number);

if (bumpType === 'major') { major++; minor = 0; patch = 0; }
else if (bumpType === 'minor') { minor++; patch = 0; }
else { patch++; }

const newVersion = `${major}.${minor}.${patch}`;

// --- package.json ---
pkg.version = newVersion;
writeFileSync('package.json', JSON.stringify(pkg, null, 2) + '\n');

// --- package-lock.json (falls vorhanden) ---
try {
  const lock = JSON.parse(readFileSync('package-lock.json', 'utf8'));
  lock.version = newVersion;
  if (lock.packages?.['']) lock.packages[''].version = newVersion;
  writeFileSync('package-lock.json', JSON.stringify(lock, null, 2) + '\n');
} catch { /* ignorieren wenn nicht vorhanden */ }

// --- src-tauri/tauri.conf.json ---
const tauriConf = JSON.parse(readFileSync('src-tauri/tauri.conf.json', 'utf8'));
tauriConf.version = newVersion;
writeFileSync('src-tauri/tauri.conf.json', JSON.stringify(tauriConf, null, 2) + '\n');

// --- RELEASE_NOTES.md ---
const today = new Date().toISOString().split('T')[0];
const content = readFileSync('RELEASE_NOTES.md', 'utf8');
const lines = content.split('\n');

const unreleasedIdx = lines.findIndex(l => l === '## Unreleased');
if (unreleasedIdx !== -1) {
  const nextSectionIdx = lines.findIndex((l, i) => i > unreleasedIdx && l.startsWith('## '));
  const unreleasedLines = nextSectionIdx === -1
    ? lines.slice(unreleasedIdx + 1)
    : lines.slice(unreleasedIdx + 1, nextSectionIdx);

  const hasContent = unreleasedLines.some(l => l.trim() !== '' && l.trim() !== '---');

  // Immer einen Versionsabschnitt erstellen — auch wenn Unreleased leer war
  const filteredUnreleased = hasContent
    ? unreleasedLines.filter(l => l.trim() !== '---')
    : ['- Kein Changelog-Eintrag.'];

  const newLines = [
    ...lines.slice(0, unreleasedIdx + 1),
    '',
    '---',
    '',
    `## v${newVersion} - ${today}`,
    ...filteredUnreleased,
    ...(nextSectionIdx === -1 ? [] : lines.slice(nextSectionIdx)),
  ];
  writeFileSync('RELEASE_NOTES.md', newLines.join('\n'));
}

console.log(`Version: ${oldVersion} → ${newVersion}`);
