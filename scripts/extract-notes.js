#!/usr/bin/env node
// Extrahiert den Release-Notes-Abschnitt für die aktuelle Version
// und schreibt ihn nach release_body.txt (für GitHub Actions)

import { readFileSync, writeFileSync } from 'fs';

const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const version = pkg.version;

const content = readFileSync('RELEASE_NOTES.md', 'utf8');
const lines = content.split('\n');

const startIdx = lines.findIndex(l => l.startsWith(`## v${version}`));
if (startIdx === -1) {
  writeFileSync('release_body.txt', `Release v${version}`);
  process.exit(0);
}

const endIdx = lines.findIndex((l, i) => i > startIdx && l.startsWith('## '));
const noteLines = endIdx === -1
  ? lines.slice(startIdx + 1)
  : lines.slice(startIdx + 1, endIdx);

writeFileSync('release_body.txt', noteLines.join('\n').trim());
