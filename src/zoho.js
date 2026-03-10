/**
 * Zoho CRM API v8 – EU Region
 * Alle API-Calls laufen über Rust-Commands (zoho_api_get / zoho_api_post),
 * um CORS und Tauri HTTP-Plugin-Probleme zu umgehen.
 */
import { invoke } from '@tauri-apps/api/core';

/** Hilfsfunktion: Fehler-Objekt aus beliebigem catch-Wert bauen */
function toError(e) {
  if (e instanceof Error) return e;
  return new Error(typeof e === 'string' ? e : JSON.stringify(e));
}

/**
 * Lädt alle Accounts (paginiert) für lokales Filtering.
 * Gibt ein Array von { id, Account_Name } zurück.
 */
export async function getAllZohoAccounts(config) {
  const { zohoClientId, zohoClientSecret, zohoRefreshToken } = config;
  console.log('[Zoho] getAllZohoAccounts (paginiert) via Rust...');
  try {
    const raw = await invoke('zoho_get_all_accounts', {
      clientId: zohoClientId,
      clientSecret: zohoClientSecret,
      refreshToken: zohoRefreshToken,
    });
    const accounts = JSON.parse(raw);
    console.log('[Zoho] Accounts geladen:', accounts.length);
    return accounts;
  } catch (e) {
    throw toError(e);
  }
}

/**
 * Sucht exakt nach einem Account-Namen oder legt ihn neu an.
 * @returns {{ account: {id, Account_Name}, created: boolean }}
 */
export async function findOrCreateZohoAccount(config, name) {
  const { zohoClientId, zohoClientSecret, zohoRefreshToken } = config;
  console.log('[Zoho] findOrCreateZohoAccount – suche:', name);

  // Exakte Suche
  let searchRaw;
  try {
    searchRaw = await invoke('zoho_api_get', {
      clientId: zohoClientId,
      clientSecret: zohoClientSecret,
      refreshToken: zohoRefreshToken,
      path: `Accounts/search?criteria=(Account_Name:equals:${encodeURIComponent(name)})&fields=id,Account_Name`,
    });
  } catch (e) {
    throw toError(e);
  }
  console.log('[Zoho] Account-Suche raw:', searchRaw);
  // Zoho gibt bei keinem Ergebnis einen leeren Body (204) zurück → kein JSON → kein Fehler, einfach weitermachen
  let searchData = {};
  if (searchRaw?.trim()) {
    try { searchData = JSON.parse(searchRaw); } catch { console.warn('[Zoho] Account-Suche: kein gültiges JSON, behandle als kein Ergebnis:', searchRaw); }
  }
  console.log('[Zoho] Account-Suche parsed:', JSON.stringify(searchData, null, 2));

  if (searchData.data?.length) {
    console.log('[Zoho] Account gefunden:', searchData.data[0].id);
    return { account: searchData.data[0], created: false };
  }

  // Neu anlegen
  console.log('[Zoho] Account nicht gefunden – lege neu an:', name);
  const createBody = JSON.stringify({ data: [{ Account_Name: name, Tag: [{ name: 'Institut' }] }] });
  console.log('[Zoho] Account-Create body:', createBody);
  let createRaw;
  try {
    createRaw = await invoke('zoho_api_post', {
      clientId: zohoClientId,
      clientSecret: zohoClientSecret,
      refreshToken: zohoRefreshToken,
      path: 'Accounts',
      body: createBody,
    });
  } catch (e) {
    throw toError(e);
  }
  console.log('[Zoho] Account-Create raw response:', createRaw);
  let createData;
  try { createData = JSON.parse(createRaw); } catch { throw new Error(`Ungültige Zoho-Antwort (Account-Erstellung): ${createRaw}`); }
  console.log('[Zoho] Account-Create parsed:', JSON.stringify(createData, null, 2));
  const entry = createData.data?.[0];
  if (!entry || entry.status === 'error') {
    const details = entry?.details ? JSON.stringify(entry.details) : '';
    throw new Error(`${entry?.message || 'Zoho Account-Erstellung fehlgeschlagen'} | details: ${details} | code: ${entry?.code}`);
  }
  return { account: { id: entry.details.id, Account_Name: name }, created: true };
}

/**
 * Erstellt einen Abschluss (Deal) an einem Zoho Account.
 */
export async function createZohoDeal(config, accountId, dealName, closingDate, description) {
  const { zohoClientId, zohoClientSecret, zohoRefreshToken } = config;

  const payload = {
    data: [{
      Deal_Name: dealName,
      Stage: 'Closed Won',
      Closing_Date: closingDate,
      Account_Name: { id: accountId },
      Description: description,
    }],
  };
  console.log('[Zoho] createZohoDeal accountId:', accountId, 'typeof:', typeof accountId);
  console.log('[Zoho] createZohoDeal payload:', JSON.stringify(payload, null, 2));

  let raw;
  try {
    raw = await invoke('zoho_api_post', {
      clientId: zohoClientId,
      clientSecret: zohoClientSecret,
      refreshToken: zohoRefreshToken,
      path: 'Deals',
      body: JSON.stringify(payload),
    });
  } catch (e) {
    throw toError(e);
  }
  console.log('[Zoho] createZohoDeal raw response:', raw);
  let data;
  try { data = JSON.parse(raw); } catch { throw new Error(`Ungültige Zoho-Antwort (Deal): ${raw}`); }
  console.log('[Zoho] createZohoDeal parsed:', JSON.stringify(data, null, 2));

  // Top-Level-Fehler (z. B. OAUTH_SCOPE_MISMATCH ohne data-Array)
  if (data.status === 'error' || data.code) {
    throw new Error(`${data.message || 'Zoho API-Fehler'} | code: ${data.code}`);
  }

  const entry = data.data?.[0];
  if (!entry || entry.status === 'error') {
    const details = entry?.details ? JSON.stringify(entry.details) : '';
    throw new Error(`${entry?.message || 'Zoho Deal-Erstellung fehlgeschlagen'} | details: ${details} | code: ${entry?.code}`);
  }
  return entry;
}

/**
 * Erstellt eine Notiz an einem Zoho Account.
 */
export async function createZohoNote(config, accountId, title, content) {
  const { zohoClientId, zohoClientSecret, zohoRefreshToken } = config;

  const payload = {
    data: [{
      Note_Title: title,
      Note_Content: content,
      Parent_Id: { id: accountId, module: { api_name: 'Accounts' } },
    }],
  };
  console.log('[Zoho] createZohoNote payload:', JSON.stringify(payload, null, 2));

  let raw;
  try {
    raw = await invoke('zoho_api_post', {
      clientId: zohoClientId,
      clientSecret: zohoClientSecret,
      refreshToken: zohoRefreshToken,
      path: 'Notes',
      body: JSON.stringify(payload),
    });
  } catch (e) {
    throw toError(e);
  }
  console.log('[Zoho] createZohoNote raw response:', raw);
  const data = JSON.parse(raw);
  console.log('[Zoho] createZohoNote parsed:', JSON.stringify(data, null, 2));
  const entry = data.data?.[0];
  if (!entry || entry.status === 'error') {
    const details = entry?.details ? JSON.stringify(entry.details) : '';
    throw new Error(`${entry?.message || 'Zoho Notiz-Erstellung fehlgeschlagen'} | details: ${details} | code: ${entry?.code}`);
  }
  return entry;
}
