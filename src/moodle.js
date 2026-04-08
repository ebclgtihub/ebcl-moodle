/**
 * Moodle REST API Integration
 *
 * Repliziert exakt die gleiche Einschreibe-Logik wie der CSV-Export:
 *  - Trainer werden in ALLE Kurse eingeschrieben (Rolle 4 = non-editing teacher)
 *    und zu ALLEN Klassen-Gruppen hinzugefügt.
 *  - Schüler werden in ihre zugewiesenen Kurse eingeschrieben (Rolle 5 = student)
 *    und zu ihrer Klassen-Gruppe hinzugefügt.
 *
 * Voraussetzungen auf der Moodle-Seite:
 *  - Web Service mit Token aktiviert
 *  - Token hat Berechtigungen für:
 *    core_user_create_users, core_user_get_users, core_user_get_users_by_field,
 *    enrol_manual_enrol_users, core_group_create_groups,
 *    core_group_get_course_groups, core_group_get_group_members, core_group_add_group_members,
 *    core_cohort_search_cohorts, core_cohort_create_cohorts, core_cohort_add_cohort_members,
 *    core_enrol_get_enrolled_users, core_enrol_get_users_courses
 *  - Kurs-URLs im Format "…/course/view.php?id=<nummer>" ODER
 *    numerische Kurs-IDs im id-Feld des Kurs-Pools
 */

/**
 * Flacht verschachtelte Parameter auf Moodle-Format ab:
 * { users: [{username: 'foo'}] } → { 'users[0][username]': 'foo' }
 */
function flattenParams(obj, prefix = '') {
  const result = {};
  for (const [key, value] of Object.entries(obj)) {
    const newKey = prefix ? `${prefix}[${key}]` : key;
    if (Array.isArray(value)) {
      value.forEach((item, i) => {
        if (typeof item === 'object' && item !== null) {
          Object.assign(result, flattenParams(item, `${newKey}[${i}]`));
        } else {
          result[`${newKey}[${i}]`] = item;
        }
      });
    } else if (typeof value === 'object' && value !== null) {
      Object.assign(result, flattenParams(value, newKey));
    } else if (value !== null && value !== undefined) {
      result[newKey] = value;
    }
  }
  return result;
}

/**
 * Führt einen Moodle REST API-Aufruf durch.
 * Wirft einen Fehler bei HTTP-Fehler oder Moodle-Exception.
 */
async function callMoodle(baseUrl, token, wsfunction, params = {}) {
  const url = `${baseUrl.replace(/\/+$/, '')}/webservice/rest/server.php`;
  const body = new URLSearchParams({
    wstoken: token,
    wsfunction,
    moodlewsrestformat: 'json',
    ...flattenParams(params),
  });

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    console.error(`[Moodle] HTTP ${response.status} für ${wsfunction}:`, text);
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  }

  const data = await response.json();
  console.debug(`[Moodle] ${wsfunction} raw response:`, data);

  if (data && typeof data === 'object' && data.exception) {
    console.error(`[Moodle] Exception bei ${wsfunction}:`, data);
    throw new Error(data.message || `Moodle-Fehler: ${data.exception}`);
  }

  return data;
}

/**
 * Teilt ein Array in Chunks der Größe n auf und verarbeitet alle parallel.
 * Gibt ein flaches Array aller Ergebnisse zurück.
 */
async function chunkedParallel(items, chunkSize, fn) {
  const results = [];
  for (let i = 0; i < items.length; i += chunkSize) {
    const chunk = items.slice(i, i + chunkSize);
    const chunkResults = await Promise.all(chunk.map(fn));
    results.push(...chunkResults);
  }
  return results;
}


/**
 * Ermittelt die numerische Moodle-Kurs-ID aus einem Kurs-Objekt.
 * Priorität: 1) direkte numerische id (aus Power-Automate-Spalte), 2) URL-Parameter ?id=123
 */
function extractMoodleCourseId(course) {
  // Zuerst: direkte numerische id aus dem Kurs-Pool (Power Automate ID-Spalte)
  const numId = parseInt(String(course?.id ?? ''), 10);
  if (!isNaN(numId) && numId > 0) return numId;
  // Fallback: aus Kurs-URL extrahieren (?id=123)
  if (course?.url) {
    const match = String(course.url).match(/[?&]id=(\d+)/);
    if (match) return parseInt(match[1], 10);
  }
  return null;
}

/**
 * Führt die vollständige Moodle-Einschreibung durch:
 * 1. Alle User anlegen (oder bestehende finden)
 * 2. Gruppen vorbereiten (bestehende laden + fehlende anlegen)
 * 3. Alle User in ihre Kurse einschreiben
 * 4. Alle User ihren Gruppen zuordnen
 *
 * @param {object} opts
 * @param {string}   opts.baseUrl           - Moodle-Basis-URL (z.B. https://moodle.schule.at)
 * @param {string}   opts.token             - Moodle Web Service Token
 * @param {object[]} opts.generatedData     - Generierte Account-Daten (aus generateList)
 * @param {object[]} opts.activeMatrixCourses - Aktive Kurse aus der Matrix
 * @param {object}   opts.config            - App-Konfiguration (institute, enrolDate, enrolPeriod, …)
 * @param {Function} [opts.onProgress]      - Callback für Fortschritts-Meldungen (string)
 * @returns {Promise<{usersCreated, enrolmentsDone, groupsCreated, warnings}>}
 */
export async function enrollInMoodle({
  baseUrl,
  token,
  generatedData,
  activeMatrixCourses,
  config,
  onProgress,
}) {
  const report = (label, pct) => onProgress?.(label, pct);

  // ── Eingaben prüfen ────────────────────────────────────────────────────────
  if (!baseUrl?.trim()) throw new Error('Moodle-URL ist nicht konfiguriert (Einstellungen → Backend).');
  if (!token?.trim()) throw new Error('Moodle API-Token ist nicht konfiguriert (Einstellungen → Backend).');
  if (!generatedData?.length) throw new Error('Keine Accountdaten vorhanden — bitte zuerst Liste generieren.');

  // Kurs-IDs prüfen
  const coursesWithIds = activeMatrixCourses.map(c => ({
    ...c,
    moodleId: extractMoodleCourseId(c),
  }));
  const missingCourseIds = coursesWithIds.filter(c => !c.moodleId);
  if (missingCourseIds.length > 0) {
    throw new Error(
      `Moodle-Kurs-ID nicht ermittelbar für: ${missingCourseIds.map(c => c.label).join(', ')}.\n` +
      'Bitte prüfen ob die Kurs-URLs das Format "…?id=123" haben oder die Kurs-IDs numerisch sind.'
    );
  }

  const enrolDateTs = Math.floor(new Date(config.enrolDate).getTime() / 1000);
  const enrolEndTs = enrolDateTs + parseInt(config.enrolPeriod, 10) * 86400;

  // ── Schritt 1: User anlegen ────────────────────────────────────────────────
  report(`${generatedData.length} Accounts vorbereiten…`, 5);

  const usersPayload = generatedData.map(u => ({
    username: u.user?.trim().toLowerCase(),
    password: u.pw?.trim(),
    firstname: u.first?.trim(),
    lastname: u.last?.trim(),
    email: u.mail?.trim().toLowerCase(),
    auth: 'manual',
  }));

  const userIdMap = {};   // username → moodle user id
  const warnings = [];
  let usersCreated = 0;

  // ── Schritt 1a: Bestehende Institut-User laden (Kohorte via E-Mail-Wildcard) ───
  // core_user_get_users mit email-Wildcard findet alle Institut-User in 1 Call.
  // Hinweis: username-Feld unterstützt kein %-Wildcard (Moodle Docs bestätigt),
  //          email/firstname/lastname dagegen schon.
  // Fallback auf core_user_get_users_by_field falls der Token die Funktion nicht hat.
  report('Bestehende Accounts prüfen…', 10);
  const instClean = (config.institute || '').replace(/\s+/g, '').toLowerCase();
  {
    let wildcardSucceeded = false;
    try {
      const res = await callMoodle(baseUrl, token, 'core_user_get_users', {
        criteria: [{ key: 'email', value: `%@${instClean}.com` }],
      });
      const users = res?.users ?? (Array.isArray(res) ? res : []);
      users.forEach(u => { if (u.username && u.id) userIdMap[u.username] = u.id; });
      wildcardSucceeded = true;
    } catch (e) { console.warn('[Moodle] Wildcard-Vorprüfung fehlgeschlagen:', e.message); }

    // Fallback: core_user_get_users_by_field mit exakten Usernamen aus generatedData
    if (!wildcardSucceeded) {
      const usernames = generatedData.map(u => u.user?.trim().toLowerCase()).filter(Boolean);
      const chunkSize = 200;
      for (let i = 0; i < usernames.length; i += chunkSize) {
        try {
          const res = await callMoodle(baseUrl, token, 'core_user_get_users_by_field', {
            field: 'username', values: usernames.slice(i, i + chunkSize),
          });
          if (Array.isArray(res)) res.forEach(u => { if (u.username && u.id) userIdMap[u.username] = u.id; });
        } catch (e) { console.warn('[Moodle] Vorprüfung chunk fehlgeschlagen:', e.message); }
      }
    }

    const existingCount = Object.keys(userIdMap).length;
    if (existingCount > 0) warnings.push(`${existingCount} User bereits vorhanden — werden wiederverwendet.`);
  }

  // ── Schritt 1b: Nur neue User anlegen ─────────────────────────────────────
  const toCreate = usersPayload.filter(u => !userIdMap[u.username]);
  if (toCreate.length > 0) {
    report(`${toCreate.length} neue Accounts anlegen…`, 20);
    const chunkSize = 50; // 50 User × ~5 Felder = 250 Parameter pro Request
    for (let i = 0; i < toCreate.length; i += chunkSize) {
      const chunk = toCreate.slice(i, i + chunkSize);
      try {
        const created = await callMoodle(baseUrl, token, 'core_user_create_users', { users: chunk });
        if (Array.isArray(created)) {
          created.forEach(u => { userIdMap[u.username] = u.id; });
          usersCreated += created.length;
        }
      } catch (err) {
        // Chunk fehlgeschlagen → jeden User einzeln versuchen
        for (const user of chunk) {
          try {
            const created = await callMoodle(baseUrl, token, 'core_user_create_users', { users: [user] });
            if (Array.isArray(created) && created[0]) { userIdMap[created[0].username] = created[0].id; usersCreated++; }
          } catch {
            // User existiert evtl. bereits — per Lookup nachholen
            try {
              const found = await callMoodle(baseUrl, token, 'core_user_get_users', { criteria: [{ key: 'username', value: user.username }] });
              const match = (found?.users ?? found ?? [])[0];
              if (match) { userIdMap[match.username] = match.id; }
              else warnings.push(`Account ${user.username} konnte weder angelegt noch gefunden werden.`);
            } catch { warnings.push(`Account ${user.username} konnte weder angelegt noch gefunden werden.`); }
          }
        }
      }
    }
  }

  const resolvedCount = Object.keys(userIdMap).length;
  if (resolvedCount === 0) {
    throw new Error('Kein Moodle-Account konnte angelegt werden. Bitte Moodle-URL und Token in den Einstellungen prüfen.');
  }

  // ── Schritt 2: Gruppen vorbereiten ─────────────────────────────────────────
  report('Klassen-Gruppen vorbereiten…', 40);

  // Welche Gruppen werden benötigt? Direkt aus generatedData[].cLabel ableiten —
  // nicht über classLabelById, da cLabel bei Gruppen-Wiederverwendung vom berechneten
  // Namen abweichen kann (z.B. "Institute-1A" statt "Institute-3A" mit Offset).
  const groupsNeededMap = new Map(); // key `${courseid}:${name}` → {courseid, name}
  const activeMoodleIdSet = new Set(coursesWithIds.map(c => c.moodleId));
  generatedData.forEach(u => {
    if (u.isT || !u.cLabel) return;
    u.courses.forEach(course => {
      const moodleId = extractMoodleCourseId(course);
      if (!moodleId || !activeMoodleIdSet.has(moodleId)) return;
      const key = `${moodleId}:${u.cLabel}`;
      if (!groupsNeededMap.has(key)) groupsNeededMap.set(key, { courseid: moodleId, name: u.cLabel });
    });
  });

  const groupIdMap = {};          // `${courseid}:${name}` → groupId (aktuelle Session)
  const allExistingGroupIds = []; // alle Gruppen-IDs aller Kurse (für Trainer im Neu-Anlegen-Modus)

  const uniqueCourseIds = [...new Set(coursesWithIds.map(c => c.moodleId))];

  // Bestehende Gruppen laden — immer (für Duplikat-Check + Neu-Anlegen-Trainer)
  for (const courseid of uniqueCourseIds) {
    try {
      const existing = await callMoodle(baseUrl, token, 'core_group_get_course_groups', { courseid });
      if (Array.isArray(existing)) {
        const institutePrefix = (config.institute?.trim() || '').toLowerCase();
        existing.forEach(g => {
          const key = `${courseid}:${g.name}`;
          if (groupsNeededMap.has(key)) groupIdMap[key] = g.id;
          // Institutsgruppen immer sammeln — Trainer brauchen sie in beiden Modi
          if (g.name?.toLowerCase().startsWith(institutePrefix))
            allExistingGroupIds.push(g.id);
        });
      }
    } catch { /* non-fatal */ }
  }

  if (groupsNeededMap.size > 0) {
    // Fehlende Gruppen anlegen
    const toCreate = Array.from(groupsNeededMap.values()).filter(
      g => !groupIdMap[`${g.courseid}:${g.name}`]
    );
    if (toCreate.length > 0) {
      try {
        const created = await callMoodle(baseUrl, token, 'core_group_create_groups', {
          groups: toCreate.map(g => ({ courseid: g.courseid, name: g.name, description: '' })),
        });
        if (Array.isArray(created)) {
          created.forEach(g => {
            groupIdMap[`${g.courseid}:${g.name}`] = g.id;
            allExistingGroupIds.push(g.id); // immer — Trainer brauchen neue Gruppen in beiden Modi
          });
        }
      } catch (e) {
        warnings.push(`Gruppen konnten nicht alle angelegt werden: ${e.message}`);
      }
    }
  }

  // ── Schritt 3: Einschreibungen ─────────────────────────────────────────────
  report('Einschreibungen durchführen…', 55);

  const enrolments = [];
  generatedData.forEach(userData => {
    const userId = userIdMap[userData.user];
    if (!userId) return;

    if (userData.isT) {
      // Trainer: in ALLE aktiven Kurse mit Rolle 4 (non-editing teacher)
      coursesWithIds.forEach(course => {
        enrolments.push({
          roleid: 4,
          userid: userId,
          courseid: course.moodleId,
          timestart: enrolDateTs,
          timeend: enrolEndTs,
        });
      });
    } else {
      // Schüler: nur in Kurse einschreiben die in activeMatrixCourses sind.
      // Kurs-IDs aus Anreicherung (fetchFullEnrollments) werden ignoriert —
      // sonst würden historische Kurse mit neuen Zeiträumen überschrieben.
      const activeMoodleIds = new Set(coursesWithIds.map(c => c.moodleId));
      userData.courses.forEach(course => {
        const moodleId = extractMoodleCourseId(course);
        if (!moodleId || !activeMoodleIds.has(moodleId)) return;
        enrolments.push({
          roleid: 5,
          userid: userId,
          courseid: moodleId,
          timestart: enrolDateTs,
          timeend: enrolEndTs,
        });
      });
    }
  });

  if (enrolments.length > 0) {
    // 100 Einschreibungen × 5 Felder = 500 Parameter — sicher unter max_input_vars
    const chunkSize = 100;
    const failedEnrolments = [];
    for (let i = 0; i < enrolments.length; i += chunkSize) {
      const chunk = enrolments.slice(i, i + chunkSize);
      try {
        await callMoodle(baseUrl, token, 'enrol_manual_enrol_users', { enrolments: chunk });
      } catch (bulkErr) {
        warnings.push(`Einschreibungs-Chunk fehlgeschlagen (${bulkErr.message}) — versuche einzeln…`);
        for (const enrolment of chunk) {
          let success = false;
          for (let attempt = 1; attempt <= 3 && !success; attempt++) {
            try { await callMoodle(baseUrl, token, 'enrol_manual_enrol_users', { enrolments: [enrolment] }); success = true; } catch { /* retry */ }
          }
          if (!success) failedEnrolments.push(enrolment);
        }
      }
    }
    if (failedEnrolments.length > 0) {
      const failedUserIds = [...new Set(failedEnrolments.map(e => e.userid))];
      const failedNames = failedUserIds.map(id => Object.keys(userIdMap).find(u => userIdMap[u] === id) ?? `ID ${id}`);
      throw new Error(`Einschreibung fehlgeschlagen für: ${failedNames.join(', ')}. Bitte Moodle-Berechtigungen und Kurs-IDs prüfen.`);
    }
  }

  // ── Schritt 4: Gruppen zuordnen ────────────────────────────────────────────
  report('Gruppen zuordnen…', 70);

  const groupMembers = [];
  generatedData.forEach(userData => {
    const userId = userIdMap[userData.user];
    if (!userId) return;

    if (userData.isT) {
      // Trainer: zu ALLEN Institutsgruppen aller Kurse hinzufügen (beide Modi)
      const trainerGroupIds = [...new Set(allExistingGroupIds)];
      trainerGroupIds.forEach(groupId => {
        groupMembers.push({ groupid: groupId, userid: userId });
      });
    } else {
      // Schüler: zur Klassen-Gruppe in jedem ihrer Kurse
      userData.courses.forEach(course => {
        const moodleId = extractMoodleCourseId(course);
        if (!moodleId) return;
        const groupName = userData.cLabel;
        if (!groupName) return;
        const groupId = groupIdMap[`${moodleId}:${groupName}`];
        if (groupId) groupMembers.push({ groupid: groupId, userid: userId });
      });
    }
  });

  if (groupMembers.length > 0) {
    const chunkSize = 200; // 200 × 2 Felder = 400 Parameter
    for (let i = 0; i < groupMembers.length; i += chunkSize) {
      const chunk = groupMembers.slice(i, i + chunkSize);
      try {
        await callMoodle(baseUrl, token, 'core_group_add_group_members', { members: chunk });
      } catch (bulkErr) {
        warnings.push(`Gruppen-Chunk fehlgeschlagen (${bulkErr.message}). Versuche einzeln…`);
        for (const member of chunk) {
          try {
            await callMoodle(baseUrl, token, 'core_group_add_group_members', { members: [member] });
          } catch (e) {
            warnings.push(`Gruppenzuweisung für User-ID ${member.userid} in Gruppe ${member.groupid} fehlgeschlagen: ${e.message}`);
          }
        }
      }
    }
  }

  // ── Schritt 5: Kohorte ────────────────────────────────────────────────────
  // Entspricht cohort1 = config.institute im CSV-Export
  report('Kohorte einrichten…', 82);

  let cohortId = null;
  let cohortCreated = false;
  const cohortName = config.institute?.trim() ?? '';
  // idnumber: keine Leerzeichen, keine Sonderzeichen
  const cohortIdNumber = cohortName.replace(/\s+/g, '-').replace(/[^\w-]/g, '');

  // Erst suchen — nur anlegen wenn nicht vorhanden
  try {
    const found = await callMoodle(baseUrl, token, 'core_cohort_search_cohorts', {
      query: cohortName,
      context: { contextlevel: 'system', instanceid: 0 },
      includes: 'all',
      limitfrom: 0,
      limitnum: 50,
    });
    const cohorts = found?.cohorts ?? found;
    if (Array.isArray(cohorts)) {
      const match = cohorts.find(
        c => c.name === cohortName || c.idnumber === cohortIdNumber
      );
      if (match) cohortId = match.id;
    }
  } catch (e) {
    warnings.push(`Kohorte-Suche fehlgeschlagen: ${e.message}`);
  }

  if (!cohortId) {
    try {
      const created = await callMoodle(baseUrl, token, 'core_cohort_create_cohorts', {
        cohorts: [{
          categorytype: { type: 'system', value: '' },
          name: cohortName,
          idnumber: cohortIdNumber,
          description: '',
        }],
      });
      if (Array.isArray(created) && created[0]?.id) {
        cohortId = created[0].id;
        cohortCreated = true;
      }
    } catch (e) {
      warnings.push(`Kohorte konnte nicht angelegt werden: ${e.message}`);
    }
  }

  if (cohortId) {
    const cohortMembers = Object.entries(userIdMap).map(([, userid]) => ({
      cohorttype: { type: 'id', value: String(cohortId) },
      usertype: { type: 'id', value: String(userid) },
    }));
    const chunkSize = 100; // 100 × 4 Felder (nested) = sicher
    for (let i = 0; i < cohortMembers.length; i += chunkSize) {
      const chunk = cohortMembers.slice(i, i + chunkSize);
      try {
        await callMoodle(baseUrl, token, 'core_cohort_add_cohort_members', { members: chunk });
      } catch (e) {
        warnings.push(`Kohorte-Zuweisung fehlgeschlagen: ${e.message}`);
      }
    }
  } else {
    warnings.push('Kohorte konnte nicht angelegt oder gefunden werden — Kohorte-Zuweisung übersprungen.');
  }

  return {
    usersCreated,
    usersResolved: resolvedCount,
    enrolmentsDone: enrolments.length,
    groupsCreated: Object.keys(groupIdMap).length,
    cohortName,
    cohortId,
    cohortCreated,
    cohortMembersAdded: cohortId ? Object.keys(userIdMap).length : 0,
    warnings,
    userIdMap,
  };
}

/**
 * Ermittelt die höchsten vorhandenen Student/Trainer/Klassen-Nummern
 * für ein Institut, damit bei "Neu anlegen" fortlaufend nummeriert werden kann.
 *
 * Gibt zusätzlich `orphanUsernames` zurück: Accounts die in Moodle existieren,
 * aber in keinem der aktiven Kurse eingeschrieben sind ("Geister-Accounts").
 * Diese können entstehen wenn ein früherer Einschreibungsversuch fehlgeschlagen ist.
 */

/**
 * Gibt alle Moodle-User des Instituts zurück (via Email-Wildcard *@instClean.com).
 * Nützlich um bestehende Trainer zu laden wenn keine Gruppenabfrage möglich ist.
 */
export async function fetchInstituteUsers(baseUrl, token, instClean) {
  try {
    const res = await callMoodle(baseUrl, token, 'core_user_get_users', {
      criteria: [{ key: 'email', value: `%@${instClean}.com` }],
    });
    return res?.users ?? [];
  } catch (e) {
    console.warn('[Moodle] fetchInstituteUsers fehlgeschlagen:', e.message);
    return [];
  }
}

/**
 * @param {number[]} activeCourseIds – Fallback: Kurse für Enrollment-Abfrage wenn Wildcard leer
 */
export async function findMaxNumbers(baseUrl, token, instClean, activeCourseIds = []) {
  let maxStudent = 0;
  let maxTrainer = 0;

  const applyUsers = users => users.forEach(u => {
    const sm = u.username?.match(new RegExp(`^${instClean}-student-(\\d+)$`, 'i'));
    if (sm) maxStudent = Math.max(maxStudent, parseInt(sm[1], 10));
    const tm = u.username?.match(new RegExp(`^${instClean}-trainer-(\\d+)$`, 'i'));
    if (tm) maxTrainer = Math.max(maxTrainer, parseInt(tm[1], 10));
  });

  // Email-Wildcard — username unterstützt kein %, email schon (Moodle-Doku).
  // Leere Antwort (0 User) = erstes Institut → maxStudent/maxTrainer = 0 → bei 1 starten. Korrekt.
  // Nur bei Exception (Token-Fehler, Netzwerk) auf Enrollment-Daten zurückfallen.
  let wildcardFailed = false;
  try {
    const res = await callMoodle(baseUrl, token, 'core_user_get_users', {
      criteria: [{ key: 'email', value: `%@${instClean}.com` }],
    });
    applyUsers(res?.users ?? []);
  } catch (e) {
    wildcardFailed = true;
    console.warn('[Moodle] findMaxNumbers: Email-Wildcard fehlgeschlagen:', e.message);
  }

  // Fallback: nur bei echtem API-Fehler — Enrollment-Daten als Notlösung
  if (wildcardFailed && activeCourseIds.length > 0) {
    console.warn('[Moodle] findMaxNumbers: Fallback auf Enrollment-Daten');
    const enrolledById = new Map();
    for (const courseid of activeCourseIds) {
      try {
        const enrolled = await callMoodle(baseUrl, token, 'core_enrol_get_enrolled_users', { courseid });
        if (Array.isArray(enrolled)) enrolled.forEach(u => { if (u.id && u.username) enrolledById.set(u.id, u); });
      } catch (e) { console.warn('[Moodle] findMaxNumbers fallback enrollment failed:', e.message); }
    }
    applyUsers([...enrolledById.values()]);
  }

  return { maxStudent, maxTrainer, orphanUsernames: [] };
}

/**
 * Reichert generatedData mit den tatsächlichen Kurseinschreibungen aus Moodle an.
 * Nützlich nach einem Aktualisieren-Lauf: bestehende User haben evtl. mehr Kurse
 * als in der aktuellen Generation konfiguriert wurde.
 *
 * @param {string}   baseUrl          - Moodle-Basis-URL
 * @param {string}   token            - Moodle Web Service Token
 * @param {object[]} generatedData    - Account-Daten aus generateList
 * @param {object}   userIdMap        - username → Moodle-User-ID (aus enrollInMoodle)
 * @param {object[]} courseDictionary - Alle bekannten Kurse der App (mit id, label, shorthand, url)
 * @returns {Promise<object[]>}       - Angereichertes generatedData
 */
export async function fetchFullEnrollments(baseUrl, token, generatedData, userIdMap, courseDictionary) {
  const courseById = {};
  courseDictionary.forEach(c => { courseById[String(c.id)] = c; });

  // 10 parallele Calls — unbegrenztes Promise.all kann Moodle bei >100 Usern überlasten
  const enriched = await chunkedParallel(generatedData, 10, async (d) => {
    const userId = userIdMap?.[d.user?.trim().toLowerCase()];
    if (!userId) return d;
    try {
      const moodleCourses = await callMoodle(baseUrl, token, 'core_enrol_get_users_courses', { userid: userId });
      if (Array.isArray(moodleCourses)) {
        if (moodleCourses.length === 0) {
          // User existiert in Moodle aber ist in keinem Kurs → Einschreibung fehlgeschlagen
          return { ...d, courses: [] };
        }
        const fullCourses = moodleCourses.map(c => courseById[String(c.id)]).filter(Boolean);
        if (fullCourses.length > 0) return { ...d, courses: fullCourses };
        // Kurse vorhanden aber keiner im App-Pool → Session-Kurse behalten
      }
    } catch (e) {
      // Fehlende Berechtigung → sofort re-throwen damit äußerer catch den Toast zeigt.
      // Einzelne Netzwerkfehler still überspringen.
      if (e.message === 'Access control exception') throw e;
      console.warn('[Moodle] fetchFullEnrollments: failed for', d.user, e.message);
    }
    return d;
  });
  return enriched;
}

/**
 * Lädt alle Klassen-Gruppen eines Instituts aus den angegebenen Kursen.
 * Gibt [{id, name, courseId, memberCount}] zurück, sortiert nach Name.
 */
export async function fetchInstituteGroups(baseUrl, token, institute, allCourseIds) {
  const prefix = institute.trim() + '-';
  const groupsMap = new Map(); // name → {id, name, courseId, existingCourseIds: Set}

  // Alle Kurse parallel abfragen statt sequenziell
  const results = await Promise.all(allCourseIds.map(async courseId => {
    try {
      const groups = await callMoodle(baseUrl, token, 'core_group_get_course_groups', { courseid: courseId });
      return { courseId, groups: Array.isArray(groups) ? groups : [] };
    } catch (e) {
      console.warn(`[Moodle] fetchInstituteGroups: Kurs ${courseId}:`, e.message);
      return { courseId, groups: [] };
    }
  }));

  for (const { courseId, groups } of results) {
    for (const g of groups) {
      if (!g.name.startsWith(prefix)) continue;
      if (!groupsMap.has(g.name)) {
        groupsMap.set(g.name, { id: g.id, name: g.name, courseId, existingCourseIds: new Set([courseId]), memberCount: 0 });
      } else {
        groupsMap.get(g.name).existingCourseIds.add(courseId);
      }
    }
  }

  if (groupsMap.size === 0) return [];

  const groups = [...groupsMap.values()];
  const instClean = institute.replace(/\s+/g, '').toLowerCase();

  // Gruppengrößen ermitteln — Trainer per Wildcard einmalig holen statt alle Enrollments pro Kurs
  try {
    const [memberships, trainerRes] = await Promise.all([
      callMoodle(baseUrl, token, 'core_group_get_group_members', { groupids: groups.map(g => g.id) }),
      callMoodle(baseUrl, token, 'core_user_get_users', {
        // email unterstützt %-Wildcard, username nicht (Moodle Docs)
        criteria: [{ key: 'email', value: `trainer%@${instClean}.com` }],
      }),
    ]);

    const trainerIds = new Set(
      (trainerRes?.users ?? trainerRes ?? []).map(u => u.id).filter(Boolean)
    );

    if (Array.isArray(memberships)) {
      memberships.forEach(m => {
        const g = groups.find(x => x.id === m.groupid);
        if (!g) return;
        g.memberCount = (m.userids || []).filter(id => !trainerIds.has(id)).length;
      });
    }
  } catch (e) {
    console.warn('[Moodle] fetchInstituteGroups: Mitgliederanzahl fehlgeschlagen:', e.message);
  }

  // Set → Array für Serialisierung
  return groups
    .map(g => ({ ...g, existingCourseIds: [...g.existingCourseIds] }))
    .sort((a, b) => a.name.localeCompare(b.name, 'de'));
}

/**
 * Lädt alle eingeschriebenen Nutzer einer Gruppe aus einem Kurs.
 * Gibt User-Objekte zurück (mit username, firstname, lastname, email).
 */
export async function fetchGroupMembers(baseUrl, token, courseId, groupId) {
  try {
    const users = await callMoodle(baseUrl, token, 'core_enrol_get_enrolled_users', {
      courseid: courseId,
      options: [{ name: 'groupid', value: groupId }],
    });
    return Array.isArray(users) ? users : [];
  } catch (e) {
    console.warn(`[Moodle] fetchGroupMembers: Gruppe ${groupId}:`, e.message);
    return [];
  }
}

/**
 * Lädt alle Kursbereiche und Kurse direkt von der Moodle-Instanz.
 * Gibt ein Array im gleichen Format wie der Power-Automate-Pool zurück:
 * { id, label, shorthand, url, tag }
 */
export async function fetchMoodleCourses(baseUrl, token) {
  const [categories, courses] = await Promise.all([
    callMoodle(baseUrl, token, 'core_course_get_categories', {}),
    callMoodle(baseUrl, token, 'core_course_get_courses', {}),
  ]);

  const catMap = {};
  if (Array.isArray(categories)) categories.forEach(c => { catMap[c.id] = c.name; });

  if (!Array.isArray(courses)) throw new Error('Keine Kurse von Moodle erhalten.');

  const base = baseUrl.replace(/\/+$/, '');
  return courses
    .filter(c => c.id > 1) // ID 1 = Site Home überspringen
    .map(c => {
      const sh = (c.shortname || '').substring(0, 10) || String(c.fullname || '').replace(/[^A-Z0-9]/g, '').substring(0, 4) || String(c.id);
      return {
        id: String(c.id),
        label: c.fullname || c.shortname || `Kurs ${c.id}`,
        shorthand: sh,
        url: `${base}/course/view.php?id=${c.id}`,
        tag: catMap[c.categoryid] || '',
      };
    });
}
