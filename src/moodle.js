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
 *    core_user_create_users, core_user_get_users_by_field,
 *    enrol_manual_enrol_users, core_group_create_groups,
 *    core_group_get_course_groups, core_group_add_group_members
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
 * @param {object[]} opts.classRows         - Klassen-Zeilen
 * @param {object}   opts.config            - App-Konfiguration (institute, enrolDate, enrolPeriod, …)
 * @param {Function} opts.getClassLabel     - Gibt den Label einer Klassen-Zeile zurück
 * @param {Function} [opts.onProgress]      - Callback für Fortschritts-Meldungen (string)
 * @returns {Promise<{usersCreated, enrolmentsDone, groupsCreated, warnings}>}
 */
export async function enrollInMoodle({
  baseUrl,
  token,
  generatedData,
  activeMatrixCourses,
  classRows,
  config,
  getClassLabel,
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

  // ── Schritt 1a: Bereits bestehende User suchen ────────────────────────────
  report('Bestehende Accounts prüfen…', 10);
  try {
    const existing = await callMoodle(baseUrl, token, 'core_user_get_users_by_field', {
      field: 'username',
      values: generatedData.map(u => u.user),
    });
    console.log('[Moodle] core_user_get_users_by_field (Vorprüfung):', existing);
    if (Array.isArray(existing)) {
      existing.forEach(u => { userIdMap[u.username] = u.id; });
      if (existing.length > 0) warnings.push(`${existing.length} User bereits vorhanden — werden wiederverwendet.`);
    }
  } catch (e) {
    console.warn('[Moodle] Vorprüfung fehlgeschlagen:', e.message);
  }

  // ── Schritt 1b: Nur neue User anlegen ─────────────────────────────────────
  const toCreate = usersPayload.filter(u => !userIdMap[u.username]);
  if (toCreate.length > 0) {
    report(`${toCreate.length} neue Accounts anlegen…`, 20);
    try {
      const created = await callMoodle(baseUrl, token, 'core_user_create_users', { users: toCreate });
      console.log('[Moodle] core_user_create_users response:', created);
      if (Array.isArray(created)) {
        created.forEach(u => { userIdMap[u.username] = u.id; });
        usersCreated = created.length;
      }
    } catch (err) {
      console.error('[Moodle] Bulk-Erstellung fehlgeschlagen:', err.message);
      warnings.push(`Bulk-Erstellung fehlgeschlagen (${err.message}). Versuche einzeln…`);
      for (const user of toCreate) {
        try {
          const created = await callMoodle(baseUrl, token, 'core_user_create_users', { users: [user] });
          console.log(`[Moodle] core_user_create_users (${user.username}):`, created);
          if (Array.isArray(created) && created[0]) {
            userIdMap[created[0].username] = created[0].id;
            usersCreated++;
          }
        } catch (singleErr) {
          console.error(`[Moodle] core_user_create_users (${user.username}) fehlgeschlagen:`, singleErr.message);
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

  // Klassen-Label-Map: classId (number) → vollständiger Gruppenname
  const classLabelById = {};
  classRows.forEach(r => {
    classLabelById[r.id] = `${config.institute?.trim()}-${getClassLabel(r)}`;
  });

  // Welche Gruppen werden benötigt? Pro Kurs eine Gruppe je Klasse, die diesen Kurs hat.
  const groupsNeededMap = new Map(); // key `${courseid}:${name}` → {courseid, name}
  coursesWithIds.forEach(course => {
    classRows.forEach(r => {
      const hasStudents = generatedData.some(
        u =>
          !u.isT &&
          u.cNum === String(r.id).padStart(2, '0') &&
          u.courses.some(uc => extractMoodleCourseId(uc) === course.moodleId)
      );
      if (hasStudents) {
        const groupName = classLabelById[r.id];
        groupsNeededMap.set(`${course.moodleId}:${groupName}`, {
          courseid: course.moodleId,
          name: groupName,
        });
      }
    });
  });

  const groupIdMap = {};   // `${courseid}:${name}` → groupId

  if (groupsNeededMap.size > 0) {
    // Bestehende Gruppen laden (um Duplikate zu vermeiden)
    const uniqueCourseIds = [...new Set(coursesWithIds.map(c => c.moodleId))];
    for (const courseid of uniqueCourseIds) {
      try {
        const existing = await callMoodle(baseUrl, token, 'core_group_get_course_groups', { courseid });
        if (Array.isArray(existing)) {
          existing.forEach(g => {
            const key = `${courseid}:${g.name}`;
            if (groupsNeededMap.has(key)) groupIdMap[key] = g.id;
          });
        }
      } catch { /* non-fatal */ }
    }

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
          created.forEach(g => { groupIdMap[`${g.courseid}:${g.name}`] = g.id; });
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
      // Schüler: in ihre zugewiesenen Kurse mit Rolle 5 (student)
      userData.courses.forEach(course => {
        const moodleId = extractMoodleCourseId(course);
        if (!moodleId) return;
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
    await callMoodle(baseUrl, token, 'enrol_manual_enrol_users', { enrolments });
  }

  // ── Schritt 4: Gruppen zuordnen ────────────────────────────────────────────
  report('Gruppen zuordnen…', 70);

  const groupMembers = [];
  generatedData.forEach(userData => {
    const userId = userIdMap[userData.user];
    if (!userId) return;

    if (userData.isT) {
      // Trainer: zu ALLEN Gruppen aller Kurse hinzufügen
      Object.values(groupIdMap).forEach(groupId => {
        groupMembers.push({ groupid: groupId, userid: userId });
      });
    } else {
      // Schüler: zur Klassen-Gruppe in jedem ihrer Kurse
      userData.courses.forEach(course => {
        const moodleId = extractMoodleCourseId(course);
        if (!moodleId) return;
        const classId = parseInt(userData.cNum, 10);
        const groupName = classLabelById[classId];
        if (!groupName) return;
        const groupId = groupIdMap[`${moodleId}:${groupName}`];
        if (groupId) groupMembers.push({ groupid: groupId, userid: userId });
      });
    }
  });

  if (groupMembers.length > 0) {
    try {
      await callMoodle(baseUrl, token, 'core_group_add_group_members', { members: groupMembers });
    } catch (e) {
      warnings.push(`Gruppen-Zuweisung teilweise fehlgeschlagen: ${e.message}`);
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
    try {
      await callMoodle(baseUrl, token, 'core_cohort_add_cohort_members', {
        members: cohortMembers,
      });
    } catch (e) {
      warnings.push(`Kohorte-Zuweisung fehlgeschlagen: ${e.message}`);
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
  };
}
