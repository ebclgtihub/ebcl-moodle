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
          // Erstellung fehlgeschlagen — User existiert möglicherweise bereits (z.B. wenn Schritt 1a
          // fehlgeschlagen ist und die Vorprüfung den Account nicht gefunden hat).
          // Nochmals per Lookup suchen, damit der Account trotzdem eingeschrieben werden kann.
          try {
            const found = await callMoodle(baseUrl, token, 'core_user_get_users_by_field', { field: 'username', values: [user.username] });
            if (Array.isArray(found) && found[0]) {
              userIdMap[found[0].username] = found[0].id;
            } else {
              warnings.push(`Account ${user.username} konnte weder angelegt noch gefunden werden.`);
            }
          } catch (lookupErr) {
            console.error(`[Moodle] ${user.username}: weder anlegen noch finden:`, lookupErr.message);
            warnings.push(`Account ${user.username} konnte weder angelegt noch gefunden werden.`);
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
          u.cLabel === classLabelById[r.id] &&
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
    let failedEnrolments = [];
    try {
      await callMoodle(baseUrl, token, 'enrol_manual_enrol_users', { enrolments });
    } catch (bulkErr) {
      // Bulk fehlgeschlagen → jeden Eintrag einzeln versuchen, bis zu 3 Versuche
      warnings.push(`Bulk-Einschreibung fehlgeschlagen (${bulkErr.message}) — versuche einzeln…`);
      for (const enrolment of enrolments) {
        let success = false;
        for (let attempt = 1; attempt <= 3 && !success; attempt++) {
          try {
            await callMoodle(baseUrl, token, 'enrol_manual_enrol_users', { enrolments: [enrolment] });
            success = true;
          } catch {
            // nach 3 Fehlversuchen als gescheitert markieren
          }
        }
        if (!success) failedEnrolments.push(enrolment);
      }
    }
    if (failedEnrolments.length > 0) {
      const failedUserIds = [...new Set(failedEnrolments.map(e => e.userid))];
      // Username rückauflösen für lesbare Fehlermeldung
      const failedNames = failedUserIds.map(id => Object.keys(userIdMap).find(u => userIdMap[u] === id) ?? `ID ${id}`);
      throw new Error(
        `Einschreibung fehlgeschlagen für: ${failedNames.join(', ')}. ` +
        `Bitte Moodle-Berechtigungen und Kurs-IDs prüfen.`
      );
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
    try {
      await callMoodle(baseUrl, token, 'core_group_add_group_members', { members: groupMembers });
    } catch (bulkErr) {
      // Bulk fehlgeschlagen → einzeln versuchen
      warnings.push(`Bulk-Gruppenzuweisung fehlgeschlagen (${bulkErr.message}). Versuche einzeln…`);
      for (const member of groupMembers) {
        try {
          await callMoodle(baseUrl, token, 'core_group_add_group_members', { members: [member] });
        } catch (e) {
          warnings.push(`Gruppenzuweisung für User-ID ${member.userid} in Gruppe ${member.groupid} fehlgeschlagen: ${e.message}`);
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
 * @param {number[]} activeCourseIds – Kurse der aktuellen Matrix (für Enrollment-Check)
 */
export async function findMaxNumbers(baseUrl, token, instClean, activeCourseIds = []) {
  let maxStudent = 0;
  let maxTrainer = 0;
  let foundUsers = [];

  // Zwei getrennte Calls — verhindert Überschreitung von PHP max_input_vars (Default: 1000).
  const studentCandidates = [];
  for (let i = 1; i <= 999; i++)
    studentCandidates.push(`${instClean}-student-${String(i).padStart(3, '0')}`);

  const trainerCandidates = [];
  for (let t = 1; t <= 99; t++)
    trainerCandidates.push(`${instClean}-trainer-${t}`);

  const [studentResult, trainerResult] = await Promise.allSettled([
    callMoodle(baseUrl, token, 'core_user_get_users_by_field', { field: 'username', values: studentCandidates }),
    callMoodle(baseUrl, token, 'core_user_get_users_by_field', { field: 'username', values: trainerCandidates }),
  ]);
  if (studentResult.status === 'fulfilled' && Array.isArray(studentResult.value)) foundUsers.push(...studentResult.value);
  if (trainerResult.status === 'fulfilled' && Array.isArray(trainerResult.value)) foundUsers.push(...trainerResult.value);
  if (studentResult.status === 'rejected') console.warn('[Moodle] findMaxNumbers: student lookup failed:', studentResult.reason?.message);
  if (trainerResult.status === 'rejected') console.warn('[Moodle] findMaxNumbers: trainer lookup failed:', trainerResult.reason?.message);

  // Enrollment-IDs aus aktiven Kursen (für maxStudent/maxTrainer — nur eingeschriebene zählen)
  const enrolledUserIds = new Set();
  let enrollmentFetchSucceeded = false;
  for (const courseid of activeCourseIds) {
    try {
      const enrolled = await callMoodle(baseUrl, token, 'core_enrol_get_enrolled_users', { courseid });
      enrollmentFetchSucceeded = true;
      if (Array.isArray(enrolled)) {
        enrolled.forEach(u => { if (u.id) enrolledUserIds.add(u.id); });
      }
    } catch (e) {
      console.warn('[Moodle] findMaxNumbers: enrollment fetch failed:', e.message);
    }
  }

  // Nur eingeschriebene Accounts zählen für maxStudent/maxTrainer.
  // Fallback auf alle foundUsers wenn kein Enrollment-Fetch geklappt hat
  // (z.B. fehlende Berechtigung) — verhindert dass maxStudent fälschlicherweise 0 wird.
  const enrolledUsers = (activeCourseIds.length > 0 && enrollmentFetchSucceeded)
    ? foundUsers.filter(u => u.id && enrolledUserIds.has(u.id))
    : foundUsers;

  enrolledUsers.forEach(u => {
    const sm = u.username?.match(/-student-(\d+)$/i);
    if (sm) maxStudent = Math.max(maxStudent, parseInt(sm[1], 10));
    const tm = u.username?.match(/-trainer-(\d+)$/i);
    if (tm) maxTrainer = Math.max(maxTrainer, parseInt(tm[1], 10));
  });

  // Geister-Accounts nur melden wenn Enrollment-Fetch erfolgreich war.
  // Sonst ist enrolledUserIds leer → alle foundUsers wären fälschlich als Geister markiert.
  const orphanUsernames = (activeCourseIds.length > 0 && enrollmentFetchSucceeded)
    ? foundUsers.filter(u => u.id && !enrolledUserIds.has(u.id)).map(u => u.username)
    : [];

  return { maxStudent, maxTrainer, orphanUsernames };
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

  const enriched = await Promise.all(generatedData.map(async (d) => {
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
  }));
  return enriched;
}

/**
 * Lädt alle Klassen-Gruppen eines Instituts aus den angegebenen Kursen.
 * Gibt [{id, name, courseId, memberCount}] zurück, sortiert nach Name.
 */
export async function fetchInstituteGroups(baseUrl, token, institute, allCourseIds) {
  const prefix = institute.trim() + '-';
  const groupsMap = new Map(); // name → {id, name, courseId, existingCourseIds: Set}

  for (const courseId of allCourseIds) {
    try {
      const groups = await callMoodle(baseUrl, token, 'core_group_get_course_groups', { courseid: courseId });
      if (!Array.isArray(groups)) continue;
      for (const g of groups) {
        if (!g.name.startsWith(prefix)) continue;
        if (!groupsMap.has(g.name)) {
          groupsMap.set(g.name, { id: g.id, name: g.name, courseId, existingCourseIds: new Set([courseId]), memberCount: 0 });
        } else {
          groupsMap.get(g.name).existingCourseIds.add(courseId);
        }
      }
    } catch (e) {
      console.warn(`[Moodle] fetchInstituteGroups: Kurs ${courseId}:`, e.message);
    }
  }

  if (groupsMap.size === 0) return [];

  const groups = [...groupsMap.values()];
  try {
    const memberships = await callMoodle(baseUrl, token, 'core_group_get_group_members', {
      groupids: groups.map(g => g.id),
    });
    if (Array.isArray(memberships)) {
      memberships.forEach(m => {
        const g = groups.find(x => x.id === m.groupid);
        if (g) g.memberCount = (m.userids || []).length;
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
