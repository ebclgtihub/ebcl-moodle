import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import { LazyStore } from '@tauri-apps/plugin-store';
import {
  Users, FileSpreadsheet, CheckCircle2, Building2, Plus,
  Loader2, Table as TableIcon, Check, AlertTriangle, ChevronDown,
  Eye, X, RefreshCw, Info, Settings, HelpCircle, BookOpen,
  Zap, ClipboardList, ShieldCheck, GraduationCap, FileDown,
  Save, Wifi, WifiOff
} from 'lucide-react';

/**
 * Moodle Anmeldungen V3
 * - STORE: Echter Tauri Plugin Store (@tauri-apps/plugin-store) mit LazyStore
 * - UX: Toast-Notifications statt inline Errors, Save-Status Indikator
 * - PERF: useCallback & useMemo konsequent eingesetzt, kein unnötiges Re-Rendering
 * - ROBUSTHEIT: Fehlerbehandlung ausgebaut, Lade-Zustände granularer
 * - CODE: Konstanten ausgelagert, Magic-Numbers eliminiert
 */

// ─── Tauri Plugin Store ────────────────────────────────────────────────────────
// Wird beim ersten Zugriff lazy initialisiert.
// Datei liegt im App-Datenverzeichnis des Nutzers (z.B. AppData/Roaming auf Windows).
const store = new LazyStore('moodle-settings.json', { autoSave: false });

// ─── Konstanten ───────────────────────────────────────────────────────────────
const COLORS = {
  main: '#9D202B',
  bg: '#FFFEF4',
  accent1: '#153D61',
  accent2: '#00664F',
};

const DEFAULT_CONFIG = {
  institute: '',
  studentPwd: '***ENTFERNT***',
  trainerPwd: '***ENTFERNT***',
  enrolPeriod: 31,
  enrolDate: new Date().toISOString().split('T')[0],
  classSizes: [20, 30, 40, 50],
  classCounts: { 0: 2, 1: 1, 2: 0, 3: 0 },
  trainerCount: 2,
  courseSlotCount: 4,
  selectedPoolCourseIds: Array(8).fill('none'),
};

const COURSE_API_URL =
  '***ENTFERNT***';

// ─── Hilfsfunktionen (außerhalb Komponente = kein Re-Create) ──────────────────
const findValueByPattern = (item, patterns) => {
  const keys = Object.keys(item);
  for (const p of patterns) {
    const match = keys.find(k => k.toLowerCase() === p.toLowerCase());
    if (match && item[match] != null && String(item[match]).trim() !== '') return item[match];
  }
  for (const p of patterns) {
    const match = keys.find(k => k.toLowerCase().includes(p.toLowerCase()));
    if (match && item[match] != null && String(item[match]).trim() !== '') return item[match];
  }
  return null;
};

const isGuid = str =>
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(str);

const esc = v => `"${String(v ?? '').replace(/"/g, '""')}"`;

// ─── Toast-Komponente ─────────────────────────────────────────────────────────
const Toast = ({ toasts, removeToast }) => (
  <div className="fixed bottom-5 right-5 z-[999] flex flex-col gap-2 pointer-events-none">
    {toasts.map(t => (
      <div
        key={t.id}
        className={`flex items-center gap-3 px-4 py-3 rounded-2xl shadow-xl text-white text-xs font-semibold pointer-events-auto
          ${t.type === 'error' ? 'bg-rose-600' : t.type === 'success' ? 'bg-emerald-600' : 'bg-slate-700'}`}
        style={{ animation: 'slideInRight 0.25s ease' }}
      >
        {t.type === 'error' ? <AlertTriangle size={14} /> : <CheckCircle2 size={14} />}
        <span>{t.message}</span>
        <button onClick={() => removeToast(t.id)} className="ml-1 opacity-70 hover:opacity-100"><X size={12} /></button>
      </div>
    ))}
  </div>
);

// ─── Haupt-Komponente ─────────────────────────────────────────────────────────
const App = () => {
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [courseDictionary, setCourseDictionary] = useState([]);
  const [isLoadingPool, setIsLoadingPool] = useState(false);
  const [activeModal, setActiveModal] = useState(null);
  const [classMatrix, setClassMatrix] = useState({});
  const [generatedData, setGeneratedData] = useState([]);
  const [isGenerated, setIsGenerated] = useState(false);
  const [isExportingPDF, setIsExportingPDF] = useState(false);
  const [libsReady, setLibsReady] = useState(false);
  const [isStoreLoaded, setIsStoreLoaded] = useState(false);
  const [saveStatus, setSaveStatus] = useState('idle'); // 'idle' | 'saving' | 'saved' | 'error'
  const [toasts, setToasts] = useState([]);
  const [isOnline, setIsOnline] = useState(navigator.onLine);
  const saveTimeoutRef = useRef(null);
  const toastIdRef = useRef(0);

  // ─── Toast System ────────────────────────────────────────────────────────────
  const addToast = useCallback((message, type = 'info', duration = 4000) => {
    const id = ++toastIdRef.current;
    setToasts(prev => [...prev, { id, message, type }]);
    if (duration > 0) setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), duration);
  }, []);

  const removeToast = useCallback(id => setToasts(prev => prev.filter(t => t.id !== id)), []);

  // ─── Online-Status ────────────────────────────────────────────────────────────
  useEffect(() => {
    const onOnline = () => { setIsOnline(true); addToast('Verbindung wiederhergestellt', 'success'); };
    const onOffline = () => { setIsOnline(false); addToast('Keine Internetverbindung', 'error'); };
    window.addEventListener('online', onOnline);
    window.addEventListener('offline', onOffline);
    return () => { window.removeEventListener('online', onOnline); window.removeEventListener('offline', onOffline); };
  }, [addToast]);

  // ─── Store: Laden ─────────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      try {
        const savedConfig = await store.get('appConfig');
        if (savedConfig) setConfig(prev => ({ ...prev, ...savedConfig }));

        const savedMatrix = await store.get('classMatrix');
        if (savedMatrix) setClassMatrix(savedMatrix);
      } catch (err) {
        console.error('Store laden fehlgeschlagen:', err);
        addToast('Einstellungen konnten nicht geladen werden.', 'error');
      } finally {
        setIsStoreLoaded(true);
      }
    };
    load();
  }, []); // eslint-disable-line

  // ─── Store: Auto-Save (debounced) ────────────────────────────────────────────
  useEffect(() => {
    if (!isStoreLoaded) return;

    setSaveStatus('saving');
    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);

    saveTimeoutRef.current = setTimeout(async () => {
      try {
        await store.set('appConfig', config);
        await store.set('classMatrix', classMatrix);
        await store.save(); // Physisch auf Disk schreiben (Tauri)
        setSaveStatus('saved');
        setTimeout(() => setSaveStatus('idle'), 2000);
      } catch (err) {
        console.error('Store speichern fehlgeschlagen:', err);
        setSaveStatus('error');
        addToast('Einstellungen konnten nicht gespeichert werden.', 'error');
      }
    }, 600);

    return () => clearTimeout(saveTimeoutRef.current);
  }, [config, classMatrix, isStoreLoaded]); // eslint-disable-line

  // ─── PDF Libs laden ───────────────────────────────────────────────────────────
  useEffect(() => {
    const loadScript = src =>
      new Promise((res, rej) => {
        if (document.querySelector(`script[src="${src}"]`)) return res();
        const s = document.createElement('script');
        s.src = src; s.onload = res; s.onerror = rej;
        document.head.appendChild(s);
      });

    Promise.all([
      loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'),
      loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js'),
    ])
      .then(() => setLibsReady(true))
      .catch(() => addToast('PDF-Bibliothek konnte nicht geladen werden.', 'error'));
  }, []); // eslint-disable-line

  // ─── Kurs-Pool laden ──────────────────────────────────────────────────────────
  const fetchCoursePool = useCallback(async () => {
    setIsLoadingPool(true);
    try {
      const response = await fetch(COURSE_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ request: 'get_courses' }),
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const rawData = await response.json();
      const items = Array.isArray(rawData)
        ? rawData
        : Array.isArray(rawData?.value)
        ? rawData.value
        : [];

      const normalized = items.map((item, index) => {
        const rawId = findValueByPattern(item, ['id', 'guid', 'key', 'ident']) || `c-${index}`;
        const label = findValueByPattern(item, ['label', 'name', 'titel', 'title', 'bezeichnung', 'kurs']) || 'Unbenannter Kurs';
        let shorthand = findValueByPattern(item, ['kurzel', 'kuerzel', 'kürzel', 'shorthand', 'short', 'code', 'kennung', 'abbr']);
        if (!shorthand || isGuid(String(shorthand))) {
          shorthand = String(label).replace(/[^a-zA-Z0-9 ]/g, '').split(' ').filter(w => w.length > 0).map(w => w[0]).join('').toUpperCase();
          if (shorthand.length < 2) shorthand = String(label).substring(0, 3).toUpperCase();
        }
        const url = findValueByPattern(item, ['url', 'link', 'hyperlink', 'moodle', 'portal']) || '';
        return { id: String(rawId), label: String(label).trim(), shorthand: String(shorthand).trim(), url: String(url).trim() };
      });

      setCourseDictionary(normalized);
      if (normalized.length > 0) {
        setConfig(prev => {
          const updatedIds = [...prev.selectedPoolCourseIds];
          normalized.slice(0, prev.courseSlotCount).forEach((c, i) => {
            if (updatedIds[i] === 'none') updatedIds[i] = c.id;
          });
          return { ...prev, selectedPoolCourseIds: updatedIds };
        });
      }
      addToast(`${normalized.length} Kurse geladen.`, 'success', 2500);
    } catch (e) {
      addToast('Verbindung zum Kurs-Pool fehlgeschlagen.', 'error');
    } finally {
      setIsLoadingPool(false);
    }
  }, [addToast]);

  useEffect(() => { fetchCoursePool(); }, [fetchCoursePool]);

  // ─── Berechnungen ─────────────────────────────────────────────────────────────
  const activeMatrixCourses = useMemo(
    () =>
      config.selectedPoolCourseIds
        .slice(0, config.courseSlotCount)
        .map(id => courseDictionary.find(c => String(c.id) === String(id)))
        .filter(c => c && c.id !== 'none'),
    [config.selectedPoolCourseIds, config.courseSlotCount, courseDictionary]
  );

  const totals = useMemo(() => {
    let std = 0;
    Object.entries(config.classCounts).forEach(([idx, count]) => {
      std += (config.classSizes[parseInt(idx)] || 0) * count;
    });
    const cls = Object.values(config.classCounts).reduce((a, b) => a + b, 0);
    return { cls, std, trainers: config.trainerCount, all: std + config.trainerCount };
  }, [config.classCounts, config.classSizes, config.trainerCount]);

  const classRows = useMemo(() => {
    const rows = []; let id = 1;
    [0, 1, 2, 3].forEach(idx => {
      const size = config.classSizes[idx];
      for (let i = 0; i < config.classCounts[idx]; i++) rows.push({ id: id++, size });
    });
    return rows;
  }, [config.classCounts, config.classSizes]);

  const rawEndDate = useMemo(() => {
    const d = new Date(config.enrolDate);
    d.setDate(d.getDate() + parseInt(config.enrolPeriod || 0));
    return d;
  }, [config.enrolDate, config.enrolPeriod]);

  const endDateDisplay = rawEndDate.toISOString().split('T')[0];
  const endDateFormatted = rawEndDate.toLocaleDateString('de-DE');

  // ─── Handler ──────────────────────────────────────────────────────────────────
  const handleInput = useCallback(e => {
    const { name, value } = e.target;
    const isNumeric = ['enrolPeriod', 'trainerCount', 'courseSlotCount'].includes(name);
    setConfig(p => ({ ...p, [name]: isNumeric ? Math.max(0, parseInt(value, 10) || 0) : value }));
  }, []);

  const handleEndDateInput = useCallback(e => {
    if (!e.target.value) return;
    const diff = Math.ceil((new Date(e.target.value) - new Date(config.enrolDate)) / 86400000);
    setConfig(p => ({ ...p, enrolPeriod: Math.max(0, diff) }));
  }, [config.enrolDate]);

  const updateClassSize = useCallback((idx, val) => {
    setConfig(p => {
      const s = [...p.classSizes];
      s[idx] = Math.max(0, parseInt(val, 10) || 0);
      return { ...p, classSizes: s };
    });
  }, []);

  const updateClassCount = useCallback((idx, val) => {
    setConfig(p => ({ ...p, classCounts: { ...p.classCounts, [idx]: Math.max(0, parseInt(val, 10) || 0) } }));
  }, []);

  const toggleCourseAssignment = useCallback((classId, courseId) => {
    setClassMatrix(prev => {
      const cur = (prev[classId] || []).map(String);
      const sid = String(courseId);
      return { ...prev, [classId]: cur.includes(sid) ? cur.filter(x => x !== sid) : [...cur, sid] };
    });
  }, []);

  const updateCourseSlot = useCallback((slotIndex, newId) => {
    setConfig(prev => {
      const ids = [...prev.selectedPoolCourseIds];
      ids[slotIndex] = newId;
      return { ...prev, selectedPoolCourseIds: ids };
    });
  }, []);

  // ─── Reset ────────────────────────────────────────────────────────────────────
  const handleReset = useCallback(async () => {
    if (!window.confirm('Alle Einstellungen zurücksetzen?')) return;
    setConfig(DEFAULT_CONFIG);
    setClassMatrix({});
    setGeneratedData([]);
    setIsGenerated(false);
    try {
      await store.clear();
      await store.save();
      addToast('Einstellungen zurückgesetzt.', 'success');
    } catch {
      addToast('Reset fehlgeschlagen.', 'error');
    }
  }, [addToast]);

  // ─── Generierung ──────────────────────────────────────────────────────────────
  const generateList = useCallback(() => {
    if (!config.institute?.trim()) return addToast('Bitte Institutsname eingeben.', 'error');
    if (classRows.length === 0 && config.trainerCount === 0) return addToast('Keine Klassen oder Trainer konfiguriert.', 'error');

    const activeIds = activeMatrixCourses.map(c => String(c.id));
    const missing = classRows.find(r => {
      const assigned = (classMatrix[r.id] || []).filter(id => activeIds.includes(String(id)));
      return assigned.length === 0;
    });
    if (missing) return addToast(`Klasse ${String(missing.id).padStart(2, '0')} hat keine Kurszuweisung.`, 'error');

    const instClean = config.institute.replace(/\s+/g, '').toLowerCase();
    const data = [];

    for (let t = 1; t <= config.trainerCount; t++) {
      data.push({
        cNum: 'ALL', isT: true, first: 'Trainer', last: config.institute,
        user: `${instClean}-trainer-${t}`, mail: `trainer${t}@${instClean}.com`,
        pw: config.trainerPwd, courses: activeMatrixCourses,
      });
    }

    let sIdx = 1;
    classRows.forEach(r => {
      const selIds = (classMatrix[r.id] || []).map(String);
      const selCourses = courseDictionary.filter(cd => selIds.includes(String(cd.id)) && activeIds.includes(String(cd.id)));
      const classLabel = `${config.institute}-Klasse-${String(r.id).padStart(2, '0')}`;
      for (let i = 0; i < r.size; i++) {
        const id = String(sIdx++).padStart(3, '0');
        data.push({
          cNum: String(r.id).padStart(2, '0'), cLabel: classLabel, isT: false,
          first: 'Schüler', last: config.institute,
          user: `${instClean}-student-${id}`, mail: `student${id}@${instClean}.com`,
          pw: config.studentPwd, courses: selCourses,
        });
      }
    });

    setGeneratedData(data);
    setIsGenerated(true);
    setActiveModal('dataPreview');
    addToast(`${data.length} Accounts generiert.`, 'success');
  }, [config, classRows, classMatrix, activeMatrixCourses, courseDictionary, addToast]);

  // ─── CSV Export ───────────────────────────────────────────────────────────────
  const downloadCSV = useCallback(() => {
    if (!generatedData.length) return;

    const rowsWithEnrols = generatedData.map(r => {
      const enrols = [];
      if (r.isT) {
        activeMatrixCourses.forEach(c =>
          classRows.forEach(cls =>
            enrols.push({ shorthand: c.shorthand, group: `${config.institute}-Klasse-${String(cls.id).padStart(2, '0')}`, role: 4, period: config.enrolPeriod })
          )
        );
      } else {
        r.courses.forEach(c => enrols.push({ shorthand: c.shorthand, group: r.cLabel, role: 5, period: config.enrolPeriod }));
      }
      return { ...r, enrols };
    });

    const maxCourses = Math.max(...rowsWithEnrols.map(r => r.enrols.length), 1);
    const headers = ['username', 'firstname', 'lastname', 'email', 'password', 'cohort1'];
    for (let i = 1; i <= maxCourses; i++) headers.push(`course${i}`, `group${i}`, `role${i}`, `enrolperiod${i}`);

    const lines = rowsWithEnrols.map(r => {
      const line = [esc(r.user), esc(r.first), esc(r.last), esc(r.mail), esc(r.pw), esc(r.last)];
      for (let i = 0; i < maxCourses; i++) {
        const e = r.enrols[i];
        line.push(e ? esc(e.shorthand) : '""', e ? esc(e.group) : '""', e ? esc(e.role) : '""', e ? esc(e.period) : '""');
      }
      return line.join(',');
    });

    const blob = new Blob(['\uFEFF', headers.join(',') + '\r\n' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: `EBCL-Moodle-Upload-${config.institute.replace(/\s+/g, '_')}-${new Date().toISOString().split('T')[0]}.csv` });
    a.click();
    URL.revokeObjectURL(a.href);
    addToast('CSV heruntergeladen.', 'success');
  }, [generatedData, activeMatrixCourses, classRows, config, addToast]);

  // ─── PDF Export ───────────────────────────────────────────────────────────────
  const downloadPDF = useCallback(async () => {
    if (!window.jspdf || !libsReady) return addToast('PDF-Bibliothek noch nicht bereit.', 'error');
    setIsExportingPDF(true);
    try {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF('p', 'mm', 'a4');
      const primary = [157, 32, 43];
      const today = new Date().toISOString().split('T')[0];
      const filename = `EBCL-Zugangsdaten-${config.institute.replace(/\s+/g, '_')}-${today}.pdf`;

      const renderHeader = (title, info) => {
        doc.setFontSize(20).setTextColor(...primary).setFont('helvetica', 'bold').text(config.institute.toUpperCase(), 15, 20);
        doc.setFontSize(12).setTextColor(60).setFont('helvetica', 'normal').text(title, 15, 28);
        doc.setFontSize(10).setTextColor(37, 99, 235).text('Portal: https://world.ebcl.eu/', 15, 34).link(15, 31, 50, 4, { url: 'https://world.ebcl.eu/' });
        doc.setFontSize(8).setTextColor(120).text(`ZEITRAUM: ${new Date(config.enrolDate).toLocaleDateString('de-DE')} bis ${endDateFormatted} | ${info}`, 15, 42);
        doc.setLineWidth(0.3).setDrawColor(0).line(15, 45, 195, 45);
      };

      const tOpts = courses => ({
        startY: 50, theme: 'grid', styles: { fontSize: 7, textColor: [0, 0, 0] },
        headStyles: { fillColor: [40, 40, 40], textColor: 255 },
        columnStyles: {
          0: { cellWidth: 45 }, 1: { cellWidth: 35 }, 2: { cellWidth: 25 },
          ...Object.fromEntries(Array.from({ length: courses.length }, (_, i) => [i + 3, { textColor: primary, fontStyle: 'bold' }])),
        },
        didDrawCell: ({ cell, column, section, row }) => {
          if (section === 'body' && column.index >= 3) {
            const c = courses[column.index - 3];
            if (c?.url) doc.link(cell.x, cell.y, cell.width, cell.height, { url: c.url });
          }
        },
      });

      const trainers = generatedData.filter(d => d.isT);
      if (trainers.length > 0) {
        renderHeader('Zugangsdaten: Trainer', `TRAINER: ${trainers.length}`);
        doc.autoTable({
          head: [['Name (Eingabefeld)', 'Username', 'Passwort', ...activeMatrixCourses.map((_, i) => `Kurs ${i + 1}`)]],
          body: trainers.map(t => ['', t.user, t.pw, ...t.courses.map(c => c.label)]),
          ...tOpts(activeMatrixCourses),
          didParseCell: d => { if (d.section === 'body') d.cell.styles.fillColor = [255, 255, 245]; },
        });
      }

      const classIds = [...new Set(generatedData.filter(d => !d.isT).map(d => d.cNum))].sort();
      classIds.forEach((id, idx) => {
        if (trainers.length > 0 || idx > 0) doc.addPage();
        const students = generatedData.filter(d => d.cNum === id);
        renderHeader(`Teilnehmerliste: Klasse-${id}`, `SCHÜLER: ${students.length}`);
        doc.autoTable({
          head: [['Name (Eingabefeld)', 'Username', 'Passwort', ...students[0].courses.map((_, i) => `Kurs ${i + 1}`)]],
          body: students.map(s => ['', s.user, s.pw, ...s.courses.map(c => c.label)]),
          ...tOpts(students[0].courses),
        });
      });

      const pageCount = doc.internal.getNumberOfPages();
      for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i).setFontSize(7).setTextColor(150).text(filename, 15, 287);
        doc.text(`Seite ${i} von ${pageCount}`, 185, 287, { align: 'right' });
      }

      doc.save(filename);
      addToast('PDF exportiert.', 'success');
    } catch (e) {
      console.error('PDF Fehler:', e);
      addToast('PDF-Export fehlgeschlagen.', 'error');
    } finally {
      setIsExportingPDF(false);
    }
  }, [generatedData, libsReady, config, activeMatrixCourses, endDateFormatted, addToast]);

  // ─── Save-Status Badge ────────────────────────────────────────────────────────
  const SaveBadge = () => {
    if (saveStatus === 'idle') return null;
    const map = {
      saving: { icon: <Loader2 size={10} className="animate-spin" />, text: 'Speichert…', color: 'text-slate-400' },
      saved:  { icon: <Check size={10} />,                            text: 'Gespeichert', color: 'text-emerald-500' },
      error:  { icon: <AlertTriangle size={10} />,                    text: 'Fehler',      color: 'text-rose-500' },
    };
    const { icon, text, color } = map[saveStatus] || {};
    return (
      <span className={`flex items-center gap-1 text-[9px] font-semibold uppercase tracking-widest ${color} transition-all`}>
        {icon}{text}
      </span>
    );
  };

  // ─── Modals ───────────────────────────────────────────────────────────────────
  const renderHelpModal = () => (
    <div className="fixed inset-0 z-[250] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-3xl overflow-hidden border border-slate-100 flex flex-col max-h-[90vh]">
        <div className="p-6 md:p-8 bg-slate-50 border-b flex justify-between items-start shrink-0 relative overflow-hidden">
          <div className="absolute top-0 right-0 p-8 opacity-5 pointer-events-none"><Zap size={140} style={{ color: COLORS.main }} /></div>
          <div className="flex items-center gap-4 relative z-10">
            <div style={{ backgroundColor: COLORS.main }} className="p-3.5 text-white rounded-2xl shadow-lg"><BookOpen size={24} /></div>
            <div>
              <h3 className="text-xl font-bold text-slate-800 tracking-tight leading-none uppercase">Einfach erklärt</h3>
              <p className="text-[10px] text-slate-500 mt-1.5 font-semibold uppercase tracking-[0.2em]">Workflow für Mitarbeiter</p>
            </div>
          </div>
          <button onClick={() => setActiveModal(null)} className="p-2 hover:bg-slate-200 rounded-full text-slate-400 transition-all active:scale-90"><X size={24} /></button>
        </div>
        <div className="flex-1 overflow-y-auto p-6 md:p-8 space-y-8">
          <section>
            <div className="flex items-center gap-2.5 mb-3">
              <div style={{ color: COLORS.main }} className="bg-red-50 p-1.5 rounded-lg"><Zap size={16} /></div>
              <h4 className="text-sm font-bold text-slate-800 uppercase tracking-widest">Was macht dieses Programm?</h4>
            </div>
            <p className="text-xs text-slate-500 leading-relaxed pl-9 border-l-2 border-slate-100">
              Als EBCL-Mitarbeiter erstellst du mit diesem Tool hunderte Moodle-Zugänge in wenigen Sekunden für unsere Partnerinstitute. Alle Einstellungen werden automatisch lokal gespeichert und sind beim nächsten Start sofort wieder verfügbar.
            </p>
          </section>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8 text-slate-500 text-[11px]">
            {[
              ['01', 'Einstellungen', 'Gib den Namen des Partner-Instituts ein und definiere Klassen, Trainer und Einschreibedauer.'],
              ['02', 'Kursverteilung', 'Wähle Kurse aus dem Moodle-Pool und weise sie den Klassen per Plus (+) zu.'],
              ['03', 'Trainer-Zugang', 'Trainer erhalten automatisch Zugriff auf alle gewählten Kurse und Klassen.'],
              ['04', 'Ergebnis sichern', 'CSV: Import-Datei für Moodle.\nPDF: Zugangsliste für das Institut.'],
            ].map(([num, title, desc]) => (
              <div key={num} className="flex gap-3 items-start">
                <span style={{ color: COLORS.accent1, backgroundColor: COLORS.accent1 + '10' }} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px]">{num}</span>
                <div className="space-y-1"><h5 className="font-bold text-slate-700 uppercase tracking-wider">{title}</h5><p className="whitespace-pre-line">{desc}</p></div>
              </div>
            ))}
          </div>
        </div>
        <div className="p-6 md:p-8 bg-slate-50 border-t flex justify-end shrink-0">
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: COLORS.accent1 }} className="text-white px-10 py-3 rounded-2xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all text-xs">Alles klar!</button>
        </div>
      </div>
    </div>
  );

  const renderSettingsModal = () => (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-xl w-full max-w-md overflow-hidden border border-slate-200">
        <div className="p-6 bg-slate-50 border-b flex justify-between items-center">
          <div className="flex items-center gap-3">
            <div style={{ backgroundColor: COLORS.accent1 }} className="p-2.5 text-white rounded-xl shadow-md"><Settings size={18} /></div>
            <h3 className="font-bold text-slate-800 uppercase tracking-tight text-sm">System-Settings</h3>
          </div>
          <button onClick={() => setActiveModal(null)} className="p-2 hover:bg-slate-200 rounded-full text-slate-400 transition-colors"><X size={20} /></button>
        </div>
        <div className="p-6 space-y-5">
          <h4 className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2"><Users size={14} /> Klassengrößen definieren</h4>
          <div className="grid grid-cols-2 gap-3">
            {config.classSizes.map((size, idx) => (
              <div key={idx} className="bg-white p-3 rounded-xl border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                <label className="text-[9px] font-semibold text-slate-400 uppercase block mb-1">Typ {idx + 1}</label>
                <input type="number" min="0" value={size} onChange={e => updateClassSize(idx, e.target.value)} style={{ color: COLORS.main }} className="w-full bg-transparent text-base font-bold outline-none" />
              </div>
            ))}
          </div>

          <div className="border-t pt-4">
            <h4 className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2 mb-3"><Save size={14} /> Datenspeicherung</h4>
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 text-[10px] text-emerald-700 flex items-start gap-2">
              <CheckCircle2 size={14} className="shrink-0 mt-0.5" />
              <div>
                <span className="font-bold block mb-0.5">Automatisch gespeichert</span>
                Alle Einstellungen werden in <code className="bg-white px-1 rounded">moodle-settings.json</code> lokal gespeichert und sind nach einem Neustart sofort verfügbar.
              </div>
            </div>
          </div>
        </div>
        <div className="p-6 bg-slate-50 border-t flex justify-between items-center">
          <button onClick={handleReset} className="text-rose-500 hover:text-rose-700 text-[10px] font-bold uppercase flex items-center gap-1.5 hover:bg-rose-50 px-3 py-2 rounded-lg transition-all">
            <RefreshCw size={12} /> Zurücksetzen
          </button>
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: COLORS.accent1 }} className="text-white px-8 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all text-xs">Schließen</button>
        </div>
      </div>
    </div>
  );

  const renderCoursePreviewModal = () => (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-xl w-full max-w-4xl overflow-hidden flex flex-col border border-slate-200 max-h-[90vh]">
        <div className="p-5 bg-slate-50 border-b flex flex-wrap justify-between items-center gap-4 shrink-0">
          <div className="flex items-center gap-3">
            <div style={{ backgroundColor: COLORS.accent1 }} className="p-2 text-white rounded-xl shadow-sm"><Eye size={18} /></div>
            <div>
              <h3 className="font-bold text-slate-800 uppercase tracking-tight text-sm">Kursübersicht</h3>
              <p className="text-[9px] text-slate-400 mt-0.5">{courseDictionary.length} Kurse verfügbar</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={fetchCoursePool} disabled={isLoadingPool} style={{ color: COLORS.accent2 }} className="flex items-center gap-2 px-3 py-1.5 bg-emerald-50 rounded-lg text-xs font-bold uppercase hover:bg-emerald-100 disabled:opacity-50 transition-colors">
              <RefreshCw size={14} className={isLoadingPool ? 'animate-spin' : ''} /> Sync
            </button>
            <button onClick={() => setActiveModal(null)} className="p-2 hover:bg-slate-200 rounded-full text-slate-400"><X size={20} /></button>
          </div>
        </div>
        <div className="flex-1 overflow-auto bg-white">
          <table className="w-full text-left border-collapse min-w-[600px]">
            <thead className="bg-slate-50 sticky top-0 z-10 border-b border-slate-200 text-[10px] font-semibold text-slate-500 uppercase tracking-widest">
              <tr>
                <th className="px-5 py-3">#</th>
                <th className="px-5 py-3">Kursname</th>
                <th className="px-5 py-3">Kürzel</th>
                <th className="px-5 py-3">Link</th>
              </tr>
            </thead>
            <tbody className="divide-y text-[11px] text-slate-600">
              {isLoadingPool ? (
                <tr><td colSpan="4" className="px-5 py-8 text-center text-slate-400"><Loader2 size={20} className="animate-spin mx-auto" /></td></tr>
              ) : courseDictionary.length === 0 ? (
                <tr><td colSpan="4" className="px-5 py-8 text-center text-slate-400 italic">Keine Kurse gefunden.</td></tr>
              ) : (
                courseDictionary.map((c, idx) => (
                  <tr key={c.id} className="hover:bg-slate-50 transition-colors">
                    <td className="px-5 py-3 text-slate-400 font-mono text-[10px]">{idx + 1}</td>
                    <td className="px-5 py-3 font-semibold text-slate-700">{c.label}</td>
                    <td className="px-5 py-3">
                      <span style={{ color: COLORS.accent1, backgroundColor: COLORS.accent1 + '1A' }} className="font-mono px-2 py-0.5 rounded text-[10px]">{c.shorthand}</span>
                    </td>
                    <td className="px-5 py-3">
                      {c.url
                        ? <a href={c.url} target="_blank" rel="noreferrer" style={{ color: COLORS.accent1 }} className="hover:underline flex items-center gap-1">Öffnen <Eye size={12} /></a>
                        : <span className="text-slate-400 italic">Kein Link</span>}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
        <div className="p-5 bg-slate-50 border-t flex justify-end">
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: COLORS.main }} className="text-white px-6 py-2 rounded-xl font-bold uppercase text-xs shadow-sm hover:brightness-110 active:scale-95 transition-all">Schließen</button>
        </div>
      </div>
    </div>
  );

  const renderDataPreviewModal = () => (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-xl w-full max-w-5xl overflow-hidden flex flex-col max-h-[90vh]">
        <div className="p-6 bg-slate-800 border-b flex justify-between items-center shrink-0">
          <div className="flex items-center gap-3 text-white">
            <div style={{ backgroundColor: COLORS.accent1 }} className="p-2.5 rounded-xl shadow-sm"><ClipboardList size={20} /></div>
            <div>
              <h3 className="text-base font-bold uppercase tracking-tight leading-none">Daten-Vorschau</h3>
              <p className="text-[10px] text-slate-400 mt-0.5">{generatedData.length} Einträge · {generatedData.filter(d => d.isT).length} Trainer · {generatedData.filter(d => !d.isT).length} Schüler</p>
            </div>
          </div>
          <button onClick={() => setActiveModal(null)} className="p-2 hover:bg-white/10 rounded-full text-slate-400 transition-colors"><X size={24} /></button>
        </div>
        <div className="flex-1 overflow-auto bg-white">
          <table className="w-full text-[11px] table-auto border-collapse text-slate-700 min-w-[700px]">
            <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm border-b font-semibold text-slate-500 uppercase tracking-widest">
              <tr>
                <th className="px-6 py-4 text-left">Klasse / Typ</th>
                <th className="px-6 py-4 text-left">Username</th>
                <th className="px-6 py-4 text-left">Zuweisungen</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {generatedData.map((row, idx) => (
                <tr key={idx} className={row.isT ? 'bg-amber-50/40 font-semibold border-l-2 border-amber-300' : 'hover:bg-slate-50/50 transition-colors'}>
                  <td className="px-6 py-3">
                    <span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${row.isT ? 'bg-amber-100 text-amber-700' : 'bg-slate-100 text-slate-500'}`}>
                      {row.isT ? 'GLOBAL (Trainer)' : `K-${row.cNum}`}
                    </span>
                  </td>
                  <td className="px-6 py-3 font-mono font-medium text-slate-700 tracking-tight">{row.user}</td>
                  <td className="px-6 py-3">
                    <div className="flex flex-wrap gap-1.5">
                      {row.courses.map((c, ci) => (
                        <span key={ci} style={{ color: COLORS.main, backgroundColor: COLORS.main + '1A' }} className="font-semibold text-[9px] px-1.5 py-0.5 rounded border border-red-50">{c.shorthand}</span>
                      ))}
                      {row.courses.length === 0 && <span className="text-slate-400 italic text-[10px]">Keine Kurse</span>}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="p-5 bg-slate-50 border-t flex justify-end gap-3">
          <button disabled={isExportingPDF} onClick={downloadPDF} style={{ backgroundColor: COLORS.accent1 }} className="text-white px-5 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm flex items-center gap-2 text-xs disabled:opacity-50">
            {isExportingPDF ? <Loader2 size={14} className="animate-spin" /> : <FileDown size={14} />} PDF
          </button>
          <button onClick={downloadCSV} style={{ backgroundColor: COLORS.accent2 }} className="text-white px-5 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm flex items-center gap-2 text-xs">
            <FileSpreadsheet size={14} /> CSV
          </button>
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: COLORS.main }} className="text-white px-8 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm text-xs">Schließen</button>
        </div>
      </div>
    </div>
  );

  // ─── Main Render ──────────────────────────────────────────────────────────────
  return (
    <div style={{ backgroundColor: COLORS.bg }} className="h-screen flex flex-col p-4 md:p-6 text-slate-800 font-sans selection:bg-red-100 overflow-hidden relative">

      {/* Toast Notifications */}
      <Toast toasts={toasts} removeToast={removeToast} />

      {/* Modals */}
      {activeModal === 'help' && renderHelpModal()}
      {activeModal === 'settings' && renderSettingsModal()}
      {activeModal === 'coursePreview' && renderCoursePreviewModal()}
      {activeModal === 'dataPreview' && renderDataPreviewModal()}

      {/* HEADER */}
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-5 shrink-0">
        <div>
          <h1 className="text-xl font-bold flex items-center gap-2 text-slate-800 tracking-tight">
            <TableIcon style={{ color: COLORS.main }} size={24} /> Moodle Anmeldungen
          </h1>
          <div className="flex items-center gap-3 mt-1">
            <p className="text-slate-500 text-xs italic flex items-center gap-1.5 font-medium">
              {isLoadingPool
                ? <Loader2 size={12} className="animate-spin text-blue-500" />
                : <CheckCircle2 style={{ color: COLORS.accent2 }} size={12} />}
              Kurse verfügbar ({courseDictionary.length})
            </p>
            <span className="text-slate-200">·</span>
            <span className="flex items-center gap-1 text-[10px] text-slate-400 font-medium">
              {isOnline ? <Wifi size={11} className="text-emerald-400" /> : <WifiOff size={11} className="text-rose-400" />}
              {isOnline ? 'Online' : 'Offline'}
            </span>
            <span className="text-slate-200">·</span>
            <SaveBadge />
          </div>
        </div>
        <div className="flex items-center gap-3 bg-white px-3 py-1.5 rounded-lg border border-slate-200 shadow-sm">
          <div className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">V3</div>
          <div className="h-3 w-px bg-slate-200"></div>
          <div className="text-[9px] font-bold text-slate-500 uppercase tracking-widest">EBCL INTERNATIONAL</div>
        </div>
      </header>

      {/* HAUPT-GRID */}
      <div className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-5 lg:gap-6 overflow-hidden min-h-0">

        {/* SIDEBAR */}
        <aside className="lg:col-span-3 flex flex-col gap-4 overflow-y-auto pr-1 scrollbar-thin">

          {/* Status Card */}
          <section style={{ backgroundColor: COLORS.accent1 }} className="rounded-3xl p-5 shadow-lg flex flex-col text-white shrink-0 relative overflow-hidden">
            <div className="absolute -right-6 -top-6 opacity-10"><Users size={100} /></div>
            <div className="flex justify-between items-end mb-3 border-b border-white/10 pb-2 relative z-10">
              <div>
                <div className="text-[9px] font-semibold text-white/70 uppercase tracking-widest mb-1">Gesamt Accounts</div>
                <div className="text-2xl font-bold leading-none">{totals.all}</div>
              </div>
              <div className="text-[9px] text-white/50 font-medium">{totals.cls} Klasse{totals.cls !== 1 ? 'n' : ''}</div>
            </div>
            <div className="space-y-2 relative z-10">
              <div className="flex justify-between items-center text-[11px] bg-white/5 px-2 py-1.5 rounded-lg">
                <span className="text-white/70 font-medium uppercase tracking-wider flex items-center gap-1.5"><ShieldCheck size={12} /> Trainer</span>
                <span className="font-bold text-amber-400 text-sm">{totals.trainers}</span>
              </div>
              <div className="flex justify-between items-center text-[11px] bg-white/5 px-2 py-1.5 rounded-lg">
                <span className="text-white/70 font-medium uppercase tracking-wider flex items-center gap-1.5"><GraduationCap size={12} /> Schüler</span>
                <span className="font-bold text-blue-300 text-sm">{totals.std}</span>
              </div>
            </div>
          </section>

          {/* Actions */}
          <section className="bg-white p-4 rounded-3xl border border-slate-200 shadow-sm shrink-0">
            <button onClick={() => setActiveModal('coursePreview')} className="w-full py-2.5 bg-slate-50 text-slate-600 border border-slate-200 rounded-xl text-[10px] font-bold uppercase tracking-widest hover:bg-slate-100 transition-all flex items-center justify-center gap-2">
              <Eye size={16} /> Kursübersicht
            </button>
            <button
              disabled={isLoadingPool || courseDictionary.length === 0}
              onClick={generateList}
              style={{ backgroundColor: COLORS.main }}
              className="w-full py-3 text-white rounded-xl font-bold shadow-md mt-3 transition-all hover:brightness-110 active:scale-95 flex items-center justify-center gap-2 disabled:opacity-50 group text-sm"
            >
              <Users size={16} className="group-hover:scale-110 transition-transform" /> Liste generieren
            </button>
            <button
              disabled={!isGenerated}
              onClick={() => setActiveModal('dataPreview')}
              style={{ color: COLORS.accent1, borderColor: COLORS.accent1 + '33' }}
              className="w-full py-2.5 bg-white border rounded-xl text-[10px] font-bold uppercase mt-3 hover:bg-slate-50 disabled:opacity-40 transition-all flex items-center justify-center gap-1.5"
            >
              <ClipboardList size={14} /> Daten-Vorschau
            </button>
            <div className="grid grid-cols-2 gap-2 pt-3 border-t border-slate-100 mt-3">
              <button disabled={!isGenerated || isExportingPDF} onClick={downloadPDF} style={{ backgroundColor: COLORS.accent1 }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1.5 transition-all hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest">
                {isExportingPDF ? <Loader2 size={14} className="animate-spin" /> : <FileDown size={14} />} PDF
              </button>
              <button disabled={!isGenerated} onClick={downloadCSV} style={{ backgroundColor: COLORS.accent2 }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1.5 transition-all hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest">
                <FileSpreadsheet size={14} /> CSV
              </button>
            </div>
          </section>

          {/* Info & Settings Footer */}
          <div style={{ backgroundColor: COLORS.accent1 + '08', borderColor: COLORS.accent1 + '15' }} className="p-4 border rounded-2xl text-[10px] leading-relaxed shadow-sm relative mt-auto">
            <div style={{ color: COLORS.accent1 }} className="flex items-center gap-1.5 mb-3 font-bold uppercase tracking-widest">
              <Info size={14} /> Workflow
            </div>
            <div className="space-y-2.5 text-slate-500 font-medium mb-4">
              {[['1', 'Konfiguration festlegen.'], ['2', 'Kurse unten zuteilen.'], ['3', 'Generieren & Export.']].map(([n, t]) => (
                <div key={n} className="flex gap-2.5 items-center">
                  <span style={{ backgroundColor: COLORS.accent1 }} className="text-white w-4 h-4 rounded-full flex items-center justify-center shrink-0 font-bold text-[8px]">{n}</span>
                  <p>{t}</p>
                </div>
              ))}
            </div>
            <div className="flex gap-2 pt-3 border-t border-slate-200/60">
              <button onClick={() => setActiveModal('settings')} className="flex-1 py-2 bg-white border border-slate-200 text-slate-500 rounded-lg hover:text-slate-800 hover:border-slate-300 hover:bg-slate-50 transition-all flex items-center justify-center gap-1.5 shadow-sm">
                <Settings size={12} /><span className="text-[9px] font-bold uppercase">Settings</span>
              </button>
              <button onClick={() => setActiveModal('help')} className="flex-1 py-2 bg-white border border-slate-200 text-slate-500 rounded-lg hover:text-blue-600 hover:border-blue-300 hover:bg-blue-50 transition-all flex items-center justify-center gap-1.5 shadow-sm">
                <HelpCircle size={12} /><span className="text-[9px] font-bold uppercase">Hilfe</span>
              </button>
            </div>
          </div>
        </aside>

        {/* HAUPTINHALT */}
        <main className="lg:col-span-9 flex flex-col gap-5 overflow-hidden min-h-0">

          {/* Konfiguration */}
          <section className="bg-white p-2 rounded-3xl shadow-sm border border-slate-200 shrink-0">
            <div className="grid grid-cols-1 md:grid-cols-12 gap-2">

              {/* Organisation */}
              <div className="md:col-span-4 bg-slate-50/50 p-4 lg:p-5 rounded-2xl">
                <h3 style={{ color: COLORS.accent1 }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><Building2 size={12} /> Organisation</h3>
                <div className="space-y-3">
                  <div>
                    <label className="text-[9px] font-semibold text-slate-500 uppercase block mb-1 ml-1">Institutsname</label>
                    <input name="institute" value={config.institute} onChange={handleInput} placeholder="z.B. Volkshochschule" className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm font-medium focus:ring-1 focus:border-blue-400 outline-none shadow-sm transition-all placeholder:font-normal placeholder:text-slate-300" />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div className="bg-white p-2.5 rounded-xl border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                      <label className="text-[8px] font-semibold text-slate-400 uppercase block mb-1">Trainer (Anz.)</label>
                      <input type="number" min="0" name="trainerCount" value={config.trainerCount} onChange={handleInput} style={{ color: COLORS.main }} className="w-full bg-transparent text-sm font-semibold outline-none" />
                    </div>
                    <div className="bg-white p-2.5 rounded-xl border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                      <label className="text-[8px] font-semibold text-slate-400 uppercase block mb-1">Kurs Anzahl</label>
                      <input type="number" min="1" max="8" name="courseSlotCount" value={config.courseSlotCount} onChange={handleInput} style={{ color: COLORS.accent1 }} className="w-full bg-transparent text-sm font-semibold outline-none" />
                    </div>
                  </div>
                </div>
              </div>

              {/* Klassen */}
              <div className="md:col-span-3 bg-white p-4 lg:p-5 rounded-2xl border border-slate-100 shadow-sm">
                <h3 className="text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><GraduationCap size={12} /> Klassen Struktur</h3>
                <div className="space-y-2">
                  {config.classSizes.map((size, idx) => (
                    <div key={idx} className="flex justify-between items-center bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200 focus-within:border-blue-300 transition-colors">
                      <span className="text-[11px] font-medium text-slate-600">{size} Plätze</span>
                      <div className="flex items-center gap-1.5">
                        <span className="text-[9px] text-slate-400">Anz:</span>
                        <input type="number" min="0" value={config.classCounts[idx]} onChange={e => updateClassCount(idx, e.target.value)} style={{ color: COLORS.main }} className="w-8 bg-transparent text-right text-sm font-semibold outline-none" />
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Zeit & Sicherheit */}
              <div className="md:col-span-5 bg-slate-50/50 p-4 lg:p-5 rounded-2xl">
                <h3 style={{ color: COLORS.main }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><ShieldCheck size={12} /> Zeit & Sicherheit</h3>
                <div className="grid grid-cols-2 gap-3 mb-2.5">
                  <div className="space-y-1">
                    <label className="text-[8px] font-semibold text-slate-500 uppercase ml-1">Einschreibung</label>
                    <input name="enrolDate" type="date" value={config.enrolDate} onChange={handleInput} className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-medium outline-none shadow-sm focus:border-blue-300 transition-colors" />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[8px] font-semibold text-slate-500 uppercase ml-1">Dauer (Tage)</label>
                    <input name="enrolPeriod" type="number" min="0" value={config.enrolPeriod} onChange={handleInput} className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-medium outline-none shadow-sm focus:border-blue-300 transition-colors" />
                  </div>
                </div>
                <div className="space-y-1 mb-3">
                  <label style={{ color: COLORS.main }} className="text-[8px] font-bold uppercase ml-1 flex justify-between">
                    <span>Eingeschrieben bis</span><span className="opacity-70 font-medium">(Autom. berechnet)</span>
                  </label>
                  <input type="date" value={endDateDisplay} onChange={handleEndDateInput} style={{ color: COLORS.main, borderColor: COLORS.main + '40' }} className="w-full px-3 py-1.5 bg-red-50 border rounded-lg text-xs font-bold outline-none shadow-sm focus:ring-1 focus:ring-red-200 transition-all" />
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div className="bg-white p-2 rounded-lg border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                    <label className="text-[8px] font-semibold text-slate-400 uppercase block mb-0.5">PW Schüler</label>
                    <input name="studentPwd" value={config.studentPwd} onChange={handleInput} className="w-full bg-transparent text-[11px] font-mono font-medium text-slate-600 outline-none" />
                  </div>
                  <div className="bg-white p-2 rounded-lg border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                    <label className="text-[8px] font-semibold text-slate-400 uppercase block mb-0.5">PW Trainer</label>
                    <input name="trainerPwd" value={config.trainerPwd} onChange={handleInput} className="w-full bg-transparent text-[11px] font-mono font-medium text-slate-600 outline-none" />
                  </div>
                </div>
              </div>
            </div>
          </section>

          {/* KURS ZUWEISUNG MATRIX */}
          <section className="bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden flex flex-col flex-1 min-h-0 relative">
            <div className="px-5 py-3 bg-white border-b border-slate-100 flex items-center justify-between shrink-0 relative z-[41]">
              <div className="flex items-center gap-1.5 text-[11px] font-bold uppercase text-slate-600 tracking-wider">
                <CheckCircle2 style={{ color: COLORS.accent2 }} size={16} /> Kurs Zuweisung Matrix
              </div>
              <div className="text-[9px] text-slate-400 font-medium bg-slate-50 px-2.5 py-1 rounded-full border border-slate-100">
                Klassen: {classRows.length} · Kurse: {config.courseSlotCount}
              </div>
            </div>

            <div className="overflow-auto flex-1 relative z-0 bg-white custom-scrollbar">
              <table className="w-full text-left border-separate border-spacing-0 min-w-max">
                <thead className="sticky top-0 z-[40]">
                  <tr>
                    <th className="px-4 py-2 w-28 text-center border-b border-slate-100 bg-white font-semibold text-slate-400 uppercase tracking-widest text-[9px] shadow-sm">Klasse</th>
                    {Array.from({ length: config.courseSlotCount }).map((_, i) => (
                      <th key={i} className="px-2 py-2.5 border-b border-slate-100 min-w-[130px] max-w-[160px] bg-white shadow-sm">
                        <div className="flex flex-col gap-1.5">
                          <span className="text-[9px] text-slate-400 uppercase tracking-widest font-semibold">Kurs {i + 1}</span>
                          <div className="relative w-full">
                            <select
                              value={config.selectedPoolCourseIds[i] || 'none'}
                              onChange={e => updateCourseSlot(i, e.target.value)}
                              className="w-full appearance-none bg-white border border-slate-200 rounded-md pl-2 pr-6 py-1.5 text-[11px] font-medium text-slate-700 outline-none cursor-pointer hover:border-blue-300 focus:border-blue-400 focus:ring-1 focus:ring-blue-50 transition-all shadow-sm truncate"
                            >
                              <option value="none">-- Nicht belegt --</option>
                              {courseDictionary.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
                            </select>
                            <ChevronDown size={14} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                          </div>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50 text-[11px]">
                  {classRows.length === 0 ? (
                    <tr>
                      <td colSpan={config.courseSlotCount + 1} className="py-10 text-center text-slate-400 italic font-medium bg-slate-50/30">
                        Bitte oben Klassen definieren.
                      </td>
                    </tr>
                  ) : (
                    classRows.map(c => (
                      <tr key={c.id} className="hover:bg-blue-50/20 transition-colors group">
                        <td className="px-4 py-3 text-center border-r border-slate-50 bg-slate-50/30 group-hover:bg-transparent transition-colors">
                          <div className="flex flex-col gap-1 items-center">
                            <span className="font-bold text-slate-700 text-[11px] tracking-tight">K-{String(c.id).padStart(2, '0')}</span>
                            <span style={{ color: COLORS.accent1 }} className="text-[9px] font-semibold uppercase tracking-tighter bg-white border border-slate-100 px-1.5 py-0.5 rounded shadow-sm">{c.size} Plätze</span>
                          </div>
                        </td>
                        {Array.from({ length: config.courseSlotCount }).map((_, i) => {
                          const cid = config.selectedPoolCourseIds[i];
                          const isActive = cid !== 'none';
                          const isSelected = isActive && (classMatrix[c.id] || []).map(String).includes(String(cid));
                          return (
                            <td key={i} className="px-4 py-2.5 text-center">
                              {isActive ? (
                                <button
                                  onClick={() => toggleCourseAssignment(c.id, cid)}
                                  title={isSelected ? 'Zuweisung entfernen' : 'Kurs zuweisen'}
                                  style={{
                                    backgroundColor: isSelected ? COLORS.accent2 : 'white',
                                    color: isSelected ? 'white' : '#94A3B8',
                                    borderColor: isSelected ? COLORS.accent2 : '#E2E8F0',
                                  }}
                                  className={`w-10 h-10 rounded-xl border-[1.5px] transition-all duration-200 flex items-center justify-center mx-auto shadow-sm ${isSelected ? 'scale-105 shadow-md shadow-emerald-900/5' : 'hover:border-slate-300 hover:bg-slate-50 hover:scale-105'}`}
                                >
                                  {isSelected ? <Check size={20} strokeWidth={2.5} /> : <Plus size={18} strokeWidth={2} />}
                                </button>
                              ) : (
                                <div className="w-10 h-10 rounded-xl border border-dashed border-slate-200 bg-slate-50/50 mx-auto flex items-center justify-center">
                                  <div className="w-1.5 h-1.5 rounded-full bg-slate-200" />
                                </div>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            <div className="absolute bottom-0 left-0 right-0 h-4 bg-gradient-to-t from-black/5 to-transparent pointer-events-none" />
          </section>
        </main>
      </div>

      <style dangerouslySetInnerHTML={{ __html: `
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f8fafc; border-radius: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        @keyframes slideInRight { from { opacity: 0; transform: translateX(20px); } to { opacity: 1; transform: translateX(0); } }
      ` }} />
    </div>
  );
};

export default App;