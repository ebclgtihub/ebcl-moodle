import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import QRCode from 'qrcode';
import { LazyStore } from '@tauri-apps/plugin-store';
import { check as checkUpdate } from '@tauri-apps/plugin-updater';
import {
  Users, FileSpreadsheet, CheckCircle2, Building2, Plus,
  Loader2, Table as TableIcon, Check, AlertTriangle, ChevronDown,
  Eye, X, RefreshCw, Info, Settings, HelpCircle, BookOpen,
  Zap, ClipboardList, ShieldCheck, GraduationCap, FileDown,
  Save, Wifi, WifiOff, Star, StarOff, Trash2, History,
  Moon, Sun, Keyboard, CheckSquare, Square, Edit3
} from 'lucide-react';

/**
 * Moodle Anmeldungen V4
 * - FAVORITEN: Institute speichern, laden, löschen + Schnellwahl-Dropdown
 * - VALIDIERUNG: Rote Matrix-Markierung bei fehlenden Kurszuweisungen
 * - ALLE ZUWEISEN: Ein-Klick alle Klassen auf alle aktiven Kurse
 * - KLASSEN-NAMEN: Anpassbar in den Settings
 * - EXPORT-HISTORY: Vollständige Liste aller Exporte
 * - DUNKELMODUS: System-synchron + manueller Toggle
 * - SHORTCUTS: ⌘G/E/P/A/? mit Erläuterung im Hilfe-Modal
 */

// ─── Store ─────────────────────────────────────────────────────────────────────
const store = new LazyStore('moodle-settings.json', { autoSave: false });

// ─── Theme ─────────────────────────────────────────────────────────────────────
const LIGHT = {
  main: '#9D202B', bg: '#FFFEF4', card: '#FFFFFF',
  border: '#E2E8F0', accent1: '#153D61', accent2: '#00664F',
  text: '#1E293B', muted: '#64748B', subtle: '#F8FAFC',
};
const DARK = {
  main: '#E05060', bg: '#0F1117', card: '#1A1D27',
  border: '#2D3148', accent1: '#4A7FBF', accent2: '#00A87A',
  text: '#E2E8F0', muted: '#94A3B8', subtle: '#1E2130',
};

// ─── Defaults ──────────────────────────────────────────────────────────────────
const DEFAULT_CONFIG = {
  institute: '',
  studentPwd: '***ENTFERNT***',
  trainerPwd: '***ENTFERNT***',
  autoPassword: false,
  enrolPeriod: 31,
  enrolDate: new Date().toISOString().split('T')[0],
  classSizes: [15, 20, 30, 40],
  classCounts: { 0: 1, 1: 1, 2: 1, 3: 1 },
  classNames: {},
  trainerCount: 2,
  courseSlotCount: 4,
  selectedPoolCourseIds: Array(8).fill('none'),
};

const COURSE_API_URL =
  '***ENTFERNT***';

const SHORTCUTS = [
  { keys: ['⌘/Ctrl', 'G'], desc: 'Liste generieren' },
  { keys: ['⌘/Ctrl', 'E'], desc: 'CSV exportieren' },
  { keys: ['⌘/Ctrl', 'P'], desc: 'PDF exportieren' },
  { keys: ['⌘/Ctrl', '⇧', 'A'], desc: 'Alle Kurse zuweisen' },
  { keys: ['⌘/Ctrl', '?'], desc: 'Hilfe öffnen' },
  { keys: ['Escape'], desc: 'Modal schließen' },
];

// ─── Hilfsfunktionen ───────────────────────────────────────────────────────────
const findValueByPattern = (item, patterns) => {
  const keys = Object.keys(item);
  for (const p of patterns) {
    const m = keys.find(k => k.toLowerCase() === p.toLowerCase());
    if (m && item[m] != null && String(item[m]).trim() !== '') return item[m];
  }
  for (const p of patterns) {
    const m = keys.find(k => k.toLowerCase().includes(p.toLowerCase()));
    if (m && item[m] != null && String(item[m]).trim() !== '') return item[m];
  }
  return null;
};
const isGuid = s => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);
const esc = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
const fmtTime = d => d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
const fmtDate = d => d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
const fmtDateTime = d => `${fmtDate(d)}, ${fmtTime(d)}`;

// ─── Toast ─────────────────────────────────────────────────────────────────────
const Toast = ({ toasts, removeToast }) => (
  <div className="fixed bottom-5 right-5 z-[999] flex flex-col gap-2 pointer-events-none">
    {toasts.map(t => (
      <div key={t.id} className={`flex items-center gap-3 px-4 py-3 rounded-2xl shadow-xl text-white text-xs font-semibold pointer-events-auto
        ${t.type === 'error' ? 'bg-rose-600' : t.type === 'success' ? 'bg-emerald-600' : 'bg-slate-700'}`}
        style={{ animation: 'slideInRight 0.25s ease' }}>
        {t.type === 'error' ? <AlertTriangle size={14} /> : <CheckCircle2 size={14} />}
        <span>{t.message}</span>
        <button onClick={() => removeToast(t.id)} className="ml-1 opacity-70 hover:opacity-100"><X size={12} /></button>
      </div>
    ))}
  </div>
);

// ─── Modal-Shells (Modul-Scope = stabiler Typ, kein Remount bei App-Re-Render) ──
const ModalShell = ({ children, C, maxW = 'max-w-3xl', zIndex = 200 }) => (
  <div className="fixed inset-0 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4" style={{ zIndex }}>
    <div style={{ backgroundColor: C.card, borderColor: C.border }} className={`rounded-3xl shadow-2xl w-full ${maxW} overflow-hidden border flex flex-col max-h-[90vh]`}>{children}</div>
  </div>
);

const ModalHeader = ({ icon, title, sub, onClose, iconBg, C }) => (
  <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-5 md:p-6 border-b flex justify-between items-center shrink-0">
    <div className="flex items-center gap-3">
      <div style={{ backgroundColor: iconBg || C.accent1 }} className="p-2.5 text-white rounded-xl shadow-md">{icon}</div>
      <div>
        <h3 style={{ color: C.text }} className="font-bold uppercase tracking-tight text-sm">{title}</h3>
        {sub && <p style={{ color: C.muted }} className="text-[9px] mt-0.5">{sub}</p>}
      </div>
    </div>
    <button onClick={onClose} style={{ color: C.muted }} className="p-2 hover:bg-black/10 rounded-full transition-colors"><X size={20} /></button>
  </div>
);

// ─── Klassenname-Eingabe (lokaler State + onBlur = kein Focus-Verlust) ─────────
const ClassNameRow = ({ row, savedValue, onUpdate, C }) => {
  const [val, setVal] = React.useState(savedValue);
  return (
    <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="flex items-center gap-3 px-3 py-2 rounded-xl border focus-within:border-blue-300 transition-colors">
      <span style={{ color: C.muted }} className="text-[10px] font-mono w-12 shrink-0">K-{String(row.id).padStart(2, '0')}</span>
      <input type="text" value={val}
        onChange={e => setVal(e.target.value)}
        onBlur={e => onUpdate(row.id - 1, e.target.value)}
        placeholder={`Klasse-${String(row.id).padStart(2, '0')}`}
        style={{ color: C.text, backgroundColor: 'transparent' }}
        className="flex-1 text-xs font-medium outline-none placeholder:opacity-25" />
    </div>
  );
};

// ─── Passwort-Generator (Moodle-konform) ───────────────────────────────────────
const generatePassword = () => {
  const U = 'ABCDEFGHJKLMNPQRSTUVWXYZ', L = 'abcdefghjkmnpqrstuvwxyz';
  const D = '23456789', S = '!@#$%&*-=?';
  const pool = U + L + D + S;
  const r = s => s[Math.floor(Math.random() * s.length)];
  return [r(U), r(L), r(D), r(S), ...Array.from({ length: 6 }, () => r(pool))]
    .sort(() => Math.random() - 0.5).join('');
};

// ─── QR-Code Cache (wird einmalig generiert und wiederverwendet) ───────────────
let _qrCache = null;
const getQrDataUrl = () => _qrCache
  ? Promise.resolve(_qrCache)
  : QRCode.toDataURL('https://world.ebcl.eu/', { width: 120, margin: 1, errorCorrectionLevel: 'M' })
      .then(url => { _qrCache = url; return url; });

// ─── App ───────────────────────────────────────────────────────────────────────
const App = () => {
  const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  const [darkMode, setDarkMode] = useState(prefersDark);
  const C = darkMode ? DARK : LIGHT;

  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [courseDictionary, setCourseDictionary] = useState([]);
  const [isLoadingPool, setIsLoadingPool] = useState(false);
  const [activeModal, setActiveModal] = useState(null);
  const [classMatrix, setClassMatrix] = useState({});
  const [generatedData, setGeneratedData] = useState([]);
  const [isGenerated, setIsGenerated] = useState(false);
  const [isExportingPDF, setIsExportingPDF] = useState(false);
  const [isStoreLoaded, setIsStoreLoaded] = useState(false);
  const [saveStatus, setSaveStatus] = useState('idle');
  const [lastSavedAt, setLastSavedAt] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [isOnline, setIsOnline] = useState(navigator.onLine);
  const [favorites, setFavorites] = useState([]);
  const [exportHistory, setExportHistory] = useState([]);
  const [invalidClassIds, setInvalidClassIds] = useState(new Set());
  const [pendingUpdate, setPendingUpdate] = useState(null);
  const [isInstalling, setIsInstalling] = useState(false);
  const [installProgress, setInstallProgress] = useState(0);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [showSessionResetConfirm, setShowSessionResetConfirm] = useState(false);
  const [appVersion, setAppVersion] = useState('...');
  const [isCheckingUpdate, setIsCheckingUpdate] = useState(false);

  const [helpTab, setHelpTab] = useState('workflow');
  const [classNamesResetKey, setClassNamesResetKey] = useState(0);
  const saveTimeoutRef = useRef(null);
  const toastIdRef = useRef(0);
  const generateRef = useRef(null);
  const csvRef = useRef(null);
  const pdfRef = useRef(null);
  const assignRef = useRef(null);

  // ─── Dark Mode System-Sync ────────────────────────────────────────────────
  useEffect(() => {
    const mq = window.matchMedia('(prefers-color-scheme: dark)');
    const h = e => setDarkMode(e.matches);
    mq.addEventListener('change', h);
    return () => mq.removeEventListener('change', h);
  }, []);

  // ─── Toast ────────────────────────────────────────────────────────────────
  const addToast = useCallback((message, type = 'info', duration = 4000) => {
    const id = ++toastIdRef.current;
    setToasts(prev => [...prev, { id, message, type }]);
    if (duration > 0) setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), duration);
  }, []);
  const removeToast = useCallback(id => setToasts(prev => prev.filter(t => t.id !== id)), []);

  // ─── Online ───────────────────────────────────────────────────────────────
  useEffect(() => {
    const on = () => { setIsOnline(true); addToast('Verbindung wiederhergestellt', 'success'); };
    const off = () => { setIsOnline(false); addToast('Keine Internetverbindung', 'error'); };
    window.addEventListener('online', on); window.addEventListener('offline', off);
    return () => { window.removeEventListener('online', on); window.removeEventListener('offline', off); };
  }, [addToast]);

  // ─── App-Version ──────────────────────────────────────────────────────────
  useEffect(() => {
    import('@tauri-apps/api/app').then(m => m.getVersion()).then(setAppVersion).catch(() => setAppVersion('0.1.6'));
  }, []); // eslint-disable-line

  // ─── Updater ──────────────────────────────────────────────────────────────
  useEffect(() => {
    if (import.meta.env.DEV) return; // Kein Auto-Check im Dev-Modus
    const t = setTimeout(async () => {
      try {
        const update = await checkUpdate();
        if (update) setPendingUpdate(update);
      } catch { /* still fail */ }
    }, 3000);
    return () => clearTimeout(t);
  }, []); // eslint-disable-line

  const handleManualUpdateCheck = useCallback(async () => {
    if (import.meta.env.DEV) {
      addToast('Updates nur in der fertigen App prüfbar (nicht im Dev-Modus).', 'info', 4000);
      return;
    }
    setIsCheckingUpdate(true);
    try {
      const update = await checkUpdate();
      if (update) {
        setPendingUpdate(update);
        setActiveModal(null);
        addToast(`Update v${update.version} verfügbar — Banner oben!`, 'success');
      } else {
        addToast('App ist bereits aktuell.', 'success', 3000);
      }
    } catch (e) {
      console.error('Update check error:', e);
      const msg = String(e?.message || e || '');
      if (msg.includes('404') || msg.includes('Not Found')) {
        addToast('Kein Update-Server gefunden (noch kein Release veröffentlicht?).', 'error');
      } else {
        addToast(`Update-Prüfung fehlgeschlagen: ${msg || 'Unbekannter Fehler'}`, 'error');
      }
    }
    finally { setIsCheckingUpdate(false); }
  }, [addToast]);

  const handleInstallUpdate = useCallback(async () => {
    if (!pendingUpdate) return;
    setIsInstalling(true);
    setInstallProgress(0);
    try {
      let downloaded = 0;
      let total = 0;
      await pendingUpdate.downloadAndInstall(event => {
        if (event.event === 'Started') { total = event.data.contentLength ?? 0; }
        if (event.event === 'Progress') {
          downloaded += event.data.chunkLength;
          setInstallProgress(total > 0 ? Math.round((downloaded / total) * 100) : 0);
        }
        if (event.event === 'Finished') { setInstallProgress(100); }
      });
      addToast('Update installiert — App wird neu gestartet…', 'success', 0);
      setPendingUpdate(null);
    } catch { addToast('Update-Installation fehlgeschlagen.', 'error'); setIsInstalling(false); }
  }, [pendingUpdate, addToast]);

  // ─── Store: Laden ─────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      try {
        const ts = await store.get('lastSavedAt');
        if (ts) setLastSavedAt(new Date(ts));
        const favs = await store.get('favorites');
        if (favs) setFavorites(favs);
        const hist = await store.get('exportHistory');
        if (hist) setExportHistory(hist);
        const dm = await store.get('darkMode');
        if (dm !== null && dm !== undefined) setDarkMode(dm);
        const sizes = await store.get('classSizes');
        if (sizes) setConfig(p => ({ ...p, classSizes: sizes }));
        const names = await store.get('classNames');
        if (names) setConfig(p => ({ ...p, classNames: names }));
        const studentPwd = await store.get('studentPwd');
        if (studentPwd) setConfig(p => ({ ...p, studentPwd }));
        const trainerPwd = await store.get('trainerPwd');
        if (trainerPwd) setConfig(p => ({ ...p, trainerPwd }));
        const autoPassword = await store.get('autoPassword');
        if (autoPassword !== null && autoPassword !== undefined) setConfig(p => ({ ...p, autoPassword }));
        const pool = await store.get('coursePool');
        if (pool?.length) setCourseDictionary(pool);
      } catch (e) {
        console.error('Store laden:', e);
        addToast('Einstellungen konnten nicht geladen werden.', 'error');
      } finally {
        setIsStoreLoaded(true);
      }
    };
    load();
  }, []); // eslint-disable-line

  // ─── Store: Auto-Save ─────────────────────────────────────────────────────
  useEffect(() => {
    if (!isStoreLoaded) return;
    setSaveStatus('saving');
    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
    saveTimeoutRef.current = setTimeout(async () => {
      try {
        const now = new Date().toISOString();
        await store.set('lastSavedAt', now);
        await store.set('favorites', favorites);
        await store.set('exportHistory', exportHistory);
        await store.set('darkMode', darkMode);
        await store.set('classSizes', config.classSizes);
        await store.set('classNames', config.classNames);
        await store.set('studentPwd', config.studentPwd);
        await store.set('trainerPwd', config.trainerPwd);
        await store.set('autoPassword', config.autoPassword);
        await store.save();
        setLastSavedAt(new Date(now));
        setSaveStatus('saved');
        setTimeout(() => setSaveStatus('idle'), 2500);
      } catch {
        setSaveStatus('error');
        addToast('Speichern fehlgeschlagen.', 'error');
      }
    }, 600);
    return () => clearTimeout(saveTimeoutRef.current);
  }, [favorites, exportHistory, darkMode, config.classSizes, config.classNames, config.studentPwd, config.trainerPwd, config.autoPassword, isStoreLoaded]); // eslint-disable-line

  // ─── Kurs-Pool ────────────────────────────────────────────────────────────
  const fetchCoursePool = useCallback(async () => {
    setIsLoadingPool(true);
    try {
      const r = await fetch(COURSE_API_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ request: 'get_courses' }) });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const raw = await r.json();
      const items = Array.isArray(raw) ? raw : Array.isArray(raw?.value) ? raw.value : [];
      const normalized = items.map((item, i) => {
        const rawId = findValueByPattern(item, ['id', 'guid', 'key', 'ident']) || `c-${i}`;
        const label = findValueByPattern(item, ['label', 'name', 'titel', 'title', 'bezeichnung', 'kurs']) || 'Unbenannter Kurs';
        let sh = findValueByPattern(item, ['kurzel', 'kuerzel', 'kürzel', 'shorthand', 'short', 'code', 'kennung', 'abbr']);
        if (!sh || isGuid(String(sh))) {
          sh = String(label).replace(/[^a-zA-Z0-9 ]/g, '').split(' ').filter(w => w).map(w => w[0]).join('').toUpperCase();
          if (sh.length < 2) sh = String(label).substring(0, 3).toUpperCase();
        }
        const url = findValueByPattern(item, ['url', 'link', 'hyperlink', 'moodle', 'portal']) || '';
        return { id: String(rawId), label: String(label).trim(), shorthand: String(sh).trim(), url: String(url).trim() };
      });
      setCourseDictionary(normalized);
      if (normalized.length > 0) {
        setConfig(prev => {
          const ids = [...prev.selectedPoolCourseIds];
          normalized.slice(0, prev.courseSlotCount).forEach((c, i) => { if (ids[i] === 'none') ids[i] = c.id; });
          return { ...prev, selectedPoolCourseIds: ids };
        });
        store.set('coursePool', normalized).then(() => store.save()).catch(() => {});
      }
      addToast(`${normalized.length} Kurse geladen.`, 'success', 2500);
    } catch { addToast('Verbindung zum Kurs-Pool fehlgeschlagen — verwende Cache.', 'error'); }
    finally { setIsLoadingPool(false); }
  }, [addToast]);
  useEffect(() => { fetchCoursePool(); }, [fetchCoursePool]);

  // ─── Berechnungen ─────────────────────────────────────────────────────────
  const activeMatrixCourses = useMemo(() =>
    config.selectedPoolCourseIds.slice(0, config.courseSlotCount)
      .map(id => courseDictionary.find(c => String(c.id) === String(id)))
      .filter(c => c && c.id !== 'none'),
    [config.selectedPoolCourseIds, config.courseSlotCount, courseDictionary]);

  const totals = useMemo(() => {
    let std = 0;
    Object.entries(config.classCounts).forEach(([i, n]) => { std += (config.classSizes[parseInt(i)] || 0) * n; });
    return { cls: Object.values(config.classCounts).reduce((a, b) => a + b, 0), std, trainers: config.trainerCount, all: std + config.trainerCount };
  }, [config.classCounts, config.classSizes, config.trainerCount]);

  const classRows = useMemo(() => {
    const rows = []; let id = 1;
    [0, 1, 2, 3].forEach(idx => { for (let i = 0; i < config.classCounts[idx]; i++) rows.push({ id: id++, size: config.classSizes[idx], typeIdx: idx }); });
    return rows;
  }, [config.classCounts, config.classSizes]);

  const rawEndDate = useMemo(() => { const d = new Date(config.enrolDate); d.setDate(d.getDate() + parseInt(config.enrolPeriod || 0)); return d; }, [config.enrolDate, config.enrolPeriod]);
  const endDateDisplay = rawEndDate.toISOString().split('T')[0];
  const endDateFormatted = rawEndDate.toLocaleDateString('de-DE');

  const totalStudents = useMemo(() => classRows.reduce((s, r) => s + r.size, 0), [classRows]);

  // false | 'dauer' | 'abgelaufen'
  const isEnrolInvalid = useMemo(() => {
    const period = parseInt(config.enrolPeriod, 10);
    if (!period || period <= 0) return 'dauer';
    const end = new Date(config.enrolDate);
    end.setDate(end.getDate() + period);
    return end < new Date(new Date().toDateString()) ? 'abgelaufen' : false;
  }, [config.enrolDate, config.enrolPeriod]);

  const getClassLabel = useCallback(row => {
    const n = config.classNames?.[row.id - 1]?.trim();
    return n || `Klasse-${String(row.id).padStart(2, '0')}`;
  }, [config.classNames]);

  // ─── Handler ──────────────────────────────────────────────────────────────
  const handleInput = useCallback(e => {
    const { name, value } = e.target;
    const isNum = ['enrolPeriod', 'trainerCount', 'courseSlotCount'].includes(name);
    const processed = name === 'institute' ? value.replace(/ /g, '_') : value;
    setConfig(p => ({ ...p, [name]: isNum ? Math.max(0, parseInt(value, 10) || 0) : processed }));
  }, []);
  const handleEndDateInput = useCallback(e => {
    if (!e.target.value) return;
    const diff = Math.ceil((new Date(e.target.value) - new Date(config.enrolDate)) / 86400000);
    setConfig(p => ({ ...p, enrolPeriod: Math.max(0, diff) }));
  }, [config.enrolDate]);
  const updateClassSize = useCallback((idx, val) => setConfig(p => { const s = [...p.classSizes]; s[idx] = Math.max(0, parseInt(val, 10) || 0); return { ...p, classSizes: s }; }), []);
  const updateClassCount = useCallback((idx, val) => setConfig(p => ({ ...p, classCounts: { ...p.classCounts, [idx]: Math.max(0, parseInt(val, 10) || 0) } })), []);
  const updateClassName = useCallback((rowIndex, val) => setConfig(p => ({ ...p, classNames: { ...p.classNames, [rowIndex]: val } })), []);
  const toggleCourseAssignment = useCallback((classId, courseId) => {
    const sid = String(courseId);
    setClassMatrix(prev => { const cur = (prev[classId] || []).map(String); return { ...prev, [classId]: cur.includes(sid) ? cur.filter(x => x !== sid) : [...cur, sid] }; });
    // Validierung wird nur beim Generieren neu gesetzt — hier nur entfernen wenn jetzt mind. 1 aktiver Kurs
    // (wird im useEffect unten reaktiv nachgeführt)
  }, []);

  // Validierung reaktiv nachführen: wenn invalidClassIds gesetzt sind, live aktualisieren
  useEffect(() => {
    if (invalidClassIds.size === 0) return;
    const activeIds = config.selectedPoolCourseIds.slice(0, config.courseSlotCount).filter(id => id !== 'none').map(String);
    const stillBad = new Set();
    classRows.forEach(r => {
      const assigned = (classMatrix[r.id] || []).map(String);
      if (!assigned.some(id => activeIds.includes(id))) stillBad.add(r.id);
    });
    setInvalidClassIds(stillBad);
  }, [classMatrix, classRows, config.selectedPoolCourseIds, config.courseSlotCount]); // eslint-disable-line
  const updateCourseSlot = useCallback((i, id) => setConfig(prev => { const ids = [...prev.selectedPoolCourseIds]; ids[i] = id; return { ...prev, selectedPoolCourseIds: ids }; }), []);

  // ─── Alle Zuweisen ────────────────────────────────────────────────────────
  const assignAll = useCallback(() => {
    const activeIds = activeMatrixCourses.map(c => String(c.id));
    if (!activeIds.length) return addToast('Keine aktiven Kurse zum Zuweisen.', 'error');
    setClassMatrix(prev => {
      const next = { ...prev };
      classRows.forEach(r => { next[r.id] = [...new Set([...(next[r.id] || []).map(String), ...activeIds])]; });
      return next;
    });
    setInvalidClassIds(new Set());
    addToast(`Alle ${classRows.length} Klassen mit ${activeIds.length} Kurs(en) belegt.`, 'success');
  }, [activeMatrixCourses, classRows, addToast]);

  // ─── Reset ────────────────────────────────────────────────────────────────
  const handleSettingsReset = useCallback(() => {
    setConfig(p => ({ ...p, classSizes: [...DEFAULT_CONFIG.classSizes], classNames: {} }));
    setClassNamesResetKey(k => k + 1);
    addToast('Einstellungen zurückgesetzt.', 'success');
  }, [addToast]);

  const handleFullReset = useCallback(async () => {
    setConfig(DEFAULT_CONFIG); setClassMatrix({}); setGeneratedData([]); setIsGenerated(false); setInvalidClassIds(new Set());
    setFavorites([]); setExportHistory([]);
    setShowDeleteConfirm(false);
    try { await store.clear(); await store.save(); addToast('Alle Daten gelöscht.', 'success'); }
    catch { addToast('Reset fehlgeschlagen.', 'error'); }
  }, [addToast]);

  const handleSessionReset = useCallback(() => {
    setConfig(p => ({
      ...p,
      institute: '',
      enrolDate: new Date().toISOString().split('T')[0],
      selectedPoolCourseIds: Array(8).fill('none'),
    }));
    setClassMatrix({});
    setGeneratedData([]);
    setIsGenerated(false);
    setInvalidClassIds(new Set());
    addToast('Neue Liste gestartet.', 'success');
  }, [addToast]);

  // ─── Favoriten ────────────────────────────────────────────────────────────
  const saveFavorite = useCallback(() => {
    if (!config.institute?.trim()) return addToast('Bitte zuerst Institutsnamen eingeben.', 'error');
    const fav = { id: Date.now().toString(), name: config.institute.trim(), config: { ...config }, matrix: { ...classMatrix }, savedAt: new Date().toISOString() };
    setFavorites(prev => [fav, ...prev.filter(f => f.name !== fav.name)]);
    addToast(`"${fav.name}" als Favorit gespeichert.`, 'success');
  }, [config, classMatrix, addToast]);

  const loadFavorite = useCallback(fav => {
    setConfig({ ...DEFAULT_CONFIG, ...fav.config, enrolDate: new Date().toISOString().split('T')[0] });
    setClassMatrix(fav.matrix || {});
    setIsGenerated(false); setInvalidClassIds(new Set());
    addToast(`"${fav.name}" geladen.`, 'success');
    setActiveModal(null);
  }, [addToast]);

  const deleteFavorite = useCallback(id => setFavorites(prev => prev.filter(f => f.id !== id)), []);

  // ─── Export-History ───────────────────────────────────────────────────────
  const addExportEntry = useCallback((type, filename) => {
    setExportHistory(prev => [{
      id: Date.now().toString(), type, institute: config.institute,
      accounts: generatedData.length, trainers: generatedData.filter(d => d.isT).length,
      students: generatedData.filter(d => !d.isT).length, date: new Date().toISOString(), filename,
    }, ...prev]);
  }, [config.institute, generatedData]);

  // ─── Generierung ──────────────────────────────────────────────────────────
  const generateList = useCallback(() => {
    if (!config.institute?.trim()) return addToast('Bitte Institutsnamen eingeben.', 'error');
    if (!classRows.length && !config.trainerCount) return addToast('Keine Klassen oder Trainer.', 'error');
    const activeIds = activeMatrixCourses.map(c => String(c.id));
    const badIds = new Set();
    classRows.forEach(r => {
      const assigned = (classMatrix[r.id] || []).map(String);
      if (!assigned.some(id => activeIds.includes(id))) badIds.add(r.id);
    });
    if (badIds.size) { setInvalidClassIds(badIds); return addToast(`${badIds.size} Klasse(n) ohne Kurszuweisung — rot markiert.`, 'error'); }
    setInvalidClassIds(new Set());
    const instClean = config.institute.replace(/\s+/g, '').toLowerCase();
    const data = [];
    for (let t = 1; t <= config.trainerCount; t++)
      data.push({ cNum: 'ALL', isT: true, first: 'Trainer', last: config.institute, user: `${instClean}-trainer-${t}`, mail: `trainer${t}@${instClean}.com`, pw: config.autoPassword ? generatePassword() : config.trainerPwd, courses: activeMatrixCourses });
    let sIdx = 1;
    classRows.forEach(r => {
      const selIds = (classMatrix[r.id] || []).map(String);
      const selCourses = courseDictionary.filter(cd => selIds.includes(String(cd.id)) && activeIds.includes(String(cd.id)));
      const classLabel = `${config.institute}-${getClassLabel(r)}`;
      for (let i = 0; i < r.size; i++) {
        const id = String(sIdx++).padStart(3, '0');
        data.push({ cNum: String(r.id).padStart(2, '0'), cLabel: classLabel, isT: false, first: 'Schüler', last: config.institute, user: `${instClean}-student-${id}`, mail: `student${id}@${instClean}.com`, pw: config.autoPassword ? generatePassword() : config.studentPwd, courses: selCourses });
      }
    });
    setGeneratedData(data); setIsGenerated(true); setActiveModal('dataPreview');
    addToast(`${data.length} Accounts generiert.`, 'success');
  }, [config, classRows, classMatrix, activeMatrixCourses, courseDictionary, getClassLabel, addToast]);

  // ─── CSV ──────────────────────────────────────────────────────────────────
  const downloadCSV = useCallback(() => {
    if (!generatedData.length) return;
    const rows = generatedData.map(r => {
      const enrols = [];
      if (r.isT) activeMatrixCourses.forEach(c => classRows.forEach(cls => enrols.push({ shorthand: c.shorthand, group: `${config.institute}-${getClassLabel(cls)}`, role: 4, period: config.enrolPeriod })));
      else r.courses.forEach(c => enrols.push({ shorthand: c.shorthand, group: r.cLabel, role: 5, period: config.enrolPeriod }));
      return { ...r, enrols };
    });
    const maxC = Math.max(...rows.map(r => r.enrols.length), 1);
    const headers = ['username', 'firstname', 'lastname', 'email', 'password', 'cohort1'];
    for (let i = 1; i <= maxC; i++) headers.push(`course${i}`, `group${i}`, `role${i}`, `enrolperiod${i}`);
    const lines = rows.map(r => {
      const line = [esc(r.user), esc(r.first), esc(r.last), esc(r.mail), esc(r.pw), esc(r.last)];
      for (let i = 0; i < maxC; i++) { const e = r.enrols[i]; line.push(e ? esc(e.shorthand) : '""', e ? esc(e.group) : '""', e ? esc(e.role) : '""', e ? esc(e.period) : '""'); }
      return line.join(',');
    });
    const fname = `EBCL-Moodle-Upload-${config.institute.replace(/\s+/g, '_')}-${new Date().toISOString().split('T')[0]}.csv`;
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(new Blob(['\uFEFF', headers.join(',') + '\r\n' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' })), download: fname });
    a.click(); URL.revokeObjectURL(a.href);
    addExportEntry('CSV', fname); addToast('CSV heruntergeladen.', 'success');
  }, [generatedData, activeMatrixCourses, classRows, config, getClassLabel, addToast, addExportEntry]);

  // ─── PDF ──────────────────────────────────────────────────────────────────
  const downloadPDF = useCallback(async () => {
    setIsExportingPDF(true);
    try {
      const qrDataUrl = await getQrDataUrl().catch(() => null);
      const doc = new jsPDF('l', 'mm', 'a4');
      const primary = [157, 32, 43];
      const fname = `EBCL-Zugangsdaten-${config.institute.replace(/\s+/g, '_')}-${new Date().toISOString().split('T')[0]}.pdf`;
      const periodStr = `${new Date(config.enrolDate).toLocaleDateString('de-DE')} – ${endDateFormatted}`;
      
      const renderHeader = (title, info) => {
        doc.setFontSize(20).setTextColor(...primary).setFont('helvetica', 'bold').text(config.institute.toUpperCase(), 15, 20);
        doc.setFontSize(12).setTextColor(60).setFont('helvetica', 'normal').text(title, 15, 28);
        doc.setFontSize(10).setTextColor(37, 99, 235).text('Lernplattform: https://world.ebcl.eu/', 15, 34).link(15, 31, 65, 4, { url: 'https://world.ebcl.eu/' });
        doc.setFontSize(8).setTextColor(120).text(`ZEITRAUM: ${periodStr} | ${info}`, 15, 42);
        doc.setLineWidth(0.3).setDrawColor(0).line(15, 45, 282, 45);
        if (qrDataUrl) doc.addImage(qrDataUrl, 'PNG', 256, 4, 26, 26);
      };

      const pageMap = {};

      // ─── Leitfaden-Seite (Seite 1) ──────────────────────────────────────────
      renderHeader('LEITFADEN FÜR TRAINER', `ZEITRAUM: ${periodStr}`);
      pageMap[doc.internal.getCurrentPageInfo().pageNumber] = 'Leitfaden für Trainer';
      {
        const col1 = 15, col2 = 152;
        const bodyColor = [50, 50, 50];
        const drawSection = (num, title, lines, x, y0) => {
          doc.setFontSize(9.5).setFont('helvetica', 'bold').setTextColor(...primary);
          doc.text(`${num}   ${title}`, x, y0);
          doc.setFontSize(7.5).setFont('helvetica', 'normal').setTextColor(...bodyColor);
          let y = y0 + 6;
          lines.forEach(([text, isUrl]) => {
            if (text === '') { y += 3; return; }
            if (isUrl) {
              doc.setTextColor(37, 99, 235);
              doc.text(text, x + 2, y);
              doc.link(x + 2, y - 3, 65, 4, { url: text.trim() });
              doc.setTextColor(...bodyColor);
            } else {
              doc.text(text, x + 2, y);
            }
            y += 5;
          });
        };

        drawSection(1, 'ANMELDUNG AUF DER LERNPLATTFORM', [
          ['Die EBCL-Lernplattform ist unter folgender Adresse erreichbar:'],
          [''],
          ['https://world.ebcl.eu/', true],
          [''],
          ['So melden Sie sich an:'],
          ['  1.  Browser öffnen und https://world.ebcl.eu/ aufrufen'],
          ['  2.  Oben rechts auf „Anmelden" klicken'],
          ['  3.  Benutzername & Passwort eingeben'],
          ['       (Zugangsdaten finden Sie auf der nächsten Seite)'],
          ['  4.  Auf „Anmelden" klicken'],
        ], col1, 53);

        drawSection(2, 'MOODLE – MEINE KURSE', [
          ['Nach der Anmeldung erscheint das Dashboard.'],
          [''],
          ['Unter „Meine Kurse" finden Sie alle zugewiesenen'],
          ['Kurse. Klicken Sie auf einen Kurstitel, um Lern-'],
          ['inhalte, Aufgaben und Aktivitäten einzusehen.'],
          [''],
          ['Navigation innerhalb eines Kurses:'],
          ['  •  Linkes Menü: Alle Kursabschnitte'],
          ['  •  Kursstartseite: Übersicht aller Aktivitäten'],
          ['  •  Ankündigungen & Neuigkeiten im Newsforum'],
        ], col1, 115);

        drawSection(3, 'GRUPPEN & KLASSEN VERWALTEN', [
          ['Ihre Schüler sind einer Gruppe (Klasse) zugewiesen.'],
          [''],
          ['Gruppenmitglieder einsehen:'],
          ['  1.  Gewünschten Kurs öffnen'],
          ['  2.  Im Kursmenü auf „Teilnehmer" klicken'],
          ['  3.  Oben die Ansicht „Gruppen" auswählen'],
          ['  4.  Auf eine Gruppe klicken'],
        ], col2, 53);

        drawSection(4, 'BEWERTUNGEN DER SCHÜLER EINSEHEN', [
          ['So sehen Sie die Leistungen Ihrer Schüler:'],
          [''],
          ['  1.  Gewünschten Kurs öffnen'],
          ['  2.  Im Kursmenü auf „Bewertungen" klicken'],
          ['  3.  Bewertungsübersicht zeigt alle Aktivitäten'],
          ['       und die erzielten Punkte der Schüler'],
        ], col2, 100);

        drawSection(5, 'SCHÜLER – PROFIL & NAME ÄNDERN', [
          ['Schüler können ihren Namen selbst anpassen:'],
          ['  1.  Oben rechts auf das Profilbild klicken'],
          ['  2.  Im Dropdown „Profil" wählen'],
          ['  3.  Links unter dem Namen auf "Profil bearbeiten" klicken'],
          ['  4.  Vor- und Nachname eintragen, speichern'],
        ], col2, 140);

        // Hinweis-Box
        doc.setFillColor(255, 248, 230).setDrawColor(...primary).setLineWidth(0.4);
        doc.roundedRect(col1, 172, 267, 12, 2, 2, 'FD');
        doc.setFontSize(7.5).setFont('helvetica', 'bold').setTextColor(...primary);
        doc.text('HINWEIS:', col1 + 4, 179);
        doc.setFont('helvetica', 'normal').setTextColor(...bodyColor);
        doc.text('Ihre persönlichen Zugangsdaten (Benutzername & Passwort) befinden sich auf der nächsten Seite dieses Dokuments.', col1 + 24, 179);
      }

      const tOpts = (courses, sectionLabel) => ({
        startY: 50, 
        theme: 'grid', 
        styles: { fontSize: 7, textColor: [0, 0, 0], cellWidth: 'wrap' },
        headStyles: { fillColor: [40, 40, 40], textColor: 255 },
        columnStyles: {
          0: { cellWidth: 40 },
          ...Object.fromEntries(courses.map((_, i) => [i + 3, { cellWidth: 'wrap', textColor: primary, fontStyle: 'bold' }]))
        },
        didDrawCell: (data) => {
          const { cell, column, section, row } = data;
          
          // 1. Link-Logik für Kurse
          if (section === 'body' && column.index >= 3 && courses[column.index - 3]?.url) {
            doc.link(cell.x, cell.y, cell.width, cell.height, { url: courses[column.index - 3].url });
          }

          // 2. NEU: Formularfeld in der "Name"-Spalte (Index 0)
          if (section === 'body' && column.index === 0) {
            const textField = new doc.AcroFormTextField();
            textField.fontSize = 7;
            // Eindeutiger Name für das Feld (Kombination aus Seite und Zeile)
            textField.fieldName = `name_${sectionLabel.replace(/[^a-z0-9]/gi, '_')}_${row.index}`;
            // Position auf der Zelle (mit kleinem Padding)
            textField.Rect = [cell.x + 1, cell.y + 1, cell.width - 2, cell.height - 2];
            doc.addField(textField);
          }
        },
        didDrawPage: () => { 
          pageMap[doc.internal.getCurrentPageInfo().pageNumber] = sectionLabel; 
        },
      });

      const trainers = generatedData.filter(d => d.isT);
      if (trainers.length) {
        doc.addPage();
        renderHeader('Zugangsdaten: Trainer', `TRAINER: ${trainers.length}`);
        autoTable(doc, { 
          head: [['Name (hier tippen)', 'Username', 'Passwort', ...activeMatrixCourses.map((_, i) => `Kurs ${i + 1}`)]], 
          body: trainers.map(t => ['', t.user, t.pw, ...t.courses.map(c => c.label)]), 
          ...tOpts(activeMatrixCourses, `Trainer — ${config.institute}`), 
          didParseCell: d => { if (d.section === 'body') d.cell.styles.fillColor = [255, 255, 245]; } 
        });
      }

      [...new Set(generatedData.filter(d => !d.isT).map(d => d.cNum))].sort().forEach((id) => {
        doc.addPage();
        const students = generatedData.filter(d => d.cNum === id);
        const row = classRows.find(r => String(r.id).padStart(2, '0') === id);
        const classLabel = row ? getClassLabel(row) : `Klasse-${id}`;
        renderHeader(`Teilnehmerliste: ${classLabel}`, `SCHÜLER: ${students.length}`);
        
        autoTable(doc, { 
          head: [['Name (hier tippen)', 'Username', 'Passwort', ...students[0].courses.map((_, i) => `Kurs ${i + 1}`)]], 
          body: students.map(s => ['', s.user, s.pw, ...s.courses.map(c => c.label)]), 
          ...tOpts(students[0].courses, `${classLabel} — ${config.institute}`) 
        });
      });

      const pc = doc.internal.getNumberOfPages();
      for (let i = 1; i <= pc; i++) {
        doc.setPage(i);
        const sectionLabel = pageMap[i] || '';
        doc.setLineWidth(0.2).setDrawColor(180).line(15, 193, 282, 193);
        doc.setFont('helvetica', 'normal').setFontSize(6.5).setTextColor(100);
        doc.text(sectionLabel, 15, 197);
        doc.setTextColor(37, 99, 235);
        doc.text('world.ebcl.eu', 148.5, 197, { align: 'center' });
        doc.setTextColor(100);
        doc.text(`Zeitraum: ${periodStr}`, 282, 197, { align: 'right' });
        doc.setFontSize(6).setTextColor(160);
        doc.text(fname, 15, 201.5);
        doc.text(`Seite ${i} von ${pc}`, 282, 201.5, { align: 'right' });
      }

      doc.save(fname);
      addExportEntry('PDF', fname); 
      addToast('PDF exportiert (interaktiv).', 'success');
    } catch (e) { 
      console.error(e); 
      addToast('PDF-Export fehlgeschlagen.', 'error'); 
    } finally { 
      setIsExportingPDF(false); 
    }
  }, [generatedData, config, activeMatrixCourses, classRows, getClassLabel, endDateFormatted, addToast, addExportEntry]);

  // ─── Shortcuts ────────────────────────────────────────────────────────────
  generateRef.current = generateList; csvRef.current = downloadCSV; pdfRef.current = downloadPDF; assignRef.current = assignAll;
  useEffect(() => {
    const h = e => {
      // Niemals in Textfeldern auslösen
      const tag = e.target?.tagName?.toLowerCase();
      if (tag === 'input' || tag === 'textarea' || tag === 'select' || e.target?.isContentEditable) {
        if (e.key === 'Escape') setActiveModal(null);
        return;
      }
      const mod = e.metaKey || e.ctrlKey;
      if (e.key === 'Escape') { setActiveModal(null); return; }
      if (!mod) return;
      if (e.key === 'g') { e.preventDefault(); generateRef.current?.(); }
      if (e.key === 'e') { e.preventDefault(); csvRef.current?.(); }
      if (e.key === 'p') { e.preventDefault(); pdfRef.current?.(); }
      if (e.key === 'a' && e.shiftKey) { e.preventDefault(); assignRef.current?.(); }
      if (e.key === '?') { e.preventDefault(); setActiveModal('help'); }
    };
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  }, []);

  // ─── Save Badge ───────────────────────────────────────────────────────────
  const SaveBadge = () => {
    if (saveStatus === 'idle') return null;
    const map = {
      saving: { icon: <Loader2 size={10} className="animate-spin" />, text: 'Speichert…', color: 'text-slate-400' },
      saved: { icon: <Check size={10} />, text: 'Gespeichert', color: 'text-emerald-500' },
      error: { icon: <AlertTriangle size={10} />, text: 'Fehler', color: 'text-rose-500' },
    };
    const { icon, text, color } = map[saveStatus] || {};
    return <span className={`flex items-center gap-1 text-[9px] font-semibold uppercase tracking-widest ${color}`}>{icon}{text}</span>;
  };

  // ─── INPUT STYLE HELPER ───────────────────────────────────────────────────
  const inp = (extra = '') => ({ style: { backgroundColor: C.card, borderColor: C.border, color: C.text }, className: `w-full px-2.5 py-1.5 border rounded-lg text-xs font-medium outline-none shadow-sm focus:border-blue-300 transition-colors ${extra}` });

  // ═══════════════════════════════════════════════════════════════════════════
  // MODALS
  // ═══════════════════════════════════════════════════════════════════════════

  const renderHelpModal = () => {
    const TABS = [
      { id: 'workflow',  label: 'Workflow',      icon: <Zap size={13} /> },
      { id: 'config',   label: 'Konfiguration', icon: <Settings size={13} /> },
      { id: 'matrix',   label: 'Kurs-Matrix',   icon: <TableIcon size={13} /> },
      { id: 'export',   label: 'Export',         icon: <FileDown size={13} /> },
      { id: 'shortcuts',label: 'Shortcuts',      icon: <Keyboard size={13} /> },
    ];
    const HSection = ({ icon, title, children }) => (
      <div className="mb-5">
        <div className="flex items-center gap-2 mb-2.5">
          <span style={{ color: C.main, backgroundColor: C.main + '18' }} className="p-1 rounded-md">{icon}</span>
          <h4 style={{ color: C.text }} className="text-[11px] font-bold uppercase tracking-widest">{title}</h4>
        </div>
        <div style={{ borderColor: C.border }} className="border-l-2 pl-4 space-y-2">{children}</div>
      </div>
    );
    const Info = ({ children }) => <p style={{ color: C.muted }} className="text-xs leading-relaxed">{children}</p>;
    const Row = ({ label, desc }) => (
      <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="flex gap-3 px-3 py-2.5 rounded-xl border">
        <span style={{ color: C.text }} className="text-[11px] font-semibold w-32 shrink-0">{label}</span>
        <span style={{ color: C.muted }} className="text-[11px]">{desc}</span>
      </div>
    );
    return (
      <ModalShell C={C} maxW="max-w-2xl" zIndex={250}>
        {/* Header */}
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="px-6 pt-6 pb-0 border-b shrink-0">
          <div className="flex justify-between items-center mb-4">
            <div className="flex items-center gap-3">
              <div style={{ backgroundColor: C.main }} className="p-2.5 text-white rounded-xl shadow-md"><BookOpen size={18} /></div>
              <div>
                <h3 style={{ color: C.text }} className="text-base font-bold uppercase tracking-tight">Hilfe & Dokumentation</h3>
                <p style={{ color: C.muted }} className="text-[10px] font-semibold uppercase tracking-widest">EBCL-Moodle {appVersion}</p>
              </div>
            </div>
            <button onClick={() => setActiveModal(null)} style={{ color: C.muted }} className="p-2 hover:bg-black/10 rounded-full transition-all"><X size={20} /></button>
          </div>
          {/* Tabs */}
          <div className="flex gap-1">
            {TABS.map(t => (
              <button key={t.id} onClick={() => setHelpTab(t.id)}
                style={helpTab === t.id
                  ? { color: C.main, borderColor: C.main, backgroundColor: C.card }
                  : { color: C.muted, borderColor: 'transparent', backgroundColor: 'transparent' }}
                className="flex items-center gap-1.5 px-3 py-2 text-[10px] font-bold uppercase tracking-widest border-b-2 transition-all hover:opacity-80 rounded-t-lg">
                {t.icon}{t.label}
              </button>
            ))}
          </div>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-6">

          {/* ── WORKFLOW ── */}
          {helpTab === 'workflow' && (
            <div>
              <HSection icon={<Zap size={14} />} title="Was macht dieses Programm?">
                <Info>EBCL-Moodle erstellt automatisch hunderte Moodle-Zugangsdaten für Partnerinstitute. Du konfigurierst einmal die Klassen- und Kursstruktur, generierst die Zugänge und exportierst sie — fertig.</Info>
                <Info>Die Kursliste wird automatisch vom EBCL-Server geladen und lokal gecacht, sodass die App auch offline funktioniert.</Info>
              </HSection>
              <HSection icon={<Users size={14} />} title="Schritt-für-Schritt Workflow">
                {[
                  ['01', C.accent1, 'Institutsname eingeben', 'Den Namen des Partnerinstituts im Feld oben links eintragen. Pflichtfeld — ohne Namen kann nicht generiert werden.'],
                  ['02', C.accent2, 'Klassen & Trainer definieren', 'Anzahl der Klassen pro Typ (15/20/30/40 Plätze) und Traineranzahl festlegen. Klassengrößen sind im Settings-Panel anpassbar.'],
                  ['03', C.main,   'Kurse zuweisen', 'In der Kurs-Matrix die gewünschten Kurse pro Spalte auswählen und jeder Klasse zuweisen. "Alle zuweisen" (⌘A) belegt alle Klassen auf einmal.'],
                  ['04', '#7C3AED','Liste generieren', 'Auf "Liste generieren" klicken (⌘G). Die App prüft ob alle Klassen mindestens einen Kurs haben — fehlende werden rot markiert.'],
                  ['05', '#B45309','Exportieren', 'CSV (⌘E) für den Moodle-Import, PDF (⌘P) als Zugangsdaten-Blatt für das Institut. Beide Formate enthalten alle Accounts.'],
                  ['06', '#059669','Favorit speichern', 'Das Institut per ⭐-Button als Favorit speichern. Beim nächsten Mal ist die komplette Konfiguration per Klick wiederhergestellt.'],
                ].map(([n, col, title, desc]) => (
                  <div key={n} className="flex gap-3 items-start">
                    <span style={{ color: col, backgroundColor: col + '18', borderColor: col + '30' }} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px] border mt-0.5">{n}</span>
                    <div><p style={{ color: C.text }} className="text-[11px] font-semibold mb-0.5">{title}</p><p style={{ color: C.muted }} className="text-[11px] leading-relaxed">{desc}</p></div>
                  </div>
                ))}
              </HSection>
              <HSection icon={<Info size={14} />} title="Was wird beim Neustart zurückgesetzt?">
                <Info>Bei jedem App-Start wird alles auf Standard zurückgesetzt: Institutsname, Datum, Kursauswahl und Klassenzuweisungen. <strong style={{ color: C.text }}>Erhalten bleiben:</strong> Favoriten, Export-History, Klassengrößen, Klassen-Namen und Dark Mode.</Info>
              </HSection>
            </div>
          )}

          {/* ── KONFIGURATION ── */}
          {helpTab === 'config' && (
            <div className="space-y-1">
              <HSection icon={<Building2 size={14} />} title="Organisation">
                <Row label="Institutsname" desc="Pflichtfeld. Wird in Dateinamen, PDF-Header und als Präfix für alle Usernamen verwendet. (z.B. 'VHS Wien' → 'vhswien-student-001')" />
                <Row label="Trainer" desc="Anzahl der Trainer-Accounts. Trainer erhalten Zugang zu allen aktiven Kursen und werden auf einer eigenen PDF-Seite aufgeführt." />
                <Row label="Kurs Anzahl" desc="Wie viele Kursspalten in der Matrix angezeigt werden (max. 8). Entspricht der Anzahl der Moodle-Kurse, die belegt werden sollen." />
              </HSection>
              <HSection icon={<GraduationCap size={14} />} title="Klassen Struktur">
                <Row label="Typen (15/20/30/40)" desc="Jeder Typ definiert eine Klassengröße. Über 'Anzahl' rechts bestimmst du, wie viele Klassen dieses Typs erstellt werden." />
                <Row label="Klassengrößen" desc="Anpassbar in den Settings (⚙). Werden gespeichert und bleiben beim Neustart erhalten." />
                <Row label="Klassen-Namen" desc="Optional in den Settings anpassbar (z.B. 'IT-Klasse A'). Standard: K-01, K-02, … Werden ebenfalls gespeichert." />
              </HSection>
              <HSection icon={<ShieldCheck size={14} />} title="Zeit & Sicherheit">
                <Row label="Einschreibedatum" desc="Ab wann die Accounts aktiv sind. Wird immer auf das heutige Datum gesetzt — auch beim Laden eines Favoriten." />
                <Row label="Dauer (Tage)" desc="Wie lange die Einschreibung gilt. Das Enddatum wird automatisch berechnet und ist direkt editierbar." />
                <Row label="PW Schüler / Trainer" desc="Passwörter für generierte Accounts. Standard: '***ENTFERNT***' / '***ENTFERNT***'. Bei jedem Start auf Standard zurückgesetzt." />
              </HSection>
              <HSection icon={<Star size={14} />} title="Favoriten">
                <Info>Ein Favorit speichert: Institutsname, alle Konfigurationsfelder und die komplette Kursmatrix. Beim Laden wird nur das Datum auf heute gesetzt.</Info>
                <Info>Favoriten bleiben auch nach einem Neustart erhalten und können im Favoriten-Modal verwaltet (gelöscht) werden.</Info>
              </HSection>
            </div>
          )}

          {/* ── KURS-MATRIX ── */}
          {helpTab === 'matrix' && (
            <div>
              <HSection icon={<BookOpen size={14} />} title="Kurs-Pool">
                <Info>Die verfügbaren Kurse werden beim Start automatisch vom EBCL-Server geladen. Der zuletzt geladene Stand wird lokal gecacht — die App funktioniert damit auch offline.</Info>
                <Row label="Kurse laden" desc="Über den 'Kurse laden'-Button in der Übersicht wird der Pool manuell neu synchronisiert (z.B. wenn neue Kurse hinzugekommen sind)." />
                <Row label="Cache-Status" desc="In den Settings siehst du wie viele Kurse aktuell gecacht sind." />
              </HSection>
              <HSection icon={<TableIcon size={14} />} title="Matrix-Bedienung">
                <Row label="Kurs auswählen" desc="Im Dropdown oben in jeder Spalte einen Kurs aus dem Pool wählen. Bis zu 8 Kurse gleichzeitig möglich (je nach 'Kurs Anzahl'-Einstellung)." />
                <Row label="Zuweisung (Checkboxen)" desc="Für jede Klasse per Checkbox bestimmen, welche Kurse sie erhält. Eine Klasse kann mehrere Kurse haben." />
                <Row label="Spalten-'Alle'-Button" desc="Der blaue 'Alle'-Button im Spaltenkopf weist alle Klassen diesem Kurs zu — oder entfernt alle wieder (Toggle)." />
                <Row label="Alle zuweisen (⌘A)" desc="Weist alle Klassen allen aktiven Kursen zu. Schnellster Weg wenn alle Klassen die gleichen Kurse bekommen sollen." />
              </HSection>
              <HSection icon={<AlertTriangle size={14} />} title="Validierung">
                <Info>Beim Klick auf "Liste generieren" prüft die App, ob jede Klasse mindestens einem aktiven Kurs zugewiesen ist. Klassen ohne Zuweisung werden <span style={{ color: '#e11d48' }}>rot markiert</span>. Die roten Markierungen verschwinden sobald eine Zuweisung gesetzt wird.</Info>
              </HSection>
            </div>
          )}

          {/* ── EXPORT ── */}
          {helpTab === 'export' && (
            <div>
              <HSection icon={<FileSpreadsheet size={14} />} title="CSV-Export (⌘E)">
                <Info>Die CSV-Datei ist das Moodle-Import-Format. Sie enthält alle generierten Accounts und Einschreibungen und kann direkt in Moodle hochgeladen werden.</Info>
                <Row label="Inhalt" desc="Username, Passwort, Vorname, Nachname, E-Mail, Kurs-Shorthand, Kursgruppe, Rolle und Einschreibedauer pro Zeile." />
                <Row label="Dateiname" desc="EBCL-Moodle-Upload-{Institut}-{Datum}.csv" />
                <Row label="Kodierung" desc="UTF-8 mit BOM — für maximale Excel-Kompatibilität (Umlaute werden korrekt angezeigt)." />
                <Row label="Verwendung" desc="In Moodle: Administration → Nutzer → Nutzer hochladen → CSV-Datei auswählen." />
              </HSection>
              <HSection icon={<FileDown size={14} />} title="PDF-Export (⌘P)">
                <Info>Das PDF ist das Zugangsdaten-Dokument für das Institut. Es enthält pro Seite eine Klasse (oder die Trainer-Liste) mit allen Accounts im Tabellenformat.</Info>
                <Row label="Format" desc="A4 Querformat (Landscape) — optimiert für viele Kursspalten nebeneinander." />
                <Row label="Aufbau" desc="Seite 1 (optional): Trainer. Folgeseiten: Eine Seite pro Klasse mit Namen-Feld, Username, Passwort und allen zugewiesenen Kurslinks." />
                <Row label="Kurslinks" desc="Kurs-Zellen im PDF sind klickbar — direkter Link zum jeweiligen Moodle-Kurs." />
                <Row label="Dateiname" desc="EBCL-Zugangsdaten-{Institut}-{Datum}.pdf" />
              </HSection>
              <HSection icon={<ClipboardList size={14} />} title="Daten-Vorschau & History">
                <Row label="Vorschau" desc="Nach dem Generieren zeigt 'Daten-Vorschau' alle Accounts in einer scrollbaren Tabelle — zur Kontrolle vor dem Export." />
                <Row label="Export-History" desc="Jeder Export wird mit Zeitstempel, Dateiname und Accountanzahl in der History gespeichert (bleibt nach Neustart erhalten)." />
              </HSection>
            </div>
          )}

          {/* ── SHORTCUTS ── */}
          {helpTab === 'shortcuts' && (
            <div>
              <HSection icon={<Keyboard size={14} />} title="Tastaturkürzel">
                <div className="space-y-1.5">
                  {SHORTCUTS.map((s, i) => (
                    <div key={i} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="flex items-center justify-between px-3 py-2.5 rounded-xl border">
                      <span style={{ color: C.muted }} className="text-[11px]">{s.desc}</span>
                      <div className="flex items-center gap-1">
                        {s.keys.map((k, j) => (
                          <React.Fragment key={j}>
                            <kbd style={{ backgroundColor: C.card, borderColor: C.border, color: C.text }} className="px-2 py-0.5 rounded-md border text-[10px] font-mono font-bold shadow-sm">{k}</kbd>
                            {j < s.keys.length - 1 && <span style={{ color: C.muted }} className="text-[10px] mx-0.5">+</span>}
                          </React.Fragment>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </HSection>
              <HSection icon={<Zap size={14} />} title="Tipps & Tricks">
                <Row label="Offline arbeiten" desc="Der Kurs-Pool wird gecacht. Du kannst ohne Internet generieren und exportieren — solange der Cache vorhanden ist." />
                <Row label="Favoriten nutzen" desc="Speichere häufig verwendete Institute als Favorit. Das Datum wird beim Laden automatisch auf heute gesetzt." />
                <Row label="Alle zuweisen" desc="Bei gleicher Kursstruktur für alle Klassen: ⌘A spart alle Einzelklicks in der Matrix." />
                <Row label="PDF direkt prüfen" desc="Daten-Vorschau (nach Generieren) zeigt alle Accounts bevor du exportierst — Tipp- oder Konfigurationsfehler früh erkennen." />
                <Row label="Dark Mode" desc="Toggle in den Settings oder über den Mond/Sonne-Button in der Sidebar. Wird gespeichert." />
                <Row label="Fenstergröße" desc="Das Fenster ist skalierbar. Die letzte Größe und Position wird beim nächsten Start automatisch wiederhergestellt." />
              </HSection>
              <HSection icon={<RefreshCw size={14} />} title="Updates">
                <Info>Die App prüft beim Start automatisch ob eine neue Version verfügbar ist. Updates können direkt in den Settings installiert werden — kein manueller Download nötig.</Info>
              </HSection>
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-4 border-t flex justify-between items-center shrink-0">
          <div className="flex gap-1">
            {TABS.map((t, i) => (
              <button key={t.id} onClick={() => setHelpTab(TABS[i > 0 ? i - 1 : 0]?.id || t.id)}
                style={{ color: helpTab === t.id ? C.main : C.muted }}
                className="hidden" />
            ))}
            <span style={{ color: C.muted }} className="text-[10px]">
              {TABS.findIndex(t => t.id === helpTab) + 1} / {TABS.length}
            </span>
          </div>
          <div className="flex gap-2">
            {TABS.findIndex(t => t.id === helpTab) > 0 && (
              <button onClick={() => setHelpTab(TABS[TABS.findIndex(t => t.id === helpTab) - 1].id)}
                style={{ borderColor: C.border, color: C.muted, backgroundColor: C.card }}
                className="px-4 py-2 rounded-xl border text-[10px] font-bold uppercase hover:opacity-80 transition-all">← Zurück</button>
            )}
            {TABS.findIndex(t => t.id === helpTab) < TABS.length - 1 ? (
              <button onClick={() => setHelpTab(TABS[TABS.findIndex(t => t.id === helpTab) + 1].id)}
                style={{ backgroundColor: C.accent1 }}
                className="text-white px-6 py-2 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all text-[10px]">Weiter →</button>
            ) : (
              <button onClick={() => setActiveModal(null)} style={{ backgroundColor: C.main }}
                className="text-white px-8 py-2 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all text-[10px]">Alles klar!</button>
            )}
          </div>
        </div>
      </ModalShell>
    );
  };

  const renderSettingsModal = () => (
    <ModalShell C={C} maxW="max-w-lg">
      <ModalHeader C={C} icon={<Settings size={18} />} title="System-Settings" onClose={() => setActiveModal(null)} />
      <div className="flex-1 overflow-y-auto p-6 space-y-6">
        {/* Darstellung */}
        <div>
          <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-3">{darkMode ? <Moon size={14} /> : <Sun size={14} />} Darstellung</h4>
          <button onClick={() => setDarkMode(d => !d)} style={{ backgroundColor: C.subtle, borderColor: C.border, color: C.text }} className="w-full flex items-center justify-between px-4 py-3 rounded-xl border shadow-sm hover:opacity-80 transition-all">
            <span className="text-[11px] font-semibold">{darkMode ? 'Dunkelmodus aktiv' : 'Hellmodus aktiv'}</span>
            <span style={{ backgroundColor: darkMode ? C.main : C.accent1 }} className="flex items-center gap-1.5 text-white text-[10px] font-bold px-2.5 py-1 rounded-lg">
              {darkMode ? <Sun size={11} /> : <Moon size={11} />} {darkMode ? 'Hell' : 'Dunkel'}
            </span>
          </button>
        </div>
        {/* Klassengrößen */}
        <div>
          <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-1"><Users size={14} /> Klassengrößen</h4>
          <p style={{ color: C.muted }} className="text-[10px] mb-3 opacity-60">Wird gespeichert und bleibt beim Neustart erhalten.</p>
          <div className="grid grid-cols-2 gap-3">
            {config.classSizes.map((size, idx) => (
              <div key={idx} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase block mb-1">Typ {idx + 1} — {size} Plätze</label>
                <input type="number" min="0" value={size} onChange={e => updateClassSize(idx, e.target.value)} style={{ color: C.main }} className="w-full bg-transparent text-base font-bold outline-none" />
              </div>
            ))}
          </div>
        </div>
        {/* Klassen-Namen */}
        <div>
          <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-1"><Edit3 size={14} /> Klassen-Namen <span className="normal-case font-normal opacity-60">(optional)</span></h4>
          <p style={{ color: C.muted }} className="text-[10px] mb-3 opacity-60">Leer lassen = Standard-Name (K-01, K-02, …). Wird gespeichert.</p>
          {classRows.length === 0
            ? <p style={{ color: C.muted }} className="text-[11px] italic">Erst Klassen im Konfigurations-Panel anlegen.</p>
            : <div key={classNamesResetKey} className="space-y-1.5">
                {classRows.map(row => (
                  <ClassNameRow key={row.id} row={row} savedValue={config.classNames?.[row.id - 1] || ''} onUpdate={updateClassName} C={C} />
                ))}
              </div>
          }
        </div>
        {/* Update */}
        <div>
          <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-3"><RefreshCw size={14} /> Software-Update</h4>
          <button onClick={handleManualUpdateCheck} disabled={isCheckingUpdate || !!pendingUpdate}
            style={{ borderColor: C.border, color: pendingUpdate ? C.accent2 : C.text, backgroundColor: C.card }}
            className="w-full flex items-center justify-center gap-2 px-4 py-2.5 border rounded-xl text-[11px] font-bold hover:opacity-80 disabled:opacity-50 transition-all">
            {isCheckingUpdate ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />}
            {isCheckingUpdate ? 'Suche nach Updates…' : pendingUpdate ? `Update v${pendingUpdate.version} verfügbar` : 'Jetzt auf Updates prüfen'}
          </button>
        </div>
      </div>
      <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-5 border-t flex justify-between items-center shrink-0">
        <div className="flex items-center gap-3">
          <button onClick={handleSettingsReset} className="text-amber-600 hover:text-amber-700 text-[10px] font-bold uppercase flex items-center gap-1.5 hover:bg-amber-50 px-3 py-2 rounded-lg transition-all"><RefreshCw size={12} /> Zurücksetzen</button>
          {showDeleteConfirm ? (
            <div className="flex items-center gap-2">
              <span className="text-[10px] text-rose-600 font-bold">Wirklich alles löschen?</span>
              <button onClick={handleFullReset} className="bg-rose-500 text-white text-[10px] font-bold px-2.5 py-1 rounded-lg hover:bg-rose-600 transition-all">Ja</button>
              <button onClick={() => setShowDeleteConfirm(false)} style={{ borderColor: C.border, color: C.muted }} className="border text-[10px] font-bold px-2.5 py-1 rounded-lg hover:bg-black/5 transition-all">Nein</button>
            </div>
          ) : (
            <button onClick={() => setShowDeleteConfirm(true)} className="text-rose-500 hover:text-rose-700 text-[10px] font-bold uppercase flex items-center gap-1.5 hover:bg-rose-50 px-3 py-2 rounded-lg transition-all"><Trash2 size={12} /> Alle Daten löschen</button>
          )}
        </div>
        <button onClick={() => setActiveModal(null)} style={{ backgroundColor: C.accent1 }} className="text-white px-8 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all text-xs">Schließen</button>
      </div>
    </ModalShell>
  );

  const renderFavoritesModal = () => (
    <ModalShell C={C} maxW="max-w-lg">
      <ModalHeader C={C} icon={<Star size={18} />} title="Favoriten" sub={`${favorites.length} gespeicherte Institute`} onClose={() => setActiveModal(null)} iconBg="#B45309" />
      <div className="flex-1 overflow-y-auto p-4 space-y-2">
        {favorites.length === 0 ? (
          <div className="py-12 text-center">
            <StarOff size={32} style={{ color: C.muted }} className="mx-auto mb-3 opacity-30" />
            <p style={{ color: C.muted }} className="text-sm">Noch keine Favoriten gespeichert.</p>
            <p style={{ color: C.muted }} className="text-[11px] mt-1 opacity-60">Unten auf "Aktuell speichern" klicken.</p>
          </div>
        ) : favorites.map(fav => (
          <div key={fav.id} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="flex items-center gap-3 p-3 rounded-2xl border hover:border-amber-300 transition-all">
            <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0 font-bold text-sm text-white" style={{ backgroundColor: '#B45309' }}>
              {fav.name.charAt(0).toUpperCase()}
            </div>
            <div className="flex-1 min-w-0">
              <p style={{ color: C.text }} className="font-bold text-sm truncate">{fav.name}</p>
              <div style={{ color: C.muted }} className="text-[9px] mt-1 space-y-0.5">
                <p>
                  {Object.entries(fav.config?.classCounts || {}).filter(([,v])=>v>0).map(([i,v])=>`${v}×${fav.config?.classSizes?.[i]||'?'}`).join(' · ')} Schüler · {fav.config?.trainerCount || 0} Trainer · {fav.config?.courseSlotCount || 0} Kurse
                </p>
                <p className="font-mono opacity-60">PW-S: {fav.config?.studentPwd || '—'} · PW-T: {fav.config?.trainerPwd || '—'}</p>
                <p className="opacity-50">{fmtDate(new Date(fav.savedAt))}</p>
              </div>
            </div>
            <button onClick={() => loadFavorite(fav)} style={{ backgroundColor: C.accent1 }} className="text-white text-[10px] font-bold uppercase px-3 py-1.5 rounded-lg hover:brightness-110 transition-all">Laden</button>
            <button onClick={() => deleteFavorite(fav.id)} className="p-1.5 text-rose-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"><Trash2 size={14} /></button>
          </div>
        ))}
      </div>
      <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-4 border-t flex justify-between items-center shrink-0">
        <button onClick={saveFavorite} style={{ backgroundColor: '#B45309' }} className="text-white px-5 py-2.5 rounded-xl font-bold uppercase text-xs flex items-center gap-2 hover:brightness-110 active:scale-95 transition-all">
          <Star size={14} /> Aktuell speichern
        </button>
        <button onClick={() => setActiveModal(null)} style={{ backgroundColor: C.accent1 }} className="text-white px-6 py-2.5 rounded-xl font-bold uppercase text-xs hover:brightness-110 active:scale-95 transition-all">Schließen</button>
      </div>
    </ModalShell>
  );

  const renderHistoryModal = () => (
    <ModalShell C={C} maxW="max-w-3xl">
      <ModalHeader C={C} icon={<History size={18} />} title="Export-History" sub={`${exportHistory.length} Exporte gesamt`} onClose={() => setActiveModal(null)} />
      <div className="flex-1 overflow-auto">
        {exportHistory.length === 0 ? (
          <div className="py-12 text-center">
            <History size={32} style={{ color: C.muted }} className="mx-auto mb-3 opacity-30" />
            <p style={{ color: C.muted }} className="text-sm">Noch keine Exporte durchgeführt.</p>
          </div>
        ) : (
          <table className="text-[11px]" style={{ minWidth: '700px', width: '100%' }}>
            <thead style={{ backgroundColor: C.subtle, borderColor: C.border }} className="sticky top-0 border-b">
              <tr>{['Typ', 'Institut', 'Accounts', 'Datum', 'Datei'].map(h => <th key={h} style={{ color: C.muted }} className="px-4 py-3 text-left text-[9px] font-semibold uppercase tracking-widest whitespace-nowrap">{h}</th>)}</tr>
            </thead>
            <tbody>
              {exportHistory.map(e => (
                <tr key={e.id} style={{ borderColor: C.border }} className="border-b hover:bg-black/5 transition-colors">
                  <td className="px-4 py-3 whitespace-nowrap"><span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${e.type === 'PDF' ? 'bg-blue-100 text-blue-700' : 'bg-emerald-100 text-emerald-700'}`}>{e.type}</span></td>
                  <td style={{ color: C.text }} className="px-4 py-3 font-semibold whitespace-nowrap">{e.institute}</td>
                  <td style={{ color: C.muted }} className="px-4 py-3 whitespace-nowrap"><span style={{ color: C.text }} className="font-bold">{e.accounts}</span> · {e.trainers}T/{e.students}S</td>
                  <td style={{ color: C.muted }} className="px-4 py-3 tabular-nums whitespace-nowrap">{fmtDateTime(new Date(e.date))}</td>
                  <td style={{ color: C.muted }} className="px-4 py-3 font-mono text-[9px] whitespace-nowrap">{e.filename}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-4 border-t flex justify-end shrink-0">
        <button onClick={() => setActiveModal(null)} style={{ backgroundColor: C.main }} className="text-white px-6 py-2 rounded-xl font-bold uppercase text-xs hover:brightness-110 active:scale-95 transition-all">Schließen</button>
      </div>
    </ModalShell>
  );

  const renderCoursePreviewModal = () => (
    <ModalShell C={C} maxW="max-w-4xl" zIndex={110}>
      <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-5 border-b flex flex-wrap justify-between items-center gap-4 shrink-0">
        <div className="flex items-center gap-3">
          <div style={{ backgroundColor: C.accent1 }} className="p-2 text-white rounded-xl"><Eye size={18} /></div>
          <div><h3 style={{ color: C.text }} className="font-bold uppercase tracking-tight text-sm">Kursübersicht</h3><p style={{ color: C.muted }} className="text-[9px] mt-0.5">{courseDictionary.length} Kurse verfügbar</p></div>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={fetchCoursePool} disabled={isLoadingPool} style={{ color: C.accent2 }} className="flex items-center gap-2 px-3 py-1.5 bg-emerald-50 rounded-lg text-xs font-bold uppercase hover:bg-emerald-100 disabled:opacity-50 transition-colors">
            <RefreshCw size={14} className={isLoadingPool ? 'animate-spin' : ''} /> Sync
          </button>
          <button onClick={() => setActiveModal(null)} style={{ color: C.muted }} className="p-2 hover:bg-black/10 rounded-full"><X size={20} /></button>
        </div>
      </div>
      <div className="flex-1 overflow-auto" style={{ backgroundColor: C.card }}>
        <table className="w-full text-left border-collapse min-w-[600px]">
          <thead style={{ backgroundColor: C.subtle, borderColor: C.border }} className="sticky top-0 z-10 border-b text-[10px] font-semibold uppercase tracking-widest">
            <tr>{['#', 'Kursname', 'Kürzel', 'Link'].map(h => <th key={h} style={{ color: C.muted }} className="px-5 py-3">{h}</th>)}</tr>
          </thead>
          <tbody style={{ borderColor: C.border }} className="divide-y text-[11px]">
            {isLoadingPool ? <tr><td colSpan="4" className="px-5 py-8 text-center"><Loader2 size={20} style={{ color: C.muted }} className="animate-spin mx-auto" /></td></tr>
              : courseDictionary.length === 0 ? <tr><td colSpan="4" style={{ color: C.muted }} className="px-5 py-8 text-center italic">Keine Kurse gefunden.</td></tr>
              : courseDictionary.map((c, i) => (
                <tr key={c.id} className="hover:bg-black/5 transition-colors">
                  <td style={{ color: C.muted }} className="px-5 py-3 font-mono text-[10px]">{i + 1}</td>
                  <td style={{ color: C.text }} className="px-5 py-3 font-semibold">{c.label}</td>
                  <td className="px-5 py-3"><span style={{ color: C.accent1, backgroundColor: C.accent1 + '1A' }} className="font-mono px-2 py-0.5 rounded text-[10px]">{c.shorthand}</span></td>
                  <td className="px-5 py-3">{c.url ? <a href={c.url} target="_blank" rel="noreferrer" style={{ color: C.accent1 }} className="hover:underline flex items-center gap-1">Öffnen <Eye size={12} /></a> : <span style={{ color: C.muted }} className="italic">Kein Link</span>}</td>
                </tr>
              ))}
          </tbody>
        </table>
      </div>
      <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-5 border-t flex justify-end">
        <button onClick={() => setActiveModal(null)} style={{ backgroundColor: C.main }} className="text-white px-6 py-2 rounded-xl font-bold uppercase text-xs hover:brightness-110 active:scale-95 transition-all">Schließen</button>
      </div>
    </ModalShell>
  );

  const renderDataPreviewModal = () => (
    <ModalShell C={C} maxW="max-w-5xl" zIndex={110}>
      <div className="p-6 bg-slate-800 border-b flex justify-between items-center shrink-0">
        <div className="flex items-center gap-3 text-white">
          <div style={{ backgroundColor: C.accent1 }} className="p-2.5 rounded-xl"><ClipboardList size={20} /></div>
          <div><h3 className="text-base font-bold uppercase tracking-tight leading-none">Daten-Vorschau</h3><p className="text-[10px] text-slate-400 mt-0.5">{generatedData.length} Einträge · {generatedData.filter(d => d.isT).length} Trainer · {generatedData.filter(d => !d.isT).length} Schüler</p></div>
        </div>
        <button onClick={() => setActiveModal(null)} className="p-2 hover:bg-white/10 rounded-full text-slate-400 transition-colors"><X size={24} /></button>
      </div>
      <div className="flex-1 overflow-auto bg-white">
        <table className="w-full text-[11px] table-auto border-collapse text-slate-700 min-w-[700px]">
          <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm border-b font-semibold text-slate-500 uppercase tracking-widest">
            <tr><th className="px-6 py-4 text-left">Klasse / Typ</th><th className="px-6 py-4 text-left">Username</th><th className="px-6 py-4 text-left">Zuweisungen</th></tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {generatedData.map((row, i) => (
              <tr key={i} className={row.isT ? 'bg-amber-50/40 font-semibold border-l-2 border-amber-300' : 'hover:bg-slate-50/50 transition-colors'}>
                <td className="px-6 py-3"><span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${row.isT ? 'bg-amber-100 text-amber-700' : 'bg-slate-100 text-slate-500'}`}>{row.isT ? 'GLOBAL (Trainer)' : `K-${row.cNum}`}</span></td>
                <td className="px-6 py-3 font-mono font-medium text-slate-700 tracking-tight">{row.user}</td>
                <td className="px-6 py-3"><div className="flex flex-wrap gap-1.5">
                  {row.courses.map((c, ci) => <span key={ci} style={{ color: C.main, backgroundColor: C.main + '1A' }} className="font-semibold text-[9px] px-1.5 py-0.5 rounded">{c.shorthand}</span>)}
                  {!row.courses.length && <span className="text-slate-400 italic text-[10px]">Keine Kurse</span>}
                </div></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-5 border-t flex justify-end gap-3">
        <button disabled={isExportingPDF} onClick={downloadPDF} style={{ backgroundColor: C.accent1 }} className="text-white px-5 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm flex items-center gap-2 text-xs disabled:opacity-50">
          {isExportingPDF ? <Loader2 size={14} className="animate-spin" /> : <FileDown size={14} />} PDF
        </button>
        <button onClick={downloadCSV} style={{ backgroundColor: C.accent2 }} className="text-white px-5 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm flex items-center gap-2 text-xs">
          <FileSpreadsheet size={14} /> CSV
        </button>
        <button onClick={() => setActiveModal(null)} style={{ backgroundColor: C.main }} className="text-white px-8 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm text-xs">Schließen</button>
      </div>
    </ModalShell>
  );

  // ═══════════════════════════════════════════════════════════════════════════
  // MAIN RENDER
  // ═══════════════════════════════════════════════════════════════════════════
  return (
    <div style={{ backgroundColor: C.bg, color: C.text }} className="h-screen flex flex-col p-4 md:p-6 font-sans overflow-hidden relative transition-colors duration-200">

      <Toast toasts={toasts} removeToast={removeToast} />

      {activeModal === 'help' && renderHelpModal()}
      {activeModal === 'settings' && renderSettingsModal()}
      {activeModal === 'favorites' && renderFavoritesModal()}
      {activeModal === 'history' && renderHistoryModal()}
      {activeModal === 'coursePreview' && renderCoursePreviewModal()}
      {activeModal === 'dataPreview' && renderDataPreviewModal()}

      {/* UPDATE BANNER */}
      {pendingUpdate && (
        <div style={{ backgroundColor: C.accent1, borderColor: C.accent1 }} className="shrink-0 mb-3 rounded-2xl px-4 py-2.5 flex items-center justify-between gap-4 text-white shadow-lg">
          <div className="flex items-center gap-2.5 text-xs font-semibold">
            <Zap size={15} className="shrink-0" />
            <span>Update verfügbar: <span className="font-bold">v{pendingUpdate.version}</span></span>
            {isInstalling && installProgress > 0 && (
              <span className="opacity-70">({installProgress}%)</span>
            )}
          </div>
          <button
            onClick={handleInstallUpdate}
            disabled={isInstalling}
            className="shrink-0 bg-white/20 hover:bg-white/30 disabled:opacity-50 px-3 py-1.5 rounded-xl text-[11px] font-bold uppercase tracking-widest transition-all flex items-center gap-1.5 active:scale-95"
          >
            {isInstalling ? <Loader2 size={12} className="animate-spin" /> : <FileDown size={12} />}
            {isInstalling ? `Installiere… ${installProgress}%` : 'Jetzt installieren'}
          </button>
        </div>
      )}

      {/* HEADER */}
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-5 shrink-0">
        <div>
          <h1 className="text-xl font-bold flex items-center gap-2 tracking-tight">
            <TableIcon style={{ color: C.main }} size={24} /> Moodle Anmeldungen
          </h1>
          <div className="flex items-center gap-3 mt-1 flex-wrap">
            <p style={{ color: C.muted }} className="text-xs italic flex items-center gap-1.5 font-medium">
              {isLoadingPool ? <Loader2 size={12} className="animate-spin text-blue-500" /> : <CheckCircle2 style={{ color: C.accent2 }} size={12} />}
              Kurse ({courseDictionary.length})
            </p>
            <span style={{ color: C.border }}>·</span>
            <span style={{ color: C.muted }} className="flex items-center gap-1 text-[10px] font-medium">
              {isOnline ? <Wifi size={11} style={{ color: C.accent2 }} /> : <WifiOff size={11} className="text-rose-400" />}
              {isOnline ? 'Online' : 'Offline'}
            </span>
            <span style={{ color: C.border }}>·</span>
            <SaveBadge />
          </div>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => setDarkMode(d => !d)} style={{ backgroundColor: C.subtle, borderColor: C.border, color: C.muted }} className="p-2 rounded-xl border hover:opacity-80 transition-all" title="Dunkelmodus umschalten">
            {darkMode ? <Sun size={16} /> : <Moon size={16} />}
          </button>
          <div style={{ backgroundColor: C.card, borderColor: C.border, color: C.muted }} className="flex items-center gap-2 px-3 py-1.5 rounded-lg border shadow-sm text-[9px] font-bold uppercase tracking-widest">
            {appVersion} <span style={{ width: 1, backgroundColor: C.border }} className="h-3 inline-block" /> EBCL INTERNATIONAL
          </div>
        </div>
      </header>

      {/* HAUPT-GRID */}
      <div className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-5 lg:gap-6 overflow-hidden min-h-0">

        {/* SIDEBAR */}
        <aside className="lg:col-span-3 flex flex-col gap-4 overflow-y-auto pr-1 custom-scrollbar">

          {/* Status Card */}
          <section style={{ backgroundColor: C.accent1 }} className="rounded-3xl p-5 shadow-lg flex flex-col text-white shrink-0 relative overflow-hidden">
            <div className="absolute -right-6 -top-6 opacity-10"><Users size={100} /></div>
            <div className="flex justify-between items-end mb-3 border-b border-white/10 pb-2 relative z-10">
              <div><div className="text-[9px] font-semibold text-white/70 uppercase tracking-widest mb-1">Gesamt Accounts</div><div className="text-2xl font-bold leading-none">{totals.all}</div></div>
              <div className="text-[9px] text-white/50">{totals.cls} Klasse{totals.cls !== 1 ? 'n' : ''}</div>
            </div>
            <div className="space-y-2 relative z-10">
              <div className="flex justify-between items-center text-[11px] bg-white/5 px-2 py-1.5 rounded-lg"><span className="text-white/70 font-medium uppercase tracking-wider flex items-center gap-1.5"><ShieldCheck size={12} /> Trainer</span><span className="font-bold text-amber-400 text-sm">{totals.trainers}</span></div>
              <div className="flex justify-between items-center text-[11px] bg-white/5 px-2 py-1.5 rounded-lg"><span className="text-white/70 font-medium uppercase tracking-wider flex items-center gap-1.5"><GraduationCap size={12} /> Schüler</span><span className="font-bold text-blue-300 text-sm">{totals.std}</span></div>
            </div>
          </section>

          {/* Actions */}
          <section style={{ backgroundColor: C.card, borderColor: C.border }} className="p-4 rounded-3xl border shadow-sm shrink-0">
              <button onClick={() => setActiveModal('coursePreview')} style={{ color: C.muted, backgroundColor: C.subtle, borderColor: C.border }} className="w-full py-2.5 border rounded-xl text-[10px] font-bold uppercase tracking-widest hover:opacity-80 transition-all flex items-center justify-center gap-2">
              <Eye size={16} /> Kursübersicht
            </button>

            <button
              disabled={isLoadingPool || !courseDictionary.length || !config.institute?.trim() || !!isEnrolInvalid}
              onClick={generateList}
              style={{ backgroundColor: isEnrolInvalid ? '#B45309' : C.main }}
              className="w-full py-3 text-white rounded-xl font-bold shadow-md mt-3 transition-all hover:brightness-110 active:scale-95 flex items-center justify-center gap-2 disabled:opacity-50 group text-sm">
              <Users size={16} className="group-hover:scale-110 transition-transform" /> Liste generieren
              <kbd className="ml-1 opacity-40 text-[9px] font-mono">⌘G</kbd>
            </button>
            {isEnrolInvalid && (
              <p className="text-[10px] font-medium text-center mt-1.5 text-amber-600">
                {isEnrolInvalid === 'dauer' ? '⚠ Dauer (Tage) muss größer 0 sein.' : '⚠ Einschreibezeitraum liegt in der Vergangenheit.'}
              </p>
            )}

            <button disabled={!isGenerated} onClick={() => setActiveModal('dataPreview')} style={{ color: C.accent1, borderColor: C.accent1 + '33' }} className="w-full py-2.5 bg-transparent border rounded-xl text-[10px] font-bold uppercase mt-3 hover:opacity-80 disabled:opacity-40 transition-all flex items-center justify-center gap-1.5">
              <ClipboardList size={14} /> Daten-Vorschau
            </button>

            <div className="grid grid-cols-2 gap-2 pt-3 border-t mt-3" style={{ borderColor: C.border }}>
              <button disabled={!isGenerated || isExportingPDF} onClick={downloadPDF} style={{ backgroundColor: C.accent1 }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1 hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest transition-all">
                {isExportingPDF ? <Loader2 size={13} className="animate-spin" /> : <FileDown size={13} />} PDF <kbd className="opacity-40 font-mono text-[8px]">⌘P</kbd>
              </button>
              <button disabled={!isGenerated} onClick={downloadCSV} style={{ backgroundColor: C.accent2 }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1 hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest transition-all">
                <FileSpreadsheet size={13} /> CSV <kbd className="opacity-40 font-mono text-[8px]">⌘E</kbd>
              </button>
            </div>
            {showSessionResetConfirm ? (
              <div className="flex items-center gap-2 mt-3 w-full">
                <span className="text-[10px] text-rose-600 font-bold flex-1">Wirklich zurücksetzen?</span>
                <button onClick={() => { handleSessionReset(); setShowSessionResetConfirm(false); }} className="bg-rose-500 text-white text-[10px] font-bold px-2.5 py-1 rounded-lg hover:bg-rose-600 transition-all">Ja</button>
                <button onClick={() => setShowSessionResetConfirm(false)} style={{ borderColor: C.border, color: C.muted }} className="border text-[10px] font-bold px-2.5 py-1 rounded-lg hover:bg-black/5 transition-all">Nein</button>
              </div>
            ) : (
              <button onClick={() => setShowSessionResetConfirm(true)} style={{ borderColor: C.border, color: C.muted }} className="w-full mt-3 py-2.5 border border-dashed rounded-xl text-[10px] font-bold uppercase tracking-widest hover:border-rose-400 hover:text-rose-500 hover:bg-rose-50/40 transition-all flex items-center justify-center gap-1.5">
                <RefreshCw size={12} /> Neue Liste / Reset
              </button>
            )}
          </section>

          {/* Workflow + Nav */}
          <div style={{ backgroundColor: C.accent1 + '08', borderColor: C.accent1 + '20', color: C.muted }} className="p-4 border rounded-2xl text-[10px] leading-relaxed shadow-sm mt-auto">
            <div style={{ color: C.accent1 }} className="flex items-center gap-1.5 mb-3 font-bold uppercase tracking-widest"><Info size={14} /> Workflow</div>
            <div className="space-y-2 font-medium mb-4">
              {[['1', 'Konfiguration festlegen.'], ['2', 'Kurse zuteilen / "Alle zuweisen".'], ['3', 'Generieren & Export.']].map(([n, t]) => (
                <div key={n} className="flex gap-2.5 items-center">
                  <span style={{ backgroundColor: C.accent1 }} className="text-white w-4 h-4 rounded-full flex items-center justify-center shrink-0 font-bold text-[8px]">{n}</span><p>{t}</p>
                </div>
              ))}
            </div>
            <div className="grid grid-cols-2 gap-2 pt-3 border-t" style={{ borderColor: C.accent1 + '20' }}>
              {[
                { icon: <Star size={12} />, label: 'Favoriten', modal: 'favorites', badge: favorites.length },
                { icon: <History size={12} />, label: 'History', modal: 'history', badge: exportHistory.length },
                { icon: <Settings size={12} />, label: 'Settings', modal: 'settings' },
                { icon: <HelpCircle size={12} />, label: 'Hilfe', modal: 'help' },
              ].map(btn => (
                <button key={btn.modal} onClick={() => setActiveModal(btn.modal)} style={{ backgroundColor: C.card, borderColor: C.border, color: C.muted }} className="py-2 border rounded-lg hover:opacity-80 transition-all flex items-center justify-center gap-1.5 shadow-sm relative">
                  {btn.icon}<span className="text-[9px] font-bold uppercase">{btn.label}</span>
                  {btn.badge > 0 && <span style={{ backgroundColor: C.accent1 }} className="absolute -top-1.5 -right-1.5 text-white text-[8px] font-bold w-4 h-4 rounded-full flex items-center justify-center">{btn.badge > 9 ? '9+' : btn.badge}</span>}
                </button>
              ))}
            </div>
          </div>
        </aside>

        {/* HAUPTINHALT */}
        <main className="lg:col-span-9 flex flex-col gap-5 overflow-hidden min-h-0">

          {/* Konfiguration */}
          <section style={{ backgroundColor: C.card, borderColor: C.border }} className="p-2 rounded-3xl shadow-sm border shrink-0">
            <div className="grid grid-cols-1 md:grid-cols-12 gap-2">

              {/* Org */}
              <div style={{ backgroundColor: C.subtle }} className="md:col-span-4 p-4 lg:p-5 rounded-2xl">
                <h3 style={{ color: C.accent1 }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><Building2 size={12} /> Organisation</h3>
                <div className="space-y-3">
                  <div>
                    <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase block mb-1 ml-1">Institutsname</label>
                    <div className="flex gap-1.5">
                      <input name="institute" value={config.institute} onChange={handleInput} placeholder="z.B. Volkshochschule" style={{ backgroundColor: C.card, borderColor: config.institute?.trim() ? C.border : '#ef4444', color: C.text }} className="flex-1 px-3 py-2 border rounded-lg text-sm font-medium focus:ring-1 outline-none shadow-sm transition-all placeholder:opacity-30" />
                      <button
                        onClick={saveFavorite}
                        title="Als Favorit speichern"
                        style={{ backgroundColor: '#B45309', flexShrink: 0 }}
                        className="px-2.5 py-2 rounded-lg text-white hover:brightness-110 active:scale-95 transition-all shadow-sm"
                      >
                        <Star size={15} />
                      </button>
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    {[{ label: 'Trainer', name: 'trainerCount', col: C.main }, { label: 'Kurs Anzahl', name: 'courseSlotCount', col: C.accent1 }].map(f => (
                      <div key={f.name} style={{ backgroundColor: C.card, borderColor: C.border }} className="p-2.5 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                        <label style={{ color: C.muted }} className="text-[8px] font-semibold uppercase block mb-1">{f.label}</label>
                        <input type="number" min="0" max={f.name === 'courseSlotCount' ? 8 : undefined} name={f.name} value={config[f.name]} onChange={handleInput} style={{ color: f.col }} className="w-full bg-transparent text-sm font-semibold outline-none" />
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Klassen */}
              <div style={{ backgroundColor: C.card, borderColor: C.border }} className="md:col-span-3 p-4 lg:p-5 rounded-2xl border shadow-sm">
                <h3 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><GraduationCap size={12} /> Klassen Struktur</h3>
                <div className="space-y-2">
                  {config.classSizes.map((size, idx) => (
                    <div key={idx} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="flex justify-between items-center px-3 py-1.5 rounded-lg border focus-within:border-blue-300 transition-colors">
                      <span style={{ color: C.muted }} className="text-[11px] font-medium">{size} Plätze</span>
                      <div className="flex items-center gap-1.5">
                        <span style={{ color: C.muted }} className="text-[9px]">Anz:</span>
                        <input type="number" min="0" value={config.classCounts[idx]} onChange={e => updateClassCount(idx, e.target.value)} style={{ color: C.main }} className="w-8 bg-transparent text-right text-sm font-semibold outline-none" />
                      </div>
                    </div>
                  ))}
                </div>
                {totalStudents > 1000 && (
                  <div style={{ backgroundColor: darkMode ? '#451a0320' : '#FEF3C7', borderColor: darkMode ? '#D97706' : '#FCD34D', color: darkMode ? '#FCD34D' : '#92400E' }} className="mt-2.5 flex items-start gap-2 border rounded-lg px-2.5 py-2 text-[10px] font-medium leading-snug">
                    <AlertTriangle size={13} className="shrink-0 mt-px" style={{ color: '#D97706' }} />
                    <span>CSV-Import von <strong>{totalStudents}</strong> Usern kann Moodle-Timeouts verursachen. In kleineren Chargen exportieren.</span>
                  </div>
                )}
              </div>

              {/* Zeit & Sicherheit */}
              <div style={{ backgroundColor: C.subtle }} className="md:col-span-5 p-4 lg:p-5 rounded-2xl">
                <h3 style={{ color: C.main }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><ShieldCheck size={12} /> Zeit & Sicherheit</h3>
                <div className="grid grid-cols-2 gap-3 mb-2.5">
                  <div className="space-y-1">
                    <label style={{ color: C.muted }} className="text-[8px] font-semibold uppercase ml-1">Einschreibung</label>
                    <input name="enrolDate" type="date" value={config.enrolDate} onChange={handleInput} style={{ backgroundColor: C.card, borderColor: C.border, color: C.text }} className="w-full px-2.5 py-1.5 border rounded-lg text-xs font-medium outline-none shadow-sm focus:border-blue-300 transition-colors" />
                  </div>
                  <div className="space-y-1">
                    <label style={{ color: C.muted }} className="text-[8px] font-semibold uppercase ml-1">Dauer (Tage)</label>
                    <input name="enrolPeriod" type="number" min="0" value={config.enrolPeriod} onChange={handleInput} style={{ backgroundColor: C.card, borderColor: C.border, color: C.text }} className="w-full px-2.5 py-1.5 border rounded-lg text-xs font-medium outline-none shadow-sm focus:border-blue-300 transition-colors" />
                  </div>
                </div>
                <div className="space-y-1 mb-3">
                  <label style={{ color: C.main }} className="text-[8px] font-bold uppercase ml-1 flex justify-between"><span>Eingeschrieben bis</span><span className="opacity-70 font-medium">(Autom. berechnet)</span></label>
                  <input type="date" value={endDateDisplay} onChange={handleEndDateInput} style={{ color: C.main, borderColor: C.main + '40', backgroundColor: C.main + '08' }} className="w-full px-3 py-1.5 border rounded-lg text-xs font-bold outline-none shadow-sm focus:ring-1 focus:ring-red-200 transition-all" />
                </div>
                {/* Passwort-Modus Toggle */}
                <div className="flex items-center justify-between mb-1.5">
                  <span style={{ color: C.muted }} className="text-[8px] font-bold uppercase tracking-widest">Starkes Passwort</span>
                  <button
                    onClick={() => setConfig(p => ({ ...p, autoPassword: !p.autoPassword }))}
                    style={{ backgroundColor: config.autoPassword ? C.accent2 : C.border }}
                    className="relative flex items-center w-10 h-5 rounded-full transition-colors shrink-0">
                    <span style={{ transform: config.autoPassword ? 'translateX(21px)' : 'translateX(2px)' }} className="absolute w-3.5 h-3.5 bg-white rounded-full shadow transition-transform" />
                  </button>
                </div>
                <p style={{ color: C.muted }} className="text-[9px] mb-2 opacity-70 leading-snug">
                  {config.autoPassword ? '⚡ Auto: Einzigartiges Passwort pro Account (Moodle-konform, min. 8 Zeichen)' : '— Vorgegeben: Gleiches Passwort für alle Accounts'}
                </p>
                <div className="grid grid-cols-2 gap-2">
                  {[{ label: 'PW Schüler', name: 'studentPwd' }, { label: 'PW Trainer', name: 'trainerPwd' }].map(f => (
                    <div key={f.name} style={{ backgroundColor: C.card, borderColor: C.border, opacity: config.autoPassword ? 0.45 : 1 }} className="p-2 rounded-lg border shadow-sm transition-opacity">
                      <label style={{ color: C.muted }} className="text-[8px] font-semibold uppercase block mb-0.5">{f.label}</label>
                      <input
                        name={f.name}
                        value={config.autoPassword ? '' : config[f.name]}
                        onChange={handleInput}
                        disabled={config.autoPassword}
                        placeholder={config.autoPassword ? 'Auto-generiert' : ''}
                        style={{ color: C.text }}
                        className="w-full bg-transparent text-[11px] font-mono font-medium outline-none disabled:cursor-not-allowed placeholder:opacity-60 placeholder:font-sans placeholder:not-italic" />
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </section>

          {/* MATRIX */}
          <section style={{ backgroundColor: C.card, borderColor: C.border }} className="rounded-3xl shadow-sm border overflow-hidden flex flex-col flex-1 min-h-0 relative">
            <div style={{ backgroundColor: C.card, borderColor: C.border }} className="px-5 py-3 border-b flex items-center justify-between shrink-0 relative z-[41]">
              <div className="flex items-center gap-2 text-[11px] font-bold uppercase tracking-wider" style={{ color: C.text }}>
                <CheckCircle2 style={{ color: C.accent2 }} size={16} /> Kurs Zuweisung Matrix
                {invalidClassIds.size > 0 && (
                  <span className="px-2 py-0.5 bg-rose-100 text-rose-600 rounded-full text-[9px] font-bold animate-pulse">
                    ⚠ {invalidClassIds.size} ohne Zuweisung
                  </span>
                )}
              </div>
              <div style={{ color: C.muted, backgroundColor: C.subtle, borderColor: C.border }} className="text-[9px] font-medium px-2.5 py-1 rounded-full border">
                {classRows.length} Klassen · {config.courseSlotCount} Kurse
              </div>
            </div>

            <div className="overflow-auto flex-1 relative z-0 custom-scrollbar" style={{ backgroundColor: C.card }}>
              <table className="w-full text-left border-separate border-spacing-0 min-w-max">
                <thead className="sticky top-0 z-[40]">
                  <tr>
                    <th style={{ backgroundColor: C.card, borderColor: C.border, color: C.muted }} className="px-4 py-2 w-32 text-center border-b font-semibold uppercase tracking-widest text-[9px] shadow-sm">Klasse</th>
                    {Array.from({ length: config.courseSlotCount }).map((_, i) => (
                      <th key={i} style={{ backgroundColor: C.card, borderColor: C.border }} className="px-2 py-2.5 border-b min-w-[140px] max-w-[170px] shadow-sm">
                        <div className="flex flex-col gap-1.5">
                          <div className="flex items-center justify-between">
                            <span style={{ color: C.muted }} className="text-[9px] uppercase tracking-widest font-semibold">Kurs {i + 1}</span>
                            {config.selectedPoolCourseIds[i] !== 'none' && (() => {
                              const cid = String(config.selectedPoolCourseIds[i]);
                              const allAssigned = classRows.length > 0 && classRows.every(r => (classMatrix[r.id] || []).map(String).includes(cid));
                              return (
                                <button
                                  onClick={() => {
                                    if (allAssigned) {
                                      setClassMatrix(prev => {
                                        const next = { ...prev };
                                        classRows.forEach(r => { next[r.id] = (next[r.id] || []).map(String).filter(x => x !== cid); });
                                        return next;
                                      });
                                      addToast(`Kurs ${i + 1} von allen Klassen entfernt.`, 'success', 2000);
                                    } else {
                                      setClassMatrix(prev => {
                                        const next = { ...prev };
                                        classRows.forEach(r => { next[r.id] = [...new Set([...(next[r.id] || []).map(String), cid])]; });
                                        return next;
                                      });
                                      setInvalidClassIds(new Set());
                                      addToast(`Alle Klassen → Kurs ${i + 1} zugewiesen.`, 'success', 2000);
                                    }
                                  }}
                                  title={allAssigned ? `Kurs ${i + 1} von allen Klassen entfernen` : `Alle Klassen diesem Kurs zuweisen`}
                                  style={{ color: allAssigned ? C.main : C.accent2, backgroundColor: (allAssigned ? C.main : C.accent2) + '15' }}
                                  className="text-[8px] font-bold uppercase px-1.5 py-0.5 rounded-md hover:brightness-95 transition-all flex items-center gap-0.5"
                                >
                                  {allAssigned ? <Square size={10} /> : <CheckSquare size={10} />} Alle
                                </button>
                              );
                            })()}
                          </div>
                          <div className="relative w-full">
                            <select value={config.selectedPoolCourseIds[i] || 'none'} onChange={e => updateCourseSlot(i, e.target.value)} style={{ backgroundColor: C.card, borderColor: C.border, color: C.text }} className="w-full appearance-none border rounded-md pl-2 pr-6 py-1.5 text-[11px] font-medium outline-none cursor-pointer hover:border-blue-300 transition-all shadow-sm truncate">
                              <option value="none">-- Nicht belegt --</option>
                              {courseDictionary.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
                            </select>
                            <ChevronDown size={14} style={{ color: C.muted }} className="absolute right-2 top-1/2 -translate-y-1/2 pointer-events-none" />
                          </div>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody style={{ borderColor: C.border }} className="divide-y text-[11px]">
                  {classRows.length === 0 ? (
                    <tr><td colSpan={config.courseSlotCount + 1} style={{ color: C.muted }} className="py-10 text-center italic font-medium">Bitte oben Klassen definieren.</td></tr>
                  ) : classRows.map(c => {
                    const isInvalid = invalidClassIds.has(c.id);
                    const customName = config.classNames?.[c.id - 1]?.trim();
                    return (
                      <tr key={c.id} style={{ backgroundColor: isInvalid ? '#FFF1F2' : undefined }} className={`transition-colors group ${!isInvalid ? 'hover:bg-blue-50/10' : ''}`}>
                        <td style={{ borderColor: isInvalid ? '#FECDD3' : C.border, backgroundColor: isInvalid ? '#FFE4E6' : C.subtle }} className="px-3 py-3 text-center border-r">
                          <div className="flex flex-col gap-0.5 items-center">
                            {isInvalid && <AlertTriangle size={11} className="text-rose-500 mb-0.5" />}
                            <span style={{ color: isInvalid ? '#BE123C' : C.text }} className="font-bold text-[11px] tracking-tight">K-{String(c.id).padStart(2, '0')}</span>
                            {customName && <span style={{ color: C.accent1, backgroundColor: C.accent1 + '15' }} className="text-[8px] font-bold px-1.5 py-0.5 rounded truncate max-w-[90px]">{customName}</span>}
                            <span style={{ color: isInvalid ? '#BE123C' : C.muted, backgroundColor: isInvalid ? '#FEE2E2' : C.card, borderColor: isInvalid ? '#FECDD3' : C.border }} className="text-[9px] font-semibold uppercase tracking-tighter border px-1.5 py-0.5 rounded shadow-sm">{c.size} Pl.</span>
                          </div>
                        </td>
                        {Array.from({ length: config.courseSlotCount }).map((_, i) => {
                          const cid = config.selectedPoolCourseIds[i];
                          const isActive = cid !== 'none';
                          const isSel = isActive && (classMatrix[c.id] || []).map(String).includes(String(cid));
                          return (
                            <td key={i} className="px-4 py-2.5 text-center">
                              {isActive ? (
                                <button onClick={() => toggleCourseAssignment(c.id, cid)} title={isSel ? 'Zuweisung entfernen' : 'Kurs zuweisen'}
                                  style={{ backgroundColor: isSel ? C.accent2 : C.card, color: isSel ? 'white' : C.muted, borderColor: isSel ? C.accent2 : isInvalid ? '#FECDD3' : C.border }}
                                  className={`w-10 h-10 rounded-xl border-[1.5px] transition-all duration-200 flex items-center justify-center mx-auto shadow-sm ${isSel ? 'scale-105 shadow-md' : 'hover:scale-105'}`}>
                                  {isSel ? <Check size={20} strokeWidth={2.5} /> : <Plus size={18} strokeWidth={2} />}
                                </button>
                              ) : (
                                <div style={{ borderColor: C.border, backgroundColor: C.subtle }} className="w-10 h-10 rounded-xl border border-dashed mx-auto flex items-center justify-center">
                                  <div style={{ backgroundColor: C.border }} className="w-1.5 h-1.5 rounded-full" />
                                </div>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div className="absolute bottom-0 left-0 right-0 h-4 bg-gradient-to-t from-black/5 to-transparent pointer-events-none" />
          </section>
        </main>
      </div>

      <style dangerouslySetInnerHTML={{ __html: `
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; border-radius: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        @keyframes slideInRight { from { opacity: 0; transform: translateX(20px); } to { opacity: 1; transform: translateX(0); } }
      ` }} />
    </div>
  );
};

export default App;