import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import { uploadToSharePoint, uploadMoodleResultToSharePoint } from './sharepoint';
import { enrollInMoodle, fetchFullEnrollments, fetchMoodleCourses, findMaxNumbers, fetchInstituteGroups, fetchGroupMembers, fetchInstituteUsers , findExistingUsernames } from './moodle';
import { getAllZohoAccounts, findOrCreateZohoAccount, createZohoDeal } from './zoho';
import { captureConsole, setSecretProvider, buildLogText, copyToClipboard, entryCount, clearLog, enablePersistence, logInfo, logWarn, logError } from './logger';
import { invoke } from '@tauri-apps/api/core';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import QRCode from 'qrcode';
import { buildAccessWorkbook, excelFileName, EXISTING_PW_TEXT } from './excelExport';
import { parseOrderCsv } from './orderImport';
import { LazyStore } from '@tauri-apps/plugin-store';
import { check as checkUpdate } from '@tauri-apps/plugin-updater';
import { relaunch } from '@tauri-apps/plugin-process';
import {
  Users, FileSpreadsheet, CheckCircle2, Building2, Plus,
  Loader2, Table as TableIcon, Check, AlertTriangle, ChevronDown,
  Eye, X, RefreshCw, Info, Settings, HelpCircle, BookOpen,
  Zap, ClipboardList, ShieldCheck, GraduationCap, FileDown,
  Save, Wifi, WifiOff, Trash2, History,
  Moon, Sun, Keyboard, CheckSquare, Square, Edit3, Upload, EyeOff, Tag, FlaskConical
} from 'lucide-react';

/**
 * Moodle Anmeldungen V4
 * - VALIDIERUNG: Rote Matrix-Markierung bei fehlenden Kurszuweisungen
 * - ALLE ZUWEISEN: Ein-Klick alle Klassen auf alle aktiven Kurse
 * - KLASSEN-NAMEN: Anpassbar in den Settings
 * - EXPORT-HISTORY: Vollständige Liste aller Exporte
 * - DUNKELMODUS: System-synchron + manueller Toggle
 * - SHORTCUTS: ⌘G/E/P/A/? mit Erläuterung im Hilfe-Modal
 */

// ─── Store ─────────────────────────────────────────────────────────────────────
const store = new LazyStore('moodle-settings.json', { autoSave: false });

// Diagnose-Protokoll: Warnungen und Fehler der Konsole mitschreiben (einmalig beim Laden)
captureConsole();

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
  showLeitfaden: true,
  enrolPeriod: 340,
  defaultEnrolPeriod: 340,
  enrolDate: new Date().toISOString().split('T')[0],
  classSizes: [20, 25, 30, 40],
  classCounts: { 0: 1, 1: 1, 2: 1, 3: 1 },
  classNames: {},
  trainerCount: 2,
  courseSlotCount: 1,
  selectedPoolCourseIds: Array(8).fill('none'),
  classCustomSizes: {},
  courseApiUrl: '',
  sharepointUrl: '',
  moodleUrl: '',
  moodleToken: '',
  zohoClientId: '',
  zohoClientSecret: '',
  zohoRefreshToken: '',
  customAccents: ['#ab0325', '#153d61', '#f59e0b', '#00664f'],
  tagColorMap: { 'Schule': 0, 'Test': 1 },
  moodleBetaEnabled: false,
  moodleBetaCourseIds: [],
};


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
const DEFAULT_ACCENTS = ['#ab0325', '#153d61', '#f59e0b', '#00664f'];
const esc = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
const fmtTime = d => d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
const fmtDate = d => d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
const fmtDateTime = d => `${fmtDate(d)}, ${fmtTime(d)}`;

// ─── Toast ─────────────────────────────────────────────────────────────────────
const Toast = ({ toasts, removeToast }) => (
  <div className="fixed bottom-5 right-5 z-[999] flex flex-col gap-2 pointer-events-none">
    {toasts.map(t => (
      <div key={t.id} className={`flex items-center gap-3 px-4 py-3 rounded-2xl shadow-xl text-white text-xs font-semibold pointer-events-auto
        ${t.type === 'error' ? 'bg-rose-600' : t.type === 'success' ? 'bg-emerald-600' : t.type === 'warning' ? 'bg-amber-600' : 'bg-slate-700'}`}
        style={{ animation: 'slideInRight 0.25s ease' }}>
        {t.type === 'error' ? <AlertTriangle size={14} /> : t.type === 'warning' ? <AlertTriangle size={14} /> : <CheckCircle2 size={14} />}
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
// Zufall kommt aus crypto.getRandomValues, nicht aus Math.random — die Passwörter
// gehen als Initialkennwort an Schüler und dürfen nicht vorhersagbar sein.
const randomInt = max => {
  // Rejection Sampling: schneidet den Rest oberhalb des größten Vielfachen von
  // max ab, sonst wären kleine Werte minimal wahrscheinlicher (Modulo-Bias).
  const limit = Math.floor(2 ** 32 / max) * max;
  const buf = new Uint32Array(1);
  let value;
  do {
    crypto.getRandomValues(buf);
    value = buf[0];
  } while (value >= limit);
  return value % max;
};

const generatePassword = () => {
  const U = 'ABCDEFGHJKLMNPQRSTUVWXYZ', L = 'abcdefghjkmnpqrstuvwxyz';
  const D = '23456789', S = '!@#$%&*-=?';
  const pool = U + L + D + S;
  const r = s => s[randomInt(s.length)];
  const chars = [r(U), r(L), r(D), r(S), ...Array.from({ length: 6 }, () => r(pool))];
  // Fisher-Yates — ein Vergleichs-Sort mit Zufallskomparator mischt nicht gleichverteilt
  for (let i = chars.length - 1; i > 0; i--) {
    const j = randomInt(i + 1);
    [chars[i], chars[j]] = [chars[j], chars[i]];
  }
  return chars.join('');
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
  // Bestellliste (Marktplatz-CSV): orderPreview = geprüft, noch nicht übernommen; orderImport = aktiv
  const [orderPreview, setOrderPreview] = useState(null);
  const [orderImport, setOrderImport] = useState(null);
  const [isGenerated, setIsGenerated] = useState(false);
  const [isExportingPDF, setIsExportingPDF] = useState(false);
  const [isExportingExcel, setIsExportingExcel] = useState(false);
  const [isUploadingSP, setIsUploadingSP] = useState(false);
  const [isMoodleEnrolling, setIsMoodleEnrolling] = useState(false);
  const [showMoodleToken, setShowMoodleToken] = useState(false);
  const [showMoodleConfirm, setShowMoodleConfirm] = useState(false);
  const [zohoAllAccounts, setZohoAllAccounts] = useState([]);
  const [zohoSearching, setZohoSearching] = useState(false);
  const [openCourseSlot, setOpenCourseSlot] = useState(null);
  const [courseSlotSearch, setCourseSlotSearch] = useState('');
  const [zohoDropdownOpen, setZohoDropdownOpen] = useState(false);
  const [zohoSelectedId, setZohoSelectedId] = useState(null);
  const [instituteSearch, setInstituteSearch] = useState('');
  const [showZohoRefreshToken, setShowZohoRefreshToken] = useState(false);
  const [showZohoClientId, setShowZohoClientId] = useState(false);
  const [showZohoClientSecret, setShowZohoClientSecret] = useState(false);
  const [showZohoTokenModal, setShowZohoTokenModal] = useState(false);
  const [zohoGrantCode, setZohoGrantCode] = useState('');
  const [isExchangingZohoToken, setIsExchangingZohoToken] = useState(false);
  const [isStoreLoaded, setIsStoreLoaded] = useState(false);
  const [saveStatus, setSaveStatus] = useState('idle');
  const [lastSavedAt, setLastSavedAt] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [isOnline, setIsOnline] = useState(navigator.onLine);
  const [exportHistory, setExportHistory] = useState([]);
  const [invalidClassIds, setInvalidClassIds] = useState(new Set());
  const [pendingUpdate, setPendingUpdate] = useState(null);
  const [isInstalling, setIsInstalling] = useState(false);
  const [installProgress, setInstallProgress] = useState(0);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [settingsTab, setSettingsTab] = useState('allgemein');
  const [enrolMode, setEnrolMode] = useState('update'); // 'update' | 'new'
  const [isFindingMaxNumbers, setIsFindingMaxNumbers] = useState(false);
  const [selectedUpdateClassIds, setSelectedUpdateClassIds] = useState(new Set());
  const [moodleGroups, setMoodleGroups] = useState([]); // [{id, name, courseId, existingCourseIds, memberCount}]
  const [isLoadingMoodleGroups, setIsLoadingMoodleGroups] = useState(false);
  const [selectedMoodleGroupIds, setSelectedMoodleGroupIds] = useState(new Set());
  const [lockedClassCourseMap, setLockedClassCourseMap] = useState({}); // rowId → Set<courseId string>
  const [editingClassSizeId, setEditingClassSizeId] = useState(null);
  const [editingClassSizeVal, setEditingClassSizeVal] = useState('');
  const [showSessionResetConfirm, setShowSessionResetConfirm] = useState(false);
  const [showGenerateConfirm, setShowGenerateConfirm] = useState(false);
  const [moodleProgress, setMoodleProgress] = useState(null); // null | { label, pct, done, error }
  const [moodleResult, setMoodleResult] = useState(null); // null | { moodleSummary, sharepoint, zoho }
  const [appVersion, setAppVersion] = useState('...');
  const [isCheckingUpdate, setIsCheckingUpdate] = useState(false);
  const [allMoodleCourses, setAllMoodleCourses] = useState([]);
  const [isLoadingMoodleCourses, setIsLoadingMoodleCourses] = useState(false);
  const [moodleBetaSearch, setMoodleBetaSearch] = useState('');

  const [helpTab, setHelpTab] = useState('workflow');
  const [classNamesResetKey, setClassNamesResetKey] = useState(0);
  const saveTimeoutRef = useRef(null);
  const classGroupOffsetRef = useRef(0);
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
    // Jede sichtbare Fehlermeldung landet auch im Diagnose-Protokoll
    if (type === 'error') logError('hinweis', message);
    else if (type === 'warning') logWarn('hinweis', message);
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
    // Diagnose-Protokoll über Neustarts behalten; Sitzungsbeginn als Trennzeile
    enablePersistence(new LazyStore('diagnose-log.json', { autoSave: false }))
      .finally(() => import('@tauri-apps/api/app').then(m => m.getVersion()).catch(() => '?')
        .then(v => logInfo('sitzung', '──── App gestartet ────', `Version ${v}`)));
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
      await relaunch();
    } catch (e) { addToast(`Update fehlgeschlagen: ${e.message}`, 'error'); setIsInstalling(false); }
  }, [pendingUpdate, addToast]);

  // ─── Diagnose-Protokoll: aktuelle Geheimwerte fürs Maskieren ────────────
  const diagConfigRef = useRef(config);
  diagConfigRef.current = config;
  useEffect(() => {
    setSecretProvider(() => {
      const c = diagConfigRef.current;
      return {
        moodleToken: c.moodleToken, zohoClientId: c.zohoClientId, zohoClientSecret: c.zohoClientSecret,
        zohoRefreshToken: c.zohoRefreshToken, studentPwd: c.studentPwd, trainerPwd: c.trainerPwd,
        courseApiUrl: c.courseApiUrl, sharepointUrl: c.sharepointUrl,
      };
    });
  }, []);

  // ─── Store: Laden ─────────────────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      try {
        const ts = await store.get('lastSavedAt');
        if (ts) setLastSavedAt(new Date(ts));
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
        const showLeitfaden = await store.get('showLeitfaden');
        if (showLeitfaden !== null && showLeitfaden !== undefined) setConfig(p => ({ ...p, showLeitfaden }));
        const pool = await store.get('coursePool');
        if (pool?.length) setCourseDictionary(pool);
        const moodleUrl = await store.get('moodleUrl');
        if (moodleUrl) setConfig(p => ({ ...p, moodleUrl }));
        const moodleToken = await store.get('moodleToken');
        if (moodleToken) setConfig(p => ({ ...p, moodleToken }));
        const courseApiUrl = await store.get('courseApiUrl');
        if (courseApiUrl) setConfig(p => ({ ...p, courseApiUrl }));
        const sharepointUrl = await store.get('sharepointUrl');
        if (sharepointUrl) setConfig(p => ({ ...p, sharepointUrl }));
        const zohoClientId = await store.get('zohoClientId');
        if (zohoClientId) setConfig(p => ({ ...p, zohoClientId }));
        const zohoClientSecret = await store.get('zohoClientSecret');
        if (zohoClientSecret) setConfig(p => ({ ...p, zohoClientSecret }));
        const zohoRefreshToken = await store.get('zohoRefreshToken');
        if (zohoRefreshToken) setConfig(p => ({ ...p, zohoRefreshToken }));
        // Schulen bekommen 340 Tage: alter Standard 180 wird einmalig umgestellt, selbst gewählte Werte bleiben
        const storedEnrolPeriod = await store.get('defaultEnrolPeriod');
        const schoolPeriodApplied = await store.get('schoolEnrolPeriodApplied');
        const defaultEnrolPeriod = !schoolPeriodApplied && Number(storedEnrolPeriod) === 180 ? 340 : storedEnrolPeriod;
        if (defaultEnrolPeriod != null) setConfig(p => ({ ...p, defaultEnrolPeriod, enrolPeriod: defaultEnrolPeriod }));
        if (!schoolPeriodApplied) await store.set('schoolEnrolPeriodApplied', true);
        const customAccents = await store.get('customAccents');
        if (Array.isArray(customAccents) && customAccents.length) setConfig(p => ({ ...p, customAccents }));
        const tagColorMap = await store.get('tagColorMap');
        if (tagColorMap && typeof tagColorMap === 'object') setConfig(p => ({ ...p, tagColorMap }));
        const moodleBetaEnabled = await store.get('moodleBetaEnabled');
        if (moodleBetaEnabled != null) setConfig(p => ({ ...p, moodleBetaEnabled: !!moodleBetaEnabled }));
        const moodleBetaCourseIds = await store.get('moodleBetaCourseIds');
        if (Array.isArray(moodleBetaCourseIds)) setConfig(p => ({ ...p, moodleBetaCourseIds }));
        const moodleCoursesCache = await store.get('moodleCoursesCache');
        if (Array.isArray(moodleCoursesCache) && moodleCoursesCache.length > 0) setAllMoodleCourses(moodleCoursesCache);
        const savedEnrolMode = await store.get('enrolMode');
        if (savedEnrolMode === 'update' || savedEnrolMode === 'new') setEnrolMode(savedEnrolMode);
        // ─── Migrations ───────────────────────────────────────────────────────
        // v1.1.1: fix classSizes[0] = 2 → 20
        setConfig(p => {
          const s = p.classSizes;
          if (Array.isArray(s) && s[0] === 2) {
            const fixed = [20, ...s.slice(1)];
            store.set('classSizes', fixed).then(() => store.save()).catch(() => {});
            return { ...p, classSizes: fixed };
          }
          return p;
        });
        // v1.1.1: apply correct default accent colors if still at old default
        setConfig(p => {
          const oldDefault = ['#6366f1', '#0ea5e9', '#f59e0b', '#10b981'];
          const isOldDefault = p.customAccents.every((c, i) => c === oldDefault[i]);
          if (isOldDefault) {
            const newAccents = DEFAULT_CONFIG.customAccents;
            store.set('customAccents', newAccents).then(() => store.save()).catch(() => {});
            return { ...p, customAccents: newAccents };
          }
          return p;
        });
        // Erst nach erfolgreichem Laden freigeben — sonst würde der Auto-Save
        // die gespeicherten Werte mit den leeren Standardwerten überschreiben.
        setIsStoreLoaded(true);
      } catch (e) {
        console.error('Store laden:', e);
        addToast('Einstellungen konnten nicht geladen werden — Änderungen werden in dieser Sitzung nicht gespeichert.', 'error');
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
        await store.set('exportHistory', exportHistory);
        await store.set('darkMode', darkMode);
        await store.set('classSizes', config.classSizes);
        await store.set('classNames', config.classNames);
        await store.set('studentPwd', config.studentPwd);
        await store.set('trainerPwd', config.trainerPwd);
        await store.set('autoPassword', config.autoPassword);
        await store.set('showLeitfaden', config.showLeitfaden);
        await store.set('moodleUrl', config.moodleUrl);
        await store.set('moodleToken', config.moodleToken);
        await store.set('courseApiUrl', config.courseApiUrl);
        await store.set('sharepointUrl', config.sharepointUrl);
        await store.set('zohoClientId', config.zohoClientId);
        await store.set('zohoClientSecret', config.zohoClientSecret);
        await store.set('zohoRefreshToken', config.zohoRefreshToken);
        await store.set('defaultEnrolPeriod', config.defaultEnrolPeriod);
        await store.set('customAccents', config.customAccents);
        await store.set('tagColorMap', config.tagColorMap);
        await store.set('moodleBetaEnabled', config.moodleBetaEnabled);
        await store.set('moodleBetaCourseIds', config.moodleBetaCourseIds);
        await store.set('enrolMode', enrolMode);
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
  }, [exportHistory, darkMode, config.classSizes, config.classNames, config.studentPwd, config.trainerPwd, config.autoPassword, config.showLeitfaden, config.moodleUrl, config.moodleToken, config.courseApiUrl, config.sharepointUrl, config.zohoClientId, config.zohoClientSecret, config.zohoRefreshToken, config.moodleBetaEnabled, config.moodleBetaCourseIds, enrolMode, isStoreLoaded]); // eslint-disable-line

  // ─── Zoho: Alle Accounts laden (einmalig wenn aktiviert) ──────────────────
  const zohoEnabled = !!(config.zohoClientId && config.zohoClientSecret && config.zohoRefreshToken);
  useEffect(() => {
    console.log('[Zoho] zohoEnabled:', zohoEnabled, {
      hasClientId: !!config.zohoClientId,
      hasClientSecret: !!config.zohoClientSecret,
      hasRefreshToken: !!config.zohoRefreshToken,
    });
    if (!zohoEnabled) { setZohoAllAccounts([]); return; }
    console.log('[Zoho] Starte Account-Ladung...');
    setZohoSearching(true);
    getAllZohoAccounts(config)
      .then(accounts => {
        const sorted = [...accounts].sort((a, b) => a.Account_Name.localeCompare(b.Account_Name, 'de'));
        console.log('[Zoho] Accounts geladen:', sorted.length, sorted.slice(0, 3));
        setZohoAllAccounts(sorted);
      })
      .catch(e => console.error('[Zoho] Accounts laden fehlgeschlagen:', e.message, e))
      .finally(() => setZohoSearching(false));
  }, [zohoEnabled]); // eslint-disable-line

  // ─── Zoho: Gefilterte Vorschläge (client-side) ─────────────────────────────
  const zohoSuggestions = useMemo(() => {
    if (!zohoEnabled || !zohoAllAccounts.length) return [];
    const q = config.institute?.trim().toLowerCase().replace(/-/g, ' ');
    console.log('[Zoho] Filter:', { q, total: zohoAllAccounts.length, treffer: zohoAllAccounts.filter(a => !q || a.Account_Name.toLowerCase().includes(q)).length });
    if (!q) return zohoAllAccounts;
    return zohoAllAccounts.filter(a =>
      a.Account_Name.toLowerCase().includes(q)
    );
  }, [zohoAllAccounts, config.institute, zohoEnabled]);

  // Breite der Tag-Spalte im Kurs-Dropdown: so breit wie das längste Tag
  const courseTagColWidth = useMemo(() => {
    const maxLen = Math.max(0, ...courseDictionary.map(c => c.tag?.length ?? 0));
    return maxLen > 0 ? `${maxLen * 7 + 16}px` : '0px';
  }, [courseDictionary]);

  const getTagColor = useCallback(tag => {
    if (!tag) return '#94a3b8';
    const palette = config.customAccents?.length ? config.customAccents : DEFAULT_ACCENTS;
    const idx = config.tagColorMap?.[tag];
    return idx != null ? (palette[idx] ?? palette[0]) : palette[0];
  }, [config.customAccents, config.tagColorMap]);

  // ─── Zoho: Grant Code → Refresh Token tauschen ────────────────────────────
  const handleZohoTokenExchange = useCallback(async () => {
    if (!zohoGrantCode.trim() || !config.zohoClientId || !config.zohoClientSecret) {
      addToast('Bitte zuerst Client ID, Client Secret und Grant Code eintragen.', 'error');
      return;
    }
    setIsExchangingZohoToken(true);
    try {
      const rawText = await invoke('zoho_exchange_token', {
        clientId: config.zohoClientId,
        clientSecret: config.zohoClientSecret,
        code: zohoGrantCode.trim(),
      });
      let data;
      try { data = JSON.parse(rawText); } catch { throw new Error(`Ungültige Antwort: ${rawText}`); }
      if (!data.refresh_token) {
        throw new Error(data.error_description || data.error || rawText);
      }
      setConfig(p => ({ ...p, zohoRefreshToken: data.refresh_token }));
      setZohoGrantCode('');
      setShowZohoTokenModal(false);
      addToast('Zoho Refresh Token erfolgreich generiert und gespeichert!', 'success');
    } catch (e) {
      console.error('[Zoho] Token-Exchange Fehler:', e);
      addToast(`Zoho Token-Exchange fehlgeschlagen: ${e?.message ?? String(e)}`, 'error', 0);
    } finally {
      setIsExchangingZohoToken(false);
    }
  }, [zohoGrantCode, config.zohoClientId, config.zohoClientSecret, addToast]);

  // ─── Diagnose-Protokoll: Export ──────────────────────────────────────────
  const logHeader = useCallback(() => ({
    'App-Version': appVersion || 'unbekannt',
    'System': typeof navigator !== 'undefined' ? navigator.userAgent : 'unbekannt',
    'Moodle': (config.moodleUrl || '').replace(/^https?:\/\//, '').replace(/\/+$/, '') || 'nicht gesetzt',
    'Moodle-Token': config.moodleToken ? 'gesetzt' : 'fehlt',
    'Kursliste': config.moodleBetaEnabled ? 'direkt aus Moodle' : (config.courseApiUrl ? 'Power Automate' : 'nur Cache'),
    'Zoho': zohoEnabled ? 'aktiv' : 'nicht eingerichtet',
    'Modus': enrolMode,
  }), [appVersion, config.moodleUrl, config.moodleToken, config.moodleBetaEnabled, config.courseApiUrl, zohoEnabled, enrolMode]);

  const handleCopyLog = useCallback(async () => {
    const ok = await copyToClipboard(buildLogText(logHeader()));
    addToast(ok ? 'Protokoll kopiert — in eine Mail einfügen und an den Support schicken.' : 'Kopieren fehlgeschlagen — bitte „Speichern" nutzen.', ok ? 'success' : 'error');
  }, [logHeader, addToast]);

  const handleSaveLog = useCallback(async () => {
    const name = 'diagnose-protokoll.json';
    try {
      const file = new LazyStore(name, { autoSave: false });
      await file.set('erstellt', new Date().toISOString());
      await file.set('protokoll', buildLogText(logHeader()).split('\n'));
      await file.save();
      const { appDataDir, join } = await import('@tauri-apps/api/path');
      const { revealItemInDir } = await import('@tauri-apps/plugin-opener');
      await revealItemInDir(await join(await appDataDir(), name));
      addToast('Protokoll gespeichert — die Datei ist im Explorer bzw. Finder markiert.', 'success');
    } catch (e) {
      logError('diagnose', 'Speichern fehlgeschlagen', e);
      addToast('Speichern fehlgeschlagen — bitte „Kopieren" nutzen.', 'error');
    }
  }, [logHeader, addToast]);

  const [, setLogRev] = useState(0);
  const handleClearLog = useCallback(() => {
    clearLog();
    logInfo('sitzung', 'Protokoll geleert');
    setLogRev(r => r + 1);
    addToast('Protokoll geleert.', 'success');
  }, [logHeader, addToast]);

  // ─── Kurs-Pool ────────────────────────────────────────────────────────────
  const fetchCoursePool = useCallback(async () => {
    if (config.moodleBetaEnabled) return; // Moodle-Modus: Kurse kommen direkt von Moodle
    if (!config.courseApiUrl?.trim()) return; // keine URL konfiguriert → Cache bleibt erhalten
    setIsLoadingPool(true);
    try {
      const text = await invoke('http_post_json', { url: config.courseApiUrl, body: JSON.stringify({ request: 'get_courses' }) });
      const raw = JSON.parse(text);
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
        const tag = findValueByPattern(item, ['tag', 'tags', 'kategorie', 'category', 'typ', 'type', 'gruppe', 'group']) || '';
        return { id: String(rawId), label: String(label).trim(), shorthand: String(sh).trim(), url: String(url).trim(), tag: String(tag).trim() };
      });
      if (normalized.length === 0) {
        addToast('Kurs-Pool ist leer — Cache bleibt erhalten.', 'error');
        return;
      }
      setCourseDictionary(normalized);
      if (normalized.length > 0) {
        setConfig(prev => {
          const ids = [...prev.selectedPoolCourseIds];
          normalized.slice(0, prev.courseSlotCount).forEach((c, i) => { if (ids[i] === 'none') ids[i] = c.id; });
          return { ...prev, selectedPoolCourseIds: ids };
        });
        store.set('coursePool', normalized).then(() => store.save()).catch(() => {});
      }
      logInfo('kursliste', `${normalized.length} Kurse aus Power Automate geladen`);
      addToast(`${normalized.length} Kurse geladen.`, 'success', 2500);
    } catch (e) {
      logError('kursliste', 'Power Automate nicht erreichbar — Cache wird verwendet', e);
      addToast('Verbindung zum Kurs-Pool fehlgeschlagen — verwende Cache.', 'error');
    }
    finally { setIsLoadingPool(false); }
  }, [addToast]);
  useEffect(() => { fetchCoursePool(); }, [fetchCoursePool]);

  // ─── Beta: Moodle Kurskatalog ─────────────────────────────────────────────
  const fetchMoodleCoursesCatalog = useCallback(async () => {
    if (!config.moodleUrl?.trim() || !config.moodleToken?.trim()) {
      addToast('Moodle-URL und Token müssen in den Einstellungen hinterlegt sein.', 'error');
      return;
    }
    setIsLoadingMoodleCourses(true);
    try {
      const courses = await fetchMoodleCourses(config.moodleUrl, config.moodleToken);
      setAllMoodleCourses(courses);
      store.set('moodleCoursesCache', courses).then(() => store.save()).catch(() => {});
      addToast(`${courses.length} Kurse von Moodle geladen.`, 'success');
    } catch (e) {
      addToast(`Moodle-Kurse konnten nicht geladen werden: ${e.message}`, 'error');
    } finally {
      setIsLoadingMoodleCourses(false);
    }
  }, [config.moodleUrl, config.moodleToken, addToast]);

  // Wenn Beta aktiv: courseDictionary = ausgewählte Moodle-Kurse
  useEffect(() => {
    if (!config.moodleBetaEnabled) return;
    const selected = allMoodleCourses.filter(c => (config.moodleBetaCourseIds ?? []).includes(c.id));
    setCourseDictionary(selected);
  }, [config.moodleBetaEnabled, config.moodleBetaCourseIds, allMoodleCourses]);

  // ─── Berechnungen ─────────────────────────────────────────────────────────
  const baseMatrixCourses = useMemo(() =>
    config.selectedPoolCourseIds.slice(0, config.courseSlotCount)
      .map(id => courseDictionary.find(c => String(c.id) === String(id)))
      .filter(c => c && c.id !== 'none'),
    [config.selectedPoolCourseIds, config.courseSlotCount, courseDictionary]);

  const classRows = useMemo(() => {
    const buildNewRows = (startId) => {
      const rows = []; let id = startId;
      [0, 1, 2, 3].forEach(idx => {
        for (let i = 0; i < config.classCounts[idx]; i++) {
          const defaultSize = config.classSizes[idx];
          const customSize = config.classCustomSizes?.[id];
          const size = customSize != null ? customSize : defaultSize;
          rows.push({ id: id++, size, typeIdx: idx, defaultSize, isNew: true });
        }
      });
      return rows;
    };
    if ((enrolMode === 'update' || enrolMode === 'both') && moodleGroups.length > 0) {
      const existingRows = moodleGroups.map((g, i) => {
        const rowId = i + 1;
        const customSize = config.classCustomSizes?.[rowId];
        // Leere Gruppen (Nutzer manuell gelöscht): Custom-Größe erlaubt neue Schüler einzuschreiben
        const size = g.memberCount === 0 && customSize != null ? customSize : g.memberCount;
        return {
          id: rowId,
          size,
          typeIdx: 0,
          defaultSize: g.memberCount,
          groupName: g.name,
          moodleGroupId: g.id,
          courseId: g.courseId,
          existingCourseIds: g.existingCourseIds || [],
          isExisting: true,
        };
      });
      if (enrolMode === 'both') {
        return [...existingRows, ...buildNewRows(existingRows.length + 1)];
      }
      return existingRows;
    }
    return buildNewRows(1);
  }, [enrolMode, moodleGroups, config.classCounts, config.classSizes, config.classCustomSizes]);

  const totals = useMemo(() => {
    const std = classRows.reduce((s, r) => s + r.size, 0);
    return { cls: classRows.length, std, trainers: config.trainerCount, all: std + config.trainerCount };
  }, [classRows, config.trainerCount]);

  const rawEndDate = useMemo(() => {
    const d = new Date(config.enrolDate);
    if (isNaN(d.getTime())) return new Date();
    d.setDate(d.getDate() + Math.max(1, parseInt(config.enrolPeriod || 1)));
    return d;
  }, [config.enrolDate, config.enrolPeriod]);
  const endDateDisplay = rawEndDate.toISOString().split('T')[0];
  const endDateFormatted = rawEndDate.toLocaleDateString('de-DE');

  const totalStudents = useMemo(() => classRows.reduce((s, r) => s + r.size, 0), [classRows]);

  // Kurse ausgewählter Klassen automatisch hinzufügen (auch wenn nicht im manuellen Pool)
  const activeMatrixCourses = useMemo(() => {
    const courses = [...baseMatrixCourses];
    const ids = new Set(courses.map(c => String(c.id)));
    const selectedRows = classRows.filter(r => r.isExisting && selectedMoodleGroupIds.has(r.moodleGroupId));
    for (const row of selectedRows) {
      for (const extId of (lockedClassCourseMap[row.id] || new Set())) {
        if (ids.has(String(extId))) continue;
        const course = courseDictionary.find(c => String(c.id) === String(extId));
        if (course) { courses.push(course); ids.add(String(extId)); }
      }
    }
    return courses;
  }, [baseMatrixCourses, classRows, lockedClassCourseMap, selectedMoodleGroupIds, courseDictionary]);

  // Wenn classRows sich ändert (Klassenanzahl geändert), alle Klassen im Update-Modus selektieren
  useEffect(() => {
    setSelectedUpdateClassIds(new Set(classRows.map(r => r.id)));
  }, [classRows]);

  // Wenn Modus wechselt oder Institut sich ändert, geladene Gruppen zurücksetzen
  useEffect(() => {
    if (enrolMode !== 'update' && enrolMode !== 'both') { setMoodleGroups([]); setSelectedMoodleGroupIds(new Set()); setLockedClassCourseMap({}); }
  }, [enrolMode]);
  useEffect(() => {
    setMoodleGroups([]); setSelectedMoodleGroupIds(new Set()); setLockedClassCourseMap({});
  }, [config.institute]);

  const unusualWarnings = useMemo(() => {
    const w = [];
    if (config.trainerCount > 20) w.push(`${config.trainerCount} Trainer (ungewöhnlich hoch)`);
    if (totalStudents > 500) w.push(`${totalStudents} Schüler (ungewöhnlich hoch)`);
    if (parseInt(config.enrolPeriod, 10) > 730) w.push(`Einschreibung ${config.enrolPeriod} Tage (> 2 Jahre)`);
    if (config.classSizes.some(s => s > 60)) w.push(`Klassengröße > 60 Schüler`);
    return w;
  }, [config.trainerCount, config.enrolPeriod, config.classSizes, totalStudents]);

  // false | 'dauer' | 'abgelaufen'
  const isEnrolInvalid = useMemo(() => {
    const period = parseInt(config.enrolPeriod, 10);
    if (!period || period <= 0) return 'dauer';
    const end = new Date(config.enrolDate);
    end.setDate(end.getDate() + period);
    return end < new Date(new Date().toDateString()) ? 'abgelaufen' : false;
  }, [config.enrolDate, config.enrolPeriod]);

  const getClassLabel = useCallback(row => {
    if (row.groupName) {
      // Moodle-Gruppe: Suffix nach "{institute}-" zurückgeben
      const prefix = config.institute?.trim() + '-';
      return row.groupName.startsWith(prefix) ? row.groupName.slice(prefix.length) : row.groupName;
    }
    // Bestellliste: Klassenname aus dem Import (nicht in classNames, das dauerhaft gespeichert wird)
    if (orderImport && !row.isExisting) {
      const g = orderImport.classes[row.id - 1]?.grade;
      if (g) return g;
    }
    const n = config.classNames?.[row.id - 1]?.trim();
    if (n) return n;
    const effectiveId = row.id + classGroupOffsetRef.current;
    return `Klasse-${String(effectiveId).padStart(2, '0')}`;
  }, [config.classNames, config.institute, orderImport]);

  // ─── Handler ──────────────────────────────────────────────────────────────
  const handleInput = useCallback(e => {
    const { name, value } = e.target;
    const isNum = ['enrolPeriod', 'trainerCount', 'courseSlotCount'].includes(name);
    const isPwd = ['studentPwd', 'trainerPwd'].includes(name);
    const isInstitute = name === 'institute';
    const replaceUmlauts = s => s
      .replace(/Ä/g, 'Ae').replace(/Ö/g, 'Oe').replace(/Ü/g, 'Ue')
      .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue')
      .replace(/ß/g, 'ss');
    const processed = isPwd ? value.trim() : isInstitute ? replaceUmlauts(value).replace(/\s+/g, '-') : value;
    const NUM_MIN = { trainerCount: 0, courseSlotCount: 1, enrolPeriod: 1 };
    setConfig(p => {
      if (isNum) {
        const raw = parseInt(value, 10) || 0;
        const lo = NUM_MIN[name] ?? 0;
        return { ...p, [name]: Math.max(lo, raw) };
      }
      return { ...p, [name]: processed };
    });
  }, []);
  const handleEndDateInput = useCallback(e => {
    if (!e.target.value) return;
    const diff = Math.ceil((new Date(e.target.value) - new Date(config.enrolDate)) / 86400000);
    setConfig(p => ({ ...p, enrolPeriod: Math.max(1, diff) }));
  }, [config.enrolDate]);
  const updateClassSize = useCallback((idx, val) => setConfig(p => { const s = [...p.classSizes]; s[idx] = Math.max(0, parseInt(val, 10) || 0); return { ...p, classSizes: s }; }), []);
  const updateClassCount = useCallback((idx, val) => {
    // classMatrix, classNames und classCustomSizes alle leeren:
    // Die Klassen-IDs sind sequenziell und verschieben sich wenn classCounts sich ändert.
    // Stale Zuweisungen wären schlimmer als ein Neustart.
    setConfig(p => ({ ...p, classCounts: { ...p.classCounts, [idx]: Math.max(0, parseInt(val, 10) || 0) }, classCustomSizes: {}, classNames: {} }));
    setClassMatrix({});
    setInvalidClassIds(new Set());
  }, []);
  const updateClassName = useCallback((rowIndex, val) => setConfig(p => ({ ...p, classNames: { ...p.classNames, [rowIndex]: val } })), []);
  const updateCustomClassSize = useCallback((rowId, val) => {
    const num = Math.max(0, parseInt(val, 10) || 0);
    setConfig(p => ({ ...p, classCustomSizes: { ...p.classCustomSizes, [rowId]: num } }));
  }, []);
  const toggleCourseAssignment = useCallback((classId, courseId) => {
    const sid = String(courseId);
    if (lockedClassCourseMap[classId]?.has(sid)) return; // gesperrter Kurs (bereits eingeschrieben)
    setClassMatrix(prev => { const cur = (prev[classId] || []).map(String); return { ...prev, [classId]: cur.includes(sid) ? cur.filter(x => x !== sid) : [...cur, sid] }; });
  }, [lockedClassCourseMap]);

  // Validierung reaktiv nachführen: wenn invalidClassIds gesetzt sind, live aktualisieren
  useEffect(() => {
    if (invalidClassIds.size === 0) return;
    const activeIds = config.selectedPoolCourseIds.slice(0, config.courseSlotCount).filter(id => id !== 'none').map(String);
    const stillBad = new Set();
    classRows.forEach(r => {
      if (r.size === 0) return;
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
    classGroupOffsetRef.current = 0;
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
    classGroupOffsetRef.current = 0;
    addToast('Neue Liste gestartet.', 'success');
  }, [addToast]);

  // ─── Export-History ───────────────────────────────────────────────────────
  const addExportEntry = useCallback((type, filename) => {
    setExportHistory(prev => [{
      id: Date.now().toString(), type, institute: config.institute,
      accounts: generatedData.length, trainers: generatedData.filter(d => d.isT).length,
      students: generatedData.filter(d => !d.isT).length, date: new Date().toISOString(), filename,
    }, ...prev]);
  }, [config.institute, generatedData]);

  // ─── Bestellliste importieren ─────────────────────────────────────────────
  const handleOrderFile = useCallback(async (e) => {
    const file = e.target.files?.[0];
    e.target.value = ''; // gleiche Datei erneut wählbar
    if (!file) return;
    try {
      const parsed = parseOrderCsv(await file.text());
      if (!parsed.persons.length) return addToast('Die Bestellliste enthält keine übernehmbaren Personen.', 'error');
      setOrderPreview({ ...parsed, fileName: file.name });
      logInfo('bestellliste', 'Bestellliste geprüft', `${parsed.persons.length} Personen, ${parsed.problems.length} Probleme`);
    } catch (err) {
      addToast(err.message, 'error', 8000);
      logWarn('bestellliste', 'Bestellliste abgelehnt', err.message);
    }
  }, [addToast]);

  const applyOrderImport = useCallback(() => {
    if (!orderPreview) return;
    // Klassenstruktur aus der Bestellung: eine Zeile pro Klasse, Größe = Anzahl Schüler:innen
    const classCustomSizes = {};
    orderPreview.classes.forEach((c, i) => { classCustomSizes[i + 1] = c.count; });
    setConfig(p => ({
      ...p,
      classCounts: { 0: orderPreview.classes.length, 1: 0, 2: 0, 3: 0 },
      classCustomSizes, classNames: {},
      trainerCount: orderPreview.teachers,
    }));
    // Zeilen-IDs werden neu vergeben — alte Kurszuweisungen würden sonst an den neuen Klassen hängen
    setClassMatrix({});
    setInvalidClassIds(new Set());
    setEnrolMode('new');
    setOrderImport(orderPreview);
    setOrderPreview(null);
    setIsGenerated(false);
    addToast(`Bestellliste übernommen: ${orderPreview.persons.length} Personen in ${orderPreview.classes.length} Klasse(n). Jetzt Kurse zuweisen.`, 'success', 6000);
  }, [orderPreview, addToast]);

  const clearOrderImport = useCallback(() => {
    setOrderImport(null);
    setConfig(p => ({ ...p, classCounts: { ...DEFAULT_CONFIG.classCounts }, classCustomSizes: {}, classNames: {} }));
    setClassMatrix({});
    setInvalidClassIds(new Set());
    setIsGenerated(false);
  }, []);

  // ─── Generierung ──────────────────────────────────────────────────────────
  const generateList = useCallback(async (confirmed = false) => {
    if (!config.institute?.trim()) return addToast('Bitte Institutsnamen eingeben.', 'error');
    if (!classRows.length && !config.trainerCount) return addToast('Keine Klassen oder Trainer.', 'error');
    if (!confirmed && unusualWarnings.length > 0) { setShowGenerateConfirm(true); return; }
    setShowGenerateConfirm(false);
    const activeIds = activeMatrixCourses.map(c => String(c.id));
    const effectiveClassRows = (enrolMode === 'update' || enrolMode === 'both') && moodleGroups.length > 0
      ? classRows.filter(r => r.isNew || (r.moodleGroupId && selectedMoodleGroupIds.has(r.moodleGroupId)))
      : enrolMode === 'update' && selectedUpdateClassIds.size < classRows.length
        ? classRows.filter(r => selectedUpdateClassIds.has(r.id))
        : classRows;
    const badIds = new Set();
    effectiveClassRows.forEach(r => {
      if (r.size === 0) return; // Klassen ohne Schüler brauchen keine Kurszuweisung
      const assigned = (classMatrix[r.id] || []).map(String);
      if (!assigned.some(id => activeIds.includes(id))) badIds.add(r.id);
    });
    if (badIds.size) { setInvalidClassIds(badIds); return addToast(`${badIds.size} Klasse(n) ohne Kurszuweisung — rot markiert.`, 'error'); }
    setInvalidClassIds(new Set());

    // ── Bestellliste: Klarnamen, Anmeldename = E-Mail, keine Nummerierung ────
    if (orderImport) {
      if (enrolMode !== 'new') return addToast('Die Bestellliste funktioniert nur im Modus „Neu anlegen“.', 'error');
      classGroupOffsetRef.current = 0;
      const pw = isT => (config.autoPassword ? generatePassword() : (isT ? config.trainerPwd : config.studentPwd));
      const person = (p, extra) => ({
        isT: p.isT, first: p.first, last: p.last, user: p.email, mail: p.email,
        displayName: `${p.first} ${p.last}`, pw: pw(p.isT), ...extra,
      });
      const data = orderImport.persons.filter(p => p.isT).map(p => person(p, { cNum: 'ALL', courses: activeMatrixCourses }));
      effectiveClassRows.forEach(r => {
        const grade = getClassLabel(r);
        const selIds = (classMatrix[r.id] || []).map(String);
        const selCourses = courseDictionary.filter(cd => selIds.includes(String(cd.id)) && activeIds.includes(String(cd.id)));
        orderImport.persons.filter(p => !p.isT && p.grade === grade).forEach(p => data.push(person(p, {
          cNum: String(r.id).padStart(2, '0'), cLabel: `${config.institute}-${grade}`, courses: selCourses,
        })));
      });
      const missing = orderImport.persons.length - data.length;
      if (missing > 0) {
        return addToast(`${missing} Person(en) aus der Bestellliste sind keiner Klasse zugeordnet — wurde die Klassenstruktur geändert? Bestellliste entfernen und neu importieren.`, 'error', 10000);
      }
      // Bestehende Konten vorab erkennen: Excel/PDF zeigen für sie kein (nie gesetztes) neues Passwort
      if (config.moodleUrl?.trim() && config.moodleToken?.trim()) {
        try {
          const existing = await findExistingUsernames(config.moodleUrl, config.moodleToken, data.map(d => d.user));
          data.forEach(d => { if (existing.has(d.user)) d.existing = true; });
          if (existing.size) addToast(`${existing.size} Konto/Konten bestehen bereits — Passwort bleibt unverändert.`, 'info', 6000);
        } catch (e) {
          addToast(`Konnte bestehende Konten nicht prüfen (${e.message}) — wird bei der Einschreibung nachgeholt.`, 'warning', 8000);
        }
      }
      setGeneratedData(data); setIsGenerated(true); setActiveModal('dataPreview');
      addToast(`${data.length} Accounts aus der Bestellliste erstellt.`, 'success');
      return;
    }

    // ── Aktualisieren / Kombiniert-Modus mit geladenen Moodle-Gruppen ─────────
    if ((enrolMode === 'update' || enrolMode === 'both') && moodleGroups.length > 0) {
      const instClean = config.institute.replace(/\s+/g, '').toLowerCase();
      const usernameSet = new Set();
      const data = [];
      const allKnownCourseMap = new Map([...courseDictionary, ...(allMoodleCourses.length > 0 ? allMoodleCourses : [])].map(c => [String(c.id), c]));

      // --- Bestehende Klassen (aus Moodle-Gruppen) ---
      const existingRows = effectiveClassRows.filter(r => r.isExisting);
      // Leere Gruppen mit Custom-Größe: neue Schüler generieren (wie neue Zeilen)
      const emptyReuseRows = existingRows.filter(r => r.defaultSize === 0 && r.size > 0);
      const nonEmptyExistingRows = existingRows.filter(r => !(r.defaultSize === 0 && r.size > 0));
      for (const r of nonEmptyExistingRows) {
        const matrixIds = new Set((classMatrix[r.id] || []).map(String));
        const allCourseIds = new Set([...(r.existingCourseIds || []).map(String), ...matrixIds]);
        // Kurse aus allKnownCourseMap holen; falls nicht vorhanden (z.B. Moodle-ID nicht in
        // courseDictionary), synthetisches Objekt mit numerischer Moodle-ID erstellen —
        // extractMoodleCourseId in moodle.js kann parseInt(id) direkt verwenden.
        const selCourses = [...allCourseIds].map(id => {
          if (allKnownCourseMap.has(id)) return allKnownCourseMap.get(id);
          const numId = parseInt(id, 10);
          if (!isNaN(numId) && numId > 0) return { id: String(numId), label: `Kurs ${numId}`, shorthand: '', url: '', tag: '' };
          return null;
        }).filter(Boolean);
        const classLabel = r.groupName;
        try {
          const members = await fetchGroupMembers(config.moodleUrl, config.moodleToken, r.courseId, r.moodleGroupId);
          for (const member of members) {
            if (!member.username || usernameSet.has(member.username)) continue;
            usernameSet.add(member.username);
            // Trainer: altes Namensmuster oder Rolle 4 (Bestelllisten-Lehrkräfte heißen wie ihre E-Mail)
            const isTrainer = member.username.includes(`${instClean}-trainer-`) || (member.roles || []).some(r => r.roleid === 4);
            data.push({
              cNum: isTrainer ? 'ALL' : String(r.id).padStart(2, '0'),
              cLabel: isTrainer ? undefined : classLabel,
              isT: isTrainer,
              first: member.firstname || (isTrainer ? 'Trainer' : 'Schüler'),
              last: member.lastname || config.institute,
              user: member.username,
              existing: true, // aus Moodle geladen — Passwort wird nicht geändert
              mail: member.email || `${member.username}@${instClean}.com`,
              pw: isTrainer ? (config.autoPassword ? generatePassword() : config.trainerPwd) : (config.autoPassword ? generatePassword() : config.studentPwd),
              courses: isTrainer ? activeMatrixCourses : selCourses,
            });
          }
        } catch (e) {
          addToast(`Fehler beim Laden von ${r.groupName}: ${e.message}`, 'error');
        }
      }

      // --- Fehlende Trainer nachladen (falls alle bestehenden Gruppen leer waren) ---
      // Wenn nonEmptyExistingRows leer war, wurden noch keine Trainer per fetchGroupMembers geladen.
      // Fallback: alle Institut-User holen, nur Trainer filtern.
      if (nonEmptyExistingRows.length === 0 && config.moodleUrl?.trim() && config.moodleToken?.trim()) {
        try {
          const allUsers = await fetchInstituteUsers(config.moodleUrl, config.moodleToken, instClean);
          for (const u of allUsers) {
            if (!u.username || !u.username.includes(`${instClean}-trainer-`) || usernameSet.has(u.username)) continue;
            usernameSet.add(u.username);
            data.push({ cNum: 'ALL', isT: true, existing: true, first: u.firstname || 'Trainer', last: u.lastname || config.institute, user: u.username, mail: u.email || `${u.username}@${instClean}.com`, pw: config.autoPassword ? generatePassword() : config.trainerPwd, courses: activeMatrixCourses });
          }
        } catch { /* nicht-kritisch — Trainer werden ggf. über trainerCount abgedeckt */ }
      }

      // --- Neue Klassen (nur im 'both'-Modus) + leere wiederverwendete Gruppen ---
      const newRows = enrolMode === 'both' ? effectiveClassRows.filter(r => r.isNew) : [];
      const rowsNeedingNewStudents = [...emptyReuseRows, ...newRows];
      if (rowsNeedingNewStudents.length > 0 || (enrolMode === 'both' && config.trainerCount > 0)) {
        let studentOffset = 0, trainerOffset = 0;
        classGroupOffsetRef.current = moodleGroups.length;
        if (config.moodleUrl?.trim() && config.moodleToken?.trim()) {
          setIsFindingMaxNumbers(true);
          const activeCourseIds = activeMatrixCourses.map(c => {
            const numId = parseInt(String(c?.id ?? ''), 10);
            if (!isNaN(numId) && numId > 0) return numId;
            if (c?.url) { const m = String(c.url).match(/[?&]id=(\d+)/); if (m) return parseInt(m[1], 10); }
            return null;
          }).filter(Boolean);
          try {
            const { maxStudent, maxTrainer } = await findMaxNumbers(config.moodleUrl, config.moodleToken, instClean, activeCourseIds);
            studentOffset = maxStudent;
            trainerOffset = maxTrainer;
          } catch (e) {
            addToast(`Moodle-Abfrage für Offset fehlgeschlagen: ${e.message} — starte bei 1.`, 'error');
          } finally {
            setIsFindingMaxNumbers(false);
          }
        }
        if (enrolMode === 'both') {
          for (let t = 1; t <= config.trainerCount; t++) {
            const tNum = trainerOffset + t;
            const uname = `${instClean}-trainer-${tNum}`;
            if (!usernameSet.has(uname)) {
              usernameSet.add(uname);
              data.push({ cNum: 'ALL', isT: true, first: 'Trainer', last: config.institute, user: uname, mail: `trainer${tNum}@${instClean}.com`, pw: config.autoPassword ? generatePassword() : config.trainerPwd, courses: activeMatrixCourses });
            }
          }
        }
        let sIdx = studentOffset + 1;
        for (const r of rowsNeedingNewStudents) {
          const selIds = (classMatrix[r.id] || []).map(String);
          // Leere Wiederverwend.-Gruppen: bestehendem Gruppen-Namen direkt verwenden (moodle.js findet Gruppe via Name)
          const classLabel = r.isExisting
            ? r.groupName
            : `${config.institute}-${getClassLabel(r)}`;
          const selCourses = r.isExisting
            ? (() => {
                const allCIds = new Set([...(r.existingCourseIds || []).map(String), ...selIds]);
                return [...allCIds].map(id => {
                  if (allKnownCourseMap.has(id)) return allKnownCourseMap.get(id);
                  const numId = parseInt(id, 10);
                  if (!isNaN(numId) && numId > 0) return { id: String(numId), label: `Kurs ${numId}`, shorthand: '', url: '', tag: '' };
                  return null;
                }).filter(Boolean);
              })()
            : courseDictionary.filter(cd => selIds.includes(String(cd.id)) && activeIds.includes(String(cd.id)));
          for (let i = 0; i < r.size; i++) {
            const id = String(sIdx++).padStart(3, '0');
            data.push({ cNum: String(r.id + classGroupOffsetRef.current).padStart(2, '0'), cLabel: classLabel, isT: false, first: 'Schüler', last: config.institute, user: `${instClean}-student-${id}`, mail: `student${id}@${instClean}.com`, pw: config.autoPassword ? generatePassword() : config.studentPwd, courses: selCourses });
          }
        }
      }

      setGeneratedData(data); setIsGenerated(true); setActiveModal('dataPreview');
      addToast(`${data.length} Accounts verarbeitet.`, 'success');
      return;
    }

    // ── Neu-Anlegen-Modus: höchste bestehende Nummern + leere Gruppen abfragen ─
    let studentOffset = 0;
    let trainerOffset = 0;
    let classOffset = 0;
    let emptyGroupsQueue = []; // bestehende leere Gruppen → werden zuerst wiederverwendet
    if (enrolMode === 'new') {
      if (!config.moodleUrl?.trim() || !config.moodleToken?.trim()) {
        addToast('Neu-Anlegen: Moodle-URL und Token fehlen — starte bei 1.', 'info', 4000);
      } else {
        setIsFindingMaxNumbers(true);
        try {
          const instClean = config.institute.replace(/\s+/g, '').toLowerCase();
          const activeCourseIds = activeMatrixCourses.map(c => {
            const numId = parseInt(String(c?.id ?? ''), 10);
            if (!isNaN(numId) && numId > 0) return numId;
            if (c?.url) { const m = String(c.url).match(/[?&]id=(\d+)/); if (m) return parseInt(m[1], 10); }
            return null;
          }).filter(Boolean);
          // Kursübergreifende Klassen-Nummerierung: alle bekannten Kurse durchsuchen
          const { maxStudent, maxTrainer, orphanUsernames } = await findMaxNumbers(config.moodleUrl, config.moodleToken, instClean, activeCourseIds);
          studentOffset = maxStudent;
          trainerOffset = maxTrainer;
          // Klassen-Offset aus höchstem Student ableiten: ceil(maxStudent / Ø-Klassengröße)
          // Setzt voraus, dass frühere Sessions dieselbe Klassengröße verwendet haben.
          if (maxStudent > 0 && classRows.length > 0) {
            const avgSize = Math.round(classRows.reduce((s, r) => s + r.size, 0) / classRows.length);
            classOffset = avgSize > 0 ? Math.ceil(maxStudent / avgSize) : 0;
          }
          if (maxStudent > 0 || maxTrainer > 0 || classOffset > 0)
            addToast(`Neu-Anlegen: Starte bei Schüler ${maxStudent + 1}, Trainer ${maxTrainer + 1}, Klasse ${classOffset + 1}.`, 'success', 5000);
          if (orphanUsernames.length > 0) {
            const preview = orphanUsernames.slice(0, 3).join(', ');
            const more = orphanUsernames.length > 3 ? ` (+${orphanUsernames.length - 3} weitere)` : '';
            addToast(`⚠ ${orphanUsernames.length} Account(s) ohne Kurseinschreibung in Moodle gefunden: ${preview}${more}. Diese wurden bei der Nummerierung berücksichtigt — bitte in Moodle prüfen und ggf. löschen.`, 'warning', 12000);
          }
          // Leere Gruppen laden — neue Schüler werden dort zuerst eingeschrieben
          if (activeCourseIds.length > 0) {
            try {
              const allGroups = await fetchInstituteGroups(config.moodleUrl, config.moodleToken, config.institute, activeCourseIds);
              emptyGroupsQueue = allGroups.filter(g => g.memberCount === 0).sort((a, b) => a.name.localeCompare(b.name));
              if (emptyGroupsQueue.length > 0)
                addToast(`${emptyGroupsQueue.length} leere Gruppe(n) gefunden — neue Schüler werden dort zuerst eingeschrieben.`, 'info', 5000);
            } catch { /* nicht-kritisch */ }
          }
        } catch (e) {
          addToast(`Moodle-Abfrage fehlgeschlagen: ${e.message} — starte bei 1.`, 'error');
        } finally {
          setIsFindingMaxNumbers(false);
        }
      }
    }

    classGroupOffsetRef.current = classOffset;
    const instClean = config.institute.replace(/\s+/g, '').toLowerCase();
    const data = [];
    for (let t = 1; t <= config.trainerCount; t++) {
      const tNum = trainerOffset + t;
      data.push({ cNum: 'ALL', isT: true, first: 'Trainer', last: config.institute, user: `${instClean}-trainer-${tNum}`, mail: `trainer${tNum}@${instClean}.com`, pw: config.autoPassword ? generatePassword() : config.trainerPwd, courses: activeMatrixCourses });
    }
    // Map: App-Kurs-ID (String) → Moodle-Kurs-ID (Number) — für kursbasierten Gruppen-Abgleich
    const appToMoodleId = new Map();
    activeMatrixCourses.forEach(c => {
      const numId = parseInt(String(c?.id ?? ''), 10);
      if (!isNaN(numId) && numId > 0) { appToMoodleId.set(String(c.id), numId); return; }
      if (c?.url) { const m = String(c.url).match(/[?&]id=(\d+)/); if (m) appToMoodleId.set(String(c.id), parseInt(m[1], 10)); }
    });
    let sIdx = studentOffset + 1;
    const remainingEmptyGroups = [...emptyGroupsQueue]; // mutable Kopie, wird bei Treffer gespleißt
    effectiveClassRows.forEach(r => {
      if (r.size === 0) return; // keine Schüler → keine Gruppe verbrauchen
      const selIds = (classMatrix[r.id] || []).map(String);
      const selCourses = courseDictionary.filter(cd => selIds.includes(String(cd.id)) && activeIds.includes(String(cd.id)));
      // Leere Gruppe nur wiederverwenden wenn sie in mindestens einem konfigurierten Kurs existiert
      const selMoodleIds = new Set(selIds.map(id => appToMoodleId.get(id)).filter(Boolean));
      const matchIdx = remainingEmptyGroups.findIndex(g =>
        (g.existingCourseIds || []).some(cId => selMoodleIds.has(Number(cId)))
      );
      const classLabel = matchIdx >= 0
        ? remainingEmptyGroups.splice(matchIdx, 1)[0].name
        : `${config.institute}-${getClassLabel(r)}`;
      for (let i = 0; i < r.size; i++) {
        const id = String(sIdx++).padStart(3, '0');
        data.push({ cNum: String(r.id + classGroupOffsetRef.current).padStart(2, '0'), cLabel: classLabel, isT: false, first: 'Schüler', last: config.institute, user: `${instClean}-student-${id}`, mail: `student${id}@${instClean}.com`, pw: config.autoPassword ? generatePassword() : config.studentPwd, courses: selCourses });
      }
    });
    setGeneratedData(data); setIsGenerated(true); setActiveModal('dataPreview');
    addToast(`${data.length} Accounts generiert.`, 'success');
  }, [config, classRows, classMatrix, activeMatrixCourses, courseDictionary, allMoodleCourses, getClassLabel, addToast, unusualWarnings, enrolMode, selectedUpdateClassIds, moodleGroups, selectedMoodleGroupIds, orderImport]); // eslint-disable-line

  // ─── CSV ──────────────────────────────────────────────────────────────────
  const buildCsvBlob = useCallback(() => {
    if (!generatedData.length) return null;
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
    return new Blob(['\uFEFF', headers.join(',') + '\r\n' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  }, [generatedData, activeMatrixCourses, classRows, config, getClassLabel]);

  const downloadCSV = useCallback(() => {
    const blob = buildCsvBlob();
    if (!blob) return;
    const fname = `EBCL-Moodle-Upload-${config.institute.replace(/\s+/g, '_')}-${new Date().toISOString().split('T')[0]}.csv`;
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: fname });
    a.click(); URL.revokeObjectURL(a.href);
    addExportEntry('CSV', fname); addToast('CSV heruntergeladen.', 'success');
  }, [buildCsvBlob, config.institute, addToast, addExportEntry]);

  // ─── Excel ────────────────────────────────────────────────────────────────
  const buildExcelBlob = useCallback(async (dataOverride = null) => {
    const data = dataOverride ?? generatedData;
    if (!data.length) return null;
    const toAccount = d => ({ user: d.user, pw: d.pw, name: d.displayName || '', existing: !!d.existing, courses: d.courses || [] });
    const classIds = [...new Set(data.filter(d => !d.isT).map(d => d.cNum))].sort();
    const classes = classIds.map(id => {
      const students = data.filter(d => !d.isT && d.cNum === id);
      const row = classRows.find(r => String(r.id + classGroupOffsetRef.current).padStart(2, '0') === id);
      return { label: row ? getClassLabel(row) : `Klasse-${id}`, accounts: students.map(toAccount) };
    });
    const buffer = await buildAccessWorkbook({
      institute: config.institute,
      dateStr: new Date().toLocaleDateString('de-DE'),
      periodStr: `${new Date(config.enrolDate).toLocaleDateString('de-DE')} – ${endDateFormatted}`,
      trainers: data.filter(d => d.isT).map(toAccount),
      classes,
      personal: data.some(d => d.displayName),
    });
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }, [generatedData, config, classRows, getClassLabel, endDateFormatted]);

  const downloadExcel = useCallback(async ({ returnBlob = false, dataOverride = null } = {}) => {
    const blob = await buildExcelBlob(dataOverride);
    if (!blob) return returnBlob ? null : undefined;
    if (returnBlob) return blob;

    setIsExportingExcel(true);
    try {
      const fname = excelFileName(config.institute, config.enrolDate);
      // WebKit (Tauri/macOS) behandelt binary Blob-URLs als Navigation statt Download —
      // über FileReader zu Data-URL konvertieren, damit der Download korrekt ausgelöst wird.
      await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          const a = document.createElement('a');
          a.href = reader.result;
          a.download = fname;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          resolve();
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
      addExportEntry('XLSX', fname);
      addToast('Excel heruntergeladen.', 'success');
    } catch (e) {
      console.error(e);
      addToast('Excel-Export fehlgeschlagen.', 'error');
    } finally {
      setIsExportingExcel(false);
    }
  }, [buildExcelBlob, config.institute, addToast, addExportEntry]);

  // ─── PDF ──────────────────────────────────────────────────────────────────
  const downloadPDF = useCallback(async ({ returnBlob = false, dataOverride = null } = {}) => {
    if (!returnBlob) setIsExportingPDF(true);
    const data = dataOverride ?? generatedData;
    try {
      const qrDataUrl = await getQrDataUrl().catch(() => null);
      const doc = new jsPDF('l', 'mm', 'a4');
      const primary = [157, 32, 43];
      const fname = `EBCL-Zugangsdaten-${config.institute.replace(/\s+/g, '_')}-${new Date().toISOString().split('T')[0]}.pdf`;
      const periodStr = `${new Date(config.enrolDate).toLocaleDateString('de-DE')} – ${endDateFormatted}`;
      
      const renderHeader = (title, info) => {
        doc.setFontSize(20).setTextColor(...primary).setFont('helvetica', 'bold').text(config.institute.toUpperCase(), 15, 20);
        doc.setFontSize(12).setTextColor(60).setFont('helvetica', 'normal').text(title, 15, 28);
        doc.setFontSize(10).setTextColor(37, 99, 235).text('Zugang: https://world.ebcl.eu/', 15, 34).link(15, 31, 65, 4, { url: 'https://world.ebcl.eu/' });
        doc.setFontSize(6.5).setTextColor(130).setFont('helvetica', 'italic').text('Hinweis: Es kann beim Kopieren der Zugangsdaten zu einer Fehlermeldung kommen. Dann bitte diese händisch eintippen.', 15, 39);
        doc.setFontSize(8).setTextColor(120).setFont('helvetica', 'normal').text(`FREISCHALTZEITRAUM: ${periodStr} | ${info}`, 15, 44);
        doc.setLineWidth(0.3).setDrawColor(0).line(15, 47, 282, 47);
        if (qrDataUrl) doc.addImage(qrDataUrl, 'PNG', 256, 4, 26, 26);
      };

      const pageMap = {};

      // ─── Leitfaden-Seite (Seite 1, optional) ───────────────────────────────
      let firstDataPage = !config.showLeitfaden; // true = erste Datenseite nutzt Seite 1
      if (config.showLeitfaden) {
      renderHeader('LEITFADEN FÜR TRAINER', `FREISCHALTZEITRAUM: ${periodStr}`);
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
        doc.roundedRect(col1, 170, 267, 17, 2, 2, 'FD');
        doc.setFontSize(6.5).setFont('helvetica', 'bold').setTextColor(...primary);
        doc.text('HINWEIS:', col1 + 4, 175);
        doc.setFont('helvetica', 'normal').setTextColor(...bodyColor);
        doc.text('•  Ihre persönlichen Zugangsdaten (Benutzername & Passwort) befinden sich auf der nächsten Seite dieses Dokuments.', col1 + 6, 180);
        doc.text('•  Je nach PDF-Reader kann beim Kopieren von Benutzername oder Passwort ein überflüssiges Leerzeichen entstehen – bitte vor dem Anmelden prüfen!', col1 + 6, 185);
      }
      } // end if showLeitfaden

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

      // Hilfsfunktion: alle einzigartigen Kurse einer Account-Gruppe (nach ID, Reihenfolge bleibt erhalten)
      const getUniqueCourses = (accounts) => {
        const seen = new Set();
        const result = [];
        accounts.forEach(a => a.courses.forEach(c => {
          if (!seen.has(String(c.id))) { seen.add(String(c.id)); result.push(c); }
        }));
        return result;
      };
      // Kurs-Label für eine bestimmte Spalte (oder leer wenn User nicht eingeschrieben)
      const courseRow = (account, colCourses) =>
        colCourses.map(cc => account.courses.find(c => String(c.id) === String(cc.id))?.label || '');

      const trainers = data.filter(d => d.isT);
      if (trainers.length) {
        const trainerCourses = getUniqueCourses(trainers);
        if (!firstDataPage) doc.addPage(); else firstDataPage = false;
        renderHeader('Zugangsdaten: Trainer', `ANZAHL: ${trainers.length}`);
        autoTable(doc, {
          head: [['Name (fakultativ)', 'Username', 'Passwort', ...trainerCourses.map((_, i) => `Kurs ${i + 1}`)]],
          body: trainers.map(t => [t.displayName || '', t.user, t.existing ? EXISTING_PW_TEXT : t.pw, ...courseRow(t, trainerCourses)]),
          ...tOpts(trainerCourses, `Trainer — ${config.institute}`),
          didParseCell: d => { if (d.section === 'body') d.cell.styles.fillColor = [255, 255, 245]; }
        });
      }

      [...new Set(data.filter(d => !d.isT).map(d => d.cNum))].sort().forEach((id) => {
        if (!firstDataPage) doc.addPage(); else firstDataPage = false;
        const students = data.filter(d => d.cNum === id);
        const classCourses = getUniqueCourses(students);
        const row = classRows.find(r => String(r.id + classGroupOffsetRef.current).padStart(2, '0') === id);
        const classLabel = row ? getClassLabel(row) : `Klasse-${id}`;
        renderHeader(classLabel, `ANZAHL: ${students.length}`);

        autoTable(doc, {
          head: [['Name (fakultativ)', 'Username', 'Passwort', ...classCourses.map((_, i) => `Kurs ${i + 1}`)]],
          body: students.map(s => [s.displayName || '', s.user, s.existing ? EXISTING_PW_TEXT : s.pw, ...courseRow(s, classCourses)]),
          ...tOpts(classCourses, `${classLabel} — ${config.institute}`)
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
        doc.text(`Freischaltungszeitraum: ${periodStr}`, 282, 197, { align: 'right' });
        doc.setFontSize(6).setTextColor(160);
        doc.text(fname, 15, 201.5);
        doc.text(`Seite ${i} von ${pc}`, 282, 201.5, { align: 'right' });
      }

      if (returnBlob) return doc.output('blob');
      doc.save(fname);
      addExportEntry('PDF', fname);
      addToast('PDF exportiert (interaktiv).', 'success');
    } catch (e) {
      console.error(e);
      if (!returnBlob) addToast('PDF-Export fehlgeschlagen.', 'error');
      throw e;
    } finally {
      if (!returnBlob) setIsExportingPDF(false);
    }
  }, [generatedData, config, activeMatrixCourses, classRows, getClassLabel, endDateFormatted, addToast, addExportEntry]);

  // ─── SharePoint ───────────────────────────────────────────────────────────
  const handleSharePointUpload = useCallback(async () => {
    setIsUploadingSP(true);
    try {
      const dateStr = new Date().toISOString().split('T')[0];
      const instClean = config.institute.replace(/\s+/g, '_');
      const folderName = `${dateStr}_${instClean}`;
      const csvName = `EBCL-Moodle-Upload-${instClean}-${dateStr}.csv`;
      const pdfName = `EBCL-Zugangsdaten-${instClean}-${dateStr}.pdf`;
      const xlsxName = excelFileName(config.institute, config.enrolDate);
      const csvBlob = buildCsvBlob();
      const [pdfBlob, xlsxBlob] = await Promise.all([downloadPDF({ returnBlob: true }), downloadExcel({ returnBlob: true })]);
      if (!csvBlob || !pdfBlob || !xlsxBlob) return addToast('Daten fehlen für Upload.', 'error');
      const ok = await uploadToSharePoint(csvBlob, pdfBlob, xlsxBlob, folderName, csvName, pdfName, xlsxName, config.sharepointUrl);
      if (ok) addToast(`SharePoint: Ordner „${folderName}" erstellt.`, 'success');
      else addToast('SharePoint Upload fehlgeschlagen.', 'error');
    } catch (e) {
      console.error(e);
      addToast('SharePoint Upload fehlgeschlagen.', 'error');
    } finally {
      setIsUploadingSP(false);
    }
  }, [config.institute, buildCsvBlob, downloadPDF, downloadExcel, addToast]);

  // ─── Moodle: Klassen aus Moodle laden (Aktualisieren-Modus) ──────────────
  const loadMoodleGroups = useCallback(async () => {
    if (!config.moodleUrl?.trim() || !config.moodleToken?.trim()) return addToast('Moodle-URL und Token fehlen (Einstellungen → Backend).', 'error');
    if (!config.institute?.trim()) return addToast('Bitte zuerst Institutsnamen eingeben.', 'error');
    setIsLoadingMoodleGroups(true);
    try {
      // Alle Moodle-Kurs-IDs bestimmen (allMoodleCourses wenn vorhanden, sonst frisch laden)
      let searchCourses = allMoodleCourses.length > 0 ? allMoodleCourses : courseDictionary;
      if (searchCourses.length === 0) {
        try {
          searchCourses = await fetchMoodleCourses(config.moodleUrl, config.moodleToken);
        } catch { searchCourses = []; }
      }
      const allCourseIds = searchCourses.map(c => {
        const numId = parseInt(String(c?.id ?? ''), 10);
        return (!isNaN(numId) && numId > 0) ? numId : null;
      }).filter(Boolean);

      if (allCourseIds.length === 0) return addToast('Keine Kurse verfügbar — bitte Kurskatalog laden oder Kurs-Pool befüllen.', 'error');

      const groups = await fetchInstituteGroups(config.moodleUrl, config.moodleToken, config.institute, allCourseIds);
      if (groups.length === 0) {
        addToast(`Keine Klassen für "${config.institute}" in Moodle gefunden.`, 'warning', 6000);
        return;
      }

      // Matrix mit bestehenden Pool-Kursen vorfüllen + Locked-Map aufbauen
      // courseDictionary.id ist die externe API-ID; Moodle-ID steckt in .url (?id=123) oder ist numerisch
      const getMoodleId = c => {
        const numId = parseInt(String(c?.id ?? ''), 10);
        if (!isNaN(numId) && numId > 0) return numId;
        if (c?.url) { const m = String(c.url).match(/[?&]id=(\d+)/); if (m) return parseInt(m[1], 10); }
        return null;
      };
      const newMatrix = {};
      const newLocked = {};
      groups.forEach((g, i) => {
        const rowId = i + 1;
        const existingMoodleIds = new Set((g.existingCourseIds || []).map(Number));
        // Finde Pool-Kurse deren Moodle-ID in den existierenden Kurs-IDs liegt
        const existingPoolIds = courseDictionary
          .filter(c => existingMoodleIds.has(getMoodleId(c)))
          .map(c => String(c.id)); // externe API-ID, passend zu selectedPoolCourseIds
        newMatrix[rowId] = existingPoolIds;
        newLocked[rowId] = new Set(existingPoolIds);
      });

      // Auto-fill pool: ensure all courses referenced in newMatrix are in selectedPoolCourseIds
      // and courseSlotCount is large enough to show them as matrix columns
      const allNeededExtIds = [...new Set(Object.values(newMatrix).flat())];
      const allExistingMoodleIds = new Set(
        groups.flatMap(g => (g.existingCourseIds || []).map(Number))
      );
      const newPoolIds = [...config.selectedPoolCourseIds];
      let newCourseSlotCount = config.courseSlotCount;
      let poolChanged = false;

      // 1. Add courses from courseDictionary that are needed but not yet in pool slots
      for (const extId of allNeededExtIds) {
        if (newPoolIds.some(id => String(id) === String(extId))) continue;
        const emptySlot = newPoolIds.findIndex(id => !id || id === 'none');
        if (emptySlot === -1) break;
        newPoolIds[emptySlot] = extId;
        newCourseSlotCount = Math.max(newCourseSlotCount, emptySlot + 1);
        poolChanged = true;
      }

      // 2. Also handle courses in allMoodleCourses not in courseDictionary (Beta / direct Moodle)
      const missingFromCatalog = [...allExistingMoodleIds].filter(mid =>
        !courseDictionary.some(c => getMoodleId(c) === mid)
      );
      if (missingFromCatalog.length > 0 && allMoodleCourses.length > 0) {
        for (const mid of missingFromCatalog) {
          const course = allMoodleCourses.find(c => parseInt(String(c.id), 10) === mid);
          if (!course) continue;
          const moodleIdStr = String(course.id);
          if (newPoolIds.some(id => String(id) === moodleIdStr)) continue;
          const emptySlot = newPoolIds.findIndex(id => !id || id === 'none');
          if (emptySlot === -1) break;
          newPoolIds[emptySlot] = moodleIdStr;
          newCourseSlotCount = Math.max(newCourseSlotCount, emptySlot + 1);
          poolChanged = true;
        }
      }

      if (poolChanged) {
        const addedCount = newPoolIds.filter((id, i) => id !== config.selectedPoolCourseIds[i] && id !== 'none').length;
        setConfig(prev => ({ ...prev, selectedPoolCourseIds: newPoolIds, courseSlotCount: newCourseSlotCount }));
        if (addedCount > 0) addToast(`${addedCount} Kurs(e) automatisch in den Pool eingetragen.`, 'info', 4000);
      }

      setMoodleGroups(groups);
      setSelectedMoodleGroupIds(new Set()); // standardmäßig keine Klasse ausgewählt
      setClassMatrix(newMatrix);
      setLockedClassCourseMap(newLocked);
      addToast(`${groups.length} Klasse(n) aus Moodle geladen.`, 'success');
    } catch (e) {
      addToast(`Fehler beim Laden: ${e.message}`, 'error');
    } finally {
      setIsLoadingMoodleGroups(false);
    }
  }, [config.moodleUrl, config.moodleToken, config.institute, allMoodleCourses, courseDictionary, addToast]);

  // ─── Moodle Einschreibung ──────────────────────────────────────────────────
  const handleMoodleEnrol = useCallback(async () => {
    const progress = (label, pct) => setMoodleProgress({ label, pct, done: false, error: false });
    setIsMoodleEnrolling(true);
    setMoodleProgress({ label: 'Verbindung aufbauen…', pct: 2, done: false, error: false });
    logInfo('ablauf', 'Moodle-Einschreibung gestartet', { institut: config.institute, modus: enrolMode, konten: generatedData.length, kurse: activeMatrixCourses.length, beginn: config.enrolDate, tage: config.enrolPeriod });

    let result;
    try {
      // ── Schritt 1–5: Moodle ───────────────────────────────────────────────
      result = await enrollInMoodle({
        baseUrl: config.moodleUrl,
        token: config.moodleToken,
        generatedData,
        activeMatrixCourses,
        classRows,
        config,
        getClassLabel,
        onProgress: progress,
        enrolMode,
      });
    } catch (e) {
      console.error('[Moodle] Fehler:', e);
      setMoodleProgress({ label: `Fehler: ${e.message}`, pct: 0, done: false, error: true });
      setIsMoodleEnrolling(false);
      return; // Abbruch — kein SharePoint, kein Zoho
    }

    result.warnings.forEach(w => { addToast(w, 'info'); logWarn('moodle', w); });
    addExportEntry('Moodle', 'Moodle-Einschreibung');
    logInfo('ablauf', 'Moodle-Einschreibung abgeschlossen', `${Object.keys(result.userIdMap).length} Konten`);

    // ── Schritt 6: Nur tatsächlich eingeschriebene User behalten ──────────
    // Basis-Filter: nur User die in Moodle erstellt/gefunden wurden (in userIdMap).
    // User die weder angelegt noch gefunden werden konnten stehen nicht im PDF.
    const enrolledUsernames = new Set(Object.keys(result.userIdMap));
    const existingSet = new Set(result.existingUsernames || []);
    const moodleData = generatedData
      .filter(u => enrolledUsernames.has(u.user?.trim().toLowerCase()))
      // Bestehende Konten behalten ihr Passwort — in Excel/PDF nicht das neu generierte zeigen
      .map(u => (existingSet.has(u.user?.trim().toLowerCase()) ? { ...u, existing: true } : u));

    // ── Schritt 7: Vollständige Kurseinschreibungen laden ─────────────────
    // Mit Anreicherung: User mit 0 echten Moodle-Kursen werden gefiltert.
    // Ohne Anreicherung (Berechtigung fehlt): moodleData als Fallback.
    let exportData = moodleData;
    try {
      progress('Kurseinschreibungen aus Moodle laden…', 86);
      const enriched = await fetchFullEnrollments(
        config.moodleUrl, config.moodleToken,
        moodleData, result.userIdMap, courseDictionary
      );
      // User mit 0 echten Moodle-Kursen rausfiltern — Einschreibung fehlgeschlagen
      exportData = enriched.filter(u => u.courses?.length > 0);
      setGeneratedData(exportData);
    } catch (e) {
      console.warn('[Moodle] Kursanreicherung fehlgeschlagen:', e.message);
      addToast('Kursanreicherung fehlgeschlagen — PDF zeigt nur erfolgreich angelegte Accounts.', 'warning', 6000);
      setGeneratedData(moodleData);
    }

    let sharepointDone = false;
    let zohoDone = false;
    const postErrors = []; // Fehler nach Moodle — alle Steps laufen durch

    // ── Schritt 7: SharePoint ─────────────────────────────────────────────
    if (config.sharepointUrl) {
      progress('SharePoint hochladen…', 90);
      const now = new Date();
      const dateStr = now.toISOString().split('T')[0];
      const instClean = config.institute.replace(/\s+/g, '_');
      const folderName = `${dateStr}_${instClean}_Moodle`;
      const existingUsers = result.usersResolved - result.usersCreated;
      const cohortStatus = result.cohortId
        ? `${result.cohortName} (ID ${result.cohortId}, ${result.cohortCreated ? 'neu angelegt' : 'bereits vorhanden'}, ${result.cohortMembersAdded} Mitglieder)`
        : 'nicht angelegt';
      const lines = [
        'EBCL Moodle-Einschreibung – Zusammenfassung',
        '==========================================',
        `Institut:        ${config.institute}`,
        `Datum:           ${fmtDateTime(now)}`,
        `Moodle-URL:      ${config.moodleUrl}`,
        '',
        'Ergebnis', '--------',
        `User gesamt:      ${result.usersResolved}`,
        `  Neu angelegt:   ${result.usersCreated}`,
        `  Bereits vorh.:  ${existingUsers}`,
        `Einschreibungen:  ${result.enrolmentsDone}`,
        `Gruppen:          ${result.groupsCreated}`,
        '',
        'Kohorte', '-------',
        `Name:    ${result.cohortName}`,
        `Status:  ${cohortStatus}`,
        '',
        'Kurse', '-----',
        ...activeMatrixCourses.map(c => `  ${c.shorthand}  ${c.label}`),
        '',
        ...(result.warnings.length ? ['Hinweise', '--------', ...result.warnings, ''] : []),
        'Accounts', '--------',
        ...exportData.map(u =>
          `${u.isT ? '[Trainer]' : '[Schüler]'}  ${u.user.padEnd(40)} PW: ${u.existing ? EXISTING_PW_TEXT : u.pw}  Kurse: ${u.courses.map(c => c.shorthand).join(', ')}`
        ),
      ];
      const txtBlob = new Blob([lines.join('\r\n')], { type: 'text/plain;charset=utf-8' });
      const txtName = `EBCL-Moodle-Import-${instClean}-${dateStr}.txt`;
      const pdfName = `EBCL-Zugangsdaten-${instClean}-${dateStr}.pdf`;
      const xlsxName = excelFileName(config.institute, config.enrolDate);
      try {
        const pdfBlob = await downloadPDF({ returnBlob: true, dataOverride: exportData });
        const xlsxBlob = await downloadExcel({ returnBlob: true, dataOverride: exportData });
        if (pdfBlob && xlsxBlob) {
          const ok = await uploadMoodleResultToSharePoint(
            txtBlob, pdfBlob, xlsxBlob,
            folderName, txtName, pdfName, xlsxName,
            config.sharepointUrl
          );
          if (!ok) throw new Error('Upload-Antwort negativ');
          sharepointDone = true;
        }
      } catch (spErr) {
        console.error('[SharePoint] Fehler:', spErr);
        postErrors.push(`SharePoint: ${spErr.message}`);
      }
    }

    // ── Schritt 8: Zoho CRM — unabhängig von SharePoint ──────────────────
    if (zohoEnabled) {
      progress('Zoho CRM aktualisieren…', 96);
      try {
        const { account } = await findOrCreateZohoAccount(config, config.institute);
        const today = new Date().toISOString().split('T')[0];
        const enrolStart = new Date(config.enrolDate).toLocaleDateString('de-DE');
        const dealName = `Moodle Einschreibung – ${enrolStart} – ${endDateFormatted}`;
        const description = [
          `Institut: ${config.institute}`,
          `Datum: ${fmtDateTime(new Date())}`,
          `Einschreibezeitraum: ${new Date(config.enrolDate).toLocaleDateString('de-DE')} – ${endDateFormatted}`,
          '',
          `User: ${result.usersResolved} gesamt (${result.usersCreated} neu, ${result.usersResolved - result.usersCreated} bereits vorhanden)`,
          `Einschreibungen: ${result.enrolmentsDone}`,
          `Gruppen: ${result.groupsCreated}`,
          `Kohorte: ${result.cohortName}${result.cohortId ? ` (ID ${result.cohortId})` : ' – nicht gefunden'}`,
          '',
          `Kurse: ${activeMatrixCourses.map(c => c.shorthand).join(', ')}`,
          `Moodle-URL: ${config.moodleUrl}`,
        ].join('\n');
        await createZohoDeal(config, account.id, dealName, today, description);
        zohoDone = true;
      } catch (e) {
        console.error('[Zoho] CRM-Fehler:', e);
        postErrors.push(`CRM: ${e?.message || String(e)}`);
      }
    }

    // Fehler aus SP/CRM als Warnings anzeigen — Moodle war OK
    postErrors.forEach(msg => addToast(msg, 'warning', 8000));

    // ── Fertig ────────────────────────────────────────────────────────────
    const existing = result.usersResolved - result.usersCreated;
    const moodleSummary = [
      result.usersCreated > 0 && `${result.usersCreated} neue Accounts`,
      existing > 0 && `${existing} bereits vorhanden`,
      `${result.enrolmentsDone} Einschreibungen`,
      result.groupsCreated > 0 && `${result.groupsCreated} Gruppen`,
    ].filter(Boolean).join(' · ');
    setMoodleProgress({ label: 'Fertig', pct: 100, done: true, error: false });
    setMoodleResult({ moodleSummary, sharepoint: sharepointDone, zoho: zohoDone });
    setIsMoodleEnrolling(false);
    setTimeout(() => setMoodleProgress(null), 3000);
  }, [config, generatedData, activeMatrixCourses, classRows, getClassLabel, addToast, addExportEntry, downloadPDF, downloadExcel, zohoEnabled, endDateFormatted]);

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
      { id: 'moodle',   label: 'Moodle',         icon: <GraduationCap size={13} /> },
      { id: 'zoho',     label: 'Zoho CRM',       icon: <Users size={13} /> },
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
      <ModalShell C={C} maxW="max-w-3xl" zIndex={250}>
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
                className="flex items-center gap-1 px-2.5 py-2 text-[9px] font-bold uppercase tracking-wider border-b-2 transition-all hover:opacity-80 rounded-t-lg whitespace-nowrap">
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
                <Info>EBCL-Moodle verwaltet Moodle-Zugangsdaten für Partnerinstitute — in drei Modi: neue Accounts anlegen, bestehende Klassen aktualisieren oder beides in einem Durchlauf.</Info>
                <Info>Die Kursliste wird vom EBCL-Server geladen und lokal gecacht, sodass Generieren und Exportieren auch offline funktioniert.</Info>
              </HSection>
              <HSection icon={<RefreshCw size={14} />} title="Modus: Aktualisieren">
                {[
                  ['01', C.accent1, 'Institut wählen & Klassen laden',  'Institutsname eingeben, dann "Klassen laden" — die App fragt Moodle ab und listet alle vorhandenen Klassen des Instituts mit Mitgliederzahl.'],
                  ['02', C.accent1, 'Klassen auswählen',                'Checkboxen aktivieren welche Klassen aktualisiert werden sollen. Abgewählte Klassen erscheinen nicht in der Matrix.'],
                  ['03', C.accent1, 'Matrix prüfen',                    'Bereits eingeschriebene Kurse sind blau mit RefreshCw-Symbol vorbelegt und können nicht abgewählt werden. Weitere Kurse können ergänzt werden.'],
                  ['04', C.accent1, 'Generieren & Einschreiben',        '⌘G — lädt die echten Moodle-Mitglieder und schreibt sie erneut ein (Zeitraum & Kurse werden aktualisiert, Passwörter bleiben).'],
                ].map(([n, col, title, desc]) => (
                  <div key={n} className="flex gap-3 items-start">
                    <span style={{ color: col, backgroundColor: col + '18', borderColor: col + '30' }} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px] border mt-0.5">{n}</span>
                    <div><p style={{ color: C.text }} className="text-[11px] font-semibold mb-0.5">{title}</p><p style={{ color: C.muted }} className="text-[11px] leading-relaxed">{desc}</p></div>
                  </div>
                ))}
              </HSection>
              <HSection icon={<Plus size={14} />} title="Modus: Neu anlegen">
                {[
                  ['01', C.accent2, 'Institut & Klassenstruktur',       'Institutsname eingeben, Anzahl der Klassen pro Typ und Traineranzahl festlegen.'],
                  ['02', C.accent2, 'Kurse zuweisen',                   'In der Matrix jeder Klasse die gewünschten Kurse zuweisen. Alle auf einmal mit ⌘⇧A.'],
                  ['03', C.accent2, 'Generieren',                       '⌘G — App fragt Moodle ab um den höchsten bestehenden Account zu finden und nummeriert neue Accounts fortlaufend dahinter (keine Lücken, keine Duplikate).'],
                  ['04', C.accent2, 'Exportieren oder Einschreiben',    'CSV/PDF/Excel speichern, zu SharePoint hochladen oder direkt per REST API in Moodle einschreiben.'],
                ].map(([n, col, title, desc]) => (
                  <div key={n} className="flex gap-3 items-start">
                    <span style={{ color: col, backgroundColor: col + '18', borderColor: col + '30' }} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px] border mt-0.5">{n}</span>
                    <div><p style={{ color: C.text }} className="text-[11px] font-semibold mb-0.5">{title}</p><p style={{ color: C.muted }} className="text-[11px] leading-relaxed">{desc}</p></div>
                  </div>
                ))}
              </HSection>
              <HSection icon={<Zap size={14} />} title="Modus: Aktualisieren & Neu anlegen">
                <Info>Kombinierter Modus: bestehende Klassen aus Moodle laden und auswählen (werden aktualisiert) und gleichzeitig neue Klassen konfigurieren — alles in einem einzigen Einschreibelauf.</Info>
                <Info>Die Matrix zeigt zuerst die bestehenden Klassen (mit gesperrten Kursen), dann nach einem Trenner die neuen Klassen. Neue Accounts werden automatisch hinter den bestehenden nummeriert.</Info>
              </HSection>
              <HSection icon={<RefreshCw size={14} />} title="Was wird beim Neustart zurückgesetzt?">
                <Info>Bei jedem App-Start zurückgesetzt: Institutsname, Datum, Kursauswahl, Klassenzuweisungen und geladene Moodle-Gruppen.</Info>
                <Info><strong style={{ color: C.text }}>Erhalten bleiben:</strong> Export-History, Klassengrößen, Klassen-Namen, Passwörter, Backend-URLs, Moodle-Token, Modus und Dark Mode.</Info>
              </HSection>
            </div>
          )}

          {/* ── KONFIGURATION ── */}
          {helpTab === 'config' && (
            <div className="space-y-1">
              <HSection icon={<Building2 size={14} />} title="Organisation">
                <Row label="Institutsname" desc="Pflichtfeld. Leerzeichen bleiben erhalten und werden nur intern für Usernamen entfernt (z.B. 'Test Institut' → Username-Prefix 'testinstitut'). Im Zoho-Dropdown werden passende CRM-Accounts vorgeschlagen." />
                <Row label="Trainer" desc="Anzahl der Trainer-Accounts. Im Neu-anlegen-Modus werden neue Trainer generiert. Im Aktualisieren-Modus werden bestehende Trainer aus der Gruppe geladen." />
                <Row label="Kurs Anzahl" desc="Anzahl der Kursspalten in der Matrix (max. 8). Entspricht der Anzahl der Moodle-Kurse pro Institut." />
              </HSection>
              <HSection icon={<GraduationCap size={14} />} title="Modus">
                <Row label="Aktualisieren" desc="Klassen werden aus Moodle geladen. Bestehende Einschreibungen werden erkannt (RefreshCw) und beim Einschreiben aktualisiert (Zeitraum, Kurse — Passwörter bleiben unverändert)." />
                <Row label="Neu anlegen" desc="Neue Accounts werden generiert, fortlaufend nach den höchsten bestehenden Nummern in Moodle. Klassen, Schüler und Trainer werden automatisch weiternummeriert." />
                <Row label="Aktualisieren & Neu anlegen" desc="Kombinierter Modus: bestehende Klassen auswählen (werden aktualisiert) und zusätzlich neue Klassen konfigurieren — in einem Durchlauf." />
              </HSection>
              <HSection icon={<GraduationCap size={14} />} title="Klassen (je nach Modus)">
                <Row label="Klassen aus Moodle" desc="Im Aktualisieren- und Kombiniert-Modus: lädt alle Gruppen des Instituts aus Moodle (kursübergreifend). Mitgliederzahl wird angezeigt (ohne Trainer). Per Checkbox auswählen welche Klassen bearbeitet werden." />
                <Row label="Klassen Struktur" desc="Im Neu-anlegen-Modus und im Kombiniert-Modus (unterer Teil): Typen mit Klassengrößen und Anzahl definieren. Größen in den Einstellungen → Allgemein anpassbar." />
                <Row label="Klassengrößen editieren" desc="Bleistift-Icon neben der Schülerzahl zum einmaligen Anpassen. X-Button setzt auf Standardgröße zurück. Im Aktualisieren-Modus nicht verfügbar (echte Mitgliederzahl)." />
                <Row label="Klassen-Namen" desc="Optional in den Einstellungen anpassbar. Werden gespeichert und bleiben beim Neustart erhalten." />
              </HSection>
              <HSection icon={<ShieldCheck size={14} />} title="Zeit & Passwörter">
                <Row label="Einschreibedatum" desc="Ab wann die Accounts in Moodle aktiv sind. Wird beim App-Start auf heute gesetzt." />
                <Row label="Dauer (Tage)" desc="Gültigkeitsdauer der Einschreibung. Das Enddatum wird automatisch berechnet und ist direkt editierbar." />
                <Row label="PW Schüler / Trainer" desc="Standard-Passwörter für generierte Accounts. Werden in den Einstellungen gesetzt und bleiben beim Neustart erhalten." />
                <Row label="Auto-Passwort" desc="Generiert für jeden Account ein einzigartiges Zufalls-Passwort statt des Standard-Passworts." />
              </HSection>
            </div>
          )}

          {/* ── KURS-MATRIX ── */}
          {helpTab === 'matrix' && (
            <div>
              <HSection icon={<BookOpen size={14} />} title="Kurs-Pool">
                <Info>Die verfügbaren Kurse werden beim Start automatisch vom EBCL-Server (Power Automate) geladen und lokal gecacht — die App funktioniert damit auch offline.</Info>
                <Row label="Kurse laden" desc="Der 'Kurse laden'-Button synchronisiert den Pool manuell neu (z.B. wenn neue Kurse hinzugekommen sind)." />
                <Row label="Kurs-ID" desc="Jeder Kurs im Pool sollte eine numerische Moodle-ID haben (für die direkte Einschreibung via REST API). Alternativ wird die ID aus der Kurs-URL extrahiert." />
                <Row label="Vorschau" desc="'Kurs-Pool anzeigen' öffnet eine Übersicht aller gecachten Kurse mit Tag, ID, Shorthand und URL." />
              </HSection>
              <HSection icon={<Tag size={14} />} title="Kurs-Tags">
                <Info>Kurse können in der Datenbank mit einem Tag versehen werden (Spalte: tag / tags / kategorie / typ). Tags werden farbig im Dropdown und in der Vorschau angezeigt — nur für den Client sichtbar, kein Einfluss auf Moodle.</Info>
                <Row label="Farben" desc="Tags wählen ihre Farbe automatisch aus der Akzentpalette. Die 4 Akzentfarben können in den Einstellungen → Anpassen geändert werden." />
                <Row label="Suche" desc="Im Kurs-Dropdown kann nach Kursname, Kürzel und Tag gefiltert werden." />
                <Row label="Ausrichtung" desc="Die Tag-Spalte passt sich automatisch an den längsten Tag an — Kursnamen stehen dadurch immer auf gleicher Höhe." />
              </HSection>
              <HSection icon={<TableIcon size={14} />} title="Matrix-Bedienung">
                <Row label="Kurs auswählen" desc="Im Dropdown oben in jeder Spalte einen Kurs aus dem Pool wählen. Mit Suche und Tag-Filterung. Bis zu 8 Kursspalten gleichzeitig möglich." />
                <Row label="Zuweisung" desc="Per Klick bestimmen welche Klassen welche Kurse erhalten. Eine Klasse kann mehrere Kurse haben." />
                <Row label="Spalten-Alle-Button" desc="Der blaue 'Alle'-Button im Spaltenkopf weist alle (sichtbaren) Klassen diesem Kurs zu oder entfernt sie (Toggle)." />
                <Row label="Alle zuweisen (⌘⇧A)" desc="Weist alle Klassen allen aktiven Kursen zu — schnellster Weg wenn alle Klassen dieselben Kurse bekommen." />
              </HSection>
              <HSection icon={<RefreshCw size={14} />} title="Gesperrte Kurse (Aktualisieren-Modus)">
                <Info>Wenn Klassen aus Moodle geladen werden, erkennt die App automatisch in welchen Pool-Kursen die Klasse bereits eingeschrieben ist.</Info>
                <Row label="RefreshCw-Symbol (blau)" desc="Kurs ist bereits eingeschrieben — wird beim nächsten Einschreiben aktualisiert (Zeitraum, Passwort). Kann nicht abgewählt werden." />
                <Row label="Grünes Häkchen" desc="Kurs wurde manuell zusätzlich zugewiesen — wird neu eingeschrieben." />
                <Row label="Trennlinie 'Neue Klassen'" desc="Im 'Aktualisieren & Neu anlegen'-Modus trennt eine Linie bestehende (oben) von neuen Klassen (unten) in der Matrix." />
              </HSection>
              <HSection icon={<AlertTriangle size={14} />} title="Validierung">
                <Info>Beim Generieren prüft die App ob jede Klasse mindestens einem Kurs zugewiesen ist. Klassen ohne Zuweisung werden <span style={{ color: '#e11d48' }}>rot markiert</span>. Die Markierung verschwindet sobald eine Zuweisung gesetzt wird.</Info>
              </HSection>
            </div>
          )}

          {/* ── EXPORT ── */}
          {helpTab === 'export' && (
            <div>
              <HSection icon={<FileSpreadsheet size={14} />} title="CSV-Export (⌘E)">
                <Info>Moodle-Import-Format. Enthält alle generierten Accounts mit Einschreibungen und kann direkt in Moodle hochgeladen werden.</Info>
                <Row label="Inhalt" desc="Username, Passwort, Vorname, Nachname, E-Mail, Kohorte, Kurs-Shorthand, Gruppe, Rolle, Einschreibedauer." />
                <Row label="Kodierung" desc="UTF-8 mit BOM — für maximale Excel-Kompatibilität (Umlaute korrekt)." />
                <Row label="Moodle-Upload" desc="Administration → Nutzer → Nutzer hochladen → CSV-Datei auswählen." />
              </HSection>
              <HSection icon={<FileDown size={14} />} title="PDF-Export (⌘P)">
                <Info>Zugangsdaten-Dokument für das Institut. Pro Seite eine Klasse (oder Trainer-Liste) im Tabellenformat.</Info>
                <Row label="Format" desc="A4 Querformat — optimiert für viele Kursspalten nebeneinander." />
                <Row label="Kurslinks" desc="Kurs-Zellen im PDF sind klickbar — direkter Link zum jeweiligen Moodle-Kurs." />
                <Row label="Leitfaden" desc="Optional: erste Seite enthält einen Leitfaden für das Institut (ein-/ausschaltbar in den Einstellungen)." />
              </HSection>
              <HSection icon={<TableIcon size={14} />} title="Excel-Export">
                <Info>Identische Inhalte wie die CSV, aber im XLSX-Format mit Tabellenformatierung — übersichtlicher für manuelle Bearbeitung.</Info>
                <Row label="Dateiname" desc="Zugangsdaten-{Institut}_{Jahr}.xlsx" />
              </HSection>
              <HSection icon={<Upload size={14} />} title="SharePoint-Upload">
                <Info>Lädt CSV, PDF und Excel gleichzeitig in einen neuen SharePoint-Ordner hoch (via Power Automate Flow).</Info>
                <Row label="Ordner" desc="Wird automatisch benannt: {Institut}_{Datum} — jedes Institut bekommt seinen eigenen Ordner." />
                <Row label="Konfiguration" desc="Die Flow-URL wird in den Einstellungen → Backend hinterlegt und bleibt gespeichert." />
                <Row label="Moodle-Upload" desc="Nach einer Moodle-Einschreibung wird statt der CSV eine TXT-Zusammenfassung hochgeladen (+ PDF + Excel)." />
              </HSection>
              <HSection icon={<ClipboardList size={14} />} title="Daten-Vorschau & History">
                <Row label="Vorschau" desc="'Daten-Vorschau' zeigt nach dem Generieren alle Accounts in einer scrollbaren Tabelle zur Kontrolle vor dem Export." />
                <Row label="Export-History" desc="Jeder Export wird mit Zeitstempel und Accountanzahl gespeichert und bleibt nach dem Neustart erhalten." />
              </HSection>
            </div>
          )}

          {/* ── MOODLE ── */}
          {helpTab === 'moodle' && (
            <div>
              <HSection icon={<GraduationCap size={14} />} title="Direkte Moodle-Einschreibung">
                <Info>Accounts werden direkt über die Moodle REST API angelegt und eingeschrieben — kein CSV-Umweg nötig. Funktioniert in allen drei Modi (Neu anlegen, Aktualisieren, Beides). Das Ergebnis wird danach automatisch zu SharePoint hochgeladen.</Info>
                <Row label="Neu anlegen" desc="Erstellt neue Accounts fortlaufend nach den höchsten bestehenden Nummern. Schüler, Trainer und Klassen werden automatisch weiternummeriert." />
                <Row label="Aktualisieren" desc="Lädt echte Moodle-Mitglieder der ausgewählten Gruppen und schreibt sie neu ein — Zeitraum und Kurse werden aktualisiert, Passwörter bleiben unverändert. Neue Accounts werden dabei nicht erstellt." />
                <Row label="Aktualisieren & Neu anlegen" desc="Kombinierter Durchlauf: bestehende Gruppen werden aktualisiert, neue Accounts für neue Klassen werden erstellt und eingeschrieben." />
              </HSection>
              <HSection icon={<Settings size={14} />} title="Einrichtung">
                <Row label="Moodle URL" desc="Die Basis-URL deiner Moodle-Instanz, z.B. https://moodle.ebcl.at. Einstellungen → Backend." />
                <Row label="API-Token" desc="Web Service Token aus Moodle: Administration → Website-Administration → Plugins → Web Services → Token verwalten." />
                <Row label="Token-Sichtbarkeit" desc="Der Token kann über das Auge-Symbol in den Einstellungen ein-/ausgeblendet werden." />
              </HSection>
              <HSection icon={<ShieldCheck size={14} />} title="Erforderliche Moodle-Berechtigungen">
                <Info>Der Web Service Token benötigt folgende Funktionen (External Service konfigurieren):</Info>
                <div className="space-y-1">
                  {[
                    'core_user_create_users',
                    'core_user_get_users',
                    'core_user_get_users_by_field',
                    'enrol_manual_enrol_users',
                    'core_group_create_groups',
                    'core_group_get_course_groups',
                    'core_group_get_group_members',
                    'core_group_add_group_members',
                    'core_cohort_create_cohorts',
                    'core_cohort_search_cohorts',
                    'core_cohort_add_cohort_members',
                    'core_enrol_get_enrolled_users',
                    'core_enrol_get_users_courses',
                    'core_course_get_courses',
                  ].map(fn => (
                    <div key={fn} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="px-3 py-1.5 rounded-lg border font-mono text-[10px]" >
                      <span style={{ color: C.accent1 }}>{fn}</span>
                    </div>
                  ))}
                </div>
              </HSection>
              <HSection icon={<Zap size={14} />} title="Was passiert beim Einschreiben? (Neu anlegen)">
                {[
                  ['1', 'Offset ermitteln', 'Die App sucht die höchsten bestehenden Account-Nummern in Moodle und nummeriert neue Accounts lückenlos dahinter.'],
                  ['2', 'Nutzer prüfen',    'Bestehende Accounts werden per Username gesucht und wiederverwendet — keine Duplikate.'],
                  ['3', 'Nutzer anlegen',   'Nur wirklich neue Accounts werden erstellt. Bei Fehler wird auf Einzel-Anlage zurückgefallen.'],
                  ['4', 'Gruppen',          'Pro Kurs und Klasse wird eine Gruppe angelegt (falls nicht vorhanden). Trainer werden zu allen Gruppen hinzugefügt.'],
                  ['5', 'Einschreiben',     'Schüler → zugewiesene Kurse (Rolle: Student). Trainer → alle Kurse (Rolle: Lehrbeauftragter). Mit Einschreibezeitraum.'],
                  ['6', 'Kohorte',          'Eine System-Kohorte mit dem Institutsnamen wird angelegt oder gefunden. Alle Accounts werden zugeordnet.'],
                ].map(([n, title, desc]) => (
                  <div key={n} className="flex gap-3 items-start">
                    <span style={{ color: '#7C3AED', backgroundColor: '#7C3AED18', borderColor: '#7C3AED30' }} className="w-6 h-6 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px] border mt-0.5">{n}</span>
                    <div><p style={{ color: C.text }} className="text-[11px] font-semibold mb-0.5">{title}</p><p style={{ color: C.muted }} className="text-[11px] leading-relaxed">{desc}</p></div>
                  </div>
                ))}
              </HSection>
              <HSection icon={<RefreshCw size={14} />} title="Was passiert beim Einschreiben? (Aktualisieren)">
                {[
                  ['1', 'Klassen laden',    'Die App lädt alle Gruppen des Instituts kursübergreifend aus Moodle (core_group_get_course_groups).'],
                  ['2', 'Mitglieder laden', 'Für jede ausgewählte Gruppe werden die echten Moodle-Mitglieder geladen (core_enrol_get_enrolled_users). Trainer werden automatisch erkannt.'],
                  ['3', 'Kurse kombinieren','Bereits eingeschriebene Kurse (gesperrt) und manuell ergänzte Kurse werden zusammengeführt.'],
                  ['4', 'Einschreiben',     'Alle Mitglieder werden erneut in alle ihre Kurse eingeschrieben — Zeitraum wird dabei aktualisiert, Passwörter bleiben unverändert.'],
                ].map(([n, title, desc]) => (
                  <div key={n} className="flex gap-3 items-start">
                    <span style={{ color: C.accent1, backgroundColor: C.accent1 + '18', borderColor: C.accent1 + '30' }} className="w-6 h-6 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px] border mt-0.5">{n}</span>
                    <div><p style={{ color: C.text }} className="text-[11px] font-semibold mb-0.5">{title}</p><p style={{ color: C.muted }} className="text-[11px] leading-relaxed">{desc}</p></div>
                  </div>
                ))}
              </HSection>
              <HSection icon={<AlertTriangle size={14} />} title="Wichtige Hinweise">
                <Row label="Bestätigung" desc="Vor jeder Einschreibung erscheint ein Bestätigungs-Popup — als Schutz vor versehentlichem Ausführen." />
                <Row label="Kurs-ID" desc="Jeder Kurs im Pool muss eine numerische Moodle-ID haben, sonst wird er beim Einschreiben übersprungen." />
                <Row label="Fehleranalyse" desc="Bei Problemen die Entwicklerkonsole öffnen (Einstellungen → Allgemein → Entwicklerkonsole). Alle API-Aufrufe werden dort geloggt." />
                <Row label="SharePoint" desc="Nach erfolgreicher Einschreibung wird eine TXT-Zusammenfassung (+ PDF + Excel) automatisch zu SharePoint hochgeladen." />
              </HSection>
            </div>
          )}

          {/* ── ZOHO CRM ── */}
          {helpTab === 'zoho' && (
            <div>
              <HSection icon={<Users size={14} />} title="Zoho CRM Integration">
                <Info>Nach jeder erfolgreichen Moodle-Einschreibung wird automatisch ein Abschluss (Deal) im Zoho CRM beim jeweiligen Account hinterlegt. Die Integration ist aktiv sobald Client ID, Client Secret und Refresh Token gesetzt sind.</Info>
              </HSection>
              <HSection icon={<Settings size={14} />} title="Einrichtung (Einstellungen → Backend)">
                <Row label="Client ID" desc="OAuth-App Client ID aus der Zoho API Console (api-console.zoho.eu)." />
                <Row label="Client Secret" desc="OAuth-App Client Secret, ebenfalls aus der API Console." />
                <Row label="Refresh Token" desc="Langlebiges Token für den API-Zugriff. Wird über den 'Token generieren'-Dialog erstellt." />
                <Row label="Token generieren" desc="Öffnet ein Popup mit Anleitung. Grant Code aus der Zoho API Console einfügen → automatischer Austausch gegen einen Refresh Token." />
              </HSection>
              <HSection icon={<ShieldCheck size={14} />} title="Erforderliche OAuth-Scopes">
                <Info>Beim Generieren des Grant Codes in der Zoho API Console folgende Scopes angeben:</Info>
                <div className="space-y-1">
                  {['ZohoCRM.modules.Accounts.READ','ZohoCRM.modules.Accounts.CREATE','ZohoCRM.modules.Deals.CREATE'].map(s => (
                    <div key={s} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="px-3 py-1.5 rounded-lg border font-mono text-[10px]">
                      <span style={{ color: C.accent1 }}>{s}</span>
                    </div>
                  ))}
                </div>
              </HSection>
              <HSection icon={<Zap size={14} />} title="Was passiert beim Einschreiben?">
                {[
                  ['1', 'Account suchen',  'Der Institutsname wird exakt im CRM gesucht. Groß-/Kleinschreibung wird berücksichtigt.'],
                  ['2', 'Account anlegen', 'Wird kein Account gefunden, wird automatisch einer angelegt — mit dem Tag "Institut".'],
                  ['3', 'Deal erstellen',  'Ein Deal mit Name "Moodle Einschreibung – [Startdatum] – [Enddatum]" wird beim Account hinterlegt. Stage: Closed Won.'],
                  ['4', 'Beschreibung',    'Der Deal enthält alle Einschreibedetails: Institut, Einschreibezeitraum, Nutzeranzahl, Gruppen, Kohorte, Kurse und Moodle-URL.'],
                ].map(([n, title, desc]) => (
                  <div key={n} className="flex gap-3 items-start">
                    <span style={{ color: '#16a34a', backgroundColor: '#16a34a18', borderColor: '#16a34a30' }} className="w-6 h-6 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px] border mt-0.5">{n}</span>
                    <div><p style={{ color: C.text }} className="text-[11px] font-semibold mb-0.5">{title}</p><p style={{ color: C.muted }} className="text-[11px] leading-relaxed">{desc}</p></div>
                  </div>
                ))}
              </HSection>
              <HSection icon={<Building2 size={14} />} title="Institutübersicht">
                <Row label="CRM-Dropdown" desc="Das Institutsfeld zeigt beim Tippen automatisch passende CRM-Accounts als Vorschlag. Auswahl übernimmt den Namen direkt." />
                <Row label="Institutübersicht" desc="Der Button unter 'Kursübersicht' öffnet eine durchsuchbare Liste aller CRM-Accounts. Über 'Übernehmen' wird ein Account als Institut gesetzt." />
                <Row label="Leerzeichen" desc="Leerzeichen im Institutsnamen werden automatisch zu Bindestrichen umgewandelt (Moodle-Kompatibilität)." />
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
                <Row label="Modus wählen" desc="'Aktualisieren' für Folgesessions mit bestehenden Klassen. 'Neu anlegen' für erste Einschreibung. 'Beides' wenn neue Klassen zu einem Institut dazukommen." />
                <Row label="Klassen schnell laden" desc="Im Aktualisieren-Modus erst Institut eingeben, dann 'Klassen laden' — die App findet automatisch alle Gruppen in allen Kursen." />
                <Row label="Gesperrte Kurse" desc="RefreshCw (blau) = bereits eingeschrieben, wird aktualisiert. Grünes Häkchen = manuell ergänzt, wird neu eingeschrieben." />
                <Row label="Offset-Erkennung" desc="Im Neu-anlegen-Modus wird Moodle abgefragt um die höchsten bestehenden Nummern zu finden — neue Accounts starten automatisch dahinter." />
                <Row label="Alle zuweisen" desc="Bei gleicher Kursstruktur für alle Klassen: ⌘⇧A spart alle Einzelklicks in der Matrix." />
                <Row label="Favoriten" desc="Häufig verwendete Institute als Favorit speichern. Das Datum wird beim Laden automatisch auf heute gesetzt." />
                <Row label="Offline arbeiten" desc="Der Kurs-Pool wird gecacht. Generieren und Exportieren funktioniert ohne Internet — solange der Cache vorhanden ist." />
                <Row label="Daten-Vorschau" desc="Vor dem Export alle Accounts nochmal prüfen — insbesondere im Aktualisieren-Modus die korrekten Moodle-Mitglieder verifizieren." />
              </HSection>
              <HSection icon={<RefreshCw size={14} />} title="Updates">
                <Info>Die App prüft beim Start automatisch auf neue Versionen. Bei einem verfügbaren Update erscheint ein Popup in der Bildschirmmitte. Nach dem Download startet die App automatisch neu.</Info>
                <Row label="Manuell prüfen" desc="Einstellungen → Allgemein → 'Jetzt auf Updates prüfen'." />
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

  const renderSettingsModal = () => {
    const STABS = [
      { id: 'allgemein', label: 'Allgemein', icon: <Settings size={12} /> },
      { id: 'klassen',   label: 'Klassen',   icon: <GraduationCap size={12} /> },
      { id: 'kurse',     label: 'Kurse',     icon: <BookOpen size={12} /> },
      { id: 'anpassen',  label: 'Anpassen',  icon: <Tag size={12} /> },
      { id: 'backend',   label: 'Backend',   icon: <Upload size={12} /> },
    ];
    return (
      <ModalShell C={C} maxW="max-w-lg">
        <ModalHeader C={C} icon={<Settings size={18} />} title="System-Settings" onClose={() => setActiveModal(null)} />
        {/* Tab-Bar */}
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="flex gap-1 px-4 py-2 border-b shrink-0">
          {STABS.map(t => (
            <button key={t.id} onClick={() => setSettingsTab(t.id)}
              style={{ backgroundColor: settingsTab === t.id ? C.card : 'transparent', color: settingsTab === t.id ? C.text : C.muted, borderColor: settingsTab === t.id ? C.border : 'transparent' }}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border text-[10px] font-bold uppercase transition-all">
              {t.icon}{t.label}
            </button>
          ))}
        </div>
        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          {/* ── Allgemein ─────────────────────────────────────────── */}
          {settingsTab === 'allgemein' && <>
            <div>
              <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-3">{darkMode ? <Moon size={14} /> : <Sun size={14} />} Darstellung</h4>
              <button onClick={() => setDarkMode(d => !d)} style={{ backgroundColor: C.subtle, borderColor: C.border, color: C.text }} className="w-full flex items-center justify-between px-4 py-3 rounded-xl border shadow-sm hover:opacity-80 transition-all">
                <span className="text-[11px] font-semibold">{darkMode ? 'Dunkelmodus aktiv' : 'Hellmodus aktiv'}</span>
                <span style={{ backgroundColor: darkMode ? C.main : C.accent1 }} className="flex items-center gap-1.5 text-white text-[10px] font-bold px-2.5 py-1 rounded-lg">
                  {darkMode ? <Sun size={11} /> : <Moon size={11} />} {darkMode ? 'Hell' : 'Dunkel'}
                </span>
              </button>
            </div>
            <div>
              <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-3"><FileDown size={14} /> PDF-Export</h4>
              <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="flex items-center justify-between px-4 py-3 rounded-xl border shadow-sm">
                <div>
                  <span style={{ color: C.text }} className="text-[11px] font-semibold">Leitfaden (Seite 1)</span>
                  <p style={{ color: C.muted }} className="text-[10px] mt-0.5 opacity-70">Trainer-Anleitung als erste PDF-Seite ein-/ausblenden</p>
                </div>
                <button onClick={() => setConfig(p => ({ ...p, showLeitfaden: !p.showLeitfaden }))}
                  style={{ backgroundColor: config.showLeitfaden ? C.accent2 : C.border }}
                  className="relative flex items-center w-10 h-5 rounded-full transition-colors shrink-0 ml-4">
                  <span style={{ transform: config.showLeitfaden ? 'translateX(21px)' : 'translateX(2px)' }} className="absolute w-3.5 h-3.5 bg-white rounded-full shadow transition-transform" />
                </button>
              </div>
            </div>
            <div>
              <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-3"><RefreshCw size={14} /> Software-Update</h4>
              <button onClick={handleManualUpdateCheck} disabled={isCheckingUpdate || !!pendingUpdate}
                style={{ borderColor: C.border, color: pendingUpdate ? C.accent2 : C.text, backgroundColor: C.card }}
                className="w-full flex items-center justify-center gap-2 px-4 py-2.5 border rounded-xl text-[11px] font-bold hover:opacity-80 disabled:opacity-50 transition-all">
                {isCheckingUpdate ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />}
                {isCheckingUpdate ? 'Suche nach Updates…' : pendingUpdate ? `Update v${pendingUpdate.version} verfügbar` : 'Jetzt auf Updates prüfen'}
              </button>
            </div>
          </>}
          {/* ── Klassen ───────────────────────────────────────────── */}
          {settingsTab === 'klassen' && <>
            <div>
              <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-1"><ShieldCheck size={14} /> Standard-Einschreibedauer</h4>
              <p style={{ color: C.muted }} className="text-[10px] mb-3 opacity-60">Wird beim App-Start und beim Laden von Favoriten als Einschreibedauer gesetzt.</p>
              <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors flex items-center gap-3">
                <input
                  type="number" min="1" max="3650"
                  value={config.defaultEnrolPeriod}
                  onChange={e => {
                    const v = Math.max(1, parseInt(e.target.value, 10) || 1);
                    setConfig(p => ({ ...p, defaultEnrolPeriod: v, enrolPeriod: v }));
                  }}
                  style={{ color: C.main }}
                  className="w-20 bg-transparent text-base font-bold outline-none"
                />
                <span style={{ color: C.muted }} className="text-[11px]">Tage</span>
              </div>
            </div>
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
          </>}
          {/* ── Kurse ─────────────────────────────────────────────── */}
          {settingsTab === 'kurse' && <>
            {/* Toggle */}
            <div>
              <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-2"><BookOpen size={14} /> Kursquelle</h4>
              <div style={{ borderColor: C.border }} className="flex rounded-xl border overflow-hidden">
                <button
                  onClick={() => setConfig(p => ({ ...p, moodleBetaEnabled: false }))}
                  style={{ backgroundColor: !config.moodleBetaEnabled ? C.accent1 : C.subtle, color: !config.moodleBetaEnabled ? '#fff' : C.muted }}
                  className="flex-1 py-2 text-[10px] font-bold uppercase tracking-wide flex items-center justify-center gap-1.5 transition-all"
                >
                  <Upload size={11} /> Excel-Liste (Power Automate)
                </button>
                <button
                  onClick={() => setConfig(p => ({ ...p, moodleBetaEnabled: true }))}
                  style={{ backgroundColor: config.moodleBetaEnabled ? C.accent1 : C.subtle, color: config.moodleBetaEnabled ? '#fff' : C.muted }}
                  className="flex-1 py-2 text-[10px] font-bold uppercase tracking-wide flex items-center justify-center gap-1.5 transition-all"
                >
                  <GraduationCap size={11} /> Moodle direkt
                </button>
              </div>
            </div>

            {!config.moodleBetaEnabled ? (
              /* ── Power Automate / Excel ── */
              <div className="space-y-3">
                <p style={{ color: C.muted }} className="text-[10px] opacity-60">Kursliste wird über Power Automate abgerufen (URL in Einstellungen → Backend).</p>
                <button
                  onClick={fetchCoursePool}
                  disabled={isLoadingPool || !config.courseApiUrl?.trim()}
                  style={{ color: C.accent1, borderColor: C.accent1 + '40' }}
                  className="flex items-center gap-2 px-3 py-1.5 rounded-lg border text-[10px] font-bold uppercase hover:opacity-80 disabled:opacity-40 transition-all"
                >
                  <RefreshCw size={12} className={isLoadingPool ? 'animate-spin' : ''} />
                  Kursliste neu laden
                </button>
                {/* Vorschau aller gecachten Kurse */}
                {courseDictionary.length > 0 && (
                  <div>
                    <p style={{ color: C.muted }} className="text-[9px] font-bold uppercase tracking-widest mb-1.5">{courseDictionary.length} Kurse im Cache</p>
                    <div style={{ borderColor: C.border }} className="border rounded-xl overflow-auto max-h-64">
                      {courseDictionary.map((c, i) => (
                        <div key={c.id} style={{ borderColor: C.border }} className="flex items-center gap-3 px-3 py-2 border-b last:border-0 hover:bg-black/5">
                          <span style={{ color: C.muted }} className="text-[9px] font-mono w-6 shrink-0 text-right">{i + 1}</span>
                          {c.tag ? <span style={{ backgroundColor: getTagColor(c.tag), color: '#fff' }} className="text-[8px] font-bold px-1.5 py-0.5 rounded shrink-0">{c.tag}</span> : null}
                          <span style={{ color: C.text }} className="text-[11px] flex-1 font-medium truncate">{c.label}</span>
                          <span style={{ color: C.muted }} className="text-[9px] font-mono shrink-0">{c.shorthand}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
                {courseDictionary.length === 0 && !isLoadingPool && (
                  <p style={{ color: C.muted }} className="text-[10px] italic opacity-60">Kein Cache — URL eingeben und „Kursliste neu laden" klicken.</p>
                )}
              </div>
            ) : (
              /* ── Moodle direkt ── */
              <div className="space-y-3">
                <p style={{ color: C.muted }} className="text-[10px] opacity-60">Kurse werden direkt von Moodle abgerufen. Wähle aus, welche in der Matrix erscheinen sollen.</p>
                <div className="flex items-center gap-2">
                  <button
                    onClick={fetchMoodleCoursesCatalog}
                    disabled={isLoadingMoodleCourses}
                    style={{ color: C.accent2, borderColor: C.accent2 + '40' }}
                    className="flex items-center gap-2 px-3 py-1.5 rounded-lg border text-[10px] font-bold uppercase hover:opacity-80 disabled:opacity-50 transition-all"
                  >
                    <RefreshCw size={12} className={isLoadingMoodleCourses ? 'animate-spin' : ''} />
                    Kurse laden
                  </button>
                  {allMoodleCourses.length > 0 && (
                    <span style={{ color: C.muted }} className="text-[10px]">
                      {allMoodleCourses.length} verfügbar · <strong style={{ color: C.text }}>{(config.moodleBetaCourseIds ?? []).length} ausgewählt</strong>
                    </span>
                  )}
                </div>
                {allMoodleCourses.length > 0 && (
                  <>
                    <input
                      type="text"
                      value={moodleBetaSearch}
                      onChange={e => setMoodleBetaSearch(e.target.value)}
                      placeholder="Kurs suchen…"
                      style={{ backgroundColor: C.subtle, borderColor: C.border, color: C.text }}
                      className="w-full px-3 py-2 rounded-lg border text-[11px] outline-none placeholder:opacity-40"
                    />
                    <div style={{ borderColor: C.border }} className="border rounded-xl overflow-auto max-h-60">
                      {(() => {
                        const search = moodleBetaSearch.toLowerCase();
                        const filtered = allMoodleCourses.filter(c =>
                          !search || c.label.toLowerCase().includes(search) || c.tag.toLowerCase().includes(search) || c.id.includes(search)
                        );
                        const groups = {};
                        filtered.forEach(c => { const cat = c.tag || 'Ohne Kategorie'; if (!groups[cat]) groups[cat] = []; groups[cat].push(c); });
                        return Object.entries(groups).map(([cat, courses]) => (
                          <div key={cat}>
                            <div style={{ backgroundColor: C.subtle, color: C.muted, borderColor: C.border }} className="px-3 py-1.5 text-[9px] font-bold uppercase tracking-widest border-b sticky top-0">
                              {cat} <span className="opacity-50">({courses.length})</span>
                            </div>
                            {courses.map(c => {
                              const selected = (config.moodleBetaCourseIds ?? []).includes(c.id);
                              return (
                                <button key={c.id}
                                  onClick={() => { const ids = config.moodleBetaCourseIds ?? []; setConfig(p => ({ ...p, moodleBetaCourseIds: selected ? ids.filter(id => id !== c.id) : [...ids, c.id] })); }}
                                  style={{ borderColor: C.border, backgroundColor: selected ? (C.accent2 + '15') : 'transparent' }}
                                  className="w-full flex items-center gap-3 px-3 py-2 border-b last:border-0 hover:bg-black/5 transition-colors text-left"
                                >
                                  <span style={{ color: selected ? C.accent2 : C.muted }}>{selected ? <CheckSquare size={14} /> : <Square size={14} />}</span>
                                  <span style={{ color: C.muted }} className="text-[9px] font-mono w-10 shrink-0">#{c.id}</span>
                                  <span style={{ color: C.text }} className="text-[11px] flex-1 font-medium">{c.label}</span>
                                  <span style={{ color: C.muted }} className="text-[9px] font-mono">{c.shorthand}</span>
                                </button>
                              );
                            })}
                          </div>
                        ));
                      })()}
                    </div>
                  </>
                )}
                {allMoodleCourses.length === 0 && !isLoadingMoodleCourses && (
                  <p style={{ color: C.muted }} className="text-[10px] italic opacity-60">Noch keine Kurse geladen — klicke auf „Kurse laden".</p>
                )}
              </div>
            )}
          </>}

          {/* ── Backend ───────────────────────────────────────────── */}
          {settingsTab === 'backend' && <>
            <div>
              <div className="flex items-center justify-between mb-1">
                <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2"><Upload size={14} /> Power Automate</h4>
              </div>
              <p style={{ color: C.muted }} className="text-[10px] mb-3 opacity-60">URLs für Kursliste und SharePoint-Export.</p>
              <div className="space-y-3">
                {[
                  { label: 'Kursliste URL (POST)', name: 'courseApiUrl', ph: 'https://prod-xx.logic.azure.com/…' },
                  { label: 'SharePoint Export', name: 'sharepointUrl', ph: 'https://prod-xx.logic.azure.com/…' },
                ].map(f => (
                  <div key={f.name} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                    <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase block mb-1.5">{f.label}</label>
                    <input type="text" value={config[f.name]}
                      onChange={e => setConfig(p => ({ ...p, [f.name]: e.target.value }))}
                      placeholder={f.ph}
                      style={{ color: C.text, backgroundColor: 'transparent' }}
                      className="w-full text-[10px] font-mono outline-none placeholder:opacity-30" />
                  </div>
                ))}
              </div>
            </div>
            <div>
              <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-1">
                <GraduationCap size={14} /> Moodle REST API
              </h4>
              <p style={{ color: C.muted }} className="text-[10px] mb-3 opacity-60">Token benötigt Berechtigungen für: core_user_create_users, enrol_manual_enrol_users, core_group_create_groups u.a.</p>
              <div className="space-y-3">
                <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                  <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase block mb-1.5">Moodle URL</label>
                  <input type="text" value={config.moodleUrl}
                    onChange={e => setConfig(p => ({ ...p, moodleUrl: e.target.value }))}
                    placeholder="https://moodle.schule.at"
                    style={{ color: C.text, backgroundColor: 'transparent' }}
                    className="w-full text-[10px] font-mono outline-none placeholder:opacity-30" />
                </div>
                <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                  <div className="flex items-center justify-between mb-1.5">
                    <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase">API-Token</label>
                    <button onClick={() => setShowMoodleToken(v => !v)} style={{ color: C.muted }} className="hover:opacity-70 transition-opacity">
                      {showMoodleToken ? <EyeOff size={12} /> : <Eye size={12} />}
                    </button>
                  </div>
                  <input type={showMoodleToken ? 'text' : 'password'} value={config.moodleToken}
                    onChange={e => setConfig(p => ({ ...p, moodleToken: e.target.value }))}
                    placeholder="abc123def456…"
                    style={{ color: C.text, backgroundColor: 'transparent' }}
                    className="w-full text-[10px] font-mono outline-none placeholder:opacity-30" />
                </div>
              </div>
            </div>
            <div className="mt-4">
              <div className="flex items-center justify-between mb-1">
                <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2">
                  <Users size={14} /> Zoho CRM
                </h4>
              </div>
              <p style={{ color: C.muted }} className="text-[10px] mb-3 opacity-60">Nach Moodle-Einschreibung wird automatisch eine Notiz im CRM hinterlegt. Aktiv sobald alle drei Felder befüllt sind.</p>
              <div className="space-y-3">
                <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                  <div className="flex items-center justify-between mb-1.5">
                    <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase">Client ID</label>
                    <button onClick={() => setShowZohoClientId(v => !v)} style={{ color: C.muted }} className="hover:opacity-70 transition-opacity">
                      {showZohoClientId ? <EyeOff size={12} /> : <Eye size={12} />}
                    </button>
                  </div>
                  <input type={showZohoClientId ? 'text' : 'password'} value={config.zohoClientId}
                    onChange={e => setConfig(p => ({ ...p, zohoClientId: e.target.value }))}
                    placeholder="1000.XXXXXXXX"
                    style={{ color: C.text, backgroundColor: 'transparent' }}
                    className="w-full text-[10px] font-mono outline-none placeholder:opacity-30" />
                </div>
                <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                  <div className="flex items-center justify-between mb-1.5">
                    <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase">Client Secret</label>
                    <button onClick={() => setShowZohoClientSecret(v => !v)} style={{ color: C.muted }} className="hover:opacity-70 transition-opacity">
                      {showZohoClientSecret ? <EyeOff size={12} /> : <Eye size={12} />}
                    </button>
                  </div>
                  <input type={showZohoClientSecret ? 'text' : 'password'} value={config.zohoClientSecret}
                    onChange={e => setConfig(p => ({ ...p, zohoClientSecret: e.target.value }))}
                    placeholder="abc123…"
                    style={{ color: C.text, backgroundColor: 'transparent' }}
                    className="w-full text-[10px] font-mono outline-none placeholder:opacity-30" />
                </div>
                <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                  <div className="flex items-center justify-between mb-1.5">
                    <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase">Refresh Token</label>
                    <button onClick={() => setShowZohoRefreshToken(v => !v)} style={{ color: C.muted }} className="hover:opacity-70 transition-opacity">
                      {showZohoRefreshToken ? <EyeOff size={12} /> : <Eye size={12} />}
                    </button>
                  </div>
                  <input type={showZohoRefreshToken ? 'text' : 'password'} value={config.zohoRefreshToken}
                    onChange={e => setConfig(p => ({ ...p, zohoRefreshToken: e.target.value }))}
                    placeholder="1000.XXXXXXXX…"
                    style={{ color: C.text, backgroundColor: 'transparent' }}
                    className="w-full text-[10px] font-mono outline-none placeholder:opacity-30" />
                </div>
                <div className="flex items-center justify-between">
                  {config.zohoRefreshToken
                    ? <p style={{ color: '#16a34a' }} className="text-[9px] font-semibold flex items-center gap-1"><CheckCircle2 size={10} /> Refresh Token vorhanden</p>
                    : <p style={{ color: C.muted }} className="text-[9px] opacity-60">Kein Refresh Token gesetzt.</p>
                  }
                  <button
                    onClick={() => setShowZohoTokenModal(true)}
                    style={{ borderColor: C.border, color: C.text }}
                    className="border px-3 py-1.5 rounded-lg text-[9px] font-bold uppercase hover:opacity-70 active:scale-95 transition-all flex items-center gap-1.5"
                  >
                    <Zap size={11} /> Token generieren
                  </button>
                </div>
              </div>
            </div>

            <div className="mt-4">
              <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-1">
                <ClipboardList size={14} /> Diagnose-Protokoll
              </h4>
              <p style={{ color: C.muted }} className="text-[10px] mb-3 opacity-60">Hält fest, was die App mit Moodle, Zoho und Power Automate macht — mit den genauen Fehlermeldungen. Zugangsdaten, Passwörter und E-Mail-Adressen werden vorher entfernt. Bei Problemen kopieren oder speichern und an den Support schicken.</p>
              <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border shadow-sm flex items-center justify-between gap-2">
                <p style={{ color: C.muted }} className="text-[9px]">{entryCount()} Einträge (letzte 14 Tage, auch nach Neustart)</p>
                <div className="flex items-center gap-2">
                  <button onClick={handleCopyLog} style={{ borderColor: C.border, color: C.text }} className="border px-3 py-1.5 rounded-lg text-[9px] font-bold uppercase hover:opacity-70 active:scale-95 transition-all flex items-center gap-1.5">
                    <ClipboardList size={11} /> Kopieren
                  </button>
                  <button onClick={handleSaveLog} style={{ borderColor: C.border, color: C.text }} className="border px-3 py-1.5 rounded-lg text-[9px] font-bold uppercase hover:opacity-70 active:scale-95 transition-all flex items-center gap-1.5">
                    <FileDown size={11} /> Speichern
                  </button>
                  <button onClick={handleClearLog} title="Alle Einträge löschen" style={{ borderColor: C.border, color: C.muted }} className="border px-3 py-1.5 rounded-lg text-[9px] font-bold uppercase hover:opacity-70 active:scale-95 transition-all flex items-center gap-1.5">
                    <Trash2 size={11} /> Leeren
                  </button>
                </div>
              </div>
            </div>
          </>}

          {/* ── Anpassen ──────────────────────────────────────────── */}
          {settingsTab === 'anpassen' && <>
            {/* Tag-Zuordnung */}
            {(() => {
              const uniqueTags = [...new Set(courseDictionary.map(c => c.tag).filter(Boolean))].sort();
              if (!uniqueTags.length) return null;
              const palette = config.customAccents?.length ? config.customAccents : DEFAULT_ACCENTS;
              return (
                <div>
                  <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-1">
                    <Tag size={14} /> Tag-Farben
                  </h4>
                  <p style={{ color: C.muted }} className="text-[10px] mb-3 opacity-60">Wähle für jeden Tag eine der 4 Akzentfarben.</p>
                  <div className="space-y-2">
                    {uniqueTags.map(tag => {
                      const currentIdx = config.tagColorMap?.[tag] ?? null;
                      return (
                        <div key={tag} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border flex items-center gap-3">
                          <span style={{ backgroundColor: getTagColor(tag), color: '#fff', minWidth: '60px' }} className="text-[8px] font-bold px-2 py-1 rounded text-center shrink-0">{tag}</span>
                          <div className="flex gap-2 ml-auto">
                            {palette.map((color, i) => (
                              <button
                                key={i}
                                onClick={() => setConfig(p => ({ ...p, tagColorMap: { ...p.tagColorMap, [tag]: i } }))}
                                style={{ backgroundColor: color, outline: currentIdx === i ? `2px solid ${color}` : 'none', outlineOffset: '2px', opacity: currentIdx === i ? 1 : 0.4 }}
                                className="w-6 h-6 rounded-lg transition-all hover:opacity-100"
                                title={`Akzent ${i + 1}`}
                              />
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })()}

            {/* Akzentfarben */}
            <div>
              <h4 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 mb-1">
                <Tag size={14} /> Akzentfarben
              </h4>
              <p style={{ color: C.muted }} className="text-[10px] mb-4 opacity-60">Diese 4 Farben werden für Kurs-Tags und Akzente verwendet. Tags wählen ihre Farbe automatisch per Zufall aus dieser Palette.</p>
              <div className="space-y-3">
                {(config.customAccents ?? DEFAULT_ACCENTS).map((color, i) => (
                  <div key={i} style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border flex items-center gap-3">
                    <span style={{ backgroundColor: color }} className="w-8 h-8 rounded-lg shrink-0 shadow-sm" />
                    <div className="flex-1">
                      <p style={{ color: C.text }} className="text-[11px] font-semibold">Akzent {i + 1}</p>
                      <p style={{ color: C.muted }} className="text-[9px] font-mono">{color}</p>
                    </div>
                    <input
                      type="color"
                      value={color}
                      onChange={e => setConfig(p => { const a = [...(p.customAccents ?? DEFAULT_ACCENTS)]; a[i] = e.target.value; return { ...p, customAccents: a }; })}
                      className="w-9 h-9 rounded-lg cursor-pointer border-0 p-0.5 bg-transparent"
                    />
                  </div>
                ))}
                <button
                  onClick={() => setConfig(p => ({ ...p, customAccents: [...DEFAULT_ACCENTS] }))}
                  style={{ color: C.muted, borderColor: C.border }}
                  className="w-full border rounded-xl py-2 text-[9px] font-bold uppercase hover:opacity-70 transition-all flex items-center justify-center gap-1.5"
                >
                  <RefreshCw size={10} /> Auf Standard zurücksetzen
                </button>
              </div>
            </div>
          </>}
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
  };

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
                  <td className="px-4 py-3 whitespace-nowrap"><span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${e.type === 'PDF' ? 'bg-blue-100 text-blue-700' : e.type === 'XLSX' ? 'bg-green-100 text-green-700' : 'bg-emerald-100 text-emerald-700'}`}>{e.type}</span></td>
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

  const renderClassPoolModal = () => {
    const prefix = config.institute?.trim() + '-';
    const getMoodleId = c => {
      const n = parseInt(String(c?.id ?? ''), 10);
      if (!isNaN(n) && n > 0) return n;
      if (c?.url) { const m = String(c.url).match(/[?&]id=(\d+)/); if (m) return parseInt(m[1], 10); }
      return null;
    };
    const courseNameById = new Map([
      ...courseDictionary.map(c => [getMoodleId(c), c.label]),
      ...(allMoodleCourses.length > 0 ? allMoodleCourses.map(c => [parseInt(c.id, 10), c.label]) : []),
    ].filter(([id]) => id));
    const poolCourseByMoodleId = new Map(courseDictionary.map(c => [getMoodleId(c), c]));

    return (
      <ModalShell C={C} maxW="max-w-4xl" zIndex={115}>
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-5 border-b flex flex-wrap justify-between items-center gap-4 shrink-0">
          <div className="flex items-center gap-3">
            <div style={{ backgroundColor: C.accent1 }} className="p-2 text-white rounded-xl"><GraduationCap size={18} /></div>
            <div>
              <h3 style={{ color: C.text }} className="font-bold uppercase tracking-tight text-sm">Klassenübersicht</h3>
              <p style={{ color: C.muted }} className="text-[9px] mt-0.5">{moodleGroups.length} Klasse(n) · {selectedMoodleGroupIds.size} ausgewählt</p>
            </div>
          </div>
          <button onClick={() => setActiveModal(null)} style={{ color: C.muted }} className="p-2 hover:bg-black/10 rounded-full"><X size={20} /></button>
        </div>

        <div className="flex-1 overflow-y-auto p-5 space-y-3 custom-scrollbar">
          {moodleGroups.length === 0 ? (
            <p style={{ color: C.muted }} className="text-center italic text-sm py-10">Keine Klassen geladen. Zuerst "Klassen laden" ausführen.</p>
          ) : moodleGroups.map((g, i) => {
            const rowId = i + 1;
            const isSelected = selectedMoodleGroupIds.has(g.id);
            const shortName = g.name.startsWith(prefix) ? g.name.slice(prefix.length) : g.name;
            const existingMoodleIds = (g.existingCourseIds || []).map(Number);
            const matrixIds = (classMatrix[rowId] || []).map(Number).map(id => isNaN(id) ? null : id).filter(Boolean);
            const lockedIds = [...(lockedClassCourseMap[rowId] || new Set())];
            const lockedMoodleIds = lockedIds.map(id => {
              const c = courseDictionary.find(cc => String(cc.id) === String(id));
              return c ? getMoodleId(c) : null;
            }).filter(Boolean);
            const additionalMatrixIds = matrixIds.filter(id => !lockedMoodleIds.includes(id));

            return (
              <div key={g.id} style={{ borderColor: isSelected ? C.accent1 : C.border, backgroundColor: isSelected ? C.accent1 + '08' : C.subtle }} className="rounded-2xl border p-4">
                <div className="flex items-start justify-between gap-3 mb-3">
                  <div className="flex items-center gap-2.5">
                    <div style={{ backgroundColor: isSelected ? C.accent1 : C.muted, opacity: isSelected ? 1 : 0.4 }} className="w-2.5 h-2.5 rounded-full shrink-0 mt-0.5" />
                    <div>
                      <span style={{ color: C.text }} className="font-bold text-sm">{shortName}</span>
                      <span style={{ color: C.muted }} className="text-[10px] ml-2">{g.name}</span>
                    </div>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    {g.memberCount === 0 ? (
                      <span style={{ backgroundColor: '#fef3c7', borderColor: '#fcd34d', color: '#92400e' }} className="text-[10px] font-semibold px-2.5 py-1 rounded-full border flex items-center gap-1.5">
                        <Users size={10} /> Leer
                      </span>
                    ) : (
                      <span style={{ backgroundColor: C.card, borderColor: C.border, color: C.text }} className="text-[10px] font-semibold px-2.5 py-1 rounded-full border flex items-center gap-1.5">
                        <Users size={10} />{g.memberCount} Schüler
                      </span>
                    )}
                    {g.trainerCount > 0 && (
                      <span style={{ backgroundColor: C.accent2 + '15', borderColor: C.accent2 + '40', color: C.accent2 }} className="text-[10px] font-semibold px-2.5 py-1 rounded-full border flex items-center gap-1.5">
                        <GraduationCap size={10} />{g.trainerCount} Trainer
                      </span>
                    )}
                    <span style={{ backgroundColor: isSelected ? C.accent1 + '20' : C.subtle, borderColor: isSelected ? C.accent1 : C.border, color: isSelected ? C.accent1 : C.muted }} className="text-[9px] font-bold px-2 py-1 rounded-full border uppercase">
                      {isSelected ? 'Ausgewählt' : 'Nicht ausgewählt'}
                    </span>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {/* Bestehende Kurse */}
                  <div>
                    <p style={{ color: C.muted }} className="text-[9px] font-bold uppercase tracking-widest mb-1.5 flex items-center gap-1"><RefreshCw size={9} /> Bereits eingeschrieben ({existingMoodleIds.length})</p>
                    {existingMoodleIds.length === 0 ? (
                      <p style={{ color: C.muted }} className="text-[10px] italic">Keine bestehenden Einschreibungen</p>
                    ) : (
                      <div className="space-y-1">
                        {existingMoodleIds.map(id => {
                          const name = courseNameById.get(id) || `Kurs-ID ${id}`;
                          const poolCourse = poolCourseByMoodleId.get(id);
                          const inPool = !!poolCourse;
                          return (
                            <div key={id} style={{ backgroundColor: inPool ? C.accent1 + '12' : C.card, borderColor: inPool ? C.accent1 + '40' : C.border }} className="flex items-center gap-2 px-2.5 py-1.5 rounded-lg border text-[10px]">
                              {inPool ? <RefreshCw size={9} style={{ color: C.accent1 }} className="shrink-0" /> : <Eye size={9} style={{ color: C.muted }} className="shrink-0" />}
                              <span style={{ color: inPool ? C.accent1 : C.text }} className="font-medium truncate">{name}</span>
                              {!inPool && <span style={{ color: C.muted }} className="ml-auto shrink-0 italic">nicht im Pool</span>}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>

                  {/* Matrix-Zuweisung (zusätzliche) */}
                  <div>
                    <p style={{ color: C.muted }} className="text-[9px] font-bold uppercase tracking-widest mb-1.5 flex items-center gap-1"><Plus size={9} /> Neu zugewiesen ({additionalMatrixIds.length})</p>
                    {additionalMatrixIds.length === 0 ? (
                      <p style={{ color: C.muted }} className="text-[10px] italic">Keine weiteren Kurse zugewiesen</p>
                    ) : (
                      <div className="space-y-1">
                        {additionalMatrixIds.map(id => {
                          const name = courseNameById.get(id) || `Kurs-ID ${id}`;
                          return (
                            <div key={id} style={{ backgroundColor: C.accent2 + '12', borderColor: C.accent2 + '40' }} className="flex items-center gap-2 px-2.5 py-1.5 rounded-lg border text-[10px]">
                              <Check size={9} style={{ color: C.accent2 }} className="shrink-0" />
                              <span style={{ color: C.accent2 }} className="font-medium truncate">{name}</span>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-4 border-t flex justify-end shrink-0">
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: C.accent1 }} className="text-white px-6 py-2 rounded-xl font-bold uppercase text-xs hover:brightness-110 active:scale-95 transition-all">Schließen</button>
        </div>
      </ModalShell>
    );
  };

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
            <tr>{['#', 'Tag', 'Kursname', 'Kürzel', 'Link'].map(h => <th key={h} style={{ color: C.muted }} className="px-5 py-3">{h}</th>)}</tr>
          </thead>
          <tbody style={{ borderColor: C.border }} className="divide-y text-[11px]">
            {isLoadingPool ? <tr><td colSpan="5" className="px-5 py-8 text-center"><Loader2 size={20} style={{ color: C.muted }} className="animate-spin mx-auto" /></td></tr>
              : courseDictionary.length === 0 ? <tr><td colSpan="5" style={{ color: C.muted }} className="px-5 py-8 text-center italic">Keine Kurse gefunden.</td></tr>
              : courseDictionary.map((c, i) => (
                <tr key={c.id} className="hover:bg-black/5 transition-colors">
                  <td style={{ color: C.muted }} className="px-5 py-3 font-mono text-[10px]">{i + 1}</td>
                  <td style={{ minWidth: courseTagColWidth }} className="px-5 py-3 whitespace-nowrap">{c.tag ? <span style={{ backgroundColor: getTagColor(c.tag), color: '#fff' }} className="text-[8px] font-bold px-2 py-0.5 rounded">{c.tag}</span> : <span style={{ color: C.muted }} className="italic text-[10px]">–</span>}</td>
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

  const renderInstitutePreviewModal = () => {
    const filtered = instituteSearch.trim()
      ? zohoAllAccounts.filter(a => a.Account_Name.toLowerCase().includes(instituteSearch.toLowerCase()))
      : zohoAllAccounts;
    return (
      <ModalShell C={C} maxW="max-w-2xl" zIndex={105}>
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-5 border-b flex flex-wrap justify-between items-center gap-4 shrink-0">
          <div className="flex items-center gap-3">
            <div style={{ backgroundColor: C.main }} className="p-2 text-white rounded-xl"><Building2 size={18} /></div>
            <div>
              <h3 style={{ color: C.text }} className="font-bold uppercase tracking-tight text-sm">Institutübersicht</h3>
              <p style={{ color: C.muted }} className="text-[9px] mt-0.5">{zohoAllAccounts.length} Institute im CRM</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => { setZohoSearching(true); getAllZohoAccounts(config).then(setZohoAllAccounts).catch(() => {}).finally(() => setZohoSearching(false)); }}
              disabled={zohoSearching}
              style={{ color: C.accent2 }}
              className="flex items-center gap-2 px-3 py-1.5 bg-emerald-50 rounded-lg text-xs font-bold uppercase hover:bg-emerald-100 disabled:opacity-50 transition-colors"
            >
              <RefreshCw size={14} className={zohoSearching ? 'animate-spin' : ''} /> Sync
            </button>
            <button onClick={() => setActiveModal(null)} style={{ color: C.muted }} className="p-2 hover:bg-black/10 rounded-full"><X size={20} /></button>
          </div>
        </div>
        <div style={{ backgroundColor: C.card, borderColor: C.border }} className="px-4 py-3 border-b">
          <input
            value={instituteSearch}
            onChange={e => setInstituteSearch(e.target.value)}
            placeholder="Institut suchen…"
            style={{ backgroundColor: C.subtle, borderColor: C.border, color: C.text }}
            className="w-full px-3 py-2 border rounded-lg text-xs outline-none focus:ring-1"
          />
        </div>
        <div className="flex-1 overflow-auto" style={{ backgroundColor: C.card }}>
          {zohoSearching ? (
            <div className="flex items-center justify-center py-12"><Loader2 size={20} style={{ color: C.muted }} className="animate-spin" /></div>
          ) : filtered.length === 0 ? (
            <p style={{ color: C.muted }} className="text-center py-12 italic text-sm">
              {zohoAllAccounts.length === 0 ? 'Keine Accounts geladen. Zoho CRM konfiguriert?' : 'Keine Treffer.'}
            </p>
          ) : (
            <table className="w-full text-left border-collapse">
              <thead style={{ backgroundColor: C.subtle, borderColor: C.border }} className="sticky top-0 z-10 border-b text-[10px] font-semibold uppercase tracking-widest">
                <tr>
                  <th style={{ color: C.muted }} className="px-5 py-3">#</th>
                  <th style={{ color: C.muted }} className="px-5 py-3">Institutname</th>
                  <th style={{ color: C.muted }} className="px-5 py-3">Aktion</th>
                </tr>
              </thead>
              <tbody style={{ borderColor: C.border }} className="divide-y text-[11px]">
                {filtered.map((acc, i) => (
                  <tr key={acc.id} className="hover:bg-black/5 transition-colors">
                    <td style={{ color: C.muted }} className="px-5 py-3 font-mono text-[10px]">{i + 1}</td>
                    <td style={{ color: C.text }} className="px-5 py-3 font-semibold">{acc.Account_Name}</td>
                    <td className="px-5 py-3">
                      <button
                        onMouseDown={() => {
                          setConfig(p => ({ ...p, institute: acc.Account_Name.replace(/\s+/g, '-') }));
                          setZohoSelectedId(acc.id);
                          setActiveModal(null);
                        }}
                        style={{ color: C.main, borderColor: C.main + '44' }}
                        className="px-2.5 py-1 border rounded-lg text-[9px] font-bold uppercase hover:opacity-80 transition-all flex items-center gap-1"
                      >
                        <CheckCircle2 size={11} /> Übernehmen
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-5 border-t flex justify-end">
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: C.main }} className="text-white px-6 py-2 rounded-xl font-bold uppercase text-xs hover:brightness-110 active:scale-95 transition-all">Schließen</button>
        </div>
      </ModalShell>
    );
  };

  const renderZohoTokenModal = () => (
    <ModalShell C={C} maxW="max-w-lg" zIndex={120}>
      <div style={{ backgroundColor: C.card, borderColor: C.border }} className="p-5 border-b flex items-center justify-between">
        <h2 style={{ color: C.text }} className="font-bold text-sm flex items-center gap-2"><Zap size={15} /> Zoho Refresh Token generieren</h2>
        <button onClick={() => setShowZohoTokenModal(false)} style={{ color: C.muted }} className="hover:opacity-60 transition-opacity"><X size={16} /></button>
      </div>
      <div style={{ backgroundColor: C.card }} className="p-5 space-y-4">
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-4 rounded-xl border space-y-1">
          <p style={{ color: C.muted }} className="text-[9px] font-bold uppercase tracking-widest mb-2">Anleitung</p>
          <p style={{ color: C.muted }} className="text-[9px] opacity-70 leading-relaxed">
            1. <span className="font-mono">api-console.zoho.eu</span> → Self Client → Tab "Generate Code"<br />
            2. Scopes: <span className="font-mono select-all">ZohoCRM.modules.Accounts.READ,ZohoCRM.modules.Accounts.CREATE,ZohoCRM.modules.Deals.CREATE</span><br />
            3. Duration: 10 Minuten → "Create" → Code kopieren<br />
            4. Code unten einfügen und "Generieren" klicken
          </p>
        </div>
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border space-y-1.5">
          <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase block">Client ID</label>
          <p style={{ color: C.text }} className="text-[10px] font-mono">{config.zohoClientId || <span className="opacity-40">–</span>}</p>
        </div>
        <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-3 rounded-xl border space-y-1.5">
          <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase block">Client Secret</label>
          <p style={{ color: C.text }} className="text-[10px] font-mono">{config.zohoClientSecret || <span className="opacity-40">–</span>}</p>
        </div>
        <div className="flex gap-2">
          <input
            type="text"
            value={zohoGrantCode}
            onChange={e => setZohoGrantCode(e.target.value)}
            placeholder="Grant Code: 1000.xxxx…"
            style={{ color: C.text, backgroundColor: C.subtle, borderColor: C.border }}
            className="flex-1 px-3 py-2 border rounded-xl text-[10px] font-mono outline-none focus:border-blue-300 transition-colors"
          />
          <button
            onClick={handleZohoTokenExchange}
            disabled={isExchangingZohoToken || !zohoGrantCode.trim() || !config.zohoClientId || !config.zohoClientSecret}
            style={{ backgroundColor: '#16a34a' }}
            className="text-white px-4 py-2 rounded-xl text-[10px] font-bold uppercase disabled:opacity-50 hover:brightness-110 active:scale-95 transition-all flex items-center gap-1.5 shrink-0"
          >
            {isExchangingZohoToken ? <Loader2 size={12} className="animate-spin" /> : <Zap size={12} />}
            Generieren
          </button>
        </div>
        {config.zohoRefreshToken && (
          <p style={{ color: '#16a34a' }} className="text-[9px] font-semibold flex items-center gap-1">
            <CheckCircle2 size={10} /> Refresh Token erfolgreich gespeichert
          </p>
        )}
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
      {showMoodleConfirm && (
        <div className="fixed inset-0 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm z-[300]">
          <div style={{ backgroundColor: C.card, borderColor: C.border }} className="rounded-2xl shadow-2xl border p-6 max-w-sm w-full mx-4">
            <div className="flex items-center gap-3 mb-3">
              <div className="p-2.5 rounded-xl text-white" style={{ backgroundColor: '#7C3AED' }}><GraduationCap size={18} /></div>
              <h3 style={{ color: C.text }} className="font-bold text-sm uppercase tracking-tight">Moodle einschreiben?</h3>
            </div>
            <p style={{ color: C.muted }} className="text-xs mb-1">Es werden <strong style={{ color: C.text }}>{generatedData.length} Accounts</strong> direkt in Moodle angelegt und eingeschrieben.</p>
            <p style={{ color: C.muted }} className="text-xs mb-5 opacity-70">Bereits bestehende User werden wiederverwendet. Dieser Vorgang kann nicht automatisch rückgängig gemacht werden.</p>
            <div className="flex gap-2 justify-end">
              <button onClick={() => setShowMoodleConfirm(false)} style={{ borderColor: C.border, color: C.muted }} className="border px-4 py-2 rounded-xl text-xs font-bold uppercase hover:bg-black/5 transition-all">Abbrechen</button>
              <button onClick={() => { setShowMoodleConfirm(false); handleMoodleEnrol(); }} style={{ backgroundColor: '#7C3AED' }} className="text-white px-5 py-2 rounded-xl text-xs font-bold uppercase hover:brightness-110 active:scale-95 transition-all flex items-center gap-2">
                <GraduationCap size={13} /> Bestätigen
              </button>
            </div>
          </div>
        </div>
      )}
      <div style={{ backgroundColor: C.subtle, borderColor: C.border }} className="p-4 border-t flex flex-col gap-3">
        {moodleProgress && (
          <div className="space-y-1.5">
            <div className="flex items-center justify-between">
              <span style={{ color: moodleProgress.error ? '#ef4444' : moodleProgress.done ? '#16a34a' : '#7C3AED' }} className="text-[10px] font-semibold flex-1 pr-2">
                {moodleProgress.done ? `✓ ${moodleProgress.label}` : moodleProgress.error ? `✕ ${moodleProgress.label}` : moodleProgress.label}
              </span>
              {!moodleProgress.error && <span style={{ color: C.muted }} className="text-[9px] font-bold shrink-0">{moodleProgress.pct}%</span>}
            </div>
            <div style={{ backgroundColor: C.border }} className="h-1.5 rounded-full overflow-hidden">
              <div style={{ width: `${moodleProgress.pct}%`, backgroundColor: moodleProgress.error ? '#ef4444' : moodleProgress.done ? '#16a34a' : '#7C3AED', transition: 'width 0.4s ease, background-color 0.3s' }} className="h-full rounded-full" />
            </div>
          </div>
        )}
        <div className="flex gap-2">
          <button
            disabled={isMoodleEnrolling || !config.moodleUrl || !config.moodleToken}
            onClick={() => setShowMoodleConfirm(true)}
            style={{ backgroundColor: '#7C3AED' }}
            className="flex-1 py-2.5 text-white rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm flex items-center justify-center gap-2 text-xs disabled:opacity-50"
            title={!config.moodleUrl || !config.moodleToken ? 'Moodle-URL und Token in Einstellungen → Backend konfigurieren' : ''}
          >
            {isMoodleEnrolling ? <Loader2 size={14} className="animate-spin" /> : <GraduationCap size={14} />}
            {isMoodleEnrolling ? 'Einschreiben läuft…' : 'Moodle einschreiben'}
          </button>
          <button onClick={() => setActiveModal(null)} style={{ borderColor: C.border, color: C.muted }} className="px-5 py-2.5 border rounded-xl font-bold uppercase tracking-widest hover:bg-black/5 active:scale-95 transition-all text-xs">Schließen</button>
        </div>
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
      {activeModal === 'history' && renderHistoryModal()}
      {activeModal === 'classPool' && renderClassPoolModal()}
      {activeModal === 'coursePreview' && renderCoursePreviewModal()}
      {activeModal === 'institutePreview' && renderInstitutePreviewModal()}
      {showZohoTokenModal && renderZohoTokenModal()}
      {activeModal === 'dataPreview' && renderDataPreviewModal()}
      {orderPreview && (
        <ModalShell C={C} maxW="max-w-2xl" zIndex={120}>
          <div className="p-6 bg-slate-800 border-b flex justify-between items-center shrink-0">
            <div className="flex items-center gap-3 text-white">
              <div style={{ backgroundColor: C.accent2 }} className="p-2.5 rounded-xl"><Upload size={20} /></div>
              <div>
                <h3 className="text-base font-bold uppercase tracking-tight leading-none">Bestellliste prüfen</h3>
                <p className="text-[10px] text-slate-400 mt-0.5">{orderPreview.fileName}{orderPreview.skz.length ? ` · Schulkennzahl ${orderPreview.skz.join(', ')}` : ''}</p>
              </div>
            </div>
            <button onClick={() => setOrderPreview(null)} className="p-2 hover:bg-white/10 rounded-full text-slate-400 transition-colors"><X size={24} /></button>
          </div>
          <div className="flex-1 overflow-auto bg-white p-6 space-y-4 text-[11px] text-slate-700">
            <table className="w-full border-collapse">
              <thead className="text-slate-500 uppercase tracking-widest text-[9px]"><tr><th className="text-left py-1">Gruppe</th><th className="text-right py-1">Personen</th></tr></thead>
              <tbody className="divide-y divide-slate-100">
                <tr><td className="py-1.5">Trainer (LehrerIn)</td><td className="text-right font-semibold">{orderPreview.teachers}</td></tr>
                {orderPreview.classes.map(c => <tr key={c.grade}><td className="py-1.5">Klasse {c.grade}</td><td className="text-right font-semibold">{c.count}</td></tr>)}
              </tbody>
            </table>
            {orderPreview.problems.length > 0 && (
              <div className="rounded-xl border border-amber-300 bg-amber-50 p-3">
                <p className="font-bold text-amber-800 mb-1.5 flex items-center gap-1.5"><AlertTriangle size={14} /> {orderPreview.problems.length} Zeile(n) werden nicht angelegt — bitte mit der Schule klären</p>
                <ul className="space-y-0.5 text-amber-900">
                  {orderPreview.problems.map(pr => <li key={pr.line}>Zeile {pr.line}: {pr.name || '—'} ({pr.email || 'ohne E-Mail'}) — {pr.reason}</li>)}
                </ul>
              </div>
            )}
            <p className="text-slate-500">Anmeldename und E-Mail in Moodle = E-Mail aus der Bestellung. Konten, die es schon gibt, werden übernommen und behalten ihr Passwort. Die bPK wird nicht gespeichert.</p>
          </div>
          <div className="p-4 border-t bg-slate-50 flex justify-end gap-2 shrink-0">
            <button onClick={() => setOrderPreview(null)} className="px-4 py-2 text-[11px] font-bold uppercase rounded-xl border text-slate-500">Abbrechen</button>
            <button onClick={applyOrderImport} style={{ backgroundColor: C.accent2 }} className="px-4 py-2 text-[11px] font-bold uppercase rounded-xl text-white">{orderPreview.persons.length} Personen übernehmen</button>
          </div>
        </ModalShell>
      )}

      {/* MOODLE ERGEBNIS POPUP */}
      {moodleResult && (
        <div className="fixed inset-0 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm z-[400]">
          <div style={{ backgroundColor: C.card, borderColor: C.border }} className="rounded-2xl shadow-2xl border p-7 max-w-sm w-full mx-4 flex flex-col gap-5">
            <div className="flex items-center gap-3">
              <div className="w-11 h-11 rounded-2xl flex items-center justify-center shrink-0 text-white text-xl" style={{ backgroundColor: '#16a34a' }}>✓</div>
              <div>
                <p style={{ color: C.text }} className="font-bold text-base">Fertig!</p>
                <p style={{ color: C.muted }} className="text-[11px] mt-0.5">{config.institute}</p>
              </div>
            </div>
            <div className="space-y-3">
              <div className="flex items-start gap-3">
                <span className="w-5 h-5 rounded-full flex items-center justify-center text-white text-[11px] font-bold shrink-0 mt-0.5" style={{ backgroundColor: '#7C3AED' }}>✓</span>
                <div>
                  <p style={{ color: C.text }} className="text-[12px] font-semibold">Moodle eingeschrieben</p>
                  <p style={{ color: C.muted }} className="text-[11px] mt-0.5">{moodleResult.moodleSummary}</p>
                </div>
              </div>
              {moodleResult.sharepoint && (
                <div className="flex items-start gap-3">
                  <span className="w-5 h-5 rounded-full flex items-center justify-center text-white text-[11px] font-bold shrink-0 mt-0.5" style={{ backgroundColor: '#0078d4' }}>✓</span>
                  <div>
                    <p style={{ color: C.text }} className="text-[12px] font-semibold">SharePoint hochgeladen</p>
                    <p style={{ color: C.muted }} className="text-[11px] mt-0.5">PDF, Excel & Zusammenfassung gespeichert</p>
                  </div>
                </div>
              )}
              {moodleResult.zoho && (
                <div className="flex items-start gap-3">
                  <span className="w-5 h-5 rounded-full flex items-center justify-center text-white text-[11px] font-bold shrink-0 mt-0.5" style={{ backgroundColor: '#B45309' }}>✓</span>
                  <div>
                    <p style={{ color: C.text }} className="text-[12px] font-semibold">CRM übertragen</p>
                    <p style={{ color: C.muted }} className="text-[11px] mt-0.5">Abschluss im Zoho CRM hinterlegt</p>
                  </div>
                </div>
              )}
            </div>
            <button onClick={() => setMoodleResult(null)} style={{ backgroundColor: '#16a34a' }} className="w-full py-2.5 text-white rounded-xl font-bold text-sm hover:brightness-110 active:scale-95 transition-all">
              Schließen
            </button>
          </div>
        </div>
      )}

      {/* UPDATE POPUP */}
      {pendingUpdate && (
        <div className="fixed inset-0 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm z-[500]">
          <div style={{ backgroundColor: C.card, borderColor: C.border }} className="rounded-2xl shadow-2xl border p-7 max-w-sm w-full mx-4 flex flex-col gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0" style={{ backgroundColor: C.accent1 }}>
                <Zap size={20} className="text-white" />
              </div>
              <div>
                <p style={{ color: C.text }} className="font-bold text-sm">Update verfügbar</p>
                <p style={{ color: C.muted }} className="text-[11px]">Version <span className="font-bold">{pendingUpdate.version}</span> ist bereit zur Installation</p>
              </div>
            </div>
            {isInstalling && (
              <div>
                <div className="flex justify-between mb-1">
                  <span style={{ color: C.muted }} className="text-[10px] font-semibold">Wird heruntergeladen…</span>
                  <span style={{ color: C.muted }} className="text-[10px] font-bold">{installProgress}%</span>
                </div>
                <div style={{ backgroundColor: C.border }} className="h-1.5 rounded-full overflow-hidden">
                  <div style={{ width: `${installProgress}%`, backgroundColor: C.accent1 }} className="h-full rounded-full transition-all duration-300" />
                </div>
              </div>
            )}
            <div className="flex gap-2 justify-end">
              <button
                onClick={() => setPendingUpdate(null)}
                disabled={isInstalling}
                style={{ borderColor: C.border, color: C.muted }}
                className="border px-4 py-2 rounded-xl text-xs font-bold uppercase hover:bg-black/5 disabled:opacity-40 transition-all"
              >
                Später
              </button>
              <button
                onClick={handleInstallUpdate}
                disabled={isInstalling}
                style={{ backgroundColor: C.accent1 }}
                className="text-white px-5 py-2 rounded-xl text-xs font-bold uppercase hover:brightness-110 active:scale-95 disabled:opacity-50 transition-all flex items-center gap-2"
              >
                {isInstalling ? <Loader2 size={13} className="animate-spin" /> : <FileDown size={13} />}
                {isInstalling ? `Installiere…` : 'Jetzt installieren'}
              </button>
            </div>
          </div>
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
            {zohoEnabled && (
              <button onClick={() => { setInstituteSearch(''); setActiveModal('institutePreview'); }} style={{ color: C.muted, backgroundColor: C.subtle, borderColor: C.border }} className="w-full mb-3 py-2 border rounded-xl text-[10px] font-bold uppercase tracking-widest hover:opacity-80 transition-all flex items-center justify-center gap-2">
                <Building2 size={14} /> Institutübersicht
                {zohoAllAccounts.length > 0 && <span style={{ backgroundColor: C.main + '22', color: C.main }} className="px-1.5 py-0.5 rounded-full text-[9px] font-bold">{zohoAllAccounts.length}</span>}
              </button>
            )}

            {showGenerateConfirm ? (
              <div style={{ backgroundColor: '#78350f18', borderColor: '#B45309' }} className="mt-3 p-3 rounded-xl border space-y-2">
                <p className="text-[10px] font-bold text-amber-700 uppercase tracking-wide">Ungewöhnliche Werte:</p>
                {unusualWarnings.map(w => <p key={w} className="text-[10px] text-amber-700">⚠ {w}</p>)}
                <div className="flex gap-2 pt-1">
                  <button onClick={() => generateList(true)} style={{ backgroundColor: '#B45309' }} className="flex-1 text-white text-[10px] font-bold py-1.5 rounded-lg hover:brightness-110 transition-all">Trotzdem generieren</button>
                  <button onClick={() => setShowGenerateConfirm(false)} style={{ borderColor: C.border, color: C.muted }} className="flex-1 border text-[10px] font-bold py-1.5 rounded-lg hover:bg-black/5 transition-all">Abbrechen</button>
                </div>
              </div>
            ) : (
              <button
                disabled={isLoadingPool || isFindingMaxNumbers || !courseDictionary.length || !config.institute?.trim() || !!isEnrolInvalid}
                onClick={() => generateList()}
                style={{ backgroundColor: isEnrolInvalid ? '#B45309' : unusualWarnings.length ? '#B45309' : C.main }}
                className="w-full py-3 text-white rounded-xl font-bold shadow-md mt-3 transition-all hover:brightness-110 active:scale-95 flex items-center justify-center gap-2 disabled:opacity-50 group text-sm">
                {isFindingMaxNumbers ? <Loader2 size={16} className="animate-spin" /> : <Users size={16} className="group-hover:scale-110 transition-transform" />}
                {isFindingMaxNumbers ? 'Moodle abfragen…' : 'Liste generieren'}
                {!isFindingMaxNumbers && unusualWarnings.length > 0 && <AlertTriangle size={14} className="opacity-80" />}
                {!isFindingMaxNumbers && <kbd className="ml-1 opacity-40 text-[9px] font-mono">⌘G</kbd>}
              </button>
            )}
            {isEnrolInvalid && !showGenerateConfirm && (
              <p className="text-[10px] font-medium text-center mt-1.5 text-amber-600">
                {isEnrolInvalid === 'dauer' ? '⚠ Dauer (Tage) muss größer 0 sein.' : '⚠ Einschreibezeitraum liegt in der Vergangenheit.'}
              </p>
            )}

            <button disabled={!isGenerated} onClick={() => setActiveModal('dataPreview')} style={{ color: C.accent1, borderColor: C.accent1 + '33' }} className="w-full py-2.5 bg-transparent border rounded-xl text-[10px] font-bold uppercase mt-3 hover:opacity-80 disabled:opacity-40 transition-all flex items-center justify-center gap-1.5">
              <ClipboardList size={14} /> Daten-Vorschau
            </button>

            <div className="grid grid-cols-3 gap-2 pt-3 border-t mt-3" style={{ borderColor: C.border }}>
              <button disabled={!isGenerated || isExportingPDF} onClick={downloadPDF} style={{ backgroundColor: C.accent1 }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1 hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest transition-all">
                {isExportingPDF ? <Loader2 size={13} className="animate-spin" /> : <FileDown size={13} />} PDF <kbd className="opacity-40 font-mono text-[8px]">⌘P</kbd>
              </button>
              <button disabled={!isGenerated || isExportingExcel} onClick={downloadExcel} style={{ backgroundColor: '#217346' }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1 hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest transition-all">
                {isExportingExcel ? <Loader2 size={13} className="animate-spin" /> : <FileSpreadsheet size={13} />} Excel
              </button>
              <button disabled={!isGenerated} onClick={downloadCSV} style={{ backgroundColor: C.accent2 }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1 hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest transition-all">
                <FileSpreadsheet size={13} /> CSV <kbd className="opacity-40 font-mono text-[8px]">⌘E</kbd>
              </button>
              <button disabled={!isGenerated || isUploadingSP || isExportingPDF} onClick={handleSharePointUpload} style={{ backgroundColor: '#0078d4' }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1 hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest transition-all col-span-3">
                {isUploadingSP ? <Loader2 size={13} className="animate-spin" /> : <Upload size={13} />} SharePoint
              </button>
              <button
                disabled={!isGenerated || isMoodleEnrolling || !config.moodleUrl || !config.moodleToken}
                onClick={() => setShowMoodleConfirm(true)}
                style={{ backgroundColor: '#7C3AED' }}
                className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1 hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest transition-all col-span-3"
                title={!config.moodleUrl || !config.moodleToken ? 'Moodle-URL und Token in Einstellungen → Backend konfigurieren' : 'Direkt in Moodle einschreiben'}
              >
                {isMoodleEnrolling ? <Loader2 size={13} className="animate-spin" /> : <GraduationCap size={13} />}
                {isMoodleEnrolling ? 'Einschreiben…' : 'Moodle einschreiben'}
              </button>
            </div>

            {/* Moodle Fortschrittsbalken */}
            {moodleProgress && (
              <div className="mt-3 space-y-1.5">
                <div className="flex items-center justify-between">
                  <span style={{ color: moodleProgress.error ? '#ef4444' : moodleProgress.done ? '#16a34a' : '#7C3AED' }} className="text-[10px] font-semibold leading-snug flex-1 pr-2">
                    {moodleProgress.done ? `✓ ${moodleProgress.label}` : moodleProgress.error ? `✕ ${moodleProgress.label}` : moodleProgress.label}
                  </span>
                  {!moodleProgress.error && (
                    <span style={{ color: C.muted }} className="text-[9px] font-bold shrink-0">{moodleProgress.pct}%</span>
                  )}
                </div>
                <div style={{ backgroundColor: C.border }} className="h-1.5 rounded-full overflow-hidden">
                  <div
                    style={{
                      width: `${moodleProgress.pct}%`,
                      backgroundColor: moodleProgress.error ? '#ef4444' : moodleProgress.done ? '#16a34a' : '#7C3AED',
                      transition: 'width 0.4s ease, background-color 0.3s',
                    }}
                    className="h-full rounded-full"
                  />
                </div>
              </div>
            )}

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
            <div className="grid grid-cols-1 md:grid-cols-12 gap-2 md:items-start">

              {/* Org */}
              <div style={{ backgroundColor: C.subtle }} className="md:col-span-4 p-4 lg:p-5 rounded-2xl">
                <h3 style={{ color: C.accent1 }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><Building2 size={12} /> Organisation</h3>
                <div className="space-y-3">
                  <div>
                    <label style={{ color: C.muted }} className="text-[9px] font-semibold uppercase block mb-1 ml-1 flex items-center gap-2">
                      Institutsname
                      {zohoEnabled && zohoSearching && <Loader2 size={10} className="animate-spin opacity-50" />}
                      {zohoEnabled && !zohoSearching && zohoSelectedId && <span className="text-[8px] font-bold px-1.5 py-0.5 rounded-full text-white" style={{ backgroundColor: '#16a34a' }}>CRM ✓</span>}
                      {zohoEnabled && !zohoSearching && !zohoSelectedId && config.institute?.trim() && <span className="text-[8px] font-bold px-1.5 py-0.5 rounded-full text-white" style={{ backgroundColor: '#ea580c' }}>Neu</span>}
                      {zohoEnabled && !zohoSearching && zohoAllAccounts.length > 0 && <span style={{ color: C.muted }} className="text-[8px]">{zohoAllAccounts.length} im CRM</span>}
                    </label>
                    <div className="flex gap-1.5 relative">
                      <input
                        name="institute"
                        value={config.institute}
                        onChange={e => { handleInput(e); setZohoSelectedId(null); setZohoDropdownOpen(true); }}
                        onFocus={() => zohoEnabled && setZohoDropdownOpen(true)}
                        onBlur={() => setTimeout(() => setZohoDropdownOpen(false), 150)}
                        placeholder="z.B. Volkshochschule Wien"
                        style={{ backgroundColor: C.card, borderColor: config.institute?.trim() ? C.border : '#ef4444', color: C.text }}
                        className="flex-1 px-3 py-2 border rounded-lg text-sm font-medium focus:ring-1 outline-none shadow-sm transition-all placeholder:opacity-30"
                      />
                      {zohoDropdownOpen && zohoEnabled && zohoSuggestions.length > 0 && (
                        <div style={{ backgroundColor: C.card, borderColor: C.border }} className="absolute top-full left-0 right-0 mt-1 border rounded-xl shadow-xl z-50 overflow-hidden max-h-52 overflow-y-auto">
                          {zohoSuggestions.map(acc => (
                            <button
                              key={acc.id}
                              onMouseDown={() => {
                                setConfig(p => ({ ...p, institute: acc.Account_Name.replace(/\s+/g, '-') }));
                                setZohoSelectedId(acc.id);
                                setZohoDropdownOpen(false);
                              }}
                              style={{ color: C.text, borderColor: C.border }}
                              className="w-full text-left px-3 py-2 text-xs hover:bg-black/5 flex items-center gap-2 border-b last:border-b-0 transition-colors"
                            >
                              <CheckCircle2 size={12} style={{ color: '#16a34a' }} className="shrink-0" />
                              {acc.Account_Name}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    {(() => {
                      const warn = config.trainerCount > 20;
                      return (
                        <div style={{ backgroundColor: C.card, borderColor: warn ? '#B45309' : C.border }} className="p-2.5 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                          <label style={{ color: warn ? '#B45309' : C.muted }} className="text-[8px] font-semibold uppercase block mb-1">Trainer{warn ? ' ⚠' : ''}</label>
                          <input type="number" min="0" name="trainerCount" value={config.trainerCount} onChange={handleInput} style={{ color: warn ? '#B45309' : C.main }} className="w-full bg-transparent text-sm font-semibold outline-none" />
                        </div>
                      );
                    })()}
                    <div style={{ backgroundColor: C.card, borderColor: C.border }} className="p-2.5 rounded-xl border shadow-sm focus-within:border-blue-300 transition-colors">
                      <label style={{ color: C.muted }} className="text-[8px] font-semibold uppercase block mb-1">Kurs Anzahl</label>
                      <input type="number" min="1" name="courseSlotCount" value={config.courseSlotCount} onChange={handleInput} style={{ color: C.main }} className="w-full bg-transparent text-sm font-semibold outline-none" />
                    </div>
                  </div>
                  {/* Aktualisieren / Neu anlegen */}
                  <div>
                    <label style={{ color: C.muted }} className="text-[8px] font-semibold uppercase block mb-1.5 ml-1">Modus</label>
                    <div className="flex flex-col gap-1.5">
                      <button
                        onClick={() => setEnrolMode('both')}
                        style={{
                          backgroundColor: enrolMode === 'both' ? C.main : 'transparent',
                          borderColor: enrolMode === 'both' ? C.main : C.border,
                          color: enrolMode === 'both' ? '#fff' : C.muted,
                        }}
                        className="w-full py-2 text-[10px] font-bold uppercase tracking-wide border rounded-xl transition-all"
                      >
                        Aktualisieren &amp; Neu anlegen
                      </button>
                      <div className="flex gap-1.5">
                        {[{ id: 'update', label: 'Aktualisieren', color: C.accent1 }, { id: 'new', label: 'Neu anlegen', color: C.accent2 }].map(m => (
                          <button
                            key={m.id}
                            onClick={() => setEnrolMode(m.id)}
                            style={{
                              backgroundColor: enrolMode === m.id ? m.color : 'transparent',
                              borderColor: enrolMode === m.id ? m.color : C.border,
                              color: enrolMode === m.id ? '#fff' : C.muted,
                            }}
                            className="flex-1 py-1.5 text-[9px] font-bold uppercase tracking-wide border rounded-xl transition-all"
                          >
                            {m.label}
                          </button>
                        ))}
                      </div>
                    </div>
                    {enrolMode === 'update' && (
                      <p style={{ color: C.accent1 }} className="text-[8px] mt-1 ml-1 leading-snug">Bestehende Klassen aus Moodle laden, auswählen und Zeitraum & Kurse aktualisieren. Passwörter bleiben unverändert.</p>
                    )}
                    {enrolMode === 'new' && (
                      <p style={{ color: C.accent2 }} className="text-[8px] mt-1 ml-1 leading-snug">Schüler/Trainer/Klassen werden fortlaufend nach bestehenden Einträgen nummeriert.</p>
                    )}
                    {enrolMode === 'both' && (
                      <p style={{ color: C.main }} className="text-[8px] mt-1 ml-1 leading-snug">Bestehende Klassen aktualisieren + neue Klassen hinzufügen — in einem Durchlauf.</p>
                    )}
                  </div>
                  {/* Bestellliste vom Marktplatz */}
                  <div>
                    <label style={{ color: C.muted }} className="text-[8px] font-semibold uppercase block mb-1.5 ml-1">Bestellliste (Marktplatz)</label>
                    {orderImport ? (
                      <div style={{ backgroundColor: C.card, borderColor: C.accent2 }} className="p-2.5 rounded-xl border flex items-center gap-2">
                        <CheckCircle2 size={14} style={{ color: C.accent2 }} className="shrink-0" />
                        <div className="flex-1 min-w-0">
                          <p style={{ color: C.text }} className="text-[10px] font-semibold truncate">{orderImport.fileName}</p>
                          <p style={{ color: C.muted }} className="text-[9px]">{orderImport.persons.length} Personen · {orderImport.classes.length} Klasse(n) · Anmeldename = E-Mail</p>
                        </div>
                        <button onClick={clearOrderImport} title="Bestellliste entfernen" style={{ color: C.muted }} className="p-1 hover:bg-black/5 rounded-lg"><X size={14} /></button>
                      </div>
                    ) : (
                      <label style={{ borderColor: C.border, color: C.muted }} className="w-full py-2 text-[10px] font-bold uppercase tracking-wide border border-dashed rounded-xl flex items-center justify-center gap-1.5 cursor-pointer hover:bg-black/5 transition-all">
                        <Upload size={12} /> Bestellliste importieren (.csv)
                        <input type="file" accept=".csv,text/csv" onChange={handleOrderFile} className="hidden" />
                      </label>
                    )}
                  </div>
                </div>
              </div>

              {/* Klassen */}
              <div style={{ backgroundColor: C.card, borderColor: C.border }} className="md:col-span-3 p-4 lg:p-5 rounded-2xl border shadow-sm overflow-y-auto max-h-72 custom-scrollbar">
                {enrolMode === 'update' || enrolMode === 'both' ? (
                  <>
                    <h3 style={{ color: C.muted }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><GraduationCap size={12} /> Klassen aus Moodle</h3>
                    {moodleGroups.length === 0 ? (
                      <>
                        <p style={{ color: C.muted }} className="text-[10px] mb-3 leading-snug">Klassen für das gewählte Institut aus Moodle laden.</p>
                        <button
                          onClick={loadMoodleGroups}
                          disabled={isLoadingMoodleGroups || !config.moodleUrl?.trim() || !config.moodleToken?.trim() || !config.institute?.trim()}
                          style={{ backgroundColor: C.accent1 }}
                          className="w-full py-2 text-white rounded-xl text-[10px] font-bold uppercase hover:brightness-110 active:scale-95 disabled:opacity-50 transition-all flex items-center justify-center gap-2"
                        >
                          {isLoadingMoodleGroups ? <Loader2 size={12} className="animate-spin" /> : <RefreshCw size={12} />}
                          {isLoadingMoodleGroups ? 'Laden…' : 'Klassen laden'}
                        </button>
                      </>
                    ) : (
                      <>
                        <div className="flex items-center justify-between mb-2">
                          <span style={{ color: C.muted }} className="text-[9px] font-semibold uppercase">{moodleGroups.length} Klasse(n)</span>
                          <div className="flex items-center gap-1.5">
                            <button onClick={() => setActiveModal('classPool')} style={{ backgroundColor: C.accent1 + '15', color: C.accent1, borderColor: C.accent1 + '40' }} className="text-[9px] hover:brightness-110 flex items-center gap-1 font-semibold px-2 py-1 rounded-lg border transition-all">
                              <GraduationCap size={10} /> Klassenübersicht
                            </button>
                            <button onClick={() => { setMoodleGroups([]); setSelectedMoodleGroupIds(new Set()); setClassMatrix({}); setLockedClassCourseMap({}); }} style={{ color: C.muted }} className="text-[9px] hover:opacity-70 flex items-center gap-1 px-1.5 py-1">
                              <X size={10} />
                            </button>
                          </div>
                        </div>
                        <div className="space-y-1 max-h-40 overflow-y-auto custom-scrollbar">
                          {moodleGroups.map(g => {
                            const sel = selectedMoodleGroupIds.has(g.id);
                            return (
                              <button
                                key={g.id}
                                onClick={() => setSelectedMoodleGroupIds(prev => { const n = new Set(prev); if (n.has(g.id)) n.delete(g.id); else n.add(g.id); return n; })}
                                style={{ backgroundColor: sel ? C.accent1 + '15' : C.subtle, borderColor: sel ? C.accent1 : C.border, color: sel ? C.accent1 : C.muted }}
                                className="w-full flex items-center justify-between px-2.5 py-1.5 rounded-lg border text-[10px] font-medium transition-all text-left"
                              >
                                <span className="flex items-center gap-1.5">
                                  {sel ? <CheckSquare size={11} /> : <Square size={11} />}
                                  {g.name.replace(config.institute?.trim() + '-', '')}
                                </span>
                                <span style={{ color: C.muted }} className="text-[9px] shrink-0 ml-1">{g.memberCount} Mitgl.</span>
                              </button>
                            );
                          })}
                        </div>
                        <button
                          onClick={loadMoodleGroups}
                          disabled={isLoadingMoodleGroups}
                          style={{ color: C.muted, borderColor: C.border }}
                          className="w-full mt-2 py-1.5 border rounded-xl text-[9px] font-bold uppercase hover:opacity-70 disabled:opacity-40 transition-all flex items-center justify-center gap-1.5"
                        >
                          <RefreshCw size={10} /> Neu laden
                        </button>
                      </>
                    )}
                    {enrolMode === 'both' && (
                      <>
                        <div style={{ borderColor: C.border }} className="my-3 border-t" />
                        <h3 style={{ color: C.main }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-2"><Plus size={12} /> Neue Klassen</h3>
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
                      </>
                    )}
                  </>
                ) : (
                  <>
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
                  </>
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
                    {(() => { const warnPeriod = parseInt(config.enrolPeriod, 10) > 730; return (<>
                      <label style={{ color: warnPeriod ? '#B45309' : C.muted }} className="text-[8px] font-semibold uppercase ml-1">Dauer (Tage){warnPeriod ? ' ⚠' : ''}</label>
                      <input name="enrolPeriod" type="number" min="1" value={config.enrolPeriod} onChange={handleInput} style={{ backgroundColor: C.card, borderColor: warnPeriod ? '#B45309' : C.border, color: warnPeriod ? '#B45309' : C.text }} className="w-full px-2.5 py-1.5 border rounded-lg text-xs font-medium outline-none shadow-sm focus:border-blue-300 transition-colors" />
                    </>); })()}
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
                {classRows.filter(c => {
                  if (enrolMode === 'update' && c.isNew) return false;
                  return !c.isExisting || selectedMoodleGroupIds.has(c.moodleGroupId);
                }).length} Klassen · {activeMatrixCourses.length} Kurse
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
                            {(() => {
                              const selId = config.selectedPoolCourseIds[i] || 'none';
                              const selCourse = courseDictionary.find(c => c.id === selId);
                              const isOpen = openCourseSlot === i;
                              const q = isOpen ? courseSlotSearch.toLowerCase() : '';
                              const filtered = courseDictionary.filter(c =>
                                !q || c.label.toLowerCase().includes(q) || c.shorthand.toLowerCase().includes(q) || c.tag.toLowerCase().includes(q)
                              );
                              return (
                                <>
                                  <button
                                    onClick={() => { setOpenCourseSlot(isOpen ? null : i); setCourseSlotSearch(''); }}
                                    style={{ backgroundColor: C.card, borderColor: C.border, color: selCourse ? C.text : C.muted }}
                                    className="w-full flex items-center gap-1.5 border rounded-md pl-2 pr-6 py-1.5 text-[11px] font-medium outline-none cursor-pointer hover:border-blue-300 transition-all shadow-sm text-left truncate"
                                  >
                                    {selCourse?.tag && <span style={{ backgroundColor: getTagColor(selCourse.tag), color: '#fff' }} className="text-[8px] font-bold px-1.5 py-0.5 rounded shrink-0">{selCourse.tag}</span>}
                                    <span className="truncate">{selCourse ? selCourse.label : '– Nicht belegt –'}</span>
                                  </button>
                                  <ChevronDown size={14} style={{ color: C.muted }} className="absolute right-2 top-1/2 -translate-y-1/2 pointer-events-none" />
                                  {isOpen && (
                                    <>
                                    <div className="fixed inset-0 z-40" onClick={() => setOpenCourseSlot(null)} />
                                    <div style={{ backgroundColor: C.card, borderColor: C.border }} className="absolute top-full left-0 right-0 mt-1 border rounded-xl shadow-xl z-50 overflow-hidden">
                                      <div style={{ borderColor: C.border }} className="p-1.5 border-b">
                                        <input
                                          autoFocus
                                          value={courseSlotSearch}
                                          onChange={e => setCourseSlotSearch(e.target.value)}
                                          placeholder="Suchen…"
                                          style={{ backgroundColor: C.subtle, color: C.text }}
                                          className="w-full px-2 py-1 rounded-lg text-[10px] outline-none placeholder:opacity-40"
                                        />
                                      </div>
                                      <div className="max-h-48 overflow-y-auto">
                                        <button
                                          onMouseDown={() => { updateCourseSlot(i, 'none'); setOpenCourseSlot(null); }}
                                          style={{ color: C.muted, borderColor: C.border }}
                                          className="w-full text-left px-3 py-2 text-[10px] hover:bg-black/5 border-b transition-colors italic"
                                        >– Nicht belegt –</button>
                                        {filtered.map(c => (
                                          <button
                                            key={c.id}
                                            onMouseDown={() => { updateCourseSlot(i, c.id); setOpenCourseSlot(null); }}
                                            style={{ color: C.text, borderColor: C.border, backgroundColor: selId === c.id ? getTagColor(c.tag) + '18' : undefined }}
                                            className="w-full text-left px-3 py-2 text-[10px] hover:bg-black/5 border-b last:border-b-0 transition-colors flex items-center gap-2"
                                          >
                                            <span style={{ minWidth: courseTagColWidth }} className="shrink-0 flex">
                                              {c.tag && <span style={{ backgroundColor: getTagColor(c.tag), color: '#fff' }} className="text-[8px] font-bold px-1.5 py-0.5 rounded">{c.tag}</span>}
                                            </span>
                                            <span className="truncate">{c.label}</span>
                                          </button>
                                        ))}
                                        {filtered.length === 0 && <p style={{ color: C.muted }} className="px-3 py-3 text-[10px] italic">Keine Treffer</p>}
                                      </div>
                                    </div>
                                    </>
                                  )}
                                </>
                              );
                            })()}
                          </div>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody style={{ borderColor: C.border }} className="divide-y text-[11px]">
                  {classRows.length === 0 ? (
                    <tr><td colSpan={config.courseSlotCount + 1} style={{ color: C.muted }} className="py-10 text-center italic font-medium">Bitte oben Klassen definieren.</td></tr>
                  ) : classRows.filter(c => {
                    if (enrolMode === 'update' && c.isNew) return false;
                    return !c.isExisting || selectedMoodleGroupIds.has(c.moodleGroupId);
                  }).map((c, rowIdx, visibleRows) => {
                    const isInvalid = invalidClassIds.has(c.id);
                    const isDeselected = !c.isExisting && enrolMode === 'update' && !selectedUpdateClassIds.has(c.id);
                    const customName = (orderImport && !c.isExisting ? orderImport.classes[c.id - 1]?.grade : null) || config.classNames?.[c.id - 1]?.trim();
                    const isFirstNewRow = enrolMode === 'both' && c.isNew && (rowIdx === 0 || !visibleRows[rowIdx - 1].isNew);
                    return (
                      <React.Fragment key={c.id}>
                      {isFirstNewRow && (
                        <tr><td colSpan={config.courseSlotCount + 1} style={{ backgroundColor: C.main + '12', borderColor: C.main + '30', color: C.main }} className="px-3 py-1 text-[9px] font-bold uppercase tracking-widest border-y"><Plus size={9} className="inline mr-1" />Neue Klassen</td></tr>
                      )}
                      <tr style={{ backgroundColor: isInvalid ? '#FFF1F2' : undefined, opacity: isDeselected ? 0.4 : 1 }} className={`transition-all group ${!isInvalid ? 'hover:bg-blue-50/10' : ''}`}>
                        <td style={{ borderColor: isInvalid ? '#FECDD3' : C.border, backgroundColor: isInvalid ? '#FFE4E6' : C.subtle }} className="px-3 py-3 text-center border-r">
                          <div className="flex flex-col gap-0.5 items-center">
                            {isInvalid && <AlertTriangle size={11} className="text-rose-500 mb-0.5" />}
                            <span style={{ color: isInvalid ? '#BE123C' : C.text }} className="font-bold text-[11px] tracking-tight">K-{String(c.id).padStart(2, '0')}</span>
                            {customName && <span style={{ color: C.accent1, backgroundColor: C.accent1 + '15' }} className="text-[8px] font-bold px-1.5 py-0.5 rounded truncate max-w-[90px]">{customName}</span>}
                            {editingClassSizeId === c.id ? (
                              <input
                                autoFocus
                                type="number"
                                min="0"
                                value={editingClassSizeVal}
                                onChange={e => setEditingClassSizeVal(e.target.value)}
                                onBlur={() => { updateCustomClassSize(c.id, editingClassSizeVal); setEditingClassSizeId(null); }}
                                onKeyDown={e => { if (e.key === 'Enter') e.target.blur(); if (e.key === 'Escape') setEditingClassSizeId(null); }}
                                style={{ color: C.main, borderColor: '#3b82f6' }}
                                className="w-14 bg-transparent text-[10px] font-bold outline-none text-center border-b"
                              />
                            ) : (
                              <div className="flex items-center gap-0.5">
                                <span style={{ color: isInvalid ? '#BE123C' : (c.isExisting && c.defaultSize === 0 ? '#92400e' : C.muted), backgroundColor: isInvalid ? '#FEE2E2' : (c.isExisting && c.defaultSize === 0 ? '#fef3c7' : C.card), borderColor: isInvalid ? '#FECDD3' : (c.isExisting && c.defaultSize === 0 ? '#fcd34d' : C.border) }} className="text-[9px] font-semibold uppercase tracking-tighter border px-1.5 py-0.5 rounded shadow-sm">{c.isExisting && c.defaultSize === 0 && c.size === 0 ? 'Leer' : `${c.size} Pl.`}</span>
                                {(!c.isExisting || c.defaultSize === 0) && <button
                                  onClick={() => { setEditingClassSizeId(c.id); setEditingClassSizeVal(String(c.size)); }}
                                  title={c.isExisting ? 'Neue Schüler in leere Gruppe einschreiben' : 'Schülerzahl bearbeiten'}
                                  style={{ color: config.classCustomSizes?.[c.id] != null ? '#3b82f6' : C.muted }}
                                  className="p-0.5 hover:opacity-70 transition-opacity"
                                >
                                  <Edit3 size={9} />
                                </button>}
                                {(!c.isExisting || c.defaultSize === 0) && config.classCustomSizes?.[c.id] != null && (
                                  <button
                                    onClick={() => setConfig(p => { const s = { ...p.classCustomSizes }; delete s[c.id]; return { ...p, classCustomSizes: s }; })}
                                    title="Custom-Größe zurücksetzen"
                                    style={{ color: '#ef4444' }}
                                    className="p-0.5 hover:opacity-70 transition-opacity"
                                  >
                                    <X size={9} />
                                  </button>
                                )}
                              </div>
                            )}
                          </div>
                        </td>
                        {Array.from({ length: config.courseSlotCount }).map((_, i) => {
                          const cid = config.selectedPoolCourseIds[i];
                          const isActive = cid !== 'none';
                          const isSel = isActive && (classMatrix[c.id] || []).map(String).includes(String(cid));
                          const isLocked = isActive && isSel && lockedClassCourseMap[c.id]?.has(String(cid));
                          return (
                            <td key={i} className="px-4 py-2.5 text-center">
                              {isActive ? (
                                <button onClick={() => toggleCourseAssignment(c.id, cid)} title={isLocked ? 'Bereits eingeschrieben (wird aktualisiert)' : isSel ? 'Zuweisung entfernen' : 'Kurs zuweisen'}
                                  style={{ backgroundColor: isLocked ? C.accent1 : isSel ? C.accent2 : C.card, color: (isSel || isLocked) ? 'white' : C.muted, borderColor: isLocked ? C.accent1 : isSel ? C.accent2 : isInvalid ? '#FECDD3' : C.border, cursor: isLocked ? 'default' : 'pointer' }}
                                  className={`w-10 h-10 rounded-xl border-[1.5px] transition-all duration-200 flex items-center justify-center mx-auto shadow-sm ${!isLocked && !isSel ? 'hover:scale-105' : 'scale-105 shadow-md'}`}>
                                  {isLocked ? <RefreshCw size={14} strokeWidth={2.5} /> : isSel ? <Check size={20} strokeWidth={2.5} /> : <Plus size={18} strokeWidth={2} />}
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
                      </React.Fragment>
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