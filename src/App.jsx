import React, { useState, useMemo, useEffect, useCallback } from 'react';
import { 
  Users, FileSpreadsheet, CheckCircle2, Building2, Plus, 
  Loader2, Table as TableIcon, Check, AlertTriangle, ChevronDown, 
  Eye, X, RefreshCw, Info, Settings, HelpCircle, BookOpen, 
  Zap, ClipboardList, ShieldCheck, GraduationCap, FileDown
} from 'lucide-react';

/**
 * Moodle Anmeldungen V2
 * - REFAKTORISIERUNG: Aufteilung in kleinere Render-Methoden für bessere Wartbarkeit.
 * - UX/UI: Filigraneres Design, dünnere Schriften, nahtloser Matrix-Header ohne Gap.
 * - LOGIK: Sicherere Type-Checks bei numerischen Inputs.
 */

const App = () => {
  // --- Konfiguration & State ---
  const [config, setConfig] = useState({
    institute: '',
    studentPwd: '***ENTFERNT***',
    trainerPwd: '***ENTFERNT***',
    enrolPeriod: 31,
    enrolDate: new Date().toISOString().split('T')[0],
    classSizes: [20, 30, 40, 50],
    classCounts: { 0: 2, 1: 1, 2: 0, 3: 0 },
    trainerCount: 2, 
    courseSlotCount: 4, 
    selectedPoolCourseIds: ['none', 'none', 'none', 'none', 'none', 'none', 'none', 'none'] 
  });

  const [courseDictionary, setCourseDictionary] = useState([]);
  const [isLoadingPool, setIsLoadingPool] = useState(false);
  const [activeModal, setActiveModal] = useState(null); // 'help', 'settings', 'coursePreview', 'dataPreview'
  
  const [classMatrix, setClassMatrix] = useState({});
  const [generatedData, setGeneratedData] = useState([]);
  const [isGenerated, setIsGenerated] = useState(false);
  const [isExportingPDF, setIsExportingPDF] = useState(false);
  const [libsReady, setLibsReady] = useState(false);
  const [validationError, setValidationError] = useState(null);

  const COURSE_API_URL = "***ENTFERNT***";

  const colors = {
    main: "#9D202B",
    bg: "#FFFEF4",
    accent1: "#153D61",
    accent2: "#00664F"
  };

  // --- Hilfsfunktionen ---
  const findValueByPattern = (item, patterns) => {
    const keys = Object.keys(item);
    for (const p of patterns) {
      const match = keys.find(k => k.toLowerCase() === p.toLowerCase());
      if (match && item[match] !== undefined && item[match] !== null && String(item[match]).trim() !== "") return item[match];
    }
    for (const p of patterns) {
      const match = keys.find(k => k.toLowerCase().includes(p.toLowerCase()));
      if (match && item[match] !== undefined && item[match] !== null && String(item[match]).trim() !== "") return item[match];
    }
    return null;
  };

  const isGuid = (str) => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(str);

  // --- API & Daten Laden ---
  const fetchCoursePool = useCallback(async () => {
    setIsLoadingPool(true);
    setValidationError(null);
    try {
      const response = await fetch(COURSE_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ request: "get_courses" }) 
      });
      
      if (!response.ok) throw new Error(`HTTP Fehler: ${response.status}`);
      const rawData = await response.json();
      const items = Array.isArray(rawData) ? rawData : (rawData.value && Array.isArray(rawData.value) ? rawData.value : []);
      
      const normalizedPool = items.map((item, index) => {
        const rawId = findValueByPattern(item, ['id', 'guid', 'key', 'ident']) || `c-${index}`;
        const label = findValueByPattern(item, ['label', 'name', 'titel', 'title', 'bezeichnung', 'kurs']) || "Unbenannter Kurs";
        let shorthand = findValueByPattern(item, ['kurzel', 'kuerzel', 'kürzel', 'shorthand', 'short', 'code', 'kennung', 'abbr']);

        if (!shorthand || isGuid(String(shorthand))) {
            shorthand = String(label).replace(/[^a-zA-Z0-9 ]/g, "").split(" ").filter(w => w.length > 0).map(word => word[0]).join("").toUpperCase();
            if (shorthand.length < 2) shorthand = String(label).substring(0, 3).toUpperCase();
        }
        const url = findValueByPattern(item, ['url', 'link', 'hyperlink', 'moodle', 'portal']) || "";
        return { id: String(rawId), label: String(label).trim(), shorthand: String(shorthand).trim(), url: String(url).trim() };
      });

      if (normalizedPool.length > 0) {
        setCourseDictionary(normalizedPool);
        setConfig(prev => {
          const updatedIds = [...prev.selectedPoolCourseIds];
          normalizedPool.slice(0, prev.courseSlotCount).forEach((c, idx) => {
            if (updatedIds[idx] === 'none') updatedIds[idx] = c.id;
          });
          return { ...prev, selectedPoolCourseIds: updatedIds };
        });
      }
    } catch (e) {
      setValidationError("Verbindung zum Kurs-Pool fehlgeschlagen. Bitte später erneut versuchen.");
    } finally {
      setIsLoadingPool(false);
    }
  }, []);

  useEffect(() => {
    const loadPdfLibs = async () => {
      const loadScript = (src) => new Promise((res, rej) => { 
        if (document.querySelector(`script[src="${src}"]`)) return res();
        const script = document.createElement('script'); 
        script.src = src; 
        script.onload = res; 
        script.onerror = rej; 
        document.head.appendChild(script); 
      });
      try {
        await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
        await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js');
        setLibsReady(true);
      } catch (e) { 
        console.error("PDF Bibliotheken konnten nicht geladen werden.", e); 
      }
    };
    loadPdfLibs();
    fetchCoursePool();
  }, [fetchCoursePool]);

  // --- Berechnungen (Memos) ---
  const activeMatrixCourses = useMemo(() => 
    config.selectedPoolCourseIds.slice(0, config.courseSlotCount)
      .map(id => courseDictionary.find(c => String(c.id) === String(id)))
      .filter(c => c && c.id !== 'none'),
    [config.selectedPoolCourseIds, config.courseSlotCount, courseDictionary]
  );

  const totals = useMemo(() => {
    let stdTotal = 0;
    Object.entries(config.classCounts).forEach(([idx, count]) => {
      stdTotal += (config.classSizes[parseInt(idx)] || 0) * count;
    });
    const clsCount = Object.values(config.classCounts).reduce((a, b) => a + b, 0);
    return { cls: clsCount, std: stdTotal, trainers: config.trainerCount, all: stdTotal + config.trainerCount };
  }, [config.classCounts, config.classSizes, config.trainerCount]);

  const classRows = useMemo(() => {
    const rows = []; 
    let curId = 1;
    [0, 1, 2, 3].forEach(idx => {
      const size = config.classSizes[idx];
      for (let i = 0; i < config.classCounts[idx]; i++) {
        rows.push({ id: curId++, size });
      }
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

  // --- Event Handler ---
  const handleInput = (e) => {
    const { name, value } = e.target;
    const isNumeric = ['enrolPeriod', 'trainerCount', 'courseSlotCount'].includes(name);
    setConfig(p => ({ ...p, [name]: isNumeric ? Math.max(0, parseInt(value, 10) || 0) : value }));
  };

  const handleEndDateInput = (e) => {
    if (!e.target.value) return;
    const picking = new Date(e.target.value);
    const starting = new Date(config.enrolDate);
    const diff = Math.ceil((picking.getTime() - starting.getTime()) / (1000 * 60 * 60 * 24));
    setConfig(p => ({ ...p, enrolPeriod: Math.max(0, diff) }));
  };

  const updateClassSize = (idx, val) => {
    const newSizes = [...config.classSizes];
    newSizes[idx] = Math.max(0, parseInt(val, 10) || 0);
    setConfig(p => ({ ...p, classSizes: newSizes }));
  };

  const updateClassCount = (idx, val) => {
    const count = Math.max(0, parseInt(val, 10) || 0);
    setConfig(p => ({...p, classCounts: {...p.classCounts, [idx]: count}}));
  };

  const toggleCourseAssignment = (classId, courseId) => {
    setClassMatrix(prev => { 
      const currentAssigned = prev[classId] || []; 
      const strCourseId = String(courseId);
      const isAssigned = currentAssigned.map(String).includes(strCourseId);
      return {
        ...prev, 
        [classId]: isAssigned 
          ? currentAssigned.filter(x => String(x) !== strCourseId) 
          : [...currentAssigned, strCourseId]
      }; 
    });
  };

  const updateCourseSlot = (slotIndex, newCourseId) => {
    setConfig(prev => { 
      const newIds = [...prev.selectedPoolCourseIds]; 
      newIds[slotIndex] = newCourseId; 
      return {...prev, selectedPoolCourseIds: newIds }; 
    });
  };

  // --- Generierung & Export ---
  const generateList = () => {
    if (!config.institute?.trim()) return setValidationError("Bitte geben Sie einen Institutsnamen an.");
    if (classRows.length === 0 && config.trainerCount === 0) return setValidationError("Keine Klassen oder Trainer konfiguriert.");
    
    const activeIds = activeMatrixCourses.map(c => String(c.id));
    
    // Validierung: Hat jede Klasse mindestens einen aktiven Kurs?
    const missing = classRows.find(r => {
      const assigned = classMatrix[r.id] || [];
      return assigned.filter(id => activeIds.includes(String(id))).length === 0;
    });
    
    if (missing) return setValidationError(`Fehler: Klasse ${String(missing.id).padStart(2, '0')} hat keine Kurszuweisung.`);
    
    setValidationError(null);
    const instClean = config.institute.replace(/\s+/g, '').toLowerCase();
    const data = [];
    
    // Trainer generieren
    for (let t = 1; t <= config.trainerCount; t++) {
      data.push({ 
        cNum: 'ALL', isT: true, first: 'Trainer', last: config.institute, 
        user: `${instClean}-trainer-${t}`, mail: `trainer${t}@${instClean}.com`, 
        pw: config.trainerPwd, courses: activeMatrixCourses 
      });
    }

    // Schüler generieren
    let sIdx = 1;
    classRows.forEach(r => {
      const selIds = (classMatrix[r.id] || []).map(String);
      const selCourses = courseDictionary.filter(cd => selIds.includes(String(cd.id)) && activeIds.includes(String(cd.id)));
      const classLabel = `${config.institute}-Klasse-${String(r.id).padStart(2, '0')}`;
      
      for (let i = 0; i < r.size; i++) {
        const id = String(sIdx++).padStart(3, '0');
        data.push({ 
          cNum: String(r.id).padStart(2, '0'), cLabel: classLabel, isT: false, 
          first: 'Schüler', last: config.institute, user: `${instClean}-student-${id}`, 
          mail: `student${id}@${instClean}.com`, pw: config.studentPwd, courses: selCourses 
        });
      }
    });

    setGeneratedData(data);
    setIsGenerated(true);
    setActiveModal('dataPreview');
  };

  const downloadCSV = () => {
    if (!generatedData.length) return;
    const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
    
    const rowsWithEnrols = generatedData.map(r => {
      const enrols = [];
      if (r.isT) {
        // Trainer werden in alle Kurse und alle Gruppen eingeschrieben
        activeMatrixCourses.forEach(c => {
          classRows.forEach(cls => {
            enrols.push({ shorthand: c.shorthand, group: `${config.institute}-Klasse-${String(cls.id).padStart(2, '0')}`, role: 4, period: config.enrolPeriod });
          });
        });
      } else {
        // Schüler in ihre spezifischen Kurse
        r.courses.forEach(c => enrols.push({ shorthand: c.shorthand, group: r.cLabel, role: 5, period: config.enrolPeriod }));
      }
      return { ...r, enrols };
    });

    const maxCourses = Math.max(...rowsWithEnrols.map(r => r.enrols.length), 1);
    const headers = ['username', 'firstname', 'lastname', 'email', 'password', 'cohort1'];
    
    for(let i = 1; i <= maxCourses; i++) {
      headers.push(`course${i}`, `group${i}`, `role${i}`, `enrolperiod${i}`);
    }

    const csvLines = rowsWithEnrols.map(r => {
      const line = [esc(r.user), esc(r.first), esc(r.last), esc(r.mail), esc(r.pw), esc(r.last)];
      for(let i = 0; i < maxCourses; i++) {
        const e = r.enrols[i];
        if (e) {
          line.push(esc(e.shorthand), esc(e.group), esc(e.role), esc(e.period));
        } else {
          line.push('""', '""', '""', '""');
        }
      }
      return line.join(',');
    });

    const blob = new Blob(["\uFEFF", headers.join(',') + '\r\n' + csvLines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `EBCL-Moodle-Upload-${config.institute.replace(/\s+/g, '_')}-${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const downloadPDF = async () => {
    if (!window.jspdf || !libsReady) {
      setValidationError("PDF Bibliothek noch nicht geladen. Bitte warten Sie einen Moment.");
      return;
    }
    
    setIsExportingPDF(true);
    try {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF('p', 'mm', 'a4');
      const primary = [157, 32, 43]; // #9D202B
      const today = new Date().toISOString().split('T')[0];

      const renderHeader = (title, info) => {
        doc.setFontSize(20).setTextColor(...primary).setFont("helvetica", "bold").text(config.institute.toUpperCase(), 15, 20);
        doc.setFontSize(12).setTextColor(60).setFont("helvetica", "normal").text(title, 15, 28);
        doc.setFontSize(10).setTextColor(37, 99, 235).text(`Portal: https://world.ebcl.eu/`, 15, 34).link(15, 31, 50, 4, { url: "https://world.ebcl.eu/" });
        doc.setFontSize(8).setTextColor(120).text(`ZEITRAUM: ${new Date(config.enrolDate).toLocaleDateString('de-DE')} bis ${endDateFormatted} | ${info}`, 15, 42);
        doc.setLineWidth(0.3).setDrawColor(0).line(15, 45, 195, 45);
      };

      const tableOptions = (courses) => ({
        startY: 50, theme: 'grid', styles: { fontSize: 7, textColor: [0, 0, 0] },
        headStyles: { fillColor: [40, 40, 40], textColor: 255 },
        columnStyles: { 
          0: { cellWidth: 45 }, 
          1: { cellWidth: 35 }, 
          2: { cellWidth: 25 }, 
          ...Object.fromEntries(Array.from({ length: courses.length }, (_, i) => [i + 3, { textColor: primary, fontStyle: 'bold' }])) 
        },
        didDrawCell: (data) => {
          const { cell, column, section, row } = data;
          if (section === 'body' && column.index === 0 && typeof doc.textField === 'function') {
            doc.textField({ name: `n_${row.index}_${doc.internal.getNumberOfPages()}`, x: cell.x + 1.5, y: cell.y + 1.5, w: cell.width - 3, h: cell.height - 3, fontSize: 7 });
          }
          if (section === 'body' && column.index >= 3) {
            const c = courses[column.index - 3];
            if (c?.url) doc.link(cell.x, cell.y, cell.width, cell.height, { url: c.url });
          }
        }
      });

      // Trainer Seite
      const trainers = generatedData.filter(d => d.isT);
      if (trainers.length > 0) {
        renderHeader("Zugangsdaten: Trainer-Team", `TRAINER: ${trainers.length}`);
        doc.autoTable({ 
          head: [["Name (Eingabefeld)", "Username", "Passwort", ...activeMatrixCourses.map((_,i) => `Kurs ${i+1}`)]], 
          body: trainers.map(t => ["", t.user, t.pw, ...t.courses.map(c => c.label)]), 
          ...tableOptions(activeMatrixCourses), 
          didParseCell: (d) => { if (d.section === 'body') d.cell.styles.fillColor = [255, 255, 245]; } 
        });
      }

      // Schüler Seiten (pro Klasse)
      const classIds = [...new Set(generatedData.filter(d => !d.isT).map(d => d.cNum))].sort();
      classIds.forEach(id => {
        if (trainers.length > 0 || classIds.indexOf(id) > 0) doc.addPage();
        
        const students = generatedData.filter(d => d.cNum === id);
        renderHeader(`Teilnehmerliste: Klasse-${id}`, `SCHÜLER: ${students.length}`);
        
        doc.autoTable({ 
          head: [["Name (Eingabefeld)", "Username", "Passwort", ...students[0].courses.map((_,i) => `Kurs ${i+1}`)]], 
          body: students.map(s => ["", s.user, s.pw, ...s.courses.map(c => c.label)]), 
          ...tableOptions(students[0].courses) 
        });
      });

      // Footer auf allen Seiten
      const pageCount = doc.internal.getNumberOfPages();
      for (let i = 1; i <= pageCount; i++) { 
        doc.setPage(i); 
        doc.setFontSize(7).setTextColor(150).text("EBCL International GmbH", 15, 287); 
        doc.text(`Seite ${i} von ${pageCount}`, 185, 287, { align: 'right' });
      }

      doc.save(`EBCL-Zugangsdaten-${config.institute.replace(/\s+/g, '_')}-${today}.pdf`);
    } catch (e) {
      console.error("PDF Export Fehler:", e);
      setValidationError("Fehler beim Erstellen des PDFs.");
    } finally {
      setIsExportingPDF(false);
    }
  };

  // --- Render Methoden für UI-Komponenten ---
  
  const renderHelpModal = () => (
    <div className="fixed inset-0 z-[250] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-3xl overflow-hidden animate-in zoom-in duration-300 border border-slate-100 flex flex-col max-h-[90vh]">
        <div className="p-6 md:p-8 bg-slate-50 border-b flex justify-between items-start shrink-0 relative overflow-hidden">
          <div className="absolute top-0 right-0 p-8 opacity-5 pointer-events-none"><Zap size={140} style={{ color: colors.main }}/></div>
          <div className="flex items-center gap-4 relative z-10">
            <div style={{ backgroundColor: colors.main }} className="p-3.5 text-white rounded-2xl shadow-lg shadow-red-900/10"><BookOpen size={24}/></div>
            <div>
              <h3 className="text-xl font-bold text-slate-800 tracking-tight leading-none uppercase">Einfach erklärt</h3>
              <p className="text-[10px] text-slate-500 mt-1.5 font-semibold uppercase tracking-[0.2em]">Bedienungsanleitung</p>
            </div>
          </div>
          <button onClick={() => setActiveModal(null)} title="Schließen" className="p-2 hover:bg-slate-200 rounded-full text-slate-400 transition-all active:scale-90"><X size={24}/></button>
        </div>
        
        <div className="flex-1 overflow-y-auto p-6 md:p-8 space-y-8">
          <section className="relative">
            <div className="flex items-center gap-2.5 mb-3"><div style={{ color: colors.main }} className="bg-red-50 p-1.5 rounded-lg"><Zap size={16}/></div><h4 className="text-sm font-bold text-slate-800 uppercase tracking-widest">Was macht dieses Programm?</h4></div>
            <p className="text-xs text-slate-500 leading-relaxed pl-9 border-l-2 border-slate-100">Mit diesem Tool erstellen Sie hunderte Moodle-Zugänge in wenigen Sekunden. Planen Sie hier alles vor und erhalten am Ende fertige Dateien zum Hochladen oder Ausdrucken.</p>
          </section>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8 pl-1 text-slate-500 text-[11px]">
            <div className="space-y-6">
              <div className="flex gap-3 items-start"><span style={{ color: colors.accent1, backgroundColor: colors.accent1 + "10" }} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px]">01</span><div className="space-y-1"><h5 className="font-bold text-slate-700 uppercase tracking-wider">Einstellungen</h5><p>Geben Sie Institutsnamen ein und legen Sie Klassenanzahl und Größen fest.</p></div></div>
              <div className="flex gap-3 items-start"><span style={{ color: colors.accent1, backgroundColor: colors.accent1 + "10" }} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px]">02</span><div className="space-y-1"><h5 className="font-bold text-slate-700 uppercase tracking-wider">Kursverteilung</h5><p>Wählen Sie oben Kurse aus und markieren Sie diese in der Tabelle per Plus (+).</p></div></div>
            </div>
            <div className="space-y-6">
              <div className="flex gap-3 items-start"><span style={{ color: colors.accent1, backgroundColor: colors.accent1 + "10" }} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px]">03</span><div className="space-y-1"><h5 className="font-bold text-slate-700 uppercase tracking-wider">Lehrer-Zugang</h5><p>Trainer erhalten automatisch Zugriff auf alle gewählten Kurse und alle Gruppen.</p></div></div>
              <div className="flex gap-3 items-start"><span style={{ color: colors.accent1, backgroundColor: colors.accent1 + "10" }} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 font-bold text-[10px]">04</span><div className="space-y-1"><h5 className="font-bold text-slate-700 uppercase tracking-wider">Ergebnis sichern</h5><p><b>CSV:</b> Für den Moodle-Upload.<br/><b>PDF:</b> Listen zum Ausdrucken.</p></div></div>
            </div>
          </div>
        </div>
        
        <div className="p-6 md:p-8 bg-slate-50 border-t flex justify-end shrink-0">
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: colors.accent1 }} className="text-white px-8 md:px-10 py-2.5 md:py-3 rounded-2xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-md text-xs">Alles klar!</button>
        </div>
      </div>
    </div>
  );

  const renderSettingsModal = () => (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-xl w-full max-w-md overflow-hidden animate-in zoom-in duration-200 border border-slate-200">
        <div className="p-6 bg-slate-50 border-b flex justify-between items-center">
          <div className="flex items-center gap-3"><div style={{ backgroundColor: colors.accent1 }} className="p-2.5 text-white rounded-xl shadow-md"><Settings size={18}/></div><h3 className="font-bold text-slate-800 uppercase tracking-tight leading-none text-sm">System-Settings</h3></div>
          <button onClick={() => setActiveModal(null)} title="Schließen" className="p-2 hover:bg-slate-200 rounded-full text-slate-400 transition-colors"><X size={20}/></button>
        </div>
        <div className="p-6 space-y-5">
          <h4 className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2"><Users size={14}/> Klassengrößen definieren</h4>
          <div className="grid grid-cols-2 gap-3">
            {config.classSizes.map((size, idx) => (
              <div key={`sz-set-${idx}`} className="bg-white p-3 rounded-xl border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                <label className="text-[9px] font-semibold text-slate-400 uppercase block mb-1">Typ {idx+1}</label>
                <input type="number" min="0" value={size} onChange={(e) => updateClassSize(idx, e.target.value)} style={{ color: colors.main }} className="w-full bg-transparent text-base font-bold outline-none"/>
              </div>
            ))}
          </div>
        </div>
        <div className="p-6 bg-slate-50 border-t flex justify-end">
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: colors.accent1 }} className="text-white px-8 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-md text-xs">Speichern</button>
        </div>
      </div>
    </div>
  );

  const renderCoursePreviewModal = () => (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-xl w-full max-w-4xl overflow-hidden flex flex-col border border-slate-200 max-h-[90vh]">
        <div className="p-5 bg-slate-50 border-b flex flex-wrap justify-between items-center gap-4 shrink-0">
          <div className="flex items-center gap-3">
            <div style={{ backgroundColor: colors.accent1 }} className="p-2 text-white rounded-xl shadow-sm"><Eye size={18}/></div>
            <h3 className="font-bold text-slate-800 uppercase tracking-tight leading-none text-sm">Kursübersicht</h3>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={fetchCoursePool} disabled={isLoadingPool} style={{ color: colors.accent2 }} className="flex items-center gap-2 px-3 py-1.5 bg-emerald-50 rounded-lg text-xs font-bold uppercase hover:bg-emerald-100 disabled:opacity-50 transition-colors">
              <RefreshCw size={14} className={isLoadingPool ? "animate-spin" : ""}/> Sync
            </button>
            <button onClick={() => setActiveModal(null)} title="Schließen" className="p-2 hover:bg-slate-200 rounded-full text-slate-400"><X size={20}/></button>
          </div>
        </div>
        
        <div className="flex-1 overflow-auto bg-white">
          <table className="w-full text-left border-collapse min-w-[600px]">
            <thead className="bg-slate-50 sticky top-0 z-10 border-b border-slate-200 text-[10px] font-semibold text-slate-500 uppercase tracking-widest">
              <tr>
                <th className="px-5 py-3">Kursname</th>
                <th className="px-5 py-3">Kürzel</th>
                <th className="px-5 py-3">Link</th>
              </tr>
            </thead>
            <tbody className="divide-y text-[11px] text-slate-600">
              {courseDictionary.map((c, idx) => (
                <tr key={`p-list-${idx}`} className="hover:bg-slate-50 transition-colors">
                  <td className="px-5 py-3 font-semibold text-slate-700">{c.label}</td>
                  <td className="px-5 py-3"><span style={{ color: colors.accent1, backgroundColor: colors.accent1 + "1A" }} className="font-mono px-2 py-0.5 rounded text-[10px]">{c.shorthand}</span></td>
                  <td className="px-5 py-3 truncate max-w-[250px]">{c.url ? <a href={c.url} target="_blank" rel="noreferrer" style={{ color: colors.accent1 }} className="hover:underline flex items-center gap-1">Öffnen <Eye size={12}/></a> : <span className="text-slate-400 italic">Kein Link</span>}</td>
                </tr>
              ))}
              {courseDictionary.length === 0 && !isLoadingPool && (
                <tr><td colSpan="3" className="px-5 py-8 text-center text-slate-400 italic">Keine Kurse gefunden.</td></tr>
              )}
            </tbody>
          </table>
        </div>
        <div className="p-5 bg-slate-50 border-t flex justify-end">
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: colors.main }} className="text-white px-6 py-2 rounded-xl font-bold uppercase text-xs shadow-sm hover:brightness-110 active:scale-95 transition-all">Schließen</button>
        </div>
      </div>
    </div>
  );

  const renderDataPreviewModal = () => (
    <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-xl w-full max-w-5xl overflow-hidden flex flex-col max-h-[90vh]">
        <div className="p-6 bg-slate-800 border-b flex justify-between items-center shrink-0">
          <div className="flex items-center gap-3 text-white">
            <div style={{ backgroundColor: colors.accent1 }} className="p-2.5 rounded-xl shadow-sm"><ClipboardList size={20}/></div>
            <h3 className="text-base font-bold uppercase tracking-tight leading-none">Daten-Vorschau</h3>
          </div>
          <button onClick={() => setActiveModal(null)} title="Schließen" className="p-2 hover:bg-white/10 rounded-full text-slate-400 transition-colors"><X size={24}/></button>
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
                <tr key={`gp-${idx}`} className={row.isT ? "bg-amber-50/40 font-semibold border-l-2 border-amber-300" : "hover:bg-slate-50/50 transition-colors"}>
                  <td className="px-6 py-3"><span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${row.isT ? 'bg-amber-100 text-amber-700' : 'bg-slate-100 text-slate-500'}`}>{row.isT ? "GLOBAL (Trainer)" : `K-${row.cNum}`}</span></td>
                  <td className="px-6 py-3 font-mono font-medium text-slate-700 tracking-tight">{row.user}</td>
                  <td className="px-6 py-3">
                    <div className="flex flex-wrap gap-1.5">
                      {row.courses.map((c, ci) => (
                        <span key={ci} style={{ color: colors.main, backgroundColor: colors.main + "1A" }} className="font-semibold text-[9px] px-1.5 py-0.5 rounded border border-red-50">{c.shorthand}</span>
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
          <button onClick={downloadCSV} style={{ backgroundColor: colors.accent2 }} className="text-white px-5 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm flex items-center gap-2 text-xs"><FileSpreadsheet size={14}/> CSV</button>
          <button onClick={() => setActiveModal(null)} style={{ backgroundColor: colors.main }} className="text-white px-8 py-2.5 rounded-xl font-bold uppercase tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-sm text-xs">Schließen</button>
        </div>
      </div>
    </div>
  );

  // --- Main Render ---
  return (
    <div style={{ backgroundColor: colors.bg }} className="h-screen flex flex-col p-4 md:p-6 text-slate-800 font-sans selection:bg-red-100 overflow-hidden relative">
      
      {/* Modals */}
      {activeModal === 'help' && renderHelpModal()}
      {activeModal === 'settings' && renderSettingsModal()}
      {activeModal === 'coursePreview' && renderCoursePreviewModal()}
      {activeModal === 'dataPreview' && renderDataPreviewModal()}

      {/* HEADER */}
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-5 shrink-0">
        <div>
          <h1 className="text-xl font-bold flex items-center gap-2 text-slate-800 tracking-tight">
            <TableIcon style={{ color: colors.main }} size={24}/> Moodle Anmeldungen
          </h1>
          <p className="text-slate-500 text-xs italic flex items-center gap-1.5 font-medium mt-1">
            {isLoadingPool ? <Loader2 size={12} className="animate-spin text-blue-500"/> : <CheckCircle2 style={{ color: colors.accent2 }} size={12}/>} 
            Kurse Verfügbar ({courseDictionary.length})
          </p>
        </div>
        <div className="flex items-center gap-3 bg-white px-3 py-1.5 rounded-lg border border-slate-200 shadow-sm">
           <div className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">V1</div>
           <div className="h-3 w-px bg-slate-200"></div>
           <div className="text-[9px] font-bold text-slate-500 uppercase tracking-widest">EBCL INTERNATIONAL</div>
        </div>
      </header>

      {/* HAUPT-GRID */}
      <div className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-5 lg:gap-6 overflow-hidden min-h-0">
        
        {/* SIDEBAR */}
        <aside className="lg:col-span-3 flex flex-col gap-4 overflow-y-auto pr-1 scrollbar-thin">
          {/* Status Card */}
          <section style={{ backgroundColor: colors.accent1 }} className="rounded-3xl p-5 shadow-lg flex flex-col text-white shrink-0 relative overflow-hidden">
             <div className="absolute -right-6 -top-6 opacity-10"><Users size={100} /></div>
             <div className="flex justify-between items-end mb-3 border-b border-white/10 pb-2 relative z-10">
                <div>
                  <div className="text-[9px] font-semibold text-white/70 uppercase tracking-widest mb-1">Gesamt Accounts</div>
                  <div className="text-2xl font-bold leading-none">{totals.all}</div>
                </div>
             </div>
             <div className="space-y-2 relative z-10">
                <div className="flex justify-between items-center text-[11px] bg-white/5 px-2 py-1.5 rounded-lg">
                  <span className="text-white/70 font-medium uppercase tracking-wider flex items-center gap-1.5"><ShieldCheck size={12}/> Trainer</span>
                  <span className="font-bold text-amber-400 text-sm">{totals.trainers}</span>
                </div>
                <div className="flex justify-between items-center text-[11px] bg-white/5 px-2 py-1.5 rounded-lg">
                  <span className="text-white/70 font-medium uppercase tracking-wider flex items-center gap-1.5"><GraduationCap size={12}/> Schüler</span>
                  <span className="font-bold text-blue-300 text-sm">{totals.std}</span>
                </div>
             </div>
          </section>

          {/* Actions */}
          <section className="bg-white p-4 rounded-3xl border border-slate-200 shadow-sm shrink-0">
             <button onClick={() => setActiveModal('coursePreview')} className="w-full py-2.5 bg-slate-50 text-slate-600 border border-slate-200 rounded-xl text-[10px] font-bold uppercase tracking-widest hover:bg-slate-100 transition-all flex items-center justify-center gap-2">
               <Eye size={16}/> Kursübersicht
             </button>
             
             <button disabled={isLoadingPool || courseDictionary.length === 0} onClick={generateList} style={{ backgroundColor: colors.main }} className="w-full py-3 text-white rounded-xl font-bold shadow-md mt-3 transition-all hover:brightness-110 active:scale-95 flex items-center justify-center gap-2 disabled:opacity-50 disabled:hover:brightness-100 group text-sm">
               <Users size={16} className="group-hover:scale-110 transition-transform"/> Liste generieren
             </button>
             
             <button disabled={!isGenerated} onClick={() => setActiveModal('dataPreview')} style={{ color: colors.accent1, borderColor: colors.accent1 + "33" }} className="w-full py-2.5 bg-white border rounded-xl text-[10px] font-bold uppercase mt-3 hover:bg-slate-50 disabled:opacity-40 transition-all flex items-center justify-center gap-1.5">
               <ClipboardList size={14}/> Daten-Vorschau
             </button>
             
             <div className="grid grid-cols-2 gap-2 pt-3 border-t border-slate-100 mt-3">
                <button disabled={!isGenerated || isExportingPDF} onClick={downloadPDF} style={{ backgroundColor: colors.accent1 }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1.5 transition-all hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest">
                  {isExportingPDF ? <Loader2 size={14} className="animate-spin"/> : <FileDown size={14}/>} PDF
                </button>
                <button disabled={!isGenerated} onClick={downloadCSV} style={{ backgroundColor: colors.accent2 }} className="py-2.5 text-white rounded-lg text-[10px] font-bold flex items-center justify-center gap-1.5 transition-all hover:brightness-110 active:scale-95 disabled:opacity-40 uppercase tracking-widest">
                  <FileSpreadsheet size={14}/> CSV
                </button>
             </div>
          </section>

          {/* Error Message */}
          {validationError && (
             <div className="p-3 bg-rose-50 border border-rose-200 rounded-xl text-[10px] text-rose-700 flex items-start gap-2 shadow-sm shrink-0 animate-in fade-in">
               <AlertTriangle size={16} className="shrink-0 text-rose-500 mt-0.5"/>
               <div><span className="font-bold uppercase block mb-0.5">Achtung</span>{validationError}</div>
             </div>
          )}

          {/* Info & Settings Footer */}
          <div style={{ backgroundColor: colors.accent1 + "08", borderColor: colors.accent1 + "15" }} className="p-4 border rounded-2xl text-[10px] leading-relaxed shadow-sm relative mt-auto">
             <div style={{ color: colors.accent1 }} className="flex items-center gap-1.5 mb-3 font-bold uppercase tracking-widest">
               <Info size={14}/> Workflow Info
             </div>
             <div className="space-y-2.5 text-slate-500 font-medium mb-4">
                <div className="flex gap-2.5 items-center"><span style={{ backgroundColor: colors.accent1 }} className="text-white w-4 h-4 rounded-full flex items-center justify-center shrink-0 font-bold text-[8px]">1</span><p>Konfiguration festlegen.</p></div>
                <div className="flex gap-2.5 items-center"><span style={{ backgroundColor: colors.accent1 }} className="text-white w-4 h-4 rounded-full flex items-center justify-center shrink-0 font-bold text-[8px]">2</span><p>Kurse unten zuteilen.</p></div>
                <div className="flex gap-2.5 items-center"><span style={{ backgroundColor: colors.accent1 }} className="text-white w-4 h-4 rounded-full flex items-center justify-center shrink-0 font-bold text-[8px]">3</span><p>Generieren & Export.</p></div>
             </div>
             <div className="flex gap-2 pt-3 border-t border-slate-200/60">
                <button onClick={() => setActiveModal('settings')} className="flex-1 py-2 bg-white border border-slate-200 text-slate-500 rounded-lg hover:text-slate-800 hover:border-slate-300 hover:bg-slate-50 transition-all flex items-center justify-center gap-1.5 shadow-sm">
                  <Settings size={12}/><span className="text-[9px] font-bold uppercase">Settings</span>
                </button>
                <button onClick={() => setActiveModal('help')} className="flex-1 py-2 bg-white border border-slate-200 text-slate-500 rounded-lg hover:text-blue-600 hover:border-blue-300 hover:bg-blue-50 transition-all flex items-center justify-center gap-1.5 shadow-sm">
                  <HelpCircle size={12}/><span className="text-[9px] font-bold uppercase">Hilfe</span>
                </button>
             </div>
          </div>
        </aside>

        {/* HAUPTINHALT (Rechts) */}
        <main className="lg:col-span-9 flex flex-col gap-5 overflow-hidden min-h-0">
          
          {/* Konfiguration Formular */}
          <section className="bg-white p-2 rounded-3xl shadow-sm border border-slate-200 shrink-0">
            <div className="grid grid-cols-1 md:grid-cols-12 gap-2">
              
              {/* Org */}
              <div className="md:col-span-4 bg-slate-50/50 p-4 lg:p-5 rounded-2xl">
                <h3 style={{ color: colors.accent1 }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><Building2 size={12}/> Organisation</h3>
                <div className="space-y-3">
                  <div>
                    <label className="text-[9px] font-semibold text-slate-500 uppercase block mb-1 ml-1">Institutsname</label>
                    <input name="institute" value={config.institute} onChange={handleInput} placeholder="z.B. Volkshochschule" className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm font-medium focus:ring-1 focus:border-blue-400 outline-none shadow-sm transition-all placeholder:font-normal"/>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div className="bg-white p-2.5 rounded-xl border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                      <label className="text-[8px] font-semibold text-slate-400 uppercase block mb-1">Trainer (Anz.)</label>
                      <input type="number" min="0" name="trainerCount" value={config.trainerCount} onChange={handleInput} style={{ color: colors.main }} className="w-full bg-transparent text-sm font-semibold outline-none"/>
                    </div>
                    <div className="bg-white p-2.5 rounded-xl border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                      <label className="text-[8px] font-semibold text-slate-400 uppercase block mb-1">Kurs (Anz.)</label>
                      <input type="number" min="1" max="8" name="courseSlotCount" value={config.courseSlotCount} onChange={handleInput} style={{ color: colors.accent1 }} className="w-full bg-transparent text-sm font-semibold outline-none"/>
                    </div>
                  </div>
                </div>
              </div>
              
              {/* Klassen */}
              <div className="md:col-span-3 bg-white p-4 lg:p-5 rounded-2xl border border-slate-100 shadow-sm">
                <h3 className="text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><GraduationCap size={12}/> Klassen Struktur</h3>
                <div className="space-y-2">
                  {config.classSizes.map((size, idx) => (
                    <div key={idx} className="flex justify-between items-center bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200 focus-within:border-blue-300 transition-colors">
                      <span className="text-[11px] font-medium text-slate-600">{size} Plätze</span>
                      <div className="flex items-center gap-1.5">
                        <span className="text-[9px] text-slate-400">Anz:</span>
                        <input type="number" min="0" value={config.classCounts[idx]} onChange={(e) => updateClassCount(idx, e.target.value)} style={{ color: colors.main }} className="w-8 bg-transparent text-right text-sm font-semibold outline-none"/>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Security & Time */}
              <div className="md:col-span-5 bg-slate-50/50 p-4 lg:p-5 rounded-2xl">
                <h3 style={{ color: colors.main }} className="text-[10px] font-bold uppercase tracking-[0.2em] flex items-center gap-1.5 mb-3"><ShieldCheck size={12}/> Zeit & Sicherheit</h3>
                <div className="grid grid-cols-2 gap-3 mb-2.5">
                  <div className="space-y-1">
                    <label className="text-[8px] font-semibold text-slate-500 uppercase ml-1">Einschreibung</label>
                    <input name="enrolDate" type="date" value={config.enrolDate} onChange={handleInput} className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-medium outline-none shadow-sm focus:border-blue-300 transition-colors"/>
                  </div>
                  <div className="space-y-1">
                    <label className="text-[8px] font-semibold text-slate-500 uppercase ml-1">Dauer (Tage)</label>
                    <input name="enrolPeriod" type="number" min="0" value={config.enrolPeriod} onChange={handleInput} className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-medium outline-none shadow-sm focus:border-blue-300 transition-colors"/>
                  </div>
                </div>
                
                <div className="space-y-1 mb-3">
                  <label style={{ color: colors.main }} className="text-[8px] font-bold uppercase ml-1 flex justify-between"><span>Eingeschrieben bis</span><span className="opacity-70 font-medium">(Autom. berechnet)</span></label>
                  <input type="date" value={endDateDisplay} onChange={handleEndDateInput} style={{ color: colors.main, borderColor: colors.main + "40" }} className="w-full px-3 py-1.5 bg-red-50 border rounded-lg text-xs font-bold outline-none shadow-sm focus:ring-1 focus:ring-red-200 transition-all"/>
                </div>
                
                <div className="grid grid-cols-2 gap-2">
                   <div className="bg-white p-2 rounded-lg border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                     <label className="text-[8px] font-semibold text-slate-400 uppercase block mb-0.5">PW Schüler</label>
                     <input name="studentPwd" value={config.studentPwd} onChange={handleInput} className="w-full bg-transparent text-[11px] font-mono font-medium text-slate-600 outline-none"/>
                   </div>
                   <div className="bg-white p-2 rounded-lg border border-slate-200 shadow-sm focus-within:border-blue-300 transition-colors">
                     <label className="text-[8px] font-semibold text-slate-400 uppercase block mb-0.5">PW Trainer</label>
                     <input name="trainerPwd" value={config.trainerPwd} onChange={handleInput} className="w-full bg-transparent text-[11px] font-mono font-medium text-slate-600 outline-none"/>
                   </div>
                </div>
              </div>
              
            </div>
          </section>

          {/* KURS ZUWEISUNG MATRIX */}
          <section className="bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden flex flex-col flex-1 min-h-0 relative">
            <div className="px-5 py-3 bg-white border-b border-slate-100 flex items-center justify-between shrink-0 relative z-[41]">
              <div className="flex items-center gap-1.5 text-[11px] font-bold uppercase text-slate-600 tracking-wider">
                <CheckCircle2 style={{ color: colors.accent2 }} size={16}/> Kurs Zuweisung Matrix
              </div>
              <div className="text-[9px] text-slate-400 font-medium bg-slate-50 px-2.5 py-1 rounded-full border border-slate-100">
                Klassen: {classRows.length} | Kurse: {config.courseSlotCount}
              </div>
            </div>
            
            <div className="overflow-auto flex-1 relative z-0 bg-white custom-scrollbar">
              <table className="w-full text-left border-separate border-spacing-0 min-w-max">
                <thead className="sticky top-0 z-[40]">
                  <tr>
                    <th className="px-4 py-2 w-28 text-center border-b border-slate-100 bg-white font-semibold text-slate-400 uppercase tracking-widest text-[9px] shadow-sm">
                      Klasse
                    </th>
                    {Array.from({ length: config.courseSlotCount }).map((_, i) => (
                      <th key={`th-${i}`} className="px-2 py-2.5 border-b border-slate-100 min-w-[130px] max-w-[160px] bg-white shadow-sm">
                        <div className="flex flex-col gap-1.5">
                          <span className="text-[9px] text-slate-400 uppercase tracking-widest font-semibold">Kurs {i+1}</span>
                          <div className="relative w-full">
                            <select 
                              value={config.selectedPoolCourseIds[i] || 'none'} 
                              onChange={(e) => updateCourseSlot(i, e.target.value)} 
                              className="w-full appearance-none bg-white border border-slate-200 rounded-md pl-2 pr-6 py-1.5 text-[11px] font-medium text-slate-700 outline-none cursor-pointer hover:border-blue-300 focus:border-blue-400 focus:ring-1 focus:ring-blue-50 transition-all shadow-sm truncate"
                            >
                              <option value="none" className="text-slate-400 italic">-- Nicht belegt --</option>
                              {courseDictionary.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
                            </select>
                            <ChevronDown size={14} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none"/>
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
                         Bitte oben in den Einstellungen Klassen definieren.
                       </td>
                     </tr>
                  ) : (
                    classRows.map(c => (
                      <tr key={`tr-${c.id}`} className="hover:bg-blue-50/20 transition-colors group">
                        <td className="px-4 py-3 text-center border-r border-slate-50 bg-slate-50/30 group-hover:bg-transparent transition-colors">
                          <div className="flex flex-col gap-1 items-center">
                             <span className="font-bold text-slate-700 text-[11px] tracking-tight">K-{String(c.id).padStart(2, '0')}</span>
                             <span style={{ color: colors.accent1 }} className="text-[9px] font-semibold uppercase tracking-tighter bg-white border border-slate-100 px-1.5 py-0.5 rounded shadow-sm">
                               {c.size} Plätze
                             </span>
                          </div>
                        </td>
                        {Array.from({ length: config.courseSlotCount }).map((_, i) => {
                          const cid = config.selectedPoolCourseIds[i];
                          const isActiveSlot = cid !== 'none';
                          const isSelected = isActiveSlot && (classMatrix[c.id] || []).map(String).includes(String(cid));
                          
                          return (
                            <td key={`td-${c.id}-${i}`} className="px-4 py-2.5 text-center">
                              {isActiveSlot ? (
                                <button 
                                  onClick={() => toggleCourseAssignment(c.id, cid)} 
                                  title={isSelected ? "Zuweisung entfernen" : "Kurs zuweisen"}
                                  style={{ 
                                    backgroundColor: isSelected ? colors.accent2 : "white", 
                                    color: isSelected ? "white" : "#94A3B8", 
                                    borderColor: isSelected ? colors.accent2 : "#E2E8F0" 
                                  }} 
                                  className={`w-10 h-10 rounded-xl border-[1.5px] transition-all duration-200 flex items-center justify-center mx-auto shadow-sm ${isSelected ? 'scale-105 shadow-md shadow-emerald-900/5' : 'hover:border-slate-300 hover:bg-slate-50 hover:scale-105'}`}
                                >
                                  {isSelected ? <Check size={20} strokeWidth={2.5}/> : <Plus size={18} strokeWidth={2}/>}
                                </button>
                              ) : (
                                <div className="w-10 h-10 rounded-xl border border-dashed border-slate-200 bg-slate-50/50 mx-auto flex items-center justify-center">
                                  <div className="w-1.5 h-1.5 rounded-full bg-slate-200"></div>
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
            {/* Scroll-Schatten Indikator */}
            <div className="absolute bottom-0 left-0 right-0 h-4 bg-gradient-to-t from-black/5 to-transparent pointer-events-none"></div>
          </section>
          
        </main>
      </div>

      {/* Globale Styles für Custom Scrollbar */}
      <style dangerouslySetInnerHTML={{__html: `
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f8fafc; border-radius: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
      `}} />
    </div>
  );
};

export default App;