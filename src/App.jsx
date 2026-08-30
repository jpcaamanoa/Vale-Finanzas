import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  Plus,
  Trash2,
  ReceiptText,
  Wallet,
  PiggyBank,
  TrendingUp,
  X,
  Download,
  Target,
  Landmark,
  Building2,
  AlertTriangle,
  Gift,
} from "lucide-react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from "recharts";

// ---------- helpers ----------
const money = (n) =>
  new Intl.NumberFormat("es-CL", { style: "currency", currency: "CLP", maximumFractionDigits: 0 }).format(
    Math.round(n || 0)
  );

const todayISO = () => new Date().toISOString().slice(0, 10);

const weekKey = (isoDate) => {
  const d = new Date(isoDate + "T00:00:00");
  const day = (d.getDay() + 6) % 7;
  const monday = new Date(d);
  monday.setDate(d.getDate() - day);
  return monday.toISOString().slice(0, 10);
};

const monthKey = (isoDate) => isoDate.slice(0, 7);

const fmtWeekLabel = (mondayISO) => {
  const start = new Date(mondayISO + "T00:00:00");
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  const opts = { day: "numeric", month: "short" };
  return `${start.toLocaleDateString("es-CL", opts)} — ${end.toLocaleDateString("es-CL", opts)}`;
};

const monthLabel = (ym) => {
  const [y, m] = ym.split("-").map(Number);
  const d = new Date(y, m - 1, 1);
  const label = d.toLocaleDateString("es-CL", { month: "long", year: "numeric" });
  return label.charAt(0).toUpperCase() + label.slice(1);
};

const clamp = (n, min, max) => Math.min(Math.max(n, min), max);
const round = (n) => Math.round(n || 0);

const nextTransferInfo = (day) => {
  if (!day) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tryDate = (y, m) => {
    const lastDayOfMonth = new Date(y, m + 1, 0).getDate();
    return new Date(y, m, Math.min(day, lastDayOfMonth));
  };
  let candidate = tryDate(today.getFullYear(), today.getMonth());
  if (candidate < today) {
    candidate = tryDate(today.getFullYear(), today.getMonth() + 1);
  }
  const diffDays = Math.round((candidate - today) / (1000 * 60 * 60 * 24));
  const label = candidate.toLocaleDateString("es-CL", { day: "numeric", month: "long" });
  return { date: candidate, diffDays, label };
};

const STORAGE_KEY = "control-honorarios:data";
const THEME_KEY = "control-honorarios:theme";

const DEFAULTS = {
  entries: [], // income: {id,date,patient,amount}
  otherIncome: [], // extra income not tied to work: {id,date,description,amount}
  boxLog: [], // {id,date,hours,rate,amount}
  miscExpenses: [], // {id,date,description,amount}
  fixedBusiness: [
    { id: "f1", name: "Publicidad Instagram", amount: 0 },
    { id: "f2", name: "F29 SII", amount: 0 },
  ],
  taxRate: 14.5,
  businessPct: 13.6,
  salaryPct: 25.3,
  marginPct: 5.6,
  bufferTarget: 0,
  transferDay: null, // day of month (1-31) chosen for the monthly Fintual transfer checkpoint
  goals: [
    { id: "g1", name: "Reserva (meses bajos)", current: 0, target: 0, pct: 11.7, sentByMonth: {} },
    { id: "g2", name: "Emergencias", current: 0, target: 0, pct: 5.8, sentByMonth: {} },
    { id: "g3", name: "Viajes", current: 0, target: 0, pct: 16.6, sentByMonth: {} },
    { id: "g4", name: "Compras grandes", current: 0, target: 0, pct: 14.6, sentByMonth: {} },
    { id: "g5", name: "Gustitos", current: 0, target: 0, pct: 3.9, sentByMonth: {} },
    { id: "g6", name: "Jubilación", current: 0, target: 0, pct: 2.9, sentByMonth: {} },
  ],
};

export default function ControlHonorarios() {
  const [data, setData] = useState(DEFAULTS);
  const [loaded, setLoaded] = useState(false);
  const [saveError, setSaveError] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState(monthKey(todayISO()));
  const [theme, setTheme] = useState("light");

  // income form
  const [fDate, setFDate] = useState(todayISO());
  const [fPatient, setFPatient] = useState("");
  const [fAmount, setFAmount] = useState("");

  // other (non-work) income form
  const [oDate, setODate] = useState(todayISO());
  const [oDesc, setODesc] = useState("");
  const [oAmount, setOAmount] = useState("");

  // box form
  const [bDate, setBDate] = useState(todayISO());
  const [bHours, setBHours] = useState("");
  const [bRate, setBRate] = useState("5000");

  // misc personal expense form
  const [mDate, setMDate] = useState(todayISO());
  const [mDesc, setMDesc] = useState("");
  const [mAmount, setMAmount] = useState("");

  // fixed business form
  const [showFixedForm, setShowFixedForm] = useState(false);
  const [xName, setXName] = useState("");
  const [xAmount, setXAmount] = useState("");

  // goal form
  const [showGoalForm, setShowGoalForm] = useState(false);
  const [gName, setGName] = useState("");
  const [gCurrent, setGCurrent] = useState("");
  const [gTarget, setGTarget] = useState("");
  const [gPct, setGPct] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          setData({ ...DEFAULTS, ...parsed });
        }
      } catch (e) {
        // fresh start
      } finally {
        setLoaded(true);
      }
      try {
        const tres = await window.storage.get(THEME_KEY, false);
        if (tres && (tres.value === "dark" || tres.value === "light")) {
          setTheme(tres.value);
        }
      } catch (e) {
        // keep default light theme
      }
    })();
  }, []);

  const toggleTheme = () => {
    const next = theme === "light" ? "dark" : "light";
    setTheme(next);
    // Preference only — best effort, no retry banner needed if this fails.
    window.storage && window.storage.set(THEME_KEY, next, false).catch(() => {});
  };

  const [saveErrorDetail, setSaveErrorDetail] = useState("");
  const [autoRetrying, setAutoRetrying] = useState(false);
  const pendingRef = React.useRef(null); // latest unsaved snapshot, if any
  const retryTimerRef = React.useRef(null);
  const retryDelayRef = React.useRef(3000); // starts at 3s, grows up to 60s

  const clearRetryTimer = () => {
    if (retryTimerRef.current) {
      clearTimeout(retryTimerRef.current);
      retryTimerRef.current = null;
    }
  };

  const scheduleRetry = useCallback((next) => {
    clearRetryTimer();
    setAutoRetrying(true);
    const delay = retryDelayRef.current;
    retryTimerRef.current = setTimeout(() => {
      attemptSaveRef.current(next, true);
    }, delay);
    // grow the delay for the next failure, capped at 60s
    retryDelayRef.current = Math.min(delay * 2, 60000);
  }, []);

  const attemptSaveRef = React.useRef(null);

  const attemptSave = useCallback(
    async (next, isBackgroundRetry = false) => {
      if (!window.storage || typeof window.storage.set !== "function") {
        setSaveError(true);
        setSaveErrorDetail("window.storage no está disponible en este entorno.");
        pendingRef.current = next;
        scheduleRetry(next);
        return false;
      }
      try {
        const result = await window.storage.set(STORAGE_KEY, JSON.stringify(next), false);
        if (!result) {
          throw new Error("La API de almacenamiento devolvió un resultado vacío.");
        }
        setSaveError(false);
        setSaveErrorDetail("");
        setAutoRetrying(false);
        pendingRef.current = null;
        clearRetryTimer();
        retryDelayRef.current = 3000; // reset backoff after success
        return true;
      } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        setSaveError(true);
        setSaveErrorDetail(msg);
        pendingRef.current = next;
        scheduleRetry(next);
        return false;
      }
    },
    [scheduleRetry]
  );

  useEffect(() => {
    attemptSaveRef.current = attemptSave;
  }, [attemptSave]);

  const persist = useCallback(
    async (next) => {
      setData(next);
      await attemptSave(next);
    },
    [attemptSave]
  );

  const [showBackup, setShowBackup] = useState(false);
  const [backupText, setBackupText] = useState("");
  const [restoreInput, setRestoreInput] = useState("");
  const [restoreMsg, setRestoreMsg] = useState("");

  const openBackup = () => {
    setBackupText(JSON.stringify(data));
    setRestoreInput("");
    setRestoreMsg("");
    setShowBackup(true);
  };

  const copyBackup = async () => {
    try {
      await navigator.clipboard.writeText(backupText);
      setRestoreMsg("Copiado al portapapeles. Pégalo en Notas u otro lugar seguro.");
    } catch (e) {
      setRestoreMsg("No se pudo copiar automáticamente. Selecciona el texto de abajo y cópialo a mano.");
    }
  };

  const restoreBackup = () => {
    try {
      const parsed = JSON.parse(restoreInput.trim());
      if (!parsed || typeof parsed !== "object") throw new Error("Formato inválido");
      setData({ ...DEFAULTS, ...parsed });
      persist({ ...DEFAULTS, ...parsed });
      setRestoreMsg("Datos restaurados en pantalla (e intentando guardar).");
    } catch (e) {
      setRestoreMsg("Ese texto no es un respaldo válido. Revisa que lo hayas pegado completo.");
    }
  };

  // Retry immediately if the connection comes back, regardless of backoff timer
  useEffect(() => {
    const onOnline = () => {
      if (pendingRef.current) {
        retryDelayRef.current = 3000;
        attemptSave(pendingRef.current);
      }
    };
    window.addEventListener("online", onOnline);
    return () => window.removeEventListener("online", onOnline);
  }, [attemptSave]);

  // Retry immediately if there's a pending save when the app/tab becomes visible again
  useEffect(() => {
    const onVisible = () => {
      if (document.visibilityState === "visible" && pendingRef.current) {
        retryDelayRef.current = 3000;
        attemptSave(pendingRef.current);
      }
    };
    document.addEventListener("visibilitychange", onVisible);
    return () => document.removeEventListener("visibilitychange", onVisible);
  }, [attemptSave]);

  // Clean up any pending retry timer on unmount
  useEffect(() => {
    return () => clearRetryTimer();
  }, []);

  // ---------- income ----------
  const addEntry = () => {
    const amt = parseFloat(fAmount);
    if (!amt || amt <= 0 || !fDate) return;
    const entry = { id: `i${Date.now()}`, date: fDate, patient: fPatient.trim(), amount: amt };
    persist({ ...data, entries: [...data.entries, entry] });
    setFPatient("");
    setFAmount("");
  };
  const removeEntry = (id) => persist({ ...data, entries: data.entries.filter((e) => e.id !== id) });

  // ---------- other (non-work) income ----------
  const addOtherIncome = () => {
    const amt = parseFloat(oAmount);
    if (!amt || amt <= 0 || !oDate) return;
    const entry = { id: `o${Date.now()}`, date: oDate, description: oDesc.trim(), amount: amt };
    persist({ ...data, otherIncome: [...data.otherIncome, entry] });
    setODesc("");
    setOAmount("");
  };
  const removeOtherIncome = (id) => persist({ ...data, otherIncome: data.otherIncome.filter((e) => e.id !== id) });

  // ---------- box log ----------
  const addBox = () => {
    const hrs = parseFloat(bHours);
    const rate = parseFloat(bRate);
    if (!hrs || hrs <= 0 || !rate || !bDate) return;
    const entry = { id: `bx${Date.now()}`, date: bDate, hours: hrs, rate, amount: hrs * rate };
    persist({ ...data, boxLog: [...data.boxLog, entry] });
    setBHours("");
  };
  const removeBox = (id) => persist({ ...data, boxLog: data.boxLog.filter((e) => e.id !== id) });

  // ---------- misc personal expenses ----------
  const addMisc = () => {
    const amt = parseFloat(mAmount);
    if (!amt || amt <= 0 || !mDate || !mDesc.trim()) return;
    const entry = { id: `m${Date.now()}`, date: mDate, description: mDesc.trim(), amount: amt };
    persist({ ...data, miscExpenses: [...data.miscExpenses, entry] });
    setMDesc("");
    setMAmount("");
  };
  const removeMisc = (id) => persist({ ...data, miscExpenses: data.miscExpenses.filter((e) => e.id !== id) });

  // ---------- fixed business ----------
  const addFixed = () => {
    const amt = parseFloat(xAmount);
    if ((!amt && amt !== 0) || !xName.trim()) return;
    persist({ ...data, fixedBusiness: [...data.fixedBusiness, { id: `fx${Date.now()}`, name: xName.trim(), amount: amt }] });
    setXName("");
    setXAmount("");
    setShowFixedForm(false);
  };
  const removeFixed = (id) => persist({ ...data, fixedBusiness: data.fixedBusiness.filter((e) => e.id !== id) });
  const updateFixedAmount = (id, val) => {
    const amt = parseFloat(val);
    persist({ ...data, fixedBusiness: data.fixedBusiness.map((e) => (e.id === id ? { ...e, amount: isNaN(amt) ? 0 : amt } : e)) });
  };

  const setTaxRate = (val) => persist({ ...data, taxRate: parseFloat(val) || 0 });
  const setBusinessPct = (val) => persist({ ...data, businessPct: parseFloat(val) || 0 });
  const setSalaryPct = (val) => persist({ ...data, salaryPct: parseFloat(val) || 0 });
  const setMarginPct = (val) => persist({ ...data, marginPct: parseFloat(val) || 0 });
  const setBufferTarget = (val) => persist({ ...data, bufferTarget: parseFloat(val) || 0 });
  const setTransferDay = (val) => {
    const n = parseInt(val, 10);
    persist({ ...data, transferDay: val === "" ? null : clamp(isNaN(n) ? 1 : n, 1, 31) });
  };

  // ---------- goals ----------
  const addGoal = () => {
    if (!gName.trim()) return;
    const goal = {
      id: `g${Date.now()}`,
      name: gName.trim(),
      current: parseFloat(gCurrent) || 0,
      target: parseFloat(gTarget) || 0,
      pct: parseFloat(gPct) || 0,
      sentByMonth: {},
    };
    persist({ ...data, goals: [...data.goals, goal] });
    setGName("");
    setGCurrent("");
    setGTarget("");
    setGPct("");
    setShowGoalForm(false);
  };
  const removeGoal = (id) => persist({ ...data, goals: data.goals.filter((g) => g.id !== id) });
  const updateGoalField = (id, field, val) => {
    const num = parseFloat(val);
    persist({ ...data, goals: data.goals.map((g) => (g.id === id ? { ...g, [field]: isNaN(num) ? 0 : num } : g)) });
  };
  const applyAccumulated = (id, amount, month) => {
    persist({
      ...data,
      goals: data.goals.map((g) => {
        if (g.id !== id) return g;
        const sentByMonth = { ...(g.sentByMonth || {}) };
        sentByMonth[month] = (sentByMonth[month] || 0) + amount;
        return { ...g, current: g.current + amount, sentByMonth };
      }),
    });
  };

  // ---------- derived: month scope ----------
  const months = useMemo(() => {
    const set = new Set(data.entries.map((e) => monthKey(e.date)));
    set.add(selectedMonth);
    return Array.from(set).sort().reverse();
  }, [data.entries, selectedMonth]);

  const monthEntries = useMemo(
    () => data.entries.filter((e) => monthKey(e.date) === selectedMonth).sort((a, b) => (a.date < b.date ? 1 : -1)),
    [data.entries, selectedMonth]
  );
  const monthOtherIncome = useMemo(
    () => data.otherIncome.filter((e) => monthKey(e.date) === selectedMonth).sort((a, b) => (a.date < b.date ? 1 : -1)),
    [data.otherIncome, selectedMonth]
  );
  const monthBoxLog = useMemo(
    () => data.boxLog.filter((e) => monthKey(e.date) === selectedMonth).sort((a, b) => (a.date < b.date ? 1 : -1)),
    [data.boxLog, selectedMonth]
  );

  const sessionsTotal = useMemo(() => monthEntries.reduce((s, e) => s + e.amount, 0), [monthEntries]);
  const otherIncomeTotal = useMemo(() => monthOtherIncome.reduce((s, e) => s + e.amount, 0), [monthOtherIncome]);
  const monthTotal = sessionsTotal + otherIncomeTotal;
  const taxTotal = (monthTotal * (data.taxRate || 0)) / 100;
  const afterTax = monthTotal - taxTotal;

  const boxTotal = useMemo(() => monthBoxLog.reduce((s, e) => s + e.amount, 0), [monthBoxLog]);
  const monthMiscExpenses = useMemo(
    () => data.miscExpenses.filter((e) => monthKey(e.date) === selectedMonth).sort((a, b) => (a.date < b.date ? 1 : -1)),
    [data.miscExpenses, selectedMonth]
  );
  const miscTotal = useMemo(() => monthMiscExpenses.reduce((s, e) => s + e.amount, 0), [monthMiscExpenses]);
  const fixedTotal = useMemo(() => data.fixedBusiness.reduce((s, e) => s + e.amount, 0), [data.fixedBusiness]);
  const actualBusiness = boxTotal + fixedTotal;

  const businessAccum = (afterTax * (data.businessPct || 0)) / 100;
  const salaryAccum = (afterTax * (data.salaryPct || 0)) / 100;
  const marginAccum = (afterTax * (data.marginPct || 0)) / 100;
  const goalsPctSum = useMemo(() => data.goals.reduce((s, g) => s + (g.pct || 0), 0), [data.goals]);
  const totalPctSum = data.businessPct + data.salaryPct + data.marginPct + goalsPctSum;
  const pctOff = Math.abs(100 - totalPctSum) > 0.5;
  const transferInfo = useMemo(() => nextTransferInfo(data.transferDay), [data.transferDay]);

  const monthlyTrend = useMemo(() => {
    const monthSet = new Set([...data.entries, ...data.otherIncome].map((e) => monthKey(e.date)));
    monthSet.add(selectedMonth);
    const sortedMonths = Array.from(monthSet).sort().slice(-6); // last 6 months with data (or fewer)
    return sortedMonths.map((m) => {
      const sessionsM = data.entries.filter((e) => monthKey(e.date) === m).reduce((s, e) => s + e.amount, 0);
      const otherM = data.otherIncome.filter((e) => monthKey(e.date) === m).reduce((s, e) => s + e.amount, 0);
      const totalM = sessionsM + otherM;
      const afterTaxM = totalM - (totalM * (data.taxRate || 0)) / 100;
      const goalsM = (afterTaxM * (goalsPctSum || 0)) / 100;
      const [y, mm] = m.split("-");
      const shortLabel = new Date(Number(y), Number(mm) - 1, 1).toLocaleDateString("es-CL", { month: "short" });
      return { month: m, label: shortLabel.replace(".", ""), ingresos: totalM, ahorro: round(goalsM) };
    });
  }, [data.entries, data.otherIncome, data.taxRate, goalsPctSum, selectedMonth]);

  const weeks = useMemo(() => {
    const map = {};
    monthEntries.forEach((e) => {
      const wk = weekKey(e.date);
      if (!map[wk]) map[wk] = [];
      map[wk].push(e);
    });
    return Object.entries(map)
      .sort((a, b) => (a[0] < b[0] ? 1 : -1))
      .map(([wk, items]) => ({ week: wk, items, total: items.reduce((s, i) => s + i.amount, 0) }));
  }, [monthEntries]);
  const weeklyAvg = weeks.length ? sessionsTotal / weeks.length : 0;

  // ---------- export ----------
  const safeSheetName = (name) => name.replace(/[:\\/?*[\]]/g, "-").slice(0, 31);

  const buildMonthSheetAOA = (m) => {
    const sessionsM = data.entries.filter((e) => monthKey(e.date) === m).sort((a, b) => (a.date < b.date ? -1 : 1));
    const otherM = data.otherIncome.filter((e) => monthKey(e.date) === m).sort((a, b) => (a.date < b.date ? -1 : 1));
    const boxM = data.boxLog.filter((e) => monthKey(e.date) === m).sort((a, b) => (a.date < b.date ? -1 : 1));
    const miscM = data.miscExpenses.filter((e) => monthKey(e.date) === m).sort((a, b) => (a.date < b.date ? -1 : 1));

    const sessionsTotalM = sessionsM.reduce((s, e) => s + e.amount, 0);
    const otherTotalM = otherM.reduce((s, e) => s + e.amount, 0);
    const boxTotalM = boxM.reduce((s, e) => s + e.amount, 0);
    const miscTotalM = miscM.reduce((s, e) => s + e.amount, 0);
    const totalM = sessionsTotalM + otherTotalM;
    const taxM = (totalM * (data.taxRate || 0)) / 100;
    const afterTaxM = totalM - taxM;
    const businessAccumM = (afterTaxM * (data.businessPct || 0)) / 100;
    const salaryAccumM = (afterTaxM * (data.salaryPct || 0)) / 100;
    const marginAccumM = (afterTaxM * (data.marginPct || 0)) / 100;

    const rows = [];
    rows.push(["Ingresos (sesiones)"]);
    rows.push(["Fecha", "Paciente", "Monto"]);
    sessionsM.forEach((e) => rows.push([e.date, e.patient || "—", e.amount]));
    rows.push(["", "Total sesiones", sessionsTotalM]);
    rows.push([]);
    rows.push(["Otros ingresos"]);
    rows.push(["Fecha", "Descripción", "Monto"]);
    otherM.forEach((e) => rows.push([e.date, e.description || "—", e.amount]));
    rows.push(["", "Total otros ingresos", otherTotalM]);
    rows.push([]);
    rows.push(["Registro de box"]);
    rows.push(["Fecha", "Horas", "Tarifa/hora", "Monto"]);
    boxM.forEach((e) => rows.push([e.date, e.hours, e.rate, e.amount]));
    rows.push(["", "", "Total box", boxTotalM]);
    rows.push([]);
    rows.push(["Otros gastos"]);
    rows.push(["Fecha", "Descripción", "Monto"]);
    miscM.forEach((e) => rows.push([e.date, e.description, e.amount]));
    rows.push(["", "Total otros gastos", miscTotalM]);
    rows.push([]);
    rows.push(["Resumen del mes"]);
    rows.push(["Ingresos totales", totalM]);
    rows.push([`Impuestos (${data.taxRate}%)`, round(taxM)]);
    rows.push(["Después de impuestos", round(afterTaxM)]);
    rows.push([`Negocio — calculado (${data.businessPct}%)`, round(businessAccumM)]);
    rows.push(["Negocio — real (box + fijos)", round(boxTotalM + fixedTotal)]);
    rows.push([`Sueldo — calculado (${data.salaryPct}%)`, round(salaryAccumM)]);
    rows.push(["Sueldo — gastado (otros gastos)", round(miscTotalM)]);
    rows.push([`Margen / colchón (${data.marginPct}%)`, round(marginAccumM)]);
    rows.push([]);
    rows.push(["Metas — acumulado este mes"]);
    rows.push(["Meta", "% reparto", "Acumulado este mes", "Ya transferido este mes", "Pendiente"]);
    data.goals.forEach((g) => {
      const accum = (afterTaxM * (g.pct || 0)) / 100;
      const applied = (g.sentByMonth && g.sentByMonth[m]) || 0;
      const pending = Math.max(0, accum - applied);
      rows.push([g.name, g.pct, round(accum), round(applied), round(pending)]);
    });
    return rows;
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();

    const monthSet = Array.from(
      new Set([...data.entries, ...data.otherIncome, ...data.boxLog, ...data.miscExpenses].map((e) => monthKey(e.date)))
    ).sort();

    const summaryRows = monthSet.map((m) => {
      const sessionsM = data.entries.filter((e) => monthKey(e.date) === m).reduce((s, e) => s + e.amount, 0);
      const otherM = data.otherIncome.filter((e) => monthKey(e.date) === m).reduce((s, e) => s + e.amount, 0);
      const total = sessionsM + otherM;
      const tax = (total * (data.taxRate || 0)) / 100;
      const at = total - tax;
      const boxM = data.boxLog.filter((e) => monthKey(e.date) === m).reduce((s, e) => s + e.amount, 0);
      return {
        Mes: monthLabel(m),
        "Ingresos sesiones": sessionsM,
        "Otros ingresos": otherM,
        Ingresos: total,
        Impuestos: round(tax),
        "Negocio (%calc)": round((at * data.businessPct) / 100),
        "Negocio real (box+fijos)": round(boxM + fixedTotal),
        Sueldo: round((at * data.salaryPct) / 100),
        Margen: round((at * data.marginPct) / 100),
      };
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Resumen");

    monthSet.forEach((m) => {
      const sheet = XLSX.utils.aoa_to_sheet(buildMonthSheetAOA(m));
      XLSX.utils.book_append_sheet(wb, sheet, safeSheetName(monthLabel(m)));
    });

    const goalRows = data.goals.map((g) => ({ Meta: g.name, "% reparto": g.pct, Ahorrado: g.current, Objetivo: g.target || "—" }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(goalRows), "Metas Fintual");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `control-honorarios-${todayISO()}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  if (!loaded) return <div style={{ fontFamily: "Inter, sans-serif", padding: 40, color: "var(--ink-2, #4a5550)" }}>Cargando…</div>;

  return (
    <div className="ch-root" data-theme={theme} style={styles.page}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        .ch-root {
          --paper: #F7F5F0; --paper-2: #EFEBDD; --panel: #FFFFFF; --rule: #DDD8CC;
          --ink: #20241F; --ink-2: #5C6B62; --ink-3: #8B9088;
          --teal: #1B3A3A; --on-teal: #F7F5F0; --green: #3E7C59; --coral: #B44B3C;
          --gold: #C08A3E; --brown: #7A6A53; --blue: #3E6E8A;
          --input-bg: #FBFAF6; --track: #EAE6D9; --shadow: rgba(27,58,58,0.08);
          --error-bg: #F6E3DE; --error-text: #8A3B28; --warn-bg: #FBF0DA; --warn-text: #8A5A1E;
        }
        .ch-root[data-theme="dark"] {
          --paper: #1A1916; --paper-2: #232220; --panel: #232220; --rule: rgba(240,237,228,0.12);
          --ink: #F0EDE4; --ink-2: #B7B3A8; --ink-3: #85827A;
          --teal: #3ECFA0; --on-teal: #10231D; --green: #3ECFA0; --coral: #F0784E;
          --gold: #D4A830; --brown: #C4A87E; --blue: #6FB0E0;
          --input-bg: #2A2926; --track: #35332F; --shadow: rgba(0,0,0,0.35);
          --error-bg: #34211B; --error-text: #F0A78A; --warn-bg: #332B12; --warn-text: #E0BF6E;
        }
        .stamp-btn:hover { transform: translateY(-1px); box-shadow: 0 4px 10px rgba(27,58,58,0.18); }
        .stamp-btn:active { transform: translateY(0); }
        .row-item { transition: background 0.15s ease; }
        .row-item:hover { background: var(--paper-2); }
        .del-btn { opacity: 0; transition: opacity 0.15s ease; }
        .row-item:hover .del-btn { opacity: 1; }
        input:focus, select:focus { outline: 2px solid var(--teal); outline-offset: 1px; }
        button:focus-visible { outline: 2px solid var(--teal); outline-offset: 2px; }
        .goal-bar-bg { background: var(--track); border-radius: 4px; overflow: hidden; height: 6px; }
        .goal-bar-fill { height: 100%; border-radius: 4px; transition: width 0.3s ease; }
        @media (prefers-reduced-motion: reduce) { * { transition: none !important; } }
      `}</style>

      <header style={styles.header}>
        <div style={styles.headerIcon}>
          <ReceiptText size={22} color="var(--on-teal)" strokeWidth={1.75} />
        </div>
        <div style={{ flex: 1 }}>
          <h1 style={styles.h1}>Control de honorarios</h1>
          <p style={styles.sub}>Cada pago se reparte solo, apenas lo registras.</p>
        </div>
        <button className="stamp-btn" onClick={toggleTheme} style={styles.themeBtn} aria-label="Cambiar tema" title={theme === "light" ? "Modo oscuro" : "Modo claro"}>
          {theme === "light" ? "🌙" : "☀️"}
        </button>
        <button className="stamp-btn" onClick={openBackup} style={styles.exportBtn}>
          <Wallet size={15} strokeWidth={2.25} /> Respaldo manual
        </button>
        <button className="stamp-btn" onClick={exportExcel} style={styles.exportBtn}>
          <Download size={15} strokeWidth={2.25} /> Exportar Excel
        </button>
      </header>

      {showBackup && (
        <div style={styles.backupPanel}>
          <div style={styles.panelHeaderRow}>
            <h2 style={styles.h2}>Respaldo manual</h2>
            <button className="stamp-btn" onClick={() => setShowBackup(false)} style={styles.smallAddBtn} aria-label="Cerrar">
              <X size={15} />
            </button>
          </div>
          <p style={styles.helpText}>
            Si el guardado automático sigue fallando, usa esto como red de seguridad: copia tus datos actuales y pégalos en Notas u otro lugar. Para recuperar, pega el texto de tu respaldo abajo y presiona "Restaurar".
          </p>
          <div style={{ marginBottom: 14 }}>
            <label style={styles.fieldLabel}>Tus datos actuales (cópialos)</label>
            <textarea readOnly value={backupText} style={styles.backupTextarea} onFocus={(e) => e.target.select()} />
            <button className="stamp-btn" onClick={copyBackup} style={{ ...styles.addBtn, marginTop: 8 }}>Copiar al portapapeles</button>
          </div>
          <div>
            <label style={styles.fieldLabel}>Pega aquí un respaldo para restaurar</label>
            <textarea
              placeholder="Pega aquí el texto que copiaste antes"
              value={restoreInput}
              onChange={(e) => setRestoreInput(e.target.value)}
              style={styles.backupTextarea}
            />
            <button className="stamp-btn" onClick={restoreBackup} style={{ ...styles.addBtn, marginTop: 8 }}>Restaurar desde este texto</button>
          </div>
          {restoreMsg && <p style={{ fontSize: 12.5, color: "var(--ink-2)", marginTop: 10 }}>{restoreMsg}</p>}
        </div>
      )}

      {saveError && (
        <div style={styles.errorBanner}>
          <div>
            No se pudo guardar el último cambio — reintentando sola en segundo plano, no necesitas hacer nada. Mientras tanto tus datos siguen en pantalla.
            {saveErrorDetail && <div style={{ fontSize: 11.5, opacity: 0.8, marginTop: 4 }}>Detalle: {saveErrorDetail}</div>}
          </div>
          <span style={styles.retryingBadge}>{autoRetrying ? "Reintentando…" : ""}</span>
        </div>
      )}

      <div style={styles.monthBar}>
        <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} style={styles.monthSelect} />
        {transferInfo && (
          <span style={styles.transferBadge}>
            Próxima transferencia: {transferInfo.label} {transferInfo.diffDays === 0 ? "(hoy)" : `(en ${transferInfo.diffDays} ${transferInfo.diffDays === 1 ? "día" : "días"})`}
          </span>
        )}
      </div>

      {/* Live summary */}
      <div style={styles.cardsGrid}>
        <SummaryCard icon={<TrendingUp size={16} />} label="Ingresos del mes" value={money(monthTotal)} accent="var(--teal)" />
        <SummaryCard icon={<PiggyBank size={16} />} label={`Impuestos (${data.taxRate}%)`} value={money(taxTotal)} accent="var(--gold)" />
        <SummaryCard icon={<Building2 size={16} />} label={`Negocio (${data.businessPct}%)`} value={money(businessAccum)} accent="var(--brown)" />
        <SummaryCard icon={<Landmark size={16} />} label={`Sueldo (${data.salaryPct}%) · gastado ${money(miscTotal)}`} value={money(salaryAccum)} accent="var(--blue)" />
        <SummaryCard icon={<Target size={16} />} label="Metas (acumulado)" value={money(afterTax - businessAccum - salaryAccum - marginAccum)} accent="var(--green)" big />
      </div>

      {pctOff && (
        <div style={styles.warnBanner}>
          <AlertTriangle size={14} strokeWidth={2} />
          Tus porcentajes suman {totalPctSum.toFixed(1)}% en vez de 100%. Ajusta la tabla "Reparto automático" (a la derecha) — ahí están negocio, sueldo, margen y cada meta juntos.
        </div>
      )}

      <div style={styles.mainGrid}>
        <div>
          <section style={styles.panel}>
            <h2 style={styles.h2}>Registrar sesión pagada</h2>
            <div style={styles.formRow}>
              <input type="date" value={fDate} onChange={(e) => setFDate(e.target.value)} style={styles.input} />
              <input type="text" placeholder="Paciente (opcional)" value={fPatient} onChange={(e) => setFPatient(e.target.value)} style={{ ...styles.input, flex: 1.4 }} />
              <input type="number" placeholder="Monto" value={fAmount} onChange={(e) => setFAmount(e.target.value)} style={styles.input} />
              <button className="stamp-btn" onClick={addEntry} style={styles.addBtn}><Plus size={16} /> Agregar</button>
            </div>
            {weeklyAvg > 0 && (
              <p style={styles.avgNote}>Promedio semanal: <strong>{money(weeklyAvg)}</strong> · {weeks.length} {weeks.length === 1 ? "semana" : "semanas"} con ingresos</p>
            )}
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Otros ingresos (no trabajo)</h2>
            <p style={styles.helpText}>Plata que entra pero no es de sesiones: regalos, reembolsos, ventas, etc. Se suma al total del mes igual que las sesiones, para que el reparto automático las incluya.</p>
            <div style={styles.formRow}>
              <input type="date" value={oDate} onChange={(e) => setODate(e.target.value)} style={styles.input} />
              <input type="text" placeholder="Descripción (ej: regalo, venta)" value={oDesc} onChange={(e) => setODesc(e.target.value)} style={{ ...styles.input, flex: 1.4 }} />
              <input type="number" placeholder="Monto" value={oAmount} onChange={(e) => setOAmount(e.target.value)} style={styles.input} />
              <button className="stamp-btn" onClick={addOtherIncome} style={styles.addBtn}><Plus size={16} /> Agregar</button>
            </div>
            {monthOtherIncome.length === 0 && <p style={{ ...styles.empty, marginTop: 10 }}>Sin otros ingresos este mes.</p>}
            {monthOtherIncome.map((item) => (
              <div key={item.id} className="row-item" style={styles.entryRow}>
                <span style={styles.entryDate}>{new Date(item.date + "T00:00:00").toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit" })}</span>
                <span style={styles.entryPatient}>{item.description || "—"}</span>
                <span style={styles.entryAmount}>{money(item.amount)}</span>
                <button className="del-btn" onClick={() => removeOtherIncome(item.id)} style={styles.iconBtn} aria-label="Eliminar"><Trash2 size={14} /></button>
              </div>
            ))}
            {monthOtherIncome.length > 0 && (
              <div style={styles.expenseTotalRow}><span>Total otros ingresos del mes</span><span>{money(otherIncomeTotal)}</span></div>
            )}
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Semana a semana — {monthLabel(selectedMonth)}</h2>
            {weeks.length === 0 && <p style={styles.empty}>Aún no hay ingresos este mes. Agrega arriba, o cambia de mes.</p>}
            {weeks.map(({ week, items, total }) => (
              <div key={week} style={styles.weekBlock}>
                <div style={styles.weekHeader}>
                  <span style={styles.weekLabel}>Semana {fmtWeekLabel(week)}</span>
                  <span style={styles.weekTotal}>{money(total)}</span>
                </div>
                {items.sort((a, b) => (a.date < b.date ? 1 : -1)).map((item) => (
                  <div key={item.id} className="row-item" style={styles.entryRow}>
                    <span style={styles.entryDate}>{new Date(item.date + "T00:00:00").toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit" })}</span>
                    <span style={styles.entryPatient}>{item.patient || "—"}</span>
                    <span style={styles.entryAmount}>{money(item.amount)}</span>
                    <button className="del-btn" onClick={() => removeEntry(item.id)} style={styles.iconBtn} aria-label="Eliminar"><Trash2 size={14} /></button>
                  </div>
                ))}
              </div>
            ))}
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Registro de box</h2>
            <p style={styles.helpText}>La tarifa varía según horario/lugar — pones las horas y el valor de esa hora, y calcula el monto solo.</p>
            <div style={styles.formRow}>
              <input type="date" value={bDate} onChange={(e) => setBDate(e.target.value)} style={styles.input} />
              <input type="number" step="0.5" placeholder="Horas" value={bHours} onChange={(e) => setBHours(e.target.value)} style={{ ...styles.input, maxWidth: 80 }} />
              <input type="number" placeholder="$/hora" value={bRate} onChange={(e) => setBRate(e.target.value)} style={{ ...styles.input, maxWidth: 100 }} />
              <button className="stamp-btn" onClick={addBox} style={styles.addBtn}><Plus size={16} /> Agregar</button>
            </div>
            {monthBoxLog.length === 0 && <p style={{ ...styles.empty, marginTop: 10 }}>Sin pagos de box este mes.</p>}
            {monthBoxLog.map((b) => (
              <div key={b.id} className="row-item" style={styles.entryRow}>
                <span style={styles.entryDate}>{new Date(b.date + "T00:00:00").toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit" })}</span>
                <span style={styles.entryPatient}>{b.hours}h × {money(b.rate)}</span>
                <span style={styles.entryAmount}>{money(b.amount)}</span>
                <button className="del-btn" onClick={() => removeBox(b.id)} style={styles.iconBtn} aria-label="Eliminar"><Trash2 size={14} /></button>
              </div>
            ))}
            {monthBoxLog.length > 0 && (
              <div style={styles.expenseTotalRow}><span>Total box del mes</span><span>{money(boxTotal)}</span></div>
            )}
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Otros gastos</h2>
            <p style={styles.helpText}>Gastos variables que no son fijos ni de negocio: suscripciones, salidas, invitar a la familia, regalos, etc.</p>
            <div style={styles.formRow}>
              <input type="date" value={mDate} onChange={(e) => setMDate(e.target.value)} style={styles.input} />
              <input type="text" placeholder="Descripción (ej: Claude, comida familia)" value={mDesc} onChange={(e) => setMDesc(e.target.value)} style={{ ...styles.input, flex: 1.6 }} />
              <input type="number" placeholder="Monto" value={mAmount} onChange={(e) => setMAmount(e.target.value)} style={styles.input} />
              <button className="stamp-btn" onClick={addMisc} style={styles.addBtn}><Plus size={16} /> Agregar</button>
            </div>
            {monthMiscExpenses.length === 0 && <p style={{ ...styles.empty, marginTop: 10 }}>Sin otros gastos registrados este mes.</p>}
            {monthMiscExpenses.map((m) => (
              <div key={m.id} className="row-item" style={styles.entryRow}>
                <span style={styles.entryDate}>{new Date(m.date + "T00:00:00").toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit" })}</span>
                <span style={styles.entryPatient}>{m.description}</span>
                <span style={styles.entryAmount}>{money(m.amount)}</span>
                <button className="del-btn" onClick={() => removeMisc(m.id)} style={styles.iconBtn} aria-label="Eliminar"><Trash2 size={14} /></button>
              </div>
            ))}
            {monthMiscExpenses.length > 0 && (
              <div style={styles.expenseTotalRow}><span>Total otros gastos del mes</span><span>{money(miscTotal)}</span></div>
            )}
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Metas de ahorro (Fintual)</h2>
            <p style={styles.helpText}>
              "Acumulado del mes" es lo que te corresponde según todo lo que llevas ganado este mes. "Pendiente por transferir" es lo nuevo desde la última vez que apretaste "Sumar a Fintual". No necesitas revisar esto seguido: elige un día del mes y transfiere todo lo pendiente ese día — el resto del tiempo puedes ignorar esta sección tranquilo.
            </p>
            <label style={{ ...styles.fieldLabel, maxWidth: 220 }}>
              Día del mes para transferir (opcional)
              <input
                type="number"
                min="1"
                max="31"
                placeholder="ej: 1"
                value={data.transferDay || ""}
                onChange={(e) => setTransferDay(e.target.value)}
                style={styles.input}
              />
            </label>
            {data.goals.map((g) => {
              const accum = (afterTax * (g.pct || 0)) / 100;
              const applied = (g.sentByMonth && g.sentByMonth[selectedMonth]) || 0;
              const pending = Math.max(0, accum - applied);
              const pct = g.target > 0 ? clamp((g.current / g.target) * 100, 0, 100) : null;
              return (
                <div key={g.id} className="row-item" style={styles.goalBlock}>
                  <div style={styles.goalTopRow}>
                    <span style={styles.goalName}>{g.name}</span>
                    <button className="del-btn" onClick={() => removeGoal(g.id)} style={styles.iconBtn} aria-label="Eliminar meta"><Trash2 size={14} /></button>
                  </div>
                  <div style={styles.goalFieldsRow}>
                    <label style={styles.miniLabel}>% reparto (ver tabla →)
                      <input type="number" step="0.1" value={g.pct} disabled style={{ ...styles.miniInput, color: "var(--ink-3)", background: "var(--paper-2)" }} />
                    </label>
                    <label style={styles.miniLabel}>Ahorrado total
                      <input type="number" value={g.current} onChange={(e) => updateGoalField(g.id, "current", e.target.value)} style={styles.miniInput} />
                    </label>
                    <label style={styles.miniLabel}>Meta ($, opcional)
                      <input type="number" value={g.target || ""} placeholder="—" onChange={(e) => updateGoalField(g.id, "target", e.target.value)} style={styles.miniInput} />
                    </label>
                  </div>
                  <div style={styles.goalAccumRow}>
                    <span>
                      Acumulado del mes: <strong>{money(accum)}</strong>
                      {applied > 0 && <span style={{ color: "var(--ink-3)" }}> · ya transferido {money(applied)}</span>}
                    </span>
                  </div>
                  <div style={styles.goalAccumRow}>
                    <span>Pendiente por transferir: <strong style={{ color: pending > 0 ? "var(--green)" : "var(--ink-3)" }}>{money(pending)}</strong></span>
                    <button
                      className="stamp-btn"
                      onClick={() => applyAccumulated(g.id, pending, selectedMonth)}
                      style={{ ...styles.applyBtnWide, opacity: pending > 0 ? 1 : 0.5, cursor: pending > 0 ? "pointer" : "not-allowed" }}
                      disabled={pending <= 0}
                      title="Sumar solo lo pendiente a lo ahorrado"
                    >
                      <Plus size={12} /> Sumar a Fintual
                    </button>
                  </div>
                  {pct !== null && (
                    <div style={{ marginTop: 6 }}>
                      <div className="goal-bar-bg"><div className="goal-bar-fill" style={{ width: `${pct}%`, background: "var(--green)" }} /></div>
                      <span style={styles.goalPct}>{pct.toFixed(0)}% de {money(g.target)}</span>
                    </div>
                  )}
                </div>
              );
            })}

            {!showGoalForm && <button className="stamp-btn" onClick={() => setShowGoalForm(true)} style={{ ...styles.addBtn, marginTop: 8 }}><Plus size={15} /> Agregar meta</button>}
            {showGoalForm && (
              <div style={styles.goalForm}>
                <input type="text" placeholder="Nombre" value={gName} onChange={(e) => setGName(e.target.value)} style={styles.input} />
                <input type="number" placeholder="Ahorrado" value={gCurrent} onChange={(e) => setGCurrent(e.target.value)} style={styles.input} />
                <input type="number" placeholder="Meta (opcional)" value={gTarget} onChange={(e) => setGTarget(e.target.value)} style={styles.input} />
                <input type="number" placeholder="% reparto" value={gPct} onChange={(e) => setGPct(e.target.value)} style={styles.input} />
                <button className="stamp-btn" onClick={addGoal} style={styles.addBtn}>Guardar</button>
                <button onClick={() => setShowGoalForm(false)} style={styles.cancelBtn}>Cancelar</button>
              </div>
            )}
          </section>
        </div>

        <div>
          <section style={styles.panel}>
            <h2 style={styles.h2}>Tendencia mensual</h2>
            <p style={styles.helpText}>Ingresos y ahorro destinado a metas, mes a mes (últimos {monthlyTrend.length > 1 ? monthlyTrend.length : ""} meses con datos).</p>
            {monthlyTrend.length === 0 ? (
              <p style={styles.empty}>Aún no hay suficientes datos para mostrar una tendencia.</p>
            ) : (
              <div style={{ width: "100%", height: 170 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={monthlyTrend} margin={{ top: 4, right: 4, left: 0, bottom: 0 }} barGap={3}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--rule)" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 10, fill: "var(--ink-3)" }} axisLine={{ stroke: "var(--rule)" }} tickLine={false} />
                    <YAxis tick={{ fontSize: 9, fill: "var(--ink-3)" }} axisLine={false} tickLine={false} width={40} tickFormatter={(v) => (v >= 1000 ? `${Math.round(v / 1000)}k` : v)} />
                    <Tooltip
                      formatter={(value, name) => [money(value), name === "ingresos" ? "Ingresos" : "Ahorro metas"]}
                      labelFormatter={(label) => label}
                      contentStyle={{ background: "var(--panel)", border: "1px solid var(--rule)", borderRadius: 6, fontSize: 12 }}
                    />
                    <Bar dataKey="ingresos" fill="var(--teal)" radius={[3, 3, 0, 0]} maxBarSize={22} />
                    <Bar dataKey="ahorro" fill="var(--green)" radius={[3, 3, 0, 0]} maxBarSize={22} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
            <div style={styles.trendLegend}>
              <span style={styles.trendLegendItem}><span style={{ ...styles.trendDot, background: "var(--teal)" }} /> Ingresos</span>
              <span style={styles.trendLegendItem}><span style={{ ...styles.trendDot, background: "var(--green)" }} /> Ahorro metas</span>
            </div>
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Reparto automático — todo en un solo lugar</h2>
            <p style={styles.helpText}>
              Impuestos sale del ingreso bruto. Todo lo demás (negocio, sueldo, margen, y cada meta) se reparte del ingreso después de impuestos — por eso deben sumar 100% entre todas las filas de abajo, ahorros incluidos.
            </p>

            <div style={styles.allocTable}>
              <div style={styles.allocRow}>
                <span style={styles.allocLabel}>Impuestos <em>(% del bruto)</em></span>
                <input type="number" step="0.5" value={data.taxRate} onChange={(e) => setTaxRate(e.target.value)} style={styles.allocInput} />
              </div>
              <div style={styles.allocDivider}>Después de impuestos ({money(afterTax)} este mes) →</div>
              <div style={styles.allocRow}>
                <span style={styles.allocLabel}>Negocio</span>
                <input type="number" step="0.1" value={data.businessPct} onChange={(e) => setBusinessPct(e.target.value)} style={styles.allocInput} />
              </div>
              <div style={styles.allocRow}>
                <span style={styles.allocLabel}>Sueldo</span>
                <input type="number" step="0.1" value={data.salaryPct} onChange={(e) => setSalaryPct(e.target.value)} style={styles.allocInput} />
              </div>
              <div style={styles.allocRow}>
                <span style={styles.allocLabel}>Margen / colchón</span>
                <input type="number" step="0.1" value={data.marginPct} onChange={(e) => setMarginPct(e.target.value)} style={styles.allocInput} />
              </div>
              {data.goals.map((g) => (
                <div key={g.id} style={styles.allocRow}>
                  <span style={styles.allocLabel}>💰 {g.name}</span>
                  <input type="number" step="0.1" value={g.pct} onChange={(e) => updateGoalField(g.id, "pct", e.target.value)} style={styles.allocInput} />
                </div>
              ))}
            </div>

            <div style={{ ...styles.expenseTotalRow, color: pctOff ? "var(--coral)" : "var(--green)" }}>
              <span>Suma total (debe dar 100%)</span>
              <span>{totalPctSum.toFixed(1)}%</span>
            </div>
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Negocio: calculado vs. real</h2>
            <div style={styles.compareRow}><span>Apartado por % este mes</span><span>{money(businessAccum)}</span></div>
            <div style={styles.compareRow}><span>Gasto real (box + fijos)</span><span>{money(actualBusiness)}</span></div>
            <div style={{ ...styles.expenseTotalRow, color: businessAccum - actualBusiness >= 0 ? "var(--green)" : "var(--coral)" }}>
              <span>{businessAccum - actualBusiness >= 0 ? "Te sobra" : "Te falta"}</span>
              <span>{money(Math.abs(businessAccum - actualBusiness))}</span>
            </div>
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Sueldo: apartado vs. gastado</h2>
            <p style={styles.helpText}>Compara lo que el reparto automático te asigna de sueldo con lo que realmente has gastado (otros gastos de esta página).</p>
            <div style={styles.compareRow}><span>Apartado por % este mes</span><span>{money(salaryAccum)}</span></div>
            <div style={styles.compareRow}><span>Gastado (otros gastos)</span><span>{money(miscTotal)}</span></div>
            <div style={{ ...styles.expenseTotalRow, color: salaryAccum - miscTotal >= 0 ? "var(--green)" : "var(--coral)" }}>
              <span>{salaryAccum - miscTotal >= 0 ? "Te sobra" : "Te pasaste"}</span>
              <span>{money(Math.abs(salaryAccum - miscTotal))}</span>
            </div>
          </section>

          <section style={styles.panel}>
            <h2 style={styles.h2}>Buffer mínimo — cuenta corriente</h2>
            <p style={styles.helpText}>El piso que nunca tocas en tu cuenta corriente, aparte de Fintual. Liquidez para semanas flojas.</p>
            <label style={styles.fieldLabel}>Buffer objetivo
              <input type="number" value={data.bufferTarget} onChange={(e) => setBufferTarget(e.target.value)} style={styles.input} />
            </label>
          </section>

          <section style={styles.panel}>
            <div style={styles.panelHeaderRow}>
              <h2 style={styles.h2}>Gastos fijos de negocio</h2>
              <button className="stamp-btn" onClick={() => setShowFixedForm((s) => !s)} style={styles.smallAddBtn} aria-label="Agregar gasto">
                {showFixedForm ? <X size={15} /> : <Plus size={15} />}
              </button>
            </div>
            {showFixedForm && (
              <div style={styles.expenseForm}>
                <input type="text" placeholder="Nombre" value={xName} onChange={(e) => setXName(e.target.value)} style={styles.input} />
                <input type="number" placeholder="Monto" value={xAmount} onChange={(e) => setXAmount(e.target.value)} style={{ ...styles.input, width: 100 }} />
                <button className="stamp-btn" onClick={addFixed} style={styles.addBtn}>Guardar</button>
              </div>
            )}
            {data.fixedBusiness.map((exp) => (
              <div key={exp.id} className="row-item" style={styles.expenseRow}>
                <span style={styles.entryPatient}>{exp.name}</span>
                <input type="number" value={exp.amount} onChange={(e) => updateFixedAmount(exp.id, e.target.value)} style={styles.expenseInput} />
                <button className="del-btn" onClick={() => removeFixed(exp.id)} style={styles.iconBtn} aria-label="Eliminar"><Trash2 size={14} /></button>
              </div>
            ))}
            <div style={styles.expenseTotalRow}><span>Total fijos</span><span>{money(fixedTotal)}</span></div>
          </section>
        </div>
      </div>
    </div>
  );
}

function SummaryCard({ icon, label, value, accent, big }) {
  return (
    <div style={{ ...styles.summaryCard, borderTop: `3px solid ${accent}` }}>
      <div style={{ ...styles.summaryIcon, color: accent }}>{icon}</div>
      <div style={styles.summaryLabel}>{label}</div>
      <div style={{ ...styles.summaryValue, fontSize: big ? 21 : 18, color: big ? accent : "var(--ink)" }}>{value}</div>
    </div>
  );
}

const styles = {
  page: { fontFamily: "'Inter', sans-serif", background: "var(--paper)", minHeight: "100vh", padding: "28px 20px 60px", color: "var(--ink)", maxWidth: 1080, margin: "0 auto" },
  header: { display: "flex", alignItems: "center", gap: 14, marginBottom: 22 },
  headerIcon: { width: 44, height: 44, borderRadius: 10, background: "var(--teal)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  h1: { fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 26, margin: 0, letterSpacing: "-0.01em" },
  sub: { margin: "2px 0 0", fontSize: 13.5, color: "var(--ink-2)" },
  exportBtn: { display: "flex", alignItems: "center", gap: 6, background: "var(--panel)", color: "var(--teal)", border: "1px solid var(--rule)", borderRadius: 8, padding: "9px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" },
  themeBtn: { display: "flex", alignItems: "center", justifyContent: "center", background: "var(--panel)", color: "var(--ink)", border: "1px solid var(--rule)", borderRadius: 8, width: 38, height: 38, fontSize: 15, cursor: "pointer", flexShrink: 0 },
  errorBanner: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, background: "var(--error-bg)", color: "var(--error-text)", padding: "10px 14px", borderRadius: 8, fontSize: 13, marginBottom: 16 },
  retryingBadge: { fontSize: 12, fontWeight: 600, color: "var(--error-text)", whiteSpace: "nowrap", flexShrink: 0, fontStyle: "italic" },
  backupPanel: { background: "var(--panel)", borderRadius: 12, padding: "18px 20px 20px", marginBottom: 18, boxShadow: "0 1px 3px var(--shadow)", border: "1.5px solid var(--gold)" },
  backupTextarea: { width: "100%", minHeight: 90, fontFamily: "monospace", fontSize: 11.5, padding: 10, borderRadius: 7, border: "1px solid var(--rule)", background: "var(--input-bg)", color: "var(--ink)", resize: "vertical" },
  warnBanner: { display: "flex", alignItems: "center", gap: 8, background: "var(--warn-bg)", color: "var(--warn-text)", padding: "10px 14px", borderRadius: 8, fontSize: 12.5, marginBottom: 18 },
  monthBar: { marginBottom: 18, display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" },
  transferBadge: { fontSize: 12, color: "var(--ink-2)", background: "var(--paper-2)", padding: "5px 10px", borderRadius: 20, fontWeight: 500 },
  monthSelect: { fontFamily: "'Fraunces', serif", fontSize: 17, fontWeight: 600, background: "transparent", border: "none", borderBottom: "2px solid var(--teal)", padding: "4px 4px 6px", color: "var(--teal)", cursor: "pointer" },
  cardsGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 12, marginBottom: 18 },
  summaryCard: { background: "var(--panel)", borderRadius: 10, padding: "13px 14px 15px", boxShadow: "0 1px 3px var(--shadow)" },
  summaryIcon: { marginBottom: 8 },
  summaryLabel: { fontSize: 10.5, color: "var(--ink-2)", marginBottom: 4, fontWeight: 500 },
  summaryValue: { fontFamily: "'Fraunces', serif", fontWeight: 600 },
  mainGrid: { display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 18 },
  panel: { background: "var(--panel)", borderRadius: 12, padding: "18px 20px 20px", marginBottom: 18, boxShadow: "0 1px 3px var(--shadow)" },
  panelHeaderRow: { display: "flex", alignItems: "center", justifyContent: "space-between" },
  h2: { fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 16, margin: "0 0 12px", color: "var(--teal)" },
  formRow: { display: "flex", gap: 8, flexWrap: "wrap" },
  input: { fontFamily: "'Inter', sans-serif", fontSize: 13.5, padding: "9px 10px", borderRadius: 7, border: "1px solid var(--rule)", background: "var(--input-bg)", color: "var(--ink)", flex: 1, minWidth: 80 },
  addBtn: { display: "flex", alignItems: "center", gap: 5, background: "var(--teal)", color: "var(--on-teal)", border: "none", borderRadius: 7, padding: "9px 14px", fontSize: 13.5, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" },
  cancelBtn: { background: "none", border: "1px solid var(--rule)", borderRadius: 7, padding: "9px 14px", fontSize: 13.5, color: "var(--ink-2)", cursor: "pointer" },
  smallAddBtn: { display: "flex", alignItems: "center", justifyContent: "center", background: "var(--teal)", color: "var(--on-teal)", border: "none", borderRadius: 7, width: 28, height: 28, cursor: "pointer" },
  avgNote: { fontSize: 12.5, color: "var(--ink-2)", marginTop: 12, marginBottom: 0 },
  empty: { fontSize: 13, color: "var(--ink-3)", fontStyle: "italic" },
  weekBlock: { marginBottom: 16 },
  weekHeader: { display: "flex", justifyContent: "space-between", borderBottom: "1.5px dashed var(--rule)", paddingBottom: 6, marginBottom: 6 },
  weekLabel: { fontSize: 12.5, fontWeight: 600, color: "var(--brown)", textTransform: "uppercase", letterSpacing: "0.02em" },
  weekTotal: { fontSize: 13.5, fontWeight: 700, color: "var(--teal)" },
  entryRow: { display: "grid", gridTemplateColumns: "44px 1fr auto 24px", alignItems: "center", gap: 8, padding: "6px 6px", borderRadius: 6, fontSize: 13.5 },
  entryDate: { color: "var(--ink-3)", fontSize: 12 },
  entryPatient: { color: "var(--ink)" },
  entryAmount: { fontWeight: 600, color: "var(--teal)", textAlign: "right" },
  iconBtn: { background: "none", border: "none", color: "var(--coral)", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", padding: 2 },
  helpText: { fontSize: 12.5, color: "var(--ink-2)", marginTop: -6, marginBottom: 12, lineHeight: 1.5 },
  fieldLabel: { display: "flex", flexDirection: "column", gap: 4, fontSize: 12, color: "var(--ink-2)", fontWeight: 600, marginBottom: 10 },
  expenseForm: { display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" },
  expenseRow: { display: "grid", gridTemplateColumns: "1fr 90px 24px", alignItems: "center", gap: 8, padding: "7px 6px", borderRadius: 6, fontSize: 13.5 },
  expenseInput: { fontFamily: "'Inter', sans-serif", fontSize: 13, padding: "5px 7px", borderRadius: 6, border: "1px solid var(--rule)", background: "var(--input-bg)", color: "var(--ink)", textAlign: "right", width: 90 },
  expenseTotalRow: { display: "flex", justifyContent: "space-between", borderTop: "1.5px solid var(--rule)", marginTop: 6, paddingTop: 8, fontWeight: 700, fontSize: 13.5, color: "var(--teal)" },
  compareRow: { display: "flex", justifyContent: "space-between", fontSize: 13, color: "var(--ink-2)", padding: "4px 0" },
  allocTable: { display: "flex", flexDirection: "column", gap: 2 },
  allocRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 2px" },
  allocLabel: { fontSize: 13, color: "var(--ink)" },
  allocInput: { fontFamily: "'Inter', sans-serif", fontSize: 13, padding: "5px 7px", borderRadius: 6, border: "1px solid var(--rule)", background: "var(--input-bg)", color: "var(--ink)", width: 64, textAlign: "right" },
  allocDivider: { fontSize: 10.5, color: "var(--ink-3)", fontStyle: "italic", margin: "4px 0 2px", borderTop: "1px dashed var(--rule)", paddingTop: 6 },
  goalBlock: { padding: "10px 6px", borderRadius: 8, marginBottom: 4 },
  goalTopRow: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  goalName: { fontSize: 13.5, fontWeight: 600, color: "var(--ink)" },
  goalFieldsRow: { display: "flex", gap: 8, marginTop: 6, alignItems: "flex-end", flexWrap: "wrap" },
  miniLabel: { display: "flex", flexDirection: "column", gap: 3, fontSize: 10.5, color: "var(--ink-3)", fontWeight: 600, flex: 1, minWidth: 80 },
  miniInput: { fontFamily: "'Inter', sans-serif", fontSize: 12.5, padding: "6px 7px", borderRadius: 6, border: "1px solid var(--rule)", background: "var(--input-bg)", color: "var(--ink)" },
  goalAccumRow: { display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8, fontSize: 12.5, color: "var(--ink)" },
  applyBtnWide: { display: "flex", alignItems: "center", gap: 4, background: "var(--green)", color: "#fff", border: "none", borderRadius: 6, padding: "5px 9px", fontSize: 11.5, fontWeight: 600, cursor: "pointer" },
  goalPct: { fontSize: 11, color: "var(--ink-2)" },
  goalForm: { display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10, marginBottom: 10 },
  trendWrap: { display: "flex", alignItems: "flex-end", gap: 8, height: 130, marginTop: 4 },
  trendBarGroup: { flex: 1, minWidth: 0, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 },
  trendBarsRow: { display: "flex", alignItems: "flex-end", gap: 3, height: 90, width: "100%", justifyContent: "center" },
  trendBar: { width: 10, borderRadius: "3px 3px 0 0", minHeight: 2, transition: "height 0.3s ease" },
  trendLabel: { fontSize: 9.5, color: "var(--ink-3)", whiteSpace: "nowrap" },
  trendLegend: { display: "flex", gap: 16, fontSize: 11, color: "var(--ink-2)", marginTop: 10 },
  trendLegendItem: { display: "flex", alignItems: "center", gap: 5 },
  trendDot: { width: 8, height: 8, borderRadius: 2, display: "inline-block" },
};
