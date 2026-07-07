/**
 * MasterScheduleGenerator
 *
 * Modal-style component that generates a master schedule using the
 * rotating queue system (5 separate loops: weekday, h12, h3, h4, h5).
 *
 * Props:
 *   year, month        — the month being generated
 *   doctors            — full doctor list [{id, name}]
 *   activeDoctorIds    — ids of doctors active this month (null = all)
 *   holidays           — array of ISO date strings already marked as holidays in config
 *   queueState         — current {weekday,h12,h3,h4,h5,WDQ,H12Q,H3Q,debt} from Supabase
 *   onConfirm(schedule, newQueueState) — called when admin confirms the preview
 *   onClose()          — called to dismiss the modal
 */

import { useState, useMemo, useCallback } from 'react';
import { X, ChevronLeft, ChevronRight, Info, Check } from 'lucide-react';
import { isoDate, dayType, daysInMonth, formatDisplayDate } from './dateUtils';

// ─── Queue definitions ────────────────────────────────────────────────────────
// Default loop order (by doctor NAME, resolved to id at runtime)
const DEFAULT_WDQ_NAMES  = ['อารีรัตน์','ชุติมา','กนกอร','ธัญลักษณ์','วัทนี','ธนวรรณ','ณัชพล','สมิตา','พสิษฐา','ณัฐธิดา','ขนิษฐา','ณัฐพล'];
const DEFAULT_H12Q_NAMES = ['พสิษฐา','ชุติมา','ณัชพล','ณัฐพล','กนกอร','ธัญลักษณ์','ณัฐธิดา','ขนิษฐา','ธนวรรณ','ณัชพล','ณัฐพล','วัทนี','สมิตา','ณัฐธิดา','ขนิษฐา'];
const DEFAULT_H3Q_NAMES  = ['ชุติมา','กนกอร','ธัญลักษณ์','วัทนี','ธนวรรณ','ณัชพล','สมิตา','พสิษฐา','ณัฐธิดา','ขนิษฐา','ณัฐพล'];

const THAI_MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
const WEEKDAY_LABELS = ['อา','จ','อ','พ','พฤ','ศ','ส'];

const LOOP_META = {
  weekday: { label: 'วันธรรมดา',       bg: '#dbeafe', tx: '#1e40af' },
  h12:     { label: 'วันหยุด 1-2 วัน', bg: '#fce7f3', tx: '#9d174d' },
  h3:      { label: 'วันหยุด 3 วัน',   bg: '#d1fae5', tx: '#065f46' },
  h4:      { label: 'วันหยุด 4 วัน',   bg: '#fef3c7', tx: '#92400e' },
  h5:      { label: 'วันหยุด 5+ วัน',  bg: '#ede9fe', tx: '#4c1d95' },
};

function ltFor(n) { return n <= 2 ? 'h12' : n === 3 ? 'h3' : n === 4 ? 'h4' : 'h5'; }

function detectGroups(year, month, holidayDates) {
  const n = daysInMonth(year, month);
  const holSet = new Set(holidayDates);
  const groups = [];
  let cur = [];
  for (let d = 1; d <= n; d++) {
    const date = isoDate(year, month, d);
    const dow = new Date(year, month, d).getDay();
    if (dow === 0 || dow === 6 || holSet.has(date)) cur.push(date);
    else if (cur.length) { groups.push(cur); cur = []; }
  }
  if (cur.length) groups.push(cur);
  return groups;
}

// Build id arrays from name arrays (duplicate names produce duplicate ids — correct for H12Q)
function resolveQueue(nameArr, doctors) {
  return nameArr.map(name => {
    const doc = doctors.find(d => d.name.trim() === name.trim());
    return doc?.id ?? null;
  }).filter(Boolean);
}

function runQueue(queue, lt, ptr, date, avail, debt) {
  // Normal queue: advances ptr, skips inactive + debt−
  const dk = lt === 'weekday' ? 'WDQ' : lt === 'h12' ? 'H12Q' : lt === 'h3' ? 'H3Q' : lt === 'h4' ? 'H4Q' : 'H5Q';
  let p = ptr;
  for (let i = 0; i < queue.length + 4; i++) {
    const id = queue[p % queue.length];
    const dv = debt[id]?.[dk] ?? 0;
    if (dv < 0) {
      if (!debt[id]) debt[id] = {};
      debt[id][dk] = dv + 1;
      p = (p + 1) % queue.length;
      continue;
    }
    const av = avail[id];
    if (!av || !av.active || date < av.start || date > av.end) {
      p = (p + 1) % queue.length;
      continue;
    }
    p = (p + 1) % queue.length;
    return { id, p };
  }
  return { id: null, p: (ptr + 1) % queue.length };
}

// Pull all debt+ doctors from queue IN LOOP ORDER, filter available on date
function drainPriorityQueue(queue, lt, date, avail, debt) {
  const dk = lt === 'weekday' ? 'WDQ' : lt === 'h12' ? 'H12Q' : lt === 'h3' ? 'H3Q' : lt === 'h4' ? 'H4Q' : 'H5Q';
  for (let i = 0; i < queue.length; i++) {
    const id = queue[i];
    const dv = debt[id]?.[dk] ?? 0;
    if (dv <= 0) continue;
    const av = avail[id];
    if (!av || !av.active || date < av.start || date > av.end) continue;
    if (!debt[id]) debt[id] = {};
    debt[id][dk] = dv - 1;
    return id;
  }
  return null;
}

function generateSchedule({ year, month, holidayDates, avail, WDQ, H12Q, H3Q, qp, debt }) {
  const n = daysInMonth(year, month);
  const groups = detectGroups(year, month, holidayDates);
  const holSet = new Set(groups.flat());
  const debtCopy = JSON.parse(JSON.stringify(debt));
  let { weekday: w, h12, h3, h4, h5 } = { ...qp };
  const schedule = {};
  const groupInfos = [];

  groups.forEach(g => {
    const lt = ltFor(g.length);
    const q = lt === 'h12' ? H12Q : H3Q;
    let ptr = lt === 'h12' ? h12 : lt === 'h3' ? h3 : lt === 'h4' ? h4 : h5;
    const assigns = [];
    g.forEach(date => {
      // debt+ doctors go first (priority), in loop order
      const priorityId = drainPriorityQueue(q, lt, date, avail, debtCopy);
      if (priorityId) {
        schedule[date] = priorityId;
        assigns.push({ date, id: priorityId });
      } else {
        const r = runQueue(q, lt, ptr, date, avail, debtCopy);
        ptr = r.p;
        if (r.id) schedule[date] = r.id;
        assigns.push({ date, id: r.id });
      }
    });
    if (lt === 'h12') h12 = ptr;
    else if (lt === 'h3') h3 = ptr;
    else if (lt === 'h4') h4 = ptr;
    else h5 = ptr;
    groupInfos.push({ dates: g, lt, assigns });
  });

  const wdAssigns = [];
  for (let d = 1; d <= n; d++) {
    const date = isoDate(year, month, d);
    if (holSet.has(date)) continue;
    const r = runQueue(WDQ, 'weekday', w, date, avail, debtCopy);
    w = r.p;
    if (r.id) schedule[date] = r.id;
    wdAssigns.push({ date, id: r.id });
  }
  groupInfos.unshift({ dates: wdAssigns.map(a => a.date), lt: 'weekday', assigns: wdAssigns });

  return { schedule, groupInfos, newQp: { weekday: w, h12, h3, h4, h5 } };
}

// ─── Sub-components ──────────────────────────────────────────────────────────

function CalPreview({ year, month, schedule, holidayDates, doctors, groupMap }) {
  const n = daysInMonth(year, month);
  const holSet = new Set(holidayDates);
  const fd = new Date(year, month, 1).getDay();
  const getDoc = id => doctors.find(d => d.id === id);

  const cells = [];
  for (let i = 0; i < fd; i++) cells.push(null);
  for (let d = 1; d <= n; d++) cells.push(isoDate(year, month, d));

  return (
    <div className="grid grid-cols-7 gap-1">
      {WEEKDAY_LABELS.map((w, i) => (
        <div key={w} className={`text-center text-[10px] font-medium py-1 ${i === 0 || i === 6 ? 'text-rose-500' : 'text-slate-400'}`}>{w}</div>
      ))}
      {cells.map((date, i) => {
        if (!date) return <div key={`e${i}`} />;
        const dow = new Date(date).getDay();
        const isWknd = dow === 0 || dow === 6;
        const isNat = !isWknd && holSet.has(date);
        const lt = groupMap[date];
        const docId = schedule[date];
        const doc = docId ? getDoc(docId) : null;
        const meta = lt ? LOOP_META[lt] : null;
        const dayNum = parseInt(date.slice(-2));

        let bg = meta ? meta.bg : isWknd ? '#fee2e2' : isNat ? '#fef3c7' : 'white';
        let border = meta ? `${meta.tx}40` : isWknd ? '#fca5a5' : isNat ? '#fcd34d' : '#e2e8f0';
        let txc = meta ? meta.tx : isWknd ? '#dc2626' : isNat ? '#92400e' : '#334155';

        return (
          <div key={date} style={{ background: bg, borderColor: border }}
            className="border rounded-md p-1 min-h-[44px] flex flex-col items-center">
            <span style={{ color: txc }} className="font-mono text-[11px] font-medium">{dayNum}</span>
            {doc && <span style={{ color: meta?.tx ?? '#334155' }} className="text-[9px] text-center leading-tight mt-0.5 word-break-all">{doc.name}</span>}
          </div>
        );
      })}
    </div>
  );
}

// ─── Main component ──────────────────────────────────────────────────────────

export default function MasterScheduleGenerator({ year, month, doctors, activeDoctorIds, holidays, queueState, onConfirm, onClose }) {
  const activeDocs = activeDoctorIds === null ? doctors : doctors.filter(d => activeDoctorIds.includes(d.id));

  // Resolve queue arrays (name → id)
  const WDQ  = useMemo(() => resolveQueue(queueState.WDQ  ?? DEFAULT_WDQ_NAMES,  doctors), [doctors, queueState.WDQ]);
  const H12Q = useMemo(() => resolveQueue(queueState.H12Q ?? DEFAULT_H12Q_NAMES, doctors), [doctors, queueState.H12Q]);
  const H3Q  = useMemo(() => resolveQueue(queueState.H3Q  ?? DEFAULT_H3Q_NAMES,  doctors), [doctors, queueState.H3Q]);

  // Phase: 'config' | 'preview'
  const [phase, setPhase] = useState('config');

  // National holidays admin marks (ISO dates)
  const [natHolidays, setNatHolidays] = useState(() => {
    // Pre-populate from existing config holidays for this month
    const prefix = `${year}-${String(month + 1).padStart(2,'0')}`;
    return new Set(holidays.filter(h => h.startsWith(prefix)));
  });

  // Doctor availability {id: {active, start, end}} — start/end are ISO dates
  const [avail, setAvail] = useState(() => {
    const n = daysInMonth(year, month);
    const endDate = isoDate(year, month, n);
    const startDate = isoDate(year, month, 1);
    const map = {};
    activeDocs.forEach(d => { map[d.id] = { active: true, start: startDate, end: endDate }; });
    return map;
  });

  // Queue pointers (local copy, admin can adjust)
  const [qp, setQp] = useState({ weekday: queueState.weekday, h12: queueState.h12, h3: queueState.h3, h4: queueState.h4, h5: queueState.h5 });

  // Debt adjustments {docId: {WDQ,H12Q,H3Q,H4Q,H5Q}}
  const [debt, setDebt] = useState(() => JSON.parse(JSON.stringify(queueState.debt ?? {})));

  // Generated result
  const [result, setResult] = useState(null);

  // All holidays (weekends + national)
  const allHolidays = useMemo(() => {
    const n = daysInMonth(year, month);
    const hols = [];
    for (let d = 1; d <= n; d++) {
      const date = isoDate(year, month, d);
      const dow = new Date(year, month, d).getDay();
      if (dow === 0 || dow === 6 || natHolidays.has(date)) hols.push(date);
    }
    return hols;
  }, [year, month, natHolidays]);

  const groups = useMemo(() => detectGroups(year, month, allHolidays), [year, month, allHolidays]);

  const groupMap = useMemo(() => {
    const m = {};
    if (result) result.groupInfos.forEach(gi => gi.dates.forEach(d => { m[d] = gi.lt; }));
    return m;
  }, [result]);

  // Calendar toggle
  const toggleNat = useCallback((date) => {
    const dow = new Date(date).getDay();
    if (dow === 0 || dow === 6) return;
    setNatHolidays(prev => { const s = new Set(prev); s.has(date) ? s.delete(date) : s.add(date); return s; });
  }, []);

  // Queue pointer adjustment
  const adjQp = (key, delta) => {
    const qmap = { weekday: WDQ, h12: H12Q, h3: H3Q, h4: H3Q, h5: H3Q };
    const len = qmap[key]?.length || 1;
    setQp(prev => ({ ...prev, [key]: (prev[key] + delta + len) % len }));
  };

  const adjDebt = (docId, loopKey, delta) => {
    setDebt(prev => {
      const cur = prev[docId]?.[loopKey] ?? 0;
      return { ...prev, [docId]: { ...prev[docId], [loopKey]: Math.max(-2, Math.min(2, cur + delta)) } };
    });
  };

  const handleGenerate = () => {
    const res = generateSchedule({
      year, month,
      holidayDates: allHolidays,
      avail,
      WDQ, H12Q, H3Q,
      qp,
      debt: JSON.parse(JSON.stringify(debt)),
    });
    setResult(res);
    setPhase('preview');
  };

  const handleConfirm = () => {
    if (!result) return;
    const newQueueState = {
      ...queueState,
      ...result.newQp,
      debt: {}, // debt is one-time correction — reset after each generation
    };
    onConfirm(result.schedule, newQueueState);
  };

  const n = daysInMonth(year, month);
  const fd = new Date(year, month, 1).getDay();
  const monthLabel = `${THAI_MONTHS[month]} ${year + 543}`;

  const DEBT_LOOPS = [
    { k: 'WDQ', nm: 'ธรรมดา' },
    { k: 'H12Q', nm: 'หยุด1-2' },
    { k: 'H3Q', nm: 'หยุด3' },
    { k: 'H4Q', nm: 'หยุด4' },
    { k: 'H5Q', nm: 'หยุด5+' },
  ];

  const qRows = [
    { k: 'weekday', q: WDQ,  label: 'วันธรรมดา' },
    { k: 'h12',     q: H12Q, label: 'วันหยุด 1-2 วัน' },
    { k: 'h3',      q: H3Q,  label: 'วันหยุด 3 วัน' },
    { k: 'h4',      q: H3Q,  label: 'วันหยุด 4 วัน' },
    { k: 'h5',      q: H3Q,  label: 'วันหยุด 5 วัน' },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center bg-black/40 overflow-y-auto py-6">
      <div className="bg-white rounded-2xl border border-slate-200 w-full max-w-3xl mx-4 shadow-lg">
        {/* Header */}
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200">
          <div>
            <p className="font-display font-semibold text-slate-800">จัดตารางเวรต้นแบบ</p>
            <p className="text-xs text-slate-400">{monthLabel}</p>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100 transition-colors">
            <X size={18} />
          </button>
        </div>

        <div className="px-5 py-4 space-y-5">

          {phase === 'config' && (
            <>
              {/* Calendar — mark national holidays */}
              <div>
                <p className="text-xs font-medium text-slate-600 mb-1">คลิกวันธรรมดาเพื่อเพิ่ม/ลบวันหยุดนักขัตฤกษ์</p>
                <p className="text-[10px] text-slate-400 mb-2 flex gap-3">
                  <span><span className="inline-block w-3 h-3 rounded-sm bg-red-100 border border-red-200 mr-1" />เสาร์-อาทิตย์ (อัตโนมัติ)</span>
                  <span><span className="inline-block w-3 h-3 rounded-sm bg-amber-100 border border-amber-300 mr-1" />นักขัตฤกษ์ (กดเลือก)</span>
                </p>
                <div className="grid grid-cols-7 gap-1">
                  {WEEKDAY_LABELS.map((w, i) => (
                    <div key={w} className={`text-center text-[10px] font-medium py-1 ${i === 0 || i === 6 ? 'text-rose-500' : 'text-slate-400'}`}>{w}</div>
                  ))}
                  {Array.from({ length: fd }, (_, i) => <div key={`e${i}`} />)}
                  {Array.from({ length: n }, (_, i) => {
                    const d = i + 1;
                    const date = isoDate(year, month, d);
                    const dow = new Date(year, month, d).getDay();
                    const isWknd = dow === 0 || dow === 6;
                    const isNat = natHolidays.has(date);
                    return (
                      <button key={date} onClick={() => toggleNat(date)}
                        className={`rounded-md border p-1 min-h-[40px] flex flex-col items-center transition-colors
                          ${isWknd ? 'bg-red-50 border-red-200 cursor-default' :
                            isNat ? 'bg-amber-50 border-amber-300 cursor-pointer' :
                            'bg-white border-slate-200 cursor-pointer hover:border-teal-300'}`}>
                        <span className={`text-[11px] font-medium ${isWknd ? 'text-rose-600' : isNat ? 'text-amber-700' : 'text-slate-700'}`}>{d}</span>
                        {isNat && <span className="text-[9px] text-amber-600">นขต.</span>}
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Detected groups */}
              <div>
                <p className="text-xs font-medium text-slate-600 mb-1">กลุ่มวันหยุดที่ตรวจพบ</p>
                <div className="flex flex-wrap gap-1.5">
                  {groups.length === 0
                    ? <p className="text-xs text-slate-400">ยังไม่มีวันหยุดนักขัตฤกษ์ — เสาร์-อาทิตย์จะตรวจพบอัตโนมัติ</p>
                    : groups.map((g, i) => {
                      const lt = ltFor(g.length);
                      const meta = LOOP_META[lt];
                      const lbl = g.length === 1 ? `วันที่ ${parseInt(g[0].slice(-2))}` : `วันที่ ${parseInt(g[0].slice(-2))}–${parseInt(g[g.length-1].slice(-2))}`;
                      return (
                        <span key={i} className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-medium border"
                          style={{ background: meta.bg, color: meta.tx, borderColor: `${meta.tx}40` }}>
                          {lbl} · {meta.label}
                        </span>
                      );
                    })
                  }
                </div>
              </div>

              {/* Queue pointers */}
              <div>
                <p className="text-xs font-medium text-slate-600 mb-2">คิวถัดไป (ปรับด้วย ‹ ›)</p>
                <div className="grid grid-cols-1 gap-1.5">
                  {qRows.map(({ k, q, label }) => {
                    const ptr = qp[k];
                    const nextId = q[ptr % q.length];
                    const nextName = doctors.find(d => d.id === nextId)?.name ?? '?';
                    return (
                      <div key={k} className="flex items-center gap-2 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg">
                        <span className="text-[10px] text-slate-500 min-w-[100px]">{label}</span>
                        <button onClick={() => adjQp(k, -1)} className="w-6 h-6 text-xs border border-slate-300 rounded bg-white hover:bg-slate-100 transition-colors">‹</button>
                        <span className="text-xs font-medium min-w-[60px] text-center text-slate-800">{nextName}</span>
                        <button onClick={() => adjQp(k, 1)} className="w-6 h-6 text-xs border border-slate-300 rounded bg-white hover:bg-slate-100 transition-colors">›</button>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Doctor availability */}
              <div>
                <p className="text-xs font-medium text-slate-600 mb-2">แพทย์ที่อยู่เวรเดือนนี้</p>
                <div className="space-y-1">
                  {activeDocs.map(doc => {
                    const av = avail[doc.id] ?? { active: true, start: isoDate(year, month, 1), end: isoDate(year, month, n) };
                    return (
                      <div key={doc.id} className="flex items-center gap-2 py-1.5 border-b border-slate-100">
                        <button onClick={() => setAvail(prev => ({ ...prev, [doc.id]: { ...av, active: !av.active } }))}
                          className={`w-14 text-[10px] px-1 py-1 rounded border text-center transition-colors ${av.active ? 'bg-emerald-50 border-emerald-300 text-emerald-700' : 'bg-slate-50 border-slate-300 text-slate-400'}`}>
                          {av.active ? 'มีเวร' : 'ไม่มีเวร'}
                        </button>
                        <span className={`text-xs min-w-[60px] ${av.active ? 'text-slate-800' : 'text-slate-400'}`}>{doc.name}</span>
                        {av.active && (
                          <>
                            <input type="number" min="1" max={n} value={parseInt(av.start.slice(-2))}
                              onChange={e => setAvail(prev => ({ ...prev, [doc.id]: { ...av, start: isoDate(year, month, parseInt(e.target.value) || 1) } }))}
                              className="w-10 text-[11px] text-center border border-slate-200 rounded px-1 py-0.5" />
                            <span className="text-[10px] text-slate-400">–</span>
                            <input type="number" min="1" max={n} value={parseInt(av.end.slice(-2))}
                              onChange={e => setAvail(prev => ({ ...prev, [doc.id]: { ...av, end: isoDate(year, month, parseInt(e.target.value) || n) } }))}
                              className="w-10 text-[11px] text-center border border-slate-200 rounded px-1 py-0.5" />
                          </>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Debt adjustments */}
              <div>
                <div className="flex items-center justify-between mb-1">
                  <p className="text-xs font-medium text-slate-600">ปรับคิวพิเศษ ±เวร <span className="text-slate-400 font-normal">(− = skip · + = เพิ่ม · cap ±2)</span></p>
                  {Object.keys(debt).length > 0 && (
                    <button onClick={() => setDebt({})} className="text-[11px] font-medium text-slate-400 hover:text-red-600 hover:bg-red-50 px-2 py-0.5 rounded transition-colors">
                      ล้างคิวพิเศษทั้งหมดเป็น 0
                    </button>
                  )}
                </div>
                <div className="overflow-x-auto">
                  <table className="text-[11px] w-full border-collapse min-w-[420px]">
                    <thead>
                      <tr>
                        <th className="text-left py-1 px-2 text-slate-500 font-medium border-b border-slate-200">แพทย์</th>
                        {DEBT_LOOPS.map(({ k, nm }) => (
                          <th key={k} className="text-center py-1 px-1 text-slate-500 font-medium border-b border-slate-200 min-w-[60px]">{nm}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {activeDocs.map(doc => (
                        <tr key={doc.id} className="border-b border-slate-100">
                          <td className="py-1 px-2 text-slate-700">{doc.name}</td>
                          {DEBT_LOOPS.map(({ k }) => {
                            const val = debt[doc.id]?.[k] ?? 0;
                            return (
                              <td key={k} className="py-1 px-1">
                                <div className="flex items-center justify-center gap-0.5">
                                  <button onClick={() => adjDebt(doc.id, k, -1)}
                                    className={`w-5 h-5 text-xs border rounded transition-colors ${val < 0 ? 'bg-red-50 border-red-300' : 'bg-white border-slate-200'}`}>−</button>
                                  <span className={`text-[11px] w-5 text-center font-medium ${val > 0 ? 'text-emerald-700' : val < 0 ? 'text-red-700' : 'text-slate-400'}`}>
                                    {val > 0 ? `+${val}` : val}
                                  </span>
                                  <button onClick={() => adjDebt(doc.id, k, 1)}
                                    className={`w-5 h-5 text-xs border rounded transition-colors ${val > 0 ? 'bg-emerald-50 border-emerald-300' : 'bg-white border-slate-200'}`}>+</button>
                                </div>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <button onClick={handleGenerate}
                className="w-full py-2.5 bg-teal-600 hover:bg-teal-700 text-white text-sm font-medium rounded-lg transition-colors">
                จัดตารางเวรต้นแบบ →
              </button>
            </>
          )}

          {phase === 'preview' && result && (
            <>
              <CalPreview year={year} month={month} schedule={result.schedule}
                holidayDates={allHolidays} doctors={doctors} groupMap={groupMap} />

              {/* Legend */}
              <div className="flex flex-wrap gap-3">
                {Object.entries(LOOP_META).map(([k, v]) => (
                  <span key={k} className="flex items-center gap-1.5 text-[11px]">
                    <span className="w-3 h-3 rounded-sm border inline-block" style={{ background: v.bg, borderColor: `${v.tx}40` }} />
                    <span className="text-slate-500">{v.label}</span>
                  </span>
                ))}
              </div>

              {/* Group details */}
              <div className="space-y-2">
                {result.groupInfos.map((gi, i) => {
                  const meta = LOOP_META[gi.lt];
                  if (gi.lt === 'weekday') return (
                    <div key={i} className="px-3 py-2 rounded-lg text-xs font-medium" style={{ background: meta.bg, color: meta.tx }}>
                      วันธรรมดา: {gi.assigns.filter(a => a.id).length} วัน
                    </div>
                  );
                  const d0 = parseInt(gi.dates[0].slice(-2));
                  const dN = parseInt(gi.dates[gi.dates.length - 1].slice(-2));
                  const lbl = gi.dates.length === 1 ? `วันที่ ${d0}` : `วันที่ ${d0}–${dN}`;
                  return (
                    <div key={i} className="px-3 py-2 rounded-lg" style={{ background: meta.bg }}>
                      <p className="text-[11px] font-medium mb-1" style={{ color: meta.tx }}>{lbl} — {meta.label}</p>
                      <div className="flex flex-wrap gap-3">
                        {gi.assigns.map(a => (
                          <span key={a.date} className="text-[10px]" style={{ color: meta.tx }}>
                            วันที่ {parseInt(a.date.slice(-2))}: {a.id ? doctors.find(d => d.id === a.id)?.name : '⚠️ ไม่มีผู้อยู่เวร'}
                          </span>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* New queue pointers */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                <p className="text-xs font-medium text-slate-600 mb-2">คิวสิ้นสุด → คิวถัดไป</p>
                {qRows.map(({ k, q, label }) => {
                  const ptr = result.newQp[k];
                  const last = doctors.find(d => d.id === q[(ptr - 1 + q.length) % q.length])?.name ?? '?';
                  const next = doctors.find(d => d.id === q[ptr % q.length])?.name ?? '?';
                  return (
                    <div key={k} className="flex justify-between text-xs py-1 border-b border-slate-100 last:border-0">
                      <span className="text-slate-500">{label}</span>
                      <span className="text-slate-700">สิ้นสุด: <b>{last}</b> → ถัดไป: <b>{next}</b></span>
                    </div>
                  );
                })}
              </div>

              <div className="flex gap-2">
                <button onClick={() => setPhase('config')}
                  className="flex-1 py-2 text-sm border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors">
                  ← กลับแก้ไข
                </button>
                <button onClick={handleConfirm}
                  className="flex-2 flex-grow-[2] py-2 text-sm bg-teal-600 hover:bg-teal-700 text-white font-medium rounded-lg transition-colors flex items-center justify-center gap-1.5">
                  <Check size={14} /> ยืนยันและบันทึกตารางต้นแบบ
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
