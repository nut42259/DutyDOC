import React, { useState, useEffect, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Calendar as CalendarIcon, CalendarCheck, Upload, Users, Repeat, Bell, Plus, X, Check,
  ChevronLeft, ChevronRight, Settings, UserCircle, Trash2,
  MessageCircle, Info, ArrowRightLeft, Tag, Shuffle, RotateCcw, LayoutDashboard
} from 'lucide-react';
import {
  getConfig, setConfig, getMonthData, setMonthData,
  getMarketplace, setMarketplace,
  getNotifications, addNotification as dbAddNotification,
  getDoctors, addDoctor, updateDoctor, deleteDoctor,
  getQueueState, setQueueState,
} from './storage';
import LoginScreen from './LoginScreen';
import MasterScheduleGenerator, {
  DEFAULT_WDQ_NAMES, DEFAULT_H12Q_NAMES, DEFAULT_H3Q_NAMES,
  resolveQueue, detectGroups, ltFor, lastNextInLoop,
} from './MasterScheduleGenerator';

/* ---------------------------------- constants ---------------------------------- */

const THAI_MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
const WEEKDAY_LABELS = ['อา','จ','อ','พ','พฤ','ศ','ส'];
const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

const DOCTOR_PALETTE = [
  { bg: 'bg-teal-600', soft: 'bg-teal-50', text: 'text-teal-700', ring: 'ring-teal-500' },
  { bg: 'bg-indigo-600', soft: 'bg-indigo-50', text: 'text-indigo-700', ring: 'ring-indigo-500' },
  { bg: 'bg-rose-600', soft: 'bg-rose-50', text: 'text-rose-700', ring: 'ring-rose-500' },
  { bg: 'bg-amber-600', soft: 'bg-amber-50', text: 'text-amber-700', ring: 'ring-amber-500' },
  { bg: 'bg-emerald-600', soft: 'bg-emerald-50', text: 'text-emerald-700', ring: 'ring-emerald-500' },
  { bg: 'bg-sky-600', soft: 'bg-sky-50', text: 'text-sky-700', ring: 'ring-sky-500' },
  { bg: 'bg-fuchsia-600', soft: 'bg-fuchsia-50', text: 'text-fuchsia-700', ring: 'ring-fuchsia-500' },
  { bg: 'bg-orange-600', soft: 'bg-orange-50', text: 'text-orange-700', ring: 'ring-orange-500' },
  { bg: 'bg-lime-600', soft: 'bg-lime-50', text: 'text-lime-700', ring: 'ring-lime-500' },
  { bg: 'bg-cyan-600', soft: 'bg-cyan-50', text: 'text-cyan-700', ring: 'ring-cyan-500' },
  { bg: 'bg-violet-600', soft: 'bg-violet-50', text: 'text-violet-700', ring: 'ring-violet-500' },
  { bg: 'bg-pink-600', soft: 'bg-pink-50', text: 'text-pink-700', ring: 'ring-pink-500' },
];
const getDoctorColor = (idx) => DOCTOR_PALETTE[idx % DOCTOR_PALETTE.length];

/* ---------------------------------- utils ---------------------------------- */

const pad2 = (n) => String(n).padStart(2, '0');
const isoDate = (y, m, d) => `${y}-${pad2(m + 1)}-${pad2(d)}`;
const genId = () => (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : `id-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
const monthKey = (y, m) => `month-${y}-${pad2(m + 1)}`;
const toCeYear = (y) => (y > 2400 ? y - 543 : y);
const hasOwn = (obj, key) => Object.prototype.hasOwnProperty.call(obj, key);

function formatDisplayDate(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  if (!y || !m || !d) return iso;
  return `${pad2(d)}-${MONTH_ABBR[m - 1]}-${y}`;
}

function dayType(dateStr, holidaySet) {
  const d = new Date(dateStr + 'T00:00:00');
  const dow = d.getDay();
  if (dow === 0 || dow === 6 || holidaySet.has(dateStr)) return 'holiday';
  return 'weekday';
}
const dayTypeLabel = (dateStr, holidaySet) => dayType(dateStr, holidaySet) === 'holiday' ? 'วันหยุด' : 'วันธรรมดา';
function daysInMonth(y, m) { return new Date(y, m + 1, 0).getDate(); }

function computeUsage(doctors, scheduleLike, holidaySet) {
  const used = {};
  doctors.forEach(d => { used[d.id] = { weekday: 0, holiday: 0 }; });
  Object.entries(scheduleLike || {}).forEach(([date, docId]) => {
    if (!docId || !used[docId]) return;
    used[docId][dayType(date, holidaySet)] += 1;
  });
  return used;
}

// Hard rule: nobody may work two calendar-adjacent days in the current schedule.
// (Only checks within the same month — no cross-month lookback in this prototype.)
function adjacentDates(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const base = new Date(y, m - 1, d);
  const prev = new Date(base); prev.setDate(base.getDate() - 1);
  const next = new Date(base); next.setDate(base.getDate() + 1);
  return [isoDate(prev.getFullYear(), prev.getMonth(), prev.getDate()), isoDate(next.getFullYear(), next.getMonth(), next.getDate())];
}
function hasAdjacentAssignment(scheduleLike, date, doctorId) {
  if (!doctorId) return false;
  const [prevIso, nextIso] = adjacentDates(date);
  return scheduleLike[prevIso] === doctorId || scheduleLike[nextIso] === doctorId;
}

// IMPORTANT: different xlsx builds (Node vs. the browser bundle used in
// artifacts) can construct { cellDates: true } Date objects slightly
// differently, which caused a silent one-day shift depending on environment.
// To avoid that entirely, we read cells WITHOUT cellDates and convert the
// raw Excel serial number to a calendar date ourselves — pure arithmetic,
// no Date-object/timezone ambiguity involved.
function excelSerialToISO(serial) {
  const utcMs = Math.round((serial - 25569) * 86400 * 1000);
  const d = new Date(utcMs);
  return isoDate(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}
function parseFlexibleDate(val) {
  if (typeof val === 'number' && isFinite(val)) return excelSerialToISO(val);
  if (val instanceof Date && !isNaN(val)) return isoDate(val.getUTCFullYear(), val.getUTCMonth(), val.getUTCDate());
  const s = String(val).trim();
  let m = s.match(/^(\d{3,4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return isoDate(toCeYear(Number(m[1])), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{3,4})$/);
  if (m) return isoDate(toCeYear(Number(m[3])), Number(m[2]) - 1, Number(m[1]));
  return null;
}

// Exhaustive constraint-satisfaction search: is there ANY assignment of every
// date to a doctor such that everyone ends up with EXACTLY their master
// quota, nobody works two calendar-adjacent days, and nobody works a date
// they marked unavailable? This is a real depth-first backtracking search
// (not a single greedy pass) — at each step it picks the *most constrained*
// remaining date (fewest legal candidates — the classic CSP "minimum
// remaining values" heuristic, which makes dead ends surface fast) and tries
// every legal candidate for it, backtracking across ANY earlier date if a
// later one turns out impossible. Deterministic: same input always produces
// the same output, no randomness. Bounded by `budget` recursive steps so a
// pathological month can't hang the browser — if the budget runs out we
// genuinely don't know whether a solution exists (as opposed to `solved:
// false` reached without hitting the budget, which proves none exists).
function exhaustiveSolveSchedule({ dates, doctors, quota, unavailSet, masterSchedule, holidaySet, budget }) {
  const remaining = {};
  doctors.forEach(d => { remaining[d.id] = { ...(quota[d.id] || { weekday: 0, holiday: 0 }) }; });
  const assign = {};
  const dateIndex = {};
  dates.forEach((d, i) => { dateIndex[d] = i; });
  const unassigned = new Set(dates);

  const neighborsOf = (date) => {
    const i = dateIndex[date];
    const out = [];
    if (i > 0) out.push(dates[i - 1]);
    if (i < dates.length - 1) out.push(dates[i + 1]);
    return out;
  };

  const candidatesFor = (date) => {
    const type = dayType(date, holidaySet);
    const nominal = masterSchedule[date];
    return doctors
      .map(d => d.id)
      .filter(id => (remaining[id]?.[type] || 0) > 0 && !unavailSet[id].has(date) && !neighborsOf(date).some(n => assign[n] === id))
      .sort((a, b) => {
        // Preference only (doesn't affect completeness): try the master's
        // nominal owner first, then whoever needs shifts of this type most.
        if (a === nominal) return -1;
        if (b === nominal) return 1;
        const qa = quota[a]?.[type] || 1, qb = quota[b]?.[type] || 1;
        return ((remaining[b]?.[type] ?? 0) / qb) - ((remaining[a]?.[type] ?? 0) / qa);
      });
  };

  let steps = 0;
  let timedOut = false;

  function backtrack() {
    if (timedOut) return false;
    if (++steps > budget) { timedOut = true; return false; }
    if (unassigned.size === 0) return true;

    let bestDate = null, bestCandidates = null;
    for (const date of unassigned) {
      const c = candidatesFor(date);
      if (bestCandidates === null || c.length < bestCandidates.length) {
        bestDate = date; bestCandidates = c;
        if (c.length === 0) break; // can't do worse than zero candidates
      }
    }
    if (bestCandidates.length === 0) return false;

    unassigned.delete(bestDate);
    const type = dayType(bestDate, holidaySet);
    for (const docId of bestCandidates) {
      assign[bestDate] = docId;
      remaining[docId][type] -= 1;
      if (backtrack()) return true;
      remaining[docId][type] += 1;
      assign[bestDate] = null;
      if (timedOut) { unassigned.add(bestDate); return false; }
    }
    unassigned.add(bestDate);
    return false;
  }

  const solved = backtrack();
  return { solved, timedOut, assign };
}

// Regenerates the ENTIRE current-month schedule from scratch. buildCurrentSchedule
// assigns each doctor EXACTLY their master-schedule quota of weekday and holiday
// shifts, respecting unavailability and the no-adjacent-days rule.
//
// Approach: first try exhaustiveSolveSchedule — a real backtracking search
// that finds a PERFECT assignment whenever one exists, and is the only way
// to actually guarantee that (a fixed heuristic can miss valid solutions a
// full search would find). Only if that search is inconclusive (budget
// exceeded) or proves no perfect assignment exists do we fall back to a
// heuristic: seed from the master schedule, then patch every date whose
// nominal owner is unavailable via an augmenting-path-style search that may
// recursively displace and relocate other doctors' assignments so displaced
// people still end up at their original quota. Only when no such chain of
// relocations exists is a date recorded as a genuine violation (the nominal
// owner is then kept in place as a last resort).
function buildCurrentSchedule({ doctors, year, month, masterSchedule, unavailability, holidaySet }) {
  const total = daysInMonth(year, month);
  const dates = Array.from({ length: total }, (_, i) => isoDate(year, month, i + 1));
  const dateIndex = {};
  dates.forEach((d, i) => { dateIndex[d] = i; });

  const hasMasterData = Object.values(masterSchedule || {}).some(Boolean);
  if (!hasMasterData) {
    const empty = {}; dates.forEach(d => { empty[d] = null; });
    return { schedule: empty, violations: [] };
  }

  // Quota = exactly what the master schedule gives each doctor, per type.
  const quota = computeUsage(doctors, masterSchedule, holidaySet);

  const unavailSet = {};
  doctors.forEach(d => { unavailSet[d.id] = new Set(unavailability[d.id] || []); });

  // Try for a mathematically perfect assignment first. Whenever one exists,
  // this is guaranteed to find it — no heuristic can promise that.
  const exhaustive = exhaustiveSolveSchedule({ dates, doctors, quota, unavailSet, masterSchedule, holidaySet, budget: 300000 });
  if (exhaustive.solved) {
    return { schedule: exhaustive.assign, violations: [] };
  }

  // No perfect assignment found within budget (or proven impossible) — fall
  // back to the chain-relocation heuristic below, which gets as close as
  // possible and honestly flags whatever it couldn't resolve.
  const remaining = {};
  doctors.forEach(d => { remaining[d.id] = { ...(quota[d.id] || { weekday: 0, holiday: 0 }) }; });

  const assign = {};
  dates.forEach(d => { assign[d] = null; });

  const neighborsOf = (date) => {
    const i = dateIndex[date];
    const out = [];
    if (i > 0) out.push(dates[i - 1]);
    if (i < dates.length - 1) out.push(dates[i + 1]);
    return out;
  };

  // Every mutation (placing/clearing a date, adjusting a remaining count) is
  // logged so a failed attempt can be undone exactly, letting us explore
  // chains of relocations without corrupting shared state.
  const log = [];
  const place = (date, docId, type) => {
    log.push(() => { assign[date] = null; remaining[docId][type] += 1; });
    assign[date] = docId;
    remaining[docId][type] -= 1;
  };
  const clear = (date, docId, type) => {
    log.push(() => { assign[date] = docId; remaining[docId][type] -= 1; });
    assign[date] = null;
    remaining[docId][type] += 1;
  };
  const rollbackTo = (mark) => { while (log.length > mark) log.pop()(); };

  // Try to (re)assign `date` to any valid doctor, possibly displacing and
  // recursively relocating whoever's in the way. `inFlight` prevents infinite
  // recursion by tracking dates already being re-solved earlier in this chain.
  function solveDate(date, inFlight, excludeDoctor) {
    if (inFlight.has(date)) return false;
    inFlight.add(date);
    const type = dayType(date, holidaySet);
    const nominal = masterSchedule[date];

    const candidates = doctors
      .map(d => d.id)
      .filter(id => id !== excludeDoctor && !unavailSet[id].has(date))
      .sort((a, b) => {
        if (a === nominal) return -1;
        if (b === nominal) return 1;
        const qa = quota[a]?.[type] || 1, qb = quota[b]?.[type] || 1;
        return ((remaining[b]?.[type] ?? 0) / qb) - ((remaining[a]?.[type] ?? 0) / qa);
      });

    for (const docId of candidates) {
      const mark = log.length;
      if (tryPlace(docId, date, type, inFlight)) return true;
      rollbackTo(mark);
    }
    return false;
  }

  // Try placing `docId` at `date`. Recursively frees capacity (by giving one
  // of docId's other same-type dates to someone else) and resolves adjacency
  // conflicts (by relocating whichever neighbouring date docId currently
  // holds) as needed. Fully rolls back on failure.
  function tryPlace(docId, date, type, inFlight) {
    const mark = log.length;

    if ((remaining[docId]?.[type] || 0) <= 0) {
      const owned = dates.filter(d => assign[d] === docId && d !== date && dayType(d, holidaySet) === type);
      let freed = false;
      for (const od of owned) {
        const innerMark = log.length;
        clear(od, docId, type);
        if (solveDate(od, new Set(inFlight), docId)) { freed = true; break; }
        rollbackTo(innerMark);
      }
      if (!freed) { rollbackTo(mark); return false; }
    }

    const blockers = neighborsOf(date).filter(n => assign[n] === docId);
    for (const b of blockers) {
      const btype = dayType(b, holidaySet);
      clear(b, docId, btype);
      if (!solveDate(b, new Set(inFlight), docId)) { rollbackTo(mark); return false; }
    }

    place(date, docId, type);
    return true;
  }

  // Seed with the master schedule — already balanced & (assuming a sane
  // master) adjacency-safe — then patch every date whose owner can't work.
  dates.forEach(date => {
    const nominal = masterSchedule[date];
    if (nominal && remaining[nominal] && !unavailSet[nominal].has(date)) {
      place(date, nominal, dayType(date, holidaySet));
    }
  });

  const violations = new Set();
  dates.forEach(date => {
    if (assign[date]) return;
    const mark = log.length;
    if (!solveDate(date, new Set(), null)) {
      rollbackTo(mark);
      violations.add(date);
      // Last resort: keep the nominal owner in place even though this date
      // couldn't be made to satisfy every rule (recorded above so the admin
      // can see exactly which dates need manual attention).
      const nominal = masterSchedule[date];
      const fallback = nominal && remaining[nominal]
        ? nominal
        : (doctors.find(d => !neighborsOf(date).some(n => assign[n] === d.id))?.id ?? doctors[0]?.id ?? null);
      if (fallback) {
        const type = dayType(date, holidaySet);
        assign[date] = fallback;
        remaining[fallback][type] = (remaining[fallback][type] ?? 0) - 1;
      }
    }
  });

  return { schedule: assign, violations: [...violations].sort() };
}


/* ---------------------------------- storage helpers (Supabase) ------------------- */

async function storageGet(key, fallback) {
  try {
    if (key === 'config') { const v = await getConfig(); return v ?? fallback; }
    if (key === 'marketplace') { const v = await getMarketplace(); return v ?? fallback; }
    if (key === 'notifications') { const v = await getNotifications(); return v ?? fallback; }
    if (key.startsWith('month-')) { const v = await getMonthData(key); return v ?? fallback; }
    return fallback;
  } catch { return fallback; }
}
async function storageSet(key, value) {
  try {
    if (key === 'config') { await setConfig(value); return true; }
    if (key === 'marketplace') { await setMarketplace(value); return true; }
    if (key === 'notifications') { return true; }
    if (key.startsWith('month-')) { await setMonthData(key, value); return true; }
  } catch (e) { console.error('storageSet failed', key, e); }
  return false;
}

/* ---------------------------------- small UI bits ---------------------------------- */

function EmptyState({ icon: Icon, title, hint }) {
  return (
    <div className="flex flex-col items-center justify-center text-center py-12 px-4 text-slate-400">
      <Icon size={32} className="mb-3 text-slate-300" />
      <p className="font-body font-medium text-slate-500">{title}</p>
      {hint && <p className="font-body text-sm mt-1 max-w-sm">{hint}</p>}
    </div>
  );
}

function ConfirmModal({ open, title, body, confirmLabel = 'ยืนยัน', onConfirm, onCancel, danger }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5 font-body">
        <h3 className="font-display font-semibold text-slate-800 text-lg mb-2">{title}</h3>
        <p className="text-sm text-slate-600 mb-5 whitespace-pre-line">{body}</p>
        <div className="flex justify-end gap-2">
          <button onClick={onCancel} className="px-4 py-2 rounded-lg text-sm font-medium text-slate-600 hover:bg-slate-100 transition-colors">ยกเลิก</button>
          <button onClick={onConfirm} className={`px-4 py-2 rounded-lg text-sm font-medium text-white transition-colors ${danger ? 'bg-red-600 hover:bg-red-700' : 'bg-teal-600 hover:bg-teal-700'}`}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

function MonthNav({ year, month, onShift }) {
  return (
    <div className="flex items-center gap-2">
      <button onClick={() => onShift(-1)} className="p-1.5 rounded-lg hover:bg-slate-100"><ChevronLeft size={18} /></button>
      <span className="font-display font-semibold text-slate-800 w-40 text-center">{THAI_MONTHS[month]} {year + 543}</span>
      <button onClick={() => onShift(1)} className="p-1.5 rounded-lg hover:bg-slate-100"><ChevronRight size={18} /></button>
    </div>
  );
}

function UsageTable({ title, doctors, usage, original }) {
  return (
    <div className="mt-6 overflow-x-auto">
      <p className="font-display font-semibold text-sm text-slate-700 mb-2">{title}</p>
      <table className="w-full text-xs font-body border-collapse">
        <thead>
          <tr className="text-slate-400 border-b border-slate-200">
            <th className="text-left py-1.5 pr-2">แพทย์</th>
            <th className="text-center py-1.5 px-2">วันธรรมดา</th>
            <th className="text-center py-1.5 px-2">วันหยุด</th>
          </tr>
        </thead>
        <tbody>
          {doctors.map((d, i) => {
            const u = usage[d.id] || { weekday: 0, holiday: 0 };
            const o = original ? (original[d.id] || { weekday: 0, holiday: 0 }) : null;
            const wChanged = o && o.weekday !== u.weekday;
            const hChanged = o && o.holiday !== u.holiday;
            const color = getDoctorColor(i);
            return (
              <tr key={d.id} className="border-b border-slate-100">
                <td className="py-1.5 pr-2 flex items-center gap-1.5"><span className={`w-2 h-2 rounded-full ${color.bg}`} />{d.name}</td>
                <td className={`text-center font-mono py-1.5 px-2 ${wChanged ? 'text-amber-600 font-semibold' : ''}`}>{u.weekday}{o ? <span className="text-slate-400 font-normal">({o.weekday})</span> : ''}</td>
                <td className={`text-center font-mono py-1.5 px-2 ${hChanged ? 'text-amber-600 font-semibold' : ''}`}>{u.holiday}{o ? <span className="text-slate-400 font-normal">({o.holiday})</span> : ''}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
      {original && <p className="text-[10px] text-slate-400 mt-1">ตัวเลขในวงเล็บ = จำนวนเดิมก่อนมีการขาย/แลกเวร · สีส้ม = มีการเปลี่ยนแปลงจากเดิม</p>}
    </div>
  );
}

/* ---------------------------------- master schedule queue summary ---------------------------------- */

const QUEUE_LOOP_LABELS = { weekday: 'วันธรรมดา', h12: 'วันหยุด 1-2 วัน', h3: 'วันหยุด 3 วัน', h4: 'วันหยุด 4 วัน', h5: 'วันหยุด 5 วัน' };

function MasterMonthSummary({ year, month, doctors, masterSchedule, holidays, queueState }) {
  const WDQ = resolveQueue(queueState.WDQ ?? DEFAULT_WDQ_NAMES, doctors);
  const H12Q = resolveQueue(queueState.H12Q ?? DEFAULT_H12Q_NAMES, doctors);
  const H3Q = resolveQueue(queueState.H3Q ?? DEFAULT_H3Q_NAMES, doctors);
  const qMap = { weekday: WDQ, h12: H12Q, h3: H3Q, h4: H3Q, h5: H3Q };

  const groups = detectGroups(year, month, holidays);
  const groupDateSet = new Set(groups.flat());
  const datesByType = { weekday: [], h12: [], h3: [], h4: [], h5: [] };
  const total = daysInMonth(year, month);
  for (let d = 1; d <= total; d++) {
    const date = isoDate(year, month, d);
    if (!groupDateSet.has(date)) datesByType.weekday.push(date);
  }
  groups.forEach(g => { datesByType[ltFor(g.length)].push(...g); });

  const rows = ['weekday', 'h12', 'h3', 'h4', 'h5'].map(key => {
    const queue = qMap[key];
    const info = lastNextInLoop(queue, queueState[key] ?? 0);
    const nextName = info ? (doctors.find(d => d.id === info.nextId)?.name ?? '?') : '?';
    const nextLabel = info?.nextHasDup ? `${nextName}${info.nextOcc}` : nextName;
    const datesThisMonth = datesByType[key].filter(d => masterSchedule[d]).sort();
    const storedLastDate = queueState.lastDate?.[key];

    if (datesThisMonth.length > 0) {
      const firstDoc = doctors.find(d => d.id === masterSchedule[datesThisMonth[0]])?.name ?? '?';
      const lastDoc = doctors.find(d => d.id === masterSchedule[datesThisMonth[datesThisMonth.length - 1]])?.name ?? '?';
      // Only trust the "next" prediction if the queue hasn't already moved
      // past this month for this loop (i.e. this month IS the most recent
      // one generated for it) — otherwise a later month already consumed it.
      const isCurrent = storedLastDate === datesThisMonth[datesThisMonth.length - 1];
      return { key, hasData: true, firstDoc, lastDoc, nextLabel, isCurrent };
    }
    const lastName = info ? (doctors.find(d => d.id === info.lastId)?.name ?? '?') : '?';
    const lastLabel = info?.lastHasDup ? `${lastName}${info.lastOcc}` : lastName;
    return { key, hasData: false, lastDateStr: storedLastDate, lastLabel, nextLabel };
  });

  return (
    <div className="mt-4 border border-slate-200 rounded-xl px-3 py-2.5">
      <p className="text-xs font-medium text-slate-700 mb-2">สรุปคิวเดือนนี้</p>
      <div className="space-y-1.5">
        {rows.map(r => (
          <p key={r.key} className="text-[11px] text-slate-600">
            <span className="font-medium text-slate-700">{QUEUE_LOOP_LABELS[r.key]}</span>{' '}
            {r.hasData ? (
              <>เริ่มที่ <b>{r.firstDoc}</b> จบที่ <b>{r.lastDoc}</b>{' '}
                {r.isCurrent ? <>เดือนต่อไปเริ่มที่ <b>{r.nextLabel}</b></> : <span className="text-slate-400">(คิวเดินต่อไปหลังจากเดือนนี้แล้ว)</span>}
              </>
            ) : (
              <>ล่าสุด{r.lastDateStr ? ` ${formatDisplayDate(r.lastDateStr)}` : ''} จบที่ <b>{r.lastLabel}</b> ดังนั้นเวรต่อไปเริ่มที่ <b>{r.nextLabel}</b>{' '}
                <span className="text-slate-400">(ไม่มีวันประเภทนี้ในเดือนนี้ จึงเป็นข้อมูลเก่า)</span>
              </>
            )}
          </p>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------- admin overview ---------------------------------- */

function OverviewTab({ year, month, doctorsWithShifts, hasMasterData, unavailabilityConfirmed, currentScheduleGenerated, scheduleStale, marketplace, unavailability, onGotoTab, onShiftMonth }) {
  const pendingDocs = doctorsWithShifts.filter(d => !unavailabilityConfirmed.includes(d.id));
  const openPosts = marketplace.filter(p => p.status === 'open');

  const total = daysInMonth(year, month);
  const lead = new Date(year, month, 1).getDay();
  const cells = [];
  for (let i = 0; i < lead; i++) cells.push(null);
  for (let d = 1; d <= total; d++) cells.push(isoDate(year, month, d));

  const countFor = (date) => doctorsWithShifts.filter(d => (unavailability[d.id] || []).includes(date)).length;

  const heatStyle = (n) => {
    if (n === 0) return { bg: 'bg-white', text: 'text-slate-300', border: 'border-slate-200' };
    const ratio = n / Math.max(doctorsWithShifts.length, 1);
    if (ratio >= 0.7) return { bg: 'bg-red-200', text: 'text-red-800', border: 'border-red-300' };
    if (ratio >= 0.4) return { bg: 'bg-amber-200', text: 'text-amber-800', border: 'border-amber-300' };
    return { bg: 'bg-amber-50', text: 'text-amber-600', border: 'border-amber-200' };
  };

  const currentStatus = !currentScheduleGenerated
    ? { label: 'ยังไม่ได้จัดเวร', tone: 'amber' }
    : scheduleStale
      ? { label: 'จัดแล้ว แต่มีข้อมูลใหม่หลังจากนั้น', tone: 'amber' }
      : { label: 'จัดเรียบร้อยแล้ว', tone: 'emerald' };

  const StatusCard = ({ label, tone, value, onClick }) => (
    <button onClick={onClick} className={`text-left rounded-xl border p-3 transition-colors ${tone === 'emerald' ? 'border-emerald-200 bg-emerald-50 hover:bg-emerald-100' : 'border-amber-200 bg-amber-50 hover:bg-amber-100'}`}>
      <p className="text-xs font-medium text-slate-500 mb-1">{label}</p>
      <p className={`text-sm font-semibold ${tone === 'emerald' ? 'text-emerald-700' : 'text-amber-700'}`}>{value}</p>
    </button>
  );

  return (
    <div className="max-w-3xl space-y-6">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <MonthNav year={year} month={month} onShift={onShiftMonth} />
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <StatusCard label="ตารางเวรต้นแบบ" tone={hasMasterData ? 'emerald' : 'amber'} value={hasMasterData ? 'ตั้งค่าแล้ว' : 'ยังไม่ได้ตั้งค่า'} onClick={() => onGotoTab('master')} />
        <StatusCard label="ยืนยันวันไม่สะดวก" tone={pendingDocs.length === 0 ? 'emerald' : 'amber'} value={`${doctorsWithShifts.length - pendingDocs.length}/${doctorsWithShifts.length} คนยืนยันแล้ว`} onClick={() => onGotoTab('unavailable')} />
        <StatusCard label="ตารางเวรปัจจุบัน" tone={currentStatus.tone} value={currentStatus.label} onClick={() => onGotoTab('current')} />
      </div>

      {openPosts.length > 0 && (
        <button onClick={() => onGotoTab('marketplace')} className="w-full text-left rounded-xl border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 p-3 transition-colors flex items-center justify-between">
          <div>
            <p className="text-xs font-medium text-slate-500 mb-1">ตลาดแลกเปลี่ยนเวร</p>
            <p className="text-sm font-semibold text-indigo-700">มี {openPosts.length} รายการเปิดอยู่ ต้องการการตอบรับ</p>
          </div>
          <ArrowRightLeft size={18} className="text-indigo-400 shrink-0" />
        </button>
      )}

      {pendingDocs.length > 0 && (
        <div className="rounded-xl border border-slate-200 p-3">
          <p className="text-xs font-medium text-slate-600 mb-2">ยังไม่ยืนยันวันไม่สะดวก ({pendingDocs.length} คน)</p>
          <div className="flex flex-wrap gap-1.5">
            {pendingDocs.map(d => <span key={d.id} className="text-xs bg-slate-100 text-slate-500 px-2 py-1 rounded-full">{d.name}</span>)}
          </div>
        </div>
      )}

      <div>
        <p className="text-xs font-medium text-slate-600 mb-2 flex items-center gap-1">
          <Info size={12} /> แผนที่ความหนาแน่นวันไม่สะดวก — ยิ่งเข้มยิ่งมีคนไม่สะดวกพร้อมกันหลายคน ดูก่อนกด "จัดเวร" เพื่อเช็ควันเสี่ยง
        </p>
        {doctorsWithShifts.length === 0 ? (
          <p className="text-xs text-slate-400">ยังไม่มีแพทย์ที่มีเวรเดือนนี้</p>
        ) : (
          <>
            <div className="grid grid-cols-7 gap-1 mb-1">
              {WEEKDAY_LABELS.map((w, i) => (<div key={w} className={`text-center text-xs font-body font-semibold py-1 ${i === 0 || i === 6 ? 'text-rose-500' : 'text-slate-400'}`}>{w}</div>))}
            </div>
            <div className="grid grid-cols-7 gap-1">
              {cells.map((date, i) => {
                if (!date) return <div key={`b-${i}`} />;
                const n = countFor(date);
                const c = heatStyle(n);
                const dayNum = Number(date.slice(-2));
                return (
                  <div key={date} className={`rounded-lg border p-1.5 min-h-[48px] flex flex-col items-center justify-center ${c.bg} ${c.border}`}>
                    <span className={`font-mono text-[11px] ${c.text}`}>{dayNum}</span>
                    {n > 0 && <span className={`text-[10px] font-semibold ${c.text}`}>{n} คน</span>}
                  </div>
                );
              })}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------- calendar grid (shared) ---------------------------------- */

function CalendarGrid({ year, month, scheduleData, editable, onAssign, allDoctors, selectableDoctors, holidaySet, unavailability, marketplace, compareTo, highlightDoctorId, originalData, violationDates }) {
  const [editingDate, setEditingDate] = useState(null);
  const getDoctor = (id) => allDoctors.find(d => d.id === id);
  const doctorIndex = (id) => allDoctors.findIndex(d => d.id === id);
  const total = daysInMonth(year, month);
  const leadBlanks = new Date(year, month, 1).getDay();
  const cells = [];
  for (let i = 0; i < leadBlanks; i++) cells.push(null);
  for (let d = 1; d <= total; d++) cells.push(isoDate(year, month, d));

  return (
    <div>
      <div className="grid grid-cols-7 gap-1 mb-1">
        {WEEKDAY_LABELS.map((w, i) => (<div key={w} className={`text-center text-xs font-body font-semibold py-1 ${i === 0 || i === 6 ? 'text-rose-500' : 'text-slate-400'}`}>{w}</div>))}
      </div>
      <div className="grid grid-cols-7 gap-1">
        {cells.map((date, i) => {
          if (!date) return <div key={`b-${i}`} />;
          const docId = scheduleData[date] || null;
          const doc = docId ? getDoctor(docId) : null;
          const dIdx = docId ? doctorIndex(docId) : -1;
          const color = dIdx >= 0 ? getDoctorColor(dIdx) : null;
          const type = dayType(date, holidaySet);
          const unavailDoctors = allDoctors.filter(d => (unavailability[d.id] || []).includes(date));
          const hasOpenPost = marketplace && marketplace.some(p => p.date === date && p.status === 'open');
          const dayNum = Number(date.slice(-2));
          const isEditing = editingDate === date;
          const compareId = compareTo ? (compareTo[date] || null) : null;
          const diverged = !!(compareTo && compareId && docId !== compareId);
          const isMine = !!(highlightDoctorId && docId === highlightDoctorId);
          const isViolation = !!(violationDates && violationDates.includes(date));
          const origId = originalData ? (originalData[date] || null) : null;
          const traded = !!(originalData && origId && origId !== docId);
          const titleParts = [];
          if (unavailDoctors.length) titleParts.push(`ไม่สะดวก: ${unavailDoctors.map(d => d.name).join(', ')}`);
          if (diverged) titleParts.push(`เดิมตามตารางต้นแบบ: ${getDoctor(compareId)?.name || '-'}`);
          if (traded) titleParts.push(`ขาย/แลกจาก: ${getDoctor(origId)?.name || '-'}`);

          return (
            <div
              key={date}
              className={`relative rounded-lg border p-1.5 min-h-[64px] flex flex-col gap-1 ${type === 'holiday' ? 'bg-rose-100 border-rose-200' : 'bg-white border-slate-200'} ${diverged ? 'border-l-4 border-l-sky-400' : ''} ${isMine ? `ring-2 ring-offset-1 ${color.ring}` : ''} ${editable ? 'cursor-pointer hover:border-teal-300' : ''}`}
              onClick={() => editable && setEditingDate(date)}
              title={titleParts.join(' · ')}
            >
              <div className="flex items-center justify-between">
                <span className={`font-mono text-[11px] ${type === 'holiday' ? 'text-rose-700' : 'text-slate-500'}`}>{dayNum}</span>
                {hasOpenPost && <span className="w-1.5 h-1.5 rounded-full bg-amber-500" />}
                {isViolation && <span className="text-[10px] leading-none" title="จัดให้ตรงเงื่อนไขไม่ได้แม้ลองสลับเวรหลายคู่แล้ว เจ้าของเวรเดิมจึงอยู่แทนไปก่อน">⚠️</span>}
              </div>
              {isEditing ? (
                <select autoFocus className="text-[11px] font-body border rounded p-0.5 w-full" value={docId || ''} onClick={(e) => e.stopPropagation()} onChange={(e) => { onAssign(date, e.target.value || null); setEditingDate(null); }} onBlur={() => setEditingDate(null)}>
                  <option value="">ว่าง</option>
                  {selectableDoctors.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
                </select>
              ) : doc ? (
                <div className="flex flex-col gap-0.5">
                  <span className={`text-[11px] font-body font-medium rounded px-1 py-0.5 truncate ${color.soft} ${color.text}`}>{doc.name}</span>
                  {traded && <span className="text-[9px] font-body text-slate-400 line-through truncate px-1">{getDoctor(origId)?.name || '-'}</span>}
                </div>
              ) : (
                <span className="text-[10px] font-body text-slate-300">ยังไม่กำหนด</span>
              )}
              {unavailDoctors.length > 0 && <span className="text-[9px] font-body text-slate-400">{unavailDoctors.length} คนไม่สะดวก</span>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------- main app ---------------------------------- */

export default function App() {
  const now = new Date();
  const [loading, setLoading] = useState(true);
  const [currentUser, setCurrentUser] = useState(null);
  const [_selectedDoctorId, _setSelectedDoctorId] = useState(null);
  const [activeTab, setActiveTab] = useState('current');

  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth());

  const [doctors, setDoctors] = useState([]);
  const [holidays, setHolidays] = useState([]);

  const [masterSchedule, setMasterSchedule] = useState({});
  const [masterOriginal, setMasterOriginal] = useState({});
  const [currentSchedule, setCurrentSchedule] = useState({}); // only changes when admin clicks "จัดเวร"
  const [currentScheduleGenerated, setCurrentScheduleGenerated] = useState(false);
  const [scheduleViolations, setScheduleViolations] = useState([]); // dates where all-unavailable fallback was used
  const [scheduleStale, setScheduleStale] = useState(false); // true once inputs changed after the last "จัดเวร"
  const [scheduleOverrides, setScheduleOverrides] = useState({});
  const [unavailability, setUnavailability] = useState({});
  const [unavailabilityConfirmed, setUnavailabilityConfirmed] = useState([]);
  const [activeDoctorIds, setActiveDoctorIds] = useState(null); // null = everyone active
  const [marketplace, setMarketplace] = useState([]);
  const [notifications, setNotifications] = useState([]);
  const [queueState, setQueueStateLocal] = useState(null);
  const [showMasterGen, setShowMasterGen] = useState(false);

  const [toast, setToast] = useState(null);
  const [confirmState, setConfirmState] = useState(null);

  const holidaySet = new Set(holidays);
  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(null), 2500); };
  const activeDoctors = activeDoctorIds === null ? doctors : doctors.filter(d => activeDoctorIds.includes(d.id));

  // The current schedule is intentionally NOT auto-recomputed in the
  // background anymore — admin generates it explicitly (via "จัดเวร") once
  // the master schedule and everyone's availability have settled, to avoid
  // it churning mid-decision. Manual per-day tweaks (scheduleOverrides) sit
  // on top of whatever was last generated.
  const effectiveSchedule = useMemo(() => {
    const eff = { ...currentSchedule };
    Object.keys(scheduleOverrides).forEach(date => { eff[date] = scheduleOverrides[date]; });
    return eff;
  }, [currentSchedule, scheduleOverrides]);

  useEffect(() => {
    (async () => {
      const [dbDoctors, cfg, qs] = await Promise.all([getDoctors(), storageGet('config', { holidays: [] }), getQueueState()]);
      setDoctors(dbDoctors);
      setHolidays(cfg.holidays || []);
      setQueueStateLocal(qs);
      setMarketplace(await storageGet('marketplace', []));
      setNotifications(await storageGet('notifications', []));
      setLoading(false);
    })();
  }, []);

  useEffect(() => {
    (async () => {
      const data = await storageGet(monthKey(year, month), null);

      // Auto-apply recurring rules (saved in queueState) for this month —
      // computed regardless of whether month_data exists yet, so a rule
      // shows up immediately even for a future month with no master
      // schedule saved at all.
      const rawUnavail = (data && data.unavailability) || {};
      const rules = (queueState || {}).recurringRules || [];
      const mergedUnavail = { ...rawUnavail };
      rules.forEach(({ docId, dow, occurrences }) => {
        const total = new Date(year, month + 1, 0).getDate();
        const pad2n = n => String(n).padStart(2,'0');
        const toAdd = [];
        let nth = 0;
        for (let d = 1; d <= total; d++) {
          if (new Date(year, month, d).getDay() === dow) {
            nth++;
            if (occurrences.length === 0 || occurrences.includes(nth)) {
              toAdd.push(`${year}-${pad2n(month+1)}-${pad2n(d)}`);
            }
          }
        }
        if (toAdd.length) {
          mergedUnavail[docId] = [...new Set([...(mergedUnavail[docId] || []), ...toAdd])].sort();
        }
      });

      if (!data) {
        setMasterSchedule({}); setMasterOriginal({}); setCurrentSchedule({}); setCurrentScheduleGenerated(false); setScheduleStale(false); setScheduleViolations([]);
        setScheduleOverrides({}); setUnavailability(mergedUnavail); setUnavailabilityConfirmed([]); setActiveDoctorIds(null);
      } else {
        const master = data.masterSchedule || data.schedule || {}; // data.schedule = legacy fallback
        setMasterSchedule(master);
        // masterOriginal = the baseline before any trades. Older saved months
        // won't have this field yet — fall back to treating the current
        // master as its own baseline (no trade history to show for those).
        setMasterOriginal(data.masterOriginal || master);
        setCurrentSchedule(data.currentSchedule || {});
        setCurrentScheduleGenerated(!!data.currentScheduleGenerated);
        setScheduleViolations(data.scheduleViolations || []);
        setScheduleStale(!!data.scheduleStale);
        setScheduleOverrides(data.scheduleOverrides || {});
        setUnavailability(mergedUnavail);
        setUnavailabilityConfirmed(data.unavailabilityConfirmed || []);
        setActiveDoctorIds(data.activeDoctorIds !== undefined ? data.activeDoctorIds : null);
      }
    })();
  }, [year, month, queueState]);

  useEffect(() => {
    if (!_selectedDoctorId && doctors.length > 0) _setSelectedDoctorId(doctors[0].id);
  }, [doctors, _selectedDoctorId]);

  const getDoctor = (id) => doctors.find(d => d.id === id);

  const addNotification = useCallback((message, lineMessage) => {
    setNotifications(prev =>
      [{ id: genId(), message, lineMessage, ts: new Date().toISOString() }, ...prev].slice(0, 100));
    dbAddNotification(message, lineMessage).catch(console.error);
  }, []);

  const saveConfig = async (next) => { await storageSet('config', next); };
  const saveMonth = async (patch) => {
    const payload = { masterSchedule, masterOriginal, currentSchedule, currentScheduleGenerated, scheduleViolations, scheduleStale, scheduleOverrides, unavailability, unavailabilityConfirmed, activeDoctorIds, ...patch };
    await storageSet(monthKey(year, month), payload);
  };

  const ensureActiveIncludes = (ids) => {
    if (activeDoctorIds === null) return null;
    const set = new Set(activeDoctorIds);
    ids.forEach(id => set.add(id));
    return [...set];
  };

  /* ---------- roster & active-this-month handlers ---------- */

  const addManualDoctor = async () => {
    const nd = await addDoctor('แพทย์ใหม่');
    if (!nd) return;
    const next = [...doctors, nd];
    setDoctors(next);
    const nextActive = ensureActiveIncludes([nd.id]);
    if (nextActive !== null) { setActiveDoctorIds(nextActive); await saveMonth({ activeDoctorIds: nextActive }); }
  };
  const removeDoctor = async (id) => {
    const next = doctors.filter(d => d.id !== id);
    setDoctors(next);
    await deleteDoctor(id);
    const nextActive = activeDoctorIds === null ? null : activeDoctorIds.filter(x => x !== id);
    setActiveDoctorIds(nextActive);
    await saveMonth({ activeDoctorIds: nextActive });
  };
  const editDoctorName = async (id, name) => {
    await updateDoctor(id, { name });
    setDoctors(prev => prev.map(d => d.id === id ? { ...d, name } : d));
  };
  const updateHolidays = async (next) => { setHolidays(next); await saveConfig({ doctors, holidays: next }); };

  const toggleDoctorActive = async (docId) => {
    const base = activeDoctorIds === null ? doctors.map(d => d.id) : activeDoctorIds;
    const next = base.includes(docId) ? base.filter(id => id !== docId) : [...base, docId];
    setActiveDoctorIds(next);
    await saveMonth({ activeDoctorIds: next });
  };

  /* ---------- master schedule handlers ---------- */

  const handleScheduleExcelUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      let nextDoctors = [...doctors];
      const nextMaster = { ...masterSchedule };
      const nextMasterOriginal = { ...masterOriginal };
      const nextOverrides = { ...scheduleOverrides };
      const monthPrefix = `${year}-${pad2(month + 1)}`;
      const newIds = [];
      let count = 0, skippedMonth = 0, skippedBad = 0;

      const firstCellIsDate = rows.length > 0 && !!parseFlexibleDate(rows[0]?.[0]);
      const dataRows = firstCellIsDate ? rows : rows.slice(1);

      dataRows.forEach(r => {
        if (!r || r[0] == null || r[1] == null) return;
        const dateStr = parseFlexibleDate(r[0]);
        const name = String(r[1]).trim();
        if (!dateStr || !name) { skippedBad++; return; }
        if (!dateStr.startsWith(monthPrefix)) { skippedMonth++; return; }
        let doc = nextDoctors.find(d => d.name.trim().toLowerCase() === name.toLowerCase());
        if (!doc) { doc = { id: genId(), name }; nextDoctors = [...nextDoctors, doc]; newIds.push(doc.id); }
        nextMaster[dateStr] = doc.id;
        nextMasterOriginal[dateStr] = doc.id;
        delete nextOverrides[dateStr];
        count++;
      });

      if (count === 0) { showToast('ไม่พบวันที่ที่ตรงกับเดือนนี้ในไฟล์ ตรวจสอบรูปแบบวันที่'); return; }

      setDoctors(nextDoctors);
      await saveConfig({ holidays });
      setMasterSchedule(nextMaster);
      setMasterOriginal(nextMasterOriginal);
      setScheduleOverrides({});
      const nextActive = ensureActiveIncludes(newIds);
      if (nextActive !== null) setActiveDoctorIds(nextActive);
      // A freshly (re)uploaded master schedule can differ from the old one
      // enough that the old current schedule is meaningless — reset it back
      // to "not generated" rather than just flagging it stale, so admin
      // makes a deliberate fresh "จัดเวร" instead of seeing a possibly very
      // wrong leftover schedule.
      setCurrentSchedule({});
      setCurrentScheduleGenerated(false);
      setScheduleViolations([]);
      setScheduleStale(false);
      await saveMonth({ masterSchedule: nextMaster, masterOriginal: nextMasterOriginal, scheduleOverrides: {}, activeDoctorIds: nextActive, currentSchedule: {}, currentScheduleGenerated: false, scheduleStale: false });
      showToast(`นำเข้าตารางเวรต้นแบบ ${count} วันสำเร็จ${skippedMonth ? ` (ข้าม ${skippedMonth} วันที่ไม่ตรงเดือนนี้)` : ''} — ต้องกดจัดเวรใหม่`);
    } catch (err) {
      showToast('อ่านไฟล์ไม่สำเร็จ ตรวจสอบรูปแบบไฟล์ (คอลัมน์ A: วันที่ · B: ชื่อแพทย์)');
    }
    e.target.value = '';
  };

  const manualAssignMaster = (date, docId) => {
    const oldEff = effectiveSchedule[date] || null;
    const newDocId = docId || null;
    const nextStale = currentScheduleGenerated ? true : scheduleStale;

    setMasterSchedule(prevMaster => {
      const nextMaster = { ...prevMaster, [date]: newDocId };
      setMasterOriginal(prevOriginal => {
        const nextOriginal = { ...prevOriginal, [date]: newDocId };
        setScheduleOverrides(prevOverrides => {
          const nextOverrides = { ...prevOverrides };
          delete nextOverrides[date];
          storageSet(monthKey(year, month), { masterSchedule: nextMaster, masterOriginal: nextOriginal, currentSchedule, currentScheduleGenerated, scheduleStale: nextStale, scheduleOverrides: nextOverrides, unavailability, unavailabilityConfirmed, activeDoctorIds });
          return nextOverrides;
        });
        return nextOriginal;
      });
      return nextMaster;
    });
    if (nextStale !== scheduleStale) setScheduleStale(nextStale);

    if (oldEff !== newDocId) {
      const oldName = oldEff ? getDoctor(oldEff)?.name : 'ว่าง';
      const newName = newDocId ? getDoctor(newDocId)?.name : 'ว่าง';
      addNotification(`ตารางเวรต้นแบบวันที่ ${formatDisplayDate(date)} กำหนดเป็น ${newName} (เดิม ${oldName})`, `🔔 ตารางเวรต้นแบบวันที่ ${formatDisplayDate(date)}: ${oldName} → ${newName}`);
    }
  };

  const manualAssignCurrent = (date, docId) => {
    const oldEff = effectiveSchedule[date] || null;
    if (docId && hasAdjacentAssignment({ ...effectiveSchedule, [date]: docId }, date, docId)) {
      showToast(`ทำไม่ได้: จะทำให้ ${getDoctor(docId)?.name} อยู่เวรติดกัน`);
      return;
    }
    setScheduleOverrides(prev => {
      const next = { ...prev, [date]: docId || null };
      storageSet(monthKey(year, month), { masterSchedule, masterOriginal, currentSchedule, currentScheduleGenerated, scheduleStale, scheduleOverrides: next, unavailability, unavailabilityConfirmed, activeDoctorIds });
      return next;
    });
    if (oldEff !== (docId || null)) {
      const oldName = oldEff ? getDoctor(oldEff)?.name : 'ว่าง';
      const newName = docId ? getDoctor(docId)?.name : 'ว่าง';
      addNotification(`ตารางเวรวันที่ ${formatDisplayDate(date)} เปลี่ยนจาก ${oldName} เป็น ${newName}`, `🔔 ตารางเวรวันที่ ${formatDisplayDate(date)} เปลี่ยนแล้ว: ${oldName} → ${newName}`);
    }
  };

  // The only place the current schedule is actually computed — admin
  // triggers this explicitly (via "จัดเวร") once the master schedule and
  // everyone's availability have settled, rather than it recomputing itself
  // in the background on every small change.
  const handleMasterGenConfirm = async (schedule, newQueueState) => {
    // schedule = { [isoDate]: doctorId }
    // Write it into masterSchedule and masterOriginal (treated as a fresh admin-set baseline)
    const nextMaster = { ...masterSchedule, ...schedule };
    const nextOriginal = { ...masterOriginal, ...schedule };
    setMasterSchedule(nextMaster);
    setMasterOriginal(nextOriginal);
    setScheduleOverrides({});
    setCurrentSchedule({});
    setCurrentScheduleGenerated(false);
    setScheduleStale(false);
    await saveMonth({ masterSchedule: nextMaster, masterOriginal: nextOriginal, scheduleOverrides: {}, currentSchedule: {}, currentScheduleGenerated: false, scheduleStale: false });
    // Persist new queue state. newQueueState.debt already reflects only what
    // was actually consumed this generation (see MasterScheduleGenerator's
    // handleConfirm) — untouched debt (e.g. a loop type with no groups this
    // month) must carry forward, not be wiped here too.
    setQueueStateLocal(newQueueState);
    await setQueueState(newQueueState);
    await addNotification(
      `จัดตารางเวรต้นแบบ ${['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][month]} ${year + 543} สำเร็จแล้ว`,
      `📅 จัดตารางเวรต้นแบบ ${['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][month]} ${year + 543} สำเร็จแล้ว`
    );
    setShowMasterGen(false);
    showToast('บันทึกตารางเวรต้นแบบเรียบร้อย');
  };

  const generateCurrentSchedule = async () => {
    const { schedule: next, violations } = buildCurrentSchedule({ doctors: activeDoctors, year, month, masterSchedule, unavailability, holidaySet });
    const violationList = [...violations].sort();
    setCurrentSchedule(next);
    setScheduleViolations(violationList);
    setCurrentScheduleGenerated(true);
    setScheduleStale(false);
    setScheduleOverrides({});
    await saveMonth({ currentSchedule: next, scheduleViolations: violationList, currentScheduleGenerated: true, scheduleStale: false, scheduleOverrides: {} });
    const msg = `จัดเวรตารางเวรปัจจุบันสำหรับเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว${violationList.length ? ` (มี ${violationList.length} วันที่จัดให้ตรงเงื่อนไขไม่ได้แม้ลองสลับเวรหลายคู่แล้ว — เจ้าของเวรเดิมต้องอยู่แทน)` : ''}`;
    await addNotification(msg, `🔀 ${msg}`);
    showToast('จัดเวรเรียบร้อย');
  };

  // Manually wipe the current schedule back to its "not generated" default —
  // admin must press "จัดเวร" again to produce a new one.
  const resetCurrentSchedule = async () => {
    setCurrentSchedule({});
    setCurrentScheduleGenerated(false);
    setScheduleViolations([]);
    setScheduleStale(false);
    setScheduleOverrides({});
    await saveMonth({ currentSchedule: {}, currentScheduleGenerated: false, scheduleViolations: [], scheduleStale: false, scheduleOverrides: {} });
    await addNotification(
      `ล้างตารางเวรปัจจุบันของเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว (ต้องกดจัดเวรใหม่)`,
      `🗑️ ล้างตารางเวรปัจจุบันของเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว`
    );
    showToast('ล้างตารางเวรปัจจุบันแล้ว');
  };

  // Wipe the master schedule (and everything derived from it — overrides,
  // current schedule) back to blank for this month. Unlike a re-upload, this
  // leaves the doctor roster untouched.
  const resetMasterSchedule = async () => {
    setMasterSchedule({});
    setMasterOriginal({});
    setScheduleOverrides({});
    setCurrentSchedule({});
    setCurrentScheduleGenerated(false);
    setScheduleViolations([]);
    setScheduleStale(false);
    await saveMonth({ masterSchedule: {}, masterOriginal: {}, scheduleOverrides: {}, currentSchedule: {}, currentScheduleGenerated: false, scheduleViolations: [], scheduleStale: false });
    await addNotification(
      `ล้างตารางเวรต้นแบบของเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว`,
      `🗑️ ล้างตารางเวรต้นแบบของเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว`
    );
    showToast('ล้างตารางเวรต้นแบบแล้ว');
  };

  /* ---------- unavailability ---------- */

  const toggleUnavailable = (date) => {
    if (!currentDoctorId) return;
    const nextStale = currentScheduleGenerated ? true : scheduleStale;
    setUnavailability(prev => {
      const mine = prev[currentDoctorId] || [];
      const nextMine = mine.includes(date) ? mine.filter(d => d !== date) : [...mine, date].sort();
      const next = { ...prev, [currentDoctorId]: nextMine };
      setUnavailabilityConfirmed(prevConfirmed => {
        const nextConfirmed = prevConfirmed.filter(id => id !== currentDoctorId); // editing again un-confirms
        storageSet(monthKey(year, month), { masterSchedule, masterOriginal, currentSchedule, currentScheduleGenerated, scheduleStale: nextStale, scheduleOverrides, unavailability: next, unavailabilityConfirmed: nextConfirmed, activeDoctorIds });
        return nextConfirmed;
      });
      return next;
    });
    if (nextStale !== scheduleStale) setScheduleStale(nextStale);
  };

  // docId defaults to the logged-in user's own id (used by the doctor-role
  // tab); the admin panel passes whichever doctor is currently selected,
  // which covers the admin's own record too since admin has no separate
  // "doctor" view of themselves.
  const toggleUnavailabilityConfirmed = (docId = currentDoctorId) => {
    if (!docId) return;
    setUnavailabilityConfirmed(prev => {
      const next = prev.includes(docId) ? prev.filter(id => id !== docId) : [...prev, docId];
      storageSet(monthKey(year, month), { masterSchedule, masterOriginal, currentSchedule, currentScheduleGenerated, scheduleStale, scheduleOverrides, unavailability, unavailabilityConfirmed: next, activeDoctorIds });
      return next;
    });
  };

  // Compute dates in a given year/month matching a dow+occurrences rule
  const recurringDatesForMonth = (y, m, dowTarget, occurrences) => {
    const total = daysInMonth(y, m);
    const dates = [];
    let nth = 0;
    for (let d = 1; d <= total; d++) {
      if (new Date(y, m, d).getDay() === dowTarget) {
        nth++;
        if (occurrences.length === 0 || occurrences.includes(nth)) {
          dates.push(isoDate(y, m, d));
        }
      }
    }
    return dates;
  };

  // Which of this month's unavailable dates came from a standing recurring
  // rule vs. a one-off manual click — used to give recurring dates a visually
  // distinct color in the calendar so it's clear which is which.
  const recurringDatesByDoctor = {};
  ((queueState || {}).recurringRules || []).forEach(({ docId, dow, occurrences }) => {
    const dates = recurringDatesForMonth(year, month, dow, occurrences);
    if (!recurringDatesByDoctor[docId]) recurringDatesByDoctor[docId] = new Set();
    dates.forEach(d => recurringDatesByDoctor[docId].add(d));
  });
  const isRecurringUnavailable = (docId, date) => recurringDatesByDoctor[docId]?.has(date) ?? false;

  // Apply recurring pattern to current month AND save rule for all future months


  const applyRecurringUnavailable = (docId, dowTarget, occurrences) => {
    if (!docId) return;
    // The rule itself must always be saved, even if THIS month happens to have
    // zero matching dates (e.g. no 5th Tuesday) — it's a standing rule meant to
    // auto-apply to whichever future month does have that occurrence.
    const toAdd = recurringDatesForMonth(year, month, dowTarget, occurrences);
    if (toAdd.length) {
      setUnavailability(prev => {
        const merged = [...new Set([...(prev[docId]||[]), ...toAdd])].sort();
        const updated = { ...prev, [docId]: merged };
        setUnavailabilityConfirmed(prevC => {
          const nextC = prevC.filter(id => id !== docId);
          storageSet(monthKey(year,month), { masterSchedule, masterOriginal, currentSchedule, currentScheduleGenerated, scheduleStale: currentScheduleGenerated?true:scheduleStale, scheduleOverrides, unavailability:updated, unavailabilityConfirmed:nextC, activeDoctorIds });
          return nextC;
        });
        return updated;
      });
    }
    setQueueStateLocal(prev => {
      if (!prev) return prev;
      const rules = (prev.recurringRules||[]).filter(r=>!(r.docId===docId&&r.dow===dowTarget));
      const next = { ...prev, recurringRules:[...rules,{docId,dow:dowTarget,occurrences}] };
      setQueueState(next).catch(console.error);
      return next;
    });
    showToast(toAdd.length ? 'บันทึกวันไม่สะดวกประจำเรียบร้อย' : 'บันทึก rule แล้ว — เดือนนี้ไม่มีวันดังกล่าว จะ apply อัตโนมัติในเดือนที่มี');
  };

  const clearUnavailableMonth = (docId) => {
    if (!docId) return;
    setUnavailability(prev => {
      const updated = { ...prev, [docId]: [] };
      setUnavailabilityConfirmed(prevC => {
        const nextC = prevC.filter(id => id !== docId);
        storageSet(monthKey(year,month), { masterSchedule, masterOriginal, currentSchedule, currentScheduleGenerated, scheduleStale: currentScheduleGenerated?true:scheduleStale, scheduleOverrides, unavailability:updated, unavailabilityConfirmed:nextC, activeDoctorIds });
        return nextC;
      });
      return updated;
    });
    showToast('ล้างวันไม่สะดวกเดือนนี้แล้ว');
  };;

  // Clear all unavailability for a doctor in the current month


  /* ---------- marketplace ---------- */

  const myAssignedDates = (docId) => {
    const source = currentScheduleGenerated ? effectiveSchedule : masterSchedule;
    return Object.entries(source).filter(([, id]) => id === docId).map(([d]) => d).sort();
  };

  const createPost = (date, type, targetDoctorId, requestedDate) => {
    const post = { id: genId(), date, posterId: currentDoctorId, type, targetDoctorId: targetDoctorId || null, requestedDate: requestedDate || null, status: 'open', takerId: null, createdAt: new Date().toISOString() };
    setMarketplace(prev => {
      const next = [post, ...prev];
      storageSet('marketplace', next);
      return next;
    });
    const posterName = getDoctor(currentDoctorId)?.name;
    const targetName = targetDoctorId ? getDoctor(targetDoctorId)?.name : null;
    const desc = type === 'swap'
      ? `แลกเวร: วันที่ ${formatDisplayDate(date)} (ของฉัน) ↔ วันที่ ${formatDisplayDate(requestedDate)} (ของ ${targetName})`
      : (targetName ? `ขายเวรวันที่ ${formatDisplayDate(date)} ให้ ${targetName} โดยเฉพาะ` : `ขายเวรวันที่ ${formatDisplayDate(date)} (เปิดให้ทุกคน)`);
    addNotification(`${posterName} ลงประกาศ${desc}`, `📢 ${posterName} ลงประกาศ${desc}`);
    showToast('ลงประกาศเรียบร้อย');
  };

  const createBulkSell = (dates, targetDoctorId) => {
    if (dates.length === 0) { showToast('ไม่มีเวรที่ยังไม่ได้ลงขาย'); return; }
    const newPosts = dates.map(date => ({
      id: genId(), date, posterId: currentDoctorId, type: 'sell', targetDoctorId: targetDoctorId || null, requestedDate: null,
      status: 'open', takerId: null, createdAt: new Date().toISOString(),
    }));
    setMarketplace(prev => {
      const next = [...newPosts, ...prev];
      storageSet('marketplace', next);
      return next;
    });
    const posterName = getDoctor(currentDoctorId)?.name;
    const targetName = targetDoctorId ? getDoctor(targetDoctorId)?.name : null;
    const desc = targetName ? `ขายเวรทั้งหมด ${dates.length} วัน ให้ ${targetName} โดยเฉพาะ` : `ขายเวรทั้งหมด ${dates.length} วัน (เปิดให้ทุกคน)`;
    addNotification(`${posterName} ลงประกาศ${desc}`, `📢 ${posterName} ลงประกาศ${desc}`);
    showToast(`ลงประกาศขายเวร ${dates.length} วันเรียบร้อย`);
  };

  const cancelPost = (postId) => {
    setMarketplace(prev => {
      const next = prev.map(p => p.id === postId ? { ...p, status: 'cancelled' } : p);
      storageSet('marketplace', next);
      return next;
    });
  };

  const declinePost = (post) => {
    setMarketplace(prev => {
      const next = prev.map(p => p.id === post.id ? { ...p, status: 'cancelled' } : p);
      storageSet('marketplace', next);
      return next;
    });
    const takerName = getDoctor(currentDoctorId)?.name;
    addNotification(`${takerName} ปฏิเสธคำขอวันที่ ${formatDisplayDate(post.date)} จาก ${getDoctor(post.posterId)?.name}`, `❌ ${takerName} ปฏิเสธคำขอวันที่ ${formatDisplayDate(post.date)}`);
  };

  // Completed trades (sell or swap) update the MASTER schedule directly —
  // a permanent reallocation. The current schedule always regenerates from
  // the master afterward (see buildCurrentSchedule), which is what keeps
  // everyone's weekday/holiday totals matching the new quota AND guarantees
  // nobody ends up with adjacent duty days — so no validation is needed here.
  // Uses the functional setState form because accepting several posts in a
  // row (very normal after a bulk sell) previously raced: each call read
  // masterSchedule from a stale closure, so a fast second accept could wipe
  // out the first one's change entirely.
  // A completed trade implies the giver is unavailable on the date(s) they
  // gave away (that's practically why they sold/swapped it), and the new
  // owner obviously must be treated as available there. Both parties'
  // confirmation status resets since their picture changed.
  const acceptPost = (post) => {
    if (post.targetDoctorId && post.targetDoctorId !== currentDoctorId) return;
    // This mutates the CURRENTLY LOADED month's in-memory state directly, so
    // a post belonging to some other month must not be accepted from here —
    // that would silently write the change into the wrong month's data. The
    // marketplace tab's own month nav makes switching to the right month a
    // single click away.
    const postDateMonthKey = monthKey(Number(post.date.slice(0, 4)), Number(post.date.slice(5, 7)) - 1);
    if (postDateMonthKey !== monthKey(year, month)) {
      showToast('กรุณาเปลี่ยนไปเดือนของเวรนี้ก่อน (ที่แท็บตลาดแลกเปลี่ยน) แล้วค่อยรับ/ยืนยันอีกครั้ง');
      return;
    }
    const takerId = currentDoctorId;
    const nextStale = currentScheduleGenerated ? true : scheduleStale;

    if (post.type === 'swap') {
      // SWAP: only touches the current schedule (overrides) and unavailability.
      // Master schedule and quotas stay exactly the same.
      // By requesting the swap, the poster implicitly can't do their original
      // date → mark it unavailable for them automatically.
      setScheduleOverrides(prevOverrides => {
        const nextOverrides = { ...prevOverrides, [post.date]: takerId, [post.requestedDate]: post.posterId };
        setUnavailability(prevUnavail => {
          const nextUnavail = { ...prevUnavail };
          // poster can't do post.date; taker can't do post.requestedDate
          const addUnavail = (docId, date) => {
            const list = nextUnavail[docId] || [];
            if (!list.includes(date)) nextUnavail[docId] = [...list, date].sort();
          };
          addUnavail(post.posterId, post.date);
          addUnavail(takerId, post.requestedDate);
          storageSet(monthKey(year, month), { masterSchedule, masterOriginal, currentSchedule, currentScheduleGenerated, scheduleViolations, scheduleStale: nextStale, scheduleOverrides: nextOverrides, unavailability: nextUnavail, unavailabilityConfirmed, activeDoctorIds });
          return nextUnavail;
        });
        return nextOverrides;
      });
    } else {
      // SELL: permanent quota transfer — must update master so the next
      // "จัดเวร" reflects the new ownership, and mark the seller unavailable
      // on that date (they gave it away for a reason).
      setMasterSchedule(prevMaster => {
        const nextMaster = { ...prevMaster, [post.date]: takerId };
        setUnavailability(prevUnavail => {
          const nextUnavail = {};
          Object.keys(prevUnavail).forEach(id => { nextUnavail[id] = [...(prevUnavail[id] || [])]; });
          const list = nextUnavail[post.posterId] || [];
          if (!list.includes(post.date)) nextUnavail[post.posterId] = [...list, post.date].sort();
          const takerList = nextUnavail[takerId] || [];
          nextUnavail[takerId] = takerList.filter(d => d !== post.date);
          setUnavailabilityConfirmed(prevConfirmed => {
            const nextConfirmed = prevConfirmed.filter(id => id !== post.posterId && id !== takerId);
            storageSet(monthKey(year, month), { masterSchedule: nextMaster, masterOriginal, currentSchedule, currentScheduleGenerated, scheduleViolations, scheduleStale: nextStale, scheduleOverrides, unavailability: nextUnavail, unavailabilityConfirmed: nextConfirmed, activeDoctorIds });
            return nextConfirmed;
          });
          return nextUnavail;
        });
        return nextMaster;
      });
    }
    if (nextStale !== scheduleStale) setScheduleStale(nextStale);

    setMarketplace(prevMarket => {
      const nextMarket = prevMarket.map(p => p.id === post.id ? { ...p, status: 'completed', takerId } : p);
      storageSet('marketplace', nextMarket);
      return nextMarket;
    });

    const posterName = getDoctor(post.posterId)?.name;
    const takerName = getDoctor(takerId)?.name;
    const msg = post.type === 'sell'
      ? `${takerName} รับเวรวันที่ ${formatDisplayDate(post.date)} ต่อจาก ${posterName} (ปรับตารางต้นแบบแล้ว)`
      : `${takerName} แลกเวรกับ ${posterName}: วันที่ ${formatDisplayDate(post.date)} ↔ วันที่ ${formatDisplayDate(post.requestedDate)} (ปรับตารางเวรปัจจุบันแล้ว ต้นแบบไม่เปลี่ยน)`;
    addNotification(msg, `✅ ${msg}`);
    showToast('บันทึกเรียบร้อย');
  };

  /* ---------- render helpers ---------- */

  if (!currentUser) return <LoginScreen doctors={doctors} onLogin={(doc) => { setCurrentUser(doc); if (doc.role === 'admin') setActiveTab('overview'); }} />;
  const role = currentUser.role;
  // currentDoctorId is always the logged-in user's own id, regardless of role
  const currentDoctorId = currentUser.id;

  const shiftMonth = (delta) => {
    let m = month + delta, y = year;
    if (m < 0) { m = 11; y -= 1; }
    if (m > 11) { m = 0; y += 1; }
    setMonth(m); setYear(y);
  };

  const currentUsage = computeUsage(activeDoctors, effectiveSchedule, holidaySet);
  const masterUsage = computeUsage(activeDoctors, masterSchedule, holidaySet);
  const masterOriginalUsage = computeUsage(activeDoctors, masterOriginal, holidaySet);
  const hasMasterData = Object.values(masterSchedule || {}).some(Boolean);
  const highlightDoctorId = currentDoctorId;  // highlight own shifts for everyone
  const doctorsWithShifts = activeDoctors.filter(d => (masterUsage[d.id]?.weekday || 0) + (masterUsage[d.id]?.holiday || 0) > 0);
  const pendingConfirmDocs = doctorsWithShifts.filter(d => !unavailabilityConfirmed.includes(d.id));

  const tabs = role === 'admin'
    ? [
        { id: 'overview', label: 'ภาพรวม', icon: LayoutDashboard },
        { id: 'current', label: 'ตารางเวรปัจจุบัน', icon: CalendarCheck },
        { id: 'master', label: 'ตารางเวรต้นแบบ', icon: CalendarIcon },
        { id: 'config', label: 'ตั้งค่า', icon: Settings },
        { id: 'unavailable', label: 'วันไม่สะดวก', icon: UserCircle, badge: hasMasterData ? pendingConfirmDocs.length : 0 },
        { id: 'marketplace', label: 'ตลาดแลกเปลี่ยน', icon: Repeat, badge: marketplace.filter(p => p.status === 'open' && (p.posterId === currentDoctorId || p.targetDoctorId === currentDoctorId)).length },
        { id: 'notifications', label: 'แจ้งเตือน', icon: Bell },
      ]
    : [
        { id: 'current', label: 'ตารางเวรปัจจุบัน', icon: CalendarCheck },
        { id: 'master', label: 'ตารางเวรต้นแบบ', icon: CalendarIcon },
        { id: 'unavailable', label: 'แจ้งวันไม่สะดวก', icon: UserCircle, badge: (currentDoctorId && !unavailabilityConfirmed.includes(currentDoctorId) && (masterUsage[currentDoctorId]?.weekday || 0) + (masterUsage[currentDoctorId]?.holiday || 0) > 0) ? 1 : 0 },
        { id: 'marketplace', label: 'ตลาดแลกเปลี่ยนเวร', icon: Repeat },
        { id: 'notifications', label: 'แจ้งเตือน', icon: Bell },
      ];
  const myMasterShiftCount = currentDoctorId ? (masterUsage[currentDoctorId]?.weekday || 0) + (masterUsage[currentDoctorId]?.holiday || 0) : 0;

  if (loading) return <div className="min-h-screen bg-slate-50 flex items-center justify-center text-slate-400 text-sm">กำลังโหลดข้อมูล…</div>;

  return (
    <div className="font-body bg-slate-50 min-h-[600px] rounded-2xl overflow-hidden border border-slate-200">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        .font-display { font-family: 'Space Grotesk', sans-serif; }
        .font-body { font-family: 'Inter', sans-serif; }
        .font-mono { font-family: 'IBM Plex Mono', monospace; }
      `}</style>

      <div className="bg-white border-b border-slate-200 px-4 sm:px-6 py-3 flex flex-wrap items-center gap-3 justify-between">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-teal-600 flex items-center justify-center"><CalendarIcon size={16} className="text-white" /></div>
          <span className="font-display font-semibold text-slate-800 text-lg">ระบบจัดเวรแพทย์ DutyDOC_PedBMA</span>
        </div>
        <div className="flex items-center gap-2">
          <div className="flex items-center gap-2">
            <span className="text-sm font-medium text-slate-700">{currentUser.name}</span>
            {role === 'admin' && <span className="text-[10px] bg-teal-100 text-teal-700 px-2 py-0.5 rounded-full font-medium">Admin</span>}
            <button onClick={() => setCurrentUser(null)} className="text-xs text-slate-400 hover:text-slate-700 px-2 py-1 rounded-lg hover:bg-slate-100 transition-colors">ออกจากระบบ</button>
          </div>
        </div>
      </div>

      <div className="bg-white border-b border-slate-200 px-4 sm:px-6 flex gap-1 overflow-x-auto">
        {tabs.map(t => {
          const Icon = t.icon;
          const active = activeTab === t.id;
          return (
            <button key={t.id} onClick={() => setActiveTab(t.id)} className={`flex items-center gap-1.5 px-3 py-2.5 text-sm font-medium border-b-2 whitespace-nowrap transition-colors ${active ? 'border-teal-600 text-teal-700' : 'border-transparent text-slate-500 hover:text-slate-700'}`}>
              <span className="relative inline-flex">
                <Icon size={14} />
                {t.badge > 0 && <span className="absolute -top-1.5 -right-2 min-w-[14px] h-3.5 bg-rose-500 text-white text-[8px] font-bold rounded-full flex items-center justify-center px-0.5 leading-none">{t.badge}</span>}
              </span>
              {t.label}
            </button>
          );
        })}
      </div>

      <div className="p-4 sm:p-6">

        {/* OVERVIEW TAB (admin) */}
        {activeTab === 'overview' && role === 'admin' && (
          <OverviewTab
            year={year} month={month}
            doctorsWithShifts={doctorsWithShifts} hasMasterData={hasMasterData}
            unavailabilityConfirmed={unavailabilityConfirmed}
            currentScheduleGenerated={currentScheduleGenerated} scheduleStale={scheduleStale}
            marketplace={marketplace} unavailability={unavailability}
            onGotoTab={setActiveTab} onShiftMonth={shiftMonth}
          />
        )}

        {/* CURRENT / EFFECTIVE SCHEDULE TAB */}
        {activeTab === 'current' && (
          <div>
            <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
              <MonthNav year={year} month={month} onShift={shiftMonth} />
              {role === 'admin' && hasMasterData && (
                <div className="flex items-center gap-2">
                  {(() => {
                    const pendingDocs = doctorsWithShifts.filter(d => !unavailabilityConfirmed.includes(d.id));
                    return (
                      <button onClick={() => setConfirmState({ type: 'rearrange', pendingCount: pendingDocs.length })}
                        className="flex items-center gap-1.5 bg-teal-600 hover:bg-teal-700 text-white text-sm font-medium px-3 py-2 rounded-lg transition-colors">
                        <Shuffle size={14} /> จัดเวร
                        {pendingDocs.length > 0 && (
                          <span className="bg-white/30 text-[10px] font-bold px-1.5 py-0.5 rounded-full leading-none">รอ {pendingDocs.length}</span>
                        )}
                      </button>
                    );
                  })()}
                  {currentScheduleGenerated && (
                    <button
                      onClick={() => setConfirmState({ type: 'clear-current' })}
                      className="flex items-center gap-1.5 text-sm font-medium text-slate-500 hover:text-red-600 hover:bg-red-50 px-3 py-2 rounded-lg transition-colors"
                    >
                      <RotateCcw size={14} /> ล้างตารางเวรปัจจุบัน
                    </button>
                  )}
                </div>
              )}
            </div>
            <p className="text-xs text-slate-400 mb-3 flex items-center gap-1"><Info size={12} /> ตารางนี้จะจัดก็ต่อเมื่อแอดมินกด "จัดเวร" เท่านั้น (ไม่จัดอัตโนมัติ) เพื่อให้รอโควตาต้นแบบและการแจ้งไม่สะดวกนิ่งก่อน — จำนวนเวรวันธรรมดา/วันหยุดของแต่ละคนจะเท่ากับตารางต้นแบบ และไม่มีใครอยู่เวรติดกัน{role === 'admin' ? ' คลิกวันที่เพื่อแก้ไขเฉพาะจุดเองได้หลังจัดแล้ว' : ''}</p>

            {role === 'admin' && currentScheduleGenerated && scheduleStale && (
              <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2 mb-4 flex items-start gap-2">
                <Info size={14} className="mt-0.5 shrink-0" />
                มีการเปลี่ยนแปลงข้อมูล (ตารางต้นแบบ แจ้งไม่สะดวก หรือขาย/แลกเวร) หลังจากจัดเวรครั้งล่าสุด กด "จัดเวร" อีกครั้งเพื่อให้ตารางเวรปัจจุบันตรงกับข้อมูลล่าสุด
              </div>
            )}

            {role === 'admin' && hasMasterData && doctorsWithShifts.length > 0 && (
              <div className="border border-slate-200 rounded-xl px-3 py-2.5 mb-4">
                <p className="text-xs font-medium text-slate-600 mb-2">
                  สถานะแจ้งวันไม่สะดวก: {doctorsWithShifts.filter(d => unavailabilityConfirmed.includes(d.id)).length}/{doctorsWithShifts.length} คนยืนยันแล้ว
                  <span className="text-slate-400 font-normal"> (นับเฉพาะคนที่มีเวรเดือนนี้)</span>
                </p>
                <div className="flex flex-wrap gap-1.5">
                  {activeDoctors.map((d, i) => {
                    const hasShift = doctorsWithShifts.some(x => x.id === d.id);
                    const confirmed = unavailabilityConfirmed.includes(d.id);
                    const color = getDoctorColor(doctors.findIndex(x => x.id === d.id));
                    return (
                      <span key={d.id} className={`flex items-center gap-1 text-xs px-2 py-1 rounded-full ${!hasShift ? 'bg-slate-50 text-slate-300' : confirmed ? 'bg-emerald-50 text-emerald-700' : 'bg-slate-100 text-slate-400'}`}>
                        <span className={`w-1.5 h-1.5 rounded-full ${color.bg}`} style={{ opacity: hasShift ? 1 : 0.3 }} />{d.name}
                        {!hasShift ? <span className="text-[10px]">ไม่มีเวร</span> : confirmed ? <Check size={11} /> : <span className="text-[10px]">รอยืนยัน</span>}
                      </span>
                    );
                  })}
                </div>
              </div>
            )}

            {activeDoctors.length === 0 ? (
              <EmptyState icon={Users} title="ยังไม่มีแพทย์ที่อยู่เวรเดือนนี้" hint="ไปที่แท็บ 'ตั้งค่า' เพื่อเพิ่ม/เลือกแพทย์ที่อยู่เวรเดือนนี้ก่อน" />
            ) : !hasMasterData ? (
              <EmptyState icon={CalendarIcon} title="ยังไม่มีตารางเวรต้นแบบของเดือนนี้" hint="ไปที่แท็บ 'ตารางเวรต้นแบบ' เพื่อกำหนดตารางเวรต้นแบบก่อน ตารางเวรปัจจุบันจะคำนวณจากตารางต้นแบบเท่านั้น" />
            ) : !currentScheduleGenerated ? (
              <EmptyState icon={Shuffle} title="ยังไม่ได้จัดตารางเวรปัจจุบัน" hint={role === 'admin' ? 'รอให้ทุกคนแจ้งวันไม่สะดวกและยืนยันครบ (ดูสถานะด้านบน) แล้วกดปุ่ม "จัดเวร" เพื่อเริ่มจัด' : 'รอแอดมินกดจัดเวร'} />
            ) : (
              <>
                <CalendarGrid
                  year={year} month={month} scheduleData={effectiveSchedule}
                  editable={role === 'admin'} onAssign={manualAssignCurrent}
                  allDoctors={doctors} selectableDoctors={activeDoctors}
                  holidaySet={holidaySet} unavailability={unavailability} marketplace={marketplace}
                  compareTo={currentSchedule} highlightDoctorId={highlightDoctorId}
                  violationDates={scheduleViolations}
                />
                {role === 'admin' && <p className="text-xs text-slate-400 mt-2 flex items-center gap-1"><Info size={12} /> แถบสีฟ้าด้านซ้ายของช่อง = วันนี้ถูกแก้ไขเฉพาะจุดด้วยมือ · ⚠️ = จัดให้ตรงเงื่อนไขไม่ได้แม้ลองสลับเวรหลายคู่แล้ว เจ้าของเวรเดิมจึงอยู่แทน</p>}

                {scheduleViolations.length > 0 && (
                  <div className="mt-3 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 flex items-start gap-2">
                    <span className="text-amber-500 text-base shrink-0">⚠️</span>
                    <div>
                      <p className="text-xs font-medium text-amber-800 mb-0.5">มี {scheduleViolations.length} วันที่จัดไม่ตรงเงื่อนไข</p>
                      <p className="text-xs text-amber-700">ระบบลองสลับ/โยกเวรหลายคู่แล้วแต่ยังหาทางจัดให้ตรงโควตาและไม่ติดกันไม่ได้สำหรับวันเหล่านี้ (อาจเพราะไม่สะดวกทับซ้อนกันหลายคน หรือโควตาไม่พอ) จึงให้เจ้าของเวรเดิมในตารางต้นแบบอยู่เวรแทนไปก่อน — ลองแก้ไขเฉพาะจุดเองด้านบน: {scheduleViolations.map(d => formatDisplayDate(d)).join(', ')}</p>
                    </div>
                  </div>
                )}

                {currentDoctorId && (() => {
                  const myDates = Object.entries(effectiveSchedule).filter(([, id]) => id === currentDoctorId).map(([d]) => d).sort();
                  const myWeekday = myDates.filter(d => dayType(d, holidaySet) === 'weekday');
                  const myHoliday = myDates.filter(d => dayType(d, holidaySet) === 'holiday');
                  return myDates.length > 0 ? (
                    <div className="mt-4 border border-slate-200 rounded-xl px-3 py-2.5">
                      <p className="text-xs font-medium text-slate-700 mb-1">เวรของคุณเดือนนี้</p>
                      {myWeekday.length > 0 && <p className="text-xs text-slate-600">วันธรรมดา: {myWeekday.map(d => formatDisplayDate(d)).join(', ')}</p>}
                      {myHoliday.length > 0 && <p className="text-xs text-slate-600">วันหยุด: {myHoliday.map(d => formatDisplayDate(d)).join(', ')}</p>}
                    </div>
                  ) : null;
                })()}

                <UsageTable title="จำนวนเวรที่จัดแล้วเดือนนี้ (จัดจริง / โควตาต้นแบบล่าสุด)" doctors={activeDoctors} usage={currentUsage} original={masterUsage} />
              </>
            )}
          </div>
        )}

        {/* MASTER SCHEDULE TAB */}
        {activeTab === 'master' && (
          <div>
            <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
              <MonthNav year={year} month={month} onShift={shiftMonth} />
              {role === 'admin' && (
                <div className="flex items-center gap-2">
                  <button onClick={() => setShowMasterGen(true)}
                    className="flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium px-3 py-1.5 rounded-lg transition-colors">
                    <CalendarCheck size={14} /> จัดตารางเวรต้นแบบ
                  </button>
                  <label className="flex items-center gap-1.5 border border-dashed border-slate-300 rounded-lg px-3 py-1.5 cursor-pointer hover:border-teal-400 transition-colors text-sm font-medium text-slate-600">
                    <Upload size={14} className="text-teal-600" /> อัปโหลด .xlsx
                    <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleScheduleExcelUpload} />
                  </label>
                  {hasMasterData && (
                    <button
                      onClick={() => setConfirmState({ type: 'clear-master' })}
                      className="flex items-center gap-1.5 text-sm font-medium text-slate-500 hover:text-red-600 hover:bg-red-50 px-3 py-1.5 rounded-lg transition-colors"
                    >
                      <RotateCcw size={14} /> ล้างตารางเวรต้นแบบ
                    </button>
                  )}
                </div>
              )}
            </div>
            <p className="text-xs text-slate-400 mb-3 flex items-center gap-1"><Info size={12} /> นี่คือตารางเวรต้นแบบที่คุณกำหนดเอง (ไฟล์ Excel: คอลัมน์ A วันที่ · B ชื่อแพทย์) การขาย/แลกเวรที่สำเร็จแล้วจะถูกปรับเข้าที่นี่โดยอัตโนมัติเพื่ออัปเดตโควตาล่าสุด{role === 'admin' ? ' — คลิกวันที่เพื่อแก้ไขได้โดยตรง' : ''}</p>

            {activeDoctors.length === 0 ? (
              <EmptyState icon={Users} title="ยังไม่มีแพทย์ที่อยู่เวรเดือนนี้" hint="ไปที่แท็บ 'ตั้งค่า' เพื่อเพิ่ม/เลือกแพทย์ที่อยู่เวรเดือนนี้ก่อน" />
            ) : (
              <>
                <CalendarGrid
                  year={year} month={month} scheduleData={masterSchedule}
                  editable={role === 'admin'} onAssign={manualAssignMaster}
                  allDoctors={doctors} selectableDoctors={activeDoctors}
                  holidaySet={holidaySet} unavailability={unavailability} marketplace={null}
                  compareTo={null} highlightDoctorId={highlightDoctorId} originalData={masterOriginal}
                />
                <p className="text-xs text-slate-400 mt-3 flex items-center gap-1"><Info size={12} /> ชื่อสีเทาขีดฆ่า = เจ้าของเวรเดิมก่อนขายเวร (ไม่ปรากฏสำหรับการแลกเวร) · ชื่อด้านบน = เจ้าของเวรปัจจุบัน</p>
                <UsageTable title="จำนวนเวรที่จัดแล้วเดือนนี้ (ปัจจุบัน(เดิมก่อนขายเวร))" doctors={activeDoctors} usage={masterUsage} original={masterOriginalUsage} />
                {role === 'admin' && hasMasterData && queueState && (
                  <MasterMonthSummary year={year} month={month} doctors={doctors} masterSchedule={masterSchedule} holidays={holidays} queueState={queueState} />
                )}
              </>
            )}
          </div>
        )}

        {/* CONFIG TAB (admin) */}
        {activeTab === 'config' && role === 'admin' && (
          <div className="space-y-8 max-w-3xl">
            <div>
              <p className="font-display font-semibold text-slate-800 mb-1">รายชื่อแพทย์ทั้งหมด</p>
              <div className="flex items-center justify-end mb-2">
                <button onClick={addManualDoctor} className="flex items-center gap-1 text-xs font-medium text-teal-700 hover:bg-teal-50 px-2 py-1 rounded-lg"><Plus size={14} /> เพิ่มแพทย์</button>
              </div>
              {doctors.length === 0 ? (
                <EmptyState icon={Users} title="ยังไม่มีรายชื่อแพทย์" hint="เพิ่มเองด้านบน หรืออัปโหลดตารางเวรต้นแบบซึ่งจะเพิ่มรายชื่อให้อัตโนมัติ" />
              ) : (
                <div className="flex flex-wrap gap-2">
                  {doctors.map((d, i) => {
                    const color = getDoctorColor(i);
                    return (
                      <div key={d.id} className="flex items-center gap-1.5 border border-slate-200 rounded-lg pl-2 pr-1 py-1">
                        <span className={`w-2 h-2 rounded-full ${color.bg}`} />
                        <input value={d.name} onChange={(e) => editDoctorName(d.id, e.target.value)} className="text-sm border-none focus:outline-none w-24" />
                        <button onClick={() => removeDoctor(d.id)} className="text-slate-300 hover:text-red-500 p-0.5"><Trash2 size={13} /></button>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div>
              <p className="font-display font-semibold text-slate-800 mb-1">แพทย์ที่อยู่เวรเดือนนี้ — {THAI_MONTHS[month]} {year + 543}</p>
              <p className="text-xs text-slate-400 mb-2">รายชื่อที่ปิดไว้จะไม่ปรากฏเป็นตัวเลือกในตารางเวร/ตลาดแลกเปลี่ยนของเดือนนี้ (ไม่กระทบเดือนอื่น) รายชื่อใหม่จะเปิดใช้งานอัตโนมัติ</p>
              {doctors.length === 0 ? (
                <p className="text-sm text-slate-400">เพิ่มรายชื่อแพทย์ด้านบนก่อน</p>
              ) : (
                <div className="flex flex-wrap gap-2">
                  {doctors.map((d, i) => {
                    const active = activeDoctorIds === null ? true : activeDoctorIds.includes(d.id);
                    const color = getDoctorColor(i);
                    return (
                      <button key={d.id} onClick={() => toggleDoctorActive(d.id)} className={`flex items-center gap-1.5 border rounded-lg pl-2 pr-2.5 py-1 text-sm transition-colors ${active ? 'border-teal-300 bg-teal-50 text-teal-700' : 'border-slate-200 text-slate-400'}`}>
                        <span className={`w-2 h-2 rounded-full ${color.bg}`} style={{ opacity: active ? 1 : 0.3 }} />
                        {d.name}
                        {active ? <Check size={12} /> : <X size={12} />}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>

            <div>
              <p className="font-display font-semibold text-slate-800 mb-2">วันหยุดนักขัตฤกษ์เพิ่มเติม</p>
              <p className="text-xs text-slate-400 mb-2">เสาร์-อาทิตย์นับเป็นวันหยุดโดยอัตโนมัติแล้ว เพิ่มเฉพาะวันหยุดนักขัตฤกษ์อื่น ๆ</p>
              <div className="flex items-center gap-2 mb-2">
                <input type="date" id="holidayInput" className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm font-mono" />
                <button onClick={() => { const el = document.getElementById('holidayInput'); if (el && el.value) { updateHolidays([...new Set([...holidays, el.value])].sort()); el.value = ''; } }} className="flex items-center gap-1 text-xs font-medium text-teal-700 hover:bg-teal-50 px-2 py-1.5 rounded-lg"><Plus size={14} /> เพิ่ม</button>
              </div>
              <div className="flex flex-wrap gap-1.5">
                {holidays.map(h => (<span key={h} className="flex items-center gap-1 bg-rose-50 text-rose-600 text-xs font-mono px-2 py-1 rounded-full">{formatDisplayDate(h)}<button onClick={() => updateHolidays(holidays.filter(x => x !== h))}><X size={12} /></button></span>))}
              </div>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-xl px-4 py-3">
              <p className="text-sm text-slate-600 flex items-start gap-2"><Info size={14} className="mt-0.5 shrink-0" /> ระยะห่างระหว่างเวรเป็นกฎตายตัวแล้ว: ห้ามอยู่เวรวันติดกันเสมอในตารางเวรปัจจุบัน ไม่ต้องตั้งค่าเพิ่ม</p>
            </div>
          </div>
        )}

        {/* UNAVAILABLE TAB (doctor) */}
        {activeTab === 'unavailable' && role === 'doctor' && (
          <div>
            {!currentDoctorId ? <EmptyState icon={UserCircle} title="ยังไม่มีแพทย์ในระบบ" /> : !hasMasterData ? (
              <EmptyState icon={CalendarIcon} title="ยังไม่มีตารางเวรต้นแบบของเดือนนี้" hint="รอแอดมินกำหนดตารางเวรต้นแบบก่อน" />
            ) : myMasterShiftCount === 0 ? (
              <EmptyState icon={UserCircle} title="คุณไม่มีเวรในเดือนนี้" hint={`${getDoctor(currentDoctorId)?.name} ไม่มีเวรอยู่ในตารางเวรต้นแบบของเดือน ${THAI_MONTHS[month]} ${year + 543} จึงไม่ต้องแจ้งวันไม่สะดวก`} />
            ) : (
              <>
                <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2 mb-4 flex items-start gap-2">
                  <Info size={14} className="mt-0.5 shrink-0" />
                  แจ้งวันไม่สะดวกให้ครบแล้วกดยืนยันด้านล่าง แอดมินจะรอให้ทุกคนยืนยันก่อนกดจัดเวร เพื่อให้จำนวนเวรวันธรรมดา/วันหยุดของทุกคนตรงกับตารางต้นแบบ และไม่มีใครอยู่เวรติดกัน — ถ้าตารางจัดไปแล้วและวันที่คุณอยู่เวรดันไม่สะดวกขึ้นมาทีหลัง คุณลงขาย/แลกเวรเองได้ที่แท็บ "ตลาดแลกเปลี่ยนเวร"
                </div>
                <RecurringUnavailablePanel
                  year={year} month={month}
                  onApply={(dow, occ) => applyRecurringUnavailable(currentDoctorId, dow, occ)}
                />
                <div className="flex items-center justify-between mb-4"><MonthNav year={year} month={month} onShift={shiftMonth} /></div>
                <p className="text-xs text-slate-400 mb-3 flex items-center gap-1"><Info size={12} /> คลิกวันที่เพื่อแจ้ง/ยกเลิกการแจ้งไม่สะดวก ({getDoctor(currentDoctorId)?.name}) · <span className="inline-block w-2.5 h-2.5 rounded-sm bg-indigo-50 border border-indigo-300" /> ไม่สะดวกประจำ (จาก rule) · <span className="inline-block w-2.5 h-2.5 rounded-sm bg-red-50 border border-red-300" /> จิ้มเลือกเอง</p>
                <div className="grid grid-cols-7 gap-1 mb-1">{WEEKDAY_LABELS.map((w, i) => <div key={w} className={`text-center text-xs font-body font-semibold py-1 ${i === 0 || i === 6 ? 'text-rose-500' : 'text-slate-400'}`}>{w}</div>)}</div>
                <div className="grid grid-cols-7 gap-1">
                  {(() => {
                    const total = daysInMonth(year, month);
                    const lead = new Date(year, month, 1).getDay();
                    const cells = [];
                    for (let i = 0; i < lead; i++) cells.push(null);
                    for (let d = 1; d <= total; d++) cells.push(isoDate(year, month, d));
                    return cells.map((date, i) => {
                      if (!date) return <div key={`b-${i}`} />;
                      const marked = (unavailability[currentDoctorId] || []).includes(date);
                      const recurring = marked && isRecurringUnavailable(currentDoctorId, date);
                      const onDuty = effectiveSchedule[date] === currentDoctorId;
                      const type = dayType(date, holidaySet);
                      return (
                        <button key={date} onClick={() => toggleUnavailable(date)} className={`rounded-lg border p-2 min-h-[56px] text-left transition-colors ${recurring ? 'bg-indigo-50 border-indigo-300' : marked ? 'bg-red-50 border-red-300' : type === 'holiday' ? 'bg-rose-100 border-rose-200 hover:border-teal-300' : 'bg-white border-slate-200 hover:border-teal-300'} ${onDuty ? 'ring-2 ring-offset-1 ring-teal-500' : ''}`}>
                          <div className="font-mono text-[11px] text-slate-500">{Number(date.slice(-2))}</div>
                          {onDuty && <div className="text-[9px] text-teal-600 font-medium mt-0.5">อยู่เวร</div>}
                          {marked && <div className={`text-[10px] font-medium mt-0.5 ${recurring ? 'text-indigo-500' : 'text-red-500'}`}>{recurring ? 'ไม่สะดวกประจำ' : 'ไม่สะดวก'}</div>}
                        </button>
                      );
                    });
                  })()}
                </div>
                <div className="mt-3 flex justify-end">
                  <button onClick={() => clearUnavailableMonth(currentDoctorId)} className="text-xs text-slate-400 hover:text-red-500 hover:bg-red-50 px-2 py-1 rounded-lg transition-colors flex items-center gap-1">
                    <span className="text-sm leading-none">🗑</span> ล้างวันไม่สะดวกเดือนนี้
                  </button>
                </div>
                <div className={`mt-2 flex items-center justify-between gap-3 rounded-lg border px-3 py-2.5 ${unavailabilityConfirmed.includes(currentDoctorId) ? 'bg-emerald-50 border-emerald-200' : 'bg-white border-slate-200'}`}>
                  <p className="text-xs text-slate-600">
                    {unavailabilityConfirmed.includes(currentDoctorId)
                      ? <span className="text-emerald-700 font-medium flex items-center gap-1"><Check size={14} /> คุณยืนยันแล้วว่าแจ้งวันไม่สะดวกครบสำหรับเดือนนี้</span>
                      : 'แจ้งวันไม่สะดวกครบแล้วหรือยัง? กดยืนยันเพื่อให้แอดมินรู้ว่าพร้อมจัดตารางแล้ว'}
                  </p>
                  <button onClick={toggleUnavailabilityConfirmed} className={`shrink-0 text-xs font-medium px-3 py-1.5 rounded-lg transition-colors ${unavailabilityConfirmed.includes(currentDoctorId) ? 'text-slate-500 hover:bg-slate-100' : 'bg-teal-600 hover:bg-teal-700 text-white'}`}>
                    {unavailabilityConfirmed.includes(currentDoctorId) ? 'ยกเลิกการยืนยัน' : 'ยืนยันว่าแจ้งครบแล้ว'}
                  </button>
                </div>
              </>
            )}
          </div>
        )}

        {/* UNAVAILABLE TAB (admin) — lets admin view and edit any doctor's unavailability */}
        {activeTab === 'unavailable' && role === 'admin' && (
          <div className="max-w-3xl">
            <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
              <MonthNav year={year} month={month} onShift={shiftMonth} />
            </div>
            <p className="text-xs text-slate-400 mb-4 flex items-center gap-1"><Info size={12} /> คลิกชื่อแพทย์เพื่อดู/แก้ไขวันไม่สะดวกของคนนั้น</p>
            {activeDoctors.length === 0 ? (
              <EmptyState icon={UserCircle} title="ยังไม่มีแพทย์เดือนนี้" />
            ) : (
              <AdminUnavailablePanel
                year={year} month={month} doctors={activeDoctors} allDoctors={doctors}
                unavailability={unavailability} effectiveSchedule={effectiveSchedule}
                holidaySet={holidaySet} masterSchedule={masterSchedule} defaultDocId={currentDoctorId}
                unavailabilityConfirmed={unavailabilityConfirmed} onToggleConfirmed={toggleUnavailabilityConfirmed}
                isRecurringUnavailable={isRecurringUnavailable}
                onToggle={(docId, date) => {
                  setUnavailability(prev => {
                    const mine = prev[docId] || [];
                    const next = mine.includes(date) ? mine.filter(d => d !== date) : [...mine, date].sort();
                    const updated = { ...prev, [docId]: next };
                    setUnavailabilityConfirmed(prevC => {
                      const nextC = prevC.filter(id => id !== docId); // editing again un-confirms
                      saveMonth({ unavailability: updated, unavailabilityConfirmed: nextC });
                      return nextC;
                    });
                    return updated;
                  });
                }}
                onApplyRecurring={(docId, dow, occ) => applyRecurringUnavailable(docId, dow, occ)}
                onClearMonth={(docId) => clearUnavailableMonth(docId)}
              />
            )}
          </div>
        )}

        {activeTab === 'marketplace' && (
          <MarketplaceTab
            role={role} currentDoctorId={currentDoctorId} doctors={activeDoctors} getDoctor={getDoctor}
            marketplace={marketplace} myAssignedDates={myAssignedDates} holidaySet={holidaySet}
            currentScheduleGenerated={currentScheduleGenerated}
            unavailability={unavailability} effectiveSchedule={effectiveSchedule}
            year={year} month={month} onShiftMonth={shiftMonth} showToast={showToast}
            createPost={createPost} createBulkSell={createBulkSell} cancelPost={cancelPost} declinePost={declinePost} acceptPost={acceptPost}
          />
        )}

        {activeTab === 'notifications' && (
          <div className="max-w-xl">
            <p className="text-xs text-slate-400 mb-4 flex items-center gap-1"><Info size={12} /> โหมด Prototype: ข้อความ LINE ด้านล่างเป็นตัวอย่างจำลอง ยังไม่ได้ส่งจริง</p>
            {notifications.length === 0 ? (
              <EmptyState icon={Bell} title="ยังไม่มีการแจ้งเตือน" hint="เมื่อมีการแก้ไขตารางเวรหรือแลกเวร ระบบจะแจ้งเตือนที่นี่" />
            ) : (
              <div className="space-y-3">
                {notifications.map(n => (
                  <div key={n.id} className="border border-slate-200 rounded-xl p-3">
                    <p className="text-sm text-slate-700 font-medium mb-1">{n.message}</p>
                    <p className="text-[11px] text-slate-400 mb-2 font-mono">{new Date(n.ts).toLocaleString('th-TH')}</p>
                    <div className="flex items-start gap-2 bg-emerald-50 border-l-4 border-emerald-500 rounded-r-lg px-3 py-2">
                      <MessageCircle size={14} className="text-emerald-600 mt-0.5 shrink-0" />
                      <div><span className="text-[10px] font-semibold text-emerald-600 uppercase tracking-wide">ตัวอย่างข้อความ LINE</span><p className="text-xs text-emerald-800">{n.lineMessage}</p></div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>

      {toast && <div className="fixed bottom-4 left-1/2 -translate-x-1/2 bg-slate-800 text-white text-sm px-4 py-2 rounded-lg shadow-lg z-50">{toast}</div>}

      <ConfirmModal
        open={confirmState?.type === 'rearrange'}
        title="จัดเวร?"
        body={`ระบบจะจัดตารางเวรปัจจุบันทั้งเดือนจากตารางต้นแบบล่าสุด (จำนวนเวรวันธรรมดา/วันหยุดของแต่ละคนเท่ากับต้นแบบ และรับประกันว่าไม่มีใครอยู่เวรติดกัน) การแก้ไขเฉพาะจุดที่เคยทำไว้ในตารางเวรปัจจุบันเดือน ${THAI_MONTHS[month]} ${year + 543} จะถูกล้างไปด้วย — การขาย/แลกเวรที่สำเร็จแล้วจะไม่ถูกยกเลิก เพราะถูกบันทึกลงตารางต้นแบบไปแล้ว${confirmState?.pendingCount > 0 ? `\n\n⚠️ ยังมี ${confirmState.pendingCount} คนที่ยังไม่ยืนยันว่าแจ้งวันไม่สะดวกครบ ต้องการจัดเวรเลยหรือรอก่อน?` : ''}`}
        confirmLabel={confirmState?.pendingCount > 0 ? 'จัดเวรเลย' : 'จัดเวร'}
        danger={confirmState?.pendingCount > 0}
        onCancel={() => setConfirmState(null)}
        onConfirm={() => { setConfirmState(null); generateCurrentSchedule(); }}
      />

      <ConfirmModal
        open={confirmState?.type === 'clear-current'}
        title="ล้างตารางเวรปัจจุบัน?"
        body={`ตารางเวรปัจจุบันของเดือน ${THAI_MONTHS[month]} ${year + 543} จะกลับไปเป็นค่าเริ่มต้น (ไม่ปรากฏตาราง) จนกว่าจะกดจัดเวรใหม่ การแก้ไขเฉพาะจุดที่เคยทำไว้จะหายไปด้วย — การขาย/แลกเวรที่สำเร็จแล้วจะไม่ถูกยกเลิก เพราะถูกบันทึกลงตารางต้นแบบไปแล้ว`}
        confirmLabel="ล้างตาราง"
        danger
        onCancel={() => setConfirmState(null)}
        onConfirm={() => { setConfirmState(null); resetCurrentSchedule(); }}
      />

      <ConfirmModal
        open={confirmState?.type === 'clear-master'}
        title="ล้างตารางเวรต้นแบบ?"
        body={`ตารางเวรต้นแบบทั้งหมดของเดือน ${THAI_MONTHS[month]} ${year + 543} จะถูกล้าง (รวมถึงตารางเวรปัจจุบันที่คำนวณจากมันด้วย) รายชื่อแพทย์และวันไม่สะดวกที่แจ้งไว้จะไม่หายไป — ใช้เมื่อต้องการเริ่มจัดตารางต้นแบบใหม่ทั้งหมดสำหรับเดือนนี้`}
        confirmLabel="ล้างตาราง"
        danger
        onCancel={() => setConfirmState(null)}
        onConfirm={() => { setConfirmState(null); resetMasterSchedule(); }}
      />

      {showMasterGen && queueState && (
        <MasterScheduleGenerator
          year={year} month={month}
          doctors={doctors}
          activeDoctorIds={activeDoctorIds}
          holidays={holidays}
          queueState={queueState}
          onConfirm={handleMasterGenConfirm}
          onClose={() => setShowMasterGen(false)}
        />
      )}
    </div>
  );
}

/* ---------------------------------- marketplace ---------------------------------- */

/* ---------------------------------- swap calendar ---------------------------------- */

function SwapCalendar({ year, month, candidateDates, selected, holidaySet, getDoctor, onSelect }) {
  const pad2 = n => String(n).padStart(2,'0');
  const isoDate = (y,m,d) => `${y}-${pad2(m+1)}-${pad2(d)}`;
  const daysInMonth = (y,m) => new Date(y,m+1,0).getDate();
  const WEEKDAY_LABELS = ['อา','จ','อ','พ','พฤ','ศ','ส'];
  const candidateMap = new Map(candidateDates.map(({d,ownerId}) => [d, ownerId]));
  const total = daysInMonth(year, month);
  const lead = new Date(year, month, 1).getDay();
  const cells = [];
  for (let i = 0; i < lead; i++) cells.push(null);
  for (let d = 1; d <= total; d++) cells.push(isoDate(year, month, d));

  return (
    <div className="border border-slate-200 rounded-xl p-3 bg-white">
      <div className="grid grid-cols-7 gap-1 mb-1">
        {WEEKDAY_LABELS.map((w,i) => <div key={w} className={`text-center text-[10px] font-semibold ${i===0||i===6?'text-rose-500':'text-slate-400'}`}>{w}</div>)}
      </div>
      <div className="grid grid-cols-7 gap-1">
        {cells.map((date, i) => {
          if (!date) return <div key={`b-${i}`} />;
          const ownerId = candidateMap.get(date);
          const isCandidate = !!ownerId;
          const isSelected = date === selected;
          const isHoliday = dayType(date, holidaySet) === 'holiday';
          const dayNum = Number(date.slice(-2));
          const ownerName = ownerId ? getDoctor(ownerId)?.name || '' : '';
          return (
            <button key={date}
              disabled={!isCandidate}
              onClick={() => isCandidate && onSelect(date)}
              title={isCandidate ? `${ownerName}` : ''}
              className={`rounded-lg border p-1 min-h-[48px] flex flex-col items-center justify-center transition-colors text-center
                ${isSelected ? 'bg-indigo-600 border-indigo-600 text-white' :
                  isCandidate ? (isHoliday ? 'bg-rose-50 border-rose-300 hover:border-indigo-400 cursor-pointer' : 'bg-white border-indigo-200 hover:border-indigo-400 cursor-pointer') :
                  'bg-slate-50 border-slate-100 opacity-30 cursor-not-allowed'}`}>
              <span className="font-mono text-[11px]">{dayNum}</span>
              {isCandidate && <span className={`text-[9px] truncate w-full text-center leading-tight ${isSelected ? 'text-white' : 'text-slate-500'}`}>{ownerName}</span>}
            </button>
          );
        })}
      </div>
      {selected && <p className="text-xs text-indigo-600 mt-2 font-medium">เลือก: {formatDisplayDate(selected)} ({getDoctor(candidateMap.get(selected))?.name})</p>}
    </div>
  );
}

function MarketplaceTab({ role, currentDoctorId, doctors, getDoctor, marketplace, myAssignedDates, holidaySet, currentScheduleGenerated, unavailability, effectiveSchedule, year, month, onShiftMonth, showToast, createPost, createBulkSell, cancelPost, declinePost, acceptPost }) {
  // The month nav here drives the SAME global year/month as every other tab
  // (not a separate local picker) — so "my shifts this month" and swap
  // candidates always come from whichever month is actually loaded, with no
  // risk of the two drifting out of sync. Posts/history are shown across all
  // time rather than filtered to a month.
  const [sellDate, setSellDate] = useState('');
  const [sellTarget, setSellTarget] = useState('');
  const [swapDate, setSwapDate] = useState('');
  const [swapTarget, setSwapTarget] = useState('');
  const [swapRequestedDate, setSwapRequestedDate] = useState('');
  const [swapSelfAdjacentConfirmed, setSwapSelfAdjacentConfirmed] = useState(false);
  const [acceptTarget, setAcceptTarget] = useState(null);
  const [buyAllConfirm, setBuyAllConfirm] = useState(false);

  const myDates = currentDoctorId ? myAssignedDates(currentDoctorId) : [];
  const otherDoctors = doctors.filter(d => d.id !== currentDoctorId);
  const sellPosts = marketplace.filter(p => p.type === 'sell' && p.status === 'open');
  const swapPosts = marketplace.filter(p => p.type === 'swap' && p.status === 'open');
  const history = marketplace.filter(p => p.status !== 'open').slice(0, 8);

  // "ซื้อทุกเวร" only targets posts for the CURRENTLY LOADED month — acceptPost
  // refuses posts from any other month (to avoid writing into the wrong
  // month's data), so bulk-buying across months would just fail silently.
  const monthPrefix = `${year}-${String(month + 1).padStart(2, '0')}`;
  const buyableSellPosts = sellPosts.filter(p =>
    p.date.startsWith(monthPrefix) && p.posterId !== currentDoctorId && (!p.targetDoctorId || p.targetDoctorId === currentDoctorId)
  );
  const buyAllSellPosts = () => {
    buyableSellPosts.forEach(p => acceptPost(p));
  };

  const submitSell = () => {
    if (!sellDate) return;
    const alreadyPosted = new Set(marketplace.filter(p => p.type === 'sell' && p.status === 'open' && p.posterId === currentDoctorId).map(p => p.date));
    if (sellDate === '__ALL__') {
      createBulkSell(myDates.filter(d => !alreadyPosted.has(d)), sellTarget || null);
    } else if (sellDate === '__WEEKDAY__') {
      createBulkSell(myDates.filter(d => !alreadyPosted.has(d) && dayType(d, holidaySet) === 'weekday'), sellTarget || null);
    } else if (sellDate === '__HOLIDAY__') {
      createBulkSell(myDates.filter(d => !alreadyPosted.has(d) && dayType(d, holidaySet) === 'holiday'), sellTarget || null);
    } else {
      createPost(sellDate, 'sell', sellTarget || null);
    }
    setSellDate(''); setSellTarget('');
  };

  // All dates in the current schedule owned by someone other than me,
  // same day-type as my chosen date, that won't make the OTHER doctor
  // adjacent if they take my date. Each entry includes the owner id.
  const candidateSwapDates = swapDate && currentScheduleGenerated
    ? (() => {
        const type = dayType(swapDate, holidaySet);
        return Object.entries(effectiveSchedule)
          .filter(([d, ownerId]) =>
            ownerId && ownerId !== currentDoctorId &&
            dayType(d, holidaySet) === type &&
            // target doctor won't be adjacent taking over our date
            !hasAdjacentAssignment({ ...effectiveSchedule, [swapDate]: ownerId, [d]: currentDoctorId }, swapDate, ownerId)
          )
          .map(([d, ownerId]) => ({ d, ownerId }))
          .sort((a, b) => a.d.localeCompare(b.d));
      })()
    : [];

  // Once requester picks a target date, derive the targetDoctorId from it
  const swapDerivedTarget = swapRequestedDate
    ? (effectiveSchedule[swapRequestedDate] || null)
    : null;

  // Check if the requester would end up adjacent after taking the target's date
  const swapSelfAdjacent = swapDate && swapRequestedDate && currentDoctorId && swapDerivedTarget
    ? hasAdjacentAssignment({ ...effectiveSchedule, [swapDate]: swapDerivedTarget, [swapRequestedDate]: currentDoctorId }, swapRequestedDate, currentDoctorId)
    : false;

  const submitSwap = () => {
    if (!swapDate || !swapRequestedDate || !swapDerivedTarget) return;
    if (swapSelfAdjacent && !swapSelfAdjacentConfirmed) { setSwapSelfAdjacentConfirmed(true); return; }
    createPost(swapDate, 'swap', swapDerivedTarget, swapRequestedDate);
    setSwapDate(''); setSwapTarget(''); setSwapRequestedDate(''); setSwapSelfAdjacentConfirmed(false);
  };

  const canAct = (p) => !p.targetDoctorId || p.targetDoctorId === currentDoctorId;

  // Selections reference specific dates, which only make sense for the month
  // they came from — clear them whenever the loaded month changes so a stale
  // date from the previous month can't linger in the form.
  useEffect(() => {
    setSellDate(''); setSellTarget('');
    setSwapDate(''); setSwapTarget(''); setSwapRequestedDate(''); setSwapSelfAdjacentConfirmed(false);
  }, [year, month]);

  return (
    <div className="max-w-2xl space-y-8">
      <MonthNav year={year} month={month} onShift={onShiftMonth} />

      {!currentScheduleGenerated && (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2 flex items-start gap-2">
          <Info size={14} className="mt-0.5 shrink-0" />
          ยังไม่ได้จัดตารางเวรปัจจุบันของเดือนนี้ วันที่แสดงด้านล่างจึงอิงจากตารางเวรต้นแบบไปก่อน หลังแอดมินกดจัดเวรแล้ว วันที่จะเปลี่ยนเป็นวันที่อยู่เวรจริง (ฟังก์ชันแลกเวรจะใช้งานได้หลังจัดเวรแล้วเท่านั้น)
        </div>
      )}

      <div className="border border-teal-200 rounded-2xl p-4 bg-teal-50/30">
        <div className="flex items-center justify-between gap-2 mb-3 flex-wrap">
          <div className="flex items-center gap-2"><Tag size={16} className="text-teal-600" /><p className="font-display font-semibold text-slate-800">ขายเวร</p></div>
          {buyableSellPosts.length > 0 && (
            <button onClick={() => setBuyAllConfirm(true)} className="flex items-center gap-1.5 bg-teal-600 hover:bg-teal-700 text-white text-xs font-medium px-3 py-1.5 rounded-lg transition-colors">
              <Check size={13} /> ซื้อทุกเวร ({buyableSellPosts.length})
            </button>
          )}
        </div>

        {(role === 'doctor' || role === 'admin') && (
          myDates.length === 0 ? <p className="text-sm text-slate-400 mb-2">คุณยังไม่มีเวรที่จัดไว้ในเดือนนี้</p> : (
            <div className="flex flex-wrap items-center gap-2 mb-4">
              <select value={sellDate} onChange={(e) => setSellDate(e.target.value)} className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm font-mono">
                <option value="">เลือกวันที่ของฉัน</option>
                <option value="__ALL__">🔁 ขายทุกเวรของฉันในเดือนที่เลือก</option>
                <option value="__WEEKDAY__">📅 ขายเฉพาะเวรวันธรรมดาของฉัน</option>
                <option value="__HOLIDAY__">🎌 ขายเฉพาะเวรวันหยุดของฉัน</option>
                {myDates.map(d => <option key={d} value={d}>{formatDisplayDate(d)} ({dayTypeLabel(d, holidaySet)})</option>)}
              </select>
              <select value={sellTarget} onChange={(e) => setSellTarget(e.target.value)} className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm">
                <option value="">ไม่ระบุ (เปิดให้ทุกคนรับ)</option>
                {otherDoctors.map(d => <option key={d.id} value={d.id}>ขายให้ {d.name} โดยเฉพาะ</option>)}
              </select>
              <button disabled={!sellDate} onClick={submitSell} className="flex items-center gap-1 bg-teal-600 hover:bg-teal-700 disabled:bg-slate-200 disabled:cursor-not-allowed text-white text-sm font-medium px-3 py-1.5 rounded-lg transition-colors"><Plus size={14} /> ลงประกาศขายเวร</button>
            </div>
          )
        )}

        {sellPosts.length === 0 ? <EmptyState icon={Tag} title="ยังไม่มีประกาศขายเวร" /> : (() => {
          // Privacy: doctor view shows only own posts + open-to-all posts.
          // Admin sees everything.
          const visiblePosts = role === 'admin' ? sellPosts : sellPosts.filter(p =>
            p.posterId === currentDoctorId || !p.targetDoctorId || p.targetDoctorId === currentDoctorId
          );
          if (visiblePosts.length === 0) return <EmptyState icon={Tag} title="ยังไม่มีประกาศขายเวรสำหรับคุณ" hint="ประกาศขายเวรที่ระบุแพทย์คนอื่นโดยเฉพาะจะมองไม่เห็น" />;
          return (
            <div className="space-y-2">
              {visiblePosts.map(p => {
                const poster = getDoctor(p.posterId);
                const target = p.targetDoctorId ? getDoctor(p.targetDoctorId) : null;
                const isMine = p.posterId === currentDoctorId;
                const eligible = canAct(p) && !isMine;
                return (
                  <div key={p.id} className="border border-slate-200 bg-white rounded-xl p-3 flex items-center justify-between flex-wrap gap-2">
                    <div>
                      <p className="text-sm font-medium text-slate-700">{poster?.name} · <span className="font-mono">{formatDisplayDate(p.date)}</span> ({dayTypeLabel(p.date, holidaySet)})</p>
                      <p className="text-xs text-slate-400">{target ? `ขายเฉพาะให้ ${target.name}` : 'เปิดให้ทุกคนรับ'}</p>
                    </div>
                    <div className="flex items-center gap-2">
                      {role === 'admin' && !isMine && (
                        <span className="text-[10px] text-slate-400 border border-slate-200 rounded px-1.5 py-0.5">
                          {target ? `รอ ${target.name}` : 'เปิดให้รับ'}
                        </span>
                      )}
                      {isMine ? (
                        <button onClick={() => cancelPost(p.id)} className="text-xs font-medium text-slate-500 hover:text-red-600 px-2 py-1">ยกเลิกประกาศ</button>
                      ) : eligible ? (
                        <button onClick={() => setAcceptTarget(p)} className="flex items-center gap-1 bg-teal-600 hover:bg-teal-700 text-white text-xs font-medium px-3 py-1.5 rounded-lg"><Check size={13} /> รับเวรนี้</button>
                      ) : (!role || role === 'doctor') && target ? (
                        <span className="text-xs text-slate-400">รอ {target.name} ตอบรับ</span>
                      ) : null}
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })()}
      </div>

      <div className="border border-indigo-200 rounded-2xl p-4 bg-indigo-50/30">
        <div className="flex items-center gap-2 mb-3"><ArrowRightLeft size={16} className="text-indigo-600" /><p className="font-display font-semibold text-slate-800">แลกเวร</p></div>
        <p className="text-xs text-slate-400 mb-3">แลกได้เฉพาะวันธรรมดากับวันธรรมดา หรือวันหยุดกับวันหยุด และใช้ได้หลังแอดมินจัดตารางเวรปัจจุบันแล้วเท่านั้น</p>

        {(role === 'doctor' || role === 'admin') && (
          !currentScheduleGenerated ? (
            <p className="text-sm text-slate-400 mb-2">รอแอดมินจัดตารางเวรปัจจุบันก่อน</p>
          ) : myDates.length === 0 ? (
            <p className="text-sm text-slate-400 mb-2">คุณยังไม่มีเวรในตารางปัจจุบัน</p>
          ) : (
            <div className="flex flex-wrap items-center gap-2 mb-4">
              <select value={swapDate} onChange={(e) => { setSwapDate(e.target.value); setSwapRequestedDate(''); setSwapSelfAdjacentConfirmed(false); }} className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm font-mono">
                <option value="">1. เลือกวันที่ของฉันที่จะแลกออก</option>
                {myDates.filter(d => !(unavailability[currentDoctorId] || []).includes(d)).map(d => <option key={d} value={d}>{formatDisplayDate(d)} ({dayTypeLabel(d, holidaySet)})</option>)}
              </select>

              {swapDate && (
                candidateSwapDates.length === 0 ? (
                  <span className="text-xs text-red-500">ไม่มีวันประเภทเดียวกันในตารางปัจจุบันที่แลกได้โดยไม่ทำให้ใครติดกัน</span>
                ) : (
                  <div className="w-full mt-1">
                    <p className="text-xs text-slate-500 mb-2">2. เลือกวันที่ต้องการแลกมา (เฉพาะ{dayType(swapDate, holidaySet) === 'holiday' ? 'วันหยุด' : 'วันธรรมดา'})</p>
                    <SwapCalendar
                      year={year} month={month}
                      candidateDates={candidateSwapDates}
                      selected={swapRequestedDate}
                      holidaySet={holidaySet}
                      getDoctor={getDoctor}
                      onSelect={(d) => { setSwapRequestedDate(d); setSwapSelfAdjacentConfirmed(false); }}
                    />
                  </div>
                )
              )}

              {swapSelfAdjacent && !swapSelfAdjacentConfirmed && (
                <div className="w-full bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2">
                  ⚠️ ถ้าแลกเวรนี้ <strong>คุณจะอยู่เวรติดกัน 2 วัน</strong> คุณยืนยันว่ารับได้ใช่ไหม? กดปุ่มด้านล่างอีกครั้งเพื่อยืนยัน
                </div>
              )}

              <button
                disabled={!swapDate || !swapRequestedDate}
                onClick={submitSwap}
                className={`flex items-center gap-1 text-white text-sm font-medium px-3 py-1.5 rounded-lg transition-colors disabled:bg-slate-200 disabled:cursor-not-allowed ${swapSelfAdjacent && !swapSelfAdjacentConfirmed ? 'bg-amber-500 hover:bg-amber-600' : 'bg-indigo-600 hover:bg-indigo-700'}`}
              >
                <Plus size={14} /> {swapSelfAdjacent && !swapSelfAdjacentConfirmed ? 'ยืนยันอยู่เวรติดกันได้' : 'ลงคำขอแลกเวร'}
              </button>
            </div>
          )
        )}

        {swapPosts.length === 0 ? <EmptyState icon={ArrowRightLeft} title="ยังไม่มีคำขอแลกเวร" /> : (
          <div className="space-y-2">
            {swapPosts.map(p => {
              const poster = getDoctor(p.posterId);
              const target = getDoctor(p.targetDoctorId);
              const isMine = p.posterId === currentDoctorId;
              const isTarget = p.targetDoctorId === currentDoctorId;
              return (
                <div key={p.id} className="border border-slate-200 bg-white rounded-xl p-3 flex items-center justify-between flex-wrap gap-2">
                  <div>
                    <p className="text-sm font-medium text-slate-700">{poster?.name} ขอแลกวันที่ <span className="font-mono">{formatDisplayDate(p.date)}</span> ({dayTypeLabel(p.date, holidaySet)}) กับวันที่ <span className="font-mono">{formatDisplayDate(p.requestedDate)}</span> ({dayTypeLabel(p.requestedDate, holidaySet)}) ของ {target?.name}</p>
                  </div>
                  {isMine ? (
                    <button onClick={() => cancelPost(p.id)} className="text-xs font-medium text-slate-500 hover:text-red-600 px-2 py-1">ยกเลิกคำขอ</button>
                  ) : isTarget ? (
                    <div className="flex items-center gap-2">
                      <button onClick={() => declinePost(p)} className="text-xs font-medium text-slate-500 hover:text-red-600 px-2 py-1">ปฏิเสธ</button>
                      <button onClick={() => setAcceptTarget(p)} className="flex items-center gap-1 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-medium px-3 py-1.5 rounded-lg"><Check size={13} /> ยืนยัน</button>
                    </div>
                  ) : (
                    <span className="text-xs text-slate-400">รอ {target?.name} ตอบรับ</span>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>

      {history.length > 0 && (
        <div>
          <p className="font-display font-semibold text-slate-800 mb-2 text-sm">ประวัติล่าสุด</p>
          <div className="space-y-1.5">
            {history.map(p => (<p key={p.id} className="text-xs text-slate-400">{formatDisplayDate(p.date)} · {p.type === 'sell' ? 'ขายเวร' : 'แลกเวร'} · {getDoctor(p.posterId)?.name} → {p.status === 'completed' ? (getDoctor(p.takerId)?.name || '-') : 'ยกเลิก/ปฏิเสธ'}</p>))}
          </div>
        </div>
      )}

      <ConfirmModal
        open={!!acceptTarget}
        title={acceptTarget?.type === 'swap' ? 'ยืนยันแลกเวรนี้?' : 'ยืนยันรับเวรนี้?'}
        body={acceptTarget ? (
          acceptTarget.type === 'sell'
            ? `คุณจะรับเวรวันที่ ${formatDisplayDate(acceptTarget.date)} แทน ${getDoctor(acceptTarget.posterId)?.name} — ตารางเวรต้นแบบและโควตาจะอัปเดตทันที`
            : `คุณจะมอบเวรวันที่ ${formatDisplayDate(acceptTarget.requestedDate)} ของคุณให้ ${getDoctor(acceptTarget.posterId)?.name} และรับเวรวันที่ ${formatDisplayDate(acceptTarget.date)} มาแทน — ตารางเวรปัจจุบันจะอัปเดตทันที (ตารางต้นแบบและโควตาไม่เปลี่ยน)`
        ) : ''}
        confirmLabel="ยืนยัน"
        onCancel={() => setAcceptTarget(null)}
        onConfirm={() => { acceptPost(acceptTarget); setAcceptTarget(null); }}
      />

      <ConfirmModal
        open={buyAllConfirm}
        title="ซื้อทุกเวรที่เปิดอยู่?"
        body={`คุณจะรับเวรทั้งหมด ${buyableSellPosts.length} วันที่เปิดขายอยู่ในเดือนนี้ (${buyableSellPosts.map(p => formatDisplayDate(p.date)).join(', ')}) — ตารางเวรต้นแบบและโควตาจะอัปเดตทันทีสำหรับทุกวัน`}
        confirmLabel="ซื้อทั้งหมด"
        onCancel={() => setBuyAllConfirm(false)}
        onConfirm={() => { buyAllSellPosts(); setBuyAllConfirm(false); showToast(`ซื้อเวรสำเร็จ ${buyableSellPosts.length} รายการ`); }}
      />
    </div>
  );
}

/* ---------------------------------- admin unavailable panel ---------------------------------- */

/* ---------------------------------- recurring unavailability panel ---------------------------------- */

const DOW_LABELS = ['อา','จ','อ','พ','พฤ','ศ','ส'];
const DOW_FULL   = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];

function RecurringUnavailablePanel({ year, month, onApply }) {
  const [open, setOpen] = useState(false);
  const [dow, setDow] = useState(5); // default ศุกร์
  const [occ, setOcc] = useState([]); // [] = ทุกครั้ง

  // Preview which dates will be marked
  const preview = (() => {
    const total = new Date(year, month + 1, 0).getDate();
    const pad2 = n => String(n).padStart(2,'0');
    const iso = d => `${year}-${pad2(month+1)}-${pad2(d)}`;
    const dates = [];
    let nth = 0;
    for (let d = 1; d <= total; d++) {
      if (new Date(year, month, d).getDay() === dow) {
        nth++;
        if (occ.length === 0 || occ.includes(nth)) dates.push(iso(d));
      }
    }
    return dates;
  })();

  const toggleOcc = (n) => setOcc(prev => prev.includes(n) ? prev.filter(x => x !== n) : [...prev, n].sort());

  if (!open) return (
    <div className="mb-4">
      <button onClick={() => setOpen(true)} className="flex items-center gap-1.5 text-xs font-medium text-indigo-600 hover:bg-indigo-50 border border-indigo-200 px-3 py-1.5 rounded-lg transition-colors">
        <span className="text-base leading-none">🔁</span> เพิ่มวันไม่สะดวกประจำ (fix schedule)
      </button>
    </div>
  );

  return (
    <div className="mb-4 border border-indigo-200 rounded-xl bg-indigo-50/40 p-3">
      <div className="flex items-center justify-between mb-3">
        <p className="text-xs font-medium text-indigo-700">เพิ่มวันไม่สะดวกประจำ</p>
        <button onClick={() => setOpen(false)} className="text-slate-400 hover:text-slate-600"><X size={14} /></button>
      </div>
      <div className="flex flex-wrap items-center gap-2 mb-3">
        <span className="text-xs text-slate-600">วัน:</span>
        <div className="flex gap-1">
          {DOW_LABELS.map((label, i) => (
            <button key={i} onClick={() => { setDow(i); setOcc([]); }}
              className={`w-8 h-8 rounded-full text-xs font-medium transition-colors ${dow === i ? 'bg-indigo-600 text-white' : 'bg-white border border-slate-200 text-slate-600 hover:border-indigo-300'}`}>
              {label}
            </button>
          ))}
        </div>
      </div>
      <div className="flex flex-wrap items-center gap-2 mb-3">
        <span className="text-xs text-slate-600">ครั้งที่:</span>
        {[1,2,3,4,5].map(n => (
          <button key={n} onClick={() => toggleOcc(n)}
            className={`w-8 h-8 rounded-full text-xs font-medium transition-colors ${occ.includes(n) ? 'bg-indigo-600 text-white' : 'bg-white border border-slate-200 text-slate-600 hover:border-indigo-300'}`}>
            {n}
          </button>
        ))}
        <button onClick={() => setOcc([])} className={`px-2 h-8 rounded-full text-xs font-medium transition-colors ${occ.length === 0 ? 'bg-indigo-600 text-white' : 'bg-white border border-slate-200 text-slate-600 hover:border-indigo-300'}`}>
          ทุกครั้ง
        </button>
      </div>
      <div className="flex flex-wrap gap-1 mb-3 min-h-[22px]">
        {preview.length > 0
          ? <><span className="text-[10px] text-slate-500 mr-1">เดือนนี้:</span>{preview.map(d => <span key={d} className="text-[10px] bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full font-mono">{d.slice(-2)} {['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][month]}</span>)}</>
          : <span className="text-[10px] text-amber-600 italic">เดือนนี้ไม่มีวันดังกล่าว — rule จะ apply อัตโนมัติในเดือนที่มี</span>
        }
      </div>
      <button onClick={() => { onApply(dow, occ); setOpen(false); }}
        className="flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-medium px-3 py-1.5 rounded-lg transition-colors">
        <Check size={12} /> {preview.length > 0 ? `บันทึก rule + เพิ่ม ${preview.length} วันเดือนนี้` : 'บันทึก rule (apply เดือนที่มีวันดังกล่าว)'}
      </button>
    </div>
  );
}

function AdminUnavailablePanel({ year, month, doctors, allDoctors, unavailability, effectiveSchedule, holidaySet, masterSchedule, defaultDocId, unavailabilityConfirmed, onToggleConfirmed, isRecurringUnavailable, onToggle, onApplyRecurring, onClearMonth }) {
  // Default to the logged-in admin's own entry (falling back to the first
  // doctor in the roster if they're not in this month's active list) so
  // admin doesn't have to re-select themselves every time.
  const [selectedDocId, setSelectedDocId] = useState(
    doctors.find(d => d.id === defaultDocId)?.id ?? doctors[0]?.id ?? null
  );
  const WEEKDAY_LABELS = ['อา','จ','อ','พ','พฤ','ศ','ส'];
  const pad2 = n => String(n).padStart(2,'0');
  const isoDate = (y, m, d) => `${y}-${pad2(m + 1)}-${pad2(d)}`;
  const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
  const total = daysInMonth(year, month);
  const lead = new Date(year, month, 1).getDay();
  const cells = [];
  for (let i = 0; i < lead; i++) cells.push(null);
  for (let d = 1; d <= total; d++) cells.push(isoDate(year, month, d));

  const getDocColor = (id) => {
    const idx = allDoctors.findIndex(d => d.id === id);
    return getDoctorColor(idx);
  };

  return (
    <div>
      <div className="flex flex-wrap gap-2 mb-4">
        {doctors.map(d => {
          const color = getDocColor(d.id);
          const active = selectedDocId === d.id;
          return (
            <button key={d.id} onClick={() => setSelectedDocId(d.id)}
              className={`flex items-center gap-1.5 text-sm px-3 py-1.5 rounded-lg border transition-colors ${active ? `${color.soft} ${color.text} border-transparent ring-2 ring-offset-1 ring-slate-300` : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300'}`}>
              <span className={`w-2 h-2 rounded-full ${color.bg}`} />{d.name}
            </button>
          );
        })}
      </div>

      {selectedDocId && (() => {
        const doc = allDoctors.find(d => d.id === selectedDocId);
        const markedDates = unavailability[selectedDocId] || [];
        return (
          <>
            <p className="text-xs text-slate-400 mb-3 flex items-center gap-1"><Info size={12} /> คลิกวันที่เพื่อเพิ่ม/ลบวันไม่สะดวกของ {doc?.name} · <span className="inline-block w-2.5 h-2.5 rounded-sm bg-violet-50 border border-violet-300" /> ไม่สะดวกประจำ (จาก rule) · <span className="inline-block w-2.5 h-2.5 rounded-sm bg-red-50 border border-red-300" /> จิ้มเลือกเอง</p>
            <RecurringUnavailablePanel
              year={year} month={month}
              onApply={(dow, occ) => onApplyRecurring(selectedDocId, dow, occ)}
            />
            <div className="grid grid-cols-7 gap-1 mb-1">
              {WEEKDAY_LABELS.map((w, i) => <div key={w} className={`text-center text-xs font-semibold py-1 ${i === 0 || i === 6 ? 'text-rose-500' : 'text-slate-400'}`}>{w}</div>)}
            </div>
            <div className="grid grid-cols-7 gap-1">
              {cells.map((date, i) => {
                if (!date) return <div key={`b-${i}`} />;
                const marked = markedDates.includes(date);
                const recurring = marked && isRecurringUnavailable(selectedDocId, date);
                const onDuty = effectiveSchedule[date] === selectedDocId;
                const inMaster = masterSchedule[date] === selectedDocId;
                const type = dayType(date, holidaySet);
                return (
                  <button key={date} onClick={() => onToggle(selectedDocId, date)}
                    className={`rounded-lg border p-2 min-h-[56px] text-left transition-colors
                      ${recurring ? 'bg-indigo-50 border-indigo-300' : marked ? 'bg-red-50 border-red-300' : type === 'holiday' ? 'bg-rose-100 border-rose-200 hover:border-teal-300' : 'bg-white border-slate-200 hover:border-teal-300'}`}>
                    <div className="font-mono text-[11px] text-slate-500">{Number(date.slice(-2))}</div>
                    {inMaster && <div className="text-[9px] text-sky-600 font-medium">ต้นแบบ</div>}
                    {onDuty && <div className="text-[9px] text-teal-600 font-medium">อยู่เวร</div>}
                    {marked && <div className={`text-[10px] font-medium ${recurring ? 'text-indigo-500' : 'text-red-500'}`}>{recurring ? 'ไม่สะดวกประจำ' : 'ไม่สะดวก'}</div>}
                  </button>
                );
              })}
            </div>
            <div className="mt-3 flex flex-wrap gap-1.5 items-center">
              {markedDates.length === 0
                ? <p className="text-xs text-slate-400">{doc?.name} ไม่มีวันที่แจ้งไม่สะดวกเดือนนี้</p>
                : markedDates.map(d => (
                  <span key={d} className="text-[11px] font-mono bg-red-50 text-red-600 px-2 py-0.5 rounded-full">{formatDisplayDate(d)}</span>
                ))
              }
              {markedDates.length > 0 && (
                <button onClick={() => onClearMonth(selectedDocId)} className="ml-auto text-xs text-slate-400 hover:text-red-500 hover:bg-red-50 px-2 py-1 rounded-lg transition-colors">
                  🗑 ล้างวันไม่สะดวกเดือนนี้
                </button>
              )}
            </div>
            {(() => {
              const confirmed = unavailabilityConfirmed.includes(selectedDocId);
              return (
                <div className={`mt-3 flex items-center justify-between gap-3 rounded-lg border px-3 py-2.5 ${confirmed ? 'bg-emerald-50 border-emerald-200' : 'bg-white border-slate-200'}`}>
                  <p className="text-xs text-slate-600">
                    {confirmed
                      ? <span className="text-emerald-700 font-medium flex items-center gap-1"><Check size={14} /> {doc?.name} ยืนยันแล้วว่าแจ้งวันไม่สะดวกครบสำหรับเดือนนี้</span>
                      : `${doc?.name} แจ้งวันไม่สะดวกครบแล้วหรือยัง? กดยืนยันเพื่อให้นับในสถานะความพร้อมจัดเวร`}
                  </p>
                  <button onClick={() => onToggleConfirmed(selectedDocId)} className={`shrink-0 text-xs font-medium px-3 py-1.5 rounded-lg transition-colors ${confirmed ? 'text-slate-500 hover:bg-slate-100' : 'bg-teal-600 hover:bg-teal-700 text-white'}`}>
                    {confirmed ? 'ยกเลิกการยืนยัน' : 'ยืนยันว่าแจ้งครบแล้ว'}
                  </button>
                </div>
              );
            })()}
          </>
        );
      })()}
    </div>
  );
}
