import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Calendar as CalendarIcon, CalendarCheck, Upload, Users, Repeat, Bell, Plus, X, Check,
  ChevronLeft, ChevronRight, Settings, UserCircle, Trash2,
  MessageCircle, Info, ArrowRightLeft, Tag, Shuffle, RotateCcw, LayoutDashboard, Download, FileText
} from 'lucide-react';
import {
  getConfig, setConfig, getMonthData, setMonthData,
  getMarketplace, setMarketplace,
  getNotifications, addNotification as dbAddNotification,
  getDoctors, addDoctor, updateDoctor, deleteDoctor,
  getQueueState, setQueueState,
} from './storage';
import LoginScreen from './LoginScreen';
import MasterScheduleGenerator, { detectGroups, ltFor, resolveQueue, DEFAULT_WDQ_NAMES, DEFAULT_H12Q_NAMES, DEFAULT_H3Q_NAMES } from './MasterScheduleGenerator';

/* ---------------------------------- constants ---------------------------------- */

const THAI_MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
const THAI_MONTHS_SHORT = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
const THAI_WEEKDAYS_FULL = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
const WEEKDAY_LABELS = ['อา','จ','อ','พ','พฤ','ศ','ส'];
const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

// Full title + name for the official duty-roster DOCX export — the
// doctors table only ever stores the short first name everyone in this
// app is addressed by day to day (see every other screen), never a title
// or surname, so this is the one place that needs the formal version.
// Keyed by that same short name; supplied directly by the department
// admin, not derivable from anything already in the database.
const DOCTOR_FULL_NAME = {
  'กนกอร': 'พญ.กนกอร จ่างจรูญโรจน์',
  'ขนิษฐา': 'พญ.ขนิษฐา เพิ่มทวี',
  'ชุติมา': 'พญ.ชุติมา พยุงธนทรัพย์',
  'ณัชพล': 'นพ.ณัชพล ทวีสกุลชัย',
  'ณัฐธิดา': 'พญ.ณัฐธิดา พฤกษ์งามพันธ์',
  'ณัฐพล': 'นพ.ณัฐพล รุ่งโรจนานนท์',
  'ธนวรรณ': 'พญ.ธนวรรณ ตีรณธาดา',
  'ธัญลักษณ์': 'พญ.ธัญลักษณ์ อัศวยนต์ชัย',
  'พสิษฐา': 'พญ.พสิษฐา เติมวรสิน',
  'วัทนี': 'พญ.วัทนี ทวีสิทธิ์',
  'สมิตา': 'พญ.สมิตา โลหะวิจารณ์',
  'อารีรัตน์': 'พญ.อารีรัตน์ ชัยเรืองยศ',
};
const DUTY_ROSTER_HEAD_NAME = 'นางสาวอารีรัตน์ ชัยเรืองยศ';
const DUTY_ROSTER_HEAD_TITLE = 'หัวหน้ากลุ่มงานกุมารเวชกรรม';

// Chosen by actually computing perceptual (CIE Lab Delta-E) distance between
// every pair of Tailwind's -600 hues, then greedily selecting the 12-color
// subset (one shade per hue family, no picking two lightnesses of the same
// hue) that maximizes the worst-case pairwise distance. The previous
// hand-picked set still had a closest pair (indigo/violet) at Delta-E 14.4 —
// visibly too close, which is exactly what got reported (ณัฐธิดา/ขนิษฐา).
// This set's closest pair is Delta-E 27.2, roughly double.
const DOCTOR_PALETTE = [
  { bg: 'bg-orange-600', soft: 'bg-orange-50', text: 'text-orange-700', ring: 'ring-orange-500' },
  { bg: 'bg-yellow-600', soft: 'bg-yellow-50', text: 'text-yellow-700', ring: 'ring-yellow-500' },
  { bg: 'bg-lime-600', soft: 'bg-lime-50', text: 'text-lime-700', ring: 'ring-lime-500' },
  { bg: 'bg-green-600', soft: 'bg-green-50', text: 'text-green-700', ring: 'ring-green-500' },
  { bg: 'bg-teal-600', soft: 'bg-teal-50', text: 'text-teal-700', ring: 'ring-teal-500' },
  { bg: 'bg-sky-600', soft: 'bg-sky-50', text: 'text-sky-700', ring: 'ring-sky-500' },
  { bg: 'bg-blue-600', soft: 'bg-blue-50', text: 'text-blue-700', ring: 'ring-blue-500' },
  { bg: 'bg-violet-600', soft: 'bg-violet-50', text: 'text-violet-700', ring: 'ring-violet-500' },
  { bg: 'bg-fuchsia-600', soft: 'bg-fuchsia-50', text: 'text-fuchsia-700', ring: 'ring-fuchsia-500' },
  { bg: 'bg-pink-600', soft: 'bg-pink-50', text: 'text-pink-700', ring: 'ring-pink-500' },
  { bg: 'bg-rose-600', soft: 'bg-rose-50', text: 'text-rose-700', ring: 'ring-rose-500' },
  { bg: 'bg-slate-700', soft: 'bg-slate-100', text: 'text-slate-700', ring: 'ring-slate-500' },
];
const getDoctorColor = (idx) => DOCTOR_PALETTE[idx % DOCTOR_PALETTE.length];

/* ---------------------------------- utils ---------------------------------- */

const pad2 = (n) => String(n).padStart(2, '0');
const isoDate = (y, m, d) => `${y}-${pad2(m + 1)}-${pad2(d)}`;
const genId = () => (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : `id-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
const monthKey = (y, m) => `month-${y}-${pad2(m + 1)}`;
// A month's "effective" schedule: whatever's actually in effect for that
// month right now — the generated current schedule (with manual overrides
// applied on top) once one exists, otherwise the master schedule as a
// best-effort stand-in.
const effectiveOf = (data) => !data ? {} : (data.currentScheduleGenerated
  ? { ...(data.currentSchedule || {}), ...(data.scheduleOverrides || {}) }
  : (data.masterSchedule || data.schedule || {}));
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

// Live check, not a stored snapshot: re-derives which dates in the CURRENT
// effective schedule (generated + manual overrides on top) break a hard
// rule right now — assigned to someone who declared that date unavailable,
// or calendar-adjacent to their own other shift. Deliberately independent
// of scheduleViolations (the one-time record of what buildCurrentSchedule's
// generator couldn't resolve at generation time) — that snapshot goes stale
// the moment an admin manually overrides a date afterward, in either
// direction: it won't flag a NEW violation a manual edit just introduced,
// and it keeps flagging an OLD one the admin already fixed by hand.
// boundary = { prevId, nextId, firstDate, lastDate } — whoever worked the
// day immediately before day 1 / immediately after the last day, in the
// NEIGHBORING month's own effective schedule. Without this, adjacency can
// only ever be checked against dates that are keys in scheduleLike itself,
// which is scoped to a single month — so nothing catches a manual edit
// that puts someone on, say, both Dec 31 and Jan 1.
function computeScheduleViolations(scheduleLike, unavailability, boundary = {}) {
  const { prevId = null, nextId = null, firstDate = null, lastDate = null } = boundary;
  const violations = new Set();
  Object.entries(scheduleLike).forEach(([date, docId]) => {
    if (!docId) return;
    if ((unavailability[docId] || []).includes(date)) violations.add(date);
    if (hasAdjacentAssignment(scheduleLike, date, docId)) violations.add(date);
    if (date === firstDate && prevId && docId === prevId) violations.add(date);
    if (date === lastDate && nextId && docId === nextId) violations.add(date);
  });
  return violations;
}

// Translates this app's own schedule model (one doctor "owns" a whole
// calendar day) into the department's real 3-shift duty roster: 00:01-
// 08:00, 08:00-16:00, 16:00-24:00. A weekday's doctor only covers the
// overnight block — 16:00 that day through 08:00 the next, since daytime
// hours are covered by regular staff, not this on-call rotation, so
// 08:00-16:00 is left blank. A holiday's doctor covers the entire day —
// 08:00 that day through 08:00 the next — filling both the 08:00-16:00
// and 16:00-24:00 slots. Either way, whoever a day's slot3 (16:00-24:00)
// belongs to is the same person covering slot1 (00:01-08:00) of the day
// immediately after, which is what carryIn threads through the loop.
// prevDayDoctorId seeds day 1's slot1 with whoever covered the tail end
// of the previous month (that doctor isn't otherwise visible in this
// month's own data at all).
function buildDutyRosterRows(year, month, effectiveSchedule, prevDayDoctorId, holidaySet) {
  const total = daysInMonth(year, month);
  const rows = [];
  let carryIn = prevDayDoctorId;
  for (let d = 1; d <= total; d++) {
    const date = isoDate(year, month, d);
    const docId = effectiveSchedule[date] || null;
    const isHoliday = dayType(date, holidaySet) === 'holiday';
    rows.push({
      date,
      day: d,
      dow: new Date(year, month, d).getDay(),
      slot1: carryIn,
      slot2: isHoliday ? docId : null,
      slot3: docId,
    });
    carryIn = docId;
  }
  return rows;
}

// Soft preference only — every caller uses this strictly as a tiebreak
// AFTER the real sort criteria (nominal owner, quota need), never in place
// of them, so it can't cause a date to go unfilled or push anyone over
// quota; it only nudges which otherwise-equally-good candidate gets tried
// first. Rewards whoever's nearest OTHER shift this month is farther from
// `date`, so a doctor's shifts naturally spread out to a few empty days
// apart instead of clustering at the tightest legally-allowed 1-day gap
// whenever there's a real choice available. d=1 (calendar-adjacent) is
// already excluded before candidates ever reach a sort, so the scale here
// starts at d=2 (one empty day between shifts, the tightest still-legal
// spacing) and rewards up to d=5 (four empty days) as progressively
// better; beyond that, or no other shift within the window at all, is
// treated as equally spread-out.
function spacingScore(assign, dateIndex, dates, date, docId) {
  const idx = dateIndex[date];
  for (let d = 2; d <= 5; d++) {
    if ((dates[idx - d] && assign[dates[idx - d]] === docId) || (dates[idx + d] && assign[dates[idx + d]] === docId)) {
      return d;
    }
  }
  return 6;
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
function exhaustiveSolveSchedule({ dates, doctors, quota, unavailSet, masterSchedule, holidaySet, budget, boundaryBlocked = () => false }) {
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
      .filter(id => (remaining[id]?.[type] || 0) > 0 && !unavailSet[id].has(date) && !neighborsOf(date).some(n => assign[n] === id) && !boundaryBlocked(date, id))
      .sort((a, b) => {
        // Preference only (doesn't affect completeness): try the master's
        // nominal owner first, then whoever needs shifts of this type most,
        // then — only among ties on those — whoever it'd spread out better.
        if (a === nominal) return -1;
        if (b === nominal) return 1;
        const qa = quota[a]?.[type] || 1, qb = quota[b]?.[type] || 1;
        const ratioDiff = ((remaining[b]?.[type] ?? 0) / qb) - ((remaining[a]?.[type] ?? 0) / qa);
        if (Math.abs(ratioDiff) > 1e-9) return ratioDiff;
        return spacingScore(assign, dateIndex, dates, date, b) - spacingScore(assign, dateIndex, dates, date, a);
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

// A recurring rule (e.g. "every Sunday", "every 2nd/4th Thursday") only
// becomes real dates once expanded against a specific month's calendar.
// unavailability as stored per month never includes these — they're
// re-derived on the fly wherever they're needed — so anything that reads a
// month's raw unavailability without also calling this is blind to whatever
// a doctor's recurring rules cover, and can schedule straight through them.
function expandRecurringUnavailability(rawUnavail, rules, year, month) {
  const merged = { ...rawUnavail };
  const total = daysInMonth(year, month);
  (rules || []).forEach(({ docId, dow, occurrences }) => {
    const toAdd = [];
    let nth = 0;
    for (let d = 1; d <= total; d++) {
      if (new Date(year, month, d).getDay() === dow) {
        nth++;
        if (occurrences.length === 0 || occurrences.includes(nth)) {
          toAdd.push(isoDate(year, month, d));
        }
      }
    }
    if (toAdd.length) {
      merged[docId] = [...new Set([...(merged[docId] || []), ...toAdd])].sort();
    }
  });
  return merged;
}

// A doctor's master-queue membership (WDQ = weekday queue, H12Q/H3Q = the two
// holiday queues, H3Q shared by the h3/h4/h5 loop types) is a STRUCTURAL fact
// about them, not something that varies month to month — someone absent from
// every holiday queue is never meant to work a holiday shift at all, the same
// way an unavailable date is never meant to be worked. A doctor's master
// quota for a type they're not queued for is always 0 (the queue rotation
// itself never assigns them one), but relying on that alone isn't enough: the
// "borrow" fallback in buildCurrentSchedule deliberately ignores quota to
// respect availability, and without this check it can just as easily hand a
// weekday-only doctor a holiday shift they were never eligible for.
function computeTypeEligibility(doctors, queueState) {
  const wdq = new Set(resolveQueue(queueState?.WDQ ?? DEFAULT_WDQ_NAMES, doctors));
  const h12q = new Set(resolveQueue(queueState?.H12Q ?? DEFAULT_H12Q_NAMES, doctors));
  const h3q = new Set(resolveQueue(queueState?.H3Q ?? DEFAULT_H3Q_NAMES, doctors));
  const eligibility = {};
  doctors.forEach(d => {
    eligibility[d.id] = { weekday: wdq.has(d.id), holiday: h12q.has(d.id) || h3q.has(d.id) };
  });
  return eligibility;
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
// boundaryPrevId/boundaryNextId = whoever worked the day immediately before
// day 1 / immediately after the last day, in the NEIGHBORING month's own
// schedule (a month this generation otherwise never looks at). Without
// these, nothing stops the same doctor landing on both sides of a month
// boundary — the normal same-month neighbor check has no way to see a day
// that isn't part of `dates` at all.
// debtIn = { [docId]: { weekday: N, holiday: N } } — a carry-forward
// adjustment to this month's quota, positive meaning "owed extra shifts"
// (fell short in a prior month of the same batch), negative meaning
// "already got extra" (should get fewer this month). Used by
// generateCurrentScheduleBatch to make a multi-month run's TOTAL assigned
// count match the TOTAL master-schedule quota across the batch, even when no
// single month in isolation has enough slack to match its own quota exactly
// (see the reported bug: a single month can be genuinely infeasible to
// balance perfectly once marketplace trades have reassigned specific dates).
function buildCurrentSchedule({ doctors, year, month, masterSchedule, unavailability, holidaySet, boundaryPrevId = null, boundaryNextId = null, debtIn = {}, eligibility = {}, futureRealQuota = {} }) {
  const total = daysInMonth(year, month);
  const dates = Array.from({ length: total }, (_, i) => isoDate(year, month, i + 1));
  const dateIndex = {};
  dates.forEach((d, i) => { dateIndex[d] = i; });
  const firstDate = dates[0], lastDate = dates[dates.length - 1];
  const boundaryBlocked = (date, id) =>
    (date === firstDate && boundaryPrevId && id === boundaryPrevId) ||
    (date === lastDate && boundaryNextId && id === boundaryNextId);

  const hasMasterData = Object.values(masterSchedule || {}).some(Boolean);
  if (!hasMasterData) {
    const empty = {}; dates.forEach(d => { empty[d] = null; });
    return { schedule: empty, violations: [], debtOut: debtIn };
  }

  // Quota = master-schedule quota, adjusted by any carried-in debt from a
  // prior month in the same batch. `target` is the *unclamped* debt-adjusted
  // figure (can go negative if someone was heavily over-serviced and this
  // month's own base quota isn't enough to absorb the correction); `quota` is
  // what actually gets fed to the solvers, clamped at 0 since neither solver
  // can be asked for a negative shift count. Keeping both matters: debtOut
  // below is computed against `target`, not `quota`, so whatever the clamp
  // trims off isn't silently lost — it just carries forward as debt again.
  const rawQuota = computeUsage(doctors, masterSchedule, holidaySet);
  // A doctor absent from a type's master queue entirely (see
  // computeTypeEligibility) can never target that type, no matter what debt
  // claims otherwise — their real master quota for it is always 0 anyway, so
  // this only matters as a safety net against stale/corrupted debt.
  const target = {};
  const quota = {};
  doctors.forEach(d => {
    const base = rawQuota[d.id] || { weekday: 0, holiday: 0 };
    const debt = debtIn[d.id] || {};
    const elig = eligibility[d.id] || { weekday: true, holiday: true };
    target[d.id] = {
      weekday: elig.weekday ? base.weekday + (debt.weekday || 0) : 0,
      holiday: elig.holiday ? base.holiday + (debt.holiday || 0) : 0,
    };
    quota[d.id] = { weekday: Math.max(0, target[d.id].weekday), holiday: Math.max(0, target[d.id].holiday) };
  });

  // debtOut = target - actually assigned, per doctor/type, computed from the
  // FINAL schedule regardless of which solver produced it. This is exact and
  // unclamped on purpose: since every date always ends up assigned to
  // exactly one doctor, sum(assigned) across doctors always equals the
  // number of dates of that type, so sum(debtOut) across doctors always
  // equals sum(target) - dates = sum(debtIn) exactly. In other words debt is
  // perfectly conserved batch-wide no matter how lopsided a single month's
  // solve is — nothing is ever gained or lost, only carried forward — which
  // is what actually lets a multi-month batch converge to zero debt for
  // every individual doctor rather than merely netting to zero on average.
  const computeDebtOut = (finalAssign) => {
    const assigned = {};
    doctors.forEach(d => { assigned[d.id] = { weekday: 0, holiday: 0 }; });
    dates.forEach(date => {
      const id = finalAssign[date];
      if (!id || !assigned[id]) return;
      assigned[id][dayType(date, holidaySet)] += 1;
    });
    const debtOut = {};
    doctors.forEach(d => {
      const w = target[d.id].weekday - assigned[d.id].weekday;
      const h = target[d.id].holiday - assigned[d.id].holiday;
      if (w !== 0 || h !== 0) debtOut[d.id] = { weekday: w, holiday: h };
    });
    return debtOut;
  };

  const unavailSet = {};
  doctors.forEach(d => { unavailSet[d.id] = new Set(unavailability[d.id] || []); });

  // Try for a mathematically perfect assignment first. Whenever one exists,
  // this is guaranteed to find it — no heuristic can promise that.
  const exhaustive = exhaustiveSolveSchedule({ dates, doctors, quota, unavailSet, masterSchedule, holidaySet, budget: 300000, boundaryBlocked });
  if (exhaustive.solved) {
    return { schedule: exhaustive.assign, violations: [], debtOut: computeDebtOut(exhaustive.assign) };
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
      .filter(id => id !== excludeDoctor && !unavailSet[id].has(date) && !boundaryBlocked(date, id))
      .sort((a, b) => {
        if (a === nominal) return -1;
        if (b === nominal) return 1;
        const qa = quota[a]?.[type] || 1, qb = quota[b]?.[type] || 1;
        const ratioDiff = ((remaining[b]?.[type] ?? 0) / qb) - ((remaining[a]?.[type] ?? 0) / qa);
        if (Math.abs(ratioDiff) > 1e-9) return ratioDiff;
        return spacingScore(assign, dateIndex, dates, date, b) - spacingScore(assign, dateIndex, dates, date, a);
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
  // master) adjacency-safe — then patch every date whose owner can't work
  // (including a boundary date whose nominal owner conflicts with the
  // neighboring month's actual assignment).
  dates.forEach(date => {
    const nominal = masterSchedule[date];
    if (nominal && remaining[nominal] && !unavailSet[nominal].has(date) && !boundaryBlocked(date, nominal)) {
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
      const type = dayType(date, holidaySet);

      // No exact-quota assignment exists for this date (solveDate exhausted
      // every relocation it could try). Before resorting to overriding
      // someone's stated unavailability, prefer "borrowing" against a
      // future month: assign anyone who is actually available, structurally
      // eligible for this type (in the relevant master queue — see
      // computeTypeEligibility), and not calendar-adjacent to an existing
      // assignment today, even though it puts them over their own quota for
      // this month specifically. That overage becomes debt (see debtOut
      // below) for the batch generator — or a future single-month
      // regeneration — to correct going forward.
      //
      // Two tiers, in order: first someone who already has real quota this
      // month (rawQuota > 0 — they used up a genuine allotment, the most
      // ordinary kind of "over quota"), then — only if nobody like that
      // exists — someone eligible whose queue turn hasn't landed THIS month
      // (rawQuota 0) but *will* land in a LATER month of this same batch
      // (futureRealQuota, computed by the caller by looking ahead across
      // the whole batch — see generateCurrentScheduleBatch). That second
      // tier is exactly what an admin reaches for by hand when the
      // alternative is forcing someone onto a day they declared
      // unavailable: e.g. borrowing a colleague's NEXT month's holiday
      // quota into this month, because paying it back later is provably
      // recoverable (there IS a later month to reduce) and a real
      // availability violation is not. Someone with no real quota anywhere
      // in the visible batch at all doesn't qualify for this tier — that's
      // not a debt that can ever actually be paid back within the window,
      // just an assignment they were never really due, so it's excluded
      // and we fall through to overriding availability below instead. Only
      // once both tiers come up empty does that happen.
      const eligibleAvailable = id => (eligibility[id]?.[type] ?? true) && !unavailSet[id].has(date) && !neighborsOf(date).some(n => assign[n] === id) && !boundaryBlocked(date, id);
      const byLeastOverQuota = (a, b) => {
        const diff = (remaining[b]?.[type] ?? 0) - (remaining[a]?.[type] ?? 0);
        if (diff !== 0) return diff;
        return spacingScore(assign, dateIndex, dates, date, b) - spacingScore(assign, dateIndex, dates, date, a);
      };
      const borrowCandidates = doctors.map(d => d.id).filter(id => eligibleAvailable(id) && (rawQuota[id]?.[type] || 0) > 0).sort(byLeastOverQuota);
      const zeroQuotaBorrowCandidates = doctors.map(d => d.id).filter(id => eligibleAvailable(id) && (rawQuota[id]?.[type] || 0) === 0 && (futureRealQuota[id]?.[type] ?? false)).sort(byLeastOverQuota);

      if (borrowCandidates.length > 0) {
        place(date, borrowCandidates[0], type);
        return;
      }
      if (zeroQuotaBorrowCandidates.length > 0) {
        place(date, zeroQuotaBorrowCandidates[0], type);
        return;
      }

      // Truly nobody is both available and non-adjacent today — last
      // resort, keep the nominal owner in place even though this date
      // couldn't be made to satisfy every rule (recorded above so the admin
      // can see exactly which dates need manual attention).
      const nominal = masterSchedule[date];
      let fallback;
      if (nominal && remaining[nominal] && !boundaryBlocked(date, nominal)) {
        fallback = nominal;
      } else {
        // The nominal owner isn't usable at all this month (inactive, or a
        // boundary conflict) — both borrow tiers above already failed too,
        // so this is a genuine "someone must be here" situation. Still work
        // down a priority order rather than grabbing the first non-adjacent
        // name blind to everything else: eligible candidates before
        // ineligible ones (someone outside the type's queue entirely is
        // never supposed to work it, the same rule enforced everywhere
        // else in this function), and among eligible candidates, anyone
        // with real or future-payable quota before someone with zero quota
        // for this entire batch (who would otherwise pick up a shift they
        // were never really due, in this single rarest of corners).
        const nonAdjacent = doctors.filter(d => !neighborsOf(date).some(n => assign[n] === d.id) && !boundaryBlocked(date, d.id));
        fallback =
          nonAdjacent.find(d => (eligibility[d.id]?.[type] ?? true) && ((rawQuota[d.id]?.[type] || 0) > 0 || (futureRealQuota[d.id]?.[type] ?? false)))?.id
          ?? nonAdjacent.find(d => (eligibility[d.id]?.[type] ?? true))?.id
          ?? nonAdjacent[0]?.id
          ?? doctors[0]?.id
          ?? null;
      }
      if (fallback) {
        assign[date] = fallback;
        remaining[fallback][type] = (remaining[fallback][type] ?? 0) - 1;
      }
    }
  });

  return { schedule: assign, violations: [...violations].sort(), debtOut: computeDebtOut(assign) };
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

// Lets admin pick a month range and generate all of them in one sequential
// run, carrying quota debt forward between months (see
// generateCurrentScheduleBatch) so the batch's total ends up matching the
// batch's total master-schedule quota even when no single month can on its
// own. Shows a per-month result + any leftover unresolved debt once done.
function BatchGenerateModal({ year, month, doctors, onClose, onRun, running }) {
  const [startYM, setStartYM] = useState({ year, month });
  const [endYM, setEndYM] = useState(() => {
    let m = month + 3, y = year;
    if (m > 11) { m -= 12; y += 1; }
    return { year: y, month: m };
  });
  const [result, setResult] = useState(null);
  // Names still owed a "ยืนยันว่าแจ้งครบแล้ว" for at least one month in the
  // chosen range — null until checked, [] once checked clean. Mirrors the
  // single-month generator's pendingCount warning (ConfirmModal further
  // down), which this batch flow otherwise has no equivalent of: without it
  // an admin can batch-generate several months while some doctors are still
  // mid-way through declaring their unavailable dates, silently locking in a
  // schedule that predates data those doctors hadn't submitted yet.
  const [pendingNames, setPendingNames] = useState(null);
  const [checking, setChecking] = useState(false);

  const shift = (which, delta) => {
    const setFn = which === 'start' ? setStartYM : setEndYM;
    setFn(prev => {
      let m = prev.month + delta, y = prev.year;
      if (m < 0) { m = 11; y -= 1; } else if (m > 11) { m = 0; y += 1; }
      return { year: y, month: m };
    });
    setPendingNames(null); // range changed — stale check, re-check before running
  };

  const monthCount = (endYM.year - startYM.year) * 12 + (endYM.month - startYM.month) + 1;
  const invalidRange = monthCount <= 0 || monthCount > 24;

  const doRun = async () => {
    const res = await onRun(startYM.year, startYM.month, endYM.year, endYM.month);
    setResult(res);
  };

  const handleRunClick = async () => {
    if (pendingNames !== null) { await doRun(); return; } // already warned (or clean) for this range
    setChecking(true);
    try {
      const monthsList = [];
      let cy = startYM.year, cm = startYM.month;
      while (true) {
        monthsList.push([cy, cm]);
        if (cy === endYM.year && cm === endYM.month) break;
        cm += 1; if (cm > 11) { cm = 0; cy += 1; }
        if (monthsList.length > 24) break;
      }
      const names = new Set();
      for (const [y, m] of monthsList) {
        const raw = await getMonthData(monthKey(y, m));
        if (!raw) continue;
        const master = raw.masterSchedule || raw.schedule || {};
        const hasShift = new Set(Object.values(master).filter(Boolean));
        const confirmed = new Set(raw.unavailabilityConfirmed || []);
        doctors.forEach(d => { if (hasShift.has(d.id) && !confirmed.has(d.id)) names.add(d.name); });
      }
      setPendingNames([...names]);
      if (names.size === 0) await doRun();
    } finally {
      setChecking(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-md w-full p-5 font-body">
        <h3 className="font-display font-semibold text-slate-800 text-lg mb-1">จัดตารางเวรหลายเดือน</h3>
        <p className="text-xs text-slate-400 mb-4">ระบบจะจัดทีละเดือนตามลำดับ ยกยอดจำนวนเวรที่ขาด/เกินไปชดเชยในเดือนถัดไป เพื่อให้รวมทั้งช่วงตรงกับโควต้าเวรล่าสุด</p>

        {!result ? (
          <>
            <div className="space-y-3 mb-4">
              <div>
                <p className="text-xs font-medium text-slate-600 mb-1">จากเดือน</p>
                <div className="flex items-center gap-2">
                  <button onClick={() => shift('start', -1)} disabled={running} className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-40"><ChevronLeft size={16} /></button>
                  <span className="text-sm text-slate-800 flex-1 text-center">{THAI_MONTHS[startYM.month]} {startYM.year + 543}</span>
                  <button onClick={() => shift('start', 1)} disabled={running} className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-40"><ChevronRight size={16} /></button>
                </div>
              </div>
              <div>
                <p className="text-xs font-medium text-slate-600 mb-1">ถึงเดือน</p>
                <div className="flex items-center gap-2">
                  <button onClick={() => shift('end', -1)} disabled={running} className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-40"><ChevronLeft size={16} /></button>
                  <span className="text-sm text-slate-800 flex-1 text-center">{THAI_MONTHS[endYM.month]} {endYM.year + 543}</span>
                  <button onClick={() => shift('end', 1)} disabled={running} className="p-1.5 rounded-lg hover:bg-slate-100 disabled:opacity-40"><ChevronRight size={16} /></button>
                </div>
              </div>
            </div>
            {invalidRange ? (
              <p className="text-xs text-red-500 mb-3">ช่วงเดือนไม่ถูกต้อง (ต้องไม่เกิน 24 เดือน และ &quot;ถึงเดือน&quot; ต้องไม่ก่อน &quot;จากเดือน&quot;)</p>
            ) : (
              <p className="text-xs text-slate-500 mb-3">รวม {monthCount} เดือน — ทุกเดือนที่มีตารางเวรอยู่แล้วจะถูกจัดใหม่ทับ</p>
            )}
            {pendingNames && pendingNames.length > 0 && (
              <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2 mb-3">
                <p className="font-medium mb-1">⚠️ ยังมี {pendingNames.length} คนที่ยังไม่ยืนยันว่าแจ้งวันไม่สะดวกครบในบางเดือนของช่วงนี้:</p>
                <p>{pendingNames.join(', ')}</p>
                <p className="mt-1">ถ้าจัดเวรตอนนี้ แล้วคนเหล่านี้แจ้งวันไม่สะดวกเพิ่มทีหลัง ตารางที่จัดไปแล้วจะไม่ปรับตามให้อัตโนมัติ ต้องจัดใหม่เอง</p>
              </div>
            )}
            <div className="flex justify-end gap-2">
              <button onClick={onClose} disabled={running || checking} className="px-4 py-2 rounded-lg text-sm font-medium text-slate-600 hover:bg-slate-100 disabled:opacity-50">ยกเลิก</button>
              <button
                onClick={handleRunClick}
                disabled={running || checking || invalidRange}
                className={`px-4 py-2 rounded-lg text-sm font-medium text-white disabled:opacity-50 disabled:cursor-wait ${pendingNames && pendingNames.length > 0 ? 'bg-red-600 hover:bg-red-700' : 'bg-teal-600 hover:bg-teal-700'}`}
              >
                {checking ? 'กำลังตรวจสอบ...' : running ? 'กำลังจัดเวร...' : (pendingNames && pendingNames.length > 0) ? 'จัดเวรเลย' : 'เริ่มจัดเวร'}
              </button>
            </div>
          </>
        ) : (
          <>
            <div className="space-y-1.5 mb-4 max-h-64 overflow-y-auto pr-1">
              {result.perMonth.map(r => (
                <p key={`${r.year}-${r.month}`} className="text-xs text-slate-600 flex items-center justify-between">
                  <span>{THAI_MONTHS[r.month]} {r.year + 543}</span>
                  <span className={r.violations > 0 ? 'text-red-500' : 'text-slate-400'}>{r.violations > 0 ? `${r.violations} วันจัดไม่ได้ตรงเงื่อนไข` : 'เรียบร้อย'}</span>
                </p>
              ))}
            </div>
            {Object.keys(result.finalDebt).length > 0 ? (
              <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2 mb-4">
                <p className="font-medium mb-1">ยังชดเชยไม่ครบภายในช่วงนี้ (เหลือติดไปเดือนถัดจากช่วงนี้):</p>
                {Object.entries(result.finalDebt).map(([docId, d]) => {
                  const doc = doctors.find(x => x.id === docId);
                  const parts = [];
                  if (d.weekday) parts.push(`วันธรรมดา ${d.weekday > 0 ? '+' : ''}${d.weekday}`);
                  if (d.holiday) parts.push(`วันหยุด ${d.holiday > 0 ? '+' : ''}${d.holiday}`);
                  return <p key={docId}>{doc?.name ?? '?'}: {parts.join(', ')}</p>;
                })}
              </div>
            ) : (
              <p className="text-xs text-emerald-600 mb-4">ยอดรวมทั้งช่วงตรงกับโควต้าเวรล่าสุดครบทุกคนแล้ว</p>
            )}
            <div className="flex justify-end">
              <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium text-white bg-teal-600 hover:bg-teal-700">เสร็จสิ้น</button>
            </div>
          </>
        )}
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

// Calendar-grid picker for marking national holidays (settings tab). Keeps
// its own month/year in local state — unlike the old single-date <input
// type="date"> + "เพิ่ม" button flow, navigating to a future month and
// clicking through several (often consecutive) days no longer resets back
// to the current month between each one, and each click toggles that date
// directly — no separate add step per day.
function HolidayPicker({ year, month, holidays, onToggle }) {
  const [viewYear, setViewYear] = useState(year);
  const [viewMonth, setViewMonth] = useState(month);
  const holidaySet = new Set(holidays);
  const total = daysInMonth(viewYear, viewMonth);
  const lead = new Date(viewYear, viewMonth, 1).getDay();
  const cells = [];
  for (let i = 0; i < lead; i++) cells.push(null);
  for (let d = 1; d <= total; d++) cells.push(isoDate(viewYear, viewMonth, d));

  const shiftView = (delta) => {
    let m = viewMonth + delta, y = viewYear;
    if (m < 0) { m = 11; y -= 1; } else if (m > 11) { m = 0; y += 1; }
    setViewMonth(m); setViewYear(y);
  };

  return (
    <div className="border border-slate-200 rounded-xl p-3 bg-white max-w-xs">
      <div className="flex items-center justify-between mb-2">
        <button type="button" onClick={() => shiftView(-1)} className="p-1 rounded hover:bg-slate-100"><ChevronLeft size={16} /></button>
        <span className="text-sm font-medium text-slate-700">{THAI_MONTHS[viewMonth]} {viewYear + 543}</span>
        <button type="button" onClick={() => shiftView(1)} className="p-1 rounded hover:bg-slate-100"><ChevronRight size={16} /></button>
      </div>
      <div className="grid grid-cols-7 gap-1 mb-1">
        {WEEKDAY_LABELS.map((w, i) => <div key={w} className={`text-center text-[10px] font-semibold ${i === 0 || i === 6 ? 'text-rose-500' : 'text-slate-400'}`}>{w}</div>)}
      </div>
      <div className="grid grid-cols-7 gap-1">
        {cells.map((date, i) => {
          if (!date) return <div key={`b-${i}`} />;
          const dow = new Date(viewYear, viewMonth, Number(date.slice(-2))).getDay();
          const isWeekend = dow === 0 || dow === 6;
          const isHoliday = holidaySet.has(date);
          const dayNum = Number(date.slice(-2));
          return (
            <button key={date} type="button" onClick={() => onToggle(date)} disabled={isWeekend}
              title={isWeekend ? 'เสาร์-อาทิตย์นับเป็นวันหยุดอัตโนมัติแล้ว' : ''}
              className={`rounded-md border text-[11px] font-mono py-1.5 transition-colors
                ${isWeekend ? 'bg-red-50 border-red-100 text-rose-300 cursor-default'
                  : isHoliday ? 'bg-rose-500 border-rose-500 text-white hover:bg-rose-600'
                  : 'bg-white border-slate-200 text-slate-600 hover:border-teal-400 cursor-pointer'}`}>
              {dayNum}
            </button>
          );
        })}
      </div>
    </div>
  );
}

// Lets admin pick a doctor to ring/highlight in the calendar below (instead
// of only ever seeing their own shifts ringed) — a quick way to recheck any
// one person's schedule by eye.
function DoctorHighlightPicker({ doctors, allDoctors, selectedId, onSelect }) {
  const colorFor = (id) => getDoctorColor(allDoctors.findIndex(d => d.id === id));
  return (
    <div className="flex flex-wrap items-center gap-1.5 mb-3">
      <span className="text-xs text-slate-500 mr-1">เช็คเวรของ:</span>
      <button onClick={() => onSelect(null)} className={`text-xs px-2.5 py-1 rounded-full border transition-colors ${!selectedId ? 'bg-teal-600 text-white border-transparent' : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300'}`}>
        ตัวฉันเอง
      </button>
      {doctors.map((d) => {
        const color = colorFor(d.id);
        const active = selectedId === d.id;
        return (
          <button key={d.id} onClick={() => onSelect(d.id)} className={`flex items-center gap-1.5 text-xs px-2.5 py-1 rounded-full border transition-colors ${active ? `${color.soft} ${color.text} border-transparent ring-1 ring-offset-1 ring-slate-300` : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300'}`}>
            <span className={`w-1.5 h-1.5 rounded-full ${color.bg}`} />{d.name}
          </button>
        );
      })}
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

// Static reference sequence per loop (not derived from live queue state) —
// the fixed rotation order the program is supposed to follow, independent
// of wherever the pointer currently sits. h3/h4/h5 all draw from the same
// underlying array (H3Q) so they share identical text here, but each still
// gets its own "เริ่มที่/จบที่" below since each has its own pointer.
const QUEUE_RUN_ORDER_TEXT = {
  weekday: 'อารีรัตน์ → ชุติมา → กนกอร → ธัญลักษณ์ → วัทนี → ธนวรรณ → ณัชพล → สมิตา → พสิษฐา → ณัฐธิดา → ขนิษฐา → ณัฐพล',
  h12: 'ชุติมา → ณัชพล → ณัฐพล → กนกอร → ธัญลักษณ์ → ณัฐธิดา → ขนิษฐา → ธนวรรณ → ณัชพล → ณัฐพล → วัทนี → สมิตา → ณัฐธิดา → ขนิษฐา → พสิษฐา',
  h3: 'ชุติมา → กนกอร → ธัญลักษณ์ → วัทนี → ธนวรรณ → ณัชพล → สมิตา → พสิษฐา → ณัฐธิดา → ขนิษฐา → ณัฐพล',
  h4: 'ชุติมา → กนกอร → ธัญลักษณ์ → วัทนี → ธนวรรณ → ณัชพล → สมิตา → พสิษฐา → ณัฐธิดา → ขนิษฐา → ณัฐพล',
  h5: 'ชุติมา → กนกอร → ธัญลักษณ์ → วัทนี → ธนวรรณ → ณัชพล → สมิตา → พสิษฐา → ณัฐธิดา → ขนิษฐา → ณัฐพล',
};

// Merges the static rotation reference (for recheck) with this month's
// actual start/end doctor per loop (previously a separate admin-only
// "สรุปคิวเดือนนี้" block) into one panel, visible to admin and doctors alike.
// Classifies every date in (y, m) into weekday/h12/h3/h4/h5, same logic used
// both for "this month" and for each historical month scanned below.
function classifyMonthDates(y, m, isHolidayDate) {
  const groups = detectGroups(y, m, isHolidayDate);
  const groupDateSet = new Set(groups.flatMap(g => g.dates));
  const datesByType = { weekday: [], h12: [], h3: [], h4: [], h5: [] };
  const total = daysInMonth(y, m);
  for (let d = 1; d <= total; d++) {
    const date = isoDate(y, m, d);
    if (!groupDateSet.has(date)) datesByType.weekday.push(date);
  }
  groups.forEach(g => { datesByType[ltFor(g.trueLength)].push(...g.dates); });
  return datesByType;
}

function QueueRunOrderSummary({ year, month, doctors, masterOriginal, holidays }) {
  const holidaySetAll = new Set(holidays);
  // Answers for ANY date, not just this month's — needed so a holiday
  // streak crossing a month boundary (e.g. 31 Dec – 3 Jan) is classified as
  // one continuous group using its true length, matching how it was
  // actually generated (see MasterScheduleGenerator's detectGroups).
  const isHolidayDate = (date) => {
    const dow = new Date(date + 'T00:00:00').getDay();
    return dow === 0 || dow === 6 || holidaySetAll.has(date);
  };

  const datesByType = classifyMonthDates(year, month, isHolidayDate);

  // "เดือนต่อไปเริ่มที่" is purely about the fixed rotation shown in
  // QUEUE_RUN_ORDER_TEXT (the "how the queue runs" reference), NOT about any
  // real generated schedule — so it's always literally the calendar month
  // right after the one being VIEWED, regardless of whether that next month
  // (or even the viewed month itself) has actually had its quota set yet.
  const nextCalMonth = month === 11 ? { y: year + 1, m: 0 } : { y: year, m: month + 1 };
  const nextMonthDatesByType = classifyMonthDates(nextCalMonth.y, nextCalMonth.m, isHolidayDate);
  const firstDateNextMonth = (key) => {
    const dates = nextMonthDatesByType[key];
    return dates.length > 0 ? [...dates].sort()[0] : null;
  };
  // Occurrence-safe "next name after lastName" lookup within the SAME fixed
  // text sequence already shown above (so it can never drift from what's
  // displayed) — duplicate names (only h12 has any) get a 1/2 suffix, same
  // convention used in the master-schedule generator. Used only as a
  // fallback when next month's real quota hasn't been set yet.
  const nextInRunOrder = (key, lastName) => {
    const arr = QUEUE_RUN_ORDER_TEXT[key].split(' → ').map(s => s.trim());
    const lastIdx = arr.indexOf(lastName);
    if (lastIdx === -1) return null;
    const nextIdx = (lastIdx + 1) % arr.length;
    const nextName = arr[nextIdx];
    const isDup = arr.filter(x => x === nextName).length > 1;
    const occ = isDup ? arr.slice(0, nextIdx + 1).filter(x => x === nextName).length : null;
    return occ ? `${nextName}${occ}` : nextName;
  };

  // Next month's REAL quota, if it's already been set — takes priority over
  // the theoretical rotation fallback above, the same way "เดือนนี้จบที่"
  // reflects real data rather than theory whenever real data exists.
  const [nextMonthMaster, setNextMonthMaster] = useState(undefined); // undefined = loading, null = no data
  useEffect(() => {
    let cancelled = false;
    setNextMonthMaster(undefined);
    getMonthData(monthKey(nextCalMonth.y, nextCalMonth.m)).then(data => {
      if (cancelled) return;
      setNextMonthMaster(data ? (data.masterOriginal || data.masterSchedule || data.schedule || {}) : null);
    });
    return () => { cancelled = true; };
  }, [nextCalMonth.y, nextCalMonth.m]);

  const nextMonthReal = {};
  if (nextMonthMaster) {
    ['weekday', 'h12', 'h3', 'h4', 'h5'].forEach(key => {
      const datesNextMonth = nextMonthDatesByType[key].filter(d => nextMonthMaster[d]).sort();
      if (datesNextMonth.length > 0) {
        nextMonthReal[key] = {
          date: datesNextMonth[0],
          name: doctors.find(d => d.id === nextMonthMaster[datesNextMonth[0]])?.name ?? '?',
        };
      }
    });
  }

  const thisMonth = {};
  ['weekday', 'h12', 'h3', 'h4', 'h5'].forEach(key => {
    const datesThisMonth = datesByType[key].filter(d => masterOriginal[d]).sort();
    if (datesThisMonth.length > 0) {
      thisMonth[key] = {
        firstDate: datesThisMonth[0],
        firstDoc: doctors.find(d => d.id === masterOriginal[datesThisMonth[0]])?.name ?? '?',
        lastDate: datesThisMonth[datesThisMonth.length - 1],
        lastDoc: doctors.find(d => d.id === masterOriginal[datesThisMonth[datesThisMonth.length - 1]])?.name ?? '?',
      };
    }
  });

  // "ล่าสุดจบที่" — the real most-recent assignment per loop, found by
  // actually reading past months' saved schedules (starting from the month
  // being viewed and walking backward) rather than trusting queueState's
  // stored lastDate, which has proven stale/wrong (e.g. showing a LATER
  // month than the one actually being viewed). Fetched in batches so a loop
  // type with no recent history doesn't require dozens of sequential
  // round-trips.
  const [lastReal, setLastReal] = useState(null);
  useEffect(() => {
    let cancelled = false;
    setLastReal(null);
    // Local to the effect (not the outer isHolidayDate closure) so the only
    // real dependency is `holidays` itself, not a function reference that's
    // new every render.
    const holidaySet = new Set(holidays);
    const isHoliday = (date) => {
      const dow = new Date(date + 'T00:00:00').getDay();
      return dow === 0 || dow === 6 || holidaySet.has(date);
    };
    (async () => {
      const found = {};
      const remaining = new Set(['weekday', 'h12', 'h3', 'h4', 'h5']);
      // Start from the month BEFORE the one being viewed, not the viewed
      // month itself — "ล่าสุดจบที่" means "where the queue stood coming
      // into this month," which should stay distinct from "เดือนนี้จบที่"
      // even when this month has its own data.
      let y = year, m = month - 1;
      if (m < 0) { m = 11; y -= 1; }
      const BATCH = 12, MAX_BATCHES = 3; // up to 36 months back
      for (let batch = 0; batch < MAX_BATCHES && remaining.size > 0 && !cancelled; batch++) {
        const keys = [], yms = [];
        for (let i = 0; i < BATCH; i++) {
          keys.push(monthKey(y, m));
          yms.push([y, m]);
          m -= 1; if (m < 0) { m = 11; y -= 1; }
        }
        const rows = await Promise.all(keys.map(k => getMonthData(k)));
        for (let i = 0; i < rows.length && remaining.size > 0; i++) {
          const data = rows[i];
          if (!data) continue;
          const [yy, mm] = yms[i];
          const master = data.masterOriginal || data.masterSchedule || data.schedule || {};
          const dbt = classifyMonthDates(yy, mm, isHoliday);
          remaining.forEach(key => {
            const datesOfType = dbt[key].filter(d => master[d]).sort();
            if (datesOfType.length > 0) {
              const lastDate = datesOfType[datesOfType.length - 1];
              found[key] = { date: lastDate, name: doctors.find(d => d.id === master[lastDate])?.name ?? '?' };
            }
          });
          [...remaining].forEach(key => { if (found[key]) remaining.delete(key); });
        }
      }
      if (!cancelled) setLastReal(found);
    })();
    return () => { cancelled = true; };
  }, [year, month, doctors, holidays]);

  const Box = ({ label, name, date, muted }) => (
    <div className="flex flex-col items-start gap-0.5">
      <span className="text-[9px] font-semibold tracking-wide text-slate-400 uppercase">{label}</span>
      <span className={`rounded-md px-2.5 py-1 text-xs font-semibold text-slate-800 ${muted ? 'border border-dashed border-slate-300' : 'border border-slate-700'}`}>{name}</span>
      <span className="text-[9.5px] text-slate-400">{date ? formatDisplayDate(date) : ''}</span>
    </div>
  );

  return (
    <div className="mt-4 border border-slate-200 rounded-xl px-3 py-3">
      <p className="text-xs font-medium text-slate-700 mb-3">วิธีการรันคิวเวร</p>
      <div>
        {['weekday', 'h12', 'h3', 'h4', 'h5'].map((key, idx) => {
          const tm = thisMonth[key];
          const lr = lastReal ? lastReal[key] : undefined; // undefined = still loading

          // "เดือนต่อไปเริ่มที่" — real next-month quota if it's already been
          // set, otherwise the theoretical rotation fallback computed from
          // whichever name is currently shown last ("เดือนนี้จบที่" if set,
          // else "ล่าสุดจบที่"). Waits for both the next-month fetch and the
          // historical scan to finish before deciding.
          const loaded = lastReal !== null && nextMonthMaster !== undefined;
          const real = loaded ? nextMonthReal[key] : null;
          const lastShownName = tm ? tm.lastDoc : (lr ? lr.name : null);
          const fallbackLabel = loaded && !real && lastShownName ? nextInRunOrder(key, lastShownName) : null;
          const nextLabel = real ? real.name : fallbackLabel;
          const nextDate = real ? real.date : (fallbackLabel ? firstDateNextMonth(key) : null);

          return (
            <div key={key} className={idx > 0 ? 'pt-3 mt-3 border-t border-slate-100' : ''}>
              <p className="text-[11px] font-medium text-slate-700 mb-1.5">{idx + 1}. เวร{QUEUE_LOOP_LABELS[key]}</p>
              <p className="text-[10.5px] text-slate-500 leading-relaxed mb-2">{QUEUE_RUN_ORDER_TEXT[key]}</p>
              <div className="flex items-end flex-wrap gap-x-2.5 gap-y-2">
                {lastReal === null ? (
                  <span className="text-[11px] text-slate-400">กำลังโหลด...</span>
                ) : lr ? (
                  <Box label="ล่าสุดจบที่" name={lr.name} date={lr.date} muted />
                ) : (
                  <span className="text-[11px] text-slate-400 italic">ยังไม่มีประวัติ</span>
                )}
                {tm && (
                  <>
                    <span className="text-slate-300 text-sm mb-2">→</span>
                    <Box label="เดือนนี้เริ่มที่" name={tm.firstDoc} date={tm.firstDate} />
                    <span className="text-slate-300 text-sm mb-2">→</span>
                    <Box label="เดือนนี้จบที่" name={tm.lastDoc} date={tm.lastDate} />
                  </>
                )}
                {nextLabel && (
                  <>
                    <span className="text-slate-300 text-sm mb-2">→</span>
                    <Box label="เดือนต่อไปเริ่มที่" name={nextLabel} date={nextDate} />
                  </>
                )}
              </div>
              {!tm && <p className="text-[11px] text-slate-400 italic mt-1.5">ไม่มีวันประเภทนี้ในเดือนนี้</p>}
            </div>
          );
        })}
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
        <StatusCard label="โควต้าเวร" tone={hasMasterData ? 'emerald' : 'amber'} value={hasMasterData ? 'ตั้งค่าแล้ว' : 'ยังไม่ได้ตั้งค่า'} onClick={() => onGotoTab('master')} />
        <StatusCard label="ยืนยันวันไม่สะดวก" tone={pendingDocs.length === 0 ? 'emerald' : 'amber'} value={`${doctorsWithShifts.length - pendingDocs.length}/${doctorsWithShifts.length} คนยืนยันแล้ว`} onClick={() => onGotoTab('unavailable')} />
        <StatusCard label="ตารางเวร" tone={currentStatus.tone} value={currentStatus.label} onClick={() => onGotoTab('current')} />
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

function CalendarGrid({ year, month, scheduleData, editable, onAssign, allDoctors, selectableDoctors, holidaySet, unavailability, marketplace, compareTo, highlightDoctorId, originalData, violationDates, hideUnavailableCount = false }) {
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
          if (diverged) titleParts.push(`เดิมตามโควต้าเวร: ${getDoctor(compareId)?.name || '-'}`);
          if (traded) titleParts.push(`ขาย/แลกจาก: ${getDoctor(origId)?.name || '-'}`);

          return (
            <div
              key={date}
              className={`relative rounded-lg border p-1.5 min-h-[64px] min-w-0 flex flex-col gap-1 ${type === 'holiday' ? 'bg-rose-100 border-rose-200' : 'bg-white border-slate-200'} ${diverged ? 'border-l-4 border-l-sky-400' : ''} ${isMine ? `ring-2 ring-offset-1 ${color.ring}` : ''} ${editable ? 'cursor-pointer hover:border-teal-300' : ''}`}
              onClick={() => editable && setEditingDate(date)}
              title={titleParts.join(' · ')}
            >
              <div className="flex items-center justify-between">
                <span className={`font-mono text-[11px] ${type === 'holiday' ? 'text-rose-700' : 'text-slate-500'}`}>{dayNum}</span>
                {hasOpenPost && <span className="w-1.5 h-1.5 rounded-full bg-amber-500" />}
                {isViolation && <span className="text-[10px] leading-none" title="วันนี้ไม่ตรงเงื่อนไข: อยู่เวรวันที่แจ้งไม่สะดวก หรืออยู่เวรติดกัน">⚠️</span>}
              </div>
              {isEditing ? (
                <select autoFocus className="text-[11px] font-body border rounded p-0.5 w-full" value={docId || ''} onClick={(e) => e.stopPropagation()} onChange={(e) => { onAssign(date, e.target.value || null); setEditingDate(null); }} onBlur={() => setEditingDate(null)}>
                  <option value="">ว่าง</option>
                  {selectableDoctors.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
                </select>
              ) : doc ? (
                // min-w-0 overrides flex items' default "won't shrink below
                // content" sizing, and w-full+block makes each name span's
                // width an explicit fact rather than something derived from
                // flex layout — needed because html2canvas doesn't reliably
                // replicate implicit flexbox shrink behavior.
                //
                // Names stay on one line and shrink to fit instead of
                // wrapping or truncating — via a real Tailwind responsive
                // breakpoint (text-[Npx] sm:text-[Mpx], an actual @media
                // rule), not clamp() (silently mis-parsed by html2canvas's
                // own CSS engine — see git history for that dead end).
                //
                // Deliberately NOT setting overflow-x: hidden here, even
                // though horizontal is the only axis that ever needs
                // containing: per the CSS overflow spec, pairing
                // overflow-x: hidden with overflow-y: visible doesn't
                // actually keep the y-axis visible — the visible value gets
                // silently forced to compute as auto instead whenever the
                // other axis is non-visible. html2canvas can't render a
                // real scrollbar in a static image, so that auto still
                // reads as clipped — which is what was actually cutting
                // Thai glyphs off this whole time, in every attempt that
                // paired the two, regardless of font-size mechanism. No
                // overflow rule at all (both axes default to visible,
                // never mismatched) avoids the trap entirely; the shrink-
                // to-fit sizing already keeps horizontal overflow rare
                // enough in practice not to need a hard clip as backup.
                //
                // leading-[Npx] instead of Tailwind's leading-relaxed
                // (line-height: 1.625, a UNITLESS ratio): html2canvas has a
                // known issue resolving unitless line-height against
                // font-size, which is what was pushing text toward the
                // bottom of its own padded box in the saved image while a
                // real browser (this preview, or the live admin view)
                // centered it correctly — same box, same padding, only the
                // renderer differed. An absolute pixel value removes the
                // ratio math html2canvas was getting wrong.
                <div className="flex flex-col gap-0.5 min-w-0">
                  <span
                    className={`block w-full text-center leading-[16px] sm:leading-[24px] font-body font-semibold rounded-md px-1 py-2 whitespace-nowrap text-[9px] sm:text-[14px] ${color.soft} ${color.text}`}
                  >{doc.name}</span>
                  {traded && (
                    <span className="block w-full text-center leading-[12px] sm:leading-[17px] font-body text-slate-400 line-through whitespace-nowrap px-1 text-[7px] sm:text-[10px]">{getDoctor(origId)?.name || '-'}</span>
                  )}
                </div>
              ) : (
                <span className="text-[10px] font-body text-slate-300">ยังไม่กำหนด</span>
              )}
              {!hideUnavailableCount && unavailDoctors.length > 0 && (
                <span className="block w-full font-body leading-[10px] sm:leading-[15px] text-slate-400 whitespace-nowrap text-[6px] sm:text-[9px]">{unavailDoctors.length} คนไม่สะดวก</span>
              )}
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
  // Snapshot of the GLOBAL queue state as it was right before THIS month's
  // own last generation (see handleMasterGenConfirm) — null if this month
  // has never been generated. Needed so re-opening "จัดโควต้าเวร" for an
  // already-generated month starts from the right place: the current global
  // queueState reflects wherever the LAST generation (which could be this
  // very month) left the pointers, not "right before this month ran".
  const [monthQueueSnapshot, setMonthQueueSnapshot] = useState(null);
  const [showMasterGen, setShowMasterGen] = useState(false);
  const [showBatchGen, setShowBatchGen] = useState(false);
  const [batchGenerating, setBatchGenerating] = useState(false);
  // Admin can pick a different doctor to highlight/recheck in the current &
  // master schedule calendars, instead of only ever seeing their own shifts
  // ringed. null = show the admin's own (default).
  const [recheckDoctorId, setRecheckDoctorId] = useState(null);

  const [toast, setToast] = useState(null);
  const [confirmState, setConfirmState] = useState(null);
  const [pendingAdjacentAssign, setPendingAdjacentAssign] = useState(null);
  // Wraps the current-schedule header + calendar so "บันทึกตารางเวร"
  // can capture exactly that region (not the whole page, not the nav
  // buttons above it) as an image.
  const currentScheduleCaptureRef = useRef(null);
  const [savingScheduleImage, setSavingScheduleImage] = useState(false);
  const [exportingDocx, setExportingDocx] = useState(false);
  const saveCurrentScheduleImage = async () => {
    const el = currentScheduleCaptureRef.current;
    if (!el) return;
    // savingScheduleImage also drives an export-only render tweak below
    // (the manual-edit blue-border indicator is hidden) that only makes
    // sense in a saved image, not the live admin view. Double rAF waits
    // for React to actually commit and paint that state change before
    // html2canvas captures the DOM — a single rAF (or none) risks the
    // capture happening before the browser has applied it.
    setSavingScheduleImage(true);
    await new Promise(resolve => requestAnimationFrame(resolve));
    await new Promise(resolve => requestAnimationFrame(resolve));
    try {
      // Lazy-loaded — this library is only needed for the rare "save as
      // image" click, no reason to bloat the main bundle everyone downloads.
      const html2canvas = (await import('html2canvas')).default;
      // html2canvas's own coordinate math doesn't account for how far the
      // page is currently scrolled, so whenever this element isn't at the
      // very top of the page the capture ends up offset — the real content
      // shifts down inside the canvas, leaving blank space above it and
      // pushing everything toward the bottom edge of the saved image.
      // Passing negative scrollX/scrollY corrects for that (the standard
      // fix for this well-known html2canvas issue).
      //
      // windowWidth forces html2canvas to lay the page out as if the
      // browser were this narrow, regardless of how wide the admin's own
      // window actually is right now — the calendar grid's cells have no
      // fixed/min pixel width, so at a phone-width viewport they naturally
      // reflow narrower (doctor names wrapping to a second line instead of
      // staying on one), which is what turns the capture from the usual
      // wide desktop layout into a portrait image that's already legible
      // at native size on a phone, with no artificial padding needed.
      const MOBILE_CAPTURE_WIDTH = 420;
      const canvas = await html2canvas(el, {
        backgroundColor: '#ffffff', scale: 2,
        scrollX: -window.scrollX, scrollY: -window.scrollY,
        windowWidth: MOBILE_CAPTURE_WIDTH,
        windowHeight: document.documentElement.scrollHeight,
      });
      const link = document.createElement('a');
      link.href = canvas.toDataURL('image/jpeg', 0.92);
      link.download = `${year + 543}_${pad2(month + 1)}.jpg`;
      link.click();
      showToast('บันทึกรูปภาพเรียบร้อย');
    } catch {
      showToast('บันทึกรูปภาพไม่สำเร็จ ลองอีกครั้ง');
    } finally {
      setSavingScheduleImage(false);
    }
  };

  // Builds the official 3-shift duty roster DOCX (see buildDutyRosterRows
  // for the day-to-shift mapping this is built on) and downloads it.
  const exportDutyRosterDocx = async () => {
    setExportingDocx(true);
    try {
      // Lazy-loaded, same reasoning as html2canvas above — only needed for
      // this one rare export click.
      const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, VerticalAlign, BorderStyle } = await import('docx');

      const DOC_FONT = 'TH SarabunPSK';
      const TABLE_SIZE = 28; // 14pt — docx sizes are in half-points
      const BODY_SIZE = 32;  // 16pt
      const PAGE_MARGIN_LR = 300;
      const PAGE_MARGIN_TB = 500;
      const CELL_MARGIN = 60;
      // Must fit "30 พ.ย. 69" (measured ~1440 twips) plus cell margins.
      const DATE_COL_WIDTH = 10;
      // Each name slot is TWO real table columns (head, surname), not one
      // column with a paragraph tab stop — a tab stop's position didn't
      // reliably separate the two in this library even with a proper Tab()
      // element and a verified-correct XML tab-stop definition (tried and
      // still failed), so this switches to something that can't have that
      // failure mode: an actual grid column boundary, which every Word-
      // compatible renderer aligns consistently across rows by construction.
      // Widths aren't tuned evenly — surnames run longer than the
      // title+first-name heads in this roster (measured), so the split
      // reflects that instead of a plain 50/50.
      const NAME_SLOT_WIDTH = (100 - DATE_COL_WIDTH * 2) / 3;
      const HEAD_COL_WIDTH = NAME_SLOT_WIDTH * 0.42;
      const SURNAME_COL_WIDTH = NAME_SLOT_WIDTH * 0.58;

      // Day 1's 00:01-08:00 slot belongs to whoever covered the tail end
      // of the PREVIOUS month — not visible anywhere in this month's own
      // data, so fetched fresh the same way boundary adjacency is
      // elsewhere in this file.
      const prevYM = month === 0 ? { y: year - 1, m: 11 } : { y: year, m: month - 1 };
      const prevData = await getMonthData(monthKey(prevYM.y, prevYM.m));
      const prevEff = effectiveOf(prevData);
      const prevLastDate = isoDate(prevYM.y, prevYM.m, daysInMonth(prevYM.y, prevYM.m));
      const prevDayDoctorId = prevEff[prevLastDate] || null;

      const rows = buildDutyRosterRows(year, month, effectiveSchedule, prevDayDoctorId, holidaySet);

      // { head: "นพ.ณัชพล", surname: "ทวีสกุลชัย" } — split on the LAST
      // space, since every DOCTOR_FULL_NAME entry is "prefix+firstname
      // surname" with exactly one space there.
      const nameOf = (docId) => {
        if (!docId) return null;
        const doc = doctors.find(d => d.id === docId);
        if (!doc) return null;
        const full = DOCTOR_FULL_NAME[doc.name] || doc.name;
        const idx = full.lastIndexOf(' ');
        return idx === -1 ? { head: full, surname: '' } : { head: full.slice(0, idx), surname: full.slice(idx + 1) };
      };

      const run = (text, opts = {}) => new TextRun({ text, font: DOC_FONT, size: opts.size ?? TABLE_SIZE, bold: !!opts.bold });

      const cellBorders = {
        top: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
        left: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
        right: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
      };
      const cell = (text, { colSpan = 1, bold = false, shading = null, width, align = AlignmentType.CENTER } = {}) => new TableCell({
        columnSpan: colSpan,
        verticalAlign: VerticalAlign.CENTER,
        borders: cellBorders,
        width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
        shading: shading ? { fill: shading } : undefined,
        children: [new Paragraph({ alignment: align, children: [run(text || '', { bold })] })],
      });
      // Head and surname are two separate real TableCells (own left-
      // aligned paragraph each), not one cell with a paragraph tab stop —
      // a fixed table-grid column boundary aligns identically across every
      // row by construction, the same way the date/day-name columns
      // already do, instead of depending on a tab stop actually landing
      // where it's told to (which didn't hold up in this library even with
      // a verified-correct tab-stop definition and a real Tab() element).
      // The shared internal edge (head's right / surname's left) is left
      // borderless so the pair still reads as one visual cell.
      const headBorders = { ...cellBorders, right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' } };
      const surnameBorders = { ...cellBorders, left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' } };
      const headSurnameCells = (name, { shading = null } = {}) => {
        if (!name) {
          return [
            new TableCell({ verticalAlign: VerticalAlign.CENTER, borders: headBorders, width: { size: HEAD_COL_WIDTH, type: WidthType.PERCENTAGE }, shading: shading ? { fill: shading } : undefined, children: [new Paragraph({ children: [] })] }),
            new TableCell({ verticalAlign: VerticalAlign.CENTER, borders: surnameBorders, width: { size: SURNAME_COL_WIDTH, type: WidthType.PERCENTAGE }, shading: shading ? { fill: shading } : undefined, children: [new Paragraph({ children: [] })] }),
          ];
        }
        return [
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            borders: headBorders,
            width: { size: HEAD_COL_WIDTH, type: WidthType.PERCENTAGE },
            shading: shading ? { fill: shading } : undefined,
            children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [run(name.head)] })],
          }),
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            borders: surnameBorders,
            width: { size: SURNAME_COL_WIDTH, type: WidthType.PERCENTAGE },
            shading: shading ? { fill: shading } : undefined,
            children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [run(name.surname)] })],
          }),
        ];
      };

      const headerRow = new TableRow({
        tableHeader: true,
        children: [
          cell('วัน เดือน ปี', { colSpan: 2, bold: true, width: DATE_COL_WIDTH * 2 }),
          cell('เวลา 00.01 น. - 08.00 น.', { colSpan: 2, bold: true, width: NAME_SLOT_WIDTH }),
          cell('เวลา 08.00 น. - 16.00 น.', { colSpan: 2, bold: true, width: NAME_SLOT_WIDTH }),
          cell('เวลา 16.00 น. - 24.00 น.', { colSpan: 2, bold: true, width: NAME_SLOT_WIDTH }),
        ],
      });

      const HOLIDAY_SHADING = 'FCE4EC';
      const dataRows = rows.map(r => {
        const shading = dayType(r.date, holidaySet) === 'holiday' ? HOLIDAY_SHADING : null;
        return new TableRow({
          children: [
            cell(`${r.day} ${THAI_MONTHS_SHORT[month]} ${String(year + 543).slice(-2)}`, { shading, width: DATE_COL_WIDTH }),
            cell(THAI_WEEKDAYS_FULL[r.dow], { shading, width: DATE_COL_WIDTH }),
            ...headSurnameCells(nameOf(r.slot1), { shading }),
            ...headSurnameCells(nameOf(r.slot2), { shading }),
            ...headSurnameCells(nameOf(r.slot3), { shading }),
          ],
        });
      });

      // 100% table width is relative to the page's own content area, so
      // shrinking the left/right margins (below) is what actually makes
      // the table run nearly edge to edge — the percentage alone doesn't.
      // Explicit (tight) cell margins, not the library's default ~115
      // twips each side, since every twip matters for the name columns'
      // one-line fit (see the width math above).
      const table = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        margins: { top: CELL_MARGIN, bottom: CELL_MARGIN, left: CELL_MARGIN, right: CELL_MARGIN },
        rows: [headerRow, ...dataRows],
      });

      const doc = new Document({
        sections: [{
          // Portrait (the library's default — not overriding orientation)
          // per explicit preference, even though the width math below
          // shows it's genuinely tight for the longest names.
          properties: {
            page: {
              margin: { top: PAGE_MARGIN_TB, bottom: PAGE_MARGIN_TB, left: PAGE_MARGIN_LR, right: PAGE_MARGIN_LR },
            },
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
              children: [run(`เวรประจำหน้าที่กลุ่มงานกุมารเวชกรรม อยู่เวร ประจำเดือน${THAI_MONTHS[month]} ${year + 543}`, { bold: true, size: BODY_SIZE })],
            }),
            table,
            new Paragraph({ spacing: { before: 400 }, children: [] }),
            new Paragraph({ spacing: { before: 200 }, children: [] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [run(`(${DUTY_ROSTER_HEAD_NAME})`, { size: BODY_SIZE })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [run(DUTY_ROSTER_HEAD_TITLE, { size: BODY_SIZE })] }),
          ],
        }],
      });

      const blob = await Packer.toBlob(doc);
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `เวรแพทย์กลุ่มงานกุมารเวชกรรม ประจำเดือน${THAI_MONTHS[month]}${year + 543}.docx`;
      link.click();
      URL.revokeObjectURL(link.href);
      showToast('บันทึกไฟล์ DOCX เรียบร้อย');
    } catch (err) {
      console.error(err);
      showToast('สร้างไฟล์ DOCX ไม่สำเร็จ ลองอีกครั้ง');
    } finally {
      setExportingDocx(false);
    }
  };

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

  // Whoever worked the day immediately before day 1 / immediately after the
  // last day, in the NEIGHBORING month's own effective schedule — fetched
  // fresh whenever the viewed month changes, same fresh-read pattern as the
  // rest of this app (no realtime sync). Needed so the live violation check
  // below can catch a manual edit that creates a cross-month adjacency
  // conflict (e.g. Dec 31 and Jan 1 both landing on the same doctor), which
  // a check scoped to a single month's own data can never see on its own.
  const [monthBoundary, setMonthBoundary] = useState({ prevId: null, nextId: null });
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const prevYM = month === 0 ? { y: year - 1, m: 11 } : { y: year, m: month - 1 };
      const nextYM = month === 11 ? { y: year + 1, m: 0 } : { y: year, m: month + 1 };
      const [prevMonthData, nextMonthData] = await Promise.all([
        getMonthData(monthKey(prevYM.y, prevYM.m)),
        getMonthData(monthKey(nextYM.y, nextYM.m)),
      ]);
      if (cancelled) return;
      setMonthBoundary({
        prevId: prevMonthData ? (effectiveOf(prevMonthData)[isoDate(prevYM.y, prevYM.m, daysInMonth(prevYM.y, prevYM.m))] || null) : null,
        nextId: nextMonthData ? (effectiveOf(nextMonthData)[isoDate(nextYM.y, nextYM.m, 1)] || null) : null,
      });
    })();
    return () => { cancelled = true; };
  }, [year, month]);

  // Recomputed on every render the effective schedule, unavailability, or
  // month boundary changes — see computeScheduleViolations for why this
  // replaced reading the stored scheduleViolations snapshot directly in the
  // calendar.
  const liveViolations = useMemo(() => {
    const total = daysInMonth(year, month);
    return computeScheduleViolations(effectiveSchedule, unavailability, {
      ...monthBoundary,
      firstDate: isoDate(year, month, 1),
      lastDate: isoDate(year, month, total),
    });
  }, [effectiveSchedule, unavailability, monthBoundary, year, month]);

  // Cross-device data (doctor roster, queue state, marketplace posts,
  // notifications) is only ever fetched here — there's no realtime sync, so
  // a browser tab left open won't see what another device did in the
  // meantime. Re-running this on login and whenever the marketplace tab is
  // opened is what actually surfaces those changes; setting queueState to a
  // freshly-fetched object also cascades into a re-fetch of the current
  // month's own data, since that effect depends on it.
  const refreshData = useCallback(async () => {
    const [dbDoctors, cfg, qs] = await Promise.all([getDoctors(), storageGet('config', { holidays: [] }), getQueueState()]);
    setDoctors(dbDoctors);
    setHolidays(cfg.holidays || []);
    setMarketplace(await storageGet('marketplace', []));
    setNotifications(await storageGet('notifications', []));
    setQueueStateLocal(qs);
  }, []);

  useEffect(() => {
    (async () => {
      await refreshData();
      setLoading(false);
    })();
  }, [refreshData]);

  // Marketplace data is the most time-sensitive to cross-device staleness
  // (someone else's sell/accept becoming invisible until reload) — refresh
  // it every time this tab is opened, not just once at page load.
  useEffect(() => {
    if (activeTab === 'marketplace' && currentUser) refreshData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTab]);

  useEffect(() => {
    (async () => {
      const data = await storageGet(monthKey(year, month), null);

      // Auto-apply recurring rules (saved in queueState) for this month —
      // computed regardless of whether month_data exists yet, so a rule
      // shows up immediately even for a future month with no master
      // schedule saved at all.
      const rawUnavail = (data && data.unavailability) || {};
      const rules = (queueState || {}).recurringRules || [];
      const mergedUnavail = expandRecurringUnavailability(rawUnavail, rules, year, month);

      if (!data) {
        setMasterSchedule({}); setMasterOriginal({}); setCurrentSchedule({}); setCurrentScheduleGenerated(false); setScheduleStale(false); setScheduleViolations([]);
        setScheduleOverrides({}); setUnavailability(mergedUnavail); setUnavailabilityConfirmed([]); setActiveDoctorIds(null);
        setMonthQueueSnapshot(null);
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
        setMonthQueueSnapshot(data.queueStateBeforeGen || null);
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
  // Fresh-read-then-merge-patch the CURRENTLY LOADED month's record — never
  // a blind whole-record overwrite built from local React state. This app
  // has no realtime sync, so a tab can sit open while another device/tab
  // writes to the same month in the meantime (a marketplace trade, another
  // doctor's report, a master-schedule generation's queue snapshot, etc).
  // Reading fresh right before writing means only the fields actually named
  // in `patch` change — everything else (including fields this call site
  // doesn't even know about) survives untouched.
  const saveMonth = async (patch) => {
    const mk = monthKey(year, month);
    const raw = (await getMonthData(mk)) || {};
    await setMonthData(mk, { ...raw, ...patch });
  };
  // Same fresh-read-then-merge pattern as saveMonth, but for callers that
  // need to compute their patch FROM the fresh data too (e.g. toggling one
  // doctor's unavailable date without clobbering everyone else's) rather
  // than just writing fixed values.
  const patchCurrentMonth = async (mutate) => {
    const mk = monthKey(year, month);
    const raw = (await getMonthData(mk)) || {};
    const patch = mutate(raw);
    await setMonthData(mk, { ...raw, ...patch });
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
      showToast(`นำเข้าโควต้าเวร ${count} วันสำเร็จ${skippedMonth ? ` (ข้าม ${skippedMonth} วันที่ไม่ตรงเดือนนี้)` : ''} — ต้องกดจัดเวรใหม่`);
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
      addNotification(`โควต้าเวรวันที่ ${formatDisplayDate(date)} กำหนดเป็น ${newName} (เดิม ${oldName})`, `🔔 โควต้าเวรวันที่ ${formatDisplayDate(date)}: ${oldName} → ${newName}`);
    }
  };

  const applyManualAssignCurrent = (date, docId) => {
    const oldEff = effectiveSchedule[date] || null;
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

  // Manual edits are a deliberate admin override, not the automatic
  // generator — so an adjacent-day conflict here is a warn-and-confirm, not
  // a hard block like it is for buildCurrentSchedule. The admin can see
  // exactly why (unavailability across the rest of the month, a trade that
  // already happened, etc.) and may have a real reason only they know
  // about to accept it anyway.
  const manualAssignCurrent = (date, docId) => {
    if (docId && hasAdjacentAssignment({ ...effectiveSchedule, [date]: docId }, date, docId)) {
      setPendingAdjacentAssign({ date, docId });
      return;
    }
    applyManualAssignCurrent(date, docId);
  };

  // The only place the current schedule is actually computed — admin
  // triggers this explicitly (via "จัดเวร") once the master schedule and
  // everyone's availability have settled, rather than it recomputing itself
  // in the background on every small change.
  // Whenever a month's master schedule is generated (first time or a re-do
  // after clearing), any LATER month that was already generated has its own
  // "before" queue snapshot computed from whatever the shared queue state
  // was prior to THIS change — now stale, regardless of what its last-used
  // date happens to say. Comparing dates alone can't catch this: adding or
  // removing a holiday changes how many slots a month actually consumes, so
  // the same calendar date can end up mapped to a different queue position
  // than before. Rather than try to detect that, just invalidate outright —
  // if a later month's snapshot is gone, clearing it later won't roll the
  // queue back to anything (safe default) instead of silently restoring a
  // now-wrong value. Scans a bounded 12-month window forward; nobody
  // realistically pre-generates further ahead than that.
  const invalidateDownstreamSnapshots = async (afterYear, afterMonth) => {
    const HORIZON = 12;
    const keys = [];
    let y = afterYear, m = afterMonth;
    for (let i = 0; i < HORIZON; i++) {
      m += 1;
      if (m > 11) { m = 0; y += 1; }
      keys.push(monthKey(y, m));
    }
    const rows = await Promise.all(keys.map(mk => getMonthData(mk)));
    await Promise.all(rows.map((raw, i) =>
      (raw && raw.queueStateBeforeGen) ? setMonthData(keys[i], { ...raw, queueStateBeforeGen: null }) : null
    ));
  };

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
    // Snapshot the GLOBAL queue state exactly as it was right before THIS
    // month ever started generating — onto this month's own record. The
    // queue pointers are shared/cumulative across all months, so both
    // clearing this month later (resetMasterSchedule) AND simply reopening
    // "จัดโควต้าเวร" to regenerate it again (see monthQueueSnapshot use
    // below) need this to resume from the prior month instead of wherever
    // this month's own last run advanced the pointers to. Prefer the
    // EXISTING snapshot if this month already has one (i.e. this isn't the
    // first generation) so re-generating repeatedly never drifts the
    // remembered "before" point forward.
    const beforeGen = monthQueueSnapshot ?? queueState;
    await saveMonth({ masterSchedule: nextMaster, masterOriginal: nextOriginal, scheduleOverrides: {}, currentSchedule: {}, currentScheduleGenerated: false, scheduleStale: false, queueStateBeforeGen: beforeGen });
    setMonthQueueSnapshot(beforeGen);
    // Persist new queue state. newQueueState.debt already reflects only what
    // was actually consumed this generation (see MasterScheduleGenerator's
    // handleConfirm) — untouched debt (e.g. a loop type with no groups this
    // month) must carry forward, not be wiped here too.
    setQueueStateLocal(newQueueState);
    await setQueueState(newQueueState);
    await invalidateDownstreamSnapshots(year, month);
    await addNotification(
      `จัดโควต้าเวร ${['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][month]} ${year + 543} สำเร็จแล้ว`,
      `📅 จัดโควต้าเวร ${['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][month]} ${year + 543} สำเร็จแล้ว`
    );
    setShowMasterGen(false);
    showToast('บันทึกโควต้าเวรเรียบร้อย');
  };

  const generateCurrentSchedule = async () => {
    // Cross-month adjacency: who worked the day immediately before/after
    // this month, in THAT month's own effective schedule (current schedule
    // if it's been generated, else its master schedule as a best-effort
    // fallback) — fetched fresh since neither neighboring month is loaded
    // into this session's state. Without this, nothing stops the exact same
    // doctor landing on both the last day of one month and the first day of
    // the next.
    const prevYM = month === 0 ? { y: year - 1, m: 11 } : { y: year, m: month - 1 };
    const nextYM = month === 11 ? { y: year + 1, m: 0 } : { y: year, m: month + 1 };
    const [prevMonthData, nextMonthData] = await Promise.all([
      getMonthData(monthKey(prevYM.y, prevYM.m)),
      getMonthData(monthKey(nextYM.y, nextYM.m)),
    ]);
    const boundaryPrevId = prevMonthData ? (effectiveOf(prevMonthData)[isoDate(prevYM.y, prevYM.m, daysInMonth(prevYM.y, prevYM.m))] || null) : null;
    const boundaryNextId = nextMonthData ? (effectiveOf(nextMonthData)[isoDate(nextYM.y, nextYM.m, 1)] || null) : null;

    const eligibility = computeTypeEligibility(activeDoctors, queueState);
    const { schedule: next, violations } = buildCurrentSchedule({ doctors: activeDoctors, year, month, masterSchedule, unavailability, holidaySet, boundaryPrevId, boundaryNextId, eligibility });
    const violationList = [...violations].sort();
    setCurrentSchedule(next);
    setScheduleViolations(violationList);
    setCurrentScheduleGenerated(true);
    setScheduleStale(false);
    setScheduleOverrides({});
    await saveMonth({ currentSchedule: next, scheduleViolations: violationList, currentScheduleGenerated: true, scheduleStale: false, scheduleOverrides: {} });
    const msg = `จัดเวรตารางเวรสำหรับเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว${violationList.length ? ` (มี ${violationList.length} วันที่จัดให้ตรงเงื่อนไขไม่ได้แม้ลองสลับเวรหลายคู่แล้ว — เจ้าของเวรเดิมต้องอยู่แทน)` : ''}`;
    await addNotification(msg, `🔀 ${msg}`);
    showToast('จัดเวรเรียบร้อย');
  };

  // Generates a run of consecutive months' current schedules in one go,
  // carrying leftover quota debt forward from each month into the next
  // (see buildCurrentSchedule's debtIn/debtOut) so the batch's TOTAL actual
  // assignments end up matching the batch's TOTAL master-schedule quota,
  // even where any single month in isolation couldn't (a real limitation —
  // marketplace trades can reassign specific dates in ways that make one
  // month, on its own, mathematically unable to hit its own quota exactly).
  const generateCurrentScheduleBatch = async (startY, startM, endY, endM) => {
    const monthsList = [];
    let cy = startY, cm = startM;
    while (true) {
      monthsList.push([cy, cm]);
      if (cy === endY && cm === endM) break;
      cm += 1; if (cm > 11) { cm = 0; cy += 1; }
      if (monthsList.length > 24) break; // safety cap — batches this long aren't a real workflow
    }

    // Queue membership is a structural property of the doctor, not the
    // month, so it's computed once against the full roster and reused —
    // each call to buildCurrentSchedule below only reads the entries for
    // that month's own active doctors.
    const eligibility = computeTypeEligibility(doctors, queueState);

    // Every month's raw data is fetched up front (not lazily per iteration)
    // so a month being solved can see whether each doctor has REAL master
    // quota (rawQuota > 0, before any debt) in a LATER month of this same
    // batch — the difference between "borrow from someone whose queue turn
    // is later this batch" (recoverable, provably paid back by an actual
    // quota reduction later) and "borrow from someone whose turn never
    // comes up in this whole window at all" (a debt with nowhere real to
    // land, sitting unresolved in finalDebt forever). See the
    // futureRealQuota computation below and its use in buildCurrentSchedule.
    const rawByMonth = await Promise.all(monthsList.map(([y, m]) => getMonthData(monthKey(y, m))));
    const rawQuotaByMonth = rawByMonth.map((raw, i) => {
      const master = (raw && (raw.masterSchedule || raw.schedule)) || {};
      return computeUsage(doctors, master, holidaySet);
    });

    let debt = {};
    let prevEffective = null; // in-memory result of the PREVIOUS batch month, for boundary adjacency without an extra round-trip
    const perMonth = [];

    for (let i = 0; i < monthsList.length; i++) {
      const [y, m] = monthsList[i];
      const mk = monthKey(y, m);
      const raw = rawByMonth[i] || {};
      const monthMaster = raw.masterSchedule || raw.schedule || {};
      const futureRealQuota = {};
      doctors.forEach(d => {
        futureRealQuota[d.id] = {
          weekday: rawQuotaByMonth.slice(i + 1).some(q => (q[d.id]?.weekday || 0) > 0),
          holiday: rawQuotaByMonth.slice(i + 1).some(q => (q[d.id]?.holiday || 0) > 0),
        };
      });
      // A month's stored unavailability never includes recurring rules —
      // those are only ever expanded on the fly against a specific month's
      // calendar (see expandRecurringUnavailability). Reading raw.unavailability
      // straight from storage the way single-month generation's live React
      // state never does would silently ignore every doctor's recurring
      // unavailable days for any month the admin hasn't individually opened
      // and saved yet — exactly the kind of month a multi-month batch run is
      // for.
      const monthUnavail = expandRecurringUnavailability(raw.unavailability || {}, queueState?.recurringRules || [], y, m);
      const monthActiveIds = raw.activeDoctorIds !== undefined ? raw.activeDoctorIds : null;
      const monthActiveDoctors = monthActiveIds === null ? doctors : doctors.filter(d => monthActiveIds.includes(d.id));

      let boundaryPrevId;
      if (i > 0) {
        const [py, pm] = monthsList[i - 1];
        boundaryPrevId = prevEffective[isoDate(py, pm, daysInMonth(py, pm))] || null;
      } else {
        const prevYM = m === 0 ? { y: y - 1, m: 11 } : { y, m: m - 1 };
        const prevData = await getMonthData(monthKey(prevYM.y, prevYM.m));
        boundaryPrevId = effectiveOf(prevData)[isoDate(prevYM.y, prevYM.m, daysInMonth(prevYM.y, prevYM.m))] || null;
      }
      // The next month, whether inside or outside this batch, hasn't been
      // (re)generated by the time we get here in a forward sequential run —
      // same best-effort fallback the single-month generator already uses
      // (its master schedule's nominal owner) when the next month has no
      // current schedule yet.
      const nextYM = m === 11 ? { y: y + 1, m: 0 } : { y, m: m + 1 };
      const nextData = await getMonthData(monthKey(nextYM.y, nextYM.m));
      const boundaryNextId = effectiveOf(nextData)[isoDate(nextYM.y, nextYM.m, 1)] || null;

      const { schedule: nextSchedule, violations, debtOut } = buildCurrentSchedule({
        doctors: monthActiveDoctors, year: y, month: m,
        masterSchedule: monthMaster, unavailability: monthUnavail, holidaySet,
        boundaryPrevId, boundaryNextId, debtIn: debt, eligibility, futureRealQuota,
      });
      const violationList = [...violations].sort();

      await setMonthData(mk, { ...raw, currentSchedule: nextSchedule, scheduleViolations: violationList, currentScheduleGenerated: true, scheduleStale: false, scheduleOverrides: {} });
      if (y === year && m === month) {
        setCurrentSchedule(nextSchedule);
        setScheduleViolations(violationList);
        setCurrentScheduleGenerated(true);
        setScheduleStale(false);
        setScheduleOverrides({});
      }

      // buildCurrentSchedule only tracks debt for doctors active THIS month —
      // preserve anyone else's carried-in debt unchanged so it isn't
      // silently dropped just because they had no shifts this month.
      const activeIdSet = new Set(monthActiveDoctors.map(d => d.id));
      const mergedDebt = { ...debtOut };
      Object.keys(debt).forEach(id => { if (!activeIdSet.has(id) && debt[id]) mergedDebt[id] = debt[id]; });

      perMonth.push({ year: y, month: m, violations: violationList.length, debtOut: mergedDebt });
      prevEffective = nextSchedule;
      debt = mergedDebt;
    }

    return { monthsList, perMonth, finalDebt: debt };
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
      `ล้างตารางเวรของเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว (ต้องกดจัดเวรใหม่)`,
      `🗑️ ล้างตารางเวรของเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว`
    );
    showToast('ล้างตารางเวรแล้ว');
  };

  // Wipe the master schedule (and everything derived from it — overrides,
  // current schedule) back to blank for this month. Unlike a re-upload, this
  // leaves the doctor roster untouched.
  const resetMasterSchedule = async () => {
    // The shared queue pointers are cumulative across months — if this
    // month's generation snapshotted "before" state (see
    // handleMasterGenConfirm), roll the pointers back to it too, otherwise
    // clearing + regenerating this same month would keep using wherever
    // this generation had already advanced them to (i.e. skip ahead into
    // the following month's territory) instead of resuming from the prior
    // month, as reported.
    const mk = monthKey(year, month);
    const raw = (await getMonthData(mk)) || {};
    const snapshot = raw.queueStateBeforeGen || null;

    setMasterSchedule({});
    setMasterOriginal({});
    setScheduleOverrides({});
    setCurrentSchedule({});
    setCurrentScheduleGenerated(false);
    setScheduleViolations([]);
    setScheduleStale(false);
    await saveMonth({ masterSchedule: {}, masterOriginal: {}, scheduleOverrides: {}, currentSchedule: {}, currentScheduleGenerated: false, scheduleViolations: [], scheduleStale: false, queueStateBeforeGen: null });
    setMonthQueueSnapshot(null);

    // Roll the shared queue pointers back to this month's snapshot — but
    // per loop and only ever BACKWARD. A blind whole-state restore breaks
    // when several months are cleared oldest-first: clearing Nov correctly
    // rewinds to post-Oct, but then clearing Dec would restore DEC's
    // snapshot (= post-Nov) and shove the queue forward past November
    // again. So each loop's pointer is restored only if the snapshot's
    // lastDate for that loop isn't LATER than where the queue currently
    // stands; and only pointer/lastDate/debt are touched — recurring rules
    // and custom queue orders always keep their current values (they're
    // settings, not queue position, and may have changed since the
    // snapshot was taken).
    let queueMsg = '';
    if (snapshot) {
      const cur = await getQueueState(); // fresh — not this tab's React copy
      const LOOP_KEYS = ['weekday', 'h12', 'h3', 'h4', 'h5'];
      const DEBT_KEY = { weekday: 'WDQ', h12: 'H12Q', h3: 'H3Q', h4: 'H4Q', h5: 'H5Q' };
      const restoredKeys = LOOP_KEYS.filter(k => {
        const snapLast = snapshot.lastDate?.[k] || null;
        const curLast = cur.lastDate?.[k] || null;
        return !(snapLast && curLast && snapLast > curLast);
      });
      if (restoredKeys.length > 0) {
        const next = { ...cur, lastDate: { ...(cur.lastDate || {}) } };
        restoredKeys.forEach(k => {
          next[k] = snapshot[k];
          next.lastDate[k] = snapshot.lastDate?.[k] || null;
        });
        const nextDebt = {};
        new Set([...Object.keys(cur.debt || {}), ...Object.keys(snapshot.debt || {})]).forEach(docId => {
          const entry = {};
          LOOP_KEYS.forEach(k => {
            const dk = DEBT_KEY[k];
            const v = restoredKeys.includes(k) ? snapshot.debt?.[docId]?.[dk] : cur.debt?.[docId]?.[dk];
            if (v) entry[dk] = v;
          });
          if (Object.keys(entry).length > 0) nextDebt[docId] = entry;
        });
        next.debt = nextDebt;
        setQueueStateLocal(next);
        await setQueueState(next);
        queueMsg = ' (คิวเวรถูกย้อนกลับไปก่อนการจัดครั้งนี้ด้วย)';
      }
    }
    await addNotification(
      `ล้างโควต้าเวรของเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว${queueMsg}`,
      `🗑️ ล้างโควต้าเวรของเดือน ${THAI_MONTHS[month]} ${year + 543} แล้ว${queueMsg}`
    );
    showToast(`ล้างโควต้าเวรแล้ว${queueMsg}`);
  };

  /* ---------- unavailability ---------- */

  // Every unavailability edit below updates local state optimistically for
  // instant feedback, then separately persists via patchCurrentMonth, which
  // re-reads the record fresh right before writing and only ever touches
  // unavailability-related fields. This app has no realtime sync, so a
  // doctor's tab can sit open for a long time while a marketplace trade or
  // another doctor's report lands in the meantime — the old code built its
  // save payload from local React state (masterSchedule, unavailability,
  // etc. all captured at click time), so a stale tab would blindly
  // overwrite whatever had changed elsewhere since page load, silently
  // erasing it. Never touching fields this code has no business changing
  // (masterSchedule, scheduleOverrides, ...) and always re-reading
  // unavailability itself fresh closes that gap.
  const toggleUnavailable = (date) => {
    if (!currentDoctorId) return;
    const docId = currentDoctorId;
    setUnavailability(prev => {
      const mine = prev[docId] || [];
      const nextMine = mine.includes(date) ? mine.filter(d => d !== date) : [...mine, date].sort();
      return { ...prev, [docId]: nextMine };
    });
    setUnavailabilityConfirmed(prev => prev.filter(id => id !== docId));
    const nextStale = currentScheduleGenerated ? true : scheduleStale;
    if (nextStale !== scheduleStale) setScheduleStale(nextStale);
    patchCurrentMonth((raw) => {
      const rawUnavail = raw.unavailability || {};
      const mine = rawUnavail[docId] || [];
      const nextMine = mine.includes(date) ? mine.filter(d => d !== date) : [...mine, date].sort();
      return {
        unavailability: { ...rawUnavail, [docId]: nextMine },
        unavailabilityConfirmed: (raw.unavailabilityConfirmed || []).filter(id => id !== docId),
        scheduleStale: raw.currentScheduleGenerated ? true : (raw.scheduleStale || false),
      };
    });
  };

  // docId defaults to the logged-in user's own id (used by the doctor-role
  // tab); the admin panel passes whichever doctor is currently selected,
  // which covers the admin's own record too since admin has no separate
  // "doctor" view of themselves.
  const toggleUnavailabilityConfirmed = (docId = currentDoctorId) => {
    if (!docId) return;
    setUnavailabilityConfirmed(prev => prev.includes(docId) ? prev.filter(id => id !== docId) : [...prev, docId]);
    patchCurrentMonth((raw) => {
      const prevConfirmed = raw.unavailabilityConfirmed || [];
      const next = prevConfirmed.includes(docId) ? prevConfirmed.filter(id => id !== docId) : [...prevConfirmed, docId];
      return { unavailabilityConfirmed: next };
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
        return { ...prev, [docId]: merged };
      });
      setUnavailabilityConfirmed(prev => prev.filter(id => id !== docId));
      patchCurrentMonth((raw) => {
        const rawUnavail = raw.unavailability || {};
        const merged = [...new Set([...(rawUnavail[docId]||[]), ...toAdd])].sort();
        return {
          unavailability: { ...rawUnavail, [docId]: merged },
          unavailabilityConfirmed: (raw.unavailabilityConfirmed || []).filter(id => id !== docId),
          scheduleStale: raw.currentScheduleGenerated ? true : (raw.scheduleStale || false),
        };
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

  const deleteRecurringRule = (docId, dow) => {
    if (!docId) return;
    setQueueStateLocal(prev => {
      if (!prev) return prev;
      const rule = (prev.recurringRules || []).find(r => r.docId === docId && r.dow === dow);
      const rules = (prev.recurringRules || []).filter(r => !(r.docId === docId && r.dow === dow));
      const next = { ...prev, recurringRules: rules };
      setQueueState(next).catch(console.error);
      // Also remove any dates this rule injected into the currently viewed
      // month — leaves manually-toggled dates untouched.
      if (rule) {
        const toRemove = new Set(recurringDatesForMonth(year, month, dow, rule.occurrences));
        setUnavailability(prevU => ({ ...prevU, [docId]: (prevU[docId] || []).filter(d => !toRemove.has(d)) }));
        patchCurrentMonth((raw) => {
          const rawUnavail = raw.unavailability || {};
          const mine = (rawUnavail[docId] || []).filter(d => !toRemove.has(d));
          return { unavailability: { ...rawUnavail, [docId]: mine } };
        });
      }
      return next;
    });
    showToast('ลบวันไม่สะดวกประจำแล้ว');
  };

  const clearUnavailableMonth = (docId) => {
    if (!docId) return;
    setUnavailability(prev => ({ ...prev, [docId]: [] }));
    setUnavailabilityConfirmed(prev => prev.filter(id => id !== docId));
    patchCurrentMonth((raw) => {
      const rawUnavail = raw.unavailability || {};
      return {
        unavailability: { ...rawUnavail, [docId]: [] },
        unavailabilityConfirmed: (raw.unavailabilityConfirmed || []).filter(id => id !== docId),
        scheduleStale: raw.currentScheduleGenerated ? true : (raw.scheduleStale || false),
      };
    });
    showToast('ล้างวันไม่สะดวกเดือนนี้แล้ว');
  };

  // Clear all unavailability for a doctor in the current month


  /* ---------- marketplace ---------- */

  const myAssignedDates = (docId) => {
    const source = currentScheduleGenerated ? effectiveSchedule : masterSchedule;
    return Object.entries(source).filter(([, id]) => id === docId).map(([d]) => d).sort();
  };

  // actorId defaults to the logged-in user, but admin can act on behalf of
  // another doctor from the marketplace tab (e.g. someone forgot to post
  // their own sell/accept) — in that case actorId is that doctor's id.
  const createPost = (date, type, targetDoctorId, requestedDate, actorId = currentDoctorId) => {
    const post = { id: genId(), date, posterId: actorId, type, targetDoctorId: targetDoctorId || null, requestedDate: requestedDate || null, status: 'open', takerId: null, createdAt: new Date().toISOString() };
    setMarketplace(prev => {
      const next = [post, ...prev];
      storageSet('marketplace', next);
      return next;
    });
    const posterName = getDoctor(actorId)?.name;
    const targetName = targetDoctorId ? getDoctor(targetDoctorId)?.name : null;
    const desc = type === 'swap'
      ? `แลกเวร: วันที่ ${formatDisplayDate(date)} (ของ${posterName}) ↔ วันที่ ${formatDisplayDate(requestedDate)} (ของ ${targetName})`
      : (targetName ? `ขายเวรวันที่ ${formatDisplayDate(date)} ให้ ${targetName} โดยเฉพาะ` : `ขายเวรวันที่ ${formatDisplayDate(date)} (เปิดให้ทุกคน)`);
    addNotification(`${posterName} ลงประกาศ${desc}`, `📢 ${posterName} ลงประกาศ${desc}`);
    showToast('ลงประกาศเรียบร้อย');
  };

  const createBulkSell = (dates, targetDoctorId, actorId = currentDoctorId) => {
    if (dates.length === 0) { showToast('ไม่มีเวรที่ยังไม่ได้ลงขาย'); return; }
    const newPosts = dates.map(date => ({
      id: genId(), date, posterId: actorId, type: 'sell', targetDoctorId: targetDoctorId || null, requestedDate: null,
      status: 'open', takerId: null, createdAt: new Date().toISOString(),
    }));
    setMarketplace(prev => {
      const next = [...newPosts, ...prev];
      storageSet('marketplace', next);
      return next;
    });
    const posterName = getDoctor(actorId)?.name;
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

  const declinePost = (post, actorId = currentDoctorId) => {
    setMarketplace(prev => {
      const next = prev.map(p => p.id === post.id ? { ...p, status: 'cancelled' } : p);
      storageSet('marketplace', next);
      return next;
    });
    const takerName = getDoctor(actorId)?.name;
    addNotification(`${takerName} ปฏิเสธคำขอวันที่ ${formatDisplayDate(post.date)} จาก ${getDoctor(post.posterId)?.name}`, `❌ ${takerName} ปฏิเสธคำขอวันที่ ${formatDisplayDate(post.date)}`);
  };

  // Completed trades (sell or swap) update the MASTER schedule directly —
  // a permanent reallocation. The current schedule always regenerates from
  // the master afterward (see buildCurrentSchedule), which is what keeps
  // everyone's weekday/holiday totals matching the new quota AND guarantees
  // nobody ends up with adjacent duty days — so no validation is needed here.
  //
  // A sell is a QUOTA transfer only — the master schedule tracks how many
  // shifts of each type a doctor owes this month, not which exact date they
  // actually work (buildCurrentSchedule decides that later, around each
  // doctor's own declared availability). So a sell never touches either
  // side's unavailability: the buyer may well have declared themselves
  // unavailable for that specific date already (still true after buying —
  // buying it doesn't mean committing to work that exact day), and the
  // seller isn't newly unavailable for it either (nothing stops them
  // legitimately landing on it later as part of their own reduced quota).
  // A swap is different — it directly exchanges two specific CURRENT-
  // schedule dates (see the 'swap' branch below), a real commitment to work
  // (or not work) those exact days, so marking both sides' availability
  // there is correct.
  //
  // Always reads each affected month FRESH from Supabase right before
  // writing, rather than trusting local React state — this app has no
  // realtime sync, so a tab left open can be stale relative to what another
  // device already did (this was the root cause of accepted trades silently
  // "disappearing": two devices both mutated their own stale in-memory copy
  // and the last write won). Reading fresh per month also means this works
  // for a post from ANY month, not just whichever one is currently loaded —
  // which is what makes cross-month swap and multi-month bulk-buy possible.
  const acceptPost = async (post, actorId = currentDoctorId) => {
    if (post.targetDoctorId && post.targetDoctorId !== actorId) return;
    const takerId = actorId;
    const dateMonthKey = monthKey(Number(post.date.slice(0, 4)), Number(post.date.slice(5, 7)) - 1);
    const loadedMonthKey = monthKey(year, month);
    const nextStaleFor = (raw) => raw.currentScheduleGenerated ? true : (raw.scheduleStale || false);

    // Fetches ONE month's stored data fresh, applies `mutate(raw) -> patch`,
    // persists the merged result, and mirrors it into React state too if
    // it's the currently loaded month (so the visible UI updates immediately).
    const patchMonth = async (mk, mutate) => {
      const raw = (await getMonthData(mk)) || {};
      const patch = mutate(raw);
      await setMonthData(mk, { ...raw, ...patch });
      if (mk === loadedMonthKey) {
        if ('masterSchedule' in patch) setMasterSchedule(patch.masterSchedule);
        if ('scheduleOverrides' in patch) setScheduleOverrides(patch.scheduleOverrides);
        if ('unavailability' in patch) setUnavailability(patch.unavailability);
        if ('unavailabilityConfirmed' in patch) setUnavailabilityConfirmed(patch.unavailabilityConfirmed);
        if ('scheduleStale' in patch) setScheduleStale(patch.scheduleStale);
      }
    };

    if (post.type === 'swap') {
      // SWAP: only touches the current schedule (overrides) and unavailability.
      // Master schedule and quotas stay exactly the same. By requesting the
      // swap, the poster implicitly can't do their original date anymore →
      // mark it unavailable for them automatically (and likewise for the
      // taker on the date they're giving up in exchange).
      const reqMonthKey = monthKey(Number(post.requestedDate.slice(0, 4)), Number(post.requestedDate.slice(5, 7)) - 1);
      const sameMonth = reqMonthKey === dateMonthKey;

      if (sameMonth) {
        await patchMonth(dateMonthKey, (raw) => {
          const nextOverrides = { ...(raw.scheduleOverrides || {}), [post.date]: takerId, [post.requestedDate]: post.posterId };
          const nextUnavail = { ...(raw.unavailability || {}) };
          const addUnavail = (docId, date) => {
            const list = nextUnavail[docId] || [];
            if (!list.includes(date)) nextUnavail[docId] = [...list, date].sort();
          };
          addUnavail(post.posterId, post.date);
          addUnavail(takerId, post.requestedDate);
          return { scheduleOverrides: nextOverrides, unavailability: nextUnavail, scheduleStale: nextStaleFor(raw) };
        });
      } else {
        await patchMonth(dateMonthKey, (raw) => {
          const nextOverrides = { ...(raw.scheduleOverrides || {}), [post.date]: takerId };
          const nextUnavail = { ...(raw.unavailability || {}) };
          const list = nextUnavail[post.posterId] || [];
          if (!list.includes(post.date)) nextUnavail[post.posterId] = [...list, post.date].sort();
          return { scheduleOverrides: nextOverrides, unavailability: nextUnavail, scheduleStale: nextStaleFor(raw) };
        });
        await patchMonth(reqMonthKey, (raw) => {
          const nextOverrides = { ...(raw.scheduleOverrides || {}), [post.requestedDate]: post.posterId };
          const nextUnavail = { ...(raw.unavailability || {}) };
          const list = nextUnavail[takerId] || [];
          if (!list.includes(post.requestedDate)) nextUnavail[takerId] = [...list, post.requestedDate].sort();
          return { scheduleOverrides: nextOverrides, unavailability: nextUnavail, scheduleStale: nextStaleFor(raw) };
        });
      }
    } else {
      // SELL: a master-schedule ownership change ONLY — this is a quota
      // transaction, not a real-world commitment. The master schedule
      // decides how many shifts of each type a doctor owes this month, not
      // which exact date they'll actually work; that's decided later, when
      // "จัดเวร" builds the current schedule around each doctor's own
      // declared availability. So this never touches either party's
      // unavailability: buying a date doesn't mean the buyer is now
      // available to work that exact date (they may well have declared
      // themselves unavailable for it specifically — that stays true), and
      // selling one doesn't make the seller newly unavailable for it either
      // (nothing stops them legitimately landing on it anyway as part of
      // their own remaining quota).
      await patchMonth(dateMonthKey, (raw) => {
        const nextMaster = { ...(raw.masterSchedule || raw.schedule || {}), [post.date]: takerId };
        return { masterSchedule: nextMaster, scheduleStale: nextStaleFor(raw) };
      });
    }

    setMarketplace(prevMarket => {
      const nextMarket = prevMarket.map(p => p.id === post.id ? { ...p, status: 'completed', takerId } : p);
      storageSet('marketplace', nextMarket);
      return nextMarket;
    });

    const posterName = getDoctor(post.posterId)?.name;
    const takerName = getDoctor(takerId)?.name;
    const msg = post.type === 'sell'
      ? `${takerName} รับเวรวันที่ ${formatDisplayDate(post.date)} ต่อจาก ${posterName} (ปรับโควต้าเวรแล้ว)`
      : `${takerName} แลกเวรกับ ${posterName}: วันที่ ${formatDisplayDate(post.date)} ↔ วันที่ ${formatDisplayDate(post.requestedDate)} (ปรับตารางเวรแล้ว โควต้าเวรไม่เปลี่ยน)`;
    addNotification(msg, `✅ ${msg}`);
    showToast('บันทึกเรียบร้อย');
  };

  // Admin-only: yank a marketplace entry back out, whether it's still open
  // or already completed. For a completed sell/swap this actually reverses
  // the schedule effects acceptPost applied (restores master ownership /
  // schedule overrides and best-effort undoes the auto-marked unavailability)
  // rather than just flipping the status — otherwise "cancelling" a done
  // trade would leave the shift with the wrong owner. For a still-open post
  // there's nothing to undo on the schedule side (acceptPost never ran), so
  // this reduces to the same fresh-read-then-write plumbing as acceptPost —
  // reused here so admin can't clobber a concurrent device's write either.
  const reversePost = async (post) => {
    const loadedMonthKey = monthKey(year, month);
    const nextStaleFor = (raw) => raw.currentScheduleGenerated ? true : (raw.scheduleStale || false);
    const patchMonth = async (mk, mutate) => {
      const raw = (await getMonthData(mk)) || {};
      const patch = mutate(raw);
      await setMonthData(mk, { ...raw, ...patch });
      if (mk === loadedMonthKey) {
        if ('masterSchedule' in patch) setMasterSchedule(patch.masterSchedule);
        if ('scheduleOverrides' in patch) setScheduleOverrides(patch.scheduleOverrides);
        if ('unavailability' in patch) setUnavailability(patch.unavailability);
        if ('scheduleStale' in patch) setScheduleStale(patch.scheduleStale);
      }
    };

    const dateMonthKey = monthKey(Number(post.date.slice(0, 4)), Number(post.date.slice(5, 7)) - 1);
    if (post.type === 'sell') {
      // Mirrors acceptPost's SELL branch: only the master-schedule
      // ownership changed, so only that needs reverting. Neither side's
      // unavailability was touched by the sale, so there's nothing to undo
      // there either — doing so would risk stripping a declaration either
      // doctor made for an unrelated reason, before or after the sale.
      await patchMonth(dateMonthKey, (raw) => {
        const nextMaster = { ...(raw.masterSchedule || raw.schedule || {}), [post.date]: post.posterId };
        return { masterSchedule: nextMaster, scheduleStale: nextStaleFor(raw) };
      });
    } else {
      const reqMonthKey = monthKey(Number(post.requestedDate.slice(0, 4)), Number(post.requestedDate.slice(5, 7)) - 1);
      await patchMonth(dateMonthKey, (raw) => {
        const nextOverrides = { ...(raw.scheduleOverrides || {}) };
        delete nextOverrides[post.date];
        const nextUnavail = { ...(raw.unavailability || {}) };
        if (nextUnavail[post.posterId]) nextUnavail[post.posterId] = nextUnavail[post.posterId].filter(d => d !== post.date);
        return { scheduleOverrides: nextOverrides, unavailability: nextUnavail, scheduleStale: nextStaleFor(raw) };
      });
      await patchMonth(reqMonthKey, (raw) => {
        const nextOverrides = { ...(raw.scheduleOverrides || {}) };
        delete nextOverrides[post.requestedDate];
        const nextUnavail = { ...(raw.unavailability || {}) };
        if (nextUnavail[post.takerId]) nextUnavail[post.takerId] = nextUnavail[post.takerId].filter(d => d !== post.requestedDate);
        return { scheduleOverrides: nextOverrides, unavailability: nextUnavail, scheduleStale: nextStaleFor(raw) };
      });
    }

    const wasCompleted = post.status === 'completed';
    setMarketplace(prevMarket => {
      const next = prevMarket.map(p => p.id === post.id ? { ...p, status: 'cancelled' } : p);
      storageSet('marketplace', next);
      return next;
    });

    const posterName = getDoctor(post.posterId)?.name;
    const takerName = post.takerId ? getDoctor(post.takerId)?.name : null;
    const kind = post.type === 'sell' ? 'ขายเวร' : 'แลกเวร';
    const msg = wasCompleted
      ? `แอดมินยกเลิกรายการ${kind}วันที่ ${formatDisplayDate(post.date)} ที่สำเร็จแล้วระหว่าง ${posterName} กับ ${takerName} (คืนตารางเดิม)`
      : `แอดมินยกเลิกประกาศ${kind}วันที่ ${formatDisplayDate(post.date)} ของ ${posterName}`;
    addNotification(msg, `↩️ ${msg}`);
    showToast('ยกเลิกรายการและคืนตารางเดิมแล้ว');
  };

  /* ---------- render helpers ---------- */

  if (!currentUser) return <LoginScreen doctors={doctors} onLogin={(doc) => { setCurrentUser(doc); if (doc.role === 'admin') setActiveTab('overview'); refreshData(); }} />;
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
  // Doctors always see their own shifts ringed; admin can instead pick
  // someone else to recheck via the picker in the current/master tabs.
  const highlightDoctorId = (role === 'admin' && recheckDoctorId) ? recheckDoctorId : currentDoctorId;
  const doctorsWithShifts = activeDoctors.filter(d => (masterUsage[d.id]?.weekday || 0) + (masterUsage[d.id]?.holiday || 0) > 0);
  const pendingConfirmDocs = doctorsWithShifts.filter(d => !unavailabilityConfirmed.includes(d.id));

  const tabs = role === 'admin'
    ? [
        { id: 'overview', label: 'ภาพรวม', icon: LayoutDashboard },
        { id: 'current', label: 'ตารางเวร', icon: CalendarCheck },
        { id: 'master', label: 'โควต้าเวร', icon: CalendarIcon },
        { id: 'config', label: 'ตั้งค่า', icon: Settings },
        { id: 'unavailable', label: 'วันไม่สะดวก', icon: UserCircle, badge: (hasMasterData && !currentScheduleGenerated) ? pendingConfirmDocs.length : 0 },
        { id: 'marketplace', label: 'ตลาดแลกเปลี่ยน', icon: Repeat, badge: marketplace.filter(p => p.status === 'open' && (p.posterId === currentDoctorId || p.targetDoctorId === currentDoctorId)).length },
        { id: 'notifications', label: 'แจ้งเตือน', icon: Bell },
      ]
    : [
        { id: 'current', label: 'ตารางเวร', icon: CalendarCheck },
        { id: 'master', label: 'โควต้าเวร', icon: CalendarIcon },
        { id: 'unavailable', label: 'แจ้งวันไม่สะดวก', icon: UserCircle, badge: (currentDoctorId && !currentScheduleGenerated && !unavailabilityConfirmed.includes(currentDoctorId)) ? 1 : 0 },
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
                  <button
                    onClick={() => setShowBatchGen(true)}
                    className="flex items-center gap-1.5 text-sm font-medium text-indigo-700 border border-indigo-200 hover:bg-indigo-50 px-3 py-2 rounded-lg transition-colors"
                  >
                    <CalendarCheck size={14} /> จัดหลายเดือน
                  </button>
                  {currentScheduleGenerated && (
                    <button
                      onClick={() => setConfirmState({ type: 'clear-current' })}
                      className="flex items-center gap-1.5 text-sm font-medium text-slate-500 hover:text-red-600 hover:bg-red-50 px-3 py-2 rounded-lg transition-colors"
                    >
                      <RotateCcw size={14} /> ล้างตารางเวร
                    </button>
                  )}
                </div>
              )}
              {currentScheduleGenerated && (
                <button
                  onClick={saveCurrentScheduleImage}
                  disabled={savingScheduleImage}
                  className="flex items-center gap-1.5 text-sm font-medium text-teal-700 hover:bg-teal-50 disabled:opacity-50 disabled:cursor-wait px-3 py-2 rounded-lg transition-colors border border-teal-200"
                >
                  <Download size={14} /> {savingScheduleImage ? 'กำลังบันทึก...' : 'บันทึกตารางเวร'}
                </button>
              )}
              {role === 'admin' && currentScheduleGenerated && (
                <button
                  onClick={exportDutyRosterDocx}
                  disabled={exportingDocx}
                  className="flex items-center gap-1.5 text-sm font-medium text-indigo-700 hover:bg-indigo-50 disabled:opacity-50 disabled:cursor-wait px-3 py-2 rounded-lg transition-colors border border-indigo-200"
                >
                  <FileText size={14} /> {exportingDocx ? 'กำลังสร้างไฟล์...' : 'Export to DOCX'}
                </button>
              )}
            </div>
            <p className="text-xs text-slate-400 mb-3 flex items-center gap-1"><Info size={12} /> ตารางนี้จะจัดก็ต่อเมื่อแอดมินกด "จัดเวร" เท่านั้น (ไม่จัดอัตโนมัติ) เพื่อให้รอโควต้าเวรและการแจ้งไม่สะดวกนิ่งก่อน — จำนวนเวรวันธรรมดา/วันหยุดของแต่ละคนจะเท่ากับโควต้าเวร และไม่มีใครอยู่เวรติดกัน{role === 'admin' ? ' คลิกวันที่เพื่อแก้ไขเฉพาะจุดเองได้หลังจัดแล้ว' : ''}</p>

            {role === 'admin' && currentScheduleGenerated && scheduleStale && (
              <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2 mb-4 flex items-start gap-2">
                <Info size={14} className="mt-0.5 shrink-0" />
                มีการเปลี่ยนแปลงข้อมูล (โควต้าเวร แจ้งไม่สะดวก หรือขาย/แลกเวร) หลังจากจัดเวรครั้งล่าสุด กด "จัดเวร" อีกครั้งเพื่อให้ตารางเวรตรงกับข้อมูลล่าสุด
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
              <EmptyState icon={CalendarIcon} title="ยังไม่มีโควต้าเวรของเดือนนี้" hint="ไปที่แท็บ 'โควต้าเวร' เพื่อกำหนดโควต้าเวรก่อน ตารางเวรจะคำนวณจากโควต้าเวรเท่านั้น" />
            ) : !currentScheduleGenerated ? (
              <EmptyState icon={Shuffle} title="ยังไม่ได้จัดตารางเวร" hint={role === 'admin' ? 'รอให้ทุกคนแจ้งวันไม่สะดวกและยืนยันครบ (ดูสถานะด้านบน) แล้วกดปุ่ม "จัดเวร" เพื่อเริ่มจัด' : 'รอแอดมินกดจัดเวร'} />
            ) : (
              <>
                {role === 'admin' && (
                  <DoctorHighlightPicker doctors={activeDoctors} allDoctors={doctors} selectedId={recheckDoctorId} onSelect={setRecheckDoctorId} />
                )}
                <div ref={currentScheduleCaptureRef} className="bg-white p-3 rounded-xl">
                  <p className="font-display font-semibold text-slate-800 text-base mb-2 text-center">ตารางเวร — {THAI_MONTHS[month]} {year + 543}</p>
                  <CalendarGrid
                    year={year} month={month} scheduleData={effectiveSchedule}
                    editable={role === 'admin'} onAssign={manualAssignCurrent}
                    allDoctors={doctors} selectableDoctors={activeDoctors}
                    holidaySet={holidaySet} unavailability={unavailability} marketplace={marketplace}
                    compareTo={savingScheduleImage ? null : currentSchedule} highlightDoctorId={highlightDoctorId}
                    violationDates={[...liveViolations]}
                    hideUnavailableCount={savingScheduleImage}
                  />
                </div>
                {role === 'admin' && <p className="text-xs text-slate-400 mt-2 flex items-center gap-1"><Info size={12} /> แถบสีฟ้าด้านซ้ายของช่อง = วันนี้ถูกแก้ไขเฉพาะจุดด้วยมือ · ⚠️ = วันนี้ไม่ตรงเงื่อนไข (อยู่เวรวันที่แจ้งไม่สะดวก หรืออยู่เวรติดกัน) ไม่ว่าจะมาจากตอนจัดเวรหรือแก้ไขเองทีหลัง</p>}

                {liveViolations.size > 0 && (
                  <div className="mt-3 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 flex items-start gap-2">
                    <span className="text-amber-500 text-base shrink-0">⚠️</span>
                    <div>
                      <p className="text-xs font-medium text-amber-800 mb-0.5">มี {liveViolations.size} วันที่ไม่ตรงเงื่อนไขอยู่ตอนนี้</p>
                      <p className="text-xs text-amber-700">แต่ละวันด้านล่างนี้ มีคนอยู่เวรวันที่ตัวเองแจ้งไม่สะดวก หรืออยู่เวรติดกัน — ไม่ว่าจะเกิดจากตอนจัดเวรอัตโนมัติหรือแก้ไขเองทีหลังก็ตาม ลองแก้ไขเฉพาะจุดเองด้านบน: {[...liveViolations].sort().map(d => formatDisplayDate(d)).join(', ')}</p>
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

                <UsageTable title="จำนวนเวรที่จัดแล้วเดือนนี้ (จัดจริง / โควต้าเวรล่าสุด)" doctors={activeDoctors} usage={currentUsage} original={masterUsage} />
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
                    <CalendarCheck size={14} /> จัดโควต้าเวร
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
                      <RotateCcw size={14} /> ล้างโควต้าเวร
                    </button>
                  )}
                </div>
              )}
            </div>
            <p className="text-xs text-slate-400 mb-3 flex items-center gap-1"><Info size={12} /> นี่คือโควต้าเวรที่คุณกำหนดเอง (ไฟล์ Excel: คอลัมน์ A วันที่ · B ชื่อแพทย์) การขาย/แลกเวรที่สำเร็จแล้วจะถูกปรับเข้าที่นี่โดยอัตโนมัติเพื่ออัปเดตโควตาล่าสุด{role === 'admin' ? ' — คลิกวันที่เพื่อแก้ไขได้โดยตรง' : ''}</p>

            {activeDoctors.length === 0 ? (
              <EmptyState icon={Users} title="ยังไม่มีแพทย์ที่อยู่เวรเดือนนี้" hint="ไปที่แท็บ 'ตั้งค่า' เพื่อเพิ่ม/เลือกแพทย์ที่อยู่เวรเดือนนี้ก่อน" />
            ) : (
              <>
                {role === 'admin' && (
                  <DoctorHighlightPicker doctors={activeDoctors} allDoctors={doctors} selectedId={recheckDoctorId} onSelect={setRecheckDoctorId} />
                )}
                <CalendarGrid
                  year={year} month={month} scheduleData={masterSchedule}
                  editable={role === 'admin'} onAssign={manualAssignMaster}
                  allDoctors={doctors} selectableDoctors={activeDoctors}
                  holidaySet={holidaySet} unavailability={unavailability} marketplace={null}
                  compareTo={null} highlightDoctorId={highlightDoctorId} originalData={masterOriginal}
                />
                <p className="text-xs text-slate-400 mt-3 flex items-center gap-1"><Info size={12} /> ชื่อสีเทาขีดฆ่า = เจ้าของเวรเดิมก่อนขายเวร (ไม่ปรากฏสำหรับการแลกเวร) · ชื่อด้านบน = เจ้าของเวรปัจจุบัน</p>
                <UsageTable title="จำนวนเวรที่จัดแล้วเดือนนี้ (ปัจจุบัน(เดิมก่อนขายเวร))" doctors={activeDoctors} usage={masterUsage} original={masterOriginalUsage} />
                {hasMasterData && queueState && (
                  <QueueRunOrderSummary year={year} month={month} doctors={doctors} masterOriginal={masterOriginal} holidays={holidays} />
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
                <EmptyState icon={Users} title="ยังไม่มีรายชื่อแพทย์" hint="เพิ่มเองด้านบน หรืออัปโหลดโควต้าเวรซึ่งจะเพิ่มรายชื่อให้อัตโนมัติ" />
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
              <p className="text-xs text-slate-400 mb-2">เสาร์-อาทิตย์นับเป็นวันหยุดโดยอัตโนมัติแล้ว เพิ่มเฉพาะวันหยุดนักขัตฤกษ์อื่น ๆ — คลิกวันที่เพื่อเพิ่ม/ลบได้เลย เลื่อนเดือนแล้วคลิกได้หลายวันติดกันโดยไม่ต้องเลื่อนกลับมาที่เดือนปัจจุบันใหม่ทุกครั้ง</p>
              <HolidayPicker
                year={year} month={month} holidays={holidays}
                onToggle={(date) => updateHolidays(holidays.includes(date) ? holidays.filter(h => h !== date) : [...holidays, date].sort())}
              />
              <div className="flex flex-wrap gap-1.5 mt-2">
                {holidays.map(h => (<span key={h} className="flex items-center gap-1 bg-rose-50 text-rose-600 text-xs font-mono px-2 py-1 rounded-full">{formatDisplayDate(h)}<button onClick={() => updateHolidays(holidays.filter(x => x !== h))}><X size={12} /></button></span>))}
              </div>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-xl px-4 py-3">
              <p className="text-sm text-slate-600 flex items-start gap-2"><Info size={14} className="mt-0.5 shrink-0" /> ระยะห่างระหว่างเวรเป็นกฎตายตัวแล้ว: ห้ามอยู่เวรวันติดกันเสมอในตารางเวร ไม่ต้องตั้งค่าเพิ่ม</p>
            </div>
          </div>
        )}

        {/* UNAVAILABLE TAB (doctor) */}
        {activeTab === 'unavailable' && role === 'doctor' && (
          <div>
            {!currentDoctorId ? <EmptyState icon={UserCircle} title="ยังไม่มีแพทย์ในระบบ" /> : (hasMasterData && myMasterShiftCount === 0) ? (
              <EmptyState icon={UserCircle} title="คุณไม่มีเวรในเดือนนี้" hint={`${getDoctor(currentDoctorId)?.name} ไม่มีเวรอยู่ในโควต้าเวรของเดือน ${THAI_MONTHS[month]} ${year + 543} จึงไม่ต้องแจ้งวันไม่สะดวก`} />
            ) : (
              <>
                {currentScheduleGenerated ? (
                  <div className="bg-slate-100 border border-slate-200 text-slate-600 text-xs rounded-lg px-3 py-2 mb-4 flex items-start gap-2">
                    <Info size={14} className="mt-0.5 shrink-0" />
                    แอดมินจัดตารางเวรของเดือนนี้แล้ว จึงล็อกไม่ให้แจ้ง/แก้ไขวันไม่สะดวกเพิ่มเติม — ถ้าวันที่คุณอยู่เวรดันไม่สะดวกขึ้นมา ให้ลงขาย/แลกเวรที่แท็บ "ตลาดแลกเปลี่ยนเวร" แทน
                  </div>
                ) : (
                  <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2 mb-4 flex items-start gap-2">
                    <Info size={14} className="mt-0.5 shrink-0" />
                    {hasMasterData
                      ? 'แจ้งวันไม่สะดวกล่วงหน้าได้เลย แล้วกดยืนยันด้านล่างเมื่อแจ้งครบ แอดมินจะรอให้ทุกคนยืนยันก่อนกดจัดเวร เพื่อให้จำนวนเวรวันธรรมดา/วันหยุดของทุกคนตรงกับโควต้าเวร และไม่มีใครอยู่เวรติดกัน — เมื่อแอดมินกดจัดเวรแล้ว จะล็อกไม่ให้แก้ไขเพิ่มอีก'
                      : 'ยังไม่มีโควต้าเวรของเดือนนี้ — แจ้งวันไม่สะดวกล่วงหน้าได้เลย แล้วกดยืนยันด้านล่างเมื่อแจ้งครบ — จะล็อกไม่ให้แก้ไขเพิ่มเมื่อแอดมินกดจัดเวรเท่านั้น ไม่ใช่ตอนตั้งโควต้าเวร'}
                  </div>
                )}
                <RecurringUnavailablePanel
                  year={year} month={month}
                  onApply={(dow, occ) => applyRecurringUnavailable(currentDoctorId, dow, occ)}
                  rules={(queueState?.recurringRules || []).filter(r => r.docId === currentDoctorId)}
                  onDelete={(dow) => deleteRecurringRule(currentDoctorId, dow)}
                />
                <div className="flex items-center justify-between mb-4"><MonthNav year={year} month={month} onShift={shiftMonth} /></div>
                <p className="text-xs text-slate-400 mb-3 flex items-center gap-1"><Info size={12} /> {currentScheduleGenerated ? 'ดูวันไม่สะดวกที่แจ้งไว้' : 'คลิกวันที่เพื่อแจ้ง/ยกเลิกการแจ้งไม่สะดวก'} ({getDoctor(currentDoctorId)?.name}) · <span className="inline-block w-2.5 h-2.5 rounded-sm bg-indigo-50 border border-indigo-300" /> ไม่สะดวกประจำ (จาก rule) · <span className="inline-block w-2.5 h-2.5 rounded-sm bg-red-50 border border-red-300" /> จิ้มเลือกเอง</p>
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
                        <button key={date} disabled={currentScheduleGenerated} onClick={() => toggleUnavailable(date)} className={`rounded-lg border p-2 min-h-[56px] text-left transition-colors ${recurring ? 'bg-indigo-50 border-indigo-300' : marked ? 'bg-red-50 border-red-300' : type === 'holiday' ? 'bg-rose-100 border-rose-200 hover:border-teal-300' : 'bg-white border-slate-200 hover:border-teal-300'} ${onDuty ? 'ring-2 ring-offset-1 ring-teal-500' : ''} ${currentScheduleGenerated ? 'cursor-default opacity-80' : ''}`}>
                          <div className="font-mono text-[11px] text-slate-500">{Number(date.slice(-2))}</div>
                          {onDuty && <div className="text-[9px] text-teal-600 font-medium mt-0.5">อยู่เวร</div>}
                          {marked && <div className={`text-[10px] font-medium mt-0.5 ${recurring ? 'text-indigo-500' : 'text-red-500'}`}>{recurring ? 'ไม่สะดวกประจำ' : 'ไม่สะดวก'}</div>}
                        </button>
                      );
                    });
                  })()}
                </div>
                {!currentScheduleGenerated && (
                  <>
                    <div className="mt-3 flex justify-end">
                      <button onClick={() => clearUnavailableMonth(currentDoctorId)} className="text-xs text-slate-400 hover:text-red-500 hover:bg-red-50 px-2 py-1 rounded-lg transition-colors flex items-center gap-1">
                        <span className="text-sm leading-none">🗑</span> ล้างวันไม่สะดวกเดือนนี้
                      </button>
                    </div>
                    <div className={`mt-2 rounded-xl border-2 px-4 py-3.5 ${unavailabilityConfirmed.includes(currentDoctorId) ? 'bg-emerald-50 border-emerald-300' : 'bg-amber-50 border-amber-400 animate-pulse'}`}>
                      <div className="flex items-center justify-between gap-3 flex-wrap">
                        <p className="text-sm text-slate-700">
                          {unavailabilityConfirmed.includes(currentDoctorId)
                            ? <span className="text-emerald-700 font-semibold flex items-center gap-1.5"><Check size={16} /> คุณยืนยันแล้วว่าแจ้งวันไม่สะดวกครบสำหรับเดือนนี้</span>
                            : <span className="text-amber-900 font-semibold">⚠️ แจ้งวันไม่สะดวกครบแล้วหรือยัง? อย่าลืมกดยืนยัน!</span>}
                        </p>
                        <button onClick={() => toggleUnavailabilityConfirmed(currentDoctorId)} className={`shrink-0 text-sm font-semibold px-4 py-2 rounded-lg transition-colors ${unavailabilityConfirmed.includes(currentDoctorId) ? 'text-slate-500 hover:bg-slate-100 bg-white border border-slate-200' : 'bg-amber-500 hover:bg-amber-600 text-white shadow-sm'}`}>
                          {unavailabilityConfirmed.includes(currentDoctorId) ? 'ยกเลิกการยืนยัน' : 'ยืนยันว่าแจ้งครบแล้ว'}
                        </button>
                      </div>
                    </div>
                  </>
                )}
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
                recurringRules={queueState?.recurringRules} onDeleteRecurring={deleteRecurringRule}
                onToggle={(docId, date) => {
                  setUnavailability(prev => {
                    const mine = prev[docId] || [];
                    const next = mine.includes(date) ? mine.filter(d => d !== date) : [...mine, date].sort();
                    return { ...prev, [docId]: next };
                  });
                  setUnavailabilityConfirmed(prev => prev.filter(id => id !== docId));
                  patchCurrentMonth((raw) => {
                    const rawUnavail = raw.unavailability || {};
                    const mine = rawUnavail[docId] || [];
                    const next = mine.includes(date) ? mine.filter(d => d !== date) : [...mine, date].sort();
                    return {
                      unavailability: { ...rawUnavail, [docId]: next },
                      unavailabilityConfirmed: (raw.unavailabilityConfirmed || []).filter(id => id !== docId),
                    };
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
            unavailability={unavailability} effectiveSchedule={effectiveSchedule} masterSchedule={masterSchedule}
            year={year} month={month} onShiftMonth={shiftMonth} showToast={showToast}
            createPost={createPost} createBulkSell={createBulkSell} cancelPost={cancelPost} declinePost={declinePost} acceptPost={acceptPost}
            reversePost={reversePost}
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
        body={`ระบบจะจัดตารางเวรทั้งเดือนจากโควต้าเวรล่าสุด (จำนวนเวรวันธรรมดา/วันหยุดของแต่ละคนเท่ากับโควต้าเวร และรับประกันว่าไม่มีใครอยู่เวรติดกัน) การแก้ไขเฉพาะจุดที่เคยทำไว้ในตารางเวรเดือน ${THAI_MONTHS[month]} ${year + 543} จะถูกล้างไปด้วย — การขาย/แลกเวรที่สำเร็จแล้วจะไม่ถูกยกเลิก เพราะถูกบันทึกลงโควต้าเวรไปแล้ว${confirmState?.pendingCount > 0 ? `\n\n⚠️ ยังมี ${confirmState.pendingCount} คนที่ยังไม่ยืนยันว่าแจ้งวันไม่สะดวกครบ ต้องการจัดเวรเลยหรือรอก่อน?` : ''}`}
        confirmLabel={confirmState?.pendingCount > 0 ? 'จัดเวรเลย' : 'จัดเวร'}
        danger={confirmState?.pendingCount > 0}
        onCancel={() => setConfirmState(null)}
        onConfirm={() => { setConfirmState(null); generateCurrentSchedule(); }}
      />

      <ConfirmModal
        open={!!pendingAdjacentAssign}
        title="จัดเวรติดกัน?"
        body={pendingAdjacentAssign ? `การจัดแบบนี้จะทำให้ ${getDoctor(pendingAdjacentAssign.docId)?.name} อยู่เวรวันที่ ${formatDisplayDate(pendingAdjacentAssign.date)} ติดกับอีกวันที่ตัวเองอยู่เวรอยู่แล้ว ซึ่งปกติระบบจะไม่จัดให้ — ยืนยันว่าต้องการแก้ไขแบบนี้จริงหรือไม่?` : ''}
        confirmLabel="ยืนยัน อยู่เวรติดกันได้"
        danger
        onCancel={() => setPendingAdjacentAssign(null)}
        onConfirm={() => { applyManualAssignCurrent(pendingAdjacentAssign.date, pendingAdjacentAssign.docId); setPendingAdjacentAssign(null); }}
      />

      <ConfirmModal
        open={confirmState?.type === 'clear-current'}
        title="ล้างตารางเวร?"
        body={`ตารางเวรของเดือน ${THAI_MONTHS[month]} ${year + 543} จะกลับไปเป็นค่าเริ่มต้น (ไม่ปรากฏตาราง) จนกว่าจะกดจัดเวรใหม่ การแก้ไขเฉพาะจุดที่เคยทำไว้จะหายไปด้วย — การขาย/แลกเวรที่สำเร็จแล้วจะไม่ถูกยกเลิก เพราะถูกบันทึกลงโควต้าเวรไปแล้ว`}
        confirmLabel="ล้างตาราง"
        danger
        onCancel={() => setConfirmState(null)}
        onConfirm={() => { setConfirmState(null); resetCurrentSchedule(); }}
      />

      <ConfirmModal
        open={confirmState?.type === 'clear-master'}
        title="ล้างโควต้าเวร?"
        body={`โควต้าเวรทั้งหมดของเดือน ${THAI_MONTHS[month]} ${year + 543} จะถูกล้าง (รวมถึงตารางเวรที่คำนวณจากมันด้วย) รายชื่อแพทย์และวันไม่สะดวกที่แจ้งไว้จะไม่หายไป — ใช้เมื่อต้องการเริ่มจัดโควต้าเวรใหม่ทั้งหมดสำหรับเดือนนี้`}
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
          // Prefer THIS month's own "before it was first generated" snapshot
          // over the raw global queueState — otherwise re-opening this to
          // regenerate an already-generated month would start from wherever
          // that generation itself left the pointers (i.e. the following
          // month's territory) instead of resuming from the prior month.
          queueState={monthQueueSnapshot || queueState}
          onConfirm={handleMasterGenConfirm}
          onClose={() => setShowMasterGen(false)}
        />
      )}

      {showBatchGen && (
        <BatchGenerateModal
          year={year} month={month} doctors={doctors}
          running={batchGenerating}
          onClose={() => { if (!batchGenerating) setShowBatchGen(false); }}
          onRun={async (sy, sm, ey, em) => {
            setBatchGenerating(true);
            try {
              const res = await generateCurrentScheduleBatch(sy, sm, ey, em);
              const totalViolations = res.perMonth.reduce((s, r) => s + r.violations, 0);
              const debtCount = Object.keys(res.finalDebt).length;
              const msg = `จัดตารางเวรหลายเดือนสำเร็จ (${res.monthsList.length} เดือน)${totalViolations ? ` มี ${totalViolations} วันจัดให้ตรงเงื่อนไขไม่ได้` : ''}${debtCount ? ` — ยังมี ${debtCount} คนชดเชยไม่ครบภายในช่วงนี้` : ''}`;
              await addNotification(msg, `🔀 ${msg}`);
              return res;
            } finally {
              setBatchGenerating(false);
            }
          }}
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

function MarketplaceTab({ role, currentDoctorId, doctors, getDoctor, marketplace, myAssignedDates, holidaySet, currentScheduleGenerated, unavailability, effectiveSchedule, masterSchedule, year, month, onShiftMonth, showToast, createPost, createBulkSell, cancelPost, declinePost, acceptPost, reversePost }) {
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
  const [reverseTarget, setReverseTarget] = useState(null);
  const [buyAllConfirm, setBuyAllConfirm] = useState(false);
  const [buyingAll, setBuyingAll] = useState(false);
  // Which month to pick the REQUESTED (incoming) date from — defaults to
  // swapDate's own month, but browsable forward to any future month so a
  // shift can be swapped against a later month's schedule. {year, month} or
  // null before a swapDate is chosen.
  const [swapTargetYM, setSwapTargetYM] = useState(null);
  const [targetMonthData, setTargetMonthData] = useState(null);
  const targetIsLoadedMonth = swapTargetYM && swapTargetYM.year === year && swapTargetYM.month === month;

  useEffect(() => {
    if (!swapDate) { setSwapTargetYM(null); return; }
    const [y, m] = swapDate.split('-').map(Number);
    setSwapTargetYM({ year: y, month: m - 1 });
  }, [swapDate]);

  useEffect(() => {
    if (!swapTargetYM || targetIsLoadedMonth) { setTargetMonthData(null); return; }
    let cancelled = false;
    getMonthData(monthKey(swapTargetYM.year, swapTargetYM.month)).then(data => {
      if (!cancelled) setTargetMonthData(data || {});
    });
    return () => { cancelled = true; };
  }, [swapTargetYM, targetIsLoadedMonth]);

  // The loaded (source) month's real assignments — current schedule if
  // generated, else fall back to its master schedule (same fallback
  // myAssignedDates already uses for "my dates"), so swapping still works
  // for a month whose current schedule hasn't been generated yet.
  const sourceEffectiveOrMaster = currentScheduleGenerated ? effectiveSchedule : masterSchedule;

  // Effective schedule for whichever month the requested date is being
  // picked from — the loaded month's own data if that's the same month,
  // otherwise fetched on demand: current schedule if that month's already
  // been generated, else its master schedule.
  const targetEffectiveSchedule = !swapTargetYM ? {}
    : targetIsLoadedMonth ? sourceEffectiveOrMaster
    : !targetMonthData ? {}
    : targetMonthData.currentScheduleGenerated
      ? { ...(targetMonthData.currentSchedule || {}), ...(targetMonthData.scheduleOverrides || {}) }
      : (targetMonthData.masterSchedule || targetMonthData.schedule || {});
  const targetScheduleGenerated = !swapTargetYM ? false
    : targetIsLoadedMonth ? currentScheduleGenerated
    : !!targetMonthData?.currentScheduleGenerated;
  const targetStillLoading = !!swapTargetYM && !targetIsLoadedMonth && !targetMonthData;
  // Merged view spanning the source month (swapDate lives here) and the
  // target month (swapRequestedDate lives here) — adjacency is plain
  // calendar-date arithmetic, so a merged lookup correctly catches a
  // conflict right at a month boundary (e.g. Aug 31 / Sep 1) too.
  const swapScheduleMerged = { ...sourceEffectiveOrMaster, ...targetEffectiveSchedule };
  // Admin can act on behalf of any doctor (e.g. they forgot to post their own
  // sell, or forgot to accept one) — defaults to admin's own id, meaning
  // nothing changes unless admin explicitly picks someone else.
  const [actingAsId, setActingAsId] = useState(currentDoctorId);
  const effectiveDoctorId = role === 'admin' ? (actingAsId || currentDoctorId) : currentDoctorId;
  const actingAsSomeoneElse = role === 'admin' && effectiveDoctorId !== currentDoctorId;

  const myDates = effectiveDoctorId ? myAssignedDates(effectiveDoctorId) : [];
  const otherDoctors = doctors.filter(d => d.id !== effectiveDoctorId);
  const sellPosts = marketplace.filter(p => p.type === 'sell' && p.status === 'open');
  const swapPosts = marketplace.filter(p => p.type === 'swap' && p.status === 'open');
  // Admin sees the full log (uncapped); doctors just see the recent few.
  const history = role === 'admin' ? marketplace.filter(p => p.status !== 'open') : marketplace.filter(p => p.status !== 'open').slice(0, 8);

  // "ซื้อทุกเวร" spans every open post across all months (not just the one
  // currently loaded) — acceptPost fetches each post's own month fresh from
  // storage, so this works regardless of which month is on screen.
  const buyableSellPosts = sellPosts.filter(p =>
    p.posterId !== effectiveDoctorId && (!p.targetDoctorId || p.targetDoctorId === effectiveDoctorId)
  );
  const buyAllSellPosts = async () => {
    // Sequential, not parallel: each acceptPost reads its month fresh right
    // before writing, so two posts landing in the same month must run one
    // after another to see each other's changes instead of racing.
    for (const p of buyableSellPosts) {
      await acceptPost(p, effectiveDoctorId);
    }
  };

  const submitSell = () => {
    if (!sellDate) return;
    const alreadyPosted = new Set(marketplace.filter(p => p.type === 'sell' && p.status === 'open' && p.posterId === effectiveDoctorId).map(p => p.date));
    if (sellDate === '__ALL__') {
      createBulkSell(myDates.filter(d => !alreadyPosted.has(d)), sellTarget || null, effectiveDoctorId);
    } else if (sellDate === '__WEEKDAY__') {
      createBulkSell(myDates.filter(d => !alreadyPosted.has(d) && dayType(d, holidaySet) === 'weekday'), sellTarget || null, effectiveDoctorId);
    } else if (sellDate === '__HOLIDAY__') {
      createBulkSell(myDates.filter(d => !alreadyPosted.has(d) && dayType(d, holidaySet) === 'holiday'), sellTarget || null, effectiveDoctorId);
    } else {
      createPost(sellDate, 'sell', sellTarget || null, null, effectiveDoctorId);
    }
    setSellDate(''); setSellTarget('');
  };

  // All dates in the TARGET month's schedule owned by someone other than the
  // acting doctor, same day-type as the chosen date, that won't make the
  // OTHER doctor adjacent if they take this date over. Each entry includes
  // the owner id. Candidates can come from a later month than swapDate's own
  // (cross-month swap) — the merged view keeps adjacency correct right at a
  // month boundary.
  const candidateSwapDates = swapDate && swapTargetYM && !targetStillLoading
    ? (() => {
        const type = dayType(swapDate, holidaySet);
        return Object.entries(targetEffectiveSchedule)
          .filter(([d, ownerId]) =>
            ownerId && ownerId !== effectiveDoctorId &&
            dayType(d, holidaySet) === type &&
            // target doctor won't be adjacent taking over our date
            !hasAdjacentAssignment({ ...swapScheduleMerged, [swapDate]: ownerId, [d]: effectiveDoctorId }, swapDate, ownerId)
          )
          .map(([d, ownerId]) => ({ d, ownerId }))
          .sort((a, b) => a.d.localeCompare(b.d));
      })()
    : [];

  // Once requester picks a target date, derive the targetDoctorId from it
  const swapDerivedTarget = swapRequestedDate
    ? (targetEffectiveSchedule[swapRequestedDate] || null)
    : null;

  // Check if the requester would end up adjacent after taking the target's date
  const swapSelfAdjacent = swapDate && swapRequestedDate && effectiveDoctorId && swapDerivedTarget
    ? hasAdjacentAssignment({ ...swapScheduleMerged, [swapDate]: swapDerivedTarget, [swapRequestedDate]: effectiveDoctorId }, swapRequestedDate, effectiveDoctorId)
    : false;

  const submitSwap = () => {
    if (!swapDate || !swapRequestedDate || !swapDerivedTarget) return;
    if (swapSelfAdjacent && !swapSelfAdjacentConfirmed) { setSwapSelfAdjacentConfirmed(true); return; }
    createPost(swapDate, 'swap', swapDerivedTarget, swapRequestedDate, effectiveDoctorId);
    setSwapDate(''); setSwapTarget(''); setSwapRequestedDate(''); setSwapSelfAdjacentConfirmed(false);
  };

  const canAct = (p) => !p.targetDoctorId || p.targetDoctorId === effectiveDoctorId;

  // Selections reference specific dates, which only make sense for the month
  // they came from — clear them whenever the loaded month or the acting-as
  // doctor changes so a stale date/selection can't linger in the form.
  useEffect(() => {
    setSellDate(''); setSellTarget('');
    setSwapDate(''); setSwapTarget(''); setSwapRequestedDate(''); setSwapSelfAdjacentConfirmed(false);
  }, [year, month, actingAsId]);

  return (
    <div className="max-w-2xl space-y-8">
      <MonthNav year={year} month={month} onShift={onShiftMonth} />

      {role === 'admin' && (
        <div className={`flex items-center gap-2 flex-wrap rounded-lg border px-3 py-2 ${actingAsSomeoneElse ? 'bg-amber-50 border-amber-300' : 'bg-slate-50 border-slate-200'}`}>
          <span className="text-xs font-medium text-slate-600">ทำรายการในนามของ:</span>
          <select value={actingAsId || ''} onChange={(e) => setActingAsId(e.target.value || currentDoctorId)} className="border border-slate-200 rounded-lg px-2 py-1 text-sm bg-white">
            {doctors.map(d => <option key={d.id} value={d.id}>{d.name}{d.id === currentDoctorId ? ' (ตัวฉันเอง)' : ''}</option>)}
          </select>
          {actingAsSomeoneElse && (
            <span className="text-xs text-amber-700 font-medium">⚠️ กำลังทำรายการแทน {getDoctor(effectiveDoctorId)?.name} — ลงขาย/แลก/รับเวรด้านล่างจะนับเป็นของคนนี้</span>
          )}
        </div>
      )}

      {!currentScheduleGenerated && (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg px-3 py-2 flex items-start gap-2">
          <Info size={14} className="mt-0.5 shrink-0" />
          ยังไม่ได้จัดตารางเวรของเดือนนี้ วันที่แสดงด้านล่างจึงอิงจากโควต้าเวรไปก่อน หลังแอดมินกดจัดเวรแล้ว วันที่จะเปลี่ยนเป็นวันที่อยู่เวรจริง
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
          myDates.length === 0 ? <p className="text-sm text-slate-400 mb-2">{actingAsSomeoneElse ? `${getDoctor(effectiveDoctorId)?.name}ยังไม่มีเวรที่จัดไว้ในเดือนนี้` : 'คุณยังไม่มีเวรที่จัดไว้ในเดือนนี้'}</p> : (
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
            p.posterId === effectiveDoctorId || !p.targetDoctorId || p.targetDoctorId === effectiveDoctorId
          );
          if (visiblePosts.length === 0) return <EmptyState icon={Tag} title="ยังไม่มีประกาศขายเวรสำหรับคุณ" hint="ประกาศขายเวรที่ระบุแพทย์คนอื่นโดยเฉพาะจะมองไม่เห็น" />;
          return (
            <div className="space-y-2">
              {visiblePosts.map(p => {
                const poster = getDoctor(p.posterId);
                const target = p.targetDoctorId ? getDoctor(p.targetDoctorId) : null;
                const isMine = p.posterId === effectiveDoctorId;
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
                      {role === 'admin' && (
                        <button onClick={() => setReverseTarget(p)} className="text-xs font-medium text-red-500 hover:text-red-700 px-2 py-1 border border-red-200 rounded-lg">ยกเลิก (admin)</button>
                      )}
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
        <p className="text-xs text-slate-400 mb-3">แลกได้เฉพาะวันธรรมดากับวันธรรมดา หรือวันหยุดกับวันหยุด — เลือกวันของฉันก่อน แล้วเลือกวันที่ต้องการแลกมาได้จากเดือนเดียวกันหรือเดือนอื่นในอนาคตก็ได้ (ถ้าเดือนนั้นยังไม่ได้จัดตารางเวร จะอิงจากโควต้าเวรไปก่อน)</p>

        {(role === 'doctor' || role === 'admin') && (
          myDates.length === 0 ? (
            <p className="text-sm text-slate-400 mb-2">คุณยังไม่มีเวรในเดือนนี้</p>
          ) : (
            <div className="flex flex-wrap items-center gap-2 mb-4">
              <select value={swapDate} onChange={(e) => { setSwapDate(e.target.value); setSwapRequestedDate(''); setSwapSelfAdjacentConfirmed(false); }} className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm font-mono">
                <option value="">1. เลือกวันที่ของฉันที่จะแลกออก</option>
                {myDates.filter(d => !(unavailability[effectiveDoctorId] || []).includes(d)).map(d => <option key={d} value={d}>{formatDisplayDate(d)} ({dayTypeLabel(d, holidaySet)})</option>)}
              </select>

              {swapDate && swapTargetYM && (() => {
                const now = new Date();
                const atFloor = swapTargetYM.year === now.getFullYear() && swapTargetYM.month === now.getMonth();
                const shiftTarget = (delta) => {
                  let m = swapTargetYM.month + delta, y = swapTargetYM.year;
                  if (m < 0) { m = 11; y -= 1; } else if (m > 11) { m = 0; y += 1; }
                  setSwapTargetYM({ year: y, month: m });
                  setSwapRequestedDate(''); setSwapSelfAdjacentConfirmed(false);
                };
                return (
                  <div className="w-full flex items-center gap-2 mt-1">
                    <span className="text-xs text-slate-500">2. เดือนที่จะแลกมา:</span>
                    <button type="button" disabled={atFloor} onClick={() => shiftTarget(-1)} className="w-6 h-6 flex items-center justify-center border border-slate-200 rounded hover:bg-slate-50 disabled:opacity-30 disabled:cursor-not-allowed"><ChevronLeft size={14} /></button>
                    <span className="text-xs font-medium text-slate-700 min-w-[90px] text-center">{THAI_MONTHS[swapTargetYM.month]} {swapTargetYM.year + 543}</span>
                    <button type="button" onClick={() => shiftTarget(1)} className="w-6 h-6 flex items-center justify-center border border-slate-200 rounded hover:bg-slate-50"><ChevronRight size={14} /></button>
                  </div>
                );
              })()}

              {swapDate && (
                targetStillLoading ? (
                  <span className="text-xs text-slate-400">กำลังโหลดตารางเดือนนั้น…</span>
                ) : candidateSwapDates.length === 0 ? (
                  <span className="text-xs text-red-500">ไม่มีวันประเภทเดียวกันในเดือนนี้ที่แลกได้โดยไม่ทำให้ใครติดกัน — ลองเปลี่ยนเดือนดู</span>
                ) : (
                  <div className="w-full mt-1">
                    <p className="text-xs text-slate-500 mb-2">3. เลือกวันที่ต้องการแลกมา (เฉพาะ{dayType(swapDate, holidaySet) === 'holiday' ? 'วันหยุด' : 'วันธรรมดา'}){!targetScheduleGenerated ? ' — เดือนนี้ยังไม่ได้จัดตารางเวร อิงจากโควต้าเวร' : ''}</p>
                    <SwapCalendar
                      year={swapTargetYM.year} month={swapTargetYM.month}
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
                  ⚠️ ถ้าแลกเวรนี้ <strong>{actingAsSomeoneElse ? getDoctor(effectiveDoctorId)?.name : 'คุณ'}จะอยู่เวรติดกัน 2 วัน</strong> ยืนยันว่ารับได้ใช่ไหม? กดปุ่มด้านล่างอีกครั้งเพื่อยืนยัน
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
              const isMine = p.posterId === effectiveDoctorId;
              const isTarget = p.targetDoctorId === effectiveDoctorId;
              return (
                <div key={p.id} className="border border-slate-200 bg-white rounded-xl p-3 flex items-center justify-between flex-wrap gap-2">
                  <div>
                    <p className="text-sm font-medium text-slate-700">{poster?.name} ขอแลกวันที่ <span className="font-mono">{formatDisplayDate(p.date)}</span> ({dayTypeLabel(p.date, holidaySet)}) กับวันที่ <span className="font-mono">{formatDisplayDate(p.requestedDate)}</span> ({dayTypeLabel(p.requestedDate, holidaySet)}) ของ {target?.name}</p>
                  </div>
                  <div className="flex items-center gap-2">
                    {isMine ? (
                      <button onClick={() => cancelPost(p.id)} className="text-xs font-medium text-slate-500 hover:text-red-600 px-2 py-1">ยกเลิกคำขอ</button>
                    ) : isTarget ? (
                      <>
                        <button onClick={() => declinePost(p, effectiveDoctorId)} className="text-xs font-medium text-slate-500 hover:text-red-600 px-2 py-1">ปฏิเสธ</button>
                        <button onClick={() => setAcceptTarget(p)} className="flex items-center gap-1 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-medium px-3 py-1.5 rounded-lg"><Check size={13} /> ยืนยัน</button>
                      </>
                    ) : (
                      <span className="text-xs text-slate-400">รอ {target?.name} ตอบรับ</span>
                    )}
                    {role === 'admin' && (
                      <button onClick={() => setReverseTarget(p)} className="text-xs font-medium text-red-500 hover:text-red-700 px-2 py-1 border border-red-200 rounded-lg">ยกเลิก (admin)</button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {history.length > 0 && (
        <div>
          <p className="font-display font-semibold text-slate-800 mb-2 text-sm">{role === 'admin' ? `Log การซื้อ/ขาย/แลกเวรทั้งหมด (${history.length})` : 'ประวัติล่าสุด'}</p>
          <div className={`space-y-1.5 ${role === 'admin' ? 'max-h-96 overflow-y-auto pr-1' : ''}`}>
            {history.map(p => (
              <div key={p.id} className="flex items-center justify-between gap-2 text-xs text-slate-400">
                <p>{formatDisplayDate(p.date)} · {p.type === 'sell' ? 'ขายเวร' : 'แลกเวร'} · {getDoctor(p.posterId)?.name} → {p.status === 'completed' ? (getDoctor(p.takerId)?.name || '-') : 'ยกเลิก/ปฏิเสธ'}</p>
                {role === 'admin' && p.status === 'completed' && (
                  <button onClick={() => setReverseTarget(p)} className="shrink-0 text-red-500 hover:text-red-700 font-medium px-2 py-0.5 border border-red-200 rounded">ยกเลิก (admin)</button>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      <ConfirmModal
        open={!!acceptTarget}
        title={acceptTarget?.type === 'swap' ? 'ยืนยันแลกเวรนี้?' : 'ยืนยันรับเวรนี้?'}
        body={acceptTarget ? (() => {
          const actorLabel = actingAsSomeoneElse ? getDoctor(effectiveDoctorId)?.name : 'คุณ';
          return acceptTarget.type === 'sell'
            ? `${actorLabel}จะรับเวรวันที่ ${formatDisplayDate(acceptTarget.date)} แทน ${getDoctor(acceptTarget.posterId)?.name} — โควต้าเวรจะอัปเดตทันที`
            : `${actorLabel}จะมอบเวรวันที่ ${formatDisplayDate(acceptTarget.requestedDate)} ให้ ${getDoctor(acceptTarget.posterId)?.name} และรับเวรวันที่ ${formatDisplayDate(acceptTarget.date)} มาแทน — ตารางเวรจะอัปเดตทันที (โควต้าเวรไม่เปลี่ยน)`;
        })() : ''}
        confirmLabel="ยืนยัน"
        onCancel={() => setAcceptTarget(null)}
        onConfirm={() => { acceptPost(acceptTarget, effectiveDoctorId); setAcceptTarget(null); }}
      />

      <ConfirmModal
        open={!!reverseTarget}
        title={reverseTarget?.status === 'completed' ? 'ยกเลิกรายการที่สำเร็จแล้ว?' : 'ยกเลิกประกาศนี้?'}
        body={reverseTarget ? (() => {
          const posterName = getDoctor(reverseTarget.posterId)?.name;
          const takerName = reverseTarget.takerId ? getDoctor(reverseTarget.takerId)?.name : null;
          if (reverseTarget.status !== 'completed') {
            return `จะยกเลิกประกาศ${reverseTarget.type === 'sell' ? 'ขายเวร' : 'แลกเวร'}วันที่ ${formatDisplayDate(reverseTarget.date)} ของ ${posterName}`;
          }
          return reverseTarget.type === 'sell'
            ? `รายการนี้สำเร็จไปแล้ว: ${takerName} รับเวรวันที่ ${formatDisplayDate(reverseTarget.date)} ต่อจาก ${posterName} — การยกเลิกจะคืนเวรวันนี้ให้ ${posterName} ในโควต้าเวรทันที (โปรดแจ้งทั้งสองฝ่ายให้ทราบด้วย)`
            : `รายการนี้สำเร็จไปแล้ว: ${posterName} แลกกับ ${takerName} วันที่ ${formatDisplayDate(reverseTarget.date)} ↔ ${formatDisplayDate(reverseTarget.requestedDate)} — การยกเลิกจะคืนตารางเวรของทั้งสองวันกลับเป็นเดิมทันที (โปรดแจ้งทั้งสองฝ่ายให้ทราบด้วย)`;
        })() : ''}
        confirmLabel="ยกเลิก/คืนตาราง"
        danger
        onCancel={() => setReverseTarget(null)}
        onConfirm={() => { reversePost(reverseTarget); setReverseTarget(null); }}
      />

      <ConfirmModal
        open={buyAllConfirm}
        title="ซื้อทุกเวรที่เปิดอยู่?"
        body={`คุณจะรับเวรทั้งหมด ${buyableSellPosts.length} วันที่เปิดขายอยู่ (ทุกเดือน): ${buyableSellPosts.map(p => formatDisplayDate(p.date)).join(', ')} — โควต้าเวรจะอัปเดตทันทีสำหรับทุกวัน`}
        confirmLabel={buyingAll ? 'กำลังซื้อ…' : 'ซื้อทั้งหมด'}
        onCancel={() => setBuyAllConfirm(false)}
        onConfirm={async () => {
          if (buyingAll) return;
          setBuyingAll(true);
          const count = buyableSellPosts.length;
          await buyAllSellPosts();
          setBuyingAll(false);
          setBuyAllConfirm(false);
          showToast(`ซื้อเวรสำเร็จ ${count} รายการ`);
        }}
      />
    </div>
  );
}

/* ---------------------------------- admin unavailable panel ---------------------------------- */

/* ---------------------------------- recurring unavailability panel ---------------------------------- */

const DOW_LABELS = ['อา','จ','อ','พ','พฤ','ศ','ส'];
const DOW_FULL   = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];

function RecurringUnavailablePanel({ year, month, onApply, rules = [], onDelete }) {
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

  const describeRule = (r) => {
    const occLabel = (!r.occurrences || r.occurrences.length === 0) ? 'ทุกครั้ง' : `ครั้งที่ ${r.occurrences.join(', ')}`;
    return `ทุกวัน${DOW_FULL[r.dow]} (${occLabel})`;
  };

  const rulesList = rules.length > 0 && (
    <div className="mb-3 space-y-1.5">
      {rules.map(r => (
        <div key={r.dow} className="flex items-center justify-between gap-2 bg-indigo-50 border border-indigo-200 rounded-lg px-3 py-2">
          <span className="text-xs text-indigo-700 flex items-center gap-1.5"><span className="text-sm leading-none">🔁</span> {describeRule(r)}</span>
          <button onClick={() => onDelete?.(r.dow)} className="text-xs text-slate-400 hover:text-red-600 hover:bg-red-50 px-2 py-1 rounded transition-colors flex items-center gap-1 shrink-0">
            <Trash2 size={12} /> ลบ
          </button>
        </div>
      ))}
    </div>
  );

  if (!open) return (
    <div className="mb-4">
      {rulesList}
      <button onClick={() => setOpen(true)} className="flex items-center gap-1.5 text-xs font-medium text-indigo-600 hover:bg-indigo-50 border border-indigo-200 px-3 py-1.5 rounded-lg transition-colors">
        <span className="text-base leading-none">🔁</span> เพิ่มวันไม่สะดวกประจำ (fix schedule)
      </button>
    </div>
  );

  return (
    <div className="mb-4 border border-indigo-200 rounded-xl bg-indigo-50/40 p-3">
      {rulesList}
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

function AdminUnavailablePanel({ year, month, doctors, allDoctors, unavailability, effectiveSchedule, holidaySet, masterSchedule, defaultDocId, unavailabilityConfirmed, onToggleConfirmed, isRecurringUnavailable, recurringRules, onToggle, onApplyRecurring, onDeleteRecurring, onClearMonth }) {
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
              rules={(recurringRules || []).filter(r => r.docId === selectedDocId)}
              onDelete={(dow) => onDeleteRecurring(selectedDocId, dow)}
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
                    {inMaster && <div className="text-[9px] text-sky-600 font-medium">โควต้าเวร</div>}
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
