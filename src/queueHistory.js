import { isoDate, daysInMonth, pad2 } from './dateUtils.js';

export function ltFor(n) { return n <= 2 ? 'h12' : n === 3 ? 'h3' : n === 4 ? 'h4' : 'h5'; }

// Native Date rolls day-of-month over/under into adjacent months correctly
// (e.g. new Date(2026, 11, 32) → 1 Jan 2027), so this works for any offset.
function offsetDate(year, month, dayOffset) {
  const d = new Date(year, month, dayOffset);
  return isoDate(d.getFullYear(), d.getMonth(), d.getDate());
}

// Scans a window padded well before day 1 and after the last day (not just
// [1, n]) so a holiday streak that crosses a month boundary — e.g. 31 Dec –
// 3 Jan — is recognized as ONE continuous group instead of being split into
// a 1-day group in December and a separate 3-day group in January. Only the
// dates that actually fall in [year, month] are returned per group (that's
// all this generation run assigns), but `trueLength` carries the FULL
// cross-boundary streak length so ltFor() classifies it correctly (as a
// 4-day group here, not 1+3).
//
// `isHoliday(date)` must be able to answer for ANY date, not just this
// month's — callers build it from a combination of live-edited state (for
// the month actually being configured) and saved config (for neighboring
// months).
export function detectGroups(year, month, isHoliday) {
  const n = daysInMonth(year, month);
  const PAD = 6; // comfortably longer than any realistic holiday streak
  const rawGroups = [];
  let cur = [];
  for (let d = 1 - PAD; d <= n + PAD; d++) {
    const date = offsetDate(year, month, d);
    if (isHoliday(date)) cur.push(date);
    else if (cur.length) { rawGroups.push(cur); cur = []; }
  }
  if (cur.length) rawGroups.push(cur);

  const monthStart = isoDate(year, month, 1);
  const monthEnd = isoDate(year, month, n);
  return rawGroups
    .map(g => ({ dates: g.filter(d => d >= monthStart && d <= monthEnd), trueLength: g.length }))
    .filter(g => g.dates.length > 0);
}

// Which of the five queue loops each date of a month belongs to.
export function classifyMonthDates(y, m, isHolidayDate) {
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

export const LOOP_KEYS = ['weekday', 'h12', 'h3', 'h4', 'h5'];

// Index in queue `q` (array of doctor ids, duplicates allowed) of the LAST
// doctor in `chainIds` (assignments in chronological order). A name that
// appears twice in the loop (h12 has four of them) can't be placed from
// the name alone, so every element is placed at once: pick the
// combination of positions with the smallest total forward distance
// around the loop, i.e. the fewest skipped turns overall. A run of a few
// consecutive assignments pins that down uniquely.
export function resolveLastIndex(q, chainIds) {
  const len = q.length;
  if (len === 0) return null;
  const positions = chainIds
    .map(id => { const p = []; q.forEach((x, i) => { if (x === id) p.push(i); }); return p; })
    .filter(p => p.length > 0);
  if (positions.length === 0) return null;
  const fwd = (i, j) => ((j - i + len) % len) || len;

  let costs = positions[0].map(() => 0);
  for (let k = 1; k < positions.length; k++) {
    const next = positions[k].map(j => {
      let best = Infinity;
      positions[k - 1].forEach((i, pi) => { best = Math.min(best, costs[pi] + fwd(i, j)); });
      return best;
    });
    costs = next;
  }
  let bestIdx = 0;
  costs.forEach((c, i) => { if (c < costs[bestIdx]) bestIdx = i; });
  return positions[positions.length - 1][bestIdx];
}

const MAX_MONTHS_BACK = 36;
const BATCH = 12;
const MIN_CHAIN = 6;

// Where every loop's queue stood immediately BEFORE (year, month) was
// generated, worked out from the real saved quota of the months before it
// — never from the shared live pointers. Those pointers are one cumulative
// value across all months, so any time a month is cleared and redone
// (or was generated before snapshotting existed) they sit somewhere
// unrelated to "the month before this one". Real history can't drift:
// it uses masterOriginal (the quota as generated, before any later trade
// or manual edit), and the pointer for a loop is simply the slot after
// whoever last held it.
//
// getMonthData(key) → saved month object or null. `queues` maps each loop
// to its resolved id array. A loop with no history at all before this
// month (a rare loop like h5 that simply hasn't come up since records
// began) hasn't been consumed yet, so it sits at the system's starting
// pointer (`defaultState`) — NOT the live pointer, which may already be
// advanced by later months. Its lastDate is only kept if the live value
// is genuinely older than this month (a hand-seeded date from before
// records began).
export async function deriveQueueStateBefore({ year, month, getMonthData, queues, liveState, defaultState, isHolidayDate, debt }) {
  const chains = {};
  LOOP_KEYS.forEach(k => { chains[k] = []; });
  const done = new Set();

  let y = year, m = month - 1;
  if (m < 0) { m = 11; y -= 1; }
  let scanned = 0;
  while (scanned < MAX_MONTHS_BACK && done.size < LOOP_KEYS.length) {
    const batch = [];
    for (let i = 0; i < BATCH && scanned < MAX_MONTHS_BACK; i++, scanned++) {
      batch.push([y, m]);
      m -= 1; if (m < 0) { m = 11; y -= 1; }
    }
    const rows = await Promise.all(batch.map(([by, bm]) => getMonthData(`month-${by}-${pad2(bm + 1)}`)));
    for (let i = 0; i < rows.length && done.size < LOOP_KEYS.length; i++) {
      const data = rows[i];
      if (!data) continue;
      const [by, bm] = batch[i];
      const master = data.masterOriginal || data.masterSchedule || data.schedule || {};
      const dbt = classifyMonthDates(by, bm, isHolidayDate);
      LOOP_KEYS.forEach(key => {
        if (done.has(key)) return;
        const dates = dbt[key].filter(d => master[d]).sort();
        chains[key] = dates.map(d => ({ date: d, id: master[d] })).concat(chains[key]);
        if (chains[key].length >= MIN_CHAIN) done.add(key);
      });
    }
  }

  const state = { ...liveState, lastDate: { ...(liveState.lastDate || {}) }, debt: debt ?? liveState.debt ?? {} };
  const derivedFrom = {};
  const monthStart = isoDate(year, month, 1);
  LOOP_KEYS.forEach(key => {
    const chain = chains[key];
    const q = queues[key];
    const lastIdx = (chain.length && q && q.length) ? resolveLastIndex(q, chain.map(c => c.id)) : null;
    if (lastIdx !== null) {
      state[key] = (lastIdx + 1) % q.length;
      state.lastDate[key] = chain[chain.length - 1].date;
      derivedFrom[key] = chain.length;
      return;
    }
    if (defaultState && defaultState[key] !== undefined) state[key] = defaultState[key];
    const liveLast = (liveState.lastDate || {})[key];
    state.lastDate[key] = liveLast && liveLast < monthStart ? liveLast : null;
  });
  return { state, derivedFrom };
}
