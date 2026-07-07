export const pad2 = n => String(n).padStart(2, '0');
export const isoDate = (y, m, d) => `${y}-${pad2(m + 1)}-${pad2(d)}`;
export const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
export const dayType = (dateStr, holidaySet) => {
  const d = new Date(dateStr + 'T00:00:00');
  const dow = d.getDay();
  if (dow === 0 || dow === 6 || holidaySet.has(dateStr)) return 'holiday';
  return 'weekday';
};
export const formatDisplayDate = (iso) => {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  const ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  if (!y || !m || !d) return iso;
  return `${d}-${ABBR[m-1]}-${y + 543}`;
};
