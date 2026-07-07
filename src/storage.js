import { supabase } from './supabase';

// ---------- config ----------

export async function getConfig() {
  const { data } = await supabase.from('config').select('value').eq('key', 'main').single();
  return data?.value ?? { holidays: [] };
}

export async function setConfig(value) {
  await supabase.from('config').upsert({ key: 'main', value }, { onConflict: 'key' });
}

// ---------- month data ----------

export async function getMonthData(monthKey) {
  const { data } = await supabase.from('month_data').select('data').eq('month_key', monthKey).single();
  return data?.data ?? null;
}

export async function setMonthData(monthKey, payload) {
  await supabase
    .from('month_data')
    .upsert({ month_key: monthKey, data: payload, updated_at: new Date().toISOString() }, { onConflict: 'month_key' });
}

// ---------- marketplace ----------
// Stored as a single row in config table (key='marketplace') to avoid
// jsonb-path upsert issues with the marketplace table.

export async function getMarketplace() {
  const { data } = await supabase.from('config').select('value').eq('key', 'marketplace').single();
  return data?.value ?? [];
}

export async function setMarketplace(posts) {
  await supabase.from('config').upsert({ key: 'marketplace', value: posts }, { onConflict: 'key' });
}

// ---------- notifications ----------

export async function getNotifications() {
  const { data } = await supabase
    .from('notifications')
    .select('id, message, line_message, created_at')
    .order('created_at', { ascending: false })
    .limit(100);
  return (data ?? []).map(n => ({ id: n.id, message: n.message, lineMessage: n.line_message, ts: n.created_at }));
}

export async function addNotification(message, lineMessage) {
  await supabase.from('notifications').insert({ message, line_message: lineMessage });
  triggerLineNotification(lineMessage || message);
}

function triggerLineNotification(text) {
  const url = `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/send-line`;
  fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${import.meta.env.VITE_SUPABASE_ANON_KEY}`,
    },
    body: JSON.stringify({ message: text }),
  }).catch(() => {});
}

// ---------- doctors ----------

export async function getDoctors() {
  const { data } = await supabase
    .from('doctors')
    .select('id, name, role, must_change_password, line_user_id')
    .order('created_at');
  return data ?? [];
}

export async function addDoctor(name) {
  const { data } = await supabase
    .from('doctors')
    .insert({ name, role: 'doctor', must_change_password: false,
              password_hash: '$2b$10$DJfi1ytqEAJP7pgjmofp/eStaN5JTRRGHfj70PtWN4.WhLoYPLqsm' })
    .select().single();
  return data;
}

export async function updateDoctor(id, patch) {
  await supabase.from('doctors').update(patch).eq('id', id);
}

export async function deleteDoctor(id) {
  await supabase.from('doctors').delete().eq('id', id);
}

// ---------- queue state (master schedule generator) ----------

const DEFAULT_QUEUE_STATE = {
  // Pointers = index of NEXT person to assign in each loop
  weekday: 0,   // next in WDQ (12-person loop), initial: after ณัฐพล → อารีรัตน์
  h12: 0,       // next in H12Q (15-slot loop), initial: after ขนิษฐา2 → พสิษฐา
  h3: 4,        // next in H3Q (11-person loop), initial: after วัทนี → ธนวรรณ
  h4: 8,        // next in H3Q, separate pointer, initial: after พสิษฐา → ณัฐธิดา
  h5: 6,        // next in H3Q, separate pointer, initial: after ณัชพล → สมิตา
  // Queue order (stored by doctor id so roster changes don't silently break)
  WDQ: null,    // null = use default order from doctors table
  H12Q: null,
  H3Q: null,
  // Per-doctor debt adjustments: { [doctorId]: { WDQ:0, H12Q:0, H3Q:0, H4Q:0, H5Q:0 } }
  debt: {},
  // recurring unavailability rules: [{docId, dow, occurrences}]
  recurringRules: [],
};

export async function getQueueState() {
  const { data } = await supabase.from('config').select('value').eq('key', 'queue_state').single();
  return { ...DEFAULT_QUEUE_STATE, ...(data?.value ?? {}) };
}

export async function setQueueState(value) {
  await supabase.from('config').upsert({ key: 'queue_state', value }, { onConflict: 'key' });
}

// ---------- auth ----------

export async function verifyPassword(doctorId, plainPassword) {
  const { data } = await supabase.from('doctors').select('password_hash').eq('id', doctorId).single();
  if (!data) return false;
  const bcrypt = await import('./bcrypt-browser');
  return bcrypt.compare(plainPassword, data.password_hash);
}

export async function setPassword(doctorId, newPlainPassword) {
  const bcrypt = await import('./bcrypt-browser');
  const hash = await bcrypt.hash(newPlainPassword, 10);
  await supabase.from('doctors').update({ password_hash: hash, must_change_password: false }).eq('id', doctorId);
}
