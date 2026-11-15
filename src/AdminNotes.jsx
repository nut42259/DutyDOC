import { useState } from 'react';
import { StickyNote, ChevronDown, ChevronUp, Plus, Trash2, Check, X } from 'lucide-react';

const THAI_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

function formatNoteDate(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  return `${d.getDate()} ${THAI_MONTHS[d.getMonth()]} ${d.getFullYear() + 543}`;
}

// Read-only list of the notes still open, shown inside the quota generator so
// they're in front of the admin at the moment they're scheduling.
export function OpenNotesReminder({ notes }) {
  const open = (notes || []).filter(n => !n.done);
  if (open.length === 0) return null;
  return (
    <div className="border border-amber-200 bg-amber-50 rounded-xl px-3.5 py-3">
      <p className="text-xs font-medium text-amber-800 mb-1.5 flex items-center gap-1.5"><StickyNote size={13} /> โน้ตช่วยจำ ({open.length})</p>
      <ul className="space-y-1">
        {open.map(n => (
          <li key={n.id} className="text-xs text-slate-700 whitespace-pre-wrap break-words flex gap-1.5">
            <span className="text-amber-500 select-none">•</span><span>{n.text}</span>
          </li>
        ))}
      </ul>
    </div>
  );
}

// Admin-only memo: add, tick off, and delete notes. onAdd/onToggle/onDelete
// return a promise resolving to false if the save failed.
export default function AdminNotes({ notes, onAdd, onToggle, onDelete }) {
  const [expanded, setExpanded] = useState(true);
  const [draft, setDraft] = useState('');
  const [saving, setSaving] = useState(false);
  const [showDone, setShowDone] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  const open = (notes || []).filter(n => !n.done);
  const done = (notes || []).filter(n => n.done);

  const submit = async () => {
    const text = draft.trim();
    if (!text || saving) return;
    setSaving(true);
    const ok = await onAdd(text);
    if (ok !== false) setDraft('');
    setSaving(false);
  };

  const Row = ({ n }) => (
    <li className="flex items-start gap-2 group">
      <button type="button" onClick={() => onToggle(n.id)} title={n.done ? 'ทำเครื่องหมายว่ายังไม่เสร็จ' : 'ทำเครื่องหมายว่าเสร็จแล้ว'}
        className={`mt-0.5 w-4 h-4 shrink-0 rounded border flex items-center justify-center transition-colors ${n.done ? 'bg-amber-500 border-amber-500 text-white' : 'bg-white border-slate-300 hover:border-amber-400'}`}>
        {n.done && <Check size={11} strokeWidth={3} />}
      </button>
      <div className="min-w-0 flex-1">
        <p className={`text-sm whitespace-pre-wrap break-words ${n.done ? 'text-slate-400 line-through' : 'text-slate-700'}`}>{n.text}</p>
        <p className="text-[10px] text-slate-400">{formatNoteDate(n.createdAt)}</p>
      </div>
      {confirmDeleteId === n.id ? (
        <span className="flex items-center gap-1 shrink-0 text-xs">
          <span className="text-slate-500">ลบ?</span>
          <button type="button" onClick={() => { setConfirmDeleteId(null); onDelete(n.id); }} className="p-1 rounded text-red-600 hover:bg-red-50" title="ยืนยันลบ"><Check size={13} /></button>
          <button type="button" onClick={() => setConfirmDeleteId(null)} className="p-1 rounded text-slate-400 hover:bg-slate-100" title="ยกเลิก"><X size={13} /></button>
        </span>
      ) : (
        <button type="button" onClick={() => setConfirmDeleteId(n.id)} title="ลบโน้ต"
          className="p-1 rounded text-slate-300 hover:text-red-600 hover:bg-red-50 shrink-0 transition-colors"><Trash2 size={13} /></button>
      )}
    </li>
  );

  return (
    <div className="border border-amber-200 bg-amber-50/60 rounded-xl mb-4">
      <button type="button" onClick={() => setExpanded(v => !v)} className="w-full flex items-center justify-between gap-2 px-3.5 py-2.5 text-left">
        <span className="flex items-center gap-2 text-sm font-medium text-amber-900">
          <StickyNote size={15} /> โน้ตช่วยจำ
          {open.length > 0 && <span className="text-[10px] font-semibold bg-amber-500 text-white rounded-full px-1.5 py-0.5 leading-none">{open.length}</span>}
          <span className="text-[10px] font-normal text-amber-700/70">เห็นเฉพาะ admin</span>
        </span>
        {expanded ? <ChevronUp size={15} className="text-amber-700" /> : <ChevronDown size={15} className="text-amber-700" />}
      </button>

      {expanded && (
        <div className="px-3.5 pb-3.5 space-y-3">
          <div className="flex items-start gap-2">
            <textarea value={draft} onChange={e => setDraft(e.target.value)} rows={2} maxLength={1000}
              onKeyDown={e => { if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) { e.preventDefault(); submit(); } }}
              placeholder="จดสิ่งที่ต้องจำสำหรับการจัดเวรรอบถัดไป เช่น ใครขอลดเวร ใครต้องชดเชย"
              className="flex-1 min-w-0 text-sm border border-amber-200 bg-white rounded-lg px-3 py-2 resize-y focus:outline-none focus:border-amber-400 focus:ring-4 focus:ring-amber-500/10" />
            <button type="button" onClick={submit} disabled={!draft.trim() || saving}
              className="flex items-center gap-1 shrink-0 bg-amber-500 hover:bg-amber-600 disabled:bg-slate-200 disabled:text-slate-400 disabled:cursor-not-allowed text-white text-sm font-medium px-3 py-2 rounded-lg transition-colors">
              <Plus size={14} /> {saving ? 'กำลังบันทึก…' : 'เพิ่ม'}
            </button>
          </div>

          {notes === null ? (
            <p className="text-xs text-slate-400">กำลังโหลด…</p>
          ) : open.length === 0 ? (
            <p className="text-xs text-slate-400">{done.length > 0 ? 'ไม่มีโน้ตที่ค้างอยู่' : 'ยังไม่มีโน้ต'}</p>
          ) : (
            <ul className="space-y-2.5">{open.map(n => <Row key={n.id} n={n} />)}</ul>
          )}

          {done.length > 0 && (
            <div>
              <button type="button" onClick={() => setShowDone(v => !v)} className="text-xs text-slate-500 hover:text-slate-700 flex items-center gap-1">
                {showDone ? <ChevronUp size={12} /> : <ChevronDown size={12} />} เสร็จแล้ว ({done.length})
              </button>
              {showDone && <ul className="space-y-2.5 mt-2">{done.map(n => <Row key={n.id} n={n} />)}</ul>}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
