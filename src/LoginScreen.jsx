import { useState } from 'react';
import { CalendarIcon, Eye, EyeOff } from 'lucide-react';
import { verifyPassword } from './storage';

export default function LoginScreen({ doctors, onLogin }) {
  const [doctorId, setDoctorId] = useState('');
  const [password, setPassword] = useState('');
  const [show, setShow] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const activeDoctors = doctors.filter(d => d);

  const handleLogin = async (e) => {
    e.preventDefault();
    if (!doctorId) { setError('กรุณาเลือกชื่อของคุณ'); return; }
    if (!password) { setError('กรุณาใส่รหัสผ่าน'); return; }
    setLoading(true); setError('');
    try {
      const ok = await verifyPassword(doctorId, password);
      if (!ok) { setError('รหัสผ่านไม่ถูกต้อง'); setLoading(false); return; }
      const doc = activeDoctors.find(d => d.id === doctorId);
      onLogin(doc);
    } catch {
      setError('เกิดข้อผิดพลาด กรุณาลองใหม่');
    }
    setLoading(false);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 w-full max-w-sm p-8">
        <div className="flex items-center gap-3 mb-8">
          <div className="w-10 h-10 rounded-xl bg-teal-600 flex items-center justify-center">
            <CalendarIcon size={20} className="text-white" />
          </div>
          <div>
            <p className="font-display font-semibold text-slate-800 text-lg leading-tight">ระบบจัดเวรแพทย์</p>
            <p className="text-xs text-slate-400">กรุณาเข้าสู่ระบบ</p>
          </div>
        </div>

        <form onSubmit={handleLogin} className="space-y-4">
          <div>
            <label className="text-xs font-medium text-slate-600 mb-1 block">ชื่อ</label>
            <select value={doctorId} onChange={e => { setDoctorId(e.target.value); setError(''); }}
              className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:border-teal-400 bg-white">
              <option value="">— เลือกชื่อของคุณ —</option>
              {activeDoctors.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-medium text-slate-600 mb-1 block">รหัสผ่าน</label>
            <div className="relative">
              <input type={show ? 'text' : 'password'} value={password} onChange={e => { setPassword(e.target.value); setError(''); }}
                className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm pr-10 focus:outline-none focus:border-teal-400" placeholder="รหัสผ่านคือเลขว." />
              <button type="button" onClick={() => setShow(v => !v)} className="absolute right-3 top-2.5 text-slate-400">{show ? <EyeOff size={16} /> : <Eye size={16} />}</button>
            </div>
          </div>
          {error && <p className="text-xs text-red-500">{error}</p>}
          <button type="submit" disabled={loading || !doctorId} className="w-full bg-teal-600 hover:bg-teal-700 disabled:bg-slate-200 disabled:cursor-not-allowed text-white font-medium py-2 rounded-lg text-sm transition-colors">
            {loading ? 'กำลังตรวจสอบ…' : 'เข้าสู่ระบบ'}
          </button>
        </form>
      </div>
    </div>
  );
}
