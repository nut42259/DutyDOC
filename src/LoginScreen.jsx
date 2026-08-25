import { useState } from 'react';
import { CalendarIcon, Eye, EyeOff, ArrowRight, Check } from 'lucide-react';
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
    <div className="min-h-screen bg-slate-50 relative overflow-hidden flex items-center justify-center p-4">
      {/* ambient decoration */}
      <div className="absolute -top-40 -left-36 w-[520px] h-[520px] rounded-full bg-teal-200/30 blur-3xl pointer-events-none" />
      <div className="absolute -bottom-52 -right-40 w-[600px] h-[600px] rounded-full bg-teal-200/20 blur-3xl pointer-events-none" />

      <div className="relative flex flex-col lg:flex-row items-center gap-12 lg:gap-16 max-w-4xl w-full">

        {/* Left: brand panel */}
        <div className="w-full lg:w-[340px] flex flex-col gap-6 px-2">
          <div className="flex items-center gap-3">
            <div className="w-11 h-11 rounded-2xl bg-gradient-to-br from-teal-400 to-teal-600 flex items-center justify-center shadow-lg shadow-teal-600/30">
              <CalendarIcon size={22} className="text-white" />
            </div>
            <div>
              <p className="font-display font-semibold text-teal-900 text-lg leading-tight">DutyDoc</p>
              <p className="text-[10.5px] text-slate-400">กุมารเวชกรรม รพ.บ้านแพ้ว</p>
            </div>
          </div>
          <div>
            <p className="font-display font-semibold text-slate-800 text-3xl leading-tight mb-3">จัดเวรแพทย์</p>
            <p className="text-sm text-slate-500 leading-relaxed">ระบบจัดตารางเวร แจ้งวันไม่สะดวก และซื้อขายและแลกเปลี่ยนเวร</p>
          </div>
          <div className="flex flex-col gap-3 mt-1">
            <div className="flex items-center gap-2.5 text-[13px] text-slate-600">
              <span className="w-5 h-5 rounded-md bg-teal-100 text-teal-700 flex items-center justify-center shrink-0"><Check size={12} strokeWidth={3} /></span>
              จัดตารางเวรให้อัตโนมัติ ไม่มีเวรวันติดกัน
            </div>
            <div className="flex items-center gap-2.5 text-[13px] text-slate-600">
              <span className="w-5 h-5 rounded-md bg-teal-100 text-teal-700 flex items-center justify-center shrink-0"><Check size={12} strokeWidth={3} /></span>
              แจ้งวันไม่สะดวก และแลกเวรกับเพื่อนร่วมทีมได้เลย
            </div>
          </div>
        </div>

        {/* Right: login card */}
        <div className="bg-white rounded-3xl shadow-xl shadow-slate-900/10 border border-slate-200 w-full max-w-sm p-8">
          <p className="font-display font-semibold text-slate-800 text-lg mb-1">เข้าสู่ระบบ</p>
          <p className="text-xs text-slate-400 mb-6">เลือกชื่อของคุณและใส่รหัสผ่าน</p>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="text-xs font-medium text-slate-600 mb-1 block">ชื่อ</label>
              <select value={doctorId} onChange={e => { setDoctorId(e.target.value); setError(''); }}
                className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:border-teal-400 focus:ring-4 focus:ring-teal-500/10 bg-white transition-shadow">
                <option value="">— เลือกชื่อของคุณ —</option>
                {activeDoctors.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
              </select>
            </div>
            <div>
              <label className="text-xs font-medium text-slate-600 mb-1 block">รหัสผ่าน</label>
              <div className="relative">
                <input type={show ? 'text' : 'password'} value={password} onChange={e => { setPassword(e.target.value); setError(''); }}
                  className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm pr-10 focus:outline-none focus:border-teal-400 focus:ring-4 focus:ring-teal-500/10 transition-shadow" placeholder="รหัสผ่านคือเลขว." />
                <button type="button" onClick={() => setShow(v => !v)} className="absolute right-3 top-2.5 text-slate-400">{show ? <EyeOff size={16} /> : <Eye size={16} />}</button>
              </div>
            </div>
            {error && <p className="text-xs text-red-500">{error}</p>}
            <button type="submit" disabled={loading || !doctorId}
              className="w-full bg-gradient-to-br from-teal-500 to-teal-600 hover:from-teal-600 hover:to-teal-700 disabled:from-slate-200 disabled:to-slate-200 disabled:cursor-not-allowed text-white font-semibold py-2.5 rounded-xl text-sm transition-all shadow-lg shadow-teal-600/25 disabled:shadow-none flex items-center justify-center gap-1.5">
              {loading ? 'กำลังตรวจสอบ…' : (<>เข้าสู่ระบบ <ArrowRight size={14} /></>)}
            </button>
            <p className="text-[10.5px] text-slate-300 text-center">รหัสผ่านคือเลขวันเกิดของคุณ 6 หลัก</p>
          </form>
        </div>
      </div>
    </div>
  );
}
