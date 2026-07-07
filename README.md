# ระบบจัดเวรแพทย์ — Deploy Guide

Stack: React + Vite · Supabase (DB + Edge Functions) · Vercel (hosting) · LINE Messaging API
ทุกอย่างฟรี ไม่มีค่าใช้จ่ายเพิ่มเติม

---

## ขั้นตอนที่ 1 — Supabase

1. สมัครที่ https://supabase.com แล้วสร้าง project ใหม่
2. ไปที่ SQL Editor → New Query แล้ว paste เนื้อหาจากไฟล์ supabase_schema.sql แล้วกด Run
3. ไปที่ Project Settings → API แล้วจด:
   - Project URL (เช่น https://xxxx.supabase.co)
   - anon public key
   - service_role key (สำหรับ Edge Function เท่านั้น)

---

## ขั้นตอนที่ 2 — LINE Messaging API

1. ไปที่ https://developers.line.biz → สร้าง Messaging API channel
2. ออก Channel access token (long-lived)
3. ให้แพทย์ทุกคน Add เป็นเพื่อน กับ LINE Official Account นั้น
4. ดึง LINE User ID ของแต่ละคน แล้วกรอกในหน้าตั้งค่าของแอป

ตั้งค่า Edge Function Secret:
  supabase login
  supabase link --project-ref YOUR_PROJECT_REF
  supabase secrets set LINE_CHANNEL_ACCESS_TOKEN=YOUR_TOKEN
  supabase functions deploy send-line --no-verify-jwt

---

## ขั้นตอนที่ 3 — .env

  cp .env.example .env

แก้ไข .env:
  VITE_SUPABASE_URL=https://YOUR_PROJECT.supabase.co
  VITE_SUPABASE_ANON_KEY=YOUR_ANON_KEY

---

## ขั้นตอนที่ 4 — Vercel

1. Push โค้ดขึ้น GitHub (อย่า commit .env)
2. ไปที่ https://vercel.com → Import Git Repository
3. ใส่ Environment Variables: VITE_SUPABASE_URL, VITE_SUPABASE_ANON_KEY
4. กด Deploy

---

## รหัสผ่านเริ่มต้น: 1234 (ทุกคน)
ใช้รหัสผ่านนี้ตลอดไป ไม่มีการบังคับให้เปลี่ยนตอน login ครั้งแรก
Admin เพิ่ม admin คนอื่นได้โดยแก้ role ใน Supabase Table Editor: doctors → role = 'admin'
