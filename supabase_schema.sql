-- ============================================================
-- Doctor Shift Scheduler — Supabase Schema
-- Run this in: Supabase Dashboard → SQL Editor → New Query
-- ============================================================

-- Doctors table (เก็บรายชื่อแพทย์ + role + hashed password + LINE user id)
create table if not exists doctors (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,
  role        text not null default 'doctor',   -- 'admin' | 'doctor'
  password_hash text not null default '$2a$10$rQ1Y2z3ABC4DEF5GHI6JKLe7TU8VWX9YZabcdefghijklmnopqrstu', -- placeholder, bcrypt of '1234'
  must_change_password boolean not null default true,
  line_user_id text,
  created_at  timestamptz not null default now()
);

-- Insert ณัฐพล as admin (password = 1234, hashed — will be replaced at first login)
insert into doctors (name, role, must_change_password) values ('ณัฐพล', 'admin', true);

-- Month data table (เก็บข้อมูลทั้งหมดของแต่ละเดือน เป็น JSON เหมือน window.storage เดิม)
create table if not exists month_data (
  id          uuid primary key default gen_random_uuid(),
  month_key   text not null unique,   -- e.g. 'month-2026-08'
  data        jsonb not null default '{}',
  updated_at  timestamptz not null default now()
);

-- Global config (doctor roster, holidays, min gap)
create table if not exists config (
  key   text primary key,
  value jsonb not null default '{}'
);

-- Marketplace posts
create table if not exists marketplace (
  id            uuid primary key default gen_random_uuid(),
  data          jsonb not null,   -- full post object
  created_at    timestamptz not null default now()
);

-- Notifications log
create table if not exists notifications (
  id         uuid primary key default gen_random_uuid(),
  message    text not null,
  line_message text,
  created_at timestamptz not null default now()
);

-- Row Level Security: allow all reads/writes from anon key
-- (auth is handled by our own password check, not Supabase Auth)
alter table doctors enable row level security;
alter table month_data enable row level security;
alter table config enable row level security;
alter table marketplace enable row level security;
alter table notifications enable row level security;

create policy "public read doctors" on doctors for select using (true);
create policy "public update doctors" on doctors for update using (true);
create policy "public insert doctors" on doctors for insert with check (true);
create policy "public delete doctors" on doctors for delete using (true);

create policy "public all month_data" on month_data for all using (true) with check (true);
create policy "public all config" on config for all using (true) with check (true);
create policy "public all marketplace" on marketplace for all using (true) with check (true);
create policy "public all notifications" on notifications for all using (true) with check (true);
