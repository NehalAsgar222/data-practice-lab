-- Data Practice Lab: database tables and security rules for Supabase.
-- Run once in Supabase → SQL Editor. Safe to run again (it drops and recreates the policies).

create table if not exists public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  display_name text check (char_length(display_name) <= 60),
  email text,
  photo text,
  provider text,
  created_at timestamptz default now(),
  last_active timestamptz,
  score int default 0,
  theory int default 0,
  practice int default 0,
  per jsonb default '{}'::jsonb,
  show_on_board boolean default true,
  certs jsonb default '{}'::jsonb
);

create table if not exists public.progress (
  id uuid primary key references auth.users (id) on delete cascade,
  keys jsonb not null default '{}'::jsonb,
  updated_at timestamptz default now()
);

create table if not exists public.leaderboard (
  id uuid primary key references auth.users (id) on delete cascade,
  name text not null check (char_length(name) <= 40),
  score int not null check (score between 0 and 2000),
  updated_at timestamptz default now()
);

create table if not exists public.certificates (
  id text primary key check (id ~ '^[A-Z0-9]{6,20}$'),
  uid uuid not null references auth.users (id) on delete cascade,
  name text not null check (char_length(name) <= 80),
  track text not null check (track in ('excel', 'sql', 'pbi')),
  level text not null check (level in ('basic', 'advanced')),
  issued_at timestamptz not null default now()
);

create table if not exists public.admins (
  id uuid primary key references auth.users (id) on delete cascade,
  role text default 'owner'
);

create index if not exists leaderboard_score_idx on public.leaderboard (score desc);

-- Helper: is the signed-in user an admin? (security definer so it can read admins under RLS)
create or replace function public.is_admin() returns boolean
language sql stable security definer set search_path = public
as $$ select exists (select 1 from public.admins where id = auth.uid()) $$;

-- Lets a signed-in person delete their own account and all their data.
create or replace function public.delete_my_account() returns void
language plpgsql security definer set search_path = public
as $$
begin
  if auth.uid() is null then raise exception 'not signed in'; end if;
  delete from auth.users where id = auth.uid();
end $$;
revoke all on function public.delete_my_account() from public, anon;
grant execute on function public.delete_my_account() to authenticated;

alter table public.profiles enable row level security;
alter table public.progress enable row level security;
alter table public.leaderboard enable row level security;
alter table public.certificates enable row level security;
alter table public.admins enable row level security;

drop policy if exists "profiles read own or admin" on public.profiles;
drop policy if exists "profiles insert own" on public.profiles;
drop policy if exists "profiles update own" on public.profiles;
drop policy if exists "profiles delete admin" on public.profiles;
create policy "profiles read own or admin" on public.profiles for select to authenticated using (id = auth.uid() or public.is_admin());
create policy "profiles insert own" on public.profiles for insert to authenticated with check (id = auth.uid());
create policy "profiles update own" on public.profiles for update to authenticated using (id = auth.uid()) with check (id = auth.uid());
create policy "profiles delete admin" on public.profiles for delete to authenticated using (public.is_admin());

drop policy if exists "progress read own or admin" on public.progress;
drop policy if exists "progress write own" on public.progress;
drop policy if exists "progress update own" on public.progress;
drop policy if exists "progress delete own" on public.progress;
create policy "progress read own or admin" on public.progress for select to authenticated using (id = auth.uid() or public.is_admin());
create policy "progress write own" on public.progress for insert to authenticated with check (id = auth.uid());
create policy "progress update own" on public.progress for update to authenticated using (id = auth.uid()) with check (id = auth.uid());
create policy "progress delete own" on public.progress for delete to authenticated using (id = auth.uid() or public.is_admin());

drop policy if exists "board read members" on public.leaderboard;
drop policy if exists "board insert own" on public.leaderboard;
drop policy if exists "board update own" on public.leaderboard;
drop policy if exists "board delete own" on public.leaderboard;
create policy "board read members" on public.leaderboard for select to authenticated using (true);
create policy "board insert own" on public.leaderboard for insert to authenticated with check (id = auth.uid());
create policy "board update own" on public.leaderboard for update to authenticated using (id = auth.uid()) with check (id = auth.uid());
create policy "board delete own" on public.leaderboard for delete to authenticated using (id = auth.uid() or public.is_admin());

-- Certificates: the owner and admins can list them; anyone can check ONE certificate by its ID
-- through verify_certificate() (so the full list of names is never public). Nobody can edit.
drop policy if exists "certs read all" on public.certificates;
drop policy if exists "certs read own or admin" on public.certificates;
drop policy if exists "certs insert own" on public.certificates;
drop policy if exists "certs delete admin" on public.certificates;
create policy "certs read own or admin" on public.certificates for select to authenticated using (uid = auth.uid() or public.is_admin());
create policy "certs insert own" on public.certificates for insert to authenticated with check (uid = auth.uid());
create policy "certs delete admin" on public.certificates for delete to authenticated using (public.is_admin());

-- Admins: each person can see whether they are an admin. Add admins by hand in the Table Editor.
drop policy if exists "admins read self" on public.admins;
create policy "admins read self" on public.admins for select to authenticated using (id = auth.uid());

create or replace function public.verify_certificate(cert_id text)
returns table (id text, name text, track text, level text, issued_at timestamptz)
language sql stable security definer set search_path = public
as $$ select c.id, c.name, c.track, c.level, c.issued_at from public.certificates c where c.id = upper(cert_id) $$;
grant execute on function public.verify_certificate(text) to anon, authenticated;
grant select, insert, update, delete on public.profiles, public.progress, public.leaderboard to authenticated;
grant select, insert, delete on public.certificates to authenticated;
grant select on public.admins to authenticated;
