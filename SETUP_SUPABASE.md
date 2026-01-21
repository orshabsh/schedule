# Развёртывание на GitHub Pages + Supabase

Сайт — статический (HTML/JS), поэтому идеально ложится на GitHub Pages. Данные и вход — в Supabase.

## 1) Supabase: таблицы + RLS

Откройте **Supabase → SQL Editor** и выполните:

```sql
-- 1) Профили пользователей (роль)
create table if not exists public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  role text not null check (role in ('admin','deputy','teacher')),
  display_name text
);

alter table public.profiles enable row level security;

-- читать профиль может только сам пользователь
drop policy if exists "profiles_read_own" on public.profiles;
create policy "profiles_read_own" on public.profiles
  for select
  to authenticated
  using (id = auth.uid());

-- обновлять профиль может только сам пользователь
drop policy if exists "profiles_update_own" on public.profiles;
create policy "profiles_update_own" on public.profiles
  for update
  to authenticated
  using (id = auth.uid());


-- 2) Единая строка с данными школы (как ваш db.json)
create table if not exists public.school_data (
  id integer primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.school_data enable row level security;

-- публичное чтение (anon) — расписание видно без входа
drop policy if exists "school_data_public_read" on public.school_data;
create policy "school_data_public_read" on public.school_data
  for select
  to anon, authenticated
  using (true);

-- обновление — только admin/deputy
drop policy if exists "school_data_write_admin_deputy" on public.school_data;
create policy "school_data_write_admin_deputy" on public.school_data
  for update
  to authenticated
  using (
    exists (
      select 1 from public.profiles p
      where p.id = auth.uid() and p.role in ('admin','deputy')
    )
  )
  with check (
    exists (
      select 1 from public.profiles p
      where p.id = auth.uid() and p.role in ('admin','deputy')
    )
  );
```

Дальше:
1. Вставьте начальные данные в `school_data` (можно взять ваш `db.json`).

```sql
insert into public.school_data (id, data)
values (1, '{}'::jsonb)
on conflict (id) do nothing;
```

2. Создайте пользователей в **Authentication → Users**.
3. Для каждого пользователя добавьте строку в `profiles` (id = uuid пользователя):

```sql
insert into public.profiles (id, role, display_name)
values ('02f8ab1d-841d-46f0-9916-abd4543098fa', 'admin', 'Администратор')
on conflict (id) do update set role = excluded.role, display_name = excluded.display_name;
```

Почему это безопасно: anon key можно отдавать в браузер, а доступ регулируется RLS‑политиками в БД. См. Supabase Docs. citeturn0search3turn0search0turn0search12

## 2) Проект: config.js

Скопируйте `config.example.js` → `config.js` и заполните:
- `url`: Supabase Project URL
- `anonKey`: Supabase anon/public key

## 3) GitHub Pages

1. Создайте репозиторий и загрузите файлы сайта.
2. Repository → **Settings → Pages**.
3. **Source: Deploy from a branch**, выберите ветку `main` и папку `/ (root)`.

GitHub описание этого процесса: citeturn0search2turn0search5

## 4) Импорт старого db.json

Зайдите под admin/deputy → нажмите **«Импорт JSON»** → выберите ваш старый `db.json`.
После импорта нажмите **«Сохранить»** — данные уйдут в Supabase.
