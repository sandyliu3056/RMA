-- 售後零件規劃平台：Supabase 資料庫結構。在 Supabase 的 SQL Editor 貼上整份執行一次即可（重複執行也安全）。
--
-- 設計
--   imports         每次匯入的 AIO 檔案一筆：資料日期、檔名、欄位順序、匯入當時的快照指標（案件數、未結案、缺料、待出貨）。
--                   complete 代表全部列都上傳完成；未完成的匯入（上傳到一半斷線）不會被讀取，可以在畫面上刪掉重存。
--   case_snapshots  每個案件在每次匯入時的版本（匯入 × CaseID），row 是整列原始值；同一案件以資料日期最新的「完成」匯入為準。
-- 流程（網頁 app.js）：建立 imports（complete=false）→ upsert_cases 分批上傳 → finish_import 一次交易：刪掉同一資料日期的舊匯入、標記完成。
-- 權限
--   只有登入的使用者（authenticated，且不是匿名登入）能讀寫；匿名（anon）什麼都看不到。
--   請在 Authentication → Sign In / Providers → User Signups 關閉「Allow new users to sign up」與「Allow anonymous sign-ins」，
--   使用者由管理員在 Authentication → Users 建立。

create table if not exists public.imports (
  id          bigint generated always as identity primary key,
  data_date   date not null,                     -- 資料日期（匯出檔最後一筆建案日）；同一天只能有一筆「完成」的匯入
  file_name   text,
  row_count   integer not null default 0,
  columns     jsonb not null default '[]'::jsonb, -- 匯出檔的欄位順序（載入時還原成表格用）
  kpis        jsonb not null default '{}'::jsonb, -- {cases, open, shortage, ready}
  imported_by text,
  imported_at timestamptz not null default now(),
  complete    boolean not null default false     -- finish_import 之後才為 true
);
-- 第一版沒有 complete 欄：升級用
alter table public.imports add column if not exists complete boolean not null default false;
alter table public.imports drop constraint if exists imports_data_date_key;
create unique index if not exists imports_data_date_complete_idx on public.imports (data_date) where complete;

create table if not exists public.case_snapshots (
  case_id     text not null,
  data_date   date not null,
  import_id   bigint not null references public.imports(id) on delete cascade,
  create_date timestamp,
  close_date  timestamp,
  status      text,
  country     text,
  row         jsonb not null,                    -- 整列原始值（空值已去掉）
  primary key (import_id, case_id)               -- 每次匯入各自一份，上傳到一半的匯入不會動到別次的資料
);
-- 第一版的主鍵是 (case_id, data_date)：升級用
do $$
begin
  if exists (select 1 from pg_indexes where schemaname = 'public' and tablename = 'case_snapshots'
             and indexname = 'case_snapshots_pkey' and indexdef like '%(case_id, data_date)%') then
    alter table public.case_snapshots drop constraint case_snapshots_pkey;
    alter table public.case_snapshots add primary key (import_id, case_id);
  end if;
end $$;
create index if not exists case_snapshots_create_date_idx on public.case_snapshots (create_date);
create index if not exists case_snapshots_latest_idx on public.case_snapshots (case_id, data_date desc);

-- 角色（Supabase 已有 anon / authenticated；這段只是讓同一份 SQL 也能在一般 PostgreSQL 上測試）
do $$
begin
  if not exists (select 1 from pg_roles where rolname = 'anon') then create role anon nologin; end if;
  if not exists (select 1 from pg_roles where rolname = 'authenticated') then create role authenticated nologin; end if;
end $$;

-- 列級安全：登入者（非匿名）可讀寫全部；匿名者沒有任何規則＝看不到。
-- request.jwt.claims 是 Supabase 替每個請求設定的 JWT 內容；一般 PostgreSQL 沒有這個設定時視為非匿名。
alter table public.imports enable row level security;
alter table public.case_snapshots enable row level security;
drop policy if exists "members: imports" on public.imports;
create policy "members: imports" on public.imports for all to authenticated
  using (coalesce(current_setting('request.jwt.claims', true)::jsonb ->> 'is_anonymous', 'false') <> 'true')
  with check (coalesce(current_setting('request.jwt.claims', true)::jsonb ->> 'is_anonymous', 'false') <> 'true');
drop policy if exists "members: case_snapshots" on public.case_snapshots;
create policy "members: case_snapshots" on public.case_snapshots for all to authenticated
  using (coalesce(current_setting('request.jwt.claims', true)::jsonb ->> 'is_anonymous', 'false') <> 'true')
  with check (coalesce(current_setting('request.jwt.claims', true)::jsonb ->> 'is_anonymous', 'false') <> 'true');

revoke all on table public.imports, public.case_snapshots from anon;
grant select, insert, update, delete on table public.imports, public.case_snapshots to authenticated;
grant usage, select on all sequences in schema public to authenticated;
grant usage on schema public to authenticated;

-- 寫入一批案件（網頁每次送 1,000 列）到「還沒完成」的匯入：p_rows = [{case_id, create_date, close_date, status, country, row}, …]
-- 日期用「YYYY-MM-DD HH:MM:SS」字串；同一批裡重複的 CaseID 以最後一列為準；同一次匯入已有的案件就覆蓋。回傳寫入列數。
create or replace function public.upsert_cases(p_import_id bigint, p_rows jsonb)
returns integer
language plpgsql
security invoker
set search_path = public
as $$
declare
  n integer;
begin
  if not exists (select 1 from public.imports where id = p_import_id and not complete) then
    raise exception 'import % not found or already complete', p_import_id;
  end if;
  insert into public.case_snapshots (case_id, data_date, import_id, create_date, close_date, status, country, row)
  select d.v->>'case_id', i.data_date, i.id,
         nullif(d.v->>'create_date', '')::timestamp,
         nullif(d.v->>'close_date', '')::timestamp,
         nullif(d.v->>'status', ''),
         nullif(d.v->>'country', ''),
         jsonb_strip_nulls(coalesce(d.v->'row', '{}'::jsonb))
  from public.imports i
  cross join lateral (
    select distinct on (r.v->>'case_id') r.v
    from jsonb_array_elements(p_rows) with ordinality as r(v, ord)
    where coalesce(r.v->>'case_id', '') <> ''
    order by r.v->>'case_id', r.ord desc
  ) as d
  where i.id = p_import_id
  on conflict (import_id, case_id) do update
    set data_date = excluded.data_date, create_date = excluded.create_date, close_date = excluded.close_date,
        status = excluded.status, country = excluded.country, row = excluded.row;
  get diagnostics n = row_count;
  update public.imports
     set row_count = (select count(*) from public.case_snapshots s where s.import_id = p_import_id)
   where id = p_import_id;
  return n;
end
$$;

-- 全部列上傳完成後呼叫：同一個交易裡刪掉同一資料日期的其他匯入（舊的或沒上傳完的，快照一起刪）、把這次標記為完成。回傳列數。
create or replace function public.finish_import(p_import_id bigint)
returns integer
language plpgsql
security invoker
set search_path = public
as $$
declare
  d date;
  n integer;
begin
  select data_date into d from public.imports where id = p_import_id and not complete;
  if d is null then
    raise exception 'import % not found or already complete', p_import_id;
  end if;
  select count(*) into n from public.case_snapshots where import_id = p_import_id;
  if n = 0 then
    raise exception 'import % has no rows', p_import_id;
  end if;
  delete from public.imports where data_date = d and id <> p_import_id;
  update public.imports set complete = true, row_count = n where id = p_import_id;
  return n;
end
$$;

-- 讀取期間內（建案日 p_from ～ p_to，含；null＝不限）的案件，只看「完成」的匯入，每個 CaseID 取資料日期最新的版本。
-- 以 CaseID 排序，p_after 是上一頁最後一個 CaseID（第一頁給空字串），每頁最多 p_limit 列（鍵集分頁，頁數再多也不會變慢）。
-- 每列多兩個欄位：_data_date（這個版本來自哪一天的匯出）、_case_id（分頁用）。
drop function if exists public.get_cases(date, date, integer, integer);
create or replace function public.get_cases(p_from date default null, p_to date default null,
                                            p_after text default '', p_limit integer default 5000)
returns jsonb
language sql
stable
security invoker
set search_path = public
as $$
  select coalesce(jsonb_agg(s.row || jsonb_build_object('_data_date', s.data_date, '_case_id', s.case_id) order by s.case_id), '[]'::jsonb)
  from (
    select distinct on (s.case_id) s.case_id, s.data_date, s.row
    from public.case_snapshots s
    join public.imports i on i.id = s.import_id and i.complete
    where s.case_id > coalesce(p_after, '')
      and (p_from is null or s.create_date >= p_from)
      and (p_to is null or s.create_date < p_to + 1)
    order by s.case_id, s.data_date desc
    limit least(greatest(p_limit, 1), 20000)
  ) s;
$$;

revoke all on function public.upsert_cases(bigint, jsonb) from public, anon;
revoke all on function public.finish_import(bigint) from public, anon;
revoke all on function public.get_cases(date, date, text, integer) from public, anon;
grant execute on function public.upsert_cases(bigint, jsonb) to authenticated;
grant execute on function public.finish_import(bigint) to authenticated;
grant execute on function public.get_cases(date, date, text, integer) to authenticated;

-- 讓 Supabase 的 API 立刻看到新的表與函式
notify pgrst, 'reload schema';
