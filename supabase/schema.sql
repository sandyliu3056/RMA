-- 售後零件規劃平台：Supabase 資料庫結構。在 Supabase 的 SQL Editor 貼上整份執行一次即可（重複執行也安全）。
--
-- 設計
--   imports         每次匯入的 AIO 檔案一筆：資料日期、檔名、欄位順序、匯入當時的快照指標（案件數、未結案、缺料、待出貨）
--   case_snapshots  每個案件在每次匯入時的版本（CaseID × 資料日期），row 是整列原始值；同一案件以資料日期最新的版本為準
-- 權限
--   只有登入的使用者（authenticated）能讀寫；匿名（anon）什麼都看不到。
--   請在 Authentication → Providers → Email 關閉自行註冊（Allow new users to sign up），使用者由管理員在 Users 建立。

create table if not exists public.imports (
  id          bigint generated always as identity primary key,
  data_date   date not null unique,             -- 資料日期（匯出檔最後一筆建案日）；同一天再匯入會覆蓋
  file_name   text,
  row_count   integer not null default 0,
  columns     jsonb not null default '[]'::jsonb, -- 匯出檔的欄位順序（載入時還原成表格用）
  kpis        jsonb not null default '{}'::jsonb, -- {cases, open, shortage, ready}
  imported_by text,
  imported_at timestamptz not null default now()
);

create table if not exists public.case_snapshots (
  case_id     text not null,
  data_date   date not null,
  import_id   bigint not null references public.imports(id) on delete cascade,
  create_date timestamp,
  close_date  timestamp,
  status      text,
  country     text,
  row         jsonb not null,                    -- 整列原始值（空值已去掉）
  primary key (case_id, data_date)
);
create index if not exists case_snapshots_create_date_idx on public.case_snapshots (create_date);
create index if not exists case_snapshots_import_idx on public.case_snapshots (import_id);

-- 角色（Supabase 已有 anon / authenticated；這段只是讓同一份 SQL 也能在一般 PostgreSQL 上測試）
do $$
begin
  if not exists (select 1 from pg_roles where rolname = 'anon') then create role anon nologin; end if;
  if not exists (select 1 from pg_roles where rolname = 'authenticated') then create role authenticated nologin; end if;
end $$;

-- 列級安全：登入者可讀寫全部，匿名者沒有任何規則＝看不到
alter table public.imports enable row level security;
alter table public.case_snapshots enable row level security;
drop policy if exists "members: imports" on public.imports;
create policy "members: imports" on public.imports for all to authenticated using (true) with check (true);
drop policy if exists "members: case_snapshots" on public.case_snapshots;
create policy "members: case_snapshots" on public.case_snapshots for all to authenticated using (true) with check (true);

revoke all on table public.imports, public.case_snapshots from anon;
grant select, insert, update, delete on table public.imports, public.case_snapshots to authenticated;
grant usage, select on all sequences in schema public to authenticated;
grant usage on schema public to authenticated;

-- 寫入一批案件（網頁每次送 1,000 列）：p_rows = [{case_id, create_date, close_date, status, country, row}, …]
-- 日期用「YYYY-MM-DD HH:MM:SS」字串；同一案件同一資料日期已存在就覆蓋。回傳寫入列數。
create or replace function public.upsert_cases(p_import_id bigint, p_rows jsonb)
returns integer
language plpgsql
security invoker
set search_path = public
as $$
declare
  n integer;
begin
  insert into public.case_snapshots (case_id, data_date, import_id, create_date, close_date, status, country, row)
  select r->>'case_id', i.data_date, i.id,
         nullif(r->>'create_date', '')::timestamp,
         nullif(r->>'close_date', '')::timestamp,
         nullif(r->>'status', ''),
         nullif(r->>'country', ''),
         jsonb_strip_nulls(coalesce(r->'row', '{}'::jsonb))
  from public.imports i
  cross join jsonb_array_elements(p_rows) as r
  where i.id = p_import_id and coalesce(r->>'case_id', '') <> ''
  on conflict (case_id, data_date) do update
    set import_id = excluded.import_id, create_date = excluded.create_date, close_date = excluded.close_date,
        status = excluded.status, country = excluded.country, row = excluded.row;
  get diagnostics n = row_count;
  update public.imports
     set row_count = (select count(*) from public.case_snapshots s where s.import_id = p_import_id)
   where id = p_import_id;
  return n;
end
$$;

-- 讀取期間內（建案日 p_from ～ p_to，含；null＝不限）的案件，每個 CaseID 取資料日期最新的版本；
-- 以 CaseID 排序分頁（p_offset / p_limit）。每列多一個 _data_date：這個版本來自哪一天的匯出。
create or replace function public.get_cases(p_from date default null, p_to date default null,
                                            p_offset integer default 0, p_limit integer default 5000)
returns jsonb
language sql
stable
security invoker
set search_path = public
as $$
  select coalesce(jsonb_agg(s.row || jsonb_build_object('_data_date', s.data_date) order by s.case_id), '[]'::jsonb)
  from (
    select t.case_id, t.data_date, t.row
    from (
      select distinct on (case_id) case_id, data_date, row
      from public.case_snapshots
      where (p_from is null or create_date >= p_from)
        and (p_to is null or create_date < p_to + 1)
      order by case_id, data_date desc
    ) t
    order by t.case_id
    offset greatest(p_offset, 0) limit least(greatest(p_limit, 1), 20000)
  ) s;
$$;

revoke all on function public.upsert_cases(bigint, jsonb) from public, anon;
revoke all on function public.get_cases(date, date, integer, integer) from public, anon;
grant execute on function public.upsert_cases(bigint, jsonb) to authenticated;
grant execute on function public.get_cases(date, date, integer, integer) to authenticated;

-- 讓 Supabase 的 API 立刻看到新的表與函式
notify pgrst, 'reload schema';
