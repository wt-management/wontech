-- ============================================================
-- 요청·신고함 조회 권한 (관리자 전용)
-- Supabase 대시보드 → SQL Editor 에 붙여넣고 1회 실행.
-- site_feedback / site_errors 는 익명 INSERT(요청·에러 수집)만 열려 있어
-- 브라우저에서 목록을 못 읽음 → 관리자(crm_profiles.role='admin')에게 SELECT 허용.
-- (INSERT 정책은 건드리지 않음 · 수집은 그대로 동작)
-- ============================================================

alter table public.site_feedback enable row level security;
alter table public.site_errors  enable row level security;

drop policy if exists "admin_read_feedback" on public.site_feedback;
create policy "admin_read_feedback" on public.site_feedback
  for select to authenticated
  using (exists (select 1 from public.crm_profiles p
                 where p.id = auth.uid() and p.role = 'admin'));

drop policy if exists "admin_read_errors" on public.site_errors;
create policy "admin_read_errors" on public.site_errors
  for select to authenticated
  using (exists (select 1 from public.crm_profiles p
                 where p.id = auth.uid() and p.role = 'admin'));
