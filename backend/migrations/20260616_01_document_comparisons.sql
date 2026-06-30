-- Migration date: 2026-06-16

-- Migration: add document_comparisons for the Document Compare / Redline engine.
-- Project-scoped metadata for a deterministic comparison between two document
-- versions. The redline .docx and the diff JSON live in object storage under
-- the comparisons/ key prefix and are referenced here by storage path. Access
-- is enforced at the application layer via checkProjectAccess (the backend
-- service-role client bypasses RLS). RLS is enabled and direct client grants
-- are revoked as defense in depth, matching the backend-only lockdown pattern
-- in 20260508_01_revoke_client_grants_backend_tables.sql.

create table if not exists public.document_comparisons (
  id uuid primary key default gen_random_uuid(),
  project_id uuid not null references public.projects(id) on delete cascade,
  base_document_id uuid not null references public.documents(id) on delete cascade,
  revised_document_id uuid not null references public.documents(id) on delete cascade,
  created_by text,
  status text not null default 'pending'
    check (status = any (array[
      'pending'::text,
      'processing'::text,
      'complete'::text,
      'error'::text
    ])),
  redline_storage_path text,
  diff_storage_path text,
  error text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_document_comparisons_project
  on public.document_comparisons(project_id);

create index if not exists idx_document_comparisons_created_by
  on public.document_comparisons(created_by);

alter table public.document_comparisons enable row level security;

-- Backend-owned, project-scoped table: revoke direct browser (anon/authenticated)
-- privileges so all access flows through the backend API after JWT verification.
revoke all privileges on table public.document_comparisons from anon, authenticated;
