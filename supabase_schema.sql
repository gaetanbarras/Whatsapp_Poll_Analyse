create table if not exists public.app_documents (
  document_type text not null,
  document_key text not null,
  payload jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default timezone('utc', now()),
  primary key (document_type, document_key)
);

create index if not exists idx_app_documents_type_updated_at
  on public.app_documents (document_type, updated_at desc);
