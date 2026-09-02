-- Construction site transactions system
-- Namespaced with cst_ prefix so it lives safely alongside the payroll-system tables
-- in the same Supabase Postgres database.

create table if not exists cst_sites (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  location text,
  created_at timestamptz not null default now()
);

create table if not exists cst_users (
  id uuid primary key default gen_random_uuid(),
  full_name text not null,
  email text unique not null,
  password_hash text not null,
  role text not null check (role in ('clerk', 'manager', 'finance', 'admin')),
  site_id uuid references cst_sites(id),
  created_at timestamptz not null default now()
);

create table if not exists cst_transactions (
  id uuid primary key default gen_random_uuid(),
  site_id uuid not null references cst_sites(id),
  category text not null check (category in ('materials', 'labor', 'equipment', 'fuel', 'other')),
  description text not null,
  amount numeric(14,2) not null check (amount > 0),
  transaction_date date not null default current_date,

  status text not null default 'pending'
    check (status in ('pending', 'approved', 'rejected', 'paid')),

  manager_approved boolean not null default false,
  manager_approved_by uuid references cst_users(id),
  manager_approved_at timestamptz,

  finance_approved boolean not null default false,
  finance_approved_by uuid references cst_users(id),
  finance_approved_at timestamptz,

  rejected_by uuid references cst_users(id),
  rejected_at timestamptz,
  rejection_reason text,

  paid_by uuid references cst_users(id),
  paid_at timestamptz,

  created_by uuid not null references cst_users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_cst_transactions_status on cst_transactions(status);
create index if not exists idx_cst_transactions_site on cst_transactions(site_id);

create table if not exists cst_approval_log (
  id uuid primary key default gen_random_uuid(),
  transaction_id uuid not null references cst_transactions(id) on delete cascade,
  actor_id uuid not null references cst_users(id),
  action text not null check (action in ('approve_manager', 'approve_finance', 'reject', 'mark_paid')),
  note text,
  created_at timestamptz not null default now()
);
