-- Profesor Liceu — partea comercială (rulează DUPĂ supabase_liceu_auth.sql). Idempotent.
-- Aplicația folosește cheia SECRETĂ (sb_secret_...), deci rolurile publice nu primesc niciun drept.

-- ── Profil, rol, vârstă și acord parental ──
alter table public.liceu_users add column if not exists role              text not null default 'student'; -- 'student' | 'admin'
alter table public.liceu_users add column if not exists clasa             text;
alter table public.liceu_users add column if not exists filiera           text;
alter table public.liceu_users add column if not exists profil_liceu      text;
alter table public.liceu_users add column if not exists specializare      text;
alter table public.liceu_users add column if not exists birth_year        integer;
alter table public.liceu_users add column if not exists birth_month       integer;
alter table public.liceu_users add column if not exists parent_email      text;
alter table public.liceu_users add column if not exists parental_status   text not null default 'not_required'; -- 'not_required' | 'pending' | 'granted'

-- Cererile de acord parental (link trimis pe emailul părintelui; în DB doar hash-ul tokenului)
create table if not exists public.liceu_parent_consents (
    user_id        uuid primary key references public.liceu_users(id) on delete cascade,
    parent_email   text not null,
    token_hash     text not null unique,
    created_at     double precision not null,
    expires_at     double precision not null,
    granted_at     double precision,
    consent_version text
);

-- Dovada acceptării documentelor legale (cine, ce versiune, când)
create table if not exists public.liceu_legal_acceptances (
    id           uuid primary key default gen_random_uuid(),
    user_id      uuid not null references public.liceu_users(id) on delete cascade,
    doc          text not null,          -- 'termeni' | 'confidentialitate'
    version      text not null,
    accepted_by  text not null,          -- 'elev' | 'parinte'
    actor_email  text,
    accepted_at  double precision not null
);
create index if not exists idx_liceu_legal_user on public.liceu_legal_acceptances(user_id);

-- Răspunsuri AI raportate de elevi
create table if not exists public.liceu_reports (
    id          uuid primary key default gen_random_uuid(),
    user_id     uuid references public.liceu_users(id) on delete set null,
    session_id  text,
    excerpt     text,                    -- primele ~2000 de caractere din răspuns
    reason      text not null,
    comment     text,
    status      text not null default 'nou',   -- 'nou' | 'rezolvat'
    created_at  double precision not null
);

-- ── Abonamente ──
create table if not exists public.liceu_plans (
    code             text primary key,       -- ex. 'proba', 'standard'
    name             text not null,
    price_ron_cents  integer not null default 0,
    interval         text not null default 'month',
    stripe_price_id  text,                   -- price_... din Stripe (null pentru probă)
    daily_msg_limit  integer not null default 100,
    model_tier       text not null default 'standard',   -- pentru alegerea modelului AI
    active           boolean not null default true,
    sort             integer not null default 0,
    description      text
);

create table if not exists public.liceu_subscriptions (
    user_id                uuid primary key references public.liceu_users(id) on delete cascade,
    plan_code              text references public.liceu_plans(code),
    status                 text not null,   -- 'trialing' | 'active' | 'past_due' | 'canceled' | 'incomplete' | 'unpaid' | 'comp'
    trial_ends_at          double precision,
    current_period_end     double precision,
    cancel_at_period_end   boolean default false,
    stripe_customer_id     text,
    stripe_subscription_id text unique,
    payer_email            text,
    withdrawal_waiver_at   double precision,   -- renunțarea la dreptul de retragere (bifată la plată)
    updated_at             double precision not null
);

create table if not exists public.liceu_usage (
    user_id   uuid not null references public.liceu_users(id) on delete cascade,
    day       date not null,
    msg_count integer not null default 0,
    primary key (user_id, day)
);

-- Evenimente Stripe deja procesate (webhook idempotent)
create table if not exists public.liceu_stripe_events (
    id          text primary key,
    type        text,
    received_at double precision not null
);

-- Incrementare atomică a contorului zilnic
create or replace function public.liceu_usage_inc(p_user uuid, p_day date)
returns integer language sql as $$
    insert into public.liceu_usage(user_id, day, msg_count) values (p_user, p_day, 1)
    on conflict (user_id, day) do update set msg_count = public.liceu_usage.msg_count + 1
    returning msg_count;
$$;

-- Planuri de pornire (prețurile/limitele se editează aici; stripe_price_id se completează după
-- crearea produselor în Stripe)
insert into public.liceu_plans(code, name, price_ron_cents, stripe_price_id, daily_msg_limit, model_tier, active, sort, description)
values ('proba', 'Perioadă de probă', 0, null, 30, 'standard', true, 0, '7 zile gratuit, fără card'),
       ('standard', 'Standard', 0, null, 100, 'standard', false, 1, 'DE COMPLETAT: preț + stripe_price_id, apoi active = true')
on conflict (code) do nothing;

alter table public.liceu_parent_consents   enable row level security;
alter table public.liceu_legal_acceptances enable row level security;
alter table public.liceu_reports           enable row level security;
alter table public.liceu_plans             enable row level security;
alter table public.liceu_subscriptions     enable row level security;
alter table public.liceu_usage             enable row level security;
alter table public.liceu_stripe_events     enable row level security;
revoke all on public.liceu_parent_consents, public.liceu_legal_acceptances, public.liceu_reports,
              public.liceu_plans, public.liceu_subscriptions, public.liceu_usage, public.liceu_stripe_events
       from anon, authenticated;
revoke all on function public.liceu_usage_inc(uuid, date) from anon, authenticated, public;

-- ROLLBACK:
-- drop function public.liceu_usage_inc(uuid, date);
-- drop table public.liceu_stripe_events, public.liceu_usage, public.liceu_subscriptions, public.liceu_plans,
--            public.liceu_reports, public.liceu_legal_acceptances, public.liceu_parent_consents;
-- alter table public.liceu_users drop column role, drop column clasa, drop column filiera, drop column profil_liceu,
--   drop column specializare, drop column birth_year, drop column birth_month, drop column parent_email,
--   drop column parental_status;
