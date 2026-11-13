-- Autentificare opțională pentru Profesor Liceu (Master) — tabele liceu_*, SEPARATE de cele ale Profesor ETTI (etti_*).
-- Același proiect Supabase; conversațiile rămân separate și prin app_id. Idempotent: se poate rula din nou.
-- Aplicația folosește cheia SECRETĂ (sb_secret_...), deci rolurile publice nu primesc niciun drept.

create table if not exists public.liceu_users (
    id            uuid primary key default gen_random_uuid(),
    email         text not null unique,          -- salvat mereu cu litere mici
    password_hash text not null,                 -- scrypt$n$r$p$salt_hex$hash_hex
    display_name  text,
    created_at    double precision not null,
    last_login    double precision
);

create table if not exists public.liceu_tokens (
    token_hash text primary key,                 -- sha256(token); tokenul brut stă doar în browser
    user_id    uuid not null references public.liceu_users(id) on delete cascade,
    created_at double precision not null,
    expires_at double precision not null,
    last_used  double precision
);
create index if not exists idx_liceu_tokens_user_id on public.liceu_tokens(user_id);

alter table public.liceu_users  enable row level security;
alter table public.liceu_tokens enable row level security;
revoke all on public.liceu_users, public.liceu_tokens from anon, authenticated;

-- Conversațiile unui cont (sessions.user_id e deja legat de users din Profesor de Liceu)
alter table public.sessions
    add column if not exists liceu_user_id uuid references public.liceu_users(id) on delete set null;
create index if not exists idx_sessions_liceu_user_id on public.sessions(liceu_user_id);

create or replace view public.session_previews with (security_invoker = true) as
select s.session_id,
       s.app_id,
       s.last_active,
       count(h.session_id) as msg_count,
       left(min(h.content) filter (where h.role = 'user'), 60) as preview,
       s.etti_user_id,          -- folosit de Profesor ETTI (nu se șterge)
       s.liceu_user_id          -- folosit de Profesor Liceu (Master)
from public.sessions s
left join public.history h on h.session_id = s.session_id and h.app_id = s.app_id
group by s.session_id, s.app_id, s.last_active, s.etti_user_id, s.liceu_user_id;

revoke all on public.session_previews from anon, authenticated;

-- ROLLBACK (dacă e nevoie) — refă view-ul FĂRĂ liceu_user_id (cel al ETTI rămâne):
-- create or replace view public.session_previews with (security_invoker = true) as
--   select s.session_id, s.app_id, s.last_active, count(h.session_id) as msg_count,
--          left(min(h.content) filter (where h.role = 'user'), 60) as preview, s.etti_user_id
--   from public.sessions s left join public.history h on h.session_id = s.session_id and h.app_id = s.app_id
--   group by s.session_id, s.app_id, s.last_active, s.etti_user_id;
-- (atenție: CREATE OR REPLACE nu poate elimina o coloană — pentru rollback: DROP VIEW apoi CREATE VIEW)
-- alter table public.sessions drop column liceu_user_id;
-- drop table public.liceu_tokens; drop table public.liceu_users;

-- ── Resetare parolă (cod prin email + cod de recuperare) ──
alter table public.liceu_users add column if not exists recovery_hash text;   -- scrypt(cod de recuperare)
create table if not exists public.liceu_password_resets (
    user_id    uuid primary key references public.liceu_users(id) on delete cascade,
    code_hash  text not null,              -- hmac-sha256 al codului de 6 cifre trimis pe email
    expires_at double precision not null,  -- 15 minute
    attempts   integer not null default 0, -- max 5 încercări
    created_at double precision not null
);
alter table public.liceu_password_resets enable row level security;
revoke all on public.liceu_password_resets from anon, authenticated;
-- ROLLBACK: drop table public.liceu_password_resets; alter table public.liceu_users drop column recovery_hash;

-- ── Cheile API Gemini ale elevului, salvate în cont (criptate în aplicație cu Fernet) ──
create table if not exists public.liceu_api_keys (
    id         uuid primary key default gen_random_uuid(),
    user_id    uuid not null references public.liceu_users(id) on delete cascade,
    label      text,                     -- ex. „Mama”, „Tata”
    key_enc    text not null,            -- „s1:<fernet>” (KEYS_ENCRYPTION_SECRET) sau „d1:<fernet>” (derivat)
    key_hint   text,                     -- ultimele 4 caractere, pentru afișare
    created_at double precision not null
);
create index if not exists idx_liceu_api_keys_user_id on public.liceu_api_keys(user_id);
alter table public.liceu_api_keys enable row level security;
revoke all on public.liceu_api_keys from anon, authenticated;
-- ROLLBACK: drop table public.liceu_api_keys;
