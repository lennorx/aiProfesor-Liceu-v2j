"""Legătura dintre modulele comerciale și aplicație: dependențe, configurare, utilitare.

app.py apelează o singură dată setup(...) cu funcțiile lui (client Supabase, log, email),
ca modulele de aici să nu importe app.py (Streamlit îl rulează ca __main__).
"""
from __future__ import annotations

import hashlib
import secrets as _secrets
import time
from datetime import datetime, timezone

import streamlit as st

_DEPS: dict = {}

# Versiunile documentelor legale. La orice modificare a textului din legal/*.md se schimbă
# versiunea aici → utilizatorii trebuie să accepte din nou (părintele, pentru elevii sub 16 ani).
LEGAL_VERSIONS = {"termeni": "2026-11-13", "confidentialitate": "2026-11-13"}
LEGAL_TITLES = {"termeni": "Termeni și condiții", "confidentialitate": "Politica de confidențialitate"}

AGE_DIGITAL_CONSENT = 16   # GDPR art. 8 — România: 16 ani
AGE_MAJORITY = 18


class AccessError(Exception):
    """Cererea AI nu e permisă (cont, acord parental, abonament, limită zilnică)."""


def setup(**deps):
    _DEPS.update(deps)


def db():
    return _DEPS["get_client"]()


def log(msg: str, exc: Exception | None = None):
    fn = _DEPS.get("log")
    if fn:
        fn(msg, "silent", exc)


def send_email(to: str, subject: str, body: str) -> bool:
    fn = _DEPS.get("send_email")
    return bool(fn and fn(to, subject, body))


def cfg(name: str, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def billing_on() -> bool:
    """BILLING_MODE = "on" în secrets activează conturile obligatorii, acordul parental și abonamentele.
    Implicit "off": aplicația merge ca înainte (util pentru dezvoltare)."""
    return str(cfg("BILLING_MODE", "off")).strip().lower() == "on"


def app_name() -> str:
    return str(cfg("APP_NAME", "Profesor Liceu"))


def base_url() -> str:
    return str(cfg("APP_BASE_URL", "")).strip().rstrip("/")


def company() -> dict:
    """Datele firmei din secrets (afișate în subsol și în documentele legale)."""
    return {
        "nume": cfg("COMPANY_NAME", "[DENUMIRE SRL]"),
        "cui": cfg("COMPANY_CUI", "[CUI]"),
        "reg_com": cfg("COMPANY_REG_COM", "[J..../..../....]"),
        "sediu": cfg("COMPANY_ADDRESS", "[ADRESA SEDIULUI]"),
        "email": cfg("COMPANY_EMAIL", cfg("SMTP_USER", "[EMAIL CONTACT]")),
        "telefon": cfg("COMPANY_PHONE", ""),
    }


def now() -> float:
    return time.time()


def today_iso() -> str:
    return datetime.now(timezone.utc).date().isoformat()


def new_token() -> str:
    return _secrets.token_urlsafe(32)


def token_hash(raw: str) -> str:
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def age_bounds(birth_year: int | None, birth_month: int | None) -> tuple[int | None, int | None]:
    """(vârsta minimă posibilă, vârsta maximă posibilă) din anul + luna nașterii (ziua nu o cerem)."""
    if not birth_year:
        return None, None
    d = datetime.now(timezone.utc).date()
    m = int(birth_month or 12)
    years = d.year - int(birth_year)
    # Ziua de naștere poate fi oricând în luna respectivă: înainte de lună → nu a împlinit;
    # după lună → a împlinit; în aceeași lună → nu știm.
    if d.month < m:
        return years - 1, years - 1
    if d.month > m:
        return years, years
    return years - 1, years


def needs_parental_consent(birth_year, birth_month) -> bool:
    lo, _ = age_bounds(birth_year, birth_month)
    return lo is None or lo < AGE_DIGITAL_CONSENT


def is_minor(birth_year, birth_month) -> bool:
    lo, _ = age_bounds(birth_year, birth_month)
    return lo is None or lo < AGE_MAJORITY
