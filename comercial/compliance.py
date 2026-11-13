"""Acord parental, acceptarea documentelor legale, raportări, export și ștergerea datelor (GDPR)."""
from __future__ import annotations

import json

from . import core

PARENT_LINK_DAYS = 7
PARENT_RESEND_SECONDS = 300


# ── Documente legale ──────────────────────────────────────────────────────────

def record_acceptance(user_id: str, accepted_by: str, actor_email: str | None, docs=None) -> bool:
    try:
        rows = [{"user_id": user_id, "doc": d, "version": core.LEGAL_VERSIONS[d], "accepted_by": accepted_by,
                 "actor_email": actor_email, "accepted_at": core.now()}
                for d in (docs or core.LEGAL_VERSIONS)]
        for r in rows:
            core.db().table("liceu_legal_acceptances").insert(r).execute()
        return True
    except Exception as e:
        core.log("Salvarea acceptării documentelor legale a eșuat", e)
        return False


def missing_acceptances(user_id: str) -> list[str]:
    """Documentele a căror versiune curentă nu a fost încă acceptată de acest cont."""
    try:
        res = core.db().table("liceu_legal_acceptances").select("doc, version").eq("user_id", user_id).execute()
        have = {(r["doc"], r["version"]) for r in (res.data or [])}
    except Exception as e:
        core.log("Citirea acceptărilor legale a eșuat", e)
        return []          # nu blocăm elevul dacă DB-ul are o problemă temporară
    return [d for d, v in core.LEGAL_VERSIONS.items() if (d, v) not in have]


# ── Acord parental (elevi sub 16 ani) ──────────────────────────────────────────

def request_parental_consent(user: dict, parent_email: str, renew: bool = False) -> tuple[bool, str]:
    """Creează (sau reînnoiește) cererea și trimite părintelui linkul de confirmare.
    renew=True: cerere nouă chiar dacă acordul a fost dat (documente legale actualizate)."""
    if not parent_email or "@" not in parent_email:
        return False, "Scrie emailul părintelui."
    if not core.base_url():
        core.log("APP_BASE_URL lipsește din secrets — nu pot trimite linkul de acord parental")
        return False, "Aplicația nu e configurată complet (APP_BASE_URL). Anunță administratorul."
    client = core.db()
    try:
        old = client.table("liceu_parent_consents").select("created_at, granted_at") \
            .eq("user_id", user["id"]).limit(1).execute()
        if old.data and old.data[0].get("granted_at") and not renew:
            return True, "Acordul părintelui a fost deja dat."
        if old.data and not old.data[0].get("granted_at") and core.now() - (old.data[0].get("created_at") or 0) < PARENT_RESEND_SECONDS:
            return False, "Am trimis deja un email acum câteva minute. Verifică și folderul Spam."
        raw = core.new_token()
        row = {"user_id": user["id"], "parent_email": parent_email, "token_hash": core.token_hash(raw),
               "created_at": core.now(), "expires_at": core.now() + PARENT_LINK_DAYS * 86400,
               "granted_at": None, "consent_version": core.LEGAL_VERSIONS["termeni"]}
        client.table("liceu_parent_consents").delete().eq("user_id", user["id"]).execute()
        client.table("liceu_parent_consents").insert(row).execute()
        client.table("liceu_users").update({"parent_email": parent_email, "parental_status": "pending"}) \
            .eq("id", user["id"]).execute()
    except Exception as e:
        core.log("Cererea de acord parental nu a putut fi salvată", e)
        return False, "Nu am putut trimite cererea acum. Încearcă din nou."
    link = f"{core.base_url()}/?parinte={raw}"
    name = user.get("name") or user.get("email")
    body = (
        f"Bună ziua,\n\n{name} ({user.get('email')}) și-a creat un cont în {core.app_name()}, "
        f"un asistent de învățare bazat pe inteligență artificială (AI) pentru elevii de liceu.\n\n"
        f"Pentru că are sub {core.AGE_DIGITAL_CONSENT} ani, legea (GDPR, art. 8) cere acordul dumneavoastră "
        f"înainte să poată folosi contul. Pe pagina de mai jos găsiți Termenii, Politica de confidențialitate "
        f"și puteți confirma sau refuza:\n\n{link}\n\n"
        f"Linkul e valabil {PARENT_LINK_DAYS} zile. Dacă nu recunoașteți această cerere, ignorați emailul — "
        f"contul rămâne blocat și datele lui se șterg.\n\n{core.app_name()} — {core.company()['nume']}\n"
    )
    if not core.send_email(parent_email, f"Acordul dumneavoastră pentru contul {core.app_name()}", body):
        return False, "Emailul către părinte nu a putut fi trimis. Verifică adresa și încearcă din nou."
    return True, f"Am trimis linkul de confirmare la {parent_email}."


def find_parent_request(raw_token: str) -> dict | None:
    if not raw_token or len(raw_token) > 200:
        return None
    try:
        res = core.db().table("liceu_parent_consents").select("*") \
            .eq("token_hash", core.token_hash(raw_token)).limit(1).execute()
        if not res.data:
            return None
        req = res.data[0]
        if req.get("expires_at", 0) < core.now() and not req.get("granted_at"):
            return None
        u = core.db().table("liceu_users").select("id, email, display_name, birth_year, birth_month") \
            .eq("id", req["user_id"]).limit(1).execute()
        if not u.data:
            return None
        req["student"] = u.data[0]
        return req
    except Exception as e:
        core.log("Citirea cererii de acord parental a eșuat", e)
        return None


def grant_parental_consent(req: dict) -> bool:
    try:
        client = core.db()
        client.table("liceu_parent_consents").update({"granted_at": core.now()}) \
            .eq("user_id", req["user_id"]).execute()
        client.table("liceu_users").update({"parental_status": "granted"}).eq("id", req["user_id"]).execute()
        return record_acceptance(req["user_id"], "parinte", req["parent_email"])
    except Exception as e:
        core.log("Confirmarea acordului parental a eșuat", e)
        return False


# ── Raportarea răspunsurilor AI ───────────────────────────────────────────────

REPORT_REASONS = [
    "Informație greșită",
    "Conținut nepotrivit pentru elevi",
    "Răspuns jignitor sau periculos",
    "Altceva",
]


def save_report(user_id: str | None, session_id: str | None, excerpt: str, reason: str, comment: str) -> bool:
    try:
        core.db().table("liceu_reports").insert({
            "user_id": user_id, "session_id": session_id, "excerpt": (excerpt or "")[:2000],
            "reason": reason, "comment": (comment or "")[:1000], "status": "nou", "created_at": core.now(),
        }).execute()
        return True
    except Exception as e:
        core.log("Raportarea nu a putut fi salvată", e)
        return False


# ── Drepturile GDPR: export și ștergere ───────────────────────────────────────

_USER_EXPORT_COLS = ("id, email, display_name, created_at, last_login, role, clasa, filiera, profil_liceu, "
                     "specializare, birth_year, birth_month, parent_email, parental_status")


def export_user_data(user_id: str) -> bytes:
    """Toate datele contului, ca JSON (dreptul de acces și portabilitate, art. 15 și 20 GDPR)."""
    client = core.db()
    out: dict = {"generat_la": core.now(), "aplicatie": core.app_name(), "operator": core.company()}

    def _get(table, cols="*", **eq):
        q = client.table(table).select(cols)
        for k, v in eq.items():
            q = q.eq(k, v)
        try:
            return q.execute().data or []
        except Exception as e:
            core.log(f"Export: {table} nu a putut fi citit", e)
            return []

    users = _get("liceu_users", _USER_EXPORT_COLS, id=user_id)
    out["cont"] = users[0] if users else {}
    sessions = _get("sessions", "session_id, app_id, created_at, last_active", liceu_user_id=user_id)
    convs = []
    for s in sessions:
        msgs = _get("history", "role, content, timestamp", session_id=s["session_id"])
        convs.append({**s, "mesaje": [m for m in msgs if m.get("role") in ("user", "assistant")]})
    out["conversatii"] = convs
    out["acceptari_documente"] = _get("liceu_legal_acceptances", "doc, version, accepted_by, actor_email, accepted_at",
                                      user_id=user_id)
    out["acord_parental"] = _get("liceu_parent_consents", "parent_email, created_at, granted_at, consent_version",
                                 user_id=user_id)
    out["abonament"] = _get("liceu_subscriptions", "plan_code, status, trial_ends_at, current_period_end, "
                            "cancel_at_period_end, payer_email, withdrawal_waiver_at", user_id=user_id)
    out["utilizare"] = _get("liceu_usage", "day, msg_count", user_id=user_id)
    out["raportari"] = _get("liceu_reports", "session_id, reason, comment, status, created_at", user_id=user_id)
    return json.dumps(out, ensure_ascii=False, indent=2, default=str).encode("utf-8")


def delete_account(user_id: str, cancel_billing=None) -> bool:
    """Șterge contul și toate conversațiile lui (dreptul la ștergere, art. 17 GDPR).

    Abonamentul Stripe (dacă există) e anulat înainte; facturile rămân la Stripe/contabilitate,
    unde legea cere păstrarea lor."""
    client = core.db()
    try:
        if cancel_billing:
            cancel_billing(user_id)
        sess = client.table("sessions").select("session_id").eq("liceu_user_id", user_id).execute().data or []
        sids = [s["session_id"] for s in sess]
        if sids:
            client.table("history").delete().in_("session_id", sids).execute()
            client.table("sessions").delete().in_("session_id", sids).execute()
        client.table("liceu_reports").update({"user_id": None}).eq("user_id", user_id).execute()
        for t in ("liceu_usage", "liceu_subscriptions", "liceu_legal_acceptances", "liceu_parent_consents",
                  "liceu_tokens", "liceu_password_resets", "liceu_api_keys"):
            try:
                client.table(t).delete().eq("user_id", user_id).execute()
            except Exception as e:
                core.log(f"Ștergere: {t}", e)
        client.table("liceu_users").delete().eq("id", user_id).execute()
        return True
    except Exception as e:
        core.log("Ștergerea contului a eșuat", e)
        return False
