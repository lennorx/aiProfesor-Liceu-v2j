"""Planuri, perioadă de probă, limite zilnice și plăți Stripe (abonamente)."""
from __future__ import annotations

import time

import streamlit as st

from . import core, compliance

TRIAL_PLAN = "proba"
ACTIVE_STATUSES = ("active", "trialing", "comp")
GRACE_SECONDS = 3 * 86400        # câteva zile de toleranță dacă plata reînnoirii întârzie


def trial_days() -> int:
    try:
        return int(core.cfg("TRIAL_DAYS", 7))
    except Exception:
        return 7


# ── Planuri ──────────────────────────────────────────────────────────────────

@st.cache_data(ttl=60, show_spinner=False)
def _plans_cached(_v: int = 0) -> list[dict]:
    try:
        res = core.db().table("liceu_plans").select("*").order("sort").execute()
        return res.data or []
    except Exception as e:
        core.log("Planurile nu au putut fi citite", e)
        return []


def plans(paid_only: bool = False) -> list[dict]:
    out = [p for p in _plans_cached() if p.get("active")]
    if paid_only:
        out = [p for p in out if p.get("stripe_price_id")]
    return out


def plan(code: str | None) -> dict | None:
    return next((p for p in _plans_cached() if p["code"] == code), None)


# ── Abonamentul utilizatorului ────────────────────────────────────────────────

def get_subscription(user_id: str) -> dict | None:
    try:
        res = core.db().table("liceu_subscriptions").select("*").eq("user_id", user_id).limit(1).execute()
        return res.data[0] if res.data else None
    except Exception as e:
        core.log("Abonamentul nu a putut fi citit", e)
        return None


def ensure_trial(user_id: str) -> dict | None:
    """Pornește perioada de probă o singură dată per cont (la primul acces permis)."""
    sub = get_subscription(user_id)
    if sub:
        return sub
    row = {"user_id": user_id, "plan_code": TRIAL_PLAN, "status": "trialing",
           "trial_ends_at": core.now() + trial_days() * 86400, "updated_at": core.now()}
    try:
        core.db().table("liceu_subscriptions").insert(row).execute()
        return row
    except Exception as e:
        core.log("Perioada de probă nu a putut fi pornită", e)
        return get_subscription(user_id)


def _sub_is_active(sub: dict | None) -> bool:
    if not sub:
        return False
    st_ = sub.get("status")
    if st_ == "comp":
        return True
    if st_ == "trialing" and not sub.get("stripe_subscription_id"):
        return (sub.get("trial_ends_at") or 0) > core.now()
    if st_ in ("active", "trialing", "past_due"):
        end = sub.get("current_period_end") or 0
        return end == 0 or end + GRACE_SECONDS > core.now()
    return False


def usage_today(user_id: str) -> int:
    try:
        res = core.db().table("liceu_usage").select("msg_count").eq("user_id", user_id) \
            .eq("day", core.today_iso()).limit(1).execute()
        return int(res.data[0]["msg_count"]) if res.data else 0
    except Exception:
        return 0


def _usage_inc(user_id: str) -> int:
    client = core.db()
    try:
        res = client.rpc("liceu_usage_inc", {"p_user": user_id, "p_day": core.today_iso()}).execute()
        if isinstance(res.data, int):
            return res.data
        if isinstance(res.data, list) and res.data:
            v = res.data[0]
            return int(v if isinstance(v, int) else list(v.values())[0])
    except Exception:
        pass
    # rezervă fără funcția SQL (citire + scriere)
    n = usage_today(user_id) + 1
    try:
        client.table("liceu_usage").upsert({"user_id": user_id, "day": core.today_iso(), "msg_count": n}).execute()
    except Exception as e:
        core.log("Contorul de utilizare nu a putut fi actualizat", e)
    return n


def entitlement(user: dict | None) -> dict:
    """Ce are voie utilizatorul acum. reason: ok | login | parental | legal | subscribe."""
    if not core.billing_on():
        return {"ok": True, "reason": "ok", "plan": None, "limit": None}
    if not user:
        return {"ok": False, "reason": "login"}
    if user.get("role") == "admin":
        return {"ok": True, "reason": "ok", "plan": None, "limit": None}
    if user.get("parental_status") == "pending" or (
        core.needs_parental_consent(user.get("birth_year"), user.get("birth_month"))
        and user.get("parental_status") != "granted"
    ):
        return {"ok": False, "reason": "parental"}
    missing = compliance.missing_acceptances(user["id"])
    if missing:
        return {"ok": False, "reason": "legal", "missing": missing}
    sub = ensure_trial(user["id"])
    if not _sub_is_active(sub):
        return {"ok": False, "reason": "subscribe", "sub": sub}
    p = plan(sub.get("plan_code")) or {}
    limit = None if sub.get("status") == "comp" and not p else p.get("daily_msg_limit")
    return {"ok": True, "reason": "ok", "plan": p, "sub": sub, "limit": limit}


def consume(user: dict | None):
    """Apelat înaintea fiecărei cereri AI. Ridică core.AccessError dacă nu e permisă."""
    if not core.billing_on():
        return
    ent = entitlement(user)
    if not ent["ok"]:
        raise core.AccessError({
            "login": "🔐 Intră în cont (sau creează unul) ca să folosești profesorul.",
            "parental": "👨‍👩‍👧 Contul așteaptă acordul părintelui (verificați emailul părintelui).",
            "legal": "📄 Acceptă Termenii și Politica de confidențialitate actualizate (bara laterală → Cont).",
            "subscribe": "⭐ Perioada de probă sau abonamentul s-a încheiat. Alege un abonament din bara laterală.",
        }.get(ent["reason"], "Accesul nu e permis acum."))
    if ent.get("limit") and user and user.get("role") != "admin":
        if usage_today(user["id"]) >= int(ent["limit"]):
            raise core.AccessError(
                f"📅 Ai folosit toate cele {ent['limit']} întrebări de azi incluse în plan. "
                "Limita se resetează la miezul nopții (ora UTC)."
            )
        _usage_inc(user["id"])


# ── Stripe ───────────────────────────────────────────────────────────────────

def stripe_ready() -> bool:
    return bool(core.cfg("STRIPE_SECRET_KEY")) and bool(core.base_url())


def _stripe():
    import stripe  # importat doar când e nevoie
    stripe.api_key = core.cfg("STRIPE_SECRET_KEY")
    return stripe


def create_checkout_url(user: dict, plan_code: str, payer_email: str) -> tuple[str | None, str]:
    """Link Stripe Checkout pentru abonament. Plătitorul (părintele, la elevii minori) bifează acolo
    acceptarea Termenilor și renunțarea la dreptul de retragere de 14 zile (serviciul începe imediat)."""
    p = plan(plan_code)
    if not p or not p.get("stripe_price_id"):
        return None, "Planul ales nu e disponibil."
    if not stripe_ready():
        return None, "Plățile nu sunt configurate încă."
    try:
        stripe = _stripe()
        sub = get_subscription(user["id"]) or {}
        params = dict(
            mode="subscription",
            line_items=[{"price": p["stripe_price_id"], "quantity": 1}],
            client_reference_id=user["id"],
            metadata={"user_id": user["id"], "plan_code": plan_code},
            subscription_data={"metadata": {"user_id": user["id"], "plan_code": plan_code}},
            success_url=f"{core.base_url()}/?plata=ok",
            cancel_url=f"{core.base_url()}/?plata=anulata",
            locale="ro",
            allow_promotion_codes=True,
            consent_collection={"terms_of_service": "required"},
            custom_text={"terms_of_service_acceptance": {"message": (
                # Textul exact și temeiul (OUG 34/2014) — de validat cu avocatul înainte de lansare.
                "Accept Termenii și condițiile. Solicit începerea imediată a serviciului și iau la cunoștință "
                "că, odată început, pierd dreptul de retragere de 14 zile, în condițiile legii."
            )}},
        )
        if sub.get("stripe_customer_id"):
            params["customer"] = sub["stripe_customer_id"]
        else:
            params["customer_email"] = payer_email
        session = stripe.checkout.Session.create(**params)
        return session.url, ""
    except Exception as e:
        core.log("Stripe Checkout nu a putut fi creat", e)
        return None, "Nu am putut pregăti plata acum. Încearcă din nou."


def create_portal_url(user_id: str) -> tuple[str | None, str]:
    sub = get_subscription(user_id) or {}
    if not sub.get("stripe_customer_id") or not stripe_ready():
        return None, "Nu există un abonament plătit pentru acest cont."
    try:
        s = _stripe().billing_portal.Session.create(customer=sub["stripe_customer_id"], return_url=core.base_url())
        return s.url, ""
    except Exception as e:
        core.log("Portalul Stripe nu a putut fi deschis", e)
        return None, "Nu am putut deschide gestionarea abonamentului."


def send_payment_link_to_parent(user: dict, parent_email: str, plan_code: str) -> tuple[bool, str]:
    url, err = create_checkout_url(user, plan_code, parent_email)
    if not url:
        return False, err
    p = plan(plan_code) or {}
    price = f"{(p.get('price_ron_cents') or 0) / 100:.2f} lei/lună"
    body = (
        f"Bună ziua,\n\n{user.get('name') or user.get('email')} a ales abonamentul „{p.get('name')}” "
        f"({price}) în {core.app_name()}.\n\nPentru elevii minori, plata se face de părinte. "
        f"Puteți plăti în siguranță (Stripe) aici:\n\n{url}\n\nLinkul expiră în 24 de ore. "
        f"Abonamentul se poate anula oricând.\n\n{core.app_name()} — {core.company()['nume']}\n"
    )
    if core.send_email(parent_email, f"Plata abonamentului {core.app_name()}", body):
        return True, f"Am trimis linkul de plată la {parent_email}."
    return False, "Emailul nu a putut fi trimis."


def send_portal_link_to_payer(user_id: str) -> tuple[bool, str]:
    sub = get_subscription(user_id) or {}
    url, err = create_portal_url(user_id)
    if not url:
        return False, err
    to = sub.get("payer_email")
    if not to:
        return False, "Nu știm emailul plătitorului."
    body = (f"Bună ziua,\n\nPuteți gestiona abonamentul {core.app_name()} (card, facturi, anulare) aici:\n\n"
            f"{url}\n\nLinkul expiră în scurt timp.\n")
    if core.send_email(to, f"Gestionarea abonamentului {core.app_name()}", body):
        return True, f"Am trimis linkul la {to}."
    return False, "Emailul nu a putut fi trimis."


def cancel_now(user_id: str):
    """La ștergerea contului: anulează imediat abonamentul Stripe (fără reînnoiri viitoare)."""
    sub = get_subscription(user_id) or {}
    if sub.get("stripe_subscription_id") and core.cfg("STRIPE_SECRET_KEY"):
        try:
            _stripe().Subscription.cancel(sub["stripe_subscription_id"])
        except Exception as e:
            core.log("Anularea abonamentului Stripe a eșuat", e)


def grant_comp(email: str, plan_code: str | None, days: int | None) -> tuple[bool, str]:
    """Admin: acces gratuit (ex. profesori, testeri)."""
    try:
        u = core.db().table("liceu_users").select("id").eq("email", email.strip().lower()).limit(1).execute()
        if not u.data:
            return False, "Nu există cont cu acest email."
        row = {"user_id": u.data[0]["id"], "plan_code": plan_code, "status": "comp",
               "current_period_end": (core.now() + days * 86400) if days else None,
               "trial_ends_at": None, "updated_at": core.now()}
        core.db().table("liceu_subscriptions").upsert(row).execute()
        return True, "Acces acordat."
    except Exception as e:
        core.log("Acordarea accesului gratuit a eșuat", e)
        return False, "Nu a mers."


def status_text(sub: dict | None) -> str:
    if not sub:
        return "fără abonament"
    p = plan(sub.get("plan_code")) or {}
    name = p.get("name") or sub.get("plan_code") or "—"

    def _d(ts):
        return time.strftime("%d.%m.%Y", time.localtime(ts)) if ts else "—"
    s = sub.get("status")
    if s == "trialing" and not sub.get("stripe_subscription_id"):
        return f"Probă gratuită până pe {_d(sub.get('trial_ends_at'))}"
    if s == "comp":
        return f"Acces gratuit ({name})" + (f" până pe {_d(sub.get('current_period_end'))}"
                                            if sub.get("current_period_end") else "")
    if s in ("active", "trialing"):
        extra = " — se oprește la final" if sub.get("cancel_at_period_end") else " — se reînnoiește"
        return f"{name}: activ până pe {_d(sub.get('current_period_end'))}{extra}"
    if s == "past_due":
        return f"{name}: plata reînnoirii a eșuat — actualizați cardul"
    return f"{name}: {s}"
