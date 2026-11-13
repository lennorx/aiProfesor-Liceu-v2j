"""Webhook Stripe pentru Profesor Liceu (rulează pe serverul Hetzner, lângă aplicația Streamlit).

Stripe trimite aici evenimentele de plată; serviciul actualizează tabelul liceu_subscriptions.
Variabile de mediu: STRIPE_SECRET_KEY, STRIPE_WEBHOOK_SECRET, SUPABASE_URL, SUPABASE_KEY.
Pornire locală:  uvicorn billing_webhook.main:app --port 8001
URL în Stripe (Developers → Webhooks):  https://DOMENIU/stripe/webhook
Evenimente: checkout.session.completed, customer.subscription.created, customer.subscription.updated,
            customer.subscription.deleted, invoice.paid, invoice.payment_failed
"""
from __future__ import annotations

import logging
import os
import time

import stripe
from fastapi import FastAPI, Header, HTTPException, Request
from supabase import create_client

log = logging.getLogger("billing_webhook")
logging.basicConfig(level=logging.INFO)

stripe.api_key = os.environ.get("STRIPE_SECRET_KEY", "")
WEBHOOK_SECRET = os.environ.get("STRIPE_WEBHOOK_SECRET", "")
_sb = None


def db():
    global _sb
    if _sb is None:
        _sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])
    return _sb


app = FastAPI(title="Profesor Liceu — Stripe webhook", docs_url=None, redoc_url=None)


@app.get("/stripe/health")
def health():
    return {"ok": True}


def _get(obj, key, default=None):
    try:
        return obj.get(key, default) if isinstance(obj, dict) else getattr(obj, key, default)
    except Exception:
        return default


def _period_end(sub) -> float | None:
    """current_period_end: pe abonament (API vechi) sau pe primul item (API 2025+)."""
    v = _get(sub, "current_period_end")
    if not v:
        items = _get(_get(sub, "items") or {}, "data") or []
        if items:
            v = _get(items[0], "current_period_end")
    return float(v) if v else None


def _plan_for_price(price_id: str | None) -> str | None:
    if not price_id:
        return None
    res = db().table("liceu_plans").select("code").eq("stripe_price_id", price_id).limit(1).execute()
    return res.data[0]["code"] if res.data else None


def _user_for_subscription(sub) -> str | None:
    uid = (_get(sub, "metadata") or {}).get("user_id")
    if uid:
        return uid
    res = db().table("liceu_subscriptions").select("user_id") \
        .eq("stripe_subscription_id", _get(sub, "id")).limit(1).execute()
    return res.data[0]["user_id"] if res.data else None


def _sync_subscription(sub, user_id: str | None = None, extra: dict | None = None):
    user_id = user_id or _user_for_subscription(sub)
    if not user_id:
        log.warning("abonament %s fără user_id — ignorat", _get(sub, "id"))
        return
    items = _get(_get(sub, "items") or {}, "data") or []
    price_id = _get(_get(items[0], "price") or {}, "id") if items else None
    row = {
        "user_id": user_id,
        "status": _get(sub, "status"),
        "current_period_end": _period_end(sub),
        "cancel_at_period_end": bool(_get(sub, "cancel_at_period_end")),
        "stripe_customer_id": _get(sub, "customer"),
        "stripe_subscription_id": _get(sub, "id"),
        "trial_ends_at": None,
        "updated_at": time.time(),
    }
    plan = _plan_for_price(price_id) or (_get(sub, "metadata") or {}).get("plan_code")
    if plan:
        row["plan_code"] = plan
    row.update(extra or {})
    db().table("liceu_subscriptions").upsert(row).execute()
    log.info("abonament %s → %s (%s)", row["stripe_subscription_id"], row["status"], user_id)


def _already_processed(event_id: str) -> bool:
    return bool(db().table("liceu_stripe_events").select("id").eq("id", event_id).limit(1).execute().data)


def _mark_processed(event_id: str, etype: str):
    # marcat DUPĂ procesare: dacă procesarea eșuează, Stripe retrimite evenimentul
    db().table("liceu_stripe_events").upsert({"id": event_id, "type": etype, "received_at": time.time()}).execute()


@app.post("/stripe/webhook")
async def webhook(request: Request, stripe_signature: str = Header(None, alias="Stripe-Signature")):
    payload = await request.body()
    try:
        event = stripe.Webhook.construct_event(payload, stripe_signature, WEBHOOK_SECRET)
    except Exception:
        raise HTTPException(status_code=400, detail="semnătură invalidă")

    etype, obj = event["type"], event["data"]["object"]
    if _already_processed(event["id"]):
        return {"ok": True, "duplicate": True}

    if etype == "checkout.session.completed" and _get(obj, "mode") == "subscription":
        user_id = _get(obj, "client_reference_id") or (_get(obj, "metadata") or {}).get("user_id")
        sub = stripe.Subscription.retrieve(_get(obj, "subscription"))
        details = _get(obj, "customer_details") or {}
        consent = _get(obj, "consent") or {}
        extra = {"payer_email": _get(details, "email")}
        if _get(consent, "terms_of_service") == "accepted":
            # plătitorul a bifat Termenii + începerea imediată (renunțarea la retragere)
            extra["withdrawal_waiver_at"] = float(_get(obj, "created") or time.time())
        _sync_subscription(sub, user_id, extra)
    elif etype in ("customer.subscription.created", "customer.subscription.updated",
                   "customer.subscription.deleted"):
        _sync_subscription(obj)
    elif etype in ("invoice.paid", "invoice.payment_failed"):
        sub_id = _get(obj, "subscription") or _get(_get(_get(obj, "parent") or {}, "subscription_details") or {},
                                                   "subscription")
        if sub_id:
            _sync_subscription(stripe.Subscription.retrieve(sub_id))
    _mark_processed(event["id"], etype)
    return {"ok": True}
