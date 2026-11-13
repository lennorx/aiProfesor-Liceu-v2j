"""Webhook-ul Stripe (billing_webhook/main.py) cu Supabase și Stripe simulate."""
import sys, time, types, fake_sb, warnings, logging
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
import supabase
supabase.create_client = fake_sb.create_client

SUBS = {}
fake_stripe = types.ModuleType("stripe")
fake_stripe.api_key = None
class _Sig(Exception): pass
def construct_event(payload, sig, secret):
    if sig != "semnatura-buna":
        raise _Sig("bad")
    import json; return json.loads(payload)
fake_stripe.Webhook = types.SimpleNamespace(construct_event=construct_event)
fake_stripe.Subscription = types.SimpleNamespace(retrieve=lambda sid: SUBS[sid])
sys.modules["stripe"] = fake_stripe

import os
os.environ.update(SUPABASE_URL="x", SUPABASE_KEY="y", STRIPE_WEBHOOK_SECRET="whsec_x")
sys.path.insert(0, ".")
from billing_webhook import main as wh
from fastapi.testclient import TestClient
DB = fake_sb.DB
ok = lambda c, m: print(("PASS " if c else "FAIL ") + m)
DB["liceu_plans"].append({"code": "standard", "stripe_price_id": "price_std", "active": True})
DB["liceu_subscriptions"].append({"user_id": "u1", "plan_code": "proba", "status": "trialing",
                                  "trial_ends_at": time.time() + 100, "updated_at": 0})
cl = TestClient(wh.app)
end = time.time() + 30 * 86400
SUBS["sub_1"] = {"id": "sub_1", "status": "active", "customer": "cus_1", "cancel_at_period_end": False,
                 "metadata": {"user_id": "u1", "plan_code": "standard"},
                 "items": {"data": [{"price": {"id": "price_std"}, "current_period_end": end}]}}
import json
def post(ev, sig="semnatura-buna"):
    return cl.post("/stripe/webhook", content=json.dumps(ev), headers={"Stripe-Signature": sig})

r = post({"id": "evt_0", "type": "x", "data": {"object": {}}}, sig="falsa")
ok(r.status_code == 400, "semnătură invalidă → 400")
ev = {"id": "evt_1", "type": "checkout.session.completed", "data": {"object": {
    "mode": "subscription", "client_reference_id": "u1", "subscription": "sub_1", "created": 1700000000,
    "customer_details": {"email": "mama@test.ro"}, "consent": {"terms_of_service": "accepted"}}}}
r = post(ev)
s = next(x for x in DB["liceu_subscriptions"] if x["user_id"] == "u1")
ok(r.status_code == 200 and s["status"] == "active" and s["plan_code"] == "standard" and s["stripe_subscription_id"] == "sub_1",
   "plata finalizată → abonament activ pe planul corect")
ok(s["payer_email"] == "mama@test.ro" and s["withdrawal_waiver_at"] == 1700000000 and abs(s["current_period_end"] - end) < 1,
   "se salvează plătitorul, renunțarea la retragere și data de expirare (API 2025+)")
ok(len([x for x in DB["liceu_subscriptions"] if x["user_id"] == "u1"]) == 1, "un singur rând per elev")
r = post(ev)
ok(r.json().get("duplicate"), "evenimentul repetat nu e procesat de două ori")
SUBS["sub_1"] = dict(SUBS["sub_1"], status="canceled", metadata={})
post({"id": "evt_2", "type": "customer.subscription.deleted", "data": {"object": SUBS["sub_1"]}})
s = next(x for x in DB["liceu_subscriptions"] if x["user_id"] == "u1")
ok(s["status"] == "canceled" and s["payer_email"] == "mama@test.ro", "anulare → status canceled (plătitorul rămâne)")
ok(cl.get("/stripe/health").json() == {"ok": True}, "health check")
