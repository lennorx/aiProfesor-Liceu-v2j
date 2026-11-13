"""Versiunea comercială (BILLING_MODE = "on"): vârstă, acord parental, documente legale, probă, limite,
raportare, export/ștergere, plată Stripe (simulată). Rulează prin tests/run_tests.py."""
import re, sys, time, types, smtplib, supabase, fake_sb, warnings, logging
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client

SENT = []
class FakeSMTP:
    def __init__(s, h, p, timeout=None): pass
    def __enter__(s): return s
    def __exit__(s, *a): pass
    def login(s, u, p): pass
    def send_message(s, m): SENT.append(m)
smtplib.SMTP_SSL = FakeSMTP

# Stripe simulat
STRIPE_CALLS = []
class _Obj(dict):
    __getattr__ = dict.get
fake_stripe = types.ModuleType("stripe")
fake_stripe.api_key = None
fake_stripe.checkout = types.SimpleNamespace(Session=types.SimpleNamespace(
    create=lambda **kw: (STRIPE_CALLS.append(("checkout", kw)), _Obj(url="https://checkout.stripe.test/c/s_123"))[1]))
fake_stripe.billing_portal = types.SimpleNamespace(Session=types.SimpleNamespace(
    create=lambda **kw: (STRIPE_CALLS.append(("portal", kw)), _Obj(url="https://billing.stripe.test/p/1"))[1]))
fake_stripe.Subscription = types.SimpleNamespace(cancel=lambda sid: STRIPE_CALLS.append(("cancel", sid)))
sys.modules["stripe"] = fake_stripe

# Gemini simulat
import google.genai as genai
class Chunk:
    def __init__(s, t): s.text = t; s.usage_metadata = None
class Models:
    def generate_content_stream(s, model, contents, config):
        return iter([Chunk("Răspuns de test de la profesor, suficient de lung pentru verificare.")])
class Caches:
    def create(s, **k): raise Exception("caching not supported")
class C:
    def __init__(s, api_key=None): s.models = Models(); s.caches = Caches(); s.files = None
genai.Client = C

import streamlit as st
from streamlit.testing.v1 import AppTest
import comercial
DB = fake_sb.DB
YEAR = time.localtime().tm_year
ok = lambda c, m: print(("PASS " if c else "FAIL ") + m)

DB["liceu_plans"] += [
    {"code": "proba", "name": "Perioadă de probă", "price_ron_cents": 0, "stripe_price_id": None,
     "daily_msg_limit": 2, "model_tier": "standard", "active": True, "sort": 0},
    {"code": "standard", "name": "Standard", "price_ron_cents": 2900, "stripe_price_id": "price_std",
     "daily_msg_limit": 100, "model_tier": "standard", "active": True, "sort": 1},
]


def app(storage=None, qp=None):
    at = AppTest.from_file("app_under_test.py", default_timeout=90)
    at.session_state["_browser_storage"] = storage or {"sid": "", "known": [], "token": ""}
    at.session_state["materie_selectata"] = "matematică"
    for k, v in (qp or {}).items():
        at.query_params[k] = v
    at.run(); assert not at.exception, [e.value for e in at.exception]
    return at
def btn(at, label): return next(b for b in at.button if b.label == label)
def run(at): at.run(); assert not at.exception, [e.value for e in at.exception]; return at
def register(at, email, by, bm, parent="", terms=False):
    at.text_input(key="_auth_reg_email").input(email)
    at.text_input(key="_auth_reg_pw").input("parola123"); at.text_input(key="_auth_reg_pw2").input("parola123")
    at.selectbox(key="_auth_reg_by").set_value(by); at.selectbox(key="_auth_reg_bm").set_value(bm)
    at.text_input(key="_auth_reg_parent").input(parent)
    at.checkbox(key="_auth_reg_terms").set_value(terms)
    btn(at, "Creează cont").click(); return run(at)
def user(email): return next((u for u in DB["liceu_users"] if u["email"] == email), None)
def ask(at, q):
    at.chat_input[0].set_value(q); return run(at)

# 1. fără cont: ecranul de acces, fără chat
a = app()
ok(not a.chat_input and any("ai nevoie de un cont" in i.value for i in a.info), "fără cont → nu se poate folosi chat-ul")
ok(any("asistent bazat pe inteligență artificială" in c.value for c in a.caption), "mesajul „vorbești cu un AI” e afișat")

# 2. elev sub 16 ani → acord parental
a = register(a, "mic@test.ro", YEAR - 14, 1, parent="mama@test.ro")
u = user("mic@test.ro")
ok(u and u["parental_status"] == "pending" and u["parent_email"] == "mama@test.ro", "cont sub 16 ani: așteaptă acordul părintelui")
ok(len(SENT) == 1 and SENT[-1]["To"] == "mama@test.ro", "emailul a plecat la părinte")
ok(not a.chat_input and any("acordul" in w.value for w in a.warning), "până la acord, chat-ul e blocat")
tok = re.search(r"\?parinte=([\w-]+)", SENT[-1].get_content()).group(1)
ok(DB["liceu_parent_consents"][0]["token_hash"] != tok, "în DB doar hash-ul linkului")

# 3. pagina părintelui
p = app(qp={"parinte": tok})
ok(any("Acordul părintelui" in t.value for t in p.title), "pagina pentru părinte se deschide din link")
btn(p, "✅ Îmi dau acordul").click(); run(p)
ok(user("mic@test.ro")["parental_status"] == "pending" and p.error, "fără bifele obligatorii acordul nu e salvat")
for cb in p.checkbox: cb.check()
btn(p, "✅ Îmi dau acordul").click(); run(p)
ok(user("mic@test.ro")["parental_status"] == "granted", "acordul părintelui activează contul")
acc = [r for r in DB["liceu_legal_acceptances"] if r["user_id"] == u["id"]]
ok(len(acc) == 2 and all(r["accepted_by"] == "parinte" and r["actor_email"] == "mama@test.ro" for r in acc),
   "acceptarea Termenilor + Confidențialității e salvată în numele părintelui, cu versiunea")

# 4. elevul: proba pornește, limita zilnică (2) se aplică
a.session_state["_auth_user_ts"] = 0; run(a)
ok(bool(a.chat_input), "după acord, elevul poate folosi profesorul")
sub = next(s for s in DB["liceu_subscriptions"] if s["user_id"] == u["id"])
ok(sub["status"] == "trialing" and sub["trial_ends_at"] > time.time() + 6 * 86400, "perioada de probă (7 zile) a pornit")
ask(a, "Cât face 2 + 2?"); ask(a, "Dar 3 + 3?")
ok(DB["liceu_usage"][0]["msg_count"] == 2, "întrebările sunt numărate pe zi")
ask(a, "Și 4 + 4?")
ok(any("toate cele 2 întrebări" in w.value for w in a.warning), "a treia întrebare peste limita planului e refuzată")

# 5. raportarea unui răspuns
k = next(s.key for s in a.selectbox if (s.key or "").startswith("_rep_r_"))
idx = k.split("_")[-1]
a.selectbox(key=k).set_value("Informație greșită"); a.text_area(key=f"_rep_c_{idx}").input("2+2 nu e 5")
a.button(key=f"_rep_s_{idx}").click(); run(a)
ok(len(DB["liceu_reports"]) == 1 and DB["liceu_reports"][0]["reason"] == "Informație greșită", "raportarea e salvată")

# 6. proba expirată → abonament
sub["trial_ends_at"] = time.time() - 10; a.session_state["_auth_user_ts"] = 0; run(a)
ok(not a.chat_input and any("Perioada de probă s-a încheiat" in i.value for i in a.info), "după probă: cere abonament")
st.cache_data.clear()
a.text_input(key="_sub_parent").input("mama@test.ro")
btn(a, "📧 Trimite linkul de plată părintelui").click(); run(a)
co = [c for c in STRIPE_CALLS if c[0] == "checkout"][-1][1]
ok(co["customer_email"] == "mama@test.ro" and co["client_reference_id"] == u["id"], "minor: linkul de plată e pentru părinte")
ok(co["consent_collection"] == {"terms_of_service": "required"} and "dreptul de retragere" in
   co["custom_text"]["terms_of_service_acceptance"]["message"], "la plată se bifează Termenii + renunțarea la retragere")
ok("checkout.stripe.test" in SENT[-1].get_content() and SENT[-1]["To"] == "mama@test.ro", "emailul cu linkul de plată a plecat la părinte")

# 7. elev de 17 ani: emailul părintelui obligatoriu, acceptă singur documentele
b = app()
b = register(b, "mare@test.ro", YEAR - 17, 1, parent="", terms=True)
ok(user("mare@test.ro") is None and any("emailul unui părinte" in e.value for e in b.error), "sub 18 ani: fără emailul părintelui nu se creează contul")
b = register(b, "mare@test.ro", YEAR - 17, 1, parent="tata@test.ro", terms=False)
ok(user("mare@test.ro") is None and any("accepți Termenii" in e.value for e in b.error), "peste 16 ani: Termenii trebuie acceptați")
b = register(b, "mare@test.ro", YEAR - 17, 1, parent="tata@test.ro", terms=True)
u2 = user("mare@test.ro")
ok(u2 and u2["parental_status"] == "not_required" and bool(b.chat_input), "17 ani: cont activ imediat")
ok(sum(1 for r in DB["liceu_legal_acceptances"] if r["user_id"] == u2["id"] and r["accepted_by"] == "elev") == 2,
   "acceptarea elevului e salvată")

# 8. documente legale actualizate → acceptare din nou
comercial.core.LEGAL_VERSIONS["termeni"] = "2099-01-01"
b.session_state["_auth_user_ts"] = 0; run(b)
ok(not b.chat_input and any("Am actualizat" in w.value for w in b.warning), "versiune nouă a Termenilor → se cere acceptarea")
b.checkbox(key="_gate_legal_chk").check(); run(b); b.button(key="_gate_legal_btn").click(); run(b)
ok(bool(b.chat_input), "după acceptare, accesul revine")

# 9. adult: plătește singur
c = app()
c = register(c, "adult@test.ro", YEAR - 25, 1, terms=True)
u3 = user("adult@test.ro")
next(s for s in DB["liceu_subscriptions"] if s["user_id"] == u3["id"])["trial_ends_at"] = time.time() - 10
c.session_state["_auth_user_ts"] = 0; run(c)
btn(c, "💳 Continuă la plată").click(); run(c)
co = [x for x in STRIPE_CALLS if x[0] == "checkout"][-1][1]
ok(co["customer_email"] == "adult@test.ro" and any(l.url and "checkout.stripe.test" in l.url for l in c.get("link_button")),
   "adult: pagina de plată Stripe se deschide direct")

# 10. export + ștergere (GDPR)
btn(b, "📦 Pregătește exportul").click(); run(b)
import json
exp = json.loads(b.session_state["_gdpr_export"])
ok(exp["cont"]["email"] == "mare@test.ro" and len(exp["acceptari_documente"]) >= 2, "exportul conține contul și acordurile")
b.text_input(key="_gdpr_pw").input("gresita"); b.text_input(key="_gdpr_conf").input("ȘTERGE")
btn(b, "🗑️ Șterge definitiv contul").click(); run(b)
ok(user("mare@test.ro") is not None and any("Parola" in e.value for e in b.error), "ștergerea cere parola corectă")
b.text_input(key="_gdpr_pw").input("parola123"); b.text_input(key="_gdpr_conf").input("ȘTERGE")
btn(b, "🗑️ Șterge definitiv contul").click(); run(b)
ok(user("mare@test.ro") is None and not any(r["user_id"] == u2["id"] for r in DB["liceu_legal_acceptances"])
   and not any(s.get("liceu_user_id") == u2["id"] for s in DB["sessions"]), "contul și datele lui au fost șterse")

# 11. admin
user("adult@test.ro")["role"] = "admin"
c.session_state["_auth_user_ts"] = 0; run(c)
ok(bool(c.chat_input), "adminul are acces fără abonament")
c.text_input(key="_adm_comp_email").input("mic@test.ro"); c.button(key="_adm_comp_btn").click(); run(c)
ok(next(s for s in DB["liceu_subscriptions"] if s["user_id"] == u["id"])["status"] == "comp", "adminul poate acorda acces gratuit")
ok(any("Informație greșită" in m.value for m in c.markdown), "adminul vede raportările")

# 12. pagini publice cu documentele
d = app(qp={"doc": "termeni"})
ok(any("Termeni și condiții" in m.value for m in d.markdown), "?doc=termeni afișează Termenii")
