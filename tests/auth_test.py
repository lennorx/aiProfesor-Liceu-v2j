# Test cu Supabase și Gemini simulate — rulează prin tests/run_tests.py
import re, supabase, fake_sb, warnings, logging
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client
from streamlit.testing.v1 import AppTest
DB = fake_sb.DB
S1, S2 = "1"*64, "2"*64
ok = lambda c, m: print(("PASS " if c else "FAIL ") + m)

def app(storage, sid_url=None):
    at = AppTest.from_file("app_under_test.py", default_timeout=90)
    at.session_state["_browser_storage"] = storage
    if sid_url: at.query_params["sid"] = sid_url
    at.run(); assert not at.exception, [e.value for e in at.exception]
    return at
def btn(at, label): return next(b for b in at.button if b.label == label)

# A. device 1, anonim, cu un mesaj
a = app({"sid": S1, "known": [S1], "token": ""})
ok(a.session_state["session_id"] == S1, "device1 restaurează SID din browser")
fake_sb.DB["history"].append({"id": 999, "session_id": S1, "role": "user", "content": "Mesaj pe device 1", "timestamp": __import__("time").time(), "app_id": "test_local"})
a.run()
a.text_input(key="_auth_reg_email").input("Student@Test.ro")
a.text_input(key="_auth_reg_pw").input("parola123")
a.text_input(key="_auth_reg_pw2").input("parola123")
btn(a, "Creează cont").click().run()
assert not a.exception, [e.value for e in a.exception]
u = a.session_state["_auth_user"]; uid = u["id"]
ok(u["email"] == "student@test.ro", "cont creat, email normalizat")
ok(DB["liceu_users"][0]["password_hash"].startswith("scrypt$") and "parola123" not in DB["liceu_users"][0]["password_hash"], "parola stocată ca hash scrypt")
ok(next(s for s in DB["sessions"] if s["session_id"] == S1).get("liceu_user_id") == uid, "conversația anonimă a fost mutată în cont")
js = a.session_state["_auth_ls_js"]; tok = re.search(r'"([^"]+)"', js).group(1)
ok(len(DB["liceu_tokens"]) == 1 and DB["liceu_tokens"][0]["token_hash"] != tok, "token creat; în DB doar hash-ul")

# B. device 2: token din localStorage -> logat automat, vede conversația de pe device 1
b = app({"sid": S2, "known": [S2], "token": tok})
ok((b.session_state["_auth_user"] or {}).get("id") == uid, "device2 logat automat din token")
side = " ".join(str(x.label) for x in b.sidebar.button)
ok("Mesaj pe device 1" in side, "device2 vede în listă conversația de pe device1")
ok(next(s for s in DB["sessions"] if s["session_id"] == S2).get("liceu_user_id") == uid, "conversația nouă de pe device2 aparține contului")

# C. device 3 anonim cu SID-ul contului (link copiat) -> NU primește conversația
c = app({"sid": S1, "known": [S1], "token": ""})
ok(c.session_state["session_id"] != S1, "anonim cu SID-ul unui cont primește conversație nouă")
c2 = app({"sid": "", "known": [], "token": ""}, sid_url=S1)
ok(c2.session_state["session_id"] != S1, "anonim cu ?sid= al unui cont în URL primește conversație nouă")
ok(not any(m["content"] == "Mesaj pe device 1" for m in c2.session_state["messages"]), "istoricul contului nu se încarcă pentru anonim")

# D. login greșit / corect pe device 3
c.text_input(key="_auth_login_email").input("student@test.ro")
c.text_input(key="_auth_login_pw").input("gresit000")
btn(c, "Intră").click().run()
ok(c.session_state["_auth_user"] if "_auth_user" in c.session_state else None is None and any("greșită" in e.value for e in c.error), "parolă greșită refuzată")
c.text_input(key="_auth_login_pw").input("parola123")
btn(c, "Intră").click().run()
ok(c.session_state["_auth_user"]["id"] == uid, "login corect pe device3")
ok(len(DB["liceu_tokens"]) == 2, "al doilea dispozitiv are tokenul lui")

# E. logout pe device 2
btn(b, "🚪 Ieșire").click().run()
assert not b.exception, [e.value for e in b.exception]
ok("_auth_user" not in b.session_state or b.session_state["_auth_user"] is None, "device2 delogat")
ok(len(DB["liceu_tokens"]) == 1, "tokenul device2 șters din DB")
ok("removeItem" in b.session_state["_auth_ls_js"], "tokenul se șterge din browser")
ok(b.session_state["session_id"] not in (S1, S2), "după ieșire pornește o conversație anonimă nouă")

# F. token revocat nu mai loghează
f = app({"sid": "", "known": [], "token": tok})
ok(("_auth_user" not in f.session_state) or not f.session_state["_auth_user"], "token revocat nu mai loghează")

# G. brute force
g = app({"sid": "", "known": [], "token": ""})
for i in range(6):
    g.text_input(key="_auth_login_email").input("student@test.ro"); g.text_input(key="_auth_login_pw").input("x"*9)
    btn(g, "Intră").click().run()
ok(any("Prea multe" in e.value for e in g.error), "blocare după 5 încercări greșite")

# H. ieșire peste tot
btn(c, "Ieșire peste tot").click().run()
ok(len(DB["liceu_tokens"]) == 0, "„Ieșire peste tot” șterge toate tokenurile")
