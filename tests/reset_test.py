# Test cu Supabase și Gemini simulate — rulează prin tests/run_tests.py
import re, time, smtplib, supabase, fake_sb, warnings, logging
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client
SENT=[]
class FakeSMTP:
    def __init__(s,h,p,timeout=None): s.h=h
    def __enter__(s): return s
    def __exit__(s,*a): pass
    def login(s,u,p): assert p=="abcdefghijklmnop", "spațiile din parola de aplicație trebuie eliminate"
    def send_message(s,m): SENT.append(m)
smtplib.SMTP_SSL=FakeSMTP
from streamlit.testing.v1 import AppTest
DB=fake_sb.DB
ok=lambda c,m: print(("PASS " if c else "FAIL ")+m)
def app(storage):
    at=AppTest.from_file("app_under_test.py", default_timeout=90); at.session_state["_browser_storage"]=storage; at.run()
    assert not at.exception,[e.value for e in at.exception]; return at
def btn(at,l): return next(b for b in at.button if b.label==l)
def fill(at, **kv):
    for k,v in kv.items(): at.text_input(key=k).input(v)

# creare cont -> cod de recuperare afișat o dată
a=app({"sid":"3"*64,"known":[],"token":""})
fill(a,_auth_reg_email="elev@test.ro",_auth_reg_pw="parola123",_auth_reg_pw2="parola123"); btn(a,"Creează cont").click().run()
rec=a.session_state["_auth_new_recovery"]
ok(re.fullmatch(r"[A-Z2-9]{4}-[A-Z2-9]{4}-[A-Z2-9]{4}",rec) is not None, "cod de recuperare generat la creare")
ok(any(rec in c.value for c in a.sidebar.code), "codul e afișat în bară")
btn(a,"✅ Am notat codul").click().run()
ok("_auth_new_recovery" not in a.session_state and not a.sidebar.code, "codul dispare după „Am notat”")
tok=re.search(r'"([^"]+)"',a.session_state["_auth_ls_js"]).group(1)

# resetare prin email
b=app({"sid":"","known":[],"token":""})
fill(b,_auth_reset_mail_email="elev@test.ro"); btn(b,"📧 Trimite codul").click().run()
ok(len(SENT)==1 and SENT[0]["To"]=="elev@test.ro", "email trimis")
code=re.search(r"\b(\d{6})\b",SENT[0].get_content()).group(1)
ok(DB["liceu_password_resets"][0]["code_hash"]!=code, "în DB doar hash-ul codului")
fill(b,_auth_reset_mail_email="elev@test.ro"); btn(b,"📧 Trimite codul").click().run()
ok(len(SENT)==1 and any("deja" in e.value for e in b.error), "al doilea email în <60s e refuzat")
b2=app({"sid":"","known":[],"token":""})
fill(b2,_auth_reset_mail_email="nimeni@test.ro"); btn(b2,"📧 Trimite codul").click().run()
ok(len(SENT)==1 and any("Dacă există" in s.value for s in b2.success), "email inexistent: același mesaj, niciun email")
fill(b,_auth_reset_email="elev@test.ro",_auth_reset_code="000000" if code!="000000" else "111111",_auth_reset_pw="parolanoua1",_auth_reset_pw2="parolanoua1")
btn(b,"Schimbă parola").click().run()
ok(any("greșit" in e.value for e in b.error), "cod greșit refuzat")
fill(b,_auth_reset_code=code); btn(b,"Schimbă parola").click().run()
assert not b.exception,[e.value for e in b.exception]
ok((b.session_state["_auth_user"] or {}).get("email")=="elev@test.ro", "parolă schimbată + logat")
ok(not DB["liceu_password_resets"], "codul de email consumat")
ok(len(DB["liceu_tokens"])==1, "celelalte dispozitive deconectate (doar tokenul nou)")
f=app({"sid":"","known":[],"token":tok})
ok(not f.session_state["_auth_user"] if "_auth_user" in f.session_state else True, "tokenul vechi nu mai merge")

# login cu parola nouă / veche
g=app({"sid":"","known":[],"token":""})
fill(g,_auth_login_email="elev@test.ro",_auth_login_pw="parola123"); btn(g,"Intră").click().run()
ok("_auth_user" not in g.session_state or not g.session_state["_auth_user"], "parola veche nu mai merge")
fill(g,_auth_login_pw="parolanoua1"); btn(g,"Intră").click().run()
ok(g.session_state["_auth_user"]["email"]=="elev@test.ro", "parola nouă merge")

# resetare cu codul de recuperare (cu litere mici și fără liniuțe)
h=app({"sid":"","known":[],"token":""})
fill(h,_auth_reset_email="ELEV@test.ro",_auth_reset_code=rec.replace("-","").lower(),_auth_reset_pw="altaparola9",_auth_reset_pw2="altaparola9")
btn(h,"Schimbă parola").click().run()
ok(h.session_state["_auth_user"]["email"]=="elev@test.ro", "resetare cu codul de recuperare")
rec2=h.session_state.get("_auth_new_recovery")
ok(rec2 and rec2!=rec, "cod de recuperare nou generat (cel vechi e consumat)")
i=app({"sid":"","known":[],"token":""})
fill(i,_auth_reset_email="elev@test.ro",_auth_reset_code=rec,_auth_reset_pw="inca1parola",_auth_reset_pw2="inca1parola")
btn(i,"Schimbă parola").click().run()
ok(any("greșit" in e.value for e in i.error), "codul de recuperare vechi nu mai merge")

# cod nou din cont
btn(h,"✅ Am notat codul").click().run()
btn(h,"Generează cod nou").click().run()
ok(h.session_state.get("_auth_new_recovery") not in (None, rec2), "generare cod de recuperare nou din cont")
