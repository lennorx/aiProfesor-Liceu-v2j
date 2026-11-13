# Test cu Supabase și Gemini simulate — rulează prin tests/run_tests.py
import re, supabase, fake_sb, warnings, logging
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client
import google.genai as genai
CALLS=[]; BAD=set(); DAYX=set()
class Chunk:
    def __init__(s,t): s.text=t; s.usage_metadata=None
class Models:
    def __init__(s,k): s.k=k
    def get(s, model):
        if s.k in BAD: raise Exception("400 INVALID_ARGUMENT API key not valid. API_KEY_INVALID")
        return object()
    def generate_content_stream(s, model, contents, config):
        CALLS.append((s.k, model))
        if s.k in DAYX: raise Exception("429 RESOURCE_EXHAUSTED quotaId: GenerateRequestsPerDayPerProjectPerModel-FreeTier")
        return iter([Chunk("Răspuns de test suficient de lung pentru verificare.")])
class Caches:
    def create(s,**k): raise Exception("caching not supported")
class C:
    def __init__(s,api_key=None): s.models=Models(api_key); s.caches=Caches(); s.files=None
genai.Client=C
from streamlit.testing.v1 import AppTest
DB=fake_sb.DB
ok=lambda c,m: print(("PASS " if c else "FAIL ")+m)
def app(storage):
    at=AppTest.from_file("app_under_test.py", default_timeout=90); at.session_state["materie_selectata"]="fizică_real"; at.session_state["_browser_storage"]=storage; at.run()  # Liceu: altfel „Automat” întreabă Real/Tehnologic
    assert not at.exception,[e.value for e in at.exception]; return at
def btn(at,l): return next(b for b in at.button if b.label==l)
a=app({"sid":"7"*64,"known":[],"token":""})
ok(any("Nicio cheie API" in e.value for e in a.error), "fără chei: aplicația cere o cheie")
a.text_input(key="_auth_reg_email").input("elev@x.ro"); a.text_input(key="_auth_reg_pw").input("parola123"); a.text_input(key="_auth_reg_pw2").input("parola123")
btn(a,"Creează cont").click().run(); assert not a.exception,[e.value for e in a.exception]
MAMA="AIzaMAMA000000000000000000000"; TATA="AIzaTATA000000000000000000000"; RAU="AIzaRAU0000000000000000000000"; BAD.add(RAU)
for lbl,k in (("Mama",MAMA),("Tata",TATA)):
    a.text_input(key="_acct_key_label").input(lbl); a.text_input(key="_acct_key_value").input(k); btn(a,"✅ Adaugă în cont").click().run()
    assert not a.exception,[e.value for e in a.exception]
ok(len(DB["liceu_api_keys"])==2, "2 chei salvate în cont")
ok(all(r["key_enc"].startswith("d1:") and MAMA not in r["key_enc"] and TATA not in r["key_enc"] for r in DB["liceu_api_keys"]), "cheile sunt criptate în DB")
ok({r["label"] for r in DB["liceu_api_keys"]}=={"Mama","Tata"} and {r["key_hint"] for r in DB["liceu_api_keys"]}=={"0000"}, "etichete + ultimele 4 caractere")
a.text_input(key="_acct_key_label").input("Unchiul"); a.text_input(key="_acct_key_value").input(RAU); btn(a,"✅ Adaugă în cont").click().run()
ok(len(DB["liceu_api_keys"])==2 and any("nu e validă" in e.value for e in a.error), "cheie invalidă respinsă (verificare la Google)")
a.text_input(key="_acct_key_label").input("Iar mama"); a.text_input(key="_acct_key_value").input(MAMA); btn(a,"✅ Adaugă în cont").click().run()
ok(len(DB["liceu_api_keys"])==2, "duplicat respins")
ok(not any("Nicio cheie API" in e.value for e in a.error), "cu chei în cont aplicația funcționează")
DAYX.add(MAMA); DAYX.add(TATA) if False else None
a.session_state["key_index"]=0
a.chat_input[0].set_value("Explică legea lui Ohm pentru un rezistor").run(); assert not a.exception,[e.value for e in a.exception]
used=[k for k,_ in CALLS]
ok(MAMA in used and used[-1]==TATA, f"cheia Mamei epuizată → trece automat pe cheia Tatei ({[('MAMA' if k==MAMA else 'TATA') for k in used][-3:]})")
tok=re.search(r'"([^"]+)"',a.session_state["_auth_ls_js"]).group(1)
b=app({"sid":"","known":[],"token":tok})
ok(not any("Nicio cheie API" in e.value for e in b.error), "alt dispozitiv: cheile din cont merg după login automat")
labels=" ".join(m.value for m in b.sidebar.markdown)
ok("Mama" in labels and "Tata" in labels, "lista de chei apare pe al doilea dispozitiv")
kid=DB["liceu_api_keys"][0]["id"]
b.button(key=f"_acct_key_del_{kid}").click().run()
ok(len(DB["liceu_api_keys"])==1, "ștergere cheie din cont")
btn(b,"🚪 Ieșire").click().run()
ok(any("Nicio cheie API" in e.value for e in b.error), "după ieșire cheile contului nu mai sunt folosite")
