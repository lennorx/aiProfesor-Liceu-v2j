"""Modurile specifice Liceu: Simulare BAC (generare + corectare), Corectare temă, Admitere."""
import supabase, fake_sb, warnings, logging, time
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client
import google.genai as genai
N={"c":0}
class Chunk:
    def __init__(s,t): s.text=t; s.usage_metadata=None
class Models:
    def generate_content_stream(s, model, contents, config):
        N["c"]+=1
        return iter([Chunk("SUBIECTUL I (30 de puncte)\n1. Calculați 2+2.\n\nSUBIECTUL al II-lea\n2. x\n\nSUBIECTUL al III-lea\n3. y\n\nBAREM\n1. 4")])
    def generate_content(s, model, contents, config=None):
        N["c"]+=1
        class R: text="ok"
        return R()
class Caches:
    def create(s,**k): raise Exception("caching not supported")
class C:
    def __init__(s,api_key=None): s.models=Models(); s.caches=Caches(); s.files=None
genai.Client=C
from streamlit.testing.v1 import AppTest
ok=lambda c,m: print(("PASS " if c else "FAIL ")+m)
at=AppTest.from_file("app_under_test.py", default_timeout=120); at.session_state["_browser_storage"]={}; at.query_params["sid"]="a"*64; at.run()
ok(not at.exception, "pornire")
b=next(b for b in at.button if "Simulare BAC" in (b.label or "")); b.click().run()
ok(not at.exception, f"intră în BAC {[e.value for e in at.exception][:1]}")
g=next((b for b in at.button if "Generează" in (b.label or "") or "Începe" in (b.label or "")), None)
if g:
    g.click().run(); ok(not at.exception, f"subiect generat {[e.value for e in at.exception][:1]}")
    ok(at.session_state["bac_active"] if "bac_active" in at.session_state else False, "bac activ")
    if "bac_ans_input" in [t.key for t in at.text_area]:
        at.text_area(key="bac_ans_input").input("1. 4").run()
        c=next(b for b in at.button if "Corectare AI" in (b.label or "")); c.click().run()
        ok(not at.exception and at.session_state["bac_corectat"], "corectare BAC")
# homework
at2=AppTest.from_file("app_under_test.py", default_timeout=120); at2.session_state["_browser_storage"]={}; at2.query_params["sid"]="c"*64; at2.run()
b=next((b for b in at2.button if "Corectare temă" in (b.label or "") or "Temă" in (b.label or "")), None)
if b:
    b.click().run(); ok(not at2.exception, "intră în temă")
    at2.text_area(key="hw_text_input").input("2+2=4").run()
    c=next(b for b in at2.button if "Corectează tema" in (b.label or "")); c.click().run()
    ok(not at2.exception and at2.session_state["hw_done"], "corectare temă")
b=next((b for b in at2.button if "Admitere" in (b.label or "")), None)
if b: b.click().run(); ok(not at2.exception, "intră în admitere")
