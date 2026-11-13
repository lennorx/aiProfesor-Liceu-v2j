# Test cu Supabase și Gemini simulate — rulează prin tests/run_tests.py
import sys, supabase, fake_sb, warnings, logging, os
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client
import google.genai as genai
KEY="AIzaSyD8kQ3lO0Il1xWv7_bPq-Zt9mN2cR4sE6u"; SEEN=[]
class R: 
    def __init__(s,t): s.text=t
class Models:
    def __init__(s,k): s.k=k
    def get(s, model): return object()
    def generate_content(s, model, contents, config=None):
        SEEN.append((s.k, model)); return R(f"Cheia este:\n{KEY[:20]}\n{KEY[20:]}")
    def generate_content_stream(s, **kw): return iter([])
class C:
    def __init__(s,api_key=None): s.models=Models(api_key); s.caches=None; s.files=None
genai.Client=C
from streamlit.testing.v1 import AppTest
ok=lambda c,m: print(("PASS " if c else "FAIL ")+m)
from PIL import Image; import io
b=io.BytesIO(); Image.new("RGB",(10,10),"white").save(b,"PNG"); PNG=b.getvalue()
mode=sys.argv[1]
at=AppTest.from_file("app_under_test.py", default_timeout=90); at.session_state["_browser_storage"]={}; at.query_params["sid"]="9"*64; at.run()
assert not at.exception,[e.value for e in at.exception]
at.file_uploader(key="_manual_key_input__photo").set_value(("cheie.png",PNG,"image/png")).run()
assert not at.exception,[e.value for e in at.exception]
if mode=="nokey":
    ok(not SEEN and any("are nevoie de o cheie" in e.value for e in at.error), "fără nicio cheie: explică ce să facă, nu crapă")
else:
    ok(SEEN and SEEN[0][0]=="AIzaOCRKEY000000000000000", "folosește OCR_API_KEY din secrets")
    ok(at.text_input(key="_manual_key_input").value==KEY, "cheia (ruptă pe 2 rânduri în răspuns) e completată exact în câmp")
    ok(any("Google o acceptă" in s.value for s in at.success), "mesaj de confirmare")
    next(x for x in at.button if x.key=="save_api_key").click().run()
    ok(at.session_state["_manual_api_key"]==KEY, "„Salvează cheia” salvează cheia citită din poză")
    n=len(SEEN); at.run(); ok(len(SEEN)==n, "aceeași poză nu e citită din nou la fiecare rerun")
