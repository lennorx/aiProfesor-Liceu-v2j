# Test cu Supabase și Gemini simulate — rulează prin tests/run_tests.py
import supabase, fake_sb, warnings, logging, itertools
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client
import google.genai as genai
LOG={"uploads":[], "deletes":[], "calls":[]}
_ids=itertools.count(1)
class F:
    def __init__(s,name,mime): s.name=name; s.uri="https://g/"+name; s.mime_type=mime; s.state="ACTIVE"
class Files:
    def upload(s,file,config): f=F(f"files/{next(_ids)}", config.mime_type); LOG["uploads"].append(f.name); return f
    def get(s,name): return F(name,"x")
    def delete(s,name): LOG["deletes"].append(name)
class Chunk:
    def __init__(s,t): s.text=t; s.usage_metadata=None
class Models:
    def generate_content_stream(s, model, contents, config):
        LOG["calls"].append(contents[-1]); return iter([Chunk("Am analizat toate fișierele primite, iată răspunsul.")])
class Caches:
    def create(s,**k): raise Exception("caching not supported")
class C:
    def __init__(s,api_key=None): s.models=Models(); s.files=Files(); s.caches=Caches()
genai.Client=C
from streamlit.testing.v1 import AppTest
ok=lambda c,m: print(("PASS " if c else "FAIL ")+m)
at=AppTest.from_file("app_under_test.py", default_timeout=90); at.session_state["materie_selectata"]="fizică_real"; at.session_state["_browser_storage"]={}; at.query_params["sid"]="d"*64; at.run()  # Liceu: altfel „Automat” întreabă Real/Tehnologic
fu=at.file_uploader(key="_main_file_uploader")
ok(fu.accept_multiple_files, "uploader acceptă mai multe fișiere")
from PIL import Image as _I; import io as _io; _b=_io.BytesIO(); _I.new("RGB",(4,4),"red").save(_b,"PNG"); PNG=_b.getvalue()
fu.set_value([("curs1.txt",b"Legea lui Ohm: U = R*I",'text/plain'),("curs2.txt",b"Kirchhoff: suma curentilor e zero","text/plain"),("schema.png",PNG,"image/png"),("lab.pdf",b"%PDF-1.4 test","application/pdf")]).run()
assert not at.exception,[e.value for e in at.exception]
ok(len(LOG["uploads"])==2, f"2 fișiere media urcate pe Google (au fost {len(LOG['uploads'])})")
ok(at.session_state["_current_uploaded_file_meta"]["type"]=="multi" and len(at.session_state["_current_uploaded_file_meta"]["files"])==4, "metadate pentru 4 fișiere")
at.chat_input[0].set_value("Compară cele două cursuri cu schema").run()
assert not at.exception,[e.value for e in at.exception]
parts=LOG["calls"][-1].parts
fd=[p for p in parts if p.file_data]; txt=" ".join(p.text for p in parts if p.text)
ok(len(fd)==2, "AI primește ambele imagini/PDF")
ok("=== FIȘIER: curs1.txt ===" in txt and "=== FIȘIER: curs2.txt ===" in txt and "Kirchhoff" in txt, "AI primește textul ambelor fișiere text, cu nume")
ok("4 fișiere" in txt, "promptul menționează toate cele 4 fișiere")
n_up=len(LOG["uploads"])
at.chat_input[0].set_value("Și altceva despre ele").run()
ok(len(LOG["uploads"])==n_up, "la al doilea mesaj nu se re-urcă nimic")
ok(len([p for p in LOG["calls"][-1].parts if p.file_data])==2, "al doilea mesaj are tot ambele fișiere")
# eliminare cu 🗑️
at.button(key="_rm_file_schema.png_%d"%len(PNG)).click().run()
assert not at.exception,[e.value for e in at.exception]
ok(len(LOG["deletes"])==1, "imaginea eliminată e ștearsă și de pe Google")
at.chat_input[0].set_value("Acum fără schemă").run()
ok(len([p for p in LOG["calls"][-1].parts if p.file_data])==1, "după eliminare se trimite doar PDF-ul")
ok(len(at.session_state["_current_uploaded_file_meta"]["files"])==3, "rămân 3 fișiere atașate")
# golire uploader
at.file_uploader(key="_main_file_uploader").set_value(None).run()
at.chat_input[0].set_value("Explică legea lui Ohm pentru un rezistor în circuit").run(); print("calls", len(LOG["calls"]))
parts=LOG["calls"][-1].parts
ok(not [p for p in parts if p.file_data] and "FIȘIER" not in " ".join(p.text for p in parts if p.text), "fără fișiere după golirea uploaderului")
# un singur SRT -> formatul vechi (fără antete multi)
at.file_uploader(key="_main_file_uploader").set_value([("film.srt",b"1\n00:00:01,000 --> 00:00:02,000\nHello\n","application/x-subrip")]).run()
ok(at.session_state["_current_uploaded_file_meta"]["name"]=="film.srt", "un singur fișier păstrează comportamentul vechi (SRT)")
assert not at.exception,[e.value for e in at.exception]
