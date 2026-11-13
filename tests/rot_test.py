# Test cu Supabase și Gemini simulate — rulează prin tests/run_tests.py
import supabase, fake_sb, warnings, logging
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client
import google.genai as genai
CALLS=[]; RULES={}   # (key, model) -> exception text
class Chunk:
    def __init__(s,t): s.text=t; s.usage_metadata=None
class Models:
    def __init__(s,k): s.k=k
    def generate_content_stream(s, model, contents, config):
        CALLS.append((s.k[-1], model, config.thinking_config))
        r=RULES.get((s.k,model)) or RULES.get(("*",model))
        if r: raise Exception(r)
        return iter([Chunk(f"Răspuns de la {model} cu cheia {s.k[-1]}, suficient de lung.")])
class Caches:
    def create(s,**k): raise Exception("caching not supported")
class C:
    def __init__(s,api_key=None): s.models=Models(api_key); s.caches=Caches(); s.files=None
genai.Client=C
from streamlit.testing.v1 import AppTest
ok=lambda c,m: print(("PASS " if c else "FAIL ")+m)
K1,K2="AIzaFAKE1111111111111111","AIzaFAKE2222222222222222"
DAY="429 RESOURCE_EXHAUSTED. Quota exceeded for metric: generate_content_free_tier_requests, limit: 500, quotaId: GenerateRequestsPerDayPerProjectPerModel-FreeTier"
MIN="429 RESOURCE_EXHAUSTED. quotaId: GenerateRequestsPerMinutePerProjectPerModel-FreeTier. Please retry in 36.2s"
at=AppTest.from_file("app_under_test.py", default_timeout=90); at.session_state["materie_selectata"]="fizică_real"; at.session_state["_browser_storage"]={}; at.query_params["sid"]="f"*64  # Liceu: altfel „Automat” întreabă Real/Tehnologic
at.session_state["key_index"]=0; at.run()
def ask(q):
    n=len(CALLS); at.chat_input[0].set_value(q).run(); assert not at.exception,[e.value for e in at.exception]
    return CALLS[n:]
RULES[(K1,"gemini-3.1-flash-lite")]=DAY; RULES[(K1,"gemini-3.5-flash-lite")]=DAY
c=ask("Explică legea lui Ohm pentru un rezistor")
ok([x[1] for x in c][-1]=="gemini-3.6-flash" and c[-1][0]=="1", f"cotă zilnică epuizată pe 2 modele → trece pe 3.6 Flash, aceeași cheie ({[(x[0],x[1]) for x in c]})")
c=ask("Și pentru un circuit serie?")
ok([(x[0],x[1]) for x in c]==[("1","gemini-3.6-flash")], f"al doilea mesaj sare direct peste combinațiile epuizate ({[(x[0],x[1]) for x in c]})")
for m in ["gemini-3.6-flash","gemini-3-flash-preview","gemini-3.8-flash","gemini-2.5-flash","gemini-2.5-flash-lite"]:
    RULES[(K1,m)]=DAY
c=ask("Dar pentru circuit paralel?")
ok(c[-1][:2]==("2","gemini-3.1-flash-lite") and at.session_state["key_index"]==1, f"toate modelele cheii 1 epuizate → rotește pe cheia 2 ({[(x[0],x[1]) for x in c]})")
c=ask("Ce e o diodă Zener?")
ok([(x[0],x[1]) for x in c]==[("2","gemini-3.1-flash-lite")], "rămâne pe cheia 2 fără reîncercări inutile")
RULES[(K2,"gemini-3.1-flash-lite")]=MIN
c=ask("Ce e un tranzistor bipolar?")
ok(c[-1][:2]==("2","gemini-3.5-flash-lite"), f"limită pe minut → modelul următor ({[(x[0],x[1]) for x in c]})")
RULES[("*","gemini-3.5-flash-lite")]="404 NOT_FOUND. models/gemini-3.5-flash-lite is not found for API version v1beta"
for m in ["gemini-3.6-flash","gemini-3-flash-preview","gemini-3.8-flash"]: RULES[(K2,m)]=DAY
c=ask("Ce e un amplificator operațional?")
ok(c[-1][:2]==("2","gemini-2.5-flash"), f"model inexistent (404) e sărit; 2.5 Flash răspunde ({[(x[0],x[1]) for x in c]})")
tc=c[-1][2]
ok(getattr(tc,"thinking_budget",None)==0 and getattr(tc,"thinking_level",None) is None, "Gemini 2.5 primește thinking_budget, nu thinking_level")
ok(getattr(CALLS[0][2],"thinking_level",None) is not None, "Gemini 3.x primește thinking_level")
RULES[(K2,"gemini-2.5-flash")]=DAY; RULES[(K2,"gemini-2.5-flash-lite")]=DAY
import time
st_=None
c=ask("Ce e un filtru trece-jos?")
errs=[e.value for e in at.error]
ok(any("limita" in e or "epuizat" in e.lower() for e in errs), f"totul epuizat → mesaj prietenos ({errs[:1]})")
