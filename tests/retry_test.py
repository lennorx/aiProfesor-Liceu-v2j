"""Butonul „🔄 Reîncercați răspunsul” după o eroare de cotă trebuie să refacă cererea."""
import supabase, fake_sb, warnings, logging
logging.disable(logging.WARNING); warnings.filterwarnings("ignore")
supabase.create_client = fake_sb.create_client
import streamlit as st
import google.genai as genai
state = {"fail": True, "calls": 0}
class Chunk:
    def __init__(s, t): s.text = t; s.usage_metadata = None
class FakeModels:
    def generate_content_stream(self, model, contents, config):
        state["calls"] += 1
        if state["fail"]: raise Exception("429 RESOURCE_EXHAUSTED quota")
        return iter([Chunk("Răspuns de test reușit, suficient de lung pentru verificare.")])
class FakeCaches:
    def create(self, **k): raise Exception("caching not supported")
class FakeClient:
    def __init__(s, api_key=None): s.models = FakeModels(); s.caches = FakeCaches(); s.files = None
genai.Client = FakeClient
from streamlit.testing.v1 import AppTest
ok = lambda c, m: print(("PASS " if c else "FAIL ") + m)
at = AppTest.from_file("app_under_test.py", default_timeout=90)
at.query_params["sid"] = "b" * 64; at.session_state["_browser_storage"] = {}
at.session_state["materie_selectata"] = "fizică_real"  # Liceu: altfel „Automat” întreabă Real/Tehnologic
at.run()
at.chat_input[0].set_value("Explică legea lui Ohm la rezistoare").run()
ok(any(b.key == "_retry_after_key_error" for b in at.button), "după eroarea de cotă apare butonul „Reîncercați”")
state["fail"] = False
# golim memoria cotelor epuizate (altfel combinațiile rămân ocolite ~1 minut)
calls_before = state["calls"]
st.cache_resource.clear()
at.button(key="_retry_after_key_error").click().run()
ok(state["calls"] > calls_before, "click pe „Reîncercați” refă cererea")
ok(at.session_state["messages"][-1]["role"] == "assistant", "răspunsul apare în conversație")
ok(not at.exception, "fără excepții")
