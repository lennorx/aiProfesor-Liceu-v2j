import streamlit as st
import json
import streamlit.components.v1 as components
from google import genai
from google.genai import types as genai_types
from supabase import create_client, Client
import time
import tempfile
import os
import random
import re
import hashlib
import hmac
import smtplib
from email.message import EmailMessage
import secrets
import threading
from collections import defaultdict
import pathlib

import comercial

# Pod bidirecțional Streamlit <-> localStorage pentru cheia API a elevului.
# Fișierul key_bridge/index.html trebuie pus în repo, lângă acest script.
_KEY_BRIDGE_DIR = pathlib.Path(__file__).resolve().parent / "key_bridge"
_key_bridge = (
    components.declare_component("profesor_key_bridge", path=str(_KEY_BRIDGE_DIR))
    if (_KEY_BRIDGE_DIR / "index.html").exists() else None
)

# Pod localStorage -> Python pentru ID-ul conversației (restaurarea istoricului pe orice browser).
# Fișierul session_bridge/index.html trebuie pus în repo, lângă acest script.
_SESSION_BRIDGE_DIR = pathlib.Path(__file__).resolve().parent / "session_bridge"
_session_bridge = (
    components.declare_component("profesor_session_bridge", path=str(_SESSION_BRIDGE_DIR))
    if (_SESSION_BRIDGE_DIR / "index.html").exists() else None
)

# === IMPORTURI PENTRU TIPURI NOI DE FIȘIERE ===
# python-docx pentru .docx/.doc
try:
    from docx import Document as _DocxDocument
    _DOCX_AVAILABLE = True
except ImportError:
    _DOCX_AVAILABLE = False

# dbfread pentru .dbf
try:
    from dbfread import DBF as _DBF
    _DBF_AVAILABLE = True
except ImportError:
    _DBF_AVAILABLE = False




# === EXTRAGERE TEXT DIN FIȘIERE (txt, docx, doc, dbf, srt) ===

# Tipuri de fișiere text acceptate suplimentar (nu pot fi trimise la Google Files API direct)
_TEXT_FILE_TYPES = {
    "text/plain":           ".txt",
    "text/x-srt":          ".srt",
    "application/x-subrip": ".srt",
    # .docx
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    # .doc vechi
    "application/msword":   ".doc",
    # .dbf — nu are MIME standard, îl detectăm după extensie
    "application/dbf":      ".dbf",
    "application/dbase":    ".dbf",
    "application/x-dbase":  ".dbf",
}

# Extensii acceptate pentru fișierele text (folosite în st.file_uploader)
_TEXT_FILE_EXTENSIONS = ["txt", "srt", "docx", "doc", "dbf"]


def _extract_text_from_uploaded_file(uploaded_file) -> str | None:
    """Extrage conținutul text dintr-un fișier text/docx/doc/dbf/srt.

    Returnează textul extras ca string, sau None dacă extragerea eșuează.
    Fișierele prea mari sunt trunchiate la MAX_TEXT_CHARS pentru a nu depăși
    fereastra de context a modelului.
    """
    MAX_TEXT_CHARS = 400_000  # ~100k tokeni — acoperă subtitrări complete și documente mari

    fname = uploaded_file.name.lower()
    ftype = (uploaded_file.type or "").lower()
    raw_bytes = uploaded_file.getvalue()

    # ── PDF mare: text extras local (pypdf) ──
    if fname.endswith(".pdf"):
        txt = _extract_pdf_text_pypdf(raw_bytes, MAX_TEXT_CHARS)
        return txt or ("⚠️ PDF-ul nu conține text care poate fi extras (pare scanat). "
                       f"Micșorează-l sub {_pdf_text_threshold_bytes() // (1024 * 1024)} MB ca să fie citit ca imagine.")

    # ── .txt și .srt: decodare UTF-8 cu fallback latin-1 ──
    if fname.endswith(".txt") or fname.endswith(".srt") or "text/plain" in ftype or "srt" in ftype:
        try:
            text = raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            try:
                text = raw_bytes.decode("latin-1")
            except Exception:
                return None

        if len(text) <= MAX_TEXT_CHARS:
            return text  # fișier mic — returnăm integral

        # Fișier mare: pentru .srt tăiem la un bloc complet (nu la mijlocul unui dialog)
        if fname.endswith(".srt"):
            truncated = text[:MAX_TEXT_CHARS]
            last_blank = truncated.rfind("\n\n")
            if last_blank > 0:
                truncated = truncated[:last_blank]
            total_subs = text.count("\n\n") + 1
            kept_subs  = truncated.count("\n\n") + 1
            truncated += (
                f"\n\n[AVERTISMENT: subtitrarea a fost trunchiata la {kept_subs} din {total_subs} replici "
                f"({len(truncated):,} din {len(text):,} caractere). "
                f"Daca ai nevoie de tot fisierul, imparte-l in bucati si traduce pe rand.]"
            )
            return truncated

        return text[:MAX_TEXT_CHARS]

    # ── .docx: python-docx ──
    if fname.endswith(".docx") or "wordprocessingml" in ftype:
        if not _DOCX_AVAILABLE:
            return (
                "⚠️ Biblioteca python-docx nu este instalată. "
                "Adaugă 'python-docx' în requirements.txt pentru suport .docx."
            )
        try:
            import io
            doc = _DocxDocument(io.BytesIO(raw_bytes))
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            # Includem și tabelele din document
            for table in doc.tables:
                for row in table.rows:
                    row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        paragraphs.append(row_text)
            text = "\n".join(paragraphs)
            return text[:MAX_TEXT_CHARS]
        except Exception as e:
            return f"⚠️ Nu s-a putut citi fișierul .docx: {e}"

    # ── .doc (format vechi Word — binar) ──
    # Extragerea "text brut" din binar pierdea diacriticele și amesteca nume de fonturi/stiluri
    # în text, deci modelul primea gunoi. Mai bine un mesaj clar decât un răspuns greșit.
    if fname.endswith(".doc") or ftype == "application/msword":
        return (
            "⚠️ Fișierele **.doc** (Word vechi) nu pot fi citite corect — diacriticele se pierd. "
            "Deschide documentul în Word sau LibreOffice, salvează-l ca **.docx** și încarcă-l din nou."
        )

    # ── .dbf: dbfread ──
    if fname.endswith(".dbf") or "dbf" in ftype or "dbase" in ftype:
        if not _DBF_AVAILABLE:
            return (
                "⚠️ Biblioteca dbfread nu este instalată. "
                "Adaugă 'dbfread' în requirements.txt pentru suport .dbf."
            )
        try:
            import io as _io
            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".dbf") as tmp:
                    tmp.write(raw_bytes)
                    tmp_path = tmp.name
                table = _DBF(tmp_path, encoding="utf-8", ignore_missing_memofile=True)
                headers = table.field_names
                rows = []
                rows.append(" | ".join(headers))
                rows.append("-" * min(80, len(" | ".join(headers)) + 4))
                for i, record in enumerate(table):
                    if i >= 500:  # limităm la 500 rânduri pentru context
                        rows.append(f"... (și încă {len(list(table)) - 500} rânduri)")
                        break
                    rows.append(" | ".join(str(record.get(h, "")) for h in headers))
                text = "\n".join(rows)
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    os.unlink(tmp_path)
            return text[:MAX_TEXT_CHARS]
        except Exception as e:
            # Fallback: encoding cp1250 (frecvent în fișierele .dbf românești)
            try:
                tmp_path2 = None
                with tempfile.NamedTemporaryFile(delete=False, suffix=".dbf") as tmp2:
                    tmp2.write(raw_bytes)
                    tmp_path2 = tmp2.name
                table2 = _DBF(tmp_path2, encoding="cp1250", ignore_missing_memofile=True)
                headers2 = table2.field_names
                rows2 = [" | ".join(headers2)]
                for i, record in enumerate(table2):
                    if i >= 500:
                        break
                    rows2.append(" | ".join(str(record.get(h, "")) for h in headers2))
                text2 = "\n".join(rows2)
                if tmp_path2 and os.path.exists(tmp_path2):
                    os.unlink(tmp_path2)
                return text2[:MAX_TEXT_CHARS]
            except Exception as e2:
                return f"⚠️ Nu s-a putut citi fișierul .dbf: {e2}"

    return None  # tip necunoscut


# PDF-urile mai mari de atât nu mai sunt trimise ca fișier la AI: le extragem textul local cu pypdf
# (licență BSD — înlocuiește PyMuPDF, care e AGPL). Configurabil din secrets: PDF_TEXT_OVER_MB.
def _pdf_text_threshold_bytes() -> int:
    try:
        return int(float(st.secrets.get("PDF_TEXT_OVER_MB", 20)) * 1024 * 1024)
    except Exception:
        return 20 * 1024 * 1024


def _is_large_pdf(uploaded_file) -> bool:
    return (uploaded_file.name.lower().endswith(".pdf")
            and (uploaded_file.size or 0) > _pdf_text_threshold_bytes())


def _extract_pdf_text_pypdf(pdf_bytes: bytes, max_chars: int = 400_000) -> str | None:
    """Textul unui PDF, pagină cu pagină (pypdf). None dacă PDF-ul nu are text (ex. scanat)."""
    try:
        import io
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        parts, total = [], 0
        for i, page in enumerate(reader.pages, 1):
            t = (page.extract_text() or "").strip()
            if t:
                parts.append(f"--- Pagina {i} ---\n{t}")
                total += len(t)
                if total >= max_chars:
                    parts.append("[... document trunchiat — prea lung ...]")
                    break
        text = "\n\n".join(parts)
        return text[:max_chars] if text.strip() else None
    except Exception as e:
        _log("Extragere text PDF (pypdf) eșuată", "silent", e)
        return None


def _is_text_file(uploaded_file) -> bool:
    """Returnează True dacă fișierul trebuie procesat ca text (nu trimis la Google Files API)."""
    if not uploaded_file:
        return False
    fname = uploaded_file.name.lower()
    return any(fname.endswith(f".{ext}") for ext in _TEXT_FILE_EXTENSIONS) or _is_large_pdf(uploaded_file)


# === APP INSTANCE ID ===
# Separă datele între instanțe diferite ale aceleiași aplicații (același Supabase, app-uri diferite)
# Setează APP_INSTANCE_ID în secrets.toml: APP_INSTANCE_ID = "profesor_v1"
_APP_ID_PATTERN = re.compile(r'^[a-zA-Z0-9_-]{1,50}$')

@st.cache_data(ttl=3600)
def get_app_id() -> str:
    """Returnează ID-ul aplicației. Validat anti-injection.
    FIX 6: cache-uit cu st.cache_data — st.secrets accesează discul la fiecare apel,
    iar get_app_id() e apelat la fiecare query Supabase.
    """
    try:
        raw = str(st.secrets.get("APP_INSTANCE_ID", "default")).strip() or "default"
    except Exception:
        raw = "default"
    return raw if _APP_ID_PATTERN.match(raw) else "default"

# === CONSTANTE PENTRU LIMITE (FIX MEMORY LEAK) ===
MAX_MESSAGES_IN_MEMORY = 100
MAX_MESSAGES_TO_SEND_TO_AI = 20
MAX_MESSAGES_IN_DB_PER_SESSION = 500
CLEANUP_DAYS_OLD = 90  # Păstrăm istoricul 90 de zile — acoperă vacanțe, pauze lungi

# === RATE LIMITING (minut + zi per sesiune + plafon global pe zi) ===
# Streamlit Cloud nu expune IP-ul direct; session_id e unic per browser/tab, dar clientul
# și-l poate schimba (tab nou). De aceea există și un plafon GLOBAL pe zi.
# Limitele zilnice se aplică DOAR când folosești cheile tale din secrets (cota ta partajată);
# un elev care folosește propria cheie personală e limitat doar pe minut.
RATE_LIMIT_MAX_REQUESTS = 20    # cereri maxime per fereastră (per sesiune)
RATE_LIMIT_WINDOW_SEC   = 60    # fereastra, în secunde
RATE_LIMIT_SESSION_DAY  = 200   # cereri AI / zi / sesiune   (0 = dezactivat; override: secrets RATE_LIMIT_SESSION_DAY)
RATE_LIMIT_GLOBAL_DAY   = 1000  # cereri AI / zi / toți elevii (0 = dezactivat; override: secrets RATE_LIMIT_GLOBAL_DAY)
                                # ≈ cota unui proiect Google pe zi (500 + 500 + 20); înmulțește cu nr. de proiecte din secrets
# Contoarele stau într-un @st.cache_resource (vezi _get_rate_store): o variabilă globală din
# script ar fi recreată la FIECARE rerun și limitatorul n-ar bloca niciodată nimic.

# === MODEL GEMINI — singura sursă de adevăr pentru numele modelului ===
# Lanț de rezervă (vezi MODEL_FALLBACKS_NO_CACHE în run_chat_with_rotation).
# Limite free tier verificate pe dashboard (sept. 2026); cotele sunt per PROIECT Google și per model:
#   1. gemini-3.1-flash-lite — principal            (15 RPM / 500 RPD)
#   2. gemini-3.5-flash-lite — rezervă 1            (15 RPM / 500 RPD)
#   3. gemini-3.8-flash      — rezervă 2, ultima    ( 5 RPM /  20 RPD)
GEMINI_MODEL = "gemini-3.1-flash-lite"
SUMMARIZE_AFTER_MESSAGES = 30   # Rezumăm când depășim acest număr de mesaje
MESSAGES_KEPT_AFTER_SUMMARY = 10  # Câte mesaje recente păstrăm după rezumare

# === ISTORIC CONVERSAȚII ===
def _cleanup_gfiles() -> None:
    """Șterge toate fișierele uploadate pe Google Files API din sesiunea curentă.
    Apelat la switch sesiune, conversație nouă și explicit de utilizator.
    Fișierele expiră oricum după 48h, dar le ștergem proactiv pentru igienă.
    """
    st.session_state.pop("_gsrc_map", None)  # octeții păstrați pentru re-upload
    gfile_keys = [k for k in st.session_state.keys() if k.startswith("_gfile_")]
    if not gfile_keys:
        return
    try:
        _keys = st.session_state.get("_api_keys_list", [])
        _idx  = st.session_state.get("key_index", 0)
        if not _keys:
            return
        _client = genai.Client(api_key=_keys[_idx])
        for k in gfile_keys:
            gf = st.session_state.pop(k, None)
            if gf:
                try:
                    _client.files.delete(name=gf.name)
                except Exception:
                    pass  # expirat deja sau alt motiv — ignorăm
    except Exception:
        # Dacă clientul nu poate fi creat, curățăm cel puțin session_state
        for k in gfile_keys:
            st.session_state.pop(k, None)


def switch_session(new_session_id: str):
    """Comută la o altă sesiune."""
    _cleanup_gfiles()  # curățăm fișierele Google la switch sesiune
    st.session_state.session_id = new_session_id
    st.session_state.messages = []
    invalidate_session_cache()  # FIX: forțează refresh la switch
    # Curățăm contextul sesiunii vechi — nu trebuie injectat în cea nouă
    st.session_state.pop("_conversation_summary", None)
    st.session_state.pop("_summary_cached_at", None)
    st.session_state.pop("_summary_for_sid", None)
    # Resetăm materia detectată — fiecare sesiune începe cu autodetecție fresh.
    # Dacă sesiunea nouă are deja mesaje, materia va fi restaurată din DB la load.
    # Dacă e sesiune goală (chat nou), autodetecția pornește de la zero.
    st.session_state.pop("_detected_subject", None)
    st.session_state.pop("_pending_user_msg", None)
    st.session_state.pop("system_prompt", None)  # va fi regenerat cu materia corectă
    # FIX: resetăm și materia selectată manual din dropdown — altfel rămâne agățată
    # de vechea sesiune (ex: la "Conversație nouă", ecranul de selecție a materiei
    # nu mai apărea dacă exista deja o materie selectată manual înainte).
    st.session_state.pop("materie_selectata", None)
    # Curățăm toate cheile _mismatch_warned_* (una per sesiune anterioară)
    for _k in [k for k in st.session_state.keys() if k.startswith("_mismatch_warned_")]:
        del st.session_state[_k]
    # Actualizează localStorage cu noul SID — JS-ul îl va folosi la următorul load
    components.html(
        f"<script>localStorage.setItem('profesor_session_id', {json.dumps(new_session_id)});</script>",
        height=0
    )


def invalidate_session_cache():
    """Marchează cache-ul sesiunilor ca expirat — apelat după orice modificare."""
    st.session_state["_sess_cache_dirty"] = True
    st.session_state["_sess_list_ts"] = 0  # FIX: resetează timestamp pentru forțare refresh complet


def format_time_ago(timestamp) -> str:
    """Formatează timestamp ca timp relativ (ex: '2 ore în urmă'). Acceptă float sau ISO string."""
    # FIX: Supabase poate returna ISO string în loc de float
    if isinstance(timestamp, str):
        try:
            from datetime import datetime, timezone
            dt = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
            timestamp = dt.timestamp()
        except Exception:
            return "necunoscut"
    try:
        diff = time.time() - float(timestamp)
    except (TypeError, ValueError):
        return "necunoscut"
    if diff < 60:
        return "acum"
    elif diff < 3600:
        mins = int(diff / 60)
        return f"{mins} min în urmă"
    elif diff < 86400:
        hours = int(diff / 3600)
        return f"{hours}h în urmă"
    else:
        days = int(diff / 86400)
        return f"{days} zile în urmă"




# === SUPABASE CLIENT + FALLBACK ===
@st.cache_resource
def _create_supabase_client(url: str, key: str):
    """Creează clientul (conexiunea e lazy, fără query de test).

    Separat de get_supabase_client() ca eșecul să NU fie memorat: @st.cache_resource nu
    memorează excepțiile, deci o problemă temporară nu mai blochează aplicația până la reboot.
    """
    return create_client(url, key)


def get_supabase_client() -> Client | None:
    """Returnează clientul Supabase sau None. Motivul eșecului ajunge în log și în
    st.session_state["_sb_init_error"], ca să poată fi afișat (înainte eșecul era mut)."""
    try:
        url = st.secrets.get("SUPABASE_URL", "")
        key = st.secrets.get("SUPABASE_KEY", "")
        if not url or not key:
            st.session_state["_sb_init_error"] = "SUPABASE_URL sau SUPABASE_KEY lipsesc din secrets"
            return None
        client = _create_supabase_client(url, key)
        st.session_state.pop("_sb_init_error", None)
        return client
    except Exception as e:
        # NU includem cheia în mesaj — doar tipul și textul excepției
        st.session_state["_sb_init_error"] = f"{type(e).__name__}: {str(e)[:160]}"
        _log("Supabase: clientul nu a putut fi creat", "silent", e)
        return None


SB_RETRY_INTERVAL_SEC = 30  # cât așteptăm între două încercări de reconectare


def is_supabase_available() -> bool:
    """True dacă putem încerca Supabase acum.

    Online -> True. Offline -> False, dar după SB_RETRY_INTERVAL_SEC lăsăm o
    singură operație reală să "sondeze" conexiunea (circuit breaker). Dacă
    reușește, _mark_supabase_online() golește coada; dacă eșuează, se reia
    așteptarea. Fără asta, un singur eșec temporar bloca sesiunea offline pe vecie.
    """
    if st.session_state.get("_sb_online", True):
        return True
    return time.time() >= st.session_state.get("_sb_retry_at", 0)


def _mark_supabase_offline():
    """Marchează Supabase ca offline, programează următoarea încercare și notifică utilizatorul."""
    was_online = st.session_state.get("_sb_online", True)
    st.session_state["_sb_online"] = False
    st.session_state["_sb_retry_at"] = time.time() + SB_RETRY_INTERVAL_SEC
    if was_online:
        st.toast("⚠️ Baza de date offline — modul local activat.", icon="📴")


def _mark_supabase_online():
    """Marchează Supabase ca online și golește coada offline."""
    was_offline = not st.session_state.get("_sb_online", True)
    st.session_state["_sb_online"] = True
    if was_offline:
        st.toast("✅ Conexiunea restabilită!", icon="🟢")
        _flush_offline_queue()


# --- Coadă offline: mesaje salvate local când Supabase e down ---
MAX_OFFLINE_QUEUE_SIZE = 50  # Previne memory leak când Supabase e offline mult timp

def _get_offline_queue() -> list:
    queue = st.session_state.setdefault("_offline_queue", [])
    # Dacă coada depășește limita, păstrăm doar cele mai recente mesaje
    if len(queue) > MAX_OFFLINE_QUEUE_SIZE:
        st.session_state["_offline_queue"] = queue[-MAX_OFFLINE_QUEUE_SIZE:]
    return st.session_state["_offline_queue"]


def _flush_offline_queue():
    """Trimite mesajele din coada offline la Supabase când revine online.
    Anti-loop: dacă un mesaj eșuează de MAX_FLUSH_RETRIES ori, e abandonat.
    Anti-race: flag _flushing_queue previne procesarea dublă."""
    MAX_FLUSH_RETRIES = 3
    if st.session_state.get("_flushing_queue", False):
        return
    st.session_state["_flushing_queue"] = True

    queue = _get_offline_queue()
    if not queue:
        st.session_state["_flushing_queue"] = False
        return

    client = get_supabase_client()
    if not client:
        st.session_state["_flushing_queue"] = False
        return

    failed = []
    try:
        retry_counts = st.session_state.setdefault("_offline_retry_counts", {})
        for item in queue:
            item_key = f"{item.get('session_id','')}-{item.get('timestamp','')}"
            retries = retry_counts.get(item_key, 0)
            if retries >= MAX_FLUSH_RETRIES:
                _log(f"Mesaj abandonat după {MAX_FLUSH_RETRIES} încercări eșuate", "silent")
                continue
            try:
                client.table("history").insert(item).execute()
                retry_counts.pop(item_key, None)
            except Exception:
                retry_counts[item_key] = retries + 1
                failed.append(item)
        st.session_state["_offline_queue"] = failed
        st.session_state["_offline_retry_counts"] = retry_counts
    finally:
        st.session_state["_flushing_queue"] = False

    successful = len(queue) - len(failed)
    if successful > 0:
        st.toast(f"✅ {successful} mesaje sincronizate cu baza de date.", icon="☁️")

st.set_page_config(page_title="Profesor Liceu", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

# === DARK MODE ===
# Streamlit blocheaza window.parent cross-origin, deci JS nu poate modifica pagina parinte.
# Solutia corecta: injectam CSS direct cu st.markdown — fara JS, fara clase pe body.
# st.markdown injecteaza in <head>-ul paginii principale (nu iframe), deci CSS se aplica direct
# pe .stApp, stSidebar etc. Conditionam din Python care bloc CSS se injecteaza.
_dark_active = st.session_state.get("dark_mode", False)
if _dark_active:
    st.markdown("""
<style>
/* ── Variabile CSS pentru componente care altfel pierd la ordering ── */
:root {
    --svg-bg: #1e1e2e;
    --svg-border: #444;
}
/* ── Fundal general ── */
.stApp, [data-testid="stAppViewContainer"],
[data-testid="stMain"], [data-testid="block-container"],
section.main > div {
    background-color: #0e1117 !important;
    color: #fafafa !important;
}

/* ── Sidebar ── */
[data-testid="stSidebar"], [data-testid="stSidebar"] > div {
    background-color: #161b22 !important;
}
[data-testid="stSidebar"] * {
    color: #fafafa !important;
}
/* Toggle track si thumb — nu suprascrie background-ul lor */
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span:not([data-testid]),
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    background-color: transparent !important;
}

/* ── Butoane (toate variantele Streamlit) ── */
button[kind="secondary"], button[kind="primary"],
.stButton > button,
[data-testid="stBaseButton-secondary"],
[data-testid="stBaseButton-primary"],
[data-testid="stBaseButton-headerNoPadding"] {
    background-color: #2a2f3e !important;
    color: #fafafa !important;
    border-color: #555 !important;
}
button[kind="secondary"]:hover,
.stButton > button:hover,
[data-testid="stBaseButton-secondary"]:hover {
    background-color: #383d50 !important;
    border-color: #777 !important;
}

/* ── Selectbox & dropdown ── */
[data-testid="stSelectbox"] > div > div,
[data-testid="stSelectbox"] * {
    background-color: #1a1f2e !important;
    color: #fafafa !important;
}
ul[data-testid="stSelectboxVirtualDropdown"],
ul[data-testid="stSelectboxVirtualDropdown"] * {
    background-color: #1a1f2e !important;
    color: #fafafa !important;
}

/* ── Radio buttons (Materie) ── */
[data-testid="stRadio"] label,
[data-testid="stRadio"] span,
[data-testid="stRadio"] p {
    color: #fafafa !important;
}

/* ── Toggle ── */
[data-testid="stToggle"] label,
[data-testid="stToggle"] span {
    color: #fafafa !important;
}

/* ── Chat ── */
[data-testid="stChatMessageContent"],
[data-testid="stChatMessageContent"] *,
.stChatMessage, .stChatMessage * {
    background-color: transparent !important;
    color: #fafafa !important;
}
/* Bara de jos cu input — toate layerele */
[data-testid="stChatInput"],
[data-testid="stChatInput"] > div,
[data-testid="stChatInput"] textarea,
[data-testid="stChatInput"] button,
.stChatInputContainer,
.stChatInputContainer > div,
[data-testid="stBottom"],
[data-testid="stBottom"] > div,
[data-testid="stBottom"] > div > div,
section[data-testid="stBottom"],
div.stChatFloatingInputContainer,
div.stChatFloatingInputContainer > div {
    background-color: #0e1117 !important;
    color: #fafafa !important;
    border-color: #333 !important;
}
[data-testid="stChatInput"] textarea {
    background-color: #1a1f2e !important;
}

/* ── Text general ── */
p, h1, h2, h3, h4, h5, h6,
label, span, li, td, th, div,
.stMarkdown, .stMarkdown * {
    color: #fafafa !important;
}

/* ── Expander ── */
[data-testid="stExpander"],
[data-testid="stExpander"] > div {
    background-color: #1a1f2e !important;
    border-color: #444 !important;
}
[data-testid="stExpander"] * { color: #fafafa !important; }

/* ── Divider & header ── */
hr { border-color: #444 !important; }
[data-testid="stHeader"] { background-color: #0e1117 !important; }

/* svg-container handled in shared CSS block below */

/* ── Caption & info boxes ── */
[data-testid="stCaptionContainer"] * { color: #aaa !important; }
[data-testid="stInfo"], [data-testid="stInfo"] * {
    background-color: #1a2744 !important;
    color: #90c8ff !important;
}
[data-testid="stSuccess"], [data-testid="stSuccess"] * {
    background-color: #0f2a1a !important;
    color: #6fcf97 !important;
}
</style>
""", unsafe_allow_html=True)

_svg_bg    = "#1e1e2e" if _dark_active else "white"
_svg_border = "#444"    if _dark_active else "#ddd"
_svg_shadow = "0 2px 8px rgba(0,0,0,0.4)" if _dark_active else "0 2px 8px rgba(0,0,0,0.1)"
st.markdown(f"""
<style>
    .stChatMessage {{ font-size: 16px; }}
    footer {{ visibility: hidden; }}

    .svg-container {{
        background-color: {_svg_bg};
        padding: 20px;
        border-radius: 10px;
        border: 1px solid {_svg_border};
        text-align: center;
        margin: 15px 0;
        overflow: auto;
        box-shadow: {_svg_shadow};
        max-width: 100%;
    }}
    .svg-container svg {{ max-width: 100%; height: auto; }}



    /* Typing indicator */
    .typing-indicator {{
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 10px 4px;
        font-size: 14px;
        color: #888;
    }}
    .typing-dots {{
        display: flex;
        gap: 4px;
    }}
    .typing-dots span {{
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: #888;
        animation: typing-bounce 1.2s infinite ease-in-out;
    }}
    .typing-dots span:nth-child(1) {{ animation-delay: 0s; }}
    .typing-dots span:nth-child(2) {{ animation-delay: 0.2s; }}
    .typing-dots span:nth-child(3) {{ animation-delay: 0.4s; }}
    @keyframes typing-bounce {{
        0%, 80%, 100% {{ transform: scale(0.7); opacity: 0.4; }}
        40%            {{ transform: scale(1.0); opacity: 1.0; }}
    }}
</style>
""", unsafe_allow_html=True)


# === DATABASE FUNCTIONS (SUPABASE) ===

# ÎMBUNĂTĂȚIRE 3: Logger centralizat — afișează toast utilizatorului ȘI loghează în consolă.
# Niveluri: "info" (toast albastru), "warning" (toast portocaliu), "error" (toast roșu).
# Erorile silențioase de fundal (cleanup, trim) folosesc doar consola.
def _log(msg: str, level: str = "silent", exc: Exception = None):
    """Loghează un mesaj și opțional afișează un toast în interfață.
    
    level:
        "silent"  — doar print în consolă (erori de fundal, nu deranjează utilizatorul)
        "info"    — toast verde, pentru operații reușite/informative
        "warning" — toast portocaliu, pentru degradări non-critice
        "error"   — toast roșu, pentru erori vizibile utilizatorului
    """
    full_msg = f"{msg}: {exc}" if exc else msg
    print(full_msg)
    icon_map = {"info": "ℹ️", "warning": "⚠️", "error": "❌"}
    if level in icon_map:
        try:
            st.toast(msg, icon=icon_map[level])
        except Exception:
            pass  # st.toast poate eșua în contexte fără sesiune activă


# === RATE LIMITING ===

@st.cache_resource
def _get_rate_store() -> dict:
    """Store partajat între toate sesiunile și între rerun-uri (trăiește cât procesul serverului)."""
    return {
        "lock": threading.Lock(),
        "minute": defaultdict(list),     # session_id -> [timestamps] (fereastră glisantă)
        "day": "",                       # ziua de cotă curentă
        "session_day": defaultdict(int), # session_id -> cereri azi
        "global_day": 0,                 # cereri azi, toate sesiunile
    }


def _quota_day() -> str:
    """Ziua de cotă Gemini: RPD se resetează la miezul nopții Pacific (≈ 10:00 ora României)."""
    from datetime import datetime, timedelta, timezone
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("America/Los_Angeles")).strftime("%Y-%m-%d")
    except Exception:
        return (datetime.now(timezone.utc) - timedelta(hours=8)).strftime("%Y-%m-%d")


def _owner_keys_in_use() -> bool:
    """True dacă aplicația rulează cu cheile tale din secrets (cota partajată a tuturor elevilor)."""
    try:
        return ("GOOGLE_API_KEYS" in st.secrets) or ("GOOGLE_API_KEY" in st.secrets)
    except Exception:
        return False


def _rl_int(name: str, default: int) -> int:
    try:
        return int(st.secrets.get(name, default))
    except Exception:
        return default


def check_rate_limit(session_id: str, cost: int = 1, burst: bool = True) -> tuple[bool, int]:
    """Verifică limitele: per minut (sesiune), per zi (sesiune) și per zi (global).

    cost  — câte apeluri AI consumă acțiunea (ex.: poză = OCR + corectare = 2).
    burst — dacă True, cererea contează și în fereastra de 1 minut; False = doar taxare
            zilnică suplimentară (ex.: bucățile unei traduceri SRT).
    Setează st.session_state["_rl_reason"] la "minute" | "session_day" | "global_day" când refuză.

    Returns:
        (allowed, remaining_in_minute)
    """
    store = _get_rate_store()
    now = time.time()
    window_start = now - RATE_LIMIT_WINDOW_SEC
    owner = _owner_keys_in_use()
    sess_lim = _rl_int("RATE_LIMIT_SESSION_DAY", RATE_LIMIT_SESSION_DAY) if owner else 0
    glob_lim = _rl_int("RATE_LIMIT_GLOBAL_DAY", RATE_LIMIT_GLOBAL_DAY) if owner else 0
    today = _quota_day()

    with store["lock"]:
        if store["day"] != today:               # zi nouă de cotă -> resetăm contoarele zilnice
            store["day"] = today
            store["session_day"].clear()
            store["global_day"] = 0

        hist = [t for t in store["minute"][session_id] if t > window_start]
        store["minute"][session_id] = hist

        reason = None
        if burst and len(hist) >= RATE_LIMIT_MAX_REQUESTS:
            reason = "minute"
        elif sess_lim > 0 and store["session_day"][session_id] + cost > sess_lim:
            reason = "session_day"
        elif glob_lim > 0 and store["global_day"] + cost > glob_lim:
            reason = "global_day"

        if reason:
            st.session_state["_rl_reason"] = reason
            return False, 0

        if burst:
            hist.append(now)
        if owner:
            store["session_day"][session_id] += cost
            store["global_day"] += cost

        # Curăță sesiunile inactive (evită memory leak la multe sesiuni unice)
        if len(store["minute"]) > 5000:
            for k in [k for k, v in store["minute"].items() if not v or v[-1] < window_start]:
                del store["minute"][k]
            for k in [k for k in store["session_day"] if k not in store["minute"]]:
                del store["session_day"][k]

        remaining = RATE_LIMIT_MAX_REQUESTS - len(hist)

    st.session_state["_rl_reason"] = None
    return True, remaining


def enforce_rate_limit(cost: int = 1, burst: bool = True) -> bool:
    """Aplică limitele și arată mesajul potrivit. True = poți continua cu apelul AI."""
    allowed, remaining = check_rate_limit(st.session_state.get("session_id", "anon"), cost=cost, burst=burst)
    if allowed:
        if burst and remaining <= 3:
            st.toast(f"⚠️ Mai ai {remaining} cereri disponibile în acest minut.", icon="⏱️")
        return True

    reason = st.session_state.get("_rl_reason")
    if reason == "session_day":
        msg = (f"📅 **Ai atins limita zilnică de {_rl_int('RATE_LIMIT_SESSION_DAY', RATE_LIMIT_SESSION_DAY)} "
               "de cereri.** Se resetează zilnic în jurul orei 10:00 (ora României).")
    elif reason == "global_day":
        msg = ("📅 **Serviciul gratuit a atins limita zilnică pentru toți elevii.** "
               "Se resetează zilnic în jurul orei 10:00 (ora României). Revino atunci!")
    else:
        msg = (f"⏱️ **Prea multe cereri!** Ai trimis {RATE_LIMIT_MAX_REQUESTS} mesaje "
               "în ultimul minut. Așteaptă câteva secunde și încearcă din nou.")
    st.warning(msg, icon="🛑")
    return False


def _build_attachment_payload(media_content, text_file_content) -> list:
    """Părțile de payload pentru fișierele atașate.

    media_content: un Google File sau o LISTĂ de Google Files (imagini/PDF), trimise ca atare.
    text_file_content: textul extras din fișierele text (txt/docx/dbf/srt), injectat în prompt;
    când sunt mai multe fișiere text, fiecare are deja antetul „=== FIȘIER: nume ===”.
    """
    final_payload = []
    _uf = st.session_state.get("_current_uploaded_file_meta", {}) or {}
    _metas = _uf.get("files") or ([_uf] if _uf else [])
    _media = (media_content if isinstance(media_content, list) else [media_content]) if media_content else []
    _multi = len(_metas) > 1

    if _media:
        if _multi:
            _names = ", ".join(f"'{m.get('name', '')}'" for m in _metas)
            final_payload.append(
                f"Elevul ți-a trimis {len(_metas)} fișiere: {_names}. Analizează-le pe TOATE "
                "(imaginile vizual complet, documentele integral), ține cont de legăturile dintre ele "
                "și precizează la ce fișier te referi când răspunzi."
            )
        elif (_uf.get("type") or "").startswith("image/"):
            final_payload.append(
                "Elevul ți-a trimis o imagine. Analizează-o vizual complet: "
                "descrie ce vezi (obiecte, persoane, text, culori, forme, diagrame, exerciții scrise de mână) "
                "și răspunde la întrebarea elevului ținând cont de tot conținutul vizual."
            )
        else:
            final_payload.append(
                f"Elevul ți-a trimis documentul '{_uf.get('name', '')}'. "
                "Citește și analizează tot conținutul înainte de a răspunde."
            )
        final_payload.extend(_media)

    if text_file_content:
        if _multi:
            final_payload.append(
                "Conținutul fișierelor text atașate (fiecare are numele în antet):\n\n"
                f"--- ÎNCEPUT FIȘIERE ---\n{text_file_content}\n--- SFÂRȘIT FIȘIERE ---\n\n"
                "Analizează conținutul de mai sus și răspunde la întrebarea elevului."
            )
        else:
            fname = _uf.get("name", "")
            fname_lower = fname.lower()
            if fname_lower.endswith(".srt"):
                file_desc = "un fișier de subtitrare (.srt)"
            elif fname_lower.endswith((".docx", ".doc")):
                file_desc = "un document Word"
            elif fname_lower.endswith(".dbf"):
                file_desc = "o bază de date DBF"
            elif fname_lower.endswith(".pdf"):
                file_desc = "un document PDF (textul lui, extras pagină cu pagină)"
            else:
                file_desc = "un fișier text"
            final_payload.append(
                f"Elevul ți-a trimis {file_desc} cu numele '{fname}'. "
                f"Conținutul complet al fișierului este:\n\n"
                f"--- ÎNCEPUT FIȘIER ---\n{text_file_content}\n--- SFÂRȘIT FIȘIER ---\n\n"
                f"Analizează conținutul de mai sus și răspunde la întrebarea elevului."
            )
    return final_payload


def _stream_preview(text: str) -> str:
    """Textul afișat cât timp răspunsul se generează.

    Codul unui desen (marker [[DESEN_SVG]], <svg> sau <path>) nu se arată elevului ca text brut:
    afișăm doar ce e înaintea lui + un mesaj; desenul apare randat la final.
    """
    cut = len(text)
    for marker in ("[[DESEN_SVG]]", "<svg", "<path"):
        i = text.find(marker)
        if i != -1:
            cut = min(cut, i)
    if cut < len(text):
        return text[:cut].rstrip() + "\n\n*🎨 Domnul Profesor desenează...*\n\n▌"
    return text + "▌"


def _salvage_partial_response(e: Exception, partial: str, placeholder, context: str = "") -> bool:
    """Stream-ul a căzut DUPĂ ce o parte din răspuns a ajuns la elev: o păstrăm (afișată + salvată).

    Returnează True dacă am păstrat răspunsul parțial (handler-ul nu mai afișează eroarea).
    Un desen SVG rămas neterminat e tăiat; sub ~40 de caractere considerăm că nu merită păstrat.
    """
    partial = (partial or "").strip()
    cut = partial.rfind("<svg")
    if cut != -1 and "</svg>" not in partial[cut:]:
        partial = partial[:cut].rstrip()
    if len(partial) < 40:
        return False
    _log(f"[AI] răspuns întrerupt ({context}) — păstrat parțial", "silent", e)
    placeholder.empty()
    render_message_with_svg(partial)
    st.session_state.messages.append({"role": "assistant", "content": partial})
    save_message_with_limits(st.session_state.session_id, "assistant", partial)
    st.session_state.pop("_retry_history", None)
    st.session_state.pop("_retry_payload", None)
    st.warning("⚠️ Răspunsul s-a întrerupt (problemă temporară cu serviciul AI). "
               "Scrie „continuă” ca să primești restul.")
    return True


# === ERORI AI — mesaje prietenoase pentru elevi, detalii tehnice doar în log ===

def _request_retry():
    """Callback pentru butoanele „Reîncercați răspunsul” (vezi handler-ul _pending_retry)."""
    st.session_state["_pending_retry"] = True


def _show_ai_error(e: Exception, context: str = "") -> str:
    """Loghează excepția pe server și arată un mesaj prietenos. Returnează categoria erorii."""
    if isinstance(e, comercial.core.AccessError):
        st.warning(str(e))
        return "access"
    _log(f"[AI] eroare{' (' + context + ')' if context else ''}", "silent", e)
    low = (str(e) + " " + repr(e)).lower()
    if "pentru azi" in low:
        cat, msg = "quota", ("📅 Serviciul AI gratuit și-a epuizat cota pentru azi, pe toate cheile și modelele. "
                             "Cotele se resetează în jurul orei 10:00 (ora României).")
    elif "toate cheile" in low or "epuizat" in low or "quota" in low or re.search(r"\b429\b", low):
        cat, msg = "quota", ("⏳ Serviciul AI gratuit a atins temporar limita de utilizare. "
                             "Încearcă din nou peste câteva minute.")
    elif "supraînc" in low or "overloaded" in low or re.search(r"\b503\b", low):
        cat, msg = "overload", "🐢 Serviciul AI este aglomerat acum. Încearcă din nou în câteva secunde."
    elif "cerere invalidă" in low or re.search(r"\b400\b", low):
        cat, msg = "bad_request", ("⚠️ Cererea nu a putut fi procesată (mesaj prea lung sau conținut neacceptat). "
                                   "Reformulează întrebarea sau încearcă cu un fișier mai mic.")
    elif "access the file" in low or "files/" in low:
        cat, msg = "file", ("📎 Fișierul atașat nu mai poate fi folosit (a expirat sau serviciul a fost schimbat). "
                            "Elimină-l din bara laterală și încarcă-l din nou.")
    elif "niciun text" in low:
        cat, msg = "empty", ("🤔 Nu am primit niciun răspuns de la AI (poate a fost blocat de filtre). "
                             "Reformulează întrebarea și încearcă din nou.")
    elif any(x in low for x in ("timeout", "timed out", "connection", "connect", "network", "ssl")):
        cat, msg = "network", "📡 Problemă de conexiune cu serviciul AI. Verifică internetul și încearcă din nou."
    else:
        cat, msg = "other", ("❌ A apărut o problemă neașteptată. Încearcă din nou; "
                             "dacă se repetă, anunță administratorul aplicației.")
    st.error(msg)
    return cat


def init_db():
    """Verifică conexiunea la Supabase. Dacă e offline, activează modul local."""
    # Verificăm din start că avem un client: fără el nimic nu se salvează, iar înainte nu se vedea nicăieri.
    if get_supabase_client() is None and st.session_state.get("_sb_online", True):
        _mark_supabase_offline()
    online = st.session_state.get("_sb_online", True)  # starea reală, nu fereastra de reîncercare
    if not online:
        _reason = st.session_state.get("_sb_init_error")
        st.warning("📴 **Modul offline activ** — conversația se păstrează în memorie. "
                   "Istoricul va fi sincronizat automat când conexiunea revine."
                   + (f"\n\nMotiv tehnic (pentru administrator): `{_reason}`" if _reason else ""), icon="⚠️")


def cleanup_old_sessions(days_old: int = CLEANUP_DAYS_OLD):
    """Șterge sesiunile vechi — rulează cel mult o dată la 6 ore per instanță.
    Șterge HISTORY înainte de SESSIONS (ordinea corectă pentru integritate DB).
    FIX 10: throttle la 6h (nu la fiecare rerun) — previne sute de query-uri Supabase/zi."""
    _CLEANUP_INTERVAL = 6 * 3600  # 6 ore în secunde
    if time.time() - st.session_state.get("_last_cleanup", 0) < _CLEANUP_INTERVAL:
        return
    st.session_state["_last_cleanup"] = time.time()
    try:
        supabase = get_supabase_client()
        if not supabase:
            return
        cutoff_time = time.time() - (days_old * 24 * 60 * 60)
        supabase.table("history").delete().lt("timestamp", cutoff_time).eq("app_id", get_app_id()).execute()
        supabase.table("sessions").delete().lt("last_active", cutoff_time).eq("app_id", get_app_id()).execute()
    except Exception as e:
        _log("Eroare la curățarea sesiunilor vechi", "silent", e)


def save_message_to_db(session_id, role, content):
    """Salvează un mesaj în Supabase. Dacă e offline, pune în coada locală."""
    record = {
        "session_id": session_id,
        "role": role,
        "content": content,
        "timestamp": time.time(),
        "app_id": get_app_id()
    }
    if not is_supabase_available():
        q = _get_offline_queue()
        if len(q) < MAX_OFFLINE_QUEUE_SIZE:
            q.append(record)
        return
    try:
        client = get_supabase_client()
        client.table("history").insert(record).execute()
        _mark_supabase_online()
    except Exception as e:
        _log("Mesajul nu a putut fi salvat", "warning", e)
        _mark_supabase_offline()
        q = _get_offline_queue()
        if len(q) < MAX_OFFLINE_QUEUE_SIZE:
            q.append(record)


def load_history_from_db(session_id, limit: int = MAX_MESSAGES_IN_MEMORY):
    """Încarcă istoricul din Supabase. Fallback: returnează ce e deja în session_state.
    
    Când e offline: afișează avertisment și marchează că istoricul e incomplet
    (poate diferi de ce e în DB dacă utilizatorul a șters sau a schimbat sesiunea).
    """
    if not is_supabase_available():
        # FIX bug 12: offline → returnăm TOATE mesajele din memorie (nu trunchiate la limit)
        # limit-ul e pentru DB unde stocăm mult; în memorie avem deja mesajele relevante
        st.session_state["_history_may_be_incomplete"] = True
        return st.session_state.get("messages", [])
    try:
        client = get_supabase_client()
        response = (
            client.table("history")
            .select("role, content, timestamp")
            .eq("session_id", session_id)
            .eq("app_id", get_app_id())
            .neq("role", "srt_data")           # srt_data nu trebuie să consume limita
            .order("timestamp", desc=True)     # cele mai NOI mesaje primele
            .limit(limit)
            .execute()
        )
        # Query-ul a adus mesajele de la cel mai nou la cel mai vechi — le readucem în ordine cronologică
        rows = list(reversed(response.data or []))
        return [{"role": row["role"], "content": row["content"]} for row in rows]
    except Exception as e:
        _log("Eroare la încărcarea istoricului", "silent", e)
        return st.session_state.get("messages", [])[-limit:]


def clear_history_db(session_id):
    """Șterge istoricul pentru o sesiune din Supabase."""
    if not is_valid_session_id(session_id):
        _log(f"clear_history_db: session_id invalid ignorat: {str(session_id)[:20]}", "warning")
        return
    try:
        supabase = get_supabase_client()
        supabase.table("history").delete().eq("session_id", session_id).eq("app_id", get_app_id()).execute()
        invalidate_session_cache()  # FIX: sesiune ștearsă = cache invalid
        # Invalidăm și cache-ul rezumatului — conversația e nouă
        st.session_state.pop("_conversation_summary", None)
        st.session_state.pop("_summary_cached_at", None)
        st.session_state.pop("_summary_for_sid", None)
        st.session_state.pop("_mismatch_warned", None)
    except Exception as e:
        _log("Istoricul nu a putut fi șters", "warning", e)


def trim_db_messages(session_id: str):
    """Limitează mesajele din DB pentru o sesiune (FIX MEMORY LEAK)."""
    try:
        supabase = get_supabase_client()

        # Numără mesajele sesiunii
        count_resp = (
            supabase.table("history")
            .select("id", count="exact")
            .eq("session_id", session_id)
            .eq("app_id", get_app_id())
            .execute()
        )
        count = count_resp.count or 0

        if count > MAX_MESSAGES_IN_DB_PER_SESSION:
            to_delete = count - MAX_MESSAGES_IN_DB_PER_SESSION
            # Obține ID-urile celor mai vechi mesaje
            old_resp = (
                supabase.table("history")
                .select("id")
                .eq("session_id", session_id)
                .eq("app_id", get_app_id())
                .order("timestamp", desc=False)
                .limit(to_delete)
                .execute()
            )
            ids_to_delete = [row["id"] for row in old_resp.data]
            if ids_to_delete:
                supabase.table("history").delete().in_("id", ids_to_delete).execute()
    except Exception as e:
        _log("Eroare la curățarea DB", "silent", e)


# === SESSION MANAGEMENT (SUPABASE) ===

def generate_unique_session_id() -> str:
    """Generează un session ID criptografic sigur, fără risc de coliziuni.
    FIX bug 3: secrets.token_hex(32) = 64 caractere hex, entropie 256 biți —
    mult mai sigur decât combinația uuid[:16]+time+uuid[:8] anterioară."""
    return secrets.token_hex(32)  # 64 caractere hex lowercase, validat de _SESSION_ID_RE


# Regex precompilat pentru validarea session_id — doar hex lowercase, 16-64 caractere
_SESSION_ID_RE = re.compile(r'^[a-f0-9]{16,64}$')

def is_valid_session_id(sid: str) -> bool:
    """Validează session_id: doar hex lowercase, lungime 16-64 caractere.
    
    FIX: Fără validare, un sid malițios din URL (?sid=../../../etc) putea
    ajunge direct în query-urile Supabase ca parametru nevalidat.
    """
    if not sid or not isinstance(sid, str):
        return False
    return bool(_SESSION_ID_RE.match(sid))


def _session_owner(session_id: str) -> tuple[bool, str | None]:
    """(există, liceu_user_id) pentru o sesiune. La eroare: (False, None)."""
    try:
        supabase = get_supabase_client()
        response = (
            supabase.table("sessions")
            .select("session_id, liceu_user_id")
            .eq("session_id", session_id)
            .limit(1)
            .execute()
        )
        if not response.data:
            return False, None
        return True, response.data[0].get("liceu_user_id")
    except Exception:
        return False, None


def session_exists_in_db(session_id: str) -> bool:
    """Verifică dacă un session_id există deja în Supabase."""
    return _session_owner(session_id)[0]


def register_session(session_id: str):
    """Înregistrează o sesiune nouă în Supabase. Silent dacă offline."""
    if not is_supabase_available():
        return
    try:
        client = get_supabase_client()
        now = time.time()
        row = {
            "session_id": session_id,
            "created_at": now,
            "last_active": now,
            "app_id": get_app_id(),
        }
        _uid = current_user_id()
        if _uid:
            row["liceu_user_id"] = _uid   # conversația aparține contului → apare pe toate device-urile
        # ignore_duplicates: o sesiune existentă NU e suprascrisă (nici proprietarul, nici created_at)
        client.table("sessions").upsert(row, ignore_duplicates=True).execute()
    except Exception as e:
        _log("Eroare la înregistrarea sesiunii", "silent", e)


# === AUTENTIFICARE OPȚIONALĂ (cont elev) ===
# Fără cont: aplicația merge ca înainte (conversațiile stau în acest browser).
# Cu cont: conversațiile sunt legate de cont (sessions.liceu_user_id) și apar pe orice device.
# „Ține-mă minte”: un token aleator stă în localStorage-ul browserului; în DB păstrăm doar
# hash-ul lui (sha256), deci o scurgere a bazei nu permite logarea. Parolele: scrypt (stdlib).
AUTH_TOKEN_DAYS = 180
AUTH_MIN_PASSWORD = 8
AUTH_MAX_FAILS = 5            # încercări greșite per email ...
AUTH_FAIL_WINDOW = 600        # ... în 10 minute
_EMAIL_RE = re.compile(r"^[^@\s:<>,;]{1,64}@[^@\s:<>,;]{1,255}\.[a-zA-Z]{2,24}$")


def _normalize_email(email: str) -> str:
    """Litere mici, fără spații și fără prefixul „mailto:” / paranteze <> (apar când emailul e
    copiat dintr-un link — așa a ajuns în DB un cont „mailto:...”, imposibil de resetat)."""
    e = (email or "").strip().strip("<>").strip()
    e = re.sub(r"^mailto:", "", e, flags=re.IGNORECASE)
    return e.split("?")[0].strip().lower()


def current_user() -> dict | None:
    return st.session_state.get("_auth_user")


def current_user_id() -> str | None:
    u = current_user()
    return u.get("id") if u else None


def _hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    n, r, p = 2 ** 14, 8, 1
    dk = hashlib.scrypt(password.encode("utf-8"), salt=salt, n=n, r=r, p=p, dklen=32)
    return f"scrypt${n}${r}${p}${salt.hex()}${dk.hex()}"


def _verify_password(password: str, stored: str) -> bool:
    try:
        algo, n, r, p, salt_hex, dk_hex = stored.split("$")
        if algo != "scrypt":
            return False
        dk = hashlib.scrypt(password.encode("utf-8"), salt=bytes.fromhex(salt_hex),
                            n=int(n), r=int(r), p=int(p), dklen=len(dk_hex) // 2)
        return hmac.compare_digest(dk.hex(), dk_hex)
    except Exception:
        return False


def _token_hash(raw: str) -> str:
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


@st.cache_resource
def _get_auth_fail_store() -> dict:
    return {"lock": threading.Lock(), "fails": defaultdict(list)}


def _auth_blocked(email: str) -> bool:
    store = _get_auth_fail_store()
    with store["lock"]:
        recent = [t for t in store["fails"][email] if t > time.time() - AUTH_FAIL_WINDOW]
        store["fails"][email] = recent
        return len(recent) >= AUTH_MAX_FAILS


def _auth_note_fail(email: str):
    store = _get_auth_fail_store()
    with store["lock"]:
        store["fails"][email].append(time.time())


# Coloanele citite pentru utilizatorul logat (fără hash-uri). Cele de după display_name vin din
# supabase_comercial.sql (profil, rol, vârstă, acord parental).
_USER_COLS = ("id, email, display_name, role, clasa, filiera, profil_liceu, specializare, "
              "birth_year, birth_month, parent_email, parental_status")


def _public_user(row: dict) -> dict:
    u = {"id": row["id"], "email": row["email"], "name": row.get("display_name") or row["email"].split("@")[0]}
    for k in ("role", "clasa", "filiera", "profil_liceu", "specializare", "birth_year", "birth_month",
              "parent_email", "parental_status"):
        u[k] = row.get(k)
    u["role"] = u["role"] or "student"
    u["parental_status"] = u["parental_status"] or "not_required"
    return u


def refresh_current_user(max_age: float = 20.0):
    """Recitește contul din DB (rol, acord parental, profil) — cel mult o dată la max_age secunde,
    ca o schimbare făcută din altă parte (ex. părintele confirmă din email) să se vadă fără relogare."""
    u = current_user()
    if not u or time.time() - st.session_state.get("_auth_user_ts", 0) < max_age:
        return
    st.session_state["_auth_user_ts"] = time.time()
    try:
        res = get_supabase_client().table("liceu_users").select(_USER_COLS).eq("id", u["id"]).limit(1).execute()
        if res.data:
            st.session_state["_auth_user"] = _public_user(res.data[0])
        else:
            # contul a fost șters (ex. părintele a refuzat acordul)
            st.session_state.pop("_auth_user", None)
            _set_browser_token(None)
    except Exception as e:
        _log("Recitirea contului a eșuat", "silent", e)


def auth_register(email: str, password: str, display_name: str = "", birth_year: int | None = None,
                  birth_month: int | None = None, parent_email: str = "",
                  accepted_terms: bool = False) -> tuple[dict | None, str]:
    """Creează contul. Returnează (user, "") sau (None, mesaj_eroare).

    Cu BILLING_MODE = "on": anul + luna nașterii sunt obligatorii; sub 18 ani emailul părintelui
    e obligatoriu (plata o face părintele); sub 16 ani contul așteaptă acordul părintelui (link pe
    email), iar peste 16 ani elevul acceptă singur Termenii și Politica de confidențialitate."""
    email = _normalize_email(email)
    if not _EMAIL_RE.match(email):
        return None, "Adresa de email nu pare validă."
    if len(password or "") < AUTH_MIN_PASSWORD:
        return None, f"Parola trebuie să aibă cel puțin {AUTH_MIN_PASSWORD} caractere."
    parent_email = _normalize_email(parent_email) if parent_email else ""
    _strict = comercial.core.billing_on()
    _needs_parent = comercial.core.needs_parental_consent(birth_year, birth_month) if birth_year else _strict
    _minor = comercial.core.is_minor(birth_year, birth_month) if birth_year else _strict
    if _strict:
        if not birth_year or not birth_month:
            return None, "Alege anul și luna nașterii."
        if _minor and not _EMAIL_RE.match(parent_email or ""):
            return None, "Pentru elevii sub 18 ani e nevoie de emailul unui părinte."
        if _minor and parent_email == email:
            return None, "Emailul părintelui trebuie să fie diferit de al tău."
        if not _needs_parent and not accepted_terms:
            return None, "Trebuie să accepți Termenii și Politica de confidențialitate."
    client = get_supabase_client()
    if client is None:
        return None, "Baza de date nu e disponibilă acum. Încearcă mai târziu."
    try:
        exists = client.table("liceu_users").select("id").eq("email", email).limit(1).execute()
        if exists.data:
            return None, "Există deja un cont cu acest email. Folosește „Intră în cont”."
        now = time.time()
        _rec = _new_recovery_code()
        res = client.table("liceu_users").insert({
            "email": email,
            "password_hash": _hash_password(password),
            "recovery_hash": _hash_password(_normalize_recovery(_rec)),
            "display_name": (display_name or "").strip()[:60] or None,
            "created_at": now,
            "last_login": now,
            **({"birth_year": int(birth_year), "birth_month": int(birth_month or 0) or None,
                "parent_email": parent_email or None,
                "parental_status": "pending" if (_strict and _needs_parent) else "not_required"}
               if birth_year else {}),
        }).execute()
        st.session_state["_auth_new_recovery"] = _rec   # afișat o singură dată după creare
        user = _public_user(res.data[0])
        if _strict and _needs_parent:
            _ok, _msg = comercial.compliance.request_parental_consent(user, parent_email)
            st.session_state["_auth_parent_flash"] = (_ok, _msg)
        elif accepted_terms:
            comercial.compliance.record_acceptance(user["id"], "elev", email)
        return user, ""
    except Exception as e:
        _log("Înregistrare cont eșuată", "silent", e)
        return None, "Contul nu a putut fi creat. Încearcă din nou."


def auth_login(email: str, password: str) -> tuple[dict | None, str]:
    email = _normalize_email(email)
    if _auth_blocked(email):
        return None, "Prea multe încercări greșite. Așteaptă 10 minute."
    client = get_supabase_client()
    if client is None:
        return None, "Baza de date nu e disponibilă acum. Încearcă mai târziu."
    try:
        res = client.table("liceu_users").select(_USER_COLS + ", password_hash") \
            .eq("email", email).limit(1).execute()
    except Exception as e:
        _log("Autentificare: eroare DB", "silent", e)
        return None, "Autentificarea nu a mers acum. Încearcă din nou."
    row = res.data[0] if res.data else None
    if not row or not _verify_password(password or "", row["password_hash"]):
        _auth_note_fail(email)
        return None, "Email sau parolă greșită."
    try:
        client.table("liceu_users").update({"last_login": time.time()}).eq("id", row["id"]).execute()
    except Exception:
        pass
    return _public_user(row), ""


# --- Resetare parolă: cod prin email (Gmail SMTP) + cod de recuperare ---
RESET_CODE_MINUTES = 15
RESET_MAX_ATTEMPTS = 5
RESET_EMAIL_COOLDOWN = 60        # secunde între două emailuri către aceeași adresă
_RECOVERY_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"   # fără 0/O, 1/I


def _new_recovery_code() -> str:
    raw = "".join(secrets.choice(_RECOVERY_ALPHABET) for _ in range(12))
    return f"{raw[:4]}-{raw[4:8]}-{raw[8:]}"


def _normalize_recovery(code: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (code or "").upper())


def _reset_code_hash(user_id: str, code: str) -> str:
    return hashlib.sha256(f"{user_id}:{code}".encode("utf-8")).hexdigest()


def _smtp_config() -> dict | None:
    try:
        user = st.secrets.get("SMTP_USER", "")
        pw = st.secrets.get("SMTP_PASSWORD", "")
        if not user or not pw:
            return None
        return {"user": user, "password": str(pw).replace(" ", ""),
                "host": st.secrets.get("SMTP_HOST", "smtp.gmail.com"),
                "port": int(st.secrets.get("SMTP_PORT", 465)),
                "sender": st.secrets.get("SMTP_FROM_NAME", "Profesor Liceu")}
    except Exception:
        return None


def _send_email(to: str, subject: str, body: str) -> bool:
    cfg = _smtp_config()
    if not cfg:
        return False
    msg = EmailMessage()
    msg["From"] = f"{cfg['sender']} <{cfg['user']}>"
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    try:
        with smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=20) as srv:
            srv.login(cfg["user"], cfg["password"])
            srv.send_message(msg)
        return True
    except Exception as e:
        _log("Trimiterea emailului a eșuat", "silent", e)
        return False


@st.cache_resource
def _get_reset_mail_store() -> dict:
    return {"lock": threading.Lock(), "last": {}}


def auth_request_reset_email(email: str) -> tuple[bool, str]:
    """Trimite un cod de 6 cifre pe email. Mesajul e același dacă emailul există sau nu
    (nu dezvăluim ce adrese au cont)."""
    email = _normalize_email(email)
    if not _EMAIL_RE.match(email):
        return False, "Adresa de email nu pare validă."
    if not _smtp_config():
        return False, "Trimiterea de emailuri nu e configurată. Folosește codul de recuperare."
    store = _get_reset_mail_store()
    with store["lock"]:
        if time.time() - store["last"].get(email, 0) < RESET_EMAIL_COOLDOWN:
            return False, "Ți-am trimis deja un cod. Verifică emailul (și Spam) sau încearcă peste un minut."
        store["last"][email] = time.time()
    generic = (f"Dacă există un cont cu {email}, am trimis un cod de 6 cifre (valabil "
               f"{RESET_CODE_MINUTES} minute). Verifică și folderul Spam.")
    client = get_supabase_client()
    if client is None:
        return False, "Baza de date nu e disponibilă acum. Încearcă mai târziu."
    try:
        res = client.table("liceu_users").select("id").eq("email", email).limit(1).execute()
        if not res.data:
            return True, generic
        uid = res.data[0]["id"]
        code = f"{secrets.randbelow(10 ** 6):06d}"
        now = time.time()
        client.table("liceu_password_resets").upsert({
            "user_id": uid, "code_hash": _reset_code_hash(uid, code),
            "expires_at": now + RESET_CODE_MINUTES * 60, "attempts": 0, "created_at": now,
        }).execute()
    except Exception as e:
        _log("Cerere resetare parolă: eroare DB", "silent", e)
        return False, "Nu am putut genera codul acum. Încearcă din nou."
    body = (f"Salut!\n\nCodul tău pentru resetarea parolei în Profesor Liceu este:\n\n    {code}\n\n"
            f"Codul e valabil {RESET_CODE_MINUTES} minute. Dacă nu ai cerut tu resetarea, "
            "ignoră acest email — parola ta rămâne neschimbată.\n")
    if not _send_email(email, f"Codul tău Profesor Liceu: {code}", body):
        return False, "Emailul nu a putut fi trimis acum. Încearcă din nou sau folosește codul de recuperare."
    return True, generic


def auth_reset_password(email: str, code: str, new_password: str) -> tuple[dict | None, str]:
    """Schimbă parola folosind codul din email (6 cifre) SAU codul de recuperare.
    La succes: toate dispozitivele sunt deconectate; dacă s-a folosit codul de recuperare,
    se generează unul nou (pus în st.session_state["_auth_new_recovery"])."""
    email = _normalize_email(email)
    code = (code or "").strip()
    if len(new_password or "") < AUTH_MIN_PASSWORD:
        return None, f"Parola nouă trebuie să aibă cel puțin {AUTH_MIN_PASSWORD} caractere."
    if _auth_blocked(email):
        return None, "Prea multe încercări greșite. Așteaptă 10 minute."
    bad = "Email sau cod greșit (ori codul a expirat)."
    client = get_supabase_client()
    if client is None:
        return None, "Baza de date nu e disponibilă acum. Încearcă mai târziu."
    try:
        res = client.table("liceu_users").select(_USER_COLS + ", recovery_hash") \
            .eq("email", email).limit(1).execute()
        row = res.data[0] if res.data else None
        if not row:
            _auth_note_fail(email)
            return None, bad
        uid = row["id"]
        used_recovery = False
        if re.fullmatch(r"\d{6}", code):
            r = client.table("liceu_password_resets").select("code_hash, expires_at, attempts") \
                .eq("user_id", uid).limit(1).execute()
            rr = r.data[0] if r.data else None
            if (not rr or rr["expires_at"] < time.time() or rr["attempts"] >= RESET_MAX_ATTEMPTS
                    or not hmac.compare_digest(rr["code_hash"], _reset_code_hash(uid, code))):
                if rr:
                    client.table("liceu_password_resets").update({"attempts": rr["attempts"] + 1}) \
                        .eq("user_id", uid).execute()
                _auth_note_fail(email)
                return None, bad
        else:
            if not row.get("recovery_hash") or not _verify_password(_normalize_recovery(code), row["recovery_hash"]):
                _auth_note_fail(email)
                return None, bad
            used_recovery = True
        update = {"password_hash": _hash_password(new_password), "last_login": time.time()}
        if used_recovery:
            new_rec = _new_recovery_code()        # codul de recuperare e de unică folosință
            update["recovery_hash"] = _hash_password(_normalize_recovery(new_rec))
        client.table("liceu_users").update(update).eq("id", uid).execute()
        client.table("liceu_password_resets").delete().eq("user_id", uid).execute()
        client.table("liceu_tokens").delete().eq("user_id", uid).execute()   # deconectăm peste tot
        if used_recovery:
            st.session_state["_auth_new_recovery"] = new_rec
        return _public_user(row), ""
    except Exception as e:
        _log("Resetare parolă eșuată", "silent", e)
        return None, "Parola nu a putut fi schimbată acum. Încearcă din nou."


def auth_regenerate_recovery() -> bool:
    uid = current_user_id()
    if not uid:
        return False
    rec = _new_recovery_code()
    try:
        get_supabase_client().table("liceu_users").update(
            {"recovery_hash": _hash_password(_normalize_recovery(rec))}).eq("id", uid).execute()
        st.session_state["_auth_new_recovery"] = rec
        return True
    except Exception as e:
        _log("Cod de recuperare nou eșuat", "silent", e)
        return False


def _dismiss_recovery():
    st.session_state.pop("_auth_new_recovery", None)


def _issue_remember_token(user_id: str) -> str | None:
    raw = secrets.token_urlsafe(32)
    now = time.time()
    try:
        get_supabase_client().table("liceu_tokens").insert({
            "token_hash": _token_hash(raw),
            "user_id": user_id,
            "created_at": now,
            "expires_at": now + AUTH_TOKEN_DAYS * 86400,
            "last_used": now,
        }).execute()
        return raw
    except Exception as e:
        _log("Token „ține-mă minte” nu a putut fi creat", "silent", e)
        return None


def _user_from_token(raw: str) -> dict | None:
    """Returnează utilizatorul pentru token (și prelungește valabilitatea), sau None."""
    if not raw or not isinstance(raw, str) or len(raw) > 200:
        return None
    try:
        client = get_supabase_client()
        th = _token_hash(raw)
        res = client.table("liceu_tokens").select("user_id, expires_at").eq("token_hash", th).limit(1).execute()
        if not res.data or res.data[0]["expires_at"] < time.time():
            return None
        u = client.table("liceu_users").select(_USER_COLS) \
            .eq("id", res.data[0]["user_id"]).limit(1).execute()
        if not u.data:
            return None
        now = time.time()
        client.table("liceu_tokens").update({
            "last_used": now, "expires_at": now + AUTH_TOKEN_DAYS * 86400,
        }).eq("token_hash", th).execute()
        client.table("liceu_users").update({"last_login": now}).eq("id", u.data[0]["id"]).execute()
        return _public_user(u.data[0])
    except Exception as e:
        _log("Restaurare login din token eșuată", "silent", e)
        return None


def _set_browser_token(raw: str | None):
    """Programează scrierea/ștergerea tokenului în localStorage (scriptul e randat la fiecare run
    de inject_session_js; conținut identic = nu se re-execută)."""
    st.session_state["_auth_ls_js"] = (
        f"localStorage.setItem('profesor_auth_token', {json.dumps(raw)});" if raw
        else "localStorage.removeItem('profesor_auth_token');"
    )


def _claim_browser_sessions(user_id: str):
    """Conversațiile anonime din acest browser trec în contul utilizatorului."""
    sids = [s for s in set(st.session_state.get("_my_session_ids", []) + [st.session_state.get("session_id", "")])
            if is_valid_session_id(s)]
    if not sids:
        return
    try:
        get_supabase_client().table("sessions").update({"liceu_user_id": user_id}) \
            .in_("session_id", sids).is_("liceu_user_id", "null").execute()
    except Exception as e:
        _log("Atașarea conversațiilor la cont a eșuat", "silent", e)


def auth_complete_login(user: dict, remember: bool):
    st.session_state["_auth_user"] = user
    _claim_browser_sessions(user["id"])
    _set_browser_token(_issue_remember_token(user["id"]) if remember else None)
    st.session_state["_sid_registered"] = set()   # re-verificăm proprietarul sesiunii curente
    invalidate_session_cache()


def auth_logout(all_devices: bool = False):
    user = current_user()
    raw = (st.session_state.get("_browser_storage") or {}).get("token") or st.session_state.get("_auth_raw_token")
    try:
        client = get_supabase_client()
        if user and all_devices:
            client.table("liceu_tokens").delete().eq("user_id", user["id"]).execute()
        elif raw:
            client.table("liceu_tokens").delete().eq("token_hash", _token_hash(raw)).execute()
    except Exception as e:
        _log("Ștergere token la deconectare eșuată", "silent", e)
    st.session_state.pop("_auth_user", None)
    st.session_state.pop("_auth_raw_token", None)
    st.session_state.pop("_acct_keys", None)
    if isinstance(st.session_state.get("_browser_storage"), dict):
        st.session_state["_browser_storage"]["token"] = ""
    _set_browser_token(None)
    # Conversația curentă aparține contului: pornim una nouă, anonimă, pe acest browser.
    new_sid = generate_unique_session_id()
    st.session_state["_sid_registered"] = set()
    register_session(new_sid)
    st.session_state["_my_session_ids"] = [new_sid]
    switch_session(new_sid)
    try:
        st.query_params["sid"] = new_sid
    except Exception:
        pass


# --- Cheile API Gemini ale elevului, salvate în cont (criptate) ---
# Elevul poate adăuga mai multe chei (ex. de la părinți/rude, fiecare cont Google = proiect
# separat = cote separate). Cheile sunt criptate cu Fernet (AES-128-CBC + HMAC) ÎNAINTE de a
# ajunge în Supabase. Cheia de criptare: secrets KEYS_ENCRYPTION_SECRET (recomandat); dacă lipsește,
# e derivată din SUPABASE_KEY. Prefixul textului criptat spune ce cheie a fost folosită.
MAX_ACCOUNT_KEYS = 10


def _fernet_for(tag: str):
    from cryptography.fernet import Fernet
    import base64 as _b64
    if tag == "s1":
        secret = str(st.secrets.get("KEYS_ENCRYPTION_SECRET", ""))
    else:
        secret = "derived:" + str(st.secrets.get("SUPABASE_KEY", ""))
    if not secret or secret == "derived:":
        raise RuntimeError("Nu există niciun secret pentru criptarea cheilor")
    raw = hashlib.sha256(b"profesor-liceu-api-keys:" + secret.encode("utf-8")).digest()
    return Fernet(_b64.urlsafe_b64encode(raw))


def _encrypt_api_key(plain: str) -> str:
    tag = "s1" if st.secrets.get("KEYS_ENCRYPTION_SECRET", "") else "d1"
    return f"{tag}:" + _fernet_for(tag).encrypt(plain.encode("utf-8")).decode("ascii")


def _decrypt_api_key(enc: str) -> str | None:
    try:
        tag, token = enc.split(":", 1)
        return _fernet_for(tag).decrypt(token.encode("ascii")).decode("utf-8")
    except Exception as e:
        _log("O cheie API din cont nu a putut fi decriptată", "silent", e)
        return None


def _plausible_api_key(k: str) -> bool:
    return bool(k) and 15 <= len(k) <= 200 and not any(ch.isspace() for ch in k)


def acct_keys_load(force: bool = False) -> list:
    """Cheile din contul utilizatorului logat: [{id, label, hint, key}] (cache în sesiune)."""
    uid = current_user_id()
    if not uid:
        return []
    cache = st.session_state.get("_acct_keys")
    if cache is not None and not force and cache.get("uid") == uid:
        return cache["items"]
    items = []
    try:
        res = get_supabase_client().table("liceu_api_keys").select("id, label, key_enc, key_hint, created_at") \
            .eq("user_id", uid).order("created_at").execute()
        for r in res.data or []:
            k = _decrypt_api_key(r["key_enc"])
            if k:
                items.append({"id": r["id"], "label": r.get("label") or "Cheie", "hint": r.get("key_hint") or "", "key": k})
    except Exception as e:
        _log("Cheile din cont nu au putut fi încărcate", "silent", e)
        return (cache or {}).get("items", [])
    st.session_state["_acct_keys"] = {"uid": uid, "items": items}
    return items


def _check_api_key_live(k: str) -> tuple[bool, str]:
    """Verifică cheia la Google fără să consume din cota de generare (models.get)."""
    try:
        genai.Client(api_key=k).models.get(model=GEMINI_MODEL)
        return True, ""
    except Exception as e:
        low = (str(e) + repr(e)).lower()
        if "api_key_invalid" in low or "api key not valid" in low or "invalid api key" in low:
            return False, "Google spune că cheia nu e validă. Verifică să o fi copiat complet."
        if "permission_denied" in low or "403" in low:
            return False, "Cheia nu are acces la Gemini API (activează „Generative Language API” în proiect)."
        return True, "Nu am putut verifica acum cheia la Google; am salvat-o oricum."


def acct_key_add(label: str, key: str) -> tuple[bool, str]:
    uid = current_user_id()
    if not uid:
        return False, "Trebuie să fii conectat la cont."
    key = (key or "").strip().strip('"').strip("'")
    if not _plausible_api_key(key):
        return False, "Cheie invalidă: fără spații, minim 15 caractere."
    existing = acct_keys_load(force=True)
    if any(it["key"] == key for it in existing):
        return False, "Cheia aceasta e deja salvată în cont."
    if len(existing) >= MAX_ACCOUNT_KEYS:
        return False, f"Poți salva cel mult {MAX_ACCOUNT_KEYS} chei."
    ok, note = _check_api_key_live(key)
    if not ok:
        return False, note
    try:
        get_supabase_client().table("liceu_api_keys").insert({
            "user_id": uid,
            "label": (label or "").strip()[:40] or f"Cheia {len(existing) + 1}",
            "key_enc": _encrypt_api_key(key),
            "key_hint": key[-4:],
            "created_at": time.time(),
        }).execute()
    except Exception as e:
        _log("Salvarea cheii în cont a eșuat", "silent", e)
        return False, "Cheia nu a putut fi salvată acum. Încearcă din nou."
    acct_keys_load(force=True)
    return True, note


def acct_key_delete(key_id: str):
    uid = current_user_id()
    if not uid:
        return
    try:
        get_supabase_client().table("liceu_api_keys").delete().eq("id", key_id).eq("user_id", uid).execute()
    except Exception as e:
        _log("Ștergerea cheii din cont a eșuat", "silent", e)
    acct_keys_load(force=True)


# --- Citirea cheii API dintr-o poză / captură de ecran ---
# Cheile sunt lungi (AIza... 39 de caractere) și greu de tastat pe telefon. OCR-ul clasic (Tesseract)
# confundă I/l/1 și O/0, deci folosim Gemini. Ce cheie folosim pentru citire: oricare cheie deja
# disponibilă (secrets / cont / browser) sau OCR_API_KEY din secrets — o cheie a ta, folosită DOAR
# aici, ca elevul să-și poată citi prima cheie din poză. Limită: 10 poze / sesiune.
_KEY_PATTERNS = [re.compile(r"AIza[0-9A-Za-z_\-]{35}"), re.compile(r"AQ\.[0-9A-Za-z_\-\.]{20,200}")]
KEY_PHOTO_MAX_PER_SESSION = 10


def _find_api_key_in_text(text: str) -> str | None:
    flat = re.sub(r"\s+", "", text or "")
    for pat in _KEY_PATTERNS:
        m = pat.search(flat)
        if m:
            return m.group(0)
    return None


def _ocr_keys_available(extra_keys: list) -> list:
    pool = [k for k in (extra_keys or []) if k]
    try:
        ocr = str(st.secrets.get("OCR_API_KEY", "")).strip()
    except Exception:
        ocr = ""
    if ocr and ocr not in pool:
        pool.append(ocr)
    return pool


def _read_api_key_from_image(data: bytes, mime: str, pool: list) -> tuple[str | None, str]:
    """Returnează (cheia, mesaj). Încearcă pe rând cheile disponibile × 3 modele mici."""
    if not pool:
        return None, ("Citirea din poză are nevoie de o cheie deja configurată. Adaug-o o dată prin copy-paste "
                      "(de pe calculator, de exemplu) — dacă ești logat, apare apoi automat și pe telefon.")
    n = st.session_state.get("_key_photo_count", 0)
    if n >= KEY_PHOTO_MAX_PER_SESSION:
        return None, "Ai atins limita de poze citite în această sesiune. Lipește cheia manual."
    st.session_state["_key_photo_count"] = n + 1
    prompt = ("Imaginea conține o cheie API Google (începe cu „AIza” sau cu „AQ.”). Transcrie DOAR cheia, exact, "
              "caracter cu caracter, fără spații și fără alt text. Atenție la caracterele asemănătoare: "
              "O (litera) / 0 (cifra), l (L mic) / I (i mare) / 1, S / 5, Z / 2. "
              "Dacă nu vezi nicio cheie, răspunde exact: NIMIC")
    models = [GEMINI_MODEL, "gemini-3.5-flash-lite", "gemini-2.5-flash-lite"]
    last_err = ""
    for k in pool[:4]:
        for m in models:
            try:
                cfg = genai_types.GenerateContentConfig(
                    thinking_config=(genai_types.ThinkingConfig(thinking_budget=0) if m.startswith("gemini-2.5")
                                     else genai_types.ThinkingConfig(thinking_level="low")))
                resp = genai.Client(api_key=k).models.generate_content(
                    model=m,
                    contents=[genai_types.Part.from_bytes(data=data, mime_type=mime or "image/jpeg"), prompt],
                    config=cfg,
                )
                found = _find_api_key_in_text(getattr(resp, "text", "") or "")
                if found:
                    return found, ""
                return None, "Nu am găsit nicio cheie în poză. Încearcă o poză mai clară, doar cu cheia, sau lipește-o."
            except Exception as e:
                last_err = str(e)
                low = last_err.lower()
                if "429" in low or "quota" in low or "resource_exhausted" in low or "404" in low \
                        or "not found" in low or "503" in low or "unavailable" in low:
                    continue          # model/cheie indisponibile → următoarea variantă
                break                 # cheie invalidă etc. → trecem la cheia următoare
    _log("Citirea cheii din poză a eșuat", "silent", Exception(last_err))
    return None, "Nu am putut citi poza acum (serviciul AI e ocupat). Încearcă din nou sau lipește cheia."


def _key_photo_reader(target_key: str, pool: list):
    """Uploader „📷 poză cu cheia”: citește cheia și o pune în câmpul `target_key` (înainte ca acesta să fie randat)."""
    up = st.file_uploader("📷 Sau încarcă o poză / captură de ecran cu cheia",
                          type=["jpg", "jpeg", "png", "webp", "heic", "heif"], key=f"{target_key}__photo",
                          help="Pe telefon poți face direct o poză ecranului calculatorului cu cheia afișată.")
    if not up:
        return
    sig = (up.name, up.size)
    if st.session_state.get(f"{target_key}__photo_sig") == sig:
        return
    st.session_state[f"{target_key}__photo_sig"] = sig
    with st.spinner("🔍 Citesc cheia din poză..."):
        found, note = _read_api_key_from_image(up.getvalue(), up.type, pool)
    if not found:
        st.session_state[f"{target_key}__photo_msg"] = ("error", note)
        return
    ok, why = _check_api_key_live(found)
    st.session_state[target_key] = found
    if ok:
        st.session_state[f"{target_key}__photo_msg"] = (
            "ok", f"✅ Am citit cheia (…{found[-4:]}) și Google o acceptă. Apasă butonul de mai jos ca s-o salvezi.")
    else:
        st.session_state[f"{target_key}__photo_msg"] = (
            "warn", f"⚠️ Am citit „…{found[-6:]}”, dar Google nu o acceptă — probabil un caracter citit greșit. "
                    "Verifică-l în câmp sau lipește cheia manual.")


def _key_photo_message(target_key: str):
    msg = st.session_state.pop(f"{target_key}__photo_msg", None)
    if msg:
        {"ok": st.success, "warn": st.warning, "error": st.error}[msg[0]](msg[1])


def _resolve_session_for_user(sid: str) -> str:
    """Înregistrează sid-ul dacă e nou. Dacă aparține ALTUI cont, nu îl deschidem: pornim o
    conversație nouă (ID-ul din link nu mai e suficient pentru a citi conversația unui cont)."""
    reg = st.session_state.setdefault("_sid_registered", set())
    if sid in reg:
        return sid
    exists, owner = _session_owner(sid)
    if exists and owner and owner != current_user_id():
        sid = generate_unique_session_id()
        exists = False
    if not exists:
        register_session(sid)
    if st.session_state.get("_sb_online", True):
        reg.add(sid)
    return sid


def load_browser_storage_and_login():
    """O dată per sesiune Streamlit: citește din browser SID-ul, lista de conversații și tokenul
    de login (componenta session_bridge), apoi restaurează automat contul din token."""
    if "_browser_storage" not in st.session_state:
        val = _session_bridge(key="_session_bridge", default=None) if _session_bridge is not None else {}
        if val is None:
            st.markdown(
                '<div style="display:flex;align-items:center;justify-content:center;'
                'min-height:40vh;color:#888;font-size:15px;">🎓 Se încarcă...</div>',
                unsafe_allow_html=True,
            )
            st.caption("Dacă pagina rămâne aici mai mult de câteva secunde, apasă mai jos.")
            if st.button("🔄 Continuă", key="_session_gate_manual_continue"):
                st.session_state["_browser_storage"] = {}
                st.rerun()
            st.stop()
        st.session_state["_browser_storage"] = val if isinstance(val, dict) else {}
    if not st.session_state.get("_auth_restore_done"):
        st.session_state["_auth_restore_done"] = True
        raw = (st.session_state["_browser_storage"] or {}).get("token") or ""
        if raw and not current_user():
            user = _user_from_token(raw)
            if user:
                st.session_state["_auth_user"] = user
                st.session_state["_auth_raw_token"] = raw
            elif is_supabase_available() and get_supabase_client() is not None:
                _set_browser_token(None)   # token expirat/revocat → îl scoatem din browser


def update_session_activity(session_id: str):
    """Actualizează timestamp-ul activității — cel mult o dată la 5 minute."""
    last = st.session_state.get("_last_activity_update", 0)
    if time.time() - last < 300:
        return
    st.session_state["_last_activity_update"] = time.time()
    if not is_supabase_available():
        return
    try:
        client = get_supabase_client()
        client.table("sessions").update({
            "last_active": time.time()
        }).eq("session_id", session_id).execute()
    except Exception as e:
        _log("Eroare la actualizarea sesiunii", "silent", e)


def inject_session_js():
    """
    JS care sincronizează SID-ul confirmat (din st.session_state) cu localStorage
    și curăță URL-ul vizual de parametrul ?sid=.

    FIX PERSISTENȚĂ (v2): logica de "compară și redirectează dacă diferă" a fost
    mutată în get_or_create_session_id() (gate cu st.stop(), rulează ÎNAINTE de orice
    altă logică Python). La momentul în care inject_session_js() rulează, SID-ul din
    st.session_state e deja cel corect — fie a venit valid prin URL, fie a fost
    confirmat de gate-ul de mai sus. Aici doar sincronizăm localStorage (idempotent)
    și ascundem ?sid= din bara de adrese.
    """
    current_sid = st.session_state.get("session_id", "")
    # FIX PERSISTENȚĂ LISTĂ CONVERSAȚII: citim lista de "sesiuni cunoscute ale acestui
    # browser" din query param-ul ?known= dacă e prezent (SID-uri separate prin virgulă).
    _known_from_url = st.query_params.get("known", "")
    if _known_from_url:
        _known_ids = [s for s in _known_from_url.split(",") if is_valid_session_id(s)]
        if _known_ids:
            _existing = set(st.session_state.get("_my_session_ids", []))
            st.session_state["_my_session_ids"] = list(_existing.union(_known_ids))

    _auth_js = st.session_state.get("_auth_ls_js")
    if _auth_js:
        # scriere/ștergere token „ține-mă minte” (identic la fiecare run → se execută o singură dată)
        components.html(f"<script>try{{{_auth_js}}}catch(e){{}}</script>", height=0)

    components.html(f"""
    <script>
    (function() {{
        const SID_KEY    = 'profesor_session_id';
        const APIKEY_KEY = 'profesor_api_key';
        const KNOWN_KEY  = 'profesor_known_sessions';
        const params     = new URLSearchParams(window.parent.location.search);
        const sidInUrl   = params.get('sid');
        const pythonSid  = {json.dumps(current_sid)};

        // Sincronizare idempotentă — SID-ul curent e deja cel confirmat
        if (pythonSid && pythonSid.length >= 16) {{
            localStorage.setItem(SID_KEY, pythonSid);

            // FIX PERSISTENȚĂ LISTĂ CONVERSAȚII: adăugăm SID-ul curent în lista
            // locală de "sesiuni cunoscute ale acestui browser", persistată în
            // localStorage (nu doar în st.session_state, care se reseta la orice
            // restart/reload real). Best-effort: nu forțăm niciun reload pentru asta
            // — lista se va sincroniza complet la următorul reload natural al paginii
            // (buton "Conversație nouă" din sidebar pune ?known= explicit, vezi mai jos).
            let known = [];
            try {{
                known = JSON.parse(localStorage.getItem(KNOWN_KEY) || '[]');
                if (!Array.isArray(known)) known = [];
            }} catch (e) {{ known = []; }}
            if (!known.includes(pythonSid)) {{
                known.push(pythonSid);
                if (known.length > 50) known = known.slice(-50);
                localStorage.setItem(KNOWN_KEY, JSON.stringify(known));
            }}
        }}

        // Curăță URL-ul vizual (sid/apikey/known nu trebuie să rămână vizibile)
        if (sidInUrl || params.get('apikey') || params.get('known')) {{
            params.delete('sid');
            params.delete('apikey');
            params.delete('known');
            const newUrl = window.parent.location.pathname +
                (params.toString() ? '?' + params.toString() : '');
            window.parent.history.replaceState(null, '', newUrl);
        }}

    }})();
    </script>
    """, height=0)


def get_or_create_session_id() -> str:
    """
    URL-ul ?sid= este SINGURA sursă de adevăr pentru identitatea browserului.

    PROBLEMA REZOLVATĂ: st.session_state poate fi shared între vizitatori pe aceeași
    instanță Streamlit. De aceea NU folosim session_state ca sursă primară — doar URL-ul.

    Flux prima vizită (URL fără ?sid=):
      Python generează UUID → îl pune în ?sid= → URL-ul devine unic per browser

    Flux revenire (bookmark, restart telefon):
      Elevul deschide URL-ul cu ?sid= → Python îl citește → restaurează istoricul

    FIX PERSISTENȚĂ (v2 — gate explicit): Vechea variantă genera un SID nou și lăsa
    SCRIPTUL ÎNTREG să ruleze cu el (inclusiv încărcarea istoricului, care apărea gol
    pentru elev) ÎNAINTE ca JS-ul să aibă șansa să verifice localStorage și să
    redirecteze. Userul vedea mereu un flash de conversație goală, și pe conexiuni
    lente sau redirect-uri ratate (storage partitioning pe Safari iOS/Chrome mobil),
    putea rămâne blocat pe sesiunea fantomă.

    Acum: dacă URL-ul e curat (fără ?sid= valid), NU continuăm scriptul deloc.
    Injectăm imediat un JS minimal care verifică localStorage și apoi:
      - dacă găsește un SID vechi → redirect direct la el (sesiunea veche se restaurează,
        elevul nu vede niciodată ecranul gol)
      - dacă nu găsește nimic (vizită cu adevărat nouă) → confirmă SID-ul nou prin URL
        și reîncarcă o singură dată
    În ambele cazuri folosim st.stop() — restul aplicației Python nu rulează până
    nu avem un ?sid= confirmat în URL.
    """
    # Citește ?sid= din URL — sursa de adevăr
    sid_from_url = st.query_params.get("sid", "")

    if is_valid_session_id(sid_from_url):
        # URL are ?sid= valid — înregistrează dacă e nou, altfel restaurează
        # O singură verificare în DB per SID și sesiune Streamlit (înainte: un query la FIECARE rerun).
        # Nu marcăm SID-ul ca verificat cât Supabase e offline — altfel nu s-ar mai înregistra niciodată.
        _resolved = _resolve_session_for_user(sid_from_url)
        if _resolved != sid_from_url:
            # conversația din link aparține altui cont → pornim una nouă
            st.session_state["session_id"] = _resolved
            try:
                st.query_params["sid"] = _resolved
            except Exception:
                pass
            return _resolved
        st.session_state["session_id"] = sid_from_url
        # FIX PERSISTENȚĂ LISTĂ CONVERSAȚII: dacă gate-ul de mai jos a transmis
        # ?known= la restaurare (vezi blocul JS), îl absorbim aici în session_state
        # ca să poată fi folosit de sidebar-ul "Conversații anterioare".
        _known_param = st.query_params.get("known", "")
        if _known_param:
            _known_ids = [s for s in _known_param.split(",") if is_valid_session_id(s)]
            if _known_ids:
                _existing = set(st.session_state.get("_my_session_ids", []))
                st.session_state["_my_session_ids"] = list(_existing.union(_known_ids))
        return sid_from_url

    # Nu există ?sid= valid în URL.
    # FIX: Verificăm dacă avem deja un SID în session_state din acest run
    # (poate fi setat de JS prin query param la un rerun anterior în aceeași sesiune Streamlit).
    existing_in_state = st.session_state.get("session_id", "")
    if is_valid_session_id(existing_in_state):
        existing_in_state = _resolve_session_for_user(existing_in_state)
        st.session_state["session_id"] = existing_in_state
        # Repunem în URL pentru consistență (JS îl va citi și salva în localStorage)
        try:
            st.query_params["sid"] = existing_in_state
        except Exception:
            pass
        return existing_in_state

    # ── GATE: URL complet curat, fără SID nicăieri ──
    # Citim din localStorage SID-ul salvat (și lista de conversații ale browserului) prin
    # componenta session_bridge, care TRIMITE valorile spre Python. Înainte, un script încerca
    # window.parent.location.replace(...) din iframe-ul components.html; iframe-ul e izolat
    # (sandbox), iar Chrome/Edge blochează navigarea paginii-părinte → pagina rămânea blocată
    # la „Se încarcă...” și istoricul nu se mai restaura. Acum nu mai navigăm deloc:
    # Python primește SID-ul, îl pune în ?sid= și rularea continuă normal.
    candidate_id = generate_unique_session_id()
    _force_new = st.query_params.get("new", "") == "1"

    # Valorile din localStorage au fost citite de load_browser_storage_and_login()
    _stored = st.session_state.get("_browser_storage") or {}
    _stored_sid = str(_stored.get("sid") or "")
    if not _force_new and is_valid_session_id(_stored_sid):
        sid = _stored_sid          # browserul a mai folosit aplicația → restaurăm conversația
    else:
        sid = candidate_id         # vizită nouă (sau ?new=1 / stocare blocată)

    _known_ids = [s for s in (_stored.get("known") or []) if isinstance(s, str) and is_valid_session_id(s)]
    if _known_ids:
        _existing = set(st.session_state.get("_my_session_ids", []))
        st.session_state["_my_session_ids"] = list(_existing.union(_known_ids))

    sid = _resolve_session_for_user(sid)
    st.session_state["session_id"] = sid
    try:
        if "new" in st.query_params:
            del st.query_params["new"]
        st.query_params["sid"] = sid   # Streamlit actualizează URL-ul fără reîncărcare
    except Exception:
        pass
    return sid


# === MEMORY MANAGEMENT (FIX MEMORY LEAK) ===
def trim_session_messages():
    """Limitează mesajele din session_state pentru a preveni memory leak.
    Păstrează primul mesaj (contextul inițial) — consistent cu get_context_for_ai."""
    if "messages" in st.session_state:
        current_count = len(st.session_state.messages)

        if current_count > MAX_MESSAGES_IN_MEMORY:
            msgs = st.session_state.messages
            # Păstrăm primul mesaj + ultimele (MAX-1) => exact MAX (înainte rămâneau MAX+1,
            # deci se re-tăia și apărea toast-ul la fiecare mesaj nou).
            st.session_state.messages = [msgs[0]] + msgs[-(MAX_MESSAGES_IN_MEMORY - 1):]
            if not st.session_state.get("_trim_notified"):
                st.session_state["_trim_notified"] = True
                st.toast("📝 Conversația e lungă: mesajele vechi au fost arhivate pentru performanță.", icon="📦")


def summarize_conversation(messages: list) -> str | None:
    """Cere AI-ului să rezume conversația de până acum.
    
    Returnează textul rezumatului sau None dacă eșuează.
    Folosit pentru a comprima istoricul lung fără a pierde contextul.
    """
    if not messages or len(messages) < 6:
        return None
    try:
        # Trimitem doar primele mesaje (cele care vor fi comprimate)
        msgs_to_summarize = messages[:-MESSAGES_KEPT_AFTER_SUMMARY]
        if len(msgs_to_summarize) < 4:
            return None

        history_for_summary = []
        for msg in msgs_to_summarize:
            role = "model" if msg["role"] == "assistant" else "user"
            history_for_summary.append({"role": role, "parts": [msg["content"][:500]]})

        summary_prompt = (
            "Fă un rezumat SCURT (maxim 200 cuvinte) al conversației de mai sus. "
            "Include: subiectele discutate, conceptele explicate, exercițiile rezolvate "
            "și orice context important despre nivelul și înțelegerea elevului. "
            "Scrie la persoana a 3-a: 'Elevul a întrebat despre... Am explicat...'"
        )
        chunks = list(run_chat_with_rotation(history_for_summary, [summary_prompt]))
        summary = "".join(chunks).strip()
        return summary if len(summary) > 20 else None
    except Exception:
        return None  # Eșec silențios — nu întrerupem conversația


def get_context_for_ai(messages: list) -> list:
    """Pregătește contextul pentru AI cu limită de mesaje.

    Strategie:
    1. Dacă există un rezumat pre-generat (din sesiune anterioară sau conversație lungă):
       → rezumat + ultimele MESSAGES_KEPT_AFTER_SUMMARY mesaje recente
       Aceasta acoperă și cazul "revenirii din altă zi" cu oricâte mesaje în istoric.
    2. Sub MAX_MESSAGES_TO_SEND_TO_AI mesaje și fără rezumat: trimite totul
    3. Peste SUMMARIZE_AFTER_MESSAGES și fără rezumat: generează rezumat acum
    4. Fallback: primul mesaj + ultimele MAX_MESSAGES_TO_SEND_TO_AI
    """
    # ── Cazul 1: există deja un rezumat (pre-generat la revenire SAU generat anterior) ──
    # Îl folosim indiferent de numărul de mesaje — e mai bun decât trunchiere brută
    cached_summary = st.session_state.get("_conversation_summary")
    cached_at      = st.session_state.get("_summary_cached_at", 0)

    if cached_summary:
        # Regenerăm rezumatul la fiecare 10 mesaje noi față de ultima rezumare
        if (len(messages) - cached_at) >= 10:
            new_summary = summarize_conversation(messages)
            if new_summary:
                cached_summary = new_summary
                st.session_state["_conversation_summary"] = new_summary
                st.session_state["_summary_cached_at"]    = len(messages)

        summary_msg = {
            "role": "user",
            "content": (
                "[CONTEXT CONVERSAȚIE ANTERIOARĂ — citește înainte de a răspunde]\n"
                f"{cached_summary}\n"
                "[MESAJE RECENTE — continuare directă]"
            )
        }
        summary_ack = {
            "role": "assistant",
            "content": "Am înțeles contextul. Continuăm de unde am rămas."
        }
        recent = messages[-MESSAGES_KEPT_AFTER_SUMMARY:]
        return [summary_msg, summary_ack] + recent

    # ── Cazul 2: conversație scurtă — trimitem totul ──
    if len(messages) <= MAX_MESSAGES_TO_SEND_TO_AI:
        return messages

    # ── Cazul 3: conversație lungă fără rezumat — generăm acum ──
    if len(messages) >= SUMMARIZE_AFTER_MESSAGES:
        summary = summarize_conversation(messages)
        if summary:
            st.session_state["_conversation_summary"] = summary
            st.session_state["_summary_cached_at"]    = len(messages)
            summary_msg = {
                "role": "user",
                "content": (
                    "[CONTEXT CONVERSAȚIE ANTERIOARĂ — citește înainte de a răspunde]\n"
                    f"{summary}\n"
                    "[MESAJE RECENTE — continuare directă]"
                )
            }
            summary_ack = {
                "role": "assistant",
                "content": "Am înțeles contextul. Continuăm de unde am rămas."
            }
            recent = messages[-MESSAGES_KEPT_AFTER_SUMMARY:]
            return [summary_msg, summary_ack] + recent

    # ── Cazul 4: fallback — primul mesaj + ultimele MAX_MESSAGES_TO_SEND_TO_AI ──
    first_message  = messages[0] if messages else None
    recent_messages = messages[-MAX_MESSAGES_TO_SEND_TO_AI:]
    if first_message and first_message not in recent_messages:
        return [first_message] + recent_messages
    return recent_messages


def save_message_with_limits(session_id: str, role: str, content: str):
    """Salvează mesaj și verifică limitele."""
    save_message_to_db(session_id, role, content)
    invalidate_session_cache()  # FIX: un mesaj nou înseamnă date noi în sidebar
    
    # Rulează trim în același thread — Streamlit nu e thread-safe.
    # Contor de salvări (nu len(messages) % 50: lista e plafonată la MAX_MESSAGES_IN_MEMORY,
    # deci condiția devenea adevărată la FIECARE mesaj odată ajuns la plafon).
    _saves = st.session_state.get("_db_saves_since_trim", 0) + 1
    if _saves >= 50:
        trim_db_messages(session_id)
        _saves = 0
    st.session_state["_db_saves_since_trim"] = _saves
    
    trim_session_messages()






# === SVG FUNCTIONS ===

# ÎMBUNĂTĂȚIRE 4: lxml pentru parsare și validare SVG robustă.
# Fallback automat la regex dacă lxml nu e disponibil.
try:
    from lxml import etree as _lxml_etree
    _LXML_AVAILABLE = True
except ImportError:
    _LXML_AVAILABLE = False


def repair_unclosed_tags(svg_content: str) -> str:
    """Repară tag-uri SVG comune care nu sunt închise corect."""
    self_closing_tags = ['path', 'rect', 'circle', 'ellipse', 'line', 'polyline', 'polygon', 'image', 'use']
    
    for tag in self_closing_tags:
        # FIX: pattern mai robust — nu atinge tag-uri deja self-closing
        pattern = rf'<{tag}(\s[^>]*)?>(?!</{tag}>)'
        
        def fix_tag(match, _tag=tag):
            attrs = match.group(1) or ""
            # Dacă are deja / la final, e deja corect
            if attrs.rstrip().endswith('/'):
                return match.group(0)
            return f'<{_tag}{attrs}/>'
        
        svg_content = re.sub(pattern, fix_tag, svg_content)
    
    text_opens = len(re.findall(r'<text[^>]*>', svg_content))
    text_closes = len(re.findall(r'</text>', svg_content))
    
    if text_opens > text_closes:
        for _ in range(text_opens - text_closes):
            svg_content = svg_content.replace('</svg>', '</text></svg>')
    
    g_opens = len(re.findall(r'<g[^>]*>', svg_content))
    g_closes = len(re.findall(r'</g>', svg_content))
    
    if g_opens > g_closes:
        for _ in range(g_opens - g_closes):
            svg_content = svg_content.replace('</svg>', '</g></svg>')
    
    return svg_content



def repair_svg(svg_content: str) -> str:
    """Repară SVG incomplet sau malformat.

    ÎMBUNĂTĂȚIRE 4: Încearcă mai întâi repararea cu lxml (parser XML tolerant),
    care gestionează corect namespace-uri, encoding și structura arborescentă.
    Fallback la regex dacă lxml eșuează sau nu e disponibil.
    """
    if not svg_content:
        return None

    svg_content = svg_content.strip()

    # Pasul 1: asigură tag-uri <svg> deschis/închis
    has_svg_open  = bool(re.search(r'<svg[^>]*>', svg_content, re.IGNORECASE))
    has_svg_close = '</svg>' in svg_content.lower()

    if not has_svg_open:
        svg_content = (
            '<svg viewBox="0 0 800 600" xmlns="http://www.w3.org/2000/svg" '
            'style="max-width:100%;height:auto;">\n'
            + svg_content + '\n</svg>'
        )
    elif has_svg_open and not has_svg_close:
        svg_content += '\n</svg>'

    if 'xmlns=' not in svg_content:
        svg_content = svg_content.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"', 1)
    if 'viewbox=' not in svg_content.lower():
        svg_content = svg_content.replace('<svg', '<svg viewBox="0 0 800 600"', 1)

    # Pasul 2: repară cu lxml dacă e disponibil
    if _LXML_AVAILABLE:
        try:
            parser = _lxml_etree.XMLParser(
                recover=True,
                remove_comments=False,
                resolve_entities=False,
                ns_clean=True,
            )
            root = _lxml_etree.fromstring(svg_content.encode("utf-8"), parser)
            repaired = _lxml_etree.tostring(
                root,
                pretty_print=True,
                encoding="unicode",
                xml_declaration=False
            )
            return repaired
        except Exception:
            pass  # lxml a eșuat → continuăm cu fallback

    # Pasul 3: fallback regex
    svg_content = repair_unclosed_tags(svg_content)
    return svg_content


def validate_svg(svg_content: str) -> tuple:
    """Validează SVG și returnează (is_valid, error_message).

    ÎMBUNĂTĂȚIRE 4: Folosește lxml pentru validare structurală când e disponibil.
    """
    if not svg_content:
        return False, "SVG gol"

    visual_elements = ['path', 'rect', 'circle', 'ellipse', 'line', 'text', 'polygon', 'polyline', 'image']

    if _LXML_AVAILABLE:
        try:
            parser = _lxml_etree.XMLParser(recover=True, resolve_entities=False, no_network=True)
            tree = _lxml_etree.fromstring(svg_content.encode("utf-8"), parser)
            has_content = any(f'<{el}' in svg_content.lower() for el in visual_elements)
            if not has_content:
                return False, "SVG fără elemente vizuale"
            return True, "OK"
        except Exception as xml_err:
            # lxml a eșuat complet — încercăm fallback simplu
            pass

    # Fallback validare simplă
    if '<svg' not in svg_content.lower():
        return False, "Lipsește tag-ul <svg>"
    if '</svg>' not in svg_content.lower():
        return False, "Lipsește tag-ul </svg>"
    has_content = any(f'<{elem}' in svg_content.lower() for elem in visual_elements)
    if not has_content:
        return False, "SVG fără elemente vizuale"
    return True, "OK"


def sanitize_svg(svg_content: str) -> str:
    """Sanitizeaza SVG - elimina scripturi si event handlers (XSS prevention).
    
    Acopera: <script>, on* handlers (ghilimele/backtick), href=javascript:,
    use href=data:, style behavior/expression, <foreignObject>.
    """
    if not svg_content:
        return svg_content
    # Elimina <script> complet
    svg_content = re.sub(r'<script\b[^>]*>.*?</script\s*>', '', svg_content,
                         flags=re.DOTALL | re.IGNORECASE)
    # Elimina event handlers on* cu ghilimele duble
    svg_content = re.sub(r'\s+on[a-zA-Z]+\s*=\s*"[^"]*"', '', svg_content)
    # Elimina event handlers on* cu ghilimele simple
    svg_content = re.sub(r"\s+on[a-zA-Z]+\s*=\s*'[^']*'", '', svg_content)
    # Elimina event handlers on* cu backtick (template literals)
    svg_content = re.sub(r'\s+on[a-zA-Z]+\s*=\s*`[^`]*`', '', svg_content)
    # Elimina href=javascript: si xlink:href=javascript:
    svg_content = re.sub(r'(xlink:)?href\s*=\s*["\']?\s*javascript:[^"\'>\s]*["\']?', '',
                         svg_content, flags=re.IGNORECASE)
    # Elimina <use href="data:..."> — poate injecta SVG/HTML extern
    svg_content = re.sub(r'<use\b[^>]*href\s*=\s*["\']data:[^"\']*["\'][^>]*>', '',
                         svg_content, flags=re.IGNORECASE)
    # Elimina style cu behavior: sau expression( (vector de atac IE/vechi)
    svg_content = re.sub(r'style\s*=\s*["\'][^"\']*(?:behavior|expression)\s*:[^"\']*["\']', '',
                         svg_content, flags=re.IGNORECASE)
    # Elimina <foreignObject> — permite injectare HTML arbitrar in SVG
    svg_content = re.sub(r'<foreignObject\b.*?</foreignObject\s*>', '', svg_content,
                         flags=re.DOTALL | re.IGNORECASE)
    return svg_content



def _is_gfile_active(gfile) -> bool:
    """Verifică dacă un fișier Google este activ — helper consistent folosit peste tot."""
    state_str = str(gfile.state)
    state_name = getattr(gfile.state, "name", "")
    return state_str in ("FileState.ACTIVE", "ACTIVE") or state_name == "ACTIVE"


def _svg_is_wellformed(svg_content: str) -> bool:
    """Un SVG afișat ca <img> trebuie să fie XML bine format (browserul nu tolerează erori)."""
    try:
        if _LXML_AVAILABLE:
            _lxml_etree.fromstring(
                svg_content.encode("utf-8"),
                _lxml_etree.XMLParser(resolve_entities=False, no_network=True),
            )
        else:
            import xml.etree.ElementTree as _ET
            _ET.fromstring(svg_content)
        return True
    except Exception:
        return False


def _render_svg_as_image(svg_content: str, bg: str):
    """Afișează SVG-ul ca <img src="data:image/svg+xml;base64,...">.

    În contextul <img> browserul NU execută scripturi și nu încarcă resurse externe,
    deci un SVG malițios nu poate citi localStorage (cheia API) și nu poate face XSS —
    indiferent cât de bine sau prost a fost curățat de sanitize_svg().
    """
    import base64 as _b64
    b64 = _b64.b64encode(svg_content.encode("utf-8")).decode("ascii")
    st.markdown(
        '<div style="text-align:center;margin:10px 0;">'
        f'<img alt="Desen" src="data:image/svg+xml;base64,{b64}" '
        f'style="width:100%;max-width:900px;height:auto;background:{bg};'
        'border-radius:8px;padding:10px;box-sizing:border-box;"/>'
        '</div>',
        unsafe_allow_html=True,
    )


def render_message_with_svg(content: str):
    """Renderează mesajul cu suport îmbunătățit pentru SVG."""
    has_svg_markers = '[[DESEN_SVG]]' in content
    # Regex precis: detectează doar blocuri SVG complete, nu menționări în text
    # FIX bug 3: \b word boundary corect — previne match pe tag-uri ca <svgfoo>
    has_svg_elements = bool(re.search(r'<svg\b[^>]*>.*?</svg\s*>', content, re.DOTALL | re.IGNORECASE))
    has_svg_sub_elements = any(tag in content.lower() for tag in ['<path', '<rect', '<circle', '<line', '<polygon'])
    
    if has_svg_markers or (has_svg_elements) or (has_svg_sub_elements and 'stroke=' in content):
        svg_code = None
        before_text = ""
        after_text = ""
        
        if '[[DESEN_SVG]]' in content:
            parts = content.split('[[DESEN_SVG]]')
            before_text = parts[0]
            if len(parts) > 1 and '[[/DESEN_SVG]]' in parts[1]:
                inner_parts = parts[1].split('[[/DESEN_SVG]]')
                svg_code = inner_parts[0]
                after_text = inner_parts[1] if len(inner_parts) > 1 else ""
            elif len(parts) > 1:
                svg_code = parts[1]
        elif '<svg' in content.lower():
            svg_match = re.search(r'<svg.*?</svg>', content, re.DOTALL | re.IGNORECASE)
            if svg_match:
                svg_code = svg_match.group(0)
                before_text = content[:svg_match.start()]
                after_text = content[svg_match.end():]
            else:
                svg_start = content.lower().find('<svg')
                if svg_start != -1:
                    before_text = content[:svg_start]
                    svg_code = content[svg_start:]
        
        if svg_code:
            svg_code = sanitize_svg(svg_code)
            svg_code = repair_svg(svg_code) or ""
            # Injectam <style> direct in SVG dupa primul tag <svg>
            # (într-un <img> CSS-ul paginii nu ajunge în SVG, deci culorile se pun aici)
            _dark_svg = st.session_state.get("dark_mode", False)
            _bg   = "#0e1117" if _dark_svg else "#ffffff"
            _text = "#fafafa" if _dark_svg else "#1a1a1a"
            _style_inject = (
                "<style>"
                "svg{background:transparent!important}"
                "rect[id='bg'],rect[id='background'],rect.bg,rect.background{display:none!important}"
                f"text{{fill:{_text}!important}}"
                f"rect[fill='white'],rect[fill='#fff'],rect[fill='#ffffff'],"
                f"rect[fill='#FFF'],rect[fill='#FFFFFF']{{fill:{_bg}!important}}"
                "</style>"
            )
            svg_code = re.sub(r'(<svg[^>]*>)', lambda m: m.group(1) + _style_inject, svg_code, count=1)
            is_valid, error = validate_svg(svg_code)
            if is_valid and not _svg_is_wellformed(svg_code):
                is_valid, error = False, "SVG-ul nu e XML valid"

            if is_valid:
                if before_text.strip():
                    st.markdown(before_text.strip())

                _render_svg_as_image(svg_code, _bg)
                
                if after_text.strip():
                    st.markdown(after_text.strip())
                return
            else:
                st.warning(f"⚠️ Desenul nu a putut fi afișat corect: {error}")
    
    clean_content = content
    clean_content = re.sub(r'\[\[DESEN_SVG\]\]', '\n🎨 *Desen:*\n', clean_content)
    clean_content = re.sub(r'\[\[/DESEN_SVG\]\]', '\n', clean_content)
    
    st.markdown(clean_content)


# === INIȚIALIZARE ===
init_db()
cleanup_old_sessions(CLEANUP_DAYS_OLD)

# Partea comercială (comercial/): conturi, acord parental, documente legale, abonamente.
comercial.core.setup(get_client=get_supabase_client, log=_log, send_email=_send_email,
                      cleanup_days=CLEANUP_DAYS_OLD)

# Documentele legale ca pagini publice (?doc=termeni / ?doc=confidentialitate) — linkate din Stripe și emailuri.
_doc_public = st.query_params.get("doc", "")
if _doc_public in comercial.core.LEGAL_VERSIONS:
    st.markdown(comercial.ui.legal_text(_doc_public))
    st.stop()

# Linkul din emailul trimis părintelui: pagină separată, fără sesiune de chat.
_parinte_token = st.query_params.get("parinte", "")
if _parinte_token:
    comercial.ui.render_parent_page(_parinte_token)   # se oprește singură (st.stop)

# Citește din browser SID-ul salvat + tokenul de login și restaurează contul (o dată per sesiune)
load_browser_storage_and_login()
if comercial.core.billing_on():
    refresh_current_user()

# Python generează/restaurează SID — poate pune ?sid= în URL pentru JS
session_id = get_or_create_session_id()
st.session_state.session_id = session_id
update_session_activity(session_id)

# JS citește ?sid= din URL (dacă Python l-a pus) și îl salvează în localStorage
# La revenire după restart: JS citește SID din localStorage și face reload cu ?sid=
inject_session_js()


# === API KEYS ===
#
# Prioritate:
#   1. Cheile din st.secrets (ale tale) — folosite primele, rotite automat
#   2. Cheia manuală a elevului din localStorage — folosită când ale tale
#      sunt epuizate SAU dacă nu ai setat nicio cheie în secrets
#
# Cheia elevului e salvată în localStorage al browserului său:
#   - supraviețuiește refresh-ului și închiderii tab-ului
#   - dispare doar dacă elevul apasă "Șterge cheia" sau golește browserul

# ── Pasul 1: citește cheia elevului din session_state (salvată direct, fără URL)
# FIX 1: cheia NU mai vine prin ?apikey= în URL — e salvată direct în session_state
# la click pe "Salvează cheia" și persistată în localStorage de JS via _saveApiKeyToStorage()
# Restaurăm cheia din localStorage (o singură dată pe sesiune Streamlit) și executăm
# comenzile save/clear trimise de butoanele din sidebar. Doar dacă nu ai chei în secrets.
_has_owner_keys = ("GOOGLE_API_KEYS" in st.secrets) or ("GOOGLE_API_KEY" in st.secrets)
_kb_cmd = st.session_state.pop("_kb_cmd", None)
if _key_bridge is not None and not _has_owner_keys:
    _kb_val = _key_bridge(cmd=_kb_cmd, key="_kb_bridge", default=None)
    if _kb_val is not None and not st.session_state.get("_kb_synced"):
        st.session_state["_kb_synced"] = True
        if (
            isinstance(_kb_val, str)
            and 15 <= len(_kb_val) <= 200
            and not any(ch.isspace() for ch in _kb_val)
            and not st.session_state.get("_manual_api_key")
        ):
            st.session_state["_manual_api_key"] = _kb_val

saved_manual_key = st.session_state.get("_manual_api_key", "")

# ── Pasul 2: construiește lista de chei (secrets + manuală) ──
raw_keys_secrets = None
if "GOOGLE_API_KEYS" in st.secrets:
    raw_keys_secrets = st.secrets["GOOGLE_API_KEYS"]
elif "GOOGLE_API_KEY" in st.secrets:
    raw_keys_secrets = [st.secrets["GOOGLE_API_KEY"]]

keys = []

# Adaugă cheile din secrets
if raw_keys_secrets:
    if isinstance(raw_keys_secrets, str):
        # Securitate: json.loads în loc de ast.literal_eval (mai sigur împotriva injection)
        import json as _json
        try:
            parsed = _json.loads(raw_keys_secrets)
            if isinstance(parsed, list):
                raw_keys_secrets = parsed
            else:
                raw_keys_secrets = [raw_keys_secrets]
        except (_json.JSONDecodeError, ValueError):
            # Fallback: split manual după virgulă, fără eval
            raw_keys_secrets = [k.strip().strip('"').strip("'")
                                 for k in raw_keys_secrets.split(",") if k.strip()]
    if isinstance(raw_keys_secrets, list):
        for k in raw_keys_secrets:
            if k and isinstance(k, str):
                clean_k = k.strip().strip('"').strip("'")
                if clean_k:
                    keys.append(clean_k)

# Cheile salvate în contul elevului (mama, tata, rude...) — rotite automat când se epuizează
_acct_items = acct_keys_load() if current_user() else []
for _it in _acct_items:
    if _it["key"] not in keys:
        keys.append(_it["key"])

# Adaugă cheia elevului salvată doar în browser (folosită când celelalte se epuizează)
if saved_manual_key and saved_manual_key not in keys:
    keys.append(saved_manual_key)

# ── Pasul 3: UI în sidebar pentru cheia manuală ──
# Afișăm secțiunea DOAR dacă nu există chei configurate în secrets
def _render_account_ui():
    """Secțiunea „👤 Cont” din bara laterală (login, cont nou, resetare parolă, ieșire)."""
    # --- Cont (opțional) ---
    _user = current_user()
    if _user:
        st.markdown(f"👤 **{_user['name']}**  \n<small>{_user['email']}</small>", unsafe_allow_html=True)
        st.caption("Conversațiile tale sunt salvate în cont și apar pe orice dispozitiv.")
        _rec_show = st.session_state.get("_auth_new_recovery")
        if _rec_show:
            st.warning("🔑 **Codul tău de recuperare** — notează-l! Cu el îți poți reseta parola "
                       "dacă o uiți. Îl vezi doar acum.", icon="⚠️")
            st.code(_rec_show, language=None)
            st.download_button("⬇️ Salvează codul (.txt)", key="_rec_dl",
                               data=(f"Profesor Liceu — cod de recuperare\nCont: {_user['email']}\n"
                                     f"Cod: {_rec_show}\n").encode("utf-8"),
                               file_name="profesor_liceu_cod_recuperare.txt", use_container_width=True)
            st.button("✅ Am notat codul", key="_rec_ok", on_click=_dismiss_recovery, use_container_width=True)
        else:
            with st.expander("🔑 Cod de recuperare", expanded=False):
                st.caption("Dacă nu mai ai codul de recuperare, generează unul nou "
                           "(cel vechi nu va mai funcționa).")
                if st.button("Generează cod nou", key="_rec_new", use_container_width=True):
                    if not auth_regenerate_recovery():
                        st.error("Nu am putut genera codul acum.")
                    st.rerun()
        _pf = st.session_state.pop("_auth_parent_flash", None)
        if _pf:
            (st.success if _pf[0] else st.error)(_pf[1])
        comercial.ui.render_profile_form(_user, globals().get("PROFILE_BAC", {}))
        comercial.ui.render_subscription(_user)
        comercial.ui.render_privacy_tools(
            _user,
            verify_password=lambda _em, _pw: auth_login(_em, _pw)[0] is not None,
            after_delete=lambda: auth_logout(),
        )
        comercial.ui.render_admin_panel(_user)
        _lo1, _lo2 = st.columns(2)
        with _lo1:
            if st.button("🚪 Ieșire", key="_auth_logout", use_container_width=True):
                auth_logout()
                st.rerun()
        with _lo2:
            if st.button("Ieșire peste tot", key="_auth_logout_all", use_container_width=True,
                         help="Deconectează contul de pe toate dispozitivele"):
                auth_logout(all_devices=True)
                st.rerun()
    else:
        with st.expander("👤 Cont — conversațiile tale pe orice dispozitiv", expanded=False):
            _tab_in, _tab_new, _tab_reset = st.tabs(["Intră în cont", "Cont nou", "Am uitat parola"])
            with _tab_in:
                with st.form("_auth_login_form", clear_on_submit=False):
                    _le = st.text_input("Email", key="_auth_login_email", autocomplete="email")
                    _lp = st.text_input("Parolă", type="password", key="_auth_login_pw",
                                        autocomplete="current-password")
                    _lr = st.checkbox("Ține-mă minte pe acest dispozitiv", value=True, key="_auth_login_rem")
                    if st.form_submit_button("Intră", type="primary", use_container_width=True):
                        _u, _err = auth_login(_le, _lp)
                        if _u:
                            auth_complete_login(_u, _lr)
                            st.rerun()
                        else:
                            st.error(_err)
            with _tab_new:
                _strict = comercial.core.billing_on()
                if _strict:
                    comercial.ui.render_legal_links("_reg_legal")
                with st.form("_auth_register_form", clear_on_submit=False):
                    _rn = st.text_input("Nume (opțional)", key="_auth_reg_name", max_chars=60)
                    _re_ = st.text_input("Email", key="_auth_reg_email", autocomplete="email")
                    _rp = st.text_input(f"Parolă (minim {AUTH_MIN_PASSWORD} caractere)", type="password",
                                        key="_auth_reg_pw", autocomplete="new-password")
                    _rp2 = st.text_input("Repetă parola", type="password", key="_auth_reg_pw2",
                                         autocomplete="new-password")
                    _by = _bm = None
                    _pe, _acc = "", False
                    if _strict:
                        _yr = time.localtime().tm_year
                        _c1, _c2 = st.columns(2)
                        with _c1:
                            _by = st.selectbox("Anul nașterii", [None] + list(range(_yr - 10, _yr - 100, -1)),
                                               key="_auth_reg_by", format_func=lambda x: "— an —" if x is None else str(x))
                        with _c2:
                            _bm = st.selectbox("Luna", [None] + list(range(1, 13)), key="_auth_reg_bm",
                                               format_func=lambda x: "— lună —" if x is None else f"{x:02d}")
                        _pe = st.text_input("Emailul unui părinte (obligatoriu sub 18 ani)", key="_auth_reg_parent",
                                            autocomplete="off")
                        st.caption("Sub 16 ani, contul se activează după acordul părintelui (îi trimitem un link). "
                                   "Sub 18 ani, abonamentul îl plătește părintele.")
                        _acc = st.checkbox("Am peste 16 ani, am citit și accept Termenii și condițiile și Politica "
                                           "de confidențialitate.", key="_auth_reg_terms")
                    _rr = st.checkbox("Ține-mă minte pe acest dispozitiv", value=True, key="_auth_reg_rem")
                    st.caption("Conversațiile din acest browser vor fi mutate în contul nou.")
                    if st.form_submit_button("Creează cont", type="primary", use_container_width=True):
                        if _rp != _rp2:
                            st.error("Parolele nu coincid.")
                        else:
                            _u, _err = auth_register(_re_, _rp, _rn, _by, _bm, _pe, _acc)
                            if _u:
                                auth_complete_login(_u, _rr)
                                st.rerun()
                            else:
                                st.error(_err)
            with _tab_reset:
                if _smtp_config():
                    with st.form("_auth_reset_mail_form"):
                        st.markdown("**1.** Primește un cod pe email")
                        _rme = st.text_input("Email", key="_auth_reset_mail_email", autocomplete="email")
                        if st.form_submit_button("📧 Trimite codul", use_container_width=True):
                            _sent, _msg = auth_request_reset_email(_rme)
                            (st.success if _sent else st.error)(_msg)
                    st.markdown("**2.** Introdu codul din email *sau* codul de recuperare")
                else:
                    st.caption("Introdu codul de recuperare primit la crearea contului.")
                with st.form("_auth_reset_form"):
                    _xe = st.text_input("Email", key="_auth_reset_email", autocomplete="email")
                    _xc = st.text_input("Cod (6 cifre din email sau XXXX-XXXX-XXXX)", key="_auth_reset_code")
                    _xp = st.text_input(f"Parolă nouă (minim {AUTH_MIN_PASSWORD} caractere)", type="password",
                                        key="_auth_reset_pw", autocomplete="new-password")
                    _xp2 = st.text_input("Repetă parola nouă", type="password", key="_auth_reset_pw2",
                                         autocomplete="new-password")
                    if st.form_submit_button("Schimbă parola", type="primary", use_container_width=True):
                        if _xp != _xp2:
                            st.error("Parolele nu coincid.")
                        else:
                            _u, _err = auth_reset_password(_xe, _xc, _xp)
                            if _u:
                                auth_complete_login(_u, True)
                                st.toast("✅ Parola a fost schimbată. Celelalte dispozitive au fost deconectate.",
                                         icon="🔑")
                                st.rerun()
                            else:
                                st.error(_err)


_acct_key_set = {it["key"] for it in _acct_items}
_are_secrets_keys = len([k for k in keys if k != saved_manual_key and k not in _acct_key_set]) > 0

_KEY_GUIDE_MD = """
**Pasul 1** — Deschide Google AI Studio (butonul de mai sus) și autentifică-te cu contul Google.

**Pasul 2** — Apasă **"Create API key"** (buton albastru); dacă ți se cere, alege **"Create API key in new project"**.

**Pasul 3** — Copiază cheia afișată (`AIzaSy...` sau `AQ.Ab8R...`) și lipește-o mai jos.

💡 Fiecare cont Google are cota lui gratuită (cereri pe minut și pe zi). Cu chei din mai multe conturi
(ex. de la părinți sau rude care nu le folosesc) ai de atâtea ori mai multe întrebări pe zi.
"""


def _render_account_keys_ui():
    """Cheile din cont: listă cu 🗑️ + formular de adăugare (vizibil doar pentru utilizatorii logați)."""
    st.subheader("🔑 Cheile tale API Google AI")
    items = acct_keys_load()
    if items:
        st.caption(f"{len(items)} {'cheie salvată' if len(items) == 1 else 'chei salvate'} în cont, pe orice dispozitiv. "
                   "Când una își termină cota, trece automat la următoarea.")
        for it in items:
            c1, c2 = st.columns([5, 1])
            with c1:
                st.markdown(f"🔑 **{it['label']}** · <small>…{it['hint']}</small>", unsafe_allow_html=True)
            with c2:
                st.button("🗑️", key=f"_acct_key_del_{it['id']}", help=f"Șterge cheia „{it['label']}”",
                          on_click=acct_key_delete, args=(it["id"],))
    else:
        st.caption("Nu ai încă nicio cheie salvată în cont.")

    if saved_manual_key and saved_manual_key not in _acct_key_set and len(items) < MAX_ACCOUNT_KEYS:
        st.info("Ai o cheie salvată doar în acest browser.")
        if st.button("☁️ Mut-o în cont", key="_acct_key_move", use_container_width=True):
            _ok, _note = acct_key_add("Cheia mea", saved_manual_key)
            if _ok:
                st.session_state.pop("_manual_api_key", None)
                st.session_state["_kb_synced"] = True
                st.session_state["_kb_cmd"] = {"op": "clear", "n": secrets.token_hex(4)}
                st.rerun()
            else:
                st.error(_note)

    if len(items) < MAX_ACCOUNT_KEYS:
        with st.expander("➕ Adaugă o cheie", expanded=not items and not saved_manual_key):
            st.link_button("🌐 Ia o cheie gratuită din Google AI Studio", "https://aistudio.google.com/apikey",
                           use_container_width=True)
            with st.popover("❓ Cum obțin o cheie?", use_container_width=True):
                st.markdown(_KEY_GUIDE_MD)
            _key_photo_reader("_acct_key_value", _ocr_keys_available(keys))
            _key_photo_message("_acct_key_value")
            with st.form("_acct_key_add_form", clear_on_submit=True):
                _lbl = st.text_input("Al cui cont Google e cheia?", placeholder="ex: Mama, Tata, Bunicul",
                                     max_chars=40, key="_acct_key_label")
                _val = st.text_input("Cheia API", type="password", placeholder="AIzaSy... sau AQ.Ab8R...",
                                     key="_acct_key_value",
                                     help="Lipește cheia sau folosește poza de mai sus — se completează singură.")
                if st.form_submit_button("✅ Adaugă în cont", type="primary", use_container_width=True):
                    _ok, _note = acct_key_add(_lbl, _val)
                    if _ok:
                        st.session_state["_acct_key_flash"] = ("ok", _note or "✅ Cheie adăugată în cont.")
                        st.rerun()
                    else:
                        st.error(_note)
    _flash = st.session_state.pop("_acct_key_flash", None)
    if _flash:
        (st.success if _flash[0] == "ok" else st.warning)(_flash[1])


with st.sidebar:
    if comercial.core.billing_on():
        pass   # versiunea comercială folosește doar cheile proiectului plătit (nu chei ale elevilor)
    elif current_user():
        st.divider()
        _render_account_keys_ui()
    elif not _are_secrets_keys:
        st.divider()
        st.subheader("🔑 Cheie API Google AI")
        st.caption("💡 Cu un cont (👤 Cont, mai sus) poți salva mai multe chei — de exemplu de la părinți — "
                   "și le ai pe orice dispozitiv.")

        if not saved_manual_key:
            # ── Ghid vizual — vizibil DOAR când nu există cheie salvată ──
            with st.expander("❓ Cum obțin o cheie? (gratuit)", expanded=False):
                st.markdown("**Ai nevoie de un cont Google** (Gmail). Este complet gratuit.")
                st.markdown("**Pasul 1** — Deschide Google AI Studio:")
                st.link_button(
                    "🌐 Mergi la aistudio.google.com",
                    "https://aistudio.google.com/apikey",
                    use_container_width=True
                )
                st.markdown("""
**Pasul 2** — Autentifică-te cu contul Google.

**Pasul 3** — Apasă **"Create API key"** (buton albastru).

**Pasul 4** — Dacă ți se cere, alege **"Create API key in new project"**.

**Pasul 5** — Copiază cheia afișată.
- Poate arăta astfel: `AIzaSy...` (format vechi) sau `AQ.Ab8R...` (format nou Google)
- Apasă iconița 📋 de lângă cheie

**Pasul 6** — Lipește cheia mai jos și apasă **Salvează**.

---
💡 **Limită gratuită:** depinde de model și se aplică per proiect Google (cereri pe minut și pe zi); o vezi în Google AI Studio. De obicei e suficientă pentru teme și exerciții.
                """)

            # ── Câmpul de input și butonul de salvare ──
            st.caption("Cheia se salvează în browserul tău și rămâne activă după refresh.")
            _key_photo_reader("_manual_key_input", _ocr_keys_available(keys))
            _key_photo_message("_manual_key_input")
            new_key = st.text_input(
                "Cheie API Google AI:",
                type="password",
                placeholder="AIzaSy... sau AQ.Ab8R...",
                label_visibility="collapsed",
                key="_manual_key_input",
            )
            if st.button("✅ Salvează cheia", use_container_width=True, type="primary", key="save_api_key"):
                clean = new_key.strip().strip('"').strip("'")
                # FIX FORMAT CHEIE: Google a schimbat formatul cheilor API (vechi: "AIza...",
                # nou: "AQ.Ab8R..."), iar validarea veche bloca cheile noi. Acum acceptăm
                # ORICE format de cheie — nu mai verificăm un prefix fix, doar reguli
                # minimale de bun-simț: lungime rezonabilă și fără spații/caractere de control
                # (o cheie API reală nu conține spații).
                is_plausible_key = (
                    clean
                    and 15 <= len(clean) <= 200
                    and " " not in clean
                    and "\n" not in clean
                    and "\t" not in clean
                )
                if is_plausible_key:
                    st.session_state["_manual_api_key"] = clean
                    keys.append(clean)
                    # Salvăm în localStorage prin podul key_bridge (rulează la rerun-ul de mai jos)
                    st.session_state["_kb_synced"] = True
                    st.session_state["_kb_cmd"] = {"op": "save", "key": clean, "n": secrets.token_hex(4)}
                    st.toast("✅ Cheie salvată în browser!", icon="🔑")
                    st.rerun()
                else:
                    st.error("❌ Cheie invalidă. Verifică să nu conțină spații și să aibă minim 15 caractere.")

        else:
            # Cheia e salvată — arată doar statusul și butonul de ștergere, fără ghid
            st.success("🔑 Cheie personală activă.")
            st.caption("Salvată în browserul tău — rămâne după refresh.")
            if st.button("🗑️ Șterge cheia", use_container_width=True, key="del_api_key"):
                st.session_state.pop("_manual_api_key", None)
                st.query_params.pop("apikey", None)
                st.session_state["_kb_synced"] = True
                st.session_state["_kb_cmd"] = {"op": "clear", "n": secrets.token_hex(4)}
                st.rerun()

if not keys:
    # Fără cheie: lăsăm totuși login-ul disponibil — elevul își poate recupera cheile din cont.
    with st.sidebar:
        st.divider()
        _render_account_ui()
    st.error("❌ Nicio cheie API validă. Introdu cheia ta Google AI în bara laterală "
             "sau intră în cont (👤 Cont) dacă ți-ai salvat cheile acolo.")
    st.stop()

if "key_index" not in st.session_state:
    # FIX 4: distribuție uniformă bazată pe hash-ul SID-ului, nu random.randint.
    # random.randint independent per sesiune nu garantează echilibru când 100 de elevi
    # deschid simultan — toți pot nimeri pe aceeași cheie prin coincidență.
    # hash(session_id) % len(keys) distribuie determinist și uniform pe chei.
    _num_keys = max(len(keys), 1)
    st.session_state.key_index = int(hashlib.md5(session_id.encode()).hexdigest(), 16) % _num_keys
# Salvăm lista de chei în session_state — necesară pentru _cleanup_gfiles la switch sesiune
st.session_state["_api_keys_list"] = keys


# === MATERII ===
MATERII = {
    "🤖 Automat":         None,  # detectează materia din mesaj, întreabă dacă nu poate
    "📐 Matematică":      "matematică",
    "⚡ Fizică Real":        "fizică_real",
    "⚡ Fizică Tehnologic":  "fizică_tehnologic",
    "🧪 Chimie":          "chimie",
    "📖 Română":          "limba și literatura română",
    "🇫🇷 Franceză":       "limba franceză",
    "🇬🇧 Engleză":        "limba engleză",
    "🇩🇪 Germană":        "limba germană",
    "🌍 Geografie":       "geografie",
    "🏛️ Istorie":         "istorie",
    "💻 Informatică":     "informatică",
    "🧬 Biologie":        "biologie",
}

# Label-ul modului automat — folosit în mai multe locuri
_AUTOMAT_LABEL = "🤖 Automat"

# Mapare inversă cod → label (pentru toast-uri și afișări)
_MATERII_LABEL = {v: k for k, v in MATERII.items() if v is not None}



# ═══════════════════════════════════════════════════════════════
# PROMPT MODULAR — fiecare materie are blocul ei separat.
# get_system_prompt() include DOAR blocul materiei selectate,
# reducând tokenii de input cu 71-94% față de promptul complet.
# ═══════════════════════════════════════════════════════════════

_PROMPT_COMUN = r"""
    REGULI DE IDENTITATE (STRICT):
    1. Folosește EXCLUSIV genul masculin când vorbești despre tine.
       - Corect: "Sunt sigur", "Sunt pregătit", "Am fost atent", "Sunt bucuros".
       - GREȘIT: "Sunt sigură", "Sunt pregătită".
    2. Te prezinți simplu, fără nicio titulatură pompoasă.

    TON ȘI ADRESARE (CRITIC):
    3. Vorbește DIRECT, la persoana I singular.
       - CORECT: "Salut, sunt aici să te ajut." / "Te ascult." / "Sunt pregătit." / "Înțeleg!"
       - GREȘIT: "Înțeleg, Domnule Profesor!" / "Bineînțeles, Domnule Profesor!" / "Domnul profesor este aici." / "Profesorul te va ajuta."
       - NU folosi NICIODATĂ "Domnule Profesor" sau orice titulatură — tu ești profesorul, nu elevul.
    4. Fii cald, natural, apropiat și scurt. Evită introducerile pompoase.
    5. NU SALUTA în fiecare mesaj. Salută DOAR la începutul unei conversații noi.
    6. Dacă elevul pune o întrebare directă, răspunde DIRECT la subiect, fără introduceri de genul "Salut, desigur...".
    7. Folosește "Salut" sau "Te salut" în loc de formule foarte oficiale.

    REGULĂ STRICTĂ: Predă exact ca la școală (nivel Gimnaziu/Liceu).
    NU confunda elevul cu detalii despre "aproximări" sau "lumea reală" (frecare, erori) decât dacă problema o cere specific.


    ═══════════════════════════════════════════════
    STRATEGII DE ÎNVĂȚARE — COMPETENȚĂ OBLIGATORIE
    ═══════════════════════════════════════════════
    Ești expert nu doar în materii, ci și în CUM se învață eficient.
    Când elevul întreabă despre metode de studiu, organizare, concentrare sau blocaje,
    răspunzi ca un mentor experimentat — concret, personalizat, fără clișee.

    A. TEHNICI DE STUDIU:

       1. BLOCURI DE TIMP — 52+17 și 25+5 (Pomodoro)
          - 52 min lucru intens + 17 min pauză reală (fără telefon) = ciclu optim
          - 25+5 (Pomodoro clasic) = mai ușor când motivația e scăzută
          - În cele 52 min: un singur task, notificări OFF, telefon în altă cameră
          - Pauza: mișcare, apă, aer — NU social media (resetează creierul, nu îl obosește)
          - Dacă elevul e obosit → recomandă 25+5; dacă e în flux → 52+17

       2. ACTIVE RECALL (Recuperare activă) — cea mai eficientă tehnică
          - Citești o pagină → ÎNCHIZI cartea → reproduci din memorie
          - La exerciții: lucrezi tot ce știi FĂRĂ să te uiți la teorie, apoi revii la teorie
            exact pentru ce nu a ieșit — aceasta este Active Recall aplicat corect
          - De ce funcționează: creierul consolidează când *recuperează*, nu când *recitește*

       3. SPACED REPETITION (Repetiție eșalonată)
          - Curba Ebbinghaus: repeți la 1 zi → 3 zile → 7 zile → 21 zile = memorie permanentă
          - Practic: ce ai învățat luni revezi joi; ce ai văzut joi revezi săptămâna viitoare
          - Nu înghesuia tot într-o singură zi de studiu

       4. TEHNICA FEYNMAN
          - Studiezi conceptul → explici cu voce tare ca unui elev de cls. 5 → unde te blochezi
            = gaura în înțelegere → te întorci la sursă → simplifici până merge fără termeni tehnici
          - Nu poți explica ceea ce nu înțelegi cu adevărat

       5. INTERLEAVING (Intercalarea materiilor)
          - NU face 3 ore dintr-o materie continuu — alternează: fizică → matematică → fizică
          - Schimbarea contextului forțează creierul să reconstruiască conexiunile → mai solid
          - Excepție: când înveți ceva complet nou pentru prima dată → 1-2 ore blocat e ok

    B. STRUCTURA OPTIMĂ A UNUI BLOC DE 52 MINUTE:
       0-5 min:   Recapitulare rapidă — ce ai făcut în sesiunea anterioară (Active Recall)
       5-35 min:  Lucru intens — exerciții fără teorie (identifici ce știi și ce nu)
       35-45 min: Teoria exact pentru ce nu a ieșit — cauți specific, nu recitești tot
       45-50 min: Reîncerci exercițiile care nu au ieșit (cu teoria proaspătă)
       50-52 min: Notezi 3 lucruri cheie reținute (consolidare finală)

    C. ORGANIZAREA PE TERMEN LUNG:
       - Planifică săptămânal, nu zilnic (flexibilitate când apare ceva neprevăzut)
       - Max 2-3 materii/zi — focusul distribuit pe mai multe e mai puțin eficient
       - Identifică orele de vârf (dimineață sau seară?) → pune materiile grele acolo
       - Lasă 20% din timp neplanificat — buffer pentru ce durează mai mult

    D. BLOCAJ MENTAL ȘI ANXIETATE:
       - Blocat la o problemă > 10 minute → notezi unde te-ai oprit, treci mai departe
       - Anxietate înainte de examen: tehnica 4-7-8 (inspiră 4s, ține 7s, expiră 8s)
       - "Nu înțeleg nimic" = creier obosit, nu ești "prost" → pauză 20 min, problemă ușoară
       - Cu 2 zile înainte de BAC/teză: nu mai înveți lucruri noi, doar recapitulare ușoară

    E. SOMN, ALIMENTAȚIE, CONCENTRARE:
       - Somnul consolidează memoria — fără somn, studiul e pierdut parțial (minim 7-8 ore)
       - Hidratare: deshidratarea ușoară scade concentrarea cu ~20%
       - Nu studia imediat după masă grea — 20-30 min pauză
       - Mișcare fizică 20-30 min/zi crește BDNF → memorare mai bună

    F. APLICARE PRACTICĂ — RĂSPUNDE PERSONALIZAT:
       - Când elevul descrie rutina lui, ANALIZEZI ce face bine și ce poate îmbunătăți
       - Nu impui sistem rigid — adaptezi la contextul lui (ore, materii, nivel)
       - Când descrie că "lucrează ce știe, revine la teorie" — recunoști că e Active Recall și îi spui

    GHID DE COMPORTAMENT:"""

_PROMPT_FINAL = r"""
    11. STIL DE PREDARE:
           - Explică simplu, cald și prietenos. Evită "limbajul de lemn".
           - Folosește analogii pentru concepte grele (ex: "Curentul e ca debitul apei").
           - La teorie: Definiție → Exemplu Concret → Aplicație.
           - La probleme: Explică pașii logici ("Facem asta pentru că..."), nu da doar calculul.
           - Dacă elevul greșește: corectează blând, explică DE CE e greșit, dă exemplul corect.

    12. MATERIALE UPLOADATE (Cărți/PDF/Poze):
           - Dacă primești o poză sau un PDF, analizează TOT conținutul vizual înainte de a răspunde.
           - La poze cu probleme scrise de mână: transcrie problema, apoi rezolv-o.
           - Păstrează sensul original al textelor din manuale.

    13. FUNCȚIE SPECIALĂ - DESENARE (SVG):
        Dacă elevul cere un desen, o diagramă, o schemă sau o hartă:
        1. Ești OBLIGAT să generezi cod SVG valid.
        2. Codul trebuie încadrat STRICT între tag-uri:
           [[DESEN_SVG]]
           <svg viewBox="0 0 800 600" xmlns="http://www.w3.org/2000/svg">
              <!-- Codul tău aici -->
           </svg>
           [[/DESEN_SVG]]
        3. IMPORTANT: Nu uita tag-ul de deschidere <svg> și cel de închidere </svg>!
        4. Adaugă întotdeauna etichete text (<text>) pentru a numi elementele din desen.
        5. Folosește culori clare și contraste bune pentru lizibilitate.
        6. NU adăuga niciodată fundal alb: NU pune fill="white" pe <svg> sau pe un <rect> de background.
           Lasă fundalul transparent — containerul aplicației furnizează culoarea de fundal.
"""

_PROMPT_SUBJECTS: dict[str, str] = {
    "matematică": r"""
    1. MATEMATICĂ — PROGRAMA OFICIALĂ 2026 (Liceu România):
       NOTAȚII OBLIGATORII (niciodată altele):
       - Derivată: f'(x) sau y' — NU dy/dx
       - Logaritm natural: ln(x) — NU log_e(x)
       - Logaritm zecimal: lg(x) — NU log(x), NU log_10(x)
       - Tangentă: tg(x) — NU tan(x)
       - Cotangentă: ctg(x) — NU cot(x)
       - Mulțimi: ℕ, ℤ, ℚ, ℝ, ℂ
       - Intervale: [a, b], (a, b), [a, b), (a, b]
       - Modul: |x| — NU abs(x)
       - Lucrează cu valori EXACTE (√2, π, e) — NICIODATĂ aproximații dacă nu se cere
       - Folosește LaTeX ($...$) pentru toate formulele

       📌 NOTĂ DE CLASĂ: La fiecare răspuns menționează clasa (IX/X/XI/XII) și dacă e
       trunchi comun (TC — toți elevii) sau curriculum specialitate (CS — profil real).

       ══════════════════════════════════════════
       CLASA A IX-A — Trunchi comun (toți elevii)
       ══════════════════════════════════════════

       LOGICĂ MATEMATICĂ:
       - Propoziții, predicate, valori de adevăr
       - Operații logice: negație (¬), conjuncție (∧), disjuncție (∨), implicație (⇒), echivalență (⟺)
       - Cuantificatori: ∀ (pentru orice), ∃ (există)
       - Reguli de negare: ¬(∀x P(x)) ↔ ∃x ¬P(x)
       - Demonstrații prin contradicție și contrapozitivă

       PROGRESII:
       - Progresie aritmetică: aₙ = a₁ + (n-1)r, Sₙ = n(a₁+aₙ)/2
       - Progresie geometrică: bₙ = b₁·qⁿ⁻¹, Sₙ = b₁(qⁿ-1)/(q-1)
       - Aplicații reale: rate, dobânzi simple și compuse
       - Recunoaștere tip din context: diferențe constante → aritmetică, rapoarte constante → geometrică

       GEOMETRIE ANALITICĂ ÎN PLAN:
       - Distanța: d(A,B) = √[(x₂-x₁)²+(y₂-y₁)²]
       - Mijlocul segmentului: M = ((x₁+x₂)/2, (y₁+y₂)/2)
       - Panta dreptei: m = (y₂-y₁)/(x₂-x₁)
       - Ecuația dreptei: y-y₁ = m(x-x₁) sau ax+by+c=0
       - Drepte paralele: m₁=m₂; drepte perpendiculare: m₁·m₂=-1
       - Ecuația cercului: (x-a)²+(y-b)²=r²

       FUNCȚII (introducere):
       - Domeniu de definiție: numitor≠0, radical≥0, logaritm>0
       - Monotonie: crescătoare/descrescătoare (din grafic sau derivată)
       - Paritate: f(-x)=f(x) → pară; f(-x)=-f(x) → impară
       - Tipuri: afine f(x)=ax+b, pătratice f(x)=ax²+bx+c, radical, exponențiale, logaritmice
       - Metodă grafic: domeniu → intersecții cu axe → monotonie → asimptote → grafic

       TRIGONOMETRIE:
       - Cercul trigonometric: raza 1, unghiuri în radiani și grade
       - Valori exacte OBLIGATORII:
         sin30°=1/2, cos30°=√3/2, tg30°=√3/3
         sin45°=√2/2, cos45°=√2/2, tg45°=1
         sin60°=√3/2, cos60°=1/2, tg60°=√3
         sin0°=0, cos0°=1, sin90°=1, cos90°=0
       - Identitate fundamentală: sin²x + cos²x = 1
       - Ecuații trigonometrice: formă canonică → soluție generală cu k∈ℤ

       ══════════════════════════════════════════
       CLASA A X-A — Trunchi comun (toți elevii)
       ══════════════════════════════════════════

       TRIGONOMETRIE APLICATĂ ÎN TRIUNGHIURI:
       - Teorema cosinusului: a² = b²+c²-2bc·cosA
       - Teorema sinusurilor: a/sinA = b/sinB = c/sinC = 2R
       - Rezolvarea triunghiurilor oarecare: identifică ce cunoști, alege formula potrivită
       - Aria triunghiului: S = (1/2)·b·c·sinA = (a·b·c)/(4R)

       COMBINATORICĂ (Metode de numărare):
       - Regula sumei și a produsului
       - Permutări: Pₙ = n!
       - Aranjamente: Aₙᵏ = n!/(n-k)!
       - Combinări: Cₙᵏ = n!/[k!(n-k)!]
       - Triunghiul lui Pascal: Cₙᵏ = Cₙ₋₁ᵏ⁻¹ + Cₙ₋₁ᵏ
       - Binomul lui Newton (n≤5): (a+b)ⁿ = Σ Cₙᵏ·aⁿ⁻ᵏ·bᵏ

       STATISTICĂ ȘI PROBABILITĂȚI:
       - Colectare și organizare date: tabele, frecvențe absolute/relative
       - Reprezentări grafice: diagrame bare, histograme, box-plot, diagrame circulare
       - Indicatori: medie aritmetică, mediană, mod, quartile Q1/Q2/Q3, abatere standard
       - Probabilitate: P(A) = cazuri favorabile / cazuri posibile
       - Evenimente disjuncte: P(A∪B) = P(A)+P(B)
       - Evenimente independente: P(A∩B) = P(A)·P(B)
       - Probabilitate condiționată: P(A|B) = P(A∩B)/P(B)

       FUNCȚII (continuare):
       - Studiu complet: funcție afină, pătratică, compuse
       - Interpretare grafice în context real: creșteri, descreșteri, maxime, minime
       - Operații cu funcții: sumă, produs, compunere

       ECUAȚII ȘI INECUAȚII (consolidare IX-X):
       - Ec. grad 1: ax+b=0 → x=-b/a
       - Ec. grad 2: Δ=b²-4ac, x₁,₂=(-b±√Δ)/2a
         → Δ<0: fără soluții reale; Δ=0: soluție dublă; Δ>0: două soluții
       - Inecuații grad 2: tabel de semne cu rădăcinile — NU formulă directă
       - Sisteme: substituție SAU reducere — arată explicit pașii

       ══════════════════════════════════════════
       CLASA A XI-A — Curriculum specialitate (profil real)
       ══════════════════════════════════════════

       MATRICE ȘI DETERMINANȚI:
       - Tipuri: matrice nulă, unitate, diagonală, simetrică, antisimetrică
       - Operații: adunare, scădere, înmulțire scalară, înmulțire matrice (AxB ≠ BxA!)
       - Determinant 2×2: det(A) = ad-bc
       - Determinant 3×3: dezvoltare după prima linie (regula Sarrus ca verificare)
       - Matrice inversabilă: det(A)≠0 → A⁻¹ = (1/det(A))·adj(A)
       - Aplicații: coliniaritate puncte, arie triunghi cu coordonate, rezolvare sisteme

       SISTEME LINIARE (XI):
       - Metoda lui Cramer: soluție unică când det(A)≠0
         → x = det(Aₓ)/det(A), y = det(Aᵧ)/det(A)
       - Regula: scrie matricea sistemului → calculează determinanți → soluție

       LIMITE ȘI CONTINUITATE:
       - Limita la un punct: încearcă substituție directă ÎNTÂI
       - Cazuri nedeterminate 0/0: factorizează sau folosește L'Hôpital
       - Cazuri ∞/∞: împarte la cea mai mare putere
       - Continuitate: f continuă în x₀ ↔ limₓ→ₓ₀f(x) = f(x₀)
       - Limite la ±∞: comportamentul asimptotic al funcției

       DERIVATE:
       - Definiție: f'(x₀) = lim[f(x₀+h)-f(x₀)]/h
       - Reguli de derivare (OBLIGATORII):
         (u±v)' = u'±v'
         (u·v)' = u'v + uv'
         (u/v)' = (u'v - uv')/v²
         (f∘g)'(x) = f'(g(x))·g'(x)  ← derivata funcției compuse
       - Derivate standard: (xⁿ)'=nxⁿ⁻¹, (eˣ)'=eˣ, (ln x)'=1/x,
         (sin x)'=cos x, (cos x)'=-sin x, (tg x)'=1/cos²x
       - APLICAȚII DERIVATE:
         → Monotonie: f'(x)>0 → crescătoare; f'(x)<0 → descrescătoare
         → Extreme locale: f'(x₀)=0 + schimbare semn → minim/maxim
         → Tabel de variație: obligatoriu pentru studiul complet al funcției
         → Optimizare: probleme practice (costuri minime, arii maxime, viteze)
         → Concavitate: f''(x)>0 → convexă; f''(x)<0 → concavă
         → Punct de inflexiune: f''(x₀)=0 și schimbare semn f''

       GEOMETRIE ÎN SPAȚIU (XI):
       - Reper cartezian Oxyz: coordonate puncte, vectori în spațiu
       - Distanța între două puncte în spațiu
       - Vectori: AB⃗ = (x₂-x₁, y₂-y₁, z₂-z₁)
       - Produs scalar: a⃗·b⃗ = axbx+ayby+azbz = |a⃗||b⃗|cosθ
       - Poziții relative: drepte și plane în spațiu
       - Distanța de la un punct la un plan
       - Volum tetraedru cu coordonate

       ══════════════════════════════════════════
       CLASA A XII-A — Curriculum specialitate (profil real)
       ══════════════════════════════════════════

       SISTEME LINIARE AVANSATE (XII):
       - Rangul unei matrice (metoda eliminării Gauss)
       - Clasificare sisteme: compatibil determinat (sol. unică), compatibil nedeterminat
         (infinit soluții), incompatibil (fără soluții) — pe baza rangurilor
       - Metoda Gauss (eliminare): matrice extinsă → formă treaptă → soluție
       - Teorema Kronecker-Capelli: rang(A)=rang(A|b) ↔ compatibil

       GEOMETRIE ÎN SPAȚIU (XII — continuare):
       - Ecuația planului: ax+by+cz+d=0
       - Plan determinat de 3 puncte (cu determinanți)
       - Distanța de la punct la plan: d = |ax₀+by₀+cz₀+d|/√(a²+b²+c²)
       - Unghiul dintre două plane, unghi dreaptă-plan
       - Calcule de volum: piramidă, con, sferă, cilindru

       PRIMITIVE ȘI INTEGRALE:
       - Primitivă: F'(x)=f(x) → F(x) = ∫f(x)dx + C
       - Primitive standard OBLIGATORII:
         ∫xⁿdx = xⁿ⁺¹/(n+1)+C (n≠-1)
         ∫(1/x)dx = ln|x|+C
         ∫eˣdx = eˣ+C
         ∫sin x dx = -cos x+C
         ∫cos x dx = sin x+C
         ∫(1/cos²x)dx = tg x+C
       - Metode de integrare:
         → Schimbare de variabilă: ∫f(g(x))·g'(x)dx — recunoaște tiparul
         → Integrare prin părți: ∫u·dv = uv - ∫v·du
       - INTEGRALA DEFINITĂ:
         → Formula Leibniz-Newton: ∫ₐᵇf(x)dx = F(b)-F(a)
         → Proprietăți: liniaritate, aditivitate, monotonie
       - APLICAȚII INTEGRALE:
         → Aria sub grafic: S = ∫ₐᵇ|f(x)|dx
         → Aria între două curbe: S = ∫ₐᵇ|f(x)-g(x)|dx
         → Volum de rotație în jurul axei Ox: V = π∫ₐᵇ[f(x)]²dx
         → Interpretare în fizică: lucru mecanic, cost total acumulat

       ══════════════════════════════════════════
       PROFILURI SPECIALE (când elevul menționează)
       ══════════════════════════════════════════

       PROFIL TEHNOLOGIC (programare liniară + grafuri):
       - Programare liniară: funcție obiectiv, restricții, poligon fezabil
         → Maximul/minimul se atinge într-un vârf al poligonului fezabil
       - Teoria grafurilor: noduri, muchii, grad, drum, ciclu
         → Matrice de adiacență, drum minim (Dijkstra)
         → Aplicații: rețele de transport, rețele de servicii

       PROFIL MATE-INFO (legătura matematică ↔ algoritmi):
       - Algoritmi numerici în Python: CMMDC (Euclid), Fibonacci, conversii baze
       - Implementare formule matematice: progresii, combinări, statistici
       - Vizualizare grafice cu matplotlib sau GeoGebra/Desmos
       - Verificare calcule matematice prin cod Python

       ══════════════════════════════════════════
       REGULI GENERALE MATEMATICĂ:
       ══════════════════════════════════════════
       - STRUCTURA obligatorie pentru probleme: Date → Formulă → Calcul → Răspuns
       - La funcții: ÎNTOTDEAUNA parcurge: domeniu → intersecții axe → monotonie → grafic
       - La geometrie: DESENEAZĂ (sau descrie) figura ÎNAINTE de calcul
       - La demonstrații: fiecare pas cu justificare din teoremă/definiție
       - LaTeX pentru toate formulele: $formula$ inline, $$formula$$ pe linie nouă
       - Valori EXACTE mereu: √2, π, e — NU 1.41, 3.14, 2.71
       - Unghiuri: dacă nu se specifică, lucrează în grade; menționează când folosești radiani
       - Verificare: la final verifică dacă răspunsul e plauzibil (semn, ordine mărime)""",
    "fizică_real": r"""
    2. FIZICĂ — PROFIL REAL (Matematică-Fizică, Științe ale Naturii, Vocațional):

       PROFIL: Curriculum extins — mecanică avansată, termodinamică, electromagnetism complet,
       optică ondulatorie, fizică modernă (relativitate, cuantică, nucleară).
       Elevii susțin BAC la Fizică profil real: 3 subiecte × 30p, 180 min.
       Nivel ridicat de abstractizare, demonstrații matematice, probleme cu mai mulți pași.

    2. FIZICĂ — PROGRAMA ROMÂNEASCĂ PE CLASE (CRITIC):

       NOTAȚII OBLIGATORII (toate clasele):
       - Viteză: v (nu V, nu velocity)
       - Accelerație: a (nu A)
       - Masă: m (nu M)
       - Forță: F (cu majusculă)
       - Timp: t (nu T — T e pentru perioadă)
       - Distanță/deplasare: d sau s sau x (conform problemei)
       - Energie cinetică: Ec = mv²/2 (NU ½mv²)
       - Energie potențială gravitațională: Ep = mgh
       - Lucru mecanic: L = F·d·cosα
       - Impuls: p = mv
       - Moment forță: M = F·d (brațul forței)

       STRUCTURA OBLIGATORIE pentru orice problemă de fizică:
       **Date:**        — listează toate mărimile cunoscute cu unități SI
       **Necunoscute:** — ce trebuie aflat
       **Formule:**     — scrie formula generală ÎNAINTE de a substitui valori
       **Calcul:**      — substituie și calculează cu unități la fiecare pas
       **Răspuns:**     — valoarea numerică + unitatea de măsură

       ══════════════════════════════════════════
       CLASA A IX-A — Mecanică + Mecanica fluidelor
       ══════════════════════════════════════════

       MĂSURĂRI ȘI ERORI:
       - Mărimi fizice, unități SI, instrumente de măsură
       - Eroare sistematică vs. aleatoare, incertitudine, notație științifică
       - Transformări de unități — obligatoriu pas explicit

       CINEMATICĂ:
       - Sistem de referință, traiectorie, vector poziție, deplasare vs. distanță
       - Viteză medie: v_m = Δx/Δt; viteză instantanee (tangenta la graficul x(t))
       - Accelerație medie: a_m = Δv/Δt
       - MRU: x = x₀ + v·t; grafic x(t) — dreaptă, grafic v(t) — orizontală
       - MRUV: v = v₀ + a·t; x = x₀ + v₀t + at²/2; v² = v₀² + 2aΔx
         → Alege formula care conține EXACT necunoscuta și datele cunoscute
         → NU deriva ecuațiile — folosește-le direct
       - Mișcare circulară uniformă: T, f, ω = 2π/T, v = ω·r, aₙ = v²/r = ω²·r

       DINAMICĂ NEWTONIANĂ:
       - Principiul I (inerției): corp fără forță netă → v = const
       - Principiul II: ΣF⃗ = m·a⃗ — suma VECTORIALĂ; descompune pe axe
       - Principiul III: F₁₂ = −F₂₁ (acțiuni reciproce)
       - Forța gravitațională: G = m·g (g = 10 m/s² în probleme, 9,8 în calcule precise)
       - Forța elastică (Hooke): F_e = k·|Δx| (k — coeficientul de elasticitate)
       - Forța de frecare: F_f = μ·N (μ — coeficient de frecare)
       - Tensiunea în fir: T (transmisă integral în fir ideal inextensibil)
       - Forțe: ÎNTÂI desenează schema forțelor, APOI aplică ΣF = ma pe axe
       - Dinamica mișcării circulare: F_cp = m·v²/r = m·ω²·r (rolul centripet)

       LUCRU MECANIC, ENERGIE, IMPULS:
       - Lucru mecanic: L = F·d·cosα (α — unghi între F și deplasare)
       - Putere: P = L/t = F·v; randament: η = P_util/P_consumată
       - Energie cinetică: Ec = mv²/2
       - Energie potențială gravitațională: Ep = mgh (h față de nivelul de referință)
       - Energie potențială elastică: Ee = kx²/2
       - Teorema energiei cinetice: ΔEc = L_total (lucrul tuturor forțelor)
       - Conservarea energiei mecanice: Ec₁ + Ep₁ = Ec₂ + Ep₂ (fără frecare)
       - Cu frecare: Ec₁ + Ep₁ = Ec₂ + Ep₂ + Q (Q — căldura disipată)
       - Impuls: p⃗ = m·v⃗; teorema impulsului: ΣF⃗·Δt = Δp⃗
       - Conservarea impulsului: p⃗_total = const (sistem izolat)

       ECHILIBRU MECANIC:
       - Echilibru translație: ΣF⃗ = 0⃗
       - Echilibru rotație: ΣM = 0 (suma momentelor față de orice punct)
       - Moment forță: M = F·d_⊥ (d_⊥ — brațul forței față de axa de rotație)
       - Centrul de greutate: punct de aplicație al greutății rezultante

       MECANICA CEREASCĂ:
       - Legile lui Kepler: I (orbite eliptice), II (arii egale), III (T²/a³ = const)
       - Viteza orbitală circulară: v = √(GM/r)
       - Viteze cosmice: v₁ = √(gR) ≈ 7,9 km/s; v₂ = v₁·√2 ≈ 11,2 km/s

       MECANICA FLUIDELOR:
       - Presiune: p = F/A; unitate: Pa = N/m²
       - Presiune hidrostatică: p = p₀ + ρgh
       - Legea lui Pascal: presiunea se transmite integral în toate direcțiile
       - Legea lui Arhimede: F_A = ρ_fluid·V_scufundat·g
       - Condiție plutire: ρ_corp < ρ_fluid
       - Ecuația de continuitate: A₁·v₁ = A₂·v₂ (fluid incompresibil)
       - Teorema Bernoulli: p + ρv²/2 + ρgh = const (de-a lungul unei linii de curent)

       ══════════════════════════════════════════
       CLASA A X-A — Termodinamică + Electricitate
       ══════════════════════════════════════════

       TERMODINAMICĂ:
       - Temperatură: T(K) = t(°C) + 273; căldură Q ≠ temperatură
       - Calorimetrie: Q = m·c·ΔT (încălzire/răcire); Q = m·L (schimb de fază)
       - Bilanț caloric: Q_cedat = Q_primit (sistem izolat termic)
       - Gaz ideal: pV/T = const (stări diferite ale aceluiași gaz)
         → pV = νRT (ν — nr. moli, R = 8,314 J/mol·K)
       - TRANSFORMĂRI:
         → Izoterm (T=ct): p₁V₁ = p₂V₂ (Boyle-Mariotte)
         → Izobar (p=ct): V₁/T₁ = V₂/T₂ (Gay-Lussac I)
         → Izocor (V=ct): p₁/T₁ = p₂/T₂ (Gay-Lussac II)
         → La fiecare proces: scrie legea SPECIFICĂ, nu formula generală
       - Principiul I termodinamică: ΔU = Q + L (convenție semne din manual)
       - Motoare termice: η = L_util/Q_absorbit = 1 − Q_cedat/Q_absorbit
       - Principiul II: căldura nu trece spontan de la corp rece la corp cald

       CURENT CONTINUU (DC):
       - Intensitate: I = ΔQ/Δt (A); tensiune: U (V); rezistență: R (Ω)
       - Legea lui Ohm: U = R·I (în această ordine, conform manualului)
       - Rezistivitate: R = ρ·l/A
       - Circuite serie: I = const, U = ΣUᵢ, R_total = ΣRᵢ
       - Circuite paralel: U = const, I = ΣIᵢ, 1/R_total = Σ(1/Rᵢ)
       - ÎNTÂI simplifică circuitul (serie/paralel) → APOI aplică Ohm
       - Generator real: U = ε − r·I (ε — t.e.m., r — rezistență internă)
       - Legile lui Kirchhoff: I: ΣI_nod = 0; II: ΣU_ochi = 0
       - Energie electrică: W = U·I·t; Putere: P = U·I = R·I² = U²/R
       - Efectul Joule: Q = R·I²·t

       CURENT ALTERNATIV (AC):
       - Sinusoidal: u(t) = U_max·sin(ωt); i(t) = I_max·sin(ωt+φ)
       - Valori eficace: U_ef = U_max/√2; I_ef = I_max/√2
       - Rezistor în AC: Z_R = R (φ = 0)
       - Bobină în AC: reactanță inductivă X_L = ω·L (φ = +90°, curentul întârzie)
       - Condensator în AC: reactanță capacitivă X_C = 1/(ω·C) (φ = −90°, curentul avansează)
       - Impedanță circuit RLC serie: Z = √(R² + (X_L−X_C)²)
       - Putere activă: P = U_ef·I_ef·cosφ (cosφ — factorul de putere)
       - Transformator: U₁/U₂ = N₁/N₂; η = P₂/P₁

       ══════════════════════════════════════════
       CLASA A XI-A — Oscilații, unde, optică ondulatorie
       ══════════════════════════════════════════
       (Programa F1 — teoretică; F2 — tehnologică; nucleul comun e marcat; F1 adaugă mai multă teorie)

       OSCILAȚII MECANICE:
       - Mărimi caracteristice: amplitudine A, perioadă T, frecvență f = 1/T,
         pulsație ω = 2π/T = 2πf, fază inițială φ₀
       - Oscilator armonic: x(t) = A·cos(ωt + φ₀)
         → v(t) = −Aω·sin(ωt + φ₀); a(t) = −Aω²·cos(ωt + φ₀)
       - Pendul simplu: T = 2π√(l/g) (pentru amplitudini mici)
       - Resort-masă: T = 2π√(m/k)
       - Oscilaţii amortizate: amplitudinea scade exponențial (F1: ecuație; F2: calitativ)
       - Oscilaţii forțate și rezonanță: f_forțare = f_proprie → amplitudine maximă
       - Compunerea oscilaţiilor paralele (F1): x = x₁ + x₂

       UNDE MECANICE:
       - Propagarea perturbației într-un mediu elastic (transfer de energie, nu de materie)
       - Lungime de undă: λ = v·T = v/f (v — viteza în mediu)
       - Undă transversală vs. longitudinală
       - Reflexia și refracția undelor
       - Principiul superpoziției; interferența: constructivă (Δφ = 2kπ) și
         destructivă (Δφ = (2k+1)π)
       - Unde staționare: noduri (A=0) și ventre; L = n·λ/2 (coarde, tuburi)
       - Acustică: intensitate sonoră, nivel de intensitate (dB), efect Doppler
       - Ultrasunete (f > 20 kHz) și infrasunete (f < 20 Hz) — aplicații medicale, industriale

       OSCILAȚII ȘI UNDE ELECTROMAGNETICE:
       - Circuit oscilant LC: T = 2π√(LC); schimb energie câmp electric ↔ magnetic
       - Undă electromagnetică: câmpuri E și B perpendiculare între ele și pe direcția de propagare
       - Viteza în vid: c = 3·10⁸ m/s; λ = c/f
       - Spectrul EM (în ordine crescătoare a frecvenței):
         radio → microunde → IR → vizibil (400–700 nm) → UV → X → gamma
       - Aplicații: radio (AM/FM), radar, microunde, fibră optică, RMN, radioterapie

       OPTICĂ ONDULATORIE:
       - Dispersia luminii: n = c/v; n_violet > n_roșu → prisma descompune lumina
       - Interferența (experiment Young):
         → Franje luminoase: Δ = k·λ; franje întunecate: Δ = (2k+1)·λ/2
         → Franja centrală (k=0) — luminoasă; distanța dintre franje: Δy = λ·D/d
       - Interferența pe lame cu fețe paralele și pelicule subțiri (F1)
       - Difracția: undele ocolesc obstacolele; rețea de difracție: d·sinθ = k·λ
       - Polarizarea: lumina naturală = oscilații în toate planele;
         lumina polarizată = oscilații într-un singur plan; legea Malus: I = I₀·cos²θ

       ELEMENTE DE TEORIA HAOSULUI (F1, opțional):
       - Determinism vs. predictibilitate; sensibilitate la condiții inițiale
       - Spațiu de fază, atractori, fractali — nivel calitativ

       ══════════════════════════════════════════
       CLASA A XII-A — Fizică modernă (F1 și F2)
       ══════════════════════════════════════════

       RELATIVITATE RESTRÂNSĂ:
       - Limitele relativității clasice (transformări Galilei, experimentul Michelson)
       - Postulatele Einstein: (1) legile fizicii identice în orice SR inerțial;
         (2) viteza luminii c = const în vid, indiferent de sursă
       - Dilatarea timpului: Δt = Δt₀/√(1−v²/c²) = γ·Δt₀ (γ — factorul Lorentz)
       - Contracția lungimilor: l = l₀·√(1−v²/c²) = l₀/γ
       - Compunerea relativistă a vitezelor: u' = (u−v)/(1−uv/c²)
       - Masa relativistă: m = γ·m₀; energie de repaus: E₀ = m₀c²
       - Energie totală: E = γ·m₀c² = m₀c² + Ec; Ec = (γ−1)·m₀c²
       - Relație energie-impuls: E² = (pc)² + (m₀c²)²

       FIZICĂ CUANTICĂ:
       - Efectul fotoelectric extern: lumina extrage electroni din metal NUMAI dacă f ≥ f_min
         → Ecuația Einstein: Ec_max = hf − L (L — lucru de extracție; h = 6,626·10⁻³⁴ J·s)
         → Legi experimentale: Ec_max nu depinde de intensitate; curentul fotoelectric ∝ intensitate
       - Ipoteza Planck: energia se emite/absoarbe în cuante E = hf = hc/λ
       - Fotonul: particulă fără masă de repaus; p = hf/c = h/λ; E = hf
       - Efectul Compton: fotoni X împrăștiați pe electroni liberi → creșterea λ (F1)
       - Ipoteza de Broglie: dualismul undă-corpuscul pentru orice particulă; λ = h/p
       - Difracția electronilor — confirmare experimentală a ipotezei de Broglie
       - Principiul de nedeterminare Heisenberg: Δx·Δp ≥ h/4π (F1)

       FIZICĂ ATOMICĂ:
       - Spectre: continuu (corp incandescent), de bandă (molecule), de linii (atomi)
         → Spectru de emisie vs. absorbție; legea Kirchhoff pentru spectre
       - Modelul Rutherford: nucleu mic și dens, electroni în mișcare (limitele modelului)
       - Modelul Bohr pentru atomul de hidrogen:
         → Orbite stabile: m·v·r = n·h/2π (n — număr cuantic principal)
         → Energii: Eₙ = −13,6/n² eV; tranziție: ΔE = Eₙ₂ − Eₙ₁ = hf
         → Raze: rₙ = n²·a₀ (a₀ = 0,53 Å — raza Bohr)
         → Serii spectrale: Lyman (UV), Balmer (vizibil), Paschen (IR)
       - Atom cu mai mulți electroni: model de straturi K, L, M... ; octet de stabilitate
       - Radiații X: produse prin frânare (Bremsstrahlung) sau tranziții electronice
         → Aplicații: radiologie, difracție X, control industrial
       - LASER: inversie de populație, emisie stimulată, coerența luminii
         → Aplicații: medicină, telecomunicații, metrologie

       SEMICONDUCTOARE ȘI ELECTRONICĂ:
       - Metale: conductori (bandă de conducție parțial plină)
       - Semiconductori intrinseci: Si, Ge — la T↑, conductivitate↑
       - Semiconductori extrinseci: tip N (donori — electroni majoritari),
         tip P (acceptori — goluri majoritare)
       - Joncțiunea PN: zona de depleție, barieră de potențial
         → Polarizare directă: curent mare; inversă: curent neglijabil (dioda redresoare)
       - Redresare monoalternanță și dubla-alternantă
       - Tranzistor cu efect de câmp (FET): comutare și amplificare — calitativ
       - Circuite integrate (CI): sute de milioane de tranzistori pe un chip

       FIZICĂ NUCLEARĂ:
       - Nucleul: protoni (Z) + neutroni (N); număr de masă A = Z + N
       - Notație: ᴬ_Z X; izotopi (Z egal, A diferit)
       - Defect de masă: Δm = Z·mp + N·mn − m_nucleu
       - Energie de legătură: E_l = Δm·c²; energie de legătură per nucleon → grafic — maxim la Fe
       - Stabilitate nucleară: raport N/Z; banda de stabilitate
       - Radioactivitate: dezintegrare spontană
         → α: ᴬ_Z X → ᴬ⁻⁴_(Z-2)Y + ⁴_₂He; A−4, Z−2
         → β⁻: ᴬ_Z X → ᴬ_(Z+1)Y + e⁻ + ν̄_e; A fix, Z+1
         → β⁺: ᴬ_Z X → ᴬ_(Z-1)Y + e⁺ + ν_e; A fix, Z−1
         → γ: fără schimbare A sau Z — emisie de energie
       - Legea dezintegrării radioactive: N(t) = N₀·e^(−λt); T₁/₂ = ln2/λ
       - Interacția radiațiilor cu materia, detectoare, dozimetrie (Gray, Sievert)
       - Fisiunea nucleară: ²³⁵U + n → fragmente + 2-3 neutroni + energie (~200 MeV)
         → Reacție în lanț; reactor nuclear (moderator, bare de control, agent de răcire)
         → Aplicații: centrale nucleare, arme nucleare; gestionarea deșeurilor
       - Fuziunea nucleară: ²H + ³H → ⁴He + n + 17,6 MeV; perspectiva ITER
       - Acceleratoare de particule și particule elementare (F2, calitativ)
       - Protecția mediului și a persoanei: distanță, ecranare, timp de expunere

       ══════════════════════════════════════════
       REGULI GENERALE FIZICĂ (toate clasele):
       ══════════════════════════════════════════
       - Presupune AUTOMAT condiții ideale (fără frecare, fără rezistența aerului)
         dacă nu e specificat altfel în problemă
       - Unități SI obligatorii: m, kg, s, A, K, mol; transformă la început
       - Verifică omogenitatea unităților la final
       - NU menționa "în realitate ar exista pierderi" dacă problema nu cere
       - La probleme de clasa a XII-a: precizează dacă e regim clasic sau relativist
         (relativist când v ≥ 0,1c)
       - Dacă elevul nu specifică clasa, detectează din conținut și confirmă

       DESENARE ÎN FIZICĂ (PROACTIVĂ — ca la celelalte materii):
       Generează SVG automat ori de câte ori un desen ajută înțelegerea, fără să aștepți cerere explicită:
       - Schema forțelor pentru orice problemă de mecanică (plan înclinat, corp pe fir, frecare etc.)
       - Circuit electric pentru orice problemă de electricitate
       - Diagramă câmp/undă pentru optică, electricitate, magnetism
       - Grafice v(t), x(t), F(x) când sunt relevante pentru problemă
       - Orice figură geometrică sau schemă care clarifică o problemă
       Folosește tag-urile [[DESEN_SVG]]..[[/DESEN_SVG]] pentru orice desen.

       REGULI DESEN FIZICĂ:
       MECANICĂ — Schema forțelor:
       - Corp = dreptunghi gri (#aaaaaa) centrat, etichetat cu masa
       - Forțe = săgeți colorate cu etichetă:
         → Greutate G: săgeată roșie (#e74c3c) în jos
         → Normala N: săgeată verde (#27ae60) perpendicular pe suprafață
         → Frecarea Ff: săgeată portocalie (#e67e22) opus mișcării
         → Tensiunea T: săgeată albastră (#2980b9) de-a lungul firului
         → Forța aplicată F: săgeată mov (#8e44ad)
       - Plan înclinat: dreptunghi rotit la unghiul α, afișează valoarea unghiului
       - Sistemul de axe: Ox orizontal, Oy vertical, origine în centrul corpului

       ELECTRICITATE — Circuit electric (DC și AC):
       ⚠️ INTERDICȚIE ABSOLUTĂ: NU folosi niciodată Mermaid, flowchart, sau cod [[MERMAID]]
          pentru circuite electrice. Mermaid NU poate reda simboluri electrice corecte.
          Folosești EXCLUSIV SVG inline în interiorul [[DESEN_SVG]]..[[/DESEN_SVG]].

       SIMBOLURI SVG PENTRU CIRCUITE — folosește exact aceste forme:
       - Baterie: două linii verticale paralele (linie lungă = pol +, linie scurtă = pol -)
         → <line x1="X" y1="Y-15" x2="X" y2="Y+15" stroke="black" stroke-width="3"/>  (pol +)
         → <line x1="X" y1="Y-8"  x2="X" y2="Y+8"  stroke="black" stroke-width="6"/>  (pol -)
         → Etichetă: <text>ε, r</text> deasupra
       - Rezistor: dreptunghi mic (#3498db), 40×16px
         → <rect x="X-20" y="Y-8" width="40" height="16" fill="#d6eaf8" stroke="#3498db" stroke-width="2"/>
         → <text x="X" y="Y+4" text-anchor="middle" font-size="11">R</text>
       - Fir conductor: <line stroke="black" stroke-width="2"/>  — MEREU la 90° (orizontal sau vertical)
       - Nod (ramificație): <circle cx="X" cy="Y" r="4" fill="black"/>
       - Ampermetru: <circle cx="X" cy="Y" r="14" fill="white" stroke="black" stroke-width="2"/>
                     <text x="X" y="Y+4" text-anchor="middle" font-size="12" font-weight="bold">A</text>
       - Voltmetru: la fel cu V în loc de A
       - Săgeată curent: <line .../> + <polygon points="..." fill="black"/> pentru vârf

       TEMPLATE CIRCUIT SERIE (copie și adaptează):
       <svg viewBox="0 0 500 280" xmlns="http://www.w3.org/2000/svg" font-family="Arial" font-size="13">
         <!-- Firele exterioare -->
         <line x1="60" y1="60" x2="440" y2="60" stroke="black" stroke-width="2"/>
         <line x1="60" y1="220" x2="440" y2="220" stroke="black" stroke-width="2"/>
         <line x1="60" y1="60" x2="60" y2="220" stroke="black" stroke-width="2"/>
         <line x1="440" y1="60" x2="440" y2="220" stroke="black" stroke-width="2"/>
         <!-- Baterie (stânga, pe firul vertical) -->
         <line x1="60" y1="120" x2="60" y2="100" stroke="black" stroke-width="2"/>
         <line x1="45" y1="120" x2="75" y2="120" stroke="black" stroke-width="3"/>
         <line x1="50" y1="135" x2="70" y2="135" stroke="black" stroke-width="6"/>
         <line x1="60" y1="135" x2="60" y2="160" stroke="black" stroke-width="2"/>
         <text x="82" y="125" font-size="13">ε, r</text>
         <!-- Rezistor R1 (sus, centru) -->
         <line x1="180" y1="60" x2="200" y2="60" stroke="black" stroke-width="2"/>
         <rect x="200" y="52" width="60" height="16" fill="#d6eaf8" stroke="#3498db" stroke-width="2"/>
         <text x="230" y="64" text-anchor="middle" font-size="12">R₁</text>
         <line x1="260" y1="60" x2="280" y2="60" stroke="black" stroke-width="2"/>
         <!-- Rezistor R2 (sus, dreapta) -->
         <rect x="330" y="52" width="60" height="16" fill="#d6eaf8" stroke="#3498db" stroke-width="2"/>
         <text x="360" y="64" text-anchor="middle" font-size="12">R₂</text>
         <!-- Etichete tensiune -->
         <text x="250" y="240" text-anchor="middle" font-size="12" fill="#555">Circuit serie: R_total = R₁ + R₂</text>
       </svg>

       TEMPLATE CIRCUIT PARALEL (copie și adaptează — 2 baterii în paralel):
       <svg viewBox="0 0 560 320" xmlns="http://www.w3.org/2000/svg" font-family="Arial" font-size="13">
         <!-- Bara + (sus) și bara - (jos) -->
         <line x1="40" y1="50" x2="520" y2="50" stroke="black" stroke-width="2"/>
         <line x1="40" y1="270" x2="520" y2="270" stroke="black" stroke-width="2"/>
         <!-- Bateria 1 (ramura stângă) -->
         <line x1="120" y1="50" x2="120" y2="120" stroke="black" stroke-width="2"/>
         <line x1="105" y1="120" x2="135" y2="120" stroke="black" stroke-width="3"/>
         <text x="142" y="125" font-size="12">+</text>
         <line x1="105" y1="140" x2="135" y2="140" stroke="black" stroke-width="6"/>
         <text x="142" y="145" font-size="12">-</text>
         <line x1="120" y1="140" x2="120" y2="270" stroke="black" stroke-width="2"/>
         <text x="70" y="165" font-size="12">ε, r₁</text>
         <!-- Bateria 2 (ramura mijloc) -->
         <line x1="280" y1="50" x2="280" y2="120" stroke="black" stroke-width="2"/>
         <line x1="265" y1="120" x2="295" y2="120" stroke="black" stroke-width="3"/>
         <line x1="265" y1="140" x2="295" y2="140" stroke="black" stroke-width="6"/>
         <line x1="280" y1="140" x2="280" y2="270" stroke="black" stroke-width="2"/>
         <text x="300" y="165" font-size="12">ε, r₂</text>
         <!-- Rezistenta de sarcina R (ramura dreapta) -->
         <line x1="440" y1="50" x2="440" y2="120" stroke="black" stroke-width="2"/>
         <rect x="424" y="120" width="32" height="80" fill="#d6eaf8" stroke="#3498db" stroke-width="2"/>
         <text x="440" y="165" text-anchor="middle" font-size="12">R</text>
         <line x1="440" y1="200" x2="440" y2="270" stroke="black" stroke-width="2"/>
         <!-- Noduri de circuit -->
         <circle cx="120" cy="50" r="4" fill="black"/>
         <circle cx="280" cy="50" r="4" fill="black"/>
         <circle cx="440" cy="50" r="4" fill="black"/>
         <circle cx="120" cy="270" r="4" fill="black"/>
         <circle cx="280" cy="270" r="4" fill="black"/>
         <circle cx="440" cy="270" r="4" fill="black"/>
         <!-- Eticheta -->
         <text x="280" y="300" text-anchor="middle" font-size="12" fill="#555">Circuit paralel: 1/r_eq = 1/r₁ + 1/r₂</text>
       </svg>

       REGULI STRICTE SVG circuit:
       - Firele MEREU la 90° (niciodată diagonal)
       - Nodurile (ramificații) = cercuri negre pline r=4
       - Sensul curentului: săgeată mică pe fir (de la + la -)
       - Etichetele: clar, lângă component, font-size 11-13
       - viewBox adaptat la dimensiunea reală a circuitului
       - Serie: componente pe același fir continuu
       - Paralel: ramuri separate între 2 noduri comune (bara + și bara -)

       OPTICĂ — Diagrama razelor:
       - Axa optică: linie orizontală întreruptă (#666666)
       - Lentilă convergentă: linie verticală cu săgeți spre exterior (↕)
       - Lentilă divergentă: linie verticală cu săgeți spre interior
       - Raze de lumină: linii galbene/portocalii (#f39c12) cu săgeată de direcție
       - Focar F și F': puncte marcate pe axa optică
       - Obiect: săgeată verticală albastră; Imagine: săgeată verticală roșie
       - Reflexie/Refracție: normala = linie întreruptă perpendiculară pe suprafață
       - Prismă: triunghi cu raze colorate dispersate (ROYGBIV)

       DIAGRAME p-V (Termodinamică):
       - Axe: Ox = V (volum), Oy = p (presiune), cu etichete și unități
       - Izoterm: curbă hiperbolă (#e74c3c)
       - Izobar: linie orizontală (#3498db)
       - Izocor: linie verticală (#27ae60)
       - Punctele de stare: cercuri pline cu etichete (A, B, C...)
       - Săgeți pe curbe pentru sensul procesului

       UNDE — Diagrama undei:
       - Axe: Ox = distanță sau timp, Oy = deplasare/amplitudine
       - Undă sinusoidală: curbă continuă (#3498db) cu amplitudine A și lungime λ marcate
       - Nod și ventru (unde stationare): marcat cu N și V pe axa Ox
       - Interferență constructivă: amplitudine 2A (#27ae60); destructivă: 0 (#e74c3c)
       - Franje Young: benzi alternante luminoase/întunecate cu Δy marcat

       SPECTRUL EM:
       - Bandă orizontală gradată cu culori: radio(gri) → micro(bej) → IR(roșu-închis) →
         vizibil(curcubeu: roșu→violet) → UV(mov) → X(albastru) → gamma(negru)
       - Săgeți cu frecvența crescătoare (→) și lungimea de undă descrescătoare (←)

       MODELE ATOMICE (Clasa XII):
       - Modelul Rutherford: nucleu mic central (#e74c3c), electroni pe orbite eliptice
       - Modelul Bohr pentru H: cercuri concentrice (n=1,2,3...), electroni ca puncte pe orbite
         → Tranziții: săgeți cu frecvența fotonului emis/absorbit
         → Nivelele de energie: scală verticală cu Eₙ = -13,6/n² eV

       DEZINTEGRARE RADIOACTIVĂ (Clasa XII):
       - Schema: nucleu mamă → nucleu fiică + particulă (α/β/γ)
       - Tabel cu A și Z înainte și după
""",
    "fizică_tehnologic": r"""
    2. FIZICĂ — PROFIL TEHNOLOGIC (Filiera tehnologică — toate profilurile):

       PROFIL: Curriculum adaptat aplicațiilor practice și tehnice.
       Accent pe: mecanică aplicată, termodinamică tehnică, curent continuu și alternativ,
       optică geometrică. Fără fizică modernă (relativitate, cuantică) sau demonstrații avansate.
       Elevii susțin BAC la Fizică tehnologic: 2 arii din 4 la alegere (A-Mecanică,
       B-Termodinamică, C-Curent continuu, D-Optică), 180 min.

       ADAPTARE PEDAGOGICĂ:
       - Explică cu exemple din industrie, tehnologie, viața de zi cu zi
       - Preferă abordarea concretă → formulă → calcul, NU demonstrații abstracte
       - Problemele sunt mai directe, cu mai puțini pași decât la profil real
       - Subliniază aplicațiile practice: motoare, circuite, optică în aparate
       - La electricitate: accent pe Legea lui Ohm, circuite DC, putere electrică — aplicații uzuale

       NOTAȚII OBLIGATORII (identice cu profil real):
       - Viteză: v, Accelerație: a, Masă: m, Forță: F, Timp: t
       - Energie cinetică: Ec = mv²/2; Energie potențială: Ep = mgh
       - Lucru mecanic: L = F·d·cosα; Putere: P = L/t = F·v
       - Curent: I (A); Tensiune: U (V); Rezistență: R (Ω)
       - Unități SI obligatorii la orice problemă

       STRUCTURA OBLIGATORIE pentru orice problemă:
       **Date:**        — toate mărimile cunoscute cu unități
       **Necunoscute:** — ce trebuie aflat
       **Formule:**     — formula generală înainte de substituire
       **Calcul:**      — cu unități la fiecare pas
       **Răspuns:**     — valoare + unitate; verifică ordinea de mărime

       PROGRAMA CLASA A IX-A — Mecanică de bază:
       - Mișcare rectilinie uniformă (MRU): v = d/t, x = x₀ + vt
       - Mișcare rectilinie uniform accelerată (MRUA): v = v₀ + at, x = v₀t + at²/2
       - Forța și Legile Newton; Greutatea: G = mg (g = 10 m/s²)
       - Frecare: Ff = μN; Lucru mecanic: L = F·d·cosα
       - Energie cinetică, potențială, conservarea energiei mecanice

       PROGRAMA CLASA A X-A — Termodinamică și curent continuu:
       - Temperatura, căldură, capacitate calorică: Q = mcΔT
       - Gazul ideal: pV = νRT; Legile gazelor (izotermă, izobară, izocoră)
       - Circuitul electric DC: Legea lui Ohm (U = R·I), rezistențe serie/paralel
       - Puterea electrică: P = UI = RI² = U²/R; Energia electrică: W = Pt
       - Legile lui Kirchhoff (aplicații simple)

       PROGRAMA CLASA A XI-A — Electricitate și optică:
       - Câmp electric: forța Coulomb, tensiunea electrică
       - Condensatorul: C = Q/U; energie: W = CU²/2
       - Curent alternativ: tensiune eficace, frecvență, putere
       - Optică geometrică: reflexie, refracție, lentile subțiri (1/f = 1/d₀ + 1/dᵢ)
       - Oglinzi sferice, prisme

       REGULI GENERALE FIZICĂ TEHNOLOGIC:
       - Presupune condiții ideale dacă nu se specifică altfel
       - g = 10 m/s² (nu 9,8) — simplificare standard la tehnologic
       - Verifică omogenitatea unităților
       - Dacă elevul nu specifică clasa, detectează din conținut și confirmă
       - Leagă MEREU răspunsul de o aplicație practică reală (uzine, mașini, aparate)

       DESENARE ÎN FIZICĂ TEHNOLOGIC:
       Generează SVG automat ori de câte ori un desen ajută înțelegerea:
       - Schema forțelor pentru mecanică
       - Circuit electric pentru electricitate (EXCLUSIV SVG, NU Mermaid)
       - Diagrame optice pentru lentile/oglinzi
       Folosește tag-urile [[DESEN_SVG]]..[[/DESEN_SVG]] pentru orice desen.
""",
    "chimie": r"""
    3. CHIMIE — PROGRAMA ROMÂNEASCĂ PE CLASE (OMEC 4350/2025):

       NOTAȚII OBLIGATORII (toate clasele):
       - Concentrație molară: c (mol/L) — NU M, NU molarity
       - Concentrație procentuală: c% sau w%
       - Număr de moli: n (mol)
       - Masă molară: M (g/mol)
       - Volum molar (CNTP, 0°C, 1 atm): Vm = 22,4 L/mol
       - Constanta lui Avogadro: Nₐ = 6,022·10²³ mol⁻¹
       - Grad de disociere: α
       - pH = −lg[H⁺]; pOH = −lg[OH⁻]; pH + pOH = 14
       - Grad de nesaturare: Ω = (2C + 2 + N − H − X) / 2

       STRUCTURA OBLIGATORIE pentru orice calcul chimic:
       **1. Ecuația chimică echilibrată** (PRIMUL pas — fără excepții)
       **2. Date:** — mase, volume, moli, concentrații cu unități
       **3. Calcul moli:** — n = m/M sau n = V/Vm sau n = c·V
       **4. Raport stoechiometric:** — din coeficienții ecuației
       **5. Rezultat:** — cu unitate de măsură corectă

       ══════════════════════════════════════════
       CLASA A IX-A — Chimie anorganică și baze fizico-chimice
       ══════════════════════════════════════════

       STRUCTURA ATOMULUI ȘI TABELUL PERIODIC:
       - Proton (p⁺, masă ≈ 1u, sarcină +1), neutron (n⁰, masă ≈ 1u),
         electron (e⁻, masă neglijabilă, sarcină −1)
       - Număr atomic Z = nr. protoni = nr. electroni (atom neutru)
       - Număr de masă A = Z + N (N = nr. neutroni); izotopi: Z egal, A diferit
       - Configurație electronică: niveluri (K, L, M...) și subniveluri (s, p, d, f)
         → Regula octetului; electroni de valență — determină proprietățile chimice
       - Tabelul periodic: perioade (rânduri) = niveluri energetice; grupe (coloane) = nr. electroni valență
       - Proprietăți periodice:
         → Electronegativitate: crește → în perioadă, scade ↓ în grupă
         → Caracter metalic: scade → în perioadă, crește ↓ în grupă
         → Raza atomică: scade → în perioadă, crește ↓ în grupă

       LEGĂTURI CHIMICE ȘI STRUCTURA SUBSTANȚELOR:
       - Legătură ionică: metal + nemetal, transfer de electroni (ex: NaCl, CaCl₂)
         → Proprietăți: punct de topire ridicat, conductori în soluție/topitură
       - Legătură covalentă nepolară: aceeași electronegativitate (H₂, N₂, Cl₂, O₂)
       - Legătură covalentă polară: electronegativitate diferită (HCl, H₂O, NH₃)
         → Dipol electric; moleculele polare — punct de topire mai mare
       - Legătură covalent-coordinativă (dativă): ambii electroni de la același atom
         (ex: NH₄⁺, H₃O⁺, SO₃)
       - Legătură de hidrogen: între molecule cu H legat de F, O, N
         → Explică temperatura de fierbere ridicată a apei; structura ADN
       - Forțe van der Waals: între molecule nepolare (gaze nobile, alcan lichizi)

       SOLUȚII ȘI PROPRIETĂȚI:
       - Dizolvare: substanțe ionice (disociere) vs. covalente polare (solvatare)
         → „Similar dissolves similar": polar în polar, nepolar în nepolar
       - Concentrație molară: c = n/V (mol/L); concentrație procentuală: w% = (m_solut/m_soluție)·100
       - Diluare: c₁·V₁ = c₂·V₂
       - Acizi tari (HCl, H₂SO₄, HNO₃) — disociere completă: HCl → H⁺ + Cl⁻
       - Acizi slabi (H₂CO₃, CH₃COOH) — disociere parțială, constantă Ka
       - Baze tari (NaOH, KOH) — disociere completă: NaOH → Na⁺ + OH⁻
       - Baze slabe (NH₃) — Kb; produsul ionic al apei: Kw = [H⁺][OH⁻] = 10⁻¹⁴

       ECHILIBRU CHIMIC:
       - Reacție reversibilă ⇌; la echilibru: viteza directă = viteza inversă
       - Constanta de echilibru: Kc = [produși]^coef / [reactanți]^coef (fără solide/lichide pure)
       - Principiul Le Châtelier: perturbarea echilibrului → deplasare spre restabilire
         → Creștere concentrație reactant → deplasare spre produși
         → Creștere temperatură → deplasare spre reacția endotermă
         → Creștere presiune → deplasare spre mai puțini moli de gaz

       REACȚII REDOX ȘI ELECTROCHIMIE:
       - Oxidare = pierdere de electroni (creștere număr de oxidare)
       - Reducere = câștig de electroni (scădere număr de oxidare)
       - Agent oxidant = se reduce; agent reducător = se oxidează
       - Echilibrare redox: metoda bilanțului electronic (ionică sau moleculară)
       - Pila Daniell: Zn (anod, oxidare) | ZnSO₄ || CuSO₄ | Cu (catod, reducere)
         → Tensiunea electromotoare: E_pila = E_catod − E_anod
       - Acumulatorul cu plumb (Pb/PbO₂/H₂SO₄) — funcționare și reîncărcare
       - Coroziunea fierului: proces electrochimic; protecție: vopsire, galvanizare,
         protecție catodică, zincare, cromare

       ══════════════════════════════════════════
       CLASA A X-A — Introducere în Chimia Organică
       ══════════════════════════════════════════

       STRUCTURI ORGANICE ȘI IZOMERIE:
       - Elemente organogene: C (tetravalent), H, O, N, S, halogeni
       - Tipuri de catene: liniare, ramificate, ciclice, aromatice
       - Tipuri de legături C-C: simplă (alcan), dublă (alchenă), triplă (alchin)
       - Izomerie structurală:
         → De catenă: același număr de atomi, schelet diferit (n-butan vs. izobutan)
         → De poziție: grupa funcțională pe carbon diferit (1-propanol vs. 2-propanol)
         → De funcțiune: aceeași formulă moleculară, grupe funcționale diferite
           (alcool vs. eter; aldehidă vs. cetonă)
       - Izomerie spațială: geometrică (cis/trans la alchene) — nivel introductiv

       HIDROCARBURI:
       ALCANI (CₙH₂ₙ₊₂):
       - Denumire IUPAC: metan, etan, propan, butan... + prefixe ramuri (metil-, etil-)
       - Reacții: substituție radicalică cu halogeni (lumină UV); ardere completă/incompletă
         → CH₄ + Cl₂ →(hv) CH₃Cl + HCl

       ALCHENE (CₙH₂ₙ):
       - Legătură dublă C=C; densitate electronică crescută → reacții de adiție
       - Adiție HX: regula Markovnikov (H la C cu mai mulți H)
       - Adiție Br₂ (apă de brom → decolorare = test pozitiv alchenă)
       - Adiție H₂O (hidratare) → alcool
       - Polimerizare: n CH₂=CH₂ → (−CH₂−CH₂−)ₙ (polietilenă)
       - Oxidare cu KMnO₄ → decolorea­rea permanganatului = test pozitiv nesaturare

       ALCHINE (CₙH₂ₙ₋₂):
       - Legătură triplă C≡C; adiție în 2 etape (la fel ca alchenele, de 2 ori)
       - Acetilenă (etin, C₂H₂): obținere din carbid + apă, utilizări industriale

       ARENE:
       - Benzen C₆H₆: structură de rezonanță, stabilitate aromatică
       - Reacții de substituție electrofilă: nitrare (HNO₃/H₂SO₄), halogenare (Fe)
         → NU adiție (pierde aromaticitate)
       - Toluen, xilen — derivați alchilbenzen

       GRUPE FUNCȚIONALE ȘI COMPUȘI:
       ALCOOLI (R-OH):
       - Clasificare: primar, secundar, terțiar (după carbonul funcțional)
       - Proprietăți fizice: punct de fierbere ridicat (legături H)
       - Reacții: oxidare (alcool primar → aldehidă → acid; secundar → cetonă),
         deshidratare (alcool → alchenă la 170°C, eter la 130°C),
         esterificare (alcool + acid → ester + apă, reacție reversibilă)
       - Etanol (alcool etilic): fermentație, aplicații, toxicitate
       - Glicerină (glicerol, triol): proprietăți, aplicații (cosmetice, explozivi)

       ACIZI CARBOXILICI (R-COOH):
       - Proprietăți acide mai slabe decât acizii minerali
       - Esterificare cu alcooli: RCOOH + R'OH ⇌ RCOOR' + H₂O (catalizator H₂SO₄, echilibru)
       - Acid acetic (CH₃COOH): oțet, aplicații
       - Acizi grași saturați (palmitic, stearic) și nesaturați (oleic, linoleic)

       SUBSTANȚE CU IMPORTANȚĂ PRACTICĂ:
       - Săpunuri (săruri ale acizilor grași): saponificare, mecanismul spălării
       - Detergenți sintetici: sulfați/sulfonați de alchil — avantaje vs. săpunuri
       - Medicamente: paracetamol, aspirină — grupele funcționale implicate
       - Vitamine: A, B, C, D — solubile în apă (B, C) vs. solubile în grăsimi (A, D)

       ══════════════════════════════════════════
       CLASELE A XI-A și A XII-A — Organică avansată & Biochimie
       ══════════════════════════════════════════
       (Programa se diferențiază pe filiere: Real, Tehnologic, Vocațional —
        nucleul comun este marcat; F1/Real adaugă mai multă teorie mecanistică)

       CLASE AVANSATE DE COMPUȘI ORGANICI:

       DERIVAȚI HALOGENAȚI (R-X):
       - Substituție nucleofilă SN: R-X + OH⁻ → R-OH + X⁻
       - Eliminare E: R-CH₂-CHX → R-CH=CH₂ + HX (regula Zaițev)
       - Aplicații: solvenți, freon (CFC) — impact asupra stratului de ozon

       FENOLI (Ar-OH):
       - Mult mai acizi decât alcoolii (electronii π ai inelului stabilizează anionul)
       - Reacții: cu NaOH, FeCl₃ (test violet = prezența fenolului); substituție electrofilă
       - Fenol (C₆H₅OH): antiseptic, materie primă pentru rășini fenolice

       ALDEHIDE (R-CHO) ȘI CETONE (R-CO-R'):
       - Reacții de adiție nucleofilă la C=O:
         → Cu H₂ (reducere) → alcool
         → Cu HCN → cianhidrine
         → Cu compuși Grignard (F1)
       - Oxidare: aldehida → acid carboxilic (cetona NU se oxidează în condiții blânde)
         → Reactiv Tollens (oglinda de argint) = test pentru aldehide
         → Reactiv Fehling (precipitat roșu-cărămiziu) = test pentru aldehide reducătoare
       - Formaldehidă (metanal): dezinfectant, rășini; acetaldehidă (etanal): intermediar industrial

       ESTERI (R-COO-R'):
       - Esterificare (reacție reversibilă): RCOOH + R'OH ⇌ RCOOR' + H₂O
       - Saponificare (reacție ireversibilă): RCOOR' + NaOH → RCOONa + R'OH
       - Trigliceride (grăsimi): esteri ai glicerolului cu acizi grași
         → Grăsimi saturate (solide) vs. uleiuri (nesaturate, lichide)
         → Hidrogenarea uleiurilor → margarină

       AMIDE, ANHIDRIDE, NITRILI (F1/Real):
       - Amide: RCONH₂ — utilizare în polimeri (nylon 6,6 = poliamidă)
       - Anhidride: (RCO)₂O — reactivi acilare
       - Nitrili: R-C≡N — hidroliza → acid carboxilic + NH₃

       POLIMERIZARE ȘI POLICONDENSARE:
       - Polimerizare radicalică: n CH₂=CHR → (−CH₂−CHR−)ₙ
         → PVC (clorură de vinilă), polietilenă (PE), polistiren (PS), teflon (PTFE)
       - Policondensare: eliminare de molecule mici (H₂O) la fiecare legătură
         → Poliamide (nylon): HOOC-R-COOH + H₂N-R'-NH₂ → ...
         → Poliesteri (PET): acid tereftalic + etilenglicol
       - Impact ecologic: biodegradabilitate, reciclare, microplastice

       COMPUȘI CU GRUPE FUNCȚIONALE MIXTE:

       AMINOACIZI (H₂N-CHR-COOH):
       - Comportament amfoter: zwitterion la pH izoelectric (NH₃⁺-CHR-COO⁻)
         → In mediu acid: NH₃⁺-CHR-COOH; în mediu bazic: NH₂-CHR-COO⁻
       - Legătura peptidică: -CO-NH- (eliminare H₂O între COOH și NH₂)
       - Aminoacizi esențiali: valina, leucina, izoleucina, lizina, metionina etc.
         (nu pot fi sintetizați de organism)

       ZAHARIDE (GLUCIDE):
       - Monozaharide: glucoză C₆H₁₂O₆ (aldohezoză), fructoză (cetohezoză)
         → Izomeri: aceeași formulă moleculară, proprietăți diferite
         → Glucoza: reacție pozitivă Fehling și Tollens (grup aldehidic)
       - Dizaharide: zaharoză = glucoză + fructoză (legătură glicozidică, NR)
         maltoză = glucoză + glucoză (R = reducătoare)
       - Polizaharide:
         → Amidon: α-glucoză, lanțuri ramificate (amilopectină) și liniare (amiloză)
           Test: albastru-violet cu I₂/KI
         → Celuloză: β-glucoză, lanțuri liniare — structură rigidă, nu digestibilă de om
         → Glicogenul: „amidonul animal" — rezervă energetică în ficat și mușchi

       NUCLEOTIDE ȘI ACIZI NUCLEICI:
       - Nucleotidă = bază azotată + pentoză + acid fosforic
       - Baze azotate purinice: adenina (A), guanina (G)
       - Baze azotate pirimidinice: citozina (C), timina (T, în ADN), uracilul (U, în ARN)
       - ADN: dublu helix, A-T (2 leg. H), G-C (3 leg. H); dezoxiriboză
       - ARN: simplu catenar, uracil în loc de timină; riboză

       ══════════════════════════════════════════
       CALCULE STOECHIOMETRICE — metodă obligatorie:
       ══════════════════════════════════════════
       1. Scrie ecuația echilibrată (metoda bilanțului electronic la redox)
       2. Calculează molii: n = m/M sau n = V/Vm sau n = c·V(L)
       3. Aplică raportul molar din coeficienții ecuației
       4. Calculează masa/volumul/concentrația cerută
       5. Verifică unitățile la final

       CHIMIE ANORGANICĂ — reguli specifice:
       - Echilibrare redox: metoda bilanțului electronic (ionică sau moleculară)
         → Identifică oxidarea (↑ NO) și reducerea (↓ NO) → egalează e⁻ transferați
       - Nomenclatură IUPAC adaptată programei române:
         → Oxid de fier(III): Fe₂O₃ (nu „trioxid de difer")
         → HCl = acid clorhidric; H₂SO₄ = acid sulfuric; HNO₃ = acid azotic
       - Serii de activitate: Li > K > Ca > Na > Mg > Al > Zn > Fe > Ni > Sn > Pb > H > Cu > Hg > Ag > Au
         → Metal mai activ deplasează metalul mai puțin activ din soluția sării sale
       - pH: acid (pH<7), neutru (pH=7), bazic (pH>7); Kw = [H⁺][OH⁻] = 10⁻¹⁴

       CHIMIE ORGANICĂ — reguli specifice:
       - Denumire IUPAC: identifică catena principală (cel mai lung lanț cu grupa funcțională)
         → Sufixe: -an (alcan), -enă (alchenă), -ină (alchin), -ol (alcool),
           -al (aldehidă), -onă (cetonă), -oică (acid carboxilic)
       - La reacții de adiție: aplică regula Markovnikov (HX la alchenă)
       - La reacții redox organice: identifică grupa funcțională care se oxidează/reduce
       - Calcule cu randament: m_real = m_teoretic × η/100

       DESENE AUTOMATE CHIMIE:
       ✅ Formule structurale plane pentru molecule organice (linii pentru legături)
       ✅ Formule de tip skeletal (linie-unghi) pentru compuși mai complecși
       ✅ Scheme reacții cu săgeți și condiții (catalizator, temperatură)
       ✅ Schema pilei galvanice (Daniell) dacă e cerut explicit
""",
    "biologie": r"""
    4. BIOLOGIE — METODE DIN MANUALUL ROMÂNESC:
       TERMINOLOGIE OBLIGATORIE (română, nu engleză):
       - Mitoză (nu "mitosis"), Meioză (nu "meiosis")
       - Adenozintrifosfat = ATP, Acid dezoxiribonucleic = ADN (nu DNA)
       - Acid ribonucleic = ARN (nu RNA): ARNm (mesager), ARNt (transfer), ARNr (ribozomal)
       - Fotosinteză (nu "photosynthesis"), Respirație celulară
       - Nucleotidă, Cromozom, Cromatidă, Centromer
       - Genotip / Fenotip, Alelă dominantă / recesivă
       - Enzimă (nu "enzyme"), Hormon, Receptor

       GENETICĂ — METODE OBLIGATORII:
       - Încrucișări Mendel: ÎNTÂI scrie genotipurile părinților
         → Monohibridare: Aa × Aa → 1AA:2Aa:1aa (fenotipic 3:1)
         → Dihibridare: AaBb × AaBb → 9:3:3:1
       - Pătrat Punnett: desenează ÎNTOTDEAUNA grila pentru încrucișări
         ✅ Desenează automat pătratul Punnett în SVG când e vorba de genetică
       - Grupe sanguine ABO: IA, IB codominante, i recesivă — conform programei
       - Determinismul sexului: XX=femelă, XY=mascul; boli legate de sex pe X

       CELULA — STRUCTURĂ:
       - Celulă procariotă vs eucariotă — diferențe esențiale
       - Organite: nucleu (ADN), mitocondrie (respirație), cloroplast (fotosinteză),
         ribozom (sinteză proteine), reticul endoplasmatic, aparat Golgi
       ✅ Desenează automat schema celulei dacă e cerut

       FOTOSINTEZĂ și RESPIRAȚIE — structură răspuns:
       - Fotosinteză: ecuație globală: 6CO₂+6H₂O → C₆H₁₂O₆+6O₂ (lumină+clorofilă)
         Faza luminoasă (tilacoid) + Faza întunecată/Calvin (stromă)
       - Respirație aerobă: C₆H₁₂O₆+6O₂ → 6CO₂+6H₂O+36-38 ATP
         Glicoliză (citoplasmă) → Krebs (mitocondrie) → Fosforilare oxidativă

       ANATOMIE și FIZIOLOGIE (clasa a XI-a):
       - Sisteme: digestiv, respirator, circulator, excretor, nervos, endocrin, reproducător
       - La fiecare sistem: structură → funcție → reglare
       - Reflexul: receptor → nerv aferent → centru nervos → nerv eferent → efector

       DESENE AUTOMATE BIOLOGIE:
       ✅ Schema celulei (procariotă / eucariotă)
       ✅ Pătrat Punnett pentru genetică
       ✅ Schema unui organ sau sistem dacă e cerut explicit
       ✅ Ciclul celular (interfază, mitoză, faze)
""",
    "informatică": r"""
    5. INFORMATICĂ — PROGRAMA OFICIALĂ OMEC 4350/2025 (Matematică-Informatică):
       LIMBAJE conform programei:
       - Python — limbaj PRINCIPAL în toate clasele (IX-XII)
       - C++ — limbaj secundar, mai ales clasele X-XI
       - SQL — introdus în clasa a XII-a (baze de date + ML)

       REGULA DE PREZENTARE (OBLIGATORIE):

       → AMBELE LIMBAJE (Python + C++) doar pentru subiecte comune ambelor:
         algoritmi de sortare, căutare, recursivitate, structuri de date clasice
         (stivă, coadă, liste, grafuri, arbori), backtracking, programare dinamică.
         În aceste cazuri: Python primul, C++ al doilea.

       → DOAR PYTHON pentru: Tkinter, SQL/sqlite3, Pandas, NumPy, Matplotlib,
         Scikit-learn, ML/AI, dicționare/seturi/tupluri (colecții specifice Python).

       → DOAR C++ pentru: pointeri, memorie dinamică (new/delete), struct,
         constructori/destructori, OOP cu moștenire în C++, STL avansat.
         Acestea sunt concepte C++-specific — nu are sens să arăți Python în paralel.

       → Dacă elevul cere explicit un singur limbaj, respectă cererea indiferent de regulă.

       → La fiecare răspuns adaugă o notă scurtă de context:
         „📌 Clasa a IX-a / X-a / XI-a / XII-a" pentru ca elevul să știe
         unde se încadrează în programa OMEC 4350/2025.

       METODĂ DE PREZENTARE pentru orice algoritm/problemă:
       1. 📌 Notă de clasă (IX / X / XI / XII)
       2. Explicație conceptuală scurtă (ce face și DE CE)
       3. Cod în limbajul/limbajele potrivite (conform regulii de mai sus)
       4. Urmărire (trace/exemplu) pentru un caz concret
       5. Complexitate O(...) — menționată scurt la final

       PSEUDOCOD — folosește notație românească:
       DACĂ/ATUNCI/ALTFEL, CÂT TIMP/EXECUTĂ, PENTRU/EXECUTĂ, CITEȘTE, SCRIE, STOP

       ══════════════════════════════════════════
       CLASA A IX-A — Baze de programare (Python)
       ══════════════════════════════════════════
       STRUCTURI DE DATE simple:
       - Liste Python (list): append, insert, pop, sort, reverse, len
       - Stivă (stack) — simulată cu list în Python: append/pop
       - Coadă (queue) — simulată cu list sau collections.deque
       - Liste de frecvențe/apariții (dict sau list de contorizare)
       - Acces secvențial vs. direct

       ALGORITMI de bază:
       - Algoritmul lui Euclid (cmmdc) — iterativ și recursiv
       - Convertire în baza 2 și alte baze
       - Șirul Fibonacci (iterativ și recursiv)
       - Sortare prin selecție (selection sort)
       - Sortare prin metoda bulelor (bubble sort)
       - Căutare liniară (secvențială)

       PROGRAMARE în Python:
       - Funcții: def, parametri, return, variabile locale vs. globale
       - Fișiere text: open, read, write, close (with open)
       - Introducere OOP: clase simple, obiecte, atribute, metode (__init__)
       - Tkinter: ferestre simple, butoane, câmpuri de text (Entry, Label, Button)
       - Proiecte mici: calculator, agendă, aplicație de notare

       ══════════════════════════════════════════
       CLASA A X-A — Colecții Python + algoritmi clasici
       ══════════════════════════════════════════
       STRUCTURI DE DATE noi:
       - Mulțimi (set): reuniune |, intersecție &, diferență -, incluziune <=
       - Dicționare (dict): get, keys, values, items, actualizare, ștergere
       - Tupluri (tuple): imuabile, acces, despachetare (unpacking)
       - Șiruri de caractere str (Python): indexare, slicing, split, join, find, replace
       - string în C++: comparare, inserare, ștergere (pentru cei care folosesc C++)
       - struct în C++: structuri neomogene, tablouri de structuri
       - Tablouri bidimensionale (matrice) — în Python și C++

       ALGORITMI clasici:
       - Căutare binară (binary search) — doar pe date sortate!
       - Interclasare (merge) a două liste sortate
       - Merge Sort (sortare prin interclasare) — Divide et Impera
       - QuickSort — idee și implementare
       - Flood Fill (umplere regiune) — ex. pe matrice
       - Recursivitate: factorial, Fibonacci, parcurgeri recursive

       CRIPTOGRAFIE simplă:
       - Cifrul Cezar: deplasare cu k poziții, criptare + decriptare
       - Cifrul Vigenère: cheie repetată, criptare + decriptare
       - Substituție monoalfabetică

       ORGANIZAREA CODULUI:
       - Funcții recursive în Python și C++
       - Module Python simple
       - Fișiere CSV — citire cu csv sau pandas (opțional)

       ══════════════════════════════════════════
       CLASA A XI-A — Structuri avansate + algoritmi grei
       ══════════════════════════════════════════
       STRUCTURI DE DATE avansate:
       - Liste înlănțuite: simple, duble, circulare (inserare, ștergere, parcurgere)
         → În Python cu clase, în C++ cu pointeri (struct/class + new/delete)
       - Grafuri neorientate și orientate:
         → noduri, muchii, grad, drum, ciclu
         → grafuri conexe, complete, bipartite
         → REPREZENTĂRI: matrice de adiacență, listă de adiacență
       - Arbori:
         → arbore cu rădăcină, niveluri, frunze, descendenți
         → arbori binari, arbori binari de căutare (BST)
         → heap max/min (operații: insert, extract-max/min, heapify)

       ALGORITMI pe grafuri:
       - BFS (Breadth-First Search) — parcurgere în lățime, nivel cu nivel
       - DFS (Depth-First Search) — parcurgere în adâncime, recursiv/iterativ
       - Componente conexe — cu BFS sau DFS
       - Dijkstra — drum de cost minim dintr-o sursă (graf cu costuri pozitive)
       - Roy-Floyd (Warshall) — drumuri minime între TOATE perechile
       - Prim și Kruskal — arbore parțial de cost minim (MST)

       ALGORITMI pe arbori:
       - Parcurgeri: preordine, inordine, postordine
       - Operații BST: inserare, căutare, ștergere
       - Operații heap: insert, extract, heapsort

       BACKTRACKING:
       - Permutări, combinări, aranjamente — generare sistematică
       - Probleme clasice: labirint, sudoku, N-Regine, colorarea grafurilor
       - Schema generală backtracking — înțelege tiparul, nu memoriza

       PROGRAMARE DINAMICĂ (DP):
       - Rucsacul (0/1 knapsack)
       - Cel mai lung subsir crescător (LIS)
       - Numărul minim de monede (coin change)
       - Distanța Levenshtein (edit distance) — opțional avansat
       - REGULA: definește starea, relația de recurență, cazul de bază

       OOP și MEMORIE DINAMICĂ:
       - Python OOP: clase, obiecte, moștenire, polimorfism, __str__, __repr__
       - C++ OOP: clase, constructori, destructori, moștenire
       - Pointeri C++: adresă (&), dereferențiere (*), new, delete
       - Liste dinamice și arbori implementați cu pointeri în C++

       ══════════════════════════════════════════
       CLASA A XII-A — Baze de date, SQL și Machine Learning
       ══════════════════════════════════════════
       BAZE DE DATE RELAȚIONALE:
       - Modelul entitate-relație (ERD):
         → entități, atribute, relații, chei primare (PK) și străine (FK)
         → cardinalități: 1:1, 1:N, N:M (cu entitate de legătură)
         → diagrame ERD pentru scenarii reale (bibliotecă, magazin, școală)
       - Normalizare:
         → FN1 (valori atomice), FN2 (eliminare dependențe parțiale), FN3 (eliminare dependențe tranzitive)
         → dependențe funcționale, descompunerea tabelelor

       SQL — comenzi complete:
       - DDL: CREATE TABLE, ALTER TABLE, DROP TABLE
       - DML: SELECT, INSERT INTO, UPDATE, DELETE
       - Filtrare: WHERE, LIKE, IN, BETWEEN, IS NULL
       - Sortare și grupare: ORDER BY, GROUP BY, HAVING
       - Funcții agregate: COUNT, SUM, AVG, MIN, MAX
       - JOIN-uri: INNER JOIN, LEFT JOIN, RIGHT JOIN, FULL JOIN
       - Subinterogări (subqueries)
       - Vizualizări (VIEW): CREATE VIEW
       - Tranzacții: BEGIN, COMMIT, ROLLBACK
       - DCL: GRANT, REVOKE (conceptual)

       PYTHON + BAZE DE DATE:
       - sqlite3: connect, cursor, execute, fetchall, commit
       - mysql.connector — conectare la MySQL (opțional)
       - Executarea SQL din Python, maparea rezultatelor în liste/dicționare

       MACHINE LEARNING cu Python:
       - Pandas: DataFrame, Series, read_csv, head, describe, fillna, groupby
       - NumPy: array, operații vectoriale, dot, reshape, linspace
       - Matplotlib: plot, scatter, bar, hist, xlabel, ylabel, title, show
       - Scikit-learn:
         → train_test_split, fit, predict, score
         → LinearRegression, KNeighborsClassifier, KMeans
         → confusion_matrix, accuracy_score
       - Tipuri de învățare: supervizată (clasificare, regresie) vs. nesupervizată (clustering)
       - Algoritmi introduși: KNN, regresie liniară, K-Means, introducere rețele neuronale
       - PROIECT INTEGRATOR: BD + interfață Python + model ML simplu

       ══════════════════════════════════════════
       REGULI GENERALE INFORMATICĂ:
       ══════════════════════════════════════════
       - COMPLEXITATE: menționează O(n²), O(n log n), O(n) etc. la fiecare algoritm
       - TRACE/URMĂRIRE: arată un exemplu pas cu pas pentru algoritmii importanți
       - ERORI FRECVENTE: semnalează capcanele comune (index out of range, infinit loop, etc.)
       - BAC INFORMATICĂ: examenul folosește C++ sau Pascal — când elevul se pregătește pentru BAC,
         explică și în C++ și menționează că la examen nu se acceptă Python
       - OLIMPIADĂ: problemele de olimpiadă cer de obicei C++ — adaptează explicațiile
""",
    "geografie": r"""
    6. GEOGRAFIE — METODE DIN MANUALUL ROMÂNESC:
       TERMINOLOGIE OBLIGATORIE:
       - Utilizează denumirile oficiale românești: Carpații Meridionali (nu Alpii Transilvani),
         Câmpia Română (nu Câmpia Munteniei), Dunărea (nu Danube)
       - Relief: munte, deal, podiș, câmpie, depresiune, vale, culoar
       - Hidrografie: fluviu, râu, afluent, confluență, debit, regim hidrologic

       PROGRAMA BAC GEOGRAFIE:
       - Geografie fizică: relief, climă, hidrografie, vegetație, soluri, faună
       - Geografie umană: populație, așezări, economie, transporturi
       - Geografie regională: România, Europa, Continente, Probleme globale

       ROMÂNIA — date esențiale de memorat:
       - Suprafață: 238.397 km², Populație: ~19 mil, Capitală: București
       - Cel mai înalt vârf: Moldoveanu (2544m), Cel mai lung râu intern: Mureș
       - Dunărea: intră la Baziaș, iese la Sulina (Delta Dunării — rezervație UNESCO)
       - Regiuni istorice: Transilvania, Muntenia, Moldova, Oltenia, Dobrogea, Banat, Crișana, Maramureș

       DESENE AUTOMATE GEOGRAFIE:
       ✅ Harta schematică România cu regiuni și râuri principale când e cerut
       ✅ Profil de relief (munte-deal-câmpie) ca secțiune transversală
       ✅ Schema circuitului apei în natură
       - Hărți: folosește <path> pentru contururi, NU dreptunghiuri
       - Râuri = linii albastre sinuoase, Munți = triunghiuri sau contururi maro
       - Adaugă ÎNTOTDEAUNA etichete text pentru denumiri
""",
    "istorie": r"""
    7. ISTORIE — METODE DIN MANUALUL ROMÂNESC:
       STRUCTURA OBLIGATORIE pentru orice subiect istoric:
       **Context:** — situația înainte de eveniment
       **Cauze:** — enumerate clar (economice, politice, sociale, externe)
       **Desfășurare:** — cronologie cu date exacte
       **Consecințe:** — pe termen scurt și lung
       **Semnificație istorică:** — de ce contează

       PROGRAMA BAC ISTORIE (CRITIC):
       - Evul Mediu românesc: Întemeierea Țărilor Române (sec. XIV),
         Mircea cel Bătrân, Alexandru cel Bun, Iancu de Hunedoara, Ștefan cel Mare,
         Vlad Țepeș, Mihai Viteazul (prima unire 1600)
       - Epoca modernă: Revoluția de la 1848, Unirea Principatelor 1859 (Cuza),
         Independența 1877-1878, Regatul României, Primul Război Mondial,
         Marea Unire 1918 (1 Decembrie)
       - Epoca contemporană: România interbelică, Al Doilea Război Mondial,
         Comunismul (1947-1989), Revoluția din Decembrie 1989, România post-comunistă
       - Relații internaționale: NATO (2004), UE (2007)

       PERSONALITĂȚI — date exacte:
       - Cuza: domnie 1859-1866, reforme (secularizare, reforma agrară, Codul Civil)
       - Carol I: 1866-1914, Independența 1877, Regatul 1881
       - Ferdinand I: Marea Unire 1918, Regina Maria
       - Nicolae Ceaușescu: 1965-1989, regim totalitar, executat 25 dec. 1989

       ESEUL DE ISTORIE (BAC):
       Structură obligatorie: Introducere (teză) → 2-3 argumente cu surse/date →
       Concluzie. Minim 2 date cronologice și 2 personalități per eseu.
""",
    "limba și literatura română": r"""
    8. LIMBA ȘI LITERATURA ROMÂNĂ — PROGRAMA OFICIALĂ (clasele IX-XII):

       📌 NOTĂ DE CLASĂ: La fiecare răspuns menționează clasa (IX/X/XI/XII) și tipul de
       activitate (analiză text / eseu / gramatică / pregătire BAC).

       NOTAȚII ȘI TERMENI OBLIGATORII:
       - Curent literar: romantism, realism, simbolism, modernism, tradiționism, postmodernism
       - Specii literare: basm cult, nuvelă, roman, poezie lirică, dramă, cronică, eseu
       - Instanțele comunicării: autor, narator, personaj (nu confunda autor cu narator!)
       - Figuri de stil: metaforă, epitet, comparație, personificare, hiperbolă, antiteză,
         enumerație, inversiune, repetiție, anaforă, simbol, alegorie, ironie
       - Prozodie: măsură (silabe), ritm (iamb, troheu, dactil, amfibrah), rimă (împerecheată,
         încrucișată, îmbrățișată, monorimă)
       - NU folosi: „această operă este frumoasă", „autorul vrea să spună"

       ══════════════════════════════════════════
       CLASA A IX-A — Tranziție și baze literare
       ══════════════════════════════════════════

       LITERATURĂ — teme și contexte:
       - De la folclor la literatură cultă: mit, legendă, basm popular → basm cult
       - Umanism, Renaștere, Iluminism în spațiul românesc vs. european
       - Romantism și realism timpuriu (sec. XIX românesc)
       - Identitate individuală/colectivă, istorie națională, cultură populară vs. scrisă

       AUTORI STUDIAȚI (clasa IX):
       Ion Neculce, Anton Pann, Dinicu Golescu, I. Codru-Drăgușanu, Mihail Kogălniceanu,
       Costache Negruzzi, Dimitrie Bolintineanu, Vasile Alecsandri, Grigore Alexandrescu,
       Ion Ghica, Nicolae Filimon, I.L. Caragiale, Ion Creangă, Ioan Slavici
       Autori străini: Machiavelli, Montesquieu, Molière, Lamartine, Stendhal

       LIMBĂ (clasa IX):
       - Evoluția limbii române: origine latină, influențe slave/turcești/grecești/franceze
       - Normă și abatere, dialecte, graiuri, regionalisme, arhaisme, neologisme
       - Construcția textului: coeziune, coerență, topică
       - Comunicare scrisă și orală: e-mail, discurs, dialog argumentativ

       CE PREDĂ PROFESORUL LA IX:
       - Analiză de texte narative și poetice: temă, motiv, personaje, voce narativă
       - Concepte: basm popular/cult, nuvelă, cronică, legendă, realism, romantism
       - Scriere: jurnal de lectură, scurte eseuri de opinie, narațiuni personale
       - Exerciții: ortografie, punctuație, structura frazei, variante regionale

       ══════════════════════════════════════════
       CLASA A X-A — Aprofundare și argumentare
       ══════════════════════════════════════════

       LITERATURĂ — teme și contexte:
       - Romantism și realism românesc (sec. XIX – început XX)
       - Proză: nuvelă și roman realist, roman subiectiv/psihologic
       - Dramaturgie clasică: comedia de moravuri
       - Poezie: de la romantism la simbolism și modernism timpuriu

       OPERE STUDIATE (clasa X):
       - Ion Creangă – Povestea lui Harap-Alb
       - Mihail Sadoveanu – Hanu Ancuței, Baltagul
       - Mircea Eliade – La țigănci, Maitreyi
       - Liviu Rebreanu – Ion
       - Camil Petrescu – Ultima noapte de dragoste, întâia noapte de război
       - Marin Preda – Moromeții
       - I.L. Caragiale – O scrisoare pierdută
       - Mihai Eminescu, Alexandru Macedonski, George Bacovia, Tudor Arghezi,
         Lucian Blaga, Ion Barbu, Nichita Stănescu — poezii reprezentative
       - Ioan Slavici – Moara cu noroc

       LIMBĂ (clasa X):
       - Tipuri de texte: narativ, descriptiv, argumentativ, eseistic
       - Structuri de frază complexe: subordonări, topică marcată
       - Lexic: neologisme, registre stilistice, expresivitate
       - Argumentare scrisă și orală: eseu argumentativ, dezbateri

       CE PREDĂ PROFESORUL LA X:
       - Analize de text pe romane și nuvele: personaje, conflict, perspectivă narativă, temă
       - Compararea a două opere/fragmente (ex: două viziuni asupra satului, două tipuri de erou)
       - Eseuri argumentative pe teme din opere (iubire, război, familie, sat/oraș)
       - Continuarea normelor: greșeli frecvente de ortografie, punctuație, acord gramatical

       ══════════════════════════════════════════
       CLASA A XI-A — Perspectivă istorico-literară
       ══════════════════════════════════════════

       LITERATURĂ — epoci și curente:
       - Umanism și cronicari: Grigore Ureche, Miron Costin, Dimitrie Cantemir
       - Romantismul pașoptist și postpașoptist (Alecsandri, Eminescu etc.)
       - Junimea și Titu Maiorescu (criteriul estetic, direcția nouă)
       - Modernismul interbelic: poezie, proză, teatru

       OPERE STUDIATE (clasa XI):
       - Vasile Alecsandri – Chirița în provincie
       - George Bacovia – poezii; Lucian Blaga – Meșterul Manole
       - Dimitrie Cantemir – Descrierea Moldovei (fragmente)
       - I.L. Caragiale – În vreme de război
       - Miron Costin – Letopisețul Țării Moldovei (fragmente)
       - George Coșbuc – poezii; Octavian Goga – poezii
       - Dacia literară (fragmente programatice — Kogălniceanu)
       - Mircea Eliade – Nuntă în cer
       - Mihai Eminescu – poezii (aprofundat)
       - Ion Neculce – O samă de cuvinte
       - Costache Negruzzi – Alexandru Lăpușneanul
       - Camil Petrescu – Patul lui Procust, Jocul ielelor
       - Liviu Rebreanu – Pădurea spânzuraților, Ciuleandra
       - Ioan Slavici – Moara cu noroc
       - Grigore Ureche – Letopisețul Țării Moldovei (fragmente)

       LIMBĂ (clasa XI):
       - Istoria limbii române: etape cronologice, documente vechi
       - Stilistică: figuri de stil aprofundate, registre de limbă
       - Tipuri de discurs: narativ, descriptiv, argumentativ, expozitiv
       - Pregătire: eseu structurat, rezumat, comentariu literar

       CE PREDĂ PROFESORUL LA XI:
       - Analize literare complexe: relația autor-narator-personaj, simboluri, viziunea autorului
       - Plasarea autorilor pe axa timpului + curentul literar aferent
       - Eseu interpretativ pe text literar — schemă apropiată de subiectele de BAC
       - Prezentarea comparativă a două curente/epoci sau doi autori

       ══════════════════════════════════════════
       CLASA A XII-A — Sinteză și pregătire BAC
       ══════════════════════════════════════════

       LITERATURĂ — recapitulare sistematică:
       Toți autorii canonici: Eminescu, Creangă, Caragiale, Sadoveanu, Slavici, Rebreanu,
       Camil Petrescu, Arghezi, Bacovia, Blaga, Barbu, Nichita Stănescu, Marin Preda,
       G. Călinescu, Marin Sorescu

       OPERE STUDIATE (clasa XII):
       - Tudor Arghezi – poezii (Testament, Flori de mucigai)
       - George Bacovia – poezii (Plumb, Lacustră)
       - Ion Barbu – poezii (Riga Crypto, Joc secund)
       - Lucian Blaga – poezii (Eu nu strivesc corola...)
       - I.L. Caragiale – O scrisoare pierdută
       - George Călinescu – Enigma Otiliei
       - Ion Creangă – Povestea lui Harap-Alb
       - Mihai Eminescu – poezii (Luceafărul, Floare albastră, O, mamă...)
       - Ion Pillat – poezii
       - Marin Preda – Moromeții, Cel mai iubit dintre pământeni
       - Liviu Rebreanu – Ion
       - Mihail Sadoveanu – Baltagul, Hanu Ancuței
       - Ioan Slavici – Moara cu noroc
       - Marin Sorescu – Iona, A treia țeapă
       - Nichita Stănescu – poezii

       ÎNCADRARE CURENTĂ LITERARĂ (STRICT pentru BAC):
       - Romantism: Eminescu — geniu/vulg, natură-oglindă, iubire ideală, timp/spațiu cosmic
       - Simbolism: Bacovia — simboluri, muzicalitate, cromatică depresivă, sinestezii
       - Modernism: Blaga, Arghezi, Barbu, Camil Petrescu — inovație formală, intelectualism
       - Tradiționism: Sadoveanu, Rebreanu (parțial) — specific național, rural, autohtonism
       - Realism: Slavici, Rebreanu, Caragiale — veridicitate, tipologie socială, obiectivitate
       - ⚠️ Creangă (Harap-Alb) = Basm Cult cu specific REALIST (oralitate, umanizarea fantasticului)

       STRUCTURA ESEULUI BAC (OBLIGATORIE):
       Introducere: încadrare autor + operă + curent literar + teză
       Cuprins:
         → Argument 1: idee + citat scurt (max 2 rânduri) + analiză
         → Argument 2: idee + citat + analiză
         → Element de structură/compoziție (titlu, incipit, final, laitmotiv etc.)
         → Limbaj artistic: minim 2 figuri de stil identificate și explicate
       Concluzie: reformularea tezei + judecată de valoare
       - La POEZIE: obligatoriu 1 element de prozodie (măsură, rimă, ritm)
       - La PROZĂ: perspectivă narativă, relație narator-personaj, tehnici narative
       - La DRAMĂ: conflict dramatic, didascalii, limbajul personajelor

       GRAMATICĂ — Subiectul I BAC:
       - Analiză morfologică: parte de vorbire + toate categoriile gramaticale relevante
         → Substantiv: gen, număr, caz, articulare
         → Verb: mod, timp, persoană, număr, diateză
         → Adjectiv: grad de comparație, gen, număr, caz
       - Analiză sintactică: parte de propoziție + funcție sintactică
       - Relații sintactice: coordonare (și, dar, sau, ci, deci, însă) / subordonare (că, să, care, când, dacă)
       - Tipuri de subordonate: subiectivă, predicativă, atributivă, completivă directă/indirectă,
         circumstanțială (de loc, timp, mod, cauză, scop, condiție, concesie)

       TIPURI DE SCRIERE exersate la română (IX-XII):
       - Rezumat: redă obiectiv acțiunea, fără opinii, la persoana a III-a
       - Caracterizare de personaj: trăsături fizice + morale, scene relevante, relații cu alte personaje
       - Comentariu literar: analiză pe text dat, figuri de stil, structură, semnificații
       - Eseu argumentativ: teză + 2 argumente + contraargument (opțional) + concluzie
       - Eseu interpretativ (BAC): schema de mai sus — obligatoriu cu citate și analiză

       CE PREDĂ PROFESORUL LA XII:
       - Recapitulări tematice: autor cu autor, curent cu curent, operă cu operă
       - Simulări de subiecte de BAC cu barem explicit și cronometrare
       - Feedback personalizat pe eseuri scrise de elev
       - Exerciții de gramatică tip Subiectul I (morfologie + sintaxă + vocabular)""",
    "limba engleză": r"""
    9. LIMBA ENGLEZĂ — PROFESOR VIRTUAL (programa MEN, liceu România)

       NIVEL ȘI PROFIL:
       - Clasa IX: consolidare A2→B1 | Clasa X: B1 solid | Clasa XI-XII: B1→B2 (și C1 intensiv)
       - Programa L1 (prima limbă) și L2 (a doua limbă) — abordare diferențiată la cerere
       - Proba C BAC: listening, reading, writing, speaking (niveluri A1–B2)

       TERMINOLOGIE GRAMATICALĂ — întotdeauna în română pentru elevii români:
       - Timp verbal (nu "tense"), Mod (nu "mood"), Voce (activă/pasivă)
       - Propoziție principală / subordonată, Complement direct/indirect

       COMPETENȚE DE EXERSAT (conform programei MEN):
       - Receptare orală: dialoguri, anunțuri, interviuri, prezentări
       - Receptare scrisă: articole, e-mailuri, texte funcționale, texte literare simple
       - Producere orală: prezentări, dezbateri, descrieri, povești
       - Producere scrisă: e-mail, eseu argumentativ, CV, scrisoare de intenție, recenzie
       - Mediere: reformularea/traducerea simplă a unui mesaj din/în română

       TEME PRINCIPALE (programa clasa IX, L1 nouă):
       - Teens' culture, Social media & AI, Books & Movies, Community life
       - Greening life, Hidden tourist gems, Personal growth, Exploring passions

       TEME CLASELE X–XII (4 domenii):
       - Personal: relații, sănătate, timp liber, sport, cultură de tineret
       - Public: societate, economie, mass-media, ecologie, democrație, drepturile omului
       - Ocupațional: profesii, CV, scrisoare de intenție, interviu, etică la locul de muncă
       - Educațional/Cultural: literatură în engleză, civilizație britanică/americană, știință

       GRAMATICĂ — ordinea predării (IX→XII):
       IX:
       - Present/Past Simple & Continuous, Future (will / going to / prez. cont.)
       - Articolele a/an/the/zero article — reguli și excepții
       - Substantiv: singular/plural neregulat, genitiv saxon
       - Pronume: personale, posesive, reflexive
       - Verbe modale de bază: can, must, have to, should
       X:
       - Present Perfect Simple & Continuous (for/since, just/already/yet)
       - Past Perfect; narrative tenses (Simple + Continuous + Perfect)
       - Gerunziu vs. infinitiv (enjoy doing vs. want to do)
       - Vocea pasivă (toate timpurile de bază)
       - Reported speech (concordanța timpurilor, pronume, expresii de timp)
       - Condiționali: tip 1 (real), tip 2 (ireal prezent), tip 3 (ireal trecut), mixt
       XI-XII:
       - Pronume relative (who, which, that, whose) — relative clause definitorii/nedef.
       - Inversiune (Had I known…, Should you need…)
       - Wish / If only (prezent, trecut, viitor)
       - Structuri avansate: despite/in spite of, although/even though, unless, provided that

       STRUCTURA ESEU ARGUMENTATIV (BAC / Cambridge style):
       Introducere (teză clară) → Paragraf 1 (argument + exemplu) →
       Paragraf 2 (argument + exemplu) → Concluzie (reformulare + opinie finală)
       - Topic sentence → development → concluding sentence pentru fiecare paragraf
       - Conectori: Furthermore, However, In addition, On the other hand,
         Despite this, As a result, In conclusion, To sum up

       GREȘELI FRECVENTE (corectează blând, cu explicație):
       - "I am agree" → "I agree" (agree nu e adjectiv)
       - "He go" → "He goes" (prezent simplu, pers. a III-a sg)
       - "more better" → "better" (comparativ neregulat)
       - "I have seen him yesterday" → "I saw him yesterday" (past simple cu moment precis)
       - "She is knowing" → "She knows" (stative verbs nu au continuous)
       - "discuss about" → "discuss" (fără prepoziție)
       - Articol zero la substantive generice: "Life is short" nu "The life is short"

       STIL DE PREDARE:
       - Dă exemple în engleză + traducere/explicație în română
       - La exerciții, arată mai întâi modelul rezolvat, apoi lasă elevul să exerseze
       - Corectura: subliniază greșeala → explică regula → oferă varianta corectă
       - Adaptează nivelul: simplu și concret pentru IX, mai abstract și nuanțat pentru XII
""",
    "limba franceză": r"""
    10. LIMBA FRANCEZĂ — PROFESOR VIRTUAL (programa MEN, liceu România)

        NIVEL ȘI PROFIL:
        - L1 (prima limbă): A2→B1 în clasa IX, țintă B1–B2 la final de liceu
        - L2 (a doua limbă — cea mai frecventă): A2 în IX, B1 în X, B1–B2 în XI-XII
        - Adaptează complexitatea la nivelul cerut de elev

        COMPETENȚE DE EXERSAT (conform programei MEN):
        - Receptare orală: dialoguri simple, anunțuri, mesaje audio/video
        - Receptare scrisă: anunțuri, e-mailuri, postări, articole scurte, broșuri
        - Producere orală: răspunsuri, prezentări scurte, dialoguri, dezbateri simple
        - Producere scrisă: e-mailuri, mesaje, descrieri, narațiuni, texte argumentative
        - Interacțiune și mediere: reformulare, traducere simplă română↔franceză

        TEME PRINCIPALE (4 domenii, programa IX–XII):
        - Personal: familie, prieteni, sănătate, alimentație, timp liber, hobby-uri,
                    universul adolescenței, planuri de viitor
        - Public: orașul/satul, regiuni francofone, călătorii, servicii publice,
                  mass-media, mediu, societate, drepturile omului
        - Ocupațional: meserii, locul de muncă, CV, scrisoare de intenție, interviu,
                       relația angajat-angajator, etică profesională
        - Educațional/Cultural: școala, personalități culturale/științifice/sportive,
                                civilizație franceză și francofonă, literatură simplă

        TEME SPECIALE CLASELE XI–XII:
        - Societate și cetățenie (fake news, globalizare, migrație, diversitate culturală)
        - Cultură și civilizație (scriitori, artiști, evenimente istorice franceze)
        - Texte argumentative complexe (avantaje/dezavantaje, eseu de opinie)

        FUNCȚII DE COMUNICARE — ordinea predării:
        IX: prezentare, salut, cerere/dare de informații, descriere persoane/locuri/obiecte,
            povestire scurtă, cerere/dare permisiune, exprimare acord/dezacord
        X: exprimarea opiniei și susținerea ei, invitație/propunere/acceptare/refuz,
            exprimarea obligației/dorinței/preferinței, indicații de traseu,
            relatare evenimente trecute, formulare ipoteze, exprimarea intenției
        XI-XII: argumentare, prezentare de proiect, dezbatere, mediere/rezumat de text

        GRAMATICĂ — ordinea predării (IX→XII):
        IX:
        - Articol: hotărât (le/la/les), nehotărât (un/une/des), partitiv (du/de la/des)
        - Substantiv + acord adjectiv (gen, număr)
        - Pronume personale subiect, complement (COD/COI), en/y
        - Verbe: présent indicatif (regulate + être, avoir, aller, faire, prendre, venir)
        - Imperativ simplu pentru instrucțiuni
        - Negație: ne...pas, ne...jamais, ne...plus, ne...rien
        - Interogație: Est-ce que…? / inversiune / intonație

        X:
        - Passé composé (avoir/être + participiu trecut) + acord participiu
          → Verbe cu être: aller, venir, partir, arriver, naître, mourir, rester + reflexive
        - Imparfait: formare + utilizare (descrieri, acțiuni repetate în trecut)
        - Passé composé vs. Imparfait — distincție și utilizare împreună în povestire
        - Futur simple: formare + utilizare
        - Condițional prezent: politețe, dorință, ipoteză (Je voudrais…, Si j'avais…)
        - Pronume relative: qui, que, dont, où
        - Adjective: comparativ și superlativ (plus…que, moins…que, le plus…)

        XI-XII:
        - Subjonctif prezent: formare + structuri (il faut que, vouloir que, bien que,
          pour que, à condition que)
        - Condițional trecut: ipoteze ireale despre trecut (Si j'avais su…)
        - Vocea pasivă de bază
        - Propoziții subordonate: cauzale (parce que, puisque), concesive (bien que + subj.),
          condiționale (si + prezent/imperfect/mai-mult-ca-perfect)
        - Discurs indirect (concordanța timpurilor la franceză)

        ACORD PARTICIPIU TRECUT — regulă detaliată (greșeală frecventă!):
        - Cu avoir: acord cu COD plasat ÎNAINTEA verbului
          → "La lettre qu'il a écrite" (COD 'que'=lettre, înainte → acord feminin)
          → "Il a écrit des lettres" (COD după verb → fără acord)
        - Cu être: acord cu subiectul (gen + număr)
          → "Elle est partie", "Ils sont partis"
        - Verbe reflexive → întotdeauna cu être

        STRUCTURA ESEU FRANCEZĂ (BAC / examene):
        Introduction (sujet amené → sujet posé → sujet divisé) →
        Développement: thèse (argument 1 + exemplu) + antithèse (argument 2 + exemplu)
                       + synthèse/nuanță →
        Conclusion (bilan + ouverture)
        - Conectori utili: D'abord, Ensuite, De plus, Cependant, En revanche,
          Néanmoins, Par conséquent, En conclusion, En définitive

        GREȘELI FRECVENTE (corectează blând, cu explicație):
        - Acord adjectiv: "une fille grand" → "une fille grande"
        - Auxiliar greșit: "J'ai allé" → "Je suis allé"
        - Acord participiu cu avoir: "Je l'ai vu" (m.) vs "Je l'ai vue" (f.)
        - Negație incompletă: "Je sais pas" → "Je ne sais pas" (registro formal)
        - Confuzie ser/avoir în expresii: "J'ai faim/froid/chaud" (nu "Je suis faim")
        - Subjonctif omis: "Il faut que tu vas" → "Il faut que tu ailles"

        STIL DE PREDARE:
        - Dă exemple în franceză + traducere/explicație în română
        - Contrastează cu româna sau engleza când ajută înțelegerea
        - La exerciții, arată mai întâi modelul rezolvat, apoi lasă elevul să exerseze
        - Corectura: subliniază greșeala → explică regula → oferă varianta corectă
        - Nivel IX-X: foarte concret, multe exemple, exerciții repetitive
        - Nivel XI-XII: text mai autentic, sarcini de producere liberă, argumentare

""",
    "limba germană": r"""
    11. LIMBA GERMANĂ — PROFESOR VIRTUAL (programa MEN, liceu România)

        NIVEL ȘI PROFIL:
        - L1 (prima limbă): consolidare A1→A2 în IX, țintă B1–B2 la final de liceu
        - L2 (a doua limbă): A2 în IX-X, B1 în XI, B1–B2 în XII
        - L3 (a treia limbă): A1 în IX, A2 în X-XI, B1 în XII
        - Adaptează complexitatea: întreabă elevul clasa și nivelul dacă nu e clar
        - Pregătire BAC: înțelegere text scris, redactare text, probă orală

        COMPETENȚE DE EXERSAT (conform programei MEN + CEFR):
        - Hörverstehen (înțelegere orală): dialoguri, anunțuri, știri scurte audio/video
        - Sprechen (exprimare orală): dialoguri, prezentări, dezbateri ghidate
        - Leseverstehen (înțelegere scrisă): mesaje, e-mailuri, articole, texte funcționale
        - Schreiben (producere scrisă): mesaje, e-mailuri, invitații, CV, eseu scurt
        - Mediation (mediere): rezumarea în română a unui text german și invers,
          traducere funcțională simplă
        - Competență interculturală: comparații România–spațiu germanofon (DE, AT, CH)

        TEME PRINCIPALE — 4 domenii (programa IX–XII):
        - Personal: eu, familia, prietenii, casa, orașul, viața zilnică, sănătate,
                    alimentație, timp liber, hobby-uri, universul adolescenței
        - Public: țări și regiuni germanofone, transport, servicii (magazin, bancă, poștă),
                  mass-media, social media, mediu, societate, cultură europeană
        - Ocupațional: meserii, locul de muncă, CV, scrisoare de intenție, interviu,
                       vocabular profesional (comerț, turism, gastronomie)
        - Educațional/Cultural: viața la liceu, civilizație germanofonă (Germania, Austria,
                                Elveția), personalități, literatură, film, muzică,
                                patrimoniu cultural european

        TEME SPECIALE PE CLASE:
        IX:  Prezentare, familie, școală, casă, oraș, timp liber, sărbători germane
        X:   Rutina zilnică, alimentație, cumpărături, călătorii, media, mediu
        XI:  Relații interpersonale, joburi, CV, teme sociale (migrație, globalizare),
             cultură germanofonă, texte de opinie
        XII: Tehnologie, Europa, Erasmus, texte argumentative, pregătire BAC

        GRAMATICĂ — ordinea predării (IX→XII):
        IX (A1→A2):
        - Genul substantivului (der/die/das) — OBLIGATORIU de memorat cu articol!
          → Sfat: învață întotdeauna substantivul cu articolul (der Tisch, die Lampe, das Buch)
        - Articole hotărâte/nehotărâte la nominativ și acuzativ (ein/eine/ein, kein/keine)
        - Declinarea articolelor la dativ (dem/der/dem, einem/einer/einem)
        - Pronume personale (ich, du, er/sie/es, wir, ihr, sie/Sie)
        - Pronume posesive (mein, dein, sein/ihr, unser, euer, ihr)
        - Verbe la prezent (Präsens): conjugare regulate + neregulate frecvente
          → Neregulate esențiale: sein (bin/bist/ist), haben (habe/hast/hat),
            werden, können, müssen, dürfen, wollen, sollen, mögen, fahren, laufen
        - Ordinea cuvintelor: SVO în propoziția enunțiativă, V la sfârșit în subordonată
        - Propoziție interogativă cu verb la loc 1 (Ist das...?) sau W-Fragen (Was, Wo, Wer...)
        - Numerale, ora, zilele săptămânii, lunile, anotimpurile
        - Imperativ simplu (Geh! Komm! Mach!)
        - Prepoziții frecvente cu acuzativ (durch, für, gegen, ohne, um) și dativ (aus, bei,
          mit, nach, seit, von, zu, gegenüber)

        X (A2→B1):
        - Perfekt (trecut compus): haben/sein + Partizip II
          → Partizip II: ge- + stem + -(e)t (regulate) sau forme neregulate (gehen→gegangen)
          → Verbe cu sein: gehen, kommen, fahren, fliegen, bleiben, sein, werden, sterben
        - Präteritum (trecut narativ): forme frecvente war, hatte, ging, kam, machte
          → În vorbire: Perfekt | În scris/narațiune: Präteritum
        - Comparație adjective: gut → besser → am besten (forme neregulate!)
          → Pozitiv, Komparativ, Superlativ: schnell, schneller, am schnellsten
        - Verbe modale complete: können, müssen, dürfen, wollen, sollen, mögen + Konjunktiv II
        - Conectori: und, aber, oder, denn, weil (V la sfârșit!), dass (V la sfârșit!),
          wenn, obwohl, damit, bevor, nachdem
        - Reflexive Verben (sich waschen, sich freuen, sich interessieren für)

        XI (B1):
        - Subordonate cu weil, dass, wenn, obwohl — VERB MEREU LA SFÂRȘIT
          → "Ich lerne Deutsch, weil es interessant ist." (nu "weil es ist interessant")
        - Pronume relative: der/die/das + declinare (dem, denen, dessen, deren)
          → "Das Buch, das ich lese, ist interessant."
        - Konjunktiv II pentru politețe și ipoteză:
          → ich würde + Infinitiv (würde gehen, würde kaufen)
          → Forme speciale: wäre (sein), hätte (haben), könnte, müsste, dürfte
        - Passiv: werden + Partizip II (Das Auto wird repariert.)
        - Plusquamperfekt (mai-mult-ca-perfect): hatte/war + Partizip II
        - Partizipialkonstruktionen de bază

        XII (B1→B2):
        - Konjunktiv I (vorbire indirectă — Indirekte Rede):
          → Er sagt, er sei krank. / Er sagt, er habe keine Zeit.
        - Structuri avansate: zweiteilige Konnektoren (entweder…oder, sowohl…als auch,
          zwar…aber, nicht nur…sondern auch, weder…noch)
        - Nominalizare și stilul academic/formal
        - Infinitivkonstruktionen cu zu (Es ist wichtig, Deutsch zu lernen.)
        - Genitivul în scris formal (wegen des Wetters, trotz des Regens)

        PARTICULARITĂȚI GERMANE — atenție specială la:
        - GENUL SUBSTANTIVELOR: nu există regulă universală — se memorează cu articol
          → Trucuri utile: -ung, -heit, -keit, -schaft → die (feminin!)
                          -chen, -lein → das (neutru!)
                          -er (agent) → der (de regulă masculin)
        - ORDINEA CUVINTELOR (Satzstellung): verbul conjugat MEREU pe poziția 2
          → "Heute gehe ich in die Schule." (nu "Heute ich gehe...")
        - SEPARABLE VERBEN: prefixul separabil merge la SFÂRȘITUL propoziției
          → "Ich rufe dich an." (anrufen → an...rufe)
        - CAZURILE (Kasus): nominativ, acuzativ, dativ, genitiv — afectează articolul!

        STRUCTURA TEXT ARGUMENTATIV (BAC germană):
        Einleitung (introducere + teză) →
        Hauptteil: Argument 1 + Beispiel → Argument 2 + Beispiel → Gegenargument + Widerlegung →
        Schluss (concluzie + opinie personală)
        - Conectori pentru eseu: Zunächst, Außerdem, Darüber hinaus, Allerdings,
          Dennoch, Obwohl, Deshalb, Zusammenfassend, Meiner Meinung nach

        GREȘELI FRECVENTE (corectează blând, cu explicație):
        - Gen greșit: "der Lampe" → "die Lampe" (atenție la gen!)
        - Verb la poziția greșită: "Heute ich gehe" → "Heute gehe ich" (V pe poz. 2!)
        - Verb la sfârșit omis în subordonată: "weil er ist krank" → "weil er krank ist"
        - Prefix separabil uitat: "Ich rufe dich" → "Ich rufe dich an" (anrufen!)
        - Auxiliar greșit la Perfekt: "Ich habe gegangen" → "Ich bin gegangen"
        - Acuzativ vs dativ confundat: "Ich gehe in dem Park" → "Ich gehe in den Park"
          (mișcare → acuzativ; locație → dativ)
        - Participiu II format greșit: "gegehnt" → "gegangen" (neregulat!)

        STIL DE PREDARE:
        - Explică ÎNTOTDEAUNA genul substantivelor noi: der/die/das + substantiv
        - Contrastează cu română și engleză (ajută mult pentru structura frazei)
        - Ordinea cuvintelor: desenează schema vizual dacă e necesar
          → [Poziția 1] [VERB] [Subiect dacă nu e pe poz.1] [...] [Verb2/Prefix la final]
        - La exerciții: model rezolvat → elev exersează → corecție cu explicație
        - Nivel IX: accent pe vocabular + prezent + câteva verbe neregulate esențiale
        - Nivel X: Perfekt vs Präteritum + modale + conectori
        - Nivel XI-XII: subordonate, Konjunktiv II, texte autentice, producere liberă
""",
}

_PROMPT_ALL_SUBJECTS = "\n    GHID DE COMPORTAMENT:\n" + "".join(_PROMPT_SUBJECTS.values())


def get_system_prompt(materie: str | None = None, pas_cu_pas: bool = False,
                      mod_strategie: bool = False, mod_bac_intensiv: bool = False, mod_avansat: bool = False) -> str:
    """Returnează System Prompt adaptat materiei selectate și modurilor active.
    
    OPTIMIZARE TOKEN: când materia e selectată explicit, include DOAR blocul acelei materii
    (economie 71-94% din tokenii de system prompt față de versiunea completă).
    Când materia e None (Toate materiile), include toate blocurile — comportament original.
    """

    if materie == "pedagogie":
        # Mod pedagogie: trimitem doar _PROMPT_COMUN + _PROMPT_FINAL (fără bloc materie)
        # Economie: ~70-90% din tokenii de system prompt față de versiunea cu materie
        rol_line = (
            "ROL: Ești un profesor și mentor de liceu din România, bărbat, cu experiență în "
            "pregătirea pentru BAC și în strategii de învățare eficientă. "
            "Elevul te întreabă despre cum să învețe mai bine — răspunde ca un mentor experimentat, "
            "concret și personalizat."
        )
    elif materie:
        rol_line = (
            f"ROL: Ești un profesor de liceu din România specializat în {materie.upper()}, "
            f"bărbat, cu experiență în pregătirea pentru BAC. "
            f"Răspunde EXCLUSIV la întrebări legate de {materie}. "
            f"Dacă elevul întreabă despre altă materie, îndrumă-l prietenos să schimbe materia din meniu."
        )
    else:
        rol_line = (
            "ROL: Ești un profesor de liceu din România, universal "
            "(Mate, Fizică, Chimie, Literatură și Gramatică Română, Franceză, Engleză, Germană, "
            "Geografie, Istorie, Informatică, Biologie), bărbat, cu experiență în pregătirea pentru BAC."
        )

    # Bloc suplimentar injectat când modul pas-cu-pas e activ
    pas_cu_pas_bloc = r"""

    ═══════════════════════════════════════════════════
    MOD ACTIV: EXPLICAȚIE PAS CU PAS (PRIORITATE MAXIMĂ)
    ═══════════════════════════════════════════════════
    Elevul a activat modul "Pas cu Pas". Respectă OBLIGATORIU aceste reguli pentru ORICE răspuns:

    FORMAT OBLIGATORIU pentru orice problemă sau explicație:
    **📋 Ce avem:**
    - Listează datele cunoscute din problemă

    **🎯 Ce căutăm:**
    - Spune clar ce trebuie aflat/demonstrat

    **🔢 Rezolvare pas cu pas:**
    **Pasul 1 — [nume pas]:** [acțiune + de ce o facem]
    **Pasul 2 — [nume pas]:** [acțiune + de ce o facem]
    ... (continuă până la final)

    **✅ Răspuns final:** [rezultatul clar, cu unități dacă e cazul]

    **💡 Reține:**
    - 1-2 idei cheie de memorat din acest exercițiu

    REGULI STRICTE în modul pas cu pas:
    1. NICIODATĂ nu sări un pas, chiar dacă pare evident.
    2. La fiecare pas explică DE CE faci acea operație, nu doar CE faci.
       - GREȘIT: "Împărțim la 2."
       - CORECT: "Împărțim la 2 pentru că vrem să izolăm variabila x."
    3. Dacă există mai multe metode, alege cea mai simplă și menționeaz-o.
    4. La final, verifică răspunsul (substituie înapoi sau estimează).
    5. Folosește emoji-uri pentru pași (1️⃣, 2️⃣, 3️⃣) dacă sunt mai mult de 3 pași.
    ═══════════════════════════════════════════════════
""" if pas_cu_pas else ""

    # Bloc mod Strategie
    mod_strategie_bloc = r"""

    ═══════════════════════════════════════════════════
    MOD ACTIV: EXPLICĂ-MI STRATEGIA (PRIORITATE MAXIMĂ)
    ═══════════════════════════════════════════════════
    Elevul vrea să înțeleagă CUM să gândească rezolvarea, nu să primească calculele gata făcute.

    PENTRU ORICE PROBLEMĂ, răspunde OBLIGATORIU în acest format:

    **🧠 Cum recunoști tipul de problemă:**
    - Ce elemente din enunț îți spun că e acest tip de exercițiu
    - Cu ce tip de problemă să nu o confunzi

    **🗺️ Strategia de rezolvare (fără calcule):**
    - Pasul 1: Ce faci primul și DE CE
    - Pasul 2: Unde vrei să ajungi
    - Pasul 3: Ce formulă/metodă folosești și de ce pe aceasta și nu alta

    **⚠️ Capcane frecvente:**
    - Greșelile tipice pe care le fac elevii la acest tip de problemă

    **✏️ Acum încearcă tu:**
    - Ghidează elevul să aplice strategia, nu îi da răspunsul direct

    REGULI STRICTE:
    1. NU calcula nimic — explică doar logica și gândirea
    2. Dacă elevul are lipsuri de teorie pentru a rezolva, explică ÎNTÂI teoria necesară
    3. Folosește analogii și exemple din viața reală pentru a face strategia memorabilă
    ═══════════════════════════════════════════════════
""" if mod_strategie else ""

    # Bloc mod BAC Intensiv
    mod_bac_intensiv_bloc = r"""

    ═══════════════════════════════════════════════════
    MOD ACTIV: PREGĂTIRE BAC INTENSIVĂ (PRIORITATE MAXIMĂ)
    ═══════════════════════════════════════════════════
    Elevul este în clasa a 12-a și se pregătește intens pentru BAC. Adaptează TOATE răspunsurile:

    PRIORITIZARE CONȚINUT:
    1. Focusează-te EXCLUSIV pe ce apare la BAC — nu preda lucruri care nu sunt în programă
    2. La fiecare răspuns, menționează: "Acesta apare frecvent la BAC" sau "Rar la BAC, dar posibil"
    3. Când explici o metodă, precizează dacă e metoda acceptată la BAC sau există variante mai scurte

    FORMAT RĂSPUNS BAC:
    - Structurează exact ca la subiectele de BAC (Subiectul I / II / III)
    - Punctaj estimativ: "Acest tip de problemă valorează ~15 puncte la BAC"
    - Timp estimativ: "La BAC ai ~8 minute pentru acest tip"

    TEORIA LIPSĂ — DETECTARE AUTOMATĂ (CRITIC):
    Dacă observi că elevul nu are baza teoretică pentru a rezolva problema:
    1. OPREȘTE-TE din rezolvare
    2. Spune explicit: "⚠️ Înainte să rezolvăm, trebuie să știi teoria din spate:"
    3. Explică teoria necesară SCURT și CLAR (definiție + formulă + exemplu simplu)
    4. Abia apoi continuă cu rezolvarea problemei originale

    SFATURI BAC specifice:
    - Reamintește elevului să verifice răspunsul când mai are timp
    - Semnalează când o problemă are "capcane" tipice de BAC
    - La Română: reamintește structura eseului și punctajul pe competențe
    ═══════════════════════════════════════════════════
""" if mod_bac_intensiv else r"""

    TEORIA LIPSĂ — DETECTARE AUTOMATĂ:
    Dacă observi că elevul nu are baza teoretică pentru a rezolva problema:
    1. OPREȘTE-TE și spune: "⚠️ Pentru asta trebuie să știi mai întâi:"
    2. Explică teoria necesară pe scurt (definiție + formulă + exemplu)
    3. Apoi continuă cu rezolvarea
"""

    mod_avansat_bloc = r"""

    ═══════════════════════════════════════════════════
    MOD ACTIV: AVANSAT (PRIORITATE MAXIMĂ)
    ═══════════════════════════════════════════════════
    Elevul știe deja bazele și NU vrea explicații de la zero.

    REGULI STRICTE în Mod Avansat:
    1. NU explica concepte de bază — presupune că le știe
    2. Mergi DIRECT la ideea cheie, metoda sau formula relevantă
    3. Răspuns scurt și dens: maxim 3-5 rânduri pentru o problemă tipică
    4. Format preferat:
       💡 **Ideea:** [ce metodă/formulă se aplică și de ce]
       ⚡ **Calcul rapid:** [doar pașii esențiali, fără explicații evidente]
       ✅ **Rezultat:** [răspunsul final]
    5. Dacă elevul greșește abordarea, corectează DIRECT: "Nu, aplică X în loc de Y."
    6. Folosește notații scurte și simboluri matematice, nu propoziții lungi
    ═══════════════════════════════════════════════════
""" if mod_avansat else ""

    # ── Selectează blocul de materie ──
    if materie == "pedagogie":
        # Mod pedagogie: fără bloc de materie — _PROMPT_COMUN conține deja tot ce trebuie
        ghid_materie = ""
    elif materie and materie in _PROMPT_SUBJECTS:
        # OPTIMIZARE: doar blocul materiei selectate
        ghid_materie = "\n    GHID DE COMPORTAMENT:\n" + _PROMPT_SUBJECTS[materie]
    else:
        # Toate materiile (sau materie necunoscută) — comportament original
        ghid_materie = _PROMPT_ALL_SUBJECTS

    return ("ROL: " + rol_line
            + pas_cu_pas_bloc
            + mod_strategie_bloc
            + mod_bac_intensiv_bloc
            + mod_avansat_bloc
            + _PROMPT_COMUN
            + ghid_materie
            + _PROMPT_FINAL)



# System prompt inițial — ține cont de modul pas cu pas dacă era deja setat
SYSTEM_PROMPT = get_system_prompt(
    materie=None,
    pas_cu_pas=st.session_state.get("pas_cu_pas", False),
    mod_avansat=st.session_state.get("mod_avansat", False),
    mod_strategie=st.session_state.get("mod_strategie", False),
    mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
)


# === DETECȚIE AUTOMATĂ MATERIE ===
# Mapare cuvinte cheie → materie (pentru detecție rapidă fără apel API)
SUBJECT_KEYWORDS = {
    "matematică": [
        "ecuație", "ecuatia", "funcție", "functie", "derivată", "derivata", "integrală", "integrala",
        "limită", "limita", "matrice", "determinant", "trigonometrie", "geometrie", "algebră", "algebra",
        "logaritm", "radical", "inecuație", "inecuatia", "probabilitate", "combinatorică",
        "vector", "plan", "dreapta", "paralelă", "perpendiculară", "triunghi", "cerc", "parabola",
        "matematica", "mate", "math", "calcul", "număr", "numărul", "numere",
    ],
    "fizică_real": [
        "forță", "forta", "viteză", "viteza", "accelerație", "acceleratie", "masă", "masa",
        "energie", "putere", "curent electric", "tensiune electrică", "rezistență electrică",
        "curent", "tensiune", "rezistenta", "rezistență", "circuit", "circuit electric",
        "circuit serie", "circuit paralel", "serie", "paralel",
        "câmp", "camp", "undă", "unda", "optică", "optica", "lentilă", "lentila",
        "termodinamică", "termodinamica", "gaz", "presiune", "volum", "temperatură", "temperatura",
        "fizica", "fizică", "mecanică", "mecanica", "electricitate", "baterie", "condensator",
        "gravitație", "gravitatie", "frecare", "pendul", "oscilatie", "oscilație",
        "rezistor", "ohm", "amper", "volt", "watt", "joule", "newton",
        "nod", "ramură", "legea lui kirchhoff", "legea lui ohm",
    ],
    "fizică_tehnologic": [
        "forță", "forta", "viteză", "viteza", "accelerație", "acceleratie", "masă", "masa",
        "energie", "putere", "curent electric", "tensiune electrică", "rezistență electrică",
        "curent", "tensiune", "rezistenta", "rezistență", "circuit", "circuit electric",
        "circuit serie", "circuit paralel", "serie", "paralel",
        "câmp", "camp", "undă", "unda", "optică", "optica", "lentilă", "lentila",
        "termodinamică", "termodinamica", "gaz", "presiune", "volum", "temperatură", "temperatura",
        "fizica", "fizică", "mecanică", "mecanica", "electricitate", "baterie", "condensator",
        "gravitație", "gravitatie", "frecare", "pendul", "oscilatie", "oscilație",
        "rezistor", "ohm", "amper", "volt", "watt", "joule", "newton",
        "nod", "ramură", "legea lui kirchhoff", "legea lui ohm",
    ],
    "chimie": [
        "atom", "moleculă", "molecula", "element chimic", "compus chimic",
        "reacție chimică", "reactie chimica", "ecuație chimică",
        "acid", "sare", "oxidare", "reducere", "electroliză", "electroliza",
        "număr de moli", "masă molară", "stoechiometrie",
        "organic", "alcan", "alchenă", "alchena", "alcool", "ester", "chimica", "chimie",
        "ph", "soluție", "solutie", "concentratie", "concentrație",
        "hidrogen", "oxigen", "carbon", "azot", "legătură chimică",
    ],
    "biologie": [
        "celulă", "celula", "adn", "arn", "proteină", "proteina", "enzimă", "enzima",
        "mitoză", "mitoza", "meioză", "meioza", "genetică", "genetica", "cromozom",
        "fotosinteza", "fotosinteză", "respiratie", "respirație", "metabolism",
        "ecosistem", "specie", "organ", "tesut", "țesut", "sistem nervos",
        "biologie", "biologic", "planta", "plantă", "animal",
    ],
    "informatică": [
        # general
        "algoritm", "cod", "program", "informatica", "informatică", "programare",
        # Python keywords
        "python", "def ", "list", "dict", "tuple", "set(", "append", "pandas", "numpy",
        "matplotlib", "scikit", "sklearn", "dataframe", "tkinter", "sqlite", "flask",
        # C++ keywords
        "c++", "cout", "cin", "#include", "vector<", "struct ", "pointer", "new ",
        # structuri de date
        "functie", "funcție", "vector", "array", "stivă", "stiva", "coada", "coadă",
        "lista inlantuita", "listă înlănțuită", "arbore", "graf", "heap",
        # algoritmi
        "backtracking", "greedy", "recursivitate", "recursiv", "sortare", "cautare",
        "bubble sort", "merge sort", "quicksort", "dijkstra", "bfs", "dfs",
        "programare dinamica", "programare dinamică", "rucsac", "backtrack",
        "complexitate", "recursie",
        # BD si SQL
        "sql", "baza de date", "bază de date", "select ", "join", "create table",
        "entitate", "normalizare", "sqlite", "mysql",
        # ML
        "machine learning", "invatare automata", "învățare automată", "knn",
        "clustering", "kmeans", "regresie", "clasificare", "neural", "scikit",
        # pseudocod
        "pseudocod", "variabila", "variabilă", "ciclu", "for ", "while ", "if ",
    ],
    "geografie": [
        "relief", "munte", "câmpie", "campie", "râu", "rau", "dunărea", "dunarea",
        "climă", "clima", "vegetatie", "vegetație", "populație", "populatie",
        "romania", "românia", "europa", "continent", "ocean", "geografie",
        "carpati", "carpații", "câmpia", "campia", "delta", "lac",
    ],
    "istorie": [
        "război", "razboi", "revoluție", "revolutie", "unire", "independenta", "independență",
        "cuza", "eminescu", "mihai viteazul", "stefan cel mare", "ștefan cel mare",
        "comunism", "comunist", "ceausescu", "ceaușescu", "bac 1918", "marea unire",
        "medieval", "evul mediu", "modern", "contemporan", "istorie", "istoric",
        "domnie", "domitor", "rege", "regat", "principat",
    ],
    "limba și literatura română": [
        "roman", "roman", "poezie", "poem", "eminescu", "rebreanu", "sadoveanu",
        "preda", "arghezi", "blaga", "bacovia", "caragiale", "creanga", "creangă",
        "eseu", "comentariu", "caracterizare", "narator", "personaj", "tema",
        "figuri de stil", "metafora", "metaforă", "epitet", "comparatie", "comparație",
        "roman", "proza", "proză", "dramaturgie", "gramatica", "gramatică",
        "romana", "română", "literatura", "literatură",
    ],
    "limba engleză": [
        # Identificatori de limbă / materie
        "english", "engleză", "engleza", "grammar", "essay", "vocabulary",
        # Structuri gramaticale exclusiv engleze (fraze compuse — fără risc de false positive)
        "present perfect", "past simple", "past tense", "future tense",
        "present tense", "conditional tense", "passive voice", "reported speech",
        "modal verb", "relative clause", "indirect speech",
        # Teme din programa IX (L1 nouă)
        "teens culture", "social media", "influencer", "personal growth",
        "community life", "tourist gems", "greening life",
        # Teme X–XII (4 domenii)
        "cover letter", "job interview", "curriculum vitae",
        "british culture", "american culture", "civilizație britanică",
        # Tipuri de texte / sarcini frecvente
        "formal letter", "informal email", "book review", "film review",
        "argumentative essay", "opinion essay", "for and against",
        # Vocabular gramatical în română, specific englezei
        "gerunziu", "infinitiv", "vocea pasivă", "inversiune",
        "propoziție relativă", "vorbire indirectă", "condiționala de tip",
    ],
    "limba franceză": [
        # Identificatori de limbă / materie
        "français", "franceză", "franceza",
        # Timpuri verbale exclusiv franceze
        "passé composé", "imparfait", "subjonctif", "futur simple",
        "conditionnel", "participe passé", "plus-que-parfait",
        # Verbe auxiliare (formă exclusiv franceză)
        "être", "avoir",
        # Articole și structuri exclusiv franceze
        "article partitif", "article défini", "article indéfini",
        "du ", "de la ", "des ",
        # Teme din programă
        "civilizație franceză", "francofonă", "francofonie", "espace francophone",
        "pays francophones",
        # Funcții de comunicare frecvente în lecții
        "accord du participe", "accord adjectif", "auxiliaire être",
        "auxiliaire avoir", "verbe pronominal", "verbe réfléchi",
        # Vocabular gramatical în română, specific francezei
        "participiu trecut", "acord participiu", "verb reflexiv",
        "propoziție relativă franceză", "subjonctiv francez",
    ],
    "limba germană": [
        # Identificatori de limbă / materie
        "germană", "germana", "deutsch", "německy", "allemand",
        # Terminologie exclusiv germană
        "der ", "die ", "das ", "ein ", "eine ", "kein", "keine",
        "umlaut", "eszett", "ß",
        # Timpuri verbale exclusiv germane
        "perfekt", "präteritum", "plusquamperfekt", "konjunktiv",
        "konjunktiv ii", "futur i", "futur ii",
        # Structuri gramaticale exclusiv germane
        "separable verben", "trennbare verben", "reflexive verben",
        "verb la sfârșitul", "verb la sfarsitul", "satzstellung",
        "partizip ii", "partizip i",
        # Cazuri germane
        "nominativ", "acuzativ", "dativ", "genitiv",
        # Verbe modale germane
        "können", "müssen", "dürfen", "wollen", "sollen", "mögen",
        # Vocabular gramatical în română specific germanei
        "genul substantivului", "articol hotărât german", "declinare germană",
        "propoziție subordonată germană", "prefix separabil",
        # Teme culturale specifice
        "spațiu germanofon", "germanofon", "hörverstehen", "leseverstehen",
        "oktoberfest", "bundesrepublik",
    ],
}


# Cuvinte care sunt exclusive unei materii — boost mare dacă apar
_STRONG_INDICATORS = {
    # IMPORTANT: folosiți doar cuvinte complete sau fraze — NU substring-uri scurte
    # care pot apărea accidental în alte cuvinte (ex: "ion" e în "funcționează").
    "informatică":  ["python", "c++", "def ", "cout", "#include", "algoritm", "recursiv",
                     "backtracking", "pandas", "sklearn", "compilator", "pseudocod"],
    "matematică":   ["ecuație", "inecuație", "derivată", "integrală", "matrice", "determinant",
                     "funcție", "progresie", "logaritm", "trigonometrie"],
    "fizică_real":       ["forță", "viteză", "accelerație", "curent electric", "tensiune electrică",
                         "rezistență electrică", "câmp magnetic", "undă", "frecvență",
                         "energie cinetică", "circuit electric", "circuit serie", "circuit paralel",
                         "lege lui ohm", "legea lui ohm", "condensator", "inductor",
                         "câmp electric", "sarcină electrică", "putere electrică"],
    "fizică_tehnologic": ["forță", "viteză", "accelerație", "curent electric", "tensiune electrică",
                         "rezistență electrică", "câmp magnetic", "undă", "frecvență",
                         "energie cinetică", "circuit electric", "circuit serie", "circuit paralel",
                         "lege lui ohm", "legea lui ohm", "condensator", "inductor",
                         "câmp electric", "sarcină electrică", "putere electrică"],
    "chimie":       ["reacție chimică", "ecuație chimică", "mol ", "moli ", "masă molară",
                     "oxidare", "reducere", "electroliză", "hidroliza",
                     "acid tare", "bază tare", "soluție tampon", "concentrație molară",
                     "legătură covalentă", "legătură ionică", "orbital"],
    "biologie":     ["celulă", "adn", "arn", "proteină", "metabolism", "fotosinteză",
                     "ecosistem", "evoluție", "genetică", "cromozom", "mitoză"],
    "istorie":      ["război mondial", "tratat de pace", "revoluție", "regat", "imperiu",
                     "dinastie", "domnie", "bătălie"],
    "geografie":    ["relief", "climă", "populație", "hidrografie", "câmpie", "munte",
                     "râu", "bazin hidrografic"],
    "limba și literatura română": ["figuri de stil", "narator", "personaj principal",
                     "comentariu literar", "caracterizare", "metaforă", "epitet",
                     "curent literar", "roman realist"],
    # Indicatori puternici pentru limbi străine — fraze exclusiv din terminologia
    # gramaticală a limbii respective, imposibil de confundat cu româna sau altă materie
    "limba engleză": [
        # Timpuri verbale în engleză (formă exclusiv engleză)
        "present perfect", "past simple", "past tense", "future tense",
        "present continuous", "past continuous", "past perfect",
        # Structuri gramaticale exclusiv engleze
        "passive voice", "reported speech", "modal verb", "relative clause",
        "conditional sentence", "indirect speech", "gerund", "infinitive",
        # Tipuri de texte / sarcini BAC engleză
        "argumentative essay", "opinion essay", "formal letter", "book review",
        "for and against essay",
        # Teme specifice programei noi clasa IX
        "teens culture", "greening life", "personal growth",
    ],
    "limba franceză": [
        # Timpuri verbale în franceză (formă exclusiv franceză)
        "passé composé", "imparfait", "subjonctif", "futur simple",
        "conditionnel présent", "conditionnel passé", "plus-que-parfait",
        # Structuri gramaticale exclusiv franceze
        "participe passé", "être ou avoir", "accord du participe",
        "article partitif", "verbe pronominal", "pronom relatif",
        # Teme specifice programei franceze
        "espace francophone", "pays francophones", "civilisation française",
        # Conectori/structuri de eseu francez
        "thèse antithèse", "plan dialectique",
    ],
    "limba germană": [
        # Timpuri verbale exclusiv germane
        "perfekt", "präteritum", "plusquamperfekt",
        "konjunktiv ii", "konjunktiv i",
        # Structuri gramaticale exclusiv germane — imposibil de confundat
        "separable verben", "trennbare verben", "partizip ii",
        "verb la sfârșitul propoziției", "satzstellung",
        # Cazuri germane (formă exclusiv germană)
        "der den dem des", "akkusativ", "nominativ kasus",
        # Verbe auxiliare în contexte germane
        "haben oder sein", "sein oder haben",
        # Conectori cu verb la sfârșit (exclusiv germani)
        "weil verb", "obwohl verb", "damit verb",
        # Teme culturale specifice germanei
        "spațiu germanofon", "germanofonă", "hörverstehen", "leseverstehen",
        "bundesrepublik", "österreich deutsch",
    ],
}

def detect_subject_from_text(text: str) -> str | None:
    """Detectează materia dintr-un text folosind cuvinte cheie cu sistem de ponderi.
    
    Folosește indicatori puternici (boost x3) + indicatori generali + penalizări încrucișate.
    Evită false positive-uri de tip 'matrice' → matematică când e informatică.

    Returnează:
      - str: materia detectată (ex: "matematică", "fizică_real")
      - "_fizica_ambigua": dacă textul e clar fizică dar profilul e necunoscut
      - None: dacă nu s-a putut detecta nimic
    """
    text_lower = text.lower()
    scores = {}

    # Scor de bază din cuvintele cheie generale
    for subject, keywords in SUBJECT_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw in text_lower)
        scores[subject] = score

    # Boost x3 pentru indicatori puternici (specific unui singur domeniu)
    for subject, indicators in _STRONG_INDICATORS.items():
        strong_hits = sum(1 for ind in indicators if ind in text_lower)
        scores[subject] = scores.get(subject, 0) + strong_hits * 3

    # Penalizare încrucișată: dacă avem indicatori puternici de informatică,
    # penalizăm matematica (ex: "matrice" în context cod → nu matematică)
    info_strong = sum(1 for ind in _STRONG_INDICATORS["informatică"] if ind in text_lower)
    if info_strong >= 2:
        scores["matematică"] = scores.get("matematică", 0) * 0.3

    # Elimină scoruri 0 și returnează maximul cu threshold minim
    scores = {s: v for s, v in scores.items() if v > 0}
    if not scores:
        return None
    best = max(scores, key=scores.get)
    sorted_scores = sorted(scores.values(), reverse=True)

    # Caz special: fizică_real și fizică_tehnologic au indicatori identici →
    # vor avea mereu scor egal. Detectăm că e fizică și returnăm un cod special
    # pentru a declanșa promptul de alegere profil în UI.
    if len(sorted_scores) >= 2 and sorted_scores[0] == sorted_scores[1]:
        # Verificăm dacă cele două cu scor maxim sunt ambele variante de fizică
        top_subjects = [s for s, v in scores.items() if v == sorted_scores[0]]
        if set(top_subjects) == {"fizică_real", "fizică_tehnologic"}:
            return "_fizica_ambigua"
        # Alt egal între materii diferite → nu detectăm
        return None

    return best


def get_detected_subject() -> str | None:
    """Returnează materia detectată din session_state sau None."""
    return st.session_state.get("_detected_subject", None)


def update_system_prompt_for_subject(materie: str | None):
    """Actualizează system prompt-ul pentru materia dată și salvează în session_state.
    Resetează și flag-ul de caching — noul prompt trebuie re-cached la primul apel.
    """
    st.session_state["_detected_subject"] = materie
    # Invalidăm caching-ul local — promptul s-a schimbat, cache-ul vechi nu mai e valid
    st.session_state["_ctx_cache_enabled"] = True   # permite re-caching cu noul prompt
    # FIX: folosim session_state în loc de variabilă globală (care se reseta la fiecare rerun)
    st.session_state["_prompt_cache_store"] = {}  # curăță toate intrările locale
    st.session_state["system_prompt"] = get_system_prompt(
        materie=materie,
        pas_cu_pas=st.session_state.get("pas_cu_pas", False),
        mod_avansat=st.session_state.get("mod_avansat", False),
        mod_strategie=st.session_state.get("mod_strategie", False),
        mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
    )




safety_settings = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
]



# ============================================================
# === SIMULARE BAC ===
# ============================================================

MATERII_BAC = {
    "📐 Matematică M1": {
        "cod": "matematica_m1",
        "profile": ["M1 - Mate-Info"],
        "subiecte": ["Numere complexe", "Funcții", "Ecuații/inecuații", "Probabilități", "Geometrie analitică", "Matrice și sisteme", "Legi de compoziție", "Derivate și monotonie", "Integrale și limite"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "6 exerciții scurte × 5p = 30p",
            "S2": "2 probleme (matrice+sisteme, lege compoziție) = 30p",
            "S3": "2 probleme (funcții+derivate, integrale/limite) = 30p",
        },
    },
    "⚡ Fizică tehnologic": {
        "cod": "fizica_tehnologic",
        "profile": ["Filiera tehnologică"],
        "subiecte": ["Mecanică", "Termodinamică", "Curent continuu", "Optică"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "arii": "4 arii (A-Mecanică, B-Termodinamică, C-Curent continuu, D-Optică)",
            "alegere": "Candidatul alege 2 arii din 4",
            "per_arie": "S.I (5 grilă × 3p) + S.II (problemă 15p) + S.III (problemă 15p)",
        },
    },
    "📖 Română real/tehn": {
        "cod": "romana_real_tehn",
        "profile": ["Real/tehnologic"],
        "subiecte": ["Text la prima vedere", "Comentariu literar", "Eseu personaj/curent"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "50p: A (5 itemi 30p) + B (text argumentativ 150+ cuvinte, 20p)",
            "S2": "10p: comentariu 50+ cuvinte pe fragment literar",
            "S3": "30p: eseu 400+ cuvinte (personaj/text narativ/curent literar)",
        },
    },
    "📐 Matematică M2": {
        "cod": "matematica_m2",
        "profile": ["M2 - Științe ale naturii"],
        "subiecte": ["Funcții", "Ecuații/inecuații", "Probabilități", "Geometrie", "Derivate", "Integrale"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "6 exerciții scurte × 5p = 30p",
            "S2": "2 probleme (matrice+sisteme sau geometrie, funcții) = 30p",
            "S3": "2 probleme (derivate+monotonie, integrale/arii) = 30p",
        },
    },
    "🧪 Chimie": {
        "cod": "chimie",
        "profile": ["Chimie anorganică", "Chimie organică"],
        "subiecte": ["Chimie anorganică", "Chimie organică"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "10 itemi grilă (30p) — noțiuni generale",
            "S2": "Probleme calcul anorganic/organic (30p)",
            "S3": "Probleme aplicative complexe (30p)",
        },
    },
    "🧬 Biologie": {
        "cod": "biologie",
        "profile": ["Biologie vegetală și animală", "Anatomie și fiziologie umană"],
        "subiecte": ["Celulă și țesuturi", "Genetică", "Fiziologie umană", "Ecologie", "Evoluție"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "10 itemi grilă (30p)",
            "S2": "Itemi semiobiectivi și eseu scurt (30p)",
            "S3": "Eseu structurat (30p)",
        },
    },
    "🏛️ Istorie": {
        "cod": "istorie",
        "profile": ["Umanist", "Pedagogic"],
        "subiecte": ["Popoare și spații istorice", "Oameni, societate, lume", "Relații internaționale", "Secolul XX în România"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "Sursă primară — 4 cerințe (30p)",
            "S2": "Eseu scurt factori/cauze/consecințe (30p)",
            "S3": "Eseu structurat 2 pagini cu argumente (30p)",
        },
    },
    "🌍 Geografie": {
        "cod": "geografie",
        "profile": ["Profiluri umaniste"],
        "subiecte": ["Relief", "Climă și hidrografie", "Vegetație și faună", "Populație și așezări", "Economie", "Europa și UE"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "Hartă + 5 cerințe (30p)",
            "S2": "Noțiuni geografice — definiții + exemple (30p)",
            "S3": "Eseu despre o regiune/fenomen geografic (30p)",
        },
    },
    "📖 Română uman/ped": {
        "cod": "romana_uman",
        "profile": ["Umanist", "Pedagogic"],
        "subiecte": ["Text la prima vedere", "Comentariu literar", "Eseu personaj/curent"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "50p: A (5 itemi 30p) + B (text argumentativ 150+ cuvinte, 20p)",
            "S2": "10p: comentariu 50+ cuvinte pe fragment literar",
            "S3": "30p: eseu 400+ cuvinte (personaj/text narativ/curent literar)",
        },
    },
    "💻 Informatică": {
        "cod": "informatica",
        "profile": ["C++", "Pascal"],
        "subiecte": ["Algoritmi", "Structuri de date", "Programare completă"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "Algoritmi și pseudocod (30p)",
            "S2": "Probleme cu tablouri/șiruri (30p)",
            "S3": "Problemă complexă de programare (30p)",
        },
    },
    "🔬 Fizică real": {
        "cod": "fizica_real",
        "profile": ["Matematică-Informatică", "Științe ale naturii"],
        "subiecte": ["Mecanică", "Termodinamică", "Electromagnetism", "Optică", "Fizică modernă"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": True,
        "structura": {
            "S1": "10 itemi grilă × 3p = 30p",
            "S2": "Probleme structurate (mecanică + termodinamică) = 30p",
            "S3": "Problemă complexă (electromagnetism/optică/fizică modernă) = 30p",
        },
    },
    "⚖️ Economie": {
        "cod": "economie",
        "profile": ["Economic", "Științe sociale"],
        "subiecte": ["Piața și mecanismele ei", "Agenții economici", "Macroeconomie", "Economie mondială"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": False,
    },
    "🧠 Psihologie": {
        "cod": "psihologie",
        "profile": ["Științe sociale", "Pedagogic"],
        "subiecte": ["Procesele psihice", "Personalitatea", "Psihologia socială", "Sănătate mentală"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": False,
    },
    "🧠 Logică și argumentare": {
        "cod": "logica",
        "profile": ["Filologie", "Științe sociale"],
        "subiecte": ["Propoziții și argumente", "Inferențe logice", "Argumentare și retorică", "Sofisme"],
        "timp_minute": 180,
        "punctaj_total": 100,
        "date_reale": False,
    },
}

# ── Structura oficială BAC România — filiere, profiluri, specializări ──────────
# Conform OMENCS 4923/2013 și modificările ulterioare
PROFILE_BAC = {
    "🎓 Filiera teoretică": {
        "📐 Profil real": {
            "Matematică-Informatică": {
                "materii_obligatorii": ["📐 Matematică M1", "📖 Română real/tehn"],
                "materii_optionale": ["💻 Informatică", "🔬 Fizică real", "🧪 Chimie", "🧬 Biologie"],
                "descriere": "Matematică M1 + o disciplină la alegere din arie"
            },
            "Științe ale naturii": {
                "materii_obligatorii": ["📐 Matematică M2", "📖 Română real/tehn"],
                "materii_optionale": ["🔬 Fizică real", "🧪 Chimie", "🧬 Biologie"],
                "descriere": "Matematică M2 + Fizică sau Chimie sau Biologie"
            },
        },
        "📚 Profil umanist": {
            "Filologie": {
                "materii_obligatorii": ["📖 Română uman/ped", "🏛️ Istorie"],
                "materii_optionale": ["🌍 Geografie", "🧠 Logică și argumentare", "🌐 Limbă modernă avansată"],
                "descriere": "Română + Istorie + o disciplină la alegere"
            },
            "Științe sociale": {
                "materii_obligatorii": ["📖 Română uman/ped", "🏛️ Istorie"],
                "materii_optionale": ["🌍 Geografie", "⚖️ Economie", "🧠 Psihologie", "🏛️ Sociologie"],
                "descriere": "Română + Istorie + Geografie sau Economie"
            },
        },
    },
    "🛠️ Filiera tehnologică": {
        "⚙️ Profil tehnic": {
            "Tehnic": {
                "materii_obligatorii": ["📖 Română real/tehn", "⚡ Fizică tehnologic"],
                "materii_optionale": ["📐 Matematică M2", "🧪 Chimie", "💻 Informatică"],
                "descriere": "Română + Fizică + o disciplină de specialitate"
            },
            "Resurse naturale și protecția mediului": {
                "materii_obligatorii": ["📖 Română real/tehn", "🧪 Chimie"],
                "materii_optionale": ["📐 Matematică M2", "🧬 Biologie", "⚡ Fizică tehnologic"],
                "descriere": "Română + Chimie sau Biologie"
            },
        },
        "🍽️ Profil servicii": {
            "Economic": {
                "materii_obligatorii": ["📖 Română real/tehn", "⚖️ Economie"],
                "materii_optionale": ["📐 Matematică M2", "🌍 Geografie", "🧠 Psihologie"],
                "descriere": "Română + Economie + o disciplină la alegere"
            },
            "Turism și alimentație": {
                "materii_obligatorii": ["📖 Română real/tehn", "🌍 Geografie"],
                "materii_optionale": ["📐 Matematică M2", "⚖️ Economie", "🧪 Chimie"],
                "descriere": "Română + Geografie sau altă disciplină de profil"
            },
            "Estetica și igiena corpului omenesc": {
                "materii_obligatorii": ["📖 Română real/tehn", "🧬 Biologie"],
                "materii_optionale": ["📐 Matematică M2", "🧪 Chimie"],
                "descriere": "Română + Biologie"
            },
        },
    },
    "🎨 Filiera vocatională": {
        "🎭 Profil artistic": {
            "Arte vizuale": {
                "materii_obligatorii": ["📖 Română uman/ped", "🎨 Istoria artelor"],
                "materii_optionale": ["🏛️ Istorie", "🌍 Geografie"],
                "descriere": "Română + Istoria artelor"
            },
            "Muzică": {
                "materii_obligatorii": ["📖 Română uman/ped", "🎵 Teorie-solfegiu-dicteu"],
                "materii_optionale": ["🏛️ Istorie", "🌍 Geografie"],
                "descriere": "Română + disciplină muzicală de specialitate"
            },
            "Coregrafie": {
                "materii_obligatorii": ["📖 Română uman/ped", "🎭 Artele spectacolului"],
                "materii_optionale": ["🏛️ Istorie"],
                "descriere": "Română + disciplină artistică"
            },
        },
        "⛪ Profil teologic": {
            "Teologie ortodoxă": {
                "materii_obligatorii": ["📖 Română uman/ped", "✝️ Religie"],
                "materii_optionale": ["🏛️ Istorie", "🌍 Geografie"],
                "descriere": "Română + Religie"
            },
            "Teologie catolică / Alte confesiuni": {
                "materii_obligatorii": ["📖 Română uman/ped", "✝️ Religie"],
                "materii_optionale": ["🏛️ Istorie"],
                "descriere": "Română + Religie"
            },
        },
        "⚽ Profil sportiv": {
            "Educație fizică și sport": {
                "materii_obligatorii": ["📖 Română uman/ped", "🏃 Educație fizică și sport"],
                "materii_optionale": ["🏛️ Istorie", "🌍 Geografie", "🧬 Biologie"],
                "descriere": "Română + Educație fizică + o disciplină la alegere"
            },
        },
        "🎓 Profil pedagogic": {
            "Învățători-educatoare": {
                "materii_obligatorii": ["📖 Română uman/ped", "🏛️ Istorie"],
                "materii_optionale": ["🌍 Geografie", "📐 Matematică M2", "🎵 Muzică"],
                "descriere": "Română + Istorie + o disciplină de specialitate"
            },
            "Bibliotecar-documentarist": {
                "materii_obligatorii": ["📖 Română uman/ped", "🏛️ Istorie"],
                "materii_optionale": ["🌍 Geografie"],
                "descriere": "Română + Istorie"
            },
        },
        "🪖 Profil militar": {
            "Matematică-Informatică (militar)": {
                "materii_obligatorii": ["📐 Matematică M1", "📖 Română real/tehn"],
                "materii_optionale": ["💻 Informatică", "🔬 Fizică real"],
                "descriere": "Profil real cu specific militar"
            },
        },
    },
}

# ── Materii disponibile la simulare (cu suport AI) ──────────────────────────
MATERII_SIMULARE_DISPONIBILE = {
    "📐 Matematică M1", "📐 Matematică M2",
    "📖 Română real/tehn", "📖 Română uman/ped",
    "⚡ Fizică tehnologic", "🔬 Fizică real",
    "🧪 Chimie", "🧬 Biologie",
    "🏛️ Istorie", "🌍 Geografie",
    "💻 Informatică", "⚖️ Economie",
    "🧠 Psihologie", "🧠 Logică și argumentare",
}

# ── Date reale BAC 2021-2025 ────────────────────────────────────────────────
BAC_DATE_REALE = {
    "matematica_m1": {
        "tipare": [
            "Numere complexe: calcul cu z, modul, argument, verificare egalități",
            "Funcții simple f(f(x)), f(x+a), f(f(m))=valoare — verificare proprietăți",
            "Ecuații exponențiale (3ˣ, 2ˣ) sau logaritmice (log₃, log)",
            "Probabilități cu numere naturale de două cifre (cifra zecilor, cifra unităților, multipli)",
            "Geometrie analitică: drepte perpendiculare/paralele, coordonate punct, distanțe",
            "Trigonometrie: triunghi dreptunghic/isoscel, arie, sin2A, cos A, tgB",
            "Matrice 3×3 cu parametru: det(A(a)), inversabilitate, proprietăți det(A·B)",
            "Lege de compoziție x★y: calcule punctuale, element neutru, inegalități, condiții",
            "Funcție cu ln sau eˣ: derivată, monotonie, extreme, ecuație f(x)=0 soluție unică",
            "Integrală definită + limită: calcul ∫, primitive, lim(1/x)∫₀ˣtf(t)dt",
        ],
        "subiecte_reale": [
            {"an": 2021, "s1": "Media aritmetică a=b=2021/2 | f(x)=2x²-3x+1, A(1,m) pe grafic | log₃(x+3)-log₃(x+2)=2 | Mulțime 16 submulțimi | M(3,0)N(8,3)P(6,3): MN⃗+MP⃗=MQ⃗ | sin2A=cosA·sinA → A=π/4", "s2": "A(a) 3×3 cu log a: det=1, inversabilă, det(A(a)·A(a+1)⁻¹)≥8 | x★y=xy+m(x+y)+m², m>0: calcule, 2★1=5→2★5=1, (3-x)★(3-x)=m", "s3": "f(x)=4x²-2x-4lnx: f'=(4x²-4x-1)/x, monotonie, exact 2 soluții f(x)=0 | f(x)=(4x²+1)/(x²+1): ∫₀¹f=11/3, asimptotă oblică, ∫₀¹G(x)dx=π/3+ln2-4/3"},
            {"an": 2022, "s1": "(8-6√6)(6√6+1)=2 | f(x)=x+3m, f(f(m))=2m | 2³ˣ·2²=4·4ˣ | P(cifra zecilor | divizor 6) | y=3x-2, A(a,a) pe dreaptă | Triunghi isoscel AB=10, cosA=0, arie=50", "s2": "A(x) 3×3: det=1, A(x)·A(y)=A(x+y), A(n)²+A(n)³+2A(n)=O | x★y=x²y²-4(x+y)²+1: 1★0=-3, e=0 neutru, x★x=4", "s3": "f(x)=x-ln(x²+x+5): f'=(x²-9x)/(x²+x+5), monotonie, f(x)=m soluție unică | f(x)=x³-3x+9: ∫₅⁹f=0, ∫₀⁴x dx=∫₀⁴(f-x)dx, limₙ Iₙ=0"},
            {"an": 2023, "s1": "z=3+i: (z²-zi)=10 | f(x)=x+5: f(x)²-f(x²)=1 | x³-3x²+2x=2 | P(5n+5 multiplu 10) | A(4,0)B(5,4): dreaptă prin origine paralelă AB | Triunghi isoscel dreptunghic în A, arie=4 → BC=4", "s2": "A(a) 3×3 + sistem: det(A(0))=8, inversabilitate, a=-2: x₀+z₀=2 | x★y=x²+y²-2x²y²: 2★3=18, e=1 neutru, x★(1-x)≤1", "s3": "f(x)=(3lnx+1)/(x-1): f'=(x²+15)/(x-1)², asimptotă oblică, (3lnx+1)/(x-1)≥1 | f(x)=x²+2x+eˣ: integrale, lim(1/x)∫₀ˣtf(t)dt=1"},
            {"an": 2024, "s1": "Progresie aritmetică a₂=14, a₃=18 → a₁=? | f(x)=x+2: f(f(5))=9 | 3ˣ+2ˣ⁺³=2ˣ⁻¹ | Numere impare 2 cifre din {1,2,3,7,9} | A(2,1): 2AB=OA | Triunghi dreptunghic BC=12, BC/AB=2 → arie=18√3", "s2": "A și B(x): det(B)=1, B(x)B(y)=B(x+y)-xyA, B(x)B(1-x)=A | f(X)=X³+2X²+X+a-2: f(1)=4, rădăcini a=2, (x₁-1)(x₂-2)(x₃-3)=4", "s3": "f(x)=x²-2x+2eˣ: f', lim f'/f, imaginea f | f(x)=(4x²+2)/(6x+1): ∫₁²f=12/11"},
            {"an": 2025, "s1": "z₁=1-i, z₂=2+i: 2z₁+iz₂=1 | f(x)=x+3: f(f(a))=9 | 2x²-3x-2=0 | P(număr 2 cifre, divizor 6²) | A(0,1)B(5,0)C(6,3)D(a,b): AC și BD același mijloc | Triunghi dreptunghic AB=2, tgB=√3 → BC=2√10", "s2": "A(x) 3×3: det(A(-1))=8, A(x)A(y)=A(x+y), A(x)²+A(x)³+2A(x)=O | f(X)=X³-3X²-6X+a: f(1)=-3, câit+rest la g=X²+X-3, (x₁+1)(x₂+1)(x₃+1)=1", "s3": "f(x)=x²+lnx+2: f'=2x+1/x, asimptotă oblică, bijectivă | f(x)=(x²+3)/(x+1): ∫₀³f=30, ∫₀¹xf=1-ln2, arie cu g=f(x)/eˣ egală cu (1/2)(e-1)/(e+1)"},
        ],
    },
    "fizica_tehnologic": {
        "tipare": {
            "mecanica": ["Mișcare rectilinie — energie, viteză, forțe pe corp", "Unități de măsură SI pentru mărimi derivate (putere, lucru mecanic, energie)", "Plan înclinat — forță de frecare, unghi, coeficient μ", "Sistem corpuri legate prin fir — tensiune, accelerație, mase", "Corp pe plan înclinat cu forță de tracțiune — lucru mecanic, energie cinetică, viteză"],
            "termodinamica": ["Transformări termodinamice — proprietăți izotermă/izobară/izochoră/adiabatică", "ΔU, căldură schimbată, lucru mecanic — formule și calcule", "Gaz ideal în cilindru cu piston — presiune, volum, temperatură, densitate", "Ciclu termodinamic p-V sau p-T — energie internă, căldură, lucru mecanic total"],
            "curent": ["Putere maximă transferată consumatorului (R=r)", "Rezistența unui conductor — ρ, lungime, secțiune", "Circuit serie-paralel cu sursă — tensiuni, intensități, rezistențe echivalente", "Două consumatoare în paralel — intensitate, energie, putere disipată"],
            "optica": ["Refracție și reflexie — relații unghiuri, indice de refracție n·sin·i=sin·r", "Lentilă convergentă — mărire, distanțe, focală, construcție imagine", "Efect fotoelectric — energia fotonului, frecvența prag, energia cinetică", "Lamă cu fețe plane paralele — drum optic, unghi refracție, viteza luminii"],
        },
    },
    "romana_real_tehn": {
        "tipare_s1_itemi": [
            "1. Indică sensul din text al cuvântului X și al secvenței Y",
            "2. Menționează o caracteristică/profesie/statut al personajului X, valorificând textul",
            "3. Precizează momentul/reacția/trăsătura morală + justifică cu o secvență din text",
            "4. Explică motivul pentru care... / reprezintă un eveniment / are loc situația X",
            "5. Prezintă în 30-50 cuvinte atmosfera/atitudinea/o situație conform textului",
        ],
        "teme_argumentativ": [
            "importanța studiului / lecturii / educației",
            "influența profesorilor / mentorilor asupra elevilor",
            "rolul culturii în formarea personalității",
            "comportamentul social / responsabilitatea individuală",
            "influența înfățișării/imaginii asupra succesului personal",
            "importanța relațiilor umane / familiei / prieteniei",
        ],
        "tipare_s2": [
            "Prezintă, în minimum 50 de cuvinte, perspectiva narativă din fragmentul de mai jos.",
            "Prezintă, în minimum 50 de cuvinte, rolul notațiilor autorului în fragmentul de mai jos.",
            "Comentează, în minimum 50 de cuvinte, relația dintre ideea poetică și mijloacele artistice în textul dat.",
        ],
        "repere_s3": [
            "1. Prezentarea statutului social, psihologic, moral al personajului ales",
            "2. Evidențierea unei trăsături prin două episoade sau secvențe comentate",
            "3. Analiza a două elemente de structură/compoziție/limbaj (acțiune, conflict, tehnici narative, modalități de caracterizare, registre stilistice)",
        ],
        "autori_opere": [
            "Ion Creangă — Ion / Harap-Alb / Amintiri din copilărie",
            "Ioan Slavici — Moara cu noroc / Mara / Popa Tanda",
            "Liviu Rebreanu — Ion / Pădurea spânzuraților",
            "Camil Petrescu — Ultima noapte de dragoste / Patul lui Procust",
            "G. Călinescu — Enigma Otiliei",
            "G.M. Zamfirescu — Domnișoara Nastasia / Maidanul cu dragoste",
            "Mihail Sadoveanu — Baltagul / Frații Jderi",
        ],
        "subiecte_reale": [
            {"an": 2022, "s1_text": "Text despre critici literari (Basil Munteanu & Vladimir Streinu)", "s1_B": "text argumentativ 150-200 cuvinte (succes, cultură etc.)", "s2": "Prezentarea rolului notațiilor autorului în fragment dramatic (50+ cuvinte)", "s3": "Eseu personaj dintr-un basm cult (ex. Harap-Alb): statut + trăsătură prin 2 episoade + 2 elemente structură/limbaj"},
            {"an": 2023, "s1_text": "Text la prima vedere — 5 întrebări standard", "s1_B": "text argumentativ 150+ cuvinte", "s2": "Comentariu relație idee poetică — mijloace artistice (50+ cuvinte)", "s3": "Eseu 400+ cuvinte personaj dintr-o nuvelă/roman din literatura română"},
            {"an": 2024, "s1_text": "Fragment memorialistic despre Sadoveanu", "s1_B": "text argumentativ despre cultură/comportament social (150+ cuvinte)", "s2": "Prezintă în min. 50 cuvinte perspectiva narativă din fragmentul dat", "s3": "Eseu 400+ cuvinte personaj dintr-un text dramatic/narativ studiat"},
            {"an": 2025, "s1_text": "Fragment despre profesorul Vasile Pârvan (Grigore Băjenaru, 'Părintele Geticei')", "s1_itemi": "1.sensul 'prielnic'+'pe timpuri' | 2.caracteristică profesori cu săli pline | 3.momentul cursului Pârvan+secvență | 4.motivul referinței la originea numelui | 5.atmosfera sălii Odobescu în 30-50 cuvinte", "s1_B": "Argumentează dacă înfățișarea poate influența succesul, cu referire la text și experiență personală/culturală (150+ cuvinte)", "s2": "Rolul notațiilor autorului în fragmentul din 'Domnișoara Nastasia' de G.M. Zamfirescu — scena cu Vulpașin și Nastasia (50+ cuvinte)", "s3": "Eseu min. 400 cuvinte: particularitățile de construcție ale unui personaj dintr-un text narativ studiat de Ion Creangă sau Ioan Slavici. Repere: statut social/psihologic/moral; trăsătură prin 2 episoade; 2 elemente structură/compoziție/limbaj"},
        ],
    },
    "matematica_m2": {
        "tipare": [
            "Funcții: domeniu, monotonie, grafic, f(x)=a soluție unică",
            "Ecuații/inecuații exponențiale sau logaritmice cu parametru",
            "Probabilități: combinatorică, evenimente independente, Bernoulli",
            "Geometrie analitică: drepte, distanțe, cercuri, tangente",
            "Derivate: calcul f'(x), extreme, monotonie, tangentă la grafic",
            "Integrale definite: calcul, arie, volum de rotație",
            "Șiruri: recurență, convergență, limită",
            "Trigonometrie: ecuații, identități, funcții trigonometrice",
        ],
        "subiecte_reale": [
            {"an": 2022, "s1": "Mulțimi și funcții: f(x)=2x-1, f(f(3)) | Ecuație exponențială 4ˣ-5·2ˣ+4=0 | Probabilitate extragere bile colorate | Geometrie: distanță punct-dreaptă | Derivată f(x)=xeˣ: extreme | Integrală ∫₀¹(2x+1)dx", "s2": "Funcție f(x)=x³-3x+2: monotonie, extreme, grafic, inegalitate f(x)≥-2 | Geometrie: cerc, tangente, punct exterior", "s3": "f(x)=ln(x²+1): asimptote, derivată, arie delimitată de axe | Integrală cu parametru, volum rotație"},
            {"an": 2023, "s1": "Progresie geometrică cu parametru | Ecuație logaritmică log₂(x+1)+log₂(x+3)=3 | Combinatorică: permutări cu restricții | Dreaptă prin 2 puncte, paralelă | Derivată f(x)=(x²+1)/(x-1) | ∫₁²(3x²-2x)dx=4", "s2": "f(x)=x·eˣ: monotonie, extreme, convexitate, tangentă în origine | Sistem de ecuații, discuție parametru", "s3": "f(x)=x²-2lnx: f', extreme, asimptote, arie între grafic și axe | Integrală improprie, limită șir Iₙ"},
            {"an": 2024, "s1": "Funcție liniară compusă f(f(x))=2x+3 | 9ˣ-4·3ˣ-5=0 | P(exact 3 succese din 5, p=1/2) | Mijlocul segmentului, punct pe dreaptă | f(x)=sin²x+cos²x-2sinx | ∫₀^(π/2)cosxdx=1", "s2": "f(x)=x³-6x²+9x: monotonie, grafic, inegalitate pe interval | Matrice 2×2: determinant, inversabilitate, sistem", "s3": "f(x)=(x+1)·e^(-x): f', extreme, inflexiune, arie delimitată cu axa Ox | Șir recursiv aₙ₊₁=2aₙ-1, limită"},
            {"an": 2025, "s1": "Șir aritmetic a₁=3, r=2: a₁₀+a₁₁ | log₃(2x-1)=log₃(x+2) | P(cel puțin 2 succese din 4, p=1/3) | Distanța de la A(2,3) la dreapta x-y+1=0 | f(x)=x·ln x: f'(e) | ∫₁^e(1/x)dx=1", "s2": "f(x)=x/(x²+1): domeniu, asimptote, monotonie, valori extreme, grafic | Geometrie analitică: parabolă y=x²-2x, tangentă paralelă cu y=x", "s3": "f(x)=2x-e^(2x): f', monotonie, concavitate, limită | ∫₀¹f(x)dx, arie cu axa Ox, volum de rotație"},
        ],
    },
    "biologie": {
        "tipare": {
            "vegetala_animala": [
                "Celula — organite, funcții, tipuri (procariote/eucariote, vegetale/animale)",
                "Diviziunea celulară — mitoza și meioza: faze, importanță, comparație",
                "Genetică mendeliană — mono/dihybridare, dominanță, codominanță, genotip/fenotip",
                "Genetica umană — grupe sanguine AB0, Rh, boli genetice, ereditate X-linked",
                "Evoluție — teorii Darwin/Lamarck, speciație, adaptare, selecție naturală",
                "Ecologie — biocenoza, biotop, lanțuri trofice, ecosisteme, ciclu carbon/azot",
                "Sistematica plantelor — cormofite, angiosperme, caracteristici clase",
                "Sistematica animalelor — nevertebrate, vertebrate: caracteristici, adaptări",
            ],
            "anatomie_fiziologie": [
                "Sistemul nervos — neuronul, sinapsa, SNC+SNP, reflexul, analizatori",
                "Sistemul endocrin — glande, hormoni, mecanisme feedback, afecțiuni",
                "Sistemul circulator — inima (structură, ciclu cardiac), vase sanguine, sânge",
                "Sistemul respirator — plămâni, ventilație pulmonară, schimburi gazoase, CV respiratorie",
                "Sistemul digestiv — enzime, absorbție, organe, reglare neuroumorală",
                "Sistemul excretor — nefronul, filtrare glomerulară, reabsorbție, osmoregulare",
                "Sistemul locomotor — oase, articulații, mușchi, contracție musculară (ATP, actina, miozina)",
                "Reproducerea — aparatul reproductor masculin/feminin, gametogeneză, ciclul menstrual",
            ],
        },
        "subiecte_reale": [
            {"an": 2021, "profil": "Anatomie", "s1": "10 itemi grilă: neuron, sinapsa, reflex | S2: Sistemul nervos — structura neuronului, tipuri de neuroni, arcul reflex, calea motorie | S3: Eseu sistemul endocrin — hipofiza, tiroidă, suprarenale, feedback negativ, diabet zaharat"},
            {"an": 2022, "profil": "Anatomie", "s1": "10 itemi: circulație, sânge, grupe sanguine | S2: Inima — structura (cavități, valve), ciclul cardiac (sistolă/diastolă), EKG, debitul cardiac | S3: Eseu sistemul respirator — ventilație, surfactant, capacitate vitală, insuficiență respiratorie"},
            {"an": 2023, "profil": "Anatomie", "s1": "10 itemi: digestie, absorbție, enzime | S2: Nefronul — filtrare, reabsorbție tubulară, secreție, compoziție urină, insuficiență renală | S3: Eseu reproducere — gametogeneză, ciclu menstrual, fertilizare, FIV"},
            {"an": 2024, "profil": "Anatomie", "s1": "10 itemi: muschi, os, articulatie | S2: Contracția musculară — teoria glisării filamentelor, ATP, Ca²⁺, tipuri de mușchi, oboseala musculară | S3: Eseu sistemul imunitar — imunitate nespecifică/specifică, anticorpi, limfocite T și B, vaccinare"},
            {"an": 2025, "profil": "Anatomie", "s1": "10 itemi: hormoni, glande endocrine | S2: Reglarea glicemiei — insulina vs glucagon, pancreas endocrin, diabet tip 1 vs 2, complicații | S3: Eseu sistem nervos vegetativ — simpatic vs parasimpatic, mediatori chimici, reglarea funcțiilor viscerale"},
            {"an": 2023, "profil": "Vegetala", "s1": "10 itemi: celula vegetala, fotosinteza | S2: Meioza — faze, crossing-over, importanță genetică, comparație cu mitoza | S3: Eseu genetică — legile lui Mendel, dihybridare, grupe sanguine, genotip parental din fenotipuri descendenți"},
            {"an": 2024, "profil": "Vegetala", "s1": "10 itemi: ecosisteme, populatii | S2: Lanțuri trofice — producători/consumatori/descompunători, piramide ecologice, flux energetic, ciclul carbonului | S3: Eseu evoluție — teoria sintetică, mecanisme evolutive, speciație, dovezi paleontologice/moleculare"},
        ],
    },
    "chimie": {
        "tipare": {
            "anorganica": [
                "Structura atomului — configurații electronice, periodicitate, proprietăți periodice",
                "Legătura chimică — ionică, covalentă (polară/nepolară), metalică, Van der Waals",
                "Reacții redox — oxidant/reducător, bilanț electronic, număr de oxidare",
                "Acizi și baze — teoria Arrhenius/Bronsted, pH, soluții tampon, hidroliză",
                "Săruri — solubilitate, produsul solubilității, reacții de precipitare",
                "Electroliză — legi Faraday, celulă electrolitică, aplicații industriale",
                "Cinetica chimică — viteza de reacție, factori, energia de activare, cataliza",
                "Echilibru chimic — legea acțiunii maselor, Kc, Kp, principiul Le Chatelier",
            ],
            "organica": [
                "Hidrocarburi — alcani, alchene, alchine, arene: proprietăți, reacții",
                "Derivați halogenați — reacții de substituție, eliminare, mecanism SN1/SN2",
                "Alcooli și eteri — clasificare, reacții oxidare, deshidratare, esterificare",
                "Aldehide și cetone — reacții de adiție, oxidare, recunoaștere (Tollens, Fehling)",
                "Acizi carboxilici — tărie, reacții, esteri, saponificare",
                "Amine și aminoacizi — bazicitate, proteine, legătura peptidică, structuri proteine",
                "Glucide — mono/di/polizaharide, structură, proprietăți, importanță biologică",
                "Polimeri — polimerizare, copolimeri, cauciuc, mase plastice, fibre sintetice",
            ],
        },
        "subiecte_reale": [
            {"an": 2021, "profil": "Chimie anorganică", "s1": "10 grilă: config. electronică, periodicitate, legătură ionică | S2: Echilibru chimic — Kc pentru N₂+3H₂⇌2NH₃, efect temperatură/presiune, calculul conversiei | S3: Electroliză CuSO₄ — masa depusă la catod, volumul gazului la anod, legea Faraday"},
            {"an": 2022, "profil": "Chimie anorganică", "s1": "10 grilă: redox, nr oxidare, balansare | S2: Acizi și baze — pH soluție HCl 0.1M, tampon CH₃COOH/CH₃COONa, Ka, grad disociere | S3: Coroziunea metalelor — pile galvanice, protecție catodică, reacții electrochimice"},
            {"an": 2023, "profil": "Chimie anorganică", "s1": "10 grilă: cinetica, viteza, cataliza | S2: Solubilitate — Kps AgCl, efect ion comun, precipitare selectivă Ag⁺ și Pb²⁺ | S3: Sinteza amoniacului Haber — echilibru, randament, condiții industriale, calcul moli"},
            {"an": 2024, "profil": "Chimie anorganică", "s1": "10 grilă: structura atomului, orbitaluri | S2: Reacții redox — bilanț electronic KMnO₄+HCl, identificarea oxidantului, masa de Cl₂ | S3: Pile galvanice Daniell Zn/Cu — semireacții, tensiunea electromotoare, masa depusă"},
            {"an": 2021, "profil": "Chimie organică", "s1": "10 grilă: hidrocarburi, izomeri, IUPAC | S2: Alcooli — reacție cu Na, oxidare etanol, deshidratare, formula moleculară din compoziție % | S3: Esteri și saponificare — reacția de esterificare, Ke, saponificarea grăsimilor, masa de NaOH"},
            {"an": 2022, "profil": "Chimie organică", "s1": "10 grilă: aldehide, cetone, recunoaștere | S2: Aminoacizi — reacție cu acizi/baze, legătura peptidică, structura glicilalaninei, proprietăți amfotere | S3: Polimeri — polietilena (addiție), Nylon (condensare), grad polimerizare, masa molară"},
            {"an": 2023, "profil": "Chimie organică", "s1": "10 grilă: acizi carboxilici, reacții | S2: Glucide — glucoza (formula Haworth, reacție Tollens, fermentatie), zaharoza (hidroliză) | S3: Benzena și derivați — nitrare toluen, sulfonare benzen, mecanism SEAr, aplicații industriale"},
            {"an": 2024, "profil": "Chimie organică", "s1": "10 grilă: alchene, polimeri, reactii | S2: Acizi grași și grăsimi — acid stearic vs oleic, esterificare cu glicerol, indice saponificare, hidrogenare | S3: Cauciuc natural și sintetic — izopren, polimerizare, vulcanizare, comparație proprietăți"},
        ],
    },
    "istorie": {
        "tipare": {
            "cerinte_sursa": [
                "1. Numiți o informație din sursa X (răspuns direct din text)",
                "2. Precizați secolul/perioada la care se referă sursa",
                "3. Menționați două acțiuni/măsuri/caracteristici prezentate în surse",
                "4. Prezentați un punct de vedere din sursă și argumentați cu o informație exterioară",
            ],
            "teme_frecvente": [
                "Autonomiile locale și instituțiile centrale medievale (cnezate, voievodate, domnie)",
                "Cruciada a IV-a și consecințele pentru spațiul românesc (1204)",
                "Întemeierea Țării Românești și Moldovei (sec. XIV)",
                "Mircea cel Bătrân, Iancu de Hunedoara, Ștefan cel Mare — relații cu Imperiul Otoman",
                "Revoluția de la 1848 în Principatele Române — programe, actori, consecințe",
                "Unirea Principatelor (1859) — context, rolul lui Cuza, reformele",
                "Primul Război Mondial — România (1916-1918), Marea Unire (1918)",
                "Al Doilea Război Mondial — România: 1939-1944, 23 august 1944",
                "Regimul comunist în România — instaurare, Dej, Ceaușescu, rezistență",
                "Revoluția din 1989 și tranziția democratică",
                "Relații internaționale sec. XX: NATO, ONU, Război Rece, UE",
                "Democrație ateniană vs republica romană — instituții, cetățenie",
            ],
            "structura_eseu": [
                "Introducere: context temporal și spațial (2-3 rânduri)",
                "Argument 1: cauze/premise + exemplu concret din surse sau cunoștințe",
                "Argument 2: desfășurare/actori principali + consecințe imediate",
                "Concluzie: importanța evenimentului în context mai larg (2-3 rânduri)",
            ],
        },
        "subiecte_reale": [
            {"an": 2021, "s1": "Surse despre autonomii locale sec. XIV (Diploma Cavalerilor Ioaniți 1247 + cronici) — 4 cerințe standard | S2: Rolul lui Mircea cel Bătrân în apărarea spațiului românesc față de expansiunea otomană | S3: Eseu: Revoluția de la 1848 în Principatele Române — cauze, desfășurare, actori (Bălcescu, Kogălniceanu), consecințe"},
            {"an": 2022, "s1": "Surse despre Unirea Principatelor (1859) — context, dubla alegere a lui Cuza, reacțiile marilor puteri | S2: România în Primul Război Mondial — intrarea în război (1916), campania militară, Pacea de la Buftea, consecințele Marii Uniri | S3: Eseu: Regimul comunist în România — instaurarea (1947-1948), caracteristici, represiunea politică, Securitatea"},
            {"an": 2023, "s1": "Surse despre Alexandru cel Bun și Iancu de Hunedoara — politica față de Imperiul Otoman, bătălii | S2: Revoluția din 1989 — cauze, desfășurare (16-22 dec.), consecințe, tranziția democratică | S3: Eseu: România în al Doilea Război Mondial — Pactul Ribbentrop-Molotov, cedările teritoriale (1940), intrarea alături de Axă, 23 august 1944"},
            {"an": 2024, "s1": "Surse despre formarea statelor medievale românești (Negru Vodă, Dragoș, Bogdan) | S2: Reformele lui Al. I. Cuza (1859-1866) — secularizarea averilor mănăstirești, reforma agrară, reforma instrucției publice | S3: Eseu: Participarea României la Primul Război Mondial (1916-1918) și consecințele: Marea Unire, tratatele de pace"},
            {"an": 2025, "s1": "Surse despre democrația ateniană (Pericle, Adunarea Poporului) vs instituțiile republicii romane (Senat, consuli) | S2: Nicolae Ceaușescu și regimul național-comunist — cultul personalității, politica externă independentă, criza economică din anii '80 | S3: Eseu: Constituirea României moderne în sec. XIX — Unirea Principatelor, domnia lui Carol I, Independența (1877), Constituția din 1866"},
        ],
    },
    "geografie": {
        "tipare": {
            "harta": [
                "Identificarea pe hartă a formelor de relief (munți, câmpii, dealuri, podișuri)",
                "Recunoașterea râurilor, lacurilor, regiunilor geografice",
                "Localizarea orașelor, județelor, regiunilor de dezvoltare",
                "Citirea legendei hărții și interpretarea simbolurilor",
            ],
            "teme_frecvente": [
                "Relieful României — Carpații (Orientali/Meridionali/Occidentali), Subcarpații, Podișuri, Câmpii, Delta Dunării",
                "Clima României — factori genetici, tipuri climatice, temperatura/precipitații pe regiuni",
                "Hidrografia — bazinele hidrografice, Dunărea, fluvii, lacuri naturale/artificiale, ape subterane",
                "Vegetația și fauna — etajarea vegetației, păduri, pajiști, stepă, zone protejate",
                "Solurile — tipuri (cernoziom, brun, podzol), distribuție, fertilitate",
                "Populația — evoluție, structură (sex, vârstă, etnie), mișcarea naturală și migratorie",
                "Așezările urbane și rurale — rețeaua de orașe, funcții urbane, urbanizare",
                "Agricultura — tipuri de culturi, regiuni agricole, probleme, politica agricolă UE",
                "Industria — ramuri, centre industriale, restructurare post-comunistă",
                "Transporturile — rețele rutiere, feroviare, fluviale, aeriene; coridoare paneuropene",
                "Turismul — resurse, tipuri, stațiuni, circuite, ecoturism",
                "Uniunea Europeană — instituții, extindere, politici, fonduri structurale, Schengen",
                "Europa — regiuni geografice, mari fluvii, caracteristici climatice, populație",
            ],
            "structura_eseu": [
                "Definiție/caracterizare generală a fenomenului/regiunii (1 paragraf)",
                "Localizare și răspândire spațială (cu exemple concrete — județe, regiuni, cifre)",
                "Cauze/factori determinanți (naturali și/sau umani)",
                "Consecințe și importanță economico-socială",
                "Concluzii și perspective (dacă se cere)",
            ],
        },
        "subiecte_reale": [
            {"an": 2021, "s1": "Hartă România fizică — identificare Munții Apuseni, Câmpia Bărăganului, râul Mureș, Lacul Bicaz | S2: Clima României — factori genetici (latitudine, relief, mase de aer), temperatura medie anuală, precipitații (distribuție, maxime/minime), inversii de temperatură | S3: Eseu: Dunărea — izvor, afluenți, sectoare (german, central, românesc), Delta Dunării (formare, Rezervația Biosferei), importanța economică"},
            {"an": 2022, "s1": "Hartă României — județe din Muntenia și Moldova, identificare orașe, drumuri europene | S2: Populația României — evoluție demografică (1900-2021), bilanț natural negativ, bilanț migrator, îmbătrânire demografică, structura pe etnii | S3: Eseu: Agricultura românească — resurse naturale (teren arabil, soluri), principalele culturi (cereale, floarea-soarelui, viță de vie), regiuni agricole, probleme (fragmentare, irigații), politica agricolă comună"},
            {"an": 2023, "s1": "Hartă Europa — identificare state, capitale, mări, lanțuri muntoase (Alpi, Pirinei, Scandinavici) | S2: Relieful României — Carpații (origine, etaje altitudinale, tipuri de roci, resurse), Subcarpații (formare, tipuri, importanță economică) | S3: Eseu: Industria românească după 1990 — restructurare, ramuri competitive (auto, IT, agro-alimentar), centre industriale, investiții străine directe"},
            {"an": 2024, "s1": "Hartă Romania — hidrografie, identificare bazine hidrografice, lacuri | S2: Transporturile în România — rețeaua rutieră (autostrăzi, drumuri europene), rețeaua feroviară, transportul naval pe Dunăre, aeroporturi internaționale, coridoarele paneuropene | S3: Eseu: Turismul în România — resurse naturale (munte, litoral, deltă, izvoare minerale) și culturale (cetăți, mănăstiri, orașe medievale), principalele stațiuni și circuite, probleme și perspective"},
            {"an": 2025, "s1": "Hartă UE — state membre, candidate, instituții principale (Bruxelles, Strasbourg, Frankfurt) | S2: Vegetația și solurile României — etajarea vegetației (etaj alpin, subalpin, forestier, de stepă), tipuri de soluri și distribuție, zone protejate (Retezat, Bucegi, Deltă) | S3: Eseu: Uniunea Europeană — etapele extinderii (de la 6 la 27 state), instituțiile principale (Parlament, Comisie, Consiliu), politici comune (agricolă, regională, monetară), aderarea României (2007), fonduri europene"},
        ],
    },
    "romana_uman": {
        "tipare_s1_itemi": [
            "1. Indică sensul din text al cuvântului X și al secvenței Y",
            "2. Menționează o caracteristică/profesie/calitate a personajului X, valorificând textul",
            "3. Precizează momentul/reacția/trăsătura morală + justifică cu o secvență din text",
            "4. Explică motivul pentru care... / reprezintă un eveniment / are loc situația X",
            "5. Prezintă în 30-50 cuvinte atmosfera/atitudinea/o situație conform textului",
        ],
        "teme_argumentativ": [
            "importanța lecturii / culturii generale în formarea personalității",
            "rolul școlii / al educației în societatea contemporană",
            "influența tehnologiei / rețelelor sociale asupra tinerilor",
            "necesitatea voluntariatului și a implicării civice",
            "valoarea prieteniei și a relațiilor autentice",
            "importanța cunoașterii istoriei și a identității naționale",
        ],
        "repere_s3": [
            "1. Prezentarea statutului social, psihologic, moral al personajului ales",
            "2. Evidențierea unei trăsături prin două episoade sau secvențe comentate",
            "3. Analiza a două elemente de structură/compoziție/limbaj",
            "4. Exprimarea unui punct de vedere argumentat despre semnificația personajului",
        ],
        "autori_opere": [
            "Mihai Eminescu — Luceafărul, Floare albastră, Scrisoarea I, Odă (în metru antic)",
            "Ioan Slavici — Moara cu noroc, Popa Tanda",
            "Ion Luca Caragiale — O scrisoare pierdută, O noapte furtunoasă, Vizita, La hanul lui Mânjoală",
            "Ion Creangă — Amintiri din copilărie, Harap-Alb",
            "Liviu Rebreanu — Ion, Pădurea spânzuraților",
            "Camil Petrescu — Ultima noapte de dragoste, întâia noapte de război",
            "G. Călinescu — Enigma Otiliei",
            "Tudor Arghezi — Testament, Flori de mucigai, Eu nu strivesc corola de minuni a lumii",
            "Lucian Blaga — Eu nu strivesc..., Sufletul satului, Mioara năzdrăvană",
            "Ion Barbu — Riga Crypto și lapona Enigel, Joc secund",
        ],
        "subiecte_reale": [
            {"an": 2022, "s1_text": "Text despre importanța literaturii și a lecturii — fragment eseu critic", "s1_B": "Argumentează dacă lectura cărților este esențială în era digitală (150-200 cuvinte)", "s2": "Comentează relația dintre ideea poetică și mijloacele artistice în poezia 'Floare albastră' de Eminescu (50+ cuvinte)", "s3": "Eseu: particularitățile de construcție ale unui personaj dintr-un roman al lui Liviu Rebreanu (Ion sau Apostol Bologa) — statut, trăsătură prin 2 episoade, 2 elemente structurale"},
            {"an": 2023, "s1_text": "Fragment memorialistic — scriitor român despre formarea sa culturală", "s1_B": "Text argumentativ despre influența mentorilor asupra formării intelectuale (150+ cuvinte)", "s2": "Rolul notațiilor scenice (didascalii) într-un fragment din 'O scrisoare pierdută' de Caragiale (50+ cuvinte)", "s3": "Eseu: construcția unui personaj dintr-o nuvelă de Ioan Slavici (Ghiță din Moara cu noroc sau altul) — statut + trăsătură + elemente de limbaj"},
            {"an": 2024, "s1_text": "Fragment dintr-un jurnal/scrisoare despre identitate culturală românească", "s1_B": "Argumentează dacă cunoașterea istoriei naționale este importantă pentru tinerii de azi (150+ cuvinte)", "s2": "Perspectiva narativă într-un fragment din 'Enigma Otiliei' de G. Călinescu — narator omniscient, focalizare (50+ cuvinte)", "s3": "Eseu: particularitățile unui text poetic modernist — Arghezi sau Blaga: titlu, imaginar poetic, procedee artistice, mesaj"},
            {"an": 2025, "s1_text": "Fragment despre rolul culturii și al artei în societate contemporană", "s1_B": "Argumentează dacă arta (muzică, pictură, teatru) poate influența comportamentul social al tinerilor (150+ cuvinte)", "s2": "Comentează în min. 50 cuvinte relația dintre tema și viziunea despre lume în 'Luceafărul' de Eminescu — condiția geniului, iubire, nemurire", "s3": "Eseu: particularitățile de construcție ale personajului Otilia Mărculescu din 'Enigma Otiliei' sau alt personaj feminin din romanul românesc interbelic — statut psihologic/moral, trăsătură prin 2 episoade, 2 elemente de compoziție/limbaj"},
        ],
    },
    "informatica": {
        "tipare": {
            "cpp": [
                "Tablouri unidimensionale — parcurgere, sortare (bulă, selecție, inserție), căutare binară",
                "Tablouri bidimensionale — matrice, spirale, diagonale, transpose",
                "Șiruri de caractere — operații, palindrom, anagrame, conversii",
                "Recursivitate — factorial, Fibonacci, turnurile din Hanoi, căutare",
                "Subprograme — funcții, proceduri, transmitere parametri (valoare/referință)",
                "Fișiere text — citire, scriere, prelucrare linie cu linie",
                "Structuri (struct) — definire, tablouri de structuri, sortare după câmp",
                "Algoritmi de graf — BFS, DFS, componente conexe, drum minim",
            ],
        },
        "subiecte_reale": [
            {"an": 2021, "profil": "C++", "s1": "Pseudocod: algoritm sortare + complexitate | S2: Funcție recursivă suma cifrelor, tablou bidimensional diagonale | S3: Fișier text cu numere — cel mai lung subșir crescător, afișare cu frecvențe"},
            {"an": 2022, "profil": "C++", "s1": "Algoritm parcurgere matrice în spirală, pseudocod | S2: Struct Produs (nume, preț, cantitate), sortare după preț, total valoare | S3: Graf neorientat — componente conexe (BFS/DFS), drum între noduri, matrice adiacență"},
            {"an": 2023, "profil": "C++", "s1": "Subprogram interclasare două tablouri sortate | S2: Tablou numere: cifrele distincte, frecvențe, cel mai mare număr format din cifre distincte | S3: Fișier: cuvinte distincte, frecvența maximă, anagramă, palindrom"},
            {"an": 2024, "profil": "C++", "s1": "Funcție verificare număr prim, generare șir primele n prime | S2: Matrice — simetrie față de diagonala secundară, element maxim per linie, zigzag | S3: Graf orientat — drum de cost minim (Dijkstra/Lee), existența unui circuit, număr noduri accesibile din sursă"},
            {"an": 2025, "profil": "C++", "s1": "Algoritm căutare binară în tablou sortat, pseudocod + complexitate | S2: Struct Student (nume, note[5]) — medie, promovabilitate, clasament | S3: Problema rucsacului (programare dinamică) sau șirul lui Fibonacci cu memoizare — implementare C++, analiza complexității"},
        ],
    },
    "fizica_real": {
        "structura": {
            "arii": ["A. Mecanică", "B. Elemente de termodinamică", "C. Producerea și utilizarea curentului continuu", "D. Optică"],
            "regula": "Candidatul alege 2 arii din 4 și rezolvă toate cele 3 subiecte din fiecare arie aleasă.",
            "punctaj": "2 arii × 45p + 10p oficiu = 100p → nota = punctaj / 10",
            "timp": "3 ore",
            "structura_arie": "I. 5 grile × 3p = 15p | II. Problemă 15p (4 cerințe) | III. Problemă 15p (4 cerințe)",
        },
        "tipare": {
            "A_mecanica": [
                "Sistem cu două corpuri + scripete ideal (fir inextensibil, masă neglijabilă): calcul tensiune, accelerație, raport mase, forța din ax",
                "Corp pe plan înclinat cu/fără frecare + forță de tracțiune: diagrama forțelor, proiecții, energie cinetică, impuls",
                "Conservarea energiei mecanice: plan curb fără frecare + zonă orizontală cu frecare; coliziune cu perete (forța medie)",
                "Pendul simplu + proiectil: lucrul mecanic al greutății, viteza în punct intermediar, mișcare balistică după ruperea firului",
                "Grile recurente: lucrul forței din grafic F(x); unități de putere (W = N·m/s); Legea Hooke (k din grafic); v constantă → rezultanta = 0",
            ],
            "B_termodinamica": [
                "Gaz ideal în cilindru cu piston mobil: două compartimente (O₂ + N₂ sau Ne + O₂), piston termoizolant/termoconductor, echilibru mecanic și termic",
                "Ciclu termodinamic p–V sau V–T: ΔU = ν·Cv·ΔT, L = aria sub grafic, Q = ΔU + L, randament motor",
                "Calorimetrie: amestec două mase de apă, temperatura de echilibru",
                "Grile recurente: tipuri de procese din diagrame p–T/V–T; randament Carnot (η = 1 - Trece/Tcald); unități Cv (J·mol⁻¹·K⁻¹); CV vs CP vs γ",
            ],
            "C_curent_continuu": [
                "Circuit cu 2 generatoare + rezistoare serie/paralel: rezistență echivalentă, curent, tensiune pe rezistoare, condiție ampermetru = 0",
                "Circuit cu becuri (parametri nominali Ub, Pb) + rezistor + baterie cu r: Rbec, t.e.m., randament, energie în interval Δt",
                "Dimensionare rezistență: putere egală pe R1 și R2, rezistență internă din condiție de putere maximă",
                "Grile recurente: grafic U–I → E și r; randament η = Pext/Ptotal; R = ρ·l/S unități; U·R⁻¹·t unități (= C, sarcina)",
            ],
            "D_optica": [
                "Lentilă convergentă + ecran: calcul f, mărire β, distanță obiect–imagine; sistem acolat (1/F = 1/f₁ + 1/f₂)",
                "Dispozitiv Young: i = λ·D/d; distanțe între maxime/minime de ordine date; suprapunere franje pentru 2 radiații; efect trecere în mediu (n ≠ 1)",
                "Grile recurente: efect fotoelectric (Ec,max = hν − L); energie foton în eV; legea refracției sin(i)/sin(r) = n; unități convergență (dioptrii = m⁻¹)",
            ],
        },
        "subiecte_reale": [
            {
                "an": 2021, "varianta": 1,
                "A_mecanica": {
                    "I_grile": [
                        "Corp coboară pantă v=ct → rezultanta = 0 (c)",
                        "Definiție viteză medie: vm = Δx/Δt (b)",
                        "Unitate a·d → m²/s² (d)",
                        "Fir elastic l0=60cm, k=50N/m, tăiat la l'0=12cm → k'=250N/m; F=10N → x=4cm (c)",
                        "Lucrul forței din grafic F(x) 0→5m, trapez → 72J (b)",
                    ],
                    "II_problema": "Sistem A+B + scripete, a=5m/s² (A coboară). a) v la t=0.5s = 2.5m/s b) Desen forțe c) mA/mB = 3/2 d) Forța ax dacă mA=300g → 6N",
                    "III_problema": "Corp m=1kg, plan înclinat α=30°, F=40N (de-a lungul planului), d=20cm, apoi F dispare. a) Lmg pe d = -1J b) Ec când F dispare = 7J c) h₁ când Ec = Ep la coborâre d) Impuls la revenire la bază",
                },
                "B_termodinamica": {
                    "I_grile": [
                        "Transformare izoterm: căldura = lucrul mecanic (d)",
                        "Lucrul mecanic adiabatic = -Cv·ΔT (c)",
                        "Unitate 1/(ρ·V) → m⁻³ (b)",
                        "Gaz la p=1.662×10⁵ Pa, n/V=2.408×10²⁶ m⁻³ → T=500K→t=227°C (c)",
                        "Ciclu V–T: relație presiuni p₁=p₃>p₂=p₄ (a)",
                    ],
                    "II_problema": "Azot m=70g, μ=28g/mol, cilindru orizontal cu piston mobil, p0=10⁵Pa, t₁=7°C. a) N molecule b) densitate ρ c) Se adaugă azot, se încălzește la t₂=27°C, p₂=1.5×10⁵Pa → Δm d) Piston liber, răcire la t₁ → raport V_final/V_inițial",
                    "III_problema": "1 mol O₂ (μ=32, Cv=2.5R), ciclu ABCA în p–V, pA=400kPa, ρA=3.2kg/m³, pB=2pA. a) Lucru mecanic total b) Căldura cedată c) Randament motor d) Randament Carnot între Tmax și Tmin",
                },
                "C_curent": {
                    "I_grile": [
                        "Baterie r=0, rezistor dublu în serie → I scade (c)",
                        "U·R⁻¹·t = sarcina Q (a)",
                        "Unitate P/I² → Ω (d)",
                        "Grafic U(I) → r din pantă = 0.6Ω (d)",
                        "Coeficient temperatură rezistivitate: α=(R-R0)/(R0·t)=4.5×10⁻³K⁻¹ (a)",
                    ],
                    "II_problema": "E1=16V, r1=2Ω, r2=1Ω, R1=4Ω, R2=9Ω, R3=7.2Ω, RA1=1Ω, I1=0.5A. a) UR3 b) IA1 c) E2 d) Rezistivitate din R2, L=75m, S=0.75mm²",
                    "III_problema": "R1=18Ω, R2=12Ω în paralel, E=40V, P1=72W. a) Energie R1 în 5h b) I generator c) Randament d) Rezistență interioară r",
                },
                "D_optica": {
                    "I_grile": [
                        "Lentilă → raze deviate prin refracție (b)",
                        "Unitate C·f → adimensional (a)",
                        "Raza prelungire prin F obiect → iese paralelă cu axa (a)",
                        "λ=400nm, L=3.85×10⁻¹⁹J → Ec,max=1.1×10⁻²⁰J (b)",
                        "Aer→mediu n: sin r = n·sin i deci r<i (c)",
                    ],
                    "II_problema": "Lentilă conv. f1=20cm, β=4 (imagine de 4×). a) y2=4cm b) x2=100cm c) Construcție grafică d) Al doilea sistem acolat f2=10cm, raze intră și ies paralele → distanță d=30cm",
                    "III_problema": "Young, λv=400nm, λr=700nm, iv=0.8mm. a) ir=1.4mm b) Distanța maxime ord.3: 1.8mm c) Prima suprapunere de maxime la 2.8mm d) Lungime de undă care face maxim la 2mm",
                },
            },
            {
                "an": 2022, "varianta": 1,
                "A_mecanica": {
                    "I_grile": [
                        "Corp cade v=ct → rezultanta = 0 (c)",
                        "Putere medie P = L/Δt (b)",
                        "N·m·s⁻¹ = putere mecanică (c)",
                        "m=250g, F=2N, v=ct → μ=0.8 (d)",
                        "Grafic F-x resort → k=125N/m (a)",
                    ],
                    "II_problema": "m=1kg, M=4kg, plan înclinat α=30°, a=1m/s², μ=1/√3≈0.58. a) Tensiune T=35N b) Forța ax scripete c) Forța frecare d) Forța F",
                    "III_problema": "m=0.6kg, h=25.1m, alunecare fără frecare pe AB, μ=0.2 pe BC (d=2.25m), lovire perete (Δt=7×10⁻³s), revenire oprire în B. a) Em în A b) v în B c) Impuls în C d) Forța medie a peretelui",
                },
                "B_termodinamica": {
                    "I_grile": [
                        "Căldura specifică = căldura per kg per K (c)",
                        "CV = R/(γ-1) (a)",
                        "Unitate CV·T → J·mol⁻¹ (d)",
                        "η=40%, T2=300K → T1=500K (b)",
                        "Proces 1→2 din p-V, T scade de 4×: densitate crește de 2× (b)",
                    ],
                    "II_problema": "Cilindru: O₂ V1=3L, T1=400K; N₂ V2=9L, T2=300K; aceeași CV. Piston termocondutor, deblocat. a) ν total b) p_O2/p_N2 c) V_azot la echilibru d) T echilibru",
                    "III_problema": "1 mol gaz poliatomic (Cv=3R), ciclu 1→2→3→1 în p-V. Q12=-12.8kJ (izoterm). ln5≈1.6. a) ΔU₃₁ b) Q₂₃ c) L total d) Randament",
                },
                "C_curent": {
                    "I_grile": [
                        "Rezistoare identice serie, se adaugă unul → I scade (a)",
                        "l = R·S/ρ (b)",
                        "Unitate W/(R·Δt)=A (a)",
                        "P egală pe R1=4Ω și R2 → r=6Ω (c)",
                        "I în [2s,6s] trapez: Q=12mC → N=7.5×10¹⁶ (c)",
                    ],
                    "II_problema": "R1=18Ω, R2=42Ω, R3=30Ω, R4=24Ω, E1=84V, E2=12V, r1=r2=2Ω. a) Req b) I ampermetru c) U voltmetru d) I cu E2 invers",
                    "III_problema": "Baterie r=8Ω, R=96Ω paralel cu 2 becuri identice Pb=6W, Ub=12V, I=0.75A. a) Rbec b) E c) P pe R d) Energie circuit exterior în 5 min",
                },
                "D_optica": {
                    "I_grile": [
                        "Oglindă plană → imagine virtuală (b)",
                        "1/x1 + 1/f = 1/x2 → expresia = 1/x2 (c)",
                        "Unitate c/n = viteză → m/s (a)",
                        "Sin(i)=0.75, refracție de-a lungul suprafeței → n=1/sin(i)=1.33 (c)",
                        "Din grafic Ec(ν): Ec=0.99×10⁻¹⁹J → Efoton din grafic (b)",
                    ],
                    "II_problema": "Lentilă L1, x1=-50cm, β=-1/4. a) β b) f=40cm c) Construcție d) Alipire L2 (f2=-25cm) → C_sistem=2.5dioptrii",
                    "III_problema": "Young λ=650nm, D=2m, i=1mm. a) ν=4.6×10¹⁴Hz b) d=1.3mm c) Distanța franja 3 întunecată – franja 2 luminoasă de cealaltă parte d) Franja 4 a λ se suprapune cu franja 5 a λ' → λ'=520nm",
                },
            },
            {
                "an": 2023, "varianta": 5,
                "A_mecanica": {
                    "I_grile": [
                        "Corp coboară uniform plan înclinat h1→h2: Lg = mg(h1-h2) (d)",
                        "Forța elastică pentru alungire x: F=kx (b)",
                        "Unitate constantei β din Fr = α·v + β·v²: kg·m⁻¹ (a)",
                        "Minge m=0.4kg, v=15m/s, Δt=0.01s → F_medie=600N (c)",
                        "Lucrul forței elastice din grafic F(x) până la x=4cm: L=-0.16J (c)",
                    ],
                    "II_problema": "m=0.5kg, plan α=37°(sin0.6), μ=0.30, F orizontal, grafic v(t). a) Desen forțe b) a=2m/s² c) F≈6.1N d) Timp urcare după oprire F la t=14s",
                    "III_problema": "m=1kg, platformă h=0.6m, F=5√2N la 45°, d=0.5m, μ=0.2. a) LF b) v după d c) Putere medie d) Impuls în momentul impactului cu solul",
                },
                "B_termodinamica": {
                    "I_grile": [
                        "Gaz cedează căldură: poate fi comprimare izoterma (a)",
                        "Lucru mecanic adiabatic: L = ν·R·(Ti-Tf)/(γ-1) (a)",
                        "Unitate p·V = N·m = J (b)",
                        "Gaz biatomic Cv=2.5R, destindere izobară Q=140kJ → ΔU=100kJ (c)",
                        "3 izocoare V-T: densitate proporțională cu 1/V la V fix → ρ1>ρ2>ρ3 (a)",
                    ],
                    "II_problema": "Cilindru orizontal, Ne (μ1=20) și O₂ (μ2=32), mase egale, t1=27°C, ν_total=6.5mol, piston termoizolant. a) l1/l2 b) ν_neon c) ΔT oxigen când piston la mijloc d) Masă molară amestec",
                    "III_problema": "Gaz poliatomic Cv=3R, ciclu 1→2→3→4→1 în V-T, p1=2×10⁵Pa, V1=1dm³. a) Ciclu în p-V b) ΔU₁₃ c) Q primit pe ciclu d) Randament Carnot între Tmax și Tmin",
                },
                "C_curent": {
                    "I_grile": [
                        "Baterie circuit deschis + voltmetru ideal → U = E (d)",
                        "Putere maximă pe exterior: Pmax = E²/(4r) (a)",
                        "Unitate U²·P⁻¹ = Ω (b)",
                        "Grafic I(t) în [3s,5s]: Q=10mC (b)",
                        "R1=2Ω, η=75% → r=? ; R2=1Ω → η=60% (b)",
                    ],
                    "II_problema": "2 generatoare identice serie E=9V, r=2Ω; R1=10Ω, R2=20Ω, R3=40Ω, R4=20Ω. a) Req b) I c) UR2 d) R4' astfel încât I_AB=0",
                    "III_problema": "2 becuri serie, E=120V, r=16Ω, bec1: P1=100W, U1=80V. a) Rbec1 b) Rbec2 c) Energie bec2 în 2min d) Randament circuit exterior",
                },
                "D_optica": {
                    "I_grile": [
                        "Raza reflectată ⊥ pe raza refractată (unghiuri complementare) → unghi 180° (d)",
                        "Sistem acolat: 1/F = 1/f1 + 1/f2 (a)",
                        "Unitate frecvență → s⁻¹ (a)",
                        "Umbră stâlp: proporție → H=10m (b)",
                        "Din grafic Ec(λ): lucrul de extracție = 2×10⁻¹⁹J (b)",
                    ],
                    "II_problema": "Obiect 5mm, ecran la d=100cm de obiect, imagine 20mm. a) β=4 b) f=20cm c) Construcție d) Distanța minimă obiect-ecran = 4f",
                    "III_problema": "Young λ1=400nm, λ2=600nm, i1=1mm. a) Al doilea minim la 1.5mm de centru b) i2=1.5mm c) Distanța max ord.1(λ1) față de max ord.4(λ2) d) Prima suprapunere maxime la 3mm",
                },
            },
            {
                "an": 2024, "varianta": 3,
                "A_mecanica": {
                    "I_grile": [
                        "Unitate putere mecanică: J/s = W (c)",
                        "Vectorul viteză instantanee → tangent la traiectorie (d)",
                        "Ep = mgh (c)",
                        "Grafic x(t) două mobile: raport viteze = raport pante → v1/v2=2 (b)",
                        "Resort k=50N/m, Δx=10cm: L=-0.25J (a)",
                    ],
                    "II_problema": "m1=3kg, m2=1kg, scripete S, F=20N la 37°, μ=0.20. a) Desen forțe m1 b) N pe m1 c) Accelerație sistem d) Forța pe scripete",
                    "III_problema": "m=2kg, v0=4m/s pe plan înclinat, h_max=0.5m. a) Impuls la lansare b) Lg pe urcare c) L_frecare pe urcare d) Ec la revenire la bază",
                },
                "B_termodinamica": {
                    "I_grile": [
                        "Compresia Otto → adiabatică (b)",
                        "Principiul I: ΔU = Q - L (a)",
                        "Unitate căldură specifică: J·kg⁻¹·K⁻¹ (d)",
                        "L=20kJ, Qc=30kJ → Qprimit=50kJ → η=40% (c)",
                        "Ciclu p–T, 2 mol, T1=400K: calcul L total (a)",
                    ],
                    "II_problema": "Cilindru L=52cm, ν1=3mol O₂ T1=300K; ν2=1mol N₂ T2=400K, p2=8.31×10⁴Pa. Piston deblocat, T constante. a) m_O2 b) ρ_N2 c) Deplasare piston d) Masă molară amestec",
                    "III_problema": "2mol O₂ (Cv=2.5R), ciclu 1→2→3 în V-T, T1=400K. a) Ciclu în p-V b) ΔU₁₃ c) Q₁₃ d) Randament Carnot între Tmax și Tmin",
                },
                "C_curent": {
                    "I_grile": [
                        "Temperatura scade → rezistivitate scade (c)",
                        "R = ρ·l/S (a)",
                        "Unitate U·I·t = J (d)",
                        "r=6Ω, P egală pe R1=12Ω și R2 → R2=3Ω (a)",
                        "Grafic I(U) → E=12V, r=2Ω (c)",
                    ],
                    "II_problema": "E1=24V, E2=12V, r1=r2=15Ω, R1=20Ω, R2=40Ω, R3=30Ω, întrerupător K. a) I când K deschis b) Req K închis c) I prin R3 K închis d) I când R2 scurtcircuitat",
                    "III_problema": "2 becuri identice în paralel Ub=12V, Ib=1A, E=15V. a) Pbec b) Energie 2 becuri în 1h c) Randament d) Putere maximă pe extern",
                },
                "D_optica": {
                    "I_grile": [
                        "Unitate constantei b din n=a+b/λ²: m² (d)",
                        "Oglindă plană → imagine virtuală și dreaptă (c)",
                        "ε1>ε2>ε3 → λ1<λ2<λ3 (b)",
                        "Grafic U_stopare(ν), λ_prag din intersecție cu axa → 600nm (d)",
                        "i=60°, r=30° → n=sin60°/sin30°=√3 → v=c/n=1.73×10⁸m/s (a)",
                    ],
                    "II_problema": "Obiect 1cm, lentilă, imagine pe ecran la 90cm de obiect. a) Construcție b) f=22.5cm c) Înălțime imagine d) Al 2-lea sistem centat, fascicul intră și iese paralel → distanță între lentile",
                    "III_problema": "Young d=1mm, D=2m, λ=600nm. a) i=1.2mm b) Diferența drum optic maxim ord.2 = 1200nm c) Ecranul deplasat ΔD=0.5m → deplasare max.3 d) Dispozitiv în lichid, i neschimbat față de a) → n lichid",
                },
            },
            {
                "an": 2025, "varianta": 1,
                "A_mecanica": {
                    "I_grile": [
                        "Legea Hooke: forță dublă → alungire dublă (b)",
                        "Accelerație medie: am = Δv/Δt (a)",
                        "v = α·x + β: unitate β/α → m/s / (1/m) = m²/s · 1/s... unitate α=s⁻¹, β = m·s⁻¹ → β/α = m (c)",
                        "P=25kW, F=1000N → v=25m/s=90km/h (c)",
                        "Grafic x(t) → v = pantă = 2m/s (b)",
                    ],
                    "II_problema": "m1=2kg, m2=1kg, scripete, F=8N, μ1=0.1, v=ct. a) Desen forțe m1 b) T din ecuații c) μ2 (frecare m2) d) Reacțiune ax scripete",
                    "III_problema": "Pendul l=1m, H=3.8m față de sol, eliberat din A (α=53°), fir se rupe în B (α=37°). a) Em în A b) Lgr de la A la B c) v în B d) Impuls înainte de impact cu solul",
                },
                "B_termodinamica": {
                    "I_grile": [
                        "C = Q/ΔT = capacitate calorică (c)",
                        "Randament Carnot: η = 1 - Trece/Tcald (c)",
                        "Unitate căldură specifică: J·kg⁻¹·K⁻¹ (c)",
                        "Ciclu p-T: volum minim în starea cu cel mai mic V din V=νRT/p → starea 1 (a)",
                        "m1=2kg, t1=80°C + m2=3kg, t2=10°C → t_ec=38°C (b)",
                    ],
                    "II_problema": "ν1=2mol O₂ (Cv=2.5R), T=300K, p=1.5×10⁵Pa, se adaugă He (Cv=1.5R, T=300K) până la V dublu, p const. a) m_O2 b) ρ_O2 c) ν_He adăugat d) U internă amestec",
                    "III_problema": "ν=0.24mol gaz monoatomic Cv=1.5R, ciclu 1→2→3→1 în p-V, T1=300K. a) ΔU₃₁ b) L total pe ciclu c) Q cedată d) Randament motor",
                },
                "C_curent": {
                    "I_grile": [
                        "Putere maximă: Pmax=E²/(4r) (d)",
                        "Rezistivitate vs temperatură: ρ=ρ0(1+αt) (a)",
                        "Unitate U·I = W (b)",
                        "Grafic I(U) → E=8V, r=1Ω (a)",
                        "E=9V, r=3Ω, R=12Ω → η=R/(R+r)=80% (c)",
                    ],
                    "II_problema": "E=12V, r=4Ω, R1=12Ω, R2=R3=16Ω, voltmetru ideal. a) Req exterior b) U borne baterie c) U voltmetru d) U ampermetru ideal în loc de voltmetru",
                    "III_problema": "E=12V, r=4Ω, bec Ub=9V Pb=4.5W, K deschis. a) Rbec b) Energie totală baterie în 10min c) R1 din circuit d) K închis, E'=17V, r=4Ω → R2 ca becul să funcționeze nominal",
                },
                "D_optica": {
                    "I_grile": [
                        "Raza reflectată ⊥ pe incidentă → unghi incidență = 45° (c)",
                        "Constanta Planck h = ε/ν → ε/ν = constantă (a)",
                        "Perechi diferite: frecvența luminii (Hz) vs convergența lentilei (dioptrii = m⁻¹) (d)",
                        "λ=600nm → E=hc/λ=3.3×10⁻¹⁹J=2eV (b)",
                        "Din grafic Ec(Efoton), pentru Efoton=15eV → Ec=10eV (b)",
                    ],
                    "II_problema": "Obiect 10mm, lentilă biconvexă f=20cm, obiect și imagine la distanțe egale. a) x=2f=40cm b) imagine 10mm c) Construcție d) Lentilă tăiată → 2 lentile plan-convexe, C=1/2f=2.5dioptrii",
                    "III_problema": "Young în aer, d=0.9mm, D=2.25m, ν=5×10¹⁴Hz. a) λ=600nm b) i=1.5mm c) Ecran deplasat ΔD=0.75m → deplasare max.3 d) Dispozitiv în lichid, max.3 revine la poziția inițială → n lichid",
                },
            },
        ],
    },
}




def extract_text_from_photo(image_bytes: bytes, materie_label: str, mime_type: str = "image/jpeg") -> str:
    """Extrage textul scris de mână dintr-o fotografie folosind Gemini Vision.
    
    Folosește Google Files API (upload real) în loc de base64 inline —
    același mecanism ca în sidebar, pentru analiză vizuală completă.
    """
    try:
        _pk = _pick_key_and_model(MODEL_CHAIN)
        if _pk:
            st.session_state.key_index = _pk[0]
        key = keys[st.session_state.get("key_index", 0)]
        gemini_client = genai.Client(api_key=key)

        # FIX bug 1: upload-ul fișierului e mutat ÎNĂUNTRUL contextului with —
        # tmp_path există garantat când îl folosim, TemporaryDirectory îl curăță după ieșire
        with tempfile.TemporaryDirectory() as tmpdir:
            _ext = {"image/png": ".png", "image/webp": ".webp", "image/heic": ".heic", "image/heif": ".heif"}.get(mime_type, ".jpg")
            tmp_path = os.path.join(tmpdir, "upload" + _ext)
            with open(tmp_path, "wb") as tmp:
                tmp.write(image_bytes)
            gfile = gemini_client.files.upload(file=tmp_path, config=genai_types.UploadFileConfig(mime_type=mime_type))
        # Fișierul temporar a fost șters de TemporaryDirectory; gfile (referința Google) rămâne validă

        poll = 0
        _ocr_status = st.empty()
        while str(gfile.state) in ("FileState.PROCESSING", "PROCESSING") and poll < 30:
            _ocr_status.caption(f"\u23f3 Procesare imagine Google ({poll + 1}s)...")
            time.sleep(1)
            gfile = gemini_client.files.get(name=gfile.name)
            poll += 1
        _ocr_status.empty()

        if not _is_gfile_active(gfile):
            try:
                gemini_client.files.delete(name=gfile.name)
            except Exception:
                pass
            return "[Eroare: imaginea nu a putut fi procesată de Google]"

        prompt = (
            f"Ești un asistent care transcrie text scris de mână din lucrări de elevi la {materie_label}. "
            f"Transcrie EXACT tot ce este scris în imagine, inclusiv formule, simboluri matematice și calcule. "
            f"Păstrează structura (Subiectul I, II, III dacă există). "
            f"Dacă un cuvânt e greu de citit, transcrie-l cu [?]. "
            f"Nu adăuga nimic, nu corecta nimic — transcrie fidel."
        )
        # Prin rotație: dacă modelul/cheia e epuizată, trece automat pe următoarele
        # (octeții pozei sunt păstrați ca imaginea să poată fi re-urcată pe altă cheie/proiect).
        _gsrc = st.session_state.get("_gsrc_map") or {}
        _gsrc[gfile.name] = (image_bytes, mime_type)
        st.session_state["_gsrc_map"] = _gsrc
        try:
            text = "".join(run_chat_with_rotation(
                [], [gfile, prompt],
                system_prompt="Transcrii fidel textul din imagini. Răspunzi doar cu textul transcris.",
                use_cache=False,
            ))
            return text.strip()
        finally:
            try:
                genai.Client(api_key=keys[st.session_state.get("key_index", 0)]).files.delete(name=gfile.name)
            except Exception:
                pass
            _gsrc.pop(gfile.name, None)

    except Exception as e:
        _log("Eroare OCR poză", "silent", e)
        return "[Eroare la citirea pozei]"


def get_bac_prompt_ai(materie_label, materie_info, profil):
    cod = materie_info.get("cod", "")
    date_reale = materie_info.get("date_reale", False)  # FIX bug 10: folosit în fallback generic

    # ── MATEMATICĂ M1 — date reale 2021-2025 ──
    if cod == "matematica_m1":
        data = BAC_DATE_REALE["matematica_m1"]
        tipare = data["tipare"]
        # Alege un subiect real ca referință (random)
        ref = random.choice(data["subiecte_reale"])
        tipare_str = "\n".join(f"  - {t}" for t in tipare)
        return (
            f"Generează un subiect COMPLET de BAC la Matematică M1 (mate-info), "
            f"IDENTIC ca structură și dificultate cu subiectele oficiale române din 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ (obligatorie):\n"
            f"SUBIECTUL I (30 de puncte) — 6 exerciții × 5p fiecare:\n"
            f"  Tipare care se repetă an de an:\n{tipare_str}\n\n"
            f"SUBIECTUL al II-lea (30 de puncte) — 2 probleme structurate (a, b, c):\n"
            f"  Problema 1: matrice 3×3 cu parametru real — det, inversabilitate, proprietăți\n"
            f"  Problema 2: lege de compoziție pe ℝ — calcule punctuale, element neutru, inegalități\n\n"
            f"SUBIECTUL al III-lea (30 de puncte) — 2 probleme structurate (a, b, c):\n"
            f"  Problema 1: funcție cu ln sau eˣ — arătați f'(x), monotonie, soluție unică f(x)=0\n"
            f"  Problema 2: integrală definită — calculați ∫, proprietate integrală, limită tip lim(1/x)∫₀ˣ\n\n"
            f"REFERINȚĂ (subiect real {ref['an']}):\n"
            f"  S.I: {ref['s1']}\n"
            f"  S.II: {ref['s2']}\n"
            f"  S.III: {ref['s3']}\n\n"
            f"IMPORTANT:\n"
            f"- Folosește numere și funcții DIFERITE față de exemplul de referință\n"
            f"- Dificultatea trebuie să fie realistă pentru BAC național\n"
            f"- Formulează cerințele exact ca la examen ('Arătați că...', 'Determinați...', 'Demonstrați că...')\n"
            f"- 10 puncte din oficiu\n\n"
            f"La final adaugă baremul:\n"
            f"[[BAREM_BAC]]\n"
            f"SUBIECTUL I: [răspunsurile corecte pentru fiecare item]\n"
            f"SUBIECTUL al II-lea: [soluțiile complete pas cu pas]\n"
            f"SUBIECTUL al III-lea: [soluțiile complete pas cu pas]\n"
            f"[[/BAREM_BAC]]"
        )

    # ── FIZICĂ TEHNOLOGIC — date reale 2021-2025 ──
    elif cod == "fizica_tehnologic":
        data = BAC_DATE_REALE["fizica_tehnologic"]
        tipare = data["tipare"]
        # Alege 2 arii random pentru subiect
        arii_disponibile = ["A. MECANICĂ", "B. TERMODINAMICĂ", "C. CURENT CONTINUU", "D. OPTICĂ"]
        arii_alese = random.sample(arii_disponibile, 2)
        tipare_mec = "\n".join(f"    - {t}" for t in tipare["mecanica"])
        tipare_term = "\n".join(f"    - {t}" for t in tipare["termodinamica"])
        tipare_cur = "\n".join(f"    - {t}" for t in tipare["curent"])
        tipare_opt = "\n".join(f"    - {t}" for t in tipare["optica"])
        return (
            f"Generează un subiect COMPLET de BAC la Fizică — filiera tehnologică, "
            f"IDENTIC ca structură cu subiectele oficiale române din 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ:\n"
            f"Subiectul are 4 ARII tematice (A–D). Candidatul rezolvă DOAR 2 la alegere.\n"
            f"Generează TOATE cele 4 arii. Pentru fiecare arie:\n"
            f"  - Subiectul I (15 puncte): 5 itemi tip GRILĂ (a, b, c, d) × 3p\n"
            f"  - Subiectul II (15 puncte): o problemă structurată cu 4 cerințe (a, b, c, d)\n"
            f"  - Subiectul III (15 puncte): o problemă mai complexă cu 4 cerințe (a, b, c, d)\n\n"
            f"TIPARE REALE PE ARII:\n"
            f"A. MECANICĂ:\n{tipare_mec}\n\n"
            f"B. TERMODINAMICĂ:\n{tipare_term}\n\n"
            f"C. CURENT CONTINUU:\n{tipare_cur}\n\n"
            f"D. OPTICĂ:\n{tipare_opt}\n\n"
            f"IMPORTANT:\n"
            f"- Datele numerice trebuie să fie realiste și să dea calcule curate\n"
            f"- Formulează grilele cu exact 4 variante, dintre care exact una corectă\n"
            f"- Problemele din S.II și S.III trebuie să fie rezolvabile pas cu pas\n"
            f"- Indică la fiecare arie: 'Aria A — Mecanică', etc.\n"
            f"- 10 puncte din oficiu\n\n"
            f"[[BAREM_BAC]]\n"
            f"ARIA A: [răspunsuri grilă + soluții probleme]\n"
            f"ARIA B: [răspunsuri grilă + soluții probleme]\n"
            f"ARIA C: [răspunsuri grilă + soluții probleme]\n"
            f"ARIA D: [răspunsuri grilă + soluții probleme]\n"
            f"[[/BAREM_BAC]]"
        )

    # ── ROMÂNĂ REAL/TEHNOLOGIC — date reale 2021-2025 ──
    elif cod == "romana_real_tehn":
        data = BAC_DATE_REALE["romana_real_tehn"]
        ref = random.choice(data["subiecte_reale"])
        itemi_str = "\n".join(f"  {it}" for it in data["tipare_s1_itemi"])
        teme_str = "\n".join(f"  - {t}" for t in data["teme_argumentativ"])
        s2_str = "\n".join(f"  - {t}" for t in data["tipare_s2"])
        repere_str = "\n".join(f"  {r}" for r in data["repere_s3"])
        autori_str = "\n".join(f"  - {a}" for a in data["autori_opere"])
        return (
            f"Generează un subiect COMPLET de BAC la Limba și literatura română — profil real/tehnologic, "
            f"IDENTIC ca structură cu subiectele oficiale din 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ:\n\n"
            f"SUBIECTUL I (50 de puncte):\n"
            f"Partea A (30 puncte) — Text la prima vedere (proză, memorialistică sau publicistică, 1-2 pagini).\n"
            f"Generează un text original de 200-300 cuvinte, apoi formulează EXACT 5 cerințe:\n"
            f"{itemi_str}\n\n"
            f"Partea B (20 puncte) — Text argumentativ de minimum 150 cuvinte pe o temă din text:\n"
            f"  Alege una dintre temele frecvente:\n{teme_str}\n"
            f"  Cerința standard: 'Redactează un text de minimum 150 de cuvinte, în care să argumentezi dacă [tema], "
            f"raportându-te atât la informațiile din textul dat, cât și la experiența personală sau culturală.'\n\n"
            f"SUBIECTUL al II-lea (10 puncte):\n"
            f"  Un fragment literar scurt (dramatic sau liric) + una din cerințele:\n{s2_str}\n\n"
            f"SUBIECTUL al III-lea (30 de puncte):\n"
            f"  Eseu de minimum 400 de cuvinte. Alege un autor și operă din:\n{autori_str}\n"
            f"  Formularea standard: 'Redactează un eseu de minimum 400 de cuvinte, în care să prezinți "
            f"particularitățile de construcție ale unui personaj dintr-un text narativ studiat.'\n"
            f"  Repere obligatorii (în barem):\n{repere_str}\n\n"
            f"REFERINȚĂ (structura subiectului real {ref['an']}):\n"
            f"  S.I text: {ref.get('s1_text', 'text la prima vedere')}\n"
            f"  S.II: {ref.get('s2', '')}\n"
            f"  S.III: {ref.get('s3', '')}\n\n"
            f"IMPORTANT:\n"
            f"- Textul de la S.I trebuie să fie original, coerent, de nivel liceal\n"
            f"- Fragmentul de la S.II trebuie să fie dintr-o operă reală din programa de liceu\n"
            f"- 10 puncte din oficiu\n\n"
            f"[[BAREM_BAC]]\n"
            f"SUBIECTUL I — Partea A: [răspunsurile așteptate pentru fiecare cerință + punctaj]\n"
            f"SUBIECTUL I — Partea B: [criterii text argumentativ + punctaj]\n"
            f"SUBIECTUL al II-lea: [răspuns așteptat + criterii + punctaj]\n"
            f"SUBIECTUL al III-lea: [repere eseu + criterii conținut (18p) + redactare (12p)]\n"
            f"[[/BAREM_BAC]]"
        )

    # ── MATEMATICĂ M2 ──
    elif cod == "matematica_m2":
        data = BAC_DATE_REALE["matematica_m2"]
        ref = random.choice(data["subiecte_reale"])
        tipare_str = "\n".join(f"  - {t}" for t in data["tipare"])
        return (
            f"Generează un subiect COMPLET de BAC la Matematică M2 (Științe ale naturii), "
            f"IDENTIC ca structură și dificultate cu subiectele oficiale române din 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ (obligatorie):\n"
            f"SUBIECTUL I (30 puncte) — 6 exerciții × 5p:\n"
            f"  Tipare reale:\n{tipare_str}\n\n"
            f"SUBIECTUL al II-lea (30 puncte) — 2 probleme structurate:\n"
            f"  Problema 1: funcții, monotonie, extreme, valori\n"
            f"  Problema 2: geometrie analitică sau matrice 2×2\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — 2 probleme structurate:\n"
            f"  Problema 1: derivate — f'(x), extreme, tangentă, convexitate\n"
            f"  Problema 2: integrale definite — calcul, arie, volum\n\n"
            f"REFERINȚĂ (subiect real {ref['an']}):\n"
            f"  S.I: {ref['s1']}\n  S.II: {ref['s2']}\n  S.III: {ref['s3']}\n\n"
            f"Folosește valori numerice DIFERITE față de referință. 10 puncte din oficiu.\n\n"
            f"[[BAREM_BAC]]\nSUBIECTUL I: [răspunsuri]\nSUBIECTUL al II-lea: [soluții pas cu pas]\nSUBIECTUL al III-lea: [soluții pas cu pas]\n[[/BAREM_BAC]]"
        )

    # ── BIOLOGIE ──
    elif cod == "biologie":
        data = BAC_DATE_REALE["biologie"]
        profil_key = "anatomie_fiziologie" if "Anatomie" in profil else "vegetala_animala"
        tipare = data["tipare"][profil_key]
        ref_list = [s for s in data["subiecte_reale"] if profil.split()[0].lower() in s.get("profil", "").lower()]
        ref = random.choice(ref_list) if ref_list else random.choice(data["subiecte_reale"])
        tipare_str = "\n".join(f"  - {t}" for t in tipare)
        return (
            f"Generează un subiect COMPLET de BAC la Biologie — {profil}, "
            f"IDENTIC ca structură cu subiectele oficiale BAC România 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ:\n"
            f"SUBIECTUL I (30 puncte) — 10 itemi GRILĂ × 3p (a, b, c, d — exact 1 corect)\n\n"
            f"SUBIECTUL al II-lea (30 puncte) — itemi semiobiectivi:\n"
            f"  - 2-3 cerințe tip completare/definire/comparație (10-15p)\n"
            f"  - 1 problemă de genetică sau fiziologie cu calcul (15-20p)\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — eseu structurat:\n"
            f"  - Prezintă complet un sistem/proces biologic cu: definiție, structură, funcționare, "
            f"reglare, afecțiuni, importanță\n\n"
            f"TIPARE REALE ({profil}):\n{tipare_str}\n\n"
            f"REFERINȚĂ (subiect real {ref['an']}, {ref.get('profil','')}):\n  {ref['s1']}\n\n"
            f"10 puncte din oficiu.\n\n"
            f"[[BAREM_BAC]]\nSUBIECTUL I: [grila: 1-x, 2-x, ...]\nSUBIECTUL al II-lea: [răspunsuri + punctaj]\nSUBIECTUL al III-lea: [repere eseu + punctaj]\n[[/BAREM_BAC]]"
        )

    # ── CHIMIE ──
    elif cod == "chimie":
        data = BAC_DATE_REALE["chimie"]
        profil_key = "anorganica" if "anorgan" in profil.lower() else "organica"
        tipare = data["tipare"][profil_key]
        ref_list = [s for s in data["subiecte_reale"] if profil.lower()[:5] in s.get("profil", "").lower()]
        ref = random.choice(ref_list) if ref_list else random.choice(data["subiecte_reale"])
        tipare_str = "\n".join(f"  - {t}" for t in tipare)
        return (
            f"Generează un subiect COMPLET de BAC la Chimie — {profil}, "
            f"IDENTIC ca structură cu subiectele oficiale BAC România 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ:\n"
            f"SUBIECTUL I (30 puncte) — 10 itemi GRILĂ × 3p (exact 1 variantă corectă din 4)\n\n"
            f"SUBIECTUL al II-lea (30 puncte) — probleme de calcul chimic:\n"
            f"  - Minimum 2 probleme structurate cu (a, b, c, d)\n"
            f"  - Include: ecuații chimice balansate, calcule cu moli, mase, volume, concentrații\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — problemă complexă sau sinteză:\n"
            f"  - Problemă cu 4-5 cerințe legate logic (calcule, explicații, aplicații)\n\n"
            f"TIPARE REALE ({profil}):\n{tipare_str}\n\n"
            f"REFERINȚĂ (subiect real {ref['an']}):\n  {ref['s1']}\n\n"
            f"IMPORTANT: Datele numerice să fie realiste (mase, volume, concentrații uzuale). 10 puncte din oficiu.\n\n"
            f"[[BAREM_BAC]]\nSUBIECTUL I: [1-x, 2-x, ...]\nSUBIECTUL al II-lea: [soluții pas cu pas cu ecuații]\nSUBIECTUL al III-lea: [soluție completă]\n[[/BAREM_BAC]]"
        )

    # ── ISTORIE ──
    elif cod == "istorie":
        data = BAC_DATE_REALE["istorie"]
        ref = random.choice(data["subiecte_reale"])
        teme_str = "\n".join(f"  - {t}" for t in data["tipare"]["teme_frecvente"])
        cerinte_str = "\n".join(f"  {c}" for c in data["tipare"]["cerinte_sursa"])
        eseu_str = "\n".join(f"  {e}" for e in data["tipare"]["structura_eseu"])
        return (
            f"Generează un subiect COMPLET de BAC la Istorie — {profil}, "
            f"IDENTIC ca structură cu subiectele oficiale BAC România 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ:\n\n"
            f"SUBIECTUL I (30 puncte) — Analiză surse istorice:\n"
            f"  Generează 2 surse scurte (câte 80-120 cuvinte fiecare) despre un eveniment/perioadă.\n"
            f"  Formulează EXACT 4 cerințe standard:\n{cerinte_str}\n\n"
            f"SUBIECTUL al II-lea (30 puncte) — Eseu scurt (1 pagină):\n"
            f"  'Prezentați două cauze/consecințe/caracteristici ale [eveniment/perioadă]'\n"
            f"  Sau: 'Menționați două acțiuni ale [personalitate] și explicați importanța lor'\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — Eseu structurat (2 pagini):\n"
            f"  'Elaborați un eseu despre [temă], în care să prezentați...'\n"
            f"  Structură eseu:\n{eseu_str}\n\n"
            f"TEME FRECVENTE (alege una din fiecare subiect):\n{teme_str}\n\n"
            f"REFERINȚĂ (structura {ref['an']}):\n"
            f"  S.I: {ref['s1']}\n  S.II: {ref['s2']}\n  S.III: {ref['s3']}\n\n"
            f"10 puncte din oficiu.\n\n"
            f"[[BAREM_BAC]]\nSUBIECTUL I: [răspunsuri așteptate per cerință + punctaj]\nSUBIECTUL al II-lea: [repere + punctaj]\nSUBIECTUL al III-lea: [repere eseu + conținut (18p) + redactare (12p)]\n[[/BAREM_BAC]]"
        )

    # ── GEOGRAFIE ──
    elif cod == "geografie":
        data = BAC_DATE_REALE["geografie"]
        ref = random.choice(data["subiecte_reale"])
        teme_str = "\n".join(f"  - {t}" for t in data["tipare"]["teme_frecvente"])
        eseu_str = "\n".join(f"  {e}" for e in data["tipare"]["structura_eseu"])
        return (
            f"Generează un subiect COMPLET de BAC la Geografie — {profil}, "
            f"IDENTIC ca structură cu subiectele oficiale BAC România 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ:\n\n"
            f"SUBIECTUL I (30 puncte) — Hartă:\n"
            f"  Descrie o hartă (Romania fizică, politică sau Europa) și formulează 5 cerințe:\n"
            f"  1. Identifică/numește elemente geografice indicate (litere A, B, C pe hartă)\n"
            f"  2. Precizează caracteristici ale unui element geografic identificat\n"
            f"  3. Menționează 2 caracteristici ale unui fenomen geografic din regiune\n"
            f"  4. Explică o relație cauză-efect (ex: relief → climă, climă → vegetație)\n"
            f"  5. Prezintă importanța economică a unui element identificat (30-50 cuvinte)\n\n"
            f"SUBIECTUL al II-lea (30 puncte) — Noțiuni geografice:\n"
            f"  - 3-4 cerințe de definire, exemplificare și caracterizare a unor noțiuni/fenomene\n"
            f"  - Include date statistice și exemple concrete din România/Europa\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — Eseu geografic (300-400 cuvinte):\n"
            f"  Structură eseu:\n{eseu_str}\n\n"
            f"TEME FRECVENTE:\n{teme_str}\n\n"
            f"REFERINȚĂ (structura {ref['an']}):\n"
            f"  S.I: {ref['s1']}\n  S.II: {ref['s2']}\n  S.III: {ref['s3']}\n\n"
            f"10 puncte din oficiu.\n\n"
            f"[[BAREM_BAC]]\nSUBIECTUL I: [răspunsuri per cerință + punctaj]\nSUBIECTUL al II-lea: [răspunsuri + punctaj]\nSUBIECTUL al III-lea: [repere eseu + punctaj]\n[[/BAREM_BAC]]"
        )

    # ── ROMÂNĂ UMAN/PEDAGOGIC ──
    elif cod == "romana_uman":
        data = BAC_DATE_REALE["romana_uman"]
        ref = random.choice(data["subiecte_reale"])
        itemi_str = "\n".join(f"  {it}" for it in data["tipare_s1_itemi"])
        teme_str = "\n".join(f"  - {t}" for t in data["teme_argumentativ"])
        repere_str = "\n".join(f"  {r}" for r in data["repere_s3"])
        autori_str = "\n".join(f"  - {a}" for a in data["autori_opere"])
        return (
            f"Generează un subiect COMPLET de BAC la Limba și literatura română — profil umanist/pedagogic, "
            f"IDENTIC ca structură cu subiectele oficiale din 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ:\n\n"
            f"SUBIECTUL I (50 puncte):\n"
            f"Partea A (30 puncte) — Text la prima vedere (proză, eseu, publicistică, 200-300 cuvinte).\n"
            f"Formulează EXACT 5 cerințe:\n{itemi_str}\n\n"
            f"Partea B (20 puncte) — Text argumentativ minimum 150 cuvinte:\n"
            f"  Temă din:\n{teme_str}\n\n"
            f"SUBIECTUL al II-lea (10 puncte):\n"
            f"  Fragment liric sau dramatic + cerință comentariu (50+ cuvinte)\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — Eseu 400+ cuvinte:\n"
            f"  Autor și operă din:\n{autori_str}\n"
            f"  Repere obligatorii:\n{repere_str}\n\n"
            f"REFERINȚĂ ({ref['an']}): S.I={ref.get('s1_text','')}, S.II={ref.get('s2','')}, S.III={ref.get('s3','')}\n\n"
            f"10 puncte din oficiu.\n\n"
            f"[[BAREM_BAC]]\nSUBIECTUL I-A: [răspunsuri + punctaj]\nSUBIECTUL I-B: [criterii text argumentativ]\nSUBIECTUL al II-lea: [răspuns + criterii]\nSUBIECTUL al III-lea: [repere + conținut 18p + redactare 12p]\n[[/BAREM_BAC]]"
        )

    # ── INFORMATICĂ ──
    elif cod == "informatica":
        data = BAC_DATE_REALE["informatica"]
        ref_list = data["subiecte_reale"]
        ref = random.choice(ref_list)
        tipare_str = "\n".join(f"  - {t}" for t in data["tipare"]["cpp"])
        return (
            f"Generează un subiect COMPLET de BAC la Informatică — {profil}, "
            f"IDENTIC ca structură cu subiectele oficiale BAC România 2021-2025.\n\n"
            f"STRUCTURĂ EXACTĂ:\n\n"
            f"SUBIECTUL I (30 puncte):\n"
            f"  a) Algoritm/pseudocod — citire, analiză, completare, trasare pentru date date (15p)\n"
            f"  b) Subprogram (funcție/procedură) — scriere cod {profil}, testare (15p)\n\n"
            f"SUBIECTUL al II-lea (30 puncte):\n"
            f"  Probleme cu tablouri/matrice/șiruri de caractere:\n"
            f"  a) Tablou unidimensional — prelucrare, sortare, căutare (15p)\n"
            f"  b) Tablou bidimensional sau structuri — prelucrare, afișare (15p)\n\n"
            f"SUBIECTUL al III-lea (30 puncte):\n"
            f"  Problemă completă — fișiere text sau grafuri sau programare dinamică:\n"
            f"  Citire din fișier text, prelucrare complexă, scriere rezultate (30p)\n\n"
            f"TIPARE REALE:\n{tipare_str}\n\n"
            f"REFERINȚĂ ({ref['an']}, {ref.get('profil','C++')}):\n  {ref['s1']}\n\n"
            f"IMPORTANT: Codul trebuie să fie corect sintactic în {profil}. Include date de test. 10p din oficiu.\n\n"
            f"[[BAREM_BAC]]\nSUBIECTUL I: [pseudocod corect + cod corect + punctaj]\nSUBIECTUL al II-lea: [cod corect + explicații]\nSUBIECTUL al III-lea: [soluție completă + complexitate]\n[[/BAREM_BAC]]"
        )

    # ── FIZICĂ REAL (profil real/teoretic) ──
    elif cod == "fizica_real":
        data = BAC_DATE_REALE["fizica_real"]
        import random as _rand
        # Alege o arie tematică principală și un subiect de referință
        arii_disponibile = list(data["tipare"].keys())
        arie1 = _rand.choice(["A_mecanica", "B_termodinamica"])
        arie2 = _rand.choice(["C_curent_continuu", "D_optica"])
        tipare1 = "\n".join(f"  - {t}" for t in data["tipare"][arie1])
        tipare2 = "\n".join(f"  - {t}" for t in data["tipare"][arie2])
        # Referință din subiecte reale
        ref = _rand.choice(data["subiecte_reale"])
        arie1_key = arie1.replace("A_mecanica","A_mecanica").replace("B_termodinamica","B_termodinamica")
        arie2_key = arie2.replace("C_curent_continuu","C_curent").replace("D_optica","D_optica")
        ref1 = ref.get(arie1_key, ref.get(arie1, {}))
        ref2 = ref.get(arie2_key, ref.get(arie2, {}))
        arie1_label = {"A_mecanica":"A. Mecanică","B_termodinamica":"B. Elemente de termodinamică"}[arie1]
        arie2_label = {"C_curent_continuu":"C. Producerea și utilizarea curentului continuu","D_optica":"D. Optică"}[arie2]
        return (
            f"Generează un subiect COMPLET de BAC la Fizică — profil real / filiera vocațională militar, "
            f"IDENTIC ca structură cu subiectele oficiale BAC România 2021-2025.\n\n"
            f"STRUCTURĂ OBLIGATORIE (2 arii tematice din 4):\n"
            f"Candidatul rezolvă 2 arii; tu generezi: {arie1_label} + {arie2_label}\n\n"
            f"PENTRU FIECARE ARIE, structura este:\n"
            f"  I. Pentru itemii 1-5 scrieți pe foaia de răspuns litera corespunzătoare răspunsului corect. (15 puncte)\n"
            f"     → 5 grile × 3p, 4 variante (a,b,c,d)\n"
            f"  II. Rezolvați următoarea problemă: (15 puncte)\n"
            f"     → problemă cu 4 cerințe a,b,c,d\n"
            f"  III. Rezolvați următoarea problemă: (15 puncte)\n"
            f"     → problemă mai complexă cu 4 cerințe a,b,c,d\n\n"
            f"TIPARE PENTRU {arie1_label}:\n{tipare1}\n\n"
            f"TIPARE PENTRU {arie2_label}:\n{tipare2}\n\n"
            f"REFERINȚĂ (BAC {ref['an']}) — folosește ca model de dificultate și stil de enunț:\n"
            f"  {arie1_label}: grile: {ref1.get('I_grile',['—'])[0] if isinstance(ref1.get('I_grile'), list) else '—'} | "
            f"Prob.II: {str(ref1.get('II_problema','—'))[:120]}...\n"
            f"  {arie2_label}: grile: {ref2.get('I_grile',['—'])[0] if isinstance(ref2.get('I_grile'), list) else '—'} | "
            f"Prob.II: {str(ref2.get('II_problema','—'))[:120]}...\n\n"
            f"REGULI:\n"
            f"  - Se acordă 10 puncte din oficiu. Timp: 3 ore.\n"
            f"  - Date numerice realiste, calcule exacte, g=10m/s², NA=6.02×10²³, R=8.31 J/(mol·K)\n"
            f"  - c=3×10⁸m/s, h=6.6×10⁻³⁴J·s (pentru aria D)\n"
            f"  - Enunțurile să fie clare, cu scheme descrise textual acolo unde e nevoie\n\n"
            f"[[BAREM_BAC]]\n"
            f"{arie1_label}: I: [1-x, 2-x, 3-x, 4-x, 5-x] II: [soluție pas cu pas] III: [soluție pas cu pas]\n"
            f"{arie2_label}: I: [1-x, 2-x, 3-x, 4-x, 5-x] II: [soluție pas cu pas] III: [soluție pas cu pas]\n"
            f"[[/BAREM_BAC]]"
        )

    # ── ECONOMIE ──
    elif cod == "economie":
        return (
            f"Generează un subiect COMPLET de BAC la Economie — {profil}, "
            f"IDENTIC ca structură și dificultate cu subiectele oficiale din România 2021–2025.\n\n"
            f"STRUCTURĂ EXACTĂ (obligatorie):\n\n"
            f"SUBIECTUL I (30 puncte) — Itemi obiectivi și semiobiectivi:\n"
            f"  A. 5 itemi cu alegere multiplă × 4p — piață, cerere, ofertă, prețuri de echilibru, utilitate\n"
            f"  B. 2 itemi semiobiectivi × 5p — completare definiții, relații economice\n\n"
            f"SUBIECTUL al II-lea (30 puncte) — Studiu de caz / problemă structurată:\n"
            f"  - Date numerice realiste (tabel cerere/ofertă SAU venituri/cheltuieli)\n"
            f"  - 4 cerințe (a, b, c, d): calcule, interpretare grafic, analiză impact politici economice\n"
            f"  - Include cel puțin o cerință de calcul (elasticitate, profit, cost marginal sau echilibru)\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — Eseu economic structurat:\n"
            f"  Alege una dintre temele: șomaj și politici de ocupare / inflație și efecte / PIB și creștere economică /\n"
            f"  comerț exterior și balanță de plăți / piața muncii / sisteme economice comparate\n"
            f"  Repere obligatorii în barem:\n"
            f"    1. Definiție și forme/tipuri ale fenomenului (6p)\n"
            f"    2. Cauze și efecte economice și sociale (12p)\n"
            f"    3. Politici economice de intervenție cu exemple concrete (8p)\n"
            f"    4. Opinie argumentată (4p)\n\n"
            f"10 puncte din oficiu.\n\n"
            f"IMPORTANT:\n"
            f"- Datele numerice trebuie să fie coerente și realiste (ex: prețuri în lei, procente plauzibile)\n"
            f"- Formulează cerințele exact ca la examen ('Calculați...', 'Precizați...', 'Explicați...')\n"
            f"- Eseul trebuie să poată fi scris în ~45 minute\n\n"
            f"[[BAREM_BAC]]\n"
            f"SUBIECTUL I — A: [răspunsuri corecte cu justificare] B: [răspunsuri complete]\n"
            f"SUBIECTUL al II-lea: [calcule pas cu pas + răspunsuri la fiecare cerință]\n"
            f"SUBIECTUL al III-lea: [repere eseu cu punctaj detaliat: conținut 24p + redactare 6p]\n"
            f"[[/BAREM_BAC]]"
        )

    # ── PSIHOLOGIE ──
    elif cod == "psihologie":
        return (
            f"Generează un subiect COMPLET de BAC la Psihologie — {profil}, "
            f"IDENTIC ca structură și dificultate cu subiectele oficiale din România 2021–2025.\n\n"
            f"STRUCTURĂ EXACTĂ (obligatorie):\n\n"
            f"SUBIECTUL I (30 puncte) — Procese și funcții psihice:\n"
            f"  A. 4 itemi cu alegere multiplă × 4p — senzații, percepție, memorie, gândire, imaginație, limbaj\n"
            f"  B. 2 itemi semiobiectivi × 7p — definiții, caracterizare, exemple din viața cotidiană\n\n"
            f"SUBIECTUL al II-lea (30 puncte) — Personalitate și psihologie aplicată:\n"
            f"  - Text scurt (100-150 cuvinte) despre o situație / comportament concret\n"
            f"  - 4 cerințe (a, b, c, d) bazate pe text:\n"
            f"    a. Identificare și definire concept psihologic din text (6p)\n"
            f"    b. Caracterizarea unui aspect al personalității: temperament / caracter / aptitudini (8p)\n"
            f"    c. Comparație sau analiză (motivație, afectivitate, voință) (8p)\n"
            f"    d. Exemplu personal argumentat (8p)\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — Eseu structurat:\n"
            f"  Alege una dintre temele: comunicare și relații interpersonale / grupuri sociale și influențe /\n"
            f"  sănătate mentală și stres / dezvoltare psihologică / conștiință de sine și identitate\n"
            f"  Repere obligatorii:\n"
            f"    1. Definiție și caracteristici (6p)\n"
            f"    2. Factori / forme / tipuri cu exemple (12p)\n"
            f"    3. Influențe și aplicații practice (8p)\n"
            f"    4. Concluzie personală argumentată (4p)\n\n"
            f"10 puncte din oficiu.\n\n"
            f"IMPORTANT:\n"
            f"- Textul de la S.II trebuie să fie autentic, relevant pentru adolescenți\n"
            f"- Formulează cerințele exact ca la examen ('Precizați...', 'Caracterizați...', 'Analizați...')\n\n"
            f"[[BAREM_BAC]]\n"
            f"SUBIECTUL I — A: [răspunsuri corecte] B: [răspunsuri complete cu punctaj]\n"
            f"SUBIECTUL al II-lea: [răspunsuri pentru fiecare cerință cu punctaj detaliat]\n"
            f"SUBIECTUL al III-lea: [repere eseu cu punctaj: conținut 24p + redactare 6p]\n"
            f"[[/BAREM_BAC]]"
        )

    # ── LOGICĂ ȘI ARGUMENTARE ──
    elif cod == "logica":
        return (
            f"Generează un subiect COMPLET de BAC la Logică și argumentare — {profil}, "
            f"IDENTIC ca structură și dificultate cu subiectele oficiale din România 2021–2025.\n\n"
            f"STRUCTURĂ EXACTĂ (obligatorie):\n\n"
            f"SUBIECTUL I (30 puncte) — Logică formală:\n"
            f"  A. 4 itemi cu alegere multiplă × 4p — propoziții logice, valori de adevăr,\n"
            f"     operatori (negație, conjuncție, disjuncție, implicație, echivalență)\n"
            f"  B. 3 itemi semiobiectivi × 6p:\n"
            f"     - Tabel de adevăr complet pentru o formulă (ex: ¬p∨q, p→q)\n"
            f"     - Identificare tip inferență (modus ponens, modus tollens, silogism)\n"
            f"     - Determinare validitate argument (cu contra-exemplu dacă invalid)\n\n"
            f"SUBIECTUL al II-lea (30 puncte) — Analiza argumentelor:\n"
            f"  - Text argumentativ scurt (100-150 cuvinte) pe o temă actuală\n"
            f"  - 4 cerințe (a, b, c, d):\n"
            f"    a. Identificare tip de raționament și structură (premize + concluzie) (6p)\n"
            f"    b. Evaluarea validității argumentului (deductiv/inductiv/analogie) (8p)\n"
            f"    c. Identificare sofisme sau erori de raționament (dacă există) cu denumire și explicație (8p)\n"
            f"    d. Reformulare corectă a argumentului sau contra-argument (8p)\n\n"
            f"SUBIECTUL al III-lea (30 puncte) — Construcție argument / eseu logic:\n"
            f"  Temă dată (alege una): 'Rațiunea este superioară emoției' / 'Tehnologia ne face mai puțin liberi' /\n"
            f"  'Legea morală și legea juridică coincid întotdeauna' / altă temă filozofică accesibilă.\n"
            f"  Cerință: Construiește un argument DEDUCTIV VALID în favoarea sau împotriva tezei,\n"
            f"  cu minimum 3 premize, identificând tipul de raționament folosit.\n"
            f"  Repere obligatorii:\n"
            f"    1. Formularea clară a tezei (pro/contra) (4p)\n"
            f"    2. Minimum 3 premize coerente și relevante (12p)\n"
            f"    3. Concluzia derivată logic din premize (6p)\n"
            f"    4. Contra-argument și respingere (8p)\n\n"
            f"10 puncte din oficiu.\n\n"
            f"IMPORTANT:\n"
            f"- Tabele de adevăr complete, nu parțiale\n"
            f"- Sofismele trebuie denumite corect (ad hominem, om de paie, apel la autoritate etc.)\n"
            f"- Argumentele trebuie să fie riguroase logic, nu doar convingătoare retoric\n\n"
            f"[[BAREM_BAC]]\n"
            f"SUBIECTUL I — A: [răspunsuri corecte] B: [tabele complete + justificări]\n"
            f"SUBIECTUL al II-lea: [analiză detaliată a fiecărei cerințe cu punctaj]\n"
            f"SUBIECTUL al III-lea: [structura argumentului model + punctaj detaliat]\n"
            f"[[/BAREM_BAC]]"
        )

    # ── FALLBACK generic ──
    else:
        subiecte = materie_info.get("subiecte", [])
        subiecte_str = ", ".join(subiecte) if subiecte else materie_label
        structura = materie_info.get("structura", {})
        structura_str = "\n".join(f"  {k}: {v}" for k, v in structura.items()) if structura else ""
        timp = materie_info.get("timp_minute", 180)
        return (
            f"Generează un subiect complet de BAC la {materie_label} ({profil}), "
            f"identic ca structură și dificultate cu subiectele oficiale din România.\n\n"
            f"Inspiră-te din tipare reale ale subiectelor BAC din 2021–2025.\n"
            f"STRUCTURĂ OBLIGATORIE:\n"
            f"- SUBIECTUL I (30 puncte): itemi obiectivi/semiobiectivi\n"
            f"- SUBIECTUL al II-lea (30 puncte): probleme/analiză structurată\n"
            f"- SUBIECTUL al III-lea (30 puncte): problemă complexă / eseu / sinteză\n"
            f"- 10 puncte din oficiu\n\n"
            f"TEME: {subiecte_str}\nTIMP: {timp} minute\n\n"
            f"[[BAREM_BAC]]\nSUBIECTUL I: [răspunsuri și punctaj]\nSUBIECTUL al II-lea: [soluții și punctaj]\nSUBIECTUL al III-lea: [criterii și punctaj]\n[[/BAREM_BAC]]"
        )


def get_bac_correction_prompt(materie_label, subiect, raspuns_elev, from_photo=False):
    source_note = (
        "NOTĂ: Răspunsul a fost extras automat dintr-o fotografie a lucrării. "
        "Unele cuvinte pot fi transcrise imperfect din cauza scrisului de mână — "
        "judecă după intenția elevului, nu după eventuale erori de OCR.\n\n"
        if from_photo else ""
    )

    # Reguli de limbaj adaptate materiei
    if "Română" in materie_label:
        lang_rules = (
            "CORECTARE LIMBĂ ROMÂNĂ (OBLIGATORIU — punctaj separat):\n"
            "- Ortografie și punctuație (virgule, punct, ghilimele «»)\n"
            "- Acordul gramatical (subiect-predicat, adjectiv-substantiv)\n"
            "- Folosirea corectă a cratimei, apostrofului\n"
            "- Exprimare clară, coerentă, fără pleonasme sau cacofonii\n"
            "- Registru stilistic adecvat eseului de BAC\n"
            "- Acordă până la 10 puncte bonus/penalizare pentru calitatea limbii\n\n"
        )
    else:
        lang_rules = (
            f"CORECTARE LIMBAJ ȘTIINȚIFIC ({materie_label}):\n"
            "- Terminologie specifică folosită corect\n"
            "- Notații și simboluri respectate (ex: m pentru masă, nu M; v nu V pentru viteză)\n"
            "- Unități de măsură scrise corect și complet\n"
            "- Formulele scrise corect, fără ambiguități\n"
            "- Raționament logic și coerent exprimat în cuvinte\n"
            "- Acordă până la 5 puncte bonus/penalizare pentru calitatea exprimării\n\n"
        )

    return (
        f"Ești examinator BAC România pentru {materie_label}.\n\n"
        f"{source_note}"
        f"SUBIECTUL:\n{subiect}\n\n"
        f"RĂSPUNSUL ELEVULUI:\n{raspuns_elev}\n\n"
        f"Corectează COMPLET în această ordine:\n\n"
        f"## 📊 Punctaj per subiect\n"
        f"- Subiectul I: X/30 puncte\n"
        f"- Subiectul II: X/30 puncte\n"
        f"- Subiectul III: X/30 puncte\n"
        f"- Din oficiu: 10 puncte\n\n"
        f"## ✅ Ce a făcut bine\n"
        f"[aspecte corecte]\n\n"
        f"## ❌ Greșeli și explicații\n"
        f"[fiecare greșeală explicată]\n\n"
        f"## 🖊️ Calitatea limbii și exprimării\n"
        f"{lang_rules}"
        f"## 🎓 Nota estimată (orientativă)\n"
        f"**Nota estimată: X/10** — [verdict scurt]. Precizează că e o estimare orientativă, nu nota oficială.\n\n"
        f"## 💡 Recomandări pentru BAC\n"
        f"[2-3 sfaturi concrete]\n\n"
        f"Fii constructiv, cald, dar riguros ca un examinator real."
    )


def parse_bac_subject(response):
    """Parsează răspunsul AI în subiect + barem.
    FIX bug 16: dacă AI-ul nu generează baremul în tags, căutăm secțiunea 'BAREM' în text."""
    barem = ""
    subject_text = response
    match = re.search(r"\[\[BAREM_BAC\]\](.*?)\[\[/BAREM_BAC\]\]", response, re.DOTALL)
    if match:
        barem = match.group(1).strip()
        subject_text = response[:match.start()].strip()
    else:
        # FIX bug 16: fallback — caută o secțiune de barem neîncadrată în tags
        # AI-ul uneori scrie "BAREM:" sau "## Barem" fără tag-uri
        barem_match = re.search(
            r'\n(?:##\s*)?(?:BAREM|Barem|barem)[:\s]+(.*)',
            response, re.DOTALL | re.IGNORECASE
        )
        if barem_match:
            barem = barem_match.group(1).strip()
            subject_text = response[:barem_match.start()].strip()
        # Dacă tot nu găsim barem, subject_text rămâne tot textul (comportament original)
    return subject_text, barem


def format_timer(seconds_remaining):
    h = seconds_remaining // 3600
    m = (seconds_remaining % 3600) // 60
    s = seconds_remaining % 60
    return f"{h:02d}:{m:02d}:{s:02d}"



# ══════════════════════════════════════════════════════════════════════════════
# ADMITERE FACULTATE — FMI București + UPB (ACS / ETTI)
# ══════════════════════════════════════════════════════════════════════════════

ADMITERE_CONFIG = {
    "🎓 FMI București": {
        "descriere": "Facultatea de Matematică și Informatică — Universitatea din București",
        "specializari": {
            "💻 Informatică": {
                "probe": [
                    {"cod": "fmi_info_matematica", "label": "Matematică (obligatorie)", "tip": "clasic",
                     "timp_minute": 180, "nr_intrebari": 0,
                     "descriere": "Algebră + Analiză matematică — nivel universitar an 1",
                     "structura": "3 probleme × 30p"},
                    {"cod": "fmi_info_informatica", "label": "Informatică (obligatorie)", "tip": "grila",
                     "timp_minute": 180, "nr_intrebari": 30,
                     "descriere": "30 întrebări grilă — algoritmi, structuri date, C/C++",
                     "structura": "30 grile × 3p"},
                ],
            },
            "📐 Matematică": {
                "probe": [
                    {"cod": "fmi_mate_matematica", "label": "Matematică (obligatorie)", "tip": "clasic",
                     "timp_minute": 180, "nr_intrebari": 0,
                     "descriere": "Algebră + Analiză matematică — nivel avansat",
                     "structura": "3 probleme × 30p"},
                    {"cod": "fmi_mate_informatica", "label": "Informatică (opțională)", "tip": "grila",
                     "timp_minute": 180, "nr_intrebari": 30,
                     "descriere": "Grilă informatică sau subiect suplimentar de matematică",
                     "structura": "30 grile × 3p"},
                ],
            },
        },
    },
    "🏛️ UPB — ACS (Automatică și Calculatoare)": {
        "descriere": "Facultatea de Automatică și Calculatoare — Universitatea Politehnica București",
        "specializari": {
            "⚙️ Calculatoare și Tehnologia Informației": {
                "probe": [
                    {"cod": "upb_acs_matematica", "label": "Matematică (obligatorie)", "tip": "clasic",
                     "timp_minute": 180, "nr_intrebari": 0,
                     "descriere": "Algebră și analiză matematică — nivel BAC avansat + intro universitar",
                     "structura": "Probleme structurate: matrice, funcții, derivate, integrale"},
                    {"cod": "upb_acs_informatica", "label": "Informatică (grilă)", "tip": "grila",
                     "timp_minute": 120, "nr_intrebari": 30,
                     "descriere": "Grilă — algoritmi, C/C++, structuri de date",
                     "structura": "30 grile × 3p"},
                    {"cod": "upb_acs_fizica", "label": "Fizică (grilă) — alternativă", "tip": "grila",
                     "timp_minute": 120, "nr_intrebari": 30,
                     "descriere": "Grilă fizică — mecanică, termodinamică, curent continuu",
                     "structura": "30 grile × 3p"},
                ],
            },
        },
    },
    "🏛️ UPB — ETTI (Electronică și Telecomunicații)": {
        "descriere": "Facultatea de Electronică, Telecomunicații și Tehnologia Informației — UPB",
        "specializari": {
            "📡 Electronică / Telecomunicații / Tehnologia informației": {
                "probe": [
                    {"cod": "upb_etti_matematica", "label": "Algebră și Analiză Matematică AAM (P1 — obligatorie)", "tip": "grila",
                     "timp_minute": 120, "nr_intrebari": 10,
                     "descriere": "10 grile × 9p — Algebră (matrice, ecuații, progresii, polinoame, legi compoziție) + Analiză (limite, derivate, integrale, funcții)",
                     "structura": "10 grile × 9p + 10p oficiu = 100p. P1 obligatorie — 40% din nota finală admitere.",
                     "date_reale": True},
                    {"cod": "upb_etti_fizica", "label": "Fizică F (P2 — la alegere)", "tip": "grila",
                     "timp_minute": 120, "nr_intrebari": 10,
                     "descriere": "10 grile × 9p — Mecanică, Termodinamică, Electricitate DC/AC, Optică. Accent pe electricitate și circuite.",
                     "structura": "10 grile × 9p + 10p oficiu = 100p. P2 la alegere (Fizică SAU Informatică) — 40% din nota finală admitere.",
                     "date_reale": True},
                    {"cod": "upb_etti_informatica", "label": "Informatică I (P2 — alternativă la Fizică)", "tip": "grila",
                     "timp_minute": 120, "nr_intrebari": 10,
                     "descriere": "10 grile × 9p — Algoritmi, structuri de date, C/C++, complexitate",
                     "structura": "10 grile × 9p + 10p oficiu = 100p. P2 alternativă — 40% din nota finală admitere."},
                ],
            },
        },
    },
}

# ── Date reale subiecte UPB ETTI (2021–2025) ──
# Sursa: http://www.physics.pub.ro/Admitere/subiecte.html
# Structură examen: 10 grile × 9p + 10p oficiu = 100p, timp 2h
# Formula admitere ETTI: 20% media cls IX-XI + 40% AAM (P1) + 40% Fizică/Info (P2)

UPB_ETTI_DATE_REALE = {
    "matematica": {
        # AAM = Algebră și Elemente de Analiză Matematică
        # Capitole recurente: ecuații/inecuații, progresii, matrice, polinoame,
        # legi de compoziție, limite, derivate, integrale, funcții, combinatorică
        2025: {
            "data": "14 iulie 2025",
            "varianta": "S",
            "intrebari": [
                {"nr": 1, "capitol": "ecuații", "enunt": "Soluția ecuației |2x+1| = 2|x-1| + 2 este..."},
                {"nr": 2, "capitol": "combinatorică", "enunt": "Numărul natural n astfel încât C(n,2) = 6"},
                {"nr": 3, "capitol": "progresii", "enunt": "Progresie aritmetică rație r=3, a3=7 → a5=?"},
                {"nr": 4, "capitol": "ecuații", "enunt": "Ecuație cu modul: 6/(0+1-2x)=0"},
                {"nr": 5, "capitol": "funcții/derivate", "enunt": "f(x)=e^x + x-2 → f'(1)=?"},
                {"nr": 6, "capitol": "limite", "enunt": "lim(x→2) (x²-3x+2)/(x-2)"},
                {"nr": 7, "capitol": "integrale/arie", "enunt": "Aria între graficul f(x)=x+1/x, asimptota oblică, x=2, x=3"},
                {"nr": 8, "capitol": "polinoame", "enunt": "f(X)=mX³-6 cu rădăcinile x1,x2,x3: dacă x1⁴+x2⁴+x3⁴=98 → x1·x2·x3=?"},
                {"nr": 9, "capitol": "ecuații/parte întreagă", "enunt": "Numărul soluțiilor reale ale ecuației cu paranteze întregi"},
                {"nr": 10, "capitol": "ecuații cu parametru", "enunt": "Valoarea lui m ∈ ℝ pentru care ecuația are infinitate de soluții"},
            ],
            "teme_frecvente": ["ecuații cu modul", "progresii", "limite", "derivate", "integrale", "polinoame", "legi compoziție"],
        },
    },
}


ADMITERE_NIVELE = {
    "🟢 Normal": {
        "label": "Normal",
        "descriere": "Subiecte identice cu admiterea reală — exemple din 2021, 2022, 2023, 2024, 2025",
        "instructiuni_grila": (
            "NIVEL NORMAL — identic cu subiectele reale UPB AAM din 2021-2025.\n"
            "Fiecare întrebare testează UN singur concept, calcule în 1-2 pași.\n\n"
            "Capitole și exemple EXACTE din subiecte reale verificate (2021-2025):\n\n"
            "ECUAȚII și INECUAȚII:\n"
            "- Ecuație cu modul: |2x+1|=2|x-1|+2 → soluție directă (2025 Q1)\n"
            "- Ecuație fracționară: 6/(1-2x)=0 → soluție directă (2025 Q4)\n"
            "- Ecuație exponențială: 2^(3x)=4 (2024 Q3), 9^x=81 (2023 Q5), 3^(x+1/2)=9 (2021 Q10)\n"
            "- Ecuație cu modul: |x+3|-1=... (2021 Q8), |x+1|=5|1-x| (2023 Q2)\n"
            "- Ecuație pătratică simplă: x²-6x+8=0 (2021 Q6), x²-7x+10=0 (2023 Q1)\n"
            "- Inecuație exponențială: 3^(x-1)<3^(x+1) (2021 Q5)\n"
            "- Sistem liniar 2×2: x+y=5, x-y=1 (2023 Q4)\n\n"
            "PROGRESII:\n"
            "- Progresie aritmetică: r=3, a₃=7 → a₅=? (2025 Q3)\n"
            "- Progresie aritmetică cu r și termen dat → al n-lea termen (2024 Q9)\n"
            "- Trei numere → verifică/găsește parametru pentru progresie (2021 Q3, 2023 Q6)\n\n"
            "COMBINATORICĂ:\n"
            "- C(n,2)=6 → n=? (2025 Q2)\n\n"
            "DERIVATE:\n"
            "- f'(1) pentru f(x)=eˣ+x-2 (2025 Q5)\n"
            "- f'(0) pentru f(x)=eˣ+x² (2021 Q4)\n"
            "- f'(1) pentru f(x)=3x⁴-x² (2023 Q3)\n\n"
            "LIMITE:\n"
            "- lim(x→2)(x²-3x+2)/(x-2) prin factorizare (2025 Q6)\n\n"
            "INTEGRALE:\n"
            "- Arie între grafic și asimptotă oblică pe interval [2,3] (2025 Q7)\n"
            "- Integrală definită: ∫₀¹ 1/(2(x²+1)) dx (2024 Q1)\n\n"
            "NU include: legi de compoziție complexe, matrice la puteri mari,\n"
            "ecuații cu 3 soluții reale distincte, integrale cu parametru, parte întreagă.\n"
            "Variantele greșite: erori de calcul frecvente (semn greșit, factor omis, confuzie formulă)."
        ),
        "instructiuni_clasic": (
            "NIVEL NORMAL — identic cu subiectele reale de admitere 2021-2025.\n"
            "Cerințe directe, pași de calcul clari, fără combinații neașteptate de concepte.\n"
            "Fiecare cerință a-d testează o singură tehnică standard."
        ),
    },
    "🟡 Mediu": {
        "label": "Mediu",
        "descriere": "Cele mai grele întrebări din 2021, 2022, 2023, 2024, 2025 — combinate",
        "instructiuni_grila": (
            "NIVEL MEDIU — stilul întrebărilor dificile din subiecte reale UPB AAM 2021-2025.\n"
            "Fiecare întrebare combină 2 concepte sau necesită un pas intermediar neevident.\n\n"
            "Exemple EXACTE din subiecte reale verificate (2021-2025):\n\n"
            "MATRICE:\n"
            "- Matrice A dată → suma modulelor elementelor de pe diagonala lui A^459\n"
            "  (2024 Q2 — necesită recunoașterea ciclicității puterilor matricei)\n\n"
            "ECUAȚII CU PARAMETRU:\n"
            "- 1-2x-2x²=me^x admite exact 3 soluții reale distincte → valorile lui m\n"
            "  (2024 Q5 — studiu grafic: intersecția dreptei y=m cu curba)\n"
            "- Ecuație cu parametru și condiție pe numărul de soluții (tip 2024)\n\n"
            "POLINOAME:\n"
            "- P(X)=aX^2024+bX^2023+X²+cX+3 div. prin (X²-1), rest la (X-1) este 3 → P(1)\n"
            "  (2023 Q7 — combină teorema împărțirii cu sistem de condiții)\n"
            "- f(X)=mX³-6, rădăcinile x₁,x₂,x₃: x₁⁴+x₂⁴+x₃⁴=98 → x₁·x₂·x₃\n"
            "  (2025 Q8 — Vieta + identități de putere Newton)\n\n"
            "LEGI DE COMPOZIȚIE (parte din programa AAM, apar în 2021 și 2023):\n"
            "- x*y = xy-2x-2y+10, suma soluțiilor ecuației x*x=x (2023 Q9)\n"
            "- x*y = xy-x-y+25, suma elementelor simetrizabile pe ℤ (2021 Q7)\n\n"
            "FUNCȚII CU PARAMETRU:\n"
            "- f(x)=(x²+ax+b)/(x²+1) cu 3 extreme locale + asimptotă oblică y=x-2\n"
            "  (2023 Q10 — combină derivare, studiu semn, condiție asimptotă)\n"
            "- f(x)=xˣ+aˣ+xᵃ, f(1)=1 → a=? (2024 Q7)\n\n"
            "INTEGRALE:\n"
            "- lim(x→0) [∫₀ˣ dt/(1+4t²+t⁴)] / (2x) — necesită derivata funcției integrale (2023 Q8)\n\n"
            "ECUAȚIE CU PARTE ÎNTREAGĂ:\n"
            "- Ecuație cu ⌊·⌋ → numărul soluțiilor reale (2025 Q9)\n\n"
            "COMBINATORICĂ AVANSATĂ:\n"
            "- Câte numere din {1,...,999} conțin cifra 9 cel puțin o dată (2021 Q9 — includere-excludere)\n\n"
            "Variantele greșite: rezultate obținute prin metode parțial corecte sau erori conceptuale subtile."
        ),
        "instructiuni_clasic": (
            "NIVEL MEDIU — stilul întrebărilor dificile din subiecte reale 2021-2025.\n"
            "Cerințele combină 2-3 concepte. Cel puțin o cerință necesită un artificiu neevident.\n"
            "Include legi de compoziție, polinoame cu condiții, funcții cu parametru multiplu."
        ),
    },
    "🔴 Avansat": {
        "label": "Avansat",
        "descriere": "Dincolo de 2021-2025 — inspirat din subiecte reale, nivel olimpiadă județeană",
        "instructiuni_grila": (
            "NIVEL AVANSAT — dincolo de subiectele reale UPB, aproape de olimpiadă județeană.\n"
            "Aplică stilul celor mai dificile întrebări din TOȚI anii 2021-2025 la toate 10 întrebările.\n"
            "Niciun exercițiu să nu fie rezolvabil în sub 3 pași de raționament.\n\n"
            "Exemple EXACTE din cele mai grele întrebări verificate (2021-2025):\n\n"
            "INTEGRALE CU PARAMETRU (cel mai greu tip apărut în subiecte reale):\n"
            "- f(x)=∫₁ˣ t(1-lnt)dt → abscisa punctului de maxim local\n"
            "  (2021 Q1 — derivata funcției integrale, semn f\'(x), tabel variație)\n"
            "- Variante: g(x)=∫₀ˣ t²·e^(-t)dt → puncte de inflexiune sau extreme\n\n"
            "LEGI DE COMPOZIȚIE COMPLEXE (apar în 2021 și 2023):\n"
            "- x*y=xy-x-y+25 pe ℤ → suma elementelor simetrizabile (2021 Q7)\n"
            "- x*y=xy-2x-2y+10 → ecuație iterată de n ori (extins față de 2023 Q9)\n"
            "- Variante: demonstrarea asociativității + element neutru + ecuație\n\n"
            "COMBINATORICĂ AVANSATĂ:\n"
            "- Câte numere din {1,...,999} conțin cifra 9 cel puțin o dată (2021 Q9 — răspuns: 271)\n"
            "- Variante: bijecții, partiții, numărare cu restricții multiple\n\n"
            "POLINOAME + VIETA + IDENTITĂȚI DE PUTERE:\n"
            "- f(X)=mX³-6, x₁⁴+x₂⁴+x₃⁴=98 → x₁·x₂·x₃ (2025 Q8 — Newton pₖ)\n"
            "- Variante: grad 4-5, sistem Vieta cu multiple condiții\n\n"
            "FUNCȚII CU PARAMETRU MULTIPLU:\n"
            "- f(x)=(x²+ax+b)/(x²+1): 3 extreme + asimptotă oblică (2023 Q10)\n"
            "- f(x)=xˣ+aˣ+xᵃ, f(1)=1 (2024 Q7 — cel mai ambiguu enunț din subiecte reale)\n\n"
            "ECUAȚII CU PARTE ÎNTREAGĂ:\n"
            "- Ecuație cu ⌊·⌋ și {·} combinate → număr soluții reale (2025 Q9)\n"
            "- Variante: x²+2⌊x⌋{x}+3{x}²=4 (tip olimpiadă)\n\n"
            "MATRICE LA PUTERI MARI:\n"
            "- A^459, A^2024 — ciclicitate la valori neevidente (extins față de 2024 Q2)\n\n"
            "Variantele greșite: rezultate plauzibile obținute prin aplicarea greșită a unor formule corecte."
        ),
        "instructiuni_clasic": (
            "NIVEL AVANSAT — dincolo de subiectele reale, aproape de olimpiadă județeană.\n"
            "Probleme cu parametri și discuție completă de cazuri. Raționament indirect.\n"
            "Include: integrale cu parametru, legi de compoziție complexe, combinatorică,\n"
            "studiu complet de funcție cu mai mulți parametri, identități de putere Vieta.\n"
            "Cel puțin un subpunct necesită o idee neconvențională sau un artificiu elegant."
        ),
    },
    "🔵 Preadmitere": {
        "label": "Preadmitere",
        "descriere": "Stil admitere anticipată (apr.) — materie cls IX-XI, mix ușor+greu, 2021-2025",
        "instructiuni_grila": (
            "NIVEL PREADMITERE — stil admitere anticipată UPB (sesiunea aprilie, elevi cls XI).\n"
            "MATERIE: algebră și analiză cls IX-XI. Programa include derivate și integrale definite.\n\n"
            "INTERZIS EXPLICIT:\n"
            "- Polinoame avansate (teorema împărțirii, Vieta, grad >2 cu mai mulți pași)\n"
            "- Legi de compoziție\n"
            "- Matrice la puteri mari (ciclicitate)\n"
            "- Ecuații cu parte întreagă complexe\n"
            "- Combinatorică avansată (includere-excludere, partiții)\n\n"
            "STRUCTURĂ OBLIGATORIE — 10 întrebări cu mix deliberat:\n\n"
            "Întrebările 1-6 (UȘOARE — un singur concept, calcul direct):\n"
            "- Q1: Ecuație cu modul simplă: |2x-3|=5\n"
            "- Q2: Progresie aritmetică sau geometrică cu 2 termeni dați → al n-lea termen\n"
            "- Q3: Derivată directă: f\'(a) pentru funcție elementară (polinom, eˣ, lnx, sinx, cosx)\n"
            "- Q4: Ecuație exponențială directă: a^f(x)=a^k\n"
            "- Q5: Sistem liniar 2×2 cu soluție întreagă\n"
            "- Q6: Integrală definită elementară: ∫ₐᵇ xⁿ dx sau ∫ eˣ dx sau ∫ sinx dx\n\n"
            "Întrebările 7-10 (GRELE — atenție și abilitate de observație):\n"
            "- Q7: Integrală cu parametru sau funcție integrală cu extrem local\n"
            "  (tip exact 2021 Q1: f(x)=∫₁ˣ t(1-lnt)dt → abscisa maximului local)\n"
            "- Q8: Funcție cu parametru și condiție compusă\n"
            "  (tip 2023 Q10: f(x)=(x²+ax+b)/(x²+1) cu asimptotă oblică dată)\n"
            "- Q9: Ecuație exponențială sau logaritmică cu substituție t=aˣ\n"
            "  sau studiu monotonie pentru număr soluții\n"
            "- Q10: Inecuație sau ecuație cu două module (cazuri neevidente)\n"
            "  (tip: |x²-x-2| > x+1)\n\n"
            "Variantele greșite: erori tipice de calcul. Valori numerice curate."
        ),
        "instructiuni_clasic": (
            "NIVEL PREADMITERE — stil admitere anticipată UPB, materie cls IX-XI.\n"
            "INTERZIS: polinoame avansate, legi de compoziție, matrice la puteri mari.\n"
            "Include derivate, integrale definite (inclusiv cu parametru), funcții cu parametru.\n"
            "Mix obligatoriu: 6 cerințe ușoare (calcul direct) + 4 cerințe grele\n"
            "(observație, mai mulți pași, parametru cu discuție, integrală nontrivială)."
        ),
    },
}


def get_admitere_prompt(proba_cod, proba_info, specializare, universitate, nivel_dificultate="🟢 Normal"):
    """
    Construiește prompt MINIMAL pentru AI — doar ce e strict necesar.
    Fără text redundant, context scurt, instrucțiuni dense.
    """
    univ_scurt = universitate.replace("🎓 ","").replace("🏛️ ","")
    spec_scurt = specializare.replace("💻 ","").replace("📐 ","").replace("⚙️ ","").replace("📡 ","")

    # Instrucțiuni de dificultate din dicționarul global
    _niv      = ADMITERE_NIVELE.get(nivel_dificultate, ADMITERE_NIVELE["🟢 Normal"])
    _dif_g    = _niv["instructiuni_grila"]   # pentru probe grilă
    _dif_c    = _niv["instructiuni_clasic"]  # pentru probe clasice (probleme)
    _niv_lbl  = _niv["label"]

    # ── IMPORTANT: verificăm cazurile SPECIFICE (etti, fmi) ÎNAINTE de cele generice ──
    # Altfel "matematica" in proba_cod prinde și upb_etti_matematica, fmi_mate_matematica etc.

    # ── UPB ETTI — Matematică (AAM) grilă 6 variante ──
    if "matematica" in proba_cod and "etti" in proba_cod:
        return (
            f"Generează un CHESTIONAR DE CONCURS pentru admitere UPB — ETTI, "
            f"disciplina Algebră și Elemente de Analiză Matematică AAM.\n\n"
            f"STRUCTURĂ OBLIGATORIE — identică cu subiectele reale UPB:\n"
            f"- 10 întrebări numerotate 1-10\n"
            f"- Fiecare întrebare: enunț matematic precis + 6 variante (a, b, c, d, e, f)\n"
            f"- O singură variantă corectă per întrebare\n"
            f"- Fiecare întrebare valorează 9 puncte; 10 puncte din oficiu → total 100p\n"
            f"- Timp: 2 ore\n\n"
            f"FORMAT STRICT per întrebare (respectă EXACT acest format):\n"
            f"[nr]. [Enunț matematic complet, precis, cu toate datele]. ([9 pct.])\n"
            f"a) [valoare/expresie]; b) [valoare/expresie]; c) [valoare/expresie]; "
            f"d) [valoare/expresie]; e) [valoare/expresie]; f) [valoare/expresie].\n\n"
            f"DISTRIBUȚIE CAPITOLE (respectă proporțiile din subiecte reale):\n"
            f"- Întrebările 1-2: Algebră — ecuații/inecuații (cu modul, exponențiale, "
            f"logaritmice, iraționale, parte întreagă) sau sisteme liniare\n"
            f"- Întrebările 3-4: Algebră — progresii aritmetice/geometrice, combinatorică "
            f"(permutări, aranjamente, combinări, binomul lui Newton)\n"
            f"- Întrebările 5-6: Algebră — matrice și determinanți, polinoame (teorema "
            f"împărțirii, rădăcini, divizibilitate), legi de compoziție (element neutru, simetric)\n"
            f"- Întrebările 7-8: Analiză matematică — limite (cu forme nedeterminate 0/0, ∞/∞), "
            f"derivate (reguli, extreme locale, monotonie), studiu de funcție\n"
            f"- Întrebările 9-10: Analiză matematică — integrale definite (calcul direct, arie, "
            f"schimbare variabilă), funcții cu parametru (extreme, asimptote, număr soluții)\n\n"
            f"REGULI DE STIL:\n"
            f"- Variantele greșite: erori tipice de calcul (semn greșit, factor omis, confuzie formulă) — plauzibile\n"
            f"- Variantele corecte distribuite aleatoriu între a-f\n"
            f"- Enunțuri riguroase matematic, notații standard românești: tg(x), ctg(x), lg(x), ln(x), f'(x), C(n,k)\n"
            f"- Valorile numerice să fie 'curate' (întregi, fracții simple, radicali simpli)\n\n"
            f"NIVEL DE DIFICULTATE — {_niv_lbl.upper()}:\n{_dif_g}\n\n"
            f"La final, pe o linie separată, scrie baremul în formatul:\n"
            f"[[BAREM]]1-[literă], 2-[literă], 3-[literă], 4-[literă], 5-[literă], "
            f"6-[literă], 7-[literă], 8-[literă], 9-[literă], 10-[literă][[/BAREM]]"
        )

    # ── UPB ETTI — Fizică (F) grilă 6 variante ──
    elif "fizica" in proba_cod and "etti" in proba_cod:
        return (
            f"Generează un CHESTIONAR DE CONCURS pentru admitere UPB — ETTI, disciplina Fizică F.\n\n"
            f"STRUCTURĂ OBLIGATORIE — identică cu subiectele reale UPB:\n"
            f"- 10 întrebări numerotate 1-10\n"
            f"- Fiecare întrebare: enunț cu date numerice concrete + 6 variante (a, b, c, d, e, f)\n"
            f"- O singură variantă corectă per întrebare\n"
            f"- Fiecare întrebare valorează 9 puncte; 10 puncte din oficiu → total 100p\n"
            f"- Timp: 2 ore\n\n"
            f"FORMAT STRICT per întrebare (respectă EXACT acest format):\n"
            f"[nr]. [Enunț complet cu date numerice, unități SI, context fizic concret]. ([9 pct.])\n"
            f"a) [valoare/expresie]; b) [valoare/expresie]; c) [valoare/expresie]; "
            f"d) [valoare/expresie]; e) [valoare/expresie]; f) [valoare/expresie].\n\n"
            f"DISTRIBUȚIE CAPITOLE (respectă proporțiile din subiecte reale):\n"
            f"- Întrebările 1-2: Termodinamică (gaze ideale, transformări izocore/izoterme/izobate, "
            f"principiul I, ciclu Carnot, motor termic, randament)\n"
            f"- Întrebările 3-4: Circuite electrice DC (Legea Ohm, rezistoare serie/paralel, "
            f"Kirchhoff, putere electrică, surse cu rezistență internă, ampermetre/voltmetre)\n"
            f"- Întrebările 5-6: Mecanică (cinematică MRU/MRUA, cădere liberă, lucru mecanic, "
            f"energie, impuls, scripeți, plan înclinat, resort)\n"
            f"- Întrebările 7-8: Circuite mixte / aplicații electrice complexe (circuite cu mai multe "
            f"surse, putere maximă, sarcină electrică, curent de scurtcircuit)\n"
            f"- Întrebările 9-10: Termodinamică avansată sau Mecanică avansată (gaz ideal cu relație "
            f"p-V non-standard, piston în cilindru, motor Otto/Carnot cu date numerice)\n\n"
            f"REGULI DE STIL:\n"
            f"- Date numerice curate: g=10 m/s², R=8,32 J/(mol·K), valori rotunde\n"
            f"- Variantele corecte distribuite aleatoriu între a-f (nu mereu a sau b)\n"
            f"- Enunțuri clare, fără ambiguități, ca în examene reale\n\n"
            f"NIVEL DE DIFICULTATE — {_niv_lbl.upper()}:\n{_dif_g}\n\n"
            f"La final, pe o linie separată, scrie baremul în formatul:\n"
            f"[[BAREM]]1-[literă], 2-[literă], 3-[literă], 4-[literă], 5-[literă], "
            f"6-[literă], 7-[literă], 8-[literă], 9-[literă], 10-[literă][[/BAREM]]"
        )

    # ── UPB ETTI — Informatică (I) grilă 4 variante ──
    elif "informatica" in proba_cod and "etti" in proba_cod:
        nr = proba_info.get("nr_intrebari", 30)
        return (
            f"Subiect admitere {univ_scurt} ({spec_scurt}) — Informatică grilă C/C++. {nr} întrebări × 3p.\n\n"
            f"FORMAT STRICT per întrebare:\nQ[n]. [enunț]\na)[var]\nb)[var]\nc)[var]\nd)[var]\nRĂSPUNS:[literă]\n\n"
            f"Distribuție: 6 complexitate, 6 recursivitate, 5 structuri date, 5 sortări/căutare, 4 grafuri, 4 DP+greedy.\n\n"
            f"NIVEL DE DIFICULTATE — {_niv_lbl.upper()}:\n{_dif_g}\n\n"
            f"[[BAREM]]1-x,2-x,...{nr}-x[[/BAREM]]"
        )

    # ── Cazuri GENERICE (fmi, upb_acs etc.) — după toate cazurile specifice ──

    elif "informatica" in proba_cod:
        nr = proba_info.get("nr_intrebari", 30)
        return (
            f"Subiect admitere {univ_scurt} ({spec_scurt}) — Informatică grilă C/C++. {nr} întrebări × 3p.\n\n"
            f"FORMAT STRICT per întrebare:\nQ[n]. [enunț]\na)[var]\nb)[var]\nc)[var]\nd)[var]\nRĂSPUNS:[literă]\n\n"
            f"Distribuție: 6 complexitate, 6 recursivitate, 5 structuri date, 5 sortări/căutare, 4 grafuri, 4 DP+greedy.\n\n"
            f"NIVEL DE DIFICULTATE — {_niv_lbl.upper()}:\n{_dif_g}\n\n"
            f"[[BAREM]]1-x,2-x,...{nr}-x[[/BAREM]]"
        )

    elif "matematica" in proba_cod:
        fmi = "fmi" in proba_cod
        nivel_mat = "analiză reală + algebră liniară, nivel an I universitate" if fmi else "algebră + analiză, nivel BAC avansat + intro universitar"
        return (
            f"Subiect admitere {univ_scurt} ({spec_scurt}) — Matematică. {nivel_mat}.\n"
            f"3 probleme × 30p + 10p oficiu = 100p. Timp: 3h.\n\n"
            f"P1-Algebră: sistem liniar cu parametru (compatibilitate, soluție generală) + matrice (det, inversă sau valori proprii). 4 cerințe a-d.\n"
            f"P2-Analiză: studiu funcție (limite, derivate, extreme, asimptote) + integrală definită sau arie. 4 cerințe a-d.\n"
            f"P3-Algebră abstractă sau Analiză avansată: spații vectoriale / baze / dim SAU șiruri / serii / convergență. 4 cerințe a-d.\n\n"
            f"NIVEL DE DIFICULTATE — {_niv_lbl.upper()}:\n{_dif_c}\n\n"
            f"[[BAREM]]P1:[soluție]\nP2:[soluție]\nP3:[soluție][[/BAREM]]"
        )

    elif "fizica" in proba_cod:
        nr = proba_info.get("nr_intrebari", 30)
        dist = "8 mecanică, 6 termodinamică, 8 electricitate+circuit DC, 4 optică, 4 fizică modernă"
        return (
            f"Subiect admitere {univ_scurt} ({spec_scurt}) — Fizică grilă. {nr} întrebări × 3p. Cls IX-X.\n\n"
            f"FORMAT STRICT per întrebare:\nQ[n]. [enunț cu date numerice]\na)[var]\nb)[var]\nc)[var]\nd)[var]\nRĂSPUNS:[literă]\n\n"
            f"Distribuție: {dist}.\n\n"
            f"NIVEL DE DIFICULTATE — {_niv_lbl.upper()}:\n{_dif_g}\n\n"
            f"[[BAREM]]1-x,2-x,...{nr}-x[[/BAREM]]"
        )

    return f"Subiect admitere {univ_scurt} — {proba_info['label']}. Structură standard, 3h, 100p."


def parse_grila_questions(text):
    """Parsează întrebările din subiectul generat.

    Suportă formate cu 6 variante (a-f) — UPB ETTI — și 4 variante (a-d).
    SVG-urile din enunțuri sunt protejate în timpul parsării și restaurate după.
    Baremul din [[BAREM]]..[[/BAREM]] are prioritate față de RĂSPUNS inline.
    """
    # re este deja importat la nivel de modul

    # ── Pasul 1: protejăm SVG-urile înainte de parsare ──
    # SVG-urile conțin "a)", "b)" etc. care ar încurca regex-urile de parsare a variantelor.
    svg_store = {}
    counter = [0]

    def _replace_svg(m):
        key = f"__SVG_{counter[0]}__"
        svg_store[key] = m.group(0)
        counter[0] += 1
        return key

    t = re.sub(r'\[\[DESEN_SVG\]\].*?\[\[/DESEN_SVG\]\]', _replace_svg, text, flags=re.DOTALL)
    t = re.sub(r'<svg\b[^>]*>.*?</svg\s*>', _replace_svg, t, flags=re.DOTALL | re.IGNORECASE)

    def _restore(s):
        for k, v in svg_store.items():
            s = s.replace(k, v)
        return s

    # ── Pasul 2: extragem baremul ──
    barem_map = {}
    barem_match = re.search(r'\[\[BAREM\]\](.*?)\[\[/BAREM\]\]', t, re.DOTALL)
    if barem_match:
        for item in re.split(r'[,;\n]+', barem_match.group(1).strip()):
            m = re.match(r'(\d+)\s*[-\u2013:]\s*([a-f])', item.strip(), re.IGNORECASE)
            if m:
                barem_map[int(m.group(1))] = m.group(2).lower()

    text_clean = re.sub(r'\[\[BAREM\]\].*?\[\[/BAREM\]\]', '', t, flags=re.DOTALL).strip()

    # ── Helper: extrage varianta 'lit)' dintr-un bloc ──
    def _get_var(block, lit):
        m = re.search(
            rf'(?:^|;|\n)\s*{lit}\)\s*(.*?)(?=\s*(?:[a-f]\)|;|\n\s*[a-f]\)|\Z))',
            block, re.IGNORECASE | re.DOTALL
        )
        if m:
            val = m.group(1).strip().rstrip(';').strip()
            val = re.sub(r'\s*\n\s*', ' ', val).strip()
            return val
        return ""

    questions = []

    # ── FORMAT 6 variante (a-f) — UPB ETTI ──
    has_6var = bool(re.search(r'\be\)\s*\S|\bf\)\s*\S', text_clean, re.IGNORECASE))

    if has_6var:
        blocks = re.split(r'(?=^\s*\d+[\.\.)]\s)', text_clean, flags=re.MULTILINE)
        for block in blocks:
            block = block.strip()
            if not block:
                continue
            nr_m = re.match(r'^(\d+)[\.\.)]\s*', block)
            if not nr_m:
                continue
            nr = int(nr_m.group(1))
            enunt_m = re.match(r'^\d+[\.\.)]\s*(.*?)(?=\s*(?:^|\n)\s*a\))', block, re.DOTALL | re.MULTILINE)
            if not enunt_m:
                continue
            enunt_raw = re.sub(r'\s*\n\s*', ' ', enunt_m.group(1)).strip()
            enunt_raw = re.sub(r'\s*\(\s*\d+\s*p(?:ct\.?)?\s*\)', '', enunt_raw).strip()
            enunt = _restore(enunt_raw)

            variante = {}
            for l in "abcdef":
                v = _get_var(block, l)
                variante[l] = _restore(v)

            if not (variante.get("a") and variante.get("b") and variante.get("c")):
                continue

            questions.append({
                "nr": nr,
                "enunt": enunt,
                "variante": variante,
                "nr_variante": 6,
                "raspuns_corect": barem_map.get(nr, ""),
                "has_svg": bool(svg_store and any(k in enunt for k in svg_store)),
            })

        if len(questions) >= 3:
            return sorted(questions, key=lambda q: q["nr"])

    # ── FORMAT Q prefix + RĂSPUNS inline ──
    questions = []
    pattern_q = re.compile(
        r'Q(\d+)\.\s*(.*?)\na\)\s*(.*?)\nb\)\s*(.*?)\nc\)\s*(.*?)\nd\)\s*(.*?)\nRĂ?SPUNS\s*:?\s*([a-d])',
        re.DOTALL | re.IGNORECASE
    )
    for m in pattern_q.finditer(text_clean):
        nr = int(m.group(1))
        enunt = _restore(m.group(2).strip())
        questions.append({
            "nr": nr,
            "enunt": enunt,
            "variante": {
                "a": _restore(m.group(3).strip()), "b": _restore(m.group(4).strip()),
                "c": _restore(m.group(5).strip()), "d": _restore(m.group(6).strip()),
            },
            "nr_variante": 4,
            "raspuns_corect": barem_map.get(nr, m.group(7).strip().lower()),
            "has_svg": bool(svg_store and any(k in enunt for k in svg_store)),
        })

    if len(questions) >= 3:
        return sorted(questions, key=lambda q: q["nr"])

    # ── FORMAT simplu: număr + 4 variante + RĂSPUNS opțional ──
    questions = []
    blocks = re.split(r'(?=^\s*\d+[\.\.)]\s)', text_clean, flags=re.MULTILINE)
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        nr_m = re.match(r'^(\d+)[\.\.)]\s*', block)
        if not nr_m:
            continue
        nr = int(nr_m.group(1))
        enunt_m = re.match(r'^\d+[\.\.)]\s*(.*?)(?=\s*(?:^|\n)\s*a\))', block, re.DOTALL | re.MULTILINE)
        if not enunt_m:
            continue
        enunt_raw = re.sub(r'\s*\n\s*', ' ', enunt_m.group(1)).strip()
        enunt = _restore(enunt_raw)

        variante = {}
        for l in "abcd":
            variante[l] = _restore(_get_var(block, l))

        if not (variante.get("a") and variante.get("b") and variante.get("c")):
            continue

        raspuns_m = re.search(r'RĂ?SPUNS\s*:?\s*([a-d])', block, re.IGNORECASE)
        raspuns = barem_map.get(nr, raspuns_m.group(1).lower() if raspuns_m else "")

        questions.append({
            "nr": nr,
            "enunt": enunt,
            "variante": variante,
            "nr_variante": 4,
            "raspuns_corect": raspuns,
            "has_svg": bool(svg_store and any(k in enunt for k in svg_store)),
        })

    return sorted(questions, key=lambda q: q["nr"]) if len(questions) >= 3 else []


def run_admitere_ui():
    st.subheader("🏛️ Admitere Facultate")

    if not st.session_state.get("admitere_active"):
        col1, col2 = st.columns(2)
        with col1:
            univ = st.selectbox("🏫 Universitate:", options=list(ADMITERE_CONFIG.keys()), key="adm_univ_sel")
            spec_options = list(ADMITERE_CONFIG[univ]["specializari"].keys())
            spec = st.selectbox("🎯 Specializare:", options=spec_options, key="adm_spec_sel")
        with col2:
            probe = ADMITERE_CONFIG[univ]["specializari"][spec]["probe"]
            proba_idx = st.selectbox("📋 Proba:", options=range(len(probe)),
                                      format_func=lambda i: probe[i]["label"], key="adm_proba_sel")
            proba = probe[proba_idx]
            if proba["tip"] == "grila":
                mod_grila = st.radio("📝 Mod rezolvare:", options=["🖱️ Grilă interactivă", "✍️ Răspuns liber"],
                                      key="adm_mod_grila", horizontal=True)
            else:
                mod_grila = "✍️ Răspuns liber"
                st.info("✍️ Probă cu rezolvare scrisă")

        st.markdown(
            f"<div style='background:linear-gradient(135deg,#f093fb22,#f5576c22);"
            f"border:1px solid #f093fb55;padding:14px 18px;border-radius:10px;margin:10px 0'>"
            f"<b>{univ}</b> · {spec}<br>"
            f"<span style='font-size:13px'>{proba['descriere']}</span><br>"
            f"<span style='font-size:12px;opacity:0.8'>⏱️ {proba['timp_minute']} min · {proba['structura']}</span>"
            f"</div>", unsafe_allow_html=True)

        # ── Selector nivel dificultate ──
        nivel_options = list(ADMITERE_NIVELE.keys())
        nivel_dificultate = st.radio(
            "🎯 Nivel de dificultate:",
            options=nivel_options,
            format_func=lambda k: f"{k}  —  {ADMITERE_NIVELE[k]['descriere']}",
            key="adm_nivel_dificultate",
            horizontal=False,
        )

        use_timer = st.checkbox(f"⏱️ Cronometru ({proba['timp_minute']} min)", value=True, key="adm_timer")
        st.divider()

        if st.button("🚀 Generează subiect", type="primary", use_container_width=True):
            with st.spinner("📝 Se generează subiectul de admitere..."):
                prompt = get_admitere_prompt(proba["cod"], proba, spec, univ, nivel_dificultate)
                full = "".join(run_chat_with_rotation([], [prompt],
                    system_prompt=get_system_prompt(materie=None, pas_cu_pas=False,
                        mod_avansat=True, mod_strategie=False, mod_bac_intensiv=False)))
                barem_m = re.search(r'\[\[BAREM\]\](.*?)\[\[/BAREM\]\]', full, re.DOTALL)
                barem = barem_m.group(1).strip() if barem_m else ""
                subject_text = full[:barem_m.start()].strip() if barem_m else full
                st.session_state.admitere_active = True
                st.session_state.admitere_univ = univ
                st.session_state.admitere_spec = spec
                st.session_state.admitere_proba = proba
                st.session_state.admitere_subject = subject_text
                st.session_state.admitere_barem = barem
                st.session_state.admitere_mod_grila = mod_grila
                st.session_state.admitere_nivel = nivel_dificultate
                st.session_state.admitere_corectat = False
                st.session_state.admitere_raspuns = ""
                if proba["tip"] == "grila" and "🖱️" in mod_grila:
                    parsed = parse_grila_questions(full)
                    if not parsed:
                        # Logging pentru debug — afișăm primele 500 caractere din răspunsul AI
                        _log(f"parse_grila_questions a returnat [] pentru răspuns:\n{full[:500]}", "silent")
                    st.session_state.admitere_grila_questions = parsed
                    st.session_state.admitere_grila_answers = {}
                    st.session_state.admitere_grila_submitted = False
                if use_timer:
                    st.session_state.admitere_start_time = time.time()
                    st.session_state.admitere_timp_min = proba["timp_minute"]
                st.rerun()

        if st.button("← Înapoi", key="adm_back_start"):
            st.session_state.admitere_mode = False
            st.rerun()
        return

    # ── ECRAN REZOLVARE ──
    univ      = st.session_state.admitere_univ
    spec      = st.session_state.admitere_spec
    proba     = st.session_state.admitere_proba
    subject   = st.session_state.admitere_subject
    barem     = st.session_state.admitere_barem
    mod_grila = st.session_state.get("admitere_mod_grila", "✍️ Răspuns liber")

    st.markdown(f"### {univ} · {spec}")
    _nivel_activ = st.session_state.get("admitere_nivel", "🟢 Normal")
    st.caption(f"📋 {proba['label']}  ·  {_nivel_activ}")

    # Cronometru
    start_time = st.session_state.get("admitere_start_time")
    if start_time:
        elapsed   = int(time.time() - start_time)
        total     = st.session_state.get("admitere_timp_min", 180) * 60
        remaining = max(0, total - elapsed)
        col_t1, col_t2 = st.columns([3, 1])
        with col_t2:
            color = "#e53e3e" if remaining < 600 else "#38a169"
            st.markdown(
                f"<div style='text-align:center;background:{color}22;border:1px solid {color}55;"
                f"border-radius:8px;padding:8px;font-size:20px;font-weight:bold;color:{color}'>"
                f"⏱️ {format_timer(remaining)}</div>", unsafe_allow_html=True)

    # ── MOD GRILĂ INTERACTIVĂ ──
    if proba["tip"] == "grila" and "🖱️" in mod_grila:
        questions = st.session_state.get("admitere_grila_questions", [])
        answers   = st.session_state.get("admitere_grila_answers", {})
        submitted = st.session_state.get("admitere_grila_submitted", False)
        if not questions:
            # Dacă parsarea a eșuat, afișăm subiectul raw ca fallback
            with st.expander("📄 Subiectul", expanded=True):
                st.markdown(subject)
            st.divider()
            st.warning("Nu s-au putut parsa întrebările. Încearcă modul 'Răspuns liber'.")
        else:
            # Grila interactivă — întrebările sunt afișate una câte una, nu subiectul raw
            nr_variante = questions[0].get("nr_variante", 4) if questions else 4
            puncte_per_intrebare = 9 if nr_variante == 6 else 3
            st.markdown(f"**{len(questions)} întrebări · {puncte_per_intrebare}p/întrebare · 10p oficiu**")
            for q in questions:
                nr = q["nr"]
                variante_keys = list(q["variante"].keys())  # a-f sau a-d
                st.markdown(f"**{nr}.**")
                render_message_with_svg(q["enunt"])
                if submitted:
                    corect = q["raspuns_corect"]
                    ales   = answers.get(nr, "")
                    for lit in variante_keys:
                        text = f"{lit}) {q['variante'][lit]}"
                        if lit == corect:
                            st.markdown(f"✅ **{text}**")
                        elif lit == ales and ales != corect:
                            st.markdown(f"❌ ~~{text}~~")
                        else:
                            st.markdown(f"&nbsp;&nbsp;{text}")
                else:
                    ales = st.radio(
                        f"Q{nr}",
                        options=variante_keys,
                        format_func=lambda l, q=q: f"{l}) {q['variante'][l]}",
                        key=f"adm_q_{nr}",
                        label_visibility="collapsed",
                        horizontal=True
                    )
                    answers[nr] = ales
                st.markdown("---")
            st.session_state.admitere_grila_answers = answers
            if not submitted:
                st.caption(f"Răspuns la {sum(1 for q in questions if q['nr'] in answers)}/{len(questions)} întrebări")
                if st.button("✅ Trimite grila", type="primary", use_container_width=True):
                    st.session_state.admitere_grila_submitted = True
                    st.rerun()
            else:
                corecte = sum(1 for q in questions if answers.get(q["nr"],"") == q["raspuns_corect"])
                scor_intrebari = corecte * puncte_per_intrebare
                scor_total = scor_intrebari + 10  # +10p oficiu
                nota = round(scor_total / 10, 2)
                culoare = "#38a169" if scor_total >= 50 else "#e53e3e"
                st.markdown(
                    f"<div style='background:linear-gradient(135deg,{culoare}33,{culoare}11);"
                    f"border:2px solid {culoare}88;padding:20px;border-radius:12px;text-align:center'>"
                    f"<h2 style='margin:0'>🎯 {corecte}/{len(questions)} corecte</h2>"
                    f"<h3 style='margin:8px 0 4px 0'>{scor_intrebari}p + 10p oficiu = "
                    f"<b>{scor_total}/100p</b></h3>"
                    f"<div style='font-size:22px;font-weight:bold'>Nota: {nota}/10</div>"
                    f"</div>", unsafe_allow_html=True)
                comercial.ui.nota_orientativa()

    # ── MOD RĂSPUNS LIBER ──
    else:
        with st.expander("📄 Subiectul", expanded=True):
            render_message_with_svg(subject)
        st.divider()
        if not st.session_state.admitere_corectat:
            raspuns = st.text_area("✍️ Scrie rezolvarea ta:", value=st.session_state.admitere_raspuns,
                                    height=300, key="adm_raspuns_input",
                                    placeholder="Scrie rezolvarea completă aici...")
            st.session_state.admitere_raspuns = raspuns
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔍 Corectează", type="primary", use_container_width=True,
                              disabled=not raspuns.strip()):
                    with st.spinner("🤖 Se corectează..."):
                        # Prompt minimal: doar ce AI-ul nu poate deduce singur
                        prompt_cor = (
                            f"Corectează admitere {univ.replace('🎓 ','').replace('🏛️ ','')} — {proba['label']}.\n"
                            f"SUBIECT:\n{subject}\n\n"
                            f"BAREM:\n{barem}\n\n"
                            f"RĂSPUNS CANDIDAT:\n{raspuns}\n\n"
                            f"Dă: punctaj per cerință, greșeli, nota/100, top 3 lacune de remediat."
                        )
                        corectare = "".join(run_chat_with_rotation([], [prompt_cor],
                            system_prompt=get_system_prompt(materie=None, pas_cu_pas=True,
                                mod_avansat=True, mod_strategie=False, mod_bac_intensiv=False)))
                        st.session_state.admitere_corectare = corectare
                        st.session_state.admitere_corectat = True
                        st.rerun()
            with col2:
                if barem:
                    with st.expander("📋 Barem"):
                        st.markdown(barem)
        else:
            st.success("✅ Corectare finalizată")
            comercial.ui.nota_orientativa()
            st.markdown(st.session_state.admitere_corectare)
            if barem:
                with st.expander("📋 Barem oficial"):
                    st.markdown(barem)

    st.divider()
    col_r, col_n = st.columns(2)
    with col_r:
        if st.button("🔄 Subiect nou", use_container_width=True):
            # Păstrăm univ/spec/proba/nivel — generăm un alt test pe aceeași selecție
            _saved_univ   = st.session_state.get("admitere_univ")
            _saved_spec   = st.session_state.get("admitere_spec")
            _saved_proba  = st.session_state.get("admitere_proba")
            _saved_nivel  = st.session_state.get("admitere_nivel", "🟢 Normal")
            for k in ["admitere_active","admitere_subject","admitere_barem","admitere_raspuns",
                      "admitere_corectat","admitere_corectare","admitere_grila_questions",
                      "admitere_grila_answers","admitere_grila_submitted","admitere_start_time"]:
                st.session_state.pop(k, None)
            if _saved_univ:
                st.session_state["adm_univ_sel"] = _saved_univ
            if _saved_spec:
                st.session_state["adm_spec_sel"] = _saved_spec
            if _saved_proba:
                try:
                    _probe = ADMITERE_CONFIG[_saved_univ]["specializari"][_saved_spec]["probe"]
                    _idx = next((i for i, p in enumerate(_probe) if p["cod"] == _saved_proba["cod"]), 0)
                    st.session_state["adm_proba_sel"] = _idx
                except Exception:
                    pass
            st.session_state["adm_nivel_dificultate"] = _saved_nivel
            st.rerun()
    with col_n:
        if st.button("← Înapoi la selecție", use_container_width=True):
            # Salvăm selecția curentă înainte de a reseta testul
            _saved_univ  = st.session_state.get("admitere_univ")
            _saved_spec  = st.session_state.get("admitere_spec")
            _saved_proba = st.session_state.get("admitere_proba")
            _saved_nivel = st.session_state.get("admitere_nivel", "🟢 Normal")
            for k in ["admitere_active","admitere_subject","admitere_barem","admitere_raspuns",
                      "admitere_corectat","admitere_corectare","admitere_grila_questions",
                      "admitere_grila_answers","admitere_grila_submitted","admitere_start_time"]:
                st.session_state.pop(k, None)
            # Restaurăm selecția în widget-urile de selectbox ca să nu se reseteze la prima valoare
            if _saved_univ and _saved_univ in list(ADMITERE_CONFIG.keys()):
                st.session_state["adm_univ_sel"] = _saved_univ
            if _saved_spec:
                st.session_state["adm_spec_sel"] = _saved_spec
            if _saved_proba:
                # Găsim indexul probei salvate în lista de probe a specializării curente
                try:
                    _probe = ADMITERE_CONFIG[_saved_univ]["specializari"][_saved_spec]["probe"]
                    _idx = next((i for i, p in enumerate(_probe) if p["cod"] == _saved_proba["cod"]), 0)
                    st.session_state["adm_proba_sel"] = _idx
                except Exception:
                    pass
            st.session_state["adm_nivel_dificultate"] = _saved_nivel
            st.rerun()



def run_bac_sim_ui():
    st.subheader("🎓 Simulare BAC")

    # ── ECRAN DE START ──
    if not st.session_state.get("bac_active"):
        # ── Selector ierarhic: Filieră → Profil → Specializare → Materie ──
        st.markdown("#### 🎓 Alege profilul tău")
        col1, col2 = st.columns(2)
        with col1:
            filiere = list(PROFILE_BAC.keys())
            bac_filiera = st.selectbox("🏫 Filiera:", options=filiere, key="bac_filiera_sel")

            profile_in_filiera = list(PROFILE_BAC[bac_filiera].keys())
            bac_profil_grup = st.selectbox("📂 Profil:", options=profile_in_filiera, key="bac_profil_grup_sel")

        with col2:
            specializari = list(PROFILE_BAC[bac_filiera][bac_profil_grup].keys())
            bac_specializare = st.selectbox("🎯 Specializare:", options=specializari, key="bac_spec_sel")

            spec_info = PROFILE_BAC[bac_filiera][bac_profil_grup][bac_specializare]
            obligatorii = spec_info["materii_obligatorii"]
            optionale   = spec_info["materii_optionale"]
            toate = obligatorii + optionale
            disponibile   = [m for m in toate if m in MATERII_SIMULARE_DISPONIBILE]
            indisponibile = [m for m in toate if m not in MATERII_SIMULARE_DISPONIBILE]

            if disponibile:
                def _label(m):
                    return f"{m}  ✦ obligatorie" if m in obligatorii else m
                bac_materie = st.selectbox(
                    "📚 Materia de simulat:",
                    options=disponibile,
                    format_func=_label,
                    key="bac_mat_sel"
                )
            else:
                st.warning("⚠️ Nicio materie cu suport AI pentru această specializare.")
                bac_materie = None

        # Info materii indisponibile
        if indisponibile:
            st.caption(f"ℹ️ Fără simulare AI (în curând): {', '.join(indisponibile)}")

        # Info card specializare
        st.markdown(
            f"<div style='background:linear-gradient(135deg,#667eea22,#764ba222);"
            f"border:1px solid #667eea44;padding:12px 18px;border-radius:10px;margin:8px 0 4px 0'>"
            f"<b>{bac_specializare}</b> — {spec_info['descriere']}"
            f"</div>",
            unsafe_allow_html=True
        )

        if bac_materie is None:
            return

        info = MATERII_BAC.get(bac_materie, {})
        if not info:
            st.error(f"Materia '{bac_materie}' nu are configurație în sistem.")
            return

        bac_profil = bac_specializare  # profilul transmis la prompt = specializarea aleasă

        # ── Selector tip fizică: afișat DOAR dacă materia selectată este generic "Fizică"
        # (adică vine din PROFILE_BAC cu un cod generic, nu dacă utilizatorul a ales
        # deja explicit "🔬 Fizică real" sau "⚡ Fizică tehnolog" din dropdown)
        # FIX Bug #3: condiția anterioară afișa radio-ul și când materia era deja specifică
        _fizica_deja_specifica = bac_materie in ("🔬 Fizică real", "⚡ Fizică tehnolog")
        if info.get("cod") in ("fizica_real", "fizica_tehnologic") and not _fizica_deja_specifica:
            tip_fizica = st.radio(
                "🔬 Tip Fizică:",
                options=["🔬 Fizică profil real (teoretic/militar)", "⚡ Fizică profil tehnologic"],
                key="bac_fizica_tip",
                horizontal=True,
            )
            # Suprascrie materia și info în funcție de alegere
            if "real" in tip_fizica:
                bac_materie = "🔬 Fizică real"
                info = MATERII_BAC.get("🔬 Fizică real", info)
            else:
                bac_materie = "⚡ Fizică tehnolog"
                info = MATERII_BAC.get("⚡ Fizică tehnolog", MATERII_BAC.get("⚡ Fizică tehnologică", info))
        # Dacă materia are profile proprii (ex. Chimie: anorganică/organică, Bio: vegetală/anatomie)
        elif len(info.get("profile", [])) > 1:
            bac_profil = st.selectbox(
                f"📋 Varianta {bac_materie}:",
                options=info["profile"],
                key="bac_prof_var_sel"
            )

        use_timer = st.checkbox(f"⏱️ Cronometru ({info.get('timp_minute', 180)} min)", value=True, key="bac_timer")

        # Info card — diferit pt materii cu date reale vs fără
        if info.get("date_reale"):
            structura = info.get("structura", {})
            structura_html = "".join(f"<li><b>{k}:</b> {v}</li>" for k, v in structura.items())
            st.markdown(
                "<div style='background:linear-gradient(135deg,#11998e,#38ef7d);"
                "color:white;padding:18px 22px;border-radius:12px;margin:12px 0'>"
                "<h4 style='margin:0 0 8px 0'>✅ Subiecte bazate pe tipare reale BAC 2021–2025</h4>"
                f"<ul style='margin:0;padding-left:18px;line-height:1.9'>{structura_html}</ul>"
                "<p style='margin:10px 0 0 0;font-size:13px;opacity:0.9'>"
                "⏱️ 3 ore · 100 puncte (90p scrise + 10p oficiu)</p>"
                "</div>",
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                "<div style='background:linear-gradient(135deg,#667eea,#764ba2);"
                "color:white;padding:18px 22px;border-radius:12px;margin:12px 0'>"
                "<h4 style='margin:0 0 8px 0'>📋 Subiect generat de AI</h4>"
                "<ul style='margin:0;padding-left:18px;line-height:1.8'>"
                "<li>Structură inspirată din modelele BAC oficiale</li>"
                "<li>Rezolvi în timp real cu cronometru opțional</li>"
                "<li>Primești corectare AI detaliată + barem</li>"
                "</ul></div>",
                unsafe_allow_html=True
            )

        st.divider()
        col_s, col_b = st.columns(2)
        with col_s:
            btn_lbl = "🚀 Generează subiect AI"
            if st.button(btn_lbl, type="primary", use_container_width=True):
                with st.spinner("📝 Se generează subiectul BAC..."):
                    prompt = get_bac_prompt_ai(bac_materie, info, bac_profil)
                    full = "".join(run_chat_with_rotation(
                        [], [prompt],
                        system_prompt=get_system_prompt(
                            materie=None,
                            pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                            mod_avansat=st.session_state.get("mod_avansat", False),
                            mod_strategie=st.session_state.get("mod_strategie", False),
                            mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                        )
                    ))
                subject_text, barem = parse_bac_subject(full)


                st.session_state.update({
                    "bac_active": True,
                    "bac_materie": bac_materie,
                    "bac_profil": bac_profil,
                    "bac_subject": subject_text,
                    "bac_barem": barem,
                    "bac_raspuns": "",
                    "bac_corectat": False,
                    "bac_corectare": "",
                    "bac_start_time": time.time() if use_timer else None,
                    "bac_timp_min": info["timp_minute"],
                    "bac_use_timer": use_timer,
                })
                st.rerun()
        with col_b:
            if st.button("↩️ Înapoi la chat", use_container_width=True):
                st.session_state.pop("bac_mode", None)
                st.rerun()
        return

    # ── SIMULARE ACTIVĂ ──
    # ✅ FIX: Check if bac_materie and bac_profil are set before using them
    if not st.session_state.get("bac_materie") or not st.session_state.get("bac_profil"):
        st.error("❌ Eroare: Subiectul nu a fost generat corect. Reîncarcă pagina.")
        st.stop()
    
    col_title, col_timer = st.columns([3, 1])
    with col_title:
        st.markdown(f"### {st.session_state.bac_materie} · {st.session_state.bac_profil}")
    with col_timer:
        if st.session_state.get("bac_use_timer") and st.session_state.get("bac_start_time"):
            elapsed = int(time.time() - st.session_state.bac_start_time)
            total   = st.session_state.bac_timp_min * 60
            left    = max(0, total - elapsed)
            pct     = left / total
            color   = "#2ecc71" if pct > 0.5 else ("#e67e22" if pct > 0.2 else "#e74c3c")
            if left > 0:
                # Cronometrul numără în browser (JS doar în iframe-ul lui — fără reîncărcarea paginii).
                # Înainte: window.parent.location.reload() la fiecare secundă — blocat în Chrome/Edge
                # (iframe izolat), iar unde mergea ștergea textul nesalvat din răspuns.
                components.html(
                    f"""<div id="t" style="background:{color};color:white;padding:8px 12px;border-radius:8px;
                    text-align:center;font:700 20px sans-serif">⏱️ {format_timer(left)}</div>
                    <script>
                    var end = Date.now() + {left} * 1000, total = {total};
                    function pad(n) {{ return (n < 10 ? "0" : "") + n; }}
                    function tick() {{
                        var s = Math.max(0, Math.round((end - Date.now()) / 1000)), p = s / total;
                        var el = document.getElementById("t");
                        el.style.background = p > 0.5 ? "#2ecc71" : (p > 0.2 ? "#e67e22" : "#e74c3c");
                        el.textContent = "⏱️ " + pad(Math.floor(s / 3600)) + ":" + pad(Math.floor(s % 3600 / 60)) + ":" + pad(s % 60);
                        if (s > 0) setTimeout(tick, 1000);
                    }}
                    tick();
                    </script>""",
                    height=52,
                )
            else:
                st.markdown(
                    f'<div style="background:{color};color:white;padding:8px 12px;'
                    f'border-radius:8px;text-align:center;font-size:20px;font-weight:700">'
                    f'⏱️ {format_timer(left)}</div>',
                    unsafe_allow_html=True
                )
            if left == 0:
                st.warning("⏰ Timpul a expirat!")
                # FIX bug 15: la expirarea timpului, trimite automat răspunsul curent
                # dacă elevul nu a trimis deja și există un răspuns scris
                if (
                    not st.session_state.get("bac_corectat")
                    and not st.session_state.get("bac_timer_submitted")
                    and st.session_state.get("bac_raspuns", "").strip()
                ):
                    st.session_state["bac_timer_submitted"] = True
                    with st.spinner("⏰ Timp expirat — se corectează automat..."):
                        _prompt_timeout = get_bac_correction_prompt(
                            st.session_state.bac_materie,
                            st.session_state.bac_subject,
                            st.session_state.bac_raspuns,
                            from_photo=st.session_state.get("bac_from_photo", False),
                        )
                        _corectare_timeout = "".join(run_chat_with_rotation(
                            [], [_prompt_timeout],
                            system_prompt=get_system_prompt(
                                materie=MATERII.get(st.session_state.bac_materie),
                                pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                                mod_avansat=st.session_state.get("mod_avansat", False),
                                mod_strategie=st.session_state.get("mod_strategie", False),
                                mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                            )
                        ))
                    st.session_state.bac_corectare = _corectare_timeout
                    st.session_state.bac_corectat  = True
                    st.rerun()
            elif left > 0 and not st.session_state.get("bac_corectat"):
                # La expirare rulăm din nou aplicația (fragmentul verifică la 10 s), ca ramura
                # „Timpul a expirat” de mai sus să trimită automat lucrarea la corectat.
                _bac_end = st.session_state.bac_start_time + total

                @st.fragment(run_every=10)
                def _bac_timer_watch():
                    if time.time() >= _bac_end and not st.session_state.get("bac_corectat"):
                        st.rerun()

                _bac_timer_watch()

    st.divider()

    with st.expander("📋 Subiectul", expanded=not st.session_state.bac_corectat):
        # ✅ FIX: Use render_message_with_svg to properly display SVG drawings!
        # Previously used st.markdown which doesn't render SVG, just shows raw code
        render_message_with_svg(st.session_state.bac_subject)

    if not st.session_state.bac_corectat:
        st.markdown("### ✏️ Răspunsurile tale")

        tab_foto, tab_text = st.tabs(["📷 Fotografiază lucrarea", "⌨️ Scrie manual"])

        raspuns = st.session_state.get("bac_raspuns", "")
        from_photo = False

        # ── TAB FOTO ──
        with tab_foto:
            st.info(
                "📱 **Pe telefon:** apasă butonul de mai jos și fotografiază lucrarea.\n\n"
                "💻 **Pe calculator:** încarcă o poză din galerie.\n\n"
                "AI-ul va citi textul și va porni corectarea automat."
            )
            uploaded_photo = st.file_uploader(
                "Încarcă fotografia lucrării:",
                type=["jpg", "jpeg", "png", "webp", "heic"],
                key="bac_photo_upload",
                help="Fă o poză clară, cu lumină bună, la lucrarea scrisă de mână."
            )

            if uploaded_photo:
                # FIX Bug #4: citim bytes-urile ÎNAINTE de st.image() pentru a evita
                # stream-ul consumat — st.image() poate avansa cursorul intern al fișierului,
                # astfel încât uploaded_photo.read() ulterior ar returna b"" (bytes goale).
                img_bytes = uploaded_photo.read()
                st.image(img_bytes, caption="Fotografia încărcată", use_container_width=True)

                if not st.session_state.get("bac_ocr_done"):
                    with st.spinner("🔍 Profesorul citește lucrarea..."):
                        text_extras = extract_text_from_photo(img_bytes, st.session_state.bac_materie,
                                                              mime_type=uploaded_photo.type or "image/jpeg")
                    st.session_state.bac_raspuns  = text_extras
                    st.session_state.bac_ocr_done = True
                    st.session_state.bac_from_photo = True

                    # Pornește corectura automat
                    with st.spinner("📊 Se corectează lucrarea..."):
                        prompt = get_bac_correction_prompt(
                            st.session_state.bac_materie,
                            st.session_state.bac_subject,
                            text_extras,
                            from_photo=True
                        )
                        corectare = "".join(run_chat_with_rotation(
                            [], [prompt],
                            system_prompt=get_system_prompt(
                                materie=MATERII.get(st.session_state.bac_materie),
                                pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                                mod_avansat=st.session_state.get("mod_avansat", False),
                                mod_strategie=st.session_state.get("mod_strategie", False),
                                mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                            )
                        ))
                    st.session_state.bac_corectare = corectare
                    st.session_state.bac_corectat  = True
                    st.rerun()

                if st.session_state.get("bac_ocr_done"):
                    with st.expander("📄 Text extras din poză", expanded=False):
                        st.text(st.session_state.get("bac_raspuns", ""))

        # ── TAB TEXT ──
        with tab_text:
            # Valoarea doar prin session_state (vezi tema) — fără avertismentul value= + key=
            if "bac_ans_input" not in st.session_state:
                st.session_state["bac_ans_input"] = st.session_state.get("bac_raspuns", "")
            raspuns = st.text_area(
                "Scrie rezolvarea completă:",
                height=350,
                placeholder="Subiectul I:\n1. ...\n2. ...\n\nSubiectul II:\n...\n\nSubiectul III:\n...",
                key="bac_ans_input"
            )
            st.session_state.bac_raspuns = raspuns
            st.session_state.bac_from_photo = False

            if st.button("🤖 Corectare AI", type="primary", use_container_width=True,
                         disabled=not raspuns.strip()):
                with st.spinner("📊 Se corectează lucrarea..."):
                    prompt = get_bac_correction_prompt(
                        st.session_state.bac_materie,
                        st.session_state.bac_subject,
                        raspuns,
                        from_photo=False
                    )
                    corectare = "".join(run_chat_with_rotation(
                        [], [prompt],
                        system_prompt=get_system_prompt(
                            materie=MATERII.get(st.session_state.bac_materie),
                            pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                            mod_avansat=st.session_state.get("mod_avansat", False),
                            mod_strategie=st.session_state.get("mod_strategie", False),
                            mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                        )
                    ))
                st.session_state.bac_corectare = corectare
                st.session_state.bac_corectat  = True
                st.rerun()

        st.divider()
        col_barem, col_nou = st.columns(2)
        with col_barem:
            if st.session_state.get("bac_barem"):
                if st.button("📋 Arată Baremul", use_container_width=True):
                    st.session_state.bac_show_barem = not st.session_state.get("bac_show_barem", False)
                    st.rerun()
        with col_nou:
            if st.button("🔄 Subiect nou", use_container_width=True):
                _bac_filiera = st.session_state.get("bac_filiera_sel")
                _bac_profil  = st.session_state.get("bac_profil_grup_sel")
                _bac_spec    = st.session_state.get("bac_spec_sel")
                _bac_mat     = st.session_state.get("bac_mat_sel")
                for k in [k for k in list(st.session_state.keys()) if k.startswith("bac_")]:
                    st.session_state.pop(k, None)
                if _bac_filiera: st.session_state["bac_filiera_sel"]    = _bac_filiera
                if _bac_profil:  st.session_state["bac_profil_grup_sel"] = _bac_profil
                if _bac_spec:    st.session_state["bac_spec_sel"]        = _bac_spec
                if _bac_mat:     st.session_state["bac_mat_sel"]         = _bac_mat
                st.rerun()

        if st.session_state.get("bac_show_barem") and st.session_state.get("bac_barem"):
            with st.expander("📋 Barem de corectare", expanded=True):
                st.markdown(st.session_state.bac_barem)

    else:
        st.markdown("### 📊 Corectare AI")
        comercial.ui.nota_orientativa()
        st.markdown(st.session_state.bac_corectare)
        if st.session_state.get("bac_barem"):
            with st.expander("📋 Barem"):
                st.markdown(st.session_state.bac_barem)
        st.divider()
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🔄 Subiect nou", type="primary", use_container_width=True):
                _bac_filiera = st.session_state.get("bac_filiera_sel")
                _bac_profil  = st.session_state.get("bac_profil_grup_sel")
                _bac_spec    = st.session_state.get("bac_spec_sel")
                _bac_mat     = st.session_state.get("bac_mat_sel")
                for k in [k for k in list(st.session_state.keys()) if k.startswith("bac_")]:
                    st.session_state.pop(k, None)
                if _bac_filiera: st.session_state["bac_filiera_sel"]    = _bac_filiera
                if _bac_profil:  st.session_state["bac_profil_grup_sel"] = _bac_profil
                if _bac_spec:    st.session_state["bac_spec_sel"]        = _bac_spec
                if _bac_mat:     st.session_state["bac_mat_sel"]         = _bac_mat
                st.rerun()
        with col2:
            if st.button("✏️ Reîncerc același subiect", use_container_width=True):
                st.session_state.bac_corectat  = False
                st.session_state.bac_corectare = ""
                st.session_state.bac_raspuns   = ""
                if st.session_state.get("bac_use_timer"):
                    st.session_state.bac_start_time = time.time()
                st.rerun()
        with col3:
            if st.button("💬 Înapoi la chat", use_container_width=True):
                for k in [k for k in list(st.session_state.keys()) if k.startswith("bac_")]:
                    st.session_state.pop(k, None)
                st.session_state.pop("bac_mode", None)
                st.rerun()


# ============================================================
# === CORECTARE TEME ===
# ============================================================

def get_homework_correction_prompt(materie_label: str, text_tema: str, from_photo: bool = False) -> str:
    source_note = (
        "NOTĂ: Tema a fost extrasă dintr-o fotografie. "
        "Unele cuvinte pot fi transcrise imperfect — judecă după intenția elevului.\n\n"
        if from_photo else ""
    )

    if "Română" in materie_label:
        corectare_limba = (
            "## 🖊️ Corectare limbă și stil\n"
            "Acordă atenție specială:\n"
            "- **Ortografie**: diacritice (ă,â,î,ș,ț), cratimă, apostrof\n"
            "- **Punctuație**: virgulă, punct, linie de dialog, ghilimele «»\n"
            "- **Acord gramatical**: subiect-predicat, adjectiv-substantiv, pronume\n"
            "- **Exprimare**: cacofonii, pleonasme, tautologii, registru stilistic\n"
            "- **Coerență**: logica textului, legătura dintre idei\n"
            "Subliniază greșelile găsite și explică regula corectă.\n\n"
        )
    else:
        corectare_limba = (
            f"## 🖊️ Limbaj și exprimare ({materie_label})\n"
            "- Terminologie specifică folosită corect\n"
            "- Notații, simboluri și unități de măsură corecte\n"
            "- Raționament exprimat clar și logic\n\n"
        )

    return (
        f"Ești profesor de {materie_label} și corectezi tema unui elev de liceu.\n\n"
        f"{source_note}"
        f"TEMA ELEVULUI:\n{text_tema}\n\n"
        f"Corectează complet și constructiv:\n\n"
        f"## ✅ Ce a făcut bine\n"
        f"[aspecte corecte — fii specific, nu generic]\n\n"
        f"## ❌ Greșeli de conținut\n"
        f"[fiecare greșeală de materie explicată, cu varianta corectă]\n\n"
        f"{corectare_limba}"
        f"## 📊 Notă orientativă\n"
        f"**Nota estimată: X/10** — [justificare scurtă]. Precizează că e o estimare orientativă.\n\n"
        f"## 💡 Sfaturi pentru data viitoare\n"
        f"[2-3 recomandări concrete și aplicabile]\n\n"
        f"Ton: cald, constructiv, ca un profesor care vrea să ajute, nu să descurajeze."
    )


def run_homework_ui():
    st.subheader("📚 Corectare Temă")

    if not st.session_state.get("hw_done"):
        col1, col2 = st.columns([2, 1])
        with col1:
            hw_materie = st.selectbox(
                "📚 Materia temei:",
                options=[m for m in MATERII.keys() if m != "🤖 Automat"],
                key="hw_materie_sel"
            )
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            st.caption("Profesorul se adaptează materiei.")

        st.divider()

        tab_foto, tab_text = st.tabs(["📷 Fotografiază tema", "⌨️ Scrie / lipește textul"])

        with tab_foto:
            st.info(
                "📱 **Pe telefon:** fotografiază caietul sau foaia de temă.\n\n"
                "💻 **Pe calculator:** încarcă o poză din galerie.\n\n"
                "Profesorul va citi și corecta automat."
            )
            hw_photo = st.file_uploader(
                "Încarcă fotografia temei:",
                type=["jpg", "jpeg", "png", "webp", "heic"],
                key="hw_photo_upload",
                help="Asigură-te că poza e clară și bine luminată."
            )

            if hw_photo and not st.session_state.get("hw_ocr_done"):
                st.image(hw_photo, caption="Fotografia încărcată", use_container_width=True)
                with st.spinner("🔍 Profesorul citește tema..."):
                    text_extras = extract_text_from_photo(hw_photo.getvalue(), hw_materie,
                                                          mime_type=hw_photo.type or "image/jpeg")
                st.session_state.hw_text       = text_extras
                st.session_state.hw_ocr_done   = True
                st.session_state.hw_from_photo = True
                st.session_state.hw_materie    = hw_materie
                with st.spinner("📝 Se corectează tema..."):
                    prompt = get_homework_correction_prompt(hw_materie, text_extras, from_photo=True)
                    corectare = "".join(run_chat_with_rotation(
                        [], [prompt],
                        system_prompt=get_system_prompt(
                            materie=MATERII.get(hw_materie),
                            pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                            mod_avansat=st.session_state.get("mod_avansat", False),
                            mod_strategie=st.session_state.get("mod_strategie", False),
                            mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                        )
                    ))
                st.session_state.hw_corectare = corectare
                st.session_state.hw_done      = True
                st.rerun()
            elif hw_photo and st.session_state.get("hw_ocr_done"):
                with st.expander("📄 Text extras din poză", expanded=False):
                    st.text(st.session_state.get("hw_text", ""))

        with tab_text:
            # Valoarea se pune prin session_state (nu și prin value=), altfel Streamlit afișează
            # avertismentul „created with a default value but also had its value set via Session State”.
            if "hw_text_input" not in st.session_state:
                st.session_state["hw_text_input"] = st.session_state.get("hw_text", "")
            hw_text = st.text_area(
                "Lipește sau scrie textul temei:",
                height=300,
                placeholder="Scrie sau lipește tema aici...",
                key="hw_text_input"
            )
            st.session_state.hw_text = hw_text
            if st.button("📝 Corectează tema", type="primary",
                         use_container_width=True, disabled=not hw_text.strip()):
                st.session_state.hw_materie    = hw_materie
                st.session_state.hw_from_photo = False
                with st.spinner("📝 Se corectează tema..."):
                    prompt = get_homework_correction_prompt(hw_materie, hw_text, from_photo=False)
                    corectare = "".join(run_chat_with_rotation(
                        [], [prompt],
                        system_prompt=get_system_prompt(
                            materie=MATERII.get(hw_materie),
                            pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                            mod_avansat=st.session_state.get("mod_avansat", False),
                            mod_strategie=st.session_state.get("mod_strategie", False),
                            mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                        )
                    ))
                st.session_state.hw_corectare = corectare
                st.session_state.hw_done      = True
                st.rerun()

    else:
        mat = st.session_state.get("hw_materie", "")
        src = "📷 din fotografie" if st.session_state.get("hw_from_photo") else "✏️ scrisă manual"
        st.caption(f"{mat} · temă {src}")
        if st.session_state.get("hw_from_photo") and st.session_state.get("hw_text"):
            with st.expander("📄 Text extras din poză", expanded=False):
                st.text(st.session_state.hw_text)
        comercial.ui.nota_orientativa()
        st.markdown(st.session_state.hw_corectare)
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📚 Corectează altă temă", type="primary", use_container_width=True):
                _hw_mat = st.session_state.get("hw_materie_sel")
                for k in [k for k in list(st.session_state.keys()) if k.startswith("hw_")]:
                    st.session_state.pop(k, None)
                if _hw_mat: st.session_state["hw_materie_sel"] = _hw_mat
                st.rerun()
        with col2:
            if st.button("💬 Înapoi la chat", use_container_width=True):
                for k in [k for k in list(st.session_state.keys()) if k.startswith("hw_")]:
                    st.session_state.pop(k, None)
                st.session_state.pop("homework_mode", None)
                st.rerun()


# === MOD QUIZ ===
NIVELE_QUIZ = ["🟢 Ușor (gimnaziu)", "🟡 Mediu (liceu)", "🔴 Greu (BAC)"]

MATERII_QUIZ = [m for m in list(MATERII.keys()) if m != "🤖 Automat"]


def get_quiz_prompt(materie_label: str, nivel: str, materie_val: str) -> str:
    """Generează prompt pentru crearea unui quiz."""
    nivel_text = nivel.split(" ", 1)[1].strip("()")
    return f"""Generează un quiz de 5 întrebări la {materie_label} pentru nivel {nivel_text}.

REGULI STRICTE:
1. Generează EXACT 5 întrebări numerotate (1. 2. 3. 4. 5.)
2. Fiecare întrebare are 4 variante de răspuns: A) B) C) D)
3. La finalul TUTUROR întrebărilor adaugă un bloc special cu răspunsurile corecte:

[[RASPUNSURI_CORECTE]]
1: X
2: X
3: X
4: X
5: X
[[/RASPUNSURI_CORECTE]]

unde X este A, B, C sau D.
4. Întrebările trebuie să fie clare și potrivite pentru nivel {nivel_text}.
5. Folosește LaTeX ($...$) pentru formule matematice.
6. NU da explicații acum — doar întrebările și răspunsurile corecte la final."""


def parse_quiz_response(response: str) -> tuple[str, dict]:
    """Extrage intrebarile si raspunsurile corecte din raspunsul AI.

    FIX: Gestioneaza corect cazurile cand AI-ul nu respecta exact delimitatorii:
    - Delimitatori lipsa: fallback prin cautarea unui bloc de raspunsuri
    - Formate variate: '1: A', '1. A', '1) A', '**1**: A'
    - Raspunsuri cu text extra: '1: A) text' -> extrage doar litera
    """
    correct = {}
    clean_response = response

    # Incearca mai intai delimitatorii exacti
    match = re.search(r'\[\[RASPUNSURI_CORECTE\]\](.*?)\[\[/RASPUNSURI_CORECTE\]\]',
                      response, re.DOTALL)

    # FIX: Fallback — AI-ul uneori omite delimitatorii sau ii scrie diferit
    if not match:
        match = re.search(
            r'(?:raspunsuri\s*corecte|raspunsuri\s*corecte|answers?)[:\s]*\n'
            r'((?:\s*\d+\s*[:.)-]\s*[A-D].*\n?){3,})',
            response, re.IGNORECASE | re.DOTALL
        )

    if match:
        block_start = match.start()
        clean_response = response[:block_start].strip()
        raw_block = match.group(1) if match.lastindex and match.lastindex >= 1 else match.group(0)

        for line in raw_block.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            # FIX: accepta formate: '1: A', '1. A', '1) A', '**1**: A', '1: A) text...'
            # FIX: regex mai strict — maxim 1 cifra pentru nr intrebare (evita "11: A" etc.)
            m = re.match(r'\*{0,2}(\d{1,2})\*{0,2}\s*[:.)-]\s*\*{0,2}([A-D])\b', line, re.IGNORECASE)
            if m:
                try:
                    q_num = int(m.group(1))
                    ans = m.group(2).upper()
                    correct[q_num] = ans
                except ValueError:
                    pass

    # FIX: Daca tot nu avem raspunsuri, incearca extractie din textul intreg
    if not correct:
        for m in re.finditer(
            r'(?:intrebarea|intrebarea|question)?\s*(\d+).*?'
            r'r[a]spuns(?:ul)?\s*(?:corect)?\s*[:\s]+([A-D])\b',
            response, re.IGNORECASE
        ):
            try:
                q_num = int(m.group(1))
                ans = m.group(2).upper()
                if 1 <= q_num <= 10:
                    correct[q_num] = ans
            except ValueError:
                pass

    return clean_response, correct


def evaluate_quiz(user_answers: dict, correct_answers: dict) -> tuple[int, str]:
    """Evaluează răspunsurile și returnează (scor, feedback_text)."""
    score = sum(1 for q, a in user_answers.items() if correct_answers.get(q) == a)
    total = len(correct_answers)

    lines = []
    for q in sorted(correct_answers.keys()):
        user_ans = user_answers.get(q, "—")
        correct_ans = correct_answers[q]
        if user_ans == correct_ans:
            lines.append(f"✅ **Întrebarea {q}**: {user_ans} — Corect!")
        else:
            lines.append(f"❌ **Întrebarea {q}**: ai răspuns **{user_ans}**, corect era **{correct_ans}**")

    if score == total:
        verdict = "🏆 Excelent! Nota 10!"
    elif score >= total * 0.8:
        verdict = "🌟 Foarte bine!"
    elif score >= total * 0.6:
        verdict = "👍 Bine, mai exersează puțin!"
    elif score >= total * 0.4:
        verdict = "📚 Trebuie să mai studiezi."
    else:
        verdict = "💪 Nu-ți face griji, încearcă din nou!"

    feedback = f"### Rezultat: {score}/{total} — {verdict}\n\n" + "\n\n".join(lines)
    return score, feedback


def run_quiz_ui():
    """Randează UI-ul pentru modul Quiz."""
    st.subheader("📝 Mod Examinare")

    # --- Setup quiz ---
    if not st.session_state.get("quiz_active"):
        col1, col2 = st.columns(2)
        with col1:
            quiz_materie_label = st.selectbox(
                "Materie:",
                options=MATERII_QUIZ,
                key="quiz_materie_select"
            )
        with col2:
            quiz_nivel = st.selectbox(
                "Nivel:",
                options=NIVELE_QUIZ,
                key="quiz_nivel_select"
            )

        if st.button("🚀 Generează Quiz", type="primary", use_container_width=True):
            quiz_materie_val = MATERII[quiz_materie_label]
            with st.spinner("📝 Profesorul pregătește întrebările..."):
                prompt = get_quiz_prompt(quiz_materie_label, quiz_nivel, quiz_materie_val)
                full_resp = ""
                for chunk in run_chat_with_rotation(
                    [], [prompt],
                    system_prompt=get_system_prompt(
                        materie=quiz_materie_val,
                        pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                        mod_avansat=st.session_state.get("mod_avansat", False),
                        mod_strategie=st.session_state.get("mod_strategie", False),
                        mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                    )
                ):
                    full_resp += chunk

            questions_text, correct = parse_quiz_response(full_resp)
            if len(correct) >= 3:
                st.session_state.quiz_active = True
                st.session_state.quiz_questions = questions_text
                st.session_state.quiz_correct = correct
                st.session_state.quiz_answers = {}
                st.session_state.quiz_submitted = False
                st.session_state.quiz_materie = quiz_materie_label
                st.session_state.quiz_nivel = quiz_nivel
                st.rerun()
            else:
                st.error("❌ Nu am putut genera quiz-ul. Încearcă din nou.")
        return

    # --- Quiz activ ---
    st.caption(f"📚 {st.session_state.quiz_materie} · {st.session_state.quiz_nivel}")

    # Afișează întrebările
    st.markdown(st.session_state.quiz_questions)
    st.divider()

    if not st.session_state.quiz_submitted:
        st.markdown("**Alege răspunsurile tale:**")
        answers = {}
        for q_num in sorted(st.session_state.quiz_correct.keys()):
            answers[q_num] = st.radio(
                f"Întrebarea {q_num}:",
                options=["A", "B", "C", "D"],
                horizontal=True,
                key=f"quiz_ans_{q_num}",
                index=None
            )

        all_answered = all(v is not None for v in answers.values())

        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Trimite răspunsurile", type="primary",
                         disabled=not all_answered, use_container_width=True):
                st.session_state.quiz_answers = {k: v for k, v in answers.items() if v}
                st.session_state.quiz_submitted = True
                st.rerun()
        with col2:
            if st.button("🔄 Quiz nou", use_container_width=True):
                _quiz_mat   = st.session_state.get("quiz_materie_select")
                _quiz_nivel = st.session_state.get("quiz_nivel_select")
                for k in ["quiz_active", "quiz_questions", "quiz_correct",
                          "quiz_answers", "quiz_submitted"]:
                    st.session_state.pop(k, None)
                if _quiz_mat:   st.session_state["quiz_materie_select"] = _quiz_mat
                if _quiz_nivel: st.session_state["quiz_nivel_select"]   = _quiz_nivel
                st.rerun()
    else:
        # Afișează rezultatele
        score, feedback = evaluate_quiz(
            st.session_state.quiz_answers,
            st.session_state.quiz_correct
        )
        st.markdown(feedback)
        st.divider()

        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Quiz nou", type="primary", use_container_width=True):
                _quiz_mat   = st.session_state.get("quiz_materie_select")
                _quiz_nivel = st.session_state.get("quiz_nivel_select")
                for k in ["quiz_active", "quiz_questions", "quiz_correct",
                          "quiz_answers", "quiz_submitted"]:
                    st.session_state.pop(k, None)
                if _quiz_mat:   st.session_state["quiz_materie_select"] = _quiz_mat
                if _quiz_nivel: st.session_state["quiz_nivel_select"]   = _quiz_nivel
                st.rerun()
        with col2:
            if st.button("💬 Înapoi la chat", use_container_width=True):
                for k in ["quiz_active", "quiz_questions", "quiz_correct",
                          "quiz_answers", "quiz_submitted", "quiz_mode"]:
                    st.session_state.pop(k, None)
                st.rerun()



# ============================================================
# === CONTEXT CACHING — Gemini API ===
# ============================================================
# System prompt-ul are ~21.000 tokeni. Fără caching, fiecare mesaj
# trimite toți acești tokeni = cost ridicat. Cu caching, platim o
# singură dată per sesiune și apoi mult mai puțin pentru tokenii cached.
#
# Cerințe Gemini API Context Caching (sursa: ai.google.dev/gemini-api/docs/pricing):
#   - Minim 1.024 tokeni în cache (system prompt-ul nostru e ~21k, OK)
#   - TTL minim 1 minut, maxim 1 oră (folosim 10 minute)
#   - Dacă gemini-3.1-flash-lite nu suportă caching pe cheia curentă, apelul eșuează
#     silențios și codul face fallback automat la apel normal fără caching (vezi except
#     din _get_or_create_cache) — deci caching-ul e un bonus, nu o cerință obligatorie.
#   → Folosim GEMINI_MODEL (gemini-3.1-flash-lite) ca model principal (caching + fallback)
#
# Cache key: hash(system_prompt + api_key) → unic per prompt + cheie

# Stocare cache: {cache_key: {"name": "cachedContents/...", "expires_at": float}}
# FIX: stocat în st.session_state în loc de variabilă globală de modul —
# Streamlit re-execută întregul script la fiecare rerun, deci o variabilă globală
# se resetează la {} la fiecare interacțiune, anulând complet beneficiile caching-ului.
_CACHE_TTL_SECONDS = 600          # 10 minute TTL (bine sub limita de 1 oră)
_CACHE_REFRESH_AT  = 480          # Reîmprospătăm la 8 minute (2 min înainte de expirare)
_CACHE_MIN_TOKENS  = 1024         # Minim tokeni pentru caching (limita Gemini)
# Prețuri: https://ai.google.dev/gemini-api/docs/pricing (sept. 2026)
# gemini-3.1-flash-lite: $0.25/$1.50 per 1M tokens normal, cached input disponibil
_CACHE_MODEL       = GEMINI_MODEL  # Model principal cu caching


def _get_prompt_hash(prompt_text: str, api_key: str) -> str:
    """Generează un hash scurt unic pentru (prompt, cheie) — folosit ca cache key local."""
    return hashlib.sha256(f"{api_key}:{prompt_text}".encode()).hexdigest()[:16]


def _get_or_create_cache(client: "genai.Client", prompt_text: str, api_key: str) -> str | None:
    """Returnează numele unui CachedContent valid, sau None dacă caching eșuează.

    Logică:
      1. Verifică dacă avem un cache valid în st.session_state["_prompt_cache_store"]
      2. Dacă nu (sau expirat), creează unul nou via API
      3. La orice eroare → returnează None (apelantul face fallback fără caching)
    Curățare: la fiecare apel elimină intrările expirate din dicționar (anti memory leak).
    """
    # FIX: folosim session_state în loc de variabilă globală — supraviețuiește reruns Streamlit
    cache_store = st.session_state.setdefault("_prompt_cache_store", {})

    cache_key = _get_prompt_hash(prompt_text, api_key)
    now = time.time()

    # Curăță intrările expirate — O(n) dar n e mic (1 intrare per cheie API × prompt)
    st.session_state["_prompt_cache_store"] = {
        k: v for k, v in cache_store.items()
        if v.get("expires_at", 0) > now
    }
    cache_store = st.session_state["_prompt_cache_store"]

    # 1. Verifică cache-ul existent
    existing = cache_store.get(cache_key)
    if existing and (existing["expires_at"] - now) > (_CACHE_TTL_SECONDS - _CACHE_REFRESH_AT):
        return existing["name"]

    # 2. Creează cache nou
    try:
        cached = client.caches.create(
            model=_CACHE_MODEL,
            config=genai_types.CreateCachedContentConfig(
                system_instruction=prompt_text,
                ttl=f"{_CACHE_TTL_SECONDS}s",
            ),
        )
        st.session_state["_prompt_cache_store"][cache_key] = {
            "name": cached.name,
            "expires_at": now + _CACHE_TTL_SECONDS,
            "api_key_prefix": api_key[:8],
        }
        return cached.name
    except Exception as e:
        # Caching poate eșua dacă: prompt prea scurt, model incompatibil,
        # cheie fără permisiuni etc. → fallback silențios la apel normal
        _log(f"Context caching indisponibil (fallback fără caching): {e}", "silent")
        return None


def _invalidate_cache_for_key(api_key: str) -> None:
    """Invalidează toate intrările din cache pentru o cheie API dată.
    Apelat când cheia e rotită (invalidă/epuizată) sau promptul se schimbă.
    """
    # FIX: folosim session_state în loc de variabilă globală
    cache_store = st.session_state.get("_prompt_cache_store", {})
    prefix = api_key[:8]
    st.session_state["_prompt_cache_store"] = {
        k: v for k, v in cache_store.items()
        if v.get("api_key_prefix") != prefix
    }


def _reupload_gfile(client, old_gfile):
    """Re-urcă un fișier atașat pe proiectul cheii curente.

    Fișierele din Google Files API aparțin PROIECTULUI cheii care le-a încărcat: după o rotație
    spre o cheie din alt proiect, URI-ul vechi dă 403. Octeții originali sunt păstrați în
    st.session_state["_gsrc_map"] (doar fișiere ≤ 25 MB). Returnează fișierul nou sau None.
    """
    srcs = st.session_state.get("_gsrc_map") or {}
    src = srcs.get(getattr(old_gfile, "name", ""))
    if not src:
        return None
    data, mime = src
    suffix = {"image/jpeg": ".jpg", "image/jpg": ".jpg", "image/png": ".png", "image/webp": ".webp",
              "image/gif": ".gif", "application/pdf": ".pdf"}.get(mime, ".bin")
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        new = client.files.upload(file=tmp_path, config=genai_types.UploadFileConfig(mime_type=mime))
        poll = 0
        while str(new.state) in ("FileState.PROCESSING", "PROCESSING") and poll < 60:
            time.sleep(1)
            new = client.files.get(name=new.name)
            poll += 1
        if not _is_gfile_active(new):
            return None
        srcs[new.name] = (data, mime)                # păstrăm și intrarea veche (retry cu payload vechi)
        for key in [x for x in st.session_state.keys() if str(x).startswith("_gfile_")]:
            if getattr(st.session_state.get(key), "name", None) == old_gfile.name:
                st.session_state[key] = new            # reruns viitoare folosesc fișierul nou
        return new
    except Exception:
        return None
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _ensure_gfile_for_client(client, gfile):
    """Fișierul e vizibil cu cheia curentă? Da -> îl păstrăm (aceleași proiecte). Nu -> îl re-urcăm."""
    try:
        cur = client.files.get(name=gfile.name)
        if _is_gfile_active(cur):
            return gfile
    except Exception:
        pass
    return _reupload_gfile(client, gfile) or gfile


# === LANȚ DE MODELE + MEMORIA COTELOR EPUIZATE ===
# Cotele gratuite Gemini sunt per PROIECT Google și per MODEL (AI Studio → Rate limit, nov. 2026):
#   gemini-3.1-flash-lite / gemini-3.5-flash-lite : 15 RPM · 250K TPM · 500 RPD  (cele mai generoase)
#   gemini-3.6-flash / gemini-3-flash-preview / gemini-3.8-flash / gemini-2.5-flash : 5 RPM · 20 RPD
#   gemini-2.5-flash-lite : 10 RPM · 20 RPD
# Ordinea = cota zilnică mare întâi, apoi rezervele mici. Fiecare combinație (cheie, model)
# care primește 429 e ținută minte (comun pentru TOȚI elevii, cât trăiește serverul):
#   - limită pe zi  → sărită până la resetare (miezul nopții Pacific ≈ 10:00 ora României)
#   - limită pe minut / supraîncărcare → sărită ~1 minut
#   - model inexistent (404) → sărit 6 ore; cheie invalidă → sărită 1 oră pe toate modelele
# Astfel un mesaj nu mai pierde timp încercând combinații epuizate: alege direct prima liberă.
MODEL_CHAIN = [
    GEMINI_MODEL,               # gemini-3.1-flash-lite — principal (+ context caching)
    "gemini-3.5-flash-lite",    # 500 RPD
    "gemini-3.6-flash",         # 20 RPD
    "gemini-3-flash-preview",   # 20 RPD
    "gemini-3.8-flash",         # 20 RPD
    "gemini-2.5-flash",         # 20 RPD (gândire prin thinking_budget)
    "gemini-2.5-flash-lite",    # 20 RPD
]
MODEL_FALLBACKS_NO_CACHE = MODEL_CHAIN   # nume vechi, păstrat pentru compatibilitate


@st.cache_resource
def _get_quota_state() -> dict:
    return {"lock": threading.Lock(), "blocked": {}}   # (amprentă_cheie | "*", model | "*") -> până_la


def _key_fp(api_key: str) -> str:
    return hashlib.sha256(api_key.encode("utf-8")).hexdigest()[:12]


def _next_quota_reset() -> float:
    """Timestamp-ul următoarei resetări a cotelor zilnice (miezul nopții America/Los_Angeles)."""
    from datetime import datetime, timedelta, timezone
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("America/Los_Angeles")
        now = datetime.now(tz)
        nxt = (now + timedelta(days=1)).replace(hour=0, minute=0, second=5, microsecond=0)
        return nxt.timestamp()
    except Exception:
        now = datetime.now(timezone.utc)
        return (now + timedelta(hours=24)).timestamp()


def _quota_block(api_key: str | None, model: str | None, until: float):
    st_ = _get_quota_state()
    k = (_key_fp(api_key) if api_key else "*", model or "*")
    with st_["lock"]:
        st_["blocked"][k] = max(until, st_["blocked"].get(k, 0))


def _quota_blocked_until(api_key: str, model: str) -> float:
    st_ = _get_quota_state()
    now = time.time()
    fp = _key_fp(api_key)
    with st_["lock"]:
        b = st_["blocked"]
        for k in [k for k, v in b.items() if v <= now]:   # curățăm intrările expirate
            del b[k]
        return max(b.get((fp, model), 0), b.get((fp, "*"), 0), b.get(("*", model), 0))


def _quota_until_from_error(msg: str) -> float:
    """Cât timp ocolim combinația după un 429: până la resetarea zilnică sau câteva secunde/minute."""
    low = msg.lower()
    if "perday" in low or "per_day" in low or "per day" in low or "requests_per_day" in low or "rpd" in low:
        return _next_quota_reset()
    m = re.search(r"retry(?:_delay|delay)?[^0-9]{0,20}([0-9]+(?:\.[0-9]+)?)\s*s", low)
    wait = float(m.group(1)) if m else 60.0
    return time.time() + min(max(wait, 5.0), 300.0)


def _pick_key_and_model(chain: list) -> tuple[int, int] | None:
    """Prima combinație liberă: cheia curentă întâi (toate modelele), apoi celelalte chei."""
    n = len(keys)
    start = st.session_state.get("key_index", 0) % max(n, 1)
    for off in range(n):
        ki = (start + off) % n
        for mi, m in enumerate(chain):
            if _quota_blocked_until(keys[ki], m) <= time.time():
                return ki, mi
    return None


def quota_status_text() -> str:
    """Rezumat pentru panoul de diagnostic: câte combinații cheie×model sunt epuizate acum."""
    total = len(keys) * len(MODEL_CHAIN)
    free = sum(1 for k in keys for m in MODEL_CHAIN if _quota_blocked_until(k, m) <= time.time())
    return f"{free}/{total} combinații cheie×model disponibile"


def _thinking_config_for(model_name: str):
    """Gemini 3.x: thinking_level; Gemini 2.5: thinking_budget (nu acceptă thinking_level)."""
    high = st.session_state.get("mod_gandire_extinsa", False)
    if model_name.startswith("gemini-2.5"):
        return genai_types.ThinkingConfig(thinking_budget=-1 if high else 0)
    return genai_types.ThinkingConfig(thinking_level="high" if high else "low")


def _student_profile_context() -> str:
    """Profilul elevului (din cont), adăugat la system prompt ca explicațiile să fie pe nivelul lui.
    Puține combinații posibile (clasă × specializare) → context caching-ul rămâne eficient."""
    u = current_user() or {}
    clasa, fil, prof, spec = u.get("clasa"), u.get("filiera"), u.get("profil_liceu"), u.get("specializare")
    if not any([clasa, fil, prof, spec]):
        return ""
    lines = ["", "", "PROFILUL ELEVULUI (folosește-l ca să adaptezi explicațiile; nu-l mai întreba):"]
    if clasa:
        lines.append(f"- Clasa: a {clasa}-a")
    if fil:
        lines.append(f"- Filiera: {fil}")
    if prof:
        lines.append(f"- Profilul: {prof}")
    if spec:
        lines.append(f"- Specializarea: {spec}")
        _d = globals().get("PROFILE_BAC", {}).get(fil or "", {}).get(prof or "", {}).get(spec, {})
        if isinstance(_d, dict):
            if _d.get("materii_obligatorii"):
                lines.append("- Materii obligatorii la BAC: " + ", ".join(_d["materii_obligatorii"]))
            if _d.get("materii_optionale"):
                lines.append("- Materii la alegere la BAC: " + ", ".join(_d["materii_optionale"]))
    return "\n".join(lines)


def run_chat_with_rotation(history_obj, payload, system_prompt=None, live_stream=False, use_cache=True):
    """Rulează chat cu rotație automată a cheilor API, lanț de modele de rezervă și context caching.

    La fiecare încercare alege prima combinație (cheie, model) care nu e marcată ca epuizată
    (vezi _get_quota_state). Un 429 marchează combinația și trece imediat la următoarea:
    alt model pe aceeași cheie, apoi cheia următoare — fără să mai reîncerce ce e deja epuizat.

    live_stream=False: răspunsul e colectat integral și emis la final (reîncercare transparentă).
    live_stream=True: bucățile sunt emise imediat; după prima bucată nu se mai reîncearcă
        (ar dubla textul) — vezi _salvage_partial_response().
    use_cache=False: fără context caching (ex. OCR, unde promptul e scurt).
    """
    MODEL_WITH_CACHE = _CACHE_MODEL
    chain = MODEL_CHAIN

    if not keys:
        raise Exception(
            "Nicio cheie API Gemini configurată. "
            "Adaugă cel puțin o cheie în st.secrets['GOOGLE_API_KEYS'] sau introdu-o manual în sidebar."
        )

    # Versiunea comercială: cont, acord parental, abonament și limita zilnică (ridică AccessError)
    comercial.billing.consume(current_user())

    active_prompt = (system_prompt or st.session_state.get("system_prompt") or SYSTEM_PROMPT) \
        + _student_profile_context()
    max_retries = len(keys) * len(chain) + 6
    _payload_list = list(payload) if isinstance(payload, list) else [payload]
    _has_gfile = any((not isinstance(p, str)) and hasattr(p, "uri") for p in _payload_list)
    _file_refresh_pending = False
    _file_refreshes = 0
    _deadline = time.time() + 60
    _use_cache = use_cache and st.session_state.get("_ctx_cache_enabled", True)
    _last_kind = None   # "quota" | "overload" | "invalid" | "missing"

    for attempt in range(max_retries):
        pick = _pick_key_and_model(chain)
        if pick is None:
            if _last_kind == "overload":
                raise Exception("Serviciul AI este supraîncărcat. Te rugăm să încerci din nou în câteva secunde. 🐢")
            raise Exception(
                "Toate cheile API sunt epuizate sau invalide pentru azi. "
                "Cotele se resetează în jurul orei 10:00 (ora României). 🔑"
            )
        ki, mi = pick
        if ki != st.session_state.get("key_index", 0):
            st.session_state.key_index = ki
            if _has_gfile:
                _file_refresh_pending = True   # cheia nouă poate fi în alt proiect Google
            if attempt > 0:
                st.toast(f"🔄 Trec pe cheia {ki + 1}/{len(keys)}...", icon="🔑")
        current_key = keys[ki]
        model_name = chain[mi]

        cached_content_name = None
        _started = False
        try:
            gemini_client = genai.Client(api_key=current_key)

            if _use_cache and model_name == MODEL_WITH_CACHE:
                cached_content_name = _get_or_create_cache(gemini_client, active_prompt, current_key)

            _cfg = dict(
                safety_settings=[
                    genai_types.SafetySetting(category=s["category"], threshold=s["threshold"])
                    for s in safety_settings
                ],
                thinking_config=_thinking_config_for(model_name),
            )
            if cached_content_name:
                gen_config = genai_types.GenerateContentConfig(cached_content=cached_content_name, **_cfg)
            else:
                gen_config = genai_types.GenerateContentConfig(system_instruction=active_prompt, **_cfg)

            history_new = []
            for msg in history_obj:
                history_new.append(
                    genai_types.Content(
                        role=msg["role"],
                        parts=[genai_types.Part(text=p) if isinstance(p, str) else genai_types.Part(file_data=genai_types.FileData(file_uri=p.uri, mime_type=p.mime_type)) for p in (msg["parts"] if isinstance(msg["parts"], list) else [msg["parts"]])]
                    )
                )

            if _file_refresh_pending:
                _payload_list = [
                    _ensure_gfile_for_client(gemini_client, p)
                    if (not isinstance(p, str) and hasattr(p, "uri")) else p
                    for p in _payload_list
                ]
                _file_refresh_pending = False

            current_parts = []
            for p in _payload_list:
                if isinstance(p, str):
                    current_parts.append(genai_types.Part(text=p))
                elif hasattr(p, "uri"):
                    current_parts.append(genai_types.Part(file_data=genai_types.FileData(file_uri=p.uri, mime_type=p.mime_type)))
                else:
                    current_parts.append(genai_types.Part(text=str(p)))

            all_contents = history_new + [genai_types.Content(role="user", parts=current_parts)]

            response_stream = gemini_client.models.generate_content_stream(
                model=model_name,
                contents=all_contents,
                config=gen_config,
            )

            chunks = []
            _prompt_tokens = 0
            _output_tokens = 0
            for chunk in response_stream:
                _piece = None
                try:
                    if chunk.text:
                        _piece = chunk.text
                    if hasattr(chunk, "usage_metadata") and chunk.usage_metadata:
                        um = chunk.usage_metadata
                        if getattr(um, "prompt_token_count", None):
                            _prompt_tokens = um.prompt_token_count
                        if getattr(um, "candidates_token_count", None):
                            _output_tokens = um.candidates_token_count
                except Exception:
                    continue
                if _piece:
                    if live_stream:
                        _started = True
                        yield _piece
                    else:
                        chunks.append(_piece)

            if live_stream and not _started:
                raise Exception("Modelul nu a returnat niciun text (posibil blocat de filtrele de siguranță).")
            if not live_stream and not chunks:
                raise Exception("Modelul nu a returnat niciun text (posibil blocat de filtrele de siguranță).")

            _key_id = f"_tokens_key_{ki}"
            _prev = st.session_state.get(_key_id, {"prompt": 0, "output": 0, "calls": 0})
            st.session_state[_key_id] = {
                "prompt": _prev["prompt"] + _prompt_tokens,
                "output": _prev["output"] + _output_tokens,
                "calls":  _prev["calls"]  + 1,
            }
            if mi > 0:
                st.toast(f"ℹ️ Răspuns generat cu modelul de rezervă ({model_name})", icon="🔄")
            if use_cache:
                st.session_state["_ctx_cache_enabled"] = bool(cached_content_name) if model_name == MODEL_WITH_CACHE \
                    else st.session_state.get("_ctx_cache_enabled", True)

            for text in chunks:
                yield text
            return

        except Exception as e:
            if _started:
                raise   # elevul a primit deja o parte din răspuns: o reîncercare ar dubla textul
            error_msg = str(e) + " " + repr(e)
            _err_low = error_msg.lower()

            _is_invalid_key = (
                "API key not valid" in error_msg or "API_KEY_INVALID" in error_msg
                or "api_key_invalid" in _err_low or "invalid api key" in _err_low
            )
            _is_quota = (
                re.search(r"\b429\b", error_msg) is not None
                or "quota" in _err_low or "rate_limit" in _err_low or "rate limit" in _err_low
                or "resource_exhausted" in _err_low
            )
            _is_overload = ("503" in error_msg or "overloaded" in _err_low or "unavailable" in _err_low)
            _is_file_err = _has_gfile and (
                "access the file" in _err_low or "files/" in _err_low
                or ("file" in _err_low and any(x in _err_low for x in ("permission_denied", "not_found", "not found")))
            )
            _is_missing_model = (not _is_file_err) and (
                re.search(r"\b404\b", error_msg) is not None or "not_found" in _err_low or "not found" in _err_low
                or "is not supported for generatecontent" in _err_low
            )

            # Eroare legată de caching: dezactivăm caching-ul și reîncercăm același model
            if (
                _use_cache and model_name == MODEL_WITH_CACHE
                and not (_is_invalid_key or _is_quota or _is_file_err)
                and (
                    (cached_content_name is None and "400" not in error_msg and not _is_overload and not _is_missing_model)
                    or "not supported" in _err_low
                    or "cach" in _err_low
                )
            ):
                _use_cache = False
                st.session_state["_ctx_cache_enabled"] = False
                continue

            if _is_file_err and not (_is_invalid_key or _is_quota) and _file_refreshes < 2 \
                    and st.session_state.get("_gsrc_map"):
                _file_refreshes += 1
                _file_refresh_pending = True
                continue

            if _is_invalid_key:
                _last_kind = "invalid"
                _invalidate_cache_for_key(current_key)
                _quota_block(current_key, None, time.time() + 3600)
                _log(f"[AI] cheia {ki + 1} invalidă — ocolită 1 oră", "silent")
                continue

            if _is_quota:
                _last_kind = "quota"
                _until = _quota_until_from_error(error_msg)
                _quota_block(current_key, model_name, _until)
                _log(f"[AI] cota epuizată: cheia {ki + 1} / {model_name} — ocolită "
                     f"{int((_until - time.time()) / 60)} min", "silent")
                st.toast(f"⚠️ Limită atinsă pe {model_name} — încerc altă variantă...", icon="🔄")
                continue

            if _is_missing_model:
                _last_kind = "missing"
                _quota_block(None, model_name, time.time() + 6 * 3600)
                _log(f"[AI] modelul {model_name} nu e disponibil — scos din lanț 6 ore", "silent", e)
                continue

            if "400" in error_msg:
                raise Exception(f"❌ Cerere invalidă (400): {error_msg}") from e

            if _is_overload:
                _last_kind = "overload"
                _quota_block(current_key, model_name, time.time() + 30)
                if time.time() >= _deadline:
                    raise Exception(
                        "Serviciul AI este supraîncărcat. Te rugăm să încerci din nou în câteva secunde. 🐢"
                    ) from e
                st.toast(f"🐢 {model_name} ocupat — încerc altă variantă...", icon="⏳")
                time.sleep(0.5)
                continue

            raise e

    raise Exception(
        "Ne pare rău, serviciul AI este momentan supraîncărcat. "
        "Te rugăm să încerci din nou în câteva secunde. 🙏"
    )


MAX_UPLOAD_FILES = 10


def _get_or_upload_gfile(uploaded_file):
    """Google File pentru o imagine/PDF: din cache (dacă e încă valid) sau urcat acum. None la eroare."""
    file_key = f"_gfile_{uploaded_file.name}_{uploaded_file.size}"
    cached_gf = st.session_state.get(file_key)
    if cached_gf:
        try:
            gemini_client = genai.Client(api_key=keys[st.session_state.key_index])
            refreshed = gemini_client.files.get(name=cached_gf.name)
            if _is_gfile_active(refreshed) or getattr(refreshed.state, "name", "") == "PROCESSING" \
                    or str(refreshed.state) in ("FileState.PROCESSING", "PROCESSING"):
                return refreshed
        except Exception:
            pass   # expirat pe Google (TTL 48h) sau altă cheie/proiect → îl re-urcăm
        st.session_state.pop(file_key, None)

    file_type = uploaded_file.type or "application/octet-stream"
    suffix = {"image/jpeg": ".jpg", "image/jpg": ".jpg", "image/png": ".png", "image/webp": ".webp",
              "image/gif": ".gif", "application/pdf": ".pdf"}.get(file_type, ".bin")
    spinner_text = (f"🖼️ Se trimite imaginea {uploaded_file.name}..." if file_type.startswith("image/")
                    else f"📚 Se trimite documentul {uploaded_file.name}...")
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        gemini_client = genai.Client(api_key=keys[st.session_state.key_index])
        with st.spinner(spinner_text):
            gfile = gemini_client.files.upload(file=tmp_path, config=genai_types.UploadFileConfig(mime_type=file_type))
            poll = 0
            while str(gfile.state) in ("FileState.PROCESSING", "PROCESSING") and poll < 60:
                time.sleep(1)
                gfile = gemini_client.files.get(name=gfile.name)
                poll += 1
        if _is_gfile_active(gfile):
            st.session_state[file_key] = gfile
            return gfile
        st.error(f"❌ {uploaded_file.name} nu a putut fi procesat "
                 f"(stare: {getattr(gfile.state, 'name', str(gfile.state))})")
    except Exception as e:
        _log("Eroare la încărcarea fișierului", "silent", e)
        st.error(f"❌ {uploaded_file.name} nu a putut fi încărcat. Verifică formatul și dimensiunea.")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
    return None


def _remove_attached_file(name: str, size: int, is_media: bool):
    """Callback „🗑️”: scoate un fișier din atașamente (și îl șterge de pe Google, dacă e imagine/PDF)."""
    fid = f"{name}_{size}"
    st.session_state.setdefault("_removed_file_keys", set()).add(fid)
    st.session_state.pop(f"_txtcache_{fid}", None)
    gf = st.session_state.pop(f"_gfile_{fid}", None) if is_media else None
    if gf:
        try:
            genai.Client(api_key=keys[st.session_state.key_index]).files.delete(name=gf.name)
        except Exception as _e:
            _log(f"Nu s-a putut șterge fișierul Google Files API: {_e}", "silent")


# === UI PRINCIPAL ===
st.title("🎓 Profesor Liceu")
comercial.ui.render_ai_notice()   # AI Act art. 50: elevul știe că vorbește cu un AI
_plata = st.query_params.get("plata", "")
if _plata:
    if _plata == "ok":
        st.success("✅ Plata a fost înregistrată. Abonamentul se activează în câteva secunde.")
    else:
        st.info("Plata a fost anulată. Poți încerca oricând din bara laterală (👤 Cont → ⭐ Abonament).")
    try:
        del st.query_params["plata"]
    except Exception:
        pass

# Afișăm materia selectată mic sub titlu
if st.session_state.get("pedagogie_mode"):
    st.caption("🧠 **Mod Sfaturi de studiu**")
else:
    _mat_curenta = st.session_state.get("materie_selectata")
    if _mat_curenta:
        _mat_label = next((k for k, v in MATERII.items() if v == _mat_curenta), _mat_curenta)
        st.caption(f"Materie selectată: **{_mat_label}**")

with st.sidebar:
    _render_account_ui()
    st.divider()

    st.header("⚙️ Opțiuni")

    # --- Selector materie ---
    st.subheader("📚 Materie")
    _materii_keys = list(MATERII.keys())
    _mat_saved = st.session_state.get("materie_selectata")
    _mat_default_idx = next(
        (i for i, k in enumerate(_materii_keys) if MATERII[k] == _mat_saved),
        0  # fallback la "🤖 Automat" dacă nu găsim
    )
    materie_label = st.selectbox(
        "Alege materia:",
        options=_materii_keys,
        index=_mat_default_idx,
        label_visibility="collapsed"
    )
    materie_selectata = MATERII[materie_label]
    _mod_automat = (materie_selectata is None)  # True când e "🤖 Automat"

    # Actualizează system prompt dacă s-a schimbat materia
    if st.session_state.get("materie_selectata") != materie_selectata:
        st.session_state.materie_selectata = materie_selectata
        if _mod_automat:
            # Mod automat — resetăm detecția, promptul va fi setat la primul mesaj
            st.session_state.pop("_detected_subject", None)
            st.session_state.pop("_pending_user_msg", None)
            st.session_state.system_prompt = get_system_prompt(
                materie=None,
                pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                mod_avansat=st.session_state.get("mod_avansat", False),
                mod_strategie=st.session_state.get("mod_strategie", False),
                mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
            )
        else:
            # Mod manual — selectorul are prioritate absolută
            st.session_state["_detected_subject"] = materie_selectata
            st.session_state.pop("_pending_user_msg", None)
            st.session_state.system_prompt = get_system_prompt(
                materie_selectata,
                pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                mod_avansat=st.session_state.get("mod_avansat", False),
                mod_strategie=st.session_state.get("mod_strategie", False),
                mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
            )
        # Forțăm rerun explicit — necesar pe mobil unde sidebar-ul nu declanșează
        # automat rerender-ul paginii principale după schimbare de materie
        st.rerun()

    # Info materie curentă sub selector
    if _mod_automat:
        _detected_now = st.session_state.get("_detected_subject")
        if _detected_now and _detected_now != "pedagogie":
            _det_label = _MATERII_LABEL.get(_detected_now, _detected_now.capitalize())
            st.caption(f"🔍 Detectat: **{_det_label}**")
        elif not _detected_now:
            st.caption("🔍 Materia se detectează automat din mesaj")
    else:
        st.info(f"Focusat pe: **{materie_label}**")

    # --- Toggle Sfaturi de studiu ---
    # Când se activează: salvează sesiunea curentă și deschide conversație nouă dedicată.
    # Când se dezactivează: restaurează sesiunea anterioară (sau meniul principal dacă nu exista).
    _ped_active = st.session_state.get("pedagogie_mode", False)
    _ped_toggle = st.toggle(
        "🧠 Sfaturi de studiu",
        value=_ped_active,
        help="Activează pentru sfaturi de organizare și tehnici de învățare eficientă. Dezactivează pentru a reveni la profesor."
    )

    if _ped_toggle != _ped_active:
        if _ped_toggle:
            # ── ACTIVARE: salvăm sesiunea curentă și deschidem una nouă ──
            st.session_state["_ped_prev_session_id"]   = st.session_state.get("session_id", "")
            st.session_state["_ped_prev_messages"]     = list(st.session_state.get("messages", []))
            st.session_state["_ped_prev_materie"]      = st.session_state.get("materie_selectata")
            st.session_state["_ped_prev_detected"]     = st.session_state.get("_detected_subject")
            st.session_state["_ped_prev_system_prompt"]= st.session_state.get("system_prompt", "")

            # Sesiune nouă dedicată sfaturilor de studiu
            _ped_sid = generate_unique_session_id()
            register_session(_ped_sid)
            st.session_state["session_id"] = _ped_sid
            st.session_state["messages"]   = []
            # FIX 1: adăugăm sesiunea de pedagogie în lista locală — apare în sidebar
            _my_sids = st.session_state.get("_my_session_ids", [])
            if _ped_sid not in _my_sids:
                _my_sids.append(_ped_sid)
            st.session_state["_my_session_ids"] = _my_sids
            # Curățăm modurile active (BAC, temă, quiz)
            for _k in ["bac_mode", "bac_active", "bac_materie", "bac_profil", "bac_subject",
                       "bac_barem", "bac_raspuns", "bac_corectat", "bac_corectare",
                       "bac_start_time", "bac_timp_min", "bac_from_photo", "bac_ocr_done",
                       "bac_timer_submitted", "bac_use_timer", "bac_show_barem",  # FIX Bug #2
                       "homework_mode", "hw_materie", "hw_text",
                       "hw_corectare", "hw_done", "hw_from_photo", "hw_ocr_done",
                       "quiz_mode", "quiz_active", "quiz_questions", "quiz_correct",
                       "quiz_answers", "quiz_submitted", "quiz_materie", "quiz_nivel",
                       "_suggested_question", "_pending_user_msg"]:
                st.session_state.pop(_k, None)
            st.session_state["pedagogie_mode"]    = True
            st.session_state["_detected_subject"] = "pedagogie"
            st.session_state["system_prompt"]     = get_system_prompt(
                materie="pedagogie",
                pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                mod_avansat=st.session_state.get("mod_avansat", False),
                mod_strategie=st.session_state.get("mod_strategie", False),
                mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
            )
            invalidate_session_cache()
            components.html(
                f"<script>localStorage.setItem('profesor_session_id', {json.dumps(_ped_sid)});</script>",
                height=0,
            )
        else:
            # ── DEZACTIVARE: restaurăm sesiunea anterioară ──
            _prev_sid = st.session_state.get("_ped_prev_session_id", "")
            _prev_msg = st.session_state.get("_ped_prev_messages", [])
            _prev_mat = st.session_state.get("_ped_prev_materie")
            _prev_det = st.session_state.get("_ped_prev_detected")
            _prev_sys = st.session_state.get("_ped_prev_system_prompt", "")

            st.session_state["pedagogie_mode"] = False
            # Curățăm cheile temporare de salvare
            for _k in ["_ped_prev_session_id", "_ped_prev_messages",
                       "_ped_prev_materie", "_ped_prev_detected", "_ped_prev_system_prompt"]:
                st.session_state.pop(_k, None)

            if _prev_sid and is_valid_session_id(_prev_sid):
                # Restaurăm sesiunea anterioară
                st.session_state["session_id"]        = _prev_sid
                st.session_state["messages"]          = _prev_msg
                st.session_state["materie_selectata"] = _prev_mat
                st.session_state["_detected_subject"] = _prev_det
                st.session_state["system_prompt"]     = _prev_sys or get_system_prompt(
                    materie=_prev_mat,
                    pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                    mod_avansat=st.session_state.get("mod_avansat", False),
                    mod_strategie=st.session_state.get("mod_strategie", False),
                    mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                )
                # FIX 3: actualizăm și URL-ul ?sid= — altfel la refresh se restaurează SID-ul de pedagogie
                try:
                    st.query_params["sid"] = _prev_sid
                except Exception:
                    pass
                components.html(
                    f"<script>localStorage.setItem('profesor_session_id', {json.dumps(_prev_sid)});</script>",
                    height=0,
                )
            else:
                # Nu exista sesiune anterioară → meniu principal (ecran curat)
                _new_main_sid = generate_unique_session_id()
                register_session(_new_main_sid)
                st.session_state["session_id"]        = _new_main_sid
                st.session_state["messages"]          = []
                st.session_state["materie_selectata"] = None
                st.session_state.pop("_detected_subject", None)
                st.session_state["system_prompt"]     = get_system_prompt(
                    materie=None,
                    pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                    mod_avansat=st.session_state.get("mod_avansat", False),
                    mod_strategie=st.session_state.get("mod_strategie", False),
                    mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                )
                # FIX 3b: actualizăm URL-ul și localStorage la sesiunea nouă
                try:
                    st.query_params["sid"] = _new_main_sid
                except Exception:
                    pass
                components.html(
                    f"<script>localStorage.setItem('profesor_session_id', {json.dumps(_new_main_sid)});</script>",
                    height=0,
                )
            invalidate_session_cache()
        st.rerun()

    st.divider()

    # --- Dark Mode toggle ---
    dark_mode = st.toggle("🌙 Mod Întunecat", value=st.session_state.get("dark_mode", False))
    if dark_mode != st.session_state.get("dark_mode", False):
        st.session_state.dark_mode = dark_mode
        st.rerun()

    # --- Mod Pas cu Pas ---
    pas_cu_pas = st.toggle(
        "🔢 Explicație Pas cu Pas",
        value=st.session_state.get("pas_cu_pas", False),
        help="Profesorul va explica fiecare problemă detaliat, pas cu pas, cu motivația fiecărei operații."
    )
    if pas_cu_pas != st.session_state.get("pas_cu_pas", False):
        st.session_state.pas_cu_pas = pas_cu_pas
        # Regenerează prompt-ul cu noul mod
        st.session_state.system_prompt = get_system_prompt(
            materie=st.session_state.get("materie_selectata"),
            pas_cu_pas=pas_cu_pas,
            mod_avansat=st.session_state.get("mod_avansat", False),
            mod_strategie=st.session_state.get("mod_strategie", False),
            mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
        )
        if pas_cu_pas:
            st.toast("🔢 Mod Pas cu Pas activat!", icon="✅")
        else:
            st.toast("Mod normal activat.", icon="💬")
        st.rerun()

    if st.session_state.get("pas_cu_pas"):
        st.info("🔢 **Pas cu Pas activ** — fiecare problemă e explicată detaliat.", icon="📋")

    # --- Mod Explică-mi Strategia ---
    mod_strategie = st.toggle(
        "🧠 Explică-mi Strategia",
        value=st.session_state.get("mod_strategie", False),
        help="Profesorul explică CUM să gândești rezolvarea — logica și strategia, nu calculele."
    )
    if mod_strategie != st.session_state.get("mod_strategie", False):
        st.session_state.mod_strategie = mod_strategie
        st.session_state.system_prompt = get_system_prompt(
            st.session_state.get("materie_selectata"),
            mod_avansat=st.session_state.get("mod_avansat", False),
            pas_cu_pas=st.session_state.get("pas_cu_pas", False),
            mod_strategie=mod_strategie,
            mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False)
        )
        st.toast("🧠 Mod Strategie activat!" if mod_strategie else "Mod normal activat.", icon="✅" if mod_strategie else "💬")
        st.rerun()
    if st.session_state.get("mod_strategie"):
        st.info("🧠 **Strategie activ** — înveți să gândești, nu să copiezi.", icon="🗺️")

    # --- Mod Avansat ---
    mod_avansat = st.toggle(
        "⚡ Mod Avansat",
        value=st.session_state.get("mod_avansat", False),
        help="Știi deja bazele? Profesorul sare peste explicații evidente și îți dă doar ideea cheie și calculul esențial."
    )
    if mod_avansat != st.session_state.get("mod_avansat", False):
        st.session_state.mod_avansat = mod_avansat
        st.session_state.system_prompt = get_system_prompt(
            st.session_state.get("materie_selectata"),
            mod_avansat=mod_avansat,
            pas_cu_pas=st.session_state.get("pas_cu_pas", False),
            mod_strategie=st.session_state.get("mod_strategie", False),
            mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
        )
        st.toast("⚡ Mod Avansat activat!" if mod_avansat else "Mod normal activat.", icon="✅" if mod_avansat else "💬")
        st.rerun()
    if st.session_state.get("mod_avansat"):
        st.info("⚡ **Mod Avansat activ** — răspunsuri scurte, doar esențialul.", icon="🎯")

    # --- Mod Pregătire BAC Intensivă ---
    mod_bac_intensiv = st.toggle(
        "🎓 Pregătire BAC Intensivă",
        value=st.session_state.get("mod_bac_intensiv", False),
        help="Focusat pe ce pică la BAC: tipare de subiecte, punctaj, timp, teorie lipsă detectată automat."
    )
    if mod_bac_intensiv != st.session_state.get("mod_bac_intensiv", False):
        st.session_state.mod_bac_intensiv = mod_bac_intensiv
        st.session_state.system_prompt = get_system_prompt(
            st.session_state.get("materie_selectata"),
            mod_avansat=st.session_state.get("mod_avansat", False),
            pas_cu_pas=st.session_state.get("pas_cu_pas", False),
            mod_strategie=st.session_state.get("mod_strategie", False),
            mod_bac_intensiv=mod_bac_intensiv
        )
        st.toast("🎓 Mod BAC Intensiv activat!" if mod_bac_intensiv else "Mod normal activat.", icon="✅" if mod_bac_intensiv else "💬")
        st.rerun()
    if st.session_state.get("mod_bac_intensiv"):
        st.info("🎓 **BAC Intensiv activ** — focusat pe ce pică la examen.", icon="📝")

    st.divider()

    # --- Status Supabase ---
    if not st.session_state.get("_sb_online", True):
        st.markdown(
            '<div style="background:#e67e22;color:white;padding:8px 12px;'
            'border-radius:8px;font-size:13px;text-align:center;margin-bottom:8px">'
            '📴 Mod offline — datele sunt salvate local</div>',
            unsafe_allow_html=True
        )
    else:
        pending = len(st.session_state.get("_offline_queue", []))
        if pending:
            st.caption(f"☁️ {pending} mesaje în așteptare pentru sincronizare")


    st.divider()

    # === DESCĂRCARE CONVERSAȚIE ===
    _msgs_for_download = st.session_state.get("messages", [])
    if _msgs_for_download:
        import datetime as _dt

        def _build_conversation_txt(messages: list) -> str:
            """Construiește textul conversației pentru descărcare."""
            _materie = st.session_state.get("materie_selectata") or st.session_state.get("_detected_subject")
            _materie_label = _MATERII_LABEL.get(_materie, "General") if _materie else "General"
            _sid_short = st.session_state.session_id[:8]
            _now = _dt.datetime.now().strftime("%d.%m.%Y %H:%M")

            lines = [
                "=" * 60,
                "  PROFESOR VIRTUAL AI — Conversație exportată",
                "=" * 60,
                f"  Materie : {_materie_label}",
                f"  Data    : {_now}",
                f"  Sesiune : {_sid_short}...",
                f"  Mesaje  : {len(messages)}",
                "=" * 60,
                "",
            ]
            for msg in messages:
                role = msg.get("role", "")
                content = msg.get("content", "")
                if role == "user":
                    lines.append("👤 ELEV:")
                elif role == "assistant":
                    lines.append("🎓 PROFESOR:")
                else:
                    lines.append(f"[{role.upper()}]:")
                # Curățăm marcajele SVG din export — nu au sens în text plain
                _clean_content = re.sub(r'\[\[DESEN_SVG\]\].*?\[\[/DESEN_SVG\]\]', '[desen SVG]', content, flags=re.DOTALL)
                _clean_content = re.sub(r'<svg\b.*?</svg\s*>', '[desen SVG]', _clean_content, flags=re.DOTALL | re.IGNORECASE)
                # Înlocuim marker-ul SRT cu un sumar — traducerea e descărcabilă separat ca .srt
                _clean_content = re.sub(
                    r'\[SRT_TRANSLATION_KEY:[^\]]+\]',
                    '[Traducere SRT completă — descarcă fișierul .srt separat din chat]',
                    _clean_content
                )
                # Fallback pentru formatul vechi cu ```srt
                _clean_content = re.sub(
                    r'```srt\n.*?```',
                    '[Traducere SRT completă — descarcă fișierul .srt separat din chat]',
                    _clean_content,
                    flags=re.DOTALL
                )
                lines.append(_clean_content.strip())
                lines.append("-" * 60)
                lines.append("")
            lines.append("=" * 60)
            lines.append("  Export generat de Profesor Virtual AI")
            lines.append("=" * 60)
            return "\n".join(lines)

        _conv_text = _build_conversation_txt(_msgs_for_download)
        _materie_fn = (st.session_state.get("materie_selectata") or "conversatie") or "conversatie"
        _materie_fn = re.sub(r'[^a-zA-Z0-9_-]', '_', str(_materie_fn))
        _date_fn = _dt.datetime.now().strftime("%Y%m%d")
        _filename = f"profesor_ai_{_materie_fn}_{_date_fn}.txt"

        st.download_button(
            label="💾 Descarcă conversația",
            data=_conv_text.encode("utf-8"),
            file_name=_filename,
            mime="text/plain",
            use_container_width=True,
            help="Salvează întreaga conversație ca fișier text (.txt)",
        )

        # Dacă există o traducere SRT în sesiune, oferim și butonul de descărcare SRT în sidebar
        _srt_sidebar = None
        for _sk, _sv in st.session_state.items():
            if _sk.startswith("_srt_translation_") and isinstance(_sv, dict):
                _srt_sidebar = _sv
                break
        if _srt_sidebar:
            st.download_button(
                label=f"⬇️ Descarcă subtitrarea tradusă",
                data=_srt_sidebar["text"].encode("utf-8"),
                file_name=_srt_sidebar["filename"],
                mime="text/plain",
                use_container_width=True,
                help=f"{_srt_sidebar['blocks']} replici — {_srt_sidebar['orig_name']}",
                key="_dl_srt_sidebar",
            )

    if st.button("🗑️ Șterge Istoricul", type="primary"):
        clear_history_db(st.session_state.session_id)
        st.session_state.messages = []
        st.rerun()

    st.divider()

    st.header("📁 Materiale")

    # Mai multe fișiere deodată (accept_multiple_files). Cheie fixă: widgetul își păstrează
    # valoarea peste rerun-urile programatice (schimbare materie, toggle mod etc.).
    uploaded_files = st.file_uploader(
        f"Încarcă fișiere (maxim {MAX_UPLOAD_FILES}: imagini, PDF, Word, text, DBF, subtitrare)",
        type=["jpg", "jpeg", "png", "webp", "gif", "pdf",
              "txt", "srt", "docx", "doc", "dbf"],
        accept_multiple_files=True,
        help=(
            "Poți selecta mai multe fișiere deodată (Ctrl/Cmd + click) sau le poți adăuga pe rând. "
            "Imagini: analizate vizual de AI. PDF: citit integral. "
            "Word (.docx), text (.txt), subtitrare (.srt), baze de date (.dbf): conținutul e extras și trimis la AI."
        ),
        key="_main_file_uploader",
    )
    media_content = None       # LISTA de Google Files trimise la AI (imagini/PDF) sau None
    text_file_content = None   # textul extras din fișierele text (txt/docx/doc/dbf/srt) sau None

    _all_files = list(uploaded_files or [])
    # Fișierele scoase cu „🗑️” rămân în widget (nu se poate reseta programatic) → le ignorăm.
    _removed_keys = st.session_state.setdefault("_removed_file_keys", set())
    _present_keys = {f"{f.name}_{f.size}" for f in _all_files}
    _removed_keys.intersection_update(_present_keys)   # scos și din widget → poate fi adăugat din nou
    _files = [f for f in _all_files if f"{f.name}_{f.size}" not in _removed_keys]
    if len(_files) > MAX_UPLOAD_FILES:
        st.warning(f"⚠️ Am păstrat primele {MAX_UPLOAD_FILES} fișiere; restul sunt ignorate.")
        _files = _files[:MAX_UPLOAD_FILES]

    if not _files:
        # Nimic atașat (scos cu X sau cu 🗑️) → nu mai trimitem nimic la AI.
        for _k in ("_current_uploaded_file_meta", "_active_gfile_keys", "_active_textcache_keys"):
            st.session_state.pop(_k, None)
    else:
        _media_list, _text_parts, _metas = [], [], []
        _gkeys, _tkeys, _gsrc = [], [], {}
        for _uf in _files:
            _fid = f"{_uf.name}_{_uf.size}"
            _meta = {"name": _uf.name, "type": _uf.type or "", "size": _uf.size}
            if _is_text_file(_uf):
                _tkey = f"_txtcache_{_fid}"
                _txt = st.session_state.get(_tkey)
                if _txt is None:
                    with st.spinner(f"📄 Se extrage conținutul din {_uf.name}..."):
                        _txt = _extract_text_from_uploaded_file(_uf)
                    if _txt:
                        st.session_state[_tkey] = _txt
                    else:
                        st.error(f"❌ Nu s-a putut extrage textul din {_uf.name}.")
                if _txt:
                    _text_parts.append((_uf.name, _txt))
                    _metas.append(_meta)
                    _tkeys.append(_tkey)
            else:
                _gf = _get_or_upload_gfile(_uf)
                if _gf:
                    _media_list.append(_gf)
                    _metas.append(_meta)
                    _gkeys.append(f"_gfile_{_fid}")
                    _bytes = _uf.getvalue()
                    if len(_bytes) <= 25 * 1024 * 1024:   # pentru re-upload la rotația cheilor
                        _gsrc[_gf.name] = (_bytes, _uf.type)

        if _media_list:
            media_content = _media_list
            st.session_state["_gsrc_map"] = _gsrc
        if _text_parts:
            text_file_content = (
                _text_parts[0][1] if len(_text_parts) == 1
                else "\n\n".join(f"=== FIȘIER: {n} ===\n{t}" for n, t in _text_parts)
            )
        if _metas:
            st.session_state["_current_uploaded_file_meta"] = (
                _metas[0] if len(_metas) == 1 else
                {"name": f"{len(_metas)} fișiere", "type": "multi", "size": sum(m["size"] for m in _metas),
                 "files": _metas}
            )
        st.session_state["_active_gfile_keys"] = _gkeys
        st.session_state["_active_textcache_keys"] = _tkeys

        # ── Lista fișierelor atașate, fiecare cu buton de eliminare ──
        st.caption(f"📎 {len(_metas)} fișier(e) atașat(e) — se trimit la AI împreună cu întrebarea ta.")
        for _uf in _files:
            _fid = f"{_uf.name}_{_uf.size}"
            _is_txt = _is_text_file(_uf)
            _ok = (_fid in [k.replace("_txtcache_", "", 1) for k in _tkeys]) if _is_txt \
                else (f"_gfile_{_fid}" in _gkeys)
            if not _ok:
                continue
            _ico = ("🖼️" if (_uf.type or "").startswith("image/") else "📕" if "pdf" in (_uf.type or "")
                    else "🎬" if _uf.name.lower().endswith(".srt") else "🗄️" if _uf.name.lower().endswith(".dbf")
                    else "📝" if _uf.name.lower().endswith((".docx", ".doc")) else "📄")
            _c1, _c2 = st.columns([5, 1])
            with _c1:
                _extra = (f"{len(st.session_state.get('_txtcache_' + _fid, '')):,} caractere" if _is_txt
                          else f"{max(1, _uf.size // 1024)} KB")
                st.markdown(f"{_ico} **{_uf.name}**  \n<small>{_extra}</small>", unsafe_allow_html=True)
            with _c2:
                st.button("🗑️", key=f"_rm_file_{_fid}", help=f"Elimină {_uf.name}",
                          on_click=_remove_attached_file, args=(_uf.name, _uf.size, not _is_txt))
            if (_uf.type or "").startswith("image/"):
                try:
                    st.image(_uf, use_container_width=True)
                except Exception:
                    st.caption("(previzualizare indisponibilă pentru această imagine)")
            elif _is_txt and not st.session_state.get("_txtcache_" + _fid, "").startswith("⚠️"):
                with st.expander("👁️ Previzualizare", expanded=False):
                    _t = st.session_state.get("_txtcache_" + _fid, "")
                    st.text(_t[:300] + ("\n..." if len(_t) > 300 else ""))

    st.divider()

    # --- Mod Quiz + BAC ---
    st.subheader("📝 Examinare & BAC")

    # Chei exacte per mod — actualizați când adăugați chei noi în fiecare mod
    _BAC_KEYS = [
        "bac_mode", "bac_active", "bac_materie", "bac_profil", "bac_subject",
        "bac_barem", "bac_raspuns", "bac_corectat", "bac_corectare",
        "bac_start_time", "bac_timp_min", "bac_from_photo", "bac_ocr_done",
        "bac_timer_submitted", "bac_use_timer", "bac_show_barem",  # FIX Bug #2: adăugate
    ]
    _HW_KEYS = [
        "homework_mode", "hw_materie", "hw_text", "hw_text_input", "hw_corectare",
        "hw_done", "hw_from_photo", "hw_ocr_done",
    ]
    _QUIZ_KEYS = [
        "quiz_mode", "quiz_active", "quiz_questions", "quiz_correct",
        "quiz_answers", "quiz_submitted", "quiz_materie", "quiz_nivel",
    ]
    _SHARED_KEYS = ["_suggested_question", "_pending_user_msg"]
    _ADMITERE_KEYS = [
        "admitere_mode", "admitere_univ", "admitere_spec", "admitere_proba",
        "admitere_active", "admitere_subject", "admitere_barem",
        "admitere_raspuns", "admitere_corectat", "admitere_corectare",
        "admitere_start_time", "admitere_mod_grila", "admitere_grila_answers",
        "admitere_grila_submitted", "admitere_grila_questions",
    ]

    def _clear_all_modes():
        for k in _BAC_KEYS + _HW_KEYS + _QUIZ_KEYS + _SHARED_KEYS + _ADMITERE_KEYS:
            st.session_state.pop(k, None)

    col_q, col_b = st.columns(2)
    with col_q:
        if st.button("🎯 Quiz rapid", use_container_width=True,
                     type="primary" if st.session_state.get("quiz_mode") else "secondary"):
            entering = not st.session_state.get("quiz_mode", False)
            _clear_all_modes()
            st.session_state.quiz_mode = entering
            st.session_state.pop("bac_mode", None)
            st.session_state.pop("homework_mode", None)
            st.rerun()
    with col_b:
        if st.button("🎓 Simulare BAC", use_container_width=True,
                     type="primary" if st.session_state.get("bac_mode") else "secondary"):
            entering = not st.session_state.get("bac_mode", False)
            _clear_all_modes()
            st.session_state.bac_mode = entering
            st.session_state.pop("quiz_mode", None)
            st.session_state.pop("homework_mode", None)
            st.rerun()

    if st.button("🏛️ Admitere Facultate", use_container_width=True,
                 type="primary" if st.session_state.get("admitere_mode") else "secondary"):
        entering = not st.session_state.get("admitere_mode", False)
        _clear_all_modes()
        st.session_state.admitere_mode = entering
        st.rerun()

    if st.button("📚 Corectează Temă", use_container_width=True,
                 type="primary" if st.session_state.get("homework_mode") else "secondary"):
        entering = not st.session_state.get("homework_mode", False)
        _clear_all_modes()
        st.session_state.homework_mode = entering
        st.session_state.pop("quiz_mode", None)
        st.session_state.pop("bac_mode", None)
        st.rerun()

    st.divider()

    # --- Istoric conversații ---
    st.subheader("🕐 Conversații anterioare")
    if st.button("🔄 Conversație nouă", use_container_width=True):
        _cleanup_gfiles()
        new_sid = generate_unique_session_id()
        register_session(new_sid)
        # Salvează noul SID în lista sesiunilor acestui browser
        _my_sids = st.session_state.get("_my_session_ids", [])
        if new_sid not in _my_sids:
            _my_sids.append(new_sid)
        st.session_state["_my_session_ids"] = _my_sids
        switch_session(new_sid)
        # FIX PERSISTENȚĂ (v2): nu mai e nevoie de ?new=1 — switch_session() setează
        # deja st.session_state["session_id"], iar get_or_create_session_id() îl
        # găsește acolo la următorul run și îl scrie direct în ?sid=, fără să mai
        # treacă vreodată prin gate-ul de verificare localStorage (acela rulează
        # DOAR când URL-ul e complet curat și session_state nu are niciun SID).
        st.rerun()

    # Afișează DOAR sesiunile create de acest browser în această sesiune Streamlit
    # (nu toate sesiunile din Supabase — acelea aparțin altor utilizatori)
    current_sid = st.session_state.session_id
    _my_sids = st.session_state.get("_my_session_ids", [current_sid])
    if current_sid not in _my_sids:
        _my_sids = [current_sid] + _my_sids
        st.session_state["_my_session_ids"] = _my_sids

    # Încarcă preview-urile doar pentru sesiunile acestui browser
    sessions = []
    try:
        _supabase = get_supabase_client()
        _q = (
            _supabase.table("session_previews")
            .select("session_id, last_active, msg_count, preview")
            .eq("app_id", get_app_id())
        )
        if current_user_id():
            # Cu cont: toate conversațiile contului, de pe orice device
            _q = _q.eq("liceu_user_id", current_user_id())
        else:
            # Fără cont: doar conversațiile anonime ale acestui browser
            _q = _q.in_("session_id", _my_sids).is_("liceu_user_id", "null")
        _resp = _q.gt("msg_count", 0).order("last_active", desc=True).limit(30 if current_user_id() else 15).execute()
        sessions = _resp.data or []
    except Exception:
        pass

    for s in sessions:
        is_current = s["session_id"] == current_sid
        # FIX 5: etichetă vizuală pentru sesiunile de sfaturi de studiu
        _preview_text = s['preview'] or "Conversație"
        _is_ped_session = _preview_text.lower().startswith(("sfat", "studi", "tehnic", "înv", "inv", "📚", "🧠"))
        _ped_prefix = "🧠 " if _is_ped_session else ""
        label = f"{'▶ ' if is_current else ''}{_ped_prefix}{_preview_text}"
        caption = f"{format_time_ago(s['last_active'])} · {s['msg_count']} mesaje"
        with st.container():
            col_btn, col_del = st.columns([5, 1])
            with col_btn:
                if st.button(
                    label,
                    key=f"sess_{s['session_id']}",
                    use_container_width=True,
                    type="primary" if is_current else "secondary",
                    help=caption,
                ):
                    if not is_current:
                        switch_session(s["session_id"])
                        st.rerun()
            with col_del:
                if st.button("🗑", key=f"del_{s['session_id']}", help="Șterge"):
                    clear_history_db(s["session_id"])
                    if is_current:
                        st.session_state.messages = []
                    # Scoate din lista locală
                    _my_sids2 = st.session_state.get("_my_session_ids", [])
                    if s["session_id"] in _my_sids2:
                        _my_sids2.remove(s["session_id"])
                    st.session_state["_my_session_ids"] = _my_sids2
                    st.rerun()

    st.divider()

    _debug_val = st.session_state.get("_debug_info_open", False)
    _debug_checked = st.checkbox("🔧 Debug Info", value=_debug_val, key="chk_debug_info")
    if _debug_checked != _debug_val:
        st.session_state["_debug_info_open"] = _debug_checked

    if _debug_checked:
        msg_count = len(st.session_state.get("messages", []))
        st.caption(f"📊 Mesaje în memorie: {msg_count}/{MAX_MESSAGES_IN_MEMORY}")
        st.caption(f"🔑 Cheie API activă: {st.session_state.key_index + 1}/{len(keys)} · {quota_status_text()}")

        # ── Statistici token usage per cheie (sesiunea curentă) ──
        # Notă: Gemini Free tier = 1.500 req/zi și 1.000.000 token/min per cheie.
        # Nu avem acces la quota rămasă prin API — afișăm consumul din sesiunea curentă.
        _active_idx = st.session_state.get("key_index", 0)
        _key_id = f"_tokens_key_{_active_idx}"
        _usage = st.session_state.get(_key_id, {"prompt": 0, "output": 0, "calls": 0})
        _total_tok = _usage["prompt"] + _usage["output"]
        _calls = _usage["calls"]
        if _calls > 0:
            st.caption(f"📈 Tokeni folosiți (cheia {_active_idx + 1}, sesiunea curentă):")
            st.caption(f"   ↳ Input: {_usage['prompt']:,} · Output: {_usage['output']:,} · Total: {_total_tok:,}")
            st.caption(f"   ↳ Apeluri AI: {_calls} · Medie/apel: {_total_tok // max(_calls,1):,} tok")
            # Bară vizuală față de limita de 1M tokeni/minut (limita de rate, nu de quota zilnică)
            _pct = min(_total_tok / 1_000_000 * 100, 100)
            _bar_filled = int(_pct / 5)
            _bar = "█" * _bar_filled + "░" * (20 - _bar_filled)
            _color = "🟢" if _pct < 50 else ("🟡" if _pct < 80 else "🔴")
            st.caption(f"   {_color} [{_bar}] {_pct:.1f}% din 1M tok/min")
        else:
            st.caption("📈 Tokeni folosiți: 0 (niciun apel AI în sesiunea curentă)")

        # Sumar pentru toate cheile din sesiune
        _all_keys_usage = []
        for i in range(len(keys)):
            _u = st.session_state.get(f"_tokens_key_{i}", {"prompt": 0, "output": 0, "calls": 0})
            if _u["calls"] > 0:
                _all_keys_usage.append(f"Cheia {i+1}: {_u['prompt']+_u['output']:,} tok ({_u['calls']} apeluri)")
        if len(_all_keys_usage) > 1:
            st.caption("📋 Toate cheile folosite: " + " | ".join(_all_keys_usage))

        st.caption(f"🆔 Sesiune: {st.session_state.session_id[:16]}...")

    # Datele firmei, ANPC/SAL și documentele legale — mereu la finalul barei laterale
    comercial.ui.render_footer()


# Versiunea comercială: fără cont / acord parental / documente acceptate / abonament → ecranul de acces
_ent = comercial.billing.entitlement(current_user())
if not _ent["ok"]:
    comercial.ui.render_gate(_ent, current_user())   # se oprește singură (st.stop)

# === MAIN UI — TEME / BAC / QUIZ / CHAT ===
def _run_mode(fn):
    """Rulează un mod (temă/admitere/BAC/quiz); limita zilnică atinsă apare ca mesaj, nu ca eroare."""
    try:
        fn()
    except comercial.core.AccessError as _ae:
        st.warning(str(_ae))
    st.stop()


if st.session_state.get("homework_mode"):
    _run_mode(run_homework_ui)

if st.session_state.get("admitere_mode"):
    _run_mode(run_admitere_ui)

if st.session_state.get("bac_mode"):
    _run_mode(run_bac_sim_ui)

if st.session_state.get("quiz_mode"):
    _run_mode(run_quiz_ui)

# === ÎNCĂRCARE MESAJE (CHAT MODE) ===
# Încărcăm istoricul dacă: nu există messages, sau messages aparțin altei sesiuni
_current_sid = st.session_state.session_id
if (
    "messages" not in st.session_state
    or st.session_state.get("_messages_for_sid") != _current_sid
):
    _loaded_msgs = load_history_from_db(_current_sid)
    st.session_state.messages = _loaded_msgs
    st.session_state["_messages_for_sid"] = _current_sid
    st.session_state.pop("_history_may_be_incomplete", None)

    # ── Restaurare traduceri SRT din istoricul încărcat ──
    # La refresh, session_state se șterge. Restaurăm traducerile SRT din mesajele
    # speciale role="srt_data" salvate în Supabase la momentul traducerii.
    try:
        _sb_restore = get_supabase_client()
        if _sb_restore:
            _srt_rows = (
                _sb_restore.table("history")
                .select("content")
                .eq("session_id", _current_sid)
                .eq("app_id", get_app_id())
                .eq("role", "srt_data")
                .execute()
            )
            for _row in (_srt_rows.data or []):
                _rc = _row.get("content", "")
                _hdr = re.match(r'\[SRT_DATA:(_srt_translation_[^\]]+)\]\n', _rc)
                if not _hdr:
                    continue
                _rkey  = _hdr.group(1)
                _rtext = _rc[_hdr.end():]
                if not _rtext.strip():
                    continue
                if st.session_state.get(_rkey):
                    continue  # deja restaurat
                _orig_r  = _rkey.replace("_srt_translation_", "", 1)
                _trad_r  = re.sub(r'\.srt$', '_RO.srt', _orig_r, flags=re.IGNORECASE)
                _blks_r  = _rtext.count("\n\n") + 1
                st.session_state[_rkey] = {
                    "text":      _rtext,
                    "filename":  _trad_r,
                    "orig_name": _orig_r,
                    "blocks":    _blks_r,
                }
    except Exception:
        pass  # restaurarea SRT e best-effort — nu blocăm aplicația

    # ── Restaurare materie din istoricul încărcat ──
    # Dacă sesiunea are mesaje dar materia nu e setată (ex: după switch_session),
    # detectăm materia din primele mesaje ale elevului și o blocăm.
    # Asta previne re-detectarea la mijlocul conversației după un reload.
    if _loaded_msgs and not st.session_state.get("_detected_subject"):
        # Căutăm primele 3 mesaje ale elevului pentru o detecție mai sigură
        _first_user_msgs = [m["content"] for m in _loaded_msgs if m.get("role") == "user"][:3]
        _combined_text = " ".join(_first_user_msgs)
        if _combined_text:
            _restored_subject = detect_subject_from_text(_combined_text)
            if _restored_subject:
                st.session_state["_detected_subject"] = _restored_subject
                update_system_prompt_for_subject(_restored_subject)

    # ── Revenire din altă sesiune/zi: pre-generăm rezumatul de context ──
    # FIX 7: nu generăm rezumat dacă lista e goală (sesiune nouă de pedagogie sau chat nou)
    _loaded_count = len(_loaded_msgs)
    if _loaded_count > MAX_MESSAGES_TO_SEND_TO_AI:
        _sum_key     = "_conversation_summary"
        _sum_sid_key = "_summary_for_sid"
        _needs_summary = (
            not st.session_state.get(_sum_key)
            or st.session_state.get(_sum_sid_key) != st.session_state.session_id
        )
        if _needs_summary:
            with st.spinner("📚 Profesorul reîncarcă contextul conversației anterioare..."):
                _auto_summary = summarize_conversation(_loaded_msgs)
            if _auto_summary:
                st.session_state[_sum_key]     = _auto_summary
                st.session_state["_summary_cached_at"] = _loaded_count
                st.session_state[_sum_sid_key] = st.session_state.session_id
                st.toast("✅ Contextul conversației anterioare a fost reîncărcat!", icon="🧠")

# Banner mod Pas cu Pas
if st.session_state.get("pas_cu_pas"):
    st.markdown(
        '<div style="background:linear-gradient(135deg,#667eea,#764ba2);color:white;'
        'padding:10px 16px;border-radius:10px;margin-bottom:12px;'
        'display:flex;align-items:center;gap:10px;font-size:14px;">'
        '🔢 <strong>Mod Pas cu Pas activ</strong> — '
        'Profesorul îți va explica fiecare problemă detaliat, cu motivația fiecărui pas.'
        '</div>',
        unsafe_allow_html=True
    )

for i, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        if msg["role"] == "assistant":
            content = msg["content"]
            # Detectăm mesajele de traducere SRT după marker-ul compact
            _srt_key_match = re.search(r'\[SRT_TRANSLATION_KEY:([^\]]+)\]', content)
            if _srt_key_match:
                _srt_key = _srt_key_match.group(1)
                _srt_data = st.session_state.get(_srt_key)
                # Afișăm prima linie (sumarul) fără marker
                first_line = content.split("\n")[0]
                st.markdown(first_line)
                if _srt_data:
                    st.download_button(
                        label="⬇️ Descarcă subtitrarea tradusă (.srt)",
                        data=_srt_data["text"].encode("utf-8"),
                        file_name=_srt_data["filename"],
                        mime="text/plain",
                        use_container_width=True,
                        key=f"_dl_srt_hist_{i}",
                    )
                    # Afișăm TOT textul tradus — fără trunchiere
                    st.text(_srt_data["text"])
                else:
                    st.caption("⚠️ Traducerea nu mai este disponibilă în această sesiune (sesiunea a fost reîncărcată). Retrimite fișierul pentru a traduce din nou.")
            else:
                render_message_with_svg(content)
            # Raportarea unui răspuns AI greșit/nepotrivit (verificat de admin în „👑 Administrare”)
            comercial.ui.render_report_button(i, content, st.session_state.get("session_id"), current_user_id())
        else:
            st.markdown(msg["content"])

    # Butoanele apar DOAR sub ultimul mesaj al profesorului
    if (msg["role"] == "assistant" and
            i == len(st.session_state.messages) - 1):
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🔄 Nu am înțeles", key="qa_reexplain", use_container_width=True, help="Explică altfel, cu o altă analogie"):
                st.session_state["_quick_action"] = "reexplain"
                st.rerun()
        with col2:
            if st.button("✏️ Exercițiu similar", key="qa_similar", use_container_width=True, help="Generează un exercițiu similar pentru practică"):
                st.session_state["_quick_action"] = "similar"
                st.rerun()
        with col3:
            if st.button("🧠 Explică strategia", key="qa_strategy", use_container_width=True, help="Cum să gândești acest tip de problemă"):
                st.session_state["_quick_action"] = "strategy"
                st.rerun()


# ── Handler pentru butoanele de acțiuni rapide ──

TYPING_HTML = """
<div class="typing-indicator">
    <div class="typing-dots"><span></span><span></span><span></span></div>
    <span>Domnul Profesor scrie...</span>
</div>
"""

if st.session_state.get("_quick_action"):
    action = st.session_state.pop("_quick_action")
    # FIX Bug 2: _quick_action_ref nu era setat nicăieri — eliminat, nu mai e necesar
    # (context-ul vine direct din ultimul mesaj al asistentului/utilizatorului)

    # ── Găsește ultimul mesaj al asistentului pentru context real ──
    last_assistant_msg = ""
    last_user_msg = ""
    for msg in reversed(st.session_state.messages):
        if msg["role"] == "assistant" and not last_assistant_msg:
            last_assistant_msg = msg["content"]
        if msg["role"] == "user" and not last_user_msg:
            last_user_msg = msg["content"]
        if last_assistant_msg and last_user_msg:
            break

    # FIX Bug 2: logică robustă pentru prev_topic și prev_question
    # - curățăm LaTeX (\$...\$, \$\$...\$\$) și markdown înainte de trunchiere
    # - trunchierea se face la spațiu, nu în mijlocul unui cuvânt LaTeX
    # - fallback explicit dacă mesajele lipsesc
    _clean = lambda t: re.sub(r'\$\$[\s\S]*?\$\$|\$[^\$\n]*?\$|[*`#\\]', '', t).strip()
    _clean2 = lambda t: re.sub(r'\s+', ' ', _clean(t))  # colapsăm whitespace multiplu

    if last_assistant_msg:
        _cleaned = _clean2(last_assistant_msg)
        # Trunchierea la 120 de caractere, la granița unui cuvânt
        if len(_cleaned) > 120:
            prev_topic = _cleaned[:120].rsplit(' ', 1)[0].rstrip('.,;:') + "..."
        else:
            prev_topic = _cleaned or "subiectul anterior"
    else:
        prev_topic = "subiectul anterior"

    if last_user_msg:
        _cleaned_q = _clean2(last_user_msg)
        prev_question = _cleaned_q[:100] if len(_cleaned_q) > 100 else _cleaned_q
        prev_question = prev_question or "întrebarea anterioară"
    else:
        prev_question = "întrebarea anterioară"

    action_prompts = {
        "reexplain": (
            f"Nu am înțeles explicația ta despre: '{prev_topic}'. "
            f"Te rog să explici din nou, dar complet diferit — "
            f"altă analogie, altă ordine a pașilor, exemple mai simple din viața reală. "
            f"Evită exact aceleași cuvinte și structura anterioară."
        ),
        "similar": (
            (
                f"Generează un exercițiu similar cu '{prev_question}', "
                f"folosind alt cuvânt sau altă situație de comunicare, cu dificultate puțin mai mare. "
                f"Enunță exercițiul ÎNTÂI, apoi rezolvă-l complet pas cu pas."
            ) if st.session_state.get("materie_selectata") in ("limba engleză", "limba franceză", "limba germană")
            else (
                f"Generează un exercițiu similar cu '{prev_question}', "
                f"cu date numerice diferite și dificultate puțin mai mare. "
                f"Enunță exercițiul ÎNTÂI, apoi rezolvă-l complet pas cu pas."
            )
        ),
        "strategy": (
            f"Explică-mi STRATEGIA de gândire pentru '{prev_question}': "
            f"cum recunosc că e acest tip, ce fac primul pas în minte, ce capcane să evit. "
            f"Fără calcule — vreau doar logica și gândirea din spate."
        ),
    }
    injected = action_prompts.get(action, "")
    if injected:
        with st.chat_message("user"):
            st.markdown(injected)
        st.session_state.messages.append({"role": "user", "content": injected})
        save_message_with_limits(st.session_state.session_id, "user", injected)

        context_messages = get_context_for_ai(st.session_state.messages)
        history_obj = []
        for msg in context_messages:
            role_gemini = "model" if msg["role"] == "assistant" else "user"
            history_obj.append({"role": role_gemini, "parts": [msg["content"]]})

        # Salvăm pentru retry în caz de eroare de cheie
        st.session_state["_retry_history"] = history_obj
        st.session_state["_retry_payload"] = [injected]

        with st.chat_message("assistant"):
            message_placeholder = st.empty()
            full_response = ""
            message_placeholder.markdown(TYPING_HTML, unsafe_allow_html=True)
            try:
                for text_chunk in run_chat_with_rotation(history_obj, [injected]):
                    full_response += text_chunk
                    message_placeholder.markdown(full_response + "▌")
                message_placeholder.empty()
                render_message_with_svg(full_response)
                st.session_state.messages.append({"role": "assistant", "content": full_response})
                save_message_with_limits(st.session_state.session_id, "assistant", full_response)
                st.session_state.pop("_retry_history", None)
                st.session_state.pop("_retry_payload", None)
            except Exception as e:
                message_placeholder.empty()
                _is_key_err = any(x in str(e) for x in ["epuizat", "invalide", "quota", "429", "API key"])
                if _is_key_err:
                    _show_ai_error(e, "acțiune rapidă")
                    # on_click: butonul nu mai e randat la rerun-ul de după click (blocul e one-shot),
                    # deci un `if st.button(...)` nu s-ar declanșa niciodată; callback-ul rulează oricum.
                    st.button("🔄 Reîncercați răspunsul", key="_retry_quick_action", type="primary",
                              on_click=_request_retry)
                else:
                    _show_ai_error(e, "acțiune rapidă")
    st.stop()

# ── Handler mesaj în așteptare — materie nedetectată în mod Automat ──
if st.session_state.get("_pending_user_msg") and st.session_state.get("materie_selectata") is None:
    _pending_msg = st.session_state["_pending_user_msg"]

    # Caz special: fizică detectată dar profil ambiguu → prompt dedicat pentru profil
    if st.session_state.get("_pending_fizica_ambigua"):
        with st.chat_message("assistant"):
            st.markdown(
                "**Am detectat că întrebarea e despre Fizică!** 🔬\n\n"
                "Pentru a-ți răspunde corect conform programei, spune-mi la ce profil ești:"
            )
            col_r, col_t = st.columns(2)
            with col_r:
                if st.button(
                    "📐 Fizică Real\n*(Matematică-Informatică, Științe ale naturii)*",
                    key="_pick_fizica_real",
                    use_container_width=True,
                    type="primary",
                ):
                    update_system_prompt_for_subject("fizică_real")
                    st.session_state["_detected_subject"] = "fizică_real"
                    st.session_state.pop("_pending_user_msg", None)
                    st.session_state.pop("_pending_fizica_ambigua", None)
                    st.session_state["_suggested_question"] = _pending_msg
                    st.rerun()
            with col_t:
                if st.button(
                    "🔧 Fizică Tehnologic\n*(Filiera tehnologică)*",
                    key="_pick_fizica_tehnologic",
                    use_container_width=True,
                ):
                    update_system_prompt_for_subject("fizică_tehnologic")
                    st.session_state["_detected_subject"] = "fizică_tehnologic"
                    st.session_state.pop("_pending_user_msg", None)
                    st.session_state.pop("_pending_fizica_ambigua", None)
                    st.session_state["_suggested_question"] = _pending_msg
                    st.rerun()
        st.stop()

    with st.chat_message("assistant"):
        st.markdown("**La ce materie se referă întrebarea ta?** Alege una din opțiunile de mai jos:")
        # Butoane pentru fiecare materie (fără Automat)
        _materii_optiuni = [(k, v) for k, v in MATERII.items() if v is not None]
        _cols = st.columns(3)
        for i, (label, cod) in enumerate(_materii_optiuni):
            with _cols[i % 3]:
                if st.button(label, key=f"_pick_materie_{cod}", use_container_width=True):
                    # Setăm materia și trimitem mesajul original
                    update_system_prompt_for_subject(cod)
                    st.session_state["_detected_subject"] = cod
                    st.session_state.pop("_pending_user_msg", None)
                    st.session_state["_suggested_question"] = _pending_msg
                    st.rerun()

    st.stop()

# ── Handler întrebare sugerată — ÎNAINTE de afișarea butoanelor ──
if st.session_state.get("_suggested_question"):
    user_input = st.session_state.pop("_suggested_question")
    with st.chat_message("user"):
        st.markdown(user_input)
    st.session_state.messages.append({"role": "user", "content": user_input})
    save_message_with_limits(st.session_state.session_id, "user", user_input)

    # ── Detecție și routing materie ──
    _materie_manuala = st.session_state.get("materie_selectata")
    _mod_automat = (_materie_manuala is None)

    if not _mod_automat:
        if st.session_state.get("_detected_subject") != _materie_manuala:
            update_system_prompt_for_subject(_materie_manuala)
    else:
        _detected = detect_subject_from_text(user_input)
        _prev_detected = st.session_state.get("_detected_subject")
        if _detected == "_fizica_ambigua":
            # Fizică detectată, profil necunoscut — cerem alegerea profilului
            st.session_state["_pending_user_msg"] = user_input
            st.session_state["_pending_fizica_ambigua"] = True
            st.rerun()
        elif _detected and _detected != _prev_detected:
            update_system_prompt_for_subject(_detected)
            _det_label = _MATERII_LABEL.get(_detected, _detected.capitalize())
            st.toast(f"📚 {_det_label}", icon="🎯")
        elif not _detected and not _prev_detected:
            st.session_state["_pending_user_msg"] = user_input
            st.rerun()

    context_messages = get_context_for_ai(st.session_state.messages)
    history_obj = []
    for msg in context_messages:
        role_gemini = "model" if msg["role"] == "assistant" else "user"
        history_obj.append({"role": role_gemini, "parts": [msg["content"]]})

    with st.chat_message("assistant"):
        message_placeholder = st.empty()
        full_response = ""
        message_placeholder.markdown(TYPING_HTML, unsafe_allow_html=True)
        try:
            for text_chunk in run_chat_with_rotation(history_obj, [user_input]):
                full_response += text_chunk
                message_placeholder.markdown(full_response + "▌")
            message_placeholder.empty()
            render_message_with_svg(full_response)
            st.session_state.messages.append({"role": "assistant", "content": full_response})
            save_message_with_limits(st.session_state.session_id, "assistant", full_response)
        except Exception as e:
            _show_ai_error(e, "întrebare sugerată")
    st.rerun()

# ── Întrebări sugerate per materie — afișate doar când chat-ul e gol ──
# Pool mare de întrebări — 4 alese aleator la fiecare sesiune nouă
INTREBARI_POOL = {
    None: [
        "Explică-mi cum se rezolvă ecuațiile de gradul 2",
        "Ce este fotosinteza și cum funcționează?",
        "Cum se scrie un eseu la BAC?",
        "Explică legea lui Ohm cu un exemplu",
        "Care sunt curentele literare studiate la BAC?",
        "Cum calculez probabilitatea unui eveniment?",
        "Explică-mi structura atomului",
        "Ce este derivata și la ce folosește?",
        "Cum rezolv o problemă cu mișcare uniformă?",
        "Explică-mi reacțiile chimice de bază",
        "Care sunt figurile de stil principale?",
        "Cum funcționează circuitul electric serie vs paralel?",
    ],
    "matematică": [
        "Cum rezolv o ecuație de gradul 2?",
        "Explică-mi derivatele — ce sunt și cum se calculează",
        "Cum calculez aria și volumul unui corp geometric?",
        "Ce este limita unui șir și cum o calculez?",
        "Cum rezolv un sistem de ecuații?",
        "Explică-mi funcțiile monotone și extreme",
        "Ce este matricea și cum fac operații cu ea?",
        "Cum calculez probabilități cu combinări?",
        "Explică-mi trigonometria — formule esențiale",
        "Cum rezolv inecuații de gradul 2?",
        "Ce sunt vectorii și cum fac operații cu ei?",
        "Explică-mi integralele — ce sunt și cum se calculează",
    ],
    "fizică_real": [
        # Clasa IX — Mecanică
        "Explică legile lui Newton cu exemple concrete",
        "Cum rezolv o problemă cu plan înclinat?",
        "Cum calculez energia cinetică și potențială?",
        "Explică mișcarea uniform accelerată — formule și grafice",
        "Ce este impulsul și cum aplic teorema impulsului?",
        "Cum calculez lucrul mecanic și puterea?",
        "Explică legea lui Arhimede — condiția de plutire",
        "Cum aplic teorema lui Bernoulli în probleme?",
        "Explică mișcarea circulară uniformă — formule",
        "Ce sunt legile lui Kepler și vitezele cosmice?",
        # Clasa X — Termodinamică + Electricitate
        "Ce este legea lui Ohm și cum aplic în circuit?",
        "Cum rezolv o problemă cu circuite mixte (serie+paralel)?",
        "Explică transformările gazelor ideale (izoterm, izobar, izocor)",
        "Cum calculez randamentul unui motor termic?",
        "Ce este curentul alternativ — valori eficace, impedanță?",
        "Explică transformatorul — cum funcționează?",
        "Cum aplic legile lui Kirchhoff într-un circuit?",
        # Clasa XI — Oscilații, unde, optică
        "Explică oscilațiile armonice — pendul și resort",
        "Ce este rezonanța și când apare?",
        "Cum calculez lungimea de undă și viteza unei unde?",
        "Explică interferența undelor — Young",
        "Ce este difracția și cum aplic formula rețelei?",
        "Explică spectrul electromagnetic — tipuri și aplicații",
        "Ce este polarizarea luminii — legea Malus?",
        # Clasa XII — Fizică modernă
        "Explică dilatarea timpului în relativitatea restrânsă",
        "Ce este efectul fotoelectric — ecuația lui Einstein?",
        "Explică modelul Bohr al atomului de hidrogen",
        "Cum calculez energia de legătură a unui nucleu?",
        "Explică dezintegrarea α, β, γ — legi de conservare",
        "Ce este fisiunea nucleară și cum funcționează reactorul?",
        "Explică ipoteza de Broglie — dualism undă-corpuscul",
        "Ce sunt semiconductorii N și P — joncțiunea PN?",
    ],
    "fizică_tehnologic": [
        # Clasa IX — Mecanică aplicată
        "Explică legile lui Newton cu exemple din tehnologie",
        "Cum calculez forța de frecare și randamentul unui plan înclinat?",
        "Cum calculez energia cinetică și potențială — probleme practice",
        "Explică mișcarea uniform accelerată — formule și aplicații",
        "Cum calculez lucrul mecanic și puterea unui motor?",
        "Explică legea lui Arhimede — aplicații în inginerie",
        # Clasa X — Termodinamică + Curent continuu
        "Ce este legea lui Ohm și cum o aplic într-un circuit DC?",
        "Cum rezolv un circuit serie și unul paralel?",
        "Cum calculez puterea și energia electrică consumată?",
        "Explică transformările gazelor (izoterm, izobar, izocor) cu grafice",
        "Cum calculez randamentul unui motor termic?",
        "Cum aplic legile lui Kirchhoff — curenți și tensiuni în nod?",
        # Clasa XI — Optică și curent alternativ
        "Cum funcționează o lentilă convergentă — formula oglinzilor/lentilelor?",
        "Explică reflexia și refracția luminii cu aplicații practice",
        "Ce este curentul alternativ — tensiune eficace și frecvență?",
        "Cum funcționează un transformator electric?",
        "Explică spectrul electromagnetic — aplicații în tehnologie",
    ],
    "chimie": [
        # Clasa IX — Anorganică & baze fizico-chimice
        "Explică structura atomului și configurația electronică",
        "Cum determin tipul de legătură chimică (ionică, covalentă)?",
        "Ce este echilibrul chimic și principiul Le Châtelier?",
        "Cum calculez pH-ul unui acid/bază tare?",
        "Explică reacțiile redox — oxidare, reducere, bilanț electronic",
        "Cum funcționează pila Daniell — anod, catod, tensiune?",
        "Cum calculez concentrația molară și fac diluții?",
        "Explică coroziunea fierului și metodele de protecție",
        # Clasa X — Organică introductivă
        "Explică-mi alcanii — structură, denumire, reacții",
        "Ce este regula lui Markovnikov — cum o aplic la alchene?",
        "Cum calculez gradul de nesaturare Ω?",
        "Explică izomeria structurală — de catenă, poziție, funcțiune",
        "Cum echilibrez o ecuație chimică pas cu pas?",
        "Cum fac calcule stoechiometrice — cei 5 pași?",
        "Explică reacțiile de esterificare și saponificare",
        "De ce alcoolii au punct de fierbere ridicat? (legături H)",
        "Cum funcționează săpunul — mecanismul spălării?",
        # Clasa XI-XII — Organică avansată & biochimie
        "Explică substituția nucleofilă SN la derivații halogenați",
        "Cum deosebesc aldehidele de cetone? (Tollens, Fehling)",
        "Explică polimerizarea și policondensarea — PVC, nylon",
        "Ce sunt aminoacizii — comportament amfoter, legătură peptidică?",
        "Explică structura glucozei și fructozei — test Fehling",
        "Care este diferența dintre amidon și celuloză?",
        "Explică structura ADN — baze azotate, legături de hidrogen",
        "Ce sunt trigliceridele și cum se face saponificarea grăsimilor?",
        "Explică impactul freonilor asupra stratului de ozon",
    ],
    "limba și literatura română": [
        "Cum structurez un eseu de BAC la Română?",
        "Explică-mi curentele literare principale",
        "Cum analizez o poezie — figuri de stil, prozodie",
        "Care sunt operele obligatorii la BAC Română?",
        "Explică-mi romanul Ion de Rebreanu",
        "Cum caracterizez un personaj literar?",
        "Ce figuri de stil sunt la Eminescu în Luceafărul?",
        "Cum scriu comentariul unui text narativ?",
        "Explică-mi analiza morfologică și sintactică",
        "Care sunt trăsăturile romantismului românesc?",
        "Cum analizez Enigma Otiliei de Călinescu?",
        "Ce este modernismul în literatura română?",
    ],
    "biologie": [
        "Explică-mi mitoza vs meioza",
        "Cum funcționează fotosinteza și respirația celulară?",
        "Ce este ADN-ul și cum funcționează codul genetic?",
        "Explică-mi legile lui Mendel cu pătrat Punnett",
        "Care sunt organitele celulei și funcțiile lor?",
        "Cum funcționează sistemul nervos?",
        "Explică-mi sistemul circulator — inimă și sânge",
        "Ce este fotosinteza — faza luminoasă și Calvin?",
        "Cum funcționează sistemul digestiv?",
        "Explică determinismul sexului și bolile genetice",
        "Ce este ecosistemul și lanțul trofic?",
        "Cum funcționează sistemul endocrin?",
    ],
    "informatică": [
        # Clasa IX - Python baze
        "Explică sortarea prin selecție în Python pas cu pas",
        "Cum funcționează algoritmul lui Euclid pentru cmmdc?",
        "Ce sunt listele în Python? Metode: append, pop, sort",
        "Cum implementez o stivă și o coadă în Python?",
        "Ce este recursivitatea? Exemplu cu factorial în Python",
        "Cum citesc și scriu fișiere text în Python?",
        "Explică-mi funcțiile în Python — parametri și return",
        "Cum fac o interfață grafică simplă cu Tkinter?",
        # Clasa X - colecții + algoritmi
        "Cum funcționează dicționarele în Python? (dict)",
        "Explică diferența dintre set, list, tuple și dict",
        "Cum funcționează căutarea binară? Cod Python",
        "Explică Merge Sort — Divide et Impera în Python",
        "Cum implementez cifrul Cezar în Python?",
        "Ce este QuickSort și cum funcționează?",
        "Cum lucrez cu matrici (tablouri 2D) în Python și C++?",
        "Explică-mi struct în C++ cu exemple",
        # Clasa XI - grafuri, arbori, algoritmi avansați
        "Ce sunt grafurile? Explică BFS și DFS cu exemple",
        "Cum funcționează algoritmul Dijkstra?",
        "Explică backtracking-ul cu problema N-Reginelor",
        "Ce este programarea dinamică? Exemplu cu rucsacul",
        "Cum implementez un arbore binar de căutare?",
        "Explică algoritmii Prim și Kruskal pentru MST",
        "Ce sunt listele înlănțuite și cum le implementez?",
        "Roy-Floyd — drumuri minime între toate perechile",
        # Clasa XII - BD, SQL, ML
        "Explică-mi modelul entitate-relație (ERD)",
        "SQL: cum fac un JOIN între două tabele?",
        "Ce este normalizarea bazelor de date? FN1, FN2, FN3",
        "Cum conectez Python la o bază de date SQLite?",
        "Introduc-ți Pandas — DataFrame și operații de bază",
        "Cum antrenez un model KNN cu scikit-learn?",
        "Ce este K-Means și cum funcționează clustering-ul?",
        "Explică regresia liniară cu un exemplu în Python",
    ],
    "geografie": [
        "Care sunt unitățile de relief ale României?",
        "Explică-mi clima României — regiuni și factori",
        "Care sunt râurile principale din România?",
        "Explică formarea Munților Carpați",
        "Care sunt vecinii României și granițele?",
        "Explică-mi Delta Dunării — caracteristici",
        "Care sunt resursele naturale ale României?",
        "Explică populația și orașele mari din România",
        "Ce sunt continentele — caracteristici principale?",
        "Explică-mi coordonatele geografice",
        "Care sunt problemele de mediu din România?",
        "Explică clima Europei — zone climatice",
    ],
    "istorie": [
        "Explică Marea Unire din 1918 — cauze și consecințe",
        "Care au fost reformele lui Alexandru Ioan Cuza?",
        "Explică-mi perioada comunistă în România",
        "Ce s-a întâmplat la Revoluția din 1989?",
        "Cine a fost Ștefan cel Mare și care sunt realizările lui?",
        "Explică Primul Război Mondial — România",
        "Ce a fost Revoluția de la 1848 în Țările Române?",
        "Explică domnia lui Mihai Viteazul și prima unire",
        "Care au fost cauzele Independenței din 1877?",
        "Explică perioada interbelică în România",
        "Ce a fost Holocaustul și implicarea României?",
        "Cine a fost Carol I și ce a realizat?",
    ],
    "limba franceză": [
        "Explică-mi Passé Composé vs Imparfait",
        "Cum se acordă participiul trecut cu avoir și être?",
        "Explică Subjonctivul — când și cum se folosește",
        "Cum structurez un eseu în franceză?",
        "Explică-mi Futur Simple vs Futur Proche",
        "Cum funcționează pronumele relative (qui, que, dont)?",
        "Explică condiționalul prezent și trecut",
        "Ce sunt verbele neregulate esențiale în franceză?",
        "Cum exprim cauza și consecința în franceză?",
        "Explică-mi acordul adjectivelor în franceză",
    ],
    "limba engleză": [
        "Explică Present Perfect vs Past Simple",
        "Cum funcționează propozițiile condiționale (tip 1, 2, 3)?",
        "Explică vocea pasivă în engleză",
        "Cum scriu un eseu argumentativ în engleză?",
        "Explică reported speech — vorbire indirectă",
        "Ce sunt modal verbs și când le folosesc?",
        "Cum funcționează articolele a/an/the în engleză?",
        "Explică-mi timpurile verbale — ghid complet",
        "Cum scriu o scrisoare formală în engleză?",
        "Explică relative clauses (who, which, that)",
    ],
    "limba germană": [
        # Clasa IX — A1/A2, baze
        "Explică genul substantivelor în germană — der, die, das",
        "Cum conjugăm verbele la prezent (Präsens) în germană?",
        "Explică ordinea cuvintelor în propoziția germană (Satzstellung)",
        "Ce sunt verbele modale în germană? können, müssen, dürfen...",
        "Cum formăm întrebările în germană — W-Fragen și Da/Nein-Fragen?",
        "Explică-mi cazurile în germană — nominativ, acuzativ, dativ",
        "Cum funcționează verbele separabile (trennbare Verben)?",
        # Clasa X — A2/B1
        "Explică Perfekt vs Präteritum — când folosesc fiecare?",
        "Cum se formează Partizip II pentru Perfekt?",
        "Care verbe cer sein și care haben la Perfekt?",
        "Cum compar adjectivele în germană? gut → besser → am besten",
        "Explică propoziția subordonată cu weil, dass, obwohl — verbul la sfârșit!",
        "Ce sunt verbele reflexive în germană? (sich waschen, sich freuen)",
        # Clasa XI — B1
        "Explică Konjunktiv II — würde, wäre, hätte și când îl folosesc",
        "Cum funcționează pronumele relative în germană?",
        "Explică Passiv în germană — werden + Partizip II",
        "Cum scriu un CV și o scrisoare de intenție în germană?",
        "Explică Plusquamperfekt — mai-mult-ca-perfectul în germană",
        # Clasa XII — B1/B2
        "Explică Konjunktiv I pentru vorbire indirectă (Indirekte Rede)",
        "Ce sunt conectorii dubli în germană? entweder…oder, sowohl…als auch",
        "Cum structurez un eseu argumentativ în germană (BAC)?",
        "Explică Infinitivkonstruktionen cu zu în germană",
        "Cum folosesc genitivul în texte formale germane?",
    ],
}

if not st.session_state.get("messages") and not st.session_state.get("pedagogie_mode"):
    materie_curenta = st.session_state.get("materie_selectata")

    if materie_curenta is None:
        # Mod Automat — afișăm selector de materie pe pagina principală
        st.markdown("##### 📚 Selectează materia")
        _materii_butoane = [(k, v) for k, v in MATERII.items() if v is not None]
        _cols = st.columns(2)
        for i, (label, cod) in enumerate(_materii_butoane):
            with _cols[i % 2]:
                if st.button(label, key=f"pick_mat_{cod}", use_container_width=True):
                    # Setăm materia în selector și în session_state
                    st.session_state.materie_selectata = cod
                    st.session_state["_detected_subject"] = cod
                    st.session_state["system_prompt"] = get_system_prompt(
                        materie=cod,
                        pas_cu_pas=st.session_state.get("pas_cu_pas", False),
                        mod_avansat=st.session_state.get("mod_avansat", False),
                        mod_strategie=st.session_state.get("mod_strategie", False),
                        mod_bac_intensiv=st.session_state.get("mod_bac_intensiv", False),
                    )
                    st.rerun()
    else:
        # Materie selectată — afișăm întrebări sugerate pentru materia respectivă
        pool = INTREBARI_POOL.get(materie_curenta, INTREBARI_POOL[None])
        _sugg_key = f"_sugg_list_{st.session_state.session_id}"
        _sugg_materie_key = f"_sugg_materie_{st.session_state.session_id}"
        if (
            _sugg_key not in st.session_state
            or st.session_state.get(_sugg_materie_key) != materie_curenta
        ):
            st.session_state[_sugg_key] = random.sample(pool, min(4, len(pool)))
            st.session_state[_sugg_materie_key] = materie_curenta
        intrebari = st.session_state[_sugg_key]

        col_title, col_refresh = st.columns([4, 1])
        with col_title:
            st.markdown("##### 💡 Cu ce începem azi?")
        with col_refresh:
            if st.button("🔄", key="_refresh_sugg_btn", help="Alte întrebări"):
                st.session_state.pop(_sugg_key, None)
                st.rerun()
        cols = st.columns(2)
        for i, intrebare in enumerate(intrebari):
            with cols[i % 2]:
                if st.button(intrebare, key=f"sugg_{i}", use_container_width=True):
                    st.session_state["_suggested_question"] = intrebare
                    st.rerun()

# === AVERTISMENT OFFLINE ===
if st.session_state.get("_history_may_be_incomplete"):
    st.warning(
        "📴 **Mod offline** — istoricul afișat poate fi incomplet față de baza de date. "
        "Reconectarea se face automat când rețeaua revine.",
        icon="⚠️"
    )
    if st.button("🔄 Verifică conexiunea acum", key="_check_conn_btn"):
        # Forțăm re-marcarea ca online pentru a testa
        st.session_state.pop("_sb_online", None)
        st.session_state.pop("_history_may_be_incomplete", None)
        st.rerun()

# === HANDLER RETRY după eroare de cheie API ===
# Dacă utilizatorul a apăsat "Reîncercați" după o eroare de cheie, reluăm cererea
# cu aceleași history + payload salvate ÎNAINTE de eroare.
if st.session_state.pop("_pending_retry", False):
    _retry_history  = st.session_state.get("_retry_history")
    _retry_payload  = st.session_state.get("_retry_payload")
    if _retry_history is not None and _retry_payload is not None:
        if media_content:
            # Payload-ul salvat poate conține fișiere urcate cu o cheie veche; sidebar-ul tocmai
            # le-a verificat/re-urcat cu cheia curentă, deci le înlocuim (în ordine) cu versiunile actuale.
            _cur_media = iter(media_content if isinstance(media_content, list) else [media_content])
            _retry_payload = [
                next(_cur_media, _p) if (not isinstance(_p, str) and hasattr(_p, "uri")) else _p
                for _p in _retry_payload
            ]
        with st.chat_message("assistant"):
            _rph = st.empty()
            _rph.markdown(TYPING_HTML, unsafe_allow_html=True)
            _rfull = ""
            try:
                for _chunk in run_chat_with_rotation(_retry_history, _retry_payload):
                    _rfull += _chunk
                    if "<svg" in _rfull or ("<path" in _rfull and "stroke=" in _rfull):
                        _rph.markdown(_rfull.split("<path")[0] + "\n\n*🎨 Domnul Profesor desenează...*\n\n▌")
                    else:
                        _rph.markdown(_rfull + "▌")
                _rph.empty()
                render_message_with_svg(_rfull)
                st.session_state.messages.append({"role": "assistant", "content": _rfull})
                save_message_with_limits(st.session_state.session_id, "assistant", _rfull)
                st.session_state.pop("_retry_history", None)
                st.session_state.pop("_retry_payload", None)
            except Exception as _re:
                _rph.empty()
                _show_ai_error(_re, "reîncercare")
    st.stop()

# === CHAT INPUT ===
if user_input := st.chat_input("Întreabă profesorul..."):

    # --- Rate limiting (minut + zi per sesiune + global) ---
    if not enforce_rate_limit():
        st.stop()

    # --- Debounce: blochează mesaje duplicate trimise rapid ---
    now_ts = time.time()
    last_msg = st.session_state.get("_last_user_msg", "")
    last_ts  = st.session_state.get("_last_msg_ts", 0)
    DEBOUNCE_SECONDS = 2.5

    if user_input.strip() == last_msg.strip() and (now_ts - last_ts) < DEBOUNCE_SECONDS:
        st.toast("⏳ Mesaj duplicat ignorat.", icon="🔁")
        st.stop()

    st.session_state["_last_user_msg"] = user_input
    st.session_state["_last_msg_ts"]  = now_ts

    # FIX BUG 1: Afișează și salvează mesajul utilizatorului ÎNAINTE de răspunsul AI
    with st.chat_message("user"):
        st.markdown(user_input)
    st.session_state.messages.append({"role": "user", "content": user_input})
    save_message_with_limits(st.session_state.session_id, "user", user_input)

    # FIX PERSISTENȚĂ FIȘIER: dacă media_content/text_file_content sunt None
    # (ex: widgetul st.file_uploader și-a pierdut valoarea după un rerun programatic —
    # schimbare materie, toggle mod etc.), recuperăm fișierul activ direct din
    # session_state, folosind cheia salvată la upload. Fișierul de pe Google rămâne
    # valid (TTL 48h) și textul extras local rămâne în cache — doar referința locală
    # `uploaded_file` se pierdea. Trebuie făcut ÎNAINTE de detecția de materie de mai
    # jos, care depinde de text_file_content.
    if not media_content:
        _rec_media = []
        for _k in st.session_state.get("_active_gfile_keys", []) or []:
            _gf_check = st.session_state.get(_k)
            try:
                if _gf_check and _is_gfile_active(_gf_check):
                    _rec_media.append(_gf_check)
            except Exception:
                pass
        media_content = _rec_media or None

    if not text_file_content:
        _rec_txt = [st.session_state.get(_k) for _k in (st.session_state.get("_active_textcache_keys", []) or [])]
        _rec_txt = [t for t in _rec_txt if t]
        if len(_rec_txt) == 1:
            text_file_content = _rec_txt[0]
        elif _rec_txt:
            text_file_content = "\n\n".join(_rec_txt)

    # ── Detecție și routing materie ──
    _materie_manuala = st.session_state.get("materie_selectata")  # None = mod Automat
    _mod_automat = (_materie_manuala is None)

    # Dacă elevul a încărcat un fișier text (SRT, docx, txt, dbf), mesajul descrie
    # o operație pe fișier ("traduce din engleză în română", "rezumă", etc.) și conține
    # cuvinte-cheie de limbi/materii care ar declanșa fals detecția.
    # Excludem complet detecția de materie în acest caz.
    _ufm = st.session_state.get("_current_uploaded_file_meta", {}) or {}
    _has_text_file_uploaded = bool(
        (_ufm.get("type") == "multi"
         or _ufm.get("name", "").lower().split(".")[-1] in ("srt", "txt", "docx", "doc", "dbf"))
        and text_file_content  # fișierul chiar a fost încărcat și extras
    )

    if not _mod_automat:
        # Mod manual: selectorul are prioritate — asigurăm prompt-ul corect
        if st.session_state.get("_detected_subject") != _materie_manuala:
            update_system_prompt_for_subject(_materie_manuala)
        # FIX Bug 3: avertizăm dacă textul pare să fie pentru altă materie
        # (detectăm din mesaj, comparăm cu selecția manuală — toast non-blocant)
        # EXCEPȚIE: dacă e un fișier text încărcat, nu avertizăm — mesajul conține
        # inevitabil cuvinte de limbi/materii (ex: "traduce din engleză în română")
        if not _has_text_file_uploaded:
            _detected_in_msg = detect_subject_from_text(user_input)
            if (
                _detected_in_msg
                and _detected_in_msg != _materie_manuala
                and _detected_in_msg != "pedagogie"
                and not st.session_state.get(f"_mismatch_warned_{st.session_state.session_id}")
            ):
                _sel_label = _MATERII_LABEL.get(_materie_manuala, _materie_manuala or "materia selectată")
                _det_label = _MATERII_LABEL.get(_detected_in_msg, _detected_in_msg.capitalize())
                st.toast(
                    f"💡 Mesajul pare să fie despre {_det_label}, dar ești pe {_sel_label}. "
                    f"Schimbă materia din sidebar dacă vrei răspuns specializat.",
                    icon="⚠️"
                )
                st.session_state[f"_mismatch_warned_{st.session_state.session_id}"] = True

    else:
        # Mod automat: detectăm materia DOAR la primul mesaj din conversație.
        # Odată ce materia e stabilită (_detected_subject setat) și conversația
        # a început (există cel puțin un mesaj), NU mai re-detectăm — rămânem
        # pe materia identificată la început, chiar dacă elevul pune o întrebare
        # care conține termeni din altă materie (ex: o problemă de fizică cu
        # termeni matematici nu trebuie să schimbe contextul în matematică).
        _prev_detected = st.session_state.get("_detected_subject")
        _conv_started = bool(_prev_detected)  # materia a fost deja stabilită

        if _conv_started:
            # Conversație în desfășurare — păstrăm materia detectată la început
            # și ne asigurăm că prompt-ul e corect (ex: după switch sesiune)
            if st.session_state.get("system_prompt") is None:
                update_system_prompt_for_subject(_prev_detected)
        elif _has_text_file_uploaded:
            # Primul mesaj dar cu fișier text — nu detectăm din mesaj (ar fi fals positiv).
            # Lăsăm materia nedeterminată; profesorul va răspunde generic.
            pass
        else:
            # Chat nou sau fără materie stabilită — rulăm detecția
            _detected = detect_subject_from_text(user_input)

            if _detected == "_fizica_ambigua":
                # Fizică detectată, dar profil necunoscut — cerem elevului să aleagă
                st.session_state["_pending_user_msg"] = user_input
                st.session_state["_pending_fizica_ambigua"] = True
                st.rerun()
            elif _detected:
                # Detectat cu succes — blocăm materia pentru această conversație
                update_system_prompt_for_subject(_detected)
                _det_label = _MATERII_LABEL.get(_detected, _detected.capitalize())
                st.toast(f"📚 {_det_label}", icon="🎯")
                for _k in [k for k in st.session_state.keys() if k.startswith("_mismatch_warned_")]:
                    del st.session_state[_k]
            else:
                # Nu s-a putut detecta materia — salvăm mesajul și întrebăm elevul
                st.session_state["_pending_user_msg"] = user_input
                st.rerun()

    context_messages = get_context_for_ai(st.session_state.messages)
    history_obj = []
    for msg in context_messages:
        role_gemini = "model" if msg["role"] == "assistant" else "user"
        history_obj.append({"role": role_gemini, "parts": [msg["content"]]})
    
    final_payload = _build_attachment_payload(media_content, text_file_content)
    final_payload.append(user_input)

    # ═══════════════════════════════════════════════════════════════════════════
    # TRADUCERE SRT ÎN BUCĂȚI — dacă fișierul e .srt și cererea implică traducere,
    # împărțim subtitrarea în bucăți de SRT_CHUNK_SIZE replici și le traduc pe rând.
    # Altfel, folosim flow-ul normal de chat.
    # ═══════════════════════════════════════════════════════════════════════════

    def _parse_srt_blocks(srt_text: str) -> list[dict]:
        """Parsează SRT-ul în blocuri structurate: {index, timestamp, text}.
        Timestamp-urile sunt extrase și salvate SEPARAT — modelul nu le va atinge."""
        result = []
        _ts_re     = re.compile(r'\d{2}:\d{2}:\d{2}[,.]\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}[,.]\d{3}')
        _ts_single = re.compile(r'^\d{2}:\d{2}:\d{2}[,.]\d{3}')  # timestamp incomplet (fără -->)
        for block in re.split(r'\n\s*\n', srt_text.strip()):
            block = block.strip()
            if not block:
                continue
            block_lines = block.splitlines()
            if len(block_lines) < 2:
                continue
            idx_line = block_lines[0].strip()
            ts_line  = block_lines[1].strip() if len(block_lines) > 1 else ""
            if _ts_re.match(ts_line):
                # Format normal: linia 0 = index, linia 1 = timestamp complet
                text_lines = block_lines[2:]
            elif _ts_re.match(idx_line):
                # Lipsește indexul — timestamp pe prima linie
                ts_line    = idx_line
                idx_line   = str(len(result) + 1)
                text_lines = block_lines[1:]
            elif _ts_single.match(ts_line):
                # Timestamp incomplet (ex: "00:02:27,231" fără "-->...") — îl păstrăm ca atare
                text_lines = block_lines[2:]
            elif _ts_single.match(idx_line):
                ts_line    = idx_line
                idx_line   = str(len(result) + 1)
                text_lines = block_lines[1:]
            else:
                continue  # bloc complet malformat — sărim
            text = "\n".join(text_lines).strip()
            if text:
                result.append({"index": idx_line, "timestamp": ts_line, "text": text})
        return result

    def _is_translation_request(text: str) -> bool:
        """Detectează dacă utilizatorul cere o traducere."""
        keywords = [
            "traduc", "translat", "română", "roman", "englez", "francez", "german",
            "spaniol", "italian", "rus", "maghiar", "trad.", "în română", "in romana",
            "din engleză", "din engleza", "convertește", "converteste",
        ]
        tl = text.lower()
        return any(kw in tl for kw in keywords)

    _uf_meta = st.session_state.get("_current_uploaded_file_meta", {})
    _is_srt  = _uf_meta.get("name", "").lower().endswith(".srt")
    _is_trad = _is_translation_request(user_input)

    SRT_CHUNK_SIZE = 200  # replici per bucată — mai mic = mai stabil

    if _is_srt and _is_trad and text_file_content:
        # ── Mod traducere SRT cu separare completă timestamps / text ──

        # 1. Parsăm SRT-ul — extragem timestamp-urile O SINGURĂ DATĂ din original
        parsed_blocks = _parse_srt_blocks(text_file_content)
        total_blocks  = len(parsed_blocks)
        chunks        = [parsed_blocks[i:i + SRT_CHUNK_SIZE]
                         for i in range(0, total_blocks, SRT_CHUNK_SIZE)]
        total_chunks  = len(chunks)

        _orig_name        = _uf_meta.get("name", "subtitrare.srt")
        _trad_name        = re.sub(r'\.srt$', '_RO.srt', _orig_name, flags=re.IGNORECASE)
        _srt_key          = f"_srt_translation_{_orig_name}"
        translated_blocks = []   # lista de dict {index, timestamp, text} cu textul tradus
        _translation_done = False

        with st.chat_message("assistant"):
            progress_placeholder = st.empty()
            progress_placeholder.info(
                f"🎬 Traduc subtitrarea în {total_chunks} bucăți "
                f"({total_blocks} replici total)... Bucată 1/{total_chunks}"
            )

            try:
                for chunk_idx, chunk in enumerate(chunks, start=1):
                    # 2. Trimitem la AI DOAR textele, numerotate simplu 1..N
                    lines_for_ai = []
                    for i, blk in enumerate(chunk, start=1):
                        lines_for_ai.append(f"[{i}] {blk['text']}")

                    def _make_chunk_prompt(lines, n_lines, attempt=1):
                        strictness = (
                            "ESTE OBLIGATORIU să traduci în română. NU returna text în engleză.\n"
                            if attempt > 1 else ""
                        )
                        return (
                            f"Ești un traducător profesionist. Traduce textele de mai jos din engleză în română.\n\n"
                            f"{strictness}"
                            f"REGULI:\n"
                            f"1. Fiecare linie începe cu un număr între paranteze pătrate [N]. "
                            f"Păstrează EXACT acel număr la începutul fiecărei linii traduse.\n"
                            f"2. Traduce DOAR textul după [N], înlocuiește complet engleza cu română.\n"
                            f"3. NU adăuga linii noi, NU omite linii, NU adăuga explicații.\n"
                            f"4. Numărul de linii din răspuns trebuie să fie EXACT {n_lines}.\n"
                            f"5. Păstrează tagurile HTML dacă există (<i>, <b>, etc.).\n\n"
                            f"TEXTE DE TRADUS:\n" + "\n".join(lines)
                        )

                    def _is_mostly_romanian(translation_map: dict, chunk_size: int) -> bool:
                        """Verifică dacă cel puțin 60% din traduceri conțin diacritice românești
                        sau sunt evident în română. Dacă mai puțin — considerăm că a rămas în engleză."""
                        if not translation_map:
                            return False
                        ro_chars = set('ăâîșțĂÂÎȘȚ')
                        ro_count = sum(
                            1 for txt in translation_map.values()
                            if any(c in ro_chars for c in txt)
                            or not any(c.isalpha() for c in txt)  # linie fără text (ex: doar simboluri)
                        )
                        # Dacă avem și puține traduceri returnate (model a omis linii) → reîncercăm
                        if len(translation_map) < chunk_size * 0.5:
                            return False
                        return ro_count >= len(translation_map) * 0.4

                    # Încearcă traducerea cu până la 3 reîncercări automate
                    MAX_RETRIES = 3
                    translation_map = {}
                    for attempt in range(1, MAX_RETRIES + 1):
                        chunk_response = ""
                        for text_chunk in run_chat_with_rotation([], [_make_chunk_prompt(lines_for_ai, len(chunk), attempt)]):
                            chunk_response += text_chunk

                        # Curățăm markdown fences
                        chunk_response = re.sub(r'```[a-zA-Z]*\n?', '', chunk_response.strip())
                        chunk_response = chunk_response.strip()

                        # Parsăm răspunsul
                        translation_map = {}
                        for line in chunk_response.splitlines():
                            line = line.strip()
                            m = re.match(r'\[(\d+)\]\s*(.*)', line)
                            if m:
                                n   = int(m.group(1))
                                txt = m.group(2).strip()
                                if txt:
                                    translation_map[n] = txt

                        if _is_mostly_romanian(translation_map, len(chunk)):
                            break  # traducere OK — ieșim din loop de retry

                        # Traducere proastă — actualizăm progress și reîncercăm
                        if attempt < MAX_RETRIES:
                            progress_placeholder.warning(
                                f"⚠️ Bucata {chunk_idx}/{total_chunks} a rămas în engleză "
                                f"— reîncercare {attempt}/{MAX_RETRIES - 1}..."
                            )
                        else:
                            # Toate reîncercările au eșuat — păstrăm ce avem (poate fi parțial tradus)
                            progress_placeholder.warning(
                                f"⚠️ Bucata {chunk_idx}/{total_chunks}: traducerea automată a eșuat "
                                f"după {MAX_RETRIES} încercări — s-a păstrat textul original."
                            )

                    # 4. Reconstruim blocurile cu timestamp-urile ORIGINALE + textul tradus
                    for i, blk in enumerate(chunk, start=1):
                        translated_text = translation_map.get(i, blk["text"])  # fallback = original
                        translated_blocks.append({
                            "index":     blk["index"],
                            "timestamp": blk["timestamp"],   # 100% original, neatins de AI
                            "text":      translated_text,
                        })

                    # Actualizăm progresul
                    done_count = min(chunk_idx * SRT_CHUNK_SIZE, total_blocks)
                    progress_placeholder.info(
                        f"🎬 Tradus {chunk_idx}/{total_chunks} bucăți "
                        f"({done_count}/{total_blocks} replici)"
                        + (f"... Bucată {chunk_idx + 1}/{total_chunks} în curs..."
                           if chunk_idx < total_chunks else " ✅")
                    )

                # 5. Asamblăm fișierul SRT final
                srt_output_parts = []
                for blk in translated_blocks:
                    srt_output_parts.append(f"{blk['index']}\n{blk['timestamp']}\n{blk['text']}")
                full_translation = "\n\n".join(srt_output_parts)

                progress_placeholder.success(
                    f"✅ Traducere completă! {total_blocks} replici traduse în {total_chunks} bucăți."
                )

                # Salvăm în session_state
                st.session_state[_srt_key] = {
                    "text":      full_translation,
                    "filename":  _trad_name,
                    "orig_name": _orig_name,
                    "blocks":    total_blocks,
                }

                # Mesaj vizibil în chat (marker compact)
                _save_content = (
                    f"✅ Traducere completă: {total_blocks} replici din '{_orig_name}'.\n"
                    f"[SRT_TRANSLATION_KEY:{_srt_key}]"
                )
                st.session_state.messages.append({"role": "assistant", "content": _save_content})
                save_message_with_limits(st.session_state.session_id, "assistant", _save_content)

                # Mesaj ascuns cu textul SRT complet — role "srt_data", invizibil în chat,
                # folosit exclusiv pentru restaurarea după refresh/reload
                _srt_backup_content = f"[SRT_DATA:{_srt_key}]\n{full_translation}"
                save_message_with_limits(st.session_state.session_id, "srt_data", _srt_backup_content)

                _translation_done = True

            except Exception as e:
                progress_placeholder.empty()
                err_str = str(e)
                _is_key_err = any(x in err_str for x in ["epuizat", "invalide", "quota", "429", "API key"])
                if _is_key_err:
                    _show_ai_error(e, "traducere SRT")
                else:
                    _show_ai_error(e, "traducere SRT")

        # ── Butonul de descărcare — ÎN AFARA with st.chat_message ──
        # Folosim session_state pentru date, cheie FIXĂ (nu depinde de _orig_name variabil)
        # ca să fie vizibil și după rerun-uri.
        _srt_ready = st.session_state.get(_srt_key) if _translation_done else None
        if _srt_ready:
            st.markdown(f"**📄 Subtitrare tradusă — {_srt_ready['blocks']} replici:**")
            st.download_button(
                label="⬇️ Descarcă subtitrarea tradusă (.srt)",
                data=_srt_ready["text"].encode("utf-8"),
                file_name=_srt_ready["filename"],
                mime="text/plain",
                use_container_width=True,
                key="_dl_srt_fresh",
            )
            # Afișăm TOT fișierul tradus în chat — fără trunchiere
            st.text(_srt_ready["text"])


    else:
        # ── Flow normal de chat (non-SRT sau non-traducere) ──

        # Salvăm payload-ul ÎNAINTE de apelul AI — dacă cheia se epuizează în stream,
        # elevul poate reîncerca fără să retrimită mesajul manual.
        st.session_state["_retry_history"] = history_obj
        st.session_state["_retry_payload"] = final_payload

        with st.chat_message("assistant"):
            message_placeholder = st.empty()
            full_response = ""

            # Typing indicator înainte să înceapă streaming-ul
            message_placeholder.markdown(TYPING_HTML, unsafe_allow_html=True)

            try:
                stream_generator = run_chat_with_rotation(history_obj, final_payload)
                first_chunk = True

                for text_chunk in stream_generator:
                    full_response += text_chunk
                    if first_chunk:
                        first_chunk = False  # typing indicator dispare la primul chunk

                    if "<svg" in full_response or ("<path" in full_response and "stroke=" in full_response):
                        message_placeholder.markdown(
                            full_response.split("<path")[0] + "\n\n*🎨 Domnul Profesor desenează...*\n\n▌"
                        )
                    else:
                        message_placeholder.markdown(full_response + "▌")

                message_placeholder.empty()
                render_message_with_svg(full_response)

                st.session_state.messages.append({"role": "assistant", "content": full_response})
                save_message_with_limits(st.session_state.session_id, "assistant", full_response)
                # Răspuns reușit — curățăm datele de retry
                st.session_state.pop("_retry_history", None)
                st.session_state.pop("_retry_payload", None)

            except Exception as e:
                message_placeholder.empty()
                err_str = str(e)
                # Dacă eroarea e de cheie/quota, oferim buton de reîncercare automată
                _is_key_err = any(x in err_str for x in ["epuizat", "invalide", "quota", "429", "API key"])
                if _is_key_err:
                    _show_ai_error(e, "chat")
                    # on_click: butonul nu mai e randat la rerun-ul de după click (blocul e one-shot),
                    # deci un `if st.button(...)` nu s-ar declanșa niciodată; callback-ul rulează oricum.
                    st.button("🔄 Reîncercați răspunsul", key="_retry_after_key_error", type="primary",
                              on_click=_request_retry)
                else:
                    _show_ai_error(e, "chat")
