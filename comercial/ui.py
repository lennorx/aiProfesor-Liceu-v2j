"""Interfața pentru partea comercială: acord parental, documente legale, abonament, GDPR, raportare, admin."""
from __future__ import annotations

import pathlib
import time

import streamlit as st

from . import billing, compliance, core

_LEGAL_DIR = pathlib.Path(__file__).resolve().parent.parent / "legal"

AI_NOTICE = ("🤖 Vorbești cu un asistent bazat pe inteligență artificială (AI), nu cu un profesor real. "
             "Poate greși — verifică informațiile importante cu manualul sau cu profesorul tău.")
NOTA_ORIENTATIVA = ("ℹ️ Nota și punctajul sunt **estimări orientative generate de AI**, nu o evaluare oficială. "
                    "Folosește-le ca să vezi ce mai ai de lucrat.")


def _set_user(user: dict):
    st.session_state["_auth_user"] = user


# ── Documente legale ──────────────────────────────────────────────────────────

def legal_text(doc: str) -> str:
    try:
        txt = (_LEGAL_DIR / f"{doc}.md").read_text(encoding="utf-8")
    except Exception:
        return f"Documentul „{core.LEGAL_TITLES.get(doc, doc)}” nu e disponibil momentan."
    c = core.company()
    for k, v in {**{f"COMPANY_{k.upper()}": str(v) for k, v in c.items()},
                 "APP_NAME": core.app_name(), "VERSIUNE": core.LEGAL_VERSIONS.get(doc, ""),
                 "APP_URL": core.base_url() or "[ADRESA SITE]", "TRIAL_DAYS": str(billing.trial_days()),
                 "CLEANUP_DAYS": str(core._DEPS.get("cleanup_days", 30))}.items():
        txt = txt.replace("{{" + k + "}}", v)
    return txt


@st.dialog("Document", width="large")
def _legal_dialog(doc: str):
    st.markdown(legal_text(doc))


def render_legal_links(key: str):
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📄 Termeni", key=f"{key}_t", use_container_width=True):
            _legal_dialog("termeni")
    with c2:
        if st.button("🔒 Confidențialitate", key=f"{key}_p", use_container_width=True):
            _legal_dialog("confidentialitate")


def render_footer():
    """Datele firmei + linkurile obligatorii ANPC (subsolul barei laterale)."""
    c = core.company()
    st.divider()
    render_legal_links("_ft")
    st.caption(
        f"{core.app_name()} este oferit de **{c['nume']}**, CUI {c['cui']}, {c['reg_com']}, {c['sediu']}. "
        f"Contact: {c['email']}" + (f", {c['telefon']}" if c.get("telefon") else "") + ".  \n"
        "[ANPC](https://anpc.ro) · "
        "[Soluționarea alternativă a litigiilor (SAL)](https://anpc.ro/ce-este-sal/)"
    )


def render_ai_notice():
    st.caption(AI_NOTICE)


def nota_orientativa():
    st.info(NOTA_ORIENTATIVA)


# ── Pagina pentru părinte (?parinte=TOKEN) ────────────────────────────────────

def render_parent_page(raw: str):
    st.title(f"👨‍👩‍👧 Acordul părintelui — {core.app_name()}")
    req = compliance.find_parent_request(raw)
    if not req:
        st.error("Linkul nu mai e valabil (a expirat sau a fost deja folosit). Cereți elevului să retrimită "
                 "emailul din aplicație.")
        st.stop()
    s = req["student"]
    if req.get("granted_at"):
        st.success(f"Ați confirmat deja acordul pentru contul {s['email']}. Mulțumim!")
        st.stop()
    st.markdown(
        f"**{s.get('display_name') or s['email']}** (`{s['email']}`) și-a creat un cont în **{core.app_name()}**, "
        f"un asistent de învățare bazat pe inteligență artificială pentru elevii de liceu.\n\n"
        f"Pentru că are sub {core.AGE_DIGITAL_CONSENT} ani, contul poate fi folosit doar cu acordul "
        f"dumneavoastră (GDPR, art. 8). Vă rugăm să citiți documentele de mai jos."
    )
    with st.expander("📄 Termeni și condiții", expanded=False):
        st.markdown(legal_text("termeni"))
    with st.expander("🔒 Politica de confidențialitate", expanded=False):
        st.markdown(legal_text("confidentialitate"))
    with st.form("_parent_consent_form"):
        a = st.checkbox("Sunt părintele sau tutorele legal al elevului.")
        b = st.checkbox("Am citit și accept Termenii și condițiile și Politica de confidențialitate, "
                        "în numele elevului.")
        c = st.checkbox("Sunt de acord ca elevul să folosească un asistent AI și înțeleg că răspunsurile "
                        "și notele lui sunt orientative și pot conține greșeli.")
        ok = st.form_submit_button("✅ Îmi dau acordul", type="primary", use_container_width=True)
    if ok:
        if not (a and b and c):
            st.error("Pentru acord trebuie bifate toate cele trei căsuțe.")
        elif compliance.grant_parental_consent(req):
            st.success("Mulțumim! Contul elevului este activ. Puteți închide această pagină.")
            st.stop()
        else:
            st.error("Acordul nu a putut fi salvat acum. Încercați din nou peste câteva minute.")
    st.divider()
    with st.expander("Nu sunt de acord / nu recunosc această cerere"):
        st.caption("Contul elevului și toate datele lui vor fi șterse definitiv.")
        if st.button("🗑️ Refuz și șterg contul", key="_parent_refuse"):
            if compliance.delete_account(req["user_id"], billing.cancel_now):
                st.success("Contul și datele au fost șterse.")
            else:
                st.error("Ștergerea nu a mers acum. Scrieți-ne la " + core.company()["email"])
    st.stop()


# ── Poarta de acces (cont / acord parental / documente / abonament) ───────────

def render_gate(ent: dict, user: dict | None):
    """Afișată în zona principală când contul nu poate folosi încă profesorul."""
    r = ent.get("reason")
    if r == "login":
        st.info("🔐 Pentru a folosi profesorul ai nevoie de un cont. Creează-l din bara laterală "
                "(secțiunea **👤 Cont**). Primești o perioadă de probă gratuită.")
    elif r == "parental":
        pe = (user or {}).get("parent_email") or ""
        st.warning(f"👨‍👩‍👧 Pentru că ai sub {core.AGE_DIGITAL_CONSENT} ani, contul se activează după ce părintele "
                   f"își dă acordul din emailul primit" + (f" la **{pe}**." if pe else "."))
        with st.form("_gate_parent_form"):
            new_pe = st.text_input("Emailul părintelui", value=pe)
            if st.form_submit_button("📧 Trimite din nou emailul", use_container_width=True):
                ok, msg = compliance.request_parental_consent(user, new_pe.strip().lower())
                (st.success if ok else st.error)(msg)
    elif r == "legal":
        under16 = core.needs_parental_consent(user.get("birth_year"), user.get("birth_month"))
        st.warning("📄 Am actualizat " + " și ".join(core.LEGAL_TITLES[d] for d in ent.get("missing", [])) + ".")
        render_legal_links("_gate_legal")
        if under16:
            st.caption("Pentru că ai sub 16 ani, părintele trebuie să accepte noile documente.")
            if st.button("📧 Trimite părintelui documentele noi", key="_gate_legal_parent"):
                ok, msg = compliance.request_parental_consent(user, user.get("parent_email") or "", renew=True)
                if ok:
                    _set_user({**user, "parental_status": "pending"})
                (st.success if ok else st.error)(msg)
        else:
            chk = st.checkbox("Am citit și accept Termenii și condițiile și Politica de confidențialitate.",
                              key="_gate_legal_chk")
            if st.button("Accept", type="primary", key="_gate_legal_btn", disabled=not chk):
                if compliance.record_acceptance(user["id"], "elev", user.get("email"), ent.get("missing")):
                    st.rerun()
                st.error("Nu am putut salva acum. Încearcă din nou.")
    elif r == "subscribe":
        st.info("⭐ Perioada de probă s-a încheiat. Alege un abonament din bara laterală (**👤 Cont → Abonament**) "
                "ca să continui.")
    st.stop()


# ── Secțiunile contului (bara laterală) ───────────────────────────────────────

def render_profile_form(user: dict, profile_bac: dict):
    with st.expander("✏️ Profilul meu (clasă, filieră, specializare)", expanded=False):
        st.caption("Profesorul își adaptează explicațiile la clasa și specializarea ta.")
        clase = ["", "IX", "X", "XI", "XII"]
        cl = st.selectbox("Clasa", clase, index=clase.index(user.get("clasa")) if user.get("clasa") in clase else 0,
                          key="_pf_clasa", format_func=lambda x: x or "— alege —")
        filiere = [""] + list(profile_bac.keys())
        fi = st.selectbox("Filiera", filiere,
                          index=filiere.index(user.get("filiera")) if user.get("filiera") in filiere else 0,
                          key="_pf_filiera", format_func=lambda x: x or "— alege —")
        profiluri = [""] + (list(profile_bac.get(fi, {}).keys()) if fi else [])
        pr = st.selectbox("Profilul", profiluri,
                          index=profiluri.index(user.get("profil_liceu")) if user.get("profil_liceu") in profiluri else 0,
                          key="_pf_profil", format_func=lambda x: x or "— alege —")
        specs = [""] + (list(profile_bac.get(fi, {}).get(pr, {}).keys()) if fi and pr else [])
        sp = st.selectbox("Specializarea", specs,
                          index=specs.index(user.get("specializare")) if user.get("specializare") in specs else 0,
                          key="_pf_spec", format_func=lambda x: x or "— alege —")
        if st.button("💾 Salvează profilul", key="_pf_save", use_container_width=True):
            upd = {"clasa": cl or None, "filiera": fi or None, "profil_liceu": pr or None, "specializare": sp or None}
            try:
                core.db().table("liceu_users").update(upd).eq("id", user["id"]).execute()
                _set_user({**user, **upd})
                st.toast("✅ Profil salvat", icon="✏️")
                st.rerun()
            except Exception as e:
                core.log("Salvarea profilului a eșuat", e)
                st.error("Profilul nu a putut fi salvat acum.")


def render_subscription(user: dict):
    if not core.billing_on():
        return
    with st.expander("⭐ Abonament", expanded=False):
        sub = billing.get_subscription(user["id"])
        st.markdown(f"**Stare:** {billing.status_text(sub)}")
        ent = billing.entitlement(user)
        if ent.get("ok") and ent.get("limit"):
            st.caption(f"Întrebări azi: {billing.usage_today(user['id'])}/{ent['limit']}")
        minor = core.is_minor(user.get("birth_year"), user.get("birth_month"))
        paid = billing.plans(paid_only=True)
        has_paid = bool(sub and sub.get("stripe_subscription_id") and sub.get("status") in ("active", "trialing", "past_due"))
        if paid and not has_paid:
            labels = {p["code"]: f"{p['name']} — {p['price_ron_cents'] / 100:.2f} lei/lună · "
                                 f"{p['daily_msg_limit']} întrebări/zi" for p in paid}
            code = st.radio("Alege planul", list(labels), format_func=labels.get, key="_sub_plan")
            if minor:
                st.caption("Pentru elevii minori plata o face părintele: îi trimitem linkul de plată pe email.")
                pe = st.text_input("Emailul părintelui", value=user.get("parent_email") or "", key="_sub_parent")
                if st.button("📧 Trimite linkul de plată părintelui", key="_sub_send", use_container_width=True):
                    ok, msg = billing.send_payment_link_to_parent(user, pe.strip().lower(), code)
                    (st.success if ok else st.error)(msg)
            else:
                if st.button("💳 Continuă la plată", key="_sub_pay", type="primary", use_container_width=True):
                    url, err = billing.create_checkout_url(user, code, user["email"])
                    if url:
                        st.session_state["_checkout_url"] = url
                    else:
                        st.error(err)
                if st.session_state.get("_checkout_url"):
                    st.link_button("Deschide pagina de plată (Stripe)", st.session_state["_checkout_url"],
                                   use_container_width=True)
            st.caption("La plată se bifează acceptarea Termenilor și începerea imediată a serviciului.")
        elif not paid and not has_paid:
            st.caption("Abonamentele vor fi disponibile în curând.")
        if has_paid:
            if minor or (sub.get("payer_email") and sub.get("payer_email") != user["email"]):
                if st.button("📧 Trimite plătitorului linkul de gestionare", key="_sub_portal_mail",
                             use_container_width=True):
                    ok, msg = billing.send_portal_link_to_payer(user["id"])
                    (st.success if ok else st.error)(msg)
            else:
                if st.button("⚙️ Gestionează abonamentul (card, facturi, anulare)", key="_sub_portal",
                             use_container_width=True):
                    url, err = billing.create_portal_url(user["id"])
                    if url:
                        st.link_button("Deschide (Stripe)", url, use_container_width=True)
                    else:
                        st.error(err)


def render_privacy_tools(user: dict, verify_password, after_delete):
    with st.expander("🔒 Datele mele (export / ștergere)", expanded=False):
        st.caption("Poți descărca toate datele contului (profil, conversații, acorduri, abonament).")
        if st.button("📦 Pregătește exportul", key="_gdpr_prep", use_container_width=True):
            st.session_state["_gdpr_export"] = compliance.export_user_data(user["id"])
        if st.session_state.get("_gdpr_export"):
            st.download_button("⬇️ Descarcă datele (JSON)", st.session_state["_gdpr_export"],
                               file_name="datele_mele_profesor_liceu.json", mime="application/json",
                               key="_gdpr_dl", use_container_width=True)
        st.divider()
        st.caption("**Ștergerea contului** e definitivă: contul, conversațiile și abonamentul (anulat) dispar.")
        with st.form("_gdpr_delete_form"):
            pw = st.text_input("Parola, pentru confirmare", type="password", key="_gdpr_pw")
            conf = st.text_input("Scrie ȘTERGE", key="_gdpr_conf")
            if st.form_submit_button("🗑️ Șterge definitiv contul"):
                if conf.strip().upper() not in ("ȘTERGE", "STERGE"):
                    st.error("Scrie ȘTERGE pentru confirmare.")
                elif not verify_password(user["email"], pw):
                    st.error("Parola nu e corectă.")
                elif compliance.delete_account(user["id"], billing.cancel_now):
                    after_delete()
                    st.success("Contul a fost șters.")
                    st.rerun()
                else:
                    st.error("Ștergerea nu a mers acum. Încearcă din nou sau scrie-ne.")


# ── Raportarea unui răspuns ───────────────────────────────────────────────────

def render_report_button(idx: int, content: str, session_id: str | None, user_id: str | None):
    with st.popover("⚑ Raportează", use_container_width=False):
        reason = st.selectbox("Ce nu e în regulă?", compliance.REPORT_REASONS, key=f"_rep_r_{idx}")
        comment = st.text_area("Detalii (opțional)", key=f"_rep_c_{idx}", max_chars=1000, height=80)
        if st.button("Trimite raportarea", key=f"_rep_s_{idx}", type="primary"):
            if compliance.save_report(user_id, session_id, content, reason, comment):
                st.success("Mulțumim! Verificăm răspunsul.")
            else:
                st.error("Raportarea nu a putut fi trimisă acum.")


# ── Panou admin ───────────────────────────────────────────────────────────────

def render_admin_panel(user: dict):
    if not user or user.get("role") != "admin":
        return
    with st.expander("👑 Administrare", expanded=False):
        client = core.db()
        try:
            users = client.table("liceu_users").select(
                "id, email, display_name, role, clasa, specializare, parental_status, created_at, last_login"
            ).order("created_at", desc=True).limit(200).execute().data or []
            subs = {s["user_id"]: s for s in (client.table("liceu_subscriptions").select("*").execute().data or [])}
        except Exception as e:
            core.log("Admin: citire conturi", e)
            st.error("Datele nu au putut fi citite.")
            return
        c1, c2, c3 = st.columns(3)
        c1.metric("Conturi", len(users))
        c2.metric("Plătite", sum(1 for s in subs.values() if s.get("stripe_subscription_id") and s.get("status") == "active"))
        c3.metric("Așteaptă părinte", sum(1 for u in users if u.get("parental_status") == "pending"))
        for u in users[:50]:
            st.caption(f"{'👑' if u.get('role') == 'admin' else '👤'} {u['email']} · "
                       f"{billing.status_text(subs.get(u['id']))}"
                       + (" · ⏳ acord parental" if u.get("parental_status") == "pending" else ""))
        st.markdown("**Acces gratuit (profesori, testeri)**")
        em = st.text_input("Email cont", key="_adm_comp_email")
        days = st.number_input("Zile (0 = nelimitat)", min_value=0, value=30, key="_adm_comp_days")
        if st.button("Acordă", key="_adm_comp_btn"):
            ok, msg = billing.grant_comp(em, None, int(days) or None)
            (st.success if ok else st.error)(msg)
        st.markdown("**Raportări noi**")
        try:
            reps = client.table("liceu_reports").select("*").eq("status", "nou") \
                .order("created_at", desc=True).limit(30).execute().data or []
        except Exception:
            reps = []
        if not reps:
            st.caption("Nicio raportare nouă.")
        for r in reps:
            st.markdown(f"**{r['reason']}** · {time.strftime('%d.%m %H:%M', time.localtime(r['created_at']))}"
                        + (f" — {r['comment']}" if r.get("comment") else ""))
            st.code((r.get("excerpt") or "")[:600], language=None)
            if st.button("✔ Rezolvat", key=f"_adm_rep_{r['id']}"):
                client.table("liceu_reports").update({"status": "rezolvat"}).eq("id", r["id"]).execute()
                st.rerun()
