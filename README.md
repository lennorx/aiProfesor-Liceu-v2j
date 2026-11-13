# Profesor Liceu — versiunea comercială (branch `comercial`)

Asistent AI pentru elevii de liceu (Streamlit + Supabase). Pornește de la codul din `Profesor-Master-v2sj`, cu toate reparațiile din noiembrie 2026, și adaugă partea necesară pentru vânzare.

## Ce conține
| | Unde |
|---|---|
| Conturi (scrypt, „ține-mă minte” revocabil, resetare parolă prin email + cod de recuperare) | `app.py` |
| Vârstă la înscriere (an + lună); **sub 16 ani: acordul părintelui** prin link pe email; **sub 18 ani: plătește părintele** | `comercial/compliance.py`, `comercial/ui.py` |
| Termeni + Politică de confidențialitate cu **versiune**; acceptarea (elev / părinte) e salvată; la o versiune nouă se cere din nou | `legal/*.md`, `comercial/core.py` (`LEGAL_VERSIONS`) |
| „🤖 Vorbești cu un AI”, buton **⚑ Raportează** la fiecare răspuns, **note orientative** la BAC / teme / admitere | `comercial/ui.py` |
| **Export date (JSON)** și **ștergerea contului** din aplicație | `comercial/compliance.py` |
| Probă gratuită, planuri cu limită zilnică, **Stripe Checkout** + portal client, renunțarea la retragere bifată la plată | `comercial/billing.py` |
| Webhook Stripe → `liceu_subscriptions` | `billing_webhook/main.py` |
| Profilul elevului (clasă / filieră / specializare) folosit de AI; panou **👑 Administrare** (conturi, acces gratuit, raportări) | `app.py`, `comercial/ui.py` |
| PDF-uri mari: text extras local cu **pypdf** (fără PyMuPDF/AGPL) | `app.py` |
| Datele firmei + ANPC / SAL în subsol; documente publice la `?doc=termeni`, `?doc=confidentialitate` | `comercial/ui.py` |
| Hetzner: Docker + Caddy (HTTPS) | `Dockerfile`, `docker-compose.yml`, `deploy/` |

Totul se activează cu `BILLING_MODE = "on"` în secrets. Cu `"off"`, aplicația merge ca Master: fără cont obligatoriu și fără abonamente.

## Instalare
1. **Supabase:** rulează `supabase_liceu_auth.sql`, apoi `supabase_comercial.sql`. Ambele sunt deja aplicate în proiectul `jqmbglcuheyhpflywmgb`.
2. **Secrets:** vezi `deploy/secrets.toml.example`.
3. **Admin:** în Supabase rulează `update liceu_users set role = 'admin' where email = '...';`.
4. **Stripe și server:** vezi `deploy/README.md`.

## Teste
`python3 tests/run_tests.py`: Supabase, Gemini și Stripe sunt simulate. `comercial_test` și `webhook_test` acoperă partea comercială.

## Înainte de lansare
- **Furnizorul AI:** termenii Gemini API interzic aplicațiile folosite de minori, iar în SEE permit doar serviciile plătite. Dezvoltarea rămâne pe Gemini, dar la lansare trebuie trecut pe un furnizor care permite minori (decizie: Anthropic / Claude), cu DPA.
- **Avocatul:** verifică `legal/termeni.md`, `legal/confidentialitate.md` și textul de renunțare la retragere din `comercial/billing.py` (punctele marcate „de verificat”).
- **e-Factura:** conectează Stripe la un program de facturare (SmartBill, Oblio etc.).
- **Conturile vechi:** cele 3 conturi din tabelul vechi `users` (v2j) nu sunt mutate automat; utilizatorii își fac cont nou sau sunt migrați manual.
