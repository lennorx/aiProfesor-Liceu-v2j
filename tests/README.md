# Teste — Profesor Liceu (versiunea comercială)

Testele pornesc aplicația întreagă cu **Streamlit AppTest**, cu **Supabase simulat** (`fake_sb.py`, o bază
în memorie) și **Gemini simulat** (în fiecare test). Nu folosesc rețeaua și nici chei reale.

```bash
pip install -r requirements.txt
python3 tests/run_tests.py              # toate (~5–10 minute)
python3 tests/run_tests.py auth_test    # doar unul
python3 tests/run_tests.py -v           # arată fiecare verificare
APP_FILE=alt_fisier.py python3 tests/run_tests.py   # alt fișier de aplicație
```

| Test | Ce verifică |
|---|---|
| `auth_test` | cont nou, login automat pe alt dispozitiv din token, conversații pe toate dispozitivele, acces refuzat la conversația altui cont, parolă greșită, blocare după 5 încercări, Ieșire / Ieșire peste tot |
| `reset_test` | cod de recuperare, cod pe email (SMTP simulat), expirare/încercări, cod consumat, deconectare peste tot după resetare |
| `multi_test` | mai multe fișiere la „Materiale”: upload o singură dată, toate ajung la AI, eliminare 🗑️, un singur `.srt` = comportament vechi |
| `rot_test` | rotația chei × modele: cotă pe zi / pe minut, model inexistent (404), saltul direct peste combinațiile epuizate, `thinking_budget` pentru Gemini 2.5 |
| `retry_test` | butonul „🔄 Reîncercați răspunsul” chiar refă cererea |
| `modes_test` | modurile Liceu: Simulare BAC (generare + corectare), Corectare temă, Admitere — pornesc fără erori |
| `keys_test` | chei API în cont: criptare, etichete, cheie invalidă/duplicat respinse, trecerea Mama → Tata la epuizare, alt dispozitiv, ștergere |
| `photo_test` | cheia citită din poză: fără nicio cheie (mesaj), cu `OCR_API_KEY`, cheie ruptă pe 2 rânduri |

Pentru altă aplicație de profesor virtual: copiază folderul `tests/`, rulează cu `APP_FILE=...` și
adaptează testele la cheile de widget-uri diferite (ex. `_main_file_uploader`, `_auth_reg_email`).

Notă Liceu: în modul „🤖 Automat” o întrebare de fizică cere alegerea Real/Tehnologic, de aceea testele de chat
pornesc cu `materie_selectata = "fizică_real"`.

| `comercial_test` | versiunea comercială: fără cont → blocat; sub 16 ani → acordul părintelui prin link; probă + limită zilnică; raportare; plata prin părinte (Stripe simulat); 17 ani / adult; documente actualizate; export + ștergere; admin |
| `webhook_test` | webhook-ul Stripe: semnătură, plată finalizată → abonament activ, eveniment duplicat, anulare |
