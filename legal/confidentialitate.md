# Politica de confidențialitate — {{APP_NAME}}

*Versiunea {{VERSIUNE}}. PROIECT — de verificat de avocat înainte de lansare.*

## 1. Operatorul
**{{COMPANY_NUME}}**, CUI {{COMPANY_CUI}}, {{COMPANY_REG_COM}}, {{COMPANY_SEDIU}}. Contact pentru protecția datelor: {{COMPANY_EMAIL}}.

## 2. Ce date prelucrăm și de ce

| Date | Scop | Temei (GDPR art. 6) |
|---|---|---|
| Email, parolă (stocată doar ca hash scrypt), nume opțional | Contul și autentificarea | Contract (b) |
| Anul și luna nașterii, emailul părintelui | Verificarea vârstei, acordul parental (sub 16 ani), plata de către părinte (sub 18 ani) | Obligație legală (c) — art. 8 GDPR; contract (b) |
| Clasa, filiera, profilul, specializarea (opționale) | Adaptarea explicațiilor | Contract (b) |
| Întrebările, conversațiile, fișierele și pozele încărcate | Furnizarea răspunsurilor AI; istoricul conversațiilor | Contract (b) |
| Acceptarea documentelor (versiune, dată, cine a acceptat) | Dovada acordului | Obligație legală (c) |
| Abonament: plan, stare, email plătitor, ID-uri Stripe | Facturare și acces | Contract (b); obligație legală (c) pentru evidența contabilă |
| Numărul de întrebări pe zi | Limitele planului | Contract (b) |
| Raportările trimise | Siguranța conținutului | Interes legitim (f) |
| Date tehnice (identificatorul conversației și tokenul de autentificare în browser — localStorage) | Funcționarea aplicației | Contract (b) |

**Nu** folosim datele pentru publicitate, nu le vindem și nu facem profilare cu efecte juridice. Nu folosim cookie-uri de marketing sau de analiză. Datele de card sunt introduse direct la Stripe; noi nu le vedem.

## 3. Copiii
Serviciul se adresează elevilor de liceu. Pentru elevii **sub 16 ani** prelucrăm datele doar după acordul părintelui sau tutorelui, confirmat prin linkul trimis pe emailul acestuia. Până la acord, contul nu poate fi folosit; dacă acordul nu este dat, contul și datele se șterg. *[De verificat de avocat: ce înseamnă „eforturi rezonabile” de verificare (art. 8 alin. 2) și obligațiile din legea majoratului digital, dacă intră în vigoare.]*

## 4. Inteligența artificială
Întrebările și materialele sunt trimise furnizorului modelului AI pentru a genera răspunsul. *[De completat la lansare: furnizorul (ex. Anthropic) și condițiile contractuale — fără folosirea datelor la antrenare, locul prelucrării, garanții pentru transfer.]* Răspunsurile sunt generate automat și pot conține greșeli; notele sunt orientative.

## 5. Cui transmitem datele (persoane împuternicite)
- **Supabase** — baza de date (UE, Irlanda);
- **Hetzner Online** — găzduirea aplicației (UE);
- **furnizorul modelului AI** — vezi punctul 4;
- **Stripe** — plăți și facturare;
- **Google (Gmail)** — trimiterea emailurilor (coduri, acord parental, linkuri de plată);
- autorități, doar când legea o cere.

Cu fiecare avem un contract de prelucrare (DPA). Pentru transferurile în afara SEE folosim clauzele contractuale standard ale Comisiei Europene sau alte garanții legale.

## 6. Cât timp păstrăm datele
- contul și conversațiile: cât timp contul există; conversațiile neactive se șterg automat după {{CLEANUP_DAYS}} zile *[de stabilit]*;
- conturile sub 16 ani fără acord parental: se șterg după cel mult 30 de zile *[de stabilit]*;
- documentele contabile (facturi): 10 ani, conform legii;
- dovada acceptării documentelor: pe durata contului și termenul de prescripție aplicabil *[de stabilit]*.

## 7. Drepturile dumneavoastră
Aveți dreptul de acces, rectificare, ștergere, restricționare, portabilitate și opoziție, precum și dreptul de a retrage acordul.

- **Export și ștergere direct din aplicație:** bara laterală → **👤 Cont** → **🔒 Datele mele**.
- Pentru celelalte drepturi, scrieți-ne la {{COMPANY_EMAIL}}.
- Puteți depune plângere la **ANSPDCP** (https://www.dataprotection.ro).

Elevii își pot exercita singuri drepturile; părinții le pot exercita pentru copiii sub 16 ani.

## 8. Securitate
- Parolele sunt stocate doar ca hash.
- Tokenurile „ține-mă minte” sunt stocate doar ca hash și pot fi revocate oricând („Ieșire peste tot”).
- Accesul la baza de date este restricționat.
- Conexiunile sunt criptate (HTTPS).

## 9. Modificări
Vă anunțăm în aplicație despre orice modificare și vă cerem acceptarea noii versiuni.
