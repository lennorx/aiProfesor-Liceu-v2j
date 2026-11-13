# Instalare pe Hetzner (VPS)

Rezultat: `https://DOMENIU` → aplicația; `https://DOMENIU/stripe/webhook` → webhook-ul Stripe. HTTPS e automat (Caddy + Let's Encrypt).

## 1. Serverul
- Hetzner Cloud: **CX22** (2 vCPU, 4 GB RAM) e suficient la început; locație **Falkenstein / Nürnberg / Helsinki** (UE).
- Imagine: Ubuntu 24.04. Adaugă cheia SSH la creare.
- Firewall Hetzner: permite doar 22 (SSH), 80 și 443.
- În DNS-ul domeniului: înregistrare **A** → IP-ul serverului.

## 2. Docker
```bash
ssh root@IP
apt update && apt install -y docker.io docker-compose-v2 git
git clone -b comercial https://github.com/lennorx/aiProfesor-Liceu-v2j.git /opt/profesor
cd /opt/profesor
```

## 3. Secretele (nu ajung niciodată în git)
```bash
cp deploy/secrets.toml.example .streamlit/secrets.toml   # completează valorile
cp deploy/webhook.env.example deploy/webhook.env          # completează valorile
chmod 600 .streamlit/secrets.toml deploy/webhook.env
echo "DOMAIN=profesorliceu.ro" > .env
```

## 4. Pornire / actualizare
```bash
docker compose up -d --build          # pornire
git pull && docker compose up -d --build   # actualizare după un push
docker compose logs -f app webhook    # jurnale
```

## 5. Stripe
1. **Products** → produs „Profesor Liceu – Standard”, preț **recurent lunar în RON**. Copiază `price_...` în `liceu_plans.stripe_price_id`, apoi pune `active = true`.
2. **Settings → Public details:** Terms of service = `https://DOMENIU/?doc=termeni`, Privacy policy = `https://DOMENIU/?doc=confidentialitate`. Fără ele, căsuța de acceptare din Checkout nu apare.
3. **Developers → Webhooks:**
   - endpoint `https://DOMENIU/stripe/webhook`;
   - evenimente: `checkout.session.completed`, `customer.subscription.created/updated/deleted`, `invoice.paid`, `invoice.payment_failed`;
   - copiază `whsec_...` în `deploy/webhook.env`.
4. **Settings → Billing → Customer portal:** activează anularea, schimbarea cardului și facturile.
5. **Facturi și e-Factura:** Stripe nu trimite facturi în SPV. Conectează un program de facturare cu integrare Stripe (ex. SmartBill sau Oblio), care emite factura și o transmite în e-Factura.

## 6. Verificare
- `https://DOMENIU/stripe/health` → `{"ok": true}`.
- Cont nou cu an de naștere sub 16 → emailul la părinte → confirmarea activează contul.
- Plată test cu `sk_test_...` + cardul `4242 4242 4242 4242` → în admin apare „activ”.
