#!/usr/bin/env python3
"""Rulează toate testele aplicației Profesor Liceu cu Supabase și Gemini SIMULATE (fără rețea, fără chei reale).

Utilizare (din rădăcina repo-ului):
    pip install -r requirements.txt
    python3 tests/run_tests.py            # toate testele
    python3 tests/run_tests.py auth_test  # doar unul

Fiecare test rulează într-un folder temporar cu: copia aplicației (app_under_test.py), componentele
session_bridge/ și key_bridge/, fake_sb.py și un .streamlit/secrets.toml de test (chei false).
"""
import os, shutil, subprocess, sys, tempfile, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
TESTS = pathlib.Path(__file__).resolve().parent
APP = os.environ.get("APP_FILE", "app.py")

BASE_SECRETS = '''SUPABASE_URL = "https://example.invalid"
SUPABASE_KEY = "sb_secret_fakefakefakefake"
APP_INSTANCE_ID = "test_local"
SMTP_USER = "profesor.test@gmail.com"
SMTP_PASSWORD = "abcd efgh ijkl mnop"
'''
OWNER_KEYS = 'GOOGLE_API_KEYS = ["AIzaFAKE1111111111111111", "AIzaFAKE2222222222222222"]\n'

# test -> (cu cheile „tale” în secrets?, argumente, secrets suplimentare)
SUITE = {
    "auth_test":   (True,  [], ""),
    "reset_test":  (True,  [], ""),
    "multi_test":  (True,  [], ""),
    "rot_test":    (True,  [], ""),
    "retry_test":  (True,  [], ""),
    "modes_test":  (True,  [], ""),
    "webhook_test": (False, [], ""),
    "comercial_test": (True, [], 'BILLING_MODE = "on"\nAPP_BASE_URL = "https://profesor.test"\n'
                                 'STRIPE_SECRET_KEY = "sk_test_x"\nCOMPANY_NAME = "Test SRL"\n'),
    "keys_test":   (False, [], ""),
    "photo_test":  (False, ["nokey"], ""),
    "photo_test:ocr": (False, ["ocr"], 'OCR_API_KEY = "AIzaOCRKEY000000000000000"\n'),
}


def run(name: str) -> tuple[int, int, str]:
    script, owner, args, extra = name.split(":")[0], *SUITE[name]
    with tempfile.TemporaryDirectory() as d:
        d = pathlib.Path(d)
        shutil.copy(ROOT / APP, d / "app_under_test.py")
        for comp in ("session_bridge", "key_bridge", "comercial", "legal", "billing_webhook"):
            if (ROOT / comp).exists():
                shutil.copytree(ROOT / comp, d / comp)
        shutil.copy(TESTS / "fake_sb.py", d / "fake_sb.py")
        shutil.copy(TESTS / f"{script}.py", d / f"{script}.py")
        (d / ".streamlit").mkdir()
        (d / ".streamlit" / "secrets.toml").write_text(BASE_SECRETS + (OWNER_KEYS if owner else "") + extra)
        p = subprocess.run([sys.executable, f"{script}.py", *args], cwd=d, capture_output=True, text=True, timeout=900)
        out = p.stdout + p.stderr
        lines = [l for l in out.splitlines() if l.startswith(("PASS", "FAIL"))]
        passed = sum(l.startswith("PASS") for l in lines)
        failed = sum(l.startswith("FAIL") for l in lines) + (1 if "Traceback" in out else 0)
        report = "\n".join("   " + l for l in lines)
        if "Traceback" in out:
            report += "\n   " + "\n   ".join(out[out.index("Traceback"):].splitlines()[-6:])
        return passed, failed, report


def main():
    wanted = [a for a in sys.argv[1:] if not a.startswith("-")]
    names = [n for n in SUITE if not wanted or n.split(":")[0] in wanted]
    tot_p = tot_f = 0
    for n in names:
        p, f, rep = run(n)
        tot_p += p; tot_f += f
        print(f"{'✅' if not f else '❌'} {n}: {p} trec, {f} pică")
        if f or "-v" in sys.argv:
            print(rep)
    print(f"\nTotal: {tot_p} trec, {tot_f} pică")
    sys.exit(1 if tot_f else 0)


if __name__ == "__main__":
    main()
