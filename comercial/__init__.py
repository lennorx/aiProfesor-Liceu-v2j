"""Partea comercială a Profesor Liceu: conturi obligatorii, acord parental, documente legale,
abonamente Stripe, raportări, export/ștergere date. Activă doar cu BILLING_MODE = "on"."""
from . import billing, compliance, core, ui  # noqa: F401
