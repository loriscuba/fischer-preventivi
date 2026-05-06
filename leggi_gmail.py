"""
Legge email Gmail con allegati PDF (conferme ordine Fischer),
estrae dati e scrive su Supabase.
"""

import imaplib
import email
import base64
import io
import os
import re
import requests
import pdfplumber
from datetime import datetime
from email.header import decode_header

GMAIL_USER     = os.environ["GMAIL_USER"]
GMAIL_PASSWORD = os.environ["GMAIL_PASSWORD"]
SUPABASE_URL   = os.environ["SUPABASE_URL"]
SUPABASE_KEY   = os.environ["SUPABASE_KEY"]

HEADERS = {
    "apikey":        SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type":  "application/json",
    "Prefer":        "resolution=merge-duplicates",
}


# ── Gmail ──────────────────────────────────────────────────────────────────────
def connetti_gmail():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_PASSWORD)
    mail.select("inbox")
    return mail


def scarica_pdf_allegati(mail) -> list[tuple[str, bytes]]:
    """Ritorna lista di (nome_file, bytes) per ogni PDF non ancora letto."""
    _, msg_ids = mail.search(None, 'UNSEEN FROM "loris.cubaiu@fischer.it"')
    risultati = []

    for msg_id in msg_ids[0].split():
        _, msg_data = mail.fetch(msg_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        for part in msg.walk():
            if part.get_content_type() == "application/pdf":
                nome = part.get_filename() or "allegato.pdf"
                nome = decode_header(nome)[0][0]
                if isinstance(nome, bytes):
                    nome = nome.decode()
                contenuto = part.get_payload(decode=True)
                risultati.append((nome, contenuto))
                print(f"📎 Trovato allegato: {nome}")

        # Marca come letta
        mail.store(msg_id, "+FLAGS", "\\Seen")

    return risultati


# ── Estrazione testo PDF ───────────────────────────────────────────────────────
def estrai_testo(pdf_bytes: bytes) -> str:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)

# ── Parsing ordine ─────────────────────────────────────────────────────────────
def cerca(pattern, testo, default=""):
    m = re.search(pattern, testo)
    return m.group(1).strip() if m else default


def converti_data(s: str) -> str | None:
    try:
        return datetime.strptime(s, "%d.%m.%Y").strftime("%Y-%m-%d")
    except:
        return None


def converti_numero(s: str) -> float | None:
    try:
        return float(s.replace(".", "").replace(",", "."))
    except:
        return None


def parse_ordine(testo: str) -> dict:
    return {
        "numero_ordine":        cerca(r"Numero d'ordine\s+(\d+)", testo),
        "data":                 converti_data(cerca(r"Data\s+(\d{2}\.\d{2}\.\d{4})", testo)),
        "numero_cliente":       cerca(r"Numero di cliente\s+(\d+)", testo),
        "destinazione":         cerca(r"Destinazione merce\s*\n(.+)", testo),
        "magazzino":            cerca(r"Magazzino:\s*(.+)", testo),
        "peso_kg":              converti_numero(cerca(r"Peso totale:\s*([\d,\.]+)\s*KG", testo)),
        "condizioni_pagamento": cerca(r"Condizioni Pagamento\s+(.+)", testo),
        "importo_netto":        converti_numero(cerca(r"Importo totale EUR\s+([\d\.,]+)", testo)),
        "totale":               converti_numero(cerca(r"Totale ordine EUR\s+([\d\.,]+)", testo)),
        "iva":                  converti_numero(cerca(r"imponibile:\s+[\d\.,]+\s+([\d\.,]+)", testo)),
    }


def parse_righe(testo: str, numero_ordine: str) -> list[dict]:
    righe = []
    pattern = re.compile(
        r"(\d{8})\s+"
        r"(.+?)\s+"
        r"(\d+)\s+"
        r"(PZ|KG|MT|CF)\s+"
        r"([\d,\.]+)\s+"
        r"\d+\s+"
        r"(?:([\d,\.]+)(?:\/\s*([\d,\.]+))?\s+)?"
        r"([\d,\.]+)\s+"
        r"(\d+)",
        re.MULTILINE
    )
    # Cerca anche date consegna
    date_consegna = re.findall(r"Data presunta consegna il (\d{2}\.\d{2}\.\d{4})", testo)

    for i, m in enumerate(pattern.finditer(testo)):
        data_consegna = converti_data(date_consegna[i]) if i < len(date_consegna) else None
        righe.append({
            "numero_ordine":    numero_ordine,
            "cod_art":          m.group(1),
            "descrizione":      m.group(2).strip(),
            "quantita":         int(m.group(3)),
            "um":               m.group(4),
            "prezzo_unitario":  converti_numero(m.group(5)),
            "sconto_1":         converti_numero(m.group(6)) if m.group(6) else None,
            "sconto_2":         converti_numero(m.group(7)) if m.group(7) else None,
            "importo":          converti_numero(m.group(8)),
            "data_consegna":    data_consegna,
        })

    return righe


# ── Supabase ───────────────────────────────────────────────────────────────────
def upsert(tabella: str, record: dict) -> None:
    # Rimuovi campi None per evitare errori
    record = {k: v for k, v in record.items() if v is not None}
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/{tabella}",
        headers=HEADERS,
        json=record
    )
    if r.status_code not in (200, 201):
        print(f"❌ Errore {tabella}: {r.status_code} - {r.text}")
    else:
        print(f"✓ {tabella}: {record.get('numero_ordine') or record.get('cod_art')}")


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    print("=== Fischer Ordini → Supabase ===")

    mail = connetti_gmail()
    allegati = scarica_pdf_allegati(mail)
    mail.logout()

    if not allegati:
        print("📭 Nessuna email non letta con PDF allegato.")
        return

    for nome_file, pdf_bytes in allegati:
        print(f"\n📄 Elaborazione: {nome_file}")

        testo = estrai_testo(pdf_bytes)
        print("=== TESTO ESTRATTO ===")
        print(testo[:3000])
        print("=== FINE TESTO ===")
        ordine = parse_ordine(testo)
        numero_ordine = ordine.get("numero_ordine")

        if not numero_ordine:
            print(f"⚠️ Numero ordine non trovato in {nome_file}, skip.")
            continue

        print(f"  Ordine: {numero_ordine} | Cliente: {ordine.get('numero_cliente')} | Totale: {ordine.get('totale')}")

        upsert("ordini", ordine)

        righe = parse_righe(testo, numero_ordine)
        print(f"  → {len(righe)} righe trovate")
        for riga in righe:
            upsert("righe_ordine", riga)

    print("\n✅ Sync completata.")


if __name__ == "__main__":
    main()
