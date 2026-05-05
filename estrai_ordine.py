"""
Riceve una Conferma d'Ordine Fischer in base64 da GitHub Actions,
estrae intestazione e righe, scrive su Supabase.
"""

import os
import sys
import base64
import json
import re
import requests
import pdfplumber
import io
from datetime import datetime

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_KEY = os.environ["SUPABASE_KEY"]

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "resolution=merge-duplicates",
}


def extract_text(pdf_base64: str) -> str:
    pdf_bytes = base64.b64decode(pdf_base64)
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def parse_ordine(testo: str) -> dict:
    def cerca(pattern, testo, group=1, default=""):
        m = re.search(pattern, testo)
        return m.group(group).strip() if m else default

    return {
        "numero_ordine":         cerca(r"Numero d'ordine\s+(\d+)", testo),
        "data":                  cerca(r"Data\s+(\d{2}\.\d{2}\.\d{4})", testo),
        "numero_cliente":        cerca(r"Numero di cliente\s+(\d+)", testo),
        "destinazione":          cerca(r"Destinazione merce\s*\n(.+)", testo),
        "magazzino":             cerca(r"Magazzino:\s*(.+)", testo),
        "peso_kg":               cerca(r"Peso totale:\s*([\d,\.]+)\s*KG", testo),
        "condizioni_pagamento":  cerca(r"Condizioni Pagamento\s+(.+)", testo),
        "importo_netto":         cerca(r"Importo totale EUR\s+([\d\.,]+)", testo),
        "totale":                cerca(r"Totale ordine EUR\s+([\d\.,]+)", testo),
        "iva":                   cerca(r"IVA Vend\.\s*[\d,]+%\s+[\d,]+\s*%\s+imponibile:\s+[\d\.,]+\s+([\d\.,]+)", testo),
    }


def parse_righe(testo: str, numero_ordine: str) -> list[dict]:
    righe = []
    # Pattern: COD_ART DESCRIZIONE QTA UM PREZZO UM [SCONTI] IMPORTO IVA
    pattern = re.compile(
        r"(\d{8})\s+"           # cod_art (8 cifre)
        r"(.+?)\s+"             # descrizione
        r"(\d+)\s+"             # quantita
        r"(PZ|KG|MT|CF)\s+"    # um
        r"([\d,\.]+)\s+"       # prezzo_unitario
        r"\d+\s+"               # um2
        r"(?:([\d,\.]+)(?:\/\s*([\d,\.]+))?\s+)?"  # sconti opzionali
        r"([\d,\.]+)\s+"       # importo
        r"(\d+)",               # iva
        re.MULTILINE
    )

    for m in pattern.finditer(testo):
        righe.append({
            "numero_ordine":   numero_ordine,
            "cod_art":         m.group(1),
            "descrizione":     m.group(2).strip(),
            "quantita":        int(m.group(3)),
            "um":              m.group(4),
            "prezzo_unitario": float(m.group(5).replace(",", ".")),
            "sconto_1":        float(m.group(6).replace(",", ".")) if m.group(6) else None,
            "sconto_2":        float(m.group(7).replace(",", ".")) if m.group(7) else None,
            "importo":         float(m.group(8).replace(",", ".")),
        })

    return righe


def converti_data(data_str: str) -> str:
    """Converte da DD.MM.YYYY a YYYY-MM-DD per Supabase."""
    try:
        return datetime.strptime(data_str, "%d.%m.%Y").strftime("%Y-%m-%d")
    except:
        return None


def converti_numero(s: str) -> float | None:
    try:
        return float(s.replace(".", "").replace(",", "."))
    except:
        return None


def upsert(tabella: str, record: dict) -> None:
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/{tabella}",
        headers=HEADERS,
        json=record
    )
    if r.status_code not in (200, 201):
        print(f"❌ Errore {tabella}: {r.status_code} - {r.text}")
    else:
        print(f"✓ {tabella}: {record.get('numero_ordine') or record.get('cod_art')}")


def main():
    payload = json.loads(os.environ["EVENT_PAYLOAD"])
    pdf_base64 = payload.get("contenuto", "")
    nome_file  = payload.get("nome_file", "")

    print(f"📄 Elaborazione: {nome_file}")

    testo = extract_text(pdf_base64)
    ordine = parse_ordine(testo)
    numero_ordine = ordine["numero_ordine"]

    if not numero_ordine:
        print("❌ Numero ordine non trovato, skip.")
        sys.exit(1)

    # Normalizza campi
    ordine["data"]          = converti_data(ordine["data"])
    ordine["peso_kg"]       = converti_numero(ordine["peso_kg"])
    ordine["importo_netto"] = converti_numero(ordine["importo_netto"])
    ordine["totale"]        = converti_numero(ordine["totale"])
    ordine["iva"]           = converti_numero(ordine["iva"])

    # Scrivi intestazione
    upsert("ordini", ordine)

    # Scrivi righe
    righe = parse_righe(testo, numero_ordine)
    print(f"  → {len(righe)} righe trovate")
    for riga in righe:
        upsert("righe_ordine", riga)

    print("✅ Sync completata.")


if __name__ == "__main__":
    main()
