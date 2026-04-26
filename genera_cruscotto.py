#!/usr/bin/env python3
"""
genera_cruscotto.py
===================
Genera il file cruscotto_data.json partendo da:
  1. ordini_mancanti_su_rolling_consuntivo*.xlsx   (fatturati mensili)
  2. *_report_penetrazione_gamma_prodotti_*.xlsx   (gamma prodotti)

Uso:
    python genera_cruscotto.py
oppure con percorsi espliciti:
    python genera_cruscotto.py --rolling mio_rolling.xlsx --gamma mia_penetrazione.xlsx

Output: cruscotto_data.json  (nella stessa cartella dello script)
"""

import json
import re
import sys
import glob
import argparse
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("❌  pandas non trovato. Esegui: pip install pandas openpyxl")
    sys.exit(1)

# ── CONFIGURAZIONE ────────────────────────────────────────────────────────────

AGENTE_NOME  = "Cubaiu Loris"
AGENTE_COD   = "400542"          # come stringa o int, lo script gestisce entrambi

MESI_IT = [
    "gennaio","febbraio","marzo","aprile","maggio","giugno",
    "luglio","agosto","settembre","ottobre","novembre","dicembre"
]

GAMMA_SHEETS = [
    "Ferr bullonerie-viterie",
    "Ferr Serr Legno",
    "Ferr Serr Alluminio",
    "Ferr Generica",
    "Ferr Carp-Fabbro",
    "ITS",
    "Elettrico",
    "Edile",
]

# Etichette brevi per i mesi nei JSON (devono combaciare con quelle nell'HTML)
def mese_key_2024(idx_0):
    nomi = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"]
    return nomi[idx_0] + "24"

def mese_key_2025(idx_0):
    nomi = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"]
    return nomi[idx_0] + "25"


# ── RICERCA FILE ──────────────────────────────────────────────────────────────

def trova_file(pattern_glob, nome_amichevole):
    """Cerca un file con glob; se ce n'è uno solo lo restituisce, altrimenti chiede."""
    matches = glob.glob(pattern_glob)
    matches = [m for m in matches if not Path(m).name.startswith("~$")]  # skip temp
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        print(f"\n⚠️  Trovati più file per '{nome_amichevole}':")
        for i, m in enumerate(matches):
            print(f"  [{i}] {m}")
        scelta = input("Quale vuoi usare? [numero]: ").strip()
        return matches[int(scelta)]
    print(f"\n❌  Nessun file trovato per '{nome_amichevole}' (pattern: {pattern_glob})")
    percorso = input("Inserisci il percorso manualmente: ").strip()
    return percorso


# ── LETTURA ROLLING ───────────────────────────────────────────────────────────

def leggi_rolling(filepath):
    """
    Legge il file rolling e restituisce un dict:
      { cod_cliente: { dati... } }
    """
    print(f"\n📊  Lettura rolling: {filepath}")
    df = pd.read_excel(filepath, sheet_name="REPORT", header=0)

    # Normalizza codice agente (può essere int o stringa)
    df["_cod_agente_str"] = df["Cod. Agente"].astype(str).str.strip().str.split(".").str[0]
    agente_df = df[df["_cod_agente_str"] == str(AGENTE_COD)].copy()

    if agente_df.empty:
        print(f"⚠️  Nessuna riga trovata per agente {AGENTE_COD}. Codici presenti:")
        print(df["_cod_agente_str"].unique()[:10])
        sys.exit(1)

    print(f"  → {len(agente_df)} righe per {AGENTE_NOME}")

    # Rileva le colonne mesi 2024 (gen-dic)
    cols_m24 = {}   # {indice_0: nome_colonna}
    cols_m25 = {}   # {indice_0: nome_colonna}
    for col in df.columns:
        s = str(col).lower()
        for i, m in enumerate(MESI_IT):
            if m in s and "2024" in s and "fatturato" in s and "tot" not in s and "progr" not in s:
                cols_m24[i] = col
            if m in s and "2025" in s and "fatturato" in s and "tot" not in s and "progr" not in s:
                cols_m25[i] = col

    # Cerca colonna mese parziale corrente (es. "Spedito + Ordinato nel mese 010.2025")
    col_mese_corrente = None
    mese_corrente_idx = None   # indice 0-based
    for col in df.columns:
        # Pattern: "Spedito + Ordinato nel mese NNN.2025"
        m = re.search(r"ordinato nel mese\s+0*(\d+)\.2025", str(col).lower())
        if m:
            mese_corrente_idx = int(m.group(1)) - 1   # 0-based
            col_mese_corrente = col
            break

    # In alternativa cerca "Consegnato" + "preparazione" + "da spedire"
    col_consegnato = None
    col_prep       = None
    col_spedire    = None
    for col in df.columns:
        s = str(col).lower()
        if "consegnato" in s and "2025" in s:
            col_consegnato = col
        if "preparazione" in s and "2025" in s:
            col_prep = col
        if "da spedire" in s and "2025" in s:
            col_spedire = col

    # Mese parziale corrente (da usare se col_mese_corrente non trovata)
    if col_mese_corrente is None and col_consegnato:
        # Prova a estrarre il mese dalla colonna consegnato
        m2 = re.search(r"(\d{2})/(\d{2})/2025", str(col_consegnato))
        if m2:
            mese_corrente_idx = int(m2.group(2)) - 1
        col_mese_corrente = "__calcolata__"

    print(f"  → Mesi 2024 trovati: {sorted(cols_m24.keys())} ({len(cols_m24)} mesi)")
    print(f"  → Mesi 2025 trovati: {sorted(cols_m25.keys())} ({len(cols_m25)} mesi)")
    if mese_corrente_idx is not None:
        print(f"  → Mese parziale corrente: {MESI_IT[mese_corrente_idx]} 2025 ({col_mese_corrente})")

    # Aggiungi mese corrente alla raccolta 2025 se non già presente
    if mese_corrente_idx is not None and mese_corrente_idx not in cols_m25:
        cols_m25[mese_corrente_idx] = col_mese_corrente

    # Data aggiornamento dal nome del file
    fname = Path(filepath).stem
    data_match = re.search(r"(\d{1,2})[-_](\d{1,2})[-_](\d{4})", fname)
    if data_match:
        g, m_num, a = data_match.groups()
        aggiornato = f"{int(g):02d}/{int(m_num):02d}/{a}"
    else:
        from datetime import date
        aggiornato = date.today().strftime("%d/%m/%Y")

    # Costruisci dict clienti
    clienti_rolling = {}
    for _, row in agente_df.iterrows():
        cod = str(row.get("Cod. Cliente", "")).strip()
        if not cod:
            continue

        # Gamma da "Unnamed: 2"
        gamma = str(row.get("Unnamed: 2", "")).strip()
        if gamma in ("nan", ""):
            gamma = str(row.get("GAV", "")).strip()

        def v(col):
            if col is None or col == "__calcolata__":
                return 0.0
            val = row.get(col, None)
            if pd.isna(val) or val is None:
                return 0.0
            return round(float(val), 2)

        # Mesi 2024
        mesi2024 = {}
        for idx in range(12):
            key = mese_key_2024(idx)
            if idx in cols_m24:
                mesi2024[key] = v(cols_m24[idx])
            else:
                mesi2024[key] = 0.0

        # Mesi 2025
        mesi2025 = {}
        max_mese_2025 = max(cols_m25.keys()) if cols_m25 else -1
        for idx in range(max_mese_2025 + 1):
            key = mese_key_2025(idx)
            if idx in cols_m25:
                if col_mese_corrente == "__calcolata__" and idx == mese_corrente_idx:
                    # Calcola dalla somma delle 3 colonne
                    val = (v(col_consegnato) + v(col_prep) + v(col_spedire))
                else:
                    mesi2025[key] = v(cols_m25[idx])
            else:
                mesi2025[key] = 0.0
            if idx == mese_corrente_idx and col_mese_corrente == "__calcolata__":
                mesi2025[key] = round(v(col_consegnato) + v(col_prep) + v(col_spedire), 2)

        fatt2023  = v("Fatturato TOTALE 2023")
        fatt2024  = v("Fatturato TOTALE 2024")
        prog2025  = round(sum(mesi2025.values()), 2)
        var_anno  = round(prog2025 - fatt2024, 2)

        clienti_rolling[cod] = {
            "cod":       cod,
            "cod_short": cod.split("/")[-1],
            "nome":      str(row.get("Rag. Sociale", "")).strip(),
            "gamma":     gamma,
            "fatt2023":  fatt2023,
            "fatt2024":  fatt2024,
            "mesi2024":  mesi2024,
            "mesi2025":  mesi2025,
            "prog2024":  fatt2024,
            "prog2025":  prog2025,
            "var_anno":  var_anno,
            "gamma_dettaglio": {},
        }

    return clienti_rolling, aggiornato, max_mese_2025


# ── LETTURA GAMMA ─────────────────────────────────────────────────────────────

def leggi_gamma(filepath, clienti_dict):
    """
    Legge tutti i sheet del file penetrazione gamma e aggiunge
    gamma_dettaglio a ogni cliente.
    """
    print(f"\n🗂️  Lettura gamma: {filepath}")
    xl = pd.ExcelFile(filepath)
    sheets_presenti = xl.sheet_names

    for sheet in GAMMA_SHEETS:
        if sheet not in sheets_presenti:
            print(f"  ⚠️  Sheet '{sheet}' non trovato, salto.")
            continue

        df = pd.read_excel(filepath, sheet_name=sheet, header=None)

        # Riga 7 (indice 7) = intestazioni dati
        header_row = 7
        headers = df.iloc[header_row].tolist()

        # Trova colonne chiave per posizione
        # Struttura: MACROAREA(0) AGENTE(1) nome(2) %IMMAN(3) %STRAT(4) CodCliente(5) RagSoc(6) Online?(7) Fatt2024(8) prodotti(9+)
        col_agente   = 1
        col_perc_imm = 3
        col_perc_str = 4
        col_cod      = 5
        col_fatt2024 = 8
        col_prod_start = 9

        # Riga flag (riga 4): Strategica-Immancabile / Strategica
        flag_row = df.iloc[4].tolist()
        # Nomi prodotti dalla riga 7
        prod_names = headers[col_prod_start:]

        # Flags per ogni prodotto
        flags = []
        for fi in range(col_prod_start, len(headers)):
            f = str(flag_row[fi]) if fi < len(flag_row) else ""
            if f in ("Strategica-Immancabile", "Strategica"):
                flags.append(f)
            else:
                flags.append("")

        # Filtra righe Cubaiu (dalla riga header+1 in poi)
        data_rows = df.iloc[header_row + 1:].copy()
        mask = data_rows.iloc[:, col_agente].astype(str).str.strip().str.split(".").str[0] == str(AGENTE_COD)
        cubaiu_rows = data_rows[mask]

        print(f"  → Sheet '{sheet}': {len(cubaiu_rows)} clienti Cubaiu")

        for _, row in cubaiu_rows.iterrows():
            cod_cliente = str(row.iloc[col_cod]).strip().split(".")[0]
            # Cerca il cliente nel dict (il cod potrebbe essere solo il numero finale)
            matched_cod = None
            for k in clienti_dict:
                if k.endswith("/" + cod_cliente) or k == cod_cliente:
                    matched_cod = k
                    break

            if matched_cod is None:
                # Cliente nella gamma ma non nel rolling (inattivo)
                continue

            perc_imm = float(row.iloc[col_perc_imm]) if pd.notna(row.iloc[col_perc_imm]) else 0.0
            perc_str = float(row.iloc[col_perc_str]) if pd.notna(row.iloc[col_perc_str]) else 0.0
            fatt2024_gamma = float(row.iloc[col_fatt2024]) if pd.notna(row.iloc[col_fatt2024]) else 0.0

            # Prodotti
            prodotti = []
            for pi, pname in enumerate(prod_names):
                if not isinstance(pname, str) or not pname.strip():
                    continue
                col_idx = col_prod_start + pi
                val = row.iloc[col_idx] if col_idx < len(row) else None
                valore = round(float(val), 2) if (val is not None and pd.notna(val) and str(val).strip() not in ("", "nan")) else 0.0
                prodotti.append({
                    "nome":   pname.strip(),
                    "flag":   flags[pi] if pi < len(flags) else "",
                    "valore": valore,
                })

            clienti_dict[matched_cod]["gamma_dettaglio"][sheet] = {
                "perc_immancabili": round(perc_imm, 6),
                "perc_strategiche": round(perc_str, 6),
                "fatt2024":         round(fatt2024_gamma, 2),
                "prodotti":         prodotti,
            }

    return clienti_dict


# ── COSTRUZIONE JSON FINALE ───────────────────────────────────────────────────

def costruisci_json(clienti_dict, aggiornato, max_mese_2025_idx):
    # Ordina clienti per fatturato 2024 desc
    clienti_list = sorted(clienti_dict.values(), key=lambda c: c["fatt2024"], reverse=True)

    # Label mesi
    mesi_labels_2024 = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"]
    nomi_brevi = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"]
    mesi_labels_2025 = nomi_brevi[:max_mese_2025_idx + 1]

    data = {
        "aggiornato":        aggiornato,
        "agente":            AGENTE_NOME,
        "cod_agente":        AGENTE_COD,
        "clienti":           clienti_list,
        "mesi_labels_2024":  mesi_labels_2024,
        "mesi_labels_2025":  mesi_labels_2025,
        "settori_gamma":     GAMMA_SHEETS,
    }
    return data


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Genera cruscotto_data.json")
    parser.add_argument("--rolling", default=None, help="Percorso file rolling .xlsx")
    parser.add_argument("--gamma",   default=None, help="Percorso file penetrazione gamma .xlsx")
    parser.add_argument("--output",  default="cruscotto_data.json", help="File di output")
    args = parser.parse_args()

    # Trova i file se non specificati
    rolling_path = args.rolling or trova_file(
        "ordini_mancanti_su_rolling_consuntivo*.xlsx",
        "fatturato rolling"
    )
    gamma_path = args.gamma or trova_file(
        "*report_penetrazione_gamma_prodotti*.xlsx",
        "penetrazione gamma"
    )

    # Lettura
    clienti_dict, aggiornato, max_mese_2025_idx = leggi_rolling(rolling_path)
    clienti_dict = leggi_gamma(gamma_path, clienti_dict)

    # Costruzione output
    data = costruisci_json(clienti_dict, aggiornato, max_mese_2025_idx)

    # Salvataggio
    output_path = args.output
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n✅  Completato!")
    print(f"   Clienti elaborati : {len(data['clienti'])}")
    print(f"   Mesi 2025         : {len(data['mesi_labels_2025'])} ({', '.join(data['mesi_labels_2025'])})")
    print(f"   Aggiornato        : {aggiornato}")
    print(f"   Output            : {output_path}")
    print(f"\n👉  Copia {output_path} nella cartella del cruscotto e ricarica la pagina.")


if __name__ == "__main__":
    main()
