
#!/usr/bin/env python3
"""
genera_cruscotto.py  –  Fischer Italia · Cubaiu Loris (Agente 400542)
Eseguito da GitHub Actions: legge data/rolling.xlsx e data/gamma.xlsx,
scrive cruscotto_data.json nella root del repo.
"""

import json, os, re
import pandas as pd

COD_AGENTE = "400542"

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(BASE, "..")

PATH_ROLLING = os.path.join(ROOT, "data", "rolling.xlsx")
PATH_GAMMA   = os.path.join(ROOT, "data", "gamma.xlsx")
PATH_OUT     = os.path.join(ROOT, "cruscotto_data.json")

# ── utils ──────────────────────────────────────────────────────────────
def clean(v):
    try:
        x = float(v)
        return round(x, 2) if not pd.isna(x) else 0.0
    except Exception:
        return 0.0

MESI24 = ["Gen24","Feb24","Mar24","Apr24","Mag24","Giu24",
          "Lug24","Ago24","Set24","Ott24","Nov24","Dic24"]
MESI25 = ["Gen25","Feb25","Mar25","Apr25","Mag25","Giu25",
          "Lug25","Ago25","Set25","Ott25","Nov25"]

# ── 1. Rolling ─────────────────────────────────────────────────────────
print("📂  Lettura rolling consuntivo...")
df = pd.read_excel(PATH_ROLLING, sheet_name="REPORT", header=2)

rename = {
    df.columns[0]:"DIV",   df.columns[1]:"GAV",    df.columns[2]:"Gamma",
    df.columns[3]:"Cod_Agente", df.columns[4]:"Agente",
    df.columns[5]:"Cod_Cliente", df.columns[6]:"RagSoc", df.columns[7]:"Nodo",
    df.columns[8]:"Fatt2023",   df.columns[9]:"Fatt2024",
    df.columns[10]:"Gen24", df.columns[11]:"Feb24", df.columns[12]:"Mar24",
    df.columns[13]:"Apr24", df.columns[14]:"Mag24", df.columns[15]:"Giu24",
    df.columns[16]:"Lug24", df.columns[17]:"Ago24", df.columns[18]:"Set24",
    df.columns[19]:"Ott24", df.columns[20]:"Nov24", df.columns[21]:"Dic24",
    df.columns[22]:"Prog2025",
    df.columns[23]:"Gen25", df.columns[24]:"Feb25", df.columns[25]:"Mar25",
    df.columns[26]:"Apr25", df.columns[27]:"Mag25", df.columns[28]:"Giu25",
    df.columns[29]:"Lug25", df.columns[30]:"Ago25", df.columns[31]:"Set25",
    df.columns[32]:"Ott25", df.columns[33]:"Nov25",
    df.columns[34]:"Dic25_consegnato",
    df.columns[41]:"Prog2024_totale",
    df.columns[42]:"Prog2025_totale",
    df.columns[43]:"Var_Anno",
}
df = df.rename(columns=rename)
df = df[df["Cod_Agente"].astype(str).str.strip() == COD_AGENTE].copy()
df["cod_short"] = df["Cod_Cliente"].astype(str).str.split("/").str[-1].str.strip()

from datetime import date
data_aggiornamento = date.today().strftime("%d/%m/%Y")
print(f"   Clienti: {len(df)}")

# ── 2. Gamma ───────────────────────────────────────────────────────────
print("📂  Lettura gamma prodotti...")
xf2 = pd.ExcelFile(PATH_GAMMA)
sheets_gamma = xf2.sheet_names[1:]   # salta Riassunto

gamma_all = {}
for sheet in sheets_gamma:
    dg = pd.read_excel(xf2, sheet_name=sheet, header=None)
    header_row = dg.iloc[7, :]
    flag_row   = dg.iloc[4, :]
    prodotti_map = []
    for i in range(9, len(header_row)):
        nome = str(header_row[i]).strip() if pd.notna(header_row[i]) else ""
        if not nome or nome == "nan": continue
        flag = str(flag_row[i]).strip() if pd.notna(flag_row[i]) else ""
        prodotti_map.append({"col": i, "nome": nome, "flag": flag})
    mask = dg.iloc[:, 1].astype(str).str.contains(COD_AGENTE, na=False)
    for _, rg in dg[mask].iterrows():
        cod = str(rg[5]).strip().split("/")[-1]
        if cod not in gamma_all: gamma_all[cod] = {}
        gamma_all[cod][sheet] = {
            "perc_immancabili": clean(rg[3]),
            "perc_strategiche": clean(rg[4]),
            "fatt2024":         clean(rg[8]),
            "prodotti": [{"nome": p["nome"], "flag": p["flag"], "valore": clean(rg[p["col"]])} for p in prodotti_map],
        }

con_gamma = sum(1 for c in df["cod_short"] if c in gamma_all)
print(f"   Clienti con gamma: {con_gamma}/{len(df)}")

# ── 3. Build ───────────────────────────────────────────────────────────
print("🔧  Costruzione JSON...")
clienti = []
for _, r in df.iterrows():
    cs = r["cod_short"]
    clienti.append({
        "cod": str(r["Cod_Cliente"]).strip(), "cod_short": cs,
        "nome": str(r["RagSoc"]).strip(), "gamma": str(r["Gamma"]).strip(),
        "fatt2023": clean(r["Fatt2023"]), "fatt2024": clean(r["Fatt2024"]),
        "mesi2024": {m: clean(r.get(m, 0)) for m in MESI24},
        "mesi2025": {m: clean(r.get(m, 0)) for m in MESI25},
        "prog2024": clean(r.get("Prog2024_totale", 0)),
        "prog2025": clean(r.get("Prog2025_totale", 0)),
        "var_anno": clean(r.get("Var_Anno", 0)),
        "gamma_dettaglio": gamma_all.get(cs, {}),
    })

output = {
    "aggiornato": data_aggiornamento,
    "agente": "Cubaiu Loris", "cod_agente": COD_AGENTE,
    "clienti": clienti,
    "mesi_labels_2024": ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"],
    "mesi_labels_2025": ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov"],
    "settori_gamma": sheets_gamma,
}

with open(PATH_OUT, "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

kb = os.path.getsize(PATH_OUT) / 1024
print(f"✅  cruscotto_data.json  ({kb:.1f} KB) — clienti: {len(clienti)}, gamma: {con_gamma}")
