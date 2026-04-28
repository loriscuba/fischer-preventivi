# Fischer Italia · Cruscotto Cubaiu

PWA per la gestione del fatturato area Cubaiu — Agente 400542.

## Struttura repo

```
├── index.html                  ← PWA principale (Preventivi, Catalogo, Griglie, Budget)
├── cruscotto_fatturato.html    ← Cruscotto fatturato (standalone)
├── cruscotto_data.json         ← Dati elaborati (generato automaticamente da Actions)
├── scripts/
│   └── genera_cruscotto.py     ← Script di elaborazione Excel → JSON
├── data/                       ← File Excel caricati dal cruscotto (branch: data)
│   ├── rolling.xlsx
│   └── gamma.xlsx
└── .github/workflows/
    └── aggiorna_cruscotto.yml  ← Workflow GitHub Actions
```

## Come aggiornare i dati

1. Apri il cruscotto su GitHub Pages
2. Vai alla sezione **Aggiorna Dati**
3. Inserisci `owner/repo` e il tuo Personal Access Token (scope: `repo`)
4. Carica i due file Excel aggiornati
5. Clicca **Pubblica su GitHub**
6. GitHub Actions elabora i file e aggiorna `cruscotto_data.json` in ~1-2 minuti

## Setup iniziale

### 1. Branch `data`
Crea la branch `data` (vuota) dove verranno depositati i file Excel:
```bash
git checkout --orphan data
git rm -rf .
git commit --allow-empty -m "init data branch"
git push origin data
git checkout main
```

### 2. Cartella `data/`
Aggiungi la cartella al repo su branch `data`:
```bash
git checkout data
mkdir data && touch data/.gitkeep
git add data/.gitkeep
git commit -m "init data folder"
git push origin data
git checkout main
```

### 3. Permessi GitHub Actions
Vai su **Settings → Actions → General → Workflow permissions**
e seleziona **Read and write permissions**.

### 4. Personal Access Token
Vai su **Settings → Developer Settings → Personal Access Tokens → Classic**
- Scope: `repo` (full)
- Nessuna scadenza (o lunga)
- Inseriscilo nel cruscotto alla sezione "Aggiorna Dati" — viene salvato nel localStorage del browser, mai nel codice

### 5. GitHub Pages
Vai su **Settings → Pages → Source: Deploy from branch → main / root**

## Dipendenze Python (usate da Actions)

```
pandas
openpyxl
```
