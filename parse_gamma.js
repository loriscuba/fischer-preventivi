/**
 * parse_gamma.js
 * Legge il file *_report_penetrazione_gamma_prodotti_*.xlsx
 * e produce un oggetto { byCliente } compatibile con caricaXlsxDaGitHub()
 *
 * Struttura foglio (0-based):
 *  Riga 0  – titolo
 *  Riga 1  – settore
 *  Riga 2  – periodo
 *  Riga 3  – vuota
 *  Riga 4  – flag prodotti: "Strategica-Immancabile" / "Strategica" / vuoto
 *  Riga 5  – nomi prodotti (col 9+)
 *  Riga 6  – codici prodotti (col 9+)
 *  Riga 7  – intestazioni: MACROAREA(0) AGENTE(1) nome(2) %IMM(3) %STR(4) Cliente(5) RagSoc(6) Online?(7) Fatt2024(8) prodotti(9+)
 *  Riga 8+ – dati clienti
 *
 *  Ultima colonna (34): codice cliente ripetuto come numero — usato come fallback
 */

(function (global) {
  'use strict';

  const COL_AGENTE    = 1;
  const COL_NOME_AG   = 2;
  const COL_PERC_IMM  = 3;
  const COL_PERC_STR  = 4;
  const COL_COD_CLI   = 5;
  const COL_RAG_SOC   = 6;
  const COL_FATT2024  = 8;
  const COL_PROD_START = 9;
  const ROW_FLAGS     = 4;
  const ROW_NAMES     = 5;
  const ROW_CODES     = 6;
  const ROW_HEADER    = 7;
  const ROW_DATA      = 8;

  const COD_AGENTE = '400542';

  const SHEETS = [
    'Ferr bullonerie-viterie',
    'Ferr Serr Legno',
    'Ferr Serr Alluminio',
    'Ferr Generica',
    'Ferr Carp-Fabbro',
    'ITS',
    'Elettrico',
    'Edile',
  ];

  function toNum(v) {
    if (v == null || v === '' || v === '#N/A') return null;
    const n = typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.'));
    return isNaN(n) ? null : n;
  }

  function normCod(v) {
    if (v == null) return '';
    return String(v).trim().split('.')[0];
  }

  /**
   * @param {ArrayBuffer} buffer  – contenuto del file .xlsx
   * @param {object}      XLSX    – libreria SheetJS
   * @param {Set}         codiciFiltro – Set di codici cliente da includere (dal rolling); se null include tutti
   * @returns {{ byCliente: object, settori: string[] }}
   *   byCliente[cod] = { nome, settori: { [sheet]: { perc_immancabili, perc_strategiche, fatt2024, prodotti[] } } }
   */
  function parseGamma(buffer, XLSX, codiciFiltro) {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: false });

    const byCliente = {};   // cod → { nome, settori: {} }

    for (const sheetName of SHEETS) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;

      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
      if (aoa.length <= ROW_DATA) continue;

      // Leggi prodotti dalla riga 5 (nomi) e riga 4 (flag)
      const headerRow = aoa[ROW_HEADER] || [];
      const namesRow  = aoa[ROW_NAMES]  || [];
      const flagsRow  = aoa[ROW_FLAGS]  || [];

      // Costruisci lista prodotti (col 9 in poi, fino all'ultima col con nome)
      const prodotti_def = [];
      for (let c = COL_PROD_START; c < headerRow.length; c++) {
        // Il nome del prodotto è in riga 5 oppure in riga 7 (header) come fallback
        const nome = namesRow[c] != null
          ? String(namesRow[c]).trim()
          : (headerRow[c] != null ? String(headerRow[c]).trim() : '');
        if (!nome || nome === 'null') continue;

        const flag = flagsRow[c] != null ? String(flagsRow[c]).trim() : '';
        prodotti_def.push({ col: c, nome, flag });
      }

      // Righe dati
      for (let r = ROW_DATA; r < aoa.length; r++) {
        const row = aoa[r];
        if (!row) continue;

        // Filtra per agente
        const codAg = normCod(row[COL_AGENTE]);
        if (codAg !== COD_AGENTE) continue;

        // Codice cliente: colonna 5 (stringa) oppure ultima col (numero)
        let cod = normCod(row[COL_COD_CLI]);
        if (!cod || cod === 'null') {
          // fallback ultima colonna
          cod = normCod(row[row.length - 1]);
        }
        if (!cod || cod === 'null') continue;

        // Se abbiamo un filtro e il cliente non è nel rolling, salta
        // (cerca sia cod esatto sia suffisso "/cod")
        if (codiciFiltro && codiciFiltro.size > 0) {
          let found = codiciFiltro.has(cod);
          if (!found) {
            for (const k of codiciFiltro) {
              if (k.endsWith('/' + cod) || k === cod) { found = true; break; }
            }
          }
          if (!found) continue;
        }

        const nome     = String(row[COL_RAG_SOC] || '').trim();
        const percImm  = toNum(row[COL_PERC_IMM]);
        const percStr  = toNum(row[COL_PERC_STR]);
        const fatt2024 = toNum(row[COL_FATT2024]) || 0;

        // Prodotti con valore
        const prodotti = prodotti_def.map(p => ({
          nome:   p.nome,
          flag:   p.flag,
          valore: toNum(row[p.col]) || 0,
        }));

        // Aggrega per cliente (stesso cliente può avere più righe per gamma diversa)
        if (!byCliente[cod]) {
          byCliente[cod] = { cod, nome, settori: {} };
        }

        if (!byCliente[cod].settori[sheetName]) {
          byCliente[cod].settori[sheetName] = {
            perc_immancabili: percImm,
            perc_strategiche: percStr,
            fatt2024,
            prodotti,
          };
        } else {
          // Aggrega se il cliente appare due volte nello stesso sheet
          const ex = byCliente[cod].settori[sheetName];
          ex.fatt2024 += fatt2024;
          prodotti.forEach((p, i) => {
            if (ex.prodotti[i]) ex.prodotti[i].valore += p.valore;
          });
        }
      }
    }

    return { byCliente, settori: SHEETS };
  }

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = parseGamma;
  } else {
    global.parseGamma = parseGamma;
  }

})(typeof globalThis !== 'undefined' ? globalThis : window);
