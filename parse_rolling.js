/**
 * parse_rolling.js
 * Legge il file ordini_mancanti_su_rolling_consuntivo*.xlsx
 * e produce un oggetto compatibile con la struttura cruscotto_data.json
 *
 * Struttura attesa del file XLSX:
 *  Riga 1  – titolo "REPORT ORDINI MANCANTI SU CONSUNTIVO"
 *  Riga 2  – AGGIORNATO: / data
 *  Riga 3  – etichette colonne (header descrittivi, usati per mappare)
 *  Riga 4  – totali SUBTOTAL (riga da saltare)
 *  Righe 5+ – una riga per cliente
 *
 * Mappatura colonne (0-based):
 *  0  DIV.
 *  1  GAV
 *  2  Gamma (DIY / Edile / Elettrico / ITS / ecc.)
 *  3  Cod. Agente
 *  4  Agente
 *  5  Cod. Cliente
 *  6  Rag. Sociale
 *  7  Fatturato TOTALE 2024
 *  8  Fatturato TOTALE 2025
 *  9  Gen 2025 … col 20 Dic 2025
 *  21 Fatturato Progressivo Gen-Mar 2026
 *  22 GTO Gen'26 … col 33 GTO Dic'26
 *  34 MESE: Consegnato
 *  35 MESE: In preparazione
 *  36 Da spedire nel mese
 *  37 Ordinato oltre mese
 *  38 Fatturato mese anno prev.
 *  39 Spedito + Ordinato mese corrente
 *  40 Variazione mese corrente vs anno prev.
 *  41 Fatturato progressivo anno prev. (gen-mese)
 *  42 Fatturato progressivo anno corr. + ordinato
 *  43 Variazione progressivo
 */

(function (global) {
  'use strict';

  /**
   * Converte un valore di cella in numero (null se non valido)
   */
  function toNum(v) {
    if (v == null || v === '' || v === '#N/A') return null;
    const n = typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.'));
    return isNaN(n) ? null : n;
  }

  /**
   * Estrae la data di aggiornamento dalla riga 2 del foglio
   * Accetta formato "DD/MM/YYYY" o oggetto Date di SheetJS
   */
  function parseDataAggiornamento(ws, XLSX) {
    // Cella C2 (col index 2, row index 1)
    const cell = ws['C2'];
    if (!cell) return null;
    if (cell.t === 'd' || cell.v instanceof Date) {
      const d = cell.v instanceof Date ? cell.v : new Date(cell.v);
      return d.toLocaleDateString('it-IT');
    }
    if (cell.t === 's' || cell.t === 'n') {
      const raw = cell.w || String(cell.v);
      // Già formattata "DD/MM/YYYY"
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) return raw;
      // Numero seriale Excel
      if (!isNaN(raw)) {
        const d = XLSX.SSF.parse_date_code(Number(raw));
        return `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
      }
      return raw;
    }
    return null;
  }

  /**
   * Parsing principale.
   * @param {ArrayBuffer} buffer  – contenuto del file .xlsx
   * @param {object}      XLSX    – libreria SheetJS (window.XLSX)
   * @returns {object}  struttura compatibile con cruscotto_data.json
   */
  function parseRolling(buffer, XLSX) {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: false, cellNF: true });
    const ws = wb.Sheets['REPORT'];
    if (!ws) throw new Error('Foglio "REPORT" non trovato nel file.');

    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });

    // ── Metadati ──────────────────────────────────────────────────────────────
    // Riga 2 (index 1): ['AGGIORNATO:', null, 'DD/MM/YYYY', ...]
    let aggiornato = null;
    if (aoa[1] && aoa[1][2]) {
      const v = aoa[1][2];
      if (typeof v === 'number') {
        // numero seriale Excel → converte
        const d = XLSX.SSF.parse_date_code(v);
        aggiornato = `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
      } else {
        aggiornato = String(v);
      }
    }

    // Riga 3 (index 2): etichette colonne
    const hdrRow = aoa[2] || [];

    // Riga 4 (index 3): totali – la saltiamo
    // Righe dati: da index 4 in poi
    const dataRows = aoa.slice(4);

    // ── Ricava dinamicamente l'anno corrente dai titoli di colonna ────────────
    // Cerca il pattern "GTO Gen'XX" per sapere l'anno GTO
    let annoGTO = '26';
    for (let c = 0; c < hdrRow.length; c++) {
      const h = String(hdrRow[c] || '');
      const m = h.match(/GTO\s+Gen'(\d{2})/i);
      if (m) { annoGTO = m[1]; break; }
    }
    const annoCorr = 2000 + parseInt(annoGTO, 10);
    const annoPrev = annoCorr - 1;
    const annoPrevPrev = annoCorr - 2;

    // ── Mappa colonne per nome (0-based) ──────────────────────────────────────
    // Mesi 2025 (anno precedente): col 9..20  → Gen-Dic annoPrev
    const MESI_PREV_COLS   = Array.from({length: 12}, (_, i) => 9 + i);
    // GTO anno corrente: col 22..33 → Gen-Dic annoCorr
    const GTO_CORR_COLS    = Array.from({length: 12}, (_, i) => 22 + i);

    const MESI_LABELS = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
    const mesi_labels_prev = MESI_LABELS.map(m => `${m} ${annoPrev}`);
    const mesi_labels_corr = MESI_LABELS.map(m => `${m} ${annoCorr}`);
    // chiavi compatibili col JSON esistente (mesi2024 / mesi2025)
    const MESI_PREV_KEYS = MESI_LABELS.map(m => `${m}${String(annoPrev).slice(-2)}`);
    const MESI_CORR_KEYS = MESI_LABELS.map(m => `${m}${String(annoCorr).slice(-2)}`);

    // ── Clienti ───────────────────────────────────────────────────────────────
    const clientiMap = {}; // cod → oggetto

    for (const row of dataRows) {
      if (!row || row[5] == null || row[6] == null) continue; // riga vuota

      const cod    = String(row[5]).trim();
      const nome   = String(row[6]).trim();
      const gamma  = String(row[2] || '').trim() || 'Altro';
      const div    = String(row[0] || '').trim();
      const gav    = String(row[1] || '').trim();

      const fatt_prev     = toNum(row[7]);   // Fatturato TOTALE anno precedente
      const fatt_corr_tot = toNum(row[8]);   // Fatturato TOTALE anno corrente (consuntivo)
      const fatt_prog_26  = toNum(row[21]);  // Progressivo Gen-Mar anno corr.

      // Mesi anno precedente
      const mesiPrev = {};
      MESI_PREV_COLS.forEach((c, i) => {
        const v = toNum(row[c]);
        if (v != null) mesiPrev[MESI_PREV_KEYS[i]] = v;
      });

      // GTO anno corrente (ordini + spedito per mese)
      const mesiCorr = {};
      GTO_CORR_COLS.forEach((c, i) => {
        const v = toNum(row[c]);
        if (v != null) mesiCorr[MESI_CORR_KEYS[i]] = v;
      });

      // Colonne mese corrente
      const consegnato     = toNum(row[34]);
      const inPrep         = toNum(row[35]);
      const daSpedire      = toNum(row[36]);
      const ordinatoOltre  = toNum(row[37]);
      const fattMesePrev   = toNum(row[38]);
      const speditoOrdCorr = toNum(row[39]);
      const varMese        = toNum(row[40]);
      const fattProgPrev   = toNum(row[41]);
      const fattProgCorr   = toNum(row[42]);
      const varProg        = toNum(row[43]);

      // Progressivo anno corrente = progressivo gen-mar + consegnato mese + in prep + da spedire
      const progCorr = fattProgCorr != null ? fattProgCorr :
        ((fatt_prog_26 || 0) + (consegnato || 0) + (inPrep || 0) + (daSpedire || 0));

      // Variazione anno (prog corr vs tot anno prev)
      let var_anno = null;
      if (fatt_prev != null && fatt_prev !== 0 && fatt_corr_tot != null) {
        var_anno = ((fatt_corr_tot - fatt_prev) / fatt_prev) * 100;
      } else if (fattProgPrev != null && fattProgPrev !== 0 && fattProgCorr != null) {
        var_anno = ((fattProgCorr - fattProgPrev) / fattProgPrev) * 100;
      }

      if (clientiMap[cod]) {
        // Aggrega righe duplicate (stesso cliente, gamma diversa)
        const ex = clientiMap[cod];
        ex.fatt2024  = (ex.fatt2024 || 0) + (fatt_prev || 0);
        ex.fatt2025  = (ex.fatt2025 || 0) + (fatt_corr_tot || 0);
        ex.prog2026  = (ex.prog2026 || 0) + progCorr;
        MESI_PREV_KEYS.forEach(k => { ex.mesi2025[k] = (ex.mesi2025[k]||0) + (mesiPrev[k]||0); });
        MESI_CORR_KEYS.forEach(k => { ex.mesi2026[k] = (ex.mesi2026[k]||0) + (mesiCorr[k]||0); });
      } else {
        clientiMap[cod] = {
          cod,
          nome,
          gamma,
          div,
          gav,
          fatt2024: fatt_prev || 0,
          fatt2025: fatt_corr_tot || 0,
          prog2026: progCorr || 0,
          mesi2025: { ...mesiPrev },
          mesi2026: { ...mesiCorr },
          var_anno,
          mese_corrente: {
            consegnato:     consegnato,
            in_preparazione: inPrep,
            da_spedire:     daSpedire,
            ordinato_oltre: ordinatoOltre,
            fatt_mese_prev: fattMesePrev,
            spedito_ord:    speditoOrdCorr,
            var_mese:       varMese,
            fatt_prog_prev: fattProgPrev,
            fatt_prog_corr: fattProgCorr,
            var_prog:       varProg,
          },
          gamma_dettaglio: {}
        };
      }
    }

    const clientiArr = Object.values(clientiMap);

    // Ricalcola var_anno aggregata
    clientiArr.forEach(c => {
      if (c.fatt2024 > 0 && c.fatt2025 != null) {
        c.var_anno = ((c.fatt2025 - c.fatt2024) / c.fatt2024) * 100;
      }
    });

    // Gamma unica per cliente: quella con fatturato più alto se ci sono righe multiple
    // (già gestito: prende la prima. Per un dataset pulito va bene.)

    return {
      aggiornato:       aggiornato || new Date().toLocaleDateString('it-IT'),
      agente:           'Cubaiu Loris',
      cod_agente:       '400542',
      anno_prev:        annoPrev,
      anno_corr:        annoCorr,
      mesi_labels_prev,
      mesi_labels_corr,
      mesi_keys_prev:   MESI_PREV_KEYS,
      mesi_keys_corr:   MESI_CORR_KEYS,
      // Compatibilità con il vecchio cruscotto_data.json
      mesi_labels_2024: mesi_labels_prev,
      mesi_labels_2025: mesi_labels_corr,
      settori_gamma: [...new Set(clientiArr.map(c => c.gamma))].filter(Boolean).sort(),
      clienti: clientiArr
    };
  }

  // Esporta
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = parseRolling;
  } else {
    global.parseRolling = parseRolling;
  }

})(typeof globalThis !== 'undefined' ? globalThis : window);
