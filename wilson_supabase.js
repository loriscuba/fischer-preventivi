// ═══════════════════════════════════════════ WILSON SUPABASE ENGINE ═══════

const SUPABASE_URL = 'https://btcnfxoqkddpfhzahhaz.supabase.co';
const SUPABASE_KEY = 'sb_publishable_bppoTX0oooogt3v2qj6zbQ_XSSyOpY2';

var WILSON_ORDINI = [];

async function wilsonCaricaDaSupabase() {
  const status = document.getElementById('wilson-status');
  status.textContent = '⏳ Caricamento ordini da Supabase…';
  status.style.color = 'var(--warn)';

  try {
    const rOrdini = await fetch(
      `${SUPABASE_URL}/rest/v1/ordini?select=*&order=data.desc`,
      { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } }
    );
    if (!rOrdini.ok) throw new Error(`Ordini: ${rOrdini.status}`);
    const ordini = await rOrdini.json();

    const rRighe = await fetch(
      `${SUPABASE_URL}/rest/v1/righe_ordine?select=*`,
      { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } }
    );
    if (!rRighe.ok) throw new Error(`Righe: ${rRighe.status}`);
    const righe = await rRighe.json();

    const righeMap = {};
    righe.forEach(r => {
      if (!righeMap[r.numero_ordine]) righeMap[r.numero_ordine] = [];
      righeMap[r.numero_ordine].push(r);
    });

    WILSON_ORDINI = ordini.map(o => ({
      id:            o.numero_ordine,
      numero_ordine: o.numero_ordine,
      data:          o.data ? o.data.split('-').reverse().join('.') : '—',
      cliente:       o.destinazione || '—',
      num_cliente:   o.numero_cliente || '—',
      totale:        parseFloat(o.totale) || 0,
      pagamento:     o.condizioni_pagamento || '—',
      n_articoli:    (righeMap[o.numero_ordine] || []).length,
      articoli:      (righeMap[o.numero_ordine] || []).map(r => ({
        cod_art:       r.cod_art,
        descrizione:   r.descrizione,
        quantita:      r.quantita,
        um:            r.um,
        prezzo_unit:   parseFloat(r.prezzo_unitario) || 0,
        sconti:        [r.sconto_1, r.sconto_2].filter(Boolean).join('/ '),
        importo:       parseFloat(r.importo) || 0,
        data_consegna: r.data_consegna
      }))
    }));

    status.textContent = `✓ ${WILSON_ORDINI.length} ordini caricati da Supabase`;
    status.style.color = 'var(--ok)';
    wilsonRender();

  } catch(e) {
    status.textContent = `✗ Errore: ${e.message}`;
    status.style.color = 'var(--danger)';
    console.error(e);
  }
}

function wilsonRender() {
  const q     = (document.getElementById('wilson-search')?.value || '').toLowerCase();
  const sort  = document.getElementById('wilson-sort')?.value || 'data-desc';
  const tbody = document.getElementById('wilson-tbody');
  if (!tbody) return;

  let lista = WILSON_ORDINI.filter(o => {
    if (!q) return true;
    return (o.numero_ordine || '').includes(q) ||
           (o.cliente || '').toLowerCase().includes(q) ||
           (o.articoli || []).some(a =>
             (a.cod_art || '').includes(q) ||
             (a.descrizione || '').toLowerCase().includes(q));
  });

  lista.sort((a, b) => {
    if (sort === 'data-desc')   return dataCmp(b.data, a.data);
    if (sort === 'data-asc')    return dataCmp(a.data, b.data);
    if (sort === 'totale-desc') return b.totale - a.totale;
    return a.totale - b.totale;
  });

  if (lista.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;color:var(--muted);font-family:var(--mono);font-size:12px;padding:40px;">${WILSON_ORDINI.length ? 'Nessun risultato' : 'Nessun ordine trovato'}</td></tr>`;
  } else {
    tbody.innerHTML = lista.map(o => `
      <tr style="cursor:pointer;" onclick="wilsonDrawer('${o.numero_ordine}')">
        <td style="text-align:center;color:var(--muted);font-size:16px;">▶</td>
        <td><span style="font-family:var(--mono);font-size:12px;font-weight:700;">${o.numero_ordine || '—'}</span></td>
        <td><span style="font-family:var(--mono);font-size:12px;">${o.data || '—'}</span></td>
        <td>${o.cliente || '—'}</td>
        <td><span style="font-family:var(--mono);font-size:11px;color:var(--muted);">${o.num_cliente || '—'}</span></td>
        <td style="text-align:right;font-family:var(--mono);font-size:12px;">${o.n_articoli}</td>
        <td style="text-align:right;font-family:var(--mono);font-size:13px;font-weight:700;color:var(--accent);">€ ${fmtEur(o.totale)}</td>
        <td style="font-family:var(--mono);font-size:10px;color:var(--muted);">${o.pagamento || '—'}</td>
      </tr>`).join('');
  }

  const totale = lista.reduce((s, o) => s + (o.totale || 0), 0);
  const artSet = new Set(lista.flatMap(o => (o.articoli || []).map(a => a.cod_art)));
  document.getElementById('wk-ordini').textContent   = lista.length;
  document.getElementById('wk-totale').textContent   = '€ ' + fmtEur(totale);
  document.getElementById('wk-articoli').textContent = artSet.size;
  document.getElementById('wk-media').textContent    = lista.length ? '€ ' + fmtEur(totale / lista.length) : '—';
}

function dataCmp(a, b) {
  const p = s => { if (!s || s === '—') return 0; const [d, m, y] = s.split('.'); return +y * 10000 + +m * 100 + +d; };
  return p(a) - p(b);
}

function fmtEur(n) {
  return (n || 0).toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function wilsonDrawer(numOrdine) {
  const ordine = WILSON_ORDINI.find(o => o.numero_ordine === numOrdine);
  if (!ordine) return;
  document.getElementById('wilson-drawer-title').textContent =
    `Ordine ${ordine.numero_ordine} · ${ordine.cliente} · ${ordine.data}`;
  const tbody = document.getElementById('wilson-art-tbody');
  tbody.innerHTML = (ordine.articoli || []).map(a => `
    <tr>
      <td><span style="font-family:var(--mono);font-size:11px;">${a.cod_art}</span></td>
      <td>${a.descrizione}</td>
      <td style="text-align:right;font-family:var(--mono);font-size:12px;">${(a.quantita || 0).toLocaleString('it-IT')}</td>
      <td style="font-family:var(--mono);font-size:11px;color:var(--muted);">${a.um}</td>
      <td style="text-align:right;font-family:var(--mono);font-size:12px;">€ ${fmtEur(a.prezzo_unit)}</td>
      <td style="font-family:var(--mono);font-size:10px;color:var(--muted);">${a.sconti || '—'}</td>
      <td style="text-align:right;font-family:var(--mono);font-size:13px;font-weight:700;">€ ${fmtEur(a.importo)}</td>
    </tr>`).join('');
  const drawer = document.getElementById('wilson-drawer');
  drawer.style.display = 'block';
  drawer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function wilsonExportCSV() {
  if (!WILSON_ORDINI.length) return;
  const rows = [['N° Ordine', 'Data', 'Cliente', 'N° Cliente', 'Cod. Art.', 'Descrizione', 'Qtà', 'UM', 'Prezzo unit.', 'Sconti %', 'Importo EUR', 'Totale Ordine', 'Pagamento']];
  WILSON_ORDINI.forEach(o => {
    (o.articoli || []).forEach(a => {
      rows.push([o.numero_ordine, o.data, o.cliente, o.num_cliente,
        a.cod_art, a.descrizione, a.quantita, a.um,
        a.prezzo_unit, a.sconti, a.importo, o.totale, o.pagamento]);
    });
  });
  const csv = rows.map(r => r.map(v => `"${String(v || '').replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `wilson_ordini_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
}

// Patch showPanel: auto-carica Supabase al click su Wilson
const _origShowPanel = window.showPanel;
window.showPanel = function(id, el) {
  if (typeof _origShowPanel === 'function') _origShowPanel(id, el);
  if (id === 'wilson') wilsonCaricaDaSupabase();
};

// Init
wilsonRender();
