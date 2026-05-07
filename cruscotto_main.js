// ═══════════════════════════════════════════════════════
//  STATO GLOBALE
// ═══════════════════════════════════════════════════════
let DATA = null;
let clienti = [];
let sortCol = 'fatt2024', sortDir = -1;
let mChart = null;

let ANNO_PREV = 2025, ANNO_CORR = 2026;
let MESI_KEYS_PREV   = [];
let MESI_KEYS_CORR   = [];
let MESI_LABELS_PREV = [];
let MESI_LABELS_CORR = [];
const MESI_SHORT = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];

// ═══════════════════════════════════════════════════════
//  UTILS
// ═══════════════════════════════════════════════════════
const fmt    = v => (v == null || isNaN(v)) ? '—' : '€\u202F' + Math.round(v).toLocaleString('it-IT');
const fmtK   = v => (v == null || isNaN(v)) ? '—' : '€\u202F' + Math.round(v/1000).toLocaleString('it-IT') + 'k';
const fmtPct = v => (v == null || isNaN(v)) ? '—' : (v >= 0 ? '+' : '') + Math.round(v) + '%';

// Funzione per ottenere il mese precedente (0-11)
function getPrevMonth() {
  const now = new Date();
  return now.getMonth() === 0 ? 11 : now.getMonth() - 1;
}

// Funzione per ottenere il nome del mese precedente
function getPrevMonthName() {
  const mesi = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
  return mesi[getPrevMonth()];
}

// Funzione per ottenere il progressivo fino al mese precedente
// Struttura: cliente.mesi2025 = { Gen25: 100, Feb25: 200, ... }
//            cliente.mesi2026 = { Gen26: 50, Feb26: 150, ... }
function getProgressivoUntilPrevMonth(cliente, anno) {
  const prevMese = getPrevMonth(); // 0-11
  const MESI_LABELS = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
  const annoSuffix = anno === 2025 ? '25' : '26'; // '25' o '26'
  const mesiObj = anno === 2025 ? (cliente.mesi2025 || {}) : (cliente.mesi2026 || {});
  
  let sum = 0;
  
  // Somma da mese 0 (Gen) fino al mese precedente
  for (let m = 0; m <= prevMese; m++) {
    const meseKey = MESI_LABELS[m] + annoSuffix; // 'Gen25', 'Feb25', ..., 'Apr25', etc.
    const val = mesiObj[meseKey];
    if (val !== undefined && val !== null) {
      sum += (val || 0);
    }
  }
  
  // Se abbiamo trovato dati, ritorna la somma
  if (sum > 0) return sum;
  
  // Fallback: usa il totale annuale se non abbiamo i dati mensili
  if (anno === 2025) {
    return cliente.fatt2025 || cliente._prev || 0;
  }
  if (anno === 2026) {
    return cliente.prog2026 || cliente._prog || 0;
  }
  
  return 0;
}

function renderOverview() {
  console.log('[DEBUG] renderOverview() called');
  console.log('[DEBUG] clienti.length:', clienti.length);
  console.log('[DEBUG] DATA:', DATA);
  
  if (!DATA || !clienti || clienti.length === 0) {
    console.warn('[DEBUG] No data or clienti empty, returning early');
    return;
  }

  const today = new Date();
  const meseCorrente = today.getMonth() + 1; // 1-12
  const annoCorrente = today.getFullYear();
  const prevMese = getPrevMonth(); // 0-11
  const prevMeseName = getPrevMonthName();
  
  // ════════════════════════════════════════════════════════
  // PROGRESSIVI (fino al mese precedente)
  // ════════════════════════════════════════════════════════
  const prog2025 = clienti.reduce((sum, c) => sum + getProgressivoUntilPrevMonth(c, 2025), 0);
  const prog2026 = clienti.reduce((sum, c) => sum + getProgressivoUntilPrevMonth(c, 2026), 0);
  const deltaProg = prog2025 === 0 ? null : Math.round(((prog2026 - prog2025) / prog2025) * 100);
  
  document.getElementById('kpi-prog-2025').textContent = fmtK(prog2025);
  document.getElementById('kpi-prog-2025-mese').textContent = `fino a ${prevMeseName.toLowerCase()}`;
  document.getElementById('kpi-prog-2026').textContent = fmtK(prog2026);
  document.getElementById('kpi-prog-2026-mese').textContent = `fino a ${prevMeseName.toLowerCase()}`;
  
  const deltaEl = document.getElementById('kpi-prog-delta');
  if (deltaProg !== null) {
    deltaEl.className = 'kpi-delta ' + deltaClass(deltaProg);
    deltaEl.textContent = fmtPct(deltaProg);
  } else {
    deltaEl.textContent = '—';
    deltaEl.className = 'kpi-delta neu';
  }
  
  // ════════════════════════════════════════════════════════
  // KPI SECONDARI
  // ════════════════════════════════════════════════════════
  const nClienti = clienti.length;
  const mediaClienti = nClienti > 0 ? prog2026 / nClienti : 0;
  
  document.getElementById('kpi-clienti').textContent = nClienti;
  document.getElementById('kpi-media-cliente').textContent = fmtK(mediaClienti);
  
  // ════════════════════════════════════════════════════════
  // MESE CORRENTE (ora - fino ad ora)
  // ════════════════════════════════════════════════════════
  let meseMese2025 = 0, meseMese2026 = 0;
  clienti.forEach(c => {
    const key2025 = `m${meseCorrente}_2025`;
    const key2026 = `m${meseCorrente}_2026`;
    meseMese2025 += (c[key2025] || 0);
    meseMese2026 += (c[key2026] || 0);
  });
  
  const deltaMese = meseMese2025 === 0 ? null : Math.round(((meseMese2026 - meseMese2025) / meseMese2025) * 100);
  
  document.getElementById('mese-2025').textContent = fmt(meseMese2025);
  document.getElementById('mese-2026').textContent = fmt(meseMese2026);
  
  const meseDeltaEl = document.getElementById('mese-delta');
  if (deltaMese !== null) {
    meseDeltaEl.className = 'kpi-delta ' + deltaClass(deltaMese);
    meseDeltaEl.textContent = fmtPct(deltaMese);
  } else {
    meseDeltaEl.textContent = '—';
    meseDeltaEl.className = 'kpi-delta neu';
  }
  
  // ════════════════════════════════════════════════════════
  // HEADER DATA
  // ════════════════════════════════════════════════════════
  document.getElementById('header-date').textContent = today.toLocaleDateString('it-IT', {weekday:'long',day:'numeric',month:'long',year:'numeric'});
  
  // ════════════════════════════════════════════════════════
  // GRAFICI
  // ════════════════════════════════════════════════════════
  renderCharts();
  
  // ════════════════════════════════════════════════════════
  // TOP CLIENTI (mese precedente)
  // ════════════════════════════════════════════════════════
  renderTopClienti('top10');
}

function renderTopClienti(tipo = 'top10') {
  if (!clienti || clienti.length === 0) return;
  
  const prevMese = getPrevMonth(); // 0-11
  const MESI_LABELS = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
  const meseLabelPrev = MESI_LABELS[prevMese]; // 'Gen', 'Feb', 'Apr', etc.
  
  // Calcola i dati per il mese precedente per entrambi gli anni
  const clientiData = clienti.map(c => {
    const mesi2025 = c.mesi2025 || {};
    const mesi2026 = c.mesi2026 || {};
    
    const val2025 = mesi2025[meseLabelPrev + '25'] || 0;
    const val2026 = mesi2026[meseLabelPrev + '26'] || 0;
    
    const delta = val2025 === 0 ? 0 : Math.round(((val2026 - val2025) / val2025) * 100);
    
    return {
      nome: c.nome || 'N/A',
      val2025,
      val2026,
      delta,
      orig: c
    };
  }).filter(c => c.val2025 > 0 || c.val2026 > 0); // Solo clienti con movimento
  
  // Ordina per 2026 decrescente
  clientiData.sort((a, b) => b.val2026 - a.val2026);
  
  // Seleziona Top 10 o Ultimi 10
  let toShow = clientiData.slice(0, 10);
  if (tipo === 'bottom10') {
    toShow = clientiData.slice(-10).reverse();
  }
  
  // Renderizza tabella
  const tbody = document.getElementById('top-clienti-tbody');
  tbody.innerHTML = toShow.map((c, i) => {
    const deltaClass = c.delta > 0 ? 'pos' : c.delta < 0 ? 'neg' : 'neu';
    return `<tr>
      <td style="text-align:center;font-weight:600;color:var(--text-muted);">${i + 1}</td>
      <td class="td-nome">${c.nome}</td>
      <td style="text-align:right;">${fmt(c.val2025)}</td>
      <td style="text-align:right;">${fmt(c.val2026)}</td>
      <td style="text-align:center;"><span class="pill ${deltaClass}">${fmtPct(c.delta)}</span></td>
    </tr>`;
  }).join('');
}

function renderCharts() {
  if (!clienti || clienti.length === 0) return;
  
  // Dati mensili per il grafico
  const andamentoData2025 = [];
  const andamentoData2026 = [];
  const progressivoData2025 = [];
  const progressivoData2026 = [];
  
  const MESI_LABELS = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
  
  let cumul2025 = 0, cumul2026 = 0;
  
  for (let m = 0; m < 12; m++) {
    const meseLbl = MESI_LABELS[m];
    const key2025 = meseLbl + '25'; // 'Gen25', 'Feb25', etc.
    const key2026 = meseLbl + '26'; // 'Gen26', 'Feb26', etc.
    
    let mese2025 = 0, mese2026 = 0;
    clienti.forEach(c => {
      const mesi2025 = c.mesi2025 || {};
      const mesi2026 = c.mesi2026 || {};
      mese2025 += (mesi2025[key2025] || 0);
      mese2026 += (mesi2026[key2026] || 0);
    });
    
    andamentoData2025.push(mese2025);
    andamentoData2026.push(mese2026);
    
    cumul2025 += mese2025;
    cumul2026 += mese2026;
    
    progressivoData2025.push(cumul2025);
    progressivoData2026.push(cumul2026);
  }
  
  const labels = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
  
  // Distruggi grafici precedenti se esistono
  if (window.chartAndamento) window.chartAndamento.destroy();
  if (window.chartProgressivo) window.chartProgressivo.destroy();
  
  // Grafico Andamento Mensile
  const ctxAnd = document.getElementById('chart-andamento');
  if (ctxAnd) {
    window.chartAndamento = new Chart(ctxAnd, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          {
            label: '2025',
            data: andamentoData2025,
            backgroundColor: '#e2e8f0',
            borderColor: '#6e7681',
            borderWidth: 0,
            borderRadius: 4
          },
          {
            label: '2026',
            data: andamentoData2026,
            backgroundColor: '#0969da',
            borderColor: '#0969da',
            borderWidth: 0,
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: true,
            position: 'bottom',
            labels: {font:{size:11,family:"'DM Mono',monospace"},color:'var(--text-muted)',boxHeight:6,padding:12}
          },
          tooltip: {
            backgroundColor: 'rgba(13,17,23,0.9)',
            titleFont: {family:"'DM Mono',monospace",size:11},
            bodyFont: {family:"'DM Mono',monospace",size:11},
            padding: 8,
            displayColors: true,
            callbacks: {
              label: ctx => ctx.dataset.label + ': ' + fmt(ctx.parsed.y)
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {font:{size:10,family:"'DM Mono',monospace"},color:'var(--text-muted)',callback:v=>''},
            grid: {color:'var(--border-subtle)',drawBorder:false},
            border: {display:false}
          },
          x: {
            ticks: {font:{size:10,family:"'DM Mono',monospace"},color:'var(--text-muted)'},
            grid: {display:false},
            border: {display:false}
          }
        }
      }
    });
  }
  
  // Grafico Progressivo
  const ctxProg = document.getElementById('chart-progressivo');
  if (ctxProg) {
    window.chartProgressivo = new Chart(ctxProg, {
      type: 'line',
      data: {
        labels,
        datasets: [
          {
            label: 'Progressivo 2025',
            data: progressivoData2025,
            borderColor: '#6e7681',
            backgroundColor: 'rgba(110,118,129,0.05)',
            borderWidth: 2,
            tension: 0.4,
            fill: true,
            pointRadius: 3,
            pointBackgroundColor: '#6e7681'
          },
          {
            label: 'Progressivo 2026',
            data: progressivoData2026,
            borderColor: '#0969da',
            backgroundColor: 'rgba(9,105,218,0.05)',
            borderWidth: 2,
            tension: 0.4,
            fill: true,
            pointRadius: 3,
            pointBackgroundColor: '#0969da'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: true,
            position: 'bottom',
            labels: {font:{size:11,family:"'DM Mono',monospace"},color:'var(--text-muted)',boxHeight:6,padding:12}
          },
          tooltip: {
            backgroundColor: 'rgba(13,17,23,0.9)',
            titleFont: {family:"'DM Mono',monospace",size:11},
            bodyFont: {family:"'DM Mono',monospace",size:11},
            padding: 8,
            displayColors: true,
            callbacks: {
              label: ctx => ctx.dataset.label + ': ' + fmt(ctx.parsed.y)
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {font:{size:10,family:"'DM Mono',monospace"},color:'var(--text-muted)',callback:v=>''},
            grid: {color:'var(--border-subtle)',drawBorder:false},
            border: {display:false}
          },
          x: {
            ticks: {font:{size:10,family:"'DM Mono',monospace"},color:'var(--text-muted)'},
            grid: {display:false},
            border: {display:false}
          }
        }
      }
    });
  }
}

function deltaClass(v) {
  if (v == null || isNaN(v)) return 'neu';
  return v > 0 ? 'pos' : v < 0 ? 'neg' : 'neu';
}

const badgeClass = g => {
  const gg = (g || '').toLowerCase();
  if (gg.includes('edile'))  return 'badge-edile';
  if (gg.includes('elet'))   return 'badge-elettrico';
  if (gg.includes('its'))    return 'badge-its';
  if (gg.includes('ferr'))   return 'badge-ferr';
  if (gg.includes('diy'))    return 'badge-diy';
  return 'badge-default';
};

// ═══════════════════════════════════════════════════════
//  TABS
// ═══════════════════════════════════════════════════════
function showPanel(id, el) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('panel-' + id).classList.add('active');
  if (el) el.classList.add('active');
  if (id === 'clienti') renderClientiTab();
  if (id === 'gamma' && DATA) renderGamma();
  if (id === 'pipeline') renderPipeline();
  if (id === 'wilson') wilsonCaricaDaSupabase();
}

// ═══════════════════════════════════════════════════════
//  PIPELINE
// ═══════════════════════════════════════════════════════
function renderPipeline() {
  const today = new Date();
  const meseCorrente = today.getMonth() + 1;
  const cards = document.getElementById('pipeline-cards');
  if (!cards) return;
  const filter = (document.getElementById('pipeline-filter') || {}).value || 'all';

  const rows = PIPELINE_STORICO.map(p => {
    const clienteRolling = clienti.find(c =>
      (c.nome || '').toLowerCase().includes(p.org.toLowerCase().split(' ')[0].toLowerCase()) ||
      p.org.toLowerCase().includes((c.nome || '').toLowerCase().split(' ')[0].toLowerCase())
    );
    const ultimo = new Date(p.ultimo);
    const giorniDaUltimo = Math.floor((today - ultimo) / 86400000);
    const atteso = p.gap || 30;
    const giorniRitardo = giorniDaUltimo - atteso;
    const giornoSuggerito = p.giorno || 15;
    const appuntamento = new Date(today.getFullYear(), today.getMonth(), Math.min(giornoSuggerito, 28));
    if (appuntamento < today) appuntamento.setMonth(appuntamento.getMonth() + 1);
    if (appuntamento.getDay() === 6) appuntamento.setDate(appuntamento.getDate() + 2);
    else if (appuntamento.getDay() === 0) appuntamento.setDate(appuntamento.getDate() + 1);
    let stato, statoColor, statolabel;
    if (p.gap && p.gap <= 60) {
      if (giorniRitardo > 7) { stato = 'overdue'; statoColor = 'var(--danger)'; statolabel = `Scaduto (${giorniRitardo}gg fa)`; }
      else if (giorniRitardo > 0) { stato = 'due_soon'; statoColor = 'var(--warn)'; statolabel = `In scadenza (${giorniRitardo}gg)`; }
      else { stato = 'ok'; statoColor = 'var(--ok)'; statolabel = `In regola (tra ${-giorniRitardo}gg)`; }
    } else {
      stato = 'irregular'; statoColor = 'var(--muted2)'; statolabel = 'Frequenza variabile';
    }
    return { ...p, stato, statoColor, statolabel, giorniDaUltimo, giorniRitardo, appuntamento, clienteRolling };
  });

  const ordine = { overdue: 0, due_soon: 1, ok: 2, irregular: 3 };
  rows.sort((a, b) => ordine[a.stato] - ordine[b.stato] || b.giorniRitardo - a.giorniRitardo);
  const filtered = filter === 'all' ? rows : rows.filter(r => r.stato === filter);
  const mesiFmt = ['','Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];

  cards.innerHTML = filtered.map(r => {
    const mesiBar = Array.from({length:12}, (_,i) => {
      const m = i + 1;
      const ha = r.mesi_count && r.mesi_count[String(m)];
      const isNow = m === meseCorrente;
      return `<span title="${mesiFmt[m]}" style="width:16px;height:16px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:8px;font-weight:700;
        background:${ha ? (isNow ? 'var(--accent)' : 'var(--ok-l)') : (isNow ? 'var(--warn-l)' : 'var(--surface2)')};
        color:${ha ? (isNow ? '#fff' : 'var(--ok)') : (isNow ? 'var(--warn)' : 'var(--muted2)')};
        border:1px solid ${ha ? (isNow ? 'var(--accent)' : 'var(--ok)') : (isNow ? 'var(--warn)' : 'var(--border)')};
        cursor:default;">${mesiFmt[m][0]}</span>`;
    }).join('');
    const progRolling = r.clienteRolling ? `
      <div style="display:flex;justify-content:space-between;font-size:11px;font-family:var(--mono);color:var(--muted);margin-top:6px;">
        <span>Progressivo anno</span>
        <span style="color:var(--accent);font-weight:600;">${fmt(r.clienteRolling._prog || 0)}</span>
      </div>` : '';
    const appStr = r.appuntamento.toLocaleDateString('it-IT', {weekday:'short',day:'numeric',month:'short'});
    return `<div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px;box-shadow:var(--shadow);border-left:3px solid ${r.statoColor};">
      <div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px;">
        <div style="flex:1;min-width:0;">
          <div style="font-weight:700;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${r.org}">${r.org}</div>
          <div style="font-size:11px;color:${r.statoColor};font-family:var(--mono);margin-top:2px;font-weight:600;">${r.statolabel}</div>
        </div>
        <button data-org="${r.org}" data-gap="${r.gap||30}" data-media="${r.media}" data-ritardo="${r.giorniRitardo||0}" data-mesi="${r.mesi}"
          onclick="chiediAI(this.dataset.org, +this.dataset.gap, +this.dataset.media, +this.dataset.ritardo, +this.dataset.mesi)"
          style="flex-shrink:0;margin-left:8px;font-family:var(--mono);font-size:10px;padding:4px 8px;border:1px solid var(--accent);background:transparent;color:var(--accent);border-radius:5px;cursor:pointer;">✦ AI</button>
      </div>
      <div style="display:flex;gap:10px;margin-bottom:10px;flex-wrap:wrap;">
        <div style="flex:1;min-width:80px;background:var(--surface2);border-radius:7px;padding:7px 10px;">
          <div style="font-size:10px;color:var(--muted);font-family:var(--mono);">Freq. media</div>
          <div style="font-weight:700;font-size:14px;">${r.gap ? r.gap+'gg' : '–'}</div>
        </div>
        <div style="flex:1;min-width:80px;background:var(--surface2);border-radius:7px;padding:7px 10px;">
          <div style="font-size:10px;color:var(--muted);font-family:var(--mono);">Media ordine</div>
          <div style="font-weight:700;font-size:14px;">${fmt(r.media)}</div>
        </div>
        <div style="flex:1;min-width:80px;background:var(--surface2);border-radius:7px;padding:7px 10px;">
          <div style="font-size:10px;color:var(--muted);font-family:var(--mono);">N° ordini</div>
          <div style="font-weight:700;font-size:14px;">${r.n}</div>
        </div>
      </div>
      ${progRolling}
      <div style="margin-top:10px;">
        <div style="font-size:10px;color:var(--muted);font-family:var(--mono);margin-bottom:5px;">Attività per mese</div>
        <div style="display:flex;gap:3px;flex-wrap:nowrap;">${mesiBar}</div>
      </div>
      <div style="margin-top:10px;padding:8px 10px;background:var(--accent-l);border-radius:7px;display:flex;justify-content:space-between;align-items:center;">
        <div>
          <div style="font-size:10px;color:var(--accent);font-family:var(--mono);">Prossimo appuntamento suggerito</div>
          <div style="font-weight:700;font-size:12px;color:var(--accent);">${appStr} (giorno ~${r.giorno})</div>
        </div>
        <div style="font-size:10px;color:var(--muted);font-family:var(--mono);text-align:right;">
          Ultimo ordine<br><strong style="color:var(--text);">${new Date(r.ultimo).toLocaleDateString('it-IT',{day:'2-digit',month:'short',year:'2-digit'})}</strong>
        </div>
      </div>
      <div id="ai-${r.org.replace(/[^a-zA-Z0-9]/g,'_')}" style="display:none;margin-top:8px;padding:8px 10px;background:var(--surface2);border:1px solid var(--border);border-radius:7px;font-size:11px;line-height:1.6;font-family:var(--mono);color:var(--text2);"></div>
    </div>`;
  }).join('') || '<div style="color:var(--muted);font-family:var(--mono);font-size:13px;padding:20px;">Nessun cliente in questa categoria.</div>';
}

async function chiediAI(org, gap, media, ritardo, mesiAttivi) {
  const safeId = org.replace(/[^a-zA-Z0-9]/g, '_');
  const el = document.getElementById('ai-' + safeId);
  if (!el) return;
  el.style.display = 'block';
  el.innerHTML = '⏳ Generando suggerimento...';
  const cr = clienti.find(c => (c.nome||'').toLowerCase().includes(org.toLowerCase().split(' ')[0].toLowerCase()));
  const progInfo = cr ? `Progressivo anno corrente: €${(cr._prog||0).toLocaleString('it-IT')}. Mese corrente: €${((cr.mese_corrente||{}).consegnato||0).toLocaleString('it-IT')}.` : '';
  const prompt = `Sei un assistente commerciale esperto. Fornisci un suggerimento pratico e conciso (3-4 righe) per il cliente "${org}" basandoti su questi dati:
- Frequenza media ordini: ogni ${gap} giorni
- Valore medio ordine: €${media}
- Mesi attivi su 12: ${mesiAttivi}
- Ritardo sull'ordine atteso: ${ritardo > 0 ? ritardo+' giorni' : 'nessuno'}
${progInfo}
Suggerisci: quando contattarlo, cosa proporre, e un'azione specifica. Sii diretto e operativo. Rispondi in italiano senza elenchi puntati.`;
  try {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'anthropic-dangerous-direct-browser-access': 'true' },
      body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 200, messages: [{ role: 'user', content: prompt }] })
    });
    const data = await res.json();
    el.innerHTML = '✦ ' + (data.content?.[0]?.text || 'Nessuna risposta');
  } catch(e) {
    el.innerHTML = '⚠ Errore AI: ' + e.message;
  }
}

async function generaAgendaMese() {
  const modal = document.getElementById('agenda-modal');
  const agendaEl = document.getElementById('agenda-content');
  modal.style.display = 'block';
  agendaEl.innerHTML = '⏳ Generando agenda ottimizzata per zona...';
  const oggi = new Date();
  const mese = oggi.toLocaleDateString('it-IT', {month:'long', year:'numeric'});
  const prioritari = PIPELINE_STORICO
    .filter(p => p.gap && p.gap <= 50)
    .map(p => {
      const giorniDaUltimo = Math.floor((oggi - new Date(p.ultimo)) / 86400000);
      const ritardo = giorniDaUltimo - (p.gap || 30);
      const addr = CLIENTI_ADDRESSES.find(a =>
        a.cliente.toLowerCase().includes(p.org.toLowerCase().split(' ')[0]) ||
        p.org.toLowerCase().includes(a.cliente.toLowerCase().split(' ')[0])
      );
      return { ...p, ritardo, citta: addr?.citta || '–', indirizzo: addr?.indirizzo_corretto || '–' };
    })
    .sort((a, b) => b.ritardo - a.ritardo)
    .slice(0, 18);
  const perZona = {};
  prioritari.forEach(p => {
    const z = p.citta || 'Altro';
    if (!perZona[z]) perZona[z] = [];
    perZona[z].push(p);
  });
  const listaZone = Object.entries(perZona)
    .map(([zona, ps]) => `ZONA ${zona}: ${ps.map(p => `${p.org} (ogni ${p.gap}gg, €${p.media} media, ritardo ${p.ritardo > 0 ? p.ritardo+'gg' : 'ok'}, indirizzo: ${p.indirizzo})`).join(' | ')}`)
    .join('\n');
  const prompt = `Sei un assistente commerciale esperto per un agente di zona in Liguria. Crea un'agenda settimanale OTTIMIZZATA per il mese di ${mese} organizzando le visite PER ZONA GEOGRAFICA per minimizzare i km percorsi.\n\nClienti:\n${listaZone}\n\nRegole:\n- Organizza per settimana (Settimana 1, 2, 3, 4)\n- Raggruppa clienti della stessa zona\n- Priorità a chi ha più ritardo\n- Max 4-5 clienti al giorno\n- Formato: Settimana X → Lunedì: [zona] cliente1 + cliente2. Martedì: ...\n- In italiano, concreto e operativo.`;
  try {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'anthropic-dangerous-direct-browser-access': 'true' },
      body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 800, messages: [{ role: 'user', content: prompt }] })
    });
    const data = await res.json();
    const testo = (data.content?.[0]?.text || 'Nessuna risposta')
      .replace(/\n/g, '<br>')
      .replace(/Settimana \d/g, m => `<br><strong style="color:var(--accent)">${m}</strong>`)
      .replace(/Lunedì:|Martedì:|Mercoledì:|Giovedì:|Venerdì:/g, m => `<br><em>${m}</em>`);
    agendaEl.innerHTML = testo;
  } catch(e) {
    agendaEl.innerHTML = '⚠ Errore: ' + e.message;
  }
}

// ═══════════════════════════════════════════════════════
//  CARICAMENTO DATI
// ═══════════════════════════════════════════════════════
async function caricaDati() {
  const repo = (localStorage.getItem('gh_repo') || 'loriscuba/fischer-preventivi').trim();
  if (repo) {
    await caricaXlsxDaGitHub(repo);
  } else {
    try {
      const r = await fetch('cruscotto_data.json');
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      applicaDati(await r.json());
    } catch(e) {
      mostraErrore('Nessun dato. Vai in "Aggiorna Dati" per caricare i file Excel.');
    }
  }
}

async function caricaXlsxDaGitHub(repo) {
  try {
    const rR = await fetch(`https://raw.githubusercontent.com/${repo}/data/data/rolling.xlsx?t=${Date.now()}`);
    if (!rR.ok) throw new Error('rolling.xlsx non trovato');
    const bufR = await rR.arrayBuffer();
    const json = parseRolling(bufR, XLSX);
    try {
      const rG = await fetch(`https://raw.githubusercontent.com/${repo}/data/data/gamma.xlsx?t=${Date.now()}`);
      if (rG.ok) {
        const bufG = await rG.arrayBuffer();
        const codici = new Set(json.clienti.map(c => c.cod));
        const { byCliente } = parseGamma(bufG, XLSX, codici);
        json.clienti.forEach(c => { if (byCliente[c.cod]) c.gamma_dettaglio = byCliente[c.cod].settori; });
      }
    } catch(_) {}
    try {
      const rC = await fetch(`https://raw.githubusercontent.com/${repo}/data/data/cedi.xlsx?t=${Date.now()}`);
      if (rC.ok) {
        const bufC = await rC.arrayBuffer();
        const wb = XLSX.read(bufC, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        const COD_VENDITORE = 400542;
        const cediMap = {};
        for (const r of rows) {
          const codVend = Number(r[7]);
          const codCli  = String(r[9]).trim().split('.')[0];
          const nomeCedi = String(r[10]).trim();
          const fatt = Number(r[11]) || 0;
          if (codVend !== COD_VENDITORE) continue;
          if (!codCli || isNaN(Number(codCli))) continue;
          if (!cediMap[codCli]) cediMap[codCli] = [];
          cediMap[codCli].push({ nome: nomeCedi, fatt });
        }
        json.clienti.forEach(c => {
          const parts = String(c.cod).trim().split('/');
          const codCorto = parts[parts.length - 1];
          if (cediMap[codCorto]) c.cedi_dettaglio = cediMap[codCorto];
        });
      }
    } catch(_) {}
    applicaDati(json);
  } catch(e) {
    mostraErrore(`Impossibile caricare dati: ${e.message}`);
  }
}

function applicaGammaBuf(buf) {
  try {
    const codici = new Set(clienti.map(c => c.cod));
    const { byCliente } = parseGamma(buf, XLSX, codici);
    let cnt = 0;
    clienti.forEach(c => { if (byCliente[c.cod]) { c.gamma_dettaglio = byCliente[c.cod].settori; cnt++; } });
    const prev = document.getElementById('preview-gamma');
    prev.style.display = 'block'; prev.style.background = 'var(--ok-l)'; prev.style.color = 'var(--ok)';
    prev.innerHTML = `✓ Gamma caricata per ${cnt} clienti`;
  } catch(e) {
    const prev = document.getElementById('preview-gamma');
    prev.style.display = 'block'; prev.style.background = 'var(--danger-l)'; prev.style.color = 'var(--danger)';
    prev.innerHTML = `✗ Errore gamma: ${e.message}`;
  }
}

function applicaDati(json) {
  DATA = json;
  clienti.length = 0;
  (DATA.clienti || []).forEach(c => {
    if (c.fatt2023 !== undefined && c.fatt2024 === undefined) c.fatt2024 = c.fatt2023;
    if (c.prog2025 !== undefined && c.prog2026 === undefined) c.prog2026 = c.prog2025;
    if (c.mesi2024 !== undefined && !c.mesi2025) c.mesi2025 = c.mesi2024;
    c._prev = c.fatt2024 || 0;
    c._prog = c.prog2026 || 0;
    c._varMese = (c.mese_corrente || {}).var_mese || null;
    clienti.push(c);
  });
  ANNO_PREV = DATA.anno_prev || 2025;
  ANNO_CORR = DATA.anno_corr || 2026;
  MESI_KEYS_PREV   = DATA.mesi_keys_prev  || MESI_SHORT.map(m => m + String(ANNO_PREV).slice(-2));
  MESI_KEYS_CORR   = DATA.mesi_keys_corr  || MESI_SHORT.map(m => m + String(ANNO_CORR).slice(-2));
  MESI_LABELS_PREV = DATA.mesi_labels_prev || DATA.mesi_labels_2024 || MESI_KEYS_PREV;
  MESI_LABELS_CORR = DATA.mesi_labels_corr || DATA.mesi_labels_2025 || MESI_KEYS_CORR;
  
  // Aggiorna header con data di aggiornamento
  const headerDate = document.getElementById('header-date');
  if (headerDate) {
    headerDate.textContent = `Aggiornato: ${DATA.aggiornato}`;
  }
  
  // Carica la nuova funzione di rendering per il layout moderno
  renderOverview();
  
  if (uploadState._gammaBuf) { applicaGammaBuf(uploadState._gammaBuf); uploadState._gammaBuf = null; }
  if (uploadState._cediMap) {
    clienti.forEach(c => {
      const key = String(c.cod).trim();
      if (uploadState._cediMap[key]) c.cedi_dettaglio = uploadState._cediMap[key];
    });
  }
}

function mostraErrore(msg) {
  console.error(msg);
  // Nel nuovo layout, mostra l'errore in console
  // Nel vecchio HTML cercava overview-loader che non esiste più
}

// ═══════════════════════════════════════════════════════
//  CLIENTI TAB (nuovo layout)
// ═══════════════════════════════════════════════════════
function renderClientiTab() {
  if (!clienti || clienti.length === 0) return;
  
  // Leggi i filtri
  const searchInput = document.getElementById('search-clienti');
  const sortSelect = document.getElementById('sort-clienti');
  
  let q = searchInput ? searchInput.value.toLowerCase() : '';
  let sortType = sortSelect ? sortSelect.value : 'nome';
  
  // Filtra
  let rows = clienti.filter(c => {
    if (q && !c.nome.toLowerCase().includes(q)) return false;
    return true;
  });
  
  // Ordina
  if (sortType === 'nome') {
    rows.sort((a, b) => a.nome.localeCompare(b.nome));
  } else if (sortType === 'fatturato') {
    rows.sort((a, b) => (b.fatt2025 || 0) - (a.fatt2025 || 0));
  } else if (sortType === 'data') {
    rows.sort((a, b) => (b.prog2026 || 0) - (a.prog2026 || 0));
  }
  
  // Renderizza tabella
  const tbody = document.getElementById('clienti-tbody');
  tbody.innerHTML = rows.map(c => {
    const fatt2025 = c.fatt2025 || 0;
    const prog2026 = c.prog2026 || 0;
    const ultimaVisita = c.mese_corrente ? 'Questo mese' : '—';
    
    return `<tr onclick="openModal('${c.cod.replace(/'/g,"\\'")}')">
      <td class="td-nome">${c.nome}<small>Cod. ${c.cod}</small></td>
      <td>${c.div || '—'}</td>
      <td>${c.gav || '—'}</td>
      <td style="text-align:right;">${fmt(fatt2025)}</td>
      <td style="text-align:right;">${fmt(prog2026)}</td>
      <td>${ultimaVisita}</td>
    </tr>`;
  }).join('');
  
  // Event listener per search e sort
  if (searchInput) {
    searchInput.oninput = () => renderClientiTab();
  }
  if (sortSelect) {
    sortSelect.onchange = () => renderClientiTab();
  }
}

function buildOverview() {
  if (!DATA || !clienti.length) return;
  const tot_prev = clienti.reduce((s,c) => s + c._prev, 0);
  const tot_prog = clienti.reduce((s,c) => s + c._prog, 0);
  const varProg  = tot_prev > 0 ? (tot_prog - tot_prev) / tot_prev * 100 : null;
  const attivi   = clienti.filter(c => c._prog > 0 || c._prev > 0).length;
  const totMesePrev = clienti.reduce((s,c) => s + ((c.mese_corrente||{}).fatt_mese_prev || 0), 0);
  const totMeseCorr = clienti.reduce((s,c) => {
    const mc = c.mese_corrente || {};
    const rolling = (mc.consegnato||0) + (mc.in_preparazione||0) + (mc.da_spedire||0);
    const cedi = (c.cedi_dettaglio||[]).reduce((t,g) => t + (g.fatt||0), 0);
    return s + rolling + cedi;
  }, 0);
  const varMeseAgg = totMesePrev > 0 ? (totMeseCorr - totMesePrev) / totMesePrev * 100 : null;
  const totProgPrev = clienti.reduce((s,c) => s + ((c.mese_corrente||{}).fatt_prog_prev || 0), 0);
  const totProgCorr = clienti.reduce((s,c) => s + ((c.mese_corrente||{}).fatt_prog_corr || c._prog || 0), 0);
  const varProgAgg  = totProgPrev > 0 ? (totProgCorr - totProgPrev) / totProgPrev * 100 : null;

  const kpis = [
    { label: `Fatturato ${ANNO_PREV}`,        val: fmt(tot_prev), sub: 'anno chiuso' },
    { label: `Progressivo ${ANNO_CORR}`,      val: fmt(tot_prog), delta: varProg, deltaLabel: `vs ${ANNO_PREV}` },
    { label: 'Clienti attivi',                val: attivi, sub: 'con dati fatturato' },
    { label: 'Media / cliente',               val: fmtK(tot_prog / (attivi||1)), sub: `progressivo ${ANNO_CORR}` },
    { label: 'Mese corrente area',            val: fmt(totMeseCorr), delta: varMeseAgg, deltaLabel: `vs stesso mese ${ANNO_PREV}` },
    { label: `Progressivo ${ANNO_CORR} area`, val: fmt(totProgCorr), delta: varProgAgg, deltaLabel: `vs prog. ${ANNO_PREV}` },
  ];
  document.getElementById('kpi-grid').innerHTML = kpis.map(k => `
    <div class="kpi">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-val">${k.val}</div>
      ${k.delta != null ? `<div class="kpi-delta ${deltaClass(k.delta)}">${fmtPct(k.delta)} ${k.deltaLabel}</div>` : ''}
      ${k.sub ? `<div class="kpi-sub">${k.sub}</div>` : ''}
    </div>`).join('');

  document.getElementById('section-mese-title').textContent = `Confronto mese corrente — ${DATA.aggiornato}`;
  const meseCols = [
    { label: 'Mese corrente area', prev: totMesePrev, corr: totMeseCorr },
    { label: `Progressivo Gen–mese`, prev: totProgPrev, corr: totProgCorr },
  ];
  document.getElementById('mese-compare-grid').innerHTML = meseCols.map(m => {
    const diff = m.corr - m.prev;
    const pct  = m.prev > 0 ? (diff / m.prev * 100) : null;
    return `<div class="mese-compare-card">
      <div class="mese-card-label">${m.label}</div>
      <div class="mese-row"><span class="mese-year">${ANNO_PREV}</span><span class="mese-val">${fmt(m.prev)}</span></div>
      <div class="mese-row"><span class="mese-year">${ANNO_CORR}</span><span class="mese-val">${fmt(m.corr)}</span></div>
      <hr class="mese-divider">
      <div class="mese-diff">
        <span class="mese-diff-label">Differenza</span>
        <span class="pill ${deltaClass(diff)}">${diff >= 0 ? '+' : ''}${fmt(diff)}  ${pct != null ? fmtPct(pct) : ''}</span>
      </div>
    </div>`;
  }).join('');

  document.getElementById('chart-mesi-title').textContent = `Andamento mensile ${ANNO_PREV} vs ${ANNO_CORR}`;
  const m_prev = MESI_KEYS_PREV.map(k => clienti.reduce((s,c) => s + ((c.mesi2025||{})[k]||0), 0));
  const m_corr = MESI_KEYS_CORR.map(k => clienti.reduce((s,c) => s + ((c.mesi2026||{})[k]||0), 0));
  const _cm = Chart.getChart('chart-mesi'); if (_cm) _cm.destroy();
  new Chart(document.getElementById('chart-mesi').getContext('2d'), {
    type: 'bar',
    data: { labels: MESI_SHORT, datasets: [
      { label: String(ANNO_PREV), data: m_prev, backgroundColor: 'rgba(107,114,128,.15)', borderColor: 'rgba(107,114,128,.6)', borderWidth: 1.5, borderRadius: 4 },
      { label: String(ANNO_CORR), data: m_corr, backgroundColor: 'rgba(26,108,255,.2)',  borderColor: 'rgba(26,108,255,.8)',  borderWidth: 1.5, borderRadius: 4 },
    ]}, options: chartOpts()
  });

  const gammaMap = {};
  clienti.forEach(c => { gammaMap[c.gamma] = (gammaMap[c.gamma]||0) + c._prog; });
  const gLabels = Object.keys(gammaMap).filter(k => gammaMap[k]>0).sort((a,b) => gammaMap[b]-gammaMap[a]);
  const PALETTE = ['#1a6cff','#0d9e6e','#d97706','#dc2626','#7c3aed','#0891b2','#be185d'];
  const _cg = Chart.getChart('chart-gamma'); if (_cg) _cg.destroy();
  new Chart(document.getElementById('chart-gamma').getContext('2d'), {
    type: 'doughnut',
    data: { labels: gLabels, datasets: [{ data: gLabels.map(k=>gammaMap[k]), backgroundColor: PALETTE, borderWidth: 2, borderColor: '#fff', hoverOffset: 6 }] },
    options: { responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { color: '#6b7280', font: { family: 'DM Mono', size: 11 }, padding: 12, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => ` ${fmt(ctx.raw)}` }, bodyFont: { family: 'DM Mono' } }
      }
    }
  });

  document.getElementById('chart-top15-title').textContent = `Top 15 clienti · ${ANNO_PREV} vs prog. ${ANNO_CORR}`;
  const top15 = [...clienti].sort((a,b) => b._prev - a._prev).slice(0,15);
  const _ct = Chart.getChart('chart-top15'); if (_ct) _ct.destroy();
  new Chart(document.getElementById('chart-top15').getContext('2d'), {
    type: 'bar',
    data: { labels: top15.map(c => c.nome.length > 30 ? c.nome.slice(0,30)+'…' : c.nome),
      datasets: [
        { label: String(ANNO_PREV), data: top15.map(c => c._prev), backgroundColor: 'rgba(107,114,128,.2)', borderColor: 'rgba(107,114,128,.7)', borderWidth: 1.5, borderRadius: 3 },
        { label: `Prog. ${ANNO_CORR}`, data: top15.map(c => c._prog), backgroundColor: 'rgba(26,108,255,.25)', borderColor: 'rgba(26,108,255,.8)', borderWidth: 1.5, borderRadius: 3 },
      ]
    }, options: { ...chartOpts(), indexAxis: 'y' }
  });
}

function chartOpts() {
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { labels: { color: '#6b7280', font: { family: 'DM Mono', size: 11 } } },
      tooltip: { callbacks: { label: ctx => ` ${fmt(ctx.raw)}` }, bodyFont: { family: 'DM Mono' }, titleFont: { family: 'Syne' } }
    },
    scales: {
      x: { ticks: { color: '#9ca3af', font: { family: 'DM Mono', size: 10 } }, grid: { color: '#e2e6ea' } },
      y: { ticks: { color: '#9ca3af', font: { family: 'DM Mono', size: 10 } }, grid: { color: '#e2e6ea' } }
    }
  };
}

// ═══════════════════════════════════════════════════════
//  TABELLA CLIENTI
// ═══════════════════════════════════════════════════════
function buildFilters() {
  const sel = document.getElementById('filter-gamma');
  while (sel.options.length > 1) sel.remove(1);
  [...new Set(clienti.map(c => c.gamma))].sort().forEach(g => {
    const o = document.createElement('option'); o.value = g; o.textContent = g; sel.appendChild(o);
  });
}

function renderTabella() {
  if (!clienti.length) return;
  const q  = document.getElementById('search-cliente').value.toLowerCase();
  const fg = document.getElementById('filter-gamma').value;
  const ft = document.getElementById('filter-trend').value;
  let rows = clienti.filter(c => {
    if (q  && !c.nome.toLowerCase().includes(q)) return false;
    if (fg && c.gamma !== fg) return false;
    if (ft === 'pos' && (c.var_anno == null || c.var_anno <= 0)) return false;
    if (ft === 'neg' && (c.var_anno == null || c.var_anno > 0))  return false;
    return true;
  });
  rows = rows.sort((a,b) => {
    const col = sortCol === 'fatt2024' ? '_prev' : sortCol === 'prog2026' ? '_prog' : sortCol;
    const va = a[col] ?? 0, vb = b[col] ?? 0;
    if (typeof va === 'string') return sortDir * va.localeCompare(vb);
    return sortDir * (vb - va);
  });
  document.querySelectorAll('thead th').forEach(th => {
    th.classList.remove('sort-asc','sort-desc');
    const sc = th.getAttribute('onclick') || '';
    if (sc.includes(`'${sortCol}'`)) th.classList.add(sortDir === 1 ? 'sort-asc' : 'sort-desc');
  });
  document.getElementById('tbl-body').innerHTML = rows.map(c => {
    const va = c.var_anno;
    const mc = c.mese_corrente || {};
    const vm = mc.var_mese;
    const vaStr = va == null ? '—' : `<span class="pill ${deltaClass(va)}">${va>=0?'▲':'▼'} ${Math.abs(Math.round(va))}%</span>`;
    const vmStr = vm == null ? '—' : `<span class="pill ${deltaClass(vm)}">${fmt(vm)}</span>`;
    return `<tr onclick="openModal('${c.cod.replace(/'/g,"\\'")}')">
      <td class="td-nome">${c.nome}<small>Cod. ${c.cod}</small></td>
      <td><span class="badge ${badgeClass(c.gamma)}">${c.gamma}</span></td>
      <td>${fmt(c._prev)}</td>
      <td>${fmt(c._prog)}</td>
      <td>${vaStr}</td>
      <td>${vmStr}</td>
      <td style="color:var(--accent);font-size:15px;">→</td>
    </tr>`;
  }).join('');
  document.getElementById('tbl-count').textContent = `${rows.length} clienti su ${clienti.length}`;
}

function sortBy(col) {
  if (sortCol === col) sortDir *= -1; else { sortCol = col; sortDir = -1; }
  renderTabella();
}

// ═══════════════════════════════════════════════════════
//  MODAL DETTAGLIO
// ═══════════════════════════════════════════════════════
function openModal(cod) {
  const c = clienti.find(x => x.cod === cod);
  if (!c) return;
  document.getElementById('m-nome').textContent = c.nome;
  document.getElementById('m-sub').textContent  = `Cod. ${c.cod} · Gamma: ${c.gamma}`;
  const mc = c.mese_corrente || {};
  const va = c.var_anno;
  document.getElementById('m-kpi').innerHTML = [
    { label: `Fatturato ${ANNO_PREV}`,   val: fmt(c._prev) },
    { label: `Progressivo ${ANNO_CORR}`, val: fmt(c._prog) },
    { label: 'Variazione anno',          val: fmtPct(va), cls: deltaClass(va) },
  ].map(k => `<div class="kpi">
    <div class="kpi-label">${k.label}</div>
    <div class="kpi-val" style="font-size:17px;">${k.cls ? `<span class="pill ${k.cls}">${k.val}</span>` : k.val}</div>
  </div>`).join('');

  const meseCorr = (mc.consegnato||0) + (mc.in_preparazione||0) + (mc.da_spedire||0);
  const mesePrev = mc.fatt_mese_prev || 0;
  const cediTot  = (c.cedi_dettaglio || []).reduce((s, g) => s + (g.fatt || 0), 0);
  const meseTot  = meseCorr + cediTot;
  const meseDiff = meseTot - mesePrev;
  const mesePct  = mesePrev > 0 ? meseDiff/mesePrev*100 : null;
  document.getElementById('m-mese-title').textContent = `Confronto mese corrente — ${DATA.aggiornato||''}`;
  document.getElementById('m-compare-block').innerHTML = `
    <div class="compare-card">
      <div class="compare-card-title">Fatturato mese</div>
      <div class="compare-row"><span class="compare-year">${ANNO_PREV}</span><span class="compare-val">${fmt(mesePrev)}</span></div>
      <div class="compare-row"><span class="compare-year">${ANNO_CORR} Rolling</span><span class="compare-val">${fmt(meseCorr)}</span></div>
      ${cediTot > 0 ? `<div class="compare-row"><span class="compare-year" style="color:var(--accent);">+ CEDI</span><span class="compare-val" style="color:var(--accent);">${fmt(cediTot)}</span></div>` : ''}
      <div class="compare-diff">
        <span class="compare-diff-label">Differenza vs ${ANNO_PREV}</span>
        <span class="pill ${deltaClass(meseDiff)}">${meseDiff>=0?'+':''}${fmt(meseDiff)} ${mesePct!=null?fmtPct(mesePct):''}</span>
      </div>
    </div>
    <div class="compare-card">
      <div class="compare-card-title">Dettaglio mese ${ANNO_CORR}</div>
      <div class="compare-row"><span class="compare-year">Consegnato</span><span class="compare-val">${fmt(mc.consegnato)}</span></div>
      <div class="compare-row"><span class="compare-year">In preparazione</span><span class="compare-val">${fmt(mc.in_preparazione)}</span></div>
      <div class="compare-row"><span class="compare-year">Da spedire</span><span class="compare-val">${fmt(mc.da_spedire)}</span></div>
      ${cediTot > 0 ? `<div class="compare-row" style="border-top:1px dashed var(--border);margin-top:4px;padding-top:6px;"><span class="compare-year" style="color:var(--accent);">CEDI</span><span class="compare-val" style="color:var(--accent);">${fmt(cediTot)}</span></div>` : ''}
    </div>`;

  const progPrev = mc.fatt_prog_prev || 0;
  const progCorr = mc.fatt_prog_corr || c._prog || 0;
  const progDiff = progCorr - progPrev;
  const progPct  = progPrev > 0 ? progDiff/progPrev*100 : null;
  document.getElementById('m-prog-title').textContent = `Confronto progressivo Gen–${DATA.aggiornato?.split('/')[1]||'mese'} ${ANNO_PREV} vs ${ANNO_CORR}`;
  document.getElementById('m-prog-block').innerHTML = `
    <div class="compare-card">
      <div class="compare-card-title">Progressivo anno</div>
      <div class="compare-row"><span class="compare-year">${ANNO_PREV}</span><span class="compare-val">${fmt(progPrev)}</span></div>
      <div class="compare-row"><span class="compare-year">${ANNO_CORR}</span><span class="compare-val">${fmt(progCorr)}</span></div>
      <div class="compare-diff"><span class="compare-diff-label">Differenza</span>
        <span class="pill ${deltaClass(progDiff)}">${progDiff>=0?'+':''}${fmt(progDiff)} ${progPct!=null?fmtPct(progPct):''}</span>
      </div>
    </div>
    <div class="compare-card">
      <div class="compare-card-title">Fatturato pieno ${ANNO_PREV} vs prog. ${ANNO_CORR}</div>
      <div class="compare-row"><span class="compare-year">Fatt. ${ANNO_PREV}</span><span class="compare-val">${fmt(c._prev)}</span></div>
      <div class="compare-row"><span class="compare-year">Prog. ${ANNO_CORR}</span><span class="compare-val">${fmt(c._prog)}</span></div>
      <div class="compare-diff"><span class="compare-diff-label">Variazione</span>
        <span class="pill ${deltaClass(va)}">${fmtPct(va)}</span>
      </div>
    </div>`;

  if (mChart) { mChart.destroy(); mChart = null; }
  const d_prev = MESI_KEYS_PREV.map(k => (c.mesi2025||{})[k]||0);
  const d_corr = MESI_KEYS_CORR.map(k => (c.mesi2026||{})[k]||null);
  mChart = new Chart(document.getElementById('m-chart').getContext('2d'), {
    type: 'bar',
    data: { labels: MESI_SHORT, datasets: [
      { label: String(ANNO_PREV), data: d_prev, backgroundColor: 'rgba(107,114,128,.18)', borderColor: 'rgba(107,114,128,.7)', borderWidth: 1.5, borderRadius: 3 },
      { label: String(ANNO_CORR), data: d_corr, backgroundColor: 'rgba(26,108,255,.22)',  borderColor: 'rgba(26,108,255,.8)',  borderWidth: 1.5, borderRadius: 3 },
    ]}, options: chartOpts()
  });

  const gd = c.gamma_dettaglio || {};
  const settori = Object.keys(gd);
  const tabsEl   = document.getElementById('m-gamma-tabs');
  const panelsEl = document.getElementById('m-gamma-panels');
  if (settori.length === 0) {
    tabsEl.innerHTML = '';
    panelsEl.innerHTML = '<div style="font-family:var(--mono);font-size:12px;color:var(--muted);padding:12px 0;">Nessun dato gamma disponibile.</div>';
  } else {
    tabsEl.innerHTML = settori.map((s,i) => `
      <button class="gamma-tab-btn${i===0?' active':''}" onclick="switchGammaTab(this,'gtab-${cod.replace(/\W/g,'')}-${i}')">${s}</button>`).join('');
    panelsEl.innerHTML = settori.map((s,i) => {
      const sg = gd[s];
      const pctImm = sg.perc_immancabili != null ? Math.round(sg.perc_immancabili * 100) : null;
      const pctStr = sg.perc_strategiche != null ? Math.round(sg.perc_strategiche * 100) : null;
      const colImm = pctImm == null ? '#9ca3af' : pctImm >= 80 ? '#0d9e6e' : pctImm >= 50 ? '#d97706' : '#dc2626';
      const prodotti = sg.prodotti || [];
      return `<div class="gamma-settore-panel${i===0?' active':''}" id="gtab-${cod.replace(/\W/g,'')}-${i}">
        <div class="gamma-header-row" style="margin-bottom:12px;">
          <div style="font-family:var(--mono);font-size:11px;color:var(--muted);">Fatt. 2024: <strong style="color:var(--text)">${fmt(sg.fatt2024)}</strong></div>
          <div class="gamma-pct-pills">
            ${pctImm != null ? `<span class="pct-pill imm">Imm. ${pctImm}%</span>` : ''}
            ${pctStr != null ? `<span class="pct-pill str">Str. ${pctStr}%</span>` : ''}
          </div>
        </div>
        ${pctImm != null ? `<div class="gauge-bar" style="margin-bottom:12px;height:4px;"><div class="gauge-fill" style="width:${Math.min(pctImm,100)}%;background:${colImm}"></div></div>` : ''}
        <div class="prod-list">
          ${prodotti.map(p => `
            <div class="prod-item ${p.valore > 0 ? 'has-value' : 'missing'}">
              <span class="prod-name">${p.nome}</span>
              <span class="prod-flag ${p.flag === 'Strategica-Immancabile' ? 'imm' : 'str'}">${p.flag === 'Strategica-Immancabile' ? 'IMM' : 'STR'}</span>
              <span class="prod-val ${p.valore > 0 ? '' : 'zero'}">${p.valore > 0 ? fmt(p.valore) : '—'}</span>
            </div>`).join('')}
        </div>
      </div>`;
    }).join('');
  }
  document.getElementById('modal').classList.add('open');
}

function switchGammaTab(btn, panelId) {
  btn.closest('.modal').querySelectorAll('.gamma-tab-btn').forEach(b => b.classList.remove('active'));
  btn.closest('.modal').querySelectorAll('.gamma-settore-panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById(panelId)?.classList.add('active');
}

function closeModalBg(e) { if (e.target.id === 'modal') closeModalDirect(); }
function closeModalDirect() {
  document.getElementById('modal').classList.remove('open');
  if (mChart) { mChart.destroy(); mChart = null; }
}

// ═══════════════════════════════════════════════════════
//  GAMMA PANEL
// ═══════════════════════════════════════════════════════
function renderGamma() {
  if (!DATA) return;
  const settori = DATA.settori_gamma || [...new Set(clienti.map(c=>c.gamma))].sort();
  const stats = {};
  clienti.forEach(c => {
    const gd = c.gamma_dettaglio || {};
    Object.keys(gd).forEach(s => {
      if (!stats[s]) stats[s] = { n:0, sumI:0, sumS:0, fatt:0, cliList:[] };
      stats[s].n++;
      if (gd[s].perc_immancabili != null) stats[s].sumI += gd[s].perc_immancabili;
      if (gd[s].perc_strategiche != null) stats[s].sumS += gd[s].perc_strategiche;
      stats[s].fatt += gd[s].fatt2024 || 0;
      stats[s].cliList.push({ nome: c.nome, val: gd[s].fatt2024||0 });
    });
    if (Object.keys(gd).length === 0) {
      if (!stats[c.gamma]) stats[c.gamma] = { n:0, sumI:0, sumS:0, fatt:0, cliList:[] };
      stats[c.gamma].n++;
      stats[c.gamma].fatt += c._prog;
      stats[c.gamma].cliList.push({ nome: c.nome, val: c._prog });
    }
  });
  document.getElementById('gamma-kpi-grid').innerHTML = [
    { label: 'Clienti totali',  val: clienti.length },
    { label: 'Settori/gamme',   val: Object.keys(stats).length },
    { label: 'Con dati gamma',  val: clienti.filter(c=>Object.keys(c.gamma_dettaglio||{}).length>0).length },
    { label: `Fatt. top gamma`, val: fmt(Math.max(...Object.values(stats).map(s=>s.fatt))) },
  ].map(k => `<div class="kpi"><div class="kpi-label">${k.label}</div><div class="kpi-val" style="font-size:20px;">${k.val}</div></div>`).join('');
  const allSettori = [...new Set([...settori, ...Object.keys(stats)])];
  document.getElementById('gamma-overview').innerHTML = allSettori.map(s => {
    const st = stats[s];
    if (!st) return '';
    const avgI = st.n > 0 && st.sumI > 0 ? Math.round(st.sumI/st.n*100) : null;
    const avgS = st.n > 0 && st.sumS > 0 ? Math.round(st.sumS/st.n*100) : null;
    const topCli = [...st.cliList].sort((a,b)=>b.val-a.val).slice(0,5);
    const colI = avgI == null ? '#9ca3af' : avgI >= 80 ? '#0d9e6e' : avgI >= 50 ? '#d97706' : '#dc2626';
    return `<div class="gamma-sector-card">
      <div class="gs-title">${s} <span class="gs-badge">${st.n} clienti</span></div>
      ${avgI!=null?`<div class="gs-stat"><span class="gs-stat-label">Imm. media</span><span class="pct-pill imm">${avgI}%</span></div>
      <div class="gauge-bar"><div class="gauge-fill" style="width:${Math.min(avgI,100)}%;background:${colI}"></div></div>`:''}
      ${avgS!=null?`<div class="gs-stat"><span class="gs-stat-label">Str. media</span><span class="pct-pill str">${avgS}%</span></div>`:''}
      <div class="gs-stat-label" style="margin:10px 0 6px;">Top clienti</div>
      ${topCli.map(c=>`<div class="gs-row"><span class="gs-name" title="${c.nome}">${c.nome}</span><span class="gs-pct">${fmt(c.val)}</span></div>`).join('')}
    </div>`;
  }).join('');
}

// ═══════════════════════════════════════════════════════
//  UPLOAD / GITHUB
// ═══════════════════════════════════════════════════════
const uploadState = { rolling: null, gamma: null, cedi: null };

function loadConfig() {
  const repoEl = document.getElementById('gh-repo');
  const tokenEl = document.getElementById('gh-token');
  if (repoEl) repoEl.value  = localStorage.getItem('gh_repo')  || '';
  if (tokenEl) tokenEl.value = localStorage.getItem('gh_token') || '';
}
function saveConfig() {
  const repoEl = document.getElementById('gh-repo');
  const tokenEl = document.getElementById('gh-token');
  if (repoEl) localStorage.setItem('gh_repo',  repoEl.value.trim());
  if (tokenEl) localStorage.setItem('gh_token', tokenEl.value.trim());
}
function handleDrop(e, tipo) {
  e.preventDefault();
  document.getElementById('upload-zone-'+tipo).classList.remove('drag');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0], tipo);
}

async function handleFile(f, tipo) {
  uploadState[tipo] = f;
  const nameEl   = document.getElementById(tipo+'-name');
  const statusEl = document.getElementById('status-'+tipo);
  nameEl.textContent = f.name; nameEl.style.color = 'var(--accent)';
  document.getElementById('upload-zone-'+tipo).style.borderColor = 'var(--accent)';
  statusEl.textContent = `✓ ${(f.size/1024).toFixed(0)} KB — file pronto`;
  statusEl.style.color = 'var(--ok)';

  if (tipo === 'rolling' && typeof parseRolling === 'function') {
    try {
      const json = parseRolling(await f.arrayBuffer(), XLSX);
      applicaDati(json);
      const repo = (localStorage.getItem('gh_repo') || 'loriscuba/fischer-preventivi').trim();
      if (repo) {
        if (!uploadState.gamma) {
          try {
            const rG = await fetch(`https://raw.githubusercontent.com/${repo}/data/data/gamma.xlsx?t=${Date.now()}`);
            if (rG.ok && typeof parseGamma === 'function') {
              const bufG = await rG.arrayBuffer();
              applicaGammaBuf(bufG);
              const prevG = document.getElementById('preview-gamma');
              prevG.style.display = 'block'; prevG.style.background = 'var(--ok-l)'; prevG.style.color = 'var(--ok)';
              prevG.innerHTML = `✓ Gamma caricata da GitHub`;
            }
          } catch(_) {}
        }
        if (!uploadState.cedi) {
          try {
            const rC = await fetch(`https://raw.githubusercontent.com/${repo}/data/data/cedi.xlsx?t=${Date.now()}`);
            if (rC.ok) {
              const bufC = await rC.arrayBuffer();
              const wb = XLSX.read(bufC, { type: 'array' });
              const ws = wb.Sheets[wb.SheetNames[0]];
              const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
              const cediMap = {};
              for (const r of rows) {
                const codVend = Number(r[7]); const codCli = String(r[9]).trim().split('.')[0];
                const nomeCedi = String(r[10]).trim(); const fatt = Number(r[11]) || 0;
                if (codVend !== 400542 || !codCli || isNaN(Number(codCli))) continue;
                if (!cediMap[codCli]) cediMap[codCli] = [];
                cediMap[codCli].push({ nome: nomeCedi, fatt });
              }
              let cnt = 0;
              clienti.forEach(c => {
                const codCorto = String(c.cod).trim().split('/').pop();
                if (cediMap[codCorto]) { c.cedi_dettaglio = cediMap[codCorto]; cnt++; }
              });
              uploadState._cediMap = cediMap;
              const prevC = document.getElementById('preview-cedi');
              prevC.style.display = 'block'; prevC.style.background = 'var(--ok-l)'; prevC.style.color = 'var(--ok)';
              prevC.innerHTML = `✓ CEDI caricati per ${cnt} clienti`;
            }
          } catch(_) {}
        }
      }
      const prev = document.getElementById('preview-rolling');
      prev.style.display = 'block';
      const tPrev = json.clienti.reduce((s,c)=>s+(c[`fatt${json.anno_prev}`]||0),0);
      const tCorr = json.clienti.reduce((s,c)=>s+(c[`prog${json.anno_corr}`]||c[`fatt${json.anno_corr}`]||0),0);
      prev.innerHTML = [`✓ Data: ${json.aggiornato}`,`✓ Clienti: ${json.clienti.length}`,`✓ Fatt. ${json.anno_prev}: ${fmt(tPrev)}`,`✓ Prog. ${json.anno_corr}: ${fmt(tCorr)}`,`→ Cruscotto aggiornato`].join('<br>');
    } catch(e) {
      const prev = document.getElementById('preview-rolling');
      prev.style.display = 'block'; prev.style.background = 'var(--danger-l)'; prev.style.color = 'var(--danger)';
      prev.innerHTML = `✗ Errore parsing: ${e.message}`;
    }
  }

  if (tipo === 'gamma' && typeof parseGamma === 'function') {
    try {
      const buf = await f.arrayBuffer();
      if (clienti.length === 0) { uploadState._gammaBuf = buf; document.getElementById('preview-gamma').style.display='block'; document.getElementById('preview-gamma').innerHTML='⏳ Gamma pronta — verrà applicata al rolling'; }
      else applicaGammaBuf(buf);
    } catch(e) { const p=document.getElementById('preview-gamma'); p.style.display='block'; p.style.background='var(--danger-l)'; p.style.color='var(--danger)'; p.innerHTML=`✗ ${e.message}`; }
  }

  if (tipo === 'cedi') {
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const cediMap = {};
      for (const r of rows) {
        const codVend=Number(r[7]); const codCli=String(r[9]).trim().split('.')[0];
        const nomeCedi=String(r[10]).trim(); const fatt=Number(r[11])||0;
        if (codVend!==400542||!codCli||codCli===''||isNaN(Number(codCli))) continue;
        if (!cediMap[codCli]) cediMap[codCli]=[];
        cediMap[codCli].push({nome:nomeCedi,fatt});
      }
      let cnt=0;
      clienti.forEach(c => { const codCorto=String(c.cod).trim().split('/').pop(); if(cediMap[codCorto]){c.cedi_dettaglio=cediMap[codCorto];cnt++;} });
      uploadState._cediMap=cediMap;
      const totFatt=Object.values(cediMap).flat().reduce((s,x)=>s+x.fatt,0);
      const prev=document.getElementById('preview-cedi');
      prev.style.display='block';
      prev.innerHTML=[`✓ CEDI caricati per ${cnt} clienti`,`✓ Totale: ${fmt(totFatt)}`].join('<br>');
    } catch(e) { const p=document.getElementById('preview-cedi'); p.style.display='block'; p.style.background='var(--danger-l)'; p.style.color='var(--danger)'; p.innerHTML=`✗ ${e.message}`; }
  }

  if (uploadState.rolling) document.getElementById('btn-pubblica').style.display = 'block';
}

function log(msg, color) {
  const el = document.getElementById('upload-log');
  el.style.display = 'block';
  el.innerHTML += `<div style="color:${color||'var(--text2)'}">${msg}</div>`;
  el.scrollTop = el.scrollHeight;
}

async function fileToBase64(file) {
  return new Promise((res, rej) => { const r=new FileReader(); r.onload=()=>res(r.result.split(',')[1]); r.onerror=rej; r.readAsDataURL(file); });
}

async function getSHA(path, repo, token, branch='data') {
  const r = await fetch(`https://api.github.com/repos/${repo}/contents/${path}?ref=${branch}`, { headers:{Authorization:`Bearer ${token}`,Accept:'application/vnd.github+json'} });
  if (r.status===404) return null;
  if (!r.ok) throw new Error(`getSHA ${path}: ${r.status}`);
  return (await r.json()).sha || null;
}

async function commitFile(path, b64, msg, repo, token, sha, branch='data') {
  const body={message:msg,content:b64,branch}; if(sha) body.sha=sha;
  const r=await fetch(`https://api.github.com/repos/${repo}/contents/${path}`,{method:'PUT',headers:{Authorization:`Bearer ${token}`,Accept:'application/vnd.github+json','Content-Type':'application/json'},body:JSON.stringify(body)});
  if (!r.ok) { const e=await r.json().catch(()=>({})); throw new Error(`${path}: ${r.status} ${e.message||''}`); }
}

async function pubblicaSuGitHub() {
  const repo  = (localStorage.getItem('gh_repo')||'').trim();
  const token = (localStorage.getItem('gh_token')||'').trim();
  if (!repo||!token) { alert('Inserisci repo e token.'); return; }
  if (!uploadState.rolling) { alert('Carica almeno il file rolling.'); return; }
  const btn = document.getElementById('btn-pubblica');
  btn.disabled=true; btn.textContent='⏳  Invio…';
  document.getElementById('upload-log').innerHTML='';
  try {
    log('🔐  Verifica token…','var(--muted)');
    const me=await fetch('https://api.github.com/user',{headers:{Authorization:`Bearer ${token}`,Accept:'application/vnd.github+json'}});
    if (!me.ok) throw new Error(`Token non valido (${me.status})`);
    log(`✓  @${(await me.json()).login}`,'var(--ok)');
    const b64R=await fileToBase64(uploadState.rolling);
    const shaR=await getSHA('data/rolling.xlsx',repo,token);
    log('⬆️  Upload rolling.xlsx…','var(--muted)');
    await commitFile('data/rolling.xlsx',b64R,`📊 rolling [${new Date().toLocaleDateString('it-IT')}]`,repo,token,shaR);
    log('✓  rolling.xlsx','var(--ok)');
    if (uploadState.gamma) {
      const b64G=await fileToBase64(uploadState.gamma); const shaG=await getSHA('data/gamma.xlsx',repo,token);
      await commitFile('data/gamma.xlsx',b64G,`🗂️ gamma [${new Date().toLocaleDateString('it-IT')}]`,repo,token,shaG);
      log('✓  gamma.xlsx','var(--ok)');
    }
    if (uploadState.cedi) {
      const b64C=await fileToBase64(uploadState.cedi); const shaC=await getSHA('data/cedi.xlsx',repo,token);
      await commitFile('data/cedi.xlsx',b64C,`🏢 cedi [${new Date().toLocaleDateString('it-IT')}]`,repo,token,shaC);
      log('✓  cedi.xlsx','var(--ok)');
    }
    log('🚀  GitHub Actions avviato — ~2 min.','var(--accent)');
    log(`🔗 <a href="https://github.com/${repo}/actions" target="_blank" style="color:var(--accent);">Monitora →</a>`,'');
    btn.textContent='✅  Pubblicato!'; btn.style.background='var(--ok)';
  } catch(err) {
    log(`✗  ${err.message}`,'var(--danger)');
    btn.disabled=false; btn.textContent='🚀  PUBBLICA SU GITHUB';
  }
}

// ═══════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════
loadConfig();
caricaDati();
