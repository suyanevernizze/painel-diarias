// ── STATE ──
let allData = [], filteredData = [], currentPage = 1, sortCol = 'valor', sortDir = 'desc';
const PS = 50;
let CH = {};
const COLORS = ['#2e9e44','#f0b429','#378ADD','#D85A30','#8b7cf6','#1D9E75','#e8843a','#d4537e','#ef4444','#4caf60'];

// ── FORMATTERS ──
// Formata valor em BRL com centavos: R$ 1.234,56
const fR = v => (Number(v) || 0).toLocaleString('pt-BR', {
  style: 'currency', currency: 'BRL',
  minimumFractionDigits: 2, maximumFractionDigits: 2
});

// Formata valor em BRL completo (igual fR, sem abreviação K/M)
const fS = v => (Number(v) || 0).toLocaleString('pt-BR', {
  style: 'currency', currency: 'BRL',
  minimumFractionDigits: 2, maximumFractionDigits: 2
});

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400 * 1000));
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

// ── DRAG & DROP ──
const dz = document.getElementById('dropZone');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
dz.addEventListener('drop', e => {
  e.preventDefault();
  dz.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) processFile(f);
});
function loadFile(i) { if (i.files[0]) processFile(i.files[0]); }

// ── LOAD FILE ──
function processFile(file) {
  document.getElementById('fileInfo').textContent = file.name;
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
    const sn = wb.SheetNames;
    const ms = sn.find(n => /diária|diaria|base/i.test(n)) || sn[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[ms], { defval: '' });
    const data = [];
    rows.forEach(row => {
      const keys = Object.keys(row);
      const get = (...ns) => {
        for (const n of ns) {
          const k = keys.find(k =>
            k.toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,'')
             .includes(n.toLowerCase().replace(/\s+/g,'').replace(/[^a-z0-9]/g,''))
          );
          if (k !== undefined && row[k] !== '') return row[k];
        }
        return '';
      };
      const valor     = parseFloat(String(get('TOTALR','TOTAL R','valor','total')).replace(',', '.')) || 0;
      const horas     = parseFloat(String(get('Totaldehoras','horas','Total de Horas','tempoespera')).replace(',', '.')) || 0;
      const filial    = String(get('Filial')).trim();
      const cliente   = String(get('ClienteKlabin','cliente','Cliente')).trim();
      const contrato  = String(get('Contrato')).trim().toUpperCase();
      const ocorrencia = String(get('TipodeOcorrncia','TipodeOcorrencia','ocorrencia','Tipo','Ocorrncia')).trim().toUpperCase();
      const status    = String(get('STATUSDASTE','STATUS','status','Status')).trim().toUpperCase();
      const placa     = String(get('Placa','placa','PLACA','veículo','veiculo')).trim().toUpperCase();
      const cte       = String(get('CTe','cte','CTE','CT-e','conhecimento')).trim();
      const nf        = String(get('NF','nf','NotaFiscal','nota','NF-e','NFe')).trim();
      const dataT     = get('DataTransporte','Data','data','DataTransp');
      if (filial || cliente || valor)
        data.push({ filial, cliente, contrato, ocorrencia, status, placa, cte, nf, horas, valor, data: dataT });
    });
    if (!data.length) {
      alert('Não foi possível identificar os dados. Verifique se a planilha contém colunas como: Filial, Cliente Klabin, Contrato, Tipo de Ocorrência, Total R$, STATUS DASTE.');
      return;
    }
    allData = data;
    document.getElementById('dropZone').classList.add('hidden');
    document.getElementById('dashboard').classList.add('visible');
    populateFilters();
    applyFilters();
  };
  reader.readAsBinaryString(file);
}

// ── FILTERS ──
function populateFilters() {
  const fs = { selFilial: 'filial', selContrato: 'contrato', selOcorrencia: 'ocorrencia', selStatus: 'status', selCliente: 'cliente' };
  Object.entries(fs).forEach(([id, field]) => {
    const sel = document.getElementById(id), cur = sel.value;
    while (sel.options.length > 1) sel.remove(1);
    [...new Set(allData.map(r => r[field]).filter(Boolean))].sort().forEach(v => {
      const o = document.createElement('option');
      o.value = v; o.text = v; sel.appendChild(o);
    });
    sel.value = cur;
  });
}

function applyQuickPeriod() {
  const val = document.getElementById('selPeriodo').value;
  if (!val) return;
  const now = new Date();
  let s, e = new Date(now);
  if (val === 'thismonth')       { s = new Date(now.getFullYear(), now.getMonth(), 1);     e = new Date(now.getFullYear(), now.getMonth() + 1, 0); }
  else if (val === 'lastmonth')  { s = new Date(now.getFullYear(), now.getMonth() - 1, 1); e = new Date(now.getFullYear(), now.getMonth(), 0); }
  else if (val === 'thisyear')   { s = new Date(now.getFullYear(), 0, 1);                  e = new Date(now.getFullYear(), 11, 31); }
  else { s = new Date(now); s.setDate(s.getDate() - parseInt(val)); }
  const fmt = d => d.toISOString().slice(0, 10);
  document.getElementById('selDtI').value = fmt(s);
  document.getElementById('selDtF').value = fmt(e);
  applyFilters();
}

function applyFilters() {
  const f     = document.getElementById('selFilial').value;
  const c     = document.getElementById('selContrato').value;
  const o     = document.getElementById('selOcorrencia').value;
  const s     = document.getElementById('selStatus').value;
  const cl    = document.getElementById('selCliente').value;
  const q     = document.getElementById('tableSearch').value.toLowerCase();
  const placa = document.getElementById('selPlaca').value.toUpperCase().trim();
  const cte   = document.getElementById('selCte').value.trim();
  const nf    = document.getElementById('selNf').value.trim();
  const di    = document.getElementById('selDtI').value;
  const df    = document.getElementById('selDtF').value;
  const dtI   = di ? new Date(di + 'T00:00:00') : null;
  const dtF   = df ? new Date(df + 'T23:59:59') : null;

  filteredData = allData.filter(r => {
    if (f     && r.filial     !== f)  return false;
    if (c     && r.contrato   !== c)  return false;
    if (o     && r.ocorrencia !== o)  return false;
    if (s     && r.status     !== s)  return false;
    if (cl    && r.cliente    !== cl) return false;
    if (placa && !r.placa.includes(placa)) return false;
    if (cte   && !r.cte.includes(cte))     return false;
    if (nf    && !r.nf.includes(nf))       return false;
    if (q) {
      const hay = `${r.filial} ${r.cliente} ${r.contrato} ${r.placa} ${r.cte} ${r.nf}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    if (dtI || dtF) {
      const rd = toDate(r.data);
      if (!rd) return false;
      if (dtI && rd < dtI) return false;
      if (dtF && rd > dtF) return false;
    }
    return true;
  });
  sortData(); currentPage = 1; renderMetrics(); renderCharts(); renderTable();
}

function resetFilters() {
  ['selFilial','selContrato','selOcorrencia','selStatus','selCliente','selPeriodo'].forEach(id => document.getElementById(id).value = '');
  ['selPlaca','selCte','selNf','tableSearch'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('selDtI').value = '';
  document.getElementById('selDtF').value = '';
  applyFilters();
}

function sortData() {
  filteredData.sort((a, b) => {
    let va = a[sortCol] || '', vb = b[sortCol] || '';
    if (sortCol === 'valor' || sortCol === 'horas') { va = Number(va) || 0; vb = Number(vb) || 0; }
    else { va = String(va).toLowerCase(); vb = String(vb).toLowerCase(); }
    return va < vb ? (sortDir === 'asc' ? -1 : 1) : va > vb ? (sortDir === 'asc' ? 1 : -1) : 0;
  });
}

function srt(col) {
  if (sortCol === col) sortDir = sortDir === 'asc' ? 'desc' : 'asc';
  else { sortCol = col; sortDir = (col === 'valor' || col === 'horas') ? 'desc' : 'asc'; }
  document.querySelectorAll('thead th').forEach(th => th.classList.remove('sa','sd'));
  const cols = ['filial','cliente','placa','cte','nf','contrato','ocorrencia','horas','valor','status'];
  const idx = cols.indexOf(col);
  if (idx >= 0) {
    const ths = document.querySelectorAll('thead th');
    if (ths[idx]) ths[idx].classList.add(sortDir === 'asc' ? 'sa' : 'sd');
  }
  sortData(); renderTable();
}

// ── METRICS ──
function renderMetrics() {
  const d = filteredData;
  if (!d.length) {
    document.getElementById('metrics').innerHTML = '<div class="metric"><div class="m-label">Sem dados</div></div>';
    return;
  }
  const vals  = d.map(r => r.valor).filter(v => v > 0);
  const hs    = d.map(r => r.horas).filter(h => h > 0);
  const total = vals.reduce((s, v) => s + v, 0);
  const fat   = d.filter(r => r.status.includes('FATURAD')).reduce((s, r) => s + r.valor, 0);
  const ag    = d.filter(r => r.status.includes('APROV')).reduce((s, r) => s + r.valor, 0);
  const agN   = d.filter(r => r.status.includes('APROV')).length;
  const minV  = vals.length ? Math.min(...vals) : 0;
  const maxV  = vals.length ? Math.max(...vals) : 0;
  const minH  = hs.length   ? Math.min(...hs)   : 0;
  const maxH  = hs.length   ? Math.max(...hs)   : 0;
  const avgH  = hs.length   ? hs.reduce((s, h) => s + h, 0) / hs.length : 0;
  const avgV  = d.length    ? total / d.length  : 0;
  const rMinV = d.find(r => r.valor === minV), rMaxV = d.find(r => r.valor === maxV);
  const rMinH = d.find(r => r.horas === minH), rMaxH = d.find(r => r.horas === maxH);

  document.getElementById('metrics').innerHTML = `
    <div class="metric"><span class="m-icon">💰</span><div class="m-label">Total geral</div><div class="m-val">${fS(total)}</div><div class="m-sub">${d.length.toLocaleString('pt-BR')} registros</div></div>
    <div class="metric"><span class="m-icon">✅</span><div class="m-label">Faturado</div><div class="m-val">${fS(fat)}</div><div class="m-sub">${total ? Math.round(fat / total * 100) : 0}% do total</div></div>
    <div class="metric c-amber"><span class="m-icon">⏳</span><div class="m-label">Ag. aprovação</div><div class="m-val">${fS(ag)}</div><div class="m-sub">${agN} pendentes</div></div>
    <div class="metric c-teal"><span class="m-icon">📊</span><div class="m-label">Ticket médio</div><div class="m-val">${fS(avgV)}</div><div class="m-sub">por ocorrência</div></div>
    <div class="metric c-blue"><span class="m-icon">🔽</span><div class="m-label">Menor diária</div><div class="m-val">${fS(minV)}</div><div class="m-sub">${rMinV ? rMinV.cliente || rMinV.filial : '—'}</div></div>
    <div class="metric c-gold"><span class="m-icon">🔼</span><div class="m-label">Maior diária</div><div class="m-val">${fS(maxV)}</div><div class="m-sub">${rMaxV ? rMaxV.cliente || rMaxV.filial : '—'}</div></div>
    <div class="metric c-blue"><span class="m-icon">🕐</span><div class="m-label">Menor espera</div><div class="m-val">${minH.toFixed(1)}h</div><div class="m-sub">${rMinH ? rMinH.cliente || rMinH.filial : '—'}</div></div>
    <div class="metric c-gold"><span class="m-icon">🕙</span><div class="m-label">Maior espera</div><div class="m-val">${maxH.toFixed(1)}h</div><div class="m-sub">${rMaxH ? rMaxH.cliente || rMaxH.filial : '—'}</div></div>
    <div class="metric c-purple"><span class="m-icon">⏱️</span><div class="m-label">Média de espera</div><div class="m-val">${avgH.toFixed(1)}h</div><div class="m-sub">por ocorrência</div></div>
  `;
}

// ── MODAL ──
function openModal(title, rows, extraMetrics) {
  document.getElementById('modalTitle').textContent = title;
  const mm = document.getElementById('modalMetrics');
  if (extraMetrics && extraMetrics.length) {
    mm.innerHTML = extraMetrics.map(m =>
      `<div class="modal-metric"><div class="mm-label">${m.label}</div><div class="mm-val">${m.val}</div>${m.sub ? `<div class="mm-sub">${m.sub}</div>` : ''}</div>`
    ).join('');
    mm.style.display = 'grid';
  } else {
    mm.innerHTML = ''; mm.style.display = 'none';
  }
  if (!rows || !rows.length) {
    document.getElementById('modalThead').innerHTML = '';
    document.getElementById('modalTbody').innerHTML = '<tr><td colspan="9" style="text-align:center;padding:2rem;color:var(--tx3)">Sem registros</td></tr>';
    document.getElementById('modalOverlay').classList.add('open');
    return;
  }
  const cols = ['Filial','Cliente','Placa','CTe','NF','Ocorrência','Horas','Total R$','Status'];
  const keys = ['filial','cliente','placa','cte','nf','ocorrencia','horas','valor','status'];
  document.getElementById('modalThead').innerHTML = '<tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr>';
  document.getElementById('modalTbody').innerHTML = rows.slice(0, 200).map(r => {
    const s = r.status || '';
    const badge = s.includes('FATURAD') ? 'bf' : s.includes('APROV') ? 'ba' : 'br';
    return '<tr>' + keys.map((k, i) =>
      i === 8 ? `<td><span class="badge ${badge}">${s || '—'}</span></td>` :
      i === 6 ? `<td class="vc">${(r.horas || 0).toFixed(1)}h</td>` :
      i === 7 ? `<td class="vc">${fR(r.valor)}</td>` :
               `<td>${r[k] || '—'}</td>`
    ).join('') + '</tr>';
  }).join('');
  document.getElementById('modalOverlay').classList.add('open');
}

function closeModal(e) {
  if (!e || e.target === document.getElementById('modalOverlay'))
    document.getElementById('modalOverlay').classList.remove('open');
}

// ── CHARTS ──
function dc(id) { if (CH[id]) { CH[id].destroy(); delete CH[id]; } }

function mkDonut(id, labels, vals, colors, field) {
  dc(id);
  CH[id] = new Chart(document.getElementById(id), {
    type: 'doughnut',
    data: { labels, datasets: [{ data: vals, backgroundColor: colors, borderWidth: 2, borderColor: 'transparent', hoverOffset: 8 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '62%',
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fR(ctx.raw) + ' (' + Math.round(ctx.raw / (vals.reduce((a, b) => a + b, 1)) * 100) + '%)' } }
      },
      onClick: (evt, el) => {
        if (!el.length) return;
        const lbl  = labels[el[0].index];
        const rows = filteredData.filter(r => (r[field] || '').trim() === lbl.trim());
        const tot  = rows.reduce((s, r) => s + r.valor, 0);
        const avgH = rows.length ? rows.reduce((s, r) => s + r.horas, 0) / rows.length : 0;
        openModal(`${field.charAt(0).toUpperCase() + field.slice(1)}: ${lbl}`,
          [...rows].sort((a, b) => b.valor - a.valor),
          [
            { label: 'Total R$',     val: fS(tot) },
            { label: 'Registros',    val: rows.length },
            { label: 'Média espera', val: avgH.toFixed(1) + 'h' },
            { label: 'Ticket médio', val: fS(rows.length ? tot / rows.length : 0) }
          ]
        );
      }
    }
  });
}

function renderCharts() {
  const d = filteredData;

  // ── Linha do tempo ──
  const bM = {}, bQ = {};
  d.forEach(r => {
    const dt = toDate(r.data); if (!dt) return;
    const k = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
    bM[k] = (bM[k] || 0) + r.valor;
    bQ[k] = (bQ[k] || 0) + 1;
  });
  const mK = Object.keys(bM).sort().slice(-24);
  const mV = mK.map(k => bM[k]);
  const mQ = mK.map(k => bQ[k] || 0);
  const mL = mK.map(k => {
    const [y, m] = k.split('-');
    return ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'][+m - 1] + '/' + y.slice(2);
  });
  dc('cMensal');
  CH['cMensal'] = new Chart(document.getElementById('cMensal'), {
    data: {
      labels: mL,
      datasets: [
        { type: 'bar',  label: 'Total R$',     data: mV, backgroundColor: 'rgba(30,125,50,.75)', borderColor: '#2e9e44', borderWidth: 1, borderRadius: 6, yAxisID: 'y' },
        { type: 'line', label: 'Ocorrências',  data: mQ, borderColor: '#f0b429', backgroundColor: 'rgba(240,180,41,.1)', pointBackgroundColor: '#f0b429', pointRadius: 5, pointHoverRadius: 7, tension: .35, fill: true, yAxisID: 'y2' }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: true, labels: { color: '#7aaa86', boxWidth: 12, font: { size: 11 } } },
        tooltip: { callbacks: { label: ctx => ctx.datasetIndex === 0 ? fR(ctx.raw) : ctx.raw + ' ocorr.' } }
      },
      scales: {
        x:  { ticks: { color: '#4d7a58', font: { size: 10 }, maxRotation: 45 }, grid: { color: 'rgba(45,82,56,.4)' } },
        y:  { ticks: { callback: v => fS(v), color: '#4d7a58', font: { size: 9 } }, grid: { color: 'rgba(45,82,56,.4)' } },
        y2: { position: 'right', ticks: { color: '#f0b429', font: { size: 9 } }, grid: { display: false } }
      },
      onClick: (evt, el) => {
        if (!el.length) return;
        const lbl  = mL[el[0].index], key = mK[el[0].index];
        const rows = d.filter(r => {
          const dt = toDate(r.data); if (!dt) return false;
          const k  = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
          return k === key;
        });
        const tot  = rows.reduce((s, r) => s + r.valor, 0);
        const avgH = rows.length ? rows.reduce((s, r) => s + r.horas, 0) / rows.length : 0;
        openModal('Mês: ' + lbl, [...rows].sort((a, b) => b.valor - a.valor), [
          { label: 'Total R$',      val: fS(tot) },
          { label: 'Ocorrências',   val: rows.length },
          { label: 'Média espera',  val: avgH.toFixed(1) + 'h' },
          { label: 'Ticket médio',  val: fS(rows.length ? tot / rows.length : 0) }
        ]);
      }
    }
  });

  // ── Top clientes ──
  const bC = {};
  d.forEach(r => { if (r.cliente) bC[r.cliente.trim()] = (bC[r.cliente.trim()] || 0) + r.valor; });
  const cK = Object.keys(bC).sort((a, b) => bC[b] - bC[a]).slice(0, 10);
  const cV = cK.map(k => bC[k]);
  dc('cCliente');
  document.getElementById('cCliente').parentElement.style.height = Math.max(cK.length * 30 + 60, 180) + 'px';
  CH['cCliente'] = new Chart(document.getElementById('cCliente'), {
    type: 'bar',
    data: {
      labels: cK,
      datasets: [{ data: cV, backgroundColor: COLORS.slice(0, cK.length).map(c => c + 'bb'), borderColor: COLORS.slice(0, cK.length), borderWidth: 1, borderRadius: 4, hoverBorderWidth: 2 }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fR(ctx.raw) } }
      },
      scales: {
        x: { ticks: { callback: v => fS(v), color: '#4d7a58', font: { size: 9 } }, grid: { color: 'rgba(45,82,56,.4)' } },
        y: { ticks: { color: '#7aaa86', font: { size: 10 } } }
      },
      onClick: (evt, el) => {
        if (!el.length) return;
        const cli  = cK[el[0].index];
        const rows = filteredData.filter(r => r.cliente.trim() === cli);
        const tot  = rows.reduce((s, r) => s + r.valor, 0);
        const avgH = rows.length ? rows.reduce((s, r) => s + r.horas, 0) / rows.length : 0;
        openModal('Cliente: ' + cli, [...rows].sort((a, b) => b.valor - a.valor), [
          { label: 'Total R$',     val: fS(tot) },
          { label: 'Registros',    val: rows.length },
          { label: 'Média espera', val: avgH.toFixed(1) + 'h' },
          { label: 'Ticket médio', val: fS(rows.length ? tot / rows.length : 0) }
        ]);
      }
    }
  });

  // ── Ocorrência ──
  const bO = {};
  d.forEach(r => { const k = (r.ocorrencia || '—').trim(); bO[k] = (bO[k] || 0) + r.valor; });
  const oK = Object.keys(bO).sort((a, b) => bO[b] - bO[a]);
  const oV = oK.map(k => bO[k]);
  const oC = COLORS.slice(0, oK.length);
  document.getElementById('lgOc').innerHTML = oK.map((k, i) => `<span><span class="ld" style="background:${oC[i]}"></span>${k}</span>`).join('');
  mkDonut('cOcorrencia', oK, oV, oC, 'ocorrencia');

  // ── Filial ──
  const bF = {};
  d.forEach(r => { if (r.filial) bF[r.filial] = (bF[r.filial] || 0) + r.valor; });
  const fK = Object.keys(bF).sort((a, b) => bF[b] - bF[a]);
  const fV = fK.map(k => bF[k]);
  const fC = fK.map((_, i) => COLORS[i % COLORS.length]);
  document.getElementById('lgFilial').innerHTML = fK.map((k, i) => `<span><span class="ld" style="background:${fC[i]}"></span>${k}</span>`).join('');
  mkDonut('cFilial', fK, fV, fC, 'filial');

  // ── Contrato ──
  const bCo = {};
  d.forEach(r => { const k = (r.contrato || '—').trim(); if (k) bCo[k] = (bCo[k] || 0) + r.valor; });
  const coK = Object.keys(bCo).sort((a, b) => bCo[b] - bCo[a]);
  const coV = coK.map(k => bCo[k]);
  const coC = coK.map((_, i) => ['#2e9e44','#f0b429','#378ADD','#D85A30'][i % 4]);
  document.getElementById('lgContrato').innerHTML = coK.map((k, i) => `<span><span class="ld" style="background:${coC[i]}"></span>${k}</span>`).join('');
  mkDonut('cContrato', coK, coV, coC, 'contrato');

  // ── Status ──
  const bSt = {};
  d.forEach(r => { const k = (r.status || '—').trim().toUpperCase(); if (k) bSt[k] = (bSt[k] || 0) + r.valor; });
  const stK = Object.keys(bSt).sort((a, b) => bSt[b] - bSt[a]);
  const stV = stK.map(k => bSt[k]);
  const stC = stK.map(k => k.includes('FATURAD') ? '#2e9e44' : k.includes('APROV') ? '#f0b429' : '#ef4444');
  document.getElementById('lgStatus').innerHTML = stK.map((k, i) => `<span><span class="ld" style="background:${stC[i]}"></span>${k}</span>`).join('');
  mkDonut('cStatus', stK, stV, stC, 'status');
}

// ── TABLE ──
function renderTable() {
  const total = filteredData.length, pages = Math.ceil(total / PS), start = (currentPage - 1) * PS;
  document.getElementById('tableCount').textContent = total.toLocaleString('pt-BR') + ' registros';
  const tb = document.getElementById('tbody'); tb.innerHTML = '';
  filteredData.slice(start, start + PS).forEach(r => {
    const ri = allData.indexOf(r), s = r.status || '';
    const badge = s.includes('FATURAD') ? 'bf' : s.includes('APROV') ? 'ba' : 'br';
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${r.filial||'—'}</td><td>${r.cliente||'—'}</td><td>${r.placa||'—'}</td><td>${r.cte||'—'}</td><td>${r.nf||'—'}</td><td>${r.contrato||'—'}</td><td>${r.ocorrencia||'—'}</td><td class="vc">${(r.horas||0).toFixed(1)}h</td><td class="vc">${fR(r.valor)}</td><td><span class="badge ${badge}">${s||'—'}</span></td><td><button class="btn-del" onclick="delRow(${ri})">✕</button></td>`;
    tb.appendChild(tr);
  });
  const pg = document.getElementById('pagination');
  if (pages <= 1) { pg.innerHTML = `<span>${total} registro${total !== 1 ? 's' : ''}</span>`; return; }
  let btns = '', sp = [];
  for (let i = 1; i <= pages; i++) if (i === 1 || i === pages || Math.abs(i - currentPage) <= 2) sp.push(i);
  let prev = 0;
  sp.forEach(p => {
    if (prev && p - prev > 1) btns += `<button class="pgb" disabled>…</button>`;
    btns += `<button class="pgb${p === currentPage ? ' active' : ''}" onclick="goPage(${p})">${p}</button>`;
    prev = p;
  });
  pg.innerHTML = `<span>${start + 1}–${Math.min(start + PS, total)} de ${total}</span><div class="pg-btns"><button class="pgb" onclick="goPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>‹</button>${btns}<button class="pgb" onclick="goPage(${currentPage + 1})" ${currentPage === pages ? 'disabled' : ''}>›</button></div>`;
}

function goPage(p) { currentPage = Math.max(1, Math.min(p, Math.ceil(filteredData.length / PS))); renderTable(); }

// ── ADD / DELETE ──
function addRow() {
  const f = document.getElementById('iFilial').value.trim();
  const c = document.getElementById('iCliente').value.trim();
  const v = parseFloat(document.getElementById('iValor').value) || 0;
  if (!f || !c || !v) { alert('Preencha ao menos Filial, Cliente e Total R$.'); return; }
  allData.push({
    filial:     f, cliente: c,
    placa:      document.getElementById('iPlaca').value.trim().toUpperCase(),
    cte:        document.getElementById('iCte').value.trim(),
    nf:         document.getElementById('iNf').value.trim(),
    contrato:   document.getElementById('iContrato').value,
    ocorrencia: document.getElementById('iOcorrencia').value,
    horas:      parseFloat(document.getElementById('iHoras').value) || 0,
    valor:      v,
    status:     document.getElementById('iStatus').value,
    data:       new Date()
  });
  ['iFilial','iCliente','iPlaca','iCte','iNf','iValor','iHoras'].forEach(id => document.getElementById(id).value = '');
  populateFilters(); applyFilters();
}

function delRow(i) {
  if (!confirm('Remover este registro?')) return;
  allData.splice(i, 1); populateFilters(); applyFilters();
}

// ── EXPORT ──
function exportCSV() {
  const h    = ['Filial','Cliente','Placa','CTe','NF','Contrato','Ocorrência','Horas','Total R$','Status'];
  const rows = filteredData.map(r =>
    [r.filial, r.cliente, r.placa, r.cte, r.nf, r.contrato, r.ocorrencia, r.horas, r.valor, r.status]
    .map(v => `"${String(v || '').replace(/"/g, '""')}"`).join(',')
  );
  dl('diarias.csv', '﻿' + h.join(',') + '\n' + rows.join('\n'), 'text/csv;charset=utf-8');
}

function exportXLSX() {
  const ws = XLSX.utils.json_to_sheet(filteredData.map(r => ({
    Filial: r.filial, Cliente: r.cliente, Placa: r.placa, CTe: r.cte, NF: r.nf,
    Contrato: r.contrato, 'Ocorrência': r.ocorrencia, Horas: r.horas,
    'Total R$': r.valor, Status: r.status
  })));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Diárias');
  XLSX.writeFile(wb, 'diarias.xlsx');
}

function dl(name, content, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([content], { type }));
  a.download = name;
  a.click();
}
