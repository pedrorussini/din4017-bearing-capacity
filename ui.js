/**
 * ui.js — Interface do usuário
 */

// Estado
let layerCount = 0;
let aboveCount = 0;

// ── Adicionar/remover camadas abaixo ─────────────────────────────────────────
function addLayer(data) {
  layerCount++;
  const tbody = document.getElementById('layers-body');
  const tr = document.createElement('tr');
  tr.id = `layer-row-${layerCount}`;
  tr.innerHTML = `
    <td>${layerCount}</td>
    <td><input type="text" class="wide" id="ldesc${layerCount}" value="${data && data.desc ? data.desc : ''}" placeholder="Descrição (opcional)" style="width:130px;padding:5px 6px;border:1px solid #cdd3db;border-radius:5px;font-size:0.82rem;" /></td>
    <td><input type="number" id="lh${layerCount}" value="${data ? data.h : 2.0}" step="0.1" min="0.1" /></td>
    <td><input type="number" id="lg${layerCount}" value="${data ? data.gamma : 18.0}" step="0.5" /></td>
    <td><input type="number" id="lp${layerCount}" value="${data ? data.phi : 30.0}" step="0.5" min="0" max="50" /></td>
    <td><input type="number" id="lc${layerCount}" value="${data ? data.c : 0.0}" step="0.5" min="0" /></td>
  `;
  tbody.appendChild(tr);
}

function removeLayer() {
  if (layerCount <= 1) return;
  const tbody = document.getElementById('layers-body');
  tbody.removeChild(tbody.lastElementChild);
  layerCount--;
}

// ── Adicionar/remover camadas acima ──────────────────────────────────────────
function addAbove(data) {
  aboveCount++;
  const tbody = document.getElementById('above-body');
  const tr = document.createElement('tr');
  tr.id = `above-row-${aboveCount}`;
  tr.innerHTML = `
    <td>${aboveCount}</td>
    <td><input type="text" style="width:130px;padding:5px 6px;border:1px solid #cdd3db;border-radius:5px;font-size:0.82rem;" id="adesc${aboveCount}" value="${data && data.desc ? data.desc : ''}" placeholder="Descrição" /></td>
    <td><input type="number" id="ah${aboveCount}" value="${data ? data.h : 1.0}" step="0.1" min="0.01" /></td>
    <td><input type="number" id="ag${aboveCount}" value="${data ? data.gamma : 18.0}" step="0.5" /></td>
  `;
  tbody.appendChild(tr);
}

function removeAbove() {
  if (aboveCount <= 1) return;
  const tbody = document.getElementById('above-body');
  tbody.removeChild(tbody.lastElementChild);
  aboveCount--;
}

// ── Carregar exemplo ELPLA (Example 11) ──────────────────────────────────────
function loadExample() {
  // Limpar camadas existentes
  document.getElementById('layers-body').innerHTML = '';
  document.getElementById('above-body').innerHTML = '';
  layerCount = 0;
  aboveCount = 0;

  // Geometria da fundação
  document.getElementById('b').value  = 4.0;
  document.getElementById('a').value  = 5.0;
  document.getElementById('tf').value = 2.0;

  // Camadas abaixo da fundação (Example 11, DIN 4017)
  // Fronteiras em 1.5m e 3.0m abaixo da fundação
  const layersData = [
    { desc: 'Areia média',    h: 1.5,  gamma: 11.0, phi: 30.0, c: 0.0 },
    { desc: 'Silte',          h: 1.5,  gamma: 12.0, phi: 25.0, c: 5.0 },
    { desc: 'Solo residual',  h: 10.0, gamma: 10.0, phi: 22.5, c: 2.0 }
  ];
  layersData.forEach(d => addLayer(d));

  // Camadas acima da fundação (Example 11)
  const aboveData = [
    { desc: 'Argila com areia', h: 0.5, gamma: 18.0 },
    { desc: 'Cascalho grosso',  h: 1.1, gamma: 18.5 },
    { desc: 'Areia média',      h: 0.4, gamma: 11.0 }
  ];
  aboveData.forEach(d => addAbove(d));

  hideResults();
  showError('');
}

// ── Ler inputs ────────────────────────────────────────────────────────────────
function readLayers() {
  const layers = [];
  for (let i = 1; i <= layerCount; i++) {
    layers.push({
      desc:  document.getElementById(`ldesc${i}`).value,
      h:     parseFloat(document.getElementById(`lh${i}`).value),
      gamma: parseFloat(document.getElementById(`lg${i}`).value),
      phi:   parseFloat(document.getElementById(`lp${i}`).value),
      c:     parseFloat(document.getElementById(`lc${i}`).value)
    });
  }
  return layers;
}

function readAbove() {
  const above = [];
  for (let i = 1; i <= aboveCount; i++) {
    above.push({
      desc:  document.getElementById(`adesc${i}`).value,
      h:     parseFloat(document.getElementById(`ah${i}`).value),
      gamma: parseFloat(document.getElementById(`ag${i}`).value)
    });
  }
  return above;
}

// ── Validação ─────────────────────────────────────────────────────────────────
function validate(b, a, tf, layers, above) {
  if (b <= 0 || a <= 0) return 'Dimensões b e a devem ser positivas.';
  if (tf < 0) return 'Profundidade de fundação tf não pode ser negativa.';
  if (layers.length === 0) return 'Adicione pelo menos uma camada de solo.';
  for (const l of layers) {
    if (isNaN(l.h) || l.h <= 0) return 'Todas as camadas devem ter espessura > 0.';
    if (isNaN(l.phi) || l.phi < 0) return 'Ângulo de atrito φ deve ser ≥ 0.';
    if (isNaN(l.c) || l.c < 0) return 'Coesão c deve ser ≥ 0.';
  }
  if (above.length === 0) return 'Adicione pelo menos uma sub-camada acima da fundação.';
  const totalAbove = above.reduce((s, l) => s + l.h, 0);
  if (Math.abs(totalAbove - tf) > 0.05 && tf > 0) {
    return `Atenção: soma das espessuras acima (${totalAbove.toFixed(2)} m) difere de tf (${tf} m). Verifique.`;
  }
  return null;
}

// ── Calcular ──────────────────────────────────────────────────────────────────
function calculate() {
  const b  = parseFloat(document.getElementById('b').value);
  const a  = parseFloat(document.getElementById('a').value);
  const tf = parseFloat(document.getElementById('tf').value);
  const layers = readLayers();
  const above  = readAbove();

  const valErr = validate(b, a, tf, layers, above);
  if (valErr) {
    // Avisos de espessura não bloqueiam o cálculo
    if (valErr.startsWith('Atenção')) {
      showError(valErr);
    } else {
      showError(valErr);
      return;
    }
  } else {
    showError('');
  }

  const res = calculateBearingCapacity(b, a, tf, layers, above);

  if (res.error) {
    showError('Erro no cálculo: ' + res.error);
    return;
  }

  showResults(res, b, a, tf, layers);
}

// ── Exibir resultados ─────────────────────────────────────────────────────────
function showResults(res, b, a, tf, layers) {
  document.getElementById('results').style.display = 'block';

  // Caixas de resultado
  const grid = document.getElementById('res-grid');
  grid.innerHTML = `
    <div class="res-box">
      <div class="label">φ<sub>m</sub></div>
      <div class="value">${res.phiM.toFixed(2)}°</div>
      <div class="unit">ângulo de atrito equivalente</div>
    </div>
    <div class="res-box">
      <div class="label">c<sub>m</sub></div>
      <div class="value">${res.cm.toFixed(2)}</div>
      <div class="unit">kN/m² — coesão equivalente</div>
    </div>
    <div class="res-box">
      <div class="label">γ<sub>m</sub></div>
      <div class="value">${res.gammaM.toFixed(2)}</div>
      <div class="unit">kN/m³ — peso esp. abaixo</div>
    </div>
    <div class="res-box">
      <div class="label">γ'<sub>m</sub></div>
      <div class="value">${res.gammaMAbove.toFixed(2)}</div>
      <div class="unit">kN/m³ — peso esp. acima</div>
    </div>
    <div class="res-box">
      <div class="label">N<sub>d</sub> / N<sub>c</sub> / N<sub>b</sub></div>
      <div class="value" style="font-size:1rem">${res.Nd.toFixed(1)} / ${res.Nc.toFixed(1)} / ${res.Nb.toFixed(1)}</div>
      <div class="unit">fatores de carga</div>
    </div>
    <div class="res-box">
      <div class="label">ν<sub>d</sub> / ν<sub>c</sub> / ν<sub>b</sub></div>
      <div class="value" style="font-size:1rem">${res.nud.toFixed(3)} / ${res.nuc.toFixed(3)} / ${res.nub.toFixed(3)}</div>
      <div class="unit">fatores de forma</div>
    </div>
    <div class="res-box" style="background:#e8f5e9;border:2px solid #2e7d32;">
      <div class="label" style="color:#1b5e20;font-weight:600;">q<sub>ult</sub></div>
      <div class="value" style="font-size:1.7rem;color:#1b5e20;">${res.qult.toFixed(0)}</div>
      <div class="unit">kN/m² — capacidade de carga última</div>
    </div>
    <div class="res-box" style="font-size:0.78rem;text-align:left;padding:10px 12px;">
      <div class="label">Parcelas (kN/m²)</div>
      <div>c·Nc·νc = <b>${res.term_c.toFixed(1)}</b></div>
      <div>γ'·tf·Nd·νd = <b>${res.term_q.toFixed(1)}</b></div>
      <div>γ·b·Nb·νb = <b>${res.term_b.toFixed(1)}</b></div>
    </div>
  `;

  // Tabela de iterações
  const tbody = document.getElementById('iter-body');
  tbody.innerHTML = '';
  for (const it of res.iterations) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${it.iter}</td>
      <td>${it.phiIn.toFixed(4)}°</td>
      <td>${it.phiOut.toFixed(4)}°</td>
      <td>${it.delta.toFixed(2)}%</td>
      <td style="color:${it.delta < 3 ? '#2e7d32' : '#c62828'}">${it.delta < 3 ? '✓ Sim' : '✗ Não'}</td>
    `;
    tbody.appendChild(tr);
  }
  if (!res.converged) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="5" style="color:#c62828;font-weight:600;">⚠ Não convergiu dentro do limite de iterações</td>`;
    tbody.appendChild(tr);
  }

  // Gráfico
  drawRupture('rupture-canvas', res.geom, layers, tf, b);
}

function hideResults() {
  document.getElementById('results').style.display = 'none';
}

function showError(msg) {
  const el = document.getElementById('error-msg');
  el.textContent = msg;
  el.style.display = msg ? 'block' : 'none';
}

// ── Inicialização ─────────────────────────────────────────────────────────────
(function init() {
  // Começa com 3 camadas e 2 sub-camadas acima como padrão
  addLayer({ h: 3.0, gamma: 18.0, phi: 30.0, c: 0.0 });
  addLayer({ h: 3.0, gamma: 18.0, phi: 25.0, c: 0.0 });
  addLayer({ h: 4.0, gamma: 17.0, phi: 20.0, c: 5.0 });
  addAbove({ h: 1.0, gamma: 18.0 });
  addAbove({ h: 1.0, gamma: 18.5 });
})();
