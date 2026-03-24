/**
 * draw.js — Visualização da superfície de ruptura (Canvas 2D)
 * Sistema: origem no centro da base da fundação, x→, y↑.
 */

const LAYER_COLORS = [
  '#c8e6c9','#fff9c4','#ffe0b2','#f8bbd0','#d1c4e9',
  '#b3e5fc','#dcedc8','#fce4ec','#e8eaf6','#e0f2f1'
];

function drawRupture(canvasId, geom, layers, tf, b, zw) {
  const canvas = document.getElementById(canvasId);
  if (!canvas || !geom || !geom.spiralPoints) return;
  const ctx = canvas.getContext('2d');
  const W = canvas.width;
  const H = canvas.height;
  ctx.clearRect(0, 0, W, H);

  const sp = geom.spiralPoints;

  // Profundidade máxima real (numericamente)
  const allY = sp.map(p => p.y).concat([geom.yP2, geom.P5.y]);
  const yMinData = Math.min(...allY);  // mais negativo = mais fundo

  // Limites do mundo (y↑)
  const xMin = -b * 0.8;
  const xMax = geom.xP6 + b * 0.5;
  const yMin = yMinData - 1.0;
  const yMax = tf + 1.5;

  const margin = { top: 30, bottom: 10, left: 55, right: 20 };
  const pw = W - margin.left - margin.right;
  const ph = H - margin.top - margin.bottom;

  function cx(x) { return margin.left + (x - xMin) / (xMax - xMin) * pw; }
  function cy(y) { return margin.top  + (yMax - y) / (yMax - yMin) * ph; }

  // ── Camadas de solo (abaixo da fundação) ────────────────────────────────
  let cumZ = 0;
  for (let k = 0; k < layers.length; k++) {
    const yTop = -cumZ;
    const yBot = -(cumZ + layers[k].h);
    ctx.fillStyle = LAYER_COLORS[k % LAYER_COLORS.length];
    ctx.fillRect(cx(xMin), cy(yTop), pw, cy(yBot) - cy(yTop));

    // Label
    const yMid = -(cumZ + layers[k].h / 2);
    ctx.fillStyle = '#333';
    ctx.font = '11px Segoe UI';
    ctx.textAlign = 'left';
    const label = layers[k].desc
      ? `${k+1}. ${layers[k].desc}  φ=${layers[k].phi}°  c=${layers[k].c} kN/m²  γ=${layers[k].gamma} kN/m³`
      : `Camada ${k+1}: φ=${layers[k].phi}°  c=${layers[k].c} kN/m²`;
    ctx.fillText(label, cx(xMin) + 4, cy(yMid) + 4);
    cumZ += layers[k].h;
  }

  // ── Solo acima da fundação (cinza) ───────────────────────────────────────
  ctx.fillStyle = '#dce3ea';
  ctx.fillRect(cx(xMin), cy(yMax), pw, cy(0) - cy(yMax));

  // ── Linhas de fronteira de camadas ───────────────────────────────────────
  ctx.strokeStyle = '#666';
  ctx.lineWidth = 1;
  ctx.setLineDash([5, 4]);
  let cumZb = 0;
  for (let k = 0; k < layers.length - 1; k++) {
    cumZb += layers[k].h;
    const yLine = -cumZb;
    ctx.beginPath();
    ctx.moveTo(cx(xMin), cy(yLine));
    ctx.lineTo(cx(xMax), cy(yLine));
    ctx.stroke();
    // Cota da fronteira
    ctx.fillStyle = '#555';
    ctx.font = '10px Segoe UI';
    ctx.textAlign = 'right';
    ctx.fillText(`${cumZb.toFixed(1)}m`, cx(xMin) - 2, cy(yLine) + 3);
  }
  ctx.setLineDash([]);

  // ── Superfície do terreno ─────────────────────────────────────────────────
  ctx.strokeStyle = '#388e3c';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(cx(xMin), cy(0));
  ctx.lineTo(cx(xMax), cy(0));
  ctx.stroke();

  // ── Fundação ──────────────────────────────────────────────────────────────
  const fL = cx(-b / 2), fR = cx(b / 2), fT = cy(0);
  ctx.fillStyle = '#455a64';
  ctx.fillRect(fL, fT - 14, fR - fL, 14);
  ctx.strokeStyle = '#263238';
  ctx.lineWidth = 1;
  ctx.strokeRect(fL, fT - 14, fR - fL, 14);
  ctx.fillStyle = '#fff';
  ctx.font = 'bold 10px Segoe UI';
  ctx.textAlign = 'center';
  ctx.fillText(`b=${b}m`, (fL + fR) / 2, fT - 4);

  // ── Superfície de ruptura ────────────────────────────────────────────────
  // Cunha ativa: (b/2, 0) → P2
  ctx.strokeStyle = '#e53935';
  ctx.lineWidth = 2.5;
  ctx.beginPath();
  ctx.moveTo(cx(b / 2), cy(0));
  ctx.lineTo(cx(geom.xP2), cy(geom.yP2));
  // Espiral
  for (const p of sp) ctx.lineTo(cx(p.x), cy(p.y));
  // Cunha passiva: P5 → P6
  ctx.lineTo(cx(geom.xP6), cy(geom.yP6));
  ctx.stroke();

  // ── Nível d'água ─────────────────────────────────────────────────────────
  if (zw != null) {
    // No sistema de coordenadas do canvas: y=0 é a base da fundação.
    // A superfície fica em y = tf. O NA fica em y = tf − zw.
    const y_wt = tf - zw;
    ctx.strokeStyle = '#1565c0';
    ctx.lineWidth = 1.5;
    ctx.setLineDash([10, 5]);
    ctx.beginPath();
    ctx.moveTo(cx(xMin), cy(y_wt));
    ctx.lineTo(cx(xMax), cy(y_wt));
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = '#1565c0';
    ctx.font = 'bold 10px Segoe UI';
    ctx.textAlign = 'left';
    ctx.fillText(`NA  (${zw}m desde superfície)`, cx(xMin) + 4, cy(y_wt) - 4);
  }

  // ── Eixo de simetria (tracejado) ─────────────────────────────────────────
  ctx.strokeStyle = '#9e9e9e';
  ctx.lineWidth = 1;
  ctx.setLineDash([4, 4]);
  ctx.beginPath();
  ctx.moveTo(cx(0), cy(0));
  ctx.lineTo(cx(0), cy(geom.yP2));
  ctx.stroke();
  ctx.setLineDash([]);

  // ── Cotas verticais (profundidade desde a superfície do terreno) ─────────
  // y = 0  →  base da fundação  →  profundidade = tf
  // y = -z →  z abaixo da base →  profundidade = tf + z
  ctx.font = '10px Segoe UI';
  ctx.textAlign = 'right';
  const depthMax = Math.ceil(-yMinData + 1);
  const step = (tf + depthMax) <= 8 ? 1 : (tf + depthMax) <= 16 ? 2 : 5;

  // Cota "0m" na superfície do terreno (y = tf no sistema do canvas)
  ctx.fillStyle = '#388e3c';
  ctx.fillText('0m', cx(xMin) - 2, cy(tf) + 3);

  // Cota da base da fundação
  ctx.fillStyle = '#555';
  ctx.fillText(`${tf}m`, cx(xMin) - 2, cy(0) + 3);

  // Cotas abaixo da fundação
  for (let z = step; z <= depthMax; z += step) {
    const yLine       = -z;                // coordenada world
    const surfaceDepth = tf + z;           // profundidade desde superfície
    // só imprimir se não colidir com o rótulo tf já desenhado
    if (Math.abs(surfaceDepth - tf) < step * 0.6) continue;
    ctx.fillStyle = '#333';
    ctx.fillText(`${surfaceDepth}m`, cx(xMin) - 2, cy(yLine) + 3);
    ctx.strokeStyle = '#e0e0e0';
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(cx(xMin), cy(yLine));
    ctx.lineTo(cx(xMax), cy(yLine));
    ctx.stroke();
  }

  // ── Legenda ───────────────────────────────────────────────────────────────
  const legendEl = document.getElementById('legend');
  if (legendEl) {
    legendEl.innerHTML =
      layers.map((l, k) =>
        `<div class="legend-item">
          <div class="legend-color" style="background:${LAYER_COLORS[k % LAYER_COLORS.length]};border:1px solid #aaa"></div>
          <span>Camada ${k+1}${l.desc ? ': ' + l.desc : ''}</span>
        </div>`
      ).join('') +
      `<div class="legend-item">
        <div class="legend-color" style="background:#e53935;height:3px;margin-top:5px;"></div>
        <span>Superfície de ruptura (DIN 4017)</span>
      </div>`;
  }
}
