/**
 * calc.js — DIN 4017:2006-03
 * Mecanismo de ruptura one-sided (DIN/ELPLA):
 *   C = (b/2, 0)  →  cunha ativa  →  espiral  →  cunha passiva
 *
 * Sistema de coordenadas: origem no centro da base da fundação, y↑ positivo.
 */

const RAD = Math.PI / 180;
const DEG = 180 / Math.PI;
const tan  = d => Math.tan(d * RAD);
const sin  = d => Math.sin(d * RAD);
const cos  = d => Math.cos(d * RAD);
const asin = r => Math.asin(r) * DEG;

// ─── Geometria da superfície de ruptura ──────────────────────────────────────
/**
 * Constrói os pontos da superfície de ruptura.
 * Mecanismo direito (one-sided), α=β=δ=0 (DIN 4017 Annexe A).
 *
 * Geometria (verificada contra ELPLA Example 11):
 *   α = 45 + φ/2  (ângulo da face da cunha ativa com a horizontal)
 *   θ = 45 − φ/2  (ângulo da face da cunha passiva com a horizontal)
 *   C  = (b/2, 0)  ← centro da espiral, borda direita da fundação
 *   r2 = b / (2·cos α)  ← raio inicial (DIN 4017 Eq. A.13)
 *   P2 = (0, −r2·sin α)  ← vértice da cunha ativa (abaixo da fundação)
 *   r1 = r2·e^(π/2·tan φ)
 *   P5 = C + r1·(cos(α−90°), sin(α−90°))  ← fim da espiral
 *   P6 = P5 + r1·(cos θ, sin θ)  ← na superfície
 *
 * @param {number} phiDeg  ângulo de atrito φ em graus
 * @param {number} b       largura da fundação (m)
 * @param {number} nPts    resolução da espiral (pontos)
 */
function buildRupturePoints(phiDeg, b, nPts) {
  nPts = nPts || 120;
  const alpha = 45 + phiDeg / 2;   // graus
  const theta = 45 - phiDeg / 2;   // graus

  const r2 = b / (2 * cos(alpha));  // Eq. A.13
  const r1 = r2 * Math.exp((Math.PI / 2) * tan(phiDeg));

  // Centro da espiral: borda direita da fundação, ao nível da base
  const Cx = b / 2;
  const Cy = 0;

  // Vértice da cunha ativa: ponto mais profundo da cunha (eixo de simetria)
  const xP2 = 0;
  const yP2 = -r2 * sin(alpha);  // negativo = abaixo da fundação

  // Ângulo do raio C→P2 (em radianos, sentido anti-horário de +x)
  const angle0 = Math.atan2(yP2 - Cy, xP2 - Cx);  // ≈ −(135°−φ/2) × RAD

  // Espiral logarítmica: rotação ANTI-HORÁRIA (ângulo aumenta) em y↑
  const spiralPoints = [];
  for (let i = 0; i <= nPts; i++) {
    const psi = (i / nPts) * (Math.PI / 2);          // 0 → π/2
    const r   = r2 * Math.exp(psi * tan(phiDeg));
    const ang = angle0 + psi;                         // anti-horário
    spiralPoints.push({
      x: Cx + r * Math.cos(ang),
      y: Cy + r * Math.sin(ang)
    });
  }

  // Ponto P5 = fim da espiral (ψ = π/2)
  const P5 = spiralPoints[nPts];

  // Ponto P6 = base da cunha passiva na superfície (y = 0)
  // Direção da face passiva: ângulo θ acima da horizontal, para direita
  // Comprimento = r1  (verificado: P6.y = 0)
  const xP6 = P5.x + r1 * cos(theta);
  const yP6 = P5.y + r1 * sin(theta);  // ≈ 0

  // Profundidade máxima da espiral (no ponto em que dy/dψ = 0, ou seja psi = α−90°... = psi em que angle0+psi aponta para baixo)
  // Ocorre quando angle0 + psi = −π/2  →  psi = −π/2 − angle0
  const psi_maxDepth = Math.max(0, Math.min(Math.PI / 2, -Math.PI / 2 - angle0));
  const r_maxDepth = r2 * Math.exp(psi_maxDepth * tan(phiDeg));
  const maxDepthY   = Cy + r_maxDepth * Math.sin(angle0 + psi_maxDepth); // y↑ (negativo)

  return {
    Cx, Cy, r2, r1, alpha, theta,
    xP2, yP2,             // vértice da cunha ativa
    P5, xP6, yP6,         // fim da espiral, superfície
    spiralPoints,
    maxDepth: maxDepthY,  // y↑ (negativo = abaixo)
    angle0
  };
}

// ─── Comprimentos da linha de ruptura por camada ─────────────────────────────
/**
 * Calcula os comprimentos li da linha de ruptura dentro de cada camada.
 *
 * A superfície de ruptura (mecanismo direito) é formada por 3 trechos:
 *   1. Cunha ativa: segmento (b/2, 0) → P2 (reta)
 *   2. Espiral:      arco de P2 → P5
 *   3. Cunha passiva: segmento P5 → P6 (reta)
 *
 * Profundidade de cada camada: camadas[0] está logo abaixo da fundação.
 */
function calcLayerLengths(phiDeg, b, layers) {
  const g = buildRupturePoints(phiDeg, b, 400);

  // Limites de profundidade z↓ de cada camada (z↓ = −y↑)
  const cumZ = [0];
  for (const lay of layers) cumZ.push(cumZ[cumZ.length - 1] + lay.h);

  const lengths = new Array(layers.length).fill(0);

  // Função que distribui comprimento de um segmento entre camadas
  function addSegment(x1, y1, x2, y2) {
    const zA = -y1;  // z↓
    const zB = -y2;
    const segLen = Math.hypot(x2 - x1, y2 - y1);
    const zMin = Math.min(zA, zB);
    const zMax = Math.max(zA, zB);
    if (zMax - zMin < 1e-12) {
      // segmento horizontal: atribui à camada correspondente
      for (let k = 0; k < layers.length; k++) {
        if (zA >= cumZ[k] && zA < cumZ[k + 1]) {
          lengths[k] += segLen;
          break;
        }
      }
      return;
    }
    for (let k = 0; k < layers.length; k++) {
      const top = cumZ[k], bot = cumZ[k + 1];
      const oTop = Math.max(zMin, top);
      const oBot = Math.min(zMax, bot);
      if (oBot <= oTop) continue;
      lengths[k] += (oBot - oTop) / (zMax - zMin) * segLen;
    }
  }

  // 1. Cunha ativa: (b/2, 0) → P2 = (xP2, yP2)
  addSegment(b / 2, 0, g.xP2, g.yP2);

  // 2. Espiral: segmentos consecutivos
  const sp = g.spiralPoints;
  for (let i = 0; i < sp.length - 1; i++) {
    addSegment(sp[i].x, sp[i].y, sp[i + 1].x, sp[i + 1].y);
  }

  // 3. Cunha passiva: P5 → P6
  addSegment(g.P5.x, g.P5.y, g.xP6, g.yP6);

  return lengths;
}

// ─── Iteração para φm ────────────────────────────────────────────────────────
function iteratePhiM(b, layers, maxIter, tol) {
  maxIter = maxIter || 20;
  tol     = tol     || 3.0;  // % (DIN 4017)

  const iterations = [];
  // Passo inicial: maior φ das camadas
  let phiIn = Math.max(...layers.map(l => l.phi));

  for (let iter = 0; iter < maxIter; iter++) {
    const ls     = calcLayerLengths(phiIn, b, layers);
    const lTotal = ls.reduce((s, v) => s + v, 0);

    if (lTotal < 1e-10) {
      return { phiM: phiIn, iterations, converged: false,
               error: 'Comprimento total da linha de ruptura = 0' };
    }

    // φm ponderado por comprimentos
    let sumTanPhi = 0;
    for (let k = 0; k < layers.length; k++) {
      sumTanPhi += ls[k] * tan(layers[k].phi);
    }
    const phiOut = Math.atan(sumTanPhi / lTotal) * DEG;
    const delta  = Math.abs(phiIn - phiOut) / (phiIn > 0 ? phiIn : 1) * 100;

    iterations.push({ iter: iter + 1, phiIn, phiOut, delta,
                      lengths: ls.slice(), lTotal });

    if (delta < tol) {
      return { phiM: (phiIn + phiOut) / 2, iterations, converged: true };
    }
    phiIn = (phiIn + phiOut) / 2;  // média (conforme ELPLA)
  }
  return { phiM: phiIn, iterations, converged: false };
}

// ─── cm (coesão equivalente) ─────────────────────────────────────────────────
function calcCm(phiM, b, layers) {
  const ls     = calcLayerLengths(phiM, b, layers);
  const lTotal = ls.reduce((s, v) => s + v, 0);
  let sumC = 0;
  for (let k = 0; k < layers.length; k++) sumC += ls[k] * layers[k].c;
  return { cm: sumC / lTotal, lengths: ls, lTotal };
}

// ─── γm (peso específico médio abaixo da fundação) ───────────────────────────
function calcGammaM(phiM, b, layers) {
  const g = buildRupturePoints(phiM, b, 600);

  // Limites de profundidade
  const cumZ = [0];
  for (const lay of layers) cumZ.push(cumZ[cumZ.length - 1] + lay.h);

  const areas  = new Array(layers.length).fill(0);
  const nInteg = 800;

  // Profundidade máxima da superfície de ruptura
  const zMax = Math.min(
    Math.max(-g.maxDepth, -g.yP2, -g.P5.y) + 0.5,
    cumZ[cumZ.length - 1]
  );

  // Para cada profundidade z, calcular largura da superfície de ruptura
  // Largura = x_right(z) − x_left(z)
  // Limite esquerdo: face da cunha ativa (de x=b/2 até x=0 em z=d_ponta)
  //   x_left(z) = b/2 * (1 − z/|yP2|)  para z ≤ |yP2|, senão 0
  // Limite direito: espiral ou face passiva, obtido por interseção numérica

  const sp   = g.spiralPoints;
  const xP2  = g.xP2;
  const yP2  = g.yP2;
  const d_ponta = -yP2;  // profundidade do vértice da cunha ativa

  for (let i = 0; i < nInteg; i++) {
    const z  = (i + 0.5) / nInteg * zMax;
    const dz = zMax / nInteg;
    const y_target = -z;  // y↑

    // Limite esquerdo (face da cunha ativa ou eixo)
    let xLeft = 0;
    if (z <= d_ponta) {
      xLeft = (b / 2) * (1 - z / d_ponta);
    }

    // Limite direito: cruzamento com espiral ou cunha passiva (x mais alto)
    let xRight = xLeft;

    // Verifica espiral
    for (let j = 0; j < sp.length - 1; j++) {
      const yA = sp[j].y, yB = sp[j + 1].y;
      if ((yA - y_target) * (yB - y_target) <= 0 && Math.abs(yA - yB) > 1e-12) {
        const t = (y_target - yA) / (yB - yA);
        const xc = sp[j].x + t * (sp[j + 1].x - sp[j].x);
        if (xc > xRight) xRight = xc;
      }
    }

    // Verifica cunha passiva (P5 → P6)
    {
      const yA = g.P5.y, yB = g.yP6;
      if ((yA - y_target) * (yB - y_target) <= 0 && Math.abs(yA - yB) > 1e-12) {
        const t  = (y_target - yA) / (yB - yA);
        const xc = g.P5.x + t * (g.xP6 - g.P5.x);
        if (xc > xRight) xRight = xc;
      }
    }

    const width = xRight - xLeft;
    if (width <= 0) continue;

    for (let k = 0; k < layers.length; k++) {
      if (z >= cumZ[k] && z < cumZ[k + 1]) {
        areas[k] += width * dz;
        break;
      }
    }
  }

  const aTotal = areas.reduce((s, v) => s + v, 0);
  let sumGamma = 0;
  for (let k = 0; k < layers.length; k++) sumGamma += areas[k] * layers[k].gamma;

  return { gammaM: aTotal > 0 ? sumGamma / aTotal : 0, areas, aTotal };
}

// ─── γ'm (peso específico acima da fundação) ─────────────────────────────────
function calcGammaMAbove(tf, aboveLayers) {
  let totalH = 0, sumGH = 0;
  for (const lay of aboveLayers) {
    totalH += lay.h;
    sumGH  += lay.h * lay.gamma;
  }
  return totalH > 0 ? sumGH / totalH : 0;
}

// ─── Fatores de capacidade de carga (DIN 4017 Eqs. 5–7) ─────────────────────
function bearingFactors(phiDeg) {
  if (phiDeg <= 0.001) return { Nd: 1, Nc: 5.14, Nb: 0 };
  const Nd = Math.tan((45 + phiDeg / 2) * RAD) ** 2 * Math.exp(Math.PI * tan(phiDeg));
  const Nc = (Nd - 1) / tan(phiDeg);
  const Nb = (Nd - 1) * tan(phiDeg);  // DIN 4017 Eq. 6
  return { Nd, Nc, Nb };
}

// ─── Fatores de forma (DIN 4017 Tabelle 2 — Rechteck) ────────────────────────
function shapeFactors(phiDeg, b, a, Nd) {
  const ratio = b / a;  // b ≤ a
  if (phiDeg <= 0.001) {
    return { nub: 1, nud: 1, nuc: 1 + 0.2 * ratio };
  }
  const nub = 1 - 0.3 * ratio;
  const nud = 1 + ratio * sin(phiDeg);
  const nuc = (nud * Nd - 1) / (Nd - 1);
  return { nub, nud, nuc };
}

// ─── Cálculo principal ────────────────────────────────────────────────────────
function calculateBearingCapacity(b, a, tf, layers, above) {
  // 1. Iteração φm
  const iterResult = iteratePhiM(b, layers);
  if (iterResult.error) return { error: iterResult.error };
  const phiM = iterResult.phiM;

  // 2. cm
  const { cm, lengths, lTotal } = calcCm(phiM, b, layers);

  // 3. γm
  const { gammaM, areas, aTotal } = calcGammaM(phiM, b, layers);

  // 4. γ'm
  const gammaMAbove = calcGammaMAbove(tf, above);

  // 5. Fatores
  const { Nd, Nc, Nb }   = bearingFactors(phiM);
  const { nub, nud, nuc } = shapeFactors(phiM, b, a, Nd);

  // 6. qult (DIN 4017 Eq. 1)
  const term_c = cm         * Nc * nuc;
  const term_q = gammaMAbove * tf * Nd * nud;
  const term_b = gammaM     * b  * Nb * nub;
  const qult   = term_c + term_q + term_b;

  return {
    phiM, cm, gammaM, gammaMAbove,
    Nd, Nc, Nb, nub, nud, nuc,
    term_c, term_q, term_b, qult,
    iterations:  iterResult.iterations,
    converged:   iterResult.converged,
    lengths, lTotal, areas, aTotal,
    geom: buildRupturePoints(phiM, b, 120)
  };
}
