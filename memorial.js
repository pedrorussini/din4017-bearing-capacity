/**
 * memorial.js — Geração de Memória de Cálculo em Word (.docx)
 * Equações no formato OMML (nativo do Microsoft Word).
 * Norma: DIN 4017:2006-03
 */

let _lastResult = null;
let _lastInputs  = null;

function setMemorialData(res, b, a, tf, layers, above, zw) {
  _lastResult = res;
  _lastInputs  = { b, a, tf, layers, above, zw };
}

async function downloadMemorial() {
  if (!_lastResult) { alert('Execute o cálculo antes de gerar a memória.'); return; }
  const D = window.docx;
  if (!D) {
    if (window._docxLoadError) {
      alert('Falha ao carregar a biblioteca docx (CDN inacessível).\nVerifique sua conexão com a internet e recarregue a página.');
    } else {
      alert('Biblioteca docx ainda não carregada.\nAguarde alguns segundos e tente novamente, ou recarregue a página.');
    }
    return;
  }

  const R = _lastResult;
  const { b, a, tf, layers, above, zw } = _lastInputs;

  // ─── Helpers ─────────────────────────────────────────────────────────────
  const H1 = (text) => new D.Paragraph({ text, heading: D.HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 140 } });
  const H2 = (text) => new D.Paragraph({ text, heading: D.HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 100 } });
  const H3 = (text) => new D.Paragraph({ text, heading: D.HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 } });
  const P  = (...runs) => new D.Paragraph({ children: runs, spacing: { after: 100 } });
  const PI = (...runs) => new D.Paragraph({ children: runs, spacing: { after: 100 },
    indent: { left: 360 } });
  const T  = (text, opts = {}) => new D.TextRun({ text, ...opts });
  const B  = (text) => new D.TextRun({ text, bold: true });
  const I  = (text) => new D.TextRun({ text, italics: true });
  const BR = () => new D.Paragraph({ spacing: { after: 60 } });

  const EQ = (...children) => new D.Paragraph({
    children: [new D.Math({ children })],
    alignment: D.AlignmentType.CENTER,
    spacing: { before: 100, after: 100 }
  });

  const MR   = (t) => new D.MathRun(t);
  const MF   = (num, den) => new D.MathFraction({ numerator: num, denominator: den });
  const MSup = (base, sup) => new D.MathSuperScript({ children: base, superScript: sup });
  const MSub = (base, sub) => new D.MathSubScript({ children: base, subScript: sub });
  const MSSub = (base, sup, sub) => new D.MathSubSuperScript({ children: base, superScript: sup, subScript: sub });

  function tbl(headers, rows) {
    const borderOpts = { style: D.BorderStyle.SINGLE, size: 4, color: '999999' };
    const borders = { top: borderOpts, bottom: borderOpts, left: borderOpts, right: borderOpts,
                      insideHorizontal: borderOpts, insideVertical: borderOpts };
    const mkCell = (text, header = false) => new D.TableCell({
      children: [new D.Paragraph({
        children: [new D.TextRun({ text: String(text), bold: header, size: header ? 19 : 18 })],
        alignment: D.AlignmentType.CENTER,
        spacing: { before: 40, after: 40 }
      })],
      borders,
      shading: header ? { fill: 'D6E4F0' } : undefined,
      margins: { top: 60, bottom: 60, left: 80, right: 80 }
    });
    return new D.Table({
      width: { size: 100, type: D.WidthType.PERCENTAGE },
      rows: [
        new D.TableRow({ children: headers.map(h => mkCell(h, true)), tableHeader: true }),
        ...rows.map(row => new D.TableRow({ children: row.map(c => mkCell(c)) }))
      ]
    });
  }

  const n2 = v => (typeof v === 'number' ? v.toFixed(2) : v);
  const n3 = v => (typeof v === 'number' ? v.toFixed(3) : v);
  const n4 = v => (typeof v === 'number' ? v.toFixed(4) : v);

  const geom = buildRupturePoints(R.phiM, b, 10);
  const gammaEffArr = R.gammaEff || layers.map(l => l.gamma);

  // ─── Construção do documento ──────────────────────────────────────────────
  const children = [];

  // ══════════════════════════════════════════════════════════════════════════
  // CAPA
  // ══════════════════════════════════════════════════════════════════════════
  children.push(new D.Paragraph({ spacing: { after: 400 } }));
  children.push(new D.Paragraph({
    children: [T('MEMÓRIA DE CÁLCULO', { bold: true, size: 36 })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 80 }
  }));
  children.push(new D.Paragraph({
    children: [T('Capacidade de Carga Última em Solo em Camadas', { size: 26, italics: true })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 80 }
  }));
  children.push(new D.Paragraph({
    children: [T('Método Iterativo — DIN 4017:2006-03', { size: 22, color: '1a3a5c' })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 80 }
  }));
  children.push(new D.Paragraph({
    children: [T(`Data: ${new Date().toLocaleDateString('pt-BR')}`, { size: 20, color: '666666' })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 600 }
  }));

  // Quadro resumo
  children.push(new D.Paragraph({
    children: [T('RESUMO DOS RESULTADOS', { bold: true, size: 22, color: '1a3a5c' })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 100 }
  }));
  children.push(tbl(
    ['Parâmetro', 'Valor', 'Unidade'],
    [
      ['Fundação b × a', `${b} × ${a}`, 'm'],
      ['Profundidade tf', tf, 'm'],
      ['φm equivalente', n2(R.phiM), '°'],
      ['cm equivalente', n2(R.cm), 'kN/m²'],
      ['γm equivalente', n2(R.gammaM), 'kN/m³'],
      ["γ'm equivalente", n2(R.gammaMAbove), 'kN/m³'],
      ['qult', R.qult.toFixed(0), 'kN/m²'],
    ]
  ));
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 1. DADOS DE ENTRADA
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('1. Dados de Entrada'));

  children.push(H2('1.1 Geometria da Fundação'));
  children.push(P(T(
    'A fundação é caracterizada por sua largura b (menor dimensão), comprimento a (maior dimensão) ' +
    'e profundidade de assentamento tf medida a partir da superfície do terreno. ' +
    'A DIN 4017 adota convenção de que b ≤ a, pois os fatores de forma são expressos em função de b/a.'
  )));
  children.push(tbl(
    ['Parâmetro', 'Símbolo', 'Valor', 'Unidade'],
    [
      ['Largura (menor dimensão)', 'b', b, 'm'],
      ['Comprimento (maior dimensão)', 'a', a, 'm'],
      ['Profundidade de fundação', 'tf', tf, 'm'],
      ['Relação b/a', 'b/a', n3(b/a), '—'],
    ]
  ));
  children.push(BR());

  children.push(H2('1.2 Perfil de Solo — Camadas Abaixo da Fundação'));
  children.push(P(T(
    'As camadas são definidas abaixo da cota de assentamento da fundação (z = 0 corresponde à base da fundação). ' +
    'Os parâmetros necessários são: espessura hi, peso específico natural γi, ' +
    'ângulo de atrito interno φi e coesão ci. Quando houver nível d\'água, ' +
    'o peso específico efetivo γ\'i = γi − γw deve ser adotado na camada submersa.'
  )));
  children.push(tbl(
    ['#', 'Descrição', 'h (m)', 'γ (kN/m³)', 'φ (°)', 'c (kN/m²)'],
    layers.map((l, i) => [i+1, l.desc || `Camada ${i+1}`, l.h, l.gamma, l.phi, l.c])
  ));
  children.push(BR());

  children.push(H2('1.3 Perfil de Solo — Camadas Acima da Fundação'));
  children.push(P(T(
    'As camadas acima da fundação contribuem como sobrecarga lateral ao mecanismo de ruptura. ' +
    'Seu peso específico médio γ\'m é calculado como média ponderada pela espessura e ' +
    'entra na parcela de sobrecarga da equação de capacidade de carga. ' +
    'Devem somar espessura total igual à profundidade de fundação tf.'
  )));
  children.push(tbl(
    ['#', 'Descrição', 'h (m)', 'γ (kN/m³)'],
    above.map((l, i) => [i+1, l.desc || `Sub-camada ${i+1}`, l.h, l.gamma])
  ));
  children.push(P(T(
    `Soma das espessuras: Σhi = ${above.reduce((s,l)=>s+l.h,0).toFixed(2)} m ` +
    `${Math.abs(above.reduce((s,l)=>s+l.h,0) - tf) < 0.06 ? '≈ tf ✓' : `≠ tf = ${tf} m — verificar!`}`
  )));
  children.push(BR());

  children.push(H2('1.4 Nível d\'Água'));
  if (zw != null) {
    children.push(P(T(
      'A presença do nível d\'água reduz a tensão efetiva no solo submerso. ' +
      'Quando o lençol freático está dentro da zona de influência da fundação, ' +
      'o peso específico natural γ deve ser substituído pelo peso efetivo (submerso) γ\', ' +
      'dado por γ\' = γ − γw, onde γw = 10 kN/m³ é o peso específico da água. ' +
      'A redução afeta tanto γm (abaixo da fundação) quanto γ\'m (acima da fundação), ' +
      'diminuindo a capacidade de carga nas parcelas de sobrecarga e peso próprio do solo.'
    )));
    children.push(tbl(
      ['Parâmetro', 'Símbolo', 'Valor', 'Unidade'],
      [
        ['Profundidade do NA (desde superfície)', 'zw', zw, 'm'],
        ['Posição relativa à fundação', '—', zw <= tf ? `${n2(tf - zw)} m acima da base` : `${n2(zw - tf)} m abaixo da base`, ''],
        ['Peso específico da água', 'γw', '10', 'kN/m³'],
      ]
    ));
  } else {
    children.push(P(T(
      'Nível d\'água não considerado. Todos os pesos específicos inseridos são utilizados diretamente, ' +
      'sem aplicação de empuxo de água. Caso o lençol freático esteja dentro da zona de influência ' +
      'da fundação (profundidade ≤ b abaixo da base), recomenda-se reanalisar com NA ativo.'
    )));
  }
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 2. FUNDAMENTOS TEÓRICOS
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('2. Fundamentos Teóricos'));

  children.push(H2('2.1 Norma de Referência'));
  children.push(P(T(
    'A DIN 4017:2006-03 — "Baugrund; Berechnung des Grundbruchwiderstands von Flachgründungen" — ' +
    'é a norma alemã para cálculo da capacidade de carga última de fundações superficiais. ' +
    'Ela estabelece o método analítico baseado na teoria da plastificação do solo e no ' +
    'equilíbrio limite, sendo amplamente utilizada em projetos geotécnicos na Alemanha e ' +
    'referenciada em normas europeias (EC7 / EN 1997-1). ' +
    'A norma é válida para fundações superficiais (tf/b ≤ 1) assentes em solos com comportamento ' +
    'Mohr-Coulomb, sujeitas a cargas predominantemente verticais.'
  )));

  children.push(H2('2.2 Hipóteses e Condições de Aplicabilidade'));
  children.push(P(T('O método baseia-se nas seguintes hipóteses fundamentais:')));
  children.push(PI(B('a) Solo rígido-perfeitamente plástico: '),
    T('o solo não apresenta deformações antes da ruptura e, após atingir a resistência de pico, ' +
      'deforma-se indefinidamente sem variação de tensão (critério de Mohr-Coulomb com c e φ constantes).')));
  children.push(PI(B('b) Superfície de ruptura pré-definida: '),
    T('o mecanismo de colapso é do tipo Prandtl, com geometria determinada pelos parâmetros de ' +
      'resistência equivalentes. Não há busca por superfície crítica: a geometria é única para cada φm.')));
  children.push(PI(B('c) Carga vertical centrada: '),
    T('a formulação padrão assume carga sem excentricidade e sem inclinação. ' +
      'Cargas inclinadas ou excêntricas requerem fatores adicionais (di, ic), não incluídos neste cálculo.')));
  children.push(PI(B('d) Terreno e soleira horizontais: '),
    T('o ângulo da superfície do terreno e o ângulo da soleira da fundação são nulos (α = β = δ = 0). ' +
      'Para situações em talude ou fundação inclinada, fatores geométricos adicionais devem ser aplicados.')));
  children.push(PI(B('e) Princípio da superposição: '),
    T('a capacidade de carga última é obtida pela soma de três parcelas independentes ' +
      '(coesão, sobrecarga e peso próprio do solo), o que introduz conservadorismo ' +
      'admitido pela norma como aceitável para uso prático.')));
  children.push(BR());

  children.push(H2('2.3 Mecanismo de Ruptura de Prandtl'));
  children.push(P(T(
    'O mecanismo de ruptura adotado pela DIN 4017 é derivado da solução de Prandtl (1920) ' +
    'para penetração de um punção rígido em material semi-infinito. ' +
    'Para o caso de fundação superficial com carga vertical, o mecanismo é composto por três zonas:'
  )));
  children.push(PI(B('Zona I — Cunha Ativa (Rankine): '),
    T('triângulo rígido logo abaixo da fundação, que se desloca verticalmente com ela. ' +
      'A face inclinada faz ângulo α = 45° + φm/2 com a horizontal. ' +
      'Nesta zona, as tensões principais são vertical (σ1) e horizontal (σ3), ' +
      'e o material está no limiar da ruptura ativa.')));
  children.push(PI(B('Zona II — Zona de Transição (Espiral Logarítmica): '),
    T('setor circular em que os planos de ruptura giram de 90°, conectando a cunha ativa à passiva. ' +
      'A fronteira com o solo vizinho segue uma espiral logarítmica r = r₀·e^(ψ·tan φm), ' +
      'única curva que mantém a condição de Mohr-Coulomb em todos os pontos e permite ' +
      'equilíbrio de momentos sem trabalho líquido dos planos de corte.')));
  children.push(PI(B('Zona III — Cunha Passiva (Rankine): '),
    T('triângulo que emerge lateralmente, com a face inferior inclinada de ϑ = 45° − φm/2 ' +
      'com a horizontal. O solo é comprimido lateralmente e resiste ao deslocamento ' +
      'como na teoria de Rankine passiva.')));
  children.push(P(T(
    'A solução é "one-sided" (semi-mecanismo), pois, para uma fundação longa, a ruptura ' +
    'ocorre de um lado apenas. O mecanismo completo seria simétrico, mas o cálculo da ' +
    'capacidade de carga por superposição usa o semi-mecanismo para cada parcela.'
  )));

  children.push(H2('2.4 Parâmetros Equivalentes para Solo em Camadas'));
  children.push(P(T(
    'Quando o perfil de solo não é homogêneo, a superfície de ruptura atravessa camadas com ' +
    'diferentes parâmetros de resistência. A DIN 4017 §7.3 (Mehrschichtiger Baugrund) propõe ' +
    'substituir o perfil real por um solo homogêneo equivalente, cujos parâmetros φm, cm e γm ' +
    'são determinados pela ponderação das contribuições de cada camada dentro do mecanismo de ruptura.'
  )));
  children.push(P(T(
    'A ideia central é que a mobilização de resistência ao longo da superfície de ruptura ' +
    'é proporcional ao comprimento de linha de ruptura em cada camada (φm, cm) ' +
    'e ao volume de solo mobilizado em cada camada (γm). ' +
    'Como a geometria da superfície de ruptura depende do próprio φm, ' +
    'a determinação de φm requer um processo iterativo.'
  )));

  // ══════════════════════════════════════════════════════════════════════════
  // 3. GEOMETRIA DA SUPERFÍCIE DE RUPTURA
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('3. Geometria da Superfície de Ruptura'));

  children.push(H2('3.1 Definição Geométrica'));
  children.push(P(T(
    'O mecanismo de ruptura é totalmente definido pelo ângulo de atrito equivalente φm e ' +
    'pela largura da fundação b. A origem do sistema de coordenadas está no centro da base ' +
    'da fundação (x = 0, y = 0), com x positivo para direita e y positivo para cima.'
  )));
  children.push(P(T(
    'O centro da espiral logarítmica C está na borda direita da fundação, ' +
    'ao nível da base (C = b/2, 0). A partir deste ponto, a espiral parte ' +
    'do vértice da cunha ativa P2 e termina no ponto P5, onde começa a cunha passiva.'
  )));

  children.push(H2('3.2 Ângulos das Cunhas (DIN 4017, Anexo A)'));
  children.push(P(T(
    'Para condição padrão (sem inclinação de carga, sem inclinação de terreno): ' +
    'o ângulo α da face da cunha ativa com a horizontal é 45° + φm/2, ' +
    'e o ângulo ϑ da face da cunha passiva é 45° − φm/2. ' +
    'O setor de transição subtende exatamente 90° (ν = π/2):'
  )));

  children.push(EQ(
    MR('α = 45° + '),
    MF([MR('φ'), MSub([MR('')],[MR('m')])],[MR('2')]),
    MR(' = 45° + '),
    MF([MR(`${n2(R.phiM)}°`)],[MR('2')]),
    MR(` = ${n2(45 + R.phiM/2)}°`)
  ));

  children.push(EQ(
    MR('ϑ = 45° − '),
    MF([MR('φ'), MSub([MR('')],[MR('m')])],[MR('2')]),
    MR(' = 45° − '),
    MF([MR(`${n2(R.phiM)}°`)],[MR('2')]),
    MR(` = ${n2(45 - R.phiM/2)}°`)
  ));

  children.push(EQ(MR('ν = 90°  (setor de transição)') ));

  children.push(H2('3.3 Raios da Espiral Logarítmica'));
  children.push(P(T(
    'O raio inicial r₂ é o comprimento C→P2 (do centro ao vértice da cunha ativa). ' +
    'Após girar 90° em torno de C, o raio cresce para r₁ = C→P5 (fim da espiral). ' +
    'A razão r₁/r₂ = e^(π/2·tan φm) depende exclusivamente de φm:'
  )));

  children.push(EQ(
    MSub([MR('r')],[MR('2')]),
    MR(' = '),
    MF([MR("b")],[MR('2 · cos α')]),
    MR(' = '),
    MF([MR(`${b}`)],[MR(`2 · cos ${n2(45+R.phiM/2)}°`)]),
    MR(` = ${n2(geom.r2)} m`)
  ));

  children.push(EQ(
    MSub([MR('r')],[MR('1')]),
    MR(' = '),
    MSub([MR('r')],[MR('2')]),
    MR(' · '),
    MSup([MR('e')],[
      MF([MR('π')],[MR('2')]),
      MR(' · tan φ'),
      MSub([MR('')],[MR('m')])
    ]),
    MR(` = ${n2(geom.r2)} · e^(π/2 · tan ${n2(R.phiM)}°) = ${n2(geom.r1)} m`)
  ));

  children.push(P(T(
    `A espiral logarítmica é descrita por r(ψ) = r₂ · e^(ψ · tan φm), com ψ ∈ [0°, 90°]. ` +
    `A profundidade máxima atingida pela superfície de ruptura ocorre quando a tangente à espiral ` +
    `é horizontal, e vale ${n2(-geom.maxDepth)} m abaixo da base da fundação. ` +
    `Este valor determina a profundidade de influência do mecanismo e, portanto, ` +
    `quais camadas participam dos cálculos de φm, cm e γm.`
  )));

  children.push(tbl(
    ['Ponto / Parâmetro', 'Coordenadas / Valor', 'Significado'],
    [
      ['C (centro espiral)', `(${n2(b/2)}, 0) m`, 'Borda direita da fundação, base'],
      ['P2 (vértice cunha ativa)', `(0 ; ${n2(geom.yP2)} m)`, 'Vértice abaixo da fundação no eixo de simetria'],
      ['P5 (fim da espiral)', `(${n2(geom.P5.x)} ; ${n2(geom.P5.y)} m)`, 'Início da cunha passiva'],
      ['P6 (na superfície)', `(${n2(geom.xP6)} ; ≈ 0 m)`, 'Extremidade da cunha passiva na superfície'],
      ['r₂', `${n2(geom.r2)} m`, 'Raio inicial da espiral (C → P2)'],
      ['r₁', `${n2(geom.r1)} m`, 'Raio final da espiral (C → P5)'],
      ['Profundidade máxima', `${n2(-geom.maxDepth)} m`, 'Ponto mais fundo da superfície de ruptura'],
    ]
  ));
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 4. ÂNGULO DE ATRITO EQUIVALENTE φm
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('4. Determinação do Ângulo de Atrito Equivalente φm'));

  children.push(H2('4.1 Fundamento Teórico'));
  children.push(P(T(
    'Em um solo homogêneo, o mecanismo de Prandtl é governado por um único ângulo de atrito φ. ' +
    'Em solo estratificado, a superfície de ruptura atravessa camadas com φi distintos, ' +
    'e nenhum valor único é diretamente aplicável. ' +
    'A DIN 4017 propõe definir φm como o ângulo de atrito de um solo fictício homogêneo ' +
    'que mobiliza a mesma resistência tangencial total ao longo de toda a superfície de ruptura.'
  )));
  children.push(P(T(
    'A resistência tangencial mobilizada em um segmento de comprimento li dentro da camada i é ' +
    'proporcional à tensão normal local e a tan φi. A ponderação por comprimento de linha ' +
    'é a aproximação adotada pela DIN 4017, que integra a contribuição de cada camada ' +
    'de forma linear pelo comprimento de ruptura contido nela:'
  )));

  children.push(EQ(
    MR('tan '),
    MSub([MR('φ')],[MR('m')]),
    MR(' = '),
    MF(
      [MSub([MR('Σ  l')],[MR('i')]), MR(' · tan '), MSub([MR('φ')],[MR('i')])],
      [MSub([MR('Σ  l')],[MR('i')])]
    )
  ));

  children.push(H2('4.2 Necessidade de Processo Iterativo'));
  children.push(P(T(
    'O problema é implícito: para calcular os comprimentos li é preciso conhecer a geometria ' +
    'da superfície de ruptura, que por sua vez depende de φm. ' +
    'O processo iterativo parte de um φm inicial (tipicamente o maior φi das camadas), ' +
    'calcula os comprimentos li com essa geometria, obtém um novo φm,out, ' +
    'e repete usando a média φm,novo = (φm,in + φm,out)/2 até convergir.'
  )));
  children.push(P(T(
    'A DIN 4017 §7.3 estabelece o critério de convergência Δ < 3%, onde:'
  )));
  children.push(EQ(
    MR('Δ = '),
    MF(
      [MR('|'), MSub([MR('φ')],[MR('m,in')]), MR(' − '), MSub([MR('φ')],[MR('m,out')]), MR('|')],
      [MSub([MR('φ')],[MR('m,in')])]
    ),
    MR(' × 100%  <  3%')
  ));

  children.push(H2('4.3 Desenvolvimento das Iterações'));
  children.push(P(T(
    `Valor inicial: φm,1 = max(φi) = ${layers.reduce((m,l)=>Math.max(m,l.phi),0).toFixed(1)}° ` +
    `(maior ângulo de atrito entre as camadas).`
  )));

  children.push(tbl(
    ['Iteração', 'φm,in (°)', 'φm,out (°)', 'Δ (%)', 'φm,novo (°)', 'Convergiu?'],
    R.iterations.map(it => [
      it.iter,
      n4(it.phiIn),
      n4(it.phiOut),
      n2(it.delta),
      it.delta < 3 ? '—' : n4((it.phiIn + it.phiOut) / 2),
      it.delta < 3 ? 'Sim ✓' : 'Não'
    ])
  ));
  children.push(BR());

  const lastIt = R.iterations[R.iterations.length - 1];
  children.push(P(
    B('φm convergido = '),
    T(`(${n4(lastIt.phiIn)}° + ${n4(lastIt.phiOut)}°) / 2 = `),
    B(`${n4(R.phiM)}°`)
  ));

  children.push(H2('4.4 Comprimentos da Linha de Ruptura por Camada'));
  children.push(P(T(
    `Calculados com φm = ${n2(R.phiM)}°. ` +
    'A linha de ruptura total é dividida em três segmentos: ' +
    'cunha ativa (segmento reto), espiral logarítmica (arco) e cunha passiva (segmento reto). ' +
    'A fração de cada segmento dentro de cada camada é determinada pelas profundidades das fronteiras.'
  )));
  children.push(tbl(
    ['Camada', 'φi (°)', 'li (m)', 'li / Σli (%)', 'li · tan φi'],
    layers.map((l, k) => [
      l.desc || `Camada ${k+1}`,
      l.phi,
      n3(R.lengths[k]),
      n2(R.lengths[k] / R.lTotal * 100),
      n4(R.lengths[k] * Math.tan(l.phi * Math.PI / 180))
    ]).concat([['Total', '—', n3(R.lTotal), '100,00', n4(layers.reduce((s,l,k)=>s+R.lengths[k]*Math.tan(l.phi*Math.PI/180),0))]])
  ));
  children.push(BR());

  const sumLtanPhi = layers.reduce((s,l,k) => s + R.lengths[k] * Math.tan(l.phi * Math.PI / 180), 0);
  children.push(P(
    T('Verificação: tan φm = '),
    T(`${n4(sumLtanPhi)} / ${n3(R.lTotal)} = ${n4(sumLtanPhi/R.lTotal)}  →  φm = arctan(${n4(sumLtanPhi/R.lTotal)}) = `),
    B(`${n4(Math.atan(sumLtanPhi/R.lTotal)*180/Math.PI)}°  ✓`)
  ));
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 5. COESÃO EQUIVALENTE cm
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('5. Determinação da Coesão Equivalente cm'));

  children.push(H2('5.1 Fundamento Teórico'));
  children.push(P(T(
    'A coesão equivalente cm representa a coesão média mobilizada ao longo de toda a superfície ' +
    'de ruptura. Assim como φm, é obtida pela ponderação linear dos valores ci de cada camada ' +
    'pelo comprimento li da linha de ruptura dentro dela. ' +
    'A ponderação por comprimento é fisicamente consistente: em uma seção de ruptura ' +
    'de comprimento li em material coesivo ci, a força resistente de coesão vale ci · li ' +
    '(por unidade de profundidade na direção perpendicular ao plano de análise).'
  )));
  children.push(P(T(
    'A mesma geometria (comprimentos li) obtida para φm é reutilizada para cm, ' +
    'pois φm e cm descrevem o mesmo solo equivalente e devem ser consistentes ' +
    'com a mesma superfície de ruptura:'
  )));

  children.push(EQ(
    MSub([MR('c')],[MR('m')]),
    MR(' = '),
    MF(
      [MSub([MR('Σ  l')],[MR('i')]), MR(' · '), MSub([MR('c')],[MR('i')])],
      [MSub([MR('Σ  l')],[MR('i')])]
    )
  ));

  children.push(H2('5.2 Desenvolvimento Numérico'));
  const cmNum = layers.map((l, k) => `${n3(R.lengths[k])} × ${l.c}`).join(' + ');
  children.push(P(
    T(`cm = (${cmNum}) / ${n3(R.lTotal)}`),
    T(layers.every(l => l.c === 0) ? ' = 0 (solo sem coesão)' : '')
  ));
  children.push(P(B(`cm = ${n2(R.cm)} kN/m²`)));

  if (layers.every(l => l.c === 0)) {
    children.push(P(I(
      'Observação: todas as camadas têm c = 0, portanto cm = 0 e a parcela de coesão na ' +
      'equação de qult é nula. Em solos granulares (areias, cascalhos), esta é a situação típica.'
    )));
  } else {
    children.push(P(I(
      `Observação: a parcela de coesão contribui com cm·Nc·νc = ${n2(R.cm)} × ${n2(R.Nc)} × ${n3(R.nuc)} = ${n2(R.term_c)} kN/m² ` +
      `(${n2(R.term_c/R.qult*100)}% de qult).`
    )));
  }
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 6. PESO ESPECÍFICO MÉDIO ABAIXO DA FUNDAÇÃO γm
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('6. Determinação do Peso Específico Médio γm'));

  children.push(H2('6.1 Fundamento Teórico'));
  children.push(P(T(
    'O peso específico γm representa a inércia média do maciço de solo contido dentro da ' +
    'superfície de ruptura. Ao contrário de φm e cm — cujas equações de equilíbrio envolvem ' +
    'forças distribuídas ao longo da linha de ruptura —, a parcela de peso próprio do solo ' +
    'envolve o peso do volume mobilizado. Por isso γm é ponderado pelas áreas Ai ' +
    '(área da seção transversal do mecanismo dentro de cada camada, por unidade de comprimento):'
  )));

  children.push(EQ(
    MSub([MR('γ')],[MR('m')]),
    MR(' = '),
    MF(
      [MSub([MR('Σ  A')],[MR('i')]), MR(' · '), MSub([MR('γ')],[MR('i')])],
      [MSub([MR('Σ  A')],[MR('i')])]
    )
  ));

  children.push(P(T(
    'As áreas são calculadas por integração numérica (método das faixas horizontais): ' +
    'para cada profundidade z, determina-se a largura do mecanismo entre o limite esquerdo ' +
    '(face da cunha ativa) e o limite direito (espiral ou cunha passiva), ' +
    'e acumula-se a área em cada camada.'
  )));

  if (zw != null) {
    children.push(P(T(
      `Com nível d'água em zw = ${zw} m desde a superfície (${n2(zw-tf)} m desde a base da fundação), ` +
      'as camadas total ou parcialmente submersas recebem peso específico efetivo ' +
      "γ'i = γi − 10 kN/m³ na integração. A tabela abaixo mostra o peso efetivo adotado para cada camada:"
    )));
  }

  children.push(H2('6.2 Áreas e Pesos por Camada'));
  children.push(tbl(
    zw != null
      ? ['Camada', 'γi natural (kN/m³)', "γ'i efetivo (kN/m³)", 'Ai (m²)', "Ai·γ'i (kN/m)"]
      : ['Camada', 'γi (kN/m³)', 'Ai (m²)', 'Ai·γi (kN/m)'],
    layers.map((l, k) => zw != null
      ? [l.desc || `Camada ${k+1}`, l.gamma, n2(gammaEffArr[k]), n3(R.areas[k]), n3(R.areas[k] * gammaEffArr[k])]
      : [l.desc || `Camada ${k+1}`, l.gamma, n3(R.areas[k]), n3(R.areas[k] * l.gamma)]
    ).concat([['Total', '—', ...(zw != null ? ['—'] : []), n3(R.aTotal),
      n3(layers.reduce((s,l,k) => s + R.areas[k] * gammaEffArr[k], 0))]])
  ));
  children.push(BR());
  children.push(P(
    T(`γm = ${n3(layers.reduce((s,l,k)=>s+R.areas[k]*gammaEffArr[k],0))} / ${n3(R.aTotal)} = `),
    B(`${n2(R.gammaM)} kN/m³`)
  ));
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 7. PESO ESPECÍFICO ACIMA DA FUNDAÇÃO γ'm
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1("7. Determinação do Peso Específico Acima da Fundação γ'm"));

  children.push(H2('7.1 Fundamento Teórico'));
  children.push(P(T(
    'O solo localizado lateralmente à fundação, entre a superfície do terreno e o nível de ' +
    'assentamento (camadas acima da fundação), atua como sobrecarga distribuída q = γ\'m · tf ' +
    'sobre o plano de assentamento. Esta sobrecarga é a condição de contorno de tensão ' +
    'na fronteira do mecanismo de Prandtl, e aparece na equação de capacidade de carga ' +
    'multiplicada pelo fator de sobrecarga Nd.'
  )));
  children.push(P(T(
    'O peso específico médio γ\'m é calculado como média ponderada pela espessura das camadas acima, ' +
    'substituindo γ por γ\' = γ − 10 nas parcelas submersas quando há nível d\'água:'
  )));

  children.push(EQ(
    MSup([MR("γ'")],[MR('')]),
    MSub([MR('')],[MR('m')]),
    MR(' = '),
    MF(
      [MSub([MR('Σ  h')],[MR('i')]), MR(' · '), MSub([MR('γ')],[MR('i')])],
      [MR('t'), MSub([MR('')],[MR('f')])]
    )
  ));

  children.push(H2('7.2 Desenvolvimento Numérico'));
  if (zw != null) {
    let cumZ2 = 0;
    const aboveRows = above.map(l => {
      const zTop = cumZ2, zBot = cumZ2 + l.h;
      cumZ2 = zBot;
      let gEff = l.gamma;
      if (zw < zBot) {
        const zWtInLayer = Math.max(zw, zTop);
        const hWet = zBot - zWtInLayer;
        const hDry = l.h - hWet;
        gEff = (hDry * l.gamma + hWet * (l.gamma - 10)) / l.h;
      }
      return [l.desc || `Sub-camada`, n2(zTop), n2(zBot), l.gamma, n2(gEff), l.h];
    });
    children.push(tbl(
      ['Camada', 'z topo (m)', 'z fundo (m)', 'γ natural (kN/m³)', "γ' efetivo (kN/m³)", 'h (m)'],
      aboveRows
    ));
    children.push(P(T(`γ'm ponderado com efeito do NA = `), B(`${n2(R.gammaMAbove)} kN/m³`)));
  } else {
    const gammaAboveNum = above.map(l => `${l.h} × ${l.gamma}`).join(' + ');
    children.push(P(T(`γ'm = (${gammaAboveNum}) / ${tf} = `), B(`${n2(R.gammaMAbove)} kN/m³`)));
  }
  children.push(P(I(
    `A sobrecarga equivalente atuante no nível de fundação é: q = γ'm · tf = ${n2(R.gammaMAbove)} × ${tf} = ${n2(R.gammaMAbove * tf)} kN/m².`
  )));
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 8. FATORES DE CAPACIDADE DE CARGA
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('8. Fatores de Capacidade de Carga'));

  children.push(H2('8.1 Origem e Significado Físico'));
  children.push(P(T(
    'Os fatores de capacidade de carga Nd, Nc e Nb são adimensionais e representam a ' +
    'amplificação da resistência do solo pelo mecanismo de ruptura de Prandtl. ' +
    'Cada fator corresponde a uma das três parcelas da equação de qult:'
  )));
  children.push(PI(B('Nd (fator de sobrecarga — Reissner, 1924): '),
    T('amplifica a sobrecarga lateral q = γ\'m·tf. Fisicamente, representa quantas vezes a ' +
      'tensão de confinamento lateral é transformada em capacidade de carga vertical ' +
      'pelo mecanismo de Prandtl. Para φm = 0°, Nd = 1 (sem amplificação).')));
  children.push(PI(B('Nc (fator de coesão — Prandtl, 1920): '),
    T('amplifica a coesão equivalente cm. Para φm = 0°, Nc = π + 2 ≈ 5,14 (solução analítica de Prandtl). ' +
      'O fator Nc e Nd estão relacionados por Nc = (Nd − 1)/tan φm (exceto para φm = 0).')));
  children.push(PI(B('Nb (fator de peso próprio — Caquot e Kérisel, 1953): '),
    T('amplifica o termo de peso próprio do solo γm·b. ' +
      'Não possui solução analítica exata; a DIN 4017 adota a expressão aproximada Nb = (Nd − 1)·tan φm, ' +
      'que fornece boa concordância com soluções numéricas.')));

  children.push(H2('8.2 Fórmulas (DIN 4017:2006-03, Equações 5–7)'));

  children.push(EQ(
    MSub([MR('N')],[MR('d')]),
    MR(' = '),
    MSup([MR('tan²(45° + '),MF([MR('φ'),MSub([MR('')],[MR('m')])],[MR('2')]),MR(')')],[MR('')]),
    MR(' · '),
    MSup([MR('e')],[MR('π · tan φ'),MSub([MR('')],[MR('m')])])
  ));

  children.push(EQ(
    MSub([MR('N')],[MR('c')]),
    MR(' = '),
    MF([MSub([MR('N')],[MR('d')]),MR(' − 1')],[MR('tan φ'),MSub([MR('')],[MR('m')])])
  ));

  children.push(EQ(
    MSub([MR('N')],[MR('b')]),
    MR(' = ('),MSub([MR('N')],[MR('d')]),MR(' − 1) · tan φ'),MSub([MR('')],[MR('m')])
  ));

  children.push(H2(`8.3 Valores Calculados para φm = ${n2(R.phiM)}°`));
  children.push(tbl(
    ['Fator', 'Expressão', 'Valor', 'Significado'],
    [
      ['Nd', `tan²(${n2(45+R.phiM/2)}°) · e^(π·tan ${n2(R.phiM)}°)`, n2(R.Nd), 'Fator de sobrecarga'],
      ['Nc', `(${n2(R.Nd)} − 1) / tan ${n2(R.phiM)}°`, n2(R.Nc), 'Fator de coesão'],
      ['Nb', `(${n2(R.Nd)} − 1) · tan ${n2(R.phiM)}°`, n2(R.Nb), 'Fator de peso próprio'],
    ]
  ));
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 9. FATORES DE FORMA
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('9. Fatores de Forma'));

  children.push(H2('9.1 Fundamento Teórico'));
  children.push(P(T(
    'Os fatores de capacidade de carga Nd, Nc e Nb são derivados para fundações contínuas ' +
    '(estado plano de deformação, b/a → 0). Em fundações de comprimento finito, ' +
    'o solo nas extremidades da fundação oferece confinamento adicional, ' +
    'aumentando Nd e Nc (efeito 3D benéfico), enquanto o peso próprio do solo ' +
    'nas extremidades é menor (efeito 3D desfavorável para Nb).'
  )));
  children.push(P(T(
    'A DIN 4017 Tabelle 2 fornece fatores de forma νd, νc e νb para fundações retangulares ' +
    '(Rechteckfundamente), em função de b/a e φm. ' +
    'Para b/a → 0 (fundação longa), todos os fatores tendem a 1,0. ' +
    'Para b/a = 1 (fundação quadrada), os fatores assumem valores máximos (νd, νc) ou mínimos (νb).'
  )));

  children.push(H2('9.2 Fórmulas (DIN 4017:2006-03, Tabelle 2 — Rechteck)'));

  children.push(EQ(
    MSub([MR('ν')],[MR('b')]),
    MR(' = 1 − 0,3 · '),
    MF([MR("b")],[MR("a")])
  ));

  children.push(EQ(
    MSub([MR('ν')],[MR('d')]),
    MR(' = 1 + '),
    MF([MR("b")],[MR("a")]),
    MR(' · sin φ'),
    MSub([MR('')],[MR('m')])
  ));

  children.push(EQ(
    MSub([MR('ν')],[MR('c')]),
    MR(' = '),
    MF(
      [MSub([MR('ν')],[MR('d')]),MR(' · '),MSub([MR('N')],[MR('d')]),MR(' − 1')],
      [MSub([MR('N')],[MR('d')]),MR(' − 1')]
    )
  ));

  children.push(H2(`9.3 Valores Calculados para b/a = ${n3(b/a)}, φm = ${n2(R.phiM)}°`));
  children.push(tbl(
    ['Fator', 'Expressão numérica', 'Valor', 'Efeito vs. fundação longa'],
    [
      ['νb', `1 − 0,3 × ${n3(b/a)} = `, n3(R.nub), R.nub < 1 ? `Redução de ${n2((1-R.nub)*100)}%` : 'Sem redução'],
      ['νd', `1 + ${n3(b/a)} × sin ${n2(R.phiM)}° = `, n3(R.nud), `Aumento de ${n2((R.nud-1)*100)}%`],
      ['νc', `(${n3(R.nud)} × ${n2(R.Nd)} − 1) / (${n2(R.Nd)} − 1) = `, n3(R.nuc), `Aumento de ${n2((R.nuc-1)*100)}%`],
    ]
  ));
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 10. CAPACIDADE DE CARGA ÚLTIMA qult
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('10. Capacidade de Carga Última qult'));

  children.push(H2('10.1 Equação Geral (DIN 4017:2006-03, Equação 1)'));
  children.push(P(T(
    'A capacidade de carga última é obtida pela superposição de três parcelas, ' +
    'cada uma associada a um mecanismo de resistência independente. ' +
    'Embora a superposição seja uma aproximação (os mecanismos não são rigorosamente independentes), ' +
    'ela é conservativa e estabelecida como procedimento padrão pela DIN 4017:'
  )));

  children.push(EQ(
    MSub([MR('q')],[MR('ult')]),
    MR(' = '),
    MSub([MR('c')],[MR('m')]),
    MR(' · '),MSub([MR('N')],[MR('c')]),MR(' · '),MSub([MR('ν')],[MR('c')]),
    MR('  +  '),
    MSup([MR("γ'")],[MR('')]),MSub([MR('')],[MR('m')]),
    MR(' · '),MSub([MR('t')],[MR('f')]),MR(' · '),MSub([MR('N')],[MR('d')]),MR(' · '),MSub([MR('ν')],[MR('d')]),
    MR('  +  '),
    MSub([MR('γ')],[MR('m')]),
    MR(' · b · '),MSub([MR('N')],[MR('b')]),MR(' · '),MSub([MR('ν')],[MR('b')])
  ));

  children.push(H2('10.2 Significado Físico das Parcelas'));
  children.push(PI(B('Parcela I — Coesão (cm · Nc · νc): '),
    T('resistência intrínseca da matriz do solo. Ativa mesmo sem peso próprio ou sobrecarga. ' +
      'Fundamental em argilas saturadas (φ = 0), onde toda a resistência vem de cu = cm.')));
  children.push(PI(B('Parcela II — Sobrecarga (γ\'m · tf · Nd · νd): '),
    T('efeito de confinamento exercido pelo solo ao redor da fundação. ' +
      'Quanto mais profunda a fundação, maior tf e maior a contribuição desta parcela. ' +
      'É a única parcela presente mesmo quando c = 0 e γm = 0.')));
  children.push(PI(B('Parcela III — Peso próprio do solo (γm · b · Nb · νb): '),
    T('resistência associada ao peso do solo dentro da zona de ruptura. ' +
      'Proporcional à largura b: fundações maiores mobilizam maior volume de solo. ' +
      'Esta parcela é a mais incerta, pois depende da integração numérica de γm.')));

  children.push(H2('10.3 Desenvolvimento Numérico'));
  children.push(tbl(
    ['Parcela', 'Expressão', 'Cálculo', 'Valor (kN/m²)', '% de qult'],
    [
      ['Coesão',     'cm · Nc · νc',     `${n2(R.cm)} × ${n2(R.Nc)} × ${n3(R.nuc)}`,     n2(R.term_c), n2(R.term_c/R.qult*100)+'%'],
      ['Sobrecarga', "γ'm · tf · Nd · νd", `${n2(R.gammaMAbove)} × ${tf} × ${n2(R.Nd)} × ${n3(R.nud)}`, n2(R.term_q), n2(R.term_q/R.qult*100)+'%'],
      ['Peso próprio','γm · b · Nb · νb', `${n2(R.gammaM)} × ${b} × ${n2(R.Nb)} × ${n3(R.nub)}`,  n2(R.term_b), n2(R.term_b/R.qult*100)+'%'],
    ]
  ));
  children.push(BR());

  children.push(new D.Paragraph({
    children: [
      T('qult = ', { size: 26, bold: true }),
      T(`${n2(R.term_c)}  +  ${n2(R.term_q)}  +  ${n2(R.term_b)}  = `, { size: 24 }),
      T(`${R.qult.toFixed(0)} kN/m²`, { size: 32, bold: true, color: '1a3a5c' })
    ],
    alignment: D.AlignmentType.CENTER,
    spacing: { before: 160, after: 160 }
  }));

  children.push(H2('10.4 Observações sobre a Capacidade de Carga Admissível'));
  children.push(P(T(
    'A capacidade de carga última qult obtida não inclui coeficientes de segurança. ' +
    'Para obter a tensão admissível de projeto, devem ser aplicados os fatores de segurança ' +
    'conforme a norma aplicável ao projeto:'
  )));
  children.push(PI(T('DIN 1054:2010 (GEO): verificação por estados limites últimos (ELU/GEO), com fatores parciais aplicados às cargas e às resistências.')));
  children.push(PI(T('ABNT NBR 6122:2022: capacidade de carga admissível = qult / FS, com FS ≥ 3,0 para fundações em serviço (Método Global).')));
  children.push(PI(T('EC7 / EN 1997-1: abordagem de projeto por Ações de Cálculo (DA1, DA2 ou DA3), com coeficientes γR para resistências de fundação.')));
  children.push(BR());

  // ══════════════════════════════════════════════════════════════════════════
  // 11. REFERÊNCIAS
  // ══════════════════════════════════════════════════════════════════════════
  children.push(H1('11. Referências'));
  children.push(P(T('DIN 4017:2006-03 — Baugrund; Berechnung des Grundbruchwiderstands von Flachgründungen. Deutsches Institut für Normung, 2006.')));
  children.push(P(T('DIN 1054:2010-12 — Baugrund — Sicherheitsnachweise im Erd- und Grundbau — Ergänzende Regelungen zu DIN EN 1997-1. Deutsches Institut für Normung, 2010.')));
  children.push(P(T('Prandtl, L. (1920). Über die Eindringungsfestigkeit plastischer Baustoffe und die Festigkeit von Schneiden. ZAMM, 1(1), 15–20.')));
  children.push(P(T('Reissner, H. (1924). Zum Erddruckproblem. Proc. 1st Int. Congress for Applied Mechanics, Delft, 295–311.')));
  children.push(P(T('Caquot, A. & Kérisel, J. (1953). Sur le terme de surface dans le calcul des fondations en milieu pulvérulent. Proc. 3rd ICSMFE, Zürich, Vol. 1, 336–337.')));
  children.push(P(T('ELPLA — Software for Analysis of Pile and Slab Foundations. Verification Examples. Example 11: DIN 4017 layered soil.')));
  children.push(P(T('ABNT NBR 6122:2022 — Projeto e execução de fundações. Associação Brasileira de Normas Técnicas, 2022.')));
  children.push(BR());

  // ── Rodapé ───────────────────────────────────────────────────────────────
  children.push(new D.Paragraph({
    children: [T('─────────────────────────────────────────────────────', { color: 'cccccc' })],
    spacing: { before: 200 }
  }));
  children.push(new D.Paragraph({
    children: [T(
      'Calculado por din4017-bearing-capacity · pedrorussini.github.io/din4017-bearing-capacity',
      { size: 16, color: '999999', italics: true }
    )],
    alignment: D.AlignmentType.CENTER
  }));

  // ── Gerar e baixar ────────────────────────────────────────────────────────
  const doc = new D.Document({
    sections: [{ children }],
    styles: {
      default: { document: { run: { font: 'Calibri', size: 22 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', run: { bold: true, size: 28, color: '1a3a5c' } },
        { id: 'Heading2', name: 'Heading 2', run: { bold: true, size: 24, color: '1a4a7a' } },
        { id: 'Heading3', name: 'Heading 3', run: { bold: true, size: 22, color: '2e6da4' } },
      ]
    }
  });

  const blob = await D.Packer.toBlob(doc);
  const url  = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href  = url;
  link.download = `memorial_din4017_${new Date().toISOString().slice(0,10)}.docx`;
  link.click();
  URL.revokeObjectURL(url);
}
