/**
 * memorial.js — Geração de Memória de Cálculo em Word (.docx)
 * Equações no formato OMML (nativo do Microsoft Word).
 * Norma: DIN 4017:2006-03
 */

// Estado global (preenchido por showResults em ui.js)
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
    spacing: { before: 300, after: 120 } });
  const H2 = (text) => new D.Paragraph({ text, heading: D.HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 } });
  const P  = (...runs) => new D.Paragraph({ children: runs, spacing: { after: 80 } });
  const T  = (text, opts = {}) => new D.TextRun({ text, ...opts });
  const B  = (text) => new D.TextRun({ text, bold: true });
  const BR = () => new D.Paragraph({ spacing: { after: 40 } });

  // Equação em bloco (centralizada)
  const EQ = (...children) => new D.Paragraph({
    children: [new D.Math({ children })],
    alignment: D.AlignmentType.CENTER,
    spacing: { before: 80, after: 80 }
  });

  // Equação inline
  const EQI = (...children) => new D.Math({ children });

  // Helpers OMML
  const MR  = (t) => new D.MathRun(t);
  const MF  = (num, den) => new D.MathFraction({ numerator: num, denominator: den });
  const MSup = (base, sup) => new D.MathSuperScript({ children: base, superScript: sup });
  const MSub = (base, sub) => new D.MathSubScript({ children: base, subScript: sub });
  const MSSub = (base, sup, sub) => new D.MathSubSuperScript({ children: base, superScript: sup, subScript: sub });

  // Tabela simples
  function tbl(headers, rows) {
    const borderOpts = {
      style: D.BorderStyle.SINGLE, size: 4, color: '888888'
    };
    const borders = { top: borderOpts, bottom: borderOpts, left: borderOpts, right: borderOpts,
                      insideHorizontal: borderOpts, insideVertical: borderOpts };
    const mkCell = (text, header = false) => new D.TableCell({
      children: [new D.Paragraph({
        children: [new D.TextRun({ text: String(text), bold: header, size: header ? 20 : 18 })],
        alignment: D.AlignmentType.CENTER
      })],
      borders,
      shading: header ? { fill: 'D0D9E8' } : undefined
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

  // ─── Construção do documento ──────────────────────────────────────────────
  const children = [];

  // ── Título ─────────────────────────────────────────────────────────────
  children.push(new D.Paragraph({
    children: [T('MEMÓRIA DE CÁLCULO', { bold: true, size: 32 })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 60 }
  }));
  children.push(new D.Paragraph({
    children: [T('Capacidade de Carga Última em Solo em Camadas', { size: 24, italics: true })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 60 }
  }));
  children.push(new D.Paragraph({
    children: [T('Norma: DIN 4017:2006-03', { size: 20, color: '444444' })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 60 }
  }));
  children.push(new D.Paragraph({
    children: [T(`Data: ${new Date().toLocaleDateString('pt-BR')}`, { size: 18, color: '666666' })],
    alignment: D.AlignmentType.CENTER, spacing: { after: 200 }
  }));

  // ── 1. Dados de Entrada ─────────────────────────────────────────────────
  children.push(H1('1. Dados de Entrada'));
  children.push(H2('1.1 Geometria da Fundação'));
  children.push(tbl(
    ['Parâmetro', 'Símbolo', 'Valor', 'Unidade'],
    [
      ['Largura', 'b', b, 'm'],
      ['Comprimento', 'a', a, 'm'],
      ['Profundidade de fundação', 'tf', tf, 'm'],
    ]
  ));
  children.push(BR());

  children.push(H2('1.2 Camadas Abaixo da Fundação'));
  children.push(tbl(
    ['#', 'Descrição', 'h (m)', 'γ (kN/m³)', 'φ (°)', 'c (kN/m²)'],
    layers.map((l, i) => [i+1, l.desc || `Camada ${i+1}`, l.h, l.gamma, l.phi, l.c])
  ));
  children.push(BR());

  children.push(H2('1.3 Camadas Acima da Fundação'));
  children.push(tbl(
    ['#', 'Descrição', 'h (m)', 'γ (kN/m³)'],
    above.map((l, i) => [i+1, l.desc || `Sub-camada ${i+1}`, l.h, l.gamma])
  ));
  children.push(BR());

  children.push(H2('1.4 Nível d\'Água'));
  if (zw != null) {
    children.push(tbl(
      ['Parâmetro', 'Símbolo', 'Valor', 'Unidade'],
      [
        ['Profundidade do NA (desde superfície)', 'zw', zw, 'm'],
        ['Peso específico da água', 'γw', '10', 'kN/m³'],
      ]
    ));
    children.push(P(T(
      `Peso efetivo aplicado: γ' = γ − γw. Camadas com topo a partir de ${zw} m de profundidade ` +
      `recebem peso específico reduzido.`
    )));
  } else {
    children.push(P(T('Nível d\'água não considerado (solo seco).')));
  }
  children.push(BR());

  // ── 2. Metodologia ──────────────────────────────────────────────────────
  children.push(H1('2. Metodologia'));
  children.push(P(T(
    'O método iterativo da DIN 4017:2006-03 determina parâmetros de resistência equivalentes ' +
    '(φm, cm, γm) para um perfil de solo em camadas, a partir da geometria da superfície de ruptura. ' +
    'O mecanismo de ruptura é do tipo Prandtl, composto por uma cunha ativa, uma zona de transição ' +
    'em espiral logarítmica e uma cunha passiva (caso α = β = δ = 0).'
  )));

  // ── 3. Geometria da Superfície de Ruptura ────────────────────────────────
  children.push(H1('3. Geometria da Superfície de Ruptura'));
  children.push(P(T(
    'Para carga vertical, terreno horizontal e soleira horizontal (α = β = δ = 0), ' +
    'os ângulos das cunhas e o arco da espiral são dados por (DIN 4017 Anexo A, Eqs. A.10–A.13):'
  )));

  // ϑ = 45 − φm/2
  children.push(EQ(
    MSub([MR('ϑ')], [MR('1')]),
    MR(' = '),
    MSub([MR('ϑ')], [MR('2')]),
    MR(' = '),
    MSub([MR('ϑ')], [MR('3')]),
    MR(' = 45° − '),
    MF([MR('φ'), MSub([MR('')], [MR('m')])], [MR('2')])
  ));

  // ν = 90°
  children.push(EQ(MR('ν = 90°')));

  // r2 = b / (2·cos(45+φm/2))
  children.push(EQ(
    MSub([MR('r')], [MR('2')]),
    MR(' = '),
    MF([MR("b'")], [MR('2 · cos(45° + '),
      MF([MR('φ'), MSub([MR('')], [MR('m')])], [MR('2')]),
      MR(')')
    ])
  ));

  // r1 = r2 · e^(π/2 · tan φm)
  children.push(EQ(
    MSub([MR('r')], [MR('1')]),
    MR(' = '),
    MSub([MR('r')], [MR('2')]),
    MR(' · '),
    MSup([MR('e')], [
      MF([MR('π')], [MR('2')]),
      MR(' · tan φ'),
      MSub([MR('')], [MR('m')])
    ])
  ));

  children.push(P(T(`Para o φm convergido = ${n4(R.phiM)}°:`)));
  const geom = buildRupturePoints(R.phiM, b, 10);
  children.push(tbl(
    ['Parâmetro', 'Fórmula', 'Valor', 'Unidade'],
    [
      ['ϑ', '45° − φm/2', n2(45 - R.phiM/2), '°'],
      ['ν', '90°', '90', '°'],
      ['r₂', "b/(2·cos(45+φm/2))", n2(geom.r2), 'm'],
      ['r₁', 'r₂·e^(π/2·tan φm)', n2(geom.r1), 'm'],
      ['Prof. max. espiral', 'r₁·e^(ψmax·tan φm)·sin(−90°+φm)', n2(-geom.maxDepth), 'm'],
    ]
  ));
  children.push(BR());

  // ── 4. Iteração para φm ────────────────────────────────────────────────
  children.push(H1('4. Determinação do Ângulo de Atrito Equivalente φm'));
  children.push(P(T(
    'O ângulo φm é determinado iterativamente pela ponderação dos ângulos de atrito ' +
    'de cada camada pelo comprimento da linha de ruptura dentro de cada camada (DIN 4017 §7.3):'
  )));

  // tan φm = Σ(li·tan φi) / Σli
  children.push(EQ(
    MR('tan '),
    MSub([MR('φ')], [MR('m')]),
    MR(' = '),
    MF(
      [MSub([MR('Σ  l')], [MR('i')]), MR(' · tan '), MSub([MR('φ')], [MR('i')])],
      [MSub([MR('Σ  l')], [MR('i')])]
    )
  ));

  children.push(P(T('Critério de convergência: Δ = |φm,in − φm,out| / φm,in × 100% < 3%')));
  children.push(P(T('A cada iteração que não converge: φm,novo = (φm,in + φm,out) / 2')));
  children.push(BR());

  // Tabela de iterações
  children.push(P(B('Tabela de Iterações:')));
  children.push(tbl(
    ['Iteração', 'φm,in (°)', 'φm,out (°)', 'Δ (%)', 'Convergiu?'],
    R.iterations.map(it => [
      it.iter,
      n4(it.phiIn),
      n4(it.phiOut),
      n2(it.delta),
      it.delta < 3 ? 'Sim ✓' : 'Não'
    ])
  ));
  children.push(BR());
  children.push(P(B('Resultado: '), T(`φm = (${n4(R.iterations[R.iterations.length-1].phiIn)}° + ${n4(R.iterations[R.iterations.length-1].phiOut)}°) / 2 = ${n4(R.phiM)}°`)));

  // Comprimentos por camada (última iteração com φm final)
  children.push(BR());
  children.push(P(B('Comprimentos da linha de ruptura por camada (φm = ' + n2(R.phiM) + '°):')));
  children.push(tbl(
    ['Camada', 'φi (°)', 'li (m)', 'li / Σli'],
    layers.map((l, k) => [
      l.desc || `Camada ${k+1}`,
      l.phi,
      n3(R.lengths[k]),
      n4(R.lengths[k] / R.lTotal)
    ]).concat([['Total', '—', n3(R.lTotal), '1.0000']])
  ));
  children.push(BR());

  // ── 5. Coesão Equivalente cm ──────────────────────────────────────────
  children.push(H1('5. Determinação da Coesão Equivalente cm'));

  children.push(EQ(
    MSub([MR('c')], [MR('m')]),
    MR(' = '),
    MF(
      [MSub([MR('Σ  l')], [MR('i')]), MR(' · '), MSub([MR('c')], [MR('i')])],
      [MSub([MR('Σ  l')], [MR('i')])]
    )
  ));

  // Desenvolvimento numérico
  const cmNum = layers.map((l, k) => `${n3(R.lengths[k])} × ${l.c}`).join(' + ');
  const cmDen = n3(R.lTotal);
  children.push(P(T(`cm = (${cmNum}) / ${cmDen} = `), B(`${n2(R.cm)} kN/m²`)));

  // ── 6. Peso Específico Médio γm ───────────────────────────────────────
  children.push(H1('6. Determinação do Peso Específico Médio γm'));
  children.push(P(T(
    'Ponderado pelas áreas Ai de cada camada contidas dentro da superfície de ruptura:'
  )));

  children.push(EQ(
    MSub([MR('γ')], [MR('m')]),
    MR(' = '),
    MF(
      [MSub([MR('Σ  A')], [MR('i')]), MR(' · '), MSub([MR('γ')], [MR('i')])],
      [MSub([MR('Σ  A')], [MR('i')])]
    )
  ));

  const gammaEffArr = R.gammaEff || layers.map(l => l.gamma);
  children.push(tbl(
    zw != null
      ? ['Camada', 'γi (kN/m³)', "γ'i efetivo (kN/m³)", 'Ai (m²)', "Ai·γ'i (kN/m)"]
      : ['Camada', 'γi (kN/m³)', 'Ai (m²)', 'Ai·γi (kN/m)'],
    layers.map((l, k) => zw != null
      ? [l.desc || `Camada ${k+1}`, l.gamma, n2(gammaEffArr[k]), n3(R.areas[k]), n3(R.areas[k] * gammaEffArr[k])]
      : [l.desc || `Camada ${k+1}`, l.gamma, n3(R.areas[k]), n3(R.areas[k] * l.gamma)]
    ).concat([['Total', '—', ...(zw != null ? ['—'] : []), n3(R.aTotal),
      n3(layers.reduce((s, l, k) => s + R.areas[k] * gammaEffArr[k], 0))]])
  ));
  children.push(BR());
  children.push(P(B(`γm = ${n2(R.gammaM)} kN/m³`)));

  // ── 7. Peso Específico γ'm ────────────────────────────────────────────
  children.push(H1("7. Determinação do Peso Específico Acima da Fundação γ'm"));
  children.push(P(T('Média ponderada pela espessura das camadas acima da fundação:')));

  children.push(EQ(
    MSup([MR("γ'")], [MR('')]),
    MSub([MR('')], [MR('m')]),
    MR(' = '),
    MF(
      [MSub([MR('Σ  h')], [MR('i')]), MR(' · '), MSub([MR('γ')], [MR('i')])],
      [MR('t'), MSub([MR('')], [MR('f')])]
    )
  ));

  if (zw != null) {
    // Mostrar peso efetivo de cada sub-camada
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
      ['Camada', 'z topo (m)', 'z fundo (m)', 'γ (kN/m³)', "γ' efetivo (kN/m³)", 'h (m)'],
      aboveRows
    ));
    children.push(P(T(`γ'm (ponderado com NA) = `), B(`${n2(R.gammaMAbove)} kN/m³`)));
  } else {
    const gammaAboveNum = above.map(l => `${l.h} × ${l.gamma}`).join(' + ');
    children.push(P(T(`γ'm = (${gammaAboveNum}) / ${tf} = `), B(`${n2(R.gammaMAbove)} kN/m³`)));
  }

  // ── 8. Fatores de Capacidade de Carga ────────────────────────────────
  children.push(H1('8. Fatores de Capacidade de Carga'));
  children.push(P(T(`Para φm = ${n2(R.phiM)}° (DIN 4017:2006-03, Eqs. 5–7):`)));

  // Nd
  children.push(EQ(
    MSub([MR('N')], [MR('d')]),
    MR(' = '),
    MSup(
      [MR('tan'), MSup([MR('')], [MR('2')]), MR('(45° + '),
       MF([MR('φ'), MSub([MR('')], [MR('m')])], [MR('2')]), MR(')')],
      [MR('')]
    ),
    MR(' · '),
    MSup([MR('e')], [MR('π · tan φ'), MSub([MR('')], [MR('m')])])
  ));

  // Nc
  children.push(EQ(
    MSub([MR('N')], [MR('c')]),
    MR(' = '),
    MF(
      [MSub([MR('N')], [MR('d')]), MR(' − 1')],
      [MR('tan φ'), MSub([MR('')], [MR('m')])]
    )
  ));

  // Nb
  children.push(EQ(
    MSub([MR('N')], [MR('b')]),
    MR(' = '),
    MR('('), MSub([MR('N')], [MR('d')]), MR(' − 1) · tan φ'),
    MSub([MR('')], [MR('m')])
  ));

  children.push(tbl(
    ['Fator', 'Valor'],
    [
      ['Nd', n2(R.Nd)],
      ['Nc', n2(R.Nc)],
      ['Nb', n2(R.Nb)]
    ]
  ));
  children.push(BR());

  // ── 9. Fatores de Forma ───────────────────────────────────────────────
  children.push(H1('9. Fatores de Forma (Fundação Retangular)'));
  children.push(P(T(`Para fundação retangular b/a = ${b}/${a} = ${n3(b/a)} (DIN 4017 Tabelle 2):`)));

  // νb
  children.push(EQ(
    MSub([MR('ν')], [MR('b')]),
    MR(' = 1 − 0,3 · '),
    MF([MR("b'")], [MR("a'")])
  ));

  // νd
  children.push(EQ(
    MSub([MR('ν')], [MR('d')]),
    MR(' = 1 + '),
    MF([MR("b'")], [MR("a'")]),
    MR(' · sin φ'),
    MSub([MR('')], [MR('m')])
  ));

  // νc
  children.push(EQ(
    MSub([MR('ν')], [MR('c')]),
    MR(' = '),
    MF(
      [MSub([MR('ν')], [MR('d')]), MR(' · '), MSub([MR('N')], [MR('d')]), MR(' − 1')],
      [MSub([MR('N')], [MR('d')]), MR(' − 1')]
    )
  ));

  children.push(tbl(
    ['Fator', 'Valor'],
    [
      ['νd', n3(R.nud)],
      ['νc', n3(R.nuc)],
      ['νb', n3(R.nub)]
    ]
  ));
  children.push(BR());

  // ── 10. Capacidade de Carga Última qult ──────────────────────────────
  children.push(H1('10. Capacidade de Carga Última qult'));
  children.push(P(T('Fórmula principal (DIN 4017:2006-03, Eq. 1):')));

  children.push(EQ(
    MSub([MR('q')], [MR('ult')]),
    MR(' = '),
    MSub([MR('c')], [MR('m')]),
    MR(' · '),
    MSub([MR('N')], [MR('c')]),
    MR(' · '),
    MSub([MR('ν')], [MR('c')]),
    MR('  +  '),
    MSup([MR("γ'")], [MR('')]),
    MSub([MR('')], [MR('m')]),
    MR(' · '),
    MSub([MR('t')], [MR('f')]),
    MR(' · '),
    MSub([MR('N')], [MR('d')]),
    MR(' · '),
    MSub([MR('ν')], [MR('d')]),
    MR('  +  '),
    MSub([MR('γ')], [MR('m')]),
    MR(' · b · '),
    MSub([MR('N')], [MR('b')]),
    MR(' · '),
    MSub([MR('ν')], [MR('b')])
  ));

  children.push(P(T('Desenvolvimento numérico:')));
  children.push(P(
    T(`Parcela coesão:  cm·Nc·νc = ${n2(R.cm)} × ${n2(R.Nc)} × ${n3(R.nuc)} = `),
    B(`${n2(R.term_c)} kN/m²`)
  ));
  children.push(P(
    T(`Parcela sobrecarga:  γ'm·tf·Nd·νd = ${n2(R.gammaMAbove)} × ${tf} × ${n2(R.Nd)} × ${n3(R.nud)} = `),
    B(`${n2(R.term_q)} kN/m²`)
  ));
  children.push(P(
    T(`Parcela base:  γm·b·Nb·νb = ${n2(R.gammaM)} × ${b} × ${n2(R.Nb)} × ${n3(R.nub)} = `),
    B(`${n2(R.term_b)} kN/m²`)
  ));
  children.push(BR());

  children.push(new D.Paragraph({
    children: [
      T('qult = ', { size: 28, bold: true }),
      T(`${n2(R.term_c)} + ${n2(R.term_q)} + ${n2(R.term_b)} = `, { size: 28 }),
      T(`${R.qult.toFixed(0)} kN/m²`, { size: 32, bold: true, color: '1a3a5c' })
    ],
    alignment: D.AlignmentType.CENTER,
    spacing: { before: 120, after: 120 }
  }));

  // ── Rodapé ───────────────────────────────────────────────────────────
  children.push(new D.Paragraph({
    children: [T('─────────────────────────────────────────', { color: 'aaaaaa' })],
    spacing: { before: 200 }
  }));
  children.push(new D.Paragraph({
    children: [T('Calculado por din4017-bearing-capacity · pedrorussini.github.io/din4017-bearing-capacity',
      { size: 16, color: '888888', italics: true })],
    alignment: D.AlignmentType.CENTER
  }));

  // ── Gerar e baixar ────────────────────────────────────────────────────
  const doc = new D.Document({
    sections: [{ children }],
    styles: {
      default: {
        document: { run: { font: 'Calibri', size: 22 } }
      },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1',
          run: { bold: true, size: 28, color: '1a3a5c' } },
        { id: 'Heading2', name: 'Heading 2',
          run: { bold: true, size: 24, color: '2e4f7a' } }
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
