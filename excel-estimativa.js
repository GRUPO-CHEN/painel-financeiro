// Excel da calculadora aduaneira: estimativa (3 cenários) + precificação, com fórmulas vivas.
// Mesmo método da calculadora (calcularAduaneiro/calcular3 em importacoes.html).
// Uso: gerarExcelEstimativa(ExcelJS, dados) → Promise<ArrayBuffer>. Não guarda nada: os dados vêm de quem chama.
(function (global) {
'use strict';

const AZUL = 'FF1F4E8C', GRAFITE = 'FF1E293B', CINZA = 'FF64748B';
const fonte = (o = {}) => ({ name: 'Calibri', size: o.size || 10, bold: !!o.b, italic: !!o.i, color: { argb: o.c || GRAFITE } });
const F = {
  tit: fonte({ size: 15, b: true, c: AZUL }), sub: fonte({ i: true, c: CINZA }), sec: fonte({ size: 11, b: true, c: 'FFFFFFFF' }),
  subsec: fonte({ b: true, c: AZUL }), lbl: fonte(), lblb: fonte({ b: true }), inp: fonte({ c: 'FF0B4FBF' }),
  nota: fonte({ size: 9, i: true, c: CINZA }), hdr: fonte({ b: true, c: AZUL }),
};
const cor = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const P = { sec: cor(AZUL), subsec: cor('FFE8EEF7'), inp: cor('FFFFF8E1'), falta: cor('FFFDE2C4'), dest: cor('FFEEF4FB'), cen: cor('FFE9E5F5') };
const fino = { style: 'thin', color: { argb: 'FFCBD5E1' } };
const BD = { top: { top: { style: 'thin', color: { argb: GRAFITE } } }, box: { left: fino, right: fino, top: fino, bottom: fino },
  hdr: { bottom: { style: 'thin', color: { argb: AZUL } } } };
const DIR = { horizontal: 'right' }, ESQ = { horizontal: 'left' };
const N2 = '#,##0.00', N3 = '#,##0.000', N4 = '0.0000', PCT = '0.00%', PCT0 = '0.0%', PCT3 = '0.000%', DT = 'dd/mm/yyyy', INT = '#,##0';
const L = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const CI = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
const num = v => (v === '' || v == null || isNaN(Number(v))) ? null : Number(v);

function gerarExcelEstimativa(ExcelJS, d) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Painel financeiro · Importações'; wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  const folha = nome => wb.addWorksheet(nome, { views: [{ showGridLines: false }],
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true } });
  const E = folha('Estimativa'), PR = folha('Precificação'), C = folha('Cálculo'), T = folha('Tabelas');
  for (const ws of [E, PR, C, T]) ws.getColumn('A').width = 2;
  C.pageSetup.fitToPage = false;

  const put = (ws, ref, v, o = {}) => {
    const c = ws.getCell(ref);
    c.value = typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : (v === undefined ? null : v);
    c.font = o.font || F.lbl;
    if (o.fmt) c.numFmt = o.fmt;
    if (o.fill) c.fill = o.fill;
    if (o.align) c.alignment = o.align;
    if (o.border) c.border = o.border;
    return c;
  };
  const faixa = (ws, row, c1, c2, fill) => { for (let c = c1; c <= c2; c++) ws.getCell(row, c).fill = fill; };
  const secao = (ws, row, txt, c1, c2) => { faixa(ws, row, c1, c2, P.sec); put(ws, L(c1) + row, txt, { font: F.sec, fill: P.sec }); };
  const subsecao = (ws, row, txt, c1, c2) => { faixa(ws, row, c1, c2, P.subsec); put(ws, L(c1) + row, txt, { font: F.subsec, fill: P.subsec }); };
  const entrada = (ws, ref, v, fmt) => put(ws, ref, v, { font: F.inp, fmt, fill: P.inp, border: BD.box });
  const nota = (ws, ref, txt) => { ws.getCell(ref).note = txt; };
  const lista = (ws, ref, formula) => { ws.getCell(ref).dataValidation = { type: 'list', allowBlank: true, formulae: [formula] }; };

  const pr = d.prem || {}, a = pr.a || {}, m = pr.m || {}, dh = pr.d || {};
  const itens = (d.itens || []).filter(x => num(x.qtd) && num(x.fob));
  const NIT = Math.max(15, itens.length);
  const margens = (d.margens && d.margens.length ? d.margens : [0.5, 0.45, 0.4, 0.3, 0.2, 0.1, 0]).slice(0, 7);
  while (margens.length < 7) margens.push(0);
  const emps = (d.empresas || []).slice(0, 2);
  while (emps.length < 2) emps.push({ nome: emps.length ? 'Empresa 2' : 'Empresa 1', regime: 'tabela', credito: true });

  // ======================= Tabelas =======================
  const tipos = d.tipos || [];
  const TL0 = 6, TL1 = Math.max(TL0 + tipos.length + 4, 20);
  const TAB = `Tabelas!$B$${TL0}:$S$${TL1}`;
  const K = { ii: 2, ipi: 3, icms: 4, pis: 5, cof: 6, iid: 7, icmsd: 8, ipis: 9, icmss: 10, piss: 11, cofs: 12, irpj: 13, tots: 14, d1: 15, d2: 16, d3: 17, dt: 18 };
  put(T, 'B1', 'Alíquotas e descontos por tipo de produto', { font: F.tit });
  put(T, 'B2', 'Importação (inclui DHL): as do painel. Saída e descontos: a precificação do painel. Desconto em cascata: 1 − (1 − d1)(1 − d2)(1 − d3).', { font: F.sub });
  put(T, 'B3', `Em laranja: falta preencher ou confirmar. Para um tipo novo, use as linhas vazias até a ${TL1}.`, { font: F.nota });
  secao(T, 4, 'Importação', 3, 9); secao(T, 4, 'Saída (venda)', 10, 15); secao(T, 4, 'Desconto (em cascata)', 16, 19);
  ['Tipo de produto', 'II', 'IPI', 'ICMS', 'PIS', 'COFINS', 'II DHL', 'ICMS DHL', 'IPI', 'ICMS', 'PIS', 'COFINS', 'IRPJ/CSLL', 'Total saída',
    'Desconto 1', 'Desconto 2', 'Desconto 3', 'Desconto total'].forEach((h, i) =>
    put(T, L(2 + i) + 5, h, { font: F.hdr, align: { horizontal: i ? 'center' : 'left', wrapText: true }, border: BD.hdr }));
  T.getRow(5).height = 28;
  tipos.forEach((t, n) => {
    const r = TL0 + n;
    put(T, 'B' + r, t.familia, { font: F.lblb });
    [['C', t.ii], ['D', t.ipi], ['E', t.icms], ['F', t.pis], ['G', t.cofins], ['H', t.ii_dhl], ['I', t.icms_dhl]].forEach(([c, v]) => entrada(T, c + r, num(v) ?? 0, PCT));
    if (num(t.ipi_saida) == null) put(T, 'J' + r, `=D${r}`, { fmt: PCT }); else entrada(T, 'J' + r, num(t.ipi_saida), PCT);
    [['K', t.icms_saida], ['L', t.pis_saida], ['M', t.cofins_saida], ['N', t.irpj_csll]].forEach(([c, v]) => {
      const x = entrada(T, c + r, num(v) ?? 0, PCT); if (!t.cadastrado) x.fill = P.falta; });
    if (!t.cadastrado) nota(T, 'K' + r, `${t.familia} ainda não tem impostos de saída cadastrados na precificação do painel: ICMS de saída = ICMS da importação e PIS, COFINS e IRPJ/CSLL no padrão. Confirmar.`);
    put(T, 'O' + r, `=SUM(J${r}:N${r})`, { font: F.lblb, fmt: PCT });
    const ds = [t.d1, t.d2, t.d3].map(num), semDesc = ds.every(x => !x);
    ['P', 'Q', 'R'].forEach((c, i) => { const x = entrada(T, c + r, ds[i], PCT); if (semDesc) x.fill = P.falta; });
    put(T, 'S' + r, `=1-(1-N(P${r}))*(1-N(Q${r}))*(1-N(R${r}))`, { font: F.lblb, fmt: PCT3 });
  });
  const ER = [TL1 + 4, TL1 + 5];                                    // linhas das duas empresas
  secao(T, TL1 + 2, 'Empresas que vendem (impostos de saída)', 2, 5);
  ['Empresa', 'Regime', 'Simples', 'Crédito IPI+ICMS'].forEach((h, i) =>
    put(T, L(2 + i) + (TL1 + 3), h, { font: F.hdr, align: { horizontal: i ? 'center' : 'left', wrapText: true }, border: BD.hdr }));
  emps.forEach((e, i) => {
    const r = ER[i];
    entrada(T, 'B' + r, e.nome || `Empresa ${i + 1}`);
    entrada(T, 'C' + r, e.regime === 'simples' ? 'Simples' : 'Por tipo'); lista(T, 'C' + r, '"Por tipo,Simples"');
    entrada(T, 'D' + r, e.regime === 'simples' ? (num(e.aliquota) ?? 0) : null, PCT);
    entrada(T, 'E' + r, e.credito ? 'Sim' : 'Não'); lista(T, 'E' + r, '"Sim,Não"');
  });
  nota(T, 'C' + ER[0], '"Por tipo": impostos de saída da tabela acima (IPI, ICMS, PIS, COFINS, IRPJ/CSLL). "Simples": uma alíquota única sobre a venda (coluna D).');
  nota(T, 'E' + ER[1], 'No Simples Nacional a empresa não aproveita crédito de ICMS e IPI da compra; por isso o custo é o cheio. Se for diferente, troque para Sim.');
  const EMP = ER.map(r => ({ nome: `Tabelas!$B$${r}`, reg: `Tabelas!$C$${r}`, sim: `Tabelas!$D$${r}`, cred: `Tabelas!$E$${r}` }));
  T.getColumn('B').width = 24;
  for (let c = 3; c <= 19; c++) T.getColumn(c).width = 10.5;
  T.views = [{ state: 'frozen', xSplit: 2, ySplit: 5, showGridLines: false }];

  // ======================= Estimativa =======================
  E.getColumn('B').width = 40;
  for (const c of 'CDE') E.getColumn(c).width = 14;
  E.getColumn('F').width = 24;
  for (const c of 'GHIJKLM') E.getColumn(c).width = 13;
  put(E, 'B1', 'Estimativa de importação', { font: F.tit });
  put(E, 'B2', '=C7&IF(C5<>""," · "&C5,"")&IF(C6<>""," · cliente "&C6,"")&" · "&TEXT(DAY(C9),"00")&"/"&TEXT(MONTH(C9),"00")&"/"&YEAR(C9)', { font: F.sub });
  put(E, 'B3', 'Gerado pela calculadora aduaneira do painel (Importações). Mesmo método de cálculo; campos amarelos são para digitar.', { font: F.nota });
  secao(E, 4, 'Dados', 2, 3);
  const dataEst = d.data ? new Date(String(d.data).slice(0, 10) + 'T12:00:00') : new Date();
  [[5, 'Fornecedor', d.fornecedor || ''], [6, 'Cliente', d.cliente || ''], [7, 'Mercadoria', d.mercadoria || ''], [8, 'Empresa', null],
    [9, 'Data da estimativa', dataEst], [10, 'Cenário para precificar', d.cenario || 'Aéreo']].forEach(([r, l, v]) => {
    put(E, 'B' + r, l, { font: r === 10 ? F.lblb : F.lbl });
    if (r !== 8) entrada(E, 'C' + r, v, r === 9 ? DT : undefined);
  });
  put(E, 'C8', `=${EMP[0].nome}&" × "&${EMP[1].nome}&" (ver Precificação)"`, { font: F.nota });
  const CEN = '$C$10';
  lista(E, 'C10', '"Aéreo,Marítimo,DHL"');
  nota(E, 'C10', 'Cenário que vai para a precificação. Troque para ver pelo marítimo ou pelo DHL.');

  const hist = (d.historico || []).slice(0, 6);
  secao(E, 4, `Últimas compras${d.fornecedor ? ' da ' + d.fornecedor : ' do fornecedor'}`, 5, 10);
  ['Pedido', 'Empresa', 'Frete', 'Fator', 'Dólar DI', 'Data DI'].forEach((h, i) =>
    put(E, L(5 + i) + 5, h, { font: F.hdr, border: BD.hdr, align: i < 3 ? ESQ : DIR }));
  if (!hist.length) put(E, 'E6', d.fornecedor ? 'Sem processos fechados deste fornecedor no painel.' : 'Simulação avulsa: sem fornecedor escolhido.', { font: F.nota });
  hist.forEach((h, i) => {
    const r = 6 + i;
    put(E, 'E' + r, h.pedido ?? '', { align: ESQ }); put(E, 'F' + r, h.empresa || ''); put(E, 'G' + r, h.modal || '');
    put(E, 'H' + r, num(h.fator), { fmt: N4 }); put(E, 'I' + r, num(h.dolar_di) ?? '—', { fmt: N4, align: DIR });
    put(E, 'J' + r, h.data_di ? new Date(String(h.data_di).slice(0, 10) + 'T12:00:00') : '—', { fmt: DT, align: DIR });
  });
  put(E, 'E12', 'Fonte: processos fechados no painel (fator final).', { font: F.nota });

  secao(E, 14, 'Premissas gerais', 2, 3);
  const DI = '$C$15', DC = '$C$16', ACR = '$C$17', SISC = '$C$20', RAT = '$C$21';
  [[15, 'Dólar DUIMP estimado', num(pr.dolar_di), N4], [16, 'Dólar do câmbio estimado', num(pr.dolar_cambio) ?? num(pr.dolar_di), N4],
    [17, 'Outros acréscimos, US$', num(pr.outros_acresc) ?? 0, N2], [18, 'Nº de adições (vazio = automático)', num(pr.adicoes) || null, '0']]
    .forEach(([r, l, v, f]) => { put(E, 'B' + r, l); entrada(E, 'C' + r, v, f); });

  // ---- itens (Estimativa) ↔ Cálculo
  const I0 = 36, I1 = I0 + NIT - 1, R0 = 6, R1 = R0 + NIT - 1, RT = R1 + 1;
  put(E, 'B19', 'Nº de adições usado');
  put(E, 'C19', `=IF(C18>0,C18,SUMPRODUCT((Cálculo!G${R0}:G${R1}>0)/(COUNTIFS(Cálculo!C${R0}:C${R1},Cálculo!C${R0}:C${R1},Cálculo!G${R0}:G${R1},">0")+(Cálculo!G${R0}:G${R1}<=0))))`, { fmt: '0' });
  nota(E, 'C19', 'Automático como no painel: uma adição por fornecedor + tipo de produto (aqui, um fornecedor só).');
  put(E, 'B20', 'Taxa Siscomex, R$ (pelas adições)');
  const n = 'ROUND(C19,0)';
  put(E, 'C20', `=IF(${n}<=0,0,115.67+MIN(${n},2)*38.56+MAX(MIN(${n}-2,3),0)*30.85+MAX(MIN(${n}-5,5),0)*23.14+MAX(MIN(${n}-10,10),0)*15.42+MAX(MIN(${n}-20,30),0)*7.71+MAX(${n}-50,0)*3.86)`, { fmt: N2 });
  put(E, 'B21', 'Rateio do frete');
  put(E, 'C21', `=IF(AND(Cálculo!H${RT}>0,COUNTIFS(Cálculo!G${R0}:G${R1},">0",Cálculo!H${R0}:H${R1},">0")=COUNTIF(Cálculo!G${R0}:G${R1},">0")),"por peso","por FOB")`);
  nota(E, 'C21', 'Igual ao painel: com o peso de todos os itens, o frete é rateado por peso líquido (como na DUIMP); senão, por FOB.');

  secao(E, 23, 'Premissas por cenário', 2, 5);
  ['Aéreo', 'Marítimo', 'DHL'].forEach((h, i) => put(E, L(3 + i) + 23, h, { font: F.sec, fill: P.sec, align: DIR }));
  const v0 = x => num(x) ?? 0;
  [[24, 'Frete internacional, US$', [v0(a.frete_usd), v0(m.frete_usd), v0(dh.frete_usd)]],
    [25, 'Despesas do frete internacional, R$', [v0(a.frete_desp), v0(m.frete_desp), null]],
    [26, 'AFRMM, R$ (aprox. 8% do frete)', [null, v0(m.afrmm), null]],
    [27, 'SDAS, R$', [v0(a.sdas), v0(m.sdas), null]], [28, 'Armazenagem, R$', [v0(a.armazenagem), v0(m.armazenagem), null]],
    [29, 'Despachante, R$', [v0(a.despachante), v0(m.despachante), null]], [30, 'Frete nacional, R$', [v0(a.frete_nacional), v0(m.frete_nacional), null]],
    [31, 'Outras despesas, R$', [v0(a.outras), v0(m.outras), null]], [32, 'Despesas DHL, R$', [null, null, v0(dh.despesas)]],
  ].forEach(([r, l, vs]) => {
    put(E, 'B' + r, l);
    vs.forEach((v, i) => { const ref = L(3 + i) + r; if (v == null) put(E, ref, '—', { font: F.nota, align: DIR }); else entrada(E, ref, v, N2); });
  });
  const PREM = {
    a: { fu: 'Estimativa!$C$24', fd: 'Estimativa!$C$25', af: '0', sd: 'Estimativa!$C$27', ar: 'Estimativa!$C$28', de: 'Estimativa!$C$29', fn: 'Estimativa!$C$30', ou: 'Estimativa!$C$31' },
    m: { fu: 'Estimativa!$D$24', fd: 'Estimativa!$D$25', af: 'Estimativa!$D$26', sd: 'Estimativa!$D$27', ar: 'Estimativa!$D$28', de: 'Estimativa!$D$29', fn: 'Estimativa!$D$30', ou: 'Estimativa!$D$31' },
  };
  const FU_D = 'Estimativa!$E$24', DESP_D = 'Estimativa!$E$32';

  secao(E, 34, 'Itens da invoice', 2, 13);
  ['Item / código', 'Qtd', 'FOB unit., US$', 'Peso líq. unit., kg', 'Tipo de produto (alíquota)', 'FOB total, US$', 'Peso total, kg',
    'Fator aéreo', 'Fator marítimo', 'Fator DHL', 'Custo unit., R$ (cenário)', 'Desconto negociado'].forEach((h, i) =>
    put(E, L(2 + i) + 35, h, { font: F.hdr, align: { horizontal: (i === 0 || i === 4) ? 'left' : 'right', wrapText: true }, border: BD.hdr }));
  E.getRow(35).height = 28;
  nota(E, 'M35', 'Vazio = desconto do tipo de produto (aba Tabelas). Preencha quando negociar outro desconto com o cliente.');
  for (let k = 0; k < NIT; k++) {
    const r = I0 + k, cr = R0 + k, x = itens[k] || {};
    entrada(E, 'B' + r, x.codigo || (itens[k] ? `Item ${k + 1}` : null));
    entrada(E, 'C' + r, num(x.qtd), INT); entrada(E, 'D' + r, num(x.fob), N2); entrada(E, 'E' + r, num(x.peso) || null, N3);
    entrada(E, 'F' + r, x.familia || null); lista(E, 'F' + r, `Tabelas!$B$${TL0}:$B$${TL1}`);
    entrada(E, 'M' + r, num(x.desc_neg), PCT);
    put(E, 'G' + r, `=IF(Cálculo!G${cr}>0,Cálculo!G${cr},"")`, { fmt: N2 });
    put(E, 'H' + r, `=IF(Cálculo!H${cr}>0,Cálculo!H${cr},"")`, { fmt: N3 });
  }
  put(E, 'B' + (I1 + 1), 'Total', { font: F.lblb, border: BD.top });
  put(E, 'C' + (I1 + 1), `=SUM(C${I0}:C${I1})`, { font: F.lblb, fmt: INT, border: BD.top });
  put(E, 'G' + (I1 + 1), `=SUM(G${I0}:G${I1})`, { font: F.lblb, fmt: N2, border: BD.top });
  put(E, 'H' + (I1 + 1), `=SUM(H${I0}:H${I1})`, { font: F.lblb, fmt: N3, border: BD.top });
  for (const c of 'DEFIJKLM') E.getCell(c + (I1 + 1)).border = BD.top;

  // ======================= Cálculo (apoio) =======================
  put(C, 'B1', 'Cálculo por item (apoio)', { font: F.tit });
  put(C, 'B2', 'Mesmo método da calculadora do painel: impostos item a item; Siscomex e AFRMM pelo valor aduaneiro; SDAS, armazenagem, despachante e outras pela base do ICMS; fretes por peso ou FOB. Não precisa mexer aqui.', { font: F.sub });
  secao(C, 4, 'Item e alíquotas', 2, 17);
  [['B', 'Item'], ['C', 'Tipo de produto'], ['D', 'Qtd'], ['E', 'FOB unit., US$'], ['F', 'Peso unit., kg'], ['G', 'FOB total, US$'], ['H', 'Peso total, kg'],
    ['I', 'II'], ['J', 'IPI'], ['K', 'ICMS'], ['L', 'PIS'], ['M', 'COFINS'], ['N', 'II DHL'], ['O', 'ICMS DHL'], ['P', 'Acréscimo rateado, US$'], ['Q', 'Parte do frete']]
    .forEach(([c, h]) => put(C, c + 5, h, { font: F.hdr, align: { horizontal: 'BC'.includes(c) ? 'left' : 'right', wrapText: true }, border: BD.hdr }));
  const lk = (r, k) => `IFERROR(VLOOKUP($C${r},${TAB},${k},0),0)`;
  for (let k = 0; k < NIT; k++) {
    const r = R0 + k, er = I0 + k;
    put(C, 'B' + r, `=IF(Estimativa!B${er}="","",Estimativa!B${er})`);
    put(C, 'C' + r, `=IF(Estimativa!F${er}="","",Estimativa!F${er})`);
    put(C, 'D' + r, `=N(Estimativa!C${er})`, { fmt: INT });
    put(C, 'E' + r, `=N(Estimativa!D${er})`, { fmt: N2 });
    put(C, 'F' + r, `=N(Estimativa!E${er})`, { fmt: N3 });
    put(C, 'G' + r, `=IF(AND(D${r}>0,E${r}>0),D${r}*E${r},0)`, { fmt: N2 });
    put(C, 'H' + r, `=IF(G${r}>0,D${r}*F${r},0)`, { fmt: N3 });
    [['I', K.ii], ['J', K.ipi], ['K', K.icms], ['L', K.pis], ['M', K.cof], ['N', K.iid], ['O', K.icmsd]].forEach(([c, kk]) => put(C, c + r, '=' + lk(r, kk), { fmt: PCT }));
    put(C, 'P' + r, `=IF($G$${RT}>0,Estimativa!${ACR}*G${r}/$G$${RT},0)`, { fmt: N2 });
    put(C, 'Q' + r, `=IF(Estimativa!${RAT}="por peso",IF($H$${RT}>0,H${r}/$H$${RT},0),IF($G$${RT}>0,G${r}/$G$${RT},0))`, { fmt: PCT });
  }
  put(C, 'B' + RT, 'Total', { font: F.lblb, border: BD.top });
  for (const c of 'GHPQ') put(C, c + RT, `=SUM(${c}${R0}:${c}${R1})`, { font: F.lblb, fmt: c === 'H' ? N3 : c === 'Q' ? PCT : N2, border: BD.top });

  const BLOCO = ['Frete rateado, US$', 'Valor aduaneiro, US$', 'Valor aduaneiro, R$', 'II', 'IPI', 'PIS', 'COFINS', 'Taxa Siscomex', 'AFRMM', 'Base do ICMS', 'ICMS',
    'Parte das despesas', 'SDAS', 'Armazenagem', 'Despachante', 'Outras despesas', 'Frete internacional, R$', 'Frete nacional', 'Câmbio, R$', 'Custo total, R$', 'Fator'];
  const CH = ['fr', 'vu', 'vb', 'ii', 'ipi', 'pis', 'cof', 'sc', 'af', 'bs', 'icm', 'sh', 'sd', 'ar', 'de', 'ou', 'fb', 'fn', 'cb', 'cu', 'ft'];
  const di = `Estimativa!${DI}`, dc = `Estimativa!${DC}`;
  function bloco(c0, nome, p) {
    const X = Object.fromEntries(CH.map((k, i) => [k, L(c0 + i)]));
    secao(C, 4, nome, c0, c0 + CH.length - 1);
    BLOCO.forEach((h, i) => put(C, L(c0 + i) + 5, h, { font: F.hdr, align: { horizontal: 'right', wrapText: true }, border: BD.hdr }));
    const vm = `($G$${RT}+${p.fu})`;
    for (let k = 0; k < NIT; k++) {
      const r = R0 + k;
      const f = {
        fr: `=${p.fu}*Q${r}`, vu: `=G${r}+${X.fr}${r}+P${r}`, vb: `=${X.vu}${r}*${di}`,
        ii: `=I${r}*${X.vb}${r}`, ipi: `=J${r}*(${X.vb}${r}+${X.ii}${r})`, pis: `=L${r}*${X.vb}${r}`, cof: `=M${r}*${X.vb}${r}`,
        sc: `=IF(${vm}>0,Estimativa!${SISC}*${X.vu}${r}/${vm},0)`, af: `=IF(${vm}>0,${p.af}*${X.vu}${r}/${vm},0)`,
        bs: `=IF(AND(G${r}>0,K${r}<1),(${X.vb}${r}+${X.ii}${r}+${X.ipi}${r}+${X.pis}${r}+${X.cof}${r}+${X.sc}${r}+${X.af}${r})/(1-K${r}),0)`,
        icm: `=${X.bs}${r}*K${r}`,
        sh: `=IF($${X.bs}$${RT}>0,${X.bs}${r}/$${X.bs}$${RT},IF($G$${RT}>0,G${r}/$G$${RT},0))`,
        sd: `=${p.sd}*${X.sh}${r}`, ar: `=${p.ar}*${X.sh}${r}`, de: `=${p.de}*${X.sh}${r}`, ou: `=${p.ou}*${X.sh}${r}`,
        fb: `=(${p.fu}*${di}+${p.fd})*Q${r}`, fn: `=${p.fn}*Q${r}`, cb: `=G${r}*${dc}`,
        cu: `=${X.cb}${r}+${X.ii}${r}+${X.ipi}${r}+${X.pis}${r}+${X.cof}${r}+${X.icm}${r}+${X.sc}${r}+${X.af}${r}+${X.sd}${r}+${X.ar}${r}+${X.de}${r}+${X.ou}${r}+${X.fb}${r}+${X.fn}${r}`,
        ft: `=IF(G${r}>0,${X.cu}${r}/G${r},"")`,
      };
      for (const key of CH) put(C, X[key] + r, f[key], { fmt: key === 'ft' ? N4 : key === 'sh' ? PCT : N2 });
    }
    for (const key of CH) {
      if (key === 'ft') put(C, X.ft + RT, `=IF(G${RT}>0,${X.cu}${RT}/G${RT},"")`, { font: F.lblb, fmt: N4, border: BD.top });
      else put(C, X[key] + RT, `=SUM(${X[key]}${R0}:${X[key]}${R1})`, { font: F.lblb, fmt: key === 'sh' ? PCT : N2, border: BD.top });
    }
    return X;
  }
  const A = bloco(CI('R'), 'Aéreo', PREM.a);
  const M = bloco(CI('R') + CH.length, 'Marítimo', PREM.m);
  const d0 = CI('R') + 2 * CH.length;
  const DK = ['bs', 'ii', 'icm', 'imp', 'de', 'fb', 'cu', 'ft'];
  const D = Object.fromEntries(DK.map((k, i) => [k, L(d0 + i)]));
  secao(C, 4, 'DHL (courier)', d0, d0 + 7);
  ['Base (FOB + frete) × dólar', 'II DHL', 'ICMS DHL', 'Impostos DHL', 'Despesas DHL', 'Frete, R$', 'Custo total, R$', 'Fator'].forEach((h, i) =>
    put(C, L(d0 + i) + 5, h, { font: F.hdr, align: { horizontal: 'right', wrapText: true }, border: BD.hdr }));
  const tr2 = x => `INT((${x})*100+0.0000001)/100`;               // a DHL trunca os centavos (igual ao painel)
  for (let k = 0; k < NIT; k++) {
    const r = R0 + k;
    const f = {
      bs: `=IF(G${r}>0,${tr2(`G${r}*${di}`)}+${tr2(`${FU_D}*Q${r}*${di}`)},0)`,
      ii: `=${tr2(`N${r}*${D.bs}${r}`)}`,
      icm: `=IF(O${r}<1,${tr2(`(${D.bs}${r}+${D.ii}${r})/(1-O${r})*O${r}`)},0)`,
      imp: `=${D.ii}${r}+${D.icm}${r}`,
      de: `=IF($${D.imp}$${RT}>0,${D.imp}${r}*${DESP_D}/$${D.imp}$${RT},IF($G$${RT}>0,${DESP_D}*G${r}/$G$${RT},0))`,
      fb: `=IF($G$${RT}>0,G${r}*${FU_D}*${di}/$G$${RT},0)`,
      cu: `=${D.imp}${r}+G${r}*${dc}+${D.de}${r}+${D.fb}${r}`,
      ft: `=IF(G${r}>0,${D.cu}${r}/G${r},"")`,
    };
    for (const key of DK) put(C, D[key] + r, f[key], { fmt: key === 'ft' ? N4 : N2 });
  }
  for (const key of DK) {
    if (key === 'ft') put(C, D.ft + RT, `=IF(G${RT}>0,${D.cu}${RT}/G${RT},"")`, { font: F.lblb, fmt: N4, border: BD.top });
    else put(C, D[key] + RT, `=SUM(${D[key]}${R0}:${D[key]}${R1})`, { font: F.lblb, fmt: N2, border: BD.top });
  }
  const e0 = d0 + 8, EK = ['cu', 'cr', 'cuu', 'cru', 'liq', 'ft'];
  const ES = Object.fromEntries(EK.map((k, i) => [k, L(e0 + i)]));
  secao(C, 4, 'Cenário escolhido (vai para a precificação)', e0, e0 + 5);
  ['Custo total, R$', 'Créditos (IPI + ICMS), R$', 'Custo unitário, R$', 'Créditos unitários, R$', 'Custo unit. sem créditos, R$', 'Fator'].forEach((h, i) =>
    put(C, L(e0 + i) + 5, h, { font: F.hdr, align: { horizontal: 'right', wrapText: true }, border: BD.hdr }));
  const cen = `Estimativa!${CEN}`;
  for (let k = 0; k < NIT; k++) {
    const r = R0 + k;
    const f = {
      cu: `=IF(${cen}="Aéreo",${A.cu}${r},IF(${cen}="Marítimo",${M.cu}${r},${D.cu}${r}))`,
      cr: `=IF(${cen}="Aéreo",${A.ipi}${r}+${A.icm}${r},IF(${cen}="Marítimo",${M.ipi}${r}+${M.icm}${r},${D.icm}${r}))`,
      cuu: `=IF(G${r}>0,${ES.cu}${r}/D${r},"")`, cru: `=IF(G${r}>0,${ES.cr}${r}/D${r},"")`,
      liq: `=IF(G${r}>0,(${ES.cu}${r}-${ES.cr}${r})/D${r},"")`, ft: `=IF(G${r}>0,${ES.cu}${r}/G${r},"")`,
    };
    for (const key of EK) put(C, ES[key] + r, f[key], { font: key === 'cuu' || key === 'liq' ? F.lblb : F.lbl, fmt: key === 'ft' ? N4 : N2 });
  }
  C.getColumn('B').width = 18; C.getColumn('C').width = 22;
  for (let c = 4; c <= e0 + 5; c++) C.getColumn(c).width = 12;
  C.getRow(5).height = 40;
  C.views = [{ state: 'frozen', xSplit: 3, ySplit: 5, showGridLines: false }];

  for (let k = 0; k < NIT; k++) {
    const r = I0 + k, cr = R0 + k;
    put(E, 'I' + r, `=IF(Cálculo!G${cr}>0,Cálculo!${A.ft}${cr},"")`, { fmt: N4 });
    put(E, 'J' + r, `=IF(Cálculo!G${cr}>0,Cálculo!${M.ft}${cr},"")`, { fmt: N4 });
    put(E, 'K' + r, `=IF(Cálculo!G${cr}>0,Cálculo!${D.ft}${cr},"")`, { fmt: N4 });
    put(E, 'L' + r, `=Cálculo!${ES.cuu}${cr}`, { font: F.lblb, fmt: N2 });
  }

  // ---- resultado (igual à tabela da calculadora)
  const RR = I1 + 3;
  secao(E, RR, 'Resultado', 2, 5);
  ['Aéreo', 'Marítimo', 'DHL'].forEach((h, i) => put(E, L(3 + i) + RR, h, { font: F.sec, fill: P.sec, align: DIR }));
  const tot = c => `Cálculo!${c}${RT}`;
  const RN = ['fob', 'camb', 'dif', 'ii', 'pis', 'cof', 'ipi', 'icms', 'imp', 'sisc', 'afrmm', 'sdas', 'arm', 'desp', 'outras', 'frete', 'fnac', 'num', 'total', 'fator', 'cred', 'real', 'fator2'];
  const rr = Object.fromEntries(RN.map((k, i) => [k, RR + 1 + i]));
  const tres = f => [f, f, f];
  const res = [
    ['fob', 'Valor total da invoice (FOB), US$', [tot('G'), tot('G'), tot('G')], N2, false],
    ['camb', 'Câmbio (FOB × dólar câmbio)', [tot(A.cb), tot(M.cb), `${tot('G')}*${DC}`], N2, false],
    ['dif', 'Diferença cambial (câmbio − DUIMP)', tres(`{c}${rr.fob}*(${DC}-${DI})`), N2, false],
    ['ii', 'II', [tot(A.ii), tot(M.ii), tot(D.ii)], N2, false],
    ['pis', 'PIS', [tot(A.pis), tot(M.pis), '0'], N2, false],
    ['cof', 'COFINS', [tot(A.cof), tot(M.cof), '0'], N2, false],
    ['ipi', 'IPI', [tot(A.ipi), tot(M.ipi), '0'], N2, false],
    ['icms', 'ICMS', [tot(A.icm), tot(M.icm), tot(D.icm)], N2, false],
    ['imp', 'Impostos total', tres(`SUM({c}${rr.ii}:{c}${rr.icms})`), N2, true],
    ['sisc', '="Taxa Siscomex ("&C19&" adição(ões))"', [tot(A.sc), tot(M.sc), '0'], N2, false],
    ['afrmm', 'AFRMM', [null, tot(M.af), null], N2, false],
    ['sdas', 'SDAS', [tot(A.sd), tot(M.sd), '0'], N2, false],
    ['arm', 'Armazenagem', [tot(A.ar), tot(M.ar), '0'], N2, false],
    ['desp', 'Despachante', [tot(A.de), tot(M.de), '0'], N2, false],
    ['outras', 'Outras despesas / despesas DHL', [tot(A.ou), tot(M.ou), tot(D.de)], N2, false],
    ['frete', 'Frete internacional (R$)', [tot(A.fb), tot(M.fb), tot(D.fb)], N2, false],
    ['fnac', 'Frete nacional', [tot(A.fn), tot(M.fn), '0'], N2, false],
    ['num', 'Numerários total', tres(`{c}${rr.imp}+SUM({c}${rr.sisc}:{c}${rr.fnac})`), N2, true],
    ['total', 'Total importação', tres(`{c}${rr.camb}+{c}${rr.num}`), N2, true],
    ['fator', 'Fator de importação', tres(`IF({c}${rr.fob}>0,{c}${rr.total}/{c}${rr.fob},"")`), N4, true],
    ['cred', 'Crédito ICMS + IPI', tres(`{c}${rr.icms}+{c}${rr.ipi}`), N2, false],
    ['real', 'Custo real (sem créditos)', tres(`{c}${rr.total}-{c}${rr.cred}`), N2, false],
    ['fator2', 'Fator 2 (sem créditos)', tres(`IF({c}${rr.fob}>0,{c}${rr.real}/{c}${rr.fob},"")`), N4, false],
  ];
  for (const [key, lab, fs, fmt, forte] of res) {
    const r = rr[key], font = forte ? F.lblb : F.lbl;
    put(E, 'B' + r, lab, { font });
    fs.forEach((fx, i) => { const c = L(3 + i); if (fx == null) put(E, c + r, '—', { font: F.nota, align: DIR }); else put(E, c + r, '=' + fx.replace(/\{c\}/g, c), { font, fmt }); });
    if (['imp', 'num', 'total'].includes(key)) for (const c of 'BCDE') E.getCell(c + r).border = BD.top;
  }
  const RF = rr.fator2;
  for (const [ref, hdr] of [['C24:E32', 23], [`C${RR + 1}:E${RF}`, RR]])
    E.addConditionalFormatting({ ref, rules: [{ type: 'expression', priority: 1, formulae: [`C$${hdr}=${CEN}`], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFE9E5F5' } } } }] });
  put(E, 'B' + (RF + 2), 'Fator e custo por item: na tabela de itens. Detalhe item a item: aba Cálculo.', { font: F.nota });
  E.views = [{ state: 'frozen', ySplit: 3, showGridLines: false }];

  // ======================= Precificação (empresa 1 × empresa 2) =======================
  PR.getColumn('B').width = 46;
  for (const c of 'CDEFGHIJ') PR.getColumn(c).width = 15;
  for (const c of 'KLMN') PR.getColumn(c).width = 12.5;
  put(PR, 'B1', `="Precificação · "&${EMP[0].nome}&" × "&${EMP[1].nome}`, { font: F.tit });
  put(PR, 'B2', '=Estimativa!B2&" · cenário "&Estimativa!C10', { font: F.sub });
  put(PR, 'B3', 'Preço de tabela = (custo − créditos) ÷ (1 − margem − impostos de saída − comissão) ÷ (1 − desconto). Margens, comissão e preço informado são para digitar; desconto negociado vai na coluna M dos itens.', { font: F.nota });
  put(PR, 'B4', 'Item a precificar', { font: F.lblb }); entrada(PR, 'C4', itens[0] ? (itens[0].codigo || 'Item 1') : null); PR.mergeCells('C4:D4');
  lista(PR, 'C4', `Estimativa!$B$${I0}:$B$${I1}`);
  put(PR, 'B5', '(linha do item na estimativa)', { font: F.nota }); put(PR, 'C5', `=MATCH(C4,Estimativa!$B$${I0}:$B$${I1},0)`, { font: F.nota, fmt: '0' });
  secao(PR, 6, 'Dados do item', 2, 3);
  const ix = col => `=INDEX(Cálculo!$${col}$${R0}:$${col}$${R1},$C$5)`;
  [[7, 'Tipo de produto', ix('C'), null, false], [8, 'Quantidade', ix('D'), INT, false], [9, 'FOB unitário, US$', ix('E'), N2, false],
    [10, 'Cenário', '=Estimativa!C10', null, false], [11, 'Fator de importação', ix(ES.ft), N4, false], [12, 'Custo unitário, R$', ix(ES.cuu), N2, true],
    [13, 'Créditos unitários (IPI + ICMS), R$', ix(ES.cru), N2, false], [14, 'Custo unitário sem créditos, R$', ix(ES.liq), N2, false]]
    .forEach(([r, l, f, fmt, b]) => { put(PR, 'B' + r, l, { font: b ? F.lblb : F.lbl }); put(PR, 'C' + r, f, { font: b ? F.lblb : F.lbl, fmt, align: r === 7 || r === 10 ? ESQ : DIR }); });
  const FOBU = '$C$9', CUSTO = '$C$12', CRED = '$C$13';

  secao(PR, 6, 'Impostos de saída', 5, 7);
  put(PR, 'E7', '', { border: BD.hdr });
  put(PR, 'F7', `=${EMP[0].nome}`, { font: F.hdr, align: DIR, border: BD.hdr }); put(PR, 'G7', `=${EMP[1].nome}`, { font: F.hdr, align: DIR, border: BD.hdr });
  const lt = k => `IFERROR(VLOOKUP($C$7,${TAB},${k},0),0)`;
  [[8, 'IPI', K.ipis], [9, 'ICMS', K.icmss], [10, 'PIS', K.piss], [11, 'COFINS', K.cofs], [12, 'IRPJ/CSLL', K.irpj]].forEach(([r, l, k]) => {
    put(PR, 'E' + r, l);
    ['F', 'G'].forEach((c, i) => put(PR, c + r, `=IF(${EMP[i].reg}="Simples","—",${lt(k)})`, { fmt: PCT, align: DIR }));
  });
  put(PR, 'E13', 'Simples Nacional');
  ['F', 'G'].forEach((c, i) => put(PR, c + '13', `=IF(${EMP[i].reg}="Simples",N(${EMP[i].sim}),"—")`, { fmt: PCT, align: DIR }));
  put(PR, 'E14', 'Total', { font: F.lblb, border: BD.top });
  ['F', 'G'].forEach((c, i) => put(PR, c + '14', `=IF(${EMP[i].reg}="Simples",N(${EMP[i].sim}),SUM(${c}8:${c}12))`, { font: F.lblb, fmt: PCT, border: BD.top }));
  put(PR, 'E15', 'Crédito IPI+ICMS');
  ['F', 'G'].forEach((c, i) => put(PR, c + '15', `=${EMP[i].cred}`, { align: DIR }));
  const IMP = ['$F$14', '$G$14'], CRF = ['IF($F$15="Sim",1,0)', 'IF($G$15="Sim",1,0)'];

  secao(PR, 6, 'Comissão e desconto', 9, 10);
  put(PR, 'I7', 'Comissão'); entrada(PR, 'J7', num(d.comissao) ?? 0, PCT);
  nota(PR, 'J7', 'Comissão desta negociação. Digite aqui.');
  put(PR, 'I8', 'Desc. do tipo'); put(PR, 'J8', '=' + lt(K.dt), { fmt: PCT3 });
  const negIx = `INDEX(Estimativa!$M$${I0}:$M$${I1},$C$5)`;
  put(PR, 'I9', 'Desc. negociado'); put(PR, 'J9', `=IF(${negIx}="","—",${negIx})`, { fmt: PCT3, align: DIR });
  put(PR, 'I10', 'Desconto usado', { font: F.lblb, border: BD.top }); put(PR, 'J10', `=IF(${negIx}="",J8,${negIx})`, { font: F.lblb, fmt: PCT3, border: BD.top });
  put(PR, 'I11', 'negociado: coluna M dos itens', { font: F.nota });
  const COM = '$J$7', DESC = '$J$10';

  const MG = ['C', 'D', 'E', 'F', 'G', 'H', 'I'];
  secao(PR, 17, 'Preço de venda por margem', 2, 10);
  put(PR, 'J17', 'Preço informado', { font: F.sec, fill: P.sec, align: DIR });
  put(PR, 'B18', 'Margem desejada', { font: F.lblb });
  MG.forEach((c, i) => entrada(PR, c + '18', margens[i], PCT0));
  put(PR, 'J18', 'digite o preço ↓', { font: F.nota, align: DIR });
  const preco = (liq, imp) => { const den = `(1-{c}$18-${imp}-${COM})`; return `=IFERROR(IF(${den}<=0,"margem alta demais",(${liq})/${den}/(1-${DESC})),"")`; };
  const marg = (prow, imp, cr) => `=IFERROR(1-${imp}-${COM}-(${CUSTO}-${CRED}*${cr})/({c}${prow}*(1-${DESC})),"")`;
  const blocoEmp = (i, r0) => [
    ['sub', r0, `="Pela "&${EMP[i].nome}&IF(${EMP[i].reg}="Simples"," · Simples Nacional sobre a venda"," · impostos de saída por tipo de produto")&IF(${EMP[i].cred}="Sim",", com crédito de IPI + ICMS",", sem crédito")`],
    [r0 + 1, 'Preço de tabela, R$', preco(`${CUSTO}-${CRED}*${CRF[i]}`, IMP[i]), N2, true],
    [r0 + 2, 'Preço com desconto (venda), R$', `=IFERROR({c}${r0 + 1}*(1-${DESC}),"")`, N2, false],
    [r0 + 3, 'Fator de venda (preço de tabela ÷ FOB unitário)', `=IFERROR({c}${r0 + 1}/${FOBU},"")`, N4, false],
    [r0 + 4, 'Margem 2, %', marg(r0 + 1, IMP[i], CRF[i]), PCT, false],
  ];
  const escreve = lst => {
    for (const b of lst) {
      if (b[0] === 'sub') { subsecao(PR, b[1], b[2], 2, 10); continue; }
      const [r, l, f, fmt, neg] = b, font = neg ? F.lblb : F.lbl;
      put(PR, 'B' + r, l, { font });
      for (const c of MG) put(PR, c + r, f.replace(/\{c\}/g, c), { font, fmt, align: DIR });
    }
  };
  escreve(blocoEmp(0, 19));
  escreve(blocoEmp(1, 24));
  escreve([
    ['sub', 29, 'Por qual empresa é melhor trazer'],
    [30, `=${EMP[1].nome}&" − "&${EMP[0].nome}&": preço, R$ (preço informado: margem)"`, '=IFERROR({c}25-{c}20,"")', N2, false],
    [31, `=${EMP[1].nome}&" − "&${EMP[0].nome}&", % do preço"`, '=IFERROR({c}25/{c}20-1,"")', '+0.0%;-0.0%', false],
    [32, 'Melhor trazer pela', `=IFERROR(IF({c}25<{c}20,${EMP[1].nome},IF({c}25>{c}20,${EMP[0].nome},"tanto faz")),"")`, null, true],
  ]);
  subsecao(PR, 33, 'Prova real', 2, 10);
  put(PR, 'B34', 'Prova real pela', { font: F.lblb }); entrada(PR, 'C34', emps[0].nome || 'Empresa 1');
  lista(PR, 'C34', `Tabelas!$B$${ER[0]}:$B$${ER[1]}`);
  const EP = '$C$34', E2 = `${EMP[1].nome}`;
  escreve([
    [35, 'Preço de tabela', `=IFERROR(IF(${EP}=${E2},{c}25,{c}20),"")`, N2, false],
    [36, '(−) Desconto', `=IFERROR({c}35*${DESC},"")`, N2, false],
    [37, '(=) Venda', '=IFERROR({c}35-{c}36,"")', N2, true],
    [38, '(−) CMV (custo unitário)', `=${CUSTO}`, N2, false],
    [39, '(−) Impostos de saída', `=IFERROR({c}37*IF(${EP}=${E2},${IMP[1]},${IMP[0]}),"")`, N2, false],
    [40, '(−) Comissão', `=IFERROR({c}37*${COM},"")`, N2, false],
    [41, '(=) Margem 1, R$', '=IFERROR({c}37-{c}38-{c}39-{c}40,"")', N2, true],
    [42, 'Margem 1, %', '=IFERROR({c}41/{c}37,"")', PCT, false],
    [43, '(+) Créditos (IPI + ICMS)', `=${CRED}*IF(${EP}=${E2},${CRF[1]},${CRF[0]})`, N2, false],
    [44, '(=) Margem 2, R$', '=IFERROR({c}41+{c}43,"")', N2, true],
    [45, 'Margem 2, %', '=IFERROR({c}44/{c}37,"")', PCT, true],
    [46, 'Confere com a margem desejada?', '=IFERROR(IF(ABS({c}45-{c}18)<0.00005,"ok","verificar"),"")', null, false],
    ['sub', 48, `=${EMP[0].nome}&" sem crédito de impostos (referência, como na planilha antiga)"`],
    [49, 'Preço de tabela, R$', preco(CUSTO, IMP[0]), N2, false],
    [50, 'Fator de venda (preço de tabela ÷ FOB unitário)', `=IFERROR({c}49/${FOBU},"")`, N4, false],
  ]);
  entrada(PR, 'J20', num(d.preco_informado), N2);
  nota(PR, 'J20', 'Digite um preço de tabela (o que o cliente pediu, o do concorrente) para ver a margem em cada empresa.');
  const se = x => `=IF(N($J$20)>0,${x},"")`;
  [[21, se(`J20*(1-${DESC})`), N2, false], [22, se(`J20/${FOBU}`), N4, false],
    [23, se(`1-${IMP[0]}-${COM}-(${CUSTO}-${CRED}*${CRF[0]})/(J20*(1-${DESC}))`), PCT, true],
    [25, se('J20'), N2, true], [26, se(`J25*(1-${DESC})`), N2, false], [27, se(`J25/${FOBU}`), N4, false],
    [28, se(`1-${IMP[1]}-${COM}-(${CUSTO}-${CRED}*${CRF[1]})/(J25*(1-${DESC}))`), PCT, true],
    [30, se('J28-J23'), '+0.0%;-0.0%', false],
    [32, se(`IF(J28>J23,${EMP[1].nome},IF(J28<J23,${EMP[0].nome},"tanto faz"))`), null, true],
    [35, se(`IF(${EP}=${E2},J25,J20)`), N2, false], [36, se(`J35*${DESC}`), N2, false], [37, se('J35-J36'), N2, true],
    [38, se(CUSTO), N2, false], [39, se(`J37*IF(${EP}=${E2},${IMP[1]},${IMP[0]})`), N2, false], [40, se(`J37*${COM}`), N2, false],
    [41, se('J37-J38-J39-J40'), N2, true], [42, se('J41/J37'), PCT, false], [43, se(`${CRED}*IF(${EP}=${E2},${CRF[1]},${CRF[0]})`), N2, false],
    [44, se('J41+J43'), N2, true], [45, se('J44/J37'), PCT, true]]
    .forEach(([r, f, fmt, neg]) => put(PR, 'J' + r, f, { font: neg ? F.lblb : F.lbl, fmt, align: DIR }));
  for (const r of [20, 25]) for (const c of MG) PR.getCell(c + r).fill = P.dest;
  for (const c of [...MG, 'J']) PR.getCell(c + 32).fill = P.cen;
  for (const r of [37, 41, 44]) for (const c of ['B', ...MG, 'J']) PR.getCell(c + r).border = BD.top;

  const PC = ['H', 'I', 'J', 'K', 'L', 'M', 'N'];
  const tabelaItens = (t0, i) => {
    secao(PR, t0, '', 2, 14);
    put(PR, 'B' + t0, `="Tabela de preços pela "&${EMP[i].nome}&" · todos os itens"`, { font: F.sec, fill: P.sec });
    ['Item (tipo de produto)', 'Qtd', 'FOB unit., US$', 'Custo unit. (menos créditos, se usa), R$', 'Impostos de saída', 'Desconto'].forEach((h, j) =>
      put(PR, L(2 + j) + (t0 + 1), h, { font: F.hdr, align: { horizontal: j ? 'right' : 'left', wrapText: true }, border: BD.hdr }));
    PC.forEach((pc, j) => put(PR, pc + (t0 + 1), `=${MG[j]}$18`, { font: F.hdr, fmt: '"Margem "0%', align: DIR, border: BD.hdr }));
    PR.getRow(t0 + 1).height = 28;
    for (let k = 0; k < NIT; k++) {
      const r = t0 + 2 + k, cr = R0 + k, er = I0 + k;
      put(PR, 'B' + r, `=IF(Cálculo!G${cr}>0,Cálculo!B${cr}&" ("&Cálculo!C${cr}&")","")`, { font: F.lblb });
      put(PR, 'C' + r, `=IF($B${r}="","",Cálculo!D${cr})`, { fmt: INT });
      put(PR, 'D' + r, `=IF($B${r}="","",Cálculo!E${cr})`, { fmt: N2 });
      put(PR, 'E' + r, `=IF($B${r}="","",Cálculo!${ES.cuu}${cr}-Cálculo!${ES.cru}${cr}*IF(${EMP[i].cred}="Sim",1,0))`, { fmt: N2 });
      put(PR, 'F' + r, `=IF($B${r}="","",IF(${EMP[i].reg}="Simples",N(${EMP[i].sim}),IFERROR(VLOOKUP(Cálculo!$C${cr},${TAB},${K.tots},0),0)))`, { fmt: PCT });
      put(PR, 'G' + r, `=IF($B${r}="","",IF(Estimativa!M${er}="",IFERROR(VLOOKUP(Cálculo!$C${cr},${TAB},${K.dt},0),0),Estimativa!M${er}))`, { fmt: PCT3 });
      for (const pc of PC) put(PR, pc + r, `=IF($B${r}="","",IFERROR(IF((1-${pc}$${t0 + 1}-$F${r}-${COM})<=0,"margem alta demais",$E${r}/(1-${pc}$${t0 + 1}-$F${r}-${COM})/(1-$G${r})),""))`, { fmt: N2 });
    }
    return t0 + 2 + NIT;
  };
  let fim = tabelaItens(52, 0);
  fim = tabelaItens(fim + 1, 1);
  put(PR, 'B' + (fim + 1), 'Comissão da célula J7. Desconto: o negociado (coluna M dos itens) ou, se vazio, o do tipo de produto (aba Tabelas).', { font: F.nota });
  PR.views = [{ state: 'frozen', ySplit: 3, showGridLines: false }];

  return wb.xlsx.writeBuffer();
}

global.gerarExcelEstimativa = gerarExcelEstimativa;
if (typeof module !== 'undefined' && module.exports) module.exports = { gerarExcelEstimativa };
})(typeof window !== 'undefined' ? window : globalThis);
