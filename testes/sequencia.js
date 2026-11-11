#!/usr/bin/env node
// Sequência do item dentro da OC (aba "original" → PEDIDOS col V → Relatorio_DB col Y).
// Regressão de 11/11/2026: a aba "original" ganhou a coluna "Lote Dilly" e a coluna N virou DIAS; a chave de
// identificação usava índices fixos e nenhum item casava (0 de 2.668) — a sequência nunca foi gravada.
//   node testes/sequencia.js [--dados <pasta>]
const L = require('./lib');

const base = L.lerBase(L.pastaDaLinhaDeComando());
const rel = L.relatorio('sequência por OC (aba original)');
const NOVA = new Date(2026, 10, 12); // depois do corte de 11/11/2026
const ANTIGA = new Date(2026, 8, 21);
const CAB = ['COD. ', 'NOME CLIENTE ', 'PEDIDO ', 'COD CLIENTE', 'COD ITEM', 'DESCRIÇÃO', 'TAM.', 'OC', 'QTD ABERTA',
  'Lote Dilly', 'COD OS', 'DT RECEB', 'DT ENTREGA', 'DIAS', 'LOTES', 'NOVA DATA'];
// linha da aba "original" no layout ATUAL (com "Lote Dilly" e DIAS)
const linhaOrig = r => ['2', r[2], r[4], r[5], r[6], r[7], r[8], r[9], r[10], '000000000 OC: ' + r[9], r[11], r[12], r[13], Math.floor(Math.random() * 90) - 20, '146413', ''];
// layout LEGADO (sem cabeçalho nem "Lote Dilly"): C..K = pedido, cód. cliente, marfim, desc, tam, OC, QTD, OS, data
const linhaLegado = r => ['2', r[2], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13], 'x', '146413'];

function escolherOcs(env) {
  const por = new Map();
  L.imp(env).forEach((r, i) => {
    if (i < 3 || !L.T(r[1]) || !L.T(r[9])) return;
    const oc = L.T(r[9]);
    if (!por.has(oc)) por.set(oc, []);
    por.get(oc).push(i);
  });
  const ocs = [...por.entries()].filter(([, v]) => v.length >= 3 && v.length <= 8).map(([k]) => k);
  return { por, ocs };
}

function cenario(modo, layout) {
  const env = L.montar(base, modo);
  const { por, ocs } = escolherOcs(env);
  if (ocs.length < 3) return null;
  const [ocNova, ocParcial, ocAntiga] = ocs;
  const imp = L.imp(env);
  por.get(ocNova).concat(por.get(ocParcial)).forEach(i => { imp[i][12] = NOVA; });
  por.get(ocAntiga).forEach(i => { imp[i][12] = ANTIGA; });
  const linhas = [];
  por.get(ocNova).forEach(i => linhas.push(imp[i]));
  por.get(ocParcial).slice(1).forEach(i => linhas.push(imp[i])); // original com UMA linha a menos que a fonte
  por.get(ocAntiga).forEach(i => linhas.push(imp[i]));
  const ss = env.ss;
  if (layout === 'atual') ss.addSheet('original', [CAB, ...linhas.map(linhaOrig)], 24);
  else ss.addSheet('original', linhas.map(linhaLegado), 14);
  L.sincronizar(env);
  const pedOc = oc => L.ped(env).filter((r, i) => i >= 3 && L.T(r[9]) === oc);
  const dbOc = oc => L.db(env).filter((r, i) => i > 0 && L.T(r[8]) === oc);
  return { env, ocNova, ocParcial, ocAntiga, pedOc, dbOc, n: por.get(ocNova).length };
}

for (const modo of ['SIMULACAO', 'ATIVO']) {
  for (const layout of ['atual', 'legado']) {
    rel.secao(`aba "original" no layout ${layout} [${modo}]`);
    const c = cenario(modo, layout);
    if (!c) { rel.pular('base sem 3 OCs de 3 a 8 itens'); continue; }
    const seqs = c.pedOc(c.ocNova).map(r => Number(r[21]) || 0).sort((a, b) => a - b);
    rel.check(`OC nova e completa: PEDIDOS col V = 1..${c.n}`, seqs.length === c.n && seqs.every((v, i) => v === i + 1), JSON.stringify(seqs));
    const dbSeq = c.dbOc(c.ocNova).map(r => Number(r[24]) || 0).filter(v => v > 0).sort((a, b) => a - b);
    rel.check('OC nova e completa: Relatorio_DB col Y propagada', dbSeq.length === c.n && dbSeq.every((v, i) => v === i + 1), JSON.stringify(dbSeq));
    rel.check('OC nova mas incompleta na aba original: sem sequência (nunca número errado)',
      c.pedOc(c.ocParcial).every(r => !Number(r[21])) && c.dbOc(c.ocParcial).every(r => !Number(r[24])));
    rel.check('OC antiga (recebida antes do corte): sem sequência',
      c.pedOc(c.ocAntiga).every(r => !Number(r[21])) && c.dbOc(c.ocAntiga).every(r => !Number(r[24])));
    // 2ª rodada: a sequência é fixa — nada muda
    const antes = JSON.stringify(c.pedOc(c.ocNova).map(r => [L.T(r[0]), r[21]]));
    L.sincronizar(c.env);
    rel.check('2ª rodada não muda a sequência', JSON.stringify(c.pedOc(c.ocNova).map(r => [L.T(r[0]), r[21]])) === antes);
  }
}
rel.fim();
