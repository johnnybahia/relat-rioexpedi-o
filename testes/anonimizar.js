#!/usr/bin/env node
// Converte exports REAIS (DADOS_IMPORTADOS, PEDIDOS, Relatorio_DB em CSV) numa base FICTÍCIA com a mesma
// estrutura, para poder ficar no repositório (que é público).
//
//   node testes/anonimizar.js <pasta_com_os_3_csv_reais> [pasta_saida=testes/dados/base]
//
// Troca, sempre pelo mesmo valor nas 3 abas: CLIENTE (mantém "DILLY"/"DAKOTA", que têm regras próprias no
// código), CÓD. FILIAL, PEDIDO, CÓD. CLIENTE e CÓD. MARFIM (parte antes do último "-"; o sufixo fica, por causa
// da normalização Dilly), DESCRIÇÃO, ORD. COMPRA, CÓD. OS (vazio e "0" ficam), LOTE, MARCA e nomes de usuário
// (col V e col Z do DB). O ID_UNICO é remontado com os valores fictícios (mesma fórmula do código), então a
// base do ID continua batendo com a linha da origem. Mantém: QTD, datas, tamanho, status, UUIDs, FAT-xxx.
// No fim confere se algum valor real sobrou nos arquivos gerados (e aborta se sobrou).
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { lerCsv, escreverCsv } = require('./gasmock');

const [,, pastaEntrada, pastaSaidaArg] = process.argv;
if (!pastaEntrada) {
  console.error('uso: node testes/anonimizar.js <pasta_com_csv_reais> [pasta_saida]');
  process.exit(2);
}
const pastaSaida = path.resolve(pastaSaidaArg || path.join(__dirname, 'dados', 'base'));

const { acharCsv } = require('./lib');
const imp = lerCsv(acharCsv(pastaEntrada, 'DADOS_IMPORTADOS'));
const ped = lerCsv(acharCsv(pastaEntrada, 'PEDIDOS'));
const db  = lerCsv(acharCsv(pastaEntrada, 'Relatorio_DB'));

const T = v => String(v === null || v === undefined ? '' : v).trim();
const RE_UUID = /^(.*?)\s*\[([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\]\s*$/;

// Todos os valores reais das colunas trocadas: um valor fictício nunca pode coincidir com um deles
// (senão a conferência final não distingue coincidência de vazamento).
const reais = new Set();
function coletar(v) {
  const s = T(v); if (!s) return;
  reais.add(s);
  const i = s.lastIndexOf('-'); if (i > 0) reais.add(s.slice(0, i));
  const m = RE_UUID.exec(s); if (m) reais.add(T(m[1]));
  if (s.startsWith('{')) { try { Object.values(JSON.parse(s)).forEach(x => reais.add(T(x))); } catch (e) { /* não é JSON */ } }
}
imp.slice(3).forEach(r => [2, 3, 4, 5, 6, 7, 9, 11, 23, 24].forEach(j => coletar(r[j])));
ped.slice(3).forEach(r => [2, 3, 4, 5, 6, 7, 9, 11, 19, 20].forEach(j => coletar(r[j])));
db.slice(1).forEach(r => [2, 3, 4, 5, 6, 8, 10, 19, 20, 21].forEach(j => coletar(r[j])));

class Mapa {
  constructor(gerar) { this.m = new Map(); this.gerar = gerar; this.seq = 0; }
  get(v) {
    const k = T(v); if (!k) return '';
    if (!this.m.has(k)) {
      let c;
      do { c = this.gerar(++this.seq, k); } while (reais.has(c));
      this.m.set(k, c);
    }
    return this.m.get(k);
  }
}
const n2 = n => String(n).padStart(2, '0');
const clientes = new Mapa((n, v) => { const u = v.toUpperCase(); return u.includes('DILLY') ? `DILLY CLIENTE ${n2(n)}` : u.includes('DAKOTA') ? `DAKOTA CLIENTE ${n2(n)}` : `CLIENTE ${n2(n)}`; });
const filiais  = new Mapa(n => String(n));
const pedidos  = new Mapa(n => String(100000 + n));
const baseCod  = new Mapa(n => String(500000 + n));
const baseMarf = new Mapa(n => String(80000 + n));
const descs    = new Mapa(n => `PRODUTO ${String(n).padStart(4, '0')}`);
const ocs      = new Mapa((n, v) => String(700000 + n) + (/[A-Z]$/i.test(v) ? v.slice(-1).toUpperCase() : ''));
const oss      = new Mapa(n => String(260000 + n));
const lotes    = new Mapa(n => String(300000 + n));
const marcas   = new Mapa(n => `MARCA ${n2(n)}`);
const usuarios = new Mapa(n => `USUARIO ${n2(n)}`);

// Código com "-": troca a parte antes do último "-" e mantém o sufixo (a normalização Dilly troca só o sufixo).
function mapCodigo(v, mapa) {
  const s = T(v); if (!s) return '';
  const i = s.lastIndexOf('-');
  return i > 0 ? mapa.get(s.slice(0, i)) + s.slice(i) : mapa.get(s);
}
const mapOs = v => { const s = T(v); return (s === '' || s === '0') ? s : oss.get(s); };
function mapDesc(v) {
  const s = String(v === null || v === undefined ? '' : v);
  const m = RE_UUID.exec(s);
  return m ? `${descs.get(m[1])} [${m[2]}]` : descs.get(s);
}
function mapUsuario(v) {
  const s = T(v); if (!s) return '';
  if (s.startsWith('{')) {
    try { const o = JSON.parse(s); Object.keys(o).forEach(k => { o[k] = usuarios.get(o[k]); }); return JSON.stringify(o); } catch (e) { /* segue */ }
  }
  return ['sistema', 'reparo', 'menu'].includes(s) ? s : usuarios.get(s);
}
function mapConferencia(v) {
  const s = T(v); if (!s) return '';
  const p = s.split('|');
  if (p.length > 1) p[1] = mapUsuario(p[1]);
  return p.join('|');
}

// ── ID_UNICO: {CLIENTE}{FILIAL}{PEDIDO}{MARFIM}{TAMANHO}{OC}{OS}{DATA yyyyMMdd}-{n}[-DUPk] ────────────────
const dataId = v => { const m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(T(v)); return m ? m[3] + m[2] + m[1] : T(v); };
const partesId = id => { const m = /^(.*?)(-\d+(?:-DUP\d+)?)$/.exec(id); return m ? [m[1], m[2]] : [id, '']; };
const baseReal = c => T(c.cliente) + T(c.filial) + T(c.pedido) + T(c.marfim) + T(c.tam) + T(c.oc) + T(c.os) + dataId(c.data);
const baseFake = c => clientes.get(c.cliente) + filiais.get(c.filial) + pedidos.get(c.pedido) + mapCodigo(c.marfim, baseMarf) +
                      T(c.tam) + ocs.get(c.oc) + mapOs(c.os) + dataId(c.data);
// OS "0" numérico entra no ID como "" (String(0 || '')): tenta as duas grafias.
const variantes = c => (T(c.os) === '0') ? [c, Object.assign({}, c, { os: '' })] : [c];
const baseParaFake = new Map();
let idsPorCampos = 0, idsPorHash = 0;
function resolverBase(base, candidatos) {
  if (baseParaFake.has(base)) return baseParaFake.get(base);
  for (const c of candidatos) for (const v of variantes(c)) {
    if (baseReal(v) === base) { const f = baseFake(v); baseParaFake.set(base, f); idsPorCampos++; return f; }
  }
  return null;
}
const idMap = new Map();
function mapId(id, candidatos) {
  const s = T(id); if (!s) return '';
  if (idMap.has(s)) return idMap.get(s);
  const [base, sufixo] = partesId(s);
  let fake = resolverBase(base, candidatos || []);
  if (!fake) {
    fake = baseParaFake.get(base);
    if (!fake) {
      fake = 'ANON' + crypto.createHash('md5').update(base).digest('hex').slice(0, 10).toUpperCase();
      baseParaFake.set(base, fake); idsPorHash++;
    }
  }
  idMap.set(s, fake + sufixo);
  return fake + sufixo;
}

// ── DADOS_IMPORTADOS (A..Z; dados a partir da linha 4) ──────────────────────────────────────────────────
const saidaImp = imp.map((r, i) => {
  if (i < 2) return r.map((v, j) => (j === 7 ? v : ''));  // linhas 1-2: só H1/H2 (datas; H2 é a trava do sync)
  if (i === 2) return r.slice();                             // linha 3: cabeçalhos
  const o = r.slice();
  while (o.length < 26) o.push('');
  o[2] = clientes.get(r[2]); o[3] = filiais.get(r[3]); o[4] = pedidos.get(r[4]);
  o[5] = mapCodigo(r[5], baseCod); o[6] = mapCodigo(r[6], baseMarf); o[7] = r[7] === '' ? '' : descs.get(r[7]);
  o[9] = ocs.get(r[9]); o[11] = mapOs(r[11]);
  for (let j = 15; j <= 22; j++) o[j] = '';                 // P..W: não usadas pelo sistema
  o[23] = marcas.get(r[23]); o[24] = lotes.get(r[24]); o[25] = '';
  return o;
});
// Linhas da origem com CARTELA, na ordem: PEDIDOS é gravado nessa mesma ordem.
const origemComCartela = imp.slice(3).filter(r => T(r[1]));

// ── PEDIDOS (A..V; dados a partir da linha 4) ───────────────────────────────────────────────────────────
let k = 0;
const saidaPed = ped.map((r, i) => {
  if (i < 3 || !r.some(v => T(v))) return r.slice();
  const o = r.slice();
  const f = T(r[1]) ? origemComCartela[k++] : null;       // linha da origem na mesma posição
  const cands = [];
  if (f) cands.push({ cliente: f[2], filial: f[3], pedido: f[4], marfim: f[6], tam: f[8], oc: f[9], os: f[11], data: f[12] });
  cands.push({ cliente: r[2], filial: r[3], pedido: r[4], marfim: r[6], tam: r[8], oc: r[9], os: r[11], data: r[12] });
  o[0] = mapId(r[0], cands);
  o[2] = clientes.get(r[2]); o[3] = filiais.get(r[3]); o[4] = pedidos.get(r[4]);
  o[5] = mapCodigo(r[5], baseCod); o[6] = mapCodigo(r[6], baseMarf); o[7] = mapDesc(r[7]);
  o[9] = ocs.get(r[9]); o[11] = mapOs(r[11]);
  o[19] = marcas.get(r[19]); o[20] = lotes.get(r[20]);
  return o;
});

// ── Relatorio_DB (cabeçalho na linha 1; sem CÓD. FILIAL) ─────────────────────────────────────────────────
const saidaDb = db.map((r, i) => {
  if (i === 0 || !T(r[0])) return r.slice();
  const o = r.slice();
  // FILIAL não existe no DB: sai do próprio ID (entre CLIENTE e o resto)
  const [base] = partesId(T(r[0]));
  const resto = T(r[3]) + T(r[5]) + T(r[7]) + T(r[8]) + T(r[10]) + dataId(r[11]);
  const cands = [];
  [T(r[10]), ...(T(r[10]) === '0' ? [''] : [])].forEach(os => {
    const restoOs = T(r[3]) + T(r[5]) + T(r[7]) + T(r[8]) + os + dataId(r[11]);
    if (base.startsWith(T(r[2])) && base.endsWith(restoOs)) {
      const filial = base.slice(T(r[2]).length, base.length - restoOs.length);
      cands.push({ cliente: r[2], filial, pedido: r[3], marfim: r[5], tam: r[7], oc: r[8], os, data: r[11] });
    }
  });
  void resto;
  o[0] = mapId(r[0], cands);
  o[2] = clientes.get(r[2]); o[3] = pedidos.get(r[3]);
  o[4] = mapCodigo(r[4], baseCod); o[5] = mapCodigo(r[5], baseMarf); o[6] = mapDesc(r[6]);
  o[8] = ocs.get(r[8]); o[10] = mapOs(r[10]);
  o[19] = marcas.get(r[19]); o[20] = lotes.get(r[20]);
  o[21] = mapUsuario(r[21]);
  if (o.length > 25) o[25] = mapConferencia(r[25]);
  return o;
});

fs.mkdirSync(pastaSaida, { recursive: true });
escreverCsv(path.join(pastaSaida, 'DADOS_IMPORTADOS.csv'), saidaImp);
escreverCsv(path.join(pastaSaida, 'PEDIDOS.csv'), saidaPed);
escreverCsv(path.join(pastaSaida, 'Relatorio_DB.csv'), saidaDb);

// ── Conferência: nenhum valor real pode ter sobrado ──────────────────────────────────────────────────────
const sensiveis = new Set();
const guardar = (v, min) => { const s = T(v); if (s.length >= min) sensiveis.add(s); };
[clientes, descs, usuarios, marcas].forEach(mp => mp.m.forEach((fake, real) => guardar(real, 3)));
[pedidos, ocs, oss, lotes, baseCod, baseMarf].forEach(mp => mp.m.forEach((fake, real) => guardar(real, 4)));
const reaisTexto = [...clientes.m.keys(), ...usuarios.m.keys(), ...descs.m.keys()].filter(s => s.length >= 5);
let vazamentos = 0;
const exemplos = [];
[saidaImp.slice(3), saidaPed.slice(3), saidaDb.slice(1)].forEach((linhas, a) => linhas.forEach(r => r.forEach((v, j) => {
  const s = T(v); if (!s) return;
  const base = s.replace(RE_UUID, '$1').trim();
  let achou = sensiveis.has(s) || sensiveis.has(base);
  if (!achou && /[A-Za-z]{3}/.test(s)) achou = reaisTexto.some(real => s.includes(real));
  if (achou) { vazamentos++; if (exemplos.length < 5) exemplos.push(`${['DADOS_IMPORTADOS', 'PEDIDOS', 'Relatorio_DB'][a]} col ${j + 1}: "${s}"`); }
})));

console.log(`Base fictícia gravada em ${pastaSaida}`);
console.log(`  linhas: DADOS_IMPORTADOS ${saidaImp.length - 3} · PEDIDOS ${saidaPed.length - 3} · Relatorio_DB ${saidaDb.length - 1}`);
console.log(`  trocados: ${clientes.m.size} clientes, ${pedidos.m.size} pedidos, ${descs.m.size} descrições, ${ocs.m.size} OCs, ${lotes.m.size} lotes, ${usuarios.m.size} usuários`);
console.log(`  IDs: ${idMap.size} (${idsPorCampos} bases remontadas pelos campos, ${idsPorHash} por código fictício)`);
if (vazamentos > 0) {
  console.error(`❌ ${vazamentos} valor(es) real(is) sobraram — NÃO subir:\n  ${exemplos.join('\n  ')}`);
  process.exit(1);
}
console.log('✅ nenhum valor real encontrado nos arquivos gerados');
