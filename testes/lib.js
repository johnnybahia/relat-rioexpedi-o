// Utilitários comuns da bateria de testes: carregar uma base (3 CSV), montar a planilha simulada,
// rodar o sync, localizar grupos de linhas-irmãs pela estrutura e relatar ✅/❌.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { criarContexto, lerCsv, celula } = require('./gasmock');

const PASTA_DADOS = path.join(__dirname, 'dados');
const BASE_PADRAO = path.join(PASTA_DADOS, 'base');
const USUARIO_TESTE = '{"BAIXA1":"USUARIO TESTE"}';

// ── Arquivos ────────────────────────────────────────────────────────────────────────────────────────────
const PADROES = {
  DADOS_IMPORTADOS: /DADOS_IMPORTADOS(_\d+| \(\d+\))?$/,
  PEDIDOS: /(^|[^A-Z])PEDIDOS(_\d+| \(\d+\))?$/,
  Relatorio_DB: /RELATORIO_DB(_\d+| \(\d+\))?$/
};
/** CSV de uma aba numa pasta (aceita os nomes de export do Google Sheets e de upload); o mais recente vence. */
function acharCsv(pasta, aba) {
  const norm = f => f.toUpperCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/\.CSV$/, '');
  const achados = fs.readdirSync(pasta).filter(f => /\.csv$/i.test(f) && PADROES[aba].test(norm(f)))
    .map(f => path.join(pasta, f))
    .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
  if (achados.length === 0) throw new Error(`CSV da aba ${aba} não encontrado em ${pasta}`);
  return achados[0];
}
const _cacheBases = new Map();
function lerBase(pasta) {
  if (!_cacheBases.has(pasta)) {
    _cacheBases.set(pasta, {
      pasta,
      imp: lerCsv(acharCsv(pasta, 'DADOS_IMPORTADOS')),
      ped: lerCsv(acharCsv(pasta, 'PEDIDOS')),
      db: lerCsv(acharCsv(pasta, 'Relatorio_DB'))
    });
  }
  return _cacheBases.get(pasta);
}
/** Bases guardadas no repositório (testes/dados/<nome>/ com os 3 CSV). */
function basesDoRepositorio() {
  return fs.readdirSync(PASTA_DADOS).map(n => path.join(PASTA_DADOS, n))
    .filter(p => fs.statSync(p).isDirectory() && fs.readdirSync(p).some(f => /\.csv$/i.test(f)))
    .sort();
}
/** --dados <pasta> na linha de comando (exports novos, fora do Git); senão a base padrão. */
function pastaDaLinhaDeComando() {
  const i = process.argv.indexOf('--dados');
  return i > 0 && process.argv[i + 1] ? path.resolve(process.argv[i + 1]) : BASE_PADRAO;
}

// ── Planilha simulada ───────────────────────────────────────────────────────────────────────────────────
/**
 * Monta as abas a partir da base. Tipos como o Sheets devolve: números e datas convertidos; PEDIDOS F/G e
 * Relatorio_DB V/W/X/Z ficam texto. Baixas_Historico começa vazia (os CSV não trazem essa aba).
 */
function montar(base, modo, opcoes = {}) {
  const env = criarContexto({ verbose: !!opcoes.verbose });
  const { ss } = env;
  const txtPed = new Set([5, 6]);
  const txtDb = new Set([21, 22, 23, 25]);
  ss.addSheet('DADOS_IMPORTADOS', base.imp.map(r => r.map(v => celula(v, false))), 26);
  ss.addSheet('PEDIDOS', base.ped.map(r => r.map((v, j) => celula(v, txtPed.has(j)))), 22);
  ss.addSheet('Relatorio_DB', base.db.map((r, i) => i === 0 ? r.slice() : r.map((v, j) => celula(v, txtDb.has(j)))), 26);
  ss.addSheet('Baixas_Historico', [['ID_ITEM', 'DATA_HORA', 'QTD_BAIXADA', 'QTD_RESTANTE', 'QTD_ORIGINAL', 'USUARIO', 'TIPO']], 7);
  ss.addSheet('CONFIGURAÇÕES', [['Hora da limpeza', 11], ['', ''], ['', ''], ['Modo', modo]], 2);
  ss.addSheet('CADASTRO', [['ANA', '1', 'TOTAL', '', 15], ['BIA', '2', 'PARCIAL']], 5);
  return env;
}
/** Um ciclo completo do sync (forçado) + gravação da auditoria. */
function sincronizar(env) {
  const r1 = env.ctx.sincronizarPedidosComFonte(true);
  const r2 = env.ctx.sincronizarDados();
  vm.runInContext('_gravarAuditoria_()', env.ctx);
  return { pedidos: r1, dados: r2 };
}
const aba = (env, nome) => env.ss.getSheetByName(nome);
const imp = env => aba(env, 'DADOS_IMPORTADOS').data;   // linhas A..Z; dados a partir do índice 3
const ped = env => aba(env, 'PEDIDOS').data;            // dados a partir do índice 3
const db  = env => aba(env, 'Relatorio_DB').data;       // índice 0 = cabeçalho
const T = v => String(v === null || v === undefined ? '' : v).trim();
const dbPorId = (env, id) => db(env).find((r, i) => i > 0 && T(r[0]) === id) || null;
const idPedidosPorLote = (env, lote) => { const r = ped(env).find((x, i) => i >= 3 && T(x[20]) === T(lote)); return r ? T(r[0]) : null; };
const linhaOrigemPorLote = (env, lote) => imp(env).findIndex((x, i) => i >= 3 && T(x[24]) === T(lote));
function alterarDb(env, id, campos) {
  const r = dbPorId(env, id);
  if (!r) throw new Error('linha do DB não encontrada: ' + id);
  while (r.length < 26) r.push('');
  Object.entries(campos).forEach(([k, v]) => { r[Number(k)] = v; });
}
function registrarBaixa(env, id, baixada, restante, original) {
  aba(env, 'Baixas_Historico').data.push([id, new Date(), baixada, restante, original, 'USUARIO TESTE', '']);
}
function auditoria(env, tipo) {
  const sh = aba(env, 'Auditoria_Sincronizacao');
  return sh ? sh.data.slice(1).filter(a => !tipo || a[2] === tipo) : [];
}
const statusFinal = s => ['Faturado', 'Finalizado', 'Excluido'].includes(T(s));
const conferencia = r => T(r[25]).split('|')[0].toUpperCase();
// Faturado sem prova de usuário: col V vazia e col Z sem "FATURADO|<alguém>|…"
const faturadoSemUsuario = r => T(r[14]) === 'Faturado' && !T(r[21]) && !(conferencia(r) === 'FATURADO' && T(T(r[25]).split('|')[1]));
const resumoLinha = r => r ? `${r[14]}|QTD=${r[9]}|P=${T(r[15])}|V=${T(r[21])}|Z=${T(r[25]).split('|').slice(0, 2).join('|')}|LOTE=${T(r[20])}` : '(ausente)';

// ── Localizar situações pela estrutura (vale para qualquer base) ────────────────────────────────────────
/** Grupos de linhas-irmãs da origem (mesma impressão digital), com LOTE único em PEDIDOS e na origem. */
function gruposDeIrmas(env) {
  const contaLotePed = new Map(), contaLoteImp = new Map();
  ped(env).forEach((r, i) => { if (i >= 3 && T(r[20])) contaLotePed.set(T(r[20]), (contaLotePed.get(T(r[20])) || 0) + 1); });
  imp(env).forEach((r, i) => { if (i >= 3 && T(r[24])) contaLoteImp.set(T(r[24]), (contaLoteImp.get(T(r[24])) || 0) + 1); });
  const grupos = new Map();
  imp(env).forEach((r, i) => {
    if (i < 3 || !T(r[1])) return;
    const fp = env.ctx._criarImpressaoDigitalFromRow_(r.slice(1), 0);
    if (!grupos.has(fp)) grupos.set(fp, []);
    grupos.get(fp).push({ i, lote: T(r[24]), qtd: Number(r[10]) || 0, linha: r });
  });
  return [...grupos.values()].filter(g => g.length >= 2
    && g.every(m => m.lote && contaLotePed.get(m.lote) === 1 && contaLoteImp.get(m.lote) === 1))
    .map(g => g.map(m => Object.assign(m, { id: idPedidosPorLote(env, m.lote) })))
    .filter(g => g.every(m => m.id && dbPorId(env, m.id)));
}
/** Linha da origem sem irmãs, com LOTE único, item Ativo e sem marcação no DB. */
function linhasSozinhas(env) {
  const grupos = new Map();
  imp(env).forEach((r, i) => {
    if (i < 3 || !T(r[1])) return;
    const fp = env.ctx._criarImpressaoDigitalFromRow_(r.slice(1), 0);
    grupos.set(fp, (grupos.get(fp) || []).concat([i]));
  });
  const lotesRepetidos = new Set();
  const vistos = new Set();
  ped(env).forEach((r, i) => { if (i >= 3 && T(r[20])) { if (vistos.has(T(r[20]))) lotesRepetidos.add(T(r[20])); vistos.add(T(r[20])); } });
  const out = [];
  grupos.forEach(idx => {
    if (idx.length !== 1) return;
    const r = imp(env)[idx[0]];
    const lote = T(r[24]);
    if (!lote || lotesRepetidos.has(lote)) return;
    const id = idPedidosPorLote(env, lote);
    const d = id && dbPorId(env, id);
    if (!d || T(d[14]) !== 'Ativo' || T(d[15]) || T(d[25])) return;
    out.push({ i: idx[0], lote, id, linha: r });
  });
  return out;
}
/** Base do ID (sem sufixo) igual à calculada da linha da origem — condição da reserva pela base do ID. */
const baseDoIdConfere = (env, m) => env.ctx._idBaseFonte_(m.linha.slice(1)) === m.id.replace(/-\d+(?:-DUP\d+)?$/, '');

// ── Relatório ───────────────────────────────────────────────────────────────────────────────────────────
function relatorio(titulo) {
  let falhas = 0, oks = 0, pulados = 0;
  console.log(`\n══ ${titulo}`);
  return {
    secao: t => console.log(`\n■ ${t}`),
    info: t => console.log(`   · ${t}`),
    check(nome, cond, detalhe) {
      if (cond) oks++; else falhas++;
      console.log(`   ${cond ? '✅' : '❌'} ${nome}${detalhe && !cond ? ' — ' + detalhe : ''}`);
      return !!cond;
    },
    pular(motivo) { pulados++; console.log(`   ⏭️  pulado: ${motivo}`); },
    fim() {
      console.log(`\nRESULTADO ${titulo}: ${oks} ok, ${falhas} falha(s), ${pulados} pulado(s)`);
      process.exitCode = falhas > 0 ? 1 : 0;
      return falhas;
    }
  };
}

module.exports = {
  PASTA_DADOS, BASE_PADRAO, USUARIO_TESTE, acharCsv, lerBase, basesDoRepositorio, pastaDaLinhaDeComando,
  montar, sincronizar, imp, ped, db, T, dbPorId, idPedidosPorLote, linhaOrigemPorLote, alterarDb, registrarBaixa,
  auditoria, statusFinal, conferencia, faturadoSemUsuario, resumoLinha, gruposDeIrmas, linhasSozinhas, baseDoIdConfere,
  relatorio
};
