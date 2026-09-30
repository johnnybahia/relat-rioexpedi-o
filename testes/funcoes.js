#!/usr/bin/env node
// Funções de apoio, nos dois modos: sentinela, reparo dos faturados sem usuário, dados que a tela recebe
// (pendentesSaida), resposta ao aviso (confirmarSaidaFonte) e limpeza diária (purgarItensFinalizados).
//   node testes/funcoes.js [--dados <pasta>]
const vm = require('vm');
const L = require('./lib');

const base = L.lerBase(L.pastaDaLinhaDeComando());
const rel = L.relatorio('FUNÇÕES (reparo, aviso, limpeza, sentinela)');
const pendentes = env => L.db(env).slice(1).filter(r => L.T(r[14]) === 'Ativo' && L.conferencia(r) === 'PENDENTE');
const contar = (env, st) => L.db(env).slice(1).filter(r => L.T(r[14]) === st).length;

for (const modo of ['SIMULACAO', 'ATIVO']) {
  rel.secao(`sentinela e reparo [${modo}]`);
  const env = L.montar(base, modo);
  const ctx = env.ctx;
  L.sincronizar(env);

  // Faturados sem usuário e quantos têm gêmea viva (mesmo PEDIDO+LOTE numa linha aberta, sem pendência)
  const vivas = new Set(L.db(env).slice(1).filter(r => !L.statusFinal(r[14]) && L.conferencia(r) !== 'PENDENTE' && L.T(r[20]))
    .map(r => `${L.T(r[3])}|${L.T(r[20])}`));
  const fsu = L.db(env).slice(1).filter(L.faturadoSemUsuario);
  const comGemea = fsu.filter(r => L.T(r[20]) && vivas.has(`${L.T(r[3])}|${L.T(r[20])}`)).length;
  const sent = JSON.parse(env.props.get('ALERTA_DUPLICATAS') || '{}');
  rel.check('sentinela conta os faturados sem usuário', (sent.faturadosSemUsuario || 0) === fsu.length, `sentinela ${sent.faturadosSemUsuario}, base ${fsu.length}`);

  const pendAntes = pendentes(env).length;
  if (fsu.length === 0) {
    rel.pular('reparo: base sem faturado sem usuário');
  } else {
    ctx.repararFaturadosSemUsuario();
    vm.runInContext('_gravarAuditoria_()', ctx);
    const aba = env.ss.getSheetByName('Reparo_Faturados');
    const acoes = {};
    (aba ? aba.data.slice(1) : []).forEach(r => { acoes[r[1]] = (acoes[r[1]] || 0) + 1; });
    rel.check(`reparo: ${fsu.length - comGemea} voltam para conferência e ${comGemea} ficam para decisão manual`,
      (acoes['VOLTOU PARA CONFERÊNCIA'] || 0) === fsu.length - comGemea && (acoes['DECISÃO MANUAL (gêmea aberta)'] || 0) === comGemea, JSON.stringify(acoes));
    rel.check('reparo: os que voltaram viram Ativo + PENDENTE', pendentes(env).length === pendAntes + fsu.length - comGemea);
  }

  // Garante pelo menos 3 pendências: 3 linhas sem marcação saem da origem
  const sozinhas = L.linhasSozinhas(env).slice(0, 3);
  const pendAntesSaida = pendentes(env).length;
  sozinhas.map(m => m.i).sort((x, y) => y - x).forEach(i => L.imp(env).splice(i, 1));
  L.sincronizar(env);
  const pend1 = pendentes(env).length;
  rel.check(`${sozinhas.length} linhas sem marcação que saíram viram PENDENTE (nenhuma faturada)`, pend1 === pendAntesSaida + sozinhas.length,
    `pendentes ${pendAntesSaida}→${pend1}`);
  L.sincronizar(env);
  rel.check('sync seguinte não refatura nem mexe nas pendências', pendentes(env).length === pend1 && L.db(env).slice(1).filter(L.faturadoSemUsuario).length === comGemea);

  rel.secao(`tela e resposta ao aviso [${modo}]`);
  const payload = ctx.fetchAllDataUnified(Date.now());
  const ps = payload.pendentesSaida || [];
  rel.check('tela recebe todas as pendências com os campos do aviso', ps.length === pend1 && payload.stats.pendentesSaida === pend1
    && ps.every(p => p.uniqueId && p.planilhaLinha && p.cliente && p.saiuEm !== undefined), `payload ${ps.length}, DB ${pend1}`);
  if (ps.length < 3) {
    rel.pular('resposta ao aviso: menos de 3 pendências');
  } else {
    const [p0, p1, p2] = ps;
    const rParcial = ctx.confirmarSaidaFonte(p0.uniqueId, p0.planilhaLinha, 'FATURADO', 'BIA');
    rel.check('usuário PARCIAL não pode responder', rParcial.success === false);
    const rFat = ctx.confirmarSaidaFonte(p0.uniqueId, p0.planilhaLinha, 'FATURADO', 'ANA');
    const rowFat = L.dbPorId(env, p0.uniqueId);
    rel.check('FATURADO → Faturado com Z = FATURADO|ANA', rFat.success && rowFat[14] === 'Faturado' && L.T(rowFat[25]).startsWith('FATURADO|ANA'), JSON.stringify(rFat));
    rel.check('segunda resposta ao mesmo aviso é recusada', ctx.confirmarSaidaFonte(p0.uniqueId, p0.planilhaLinha, 'ABERTO', 'ANA').success === false);
    const rCan = ctx.confirmarSaidaFonte(p1.uniqueId, p1.planilhaLinha + 1, 'CANCELADO', 'ANA'); // número de linha desatualizado
    const rowCan = L.dbPorId(env, p1.uniqueId);
    rel.check('CANCELADO com linha desatualizada grava na linha certa (Excluido)', rCan.success && rowCan[14] === 'Excluido' && L.T(rowCan[25]).startsWith('CANCELADO|ANA'));
    const rAb = ctx.confirmarSaidaFonte(p2.uniqueId, p2.planilhaLinha, 'ABERTO', 'ANA');
    rel.check('ABERTO → continua Ativo com Z = ABERTO|ANA', rAb.success && L.dbPorId(env, p2.uniqueId)[14] === 'Ativo' && L.T(L.dbPorId(env, p2.uniqueId)[25]).startsWith('ABERTO|ANA'));
    L.sincronizar(env);
    rel.check('sync seguinte não pergunta de novo o ABERTO', L.conferencia(L.dbPorId(env, p2.uniqueId)) === 'ABERTO');
    rel.check('tela passa a mostrar 3 pendências a menos', (ctx.fetchAllDataUnified(Date.now()).pendentesSaida || []).length === ps.length - 3);
  }

  rel.secao(`limpeza diária [${modo}]`);
  const retidos = L.db(env).slice(1).filter(L.faturadoSemUsuario).length;
  const ativosAntes = contar(env, 'Ativo');
  const dias = vm.runInContext('DIAS_RETENCAO', ctx);
  const limite = new Date(); limite.setDate(limite.getDate() - dias);
  const esperadoSimulacao = L.db(env).slice(1).filter(r => L.statusFinal(r[14]) && !L.faturadoSemUsuario(r)
    && r[16] instanceof Date && r[16] < limite).length;
  const esperadoAtivo = L.db(env).slice(1).filter(r => ['Faturado', 'Excluido'].includes(L.T(r[14])) && !L.faturadoSemUsuario(r)).length;
  const antes = L.db(env).length;
  const res = ctx.purgarItensFinalizados();
  vm.runInContext('_gravarAuditoria_()', ctx);
  const apagados = antes - L.db(env).length;
  if (modo === 'ATIVO') {
    rel.check(`apaga todo Faturado/Excluido com usuário (${esperadoAtivo})`, apagados === esperadoAtivo && res.purgados === esperadoAtivo, `apagou ${apagados}`);
    rel.check(`mantém os ${retidos} Faturado sem usuário`, contar(env, 'Faturado') === retidos && contar(env, 'Excluido') === 0);
    rel.check('nenhum Ativo apagado', contar(env, 'Ativo') === ativosAntes);
    rel.check('auditoria LIMPEZA', L.auditoria(env, 'LIMPEZA').length === 1);
  } else {
    rel.check(`regra antiga: só finalizados com ${dias}+ dias (${esperadoSimulacao})`, apagados === esperadoSimulacao, `apagou ${apagados}`);
    rel.check('auditoria LIMPEZA_SIMULADA', L.auditoria(env, 'LIMPEZA_SIMULADA').length === 1);
    rel.check('nenhum Ativo apagado', contar(env, 'Ativo') === ativosAntes);
  }
}
rel.fim();
