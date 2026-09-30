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
    rel.check('histórico: um REABERTO por item reaberto', L.historico(env, 'REABERTO').length === fsu.length - comGemea,
      `histórico ${L.historico(env, 'REABERTO').length}`);
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
  let idFatAviso = null;
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
    const hAviso = L.historico(env).filter(h => h[3] === 'aviso de conferência (tela)');
    const tem = (id, ev) => hAviso.some(h => L.T(h[4]) === id && h[1] === ev && h[2] === 'ANA');
    rel.check('histórico: só as 3 respostas aceitas, com o login (FATURADO, CANCELADO, ABERTO)',
      hAviso.length === 3 && tem(p0.uniqueId, 'FATURADO') && tem(p1.uniqueId, 'CANCELADO') && tem(p2.uniqueId, 'ABERTO'),
      JSON.stringify(hAviso.map(h => [h[1], h[2], h[4]])));
    rel.check('trava do documento liberada depois da resposta', vm.runInContext('_travaDocumentoObtida_', ctx) === false);
    idFatAviso = p0.uniqueId;
  }

  rel.secao(`histórico: marcação, exclusão e gravação sem trava [${modo}]`);
  let idFatMarcacao = null;
  const livres = L.linhasSozinhas(env);
  if (livres.length < 4) {
    rel.pular('menos de 4 linhas sem irmãs livres na base');
  } else {
    const [mFat, mCanc, mExc, mExc2] = livres;
    // Marcado + saiu da origem → Faturado + histórico; outro marcado é desmarcado DURANTE o sync → nada
    L.alterarDb(env, mFat.id, { 15: 'SIM', 21: L.USUARIO_TESTE });
    L.alterarDb(env, mCanc.id, { 15: 'SIM', 21: L.USUARIO_TESTE });
    [mFat, mCanc, mExc, mExc2].map(m => m.i).sort((x, y) => y - x).forEach(i => L.imp(env).splice(i, 1));
    const mesclarOriginal = ctx._mesclarAlteracoesConcorrentes_;
    ctx._mesclarAlteracoesConcorrentes_ = (sheet, dados, updates) => {
      L.alterarDb(env, mCanc.id, { 15: '', 21: '' }); // usuário desmarca no meio do sync
      return mesclarOriginal(sheet, dados, updates);
    };
    const hAntes = L.historico(env).length;
    try { L.sincronizar(env); } finally { ctx._mesclarAlteracoesConcorrentes_ = mesclarOriginal; }
    const novos = L.historico(env).slice(hAntes);
    rel.check('marcado + saiu da origem → Faturado e histórico com o nome da coluna V',
      L.T(L.dbPorId(env, mFat.id)[14]) === 'Faturado'
      && novos.some(h => h[1] === 'FATURADO' && L.T(h[4]) === mFat.id && h[2] === 'USUARIO TESTE (BAIXA1)'),
      JSON.stringify(novos.map(h => [h[1], h[2], h[4]])));
    rel.check('desmarcado durante o sync → não fatura e não entra no histórico',
      L.T(L.dbPorId(env, mCanc.id)[14]) !== 'Faturado' && !novos.some(h => L.T(h[4]) === mCanc.id)
      && L.auditoria(env, 'FATURAMENTO_CANCELADO').length > 0);
    idFatMarcacao = mFat.id;

    // Exclusão pela tela: com o login e, página antiga, sem usuário
    const linhaDb = id => L.db(env).findIndex(r => L.T(r[0]) === id) + 1;
    L.alterarDb(env, mExc.id, { 14: 'Inativo' });
    const rEx = ctx.excluirMultiplosItens([{ uniqueId: mExc.id, planilhaLinha: linhaDb(mExc.id) }], 'ANA');
    const hEx = L.historico(env, 'EXCLUIDO').filter(h => L.T(h[4]) === mExc.id);
    rel.check('excluir pela tela → Excluido e histórico com o login', rEx.success && L.T(L.dbPorId(env, mExc.id)[14]) === 'Excluido'
      && hEx.length === 1 && hEx[0][2] === 'ANA' && /Inativo/.test(hEx[0][14]), JSON.stringify(hEx.map(h => [h[2], h[14]])));
    ctx.excluirItem(mExc2.id, linhaDb(mExc2.id));
    rel.check('exclusão sem usuário (página antiga) fica como "não informado"',
      L.historico(env, 'EXCLUIDO').some(h => L.T(h[4]) === mExc2.id && h[2] === 'não informado'));

    // Sem a trava do documento: grava linha a linha com appendRow — nada se perde
    const shH = env.ss.getSheetByName('Historico_Decisoes');
    rel.check('aba do histórico protegida com aviso contra edição manual', !!(shH.protecao && shH.protecao.warningOnly));
    let appends = 0;
    shH.appendRow = v => { appends++; Object.getPrototypeOf(shH).appendRow.call(shH, v); };
    const lockOriginal = ctx.LockService.getDocumentLock;
    ctx.LockService.getDocumentLock = () => ({ tryLock: () => false, waitLock() { throw new Error('ocupada'); }, releaseLock() {}, hasLock: () => false });
    const n0 = L.historico(env).length;
    let gravou;
    try {
      gravou = vm.runInContext(`_registrarDecisao_('TESTE', 'X', 'teste', [], 'a'); _registrarDecisao_('TESTE', 'X', 'teste', [], 'b'); _gravarDecisoes_()`, ctx);
    } finally {
      ctx.LockService.getDocumentLock = lockOriginal;
      delete shH.appendRow;
    }
    rel.check('sem a trava do documento grava linha a linha (appendRow) sem perder nada',
      gravou === true && appends === 2 && L.historico(env).length === n0 + 2);
  }

  rel.secao(`limpeza diária [${modo}]`);
  if (modo === 'SIMULACAO') { // base recente: envelhece 5 finalizados para a regra antiga ter o que apagar
    const velho = new Date(); velho.setDate(velho.getDate() - 30);
    L.db(env).slice(1).filter(r => L.statusFinal(r[14]) && !L.faturadoSemUsuario(r)).slice(0, 5).forEach(r => { r[16] = velho; });
  }
  // Cópia para o histórico falhando → nada é apagado e o processo automático não marca o dia
  if (!env.ss.getSheetByName('Historico_Decisoes')) vm.runInContext(`_registrarDecisao_('TESTE', 'X', 'teste', [], 'cria a aba'); _gravarDecisoes_()`, ctx);
  const shHist = env.ss.getSheetByName('Historico_Decisoes');
  const quebrarHistorico = () => { shHist.getRange = () => { throw new Error('falha simulada'); }; shHist.appendRow = shHist.getRange; };
  const consertarHistorico = () => { delete shHist.getRange; delete shHist.appendRow; };
  const linhasAntesFalha = L.db(env).length;
  const diasRet = vm.runInContext('DIAS_RETENCAO', ctx);
  const limiteRet = new Date(); limiteRet.setDate(limiteRet.getDate() - diasRet);
  const aApagar = L.db(env).slice(1).filter(r => !L.faturadoSemUsuario(r) && (modo === 'ATIVO'
    ? ['Faturado', 'Excluido'].includes(L.T(r[14])) : (L.statusFinal(r[14]) && r[16] instanceof Date && r[16] < limiteRet))).length;
  if (aApagar > 0) {
    quebrarHistorico();
    let rAdiada;
    try { rAdiada = ctx.purgarItensFinalizados(); } finally { consertarHistorico(); }
    vm.runInContext('_gravarAuditoria_()', ctx);
    rel.check('cópia para o histórico falha → nada apagado (LIMPEZA_ADIADA)', rAdiada.adiada === true && rAdiada.purgados === 0
      && L.db(env).length === linhasAntesFalha && L.auditoria(env, 'LIMPEZA_ADIADA').length === 1
      && L.auditoria(env, 'HISTORICO_FALHOU').length >= 1, JSON.stringify(rAdiada));
    const deveOriginal = ctx._deveLimparFaturadosAgora_;
    ctx._deveLimparFaturadosAgora_ = () => true;
    env.props.delete('ULTIMA_LIMPEZA_FATURADOS_DATA');
    quebrarHistorico();
    try { ctx.processoAutomaticoCompleto(); } finally { consertarHistorico(); ctx._deveLimparFaturadosAgora_ = deveOriginal; }
    rel.check('processo automático com limpeza adiada não marca o dia (tenta de novo na próxima)',
      !env.props.get('ULTIMA_LIMPEZA_FATURADOS_DATA') && L.auditoria(env, 'LIMPEZA_ADIADA').length === 2);
  } else {
    rel.pular('limpeza: nada a apagar nesta base — teste de falha na cópia não rodou');
  }

  const retidos = L.db(env).slice(1).filter(L.faturadoSemUsuario).length;
  const ativosAntes = contar(env, 'Ativo');
  const dias = vm.runInContext('DIAS_RETENCAO', ctx);
  const limite = new Date(); limite.setDate(limite.getDate() - dias);
  const esperadoSimulacao = L.db(env).slice(1).filter(r => L.statusFinal(r[14]) && !L.faturadoSemUsuario(r)
    && r[16] instanceof Date && r[16] < limite).length;
  const esperadoAtivo = L.db(env).slice(1).filter(r => ['Faturado', 'Excluido'].includes(L.T(r[14])) && !L.faturadoSemUsuario(r)).length;
  const antes = L.db(env).length;
  const idsAntes = new Set(L.db(env).slice(1).map(r => L.T(r[0])));
  const histAntes = L.historico(env).length;
  const decisoesAntes = L.historico(env).filter(h => h[1] !== 'APAGADO_NA_LIMPEZA').length;
  const limpezasAntes = L.auditoria(env, modo === 'ATIVO' ? 'LIMPEZA' : 'LIMPEZA_SIMULADA').length;
  const res = ctx.purgarItensFinalizados();
  vm.runInContext('_gravarAuditoria_()', ctx);
  const apagados = antes - L.db(env).length;
  const copias = L.historico(env).slice(histAntes).filter(h => h[1] === 'APAGADO_NA_LIMPEZA');
  const idsDepois = new Set(L.db(env).slice(1).map(r => L.T(r[0])));
  rel.check(`cópia no histórico de cada linha apagada (${apagados})`, copias.length === apagados
    && copias.every(h => idsAntes.has(L.T(h[4])) && !idsDepois.has(L.T(h[4]))), `cópias ${copias.length}`);
  rel.check('decisões anteriores continuam no histórico depois da limpeza',
    L.historico(env).filter(h => h[1] !== 'APAGADO_NA_LIMPEZA').length === decisoesAntes);
  if (modo === 'ATIVO' && idFatAviso && idFatMarcacao) {
    const copia = id => copias.find(h => L.T(h[4]) === id);
    rel.check('a cópia guarda quem decidiu (aviso: login; marcação: coluna V)',
      copia(idFatAviso) && copia(idFatAviso)[2] === 'ANA' && copia(idFatMarcacao) && copia(idFatMarcacao)[2] === 'USUARIO TESTE (BAIXA1)',
      JSON.stringify([copia(idFatAviso), copia(idFatMarcacao)].map(h => h && h[2])));
  }
  if (modo === 'ATIVO') {
    rel.check(`apaga todo Faturado/Excluido com usuário (${esperadoAtivo})`, apagados === esperadoAtivo && res.purgados === esperadoAtivo, `apagou ${apagados}`);
    rel.check(`mantém os ${retidos} Faturado sem usuário`, contar(env, 'Faturado') === retidos && contar(env, 'Excluido') === 0);
    rel.check('nenhum Ativo apagado', contar(env, 'Ativo') === ativosAntes);
    rel.check('auditoria LIMPEZA', L.auditoria(env, 'LIMPEZA').length === limpezasAntes + 1);
  } else {
    rel.check(`regra antiga: só finalizados com ${dias}+ dias (${esperadoSimulacao})`, apagados === esperadoSimulacao, `apagou ${apagados}`);
    rel.check('auditoria LIMPEZA_SIMULADA', L.auditoria(env, 'LIMPEZA_SIMULADA').length === limpezasAntes + 1);
    rel.check('nenhum Ativo apagado', contar(env, 'Ativo') === ativosAntes);
  }
}
rel.fim();
