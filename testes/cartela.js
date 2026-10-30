#!/usr/bin/env node
// Item com CARTELA = NÃO marcado para faturar: o servidor só aceita com a supervisão identificada e a
// ciência do risco, grava o registro permanente (Historico_Decisoes) e a coluna AA (LIBERADO_SEM_CARTELA),
// que o relatório de faturamento usa. Desfeito se o registro falhar; apagado ao desmarcar e no reset de ciclo.
//   node testes/cartela.js [--dados <pasta>]
const L = require('./lib');

const base = L.lerBase(L.pastaDaLinhaDeComando());
const rel = L.relatorio('CARTELA = NÃO (liberação pela supervisão)');
const AA = 26;
const linhaDb = (env, id) => L.db(env).findIndex(r => L.T(r[0]) === id) + 1;
const livres = env => L.db(env).slice(1).filter(r => L.T(r[14]) === 'Ativo' && !L.T(r[15]) && !L.T(r[21]) && Number(r[9]) > 0);
const marcar = (env, r, aut, usuario = 'ANA', perfil = 'BAIXA1') =>
  env.ctx.marcarParaFaturar(L.T(r[0]), linhaDb(env, L.T(r[0])), true, usuario, perfil, aut);
const desmarcar = (env, r, usuario = 'ANA', perfil = 'BAIXA1') =>
  env.ctx.marcarParaFaturar(L.T(r[0]), linhaDb(env, L.T(r[0])), false, usuario, perfil);
const intacta = r => !L.T(r[15]) && !L.T(r[21]) && !L.T(r[AA]);
const sc = env => L.historico(env, 'FATURAR_SEM_CARTELA');

for (const modo of ['SIMULACAO', 'ATIVO']) {
  rel.secao(`marcação com CARTELA = NÃO [${modo}]`);
  const env = L.montar(base, modo);
  L.sincronizar(env);
  const fsuAntes = L.db(env).slice(1).filter(L.faturadoSemUsuario).length;
  const abertos = livres(env);
  if (abertos.length < 5) { rel.pular('menos de 5 itens livres na base'); continue; }
  const [nao, ok, falha, reset, semRegistro] = abertos;
  const idNao = L.T(nao[0]);
  L.alterarDb(env, idNao, { 1: 'NÃO' });
  L.alterarDb(env, L.T(ok[0]), { 1: 'OK' });

  const sem = marcar(env, nao, null);
  rel.check('sem autorização → recusado e linha intacta', sem.success === false && sem.exigeSupervisor === true && intacta(nao) && sc(env).length === 0, JSON.stringify(sem));
  const invalidas = [{ supervisor: 'MARIA', ciente: false }, { supervisor: 'MARIA' }, { supervisor: 'MA', ciente: true },
    { supervisor: '   ', ciente: true }, { supervisor: 'MARIA', ciente: 'true' }, { ciente: true }, 'MARIA'];
  rel.check('supervisor curto/vazio ou sem ciência → recusado', invalidas.every(a => marcar(env, nao, a).success === false) && intacta(nao) && sc(env).length === 0);

  for (const valor of ['Não', 'NAO', ' não ', 'nao']) {
    L.alterarDb(env, idNao, { 1: valor });
    if (marcar(env, nao, null).success !== false) rel.check(`valor "${valor}" também exige autorização`, false);
  }
  rel.check('variações de "NÃO" (acento, caixa, espaços) exigem autorização', intacta(nao));
  L.alterarDb(env, idNao, { 1: 'NÃO' });

  const r1 = marcar(env, nao, { supervisor: '  Maria   Souza|x ', ciente: true });
  const aa = JSON.parse(L.T(nao[AA]) || '{}');
  const [supAA, usrAA, dataAA] = L.T(aa.BAIXA1).split('|');
  rel.check('com supervisão + ciência → marca (P=SIM, V com o usuário)', r1.success === true && L.T(nao[15]) === 'SIM' && /ANA/.test(L.T(nao[21])), JSON.stringify(r1));
  rel.check('coluna AA guarda supervisor, usuário e data por perfil', supAA === 'Maria Souza/x' && usrAA === 'ANA' && /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}$/.test(dataAA), L.T(nao[AA]));
  rel.check('cabeçalho LIBERADO_SEM_CARTELA criado na coluna AA', L.T(L.db(env)[0][AA]) === 'LIBERADO_SEM_CARTELA');
  const h = sc(env);
  rel.check('histórico permanente: FATURAR_SEM_CARTELA com login, item e supervisão',
    h.length === 1 && h[0][2] === 'ANA' && L.T(h[0][4]) === idNao && /Maria Souza\/x/.test(h[0][14]) && /cartela NÃO/.test(h[0][14]) && /devolução/.test(h[0][14]), JSON.stringify(h.map(x => [x[2], x[14]])));

  const itens = env.ctx.obterItensMarcadosParaFaturar('BAIXA1').items;
  const itNao = itens.find(i => i.uniqueId === idNao);
  rel.check('relatório recebe SEM_CARTELA e a liberação do perfil', itNao && itNao.SEM_CARTELA === true && /^Maria Souza\/x\|ANA\|/.test(itNao.LIBERADO_SEM_CARTELA), JSON.stringify(itNao && [itNao.SEM_CARTELA, itNao.LIBERADO_SEM_CARTELA]));
  rel.check('o outro perfil não enxerga a liberação do BAIXA1', (env.ctx.obterItensMarcadosParaFaturar('BAIXA2').items || []).every(i => i.uniqueId !== idNao));

  const d = desmarcar(env, nao);
  rel.check('desmarcar não pede supervisão e limpa P, V e AA', d.success === true && intacta(nao), JSON.stringify(d));
  rel.check('desmarcar não grava novo registro', sc(env).length === 1);
  rel.check('marcar de novo exige nova autorização (e gera novo registro)',
    marcar(env, nao, null).success === false && marcar(env, nao, { supervisor: 'JOSE', ciente: true }).success === true && sc(env).length === 2);
  desmarcar(env, nao);

  // Dois perfis marcando o mesmo item: cada um com a sua liberação
  marcar(env, nao, { supervisor: 'MARIA', ciente: true }, 'ANA', 'BAIXA1');
  marcar(env, nao, { supervisor: 'JOSE', ciente: true }, 'ANA', 'BAIXA2');
  const doisPerfis = JSON.parse(L.T(nao[AA]));
  rel.check('dois perfis: AA guarda a liberação de cada um', /^MARIA\|/.test(doisPerfis.BAIXA1) && /^JOSE\|/.test(doisPerfis.BAIXA2), L.T(nao[AA]));
  desmarcar(env, nao, 'ANA', 'BAIXA1');
  rel.check('desmarcar um perfil mantém a liberação do outro', !JSON.parse(L.T(nao[AA])).BAIXA1 && /^JOSE\|/.test(JSON.parse(L.T(nao[AA])).BAIXA2));
  desmarcar(env, nao, 'ANA', 'BAIXA2');
  rel.check('desmarcar o último perfil apaga a coluna AA', !L.T(nao[AA]));

  const rOk = marcar(env, ok, null);
  rel.check('CARTELA = OK marca normal, sem AA', rOk.success === true && L.T(ok[15]) === 'SIM' && !L.T(ok[AA]), JSON.stringify(rOk));
  rel.check('CARTELA = OK não gera FATURAR_SEM_CARTELA', sc(env).every(x => L.T(x[4]) !== L.T(ok[0])));
  const itOk = env.ctx.obterItensMarcadosParaFaturar('BAIXA1').items.find(i => i.uniqueId === L.T(ok[0]));
  rel.check('relatório: item com cartela OK não é marcado como SEM_CARTELA', itOk && itOk.SEM_CARTELA === false && itOk.LIBERADO_SEM_CARTELA === '');

  // Item marcado antes desta regra (sem AA) e com cartela NÃO: aparece no relatório sem autorização registrada
  L.alterarDb(env, L.T(falha[0]), { 1: 'NÃO', 15: 'SIM', 21: '{"BAIXA1":"ANA"}', 22: '', 23: 'BAIXA1' });
  const itLegado = env.ctx.obterItensMarcadosParaFaturar('BAIXA1').items.find(i => i.uniqueId === L.T(falha[0]));
  rel.check('marcado antes da regra: SEM_CARTELA sem liberação registrada', itLegado && itLegado.SEM_CARTELA === true && itLegado.LIBERADO_SEM_CARTELA === '');
  L.alterarDb(env, L.T(falha[0]), { 1: 'OK', 15: '', 21: '', 23: '' });

  // Registro permanente falha → a marcação é desfeita (nunca marca sem rastro)
  const shH = env.ss.getSheetByName('Historico_Decisoes');
  L.alterarDb(env, L.T(semRegistro[0]), { 1: 'NÃO' });
  const quebrar = () => { shH.getRange = () => { throw new Error('falha simulada'); }; shH.appendRow = shH.getRange; };
  const consertar = () => { delete shH.getRange; delete shH.appendRow; };
  quebrar();
  let rFalha;
  try { rFalha = marcar(env, semRegistro, { supervisor: 'MARIA', ciente: true }); } finally { consertar(); }
  rel.check('histórico indisponível → recusa e deixa a linha como estava', rFalha.success === false && intacta(semRegistro) && /NÃO foi marcado/.test(rFalha.error), JSON.stringify(rFalha));
  rel.check('trava do documento liberada depois da marcação', require('vm').runInContext('_travaDocumentoObtida_', env.ctx) === false);

  // Reset de ciclo no sync (QTD da origem mudou depois de baixa) apaga a liberação junto com a marcação
  const resetId = L.T(reset[0]);
  const qtd = Number(reset[9]);
  const iOrigem = L.imp(env).findIndex((x, i) => i >= 3 && L.T(x[24]) && L.T(x[24]) === L.T(reset[20]));
  if (qtd < 2 || iOrigem < 0 || L.idPedidosPorLote(env, L.T(reset[20])) !== resetId) {
    rel.pular('reset de ciclo: item sem LOTE único ou com QTD < 2');
  } else {
    L.imp(env)[iOrigem][1] = 'NÃO';
    L.alterarDb(env, resetId, { 1: 'NÃO' });
    marcar(env, reset, { supervisor: 'MARIA', ciente: true });
    L.registrarBaixa(env, resetId, 1, qtd - 1, qtd);
    L.alterarDb(env, resetId, { 9: qtd - 1 });
    L.imp(env)[iOrigem][10] = qtd - 1;
    L.sincronizar(env);
    const r = L.dbPorId(env, resetId);
    rel.check('reset de ciclo no sync limpa marcação, usuário e liberação (AA)', r && !L.T(r[15]) && !L.T(r[21]) && !L.T(r[AA]), L.resumoLinha(r) + ' AA=' + L.T(r && r[AA]));
  }
  rel.check('nenhum Faturado sem usuário novo', L.db(env).slice(1).filter(L.faturadoSemUsuario).length === fsuAntes);
}
rel.fim();
