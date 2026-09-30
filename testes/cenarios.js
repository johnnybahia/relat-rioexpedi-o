#!/usr/bin/env node
// Cenários de identidade e faturamento, nos dois modos (SIMULACAO e ATIVO). Cada cenário acha na base a
// situação de que precisa (grupo de linhas-irmãs, linha sem irmãs, faturado fora da origem…), altera a origem
// e/ou o DB como aconteceria no dia a dia, roda o sync e confere o resultado. Base sem a situação → pulado.
//   node testes/cenarios.js [--dados <pasta>]
// Todo bug corrigido deve ganhar um cenário aqui (CLAUDE.md, seção 18).
const L = require('./lib');

const base = L.lerBase(L.pastaDaLinhaDeComando());
const rel = L.relatorio('CENÁRIOS');
const contarFsu = env => L.db(env).slice(1).filter(L.faturadoSemUsuario).length;
const contarZ = (env, tipo) => L.db(env).slice(1).filter(r => L.conferencia(r) === tipo).length;
const dbAbertosPorLote = (env, lote) => L.db(env).slice(1).filter(r => L.T(r[20]) === L.T(lote) && !L.statusFinal(r[14]));

function cenario(titulo, achar, preparar, conferir) {
  for (const modo of ['SIMULACAO', 'ATIVO']) {
    rel.secao(`${titulo} [${modo}]`);
    const env = L.montar(base, modo);
    L.sincronizar(env);                       // estabiliza a base; o cenário mede só o efeito da mudança
    const aud = env.ss.getSheetByName('Auditoria_Sincronizacao');
    if (aud) aud.data = aud.data.slice(0, 1);
    env.props.delete('ULTIMA_DIF_IDENTIDADE');
    const alvo = achar(env);
    if (!alvo) { rel.pular('a base não tem a situação necessária'); continue; }
    const fsuAntes = contarFsu(env);
    const extra = preparar(env, alvo) || {};
    const fatAntes = new Set(L.db(env).slice(1).filter(r => L.T(r[14]) === 'Faturado').map(r => L.T(r[0])));
    const histAntes = L.historico(env).length;
    L.sincronizar(env);
    conferir(env, modo, alvo, extra);
    rel.check('nenhum Faturado sem usuário novo', contarFsu(env) <= fsuAntes, `antes ${fsuAntes}, depois ${contarFsu(env)}`);
    // Historico_Decisoes: todo Faturado novo registrado com o usuário; nada registrado que não aconteceu
    const novosFat = L.db(env).slice(1).filter(r => L.T(r[14]) === 'Faturado' && !fatAntes.has(L.T(r[0]))).map(r => L.T(r[0]));
    const histFat = L.historico(env).slice(histAntes).filter(h => h[1] === 'FATURADO');
    const comUsuario = new Set(histFat.filter(h => L.T(h[2]) && h[2] !== 'não informado').map(h => L.T(h[4])));
    rel.check('histórico: todo Faturado novo registrado com o usuário', novosFat.every(id => comUsuario.has(id)),
      `sem registro: ${novosFat.filter(id => !comUsuario.has(id)).join(', ')}`);
    rel.check('histórico: nenhum faturamento registrado que não aconteceu',
      histFat.every(h => { const r = L.dbPorId(env, L.T(h[4])); return r && L.T(r[14]) === 'Faturado'; }));
  }
}
const aud = (env, tipo) => L.auditoria(env, tipo);
const detalhesAud = (env, tipo) => aud(env, tipo).map(a => a[4]).join(' || ');
function prepararIrmas(env, grupo, especial) {
  grupo.forEach(m => L.alterarDb(env, m.id, { 14: 'Ativo', 9: m.qtd, 15: '', 21: '', 25: '' }));
  if (especial) L.alterarDb(env, especial.id, especial.campos);
}
// Grupo com duas irmãs de mesma QTD (a antes de b na origem)
function acharQtdEmpatada(env) {
  for (const g of L.gruposDeIrmas(env)) {
    for (let x = 0; x < g.length; x++) for (let y = x + 1; y < g.length; y++) {
      if (g[x].qtd === g[y].qtd && g[x].qtd > 0) return { grupo: g, a: g[x], b: g[y] };
    }
  }
  return null;
}

// 1) Incidente de 28/09: a irmã marcada (e zerada pela baixa) sai da origem; outra irmã tem a mesma QTD.
cenario('1. irmã marcada sai da origem (QTD empatada com outra irmã)', acharQtdEmpatada, (env, { grupo, a }) => {
  prepararIrmas(env, grupo, { id: a.id, campos: { 9: 0, 15: 'SIM', 21: L.USUARIO_TESTE } });
  L.registrarBaixa(env, a.id, a.qtd, 0, a.qtd);
  L.imp(env).splice(a.i, 1);
}, (env, modo, { grupo, a }) => {
  const outras = grupo.filter(m => m !== a);
  if (modo === 'ATIVO') {
    const ra = L.dbPorId(env, a.id);
    rel.check('irmã marcada → Faturado com usuário', ra[14] === 'Faturado' && L.conferencia(ra) === 'FATURADO', L.resumoLinha(ra));
    outras.forEach(m => {
      const r = L.dbPorId(env, m.id);
      rel.check(`lote ${m.lote} fica com o seu ID, Ativo, sem pendência`,
        L.idPedidosPorLote(env, m.lote) === m.id && r[14] === 'Ativo' && !L.T(r[25]) && Number(r[9]) === m.qtd, L.resumoLinha(r));
    });
  } else {
    rel.check('simulação registra que o casamento antigo troca a irmã', aud(env, 'IDENTIDADE_SIMULADA').length > 0);
  }
});

// 2) DESCRIÇÃO corrigida na origem (linha sem irmãs).
cenario('2. DESCRIÇÃO corrigida na origem', env => L.linhasSozinhas(env).find(m => L.baseDoIdConfere(env, m)) || null, (env, m) => {
  L.imp(env)[m.i][7] = String(L.imp(env)[m.i][7]) + ' (CORRIGIDA)';
}, (env, modo, m) => {
  const r = L.dbPorId(env, m.id);
  if (modo === 'ATIVO') {
    rel.check('mesmo ID, DESCRIÇÃO atualizada, sem pendência', L.idPedidosPorLote(env, m.lote) === m.id && r && r[14] === 'Ativo'
      && String(r[6]).includes('(CORRIGIDA)') && !L.T(r[25]), L.resumoLinha(r));
    rel.check('nenhuma linha nova no DB', dbAbertosPorLote(env, m.lote).length === 1);
    rel.check('auditoria ID_POR_BASE', aud(env, 'ID_POR_BASE').length >= 1);
  } else {
    rel.check('simulação registra que o ATIVO manteria o ID', aud(env, 'IDENTIDADE_SIMULADA').length > 0);
    rel.check('linha antiga não vira Faturado', r && r[14] !== 'Faturado');
  }
});

// 3) Lote novo com a mesma impressão digital de um item já Faturado (e fora da origem).
function acharFaturadoForaDaOrigem(env) {
  const ctx = env.ctx;
  const fpsOrigem = new Set(L.imp(env).slice(3).filter(r => L.T(r[1])).map(r => ctx._criarImpressaoDigitalFromRow_(r.slice(1), 0)));
  const idsPed = new Set(L.ped(env).slice(3).map(r => L.T(r[0])));
  for (const d of L.db(env).slice(1)) {
    if (L.T(d[14]) !== 'Faturado' || !L.T(d[21]) || idsPed.has(L.T(d[0]))) continue;
    const fp = ctx._criarImpressaoDigital_(d, true);
    if (fpsOrigem.has(fp)) continue;
    const vizinha = L.imp(env).find((r, i) => i >= 3 && L.T(r[1]) && L.T(r[2]) === L.T(d[2]) && L.T(r[4]) === L.T(d[3]) && L.T(r[9]) === L.T(d[8]));
    if (!vizinha) continue;
    const novo = vizinha.slice();
    novo[5] = d[4]; novo[6] = d[5]; novo[7] = ctx._descBase_(d[6]); novo[8] = d[7]; novo[11] = d[10]; novo[12] = d[11];
    novo[10] = 500; novo[24] = '999001';
    if (ctx._criarImpressaoDigitalFromRow_(novo.slice(1), 0) !== fp) continue;
    return { fatId: L.T(d[0]), novo };
  }
  return null;
}
cenario('3. lote novo com a mesma impressão digital de um item Faturado', acharFaturadoForaDaOrigem, (env, { novo }) => {
  L.imp(env).push(novo);
}, (env, modo, { fatId }) => {
  const idNovo = L.idPedidosPorLote(env, '999001');
  const linhaNova = dbAbertosPorLote(env, '999001');
  if (modo === 'ATIVO') {
    rel.check('lote novo ganha ID próprio (não herda o faturado)', idNovo && idNovo !== fatId, `ID ${idNovo}`);
    rel.check('lote novo aparece no DB como Ativo, QTD 500', linhaNova.length === 1 && linhaNova[0][14] === 'Ativo' && Number(linhaNova[0][9]) === 500);
  } else {
    const registrou = aud(env).some(a => a[2] === 'NOVO_BARRADO_SIMULADO' || a[2] === 'FINALIZADO_CAPTURA_SIMULADO' || /herda o ID/.test(a[4]));
    rel.check('simulação registra o risco (ou o lote aparece)', registrou || linhaNova.length === 1);
  }
});

// 4) Linha sem marcação sai da origem → PENDENTE, nunca Faturado.
cenario('4. linha sem marcação sai da origem', env => L.linhasSozinhas(env)[0] || null, (env, m) => {
  L.imp(env).splice(m.i, 1);
}, (env, modo, m) => {
  const r = L.dbPorId(env, m.id);
  rel.check('continua Ativo com a coluna Z = PENDENTE', r && r[14] === 'Ativo' && L.conferencia(r) === 'PENDENTE', L.resumoLinha(r));
});

// 5) Linha marcada pelo usuário sai da origem → Faturado com o usuário.
cenario('5. linha marcada sai da origem', env => L.linhasSozinhas(env)[1] || null, (env, m) => {
  L.alterarDb(env, m.id, { 15: 'SIM', 21: L.USUARIO_TESTE });
  L.imp(env).splice(m.i, 1);
}, (env, modo, m) => {
  const r = L.dbPorId(env, m.id);
  rel.check('Faturado com Z = FATURADO|usuário', r && r[14] === 'Faturado' && L.conferencia(r) === 'FATURADO' && L.T(r[25]).includes('USUARIO TESTE'), L.resumoLinha(r));
});

// 6) Importação truncada (60% das linhas somem) → nada é faturado nem sinalizado.
cenario('6. importação truncada', env => ({ pend: contarZ(env, 'PENDENTE'), fat: contarZ(env, 'FATURADO') }), env => {
  const d = L.imp(env);
  const manter = d.slice(0, 3).concat(d.slice(3, 3 + Math.floor((d.length - 3) * 0.4)));
  d.length = 0; manter.forEach(x => d.push(x));
}, (env, modo, antes) => {
  rel.check('nenhuma pendência nem faturamento novo', contarZ(env, 'PENDENTE') === antes.pend && contarZ(env, 'FATURADO') === antes.fat,
    `PENDENTE ${antes.pend}→${contarZ(env, 'PENDENTE')}, FATURADO ${antes.fat}→${contarZ(env, 'FATURADO')}`);
});

// 7) Faturamento parcial entre irmãs: a irmã A teve baixa e marcação; a origem reduz a QTD dela para o saldo.
function acharParcial(env) {
  for (const g of L.gruposDeIrmas(env)) {
    for (let x = 0; x < g.length; x++) for (let y = x + 1; y < g.length; y++) {
      const A = g[x], B = g[y];
      if (A.qtd > B.qtd && B.qtd >= 2 && new Set(g.map(m => m.qtd)).size === g.length) return { grupo: g, A, B, novaQtd: B.qtd - 1 };
    }
  }
  return null;
}
cenario('7. faturamento parcial entre irmãs', acharParcial, (env, { grupo, A, novaQtd }) => {
  prepararIrmas(env, grupo, { id: A.id, campos: { 9: novaQtd, 15: 'SIM', 21: L.USUARIO_TESTE } });
  L.registrarBaixa(env, A.id, A.qtd - novaQtd, novaQtd, A.qtd);
  L.imp(env)[A.i][10] = novaQtd;
}, (env, modo, { A, B, novaQtd }) => {
  if (modo === 'ATIVO') {
    const ra = L.dbPorId(env, A.id), rb = L.dbPorId(env, B.id);
    rel.check('irmã A fica com o seu ID (QTD do saldo, novo ciclo: marcação limpa)',
      L.idPedidosPorLote(env, A.lote) === A.id && Number(ra[9]) === novaQtd && !L.T(ra[15]), L.resumoLinha(ra));
    rel.check('irmã B fica com o seu ID e a sua QTD', L.idPedidosPorLote(env, B.lote) === B.id && Number(rb[9]) === B.qtd, L.resumoLinha(rb));
  } else {
    rel.check('simulação registra a troca (ou o casamento antigo acertou)',
      L.idPedidosPorLote(env, A.lote) === A.id || aud(env, 'IDENTIDADE_SIMULADA').length > 0);
  }
});

// 8) Linha sinalizada volta à origem → a pendência some sozinha, mesmo ID.
cenario('8. linha sinalizada volta à origem', env => L.linhasSozinhas(env)[2] || null, (env, m) => {
  const linha = L.imp(env)[m.i];
  L.imp(env).splice(m.i, 1);
  L.sincronizar(env);
  const zSaiu = L.conferencia(L.dbPorId(env, m.id));
  L.imp(env).splice(m.i, 0, linha);
  return { zSaiu };
}, (env, modo, m, { zSaiu }) => {
  const r = L.dbPorId(env, m.id);
  rel.check('pendência criada ao sair e apagada ao voltar, mesmo ID', zSaiu === 'PENDENTE' && !L.T(r[25]) && r[14] === 'Ativo'
    && L.idPedidosPorLote(env, m.lote) === m.id, `ao sair: ${zSaiu} · ao voltar: ${L.resumoLinha(r)}`);
});

// 9) DESCRIÇÃO corrigida em todas as irmãs de um grupo.
cenario('9. DESCRIÇÃO corrigida no grupo inteiro de irmãs', env => L.gruposDeIrmas(env).find(g => g.every(m => L.baseDoIdConfere(env, m))) || null, (env, g) => {
  g.forEach(m => { L.imp(env)[m.i][7] = String(L.imp(env)[m.i][7]) + ' NOVA'; });
}, (env, modo, g) => {
  if (modo === 'ATIVO') {
    rel.check('todas as irmãs mantêm o ID e nenhuma vira pendência',
      g.every(m => L.idPedidosPorLote(env, m.lote) === m.id && !L.T(L.dbPorId(env, m.id)[25])));
  } else {
    rel.check('simulação registra uma diferença por irmã', aud(env, 'IDENTIDADE_SIMULADA').length >= g.length, detalhesAud(env, 'IDENTIDADE_SIMULADA'));
  }
});

// 10) A origem reordena duas irmãs de QTD empatada.
cenario('10. origem reordena irmãs de QTD empatada', acharQtdEmpatada, (env, { a, b }) => {
  const d = L.imp(env);
  [d[a.i], d[b.i]] = [d[b.i], d[a.i]];
}, (env, modo, { a, b }) => {
  const mantidos = L.idPedidosPorLote(env, a.lote) === a.id && L.idPedidosPorLote(env, b.lote) === b.id;
  if (modo === 'ATIVO') rel.check('cada lote mantém o seu ID', mantidos);
  else rel.check('simulação registra a troca (ou o casamento antigo acertou)', mantidos || aud(env, 'IDENTIDADE_SIMULADA').length > 0);
});

// 11) Lote novo entra no topo de um grupo de irmãs, com QTD quase igual à da primeira irmã.
cenario('11. lote novo entra no grupo de irmãs', env => L.gruposDeIrmas(env).find(g => g[0].qtd >= 2) || null, (env, g) => {
  const nova = L.imp(env)[g[0].i].slice();
  nova[10] = g[0].qtd - 1; nova[24] = '999002';
  L.imp(env).splice(g[0].i, 0, nova);
}, (env, modo, g) => {
  const linhaNova = dbAbertosPorLote(env, '999002');
  if (modo === 'ATIVO') {
    rel.check('irmãs existentes mantêm o ID', g.every(m => L.idPedidosPorLote(env, m.lote) === m.id));
    rel.check('lote novo aparece como Ativo com a sua QTD', linhaNova.length === 1 && Number(linhaNova[0][9]) === g[0].qtd - 1);
  } else {
    rel.check('simulação registra a troca (ou o lote aparece)', linhaNova.length === 1 || aud(env, 'IDENTIDADE_SIMULADA').length > 0);
  }
});

// 12) Dois lotes do mesmo grupo mudam de ORD. COMPRA na origem.
cenario('12. dois lotes mudam de OC', env => L.gruposDeIrmas(env)[0] || null, (env, g) => {
  g.slice(0, 2).forEach(m => { L.imp(env)[m.i][9] = 'OCNOVA1'; });
}, (env, modo, g) => {
  const movidos = g.slice(0, 2);
  if (modo === 'ATIVO') {
    const ok = movidos.every(m => {
      const linhas = dbAbertosPorLote(env, m.lote);
      return linhas.length === 1 && L.T(linhas[0][8]) === 'OCNOVA1' && linhas[0][14] === 'Ativo' && Number(linhas[0][9]) === m.qtd;
    });
    rel.check('as 2 linhas do DB seguem para a OC nova pelo LOTE, sem duplicar', ok,
      movidos.map(m => dbAbertosPorLote(env, m.lote).map(L.resumoLinha).join(' / ')).join(' || '));
  } else {
    rel.check('simulação registra o que faria', aud(env, 'TROCA_OC_SIMULADA').length >= 1);
  }
});

rel.fim();
