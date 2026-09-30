#!/usr/bin/env node
// Invariantes do sync sobre uma base inteira (qualquer export), nos dois modos:
//   • rodar o sync duas vezes com a mesma origem não muda nada (IDs, linhas, status, pendências);
//   • nenhum Faturado sem usuário é criado; nenhum ID/UUID repetido novo;
//   • na base padrão do repositório: a primeira rodada não troca nenhum ID de PEDIDOS.
// Sem --dados roda em todas as bases de testes/dados/; com --dados <pasta> roda só nela (exports novos).
const vm = require('vm');
const path = require('path');
const L = require('./lib');

const unica = process.argv.includes('--dados');
const pastas = unica ? [L.pastaDaLinhaDeComando()] : L.basesDoRepositorio();
const rel = L.relatorio('DADOS REAIS (invariantes do sync)');

const idsPedidos = env => L.ped(env).slice(3).filter(r => L.T(r[1])).map(r => L.T(r[0]));
const fotoDb = env => new Map(L.db(env).slice(1).filter(r => L.T(r[0])).map(r => [L.T(r[0]) + '#' + L.T(r[18]), `${L.T(r[14])}|${L.conferencia(r)}`]));
const sentinela = env => { const s = env.ctx._verificarIntegridadeDuplicatas_(); return { ids: s.idsRepetidos || 0, uuids: s.uuidsRepetidos || 0 }; };

for (const pasta of pastas) {
  const base = L.lerBase(pasta);
  const nome = path.relative(process.cwd(), pasta) || pasta;
  for (const modo of ['SIMULACAO', 'ATIVO']) {
    rel.secao(`${nome} [${modo}]`);
    const env = L.montar(base, modo);
    const idsOriginais = idsPedidos(env);
    const fsuAntes = L.db(env).slice(1).filter(L.faturadoSemUsuario).length;
    const dupAntes = sentinela(env);

    let t0 = Date.now();
    L.sincronizar(env);
    const ms1 = Date.now() - t0;
    const ids1 = idsPedidos(env);
    const trocados = ids1.filter((id, k) => idsOriginais[k] !== undefined && id !== idsOriginais[k]).length;
    const linhas1 = L.db(env).length;
    const foto1 = fotoDb(env);

    t0 = Date.now();
    L.sincronizar(env);
    const ms2 = Date.now() - t0;
    const ids2 = idsPedidos(env);
    const foto2 = fotoDb(env);
    const mudou = [...foto2.entries()].filter(([k, v]) => foto1.get(k) !== v).length + [...foto1.keys()].filter(k => !foto2.has(k)).length;
    const fsuDepois = L.db(env).slice(1).filter(L.faturadoSemUsuario).length;
    const dupDepois = sentinela(env);

    rel.info(`${ids1.length} linhas em PEDIDOS · ${linhas1 - 1} no DB · 1ª rodada ${ms1} ms, 2ª ${ms2} ms · IDs trocados na 1ª rodada: ${trocados}`);
    if (!unica && path.resolve(pasta) === path.resolve(L.BASE_PADRAO)) {
      rel.check('1ª rodada não troca nenhum ID de PEDIDOS (base padrão)', trocados === 0, `${trocados} trocado(s)`);
    }
    rel.check('2ª rodada com a mesma origem: mesmos IDs em PEDIDOS', ids2.length === ids1.length && ids2.every((id, k) => id === ids1[k]));
    rel.check('2ª rodada: mesmas linhas, status e pendências no DB', L.db(env).length === linhas1 && mudou === 0, `${mudou} linha(s) mudaram`);
    rel.check('nenhum Faturado sem usuário criado pelo sync', fsuDepois <= fsuAntes, `antes ${fsuAntes}, depois ${fsuDepois}`);
    rel.check('nenhum ID_UNICO/UUID repetido novo', dupDepois.ids <= dupAntes.ids && dupDepois.uuids <= dupAntes.uuids,
      `IDs ${dupAntes.ids}→${dupDepois.ids}, UUIDs ${dupAntes.uuids}→${dupDepois.uuids}`);
    const aud = L.auditoria(env);
    const porTipo = {};
    aud.forEach(a => { porTipo[a[2]] = (porTipo[a[2]] || 0) + 1; });
    if (Object.keys(porTipo).length) rel.info('auditoria: ' + JSON.stringify(porTipo));
    void vm;
  }
}
rel.fim();
