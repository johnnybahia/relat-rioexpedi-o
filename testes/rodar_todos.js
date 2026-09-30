#!/usr/bin/env node
// Roda a bateria inteira de simulação e devolve código 1 se algo falhar.
//   node testes/rodar_todos.js                  → bases do repositório (testes/dados/*)
//   node testes/rodar_todos.js --dados <pasta>  → também/só com exports novos (fora do Git)
//   --resumo  mostra só falhas e o placar (usado pelo hook de commit)
const { spawnSync } = require('child_process');
const path = require('path');

const args = process.argv.slice(2);
const resumo = args.includes('--resumo');
const repasse = args.filter(a => a !== '--resumo');
const ETAPAS = [
  ['dados_reais.js', 'invariantes do sync nas bases'],
  ['cenarios.js', '12 cenários de identidade e faturamento'],
  ['funcoes.js', 'reparo, aviso, limpeza e sentinela'],
  ['tela.js', 'tela no Chromium']
];

const t0 = Date.now();
const placar = [];
for (const [arq, descricao] of ETAPAS) {
  const ini = Date.now();
  const r = spawnSync(process.execPath, [path.join(__dirname, arq), ...repasse], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  const saida = (r.stdout || '') + (r.stderr || '');
  const linhaResultado = (saida.match(/^RESULTADO .*$/m) || [''])[0];
  const pulou = / ([1-9]\d*) pulado/.test(linhaResultado);
  const ok = r.status === 0;
  placar.push({ arq, descricao, ok, pulou, linhaResultado, ms: Date.now() - ini });
  if (resumo) {
    const falhas = saida.split('\n').filter(l => /❌|⏭️|Error|RESULTADO/.test(l));
    if (!ok || pulou) console.log(`\n── ${arq}\n${falhas.join('\n')}`);
  } else {
    process.stdout.write(saida);
  }
}
console.log('\n════════ PLACAR ════════');
placar.forEach(p => console.log(`${p.ok ? (p.pulou ? '⚠️ ' : '✅') : '❌'} ${p.arq.padEnd(15)} ${String(Math.round(p.ms / 1000)).padStart(3)} s  ${p.descricao}${p.linhaResultado ? ' — ' + p.linhaResultado.replace(/^RESULTADO [^:]*: /, '') : ''}`));
const falhou = placar.filter(p => !p.ok);
console.log(falhou.length === 0
  ? `\nTUDO VERDE em ${Math.round((Date.now() - t0) / 1000)} s`
  : `\n${falhou.length} ETAPA(S) FALHARAM: ${falhou.map(p => p.arq).join(', ')}`);
process.exitCode = falhou.length === 0 ? 0 : 1;
