#!/usr/bin/env node
// Hook do Claude Code (PreToolUse/Bash, em .claude/settings.json): antes de um "git commit" com mudança em
// Código.gs, index.html, testes/, .claude/ ou .github/, roda testes/rodar_todos.js. Se algo falhar, bloqueia o
// commit (código 2) e devolve o resumo ao Claude. Qualquer outro comando passa direto.
const path = require('path');
const { execFileSync, spawnSync } = require('child_process');

let entrada = '';
process.stdin.on('data', d => { entrada += d; });
process.stdin.on('end', () => {
  let comando = '';
  try { comando = String((JSON.parse(entrada).tool_input || {}).command || ''); } catch (e) { process.exit(0); }
  if (!/\bgit\s+(?:-[Cc]\s+\S+\s+)*commit\b/.test(comando)) process.exit(0);

  const raiz = process.env.CLAUDE_PROJECT_DIR || path.join(__dirname, '..');
  let mudancas = '';
  try {
    mudancas = execFileSync('git', ['status', '--porcelain', '--', 'Código.gs', 'index.html', 'testes', '.claude', '.github'],
      { cwd: raiz, encoding: 'utf8' }).trim();
  } catch (e) { process.exit(0); }
  if (!mudancas) process.exit(0);

  const r = spawnSync(process.execPath, [path.join(raiz, 'testes', 'rodar_todos.js'), '--resumo'],
    { cwd: raiz, encoding: 'utf8', timeout: 540000, maxBuffer: 64 * 1024 * 1024 });
  if (r.status === 0) process.exit(0);
  process.stderr.write('Commit bloqueado: a bateria de simulação (node testes/rodar_todos.js) falhou. ' +
    'Corrija o código (ou o cenário) e rode de novo antes de commitar.\n\n' +
    ((r.stdout || '') + (r.stderr || '')).slice(-6000));
  process.exit(2);
});
