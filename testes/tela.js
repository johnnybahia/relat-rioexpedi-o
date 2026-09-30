#!/usr/bin/env node
// Tela (index.html) no Chromium, com google.script.run simulado e dados gerados pelo próprio Código.gs:
// aviso de conferência (abre sozinho para TOTAL, respostas, erros, "Responder depois", pendência nova,
// selo no item) e usuário PARCIAL sem aviso. Sem Playwright/Chromium no ambiente → pulado (não falha).
//   node testes/tela.js [--dados <pasta>] [--capturas]   (capturas vão para testes/saida/)
const fs = require('fs');
const os = require('os');
const path = require('path');
const L = require('./lib');

const rel = L.relatorio('TELA (index.html no Chromium)');

function carregarPlaywright() {
  const tentativas = ['playwright'];
  try { tentativas.push(path.join(require('child_process').execSync('npm root -g', { encoding: 'utf8' }).trim(), 'playwright')); } catch (e) { /* sem npm */ }
  for (const t of tentativas) { try { return require(t); } catch (e) { /* tenta a próxima */ } }
  return null;
}
async function abrirNavegador(chromium) {
  const caminhos = [undefined, process.env.CHROMIUM_PATH, '/opt/pw-browsers/chromium'].filter((c, i) => i === 0 || c);
  for (const executablePath of caminhos) {
    try { return await chromium.launch(executablePath ? { executablePath } : {}); } catch (e) { /* tenta o próximo */ }
  }
  return null;
}

function gerarPayload() {
  const env = L.montar(L.lerBase(L.pastaDaLinhaDeComando()), 'SIMULACAO');
  L.sincronizar(env);
  if (L.db(env).slice(1).some(L.faturadoSemUsuario)) env.ctx.repararFaturadosSemUsuario();
  L.linhasSozinhas(env).slice(0, 3).map(m => m.i).sort((x, y) => y - x).forEach(i => L.imp(env).splice(i, 1));
  L.sincronizar(env);
  return env.ctx.fetchAllDataUnified(Date.now());
}

(async () => {
  const pw = carregarPlaywright();
  if (!pw) { rel.pular('Playwright não instalado (npm i playwright) — teste da tela não rodou'); rel.fim(); return; }
  const browser = await abrirNavegador(pw.chromium);
  if (!browser) { rel.pular('Chromium não encontrado (npx playwright install chromium)'); rel.fim(); return; }

  const payload = gerarPayload();
  const N = (payload.pendentesSaida || []).length;
  if (N < 3) { rel.pular(`só ${N} pendência(s) na base — preciso de 3`); await browser.close(); rel.fim(); return; }
  if (!payload.pendentesSaida.some(p => p.gemeo)) payload.pendentesSaida[1].gemeo = { uniqueId: 'X', oc: '999', qtdAberta: 12 };
  const idxGemeo = payload.pendentesSaida.map(p => !!p.gemeo);

  const stub = `<script>
  window.__payload = ${JSON.stringify(payload)};
  window.__calls = []; window.__confirmResp = true;
  window.confirm = m => { window.__calls.push(['confirm', [m]]); return window.__confirmResp; };
  window.__respostas = {
    autenticarLogin: u => ({ success: true, nivel: u === 'BIA' ? 'PARCIAL' : 'TOTAL', tempoSessao: 15 }),
    obterNivelUsuario: u => ({ success: true, nivel: u === 'BIA' ? 'PARCIAL' : 'TOTAL', tempoSessao: 15 }),
    fetchAllDataUnified: () => JSON.parse(JSON.stringify(window.__payload)),
    confirmarSaidaFonte: (id, linha, dec) => {
      if (id === window.__falharId) return { success: false, error: 'Este item não está mais pendente (já foi respondido ou voltou à origem). Atualize a tela.' };
      window.__payload.pendentesSaida = window.__payload.pendentesSaida.filter(p => p.uniqueId !== id);
      window.__payload.stats.pendentesSaida = window.__payload.pendentesSaida.length;
      return { success: true, id: id, decisao: dec, linha: linha };
    }
  };
  function __runner(succ, fail) {
    return new Proxy({}, { get(_, nome) {
      if (nome === 'withSuccessHandler') return fn => __runner(fn, fail);
      if (nome === 'withFailureHandler') return fn => __runner(succ, fn);
      if (nome === 'withUserObject') return () => __runner(succ, fail);
      return (...args) => { window.__calls.push([nome, args]); const f = window.__respostas[nome]; const r = f ? f(...args) : { success: true }; setTimeout(() => { if (succ) succ(r); }, 20); };
    } });
  }
  window.google = { script: { run: __runner(null, null) } };
  </script>`;
  const html = fs.readFileSync(process.env.INDEX_HTML || path.join(__dirname, '..', 'index.html'), 'utf8').replace('<head>', '<head>' + stub);
  const arq = path.join(os.tmpdir(), `tela_teste_${process.pid}.html`);
  fs.writeFileSync(arq, html);
  const capturas = process.argv.includes('--capturas') ? path.join(__dirname, 'saida') : null;
  if (capturas) fs.mkdirSync(capturas, { recursive: true });

  try {
    for (const [usuario, nivel] of [['ANA', 'TOTAL'], ['BIA', 'PARCIAL']]) {
      rel.secao(`login ${usuario} (${nivel}) · ${N} pendências`);
      const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
      const erros = [];
      page.on('pageerror', e => erros.push(e.message));
      await page.goto('file://' + arq);
      await page.fill('#loginUsuario', usuario);
      await page.fill('#loginSenha', 'x');
      await page.click('#loginBtn');
      await page.waitForTimeout(1500);
      const aberto = () => page.$eval('#conferenciaSaidaModal', el => !el.classList.contains('hidden') && getComputedStyle(el).display !== 'none');
      const pill = await page.$eval('#pendentes-saida-pill', el => ({ txt: el.textContent, vis: getComputedStyle(el).display !== 'none' }));
      const chamadas = nome => page.evaluate(n => window.__calls.filter(c => c[0] === n), nome);
      if (nivel === 'TOTAL') {
        rel.check('aviso abre sozinho após o login', await aberto());
        rel.check(`contador mostra ${N}`, pill.vis && pill.txt.includes(`${N} `), pill.txt);
        rel.check(`${N} cartões no aviso`, (await page.$$eval('.conf-card', els => els.length)) === N);
        const comDup = await page.$$eval('.conf-card', els => els.map(e => !!e.querySelector('[data-decisao="DUPLICATA"]')));
        rel.check('botão Duplicata só nos cartões com gêmea', comDup.every((v, k) => v === idxGemeo[k]));
        if (capturas) await page.screenshot({ path: path.join(capturas, 'aviso.png') });

        await page.click('.conf-card:nth-child(1) [data-decisao="FATURADO"]');
        await page.waitForTimeout(300);
        const c = (await chamadas('confirmarSaidaFonte')).pop();
        rel.check('Faturado chama confirmarSaidaFonte(id, linha, FATURADO, login)', c && c[1][2] === 'FATURADO' && c[1][3] === usuario && Number(c[1][1]) > 0, JSON.stringify(c && c[1]));
        rel.check('cartão respondido some', (await page.$$eval('.conf-card', els => els.length)) === N - 1);

        const nConf = (await chamadas('confirm')).length;
        await page.click('.conf-card:nth-child(1) [data-decisao="ABERTO"]');
        await page.waitForTimeout(300);
        rel.check('"Continua aberto" não pede confirmação', (await chamadas('confirm')).length === nConf);

        await page.evaluate(() => { window.__confirmResp = false; });
        const antes = (await chamadas('confirmarSaidaFonte')).length;
        await page.click('.conf-card:nth-child(1) [data-decisao="CANCELADO"]');
        await page.waitForTimeout(300);
        rel.check('desistir na confirmação não envia nada', (await chamadas('confirmarSaidaFonte')).length === antes);
        await page.evaluate(() => { window.__confirmResp = true; });

        const idFalha = await page.$eval('.conf-card:nth-child(1)', el => el.getAttribute('data-id'));
        await page.evaluate(id => { window.__falharId = id; }, idFalha);
        await page.click('.conf-card:nth-child(1) [data-decisao="CANCELADO"]');
        await page.waitForTimeout(300);
        const erro = await page.$eval('.conf-card:nth-child(1) .conf-erro', el => el.style.display !== 'none' ? el.textContent : '');
        const travado = await page.$eval('.conf-card:nth-child(1) [data-decisao="FATURADO"]', el => el.disabled);
        rel.check('erro do servidor aparece no cartão e trava os botões', /não está mais pendente/.test(erro) && travado, erro);

        await page.click('#conferencia-saida-depois-btn');
        await page.waitForTimeout(2600);
        await page.evaluate(() => carregarListaOcs(true, true));
        await page.waitForTimeout(600);
        const pill2 = await page.$eval('#pendentes-saida-pill', el => el.textContent);
        rel.check('"Responder depois" fecha e a recarga não reabre', !(await aberto()));
        rel.check(`contador atualizado (${N - 2})`, pill2.includes(`${N - 2} `), pill2);

        await page.evaluate(() => {
          const n = JSON.parse(JSON.stringify(window.__payload.pendentesSaida[0]));
          n.uniqueId = 'NOVO-ID-TESTE'; window.__payload.pendentesSaida.push(n);
        });
        await page.evaluate(() => carregarListaOcs(true, true));
        await page.waitForTimeout(600);
        rel.check('pendência nova numa recarga reabre o aviso', await aberto());
        await page.click('#conferencia-saida-depois-btn');
        const selos = await page.$$eval('.badge-conferir', els => els.length);
        if (selos > 0) {
          await page.evaluate(() => document.querySelector('.badge-conferir').click());
          await page.waitForTimeout(200);
          rel.check('selo do item abre o aviso', await aberto());
        } else {
          rel.pular('nenhum selo visível na primeira página de OCs');
        }
        if (capturas) {
          await page.setViewportSize({ width: 390, height: 800 });
          await page.waitForTimeout(200);
          await page.screenshot({ path: path.join(capturas, 'aviso_celular.png') });
        }
      } else {
        rel.check('usuário PARCIAL não vê o aviso nem o contador', !(await aberto()) && !pill.vis);
      }
      rel.check('sem erro de JavaScript na página', erros.length === 0, erros.join(' | '));
      await page.close();
    }
  } finally {
    await browser.close();
    fs.unlinkSync(arq);
  }
  rel.fim();
})().catch(e => { console.error(e); process.exitCode = 1; });
