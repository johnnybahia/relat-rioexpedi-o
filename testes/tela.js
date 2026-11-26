#!/usr/bin/env node
// Tela (index.html) no Chromium, com google.script.run simulado e dados gerados pelo próprio Código.gs:
// aviso de conferência (abre sozinho para TOTAL, respostas, erros, "Responder depois", pendência nova,
// selo no item), exclusão de inativos com o login e usuário PARCIAL sem aviso. Sem Playwright/Chromium no ambiente → pulado (não falha).
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
  const payload = env.ctx.fetchAllDataUnified(Date.now());
  // Alerta de cartela: o primeiro item de cada OC passa a ter CARTELA = NÃO (a base anonimizada é quase toda OK)
  (payload.ordCompras || []).forEach(o => { if (o.items && o.items[0]) o.items[0].CARTELA = 'NÃO'; });
  return payload;
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
    excluirMultiplosItens: items => ({ success: true, processados: items.length, falhas: 0, results: [] }),
    // Servidor com estado: baixa, estorno e marcação mudam o item do payload (a próxima leitura já vê a mudança)
    aplicarBaixa: (id, linha, qtd) => {
      const it = window.__itemPayload(id);
      if (!it) return { success: true, id: id, linha: linha, novaQtd: 0, zerou: true };
      const atual = Number(it['QTD. ABERTA']) || 0;
      if (qtd > atual) return { success: false, error: 'Quantidade de baixa (' + qtd + ') maior que disponível (' + atual + ')' };
      it['QTD. ABERTA'] = atual - qtd; it['QTD. ORIGINAL'] = atual;
      return { success: true, id: id, linha: linha, novaQtd: atual - qtd, zerou: atual === qtd, qtdOriginal: atual };
    },
    estornarBaixa: (id, linha, linhaHist, qtd) => {
      const it = window.__itemPayload(id);
      if (!it) return { success: false, error: 'item não encontrado' };
      it['QTD. ABERTA'] = (Number(it['QTD. ABERTA']) || 0) + Number(qtd); it['QTD. ORIGINAL'] = it['QTD. ABERTA'];
      return { success: true, novaQtd: it['QTD. ABERTA'], qtdAnterior: it['QTD. ABERTA'] - Number(qtd), qtdOriginal: it['QTD. ORIGINAL'] };
    },
    marcarParaFaturar: (id, linha, marcar, usuario, perfil) => {
      const it = window.__itemPayload(id);
      if (!it) return { success: true, id: id, linha: linha, marcado: marcar, usuario: usuario };
      it.MARCAR_FATURAR = marcar ? 'SIM' : ''; it.PERFIL_EMISSAO = marcar ? perfil : '';
      it.MARCAR_FATURAR_USUARIO = marcar ? JSON.stringify({ [perfil]: usuario }) : '';
      return { success: true, id: id, linha: linha, marcado: marcar, usuario: marcar ? usuario : '',
        estado: { MARCAR_FATURAR: it.MARCAR_FATURAR, MARCAR_FATURAR_USUARIO: it.MARCAR_FATURAR_USUARIO, PERFIL_EMISSAO: it.PERFIL_EMISSAO } };
    },
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
      // A resposta é calculada na chamada (o servidor leu/gravou nessa hora); __segurar[nome] retém a entrega até __liberar(nome)
      return (...args) => {
        window.__calls.push([nome, args]); const f = window.__respostas[nome]; const r = f ? f(...args) : { success: true };
        const entregar = () => { if (succ) succ(r); };
        if (window.__segurar[nome]) window.__fila.push({ nome: nome, entregar: entregar }); else setTimeout(entregar, 20);
      };
    } });
  }
  window.__segurar = {}; window.__fila = [];
  window.__liberar = nome => { const i = window.__fila.findIndex(x => x.nome === nome); if (i < 0) return false; window.__fila.splice(i, 1)[0].entregar(); return true; };
  window.__itemPayload = id => { for (const o of window.__payload.ordCompras || []) for (const it of o.items || []) if (it.uniqueId === id) return it; return null; };
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
        // ---- alerta de CARTELA = NÃO ao marcar para faturar ----
        if (await aberto()) { await page.click('#conferencia-saida-depois-btn'); await page.waitForTimeout(200); }
        const cartelaAberto = () => page.$eval('#cartelaModal', el => !el.classList.contains('hidden'));
        const confirmarOff = () => page.$eval('#cartela-confirmar-btn', b => b.disabled);
        const marcarChamadas = () => chamadas('marcarParaFaturar');
        const alvoCartela = await page.evaluate(() => {
          const sem = v => String(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
          const livre = c => !c.checked && !c.disabled;
          const nao = [...document.querySelectorAll('tr[data-item-cartela] .checkbox-faturar')].find(c => livre(c) && sem(c.closest('tr').dataset.itemCartela) === 'NAO');
          const ok = [...document.querySelectorAll('tr[data-item-cartela] .checkbox-faturar')].find(c => livre(c) && sem(c.closest('tr').dataset.itemCartela) !== 'NAO');
          if (nao) nao.dataset.alvoNao = '1';
          if (ok) ok.dataset.alvoOk = '1';
          return nao ? { id: nao.closest('tr').dataset.uniqueId, ok: !!ok } : null;
        });
        if (!alvoCartela) {
          rel.pular('nenhum item com CARTELA NÃO e checkbox livre na primeira página de OCs');
        } else {
          const clicar = (attr) => page.evaluate(a => document.querySelector(`[data-${a}="1"]`).click(), attr);
          const n0 = (await marcarChamadas()).length;
          await clicar('alvo-nao');
          await page.waitForTimeout(200);
          rel.check('marcar item com CARTELA NÃO abre o alerta e ainda não chama o servidor', await cartelaAberto() && (await marcarChamadas()).length === n0);
          rel.check('alerta mostra descrição, cartela NÃO e o risco de devolução',
            /NÃO/.test(await page.$eval('#cartela-item-dados', el => el.textContent)) && /devolução/.test(await page.$eval('#cartelaModal', el => el.textContent)));
          await page.keyboard.press('Escape');
          await page.mouse.click(5, 5);
          rel.check('ESC e clique fora não fecham o alerta', await cartelaAberto());
          rel.check('Confirmar começa desabilitado', await confirmarOff());
          await page.fill('#cartela-supervisor', 'MA');
          await page.check('#cartela-ciente');
          rel.check('supervisor com 2 letras mantém Confirmar desabilitado', await confirmarOff());
          await page.fill('#cartela-supervisor', 'Maria  Souza');
          rel.check('supervisor + ciência habilita Confirmar', !(await confirmarOff()));
          await page.uncheck('#cartela-ciente');
          rel.check('sem a ciência desabilita de novo', await confirmarOff());
          if (capturas) { await page.check('#cartela-ciente'); await page.screenshot({ path: path.join(capturas, 'cartela.png') }); await page.uncheck('#cartela-ciente'); }
          await page.click('#cartela-cancelar-btn');
          await page.waitForTimeout(100);
          rel.check('Cancelar fecha, desmarca o checkbox e não chama o servidor',
            !(await cartelaAberto()) && !(await page.$eval('[data-alvo-nao="1"]', c => c.checked)) && (await marcarChamadas()).length === n0);
          await clicar('alvo-nao');
          await page.waitForTimeout(100);
          rel.check('o alerta abre de novo e começa limpo', await cartelaAberto() && (await page.$eval('#cartela-supervisor', i => i.value)) === '' && await confirmarOff());
          await page.fill('#cartela-supervisor', 'Maria  Souza');
          await page.check('#cartela-ciente');
          await page.click('#cartela-confirmar-btn');
          await page.waitForTimeout(600);
          const m = (await marcarChamadas()).pop();
          rel.check('Confirmar envia ao servidor a supervisão (espaços normalizados) e a ciência',
            m && m[1][0] === alvoCartela.id && m[1][2] === true && m[1][3] === usuario && m[1][5] && m[1][5].supervisor === 'Maria Souza' && m[1][5].ciente === true, JSON.stringify(m && m[1]));
          rel.check('depois de confirmar o alerta está fechado', !(await cartelaAberto()));
          if (alvoCartela.ok) {
            const nOk = (await marcarChamadas()).length;
            await clicar('alvo-ok');
            await page.waitForTimeout(500);
            const mOk = (await marcarChamadas()).slice(nOk).pop();
            rel.check('item com CARTELA OK marca direto, sem alerta e sem autorização', !(await cartelaAberto()) && mOk && !mOk[1][5], JSON.stringify(mOk && mOk[1]));
          } else {
            rel.pular('sem item com CARTELA OK livre na página');
          }
        }

        // ---- relatório de faturamento: faixa + linha em vermelho + supervisão ----
        const relHtml = await page.evaluate(() => {
          let html = '';
          const abrir = window.open;
          window.open = () => ({ document: { write: h => { html = h; }, close() {} }, print() {} });
          const base = { CLIENTE: 'CLIENTE 01', 'ORD. COMPRA': '16078984', PEDIDO: '143086', 'CÓD. CLIENTE': '1', TAMANHO: '16MM', SALDO: 130,
            'CÓD. OS': '1', PRAZO: 3, MARCAR_FATURAR_USUARIO: 'ANA', SEQUENCIA: '4', LOTE_EMISSAO: '', INFO_X: 'MARCA' };
          _gerarRelatorioFaturamentoPDF([
            Object.assign({}, base, { uniqueId: 'a', 'DESCRIÇÃO': 'ITEM COM CARTELA', SEM_CARTELA: false, LIBERADO_SEM_CARTELA: '' }),
            Object.assign({}, base, { uniqueId: 'b', 'DESCRIÇÃO': 'ITEM LIBERADO', SEM_CARTELA: true, LIBERADO_SEM_CARTELA: 'MARIA <b>X</b>|ANA|30/10/2026 10:38' }),
            Object.assign({}, base, { uniqueId: 'c', 'DESCRIÇÃO': 'ITEM LEGADO', SEM_CARTELA: true, LIBERADO_SEM_CARTELA: '' })
          ], 'ANA', [], 'FAT-001');
          window.open = abrir;
          return html;
        });
        rel.check('relatório: faixa avisa 2 itens sem cartela (1 liberado, 1 sem autorização registrada)',
          /2 ITEM\(NS\) SEM CARTELA — 1 liberado\(s\) pela supervisão, 1 SEM autorização registrada/.test(relHtml), relHtml.slice(0, 200));
        rel.check('relatório: linhas sem cartela em vermelho com selo; item com cartela segue verde',
          (relHtml.match(/class="sem-cartela"/g) || []).length === 2 && (relHtml.match(/class="highlight"/g) || []).length === 1
          && (relHtml.match(/>SEM CARTELA</g) || []).length === 2);
        rel.check('relatório: mostra a supervisão e a data, com HTML escapado',
          /supervisão: MARIA &lt;b&gt;X&lt;\/b&gt; · 30\/10\/2026 10:38/.test(relHtml) && !/<b>X<\/b>/.test(relHtml));
        rel.check('relatório: item legado mostra "autorização NÃO registrada"', /autorização da supervisão NÃO registrada/.test(relHtml));

        await page.evaluate(() => {
          document.getElementById('deleteAllModal').dataset.items = JSON.stringify([{ uniqueId: 'ID-INATIVO-TESTE', linha: 5, info: 'teste' }]);
          confirmExcluirNaoMarcados();
        });
        await page.waitForTimeout(300);
        const ex = (await chamadas('excluirMultiplosItens')).pop();
        rel.check('excluir inativos envia o login (histórico de decisões)', ex && ex[1][1] === usuario, JSON.stringify(ex && ex[1]));
        // A janela de sucesso da exclusão recarrega a página em 3 s: no meio dos testes seguintes zerava window.__calls
        // e o estado da tela (era a falha intermitente do teste do teclado)
        await page.evaluate(() => { clearInterval(_successCountdownTimer); document.getElementById('successModal').classList.add('hidden'); });

        // ---- marcar item com baixa parcial: TOTAL / PARCIAL / CANCELAR (o confirm antigo marcava no "Cancelar") ----
        await page.evaluate(() => {
          document.body.insertAdjacentHTML('beforeend', '<table><tbody><tr id="item-linha-9977" data-item-qtd="12" data-item-qtd-original="323" data-item-cartela="OK"><td><input type="checkbox" class="checkbox-faturar" checked></td></tr></tbody></table>');
          handleMarcarFaturar('ID-PARCIAL', 9977, true);
        });
        await page.waitForTimeout(100);
        const nMarcar = (await chamadas('marcarParaFaturar')).length;
        const janelaParcial = await page.$eval('#baixaConfirmacaoModal', el => !el.classList.contains('hidden'));
        const textoParcial = await page.$eval('#baixa-confirmacao-msg', el => el.textContent);
        await page.click('#baixa-confirmacao-cancelar-btn', { timeout: 2000 }).catch(() => {});
        await page.waitForTimeout(100);
        rel.check('marcar item parcial abre TOTAL / PARCIAL / CANCELAR com o baixado e o saldo',
          janelaParcial && /Baixado neste ciclo: 311/.test(textoParcial) && /Saldo em aberto: 12/.test(textoParcial), textoParcial);
        rel.check('CANCELAR não marca e desmarca a caixa',
          (await chamadas('marcarParaFaturar')).length === nMarcar && !(await page.$eval('#item-linha-9977 .checkbox-faturar', c => c.checked)));
        await page.evaluate(() => { document.querySelector('#item-linha-9977 .checkbox-faturar').checked = true; handleMarcarFaturar('ID-PARCIAL', 9977, true); });
        await page.waitForTimeout(100);
        await page.click('#baixa-confirmacao-parcial-btn', { timeout: 2000 }).catch(() => {});
        await page.waitForTimeout(300);
        const mParcial = (await chamadas('marcarParaFaturar')).pop();
        rel.check('PARCIAL marca só o baixado', mParcial && mParcial[1][0] === 'ID-PARCIAL' && mParcial[1][2] === true, JSON.stringify(mParcial && mParcial[1]));
        await page.evaluate(() => showBaixaConfirmacaoModal('x', () => {}, () => {}));
        rel.check('a janela depois da baixa continua só com TOTAL e PARCIAL',
          await page.evaluate(() => { const b = document.getElementById('baixa-confirmacao-cancelar-btn'); return !!b && getComputedStyle(b).display === 'none'; }));
        await page.evaluate(() => hideBaixaConfirmacaoModal());

        // ---- atualização automática (antes: canAutoRefresh dava erro — $confirmModal não existe — e o refresh morria) ----
        const refresh = await page.evaluate(() => {
          document.querySelectorAll('.modal-overlay').forEach(m => m.classList.add('hidden'));
          $loader.classList.add('hidden');
          if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
          const out = { intervalo: AUTO_REFRESH_INTERVAL };
          try { out.livre = canAutoRefresh(); } catch (e) { out.erro = e.message; }
          document.getElementById('successModal').classList.remove('hidden');
          try { out.comJanela = canAutoRefresh(); } catch (e) { out.erro = e.message; }
          document.getElementById('successModal').classList.add('hidden');
          const busca = document.createElement('input'); document.body.appendChild(busca); busca.focus();
          try { out.digitando = canAutoRefresh(); } catch (e) { out.erro = e.message; }
          busca.remove();
          const st = window.setTimeout, capturados = [];
          window.setTimeout = (fn, ms, ...a) => { if (ms === AUTO_REFRESH_INTERVAL) { capturados.push(fn); return 0; } return st(fn, ms, ...a); };
          try { scheduleAutoRefresh(); } finally { window.setTimeout = st; }
          const antes = window.__calls.filter(c => c[0] === 'fetchAllDataUnified').length;
          try { capturados.forEach(fn => fn()); } catch (e) { out.erro = e.message; }
          out.agendou = capturados.length;
          out.buscou = window.__calls.filter(c => c[0] === 'fetchAllDataUnified').length - antes;
          return out;
        });
        rel.check('atualização automática: sem erro; adia com janela aberta ou com alguém digitando',
          !refresh.erro && refresh.livre === true && refresh.comJanela === false && refresh.digitando === false, JSON.stringify(refresh));
        rel.check('atualização automática a cada 2 min dispara e busca os dados no servidor',
          refresh.intervalo === 120000 && refresh.agendou === 1 && refresh.buscou === 1, JSON.stringify(refresh));
        await page.waitForTimeout(300);
        const nMarcarTeclas = (await chamadas('marcarParaFaturar')).length;
        await page.keyboard.press('ArrowDown'); await page.keyboard.press(' '); await page.keyboard.press('Escape');
        await page.waitForTimeout(100);
        rel.check('teclado: atalhos de linha desligados (nenhuma tecla marca para faturar)', (await chamadas('marcarParaFaturar')).length === nMarcarTeclas, JSON.stringify((await chamadas('marcarParaFaturar')).slice(nMarcarTeclas)) + ' foco: ' + await page.evaluate(() => document.activeElement && (document.activeElement.id || document.activeElement.className || document.activeElement.tagName)));

        // ---- marcação × tela redesenhada no meio (incidente de 26/11/2026: nome do usuário com a caixa vazia e a Qtd. antiga) ----
        await page.evaluate(() => { document.querySelectorAll('.modal-overlay').forEach(m => m.classList.add('hidden')); $loader.classList.add('hidden'); });
        await page.waitForTimeout(2500); // recargas agendadas pelos testes anteriores terminam antes
        const linhaDe = id => page.evaluate(id => {
          const tr = [...document.querySelectorAll('tr[data-unique-id]')].find(t => t.dataset.uniqueId === id);
          if (!tr) return null;
          const cb = tr.querySelector('.checkbox-faturar'), b = tr.querySelector('.faturar-badge-usuario');
          return { marcado: !!(cb && cb.checked), travado: !!(cb && cb.disabled), nome: b ? b.textContent.trim() : '',
            qtd: Number(tr.querySelector('td.qtd-aberta').textContent.replace(/\D/g, '')), cor: !!tr.style.outline,
            qtdNoBotaoBaixa: (() => { const m = /"QTD\. ABERTA":(\d+)/.exec(tr.querySelector('.btn-baixa').getAttribute('onclick')); return m ? Number(m[1]) : null; })() };
        }, id);
        const naLinha = (id, sel) => page.evaluate(([id, sel]) => [...document.querySelectorAll('tr[data-unique-id]')].find(t => t.dataset.uniqueId === id).querySelector(sel).click(), [id, sel]);
        const segurar = nomes => page.evaluate(n => { window.__segurar = {}; n.forEach(x => { window.__segurar[x] = true; }); }, nomes);
        const liberar = nome => page.evaluate(n => window.__liberar(n), nome);
        const livres = await page.evaluate(() => [...document.querySelectorAll('tr[data-unique-id]')].filter(tr => {
          const c = tr.querySelector('.checkbox-faturar');
          return c && !c.checked && !c.disabled && tr.dataset.itemCartela === 'OK' && Number(tr.dataset.itemQtd) > 1
            && Number(tr.dataset.itemQtd) === Number(tr.dataset.itemQtdOriginal);
        }).slice(0, 3).map(tr => ({ id: tr.dataset.uniqueId, qtd: Number(tr.dataset.itemQtd) })));
        if (livres.length < 3) {
          rel.pular('menos de 3 itens livres (CARTELA OK, sem baixa) na tela');
        } else {
          const [a, d, f] = livres;
          await page.evaluate(() => { window.__confirmResp = true; });
          // A — a foto: a atualização lê antes da baixa, o usuário marca e a resposta da atualização chega no meio
          await segurar(['fetchAllDataUnified', 'aplicarBaixa', 'marcarParaFaturar']);
          await page.evaluate(() => carregarListaOcs(true, true));
          await naLinha(a.id, '.checkbox-faturar');
          await liberar('aplicarBaixa');
          await liberar('fetchAllDataUnified');
          await liberar('marcarParaFaturar');
          await segurar([]);
          const eA = await linhaDe(a.id);
          rel.check('atualização no meio da marcação: caixa marcada, com o nome e a Qtd. zerada (antes: nome com a caixa vazia e a Qtd. antiga)',
            eA && eA.marcado && eA.nome === usuario && eA.qtd === 0 && eA.qtdNoBotaoBaixa === 0 && !eA.travado, JSON.stringify(eA));
          await page.evaluate(() => aplicarFiltros());
          const eB = await linhaDe(a.id);
          rel.check('busca/filtro depois de marcar não volta a linha ao estado antigo', eB && eB.marcado && eB.nome === usuario && eB.qtd === 0, JSON.stringify(eB));
          await page.evaluate(() => carregarListaOcs(true, true));
          await page.waitForTimeout(300);
          const eA2 = await linhaDe(a.id);
          rel.check('a atualização seguinte (dados do servidor) mostra o mesmo', eA2 && eA2.marcado && eA2.nome === usuario && eA2.qtd === 0, JSON.stringify(eA2));

          // C — desmarcar com a atualização chegando no meio
          await segurar(['fetchAllDataUnified', 'marcarParaFaturar']);
          await page.evaluate(() => carregarListaOcs(true, true));
          await naLinha(a.id, '.checkbox-faturar');
          await liberar('fetchAllDataUnified');
          await liberar('marcarParaFaturar');
          await segurar([]);
          const eC = await linhaDe(a.id);
          rel.check('desmarcar com atualização no meio: caixa vazia, sem nome e sem a cor', eC && !eC.marcado && !eC.nome && !eC.cor && !eC.travado, JSON.stringify(eC));

          // D — enquanto grava, a caixa fica marcada e travada mesmo com a tela redesenhada (busca/filtro)
          await segurar(['aplicarBaixa', 'marcarParaFaturar']);
          await naLinha(d.id, '.checkbox-faturar');
          await page.evaluate(() => aplicarFiltros());
          const eD1 = await linhaDe(d.id);
          await liberar('aplicarBaixa');
          const eD2 = await linhaDe(d.id);
          await liberar('marcarParaFaturar');
          await segurar([]);
          const eD3 = await linhaDe(d.id);
          rel.check('enquanto grava: caixa marcada e travada, sem nome, mesmo com a tela redesenhada',
            eD1 && eD1.marcado && eD1.travado && !eD1.nome && eD2 && eD2.marcado && eD2.travado && eD2.qtd === 0, JSON.stringify([eD1, eD2]));
          rel.check('terminada a gravação: marcada, com o nome e destravada', eD3 && eD3.marcado && eD3.nome === usuario && !eD3.travado, JSON.stringify(eD3));

          // E — chamada que não volta não segura a atualização para sempre
          const teto = await page.evaluate(() => {
            try {
              document.querySelectorAll('.modal-overlay').forEach(m => m.classList.add('hidden'));
              $loader.classList.add('hidden');
              if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
              const r = [canAutoRefresh()];
              const t = _iniciarAcao_('ID-TETO');
              r.push(canAutoRefresh(), _haAcoesPendentes_(Date.now() + 61000));
              _terminarAcao_(t);
              r.push(canAutoRefresh());
              return r;
            } catch (e) { return ['erro: ' + e.message]; }
          });
          rel.check('atualização espera a ação em andamento; sem resposta em 60 s deixa de esperar',
            teto[0] === true && teto[1] === false && teto[2] === false && teto[3] === true, JSON.stringify(teto));

          // F — baixa pela janela 📦 Baixa com a atualização chegando no meio; G — estorno
          await segurar(['fetchAllDataUnified', 'aplicarBaixa']);
          await page.evaluate(() => carregarListaOcs(true, true));
          await naLinha(f.id, '.btn-baixa');
          await page.evaluate(() => { document.getElementById('nova-baixa-input').value = '1'; document.getElementById('baixa-confirm-btn').click(); });
          await liberar('aplicarBaixa');
          await page.evaluate(() => document.getElementById('baixa-confirmacao-parcial-btn').click());
          await liberar('fetchAllDataUnified');
          await segurar([]);
          const eF = await linhaDe(f.id);
          rel.check('baixa pela janela com atualização no meio: linha e botão Baixa com o saldo novo',
            eF && eF.qtd === f.qtd - 1 && eF.qtdNoBotaoBaixa === f.qtd - 1, JSON.stringify(eF));
          await page.evaluate(id => {
            const it = [].concat(...Object.values(ALL_OCS_ITEMS).map(o => o.items || [])).find(i => i.uniqueId === id);
            currentBaixaItem = JSON.parse(JSON.stringify(it));
            _executarEstorno(2, 1, it.planilhaLinha, null);
          }, f.id);
          await page.waitForTimeout(150);
          await page.evaluate(() => aplicarFiltros());
          const eG = await linhaDe(f.id);
          rel.check('estorno: a linha mostra a Qtd. devolvida (também depois da busca)', eG && eG.qtd === f.qtd && eG.qtdNoBotaoBaixa === f.qtd, JSON.stringify(eG));
          await page.waitForTimeout(2300); // recarga agendada pela baixa
          const quebradas = await page.evaluate(() => [...document.querySelectorAll('tr[data-unique-id]')]
            .filter(tr => { const cb = tr.querySelector('.checkbox-faturar'); return cb && !cb.checked && tr.querySelector('.faturar-badge-usuario'); }).length);
          rel.check('nenhuma linha com o nome do usuário e a caixa vazia', quebradas === 0, `${quebradas} linha(s)`);
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
