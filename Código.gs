
// ==================================================== 
// SISTEMA DE RELATÓRIO DE PEDIDOS - v15.5 OTIMIZADA
// COM CACHE - CARREGAMENTO RÁPIDO
// ====================================================

// ====== CONFIGURAÇÃO ======
const SS = SpreadsheetApp.openById("1qPJ8c7cq7qb86VJJ-iByeiaPnALOBcDPrPMeL75N2EI");
const FONTE_SHEET_NAME = "PEDIDOS";
const DB_SHEET_NAME = "Relatorio_DB";
const FONTE_DATA_START_ROW = 4;
const TZ = 'America/Fortaleza';
const APP_VERSION = '15.5-OTIMIZADA';

// CACHE (10 minutos)
const CACHE_DURATION = 600; // 10 minutos em segundos

// Índices de colunas - ABA PEDIDOS (para sincronização)
const ID_COL = 0;        // A
const CARTELA_COL = 1;   // B
const CLIENTE_COL = 2;   // C
const PEDIDO_COL = 4;    // E
const CODCLI_COL = 5;    // F
const MARFIM_COL = 6;    // G
const DESC_COL = 7;      // H
const TAM_COL = 8;       // I
const OC_COL = 9;        // J
const QTD_COL = 10;      // K
const OS_COL = 11;       // L
const DTREC_COL = 12;    // M
const DTENT_COL = 13;    // N
const PRAZO_COL = 14;    // O (na aba PEDIDOS)

// Índices de colunas - ABA Relatorio_DB
// Status é sempre a coluna O (índice 14 no array, coluna 15 na planilha)
const STATUS_COL = 14;   // O (coluna 15 ao contar a partir de 1)
const MARCAR_FATURAR_COL = 15; // P (coluna 16 ao contar a partir de 1) - Nova coluna para marcar itens para faturamento

// ====== BAIXAS PARCIAIS ======
const BAIXAS_SHEET_NAME = "Baixas_Historico";

// ====== FUNÇÃO WEB APP ======
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Relatório de Pedidos v15.5')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ====== FUNÇÃO DE TESTE ======
function testarHistorico() {
  Logger.clear();
  Logger.log("=== TESTE DE HISTÓRICO ===\n");

  const sheet = _getBaixasSheet_();
  const lastRow = sheet.getLastRow();

  Logger.log(`Total de linhas na planilha: ${lastRow}`);

  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, Math.min(10, lastRow - 1), sheet.getLastColumn()).getValues();
    Logger.log("\nPrimeiros registros:");
    data.forEach((row, idx) => {
      Logger.log(`${idx + 2}: ID="${row[0]}" | Qtd=${row[2]} | Data=${row[1]}`);
    });
  }
}

// ====== FUNÇÕES AUXILIARES ======
function _asDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function _fmtBR_(d) {
  if (!d) return '';
  const dt = _asDate_(d);
  return dt ? Utilities.formatDate(dt, TZ, 'dd/MM/yyyy') : '';
}

function _fmtBRDateTime_(d) {
  if (!d) return '';
  const dt = _asDate_(d);
  return dt ? Utilities.formatDate(dt, TZ, 'dd/MM/yyyy HH:mm') : '';
}

function _toNumber_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/[^\d,.-]/g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// ====== FUNÇÕES DE BAIXAS PARCIAIS ======

function _getBaixasSheet_() {
  let sheet = SS.getSheetByName(BAIXAS_SHEET_NAME);
  if (!sheet) {
    Logger.log(`📝 Criando aba ${BAIXAS_SHEET_NAME}...`);
    sheet = SS.insertSheet(BAIXAS_SHEET_NAME);
    // Criar cabeçalho
    sheet.getRange(1, 1, 1, 6).setValues([[
      'ID_ITEM', 'DATA_HORA', 'QTD_BAIXADA', 'QTD_RESTANTE', 'QTD_ORIGINAL', 'USUARIO'
    ]]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f0f2f5');
    sheet.setFrozenRows(1);
    Logger.log(`✅ Aba ${BAIXAS_SHEET_NAME} criada com sucesso`);
  } else {
    // Verifica se a coluna QTD_ORIGINAL existe
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('QTD_ORIGINAL')) {
      Logger.log(`📝 Adicionando coluna QTD_ORIGINAL...`);
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue('QTD_ORIGINAL').setFontWeight('bold').setBackground('#f0f2f5');
    }
  }
  return sheet;
}

function registrarBaixa(uniqueId, qtdBaixada, qtdRestante) {
  try {
    Logger.log(`📦 Registrando baixa para ID: "${uniqueId}"`);
    const sheet = _getBaixasSheet_();
    const now = new Date();
    const usuario = Session.getActiveUser().getEmail() || 'Sistema';

    const numCols = sheet.getLastColumn();

    // LÊ O CABEÇALHO para saber a ordem das colunas
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    Logger.log(`   Cabeçalho: ${headers.join(', ')}`);

    // Mapeia índices
    const colMap = {};
    headers.forEach((h, i) => {
      colMap[String(h).trim()] = i;
    });

    // Verifica se já existe histórico para calcular QTD_ORIGINAL
    const lastRow = sheet.getLastRow();
    let qtdOriginal = qtdRestante + qtdBaixada; // Padrão: primeira baixa

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
      const primeiraEntrada = data.find(row => String(row[0]).trim() === String(uniqueId).trim());

      if (primeiraEntrada && colMap['QTD_ORIGINAL'] !== undefined) {
        const qtdOrigPlanilha = primeiraEntrada[colMap['QTD_ORIGINAL']];
        if (qtdOrigPlanilha !== undefined && qtdOrigPlanilha !== '') {
          qtdOriginal = _toNumber_(qtdOrigPlanilha);
          Logger.log(`   ✓ Histórico existente, QTD_ORIGINAL: ${qtdOriginal}`);
        }
      }
    }

    // Cria array na ORDEM DO CABEÇALHO
    const novaLinha = new Array(numCols).fill('');
    novaLinha[colMap['ID_ITEM']] = uniqueId;
    novaLinha[colMap['DATA_HORA']] = now;
    novaLinha[colMap['QTD_BAIXADA']] = qtdBaixada;
    novaLinha[colMap['QTD_RESTANTE']] = qtdRestante;
    novaLinha[colMap['QTD_ORIGINAL']] = qtdOriginal;
    novaLinha[colMap['USUARIO']] = usuario;

    Logger.log(`   Salvando: [${novaLinha.join(', ')}]`);

    sheet.appendRow(novaLinha);
    SpreadsheetApp.flush();
    Logger.log(`✅ Baixa registrada na linha ${sheet.getLastRow()}`);

    _qtdOriginalCache_ = null;

    return { success: true, timestamp: now.toISOString() };
  } catch (e) {
    Logger.log(`❌ Erro ao registrar baixa: ${e.message}`);
    Logger.log(`   Stack: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

function obterHistoricoBaixas(uniqueId) {
  // VERSÃO ULTRA-DEFENSIVA - Lê cabeçalho dinamicamente
  Logger.log(`📋 [INICIO] obterHistoricoBaixas("${uniqueId}")`);

  try {
    if (!uniqueId) {
      Logger.log('⚠️ ID vazio, retornando array vazio');
      return { success: true, historico: [] };
    }

    const sheet = _getBaixasSheet_();
    if (!sheet) {
      Logger.log('❌ Aba não encontrada, retornando array vazio');
      return { success: true, historico: [] };
    }

    const lastRow = sheet.getLastRow();
    Logger.log(`   Última linha: ${lastRow}`);

    if (lastRow < 2) {
      Logger.log('⚠️ Sem dados, retornando array vazio');
      return { success: true, historico: [] };
    }

    const numCols = sheet.getLastColumn();

    // LÊ O CABEÇALHO para mapear as colunas corretamente
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    Logger.log(`   Cabeçalho: ${headers.join(', ')}`);

    // Mapeia índices das colunas
    const colMap = {};
    headers.forEach((h, i) => {
      colMap[String(h).trim()] = i;
    });

    Logger.log(`   ID_ITEM=${colMap['ID_ITEM']}, USUARIO=${colMap['USUARIO']}, QTD_ORIGINAL=${colMap['QTD_ORIGINAL']}`);

    // Lê os dados
    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    Logger.log(`   Leu ${data.length} linhas`);

    const idBusca = String(uniqueId).trim();
    const historico = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const idPlanilha = String(row[colMap['ID_ITEM']] || '').trim();

      if (idPlanilha === idBusca) {
        Logger.log(`   ✓ Match linha ${i + 2}`);

        // Lê cada coluna pelo nome (não pela posição)
        const dataHora = row[colMap['DATA_HORA']];
        const dataFormatada = dataHora ? _fmtBRDateTime_(dataHora) : '';

        const qtdBaixada = _toNumber_(row[colMap['QTD_BAIXADA']]);
        const qtdRestante = _toNumber_(row[colMap['QTD_RESTANTE']]);

        // QTD_ORIGINAL e USUARIO podem estar em qualquer ordem
        const qtdOriginal = colMap['QTD_ORIGINAL'] !== undefined ?
          _toNumber_(row[colMap['QTD_ORIGINAL']]) : 0;

        const usuario = colMap['USUARIO'] !== undefined ?
          String(row[colMap['USUARIO']] || 'Sistema') : 'Sistema';

        Logger.log(`      -> Qtd: ${qtdBaixada}, Usuario: "${usuario}", Original: ${qtdOriginal}`);

        historico.push({
          idItem: String(row[colMap['ID_ITEM']] || ''),
          dataHora: dataFormatada,
          dataHoraFormatada: dataFormatada,
          qtdBaixada: Number(qtdBaixada),
          qtdRestante: Number(qtdRestante),
          qtdOriginal: Number(qtdOriginal),
          usuario: usuario
        });
      }
    }

    Logger.log(`📋 Encontrados: ${historico.length} registros`);

    const resultado = {
      success: true,
      historico: historico
    };

    // Testa serialização
    try {
      const teste = JSON.stringify(resultado);
      Logger.log(`✅ Serialização OK (${teste.length} chars)`);
    } catch (jsonErr) {
      Logger.log(`❌ ERRO na serialização: ${jsonErr.message}`);
      return { success: true, historico: [] };
    }

    Logger.log('📤 [FIM] Retornando resultado');
    return resultado;

  } catch (e) {
    Logger.log(`❌ ERRO FATAL: ${e.message}`);
    Logger.log(`   Stack: ${e.stack}`);
    return { success: false, error: String(e.message), historico: [] };
  }
}

function editarUltimaBaixa(uniqueId, planilhaLinha, novaQtdBaixada) {
  try {
    const sheet = _getBaixasSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error('Nenhum histórico encontrado');
    }

    const numCols = sheet.getLastColumn();

    // Lê cabeçalho
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => {
      colMap[String(h).trim()] = i;
    });

    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    let ultimaLinha = -1;

    // Encontra a última baixa deste item
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][colMap['ID_ITEM']]).trim() === String(uniqueId).trim()) {
        ultimaLinha = i + 2;
        break;
      }
    }

    if (ultimaLinha === -1) {
      throw new Error('Nenhuma baixa encontrada para este item');
    }

    const linhaAtual = sheet.getRange(ultimaLinha, 1, 1, numCols).getValues()[0];
    const qtdRestanteAnterior = _toNumber_(linhaAtual[colMap['QTD_RESTANTE']]);
    const qtdBaixadaAnterior = _toNumber_(linhaAtual[colMap['QTD_BAIXADA']]);

    // Calcula nova quantidade restante
    const diferenca = novaQtdBaixada - qtdBaixadaAnterior;
    const novaQtdRestante = qtdRestanteAnterior - diferenca;

    Logger.log(`✏️ Editando baixa: ${qtdBaixadaAnterior} → ${novaQtdBaixada}, Restante: ${novaQtdRestante}`);

    if (novaQtdRestante < 0) {
      throw new Error('Quantidade restante não pode ser negativa');
    }

    // Atualiza o histórico usando índices do cabeçalho
    sheet.getRange(ultimaLinha, colMap['QTD_BAIXADA'] + 1).setValue(novaQtdBaixada);
    sheet.getRange(ultimaLinha, colMap['QTD_RESTANTE'] + 1).setValue(novaQtdRestante);
    sheet.getRange(ultimaLinha, colMap['DATA_HORA'] + 1).setValue(new Date());

    // Atualiza a QTD. ABERTA na planilha Relatorio_DB
    const dbSheet = SS.getSheetByName(DB_SHEET_NAME);
    if (dbSheet && planilhaLinha) {
      const dbHeaders = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0];
      const dbColMap = _getColumnIndexes_(dbHeaders);
      const qtdCol = dbColMap['QTD. ABERTA'];

      if (qtdCol !== undefined) {
        dbSheet.getRange(planilhaLinha, qtdCol + 1).setValue(novaQtdRestante);
        Logger.log(`✅ QTD. ABERTA atualizada: ${novaQtdRestante}`);
      }
    }

    SpreadsheetApp.flush();
    _qtdOriginalCache_ = null;
    limparCache();

    Logger.log(`✅ Edição concluída: ${uniqueId} | Qtd: ${novaQtdBaixada} | Restante: ${novaQtdRestante}`);

    return {
      success: true,
      novaQtdRestante: novaQtdRestante,
      qtdBaixada: novaQtdBaixada
    };
  } catch (e) {
    Logger.log(`❌ Erro ao editar última baixa: ${e.message}`);
    return { success: false, error: e.message };
  }
}

function aplicarBaixa(uniqueId, planilhaLinha, qtdBaixa) {
  try {
    const sheet = SS.getSheetByName(DB_SHEET_NAME);
    const linhaNum = Number(planilhaLinha);

    if (!sheet) throw new Error("Aba DB não encontrada");
    if (!isFinite(linhaNum) || linhaNum < 2 || linhaNum > sheet.getLastRow()) {
      throw new Error(`Linha inválida: ${planilhaLinha}`);
    }

    // Lê cabeçalhos para encontrar colunas corretas
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = _getColumnIndexes_(headers);

    const qtdCol = colMap['QTD. ABERTA'];
    const statusCol = colMap['Status'];

    if (qtdCol === undefined) {
      throw new Error("Coluna 'QTD. ABERTA' não encontrada");
    }

    // Lê quantidade atual
    const qtdAtual = sheet.getRange(linhaNum, qtdCol + 1).getValue();
    const qtdAtualNum = _toNumber_(qtdAtual);

    Logger.log(`📊 Aplicando baixa - Linha: ${linhaNum}, Qtd Atual: ${qtdAtualNum}, Baixa: ${qtdBaixa}`);

    // Valida
    if (qtdBaixa > qtdAtualNum) {
      throw new Error(`Quantidade de baixa (${qtdBaixa}) maior que disponível (${qtdAtualNum})`);
    }

    // Calcula nova quantidade
    const novaQtd = qtdAtualNum - qtdBaixa;

    // Atualiza na planilha
    sheet.getRange(linhaNum, qtdCol + 1).setValue(novaQtd);

    // Registra no histórico
    const resultHistorico = registrarBaixa(uniqueId, qtdBaixa, novaQtd);

    // Se zerou, marca como Faturado
    if (novaQtd === 0 && statusCol !== undefined) {
      sheet.getRange(linhaNum, statusCol + 1).setValue("Faturado");
      Logger.log(`✅ Item ${uniqueId} zerado e marcado como Faturado`);
    }

    SpreadsheetApp.flush();
    limparCache();
    Logger.log(`✅ Baixa aplicada: ${uniqueId} | -${qtdBaixa} | Nova Qtd: ${novaQtd}`);

    return {
      success: true,
      id: uniqueId,
      linha: linhaNum,
      novaQtd: novaQtd,
      zerou: novaQtd === 0
    };
  } catch (e) {
    Logger.log(`❌ aplicarBaixa: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

// Cache para quantidades originais (evita leituras múltiplas)
let _qtdOriginalCache_ = null;

function _buildQtdOriginalCache_() {
  try {
    const sheet = _getBaixasSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {}; // Sem histórico
    }

    const numCols = sheet.getLastColumn();

    // Lê cabeçalho
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => {
      colMap[String(h).trim()] = i;
    });

    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    const cache = {};

    // Para cada item, pega a QTD_ORIGINAL da primeira entrada
    data.forEach(row => {
      const id = row[colMap['ID_ITEM']];
      const qtdOriginal = row[colMap['QTD_ORIGINAL']];

      if (!cache[id] && qtdOriginal !== undefined && qtdOriginal !== '') {
        cache[id] = _toNumber_(qtdOriginal);
      }
    });

    Logger.log(`📦 Cache de quantidades construído: ${Object.keys(cache).length} itens`);
    return cache;
  } catch (e) {
    Logger.log(`⚠️ Erro ao construir cache: ${e.message}`);
    return {};
  }
}

function calcularQtdOriginal(uniqueId, qtdAbertaAtual) {
  try {
    // Usa cache se disponível
    if (!_qtdOriginalCache_) {
      _qtdOriginalCache_ = _buildQtdOriginalCache_();
    }

    // Se existe no histórico, usa o valor armazenado
    if (_qtdOriginalCache_[uniqueId]) {
      return _qtdOriginalCache_[uniqueId];
    }

    // Se não tem histórico, a quantidade atual É a original
    return qtdAbertaAtual;
  } catch (e) {
    Logger.log(`❌ Erro ao calcular qtd original: ${e.message}`);
    return qtdAbertaAtual;
  }
}

// ====== GERAR IDs COM SUFIXO NUMÉRICO ======

/**
 * Cria um menu personalizado na planilha ao abri-la.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('IDs Personalizados')
    .addItem('1. Gerar IDs Faltantes', 'gerarIDsUnicos')
    .addSeparator()
    .addItem('2. Ativar Geração Automática (a cada 5 min)', 'instalarTriggerAutomatico')
    .addItem('3. Desativar Geração Automática', 'desinstalarTriggerAutomatico')
    .addItem('4. Status do Trigger', 'mostrarStatusTrigger')
    .addToUi();
}

/**
 * Função principal para gerar os IDs únicos e estáticos com sufixo numérico.
 * Esta função é chamada manualmente ou pelo trigger automático.
 *
 * IMPORTANTE: Para evitar desalinhamento com IMPORTRANGE, esta função:
 * 1. LIMPA toda a coluna A (remove IDs antigos)
 * 2. LÊ dados atuais do IMPORTRANGE
 * 3. GERA novos IDs alinhados com os dados atuais
 *
 * Os IDs são baseados em dados + sufixo numérico sequencial.
 */
function gerarIDsUnicos() {
  Logger.log("=== GERANDO IDs COM SUFIXO NUMÉRICO ===");

  const sheet = SS.getSheetByName(FONTE_SHEET_NAME);

  if (!sheet) {
    Logger.log('❌ A aba "' + FONTE_SHEET_NAME + '" não foi encontrada!');
    return { gerados: 0, erro: 'Aba não encontrada' };
  }

  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < FONTE_DATA_START_ROW) {
    Logger.log('⚠️ Não há dados para processar na aba "' + FONTE_SHEET_NAME + '".');
    return { gerados: 0, erro: 'Sem dados' };
  }

  // PASSO 1: LIMPAR coluna A (IDs antigos) para evitar desalinhamento
  Logger.log(`🧹 Limpando coluna A (linhas ${FONTE_DATA_START_ROW} até ${ultimaLinha})...`);
  const rangeParaLimpar = sheet.getRange(FONTE_DATA_START_ROW, 1, ultimaLinha - FONTE_DATA_START_ROW + 1, 1);
  rangeParaLimpar.clearContent();
  SpreadsheetApp.flush();

  // PASSO 2: LER dados atuais do IMPORTRANGE (colunas B+)
  const intervalo = sheet.getRange(FONTE_DATA_START_ROW, 1, ultimaLinha - FONTE_DATA_START_ROW + 1, sheet.getLastColumn());
  const valores = intervalo.getValues();

  const contagemIDs = {};
  const novosValores = [];
  let idsGerados = 0;

  // PASSO 3: GERAR IDs para TODAS as linhas com dados
  valores.forEach(function(linha, i) {
    // Verifica se linha tem dados (coluna CARTELA preenchida)
    const cartela = linha[CARTELA_COL];

    if (!cartela || String(cartela).trim() === "") {
      novosValores.push([""]);
      return;
    }

    // Concatenação das colunas para criar ID base: C + D + E + F + H + I + G + J + L + M
    // Trata data de forma consistente (formata para yyyyMMdd se for Date)
    const dataReceb = linha[12]; // Coluna M - DATA RECEB.
    const dataFormatada = dataReceb instanceof Date ?
      Utilities.formatDate(dataReceb, TZ, 'yyyyMMdd') :
      String(dataReceb || '').trim();

    const idBase = "" +
      String(linha[2] || '').trim() + // Coluna C - CLIENTE
      String(linha[3] || '').trim() + // Coluna D
      String(linha[4] || '').trim() + // Coluna E - PEDIDO
      String(linha[5] || '').trim() + // Coluna F - CÓD. CLIENTE
      String(linha[7] || '').trim() + // Coluna H - DESCRIÇÃO
      String(linha[8] || '').trim() + // Coluna I - TAMANHO
      String(linha[6] || '').trim() + // Coluna G - CÓD. MARFIM
      String(linha[9] || '').trim() + // Coluna J - ORD. COMPRA
      String(linha[11] || '').trim() + // Coluna L - CÓD. OS
      dataFormatada;  // Coluna M - DATA RECEB. (formatada)

    if (idBase.trim() === "") {
      novosValores.push([""]);
      return;
    }

    // Gera sufixo sequencial (1, 2, 3...) para itens com mesmos dados
    const sufixoAtual = contagemIDs[idBase] || 0;
    const novoSufixo = sufixoAtual + 1;
    contagemIDs[idBase] = novoSufixo;

    const novoID = idBase + "-" + novoSufixo;

    novosValores.push([novoID]);
    idsGerados++;
    Logger.log(`  ✓ Linha ${i + FONTE_DATA_START_ROW}: ${novoID} (novo)`);
  });

  // PASSO 4: Escrever IDs na coluna A (agora alinhados com IMPORTRANGE)
  if (idsGerados > 0) {
    sheet.getRange(FONTE_DATA_START_ROW, 1, novosValores.length, 1).setValues(novosValores);
    SpreadsheetApp.flush();
    Logger.log(`✅ ${idsGerados} IDs gerados com sucesso (coluna A alinhada com IMPORTRANGE)!`);
    limparCache();
    return { gerados: idsGerados, erro: null };
  } else {
    Logger.log('⚠️ Nenhum ID gerado (sem dados válidos).');
    return { gerados: 0, erro: null };
  }
}

/**
 * Função INTELIGENTE que só regenera IDs quando REALMENTE necessário.
 * Usada pelo trigger automático.
 *
 * OTIMIZAÇÃO: Verifica se há mudanças antes de regenerar (performance!)
 * - Compara quantidade de linhas
 * - Verifica se há IDs faltantes
 * - Só regenera se detectar inconsistência
 */
function verificarEGerarIDs() {
  try {
    const sheet = SS.getSheetByName(FONTE_SHEET_NAME);
    if (!sheet) return { regenerou: false, motivo: 'Aba não encontrada' };

    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < FONTE_DATA_START_ROW) {
      return { regenerou: false, motivo: 'Sem dados' };
    }

    // PASSO 1: Verificar se realmente precisa regenerar
    Logger.log("🔍 Verificando se precisa regenerar IDs...");

    const numLinhas = ultimaLinha - FONTE_DATA_START_ROW + 1;

    // Lê apenas colunas A (ID) e B (CARTELA) para performance
    const range = sheet.getRange(FONTE_DATA_START_ROW, 1, numLinhas, 2);
    const dados = range.getValues();

    let linhasComDados = 0;
    let linhasComId = 0;
    let linhasSemIdMasComDados = 0;

    dados.forEach(row => {
      const id = row[0];
      const cartela = row[1];

      if (cartela && String(cartela).trim() !== '') {
        linhasComDados++;
        if (id && String(id).trim() !== '') {
          linhasComId++;
        } else {
          linhasSemIdMasComDados++;
        }
      }
    });

    Logger.log(`   📊 Estatísticas:`);
    Logger.log(`      - Linhas com dados: ${linhasComDados}`);
    Logger.log(`      - Linhas com ID: ${linhasComId}`);
    Logger.log(`      - Linhas sem ID mas com dados: ${linhasSemIdMasComDados}`);

    // DECISÃO: Só regenera se houver linhas sem ID
    if (linhasSemIdMasComDados === 0 && linhasComDados === linhasComId) {
      Logger.log("   ✅ Todos os IDs estão OK - NADA A FAZER");
      Logger.log("   🚀 Performance: Regeneração não necessária!");
      return { regenerou: false, motivo: 'IDs já estão corretos' };
    }

    // PRECISA REGENERAR
    Logger.log(`   ⚠️ Encontradas ${linhasSemIdMasComDados} linhas sem ID`);
    Logger.log("   🔄 Regenerando IDs...");

    const resultado = gerarIDsUnicos();

    if (resultado.gerados > 0) {
      Logger.log(`   ✅ ${resultado.gerados} IDs regenerados com sucesso`);
      return { regenerou: true, gerados: resultado.gerados };
    } else {
      Logger.log("   ✓ Nenhum ID gerado");
      return { regenerou: false, motivo: 'Sem dados válidos' };
    }
  } catch (e) {
    Logger.log(`❌ Erro na regeneração de IDs: ${e.message}`);
    return { regenerou: false, erro: e.message };
  }
}

/**
 * PROCESSO AUTOMÁTICO COMPLETO OTIMIZADO
 * Executa a cada 5 minutos via trigger
 *
 * OTIMIZAÇÕES:
 * 1. Só regenera IDs se necessário (performance!)
 * 2. Só limpa cache se houve mudanças (UX!)
 * 3. Log detalhado de performance
 */
function processoAutomaticoCompleto() {
  const inicioProcesso = Date.now();
  Logger.log("=" .repeat(70));
  Logger.log(`⏰ PROCESSO AUTOMÁTICO INICIADO - ${new Date().toLocaleString('pt-BR')}`);
  Logger.log("=".repeat(70));

  let houveMudancas = false;

  try {
    // ETAPA 1: Verificar e gerar IDs faltantes
    Logger.log("\n🔑 ETAPA 1: Verificação de IDs");
    const resultadoIds = verificarEGerarIDs();

    if (resultadoIds.regenerou) {
      Logger.log(`   ✅ IDs regenerados: ${resultadoIds.gerados || 0}`);
      houveMudancas = true;
    } else {
      Logger.log(`   ✓ ${resultadoIds.motivo || 'Nenhuma alteração necessária'}`);
    }

    // ETAPA 2: Sincronizar dados
    Logger.log("\n🔄 ETAPA 2: Sincronização de dados");
    const resultadoSync = sincronizarDadosOtimizado();

    if (resultadoSync.houveMudancas) {
      Logger.log(`   ✅ Mudanças detectadas na sincronização`);
      houveMudancas = true;
    } else {
      Logger.log(`   ✓ Nenhuma mudança - dados já sincronizados`);
    }

    // ETAPA 3: Limpar cache APENAS se houve mudanças
    Logger.log("\n🗑️ ETAPA 3: Limpeza de cache");
    if (houveMudancas) {
      limparCache();
      Logger.log("   ✅ Cache limpo (houve mudanças)");
    } else {
      Logger.log("   ⏭️  Cache mantido (sem mudanças - melhor performance para usuários!)");
    }

    const tempoTotal = Date.now() - inicioProcesso;
    Logger.log("\n" + "=".repeat(70));
    Logger.log(`✅ PROCESSO AUTOMÁTICO CONCLUÍDO EM ${tempoTotal}ms`);
    if (!houveMudancas) {
      Logger.log(`🚀 OTIMIZAÇÃO: Nenhuma mudança detectada - usuários não afetados!`);
    }
    Logger.log("=".repeat(70));

  } catch (erro) {
    Logger.log("\n❌ ERRO NO PROCESSO AUTOMÁTICO:");
    Logger.log(`   Mensagem: ${erro.message}`);
    Logger.log(`   Stack: ${erro.stack}`);
    Logger.log("=".repeat(70));

    // Envia email de notificação em caso de erro (opcional)
    // MailApp.sendEmail({
    //   to: Session.getEffectiveUser().getEmail(),
    //   subject: "⚠️ Erro no Processo Automático",
    //   body: `Erro: ${erro.message}\n\nDetalhes: ${erro.stack}`
    // });
  }
}

/**
 * Instala o trigger automático SEM ALERTAS (para executar pelo Apps Script).
 * Use esta função quando executar pelo Apps Script Editor.
 */
function instalarTriggerAutomaticoSilencioso() {
  try {
    Logger.log("🔄 Instalando trigger automático...");

    // Remove triggers antigos
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;

    triggers.forEach(trigger => {
      const funcao = trigger.getHandlerFunction();
      if (funcao === 'verificarEGerarIDs' || funcao === 'processoAutomaticoCompleto') {
        ScriptApp.deleteTrigger(trigger);
        removidos++;
        Logger.log(`   ✓ Removido trigger: ${funcao}`);
      }
    });

    if (removidos > 0) {
      Logger.log(`✅ ${removidos} trigger(s) antigo(s) removido(s)`);
    }

    // Cria novo trigger
    ScriptApp.newTrigger('processoAutomaticoCompleto')
      .timeBased()
      .everyMinutes(5)
      .create();

    Logger.log("✅ TRIGGER INSTALADO COM SUCESSO!");
    Logger.log("📋 Detalhes:");
    Logger.log("   • Função: processoAutomaticoCompleto");
    Logger.log("   • Frequência: A cada 5 minutos");
    Logger.log("   • Status: ATIVO");
    Logger.log("");
    Logger.log("🎯 O sistema automático está rodando!");
    Logger.log("   • Gera IDs faltantes automaticamente");
    Logger.log("   • Sincroniza PEDIDOS → Relatorio_DB");
    Logger.log("   • Mantém dados sempre atualizados");

    return {
      success: true,
      message: 'Trigger instalado com sucesso',
      funcao: 'processoAutomaticoCompleto',
      frequencia: '5 minutos'
    };

  } catch (e) {
    Logger.log(`❌ ERRO ao instalar trigger: ${e.message}`);
    Logger.log(`   Stack: ${e.stack}`);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Instala o trigger automático que executa a cada 5 minutos
 * IMPORTANTE: Este trigger chama processoAutomaticoCompleto() que faz TUDO
 */
function instalarTriggerAutomatico() {
  try {
    // Remove triggers antigos para evitar duplicatas
    desinstalarTriggerAutomatico();

    // Cria novo trigger que executa o processo completo
    ScriptApp.newTrigger('processoAutomaticoCompleto')
      .timeBased()
      .everyMinutes(5)
      .create();

    SpreadsheetApp.getUi().alert(
      '✅ Trigger Automático Ativado!',
      'O sistema automático está ativo e executará a cada 5 minutos:\n\n' +
      '• Gera IDs faltantes automaticamente\n' +
      '• Sincroniza PEDIDOS → Relatorio_DB\n' +
      '• Mantém dados sempre atualizados\n\n' +
      'Para desativar, use o menu: IDs Personalizados > Desativar Geração Automática',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    Logger.log("✅ Trigger automático completo instalado com sucesso");
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Erro ao instalar trigger: ' + e.message);
    Logger.log(`❌ Erro ao instalar trigger: ${e.message}`);
  }
}

/**
 * Remove o trigger automático
 * Remove triggers de verificarEGerarIDs e processoAutomaticoCompleto
 */
function desinstalarTriggerAutomatico() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;

    triggers.forEach(trigger => {
      const funcao = trigger.getHandlerFunction();
      if (funcao === 'verificarEGerarIDs' || funcao === 'processoAutomaticoCompleto') {
        ScriptApp.deleteTrigger(trigger);
        removidos++;
        Logger.log(`   Removido trigger: ${funcao}`);
      }
    });

    if (removidos > 0) {
      SpreadsheetApp.getUi().alert(
        '✅ Trigger Desativado!',
        `O sistema automático foi desativado.\n\n` +
        `${removidos} trigger(s) removido(s).\n\n` +
        'Você ainda pode:\n' +
        '• Gerar IDs manualmente: IDs Personalizados > Gerar IDs Faltantes\n' +
        '• Sincronizar manualmente: Use a função sincronizarDados()',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      Logger.log(`✅ ${removidos} trigger(s) removido(s)`);
    } else {
      SpreadsheetApp.getUi().alert(
        'ℹ️ Nenhum Trigger Ativo',
        'Não há triggers automáticos instalados.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      Logger.log("ℹ️ Nenhum trigger encontrado para remover");
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Erro ao desinstalar trigger: ' + e.message);
    Logger.log(`❌ Erro ao desinstalar trigger: ${e.message}`);
  }
}

/**
 * Mostra o status dos triggers instalados
 */
function mostrarStatusTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const triggersAtivos = triggers.filter(t =>
      t.getHandlerFunction() === 'verificarEGerarIDs' ||
      t.getHandlerFunction() === 'processoAutomaticoCompleto'
    );

    if (triggersAtivos.length > 0) {
      const trigger = triggersAtivos[0];
      const funcao = trigger.getHandlerFunction();
      const eventType = trigger.getEventType();

      const descricao = funcao === 'processoAutomaticoCompleto'
        ? 'Processo Completo (Gera IDs + Sincroniza)'
        : 'Geração de IDs';

      SpreadsheetApp.getUi().alert(
        '✅ Trigger Ativo',
        `Status: ATIVO\n` +
        `Função: ${funcao}\n` +
        `Descrição: ${descricao}\n` +
        `Tipo: ${eventType}\n` +
        `Frequência: A cada 5 minutos\n` +
        `Triggers instalados: ${triggersAtivos.length}\n\n` +
        'O sistema automático está rodando:\n' +
        '• Gera IDs faltantes\n' +
        '• Sincroniza PEDIDOS → Relatorio_DB\n' +
        '• Mantém dados sempre atualizados',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'ℹ️ Trigger Inativo',
        'Status: INATIVO\n\n' +
        'O sistema automático não está ativo.\n\n' +
        'Para ativar: IDs Personalizados > Ativar Geração Automática',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Erro ao verificar status: ' + e.message);
    Logger.log(`❌ Erro ao verificar status: ${e.message}`);
  }
}

// ====== FUNÇÃO LEGADA (mantida para compatibilidade) ======

// Gera ID composto baseado nas colunas C,D,E,F,G,J,L,M
function _gerarIdComposto_(row) {
  // Colunas solicitadas: C,D,E,F,G,J,L,M
  const colC = String(row[2] || '').trim();  // C = CLIENTE
  const colD = String(row[3] || '').trim();  // D = (coluna entre Cliente e Pedido)
  const colE = String(row[4] || '').trim();  // E = PEDIDO
  const colF = String(row[5] || '').trim();  // F = CÓD. CLIENTE
  const colG = String(row[6] || '').trim();  // G = CÓD. MARFIM
  const colJ = String(row[9] || '').trim();  // J = ORD. COMPRA
  const colL = String(row[11] || '').trim(); // L = CÓD. OS
  const colM = row[12]; // M = DATA RECEB.

  // Remove caracteres especiais e espaços
  const clean = (str) => str.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

  // Trata data especialmente
  const cleanM = colM instanceof Date ?
    Utilities.formatDate(colM, TZ, 'yyyyMMdd') :
    clean(String(colM || ''));

  // Concatena todas as colunas: C_D_E_F_G_J_L_M
  const id = `${clean(colC)}_${clean(colD)}_${clean(colE)}_${clean(colF)}_${clean(colG)}_${clean(colJ)}_${clean(colL)}_${cleanM}`;

  return id;
}

/**
 * Função legada - mantida para compatibilidade
 * Use gerarIDsUnicos() para o novo formato com sufixos numéricos
 */
function gerarIdsFaltantes() {
  Logger.clear();
  Logger.log("=== GERANDO IDs COMPOSTOS (FORMATO LEGADO) ===");

  const sheet = SS.getSheetByName(FONTE_SHEET_NAME);
  if (!sheet) { Logger.log("❌ Aba PEDIDOS não encontrada"); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < FONTE_DATA_START_ROW) { Logger.log("Sem dados"); return; }

  // Lê todas as colunas necessárias para gerar o ID
  const numCols = sheet.getLastColumn();
  const data = sheet.getRange(FONTE_DATA_START_ROW, 1, lastRow - FONTE_DATA_START_ROW + 1, numCols).getValues();

  let gerados = 0;
  let atualizados = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const idAtual = row[ID_COL];
    const idComposto = _gerarIdComposto_(row);

    // Se não tem ID ou o ID está no formato antigo, gera/atualiza
    const isFormatoAntigo = idAtual && (String(idAtual).startsWith('ID_') || String(idAtual).startsWith('CART_'));

    if (!idAtual || isFormatoAntigo) {
      sheet.getRange(i + FONTE_DATA_START_ROW, 1).setValue(idComposto);

      if (!idAtual) {
        gerados++;
        Logger.log(`  Linha ${i + FONTE_DATA_START_ROW}: ${idComposto} (novo)`);
      } else {
        atualizados++;
        Logger.log(`  Linha ${i + FONTE_DATA_START_ROW}: ${idAtual} → ${idComposto} (atualizado)`);
      }
    }
  }

  if (gerados > 0 || atualizados > 0) {
    SpreadsheetApp.flush();
    Logger.log(`✅ ${gerados} IDs novos gerados, ${atualizados} IDs atualizados para formato composto`);
    limparCache();
  } else {
    Logger.log("✅ Todos os IDs já estão no formato composto");
  }
}

// ====== FUNÇÕES AUXILIARES PARA SINCRONIZAÇÃO ======

/**
 * Cria uma "impressão digital" única dos dados para identificar itens.
 * Usado para comparar itens mesmo quando IDs mudam (devido ao IMPORTRANGE).
 *
 * Retorna uma string única baseada em: CARTELA + CLIENTE + PEDIDO + MARFIM + OC + OS + DATA
 */
function _criarImpressaoDigital_(row) {
  const partes = [
    String(row[CARTELA_COL] || '').trim(),
    String(row[CLIENTE_COL] || '').trim(),
    String(row[PEDIDO_COL] || '').trim(),
    String(row[MARFIM_COL] || '').trim(),
    String(row[OC_COL] || '').trim(),
    String(row[OS_COL] || '').trim(),
    row[DTREC_COL] instanceof Date ? row[DTREC_COL].toISOString() : String(row[DTREC_COL] || '')
  ];
  return partes.join('|');
}

/**
 * Cria um Map de impressões digitais do Relatorio_DB.
 * Retorna: Map<impressao_digital, {id, linha, row}>
 */
function _criarMapImpressoes_(dbData) {
  const map = new Map();
  dbData.forEach((row, idx) => {
    const impressao = _criarImpressaoDigital_(row);
    if (impressao && impressao !== '||||||') { // ignora linhas vazias
      map.set(impressao, {
        id: row[ID_COL],
        linha: idx + 2, // linha na planilha (primeira linha de dados = 2)
        row: row
      });
    }
  });
  return map;
}

// ====== SINCRONIZAÇÃO ======

/**
 * Versão otimizada da sincronização que retorna se houve mudanças.
 * Usada pelo processo automático para decidir se limpa cache.
 */
function sincronizarDadosOtimizado() {
  const resultado = sincronizarDados();
  const houveMudancas = resultado.novos > 0 || resultado.updates > 0 || resultado.inativos > 0;
  return { houveMudancas: houveMudancas, ...resultado };
}

function sincronizarDados() {
  Logger.clear();
  Logger.log("=".repeat(70));
  Logger.log(`SINCRONIZAÇÃO v${APP_VERSION} - ${new Date().toLocaleString('pt-BR')}`);
  Logger.log("=".repeat(70));
  
  const startTime = Date.now();
  
  try {
    const fonteSheet = SS.getSheetByName(FONTE_SHEET_NAME);
    const dbSheet = SS.getSheetByName(DB_SHEET_NAME);
    
    if (!fonteSheet || !dbSheet) { Logger.log("❌ Planilhas não encontradas"); return; }
    
    // 1) LER PEDIDOS (usa IDs que estão na planilha)
    Logger.log("\n📖 1. LENDO PEDIDOS");
    const allFonte = fonteSheet.getDataRange().getValues();
    const fonteData = allFonte.slice(FONTE_DATA_START_ROW - 1);

    const fonteMap = new Map();
    let semId = 0;
    let semCartela = 0;

    fonteData.forEach((row, idx) => {
      const id = row[ID_COL];
      const cartela = row[CARTELA_COL];

      // Ignora registros sem dados na coluna CARTELA
      if (!cartela || String(cartela).trim() === '') {
        semCartela++;
        if (id && String(id).trim()) {
          Logger.log(`   ⚠️ Linha ${idx + FONTE_DATA_START_ROW}: ID="${String(id).trim()}" sem CARTELA - será ignorado`);
        }
        return;
      }

      if (id && String(id).trim()) {
        const idStr = String(id).trim();
        fonteMap.set(idStr, row);
        Logger.log(`   ✓ PEDIDOS: ID="${idStr}", CARTELA="${cartela}"`);
      } else {
        semId++;
        Logger.log(`   ⚠️ Linha ${idx + FONTE_DATA_START_ROW}: SEM ID mas tem CARTELA="${cartela}"`);
      }
    });

    const totalFonte = fonteMap.size;
    Logger.log(`   ${totalFonte} itens com ID e CARTELA`);
    if (semId > 0) Logger.log(`   ⚠️ ${semId} sem ID - insira IDs manualmente na coluna A`);
    if (semCartela > 0) Logger.log(`   ⚠️ ${semCartela} sem CARTELA - ignorados`);

    // 2) LER Relatorio_DB
    Logger.log("\n📖 2. LENDO Relatorio_DB");
    const dbRows = dbSheet.getLastRow() - 1;
    let dbData = [];

    if (dbRows > 0) {
      // Lê 16 colunas: A-P (ID até MARCAR_FATURAR)
      // Status está na coluna O (índice 14 do array)
      // MARCAR_FATURAR está na coluna P (índice 15 do array)
      dbData = dbSheet.getRange(2, 1, dbRows, 16).getValues();
    }

    const dbMap = new Map();
    const statusCount = { Ativo: 0, Inativo: 0, Faturado: 0, Excluido: 0 };

    dbData.forEach((row, idx) => {
      const id = row[ID_COL];  // Coluna A (índice 0)
      if (id && String(id).trim()) {
        const idStr = String(id).trim();
        dbMap.set(idStr, { row: row, linha: idx + 2 });
        const st = row[STATUS_COL];  // Coluna O (índice 14)
        Logger.log(`   ✓ Relatorio_DB: ID="${idStr}", Status="${st}", Linha=${idx + 2}`);
        if (st === 'Ativo') statusCount.Ativo++;
        else if (st === 'Inativo') statusCount.Inativo++;
        else if (st === 'Faturado') statusCount.Faturado++;
        else if (st === 'Excluido') statusCount.Excluido++;
      }
    });

    const totalDB = dbMap.size;
    Logger.log(`   ${totalDB} itens`);
    Logger.log(`   Status: ${statusCount.Ativo} Ativo, ${statusCount.Inativo} Inativo, ${statusCount.Faturado} Faturado, ${statusCount.Excluido} Excluido`);

    // 2.5) CRIAR MAPS DE IMPRESSÕES DIGITAIS
    Logger.log("\n🔍 2.5. CRIANDO IMPRESSÕES DIGITAIS");

    // Map<impressao, {id, row}> para PEDIDOS
    const fonteImpressoes = new Map();
    for (let [id, row] of fonteMap.entries()) {
      const impressao = _criarImpressaoDigital_(row);
      fonteImpressoes.set(impressao, { id: id, row: row });
    }
    Logger.log(`   ✓ ${fonteImpressoes.size} impressões digitais criadas para PEDIDOS`);

    // Map<impressao, {id, linha, row}> para Relatorio_DB
    const dbImpressoes = new Map();
    for (let [id, dbItem] of dbMap.entries()) {
      const impressao = _criarImpressaoDigital_(dbItem.row);
      dbImpressoes.set(impressao, { id: id, linha: dbItem.linha, row: dbItem.row });
    }
    Logger.log(`   ✓ ${dbImpressoes.size} impressões digitais criadas para Relatorio_DB`);

    // 3) PROCESSAR
    Logger.log("\n🔄 3. PROCESSANDO");

    let novos = [];
    let updates = [];
    let marcaInativos = [];
    let idsAtualizados = [];

    for (let [id, dbItem] of dbMap.entries()) {
      const statusAtual = dbItem.row[STATUS_COL];  // Coluna O (índice 14)
      if (statusAtual === "Excluido") continue;

      // PRIMEIRA TENTATIVA: Buscar por ID
      if (fonteMap.has(id)) {
        Logger.log(`   🔄 Match encontrado: ID="${id}" existe em PEDIDOS e Relatorio_DB`);
        const fonteRow = fonteMap.get(id);

        // Array de 16 elementos (índices 0-15)
        // Posição 14 é Status na coluna O
        // Posição 15 é MARCAR_FATURAR na coluna P
        const marcarFaturarAtual = dbItem.row[MARCAR_FATURAR_COL] || "";
        const novaLinha = [
          fonteRow[ID_COL],      fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
          fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
          fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
          fonteRow[QTD_COL],     fonteRow[OS_COL],      fonteRow[DTREC_COL],
          fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "",                    marcarFaturarAtual
        ];

        let mudou = false;
        // Compara as 14 primeiras colunas (0-13), excluindo Status e MARCAR_FATURAR
        for (let i = 0; i < STATUS_COL; i++) {
          let dbVal = (dbItem.row[i] instanceof Date) ? dbItem.row[i].toISOString() : dbItem.row[i];
          let novoVal = (novaLinha[i] instanceof Date) ? novaLinha[i].toISOString() : novaLinha[i];
          if (dbVal != novoVal) { mudou = true; break; }
        }

        if (mudou || statusAtual === "Inativo") {
          const novoStatus = (statusAtual === "Faturado") ? "Faturado" : "Ativo";
          novaLinha[STATUS_COL] = novoStatus;  // Coluna O (índice 14)
          Logger.log(`   📝 Update: ID="${id}" Linha=${dbItem.linha} Status: ${statusAtual} → ${novoStatus}`);
          updates.push({ linha: dbItem.linha, dados: novaLinha, de: statusAtual, para: novoStatus });
        }

        fonteMap.delete(id);

      } else {
        // SEGUNDA TENTATIVA: Buscar por IMPRESSÃO DIGITAL (dados)
        const impressaoDB = _criarImpressaoDigital_(dbItem.row);
        const fonteItem = fonteImpressoes.get(impressaoDB);

        if (fonteItem) {
          // ENCONTROU POR DADOS! O ID mudou devido ao IMPORTRANGE
          const novoId = fonteItem.id;
          Logger.log(`   🔄 Item encontrado por dados: ID mudou "${id}" → "${novoId}"`);
          Logger.log(`      Linha=${dbItem.linha}, Status="${statusAtual}"`);
          Logger.log(`      Atualizando ID e dados no Relatorio_DB...`);

          const fonteRow = fonteItem.row;
          const marcarFaturarAtual = dbItem.row[MARCAR_FATURAR_COL] || "";

          // Atualiza com NOVO ID
          const novaLinha = [
            novoId,                fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
            fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
            fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
            fonteRow[QTD_COL],     fonteRow[OS_COL],      fonteRow[DTREC_COL],
            fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "",                    marcarFaturarAtual
          ];

          const novoStatus = (statusAtual === "Faturado") ? "Faturado" : "Ativo";
          novaLinha[STATUS_COL] = novoStatus;

          updates.push({ linha: dbItem.linha, dados: novaLinha, de: statusAtual, para: novoStatus });
          idsAtualizados.push({ de: id, para: novoId, linha: dbItem.linha });

          // Remove do fonteMap para não adicionar como novo depois
          fonteMap.delete(novoId);

        } else {
          // NÃO ENCONTROU nem por ID nem por dados - item realmente sumiu
          Logger.log(`   ❌ ID="${id}" não encontrado em PEDIDOS (nem por ID nem por dados)`);
          Logger.log(`      Status atual: "${statusAtual}", Linha: ${dbItem.linha}`);

          if (statusAtual !== "Faturado" && statusAtual !== "Inativo") {
            Logger.log(`   ⚠️ Será marcado como Inativo`);
            marcaInativos.push({ linha: dbItem.linha, id: id, de: statusAtual });
          } else {
            Logger.log(`   ℹ️ Não será alterado (já é ${statusAtual})`);
          }
        }
      }
    }
    
    // Novos itens que estão em PEDIDOS mas não em Relatorio_DB
    for (let [id, fonteRow] of fonteMap.entries()) {
      Logger.log(`   🆕 Novo item: ID="${id}" está em PEDIDOS mas não em Relatorio_DB - será adicionado como Ativo`);
      Logger.log(`      CARTELA="${fonteRow[CARTELA_COL]}", CLIENTE="${fonteRow[CLIENTE_COL]}", OC="${fonteRow[OC_COL]}"`);

      // Array de 16 elementos, Status (índice 14) = "Ativo", MARCAR_FATURAR (índice 15) = ""
      const novaLinha = [
        fonteRow[ID_COL],      fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
        fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
        fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
        fonteRow[QTD_COL],     fonteRow[OS_COL],      fonteRow[DTREC_COL],
        fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "Ativo",               ""
      ];
      novos.push(novaLinha);
    }
    
    Logger.log(`   🆕 Novos: ${novos.length}`);
    Logger.log(`   📝 Atualizar: ${updates.length}`);
    Logger.log(`   🔄 IDs Atualizados: ${idsAtualizados.length}`);
    Logger.log(`   ⚠️ Marcar Inativo: ${marcaInativos.length}`);

    // 4) VALIDAÇÃO ANTI-DUPLICATA
    Logger.log("\n🔍 3.5. VALIDAÇÃO ANTI-DUPLICATA");
    const novosValidados = [];
    const idsExistentes = new Set(dbMap.keys());
    const idsJaAdicionados = new Set();

    novos.forEach(item => {
      const id = String(item[ID_COL]).trim();

      // Verifica se já existe no DB
      if (idsExistentes.has(id)) {
        Logger.log(`   ⚠️ DUPLICATA EVITADA: ID="${id}" já existe no Relatorio_DB`);
        return;
      }

      // Verifica se já foi adicionado nesta rodada
      if (idsJaAdicionados.has(id)) {
        Logger.log(`   ⚠️ DUPLICATA EVITADA: ID="${id}" já foi processado nesta sincronização`);
        return;
      }

      // Valida se tem dados essenciais
      if (!item[CARTELA_COL] || String(item[CARTELA_COL]).trim() === '') {
        Logger.log(`   ⚠️ ITEM REJEITADO: ID="${id}" sem CARTELA`);
        return;
      }

      // Item válido - adiciona
      novosValidados.push(item);
      idsJaAdicionados.add(id);
    });

    const duplicatasEvitadas = novos.length - novosValidados.length;
    if (duplicatasEvitadas > 0) {
      Logger.log(`   🛡️ ${duplicatasEvitadas} duplicata(s) evitada(s)`);
    }
    Logger.log(`   ✓ ${novosValidados.length} itens validados para inserção`);

    // 5) APLICAR
    Logger.log("\n💾 4. APLICANDO");
    if (novosValidados.length > 0) {
      const proxLinha = dbSheet.getLastRow() + 1;
      dbSheet.getRange(proxLinha, 1, novosValidados.length, 16).setValues(novosValidados);
      Logger.log(`   ✅ ${novosValidados.length} novos adicionados`);
    }
    if (updates.length > 0) {
      updates.forEach(u => {
        dbSheet.getRange(u.linha, 1, 1, 16).setValues([u.dados]);
        Logger.log(`   ✅ Linha ${u.linha}: ${u.de} → ${u.para}`);
      });
    }
    if (marcaInativos.length > 0) {
      marcaInativos.forEach(m => {
        // STATUS_COL = 14 (índice do array)
        // +1 porque getRange usa índice baseado em 1, então coluna O = 15
        dbSheet.getRange(m.linha, STATUS_COL + 1, 1, 1).setValue("Inativo");
        Logger.log(`   ⚠️ Linha ${m.linha}: ${m.de} → Inativo (ID: ${m.id})`);
      });
    }
    
    SpreadsheetApp.flush();
    if (novosValidados.length > 0 || updates.length > 0 || marcaInativos.length > 0) {
      limparCache();
      Logger.log("   🗑️ Cache limpo");
    }

    const execTime = Date.now() - startTime;
    Logger.log("\n" + "=".repeat(70));
    Logger.log(`✅ SINCRONIZAÇÃO CONCLUÍDA (${execTime}ms)`);
    Logger.log("=".repeat(70));
    Logger.log("\n📊 RESUMO:");
    Logger.log(`   • ${totalFonte} itens lidos de PEDIDOS (com ID + CARTELA)`);
    Logger.log(`   • ${totalDB} itens lidos de Relatorio_DB`);
    Logger.log(`   • ${novosValidados.length} novos itens adicionados ao Relatorio_DB como Ativo`);
    Logger.log(`   • ${updates.length} itens atualizados no Relatorio_DB`);
    Logger.log(`   • ${marcaInativos.length} itens marcados como Inativo (não encontrados em PEDIDOS)`);
    if (idsAtualizados.length > 0) {
      Logger.log(`   🔄 ${idsAtualizados.length} IDs atualizados (por mudança de posição no IMPORTRANGE):`);
      idsAtualizados.forEach(ida => {
        Logger.log(`      - Linha ${ida.linha}: "${ida.de}" → "${ida.para}"`);
      });
    }
    if (duplicatasEvitadas > 0) Logger.log(`   🛡️ ${duplicatasEvitadas} duplicata(s) evitada(s)`);
    if (semId > 0) Logger.log(`   ⚠️ ${semId} linhas em PEDIDOS sem ID (ignoradas)`);
    if (semCartela > 0) Logger.log(`   ⚠️ ${semCartela} linhas em PEDIDOS sem CARTELA (ignoradas)`);
    Logger.log("=".repeat(70));

    // Retorna contadores para o processo automático decidir se limpa cache
    return {
      novos: novosValidados.length,
      updates: updates.length,
      inativos: marcaInativos.length,
      idsAtualizados: idsAtualizados.length
    };

  } catch (error) {
    Logger.log("\n❌ ERRO: " + error.message);
    throw error;
  }
}

// ====== CACHE ======
function limparCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('dados_completos');
    cache.remove('timestamp_dados');
    Logger.log("🗑️ Cache limpo");
  } catch (e) {
    Logger.log("⚠️ Erro ao limpar cache: " + e.message);
  }
}

function obterDadosCache() {
  try {
    const cache = CacheService.getScriptCache();
    const dadosStr = cache.get('dados_completos');
    const timestamp = cache.get('timestamp_dados');
    if (dadosStr && timestamp) {
      const dados = JSON.parse(dadosStr);
      const idade = Date.now() - parseInt(timestamp);
      Logger.log(`📦 Cache hit! Idade: ${Math.floor(idade/1000)}s`);
      return dados;
    }
  } catch (e) {
    Logger.log("⚠️ Erro ao ler cache: " + e.message);
  }
  return null;
}

function salvarDadosCache(dados) {
  try {
    const cache = CacheService.getScriptCache();
    const dadosStr = JSON.stringify(dados);
    if (dadosStr.length > 100000) {
      Logger.log("⚠️ Dados muito grandes para cache");
      return false;
    }
    cache.put('dados_completos', dadosStr, CACHE_DURATION);
    cache.put('timestamp_dados', Date.now().toString(), CACHE_DURATION);
    Logger.log(`💾 Cache salvo (${Math.floor(dadosStr.length/1024)}KB, válido por ${CACHE_DURATION/60}min)`);
    return true;
  } catch (e) {
    Logger.log("⚠️ Erro ao salvar cache: " + e.message);
    return false;
  }
}

// ====== SISTEMA WEB OTIMIZADO ======
function _readAllData_() {
  const sheet = SS.getSheetByName(DB_SHEET_NAME);
  if (!sheet) throw new Error(`Aba '${DB_SHEET_NAME}' não encontrada`);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { headers: [], rows: [], displayRows: [] };
  const lastCol = sheet.getLastColumn();

  // Valores crus (para números/datas) + valores exibidos (para códigos/IDs/textos)
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();
  const display = range.getDisplayValues();

  return {
    headers: values[0],
    rows: values.slice(1),
    displayRows: display.slice(1)
  };
}

function _getColumnIndexes_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i;
  });
  return map;
}

// Usa displayRow para campos textuais/identificadores (evita virar Data / perder zeros à esquerda)
function _rowToItem_(row, displayRow, colMap, rowIndex) {
  const get = (colName, def = '') => {
    const idx = colMap[colName];
    return (idx !== undefined && idx < row.length) ? row[idx] : def;
  };
  const getDisp = (colName, def = '') => {
    const idx = colMap[colName];
    return (idx !== undefined && idx < displayRow.length) ? displayRow[idx] : def;
  };

  const uniqueId = getDisp('ID_UNICO');
  const qtdAberta = _toNumber_(get('QTD. ABERTA', 0));

  const item = {
    uniqueId: uniqueId,                 // id textual
    planilhaLinha: rowIndex + 2,

    // TEXTUAIS/IDENTIFICADORES via display
    CARTELA: getDisp('CARTELA', 'N/A'),
    'CÓD. CLIENTE': getDisp('CÓD. CLIENTE', 'N/A'),
    'DESCRIÇÃO': getDisp('DESCRIÇÃO', 'N/A'),
    'TAMANHO': getDisp('TAMANHO', 'N/A'),
    'CÓD. MARFIM': getDisp('CÓD. MARFIM', 'N/A'),
    'CÓD. OS': getDisp('CÓD. OS', 'N/A'),
    'ORD. COMPRA': getDisp('ORD. COMPRA', 'SEM OC'),
    CLIENTE: getDisp('CLIENTE', 'SEM CLIENTE'),
    PEDIDO: getDisp('PEDIDO', 'N/A'),

    // NÚMEROS/DATA cruas
    'QTD. ABERTA': qtdAberta,
    'QTD. ORIGINAL': calcularQtdOriginal(uniqueId, qtdAberta),
    'PRAZO': get('PRAZO', null),
    'DT. ENTREGA': get('DT. ENTREGA', null),
    'DATA RECEB.': get('DATA RECEB.', null),

    Status: getDisp('Status', 'Desconhecido'),
    MARCAR_FATURAR: getDisp('MARCAR_FATURAR', '') // Nova coluna para marcação de faturamento
  };

  if (!item.uniqueId) return null;
  return item;
}

function _organizeByOC_(items) {
  const byOC = {};
  items.forEach(item => {
    const oc = item['ORD. COMPRA'] || 'SEM OC';
    if (!byOC[oc]) {
      byOC[oc] = {
        ordCompraId: oc,
        ordCompra: oc,      // alias para compatibilidade com o front
        cliente: item.CLIENTE,
        items: []
      };
    }
    byOC[oc].items.push(item);
  });
  return Object.values(byOC);
}

function _getAccessCount_() {
  try {
    const cache = CacheService.getScriptCache();
    const key = 'accessCount';
    let count = parseInt(cache.get(key) || '0');
    count++;
    cache.put(key, count.toString(), 21600); // 6h
    return count;
  } catch (e) {
    return 0;
  }
}

function fetchAllDataUnified(cacheBuster) {
  const startTime = Date.now();
  Logger.log(`🚀 FETCH v${APP_VERSION} - ${new Date().toLocaleTimeString('pt-BR')}`);
  
  try {
    // TENTAR CACHE PRIMEIRO
    if (!cacheBuster) {
      const cached = obterDadosCache();
      if (cached) {
        cached.meta.fromCache = true;
        cached.meta.cacheHit = true;
        cached.meta.executionTime = Date.now() - startTime;
        Logger.log(`✅ Retornado do cache em ${cached.meta.executionTime}ms`);
        return cached;
      }
    }
    
    Logger.log("📊 Cache miss - lendo planilha...");
    const { headers, rows, displayRows } = _readAllData_();
    
    if (rows.length === 0) {
      const emptyResult = {
        success: true,
        ordCompras: [],
        stats: { totalItems: 0, totalOCs: 0, ativos: 0, inativos: 0, faturados: 0, excluidos: 0 },
        meta: {
          version: APP_VERSION,
          timestamp: new Date().toISOString(),
          displayTime: _fmtBRDateTime_(new Date()),
          executionTime: Date.now() - startTime,
          accessCount: _getAccessCount_(),
          fromCache: false
        }
      };
      salvarDadosCache(emptyResult);
      return JSON.parse(JSON.stringify(emptyResult));
    }
    
    const colMap = _getColumnIndexes_(headers);
    const itemsWeb = rows
      .map((row, idx) => _rowToItem_(row, displayRows[idx], colMap, idx))
      .filter(item => item !== null);
    
    const ordCompras = _organizeByOC_(itemsWeb);
    
    const stats = {
      totalItems: itemsWeb.length,
      totalOCs: ordCompras.length,
      ativos: itemsWeb.filter(i => i.Status === 'Ativo').length,
      inativos: itemsWeb.filter(i => i.Status === 'Inativo').length,
      faturados: itemsWeb.filter(i => i.Status === 'Faturado').length,
      excluidos: itemsWeb.filter(i => i.Status === 'Excluido').length
    };
    
    const result = {
      success: true,
      ordCompras: ordCompras, // payload enxuto
      stats: stats,
      meta: {
        version: APP_VERSION,
        timestamp: new Date().toISOString(),
        displayTime: _fmtBRDateTime_(new Date()),
        executionTime: Date.now() - startTime,
        accessCount: _getAccessCount_(),
        fromCache: false,
        itemCount: itemsWeb.length
      }
    };
    
    salvarDadosCache(result);
    return JSON.parse(JSON.stringify(result)); // garante tipos JSON puros
    
  } catch (error) {
    Logger.log(`❌ ${error.message}`);
    return {
      success: false,
      error: error.message,
      ordCompras: [],
      stats: { totalItems: 0, totalOCs: 0, ativos: 0, inativos: 0, faturados: 0, excluidos: 0 },
      meta: {
        version: APP_VERSION,
        timestamp: new Date().toISOString(),
        executionTime: Date.now() - startTime,
        fromCache: false
      }
    };
  }
}

// ====== COMPATIBILIDADE ======
function getOrdCompraList() {
  const data = fetchAllDataUnified();
  if (!data.success) return [];
  return data.ordCompras.map(oc => ({
    ordCompraId: oc.ordCompraId,
    cliente: oc.cliente,
    itemCount: oc.items.length
  }));
}

function getItensForOrdCompra(ordCompraId) {
  const data = fetchAllDataUnified();
  if (!data.success) return [];
  const oc = data.ordCompras.find(o => o.ordCompraId === ordCompraId || o.ordCompra === ordCompraId);
  return oc ? oc.items : [];
}

// ====== AÇÕES (com validação de linha e batches tolerantes) ======
function marcarFaturado(uniqueId, planilhaLinha) {
  try {
    const sheet = SS.getSheetByName(DB_SHEET_NAME);
    const linhaNum = Number(planilhaLinha);
    if (!sheet) throw new Error("Aba DB não encontrada");
    if (!isFinite(linhaNum) || linhaNum < 2 || linhaNum > sheet.getLastRow()) {
      throw new Error(`Linha inválida: ${planilhaLinha}`);
    }

    // Lê cabeçalhos para encontrar coluna Status dinamicamente
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = _getColumnIndexes_(headers);
    const statusCol = colMap['Status'];

    if (statusCol === undefined) {
      throw new Error("Coluna 'Status' não encontrada");
    }

    sheet.getRange(linhaNum, statusCol + 1).setValue("Faturado");
    limparCache();
    Logger.log(`💰 ${uniqueId || 'sem-id'} → Faturado (linha ${linhaNum}, coluna ${statusCol + 1})`);
    return { success: true, id: uniqueId || null, linha: linhaNum };
  } catch (e) {
    Logger.log(`❌ marcarFaturado: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

function excluirItem(uniqueId, planilhaLinha) {
  try {
    const sheet = SS.getSheetByName(DB_SHEET_NAME);
    const linhaNum = Number(planilhaLinha);
    if (!sheet) throw new Error("Aba DB não encontrada");
    if (!isFinite(linhaNum) || linhaNum < 2 || linhaNum > sheet.getLastRow()) {
      throw new Error(`Linha inválida: ${planilhaLinha}`);
    }

    // Lê cabeçalhos para encontrar coluna Status dinamicamente
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = _getColumnIndexes_(headers);
    const statusCol = colMap['Status'];

    if (statusCol === undefined) {
      throw new Error("Coluna 'Status' não encontrada");
    }

    sheet.getRange(linhaNum, statusCol + 1).setValue("Excluido");
    limparCache();
    Logger.log(`🗑️ ${uniqueId || 'sem-id'} → Excluido (linha ${linhaNum}, coluna ${statusCol + 1})`);
    return { success: true, id: uniqueId || null, linha: linhaNum };
  } catch (e) {
    Logger.log(`❌ excluirItem: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

function finalizarItem(uniqueId, planilhaLinha) {
  try {
    const sheet = SS.getSheetByName(DB_SHEET_NAME);
    const linhaNum = Number(planilhaLinha);
    if (!sheet) throw new Error("Aba DB não encontrada");
    if (!isFinite(linhaNum) || linhaNum < 2 || linhaNum > sheet.getLastRow()) {
      throw new Error(`Linha inválida: ${planilhaLinha}`);
    }

    // Lê cabeçalhos para encontrar coluna Status dinamicamente
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = _getColumnIndexes_(headers);
    const statusCol = colMap['Status'];

    if (statusCol === undefined) {
      throw new Error("Coluna 'Status' não encontrada");
    }

    sheet.getRange(linhaNum, statusCol + 1).setValue("Finalizado");
    limparCache();
    Logger.log(`✅ ${uniqueId || 'sem-id'} → Finalizado (linha ${linhaNum}, coluna ${statusCol + 1})`);
    return { success: true, id: uniqueId || null, linha: linhaNum };
  } catch (e) {
    Logger.log(`❌ finalizarItem: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

// --------- Batches tolerantes a 'linha' ou 'planilhaLinha' e com resumo ---------
function marcarMultiplosFaturados(items) {
  let ok = 0, fail = 0;
  const results = [];
  (items || []).forEach(it => {
    const linha = (it && it.planilhaLinha != null) ? it.planilhaLinha : (it ? it.linha : null);
    const id = (it && (it.uniqueId || it.id)) || null;
    const r = marcarFaturado(id, linha);
    results.push(r);
    r.success ? ok++ : fail++;
  });
  return { success: fail === 0, processados: ok, falhas: fail, results };
}

function excluirMultiplosItens(items) {
  let ok = 0, fail = 0;
  const results = [];
  (items || []).forEach(it => {
    const linha = (it && it.planilhaLinha != null) ? it.planilhaLinha : (it ? it.linha : null);
    const id = (it && (it.uniqueId || it.id)) || null;
    const r = excluirItem(id, linha);
    results.push(r);
    r.success ? ok++ : fail++;
  });
  return { success: fail === 0, processados: ok, falhas: fail, results };
}

function finalizarMultiplosItens(items) {
  let ok = 0, fail = 0;
  const results = [];
  (items || []).forEach(it => {
    const linha = (it && it.planilhaLinha != null) ? it.planilhaLinha : (it ? it.linha : null);
    const id = (it && (it.uniqueId || it.id)) || null;
    const r = finalizarItem(id, linha);
    results.push(r);
    r.success ? ok++ : fail++;
  });
  return { success: fail === 0, processados: ok, falhas: fail, results };
}

// ====== FUNÇÕES PARA MARCAR ITENS PARA FATURAMENTO ======

function marcarParaFaturar(uniqueId, planilhaLinha, marcar) {
  try {
    const sheet = SS.getSheetByName(DB_SHEET_NAME);
    const linhaNum = Number(planilhaLinha);

    if (!sheet) throw new Error("Aba DB não encontrada");
    if (!isFinite(linhaNum) || linhaNum < 2 || linhaNum > sheet.getLastRow()) {
      throw new Error(`Linha inválida: ${planilhaLinha}`);
    }

    // Lê cabeçalhos - força leitura de pelo menos 16 colunas (A-P)
    const numCols = Math.max(sheet.getLastColumn(), 16);
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = _getColumnIndexes_(headers);
    let marcarCol = colMap['MARCAR_FATURAR'];

    Logger.log(`📋 DEBUG marcarParaFaturar - Colunas lidas: ${numCols}, Headers: ${headers.length}`);
    Logger.log(`📋 DEBUG - Coluna P1 contém: "${headers[15] || 'VAZIO'}"`);
    Logger.log(`📋 DEBUG - MARCAR_FATURAR encontrada no índice: ${marcarCol}`);

    if (marcarCol === undefined) {
      Logger.log("⚠️ Coluna 'MARCAR_FATURAR' não encontrada - criando automaticamente...");

      // Cria a coluna MARCAR_FATURAR no cabeçalho (coluna P = 16)
      sheet.getRange(1, 16).setValue('MARCAR_FATURAR');
      marcarCol = 15; // índice da coluna P (base 0)

      Logger.log("✅ Coluna 'MARCAR_FATURAR' criada na coluna P");
    }

    // Marca ou desmarca
    const valor = marcar ? "SIM" : "";
    sheet.getRange(linhaNum, marcarCol + 1).setValue(valor);

    SpreadsheetApp.flush();
    limparCache();

    Logger.log(`✓ ${uniqueId} → Marcado para faturar: ${marcar} (linha ${linhaNum})`);
    return { success: true, id: uniqueId, linha: linhaNum, marcado: marcar };
  } catch (e) {
    Logger.log(`❌ marcarParaFaturar: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

function obterItensMarcadosParaFaturar() {
  Logger.log("🔍 INÍCIO obterItensMarcadosParaFaturar");

  try {
    const sheet = SS.getSheetByName(DB_SHEET_NAME);
    if (!sheet) {
      Logger.log("❌ Aba DB não encontrada");
      return { success: false, error: "Aba DB não encontrada", items: [] };
    }

    const lastRow = sheet.getLastRow();
    Logger.log(`📊 Total de linhas na planilha: ${lastRow}`);

    if (lastRow < 2) {
      Logger.log("⚠️ Planilha vazia (sem dados)");
      return { success: true, items: [] };
    }

    // Força leitura de pelo menos 16 colunas (A-P)
    const lastCol = Math.max(sheet.getLastColumn(), 16);
    Logger.log(`📊 Lendo ${lastCol} colunas (forçado mínimo 16)`);

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log(`📋 Headers lidos: ${headers.length} colunas`);
    Logger.log(`📋 Coluna P1 (índice 15) contém: "${headers[15] || 'VAZIO'}"`);

    const colMap = _getColumnIndexes_(headers);
    const marcarCol = colMap['MARCAR_FATURAR'];

    Logger.log(`📋 MARCAR_FATURAR encontrada no índice: ${marcarCol}`);

    if (marcarCol === undefined) {
      Logger.log("⚠️ Coluna 'MARCAR_FATURAR' não encontrada - criando automaticamente...");

      // Cria a coluna MARCAR_FATURAR no cabeçalho (coluna P = 16)
      sheet.getRange(1, 16).setValue('MARCAR_FATURAR');
      SpreadsheetApp.flush();

      Logger.log("✅ Coluna 'MARCAR_FATURAR' criada na coluna P");

      // Retorna lista vazia já que a coluna foi acabada de criar
      return { success: true, items: [], message: 'Coluna MARCAR_FATURAR criada. Clique novamente no botão.' };
    }

    // Lê dados completos
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getValues();
    const displayValues = range.getDisplayValues();

    const itensMarcados = [];

    values.forEach((row, idx) => {
      const marcarFaturar = String(row[marcarCol] || '').trim().toUpperCase();

      if (marcarFaturar === 'SIM') {
        const displayRow = displayValues[idx];
        const item = _rowToItem_(row, displayRow, colMap, idx);

        if (item) {
          // Calcula o saldo (soma das baixas)
          const qtdOriginal = item['QTD. ORIGINAL'] || 0;
          const qtdAberta = item['QTD. ABERTA'] || 0;
          const saldo = qtdOriginal - qtdAberta; // Total baixado

          // Serializa o item para JSON (converte Date objects para strings)
          const itemSerializado = {
            uniqueId: item.uniqueId,
            planilhaLinha: item.planilhaLinha,
            CARTELA: item.CARTELA,
            'CÓD. CLIENTE': item['CÓD. CLIENTE'],
            'DESCRIÇÃO': item['DESCRIÇÃO'],
            'TAMANHO': item['TAMANHO'],
            'CÓD. MARFIM': item['CÓD. MARFIM'],
            'CÓD. OS': item['CÓD. OS'],
            'ORD. COMPRA': item['ORD. COMPRA'],
            CLIENTE: item.CLIENTE,
            PEDIDO: item.PEDIDO,
            'QTD. ABERTA': item['QTD. ABERTA'],
            'QTD. ORIGINAL': item['QTD. ORIGINAL'],
            'PRAZO': _fmtBR_(item['PRAZO']),              // Converte Date para string
            'DT. ENTREGA': _fmtBR_(item['DT. ENTREGA']),  // Converte Date para string
            'DATA RECEB.': _fmtBR_(item['DATA RECEB.']),  // Converte Date para string
            Status: item.Status,
            MARCAR_FATURAR: item.MARCAR_FATURAR,
            SALDO: saldo
          };

          itensMarcados.push(itemSerializado);
        }
      }
    });

    Logger.log(`📋 Encontrados ${itensMarcados.length} itens marcados para faturar`);

    // Retorna com JSON.parse(JSON.stringify()) para garantir tipos JSON puros
    const result = { success: true, items: itensMarcados };
    return JSON.parse(JSON.stringify(result));

  } catch (e) {
    Logger.log(`❌ ERRO obterItensMarcadosParaFaturar: ${e.message}`);
    Logger.log(`❌ Stack: ${e.stack}`);
    return { success: false, error: e.message || 'Erro desconhecido', items: [] };
  } finally {
    Logger.log("🏁 FIM obterItensMarcadosParaFaturar");
  }
}
