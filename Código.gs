
// ==================================================== 
// SISTEMA DE RELATÓRIO DE PEDIDOS - v15.5 OTIMIZADA
// COM CACHE - CARREGAMENTO RÁPIDO
// ====================================================

// ====== CONFIGURAÇÃO ======
// Lazy getter — evita chamada de API no escopo global, que impede simple triggers (onOpen) de inicializar.
// getActiveSpreadsheet() funciona no contexto de script vinculado; openById() funciona no web app (doGet).
let _SS = null;
function getSpreadsheet_() {
  if (!_SS) _SS = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById("1qPJ8c7cq7qb86VJJ-iByeiaPnALOBcDPrPMeL75N2EI");
  return _SS;
}
const FONTE_SHEET_NAME = "PEDIDOS";
const IMPORTRANGE_SHEET_NAME = "DADOS_IMPORTADOS"; // Nova aba intermediária
const DB_SHEET_NAME = "Relatorio_DB";
const FONTE_DATA_START_ROW = 4;
const TZ = 'America/Fortaleza';
const APP_VERSION = '15.6-SINCRONIZACAO';

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
const TIMESTAMP_COL = 15; // P (na aba PEDIDOS) - Timestamp de criação do ID

// Índices de colunas - ABA Relatorio_DB
// O DB compacta as colunas: não tem col D do PEDIDOS, então índices diferem a partir de PEDIDO.
// DB:  [0]=ID, [1]=CARTELA, [2]=CLIENTE, [3]=PEDIDO, [4]=CODCLI, [5]=MARFIM,
//      [6]=DESC, [7]=TAM, [8]=OC, [9]=QTD, [10]=OS, [11]=DTREC, [12]=DTENT, [13]=PRAZO,
//      [14]=Status, [15]=MARCAR_FATURAR, [16]=DATA_STATUS
const DB_PEDIDO_COL  = 3;
const DB_CODCLI_COL  = 4;
const DB_MARFIM_COL  = 5;
const DB_DESC_COL    = 6;
const DB_TAM_COL     = 7;
const DB_OC_COL      = 8;
const DB_QTD_COL     = 9;  // QTD. ABERTA (≠ QTD_COL=10 que é da aba PEDIDOS)
const STATUS_COL = 14;   // O (coluna 15 ao contar a partir de 1)
const MARCAR_FATURAR_COL = 15; // P (coluna 16 ao contar a partir de 1) - Nova coluna para marcar itens para faturamento
const DATA_STATUS_COL = 16;    // Q (coluna 17) - Data em que o status foi alterado para Faturado/Finalizado/Excluido
const MARCAR_FATURAR_USUARIO_COL = 21; // V (coluna 22) - Usuário que marcou o item para faturamento
const PEDIDOS_CODIGO_FIXO_COL = 18; // S (coluna 19) — UUID fixo por item, gerado uma vez e preservado para sempre
const DB_CODIGO_FIXO_COL      = 18; // S (coluna 19) — mesmo UUID propagado do PEDIDOS para o Relatorio_DB
const PEDIDOS_POSICAO_FONTE_COL = 16; // Q (coluna 17) — índice do item em DADOS_IMPORTADOS (para manter ordem original)
const DB_POSICAO_FONTE_COL      = 17; // R (coluna 18) — posição propagada do PEDIDOS para o Relatorio_DB
const PEDIDOS_COLX_COL          = 19; // T (coluna 20) — campo da coluna X da fonte (informação adicional da OC)
const DB_COLX_COL               = 19; // T (coluna 20) — campo da coluna X propagado do PEDIDOS para o Relatorio_DB
const PEDIDOS_LOTE_COL          = 20; // U (coluna 21) — número de lote da coluna Y da fonte
const DB_LOTE_COL               = 20; // U (coluna 21) — número de lote propagado do PEDIDOS para o Relatorio_DB
const LOTE_DILLY_SHEET_NAME     = 'LOTE DILLY'; // aba com mapeamento OC→Lote para o cliente Dilly

// ====== ABA ORIGINAL (fonte primária para ordenação dos itens dentro de cada OC) ======
const ORIGINAL_SHEET_NAME = 'original'; // nome exato da aba
const ORIG_OC_COL   = 7;  // H — Ordem de Compra
const ORIG_DESC_COL = 5;  // F — Descrição
const ORIG_TAM_COL  = 6;  // G — Tamanho
const ORIG_QTD_COL  = 8;  // I — Quantidade
const ORIG_DATA_COL = 11; // L — Data
const DIAS_RETENCAO = 15;      // Itens com status final são purgados após este número de dias

// ====== BAIXAS PARCIAIS ======
const BAIXAS_SHEET_NAME = "Baixas_Historico";

// ====== FUNÇÕES AUXILIARES SEGURAS ======
/**
 * Converte um valor Date para ISO string de forma segura.
 * Retorna string vazia se a data for inválida.
 * @param {Date} date - Objeto Date para converter
 * @returns {string} String ISO ou vazio se inválido
 */
function _toISOStringSafe_(date) {
  if (!(date instanceof Date)) return '';
  // Verifica se a data é válida
  if (isNaN(date.getTime())) return '';
  try {
    return date.toISOString();
  } catch (e) {
    Logger.log(`⚠️ Erro ao converter data para ISO: ${e.message}`);
    return '';
  }
}

// ====== FUNÇÃO WEB APP ======
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Relatório de Pedidos v15.6')
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
  if (v instanceof Date) return 0; // Date objects produzem strings enormes após regex → sempre 0
  const s = String(v || '').replace(/[^\d,.-]/g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// ====== FUNÇÕES DE BAIXAS PARCIAIS ======

function _getBaixasSheet_() {
  let sheet = getSpreadsheet_().getSheetByName(BAIXAS_SHEET_NAME);
  if (!sheet) {
    Logger.log(`📝 Criando aba ${BAIXAS_SHEET_NAME}...`);
    sheet = getSpreadsheet_().insertSheet(BAIXAS_SHEET_NAME);
    // Criar cabeçalho
    sheet.getRange(1, 1, 1, 6).setValues([[
      'ID_ITEM', 'DATA_HORA', 'QTD_BAIXADA', 'QTD_RESTANTE', 'QTD_ORIGINAL', 'USUARIO'
    ]]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f0f2f5');
    sheet.setFrozenRows(1);
    Logger.log(`✅ Aba ${BAIXAS_SHEET_NAME} criada com sucesso`);
  } else {
    // Verifica se as colunas QTD_ORIGINAL e TIPO existem
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('QTD_ORIGINAL')) {
      Logger.log(`📝 Adicionando coluna QTD_ORIGINAL...`);
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue('QTD_ORIGINAL').setFontWeight('bold').setBackground('#f0f2f5');
    }
    if (!headers.includes('TIPO')) {
      Logger.log(`📝 Adicionando coluna TIPO...`);
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue('TIPO').setFontWeight('bold').setBackground('#f0f2f5');
    }
  }
  return sheet;
}

function registrarBaixa(uniqueId, qtdBaixada, qtdRestante, usuarioHtml) {
  try {
    Logger.log(`📦 Registrando baixa para ID: "${uniqueId}"`);
    const sheet = _getBaixasSheet_();
    const now = new Date();
    const usuario = (usuarioHtml && String(usuarioHtml).trim()) ? String(usuarioHtml).trim() : (Session.getActiveUser().getEmail() || 'Sistema');

    const numCols = sheet.getLastColumn();

    // LÊ O CABEÇALHO para saber a ordem das colunas
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    Logger.log(`   Cabeçalho: ${headers.join(', ')}`);

    // Mapeia índices
    const colMap = {};
    headers.forEach((h, i) => {
      colMap[String(h).trim()] = i;
    });

    // Cria array na ORDEM DO CABEÇALHO
    const novaLinha = new Array(numCols).fill('');
    novaLinha[colMap['ID_ITEM']]      = uniqueId;
    novaLinha[colMap['DATA_HORA']]    = now;
    novaLinha[colMap['QTD_BAIXADA']]  = qtdBaixada;
    novaLinha[colMap['QTD_RESTANTE']] = qtdRestante;
    novaLinha[colMap['QTD_ORIGINAL']] = qtdRestante + qtdBaixada; // qty antes desta baixa
    novaLinha[colMap['USUARIO']]      = usuario;

    Logger.log(`   Salvando: [${novaLinha.join(', ')}]`);

    sheet.appendRow(novaLinha);
    SpreadsheetApp.flush();
    Logger.log(`✅ Baixa registrada na linha ${sheet.getLastRow()}`);

    _qtdOriginalCache_ = null;
    _saldoEfetivoCache_ = null;
    _ultimaQtdOriginalCache_ = null;

    return { success: true, timestamp: now.toISOString() };
  } catch (e) {
    Logger.log(`❌ Erro ao registrar baixa: ${e.message}`);
    Logger.log(`   Stack: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

// Registra um checkpoint no histórico de baixas quando um item é faturado parcialmente.
// O checkpoint redefine a base de QTD_ORIGINAL para os próximos cálculos de saldo.
function _registrarCheckpointFaturamento_(uniqueId, qtdAberta) {
  try {
    if (!uniqueId || qtdAberta <= 0) return;
    Logger.log(`🏁 Registrando checkpoint de faturamento para ID: "${uniqueId}" | QTD base: ${qtdAberta}`);

    const sheet = _getBaixasSheet_();
    const numCols = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

    if (colMap['TIPO'] === undefined) {
      Logger.log(`⚠️ Coluna TIPO não encontrada no Baixas_Historico, checkpoint não registrado`);
      return;
    }

    const novaLinha = new Array(numCols).fill('');
    novaLinha[colMap['ID_ITEM']]      = uniqueId;
    novaLinha[colMap['DATA_HORA']]    = new Date();
    novaLinha[colMap['QTD_BAIXADA']]  = 0;
    novaLinha[colMap['QTD_RESTANTE']] = qtdAberta;
    novaLinha[colMap['QTD_ORIGINAL']] = qtdAberta; // nova base para cálculo de saldo
    novaLinha[colMap['USUARIO']]      = 'FATURAMENTO';
    novaLinha[colMap['TIPO']]         = 'CHECKPOINT';

    sheet.appendRow(novaLinha);
    SpreadsheetApp.flush();
    _qtdOriginalCache_ = null;
    _saldoEfetivoCache_ = null;
    _ultimaQtdOriginalCache_ = null;
    Logger.log(`✅ Checkpoint registrado para "${uniqueId}": nova base = ${qtdAberta}`);
  } catch (e) {
    Logger.log(`❌ _registrarCheckpointFaturamento_: ${e.message}`);
  }
}

// Constrói (ou retorna do cache) um mapa {uniqueId → soma de QTD_BAIXADA desde o último CHECKPOINT}.
// Usado por calcularQtdOriginal para derivar o valor "original do ciclo atual" a partir do DB real.
function _getSaldoEfetivoCache_() {
  if (_saldoEfetivoCache_) return _saldoEfetivoCache_;
  try {
    const sheet = _getBaixasSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { _saldoEfetivoCache_ = {}; return {}; }

    const numCols = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

    const idIdx         = colMap['ID_ITEM'];
    const qtdBaixadaIdx = colMap['QTD_BAIXADA'];
    const tipoIdx       = colMap['TIPO'];
    if (idIdx === undefined || qtdBaixadaIdx === undefined) { _saldoEfetivoCache_ = {}; return {}; }

    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    const byId = {};
    data.forEach(row => {
      const id = String(row[idIdx] || '').trim();
      if (!id) return;
      if (!byId[id]) byId[id] = [];
      byId[id].push(row);
    });

    const cache = {};
    Object.keys(byId).forEach(id => {
      const entries = byId[id];
      let checkpointPos = -1;
      if (tipoIdx !== undefined) {
        for (let i = entries.length - 1; i >= 0; i--) {
          if (String(entries[i][tipoIdx] || '').trim() === 'CHECKPOINT') {
            checkpointPos = i;
            break;
          }
        }
      }
      let soma = 0;
      for (let i = checkpointPos + 1; i < entries.length; i++) {
        const tipo = tipoIdx !== undefined ? String(entries[i][tipoIdx] || '').trim() : '';
        if (tipo === 'CHECKPOINT') continue;
        soma += _toNumber_(entries[i][qtdBaixadaIdx]);
      }
      cache[id] = soma;
    });

    _saldoEfetivoCache_ = cache;
    return cache;
  } catch (e) {
    Logger.log(`⚠️ _getSaldoEfetivoCache_: ${e.message}`);
    _saldoEfetivoCache_ = {};
    return {};
  }
}

// Constrói (ou retorna do cache) um mapa {uniqueId → QTD_ORIGINAL da entrada mais recente
// não-CHECKPOINT após o último CHECKPOINT}. Representa o saldo real antes da última baixa
// do ciclo atual — imune a entradas antigas de testes, pois usa apenas a entrada mais recente.
function _getUltimaQtdOriginalCache_() {
  if (_ultimaQtdOriginalCache_) return _ultimaQtdOriginalCache_;
  try {
    const sheet = _getBaixasSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { _ultimaQtdOriginalCache_ = {}; return {}; }

    const numCols = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

    const idIdx          = colMap['ID_ITEM'];
    const qtdOriginalIdx = colMap['QTD_ORIGINAL'];
    const tipoIdx        = colMap['TIPO'];
    if (idIdx === undefined || qtdOriginalIdx === undefined) {
      _ultimaQtdOriginalCache_ = {}; return {};
    }

    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    // Agrupa por ID mantendo a ordem de chegada
    const byId = {};
    data.forEach(row => {
      const id = String(row[idIdx] || '').trim();
      if (!id) return;
      if (!byId[id]) byId[id] = [];
      byId[id].push(row);
    });

    const cache = {};
    Object.keys(byId).forEach(id => {
      const entries = byId[id];
      // Encontra a posição do último CHECKPOINT
      let checkpointPos = -1;
      if (tipoIdx !== undefined) {
        for (let i = entries.length - 1; i >= 0; i--) {
          if (String(entries[i][tipoIdx] || '').trim() === 'CHECKPOINT') {
            checkpointPos = i;
            break;
          }
        }
      }
      // Busca a entrada mais recente não-CHECKPOINT após o último CHECKPOINT
      for (let i = entries.length - 1; i > checkpointPos; i--) {
        const tipo = tipoIdx !== undefined ? String(entries[i][tipoIdx] || '').trim() : '';
        if (tipo === 'CHECKPOINT') continue;
        const val = entries[i][qtdOriginalIdx];
        if (val !== undefined && val !== '') {
          cache[id] = _toNumber_(val);
          break;
        }
      }
    });

    _ultimaQtdOriginalCache_ = cache;
    return cache;
  } catch (e) {
    Logger.log(`⚠️ _getUltimaQtdOriginalCache_: ${e.message}`);
    _ultimaQtdOriginalCache_ = {};
    return {};
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
        // Ignora entradas internas de checkpoint — não são baixas reais
        const tipoEntry = colMap['TIPO'] !== undefined ? String(row[colMap['TIPO']] || '').trim() : '';
        if (tipoEntry === 'CHECKPOINT') continue;

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

    // Lê QTD. ABERTA atual da planilha (fresh) para garantir que o modal mostre valor correto
    let qtdAbertaAtual = null;
    try {
      const ssLive = SpreadsheetApp.openById("1qPJ8c7cq7qb86VJJ-iByeiaPnALOBcDPrPMeL75N2EI");
      const dbSheet = ssLive.getSheetByName(DB_SHEET_NAME);
      if (dbSheet) {
        const dbHeaders = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0];
        const dbColMap = _getColumnIndexes_(dbHeaders);
        const idCol = dbColMap['ID_UNICO'];
        const qtdCol = dbColMap['QTD. ABERTA'];
        if (idCol !== undefined && qtdCol !== undefined) {
          const lastDbRow = dbSheet.getLastRow();
          if (lastDbRow >= 2) {
            const dbData = dbSheet.getRange(2, 1, lastDbRow - 1, dbSheet.getLastColumn()).getValues();
            const row = dbData.find(r => String(r[idCol]).trim() === idBusca);
            if (row) qtdAbertaAtual = _toNumber_(row[qtdCol]);
          }
        }
      }
    } catch (dbErr) {
      Logger.log(`⚠️ Não foi possível ler QTD. ABERTA: ${dbErr.message}`);
    }

    const resultado = {
      success: true,
      historico: historico,
      qtdAbertaAtual: qtdAbertaAtual
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

function editarUltimaBaixa(uniqueId, planilhaLinha, novaQtdBaixada, usuarioHtml) {
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
    if (usuarioHtml && colMap['USUARIO'] !== undefined) {
      sheet.getRange(ultimaLinha, colMap['USUARIO'] + 1).setValue(String(usuarioHtml).trim());
    }

    // Atualiza a QTD. ABERTA na planilha Relatorio_DB
    const dbSheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
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
    _saldoEfetivoCache_ = null;
    _ultimaQtdOriginalCache_ = null;
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

// Retorna apenas as baixas da sessão atual (após o último CHECKPOINT) para exibição no modal.
// Inclui linhaHistorico (linha real na planilha) para permitir estorno preciso por linha.
function obterHistoricoSessaoBaixas(uniqueId) {
  try {
    if (!uniqueId) return { success: true, historico: [] };

    const sheet = _getBaixasSheet_();
    if (!sheet || sheet.getLastRow() < 2) return { success: true, historico: [] };

    const numCols = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

    const idBusca = String(uniqueId).trim();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();

    // Coleta todas as entradas deste item com o número real da linha na planilha
    const entradasItem = [];
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][colMap['ID_ITEM']] || '').trim() !== idBusca) continue;
      entradasItem.push({ rowNum: i + 2, row: data[i] });
    }

    // Encontra a posição do último CHECKPOINT dentro das entradas do item
    let lastCpIdx = -1;
    for (let i = entradasItem.length - 1; i >= 0; i--) {
      const tipo = colMap['TIPO'] !== undefined ? String(entradasItem[i].row[colMap['TIPO']] || '').trim() : '';
      if (tipo === 'CHECKPOINT') { lastCpIdx = i; break; }
    }

    // Sessão atual = entradas após o último CHECKPOINT, excluindo CHECKPOINTs
    const historico = entradasItem
      .slice(lastCpIdx + 1)
      .filter(e => {
        const tipo = colMap['TIPO'] !== undefined ? String(e.row[colMap['TIPO']] || '').trim() : '';
        return tipo !== 'CHECKPOINT';
      })
      .map(e => {
        const row = e.row;
        const dataHora = row[colMap['DATA_HORA']];
        return {
          linhaHistorico: e.rowNum,
          dataHora: dataHora ? _fmtBRDateTime_(dataHora) : '',
          qtdBaixada: _toNumber_(row[colMap['QTD_BAIXADA']]),
          qtdRestante: _toNumber_(row[colMap['QTD_RESTANTE']]),
          usuario: colMap['USUARIO'] !== undefined ? String(row[colMap['USUARIO']] || 'Sistema') : 'Sistema'
        };
      });

    return { success: true, historico };
  } catch (e) {
    Logger.log(`❌ obterHistoricoSessaoBaixas: ${e.message}`);
    return { success: false, error: String(e.message), historico: [] };
  }
}

// Remove uma baixa específica do Baixas_Historico e restaura a QTD ao Relatorio_DB.
// linhaHistorico: linha real na planilha (1-based). qtdEstornada: quantidade a restaurar.
function estornarBaixa(uniqueId, planilhaLinha, linhaHistorico, qtdEstornada) {
  try {
    Logger.log(`↩️ estornarBaixa: ID="${uniqueId}" hist_linha=${linhaHistorico} qtd=${qtdEstornada} db_linha=${planilhaLinha}`);

    const sheet = _getBaixasSheet_();
    const numCols = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

    // Valida que a linha pertence ao item correto antes de apagar
    if (linhaHistorico < 2 || linhaHistorico > sheet.getLastRow()) {
      throw new Error(`Linha ${linhaHistorico} fora do intervalo válido`);
    }
    const rowData = sheet.getRange(linhaHistorico, 1, 1, numCols).getValues()[0];
    const idLinha = String(rowData[colMap['ID_ITEM']] || '').trim();
    if (idLinha !== String(uniqueId).trim()) {
      throw new Error(`Linha ${linhaHistorico} pertence a "${idLinha}", não a "${uniqueId}"`);
    }

    // Remove a linha do histórico
    sheet.deleteRow(linhaHistorico);
    Logger.log(`   ✅ Linha ${linhaHistorico} removida do Baixas_Historico`);

    // Restaura QTD no Relatorio_DB lendo fresh para evitar cache
    const ssLive = SpreadsheetApp.openById(getSpreadsheet_().getId());
    const dbSheet = ssLive.getSheetByName(DB_SHEET_NAME);
    if (!dbSheet) throw new Error('Aba Relatorio_DB não encontrada');

    const dbHeaders = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0];
    const dbColMap = _getColumnIndexes_(dbHeaders);
    const qtdCol = dbColMap['QTD. ABERTA'];
    if (qtdCol === undefined) throw new Error('Coluna QTD. ABERTA não encontrada no DB');

    const linhaNum = Number(planilhaLinha);
    const qtdAtual = Number(dbSheet.getRange(linhaNum, qtdCol + 1).getValue() || 0);
    const novaQtd  = qtdAtual + Number(qtdEstornada);
    dbSheet.getRange(linhaNum, qtdCol + 1).setValue(novaQtd);

    SpreadsheetApp.flush();
    _qtdOriginalCache_ = null;
    _saldoEfetivoCache_ = null;
    _ultimaQtdOriginalCache_ = null;
    limparCache();

    Logger.log(`   ✅ QTD restaurada: ${qtdAtual} → ${novaQtd} (DB linha ${linhaNum})`);
    return { success: true, novaQtd, qtdAnterior: qtdAtual };

  } catch (e) {
    Logger.log(`❌ estornarBaixa: ${e.message}`);
    return { success: false, error: e.message };
  }
}

function aplicarBaixa(uniqueId, planilhaLinha, qtdBaixa, usuarioHtml) {
  try {
    // Abre fresh para evitar leitura em cache de container do Apps Script
    const ssLive = SpreadsheetApp.openById("1qPJ8c7cq7qb86VJJ-iByeiaPnALOBcDPrPMeL75N2EI");
    const sheet = ssLive.getSheetByName(DB_SHEET_NAME);
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
    const resultHistorico = registrarBaixa(uniqueId, qtdBaixa, novaQtd, usuarioHtml);

    // Baixa não altera status — Faturado é definido apenas pela sincronização
    // quando o item desaparece do PEDIDOS/fonte (sincronizarDados).

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
let _saldoEfetivoCache_ = null;
let _ultimaQtdOriginalCache_ = null;

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
    const tipoIdx = colMap['TIPO'];

    // Primeira passagem: registra o QTD_ORIGINAL da primeira entrada para cada ID
    data.forEach(row => {
      const id = String(row[colMap['ID_ITEM']] || '').trim();
      if (!id) return;
      const qtdOriginal = row[colMap['QTD_ORIGINAL']];
      if (!cache[id] && qtdOriginal !== undefined && qtdOriginal !== '') {
        cache[id] = _toNumber_(qtdOriginal);
      }
    });

    // Segunda passagem: sobrescreve com o QTD_ORIGINAL do último checkpoint de faturamento
    if (tipoIdx !== undefined) {
      data.forEach(row => {
        const id = String(row[colMap['ID_ITEM']] || '').trim();
        if (!id) return;
        if (String(row[tipoIdx] || '').trim() !== 'CHECKPOINT') return;
        const qtdOriginal = row[colMap['QTD_ORIGINAL']];
        if (qtdOriginal !== undefined && qtdOriginal !== '') {
          cache[id] = _toNumber_(qtdOriginal); // sobrescreve: última ocorrência vence
        }
      });
    }

    Logger.log(`📦 Cache de quantidades construído: ${Object.keys(cache).length} itens`);
    return cache;
  } catch (e) {
    Logger.log(`⚠️ Erro ao construir cache: ${e.message}`);
    return {};
  }
}

// Retorna o "Saldo Aberto" do ciclo atual: QTD_ORIGINAL da entrada mais recente
// não-CHECKPOINT após o último CHECKPOINT. Reflete o valor do DB imediatamente
// antes da última baixa manual, imune a entradas antigas de testes que não têm
// CHECKPOINT separando-as.
function calcularQtdOriginal(uniqueId, qtdAbertaAtual) {
  try {
    const cache = _getUltimaQtdOriginalCache_();
    return cache[uniqueId] !== undefined ? cache[uniqueId] : qtdAbertaAtual;
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
    .addSeparator()
    .addItem('⏸ Pausar sistema (para editar dados)', 'pausarSistema')
    .addItem('▶ Retomar sistema', 'retomarSistema')
    .addSeparator()
    .addItem('🧹 Confirmar todos os alertas de faturamento (testes)', 'confirmarTodosAlertasMenu')
    .addItem('🔧 Corrigir Faturados com saldo aberto (reverter para Ativo)', 'corrigirFaturadosComSaldoAberto')
    .addSeparator()
    .addItem('⚠️ RESET COMPLETO (apaga DB + regenera IDs)', 'resetarEReprocessar')
    .addSeparator()
    .addItem('🔄 Forçar resync agora (aplica correções Dilly)', 'forcarResyncDilly')
    .addToUi();
}

/**
 * Força uma sincronização completa ignorando o guard de timestamp.
 * Use após aplicar atualizações de código que alteram campos calculados
 * (ex: correção do CÓD. MARFIM Dilly, preenchimento de OS via LOTE DILLY).
 */
function forcarResyncDilly() {
  const ui = SpreadsheetApp.getUi();
  try {
    Logger.log('⚡ Forçando resync completo (forcarExecucao=true)...');
    const resultado = sincronizarPedidosComFonte(true);
    if (resultado.erro) {
      ui.alert('Erro no resync', resultado.erro, ui.ButtonSet.OK);
      return;
    }
    // Propaga imediatamente para o Relatorio_DB
    processoAutomaticoCompleto();
    ui.alert(
      '✅ Resync concluído',
      `PEDIDOS sincronizado.\nNovos: ${resultado.novos || 0} | Atualizados: ${resultado.atualizados || 0} | Total: ${resultado.total || 0}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Reset completo: limpa o Relatorio_DB, limpa os IDs antigos da aba PEDIDOS
 * e roda o processo completo para gerar tudo do zero com a fórmula atual de IDs.
 */
function resetarEReprocessar() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '⚠️ RESET COMPLETO',
    'Isso irá:\n\n' +
    '1. Apagar TODOS os dados do Relatorio_DB\n' +
    '2. Apagar os IDs antigos da aba PEDIDOS\n' +
    '3. Regenerar todos os IDs com a nova fórmula\n' +
    '4. Repopular o Relatorio_DB do zero\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    Logger.log('Reset cancelado pelo usuário.');
    return;
  }

  Logger.log('=== RESET COMPLETO INICIADO ===');

  // 1. Limpa Relatorio_DB (mantém cabeçalho na linha 1)
  const dbSheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
  if (dbSheet && dbSheet.getLastRow() > 1) {
    dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, dbSheet.getLastColumn()).clearContent();
    SpreadsheetApp.flush();
    Logger.log('✅ Relatorio_DB limpo.');
  } else {
    Logger.log('ℹ️ Relatorio_DB já estava vazio.');
  }
  limparCache();

  // 2. Limpa coluna A (IDs) da aba PEDIDOS para forçar regeneração com nova fórmula
  const pedidosSheet = getSpreadsheet_().getSheetByName(FONTE_SHEET_NAME);
  if (pedidosSheet && pedidosSheet.getLastRow() >= FONTE_DATA_START_ROW) {
    const numLinhas = pedidosSheet.getLastRow() - FONTE_DATA_START_ROW + 1;
    pedidosSheet.getRange(FONTE_DATA_START_ROW, 1, numLinhas, 1).clearContent();
    SpreadsheetApp.flush();
    Logger.log('✅ IDs antigos removidos da aba PEDIDOS.');
  }

  // 3. Roda processo completo (sincroniza PEDIDOS + gera IDs + popula DB)
  Logger.log('🔄 Iniciando processo completo...');
  processoAutomaticoCompleto();

  Logger.log('=== RESET COMPLETO FINALIZADO ===');
  ui.alert('✅ Reset concluído!', 'IDs regenerados e Relatorio_DB repopulado com sucesso.', ui.ButtonSet.OK);
}

/**
 * Mesma lógica de resetarEReprocessar() mas SEM alertas de UI.
 * Execute esta função diretamente pelo Apps Script Editor (▶ Executar).
 * Acompanhe o progresso em: Ver > Registros de execução (Ctrl+Enter).
 */
function resetarEReprocessarSilencioso() {
  Logger.log('=== RESET COMPLETO INICIADO (modo silencioso) ===');

  // 1. Limpa Relatorio_DB (mantém cabeçalho na linha 1)
  const dbSheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
  if (dbSheet && dbSheet.getLastRow() > 1) {
    dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, dbSheet.getLastColumn()).clearContent();
    SpreadsheetApp.flush();
    Logger.log('✅ Relatorio_DB limpo.');
  } else {
    Logger.log('ℹ️ Relatorio_DB já estava vazio.');
  }
  limparCache();

  // 2. Limpa coluna A (IDs) da aba PEDIDOS para forçar regeneração com nova fórmula
  const pedidosSheet = getSpreadsheet_().getSheetByName(FONTE_SHEET_NAME);
  if (pedidosSheet && pedidosSheet.getLastRow() >= FONTE_DATA_START_ROW) {
    const numLinhas = pedidosSheet.getLastRow() - FONTE_DATA_START_ROW + 1;
    pedidosSheet.getRange(FONTE_DATA_START_ROW, 1, numLinhas, 1).clearContent();
    SpreadsheetApp.flush();
    Logger.log('✅ IDs antigos removidos da aba PEDIDOS.');
  }

  // 3. Roda processo completo (sincroniza PEDIDOS + gera IDs + popula DB)
  Logger.log('🔄 Iniciando processo completo...');
  processoAutomaticoCompleto();

  Logger.log('=== RESET COMPLETO FINALIZADO ===');
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

  const sheet = getSpreadsheet_().getSheetByName(FONTE_SHEET_NAME);

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

    // FIX: CARTELA (col B) e DESCRIÇÃO (col H) removidos do ID base.
    // São campos mutáveis que podem ser atualizados pelo sistema de origem.
    // Mantém apenas campos estáveis como identidade do pedido, consistente
    // com a lógica de sincronizarPedidosComFonte() e _criarImpressaoDigital_().
    const idBase = "" +
      String(linha[2] || '').trim() + // Coluna C - CLIENTE
      String(linha[3] || '').trim() + // Coluna D - CÓD. FILIAL
      String(linha[4] || '').trim() + // Coluna E - PEDIDO
      String(linha[6] || '').trim() + // Coluna G - CÓD. MARFIM
      String(linha[8] || '').trim() + // Coluna I - TAMANHO
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

    // Log apenas a cada 100 linhas para evitar timeout
    if (idsGerados % 100 === 0) {
      Logger.log(`  ✓ Processadas ${idsGerados} linhas...`);
    }
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
    const sheet = getSpreadsheet_().getSheetByName(FONTE_SHEET_NAME);
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

// ====== IMPORTAÇÃO DIRETA DA PLANILHA EXTERNA ======
/**
 * Substitui o IMPORTRANGE: lê diretamente da planilha externa via Apps Script
 * e grava os dados em DADOS_IMPORTADOS a partir da coluna B, atualizando H2.
 * Muito mais rápido e confiável que o IMPORTRANGE nativo.
 */
function importarDadosExternos() {
  const SOURCE_ID    = '1GtYG4Ahy5XJyJjE37S27u8RyELdRkct8nDAVGIBRI-w';
  const SOURCE_SHEET = 'RELATÓRIO GERAL DA PRODUÇÃO1';

  // Coluna J (índice 9): valores como "82249D" e "14660U" têm formatação customizada
  // no Sheets — getValues() retorna só o número cru (sem sufixo). Por isso usamos
  // getDisplayValues() APENAS nessa coluna (range estreito) para preservar o valor visível.
  const COL_J = 9;

  try {
    Logger.log(`📡 importarDadosExternos: abrindo planilha externa...`);
    const sourceSheet = SpreadsheetApp.openById(SOURCE_ID).getSheetByName(SOURCE_SHEET);
    if (!sourceSheet) throw new Error(`Aba "${SOURCE_SHEET}" não encontrada na planilha externa.`);

    // Usa lastRow/lastColumn dinâmico — evita ler 5000 linhas vazias
    const srcLastRow = sourceSheet.getLastRow();
    const srcLastCol = sourceSheet.getLastColumn();
    if (srcLastRow < 1) {
      Logger.log('⚠️ importarDadosExternos: nenhum dado encontrado na fonte.');
      return { success: false, erro: 'Sem dados na fonte' };
    }

    // Leitura única com getValues() — muito mais rápido que getDisplayValues() na range grande
    const rangeRef = sourceSheet.getRange(1, 1, srcLastRow, srcLastCol);
    const dados = rangeRef.getValues();

    // Lê display values SOMENTE da coluna J (range de 1 coluna × N linhas = muito menor)
    const colJDisplay = sourceSheet.getRange(1, COL_J + 1, srcLastRow, 1).getDisplayValues();

    // Remove linhas em branco do final
    let ultimaLinha = dados.length;
    while (ultimaLinha > 0 && dados[ultimaLinha - 1].every(c => c === '' || c === null || c === undefined)) {
      ultimaLinha--;
    }
    const dadosFiltrados = dados.slice(0, ultimaLinha);

    if (dadosFiltrados.length === 0) {
      Logger.log('⚠️ importarDadosExternos: nenhum dado encontrado na fonte.');
      return { success: false, erro: 'Sem dados na fonte' };
    }

    // Substitui coluna J pelo valor de display (preserva sufixos "D"/"U")
    dadosFiltrados.forEach((row, i) => {
      row[COL_J] = colJDisplay[i][0]; // colJDisplay é array de 1 coluna → índice 0
    });

    const destSheet = getSpreadsheet_().getSheetByName(IMPORTRANGE_SHEET_NAME);
    if (!destSheet) throw new Error(`Aba "${IMPORTRANGE_SHEET_NAME}" não encontrada.`);

    const numCols   = dadosFiltrados[0].length;
    const destLastRow = destSheet.getLastRow();
    const clearRows = Math.max(destLastRow, dadosFiltrados.length);

    // Limpa área anterior (garante pelo menos 25 colunas para cobrir colunas X e Y)
    if (clearRows > 0) {
      destSheet.getRange(1, 1, clearRows, Math.max(numCols, 25)).clearContent();
    }

    // Formata coluna F como texto — apenas as linhas que serão escritas (não a planilha inteira)
    destSheet.getRange(1, 6, dadosFiltrados.length, 1).setNumberFormat('@'); // F: CÓD. CLIENTE

    // Grava dados processados
    destSheet.getRange(1, 1, dadosFiltrados.length, numCols).setValues(dadosFiltrados);

    // Copia explícita das colunas X (24) e Y (25) da fonte → DADOS_IMPORTADOS
    // Necessário pois getLastColumn() pode não detectar colunas sem cabeçalho
    const srcRowsCopia = Math.min(5000, srcLastRow);
    if (srcRowsCopia >= 1) {
      const colXYData = sourceSheet.getRange(1, 24, srcRowsCopia, 2).getValues();
      destSheet.getRange(1, 24, srcRowsCopia, 2).setValues(colXYData);
      Logger.log(`✅ Colunas X e Y copiadas da fonte: ${srcRowsCopia} linhas → DADOS_IMPORTADOS cols X/Y`);
    }

    // Atualiza timestamp em H2 — detectado pelo guard em sincronizarPedidosComFonte()
    const ts = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
    destSheet.getRange('H2').setValue(ts);

    SpreadsheetApp.flush();
    Logger.log(`✅ importarDadosExternos: ${dadosFiltrados.length} linhas gravadas. Timestamp: ${ts}`);
    return { success: true, linhas: dadosFiltrados.length, timestamp: ts };

  } catch (e) {
    Logger.log(`❌ importarDadosExternos: ${e.message}`);
    return { success: false, erro: e.message };
  }
}

/**
 * SINCRONIZAÇÃO INTELIGENTE: DADOS_IMPORTADOS → PEDIDOS
 *
 * Sincroniza dados do IMPORTRANGE com aba PEDIDOS mantendo IDs estáveis.
 * - Identifica itens por "impressão digital" (CARTELA+CLIENTE+PEDIDO+etc)
 * - Preserva IDs e timestamps de itens existentes
 * - Adiciona novos itens com novos IDs
 * - Atualiza dados de itens existentes
 * - Lida com itens 100% idênticos usando ordem + timestamp
 *
 * @returns {Object} {houveMudancas: boolean, novos: number, atualizados: number, erro: string}
 */
function sincronizarPedidosComFonte(forcarExecucao) {
  const inicioSync = Date.now();
  Logger.log("=" .repeat(70));
  Logger.log(`🔄 SINCRONIZAÇÃO DADOS_IMPORTADOS → PEDIDOS`);
  Logger.log("=".repeat(70));

  try {
    // PASSO 1: Ler aba DADOS_IMPORTADOS (fonte com IMPORTRANGE)
    const fonteSheet = getSpreadsheet_().getSheetByName(IMPORTRANGE_SHEET_NAME);
    if (!fonteSheet) {
      Logger.log(`❌ Aba ${IMPORTRANGE_SHEET_NAME} não encontrada!`);
      Logger.log(`   Crie a aba e configure o IMPORTRANGE primeiro.`);
      return { houveMudancas: false, erro: 'Aba DADOS_IMPORTADOS não existe' };
    }

    // GUARDA DE DADOS: verifica B1 (primeiro dado) e H2 (timestamp da última importação).
    // B1 contém o primeiro valor importado — se iniciar com '#' indica erro na importação.
    // H2 contém o horário da última importação gravado por importarDadosExternos().
    // Se B1 tem erro, ou H2 está vazio/com erro, ou H2 não mudou → aborta.
    const a1Val = fonteSheet.getRange('B1').getDisplayValue().trim();
    if (a1Val.startsWith('#')) {
      Logger.log(`⚠️ Dado inválido em B1="${a1Val}". Sync ignorado.`);
      return { houveMudancas: false, motivo: 'importrange_erro_b1' };
    }
    const tsAtual = fonteSheet.getRange('H2').getDisplayValue().trim();
    if (!tsAtual || tsAtual.startsWith('#')) {
      Logger.log(`⚠️ IMPORTRANGE não concluído — H2="${tsAtual}". Sync ignorado.`);
      return { houveMudancas: false, motivo: 'importrange_nao_pronto' };
    }
    const props = PropertiesService.getScriptProperties();
    const tsAnterior = props.getProperty('ULTIMO_IMPORTRANGE_TS') || '';
    if (tsAtual === tsAnterior && !forcarExecucao) {
      Logger.log(`ℹ️ Dados não alterados (H2="${tsAtual}"). Sync ignorado.`);
      return { houveMudancas: false, motivo: 'sem_nova_importacao' };
    }
    if (forcarExecucao && tsAtual === tsAnterior) {
      Logger.log(`⚡ Execução forçada — ignorando guard de timestamp (H2="${tsAtual}")`);
    } else {
      Logger.log(`🆕 Novo timestamp detectado: "${tsAnterior || 'nenhum'}" → "${tsAtual}"`);
    }

    const fonteLastRow = fonteSheet.getLastRow();
    if (fonteLastRow < FONTE_DATA_START_ROW) {
      Logger.log(`⚠️ Sem dados em ${IMPORTRANGE_SHEET_NAME}`);
      return { houveMudancas: false, erro: 'Sem dados na fonte' };
    }

    // Lê dados da fonte a partir da coluna B (dados começam em B, coluna A ignorada)
    // Garante pelo menos 24 colunas (B→Y) para cobrir fonteRow[23] = col Y = LOTE
    const fonteNumCols = Math.max(fonteSheet.getLastColumn() - 1, 24);
    const fonteData = fonteSheet.getRange(FONTE_DATA_START_ROW, 2, fonteLastRow - FONTE_DATA_START_ROW + 1, fonteNumCols).getValues();
    Logger.log(`📥 Leu ${fonteData.length} linhas de ${IMPORTRANGE_SHEET_NAME}`);

    // PASSO 2: Ler aba PEDIDOS (atual com IDs)
    const pedidosSheet = getSpreadsheet_().getSheetByName(FONTE_SHEET_NAME);
    if (!pedidosSheet) {
      Logger.log(`❌ Aba ${FONTE_SHEET_NAME} não encontrada!`);
      return { houveMudancas: false, erro: 'Aba PEDIDOS não existe' };
    }

    const pedidosLastRow = pedidosSheet.getLastRow();
    let pedidosData = [];
    let pedidosMap = new Map(); // impressao_digital → {id, timestamp, row, linhaOriginal}

    if (pedidosLastRow >= FONTE_DATA_START_ROW) {
      // Lê dados atuais de PEDIDOS (com ID e timestamp)
      const pedidosNumCols = Math.max(21, pedidosSheet.getLastColumn()); // Garante até coluna U (LOTE)
      pedidosData = pedidosSheet.getRange(FONTE_DATA_START_ROW, 1, pedidosLastRow - FONTE_DATA_START_ROW + 1, pedidosNumCols).getValues();

      Logger.log(`📋 Leu ${pedidosData.length} linhas de ${FONTE_SHEET_NAME}`);

      // Cria mapa de itens existentes em PEDIDOS
      pedidosData.forEach((row, idx) => {
        const id = row[ID_COL];
        const cartela = row[CARTELA_COL];

        // Ignora linhas sem dados
        if (!cartela || String(cartela).trim() === '') return;

        // Cria impressão digital (colunas B até O em PEDIDOS = índices 1-14)
        const impressao = _criarImpressaoDigitalFromRow_(row, 1); // offset 1 porque ID está em 0
        const timestamp = row[TIMESTAMP_COL] || null;

        // Para itens com mesma impressão, guarda em array
        if (!pedidosMap.has(impressao)) {
          pedidosMap.set(impressao, []);
        }
        pedidosMap.get(impressao).push({
          id: id,
          timestamp: timestamp,
          row: row,
          codigoFixo: String(row[PEDIDOS_CODIGO_FIXO_COL] || '').trim(),
          linhaOriginal: idx + FONTE_DATA_START_ROW,
          usado: false
        });
      });

      Logger.log(`🔑 Mapeou ${pedidosMap.size} impressões digitais únicas`);
    } else {
      Logger.log(`📋 PEDIDOS está vazio (primeira sincronização)`);
    }

    // PASSO 3: Processar cada linha da fonte
    const novasPedidosData = [];
    let novosItens = 0;
    let itensAtualizados = 0;
    let itensComMudancaReal = 0;

    // Pré-carrega IDs do Relatorio_DB no Set (previne colisões) E
    // constrói mapa fingerprint→ID para RECUPERAR IDs quando PEDIDOS perde a coluna A.
    // Isso evita que itens ganhem novos IDs após o usuário apagar PEDIDOS e o IMPORTRANGE
    // repopular só as colunas B-O (sem a coluna A que é gerenciada por script).
    const idsUsados = new Set();
    const dbFingerprintMap = new Map(); // fingerprint → [id, ...] (array FIFO — suporta itens 100% idênticos)
    const dbCodigoFixoMap  = new Map(); // id → codigoFixo (reutilizar UUID já gravado no DB)
    const dbSheetRef = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
    if (dbSheetRef && dbSheetRef.getLastRow() >= 2) {
      const numDbRows = dbSheetRef.getLastRow() - 1;
      // Lê 19 colunas (A até S) — inclui CÓDIGO_FIXO na coluna S (índice 18)
      const dbRange = dbSheetRef.getRange(2, 1, numDbRows, 19).getValues();
      dbRange.forEach(dbRow => {
        const dbId = String(dbRow[0] || '').trim();
        if (dbId) {
          idsUsados.add(dbId);
          const fp = _criarImpressaoDigital_(dbRow, true);
          if (fp) {
            // Array FIFO por fingerprint — para itens 100% idênticos cada um tem seu próprio slot.
            // shift() consome um ID por vez, garantindo que cada item da fonte recupere um ID distinto.
            if (!dbFingerprintMap.has(fp)) dbFingerprintMap.set(fp, []);
            dbFingerprintMap.get(fp).push(dbId);
          }
          const cf = String(dbRow[DB_CODIGO_FIXO_COL] || '').trim();
          if (cf) dbCodigoFixoMap.set(dbId, cf); // UUID fixo já gravado no DB para este item
        }
      });
      Logger.log(`🔒 ${idsUsados.size} IDs do Relatorio_DB carregados (colisões + recuperação)`);
      Logger.log(`🔑 ${dbFingerprintMap.size} fingerprints do DB indexadas para recuperação de ID`);
    }

    // PASSO 2.5: Ler aba "original" para determinar a sequência correta de itens dentro de cada OC.
    // Chave: "OC|DESC|TAM|QTD|DATA" → índice global da linha (usado para ordenar itens no HTML).
    // Regra: uma vez encontrada, a posição NUNCA muda — mesmo que QTD seja alterada depois.
    const originalPosMap = new Map();
    try {
      const origSheet = getSpreadsheet_().getSheetByName(ORIGINAL_SHEET_NAME);
      if (origSheet && origSheet.getLastRow() > 1) {
        const origLastRow = origSheet.getLastRow() - 1; // exclui cabeçalho
        const origData = origSheet.getRange(2, 1, origLastRow, 12).getValues();
        origData.forEach((row, origIdx) => {
          const oc = String(row[ORIG_OC_COL] || '').trim();
          if (!oc) return;
          const desc = String(row[ORIG_DESC_COL] || '').trim();
          const tam  = String(row[ORIG_TAM_COL]  || '').trim();
          const qtd  = String(row[ORIG_QTD_COL]  || '').trim();
          const data = row[ORIG_DATA_COL] instanceof Date
            ? Utilities.formatDate(row[ORIG_DATA_COL], TZ, 'yyyy-MM-dd')
            : String(row[ORIG_DATA_COL] || '').trim();
          const chave = `${oc}|${desc}|${tam}|${qtd}|${data}`;
          if (!originalPosMap.has(chave)) { // guarda apenas a primeira ocorrência
            originalPosMap.set(chave, origIdx);
          }
        });
        Logger.log(`📐 ${originalPosMap.size} posições mapeadas da aba "${ORIGINAL_SHEET_NAME}"`);
      } else {
        Logger.log(`⚠️ Aba "${ORIGINAL_SHEET_NAME}" não encontrada ou vazia — usando índice de DADOS_IMPORTADOS como fallback`);
      }
    } catch (e) {
      Logger.log(`⚠️ Erro ao ler aba "${ORIGINAL_SHEET_NAME}": ${e.message} — usando fallback`);
    }

    // PASSO 2.6: Ler aba LOTE DILLY e montar fila FIFO de Lotes por chave OC|Código|Tamanho|Qtd.
    // Para cada item Dilly sem correspondência exata, o primeiro Lote disponível na fila é
    // consumido (shift) na ordem em que os itens aparecem em DADOS_IMPORTADOS — garantindo
    // associação estável mesmo quando OC + Código + Tamanho + Qtd não são únicos.
    const loteDillyMap = new Map(); // "OC|Código|Tam|Qtd" → [Lote, Lote, ...] (FIFO)
    try {
      const loteDillySheet = getSpreadsheet_().getSheetByName(LOTE_DILLY_SHEET_NAME);
      if (loteDillySheet && loteDillySheet.getLastRow() > 1) {
        const ldData = loteDillySheet.getRange(2, 1, loteDillySheet.getLastRow() - 1, 7).getValues();
        ldData.forEach(ldRow => {
          const ldOc   = String(ldRow[0] || '').trim(); // A: OC
          const ldLote = String(ldRow[1] || '').trim(); // B: Lote → vai para CÓD. OS
          const ldCod  = String(ldRow[2] || '').trim(); // C: Código do item
          const ldTam  = String(ldRow[4] || '').trim().replace(/[^0-9]/g, ''); // E: Tamanho (numérico)
          const ldQtd  = String(ldRow[6] || '').trim(); // G: Qtd
          if (!ldOc || !ldLote || !ldCod) return;
          const ldChave = `${ldOc}|${ldCod}|${ldTam}|${ldQtd}`;
          if (!loteDillyMap.has(ldChave)) loteDillyMap.set(ldChave, []);
          loteDillyMap.get(ldChave).push(ldLote);
        });
        Logger.log(`📦 LOTE DILLY: ${ldData.length} linhas → ${loteDillyMap.size} chaves indexadas`);
      } else {
        Logger.log(`ℹ️ Aba "${LOTE_DILLY_SHEET_NAME}" vazia ou inexistente — sem enriquecimento de OS para Dilly`);
      }
    } catch (e) {
      Logger.log(`⚠️ Erro ao ler "${LOTE_DILLY_SHEET_NAME}": ${e.message}`);
    }

    fonteData.forEach((fonteRow, idx) => {
      const cartela = fonteRow[0]; // Em DADOS_IMPORTADOS, CARTELA é coluna B (índice 0, lido a partir de B)

      // Ignora linhas vazias
      if (!cartela || String(cartela).trim() === '') {
        return;
      }

      // Cria impressão digital da linha fonte (offset 0 porque não tem coluna ID)
      const impressao = _criarImpressaoDigitalFromRow_(fonteRow, 0);

      // Procura match em PEDIDOS
      const matches = pedidosMap.get(impressao);

      let idFinal = null;
      let timestampFinal = null;
      let isNovo = false;
      let codigoFixo = ''; // UUID fixo por item — gerado uma vez, preservado para sempre
      let matchEscolhido = null; // declarado no escopo externo para uso na resolução de posicaoFonte

      if (matches && matches.length > 0) {
        // TEM MATCH(ES) em PEDIDOS - Reusar ID existente

        // Encontra match não usado.
        // Se há múltiplos candidatos com a mesma fingerprint (mesmo produto, QTDs diferentes),
        // usa QTD como critério de desempate para garantir matching estável mesmo quando o
        // IMPORTRANGE reordena as linhas — evita troca de IDs entre itens duplicados.
        const unusedMatches = matches.filter(m => !m.usado);
        if (unusedMatches.length > 1) {
          const fonteQtd = Number(fonteRow[9] || 0); // QTD em DADOS_IMPORTADOS (índice 9)
          matchEscolhido = unusedMatches.reduce((best, m) => {
            const diffBest = Math.abs(Number(best.row[QTD_COL] || 0) - fonteQtd);
            const diffCurr = Math.abs(Number(m.row[QTD_COL] || 0) - fonteQtd);
            return diffCurr < diffBest ? m : best;
          }, unusedMatches[0]);
          Logger.log(`   🔀 Desempate por QTD: fonte=${fonteQtd} → PEDIDOS QTD=${Number(matchEscolhido.row[QTD_COL]||0)} (${unusedMatches.length} candidatos com mesma fingerprint)`);
        } else {
          matchEscolhido = unusedMatches[0] || null;
        }

        if (!matchEscolhido) {
          // Todos os matches já foram usados (mais itens na fonte que em PEDIDOS)
          isNovo = true;
        } else {
          matchEscolhido.usado = true;
          codigoFixo = matchEscolhido.codigoFixo || ''; // reutiliza UUID já existente em PEDIDOS

          // Se a coluna A do PEDIDOS estava vazia (usuário apagou o PEDIDOS e o IMPORTRANGE
          // trouxe de volta só os dados B-O), tenta recuperar o ID do Relatorio_DB antes de
          // gerar um novo — evita que itens existentes ganhem novos IDs.
          const idExistente = String(matchEscolhido.id || '').trim();
          if (!idExistente) {
            const _fpList1_ = dbFingerprintMap.get(impressao);
            const idRecuperado = (_fpList1_ && _fpList1_.length > 0) ? _fpList1_.shift() : null;
            if (idRecuperado) {
              Logger.log(`   🔄 ID recuperado do DB para item sem ID em PEDIDOS: "${idRecuperado}"`);
              idFinal = idRecuperado;
              timestampFinal = matchEscolhido.timestamp || new Date();
            } else {
              isNovo = true; // nunca esteve no DB: gerar novo ID normalmente
            }
          } else {
            idFinal = idExistente;
            timestampFinal = matchEscolhido.timestamp;
          }

          if (!isNovo) {
            itensAtualizados++;
            // Detecta mudança real nos dados
            for (let fi = 0; fi <= 13; fi++) {
              const fv = fonteRow[fi];
              const pv = matchEscolhido.row[fi + 1];
              const fStr = fv instanceof Date ? _toISOStringSafe_(fv) : String(fv || '');
              const pStr = pv instanceof Date ? _toISOStringSafe_(pv) : String(pv || '');
              if (fStr !== pStr) { itensComMudancaReal++; break; }
            }
          }
        }
      } else {
        // NÃO TEM MATCH em PEDIDOS - pode ser item novo ou PEDIDOS estava vazio
        // Tenta recuperar ID do DB pela fingerprint antes de gerar novo.
        // shift() consome o primeiro slot disponível — cada item idêntico pega seu próprio ID.
        const _fpList2_ = dbFingerprintMap.get(impressao);
        const idRecuperado = (_fpList2_ && _fpList2_.length > 0) ? _fpList2_.shift() : null;
        if (idRecuperado) {
          Logger.log(`   🔄 ID recuperado do DB (sem match em PEDIDOS): "${idRecuperado}"`);
          idFinal = idRecuperado;
          timestampFinal = new Date();
          itensAtualizados++;
        } else {
          isNovo = true;
        }
      }

      // Se é novo item, gera ID e timestamp
      if (isNovo) {
        // Gera ID usando lógica existente (concatenação + sufixo)
        const dataReceb = fonteRow[11]; // Coluna L em DADOS_IMPORTADOS = DATA RECEB. (índice 11)
        const dataFormatada = dataReceb instanceof Date ?
          Utilities.formatDate(dataReceb, TZ, 'yyyyMMdd') :
          String(dataReceb || '').trim();

        // FIX: CARTELA (fonteRow[0]) e DESCRIÇÃO (fonteRow[6]) removidos do ID base.
        // Ambos são campos mutáveis - podem ser atualizados pelo sistema de origem.
        // O ID usa apenas campos estáveis que identificam o pedido de forma permanente.
        const idBase = "" +
          String(fonteRow[1] || '').trim() +  // CLIENTE
          String(fonteRow[2] || '').trim() +  // CÓD. FILIAL
          String(fonteRow[3] || '').trim() +  // PEDIDO
          String(fonteRow[5] || '').trim() +  // CÓD. MARFIM
          String(fonteRow[7] || '').trim() +  // TAMANHO
          String(fonteRow[8] || '').trim() +  // ORD. COMPRA
          String(fonteRow[10] || '').trim() + // CÓD. OS
          dataFormatada;                       // DATA RECEBIMENTO (col M)

        // Gera sufixo único
        let sufixo = 1;
        while (idsUsados.has(idBase + "-" + sufixo)) {
          sufixo++;
        }

        idFinal = idBase + "-" + sufixo;
        timestampFinal = new Date();
        novosItens++;
      }

      idsUsados.add(idFinal);

      // Resolve CÓDIGO_FIXO: reutiliza o que já existe (PEDIDOS ou DB), senão gera novo UUID
      if (!codigoFixo) {
        codigoFixo = dbCodigoFixoMap.get(idFinal) || Utilities.getUuid();
      }

      // Resolve POSICAO_FONTE: procura na aba "original" pela chave OC|DESC|TAM|QTD|DATA.
      // Uma vez encontrada, a posição é preservada para sempre (matchEscolhido já tem o valor).
      // Nunca sobrescreve uma posição válida já gravada em PEDIDOS (< 500000).
      let posicaoFonte = null;
      if (matchEscolhido && !isNovo) {
        const posExistente = matchEscolhido.row[PEDIDOS_POSICAO_FONTE_COL];
        if (typeof posExistente === 'number' && posExistente < 500000) {
          posicaoFonte = posExistente; // já encontrado antes — preserva
        }
      }
      if (posicaoFonte === null) {
        // Primeira vez: procura na aba "original"
        const oc   = String(fonteRow[8]  || '').trim(); // J: ORD. COMPRA
        const desc = String(fonteRow[6]  || '').trim(); // H: DESCRIÇÃO
        const tam  = String(fonteRow[7]  || '').trim(); // I: TAMANHO
        const qtd  = String(fonteRow[9]  || '').trim(); // K: QTD. ABERTA
        const data = fonteRow[11] instanceof Date
          ? Utilities.formatDate(fonteRow[11], TZ, 'yyyy-MM-dd')
          : String(fonteRow[11] || '').trim();          // M: DATA RECEB.
        const chave = `${oc}|${desc}|${tam}|${qtd}|${data}`;
        const posOriginal = originalPosMap.get(chave);
        posicaoFonte = (posOriginal !== undefined) ? posOriginal : (900000 + idx);
        // 900000+idx como fallback: mantém ordem relativa de DADOS_IMPORTADOS para itens não encontrados
      }

      // Para o cliente Dilly: corrige CÓD. MARFIM (col G) — veja _normalizarMarfimDilly_
      const _codMarfimFinal_ = _normalizarMarfimDilly_(
        String(fonteRow[5] || '').trim(),
        String(fonteRow[7] || '').trim(),
        String(fonteRow[1] || '').trim()
      );

      // Para o cliente Dilly: corrige CÓD. CLIENTE (col F / DB col E) — mesmo padrão.
      // É neste campo que aparece "196338-120"; CÓD. MARFIM costuma ser número simples ("70893").
      const _clienteStr_    = String(fonteRow[1] || '').trim();
      const _codClienteRaw_ = String(fonteRow[4] !== null && fonteRow[4] !== undefined ? fonteRow[4] : '');
      const _codClienteFinal_ = _normalizarMarfimDilly_(
        _codClienteRaw_,
        String(fonteRow[7] || '').trim(),
        _clienteStr_
      );

      // Para o cliente Dilly: determina CÓD. OS via aba LOTE DILLY (sobrescreve DADOS_IMPORTADOS).
      // Chave: OC | código-base do CÓD. CLIENTE (prefixo antes do último "-") | tamanho numérico | qtd
      // FIFO: cada item consome o próximo Lote disponível na fila para essa chave.
      let _osFinal_ = fonteRow[10]; // padrão: CÓD. OS de DADOS_IMPORTADOS (col L)
      const _ehDilly_ = _clienteStr_.toUpperCase().includes('DILLY');
      if (_ehDilly_ && loteDillyMap.size > 0) {
        // Usa o CÓD. CLIENTE corrigido como base da chave (ex: "196338-110" → "196338")
        const _codBase_ = _codClienteFinal_.includes('-')
          ? _codClienteFinal_.substring(0, _codClienteFinal_.lastIndexOf('-'))
          : _codClienteFinal_;
        const _tamNum_  = String(fonteRow[7] || '').trim().replace(/[^0-9]/g, '');
        const _ocStr_   = String(fonteRow[8] || '').trim();
        const _qtdStr_  = String(fonteRow[9] || '').trim();
        const _ldChave_ = `${_ocStr_}|${_codBase_}|${_tamNum_}|${_qtdStr_}`;
        const _ldFila_  = loteDillyMap.get(_ldChave_);
        if (_ldFila_ && _ldFila_.length > 0) {
          _osFinal_ = _ldFila_.shift(); // consome o próximo lote disponível
        }
      }

      // Monta linha completa para PEDIDOS (A até T)
      const novaLinha = [
        idFinal,           // A: ID_UNICO
        fonteRow[0],       // B: CARTELA
        fonteRow[1],       // C: CLIENTE
        fonteRow[2],       // D: CÓD. FILIAL
        fonteRow[3],       // E: PEDIDO
        _codClienteFinal_, // F: CÓD. CLIENTE (Dilly: corrigido com numérico do TAMANHO; guard ≥2 chars protege "7490-1")
        _codMarfimFinal_,  // G: CÓD. MARFIM (Dilly: corrigido se tiver traço com sufixo ≥2 chars)
        String(fonteRow[6] || '') + (codigoFixo ? ' [' + codigoFixo + ']' : ''),  // H: DESCRIÇÃO [UUID] — âncora de identidade visível na planilha
        fonteRow[7],       // I: TAMANHO
        fonteRow[8],       // J: ORD. COMPRA
        fonteRow[9],       // K: QTD. ABERTA
        _osFinal_,         // L: CÓD. OS (Dilly: valor do Lote via LOTE DILLY; demais: DADOS_IMPORTADOS)
        (fonteRow[11] instanceof Date ? fonteRow[11] : (!fonteRow[11] || fonteRow[11] === 0 ? '' : fonteRow[11])), // M: DATA RECEB. — 0 vira '' (evita 30/12/1899)
        fonteRow[12],      // N: DT. ENTREGA
        String(fonteRow[13] !== null && fonteRow[13] !== undefined && fonteRow[13] !== '' ? fonteRow[13] : ''), // O: PRAZO — string evita -1/2 virar data
        timestampFinal,    // P: TIMESTAMP_CRIACAO
        posicaoFonte,      // Q: POSICAO_FONTE — índice na aba "original" (fixo, nunca muda)
        '',                // R: (reservado)
        codigoFixo,        // S: CÓDIGO_FIXO — UUID imutável por item
        fonteRow[22] !== undefined ? fonteRow[22] : '',  // T: INFO_X — coluna X da fonte (informação adicional da OC)
        fonteRow[23] !== undefined ? fonteRow[23] : ''   // U: LOTE — coluna Y da fonte (número de lote)
      ];

      novasPedidosData.push(novaLinha);
    });

    // PASSO 4: Escrever dados em PEDIDOS
    if (novasPedidosData.length > 0) {
      // Limpa dados antigos
      if (pedidosLastRow >= FONTE_DATA_START_ROW) {
        pedidosSheet.getRange(FONTE_DATA_START_ROW, 1, pedidosLastRow - FONTE_DATA_START_ROW + 1, pedidosSheet.getLastColumn()).clearContent();
      }

      // Formata colunas como texto puro antes de escrever —
      // clearContent() não remove formatação, então células antigas poderiam
      // ainda estar como "data" e converter os novos valores silenciosamente.
      // Limita ao máximo entre as linhas antigas e as novas (não a planilha inteira).
      const txtFmt = '@';
      const rowsToFormat = Math.max(
        pedidosLastRow >= FONTE_DATA_START_ROW ? pedidosLastRow : FONTE_DATA_START_ROW,
        FONTE_DATA_START_ROW + novasPedidosData.length - 1
      );
      pedidosSheet.getRange(1, 6,  rowsToFormat, 2).setNumberFormat(txtFmt); // F, G
      pedidosSheet.getRange(1, 12, rowsToFormat, 1).setNumberFormat(txtFmt); // L
      pedidosSheet.getRange(1, 20, rowsToFormat, 2).setNumberFormat(txtFmt); // T (INFO_X), U (LOTE)

      // Escreve novos dados
      pedidosSheet.getRange(FONTE_DATA_START_ROW, 1, novasPedidosData.length, 21).setValues(novasPedidosData);

      SpreadsheetApp.flush();

      const tempoTotal = Date.now() - inicioSync;
      Logger.log("\n" + "=".repeat(70));
      Logger.log(`✅ SINCRONIZAÇÃO CONCLUÍDA EM ${tempoTotal}ms`);
      Logger.log(`   📊 Total de linhas: ${novasPedidosData.length}`);
      Logger.log(`   🆕 Itens novos: ${novosItens}`);
      Logger.log(`   🔄 Itens correspondidos: ${itensAtualizados} (${itensComMudancaReal} com dados alterados)`);
      Logger.log("=".repeat(70));

      // BUG FIX: houveMudancas agora só é true se há itens NOVOS ou com dados
      // realmente alterados. Antes era true para qualquer item correspondido,
      // causando limpeza desnecessária de cache a cada execução do trigger.
      // Grava timestamp do IMPORTRANGE processado — próxima execução só roda se G2 mudar
      props.setProperty('ULTIMO_IMPORTRANGE_TS', tsAtual);

      const houveMudancas = novosItens > 0 || itensComMudancaReal > 0;
      return {
        houveMudancas: houveMudancas,
        novos: novosItens,
        atualizados: itensAtualizados,
        total: novasPedidosData.length
      };
    } else {
      Logger.log("⚠️ Nenhum dado para sincronizar");
      return { houveMudancas: false, novos: 0, atualizados: 0 };
    }

  } catch (e) {
    Logger.log(`❌ ERRO na sincronização: ${e.message}`);
    Logger.log(`   Stack: ${e.stack}`);
    return { houveMudancas: false, erro: e.message };
  }
}

/**
 * Normaliza CÓD. MARFIM para o cliente Dilly substituindo o sufixo após o último
 * "-" pelo valor numérico extraído do TAMANHO. Operação idempotente: aplicar em
 * um valor já corrigido retorna o mesmo valor.
 * Guard: só corrige se o sufixo tiver 2+ caracteres — protege códigos no formato
 * "7490-1" (sufixo "1" = código de versão, não de tamanho).
 * Ex.: ("196338-120", "110CM", "DILLY NORDESTE") → "196338-110"
 *      ("7490-1",     "100CM", "DILLY NORDESTE") → "7490-1"  (sufixo curto, não altera)
 */
function _normalizarMarfimDilly_(marfim, tam, cliente) {
  if (!cliente.toUpperCase().includes('DILLY')) return marfim;
  const dashIdx = marfim.lastIndexOf('-');
  if (dashIdx === -1) return marfim;
  const sufixo = marfim.substring(dashIdx + 1);
  if (sufixo.length < 2) return marfim; // "7490-1" style: sufixo curto, não é código de tamanho
  const numTam = tam.replace(/[^0-9]/g, '');
  if (numTam) return marfim.substring(0, dashIdx + 1) + numTam;
  return marfim;
}

/**
 * Cria impressão digital de uma row com offset configurável
 * @param {Array} row - Array com dados da linha
 * @param {Number} offset - Offset das colunas (0 para DADOS_IMPORTADOS, 1 para PEDIDOS)
 */
function _criarImpressaoDigitalFromRow_(row, offset) {
  // FIX: CARTELA foi removida da impressão digital pois pode ser atualizada
  // pelo sistema de origem. Usar CARTELA causava falsos "novos itens" quando
  // ela mudava, perdendo todo o histórico e marcações do item no Relatorio_DB.
  //
  // Campos estáveis usados como identidade (inclui TAMANHO para distinguir tamanhos):
  // DADOS_IMPORTADOS (offset=0): lido a partir col B → CLIENTE=1, PEDIDO=3, MARFIM=5, TAM=7, OC=8, OS=10, DATA=11
  // PEDIDOS          (offset=1): CLIENTE=2, PEDIDO=4, MARFIM=6, TAM=8, OC=9, OS=11, DATA=12

  const cliente = String(row[1 + offset] || '').trim();
  const pedido  = String(row[3 + offset] || '').trim();
  const tam     = String(row[7 + offset] || '').trim();  // TAMANHO — diferencia itens do mesmo pedido em tamanhos distintos
  const oc      = String(row[8 + offset] || '').trim();
  const os      = String(row[10 + offset] || '').trim();
  const dataStr = _normalizarData_(row[11 + offset]);    // normalizado para Date e número serial

  // Normaliza CÓD. MARFIM para Dilly: garante fingerprint idêntica tanto ao ler
  // de DADOS_IMPORTADOS (valor original, ex: "202480-105") quanto de PEDIDOS
  // (valor já corrigido, ex: "202480-100"). A correção é idempotente.
  const marfim = _normalizarMarfimDilly_(String(row[5 + offset] || '').trim(), tam, cliente);

  return `${cliente}|${pedido}|${marfim}|${tam}|${oc}|${os}|${dataStr}`;
}

// ─── CONTROLE DE PAUSA ───────────────────────────────────────────────────────

/** Retorna true se o sistema estiver pausado pelo usuário */
function _sistemaPausado_() {
  return PropertiesService.getScriptProperties().getProperty('SISTEMA_PAUSADO') === 'true';
}

/** Pausa todas as atualizações automáticas (import + sync) */
function pausarSistema() {
  PropertiesService.getScriptProperties().setProperty('SISTEMA_PAUSADO', 'true');
  SpreadsheetApp.getUi().alert(
    '⏸ Sistema Pausado',
    'As atualizações automáticas foram pausadas.\n' +
    'Você pode editar os dados livremente.\n\n' +
    'Use "▶ Retomar sistema" no menu para reativar.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  Logger.log('⏸ Sistema pausado pelo usuário.');
}

/** Retoma as atualizações automáticas */
function retomarSistema() {
  PropertiesService.getScriptProperties().deleteProperty('SISTEMA_PAUSADO');
  SpreadsheetApp.getUi().alert(
    '▶ Sistema Retomado',
    'As atualizações automáticas foram reativadas.\n' +
    'O sistema voltará a importar e sincronizar normalmente.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  Logger.log('▶ Sistema retomado pelo usuário.');
}

/**
 * Agenda processoAutomaticoCompleto para rodar após N segundos (one-time trigger).
 * Garante que o sync sempre rode DEPOIS do import, em sequência.
 * Salva o ID do trigger no PropertiesService para cancelar se necessário.
 */
function _agendarSincronizacao_(segundos) {
  const props = PropertiesService.getScriptProperties();
  // Cancela agendamento anterior pendente
  const anteriorId = props.getProperty('SYNC_ONETIME_TRIGGER_ID');
  if (anteriorId) {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getUniqueId() === anteriorId)
      .forEach(t => { try { ScriptApp.deleteTrigger(t); } catch (_) {} });
    props.deleteProperty('SYNC_ONETIME_TRIGGER_ID');
  }
  // Cria trigger one-time
  const trigger = ScriptApp.newTrigger('processoAutomaticoCompleto')
    .timeBased()
    .after(segundos * 1000)
    .create();
  props.setProperty('SYNC_ONETIME_TRIGGER_ID', trigger.getUniqueId());
  Logger.log(`   ⏱️ Sincronização agendada em ${segundos}s`);
}

// ─── TRIGGERS AUTOMÁTICOS ─────────────────────────────────────────────────────

/**
 * IMPORTAÇÃO AUTOMÁTICA — Trigger recorrente (a cada 5 min)
 * Responsável apenas por: planilha externa → DADOS_IMPORTADOS
 * Ao concluir, agenda processoAutomaticoCompleto para 90s depois (sequência garantida).
 */
function processoImportacao() {
  if (_sistemaPausado_()) {
    Logger.log('⏸ Sistema pausado — importação ignorada.');
    return;
  }
  const inicio = Date.now();
  Logger.log("=" .repeat(70));
  Logger.log(`📡 IMPORTAÇÃO AUTOMÁTICA INICIADA - ${new Date().toLocaleString('pt-BR')}`);
  Logger.log("=".repeat(70));
  try {
    const resultado = importarDadosExternos();
    if (resultado.success) {
      Logger.log(`✅ ${resultado.linhas} linhas importadas em ${Date.now() - inicio}ms`);
      // Agenda sincronização 90s após o import — garante que DADOS_IMPORTADOS
      // já esteja gravado antes de o sync ler os dados.
      _agendarSincronizacao_(90);
    } else {
      Logger.log(`⚠️ Falha na importação: ${resultado.erro}`);
    }
  } catch (e) {
    Logger.log(`❌ Erro em processoImportacao: ${e.message}`);
  }
  Logger.log("=".repeat(70));
}

/**
 * PROCESSO AUTOMÁTICO DE SINCRONIZAÇÃO — Trigger separado (a cada 5 min)
 * Responsável por: DADOS_IMPORTADOS → PEDIDOS → Relatorio_DB → purga → cache
 * Separado de processoImportacao para respeitar o limite de 6 minutos.
 *
 * OTIMIZAÇÕES:
 * 1. Só roda a sync se H2 mudou (guard de timestamp)
 * 2. Só regenera IDs se necessário
 * 3. Só limpa cache se houve mudanças
 */
function processoAutomaticoCompleto() {
  if (_sistemaPausado_()) {
    Logger.log('⏸ Sistema pausado — sincronização ignorada.');
    return;
  }

  // Previne execuções simultâneas que causam duplicação de dados no DB
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Aguarda até 30s para adquirir o lock
  } catch (e) {
    Logger.log("⚠️ Não foi possível adquirir o lock - outra execução já está em andamento. Abortando.");
    return;
  }

  const inicioProcesso = Date.now();
  Logger.log("=" .repeat(70));
  Logger.log(`⏰ PROCESSO AUTOMÁTICO INICIADO - ${new Date().toLocaleString('pt-BR')}`);
  Logger.log("=".repeat(70));

  let houveMudancas = false;

  try {
    // ETAPA 0: Sincronizar DADOS_IMPORTADOS → PEDIDOS
    // forcarExecucao=false: usa o guard de H2 — processoImportacao() atualiza H2
    // quando importa dados novos; se H2 não mudou, a sync é ignorada (eficiente).
    Logger.log("\n📥 ETAPA 0: Sincronização DADOS_IMPORTADOS → PEDIDOS");
    const resultadoSyncFonte = sincronizarPedidosComFonte(false);

    if (resultadoSyncFonte.houveMudancas) {
      Logger.log(`   ✅ Sincronização concluída:`);
      Logger.log(`      🆕 Novos: ${resultadoSyncFonte.novos || 0}`);
      Logger.log(`      🔄 Atualizados: ${resultadoSyncFonte.atualizados || 0}`);
      houveMudancas = true;
    } else if (resultadoSyncFonte.erro) {
      Logger.log(`   ⚠️ Erro: ${resultadoSyncFonte.erro}`);
      // Continua processo mesmo com erro (pode ser primeira execução)
    } else {
      Logger.log(`   ✓ Nenhuma mudança detectada`);
    }

    // ETAPA 1: Verificar e gerar IDs faltantes (mantido por compatibilidade)
    Logger.log("\n🔑 ETAPA 1: Verificação de IDs");
    const resultadoIds = verificarEGerarIDs();

    if (resultadoIds.regenerou) {
      Logger.log(`   ✅ IDs regenerados: ${resultadoIds.gerados || 0}`);
      houveMudancas = true;
    } else {
      Logger.log(`   ✓ ${resultadoIds.motivo || 'Nenhuma alteração necessária'}`);
    }

    // ETAPA 2: Sincronizar PEDIDOS → Relatorio_DB
    Logger.log("\n🔄 ETAPA 2: Sincronização PEDIDOS → Relatorio_DB");
    const resultadoSync = sincronizarDadosOtimizado();

    if (resultadoSync.houveMudancas) {
      Logger.log(`   ✅ Mudanças detectadas na sincronização`);
      houveMudancas = true;
    } else {
      Logger.log(`   ✓ Nenhuma mudança - dados já sincronizados`);
    }

    // ETAPA 3: Purgar itens finalizados com mais de DIAS_RETENCAO dias
    Logger.log(`\n🗑️ ETAPA 3: Purga de itens finalizados (>${DIAS_RETENCAO} dias)`);
    try {
      const resultadoPurga = purgarItensFinalizados();
      if (resultadoPurga.purgados > 0) {
        Logger.log(`   ✅ ${resultadoPurga.purgados} item(ns) purgado(s) do DB`);
        houveMudancas = true;
      } else {
        Logger.log(`   ✓ Nenhum item para purgar`);
      }
    } catch (ePurga) {
      Logger.log(`   ⚠️ Erro na purga (não crítico): ${ePurga.message}`);
    }

    // ETAPA 4: Limpar cache APENAS se houve mudanças
    Logger.log("\n🗑️ ETAPA 4: Limpeza de cache");
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
    try {
      Logger.log("\n❌ ERRO NO PROCESSO AUTOMÁTICO:");
      Logger.log(`   Mensagem: ${erro ? erro.message : 'Erro desconhecido'}`);
      if (erro && erro.stack) {
        Logger.log(`   Stack: ${erro.stack}`);
      }
      Logger.log("=".repeat(70));
    } catch (logError) {
      // Se até o log falhar, tenta console.log
      console.log("Erro crítico:", erro);
    }

    // Envia email de notificação em caso de erro (opcional)
    // MailApp.sendEmail({
    //   to: Session.getEffectiveUser().getEmail(),
    //   subject: "⚠️ Erro no Processo Automático",
    //   body: `Erro: ${erro.message}\n\nDetalhes: ${erro.stack}`
    // });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Instala o trigger automático SEM ALERTAS (para executar pelo Apps Script).
 * Use esta função quando executar pelo Apps Script Editor.
 */
function instalarTriggerAutomaticoSilencioso() {
  try {
    Logger.log("🔄 Instalando triggers automáticos...");

    // Remove triggers antigos
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;

    triggers.forEach(trigger => {
      const funcao = trigger.getHandlerFunction();
      if (funcao === 'verificarEGerarIDs' || funcao === 'processoAutomaticoCompleto' || funcao === 'processoImportacao') {
        ScriptApp.deleteTrigger(trigger);
        removidos++;
        Logger.log(`   ✓ Removido trigger: ${funcao}`);
      }
    });

    if (removidos > 0) {
      Logger.log(`✅ ${removidos} trigger(s) antigo(s) removido(s)`);
    }

    // Trigger 1: importação da planilha externa → DADOS_IMPORTADOS
    ScriptApp.newTrigger('processoImportacao')
      .timeBased()
      .everyMinutes(5)
      .create();

    // Trigger 2: sincronização DADOS_IMPORTADOS → PEDIDOS → DB
    ScriptApp.newTrigger('processoAutomaticoCompleto')
      .timeBased()
      .everyMinutes(5)
      .create();

    Logger.log("✅ TRIGGERS INSTALADOS COM SUCESSO!");
    Logger.log("   • processoImportacao       → a cada 5 min (importa dados externos)");
    Logger.log("   • processoAutomaticoCompleto → a cada 5 min (sincroniza PEDIDOS → DB)");

    return { success: true };

  } catch (e) {
    Logger.log(`❌ ERRO ao instalar trigger: ${e.message}`);
    Logger.log(`   Stack: ${e.stack}`);
    return { success: false, error: e.message };
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

    // Trigger 1: importação da planilha externa → DADOS_IMPORTADOS (a cada 5 min)
    ScriptApp.newTrigger('processoImportacao')
      .timeBased()
      .everyMinutes(5)
      .create();

    // Trigger 2: sincronização DADOS_IMPORTADOS → PEDIDOS → DB (a cada 5 min)
    ScriptApp.newTrigger('processoAutomaticoCompleto')
      .timeBased()
      .everyMinutes(5)
      .create();

    SpreadsheetApp.getUi().alert(
      '✅ Triggers Automáticos Ativados!',
      'Dois triggers foram instalados (cada um a cada 5 minutos):\n\n' +
      '• processoImportacao: importa dados da planilha externa\n' +
      '• processoAutomaticoCompleto: sincroniza PEDIDOS → DB\n\n' +
      'Para desativar, use o menu: IDs Personalizados > Desativar Geração Automática',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    Logger.log("✅ Dois triggers automáticos instalados com sucesso");
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
    // Limpa também o trigger one-time de sincronização, se existir
    const props = PropertiesService.getScriptProperties();
    const oneTimeId = props.getProperty('SYNC_ONETIME_TRIGGER_ID');
    if (oneTimeId) {
      ScriptApp.getProjectTriggers()
        .filter(t => t.getUniqueId() === oneTimeId)
        .forEach(t => { try { ScriptApp.deleteTrigger(t); } catch (_) {} });
      props.deleteProperty('SYNC_ONETIME_TRIGGER_ID');
    }

    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;

    triggers.forEach(trigger => {
      const funcao = trigger.getHandlerFunction();
      if (funcao === 'verificarEGerarIDs' || funcao === 'processoAutomaticoCompleto' || funcao === 'processoImportacao') {
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

  const sheet = getSpreadsheet_().getSheetByName(FONTE_SHEET_NAME);
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
 * @param {Array} row - linha de dados
 * @param {boolean} isDbRow - true se a row vier do Relatorio_DB (layout sem gap da col D de PEDIDOS)
 * Retorna uma string única baseada em: CLIENTE + PEDIDO + MARFIM + OC + OS + DATA
 */
function _normalizarData_(val) {
  // Normaliza datas para string comparável, independente do tipo (Date, número serial ou string).
  // O IMPORTRANGE pode entregar datas como números seriais (ex: 45123) enquanto o PEDIDOS
  // as lê de volta como objetos Date. Sem normalização, a fingerprint muda a cada sync.
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, TZ, 'yyyyMMdd');
  }
  if (typeof val === 'number' && val > 40000) {
    // Número serial do Sheets → converter para Date e formatar
    try {
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      return Utilities.formatDate(d, TZ, 'yyyyMMdd');
    } catch (e) { return String(val); }
  }
  return String(val || '').trim();
}

function _criarImpressaoDigital_(row, isDbRow) {
  // PEDIDOS tem um "gap" na coluna D (índice 3) que não é gravado no Relatorio_DB.
  // Por isso os índices no DB são -1 em relação aos do PEDIDOS a partir de PEDIDO_COL.
  // isDbRow=true → usa índices relativos ao layout do Relatorio_DB.
  const pedidoIdx = isDbRow ? 3  : PEDIDO_COL;   // PEDIDOS=4, DB=3
  const marfimIdx = isDbRow ? 5  : MARFIM_COL;   // PEDIDOS=6, DB=5
  const tamIdx    = isDbRow ? 7  : TAM_COL;       // PEDIDOS=8, DB=7  ← NOVO: inclui TAMANHO
  const ocIdx     = isDbRow ? 8  : OC_COL;       // PEDIDOS=9, DB=8
  const osIdx     = isDbRow ? 10 : OS_COL;       // PEDIDOS=11, DB=10
  const dtrecIdx  = isDbRow ? 11 : DTREC_COL;    // PEDIDOS=12, DB=11

  const _cliente_ = String(row[CLIENTE_COL] || '').trim();
  const _tam_     = String(row[tamIdx]      || '').trim();
  const _marfim_  = _normalizarMarfimDilly_(String(row[marfimIdx] || '').trim(), _tam_, _cliente_);

  const partes = [
    _cliente_,                                     // índice 2 em ambos
    String(row[pedidoIdx]   || '').trim(),
    _marfim_,                                      // normalizado para Dilly (idempotente)
    _tam_,                                         // TAMANHO — distingue itens do mesmo pedido em tamanhos diferentes
    String(row[ocIdx]       || '').trim(),
    String(row[osIdx]       || '').trim(),
    _normalizarData_(row[dtrecIdx])                // normalizado para lidar com Date e número serial
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
    const impressao = _criarImpressaoDigital_(row, true); // rows do Relatorio_DB
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
  // Logger.clear() removido — apagava logs de sincronizarPedidosComFonte (ETAPA 0) que rodam antes
  Logger.log("=".repeat(70));
  Logger.log(`SINCRONIZAÇÃO v${APP_VERSION} - ${new Date().toLocaleString('pt-BR')}`);
  Logger.log("=".repeat(70));
  
  const startTime = Date.now();
  
  try {
    const fonteSheet = getSpreadsheet_().getSheetByName(FONTE_SHEET_NAME);
    const dbSheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
    
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
      // Lê até o número real de colunas do DB (mín. 21 para cobrir LOTE em U)
      // Status em O (índice 14), MARCAR_FATURAR em P (índice 15),
      // DATA_STATUS em Q (índice 16), CÓDIGO_FIXO em S (índice 18), INFO_X em T (índice 19), LOTE em U (índice 20)
      const dbNumCols = Math.max(21, dbSheet.getLastColumn());
      dbData = dbSheet.getRange(2, 1, dbRows, dbNumCols).getValues();
    }

    const dbMap = new Map();
    const statusCount = { Ativo: 0, Inativo: 0, Faturado: 0, Excluido: 0 };

    dbData.forEach((row, idx) => {
      // Calcula a linha real ANTES de verificar se tem ID
      // Isso garante que linhas vazias também sejam contabilizadas
      const linhaReal = idx + 2;  // +2 porque linha 1 é cabeçalho e idx começa em 0

      const id = row[ID_COL];  // Coluna A (índice 0)
      if (id && String(id).trim()) {
        const idStr = String(id).trim();
        dbMap.set(idStr, { row: row, linha: linhaReal });
        const st = row[STATUS_COL];  // Coluna O (índice 14)
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

    // Map<impressao, [{id, row, usado}]> para PEDIDOS — multi-map para suportar itens com fingerprint idêntica
    // FIX Bug 2: era Map simples (sobrescrevia duplicatas); agora guarda array e consome um slot por vez.
    const fonteImpressoes = new Map();
    const fonteImpressoesCount = new Map(); // fingerprint → quantas vezes aparece em PEDIDOS
    for (let [id, row] of fonteMap.entries()) {
      const impressao = _criarImpressaoDigital_(row);
      if (!fonteImpressoes.has(impressao)) fonteImpressoes.set(impressao, []);
      fonteImpressoes.get(impressao).push({ id: id, row: row, usado: false });
      fonteImpressoesCount.set(impressao, (fonteImpressoesCount.get(impressao) || 0) + 1);
    }
    Logger.log(`   ✓ ${fonteImpressoes.size} impressões digitais únicas criadas para PEDIDOS`);

    // Map<impressao, {id, linha, row}> para Relatorio_DB (usado só para cheque de duplicatas em novos itens)
    const dbImpressoes = new Map();
    for (let [id, dbItem] of dbMap.entries()) {
      const impressao = _criarImpressaoDigital_(dbItem.row, true); // rows do Relatorio_DB
      dbImpressoes.set(impressao, { id: id, linha: dbItem.linha, row: dbItem.row });
    }
    Logger.log(`   ✓ ${dbImpressoes.size} impressões digitais criadas para Relatorio_DB`);

    // Map<codigoFixo, {id, row}> para PEDIDOS — busca por UUID imutável
    // Fonte primária: col S (CÓDIGO_FIXO). Fallback: UUID embutido na DESC (col H).
    // Dois lugares para o mesmo UUID garante localização mesmo se uma das fontes for perdida.
    const fonteCodigoFixoMap = new Map();
    const _uuidRegexCf_ = /\[([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\]$/i;
    for (let [id, row] of fonteMap.entries()) {
      const cf = String(row[PEDIDOS_CODIGO_FIXO_COL] || '').trim();
      if (cf && !fonteCodigoFixoMap.has(cf)) fonteCodigoFixoMap.set(cf, { id: id, row: row });
      // Fallback: UUID visível na DESC — redundância para resistir a limpeza acidental da col S
      const m = String(row[DESC_COL] || '').match(_uuidRegexCf_);
      if (m && !fonteCodigoFixoMap.has(m[1])) fonteCodigoFixoMap.set(m[1], { id: id, row: row });
    }
    Logger.log(`   ✓ ${fonteCodigoFixoMap.size} UUIDs fixos indexados de PEDIDOS (col S + DESC)`);

    // 2.6) PROTEÇÃO ANTI-FATURAMENTO INDEVIDO
    // Lê OC+OS de DADOS_IMPORTADOS para distinguir dois cenários:
    //   A) OC|OS sumiu de DADOS_IMPORTADOS → item saiu do sistema de origem → faturamento OK
    //   B) OC|OS ainda existe → item ainda está no sistema, só o ID/fingerprint mudou
    //      (ex: rebuild zerou IDs, dado alterado na fonte) → NÃO marcar Faturado
    // IMPORTANTE: usa OC+OS (não só OC) — uma OC pode ter muitos itens; só porque
    // alguns sumiram não significa que todos sumiram. OS identifica a linha individual.
    Logger.log("\n🛡️ 2.6. OC+OS EM DADOS_IMPORTADOS (proteção anti-faturamento)");
    // Map<"OC|OS", count> — contagem proporcional: cada match bem-sucedido decrementa 1.
    // Proteção ativa apenas enquanto ainda há slots não consumidos (count > 0).
    // Isso resolve o caso de itens 100% idênticos: se DB tem 3 e PEDIDOS tem 2,
    // após 2 matches o count cai a 0 e o 3º item NÃO é mais protegido → faturamento correto.
    const dadosImportadosOcs = new Map(); // chave: "OC|OS" → contagem de ocorrências
    try {
      const importSheet = getSpreadsheet_().getSheetByName(IMPORTRANGE_SHEET_NAME);
      if (importSheet && importSheet.getLastRow() >= FONTE_DATA_START_ROW) {
        const impLastRow = importSheet.getLastRow();
        // Col J (10) = OC, Col L (12) = OS — lê 3 colunas (J, K, L) e usa índices 0 e 2
        // Usa getDisplayValues() para preservar sufixos como "13807U", "14660U"
        const vals = importSheet.getRange(FONTE_DATA_START_ROW, 10, impLastRow - FONTE_DATA_START_ROW + 1, 3).getDisplayValues();
        vals.forEach(([oc, , os]) => {
          const ocStr = String(oc || '').trim();
          const osStr = String(os || '').trim();
          if (ocStr || osStr) {
            const k = `${ocStr}|${osStr}`;
            dadosImportadosOcs.set(k, (dadosImportadosOcs.get(k) || 0) + 1);
          }
        });
        Logger.log(`   ✓ ${dadosImportadosOcs.size} pares OC+OS únicos em DADOS_IMPORTADOS`);
      } else {
        Logger.log(`   ⚠️ DADOS_IMPORTADOS vazio — proteção desabilitada nesta execução`);
      }
    } catch (eImp) {
      Logger.log(`   ⚠️ Erro ao carregar OC+OS de DADOS_IMPORTADOS: ${eImp.message} — proteção desabilitada`);
    }
    // Pré-computa contagem de itens ATIVOS no DB por OC+OS (independente de ordem de processamento).
    // Comparado com dadosImportadosOcs para detectar itens em excesso no DB:
    //   fonte >= DB ativo → todos DB-itens com este OC+OS têm correspondência possível na fonte → PROTEGE
    //   fonte <  DB ativo → DB tem itens a mais que a fonte → o excedente é genuinamente órfão → NÃO PROTEGE
    // Isso resolve o problema de ordem: o cálculo é feito antes do loop principal.
    const dbActiveOcOsCount = new Map();
    for (const [, dbi] of dbMap.entries()) {
      const s = String(dbi.row[STATUS_COL] || '').trim();
      if (s === 'Faturado' || s === 'Finalizado' || s === 'Excluido') continue;
      const oc = String(dbi.row[DB_OC_COL] || '').trim();
      const os = String(dbi.row[10] || '').trim();
      const k = `${oc}|${os}`;
      dbActiveOcOsCount.set(k, (dbActiveOcOsCount.get(k) || 0) + 1);
    }
    Logger.log(`   ✓ ${dbActiveOcOsCount.size} pares OC+OS únicos em DB (itens ativos)`);

    // 3) PROCESSAR
    Logger.log("\n🔄 3. PROCESSANDO");

    let novos = [];
    let updates = [];
    let avisos = [];
    let idsAtualizados = [];
    let autoExcluidos = 0;

    // Itens que saíram do PEDIDOS e precisam ser marcados Faturado.
    // São coletados primeiro e processados depois, ordenados por QTD.ABERTA crescente,
    // garantindo que QTD=0 (totalmente baixados) sejam tratados antes dos parciais.
    const itensFaturarPendentes = [];

    // Rastreia fingerprints de itens do DB que foram "consumidos" (matched por ID ou por fingerprint).
    // Usado para liberar a mesma fingerprint a novos itens legítimos (ex: vários itens iguais na mesma OC).
    const consumedFingerprints = new Set();

    // Precarrega IDs com histórico de baixas para detectar itens parcializados removidos
    const idsBaixados = new Set();
    try {
      const baixasSheet = getSpreadsheet_().getSheetByName(BAIXAS_SHEET_NAME);
      if (baixasSheet && baixasSheet.getLastRow() > 1) {
        const baixasIds = baixasSheet.getRange(2, 1, baixasSheet.getLastRow() - 1, 1).getValues();
        baixasIds.forEach(r => { if (r[0]) idsBaixados.add(String(r[0]).trim()); });
      }
    } catch (e) {
      Logger.log(`   ⚠️ Não foi possível carregar IDs de baixas: ${e.message}`);
    }
    Logger.log(`   ${idsBaixados.size} IDs com histórico de baixas`);

    // Baseline de QTD por ID: valor da fonte quando o ciclo de baixas começou (ou após último CHECKPOINT).
    // Usado para detectar quando a fonte atualizou QTD → reset do ciclo de faturamento.
    const qtdBaselineCicloMap = _buildQtdOriginalCache_();
    Logger.log(`   ${Object.keys(qtdBaselineCicloMap).length} baselines de ciclo carregados`);

    // Itens que sofreram reset nesta sync — precisam ter MARCAR_FATURAR_USUARIO (col V) limpo separadamente,
    // pois novaLinha só cobre 21 colunas (A-U) e não alcança a col V.
    const idsParaLimparUsuario = [];

    // IDs com alerta ativo = MARCAR_FATURAR=SIM foi setado pelo SYNC (saída inesperada).
    // IDs sem alerta ativo = MARCAR_FATURAR=SIM foi setado pelo USUÁRIO (fluxo correto).
    const idsComAlertaAtivo = new Set();
    try {
      const alertas = JSON.parse(PropertiesService.getScriptProperties().getProperty(ALERTAS_PROP_KEY) || '[]');
      alertas.forEach(a => { if (a.itemId) idsComAlertaAtivo.add(String(a.itemId).trim()); });
    } catch (e) {
      Logger.log(`   ⚠️ Não foi possível carregar alertas ativos: ${e.message}`);
    }
    Logger.log(`   ${idsComAlertaAtivo.size} IDs com alerta de faturamento ativo`);

    for (let [id, dbItem] of dbMap.entries()) {
      const statusAtual = dbItem.row[STATUS_COL];  // Coluna O (índice 14)

      // Se item está Excluido mas voltou ao PEDIDOS: reativa e remove do fonteMap
      // (sem isso, fonteMap.delete nunca seria chamado e o item seria duplicado)
      if (statusAtual === "Excluido") {
        if (fonteMap.has(id)) {
          Logger.log(`   ♻️ ID="${id}" estava Excluido mas voltou ao PEDIDOS - reativando como Ativo`);
          const fonteRow = fonteMap.get(id);
          const marcarFaturarAtual = dbItem.row[MARCAR_FATURAR_COL] || "";
          const cfReativa = dbItem.row[DB_CODIGO_FIXO_COL] || fonteRow[PEDIDOS_CODIGO_FIXO_COL] || '';
          const novaLinha = [
            fonteRow[ID_COL],      fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
            fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
            fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
            fonteRow[QTD_COL],     fonteRow[OS_COL],      fonteRow[DTREC_COL],
            fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "Ativo",               marcarFaturarAtual,
            '',  fonteRow[PEDIDOS_POSICAO_FONTE_COL] ?? '',  cfReativa,   // Q: DATA_STATUS vazia ao reativar, R: POSICAO_FONTE, S: CÓDIGO_FIXO
            fonteRow[PEDIDOS_COLX_COL] || '',  // T: INFO_X
            fonteRow[PEDIDOS_LOTE_COL] || ''   // U: LOTE
          ];
          updates.push({ linha: dbItem.linha, dados: novaLinha, de: statusAtual, para: "Ativo", id: id });
          fonteMap.delete(id);
        }
        continue;
      }

      // PRIMEIRA TENTATIVA: Buscar por ID
      if (fonteMap.has(id)) {
        const fonteRow = fonteMap.get(id);

        // Array de 16 elementos (índices 0-15)
        // Posição 14 é Status na coluna O
        // Posição 15 é MARCAR_FATURAR na coluna P
        let marcarFaturarAtual = dbItem.row[MARCAR_FATURAR_COL] || "";
        const cfMatch = dbItem.row[DB_CODIGO_FIXO_COL] || fonteRow[PEDIDOS_CODIGO_FIXO_COL] || '';

        // Determina QTD antes de montar novaLinha para que o loop "mudou" detecte a diferença.
        // • Sem baixas → fonte é autoritativa.
        // • Com baixas e QTD fonte inalterada → preserva DB.
        // • Com baixas e QTD fonte mudou (qualquer direção, ≠ 0) → reset ciclo faturamento.
        const pedidosQtd      = Number(fonteRow[QTD_COL] || 0);
        const dbQtd           = Number(dbItem.row[DB_QTD_COL] || 0);
        const temBaixasId     = idsBaixados.has(id);
        const baselineCicloId = qtdBaselineCicloMap[id];
        const resetCiclo = temBaixasId && pedidosQtd > 0
          && baselineCicloId !== undefined && pedidosQtd !== baselineCicloId
          && statusAtual !== "Faturado" && statusAtual !== "Finalizado";
        const qtdParaDB = (temBaixasId && !resetCiclo) ? dbQtd : pedidosQtd;
        if (resetCiclo) {
          marcarFaturarAtual = '';
          _registrarCheckpointFaturamento_(id, pedidosQtd);
          _removerAlertasDoItem_(id);
          idsParaLimparUsuario.push({ id: id, linha: dbItem.linha });
          Logger.log(`   🔁 RESET ciclo faturamento: ID="${id}" baseline=${baselineCicloId} → fonte=${pedidosQtd}`);
        } else if (!temBaixasId && pedidosQtd !== dbQtd) {
          Logger.log(`   🔄 QTD sincronizada com fonte: ${dbQtd} → ${pedidosQtd} (sem baixas registradas, ID="${id}")`);
        }

        const novaLinha = [
          fonteRow[ID_COL],      fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
          fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
          fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
          qtdParaDB,             fonteRow[OS_COL],      fonteRow[DTREC_COL],
          fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "",                    marcarFaturarAtual,
          dbItem.row[DATA_STATUS_COL] || '', fonteRow[PEDIDOS_POSICAO_FONTE_COL] ?? '', cfMatch,  // Q: DATA_STATUS preservado, R: POSICAO_FONTE, S: CÓDIGO_FIXO
          fonteRow[PEDIDOS_COLX_COL] || '',  // T: INFO_X
          fonteRow[PEDIDOS_LOTE_COL] || ''   // U: LOTE
        ];

        let mudou = false;
        // Compara as 14 primeiras colunas (0-13) mais INFO_X (19), excluindo Status e MARCAR_FATURAR
        for (let i = 0; i < STATUS_COL; i++) {
          let dbVal = (dbItem.row[i] instanceof Date) ? _toISOStringSafe_(dbItem.row[i]) : dbItem.row[i];
          let novoVal = (novaLinha[i] instanceof Date) ? _toISOStringSafe_(novaLinha[i]) : novaLinha[i];
          if (dbVal != novoVal) { mudou = true; break; }
        }
        if (!mudou) {
          const dbInfoX  = String(dbItem.row[DB_COLX_COL]  !== undefined ? dbItem.row[DB_COLX_COL]  : '');
          const novInfoX = String(novaLinha[DB_COLX_COL] !== undefined ? novaLinha[DB_COLX_COL] : '');
          if (dbInfoX !== novInfoX) mudou = true;
        }
        if (!mudou) {
          const dbLote  = String(dbItem.row[DB_LOTE_COL]  !== undefined ? dbItem.row[DB_LOTE_COL]  : '');
          const novLote = String(novaLinha[DB_LOTE_COL] !== undefined ? novaLinha[DB_LOTE_COL] : '');
          if (dbLote !== novLote) mudou = true;
        }

        if (mudou || resetCiclo || statusAtual === "Inativo") {
          // FIX: preserva "Faturado" e "Finalizado" - não regride para "Ativo" se o item
          // voltou ao DADOS_IMPORTADOS após já ter sido processado pelo usuário do HTML.
          const novoStatus = (statusAtual === "Faturado" || statusAtual === "Finalizado") ? statusAtual : "Ativo";
          novaLinha[STATUS_COL] = novoStatus;  // Coluna O (índice 14)
          updates.push({ linha: dbItem.linha, dados: novaLinha, de: statusAtual, para: novoStatus, id: id });
        }

        // Detecta divergência de QTD quando o usuário tem baixas (sem-baixas é auto-corrigido acima).
        // Alerta apenas quando: item tem baixas E fonte reduziu abaixo do DB E fingerprint é única.
        // Não alerta quando resetCiclo — a mudança já foi tratada como início de novo ciclo.
        const fpFonte      = _criarImpressaoDigital_(fonteRow);
        const fpDuplicados = (fonteImpressoesCount.get(fpFonte) || 1) > 1;
        if (temBaixasId && !resetCiclo && pedidosQtd < dbQtd && statusAtual !== "Faturado" && statusAtual !== "Finalizado") {
          if (fpDuplicados) {
            Logger.log(`   ⚠️ QTD difere mas fingerprint tem ${fonteImpressoesCount.get(fpFonte)} itens em PEDIDOS: ID="${id}" PEDIDOS=${pedidosQtd} DB=${dbQtd} — desalinhamento esperado entre itens duplicados, ignorado`);
          } else {
            Logger.log(`   ⚠️ DIVERGÊNCIA QTD (com baixas): ID="${id}" PEDIDOS=${pedidosQtd} < DB=${dbQtd} — fonte reduziu mas há baixas registradas`);
            _registrarAlertaFaturamento_({
              tipo: 'divergencia_qtd',
              id: id,
              cartela: String(dbItem.row[CARTELA_COL]   || ''),
              cliente: String(dbItem.row[CLIENTE_COL]   || ''),
              pedido:  String(dbItem.row[DB_PEDIDO_COL] || ''),
              oc:      String(dbItem.row[DB_OC_COL]     || ''),
              desc:    String(dbItem.row[DB_DESC_COL]   || ''),
              tam:     String(dbItem.row[DB_TAM_COL]    || ''),
              pedidosQtd: pedidosQtd,
              dbQtd:      dbQtd,
              dataEvento: new Date().toISOString()
            });
          }
        }

        fonteMap.delete(id);
        // Marca slot no fonteImpressoes como usado — sem isso, itens excluídos da fonte
        // conseguem fazer fingerprint match com este slot e não são marcados como Faturado.
        const fpFonteId = _criarImpressaoDigital_(fonteRow);
        const fpListId = fonteImpressoes.get(fpFonteId);
        if (fpListId) { const fi = fpListId.find(i => i.id === id); if (fi) fi.usado = true; }
        consumedFingerprints.add(_criarImpressaoDigital_(dbItem.row, true)); // libera fingerprint para novos itens idênticos legítimos

      } else {
        // SEGUNDA TENTATIVA: Buscar por CÓDIGO_FIXO (UUID imutável por item — ideia do usuário)
        // Mais robusto que fingerprint: funciona mesmo que campos de dados mudem junto com a linha.
        const cfDb = String(dbItem.row[DB_CODIGO_FIXO_COL] || '').trim();
        const fonteItemByCf = (cfDb && fonteCodigoFixoMap.has(cfDb)) ? fonteCodigoFixoMap.get(cfDb) : null;

        if (fonteItemByCf) {
          // ENCONTROU POR UUID! O ID mudou mas o UUID é o mesmo
          const novoId = fonteItemByCf.id;
          Logger.log(`   🔑 ID atualizado por UUID: "${id}" → "${novoId}" (Linha ${dbItem.linha})`);

          const fonteRow = fonteItemByCf.row;
          let marcarFaturarAtual = dbItem.row[MARCAR_FATURAR_COL] || "";

          const pedidosQtdCf    = Number(fonteRow[QTD_COL] || 0);
          const dbQtdCf         = Number(dbItem.row[DB_QTD_COL] || 0);
          const temBaixasCf     = idsBaixados.has(id);
          const baselineCicloCf = qtdBaselineCicloMap[id];
          const resetCicloCf    = temBaixasCf && pedidosQtdCf > 0
            && baselineCicloCf !== undefined && pedidosQtdCf !== baselineCicloCf
            && statusAtual !== "Faturado" && statusAtual !== "Finalizado";
          const qtdParaDBCf     = (temBaixasCf && !resetCicloCf) ? dbQtdCf : pedidosQtdCf;
          if (resetCicloCf) {
            marcarFaturarAtual = '';
            _registrarCheckpointFaturamento_(id, pedidosQtdCf);
            _removerAlertasDoItem_(id);
            idsParaLimparUsuario.push({ id: id, linha: dbItem.linha });
            Logger.log(`   🔁 RESET ciclo faturamento (UUID): ID="${id}" baseline=${baselineCicloCf} → fonte=${pedidosQtdCf}`);
          } else if (!temBaixasCf && pedidosQtdCf !== dbQtdCf) {
            Logger.log(`   🔄 QTD sincronizada com fonte (UUID): ${dbQtdCf} → ${pedidosQtdCf} (sem baixas registradas, ID="${id}")`);
          }

          const novaLinha = [
            novoId,                fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
            fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
            fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
            qtdParaDBCf,           fonteRow[OS_COL],      fonteRow[DTREC_COL],
            fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "",                    marcarFaturarAtual,
            dbItem.row[DATA_STATUS_COL] || '', fonteRow[PEDIDOS_POSICAO_FONTE_COL] ?? '', cfDb,  // Q: DATA_STATUS preservado, R: POSICAO_FONTE, S: UUID preservado
            fonteRow[PEDIDOS_COLX_COL] || '',  // T: INFO_X
            fonteRow[PEDIDOS_LOTE_COL] || ''   // U: LOTE
          ];

          const novoStatus = (statusAtual === "Faturado" || statusAtual === "Finalizado") ? statusAtual : "Ativo";
          novaLinha[STATUS_COL] = novoStatus;

          updates.push({ linha: dbItem.linha, dados: novaLinha, de: statusAtual, para: novoStatus });
          idsAtualizados.push({ de: id, para: novoId, linha: dbItem.linha });

          fonteMap.delete(novoId);
          fonteCodigoFixoMap.delete(cfDb); // consome este UUID para não reutilizar

          // Marca como usado no fonteImpressoes para evitar double-match por fingerprint
          const fpFonte = _criarImpressaoDigital_(fonteRow);
          const fpList = fonteImpressoes.get(fpFonte);
          if (fpList) { const fi = fpList.find(i => i.id === novoId); if (fi) fi.usado = true; }
          consumedFingerprints.add(_criarImpressaoDigital_(dbItem.row, true));

        } else {
          // TERCEIRA TENTATIVA: Buscar por IMPRESSÃO DIGITAL (fallback para itens sem UUID ou UUID ausente)
          const impressaoDB = _criarImpressaoDigital_(dbItem.row, true); // row do Relatorio_DB
          // FIX Bug 2: era fonteImpressoes.get() (sobrescrevia duplicatas); agora encontra primeiro slot livre.
          const fonteItens = fonteImpressoes.get(impressaoDB);
          const fonteItem = fonteItens ? fonteItens.find(i => !i.usado) : null;

          if (fonteItem) {
            // ENCONTROU POR FINGERPRINT! O ID mudou devido ao IMPORTRANGE
            fonteItem.usado = true; // consome este slot sem apagar outros com a mesma fingerprint
            const novoId = fonteItem.id;
            Logger.log(`   🔄 ID atualizado por fingerprint: "${id}" → "${novoId}" (Linha ${dbItem.linha})`);

            const fonteRow = fonteItem.row;
            let marcarFaturarAtual = dbItem.row[MARCAR_FATURAR_COL] || "";

            const pedidosQtdFp    = Number(fonteRow[QTD_COL] || 0);
            const dbQtdFp         = Number(dbItem.row[DB_QTD_COL] || 0);
            const temBaixasFp     = idsBaixados.has(id);
            const baselineCicloFp = qtdBaselineCicloMap[id];
            const resetCicloFp    = temBaixasFp && pedidosQtdFp > 0
              && baselineCicloFp !== undefined && pedidosQtdFp !== baselineCicloFp
              && statusAtual !== "Faturado" && statusAtual !== "Finalizado";
            const qtdParaDBFp     = (temBaixasFp && !resetCicloFp) ? dbQtdFp : pedidosQtdFp;
            if (resetCicloFp) {
              marcarFaturarAtual = '';
              _registrarCheckpointFaturamento_(id, pedidosQtdFp);
              _removerAlertasDoItem_(id);
              idsParaLimparUsuario.push({ id: id, linha: dbItem.linha });
              Logger.log(`   🔁 RESET ciclo faturamento (fp): ID="${id}" baseline=${baselineCicloFp} → fonte=${pedidosQtdFp}`);
            } else if (!temBaixasFp && pedidosQtdFp !== dbQtdFp) {
              Logger.log(`   🔄 QTD sincronizada com fonte (fingerprint): ${dbQtdFp} → ${pedidosQtdFp} (sem baixas registradas, ID="${id}")`);
            }

            const cfFp = dbItem.row[DB_CODIGO_FIXO_COL] || fonteRow[PEDIDOS_CODIGO_FIXO_COL] || '';
            const novaLinha = [
              novoId,                fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
              fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
              fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
              qtdParaDBFp,           fonteRow[OS_COL],      fonteRow[DTREC_COL],
              fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "",                    marcarFaturarAtual,
              dbItem.row[DATA_STATUS_COL] || '', fonteRow[PEDIDOS_POSICAO_FONTE_COL] ?? '', cfFp,  // Q: DATA_STATUS preservado, R: POSICAO_FONTE, S: CÓDIGO_FIXO
              fonteRow[PEDIDOS_COLX_COL] || '',  // T: INFO_X
              fonteRow[PEDIDOS_LOTE_COL] || ''   // U: LOTE
            ];

            const novoStatus = (statusAtual === "Faturado" || statusAtual === "Finalizado") ? statusAtual : "Ativo";
            novaLinha[STATUS_COL] = novoStatus;

            updates.push({ linha: dbItem.linha, dados: novaLinha, de: statusAtual, para: novoStatus });
            idsAtualizados.push({ de: id, para: novoId, linha: dbItem.linha });

            // Remove do fonteMap para não adicionar como novo depois.
            // Não apaga a chave do fonteImpressoes — apenas o slot foi marcado como usado,
            // permitindo que outros DB-items com a mesma fingerprint ainda encontrem seus slots.
            fonteMap.delete(novoId);
            consumedFingerprints.add(impressaoDB); // libera fingerprint para novos itens idênticos legítimos

          } else {
            // NÃO ENCONTROU por ID, UUID nem fingerprint — item não está em PEDIDOS
            const ocDB = String(dbItem.row[DB_OC_COL] || '').trim();
            const osDB = String(dbItem.row[10]          || '').trim(); // índice 10 = CÓD. OS no DB
            Logger.log(`   ❌ ID="${id}" não encontrado em PEDIDOS — OC="${ocDB}" OS="${osDB}", Status="${statusAtual}"`);

            // PROTEÇÃO ANTI-FATURAMENTO INDEVIDO:
            // Verifica se o par OC+OS ainda existe em DADOS_IMPORTADOS.
            // Usar só OC era amplo demais: se uma OC tem 18 itens e 8 sumiram, os 8 ficavam
            // bloqueados porque os outros 10 mantinham o OC presente.
            // Com OC+OS, cada linha é identificada individualmente (OS é único por linha).
            //   • OC+OS presente → item ainda no sistema, só ID/fingerprint mudou (rebuild/dado alterado)
            //                      → NÃO marcar Faturado (sync vai reconsolidar na próxima rodada)
            //   • OC+OS ausente  → item genuinamente saiu do sistema de origem → prossegue normal
            const chaveOcOs = `${ocDB}|${osDB}`;

            // PROTEÇÃO ANTI-FATURAMENTO INDEVIDO (proporcional, sem bloquear ações do usuário):
            // Compara contagem de slots em DADOS_IMPORTADOS vs DB ativo por OC+OS.
            //   fonte >= DB ativo → possível mismatch de ID (rebuild) → proteção ativa
            //   fonte <  DB ativo → item em excesso → proteção inativa (prossegue)
            //   fonte = 0         → item saiu do sistema → proteção inativa (prossegue)
            //
            // A proteção NÃO bloqueia:
            //   • MARCAR_FATURAR=SIM já setado (usuário marcou explicitamente)
            //   • QTD=0 (baixas completas — item entregue, pode faturar)
            //   • idsBaixados (usuário fez ao menos uma baixa — item sendo trabalhado)
            // A proteção BLOQUEIA apenas:
            //   • QTD>0 + sem baixas + sem marcação → possível ID-mismatch, aguarda reconsolidação
            let protecaoAtiva = false;
            if (dadosImportadosOcs.size > 0) {
              const slotsOrigem   = dadosImportadosOcs.get(chaveOcOs) || 0;
              const slotsDbAtivos = dbActiveOcOsCount.get(chaveOcOs)  || 0;
              if (slotsOrigem > 0 && slotsOrigem >= slotsDbAtivos) {
                protecaoAtiva = true;
                Logger.log(`   🛡️ Proteção: OC+OS "${chaveOcOs}" fonte=${slotsOrigem} ≥ DB ativo=${slotsDbAtivos}`);
              } else if (slotsOrigem > 0) {
                Logger.log(`   ⚠️ OC+OS "${chaveOcOs}" fonte=${slotsOrigem} < DB ativo=${slotsDbAtivos} → item em excesso no DB, prossegue`);
              }
            }

            // FIX: Se o usuário do HTML já marcou o item para faturar (MARCAR_FATURAR=SIM),
            // o item pode ter sido fechado/removido pelo sistema de origem mas ainda não foi
            // faturado. Manter visível e ativo para o usuário do HTML concluir o processo.
            // SEMPRE processado, independente da proteção.
            const marcarFaturar = String(dbItem.row[MARCAR_FATURAR_COL] || '').trim().toUpperCase();
            const aguardandoNF = marcarFaturar === 'SIM';
            const qtdAberta    = Number(dbItem.row[DB_QTD_COL] || 0);
            const temBaixas    = idsBaixados.has(id); // usuário já processou pelo menos uma baixa

            // Proteção extra: se MARCAR_FATURAR=SIM, verifica fingerprint para detectar
            // ID mudado (item ainda existe em PEDIDOS com outro ID) — mantém Ativo nesses casos.
            let itemAindaEmPedidosPorFingerprint = false;
            if (aguardandoNF) {
              const fpDB = _criarImpressaoDigital_(dbItem.row, true);
              itemAindaEmPedidosPorFingerprint = fonteImpressoes.has(fpDB) && fonteImpressoes.get(fpDB).some(i => !i.usado);
              if (itemAindaEmPedidosPorFingerprint) {
                Logger.log(`   ⚠️ DUPLICATA: Item aguardando NF existe em PEDIDOS com ID diferente — mantido Ativo OC="${dbItem.row[DB_OC_COL]}" QTD=${qtdAberta} ID="${id}"`);

              }
            }

            if (!itemAindaEmPedidosPorFingerprint && statusAtual !== "Faturado" && statusAtual !== "Finalizado" && statusAtual !== "Excluido") {
              if (aguardandoNF && !idsComAlertaAtivo.has(id)) {
                // MARCAR_FATURAR=SIM posto pelo USUÁRIO (sem alerta ativo) → saída esperada → Faturado direto
                Logger.log(`   ✋→✅ Marcado pelo usuário + saiu do PEDIDOS → Faturado direto (QTD=${qtdAberta}, ID="${id}")`);
                itensFaturarPendentes.push({ id: id, linha: dbItem.linha, row: dbItem.row, statusAtual: statusAtual, qtdAberta: 0, marcadoParaNF: true });
              } else if (aguardandoNF && idsComAlertaAtivo.has(id)) {
                // MARCAR_FATURAR=SIM posto pelo SYNC (alerta ativo) → mantém Ativo, aguarda confirmação do usuário
                Logger.log(`   ✋ Alerta ativo pendente de confirmação — mantido Ativo (QTD=${qtdAberta}, ID="${id}")`);
              } else if (qtdAberta === 0) {
                // QTD=0 sem marcação: baixas zeraram o item → faturado silencioso
                itensFaturarPendentes.push({ id: id, linha: dbItem.linha, row: dbItem.row, statusAtual: statusAtual, qtdAberta: 0 });
              } else if (!protecaoAtiva || temBaixas) {
                // QTD>0 sem marcação: saída inesperada → gera alerta
                itensFaturarPendentes.push({ id: id, linha: dbItem.linha, row: dbItem.row, statusAtual: statusAtual, qtdAberta: qtdAberta });
              } else {
                // QTD>0 + proteção ativa + sem baixas → aguarda reconsolidação de ID
                Logger.log(`   ⏭️ Proteção ativa (QTD=${qtdAberta}, sem baixas) → aguarda reconsolidação — OC+OS="${chaveOcOs}"`);
              }
            } else if (!itemAindaEmPedidosPorFingerprint) {
              Logger.log(`   ℹ️ Não alterado (status: ${statusAtual})`);
            }
          }
        }
      }
    }
    
    // === PROCESSAR ITENS QUE SAÍRAM DO PEDIDOS (ordenado por QTD.ABERTA crescente) ===
    // Ordena: QTD=0 primeiro (faturamento silencioso), QTD>0 por último (gera alerta).
    // Com múltiplos itens idênticos na mesma OC, isso garante que os totalmente
    // baixados sejam os primeiros a sair, e só gera alerta se realmente há QTD parcial.
    itensFaturarPendentes.sort((a, b) => a.qtdAberta - b.qtdAberta);

    if (itensFaturarPendentes.length > 0) {
      Logger.log(`\n🔄 Processando ${itensFaturarPendentes.length} item(ns) que saíram do PEDIDOS (ordenado por QTD.ABERTA):`);
    }

    itensFaturarPendentes.forEach(({ id, linha, row, statusAtual, qtdAberta, marcadoParaNF }) => {
      const linhaAtualizar = [...row];

      if (qtdAberta === 0 || marcadoParaNF) {
        // QTD=0: baixa zerou o item  |  marcadoParaNF: usuário marcou + saiu do PEDIDOS → Faturado direto
        linhaAtualizar[STATUS_COL]       = "Faturado";
        linhaAtualizar[DATA_STATUS_COL]  = new Date();
        linhaAtualizar[MARCAR_FATURAR_COL] = "";
        updates.push({ linha: linha, dados: linhaAtualizar, de: statusAtual, para: "Faturado", id: id });
        autoExcluidos++;
        Logger.log(`   ✅ ${marcadoParaNF ? 'Marcado p/ NF + saiu' : 'QTD.ABERTA=0'} → Faturado (ID="${id}")`);
      } else {
        // QTD>0 sem marcação: saída inesperada → sinaliza e gera alerta
        linhaAtualizar[MARCAR_FATURAR_COL] = "SIM";
        updates.push({ linha: linha, dados: linhaAtualizar, de: statusAtual, para: statusAtual, id: id });
        Logger.log(`   ⚠️ QTD.ABERTA=${qtdAberta} → MARCAR_FATURAR=SIM + ALERTA, mantido ${statusAtual} (ID="${id}")`);
        _registrarAlertaFaturamento_({
          id: id,
          cartela: String(row[CARTELA_COL]    || ''),
          cliente: String(row[CLIENTE_COL]    || ''),
          pedido:  String(row[DB_PEDIDO_COL]  || ''),
          oc:      String(row[DB_OC_COL]      || ''),
          desc:    String(row[DB_DESC_COL]    || ''),
          tam:     String(row[DB_TAM_COL]     || ''),
          qtdAberta: qtdAberta,
          dataEvento: new Date().toISOString()
        });
      }
    });

    // Novos itens que estão em PEDIDOS mas não em Relatorio_DB
    const duplicatasDebug = []; // acumula itens descartados para auditoria
    for (let [id, fonteRow] of fonteMap.entries()) {
      // Proteção extra: verifica por impressão digital mesmo que o ID seja "novo".
      // Evita duplicação quando sincronizarPedidosComFonte gera ID diferente para
      // um item que já existe no DB (ex: por inconsistência de dados na source).
      const impressaoFonte = _criarImpressaoDigital_(fonteRow);
      // Só rejeita se o item do DB com esta fingerprint NÃO foi consumido (matched).
      // Se foi consumido, a "vaga" foi usada pelo item correspondente e este é um novo item legítimo
      // (ex: segunda unidade de um item idêntico dentro da mesma OC).
      if (dbImpressoes.has(impressaoFonte) && !consumedFingerprints.has(impressaoFonte)) {
        const existente = dbImpressoes.get(impressaoFonte);
        Logger.log(`   ⚠️ DUPLICATA EVITADA POR FINGERPRINT: ID="${id}" já existe no DB como ID="${existente.id}" - ignorado`);
        duplicatasDebug.push([
          new Date(), 'Fingerprint idêntica ao DB', id, existente.id,
          fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL], fonteRow[PEDIDO_COL],
          fonteRow[OC_COL], fonteRow[DESC_COL], fonteRow[TAM_COL]
        ]);
        continue;
      }

      Logger.log(`   🆕 Novo item: ID="${id}" está em PEDIDOS mas não em Relatorio_DB - será adicionado como Ativo`);
      Logger.log(`      CARTELA="${fonteRow[CARTELA_COL]}", CLIENTE="${fonteRow[CLIENTE_COL]}", OC="${fonteRow[OC_COL]}"`);

      // Array de 21 elementos: A-P dados, Q=DATA_STATUS vazio, R=POSICAO_FONTE, S=CÓDIGO_FIXO, T=INFO_X, U=LOTE
      const novaLinha = [
        fonteRow[ID_COL],      fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
        fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
        fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
        fonteRow[QTD_COL],     fonteRow[OS_COL],      fonteRow[DTREC_COL],
        fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "Ativo",               "",
        '',  fonteRow[PEDIDOS_POSICAO_FONTE_COL] ?? '',  fonteRow[PEDIDOS_CODIGO_FIXO_COL] || '',  // Q: DATA_STATUS vazio, R: POSICAO_FONTE, S: CÓDIGO_FIXO
        fonteRow[PEDIDOS_COLX_COL] || '',  // T: INFO_X
        fonteRow[PEDIDOS_LOTE_COL] || ''   // U: LOTE
      ];
      novos.push(novaLinha);
    }
    
    Logger.log(`   🆕 Novos: ${novos.length}`);
    Logger.log(`   📝 Atualizar: ${updates.length}`);
    Logger.log(`   🔄 IDs Atualizados: ${idsAtualizados.length}`);

    // 4) VALIDAÇÃO ANTI-DUPLICATA
    Logger.log("\n🔍 3.5. VALIDAÇÃO ANTI-DUPLICATA");
    const novosValidados = [];
    const idsExistentes = new Set(dbMap.keys());
    const idsJaAdicionados = new Set();

    // Conta quantas vezes cada fingerprint aparece no DB, descontando itens já consumidos
    // (matched por ID ou por fingerprint na fase anterior). Cada "slot" disponível representa
    // uma vaga de duplicata no DB que já está coberta. Itens com fingerprint além dessas vagas
    // são novos legítimos (ex: segunda unidade idêntica na mesma OC).
    const fpDisponiveisDB = new Map(); // fingerprint → quantidade de itens NÃO consumidos no DB
    for (const [, dbItem] of dbMap.entries()) {
      const fp = _criarImpressaoDigital_(dbItem.row, true);
      if (!consumedFingerprints.has(fp)) {
        fpDisponiveisDB.set(fp, (fpDisponiveisDB.get(fp) || 0) + 1);
      }
    }

    novos.forEach(item => {
      const id = String(item[ID_COL]).trim();
      // helper inline para registrar rejeição no debug (usa DB layout: OC=[8], DESC=[6], TAM=[7])
      const _regDebug_ = (motivo, idExistente) => duplicatasDebug.push([
        new Date(), motivo, id, idExistente || '',
        item[1], item[2], item[3], item[8], item[6], item[7]
      ]);

      // Verifica se já existe no DB por ID
      if (idsExistentes.has(id)) {
        Logger.log(`   ⚠️ DUPLICATA EVITADA (ID): ID="${id}" já existe no Relatorio_DB`);
        _regDebug_('ID já existe no DB');
        return;
      }

      // Verifica se já foi adicionado nesta rodada por ID
      if (idsJaAdicionados.has(id)) {
        Logger.log(`   ⚠️ DUPLICATA EVITADA (ID rodada): ID="${id}" já foi processado nesta sincronização`);
        _regDebug_('ID duplicado na mesma rodada');
        return;
      }

      // Verifica se ainda há vagas no DB para esta fingerprint (itens idênticos não consumidos).
      // Se vagas > 0 o item do DB já cobre esta "instância" → rejeita como duplicata real.
      // Se vagas = 0 (todos consumidos/matched) → é um novo item legítimo e pode ser adicionado.
      const fp = _criarImpressaoDigital_(item);
      const vagasDB = fpDisponiveisDB.get(fp) || 0;
      if (vagasDB > 0) {
        Logger.log(`   ⚠️ DUPLICATA EVITADA (fingerprint): ID="${id}" tem mesmos dados de item já existente no DB (vagas=${vagasDB})`);
        _regDebug_('Fingerprint idêntica (validação final)');
        fpDisponiveisDB.set(fp, vagasDB - 1); // consome a vaga para não bloquear mais do que o necessário
        return;
      }

      // Valida se tem dados essenciais
      if (!item[CARTELA_COL] || String(item[CARTELA_COL]).trim() === '') {
        Logger.log(`   ⚠️ ITEM REJEITADO: ID="${id}" sem CARTELA`);
        _regDebug_('Sem CARTELA');
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
      // getLastRow() pode retornar linhas com formatação residual mas sem dado real.
      // Buscamos a última linha com ID_UNICO preenchido na coluna A para inserir logo após.
      const totalRows = dbSheet.getLastRow();
      let proxLinha = 2; // mínimo: logo após o cabeçalho
      if (totalRows >= 2) {
        const colA = dbSheet.getRange(2, 1, totalRows - 1, 1).getValues();
        for (let i = colA.length - 1; i >= 0; i--) {
          if (String(colA[i][0]).trim() !== '') {
            proxLinha = i + 3; // +2 por 0-index e cabeçalho, +1 para próxima linha
            break;
          }
        }
      }
      dbSheet.getRange(proxLinha, 1, novosValidados.length, novosValidados[0].length).setValues(novosValidados);
      Logger.log(`   ✅ ${novosValidados.length} novos adicionados`);
    }
    if (updates.length > 0) {
      updates.forEach(u => {
        dbSheet.getRange(u.linha, 1, 1, u.dados.length).setValues([u.dados]);
        Logger.log(`   ✅ Linha ${u.linha}: ${u.de} → ${u.para} | ID: ${u.id}`);
      });
    }
    // Persiste avisos de itens parcializados que saíram do PEDIDOS
    if (avisos.length > 0) {
      const sp = PropertiesService.getScriptProperties();
      const existentes = JSON.parse(sp.getProperty('AVISOS_PENDENTES') || '[]');
      sp.setProperty('AVISOS_PENDENTES', JSON.stringify(existentes.concat(avisos)));
      Logger.log(`   ⚠️ ${avisos.length} aviso(s) de itens com baixa gravados`);
    }

    // Atualiza IDs no Baixas_Historico quando IDs mudam (UUID ou fingerprint match).
    // Necessário para preservar QTD.ABERTA=0 de itens com baixa cujo ID foi trocado.
    // Caso específico: itens Dilly têm CÓD.OS substituído pelo Lote no PEDIDOS, causando
    // mismatch de fingerprint a cada sync e geração de novo ID (sufixo incremental).
    // Sem este fix, na 2ª sync após a baixa, temBaixasId=false e QTD volta ao valor da fonte.
    if (idsAtualizados.length > 0) {
      try {
        const baixasSheetRef = getSpreadsheet_().getSheetByName(BAIXAS_SHEET_NAME);
        if (baixasSheetRef && baixasSheetRef.getLastRow() > 1) {
          const bNumCols = baixasSheetRef.getLastColumn();
          const bHeaders = baixasSheetRef.getRange(1, 1, 1, bNumCols).getValues()[0];
          const bIdCol = bHeaders.findIndex(h => String(h).trim() === 'ID_ITEM');
          if (bIdCol >= 0) {
            const bNumRows = baixasSheetRef.getLastRow() - 1;
            const bColValues = baixasSheetRef.getRange(2, bIdCol + 1, bNumRows, 1).getValues();
            const aliasMap = {};
            idsAtualizados.forEach(({ de, para }) => { aliasMap[String(de).trim()] = para; });
            let bChanged = false;
            bColValues.forEach((row, i) => {
              const oldId = String(row[0] || '').trim();
              if (aliasMap[oldId]) {
                bColValues[i][0] = aliasMap[oldId];
                bChanged = true;
              }
            });
            if (bChanged) {
              baixasSheetRef.getRange(2, bIdCol + 1, bNumRows, 1).setValues(bColValues);
              Logger.log(`   🔄 IDs atualizados no Baixas_Historico: ${idsAtualizados.map(a => `"${a.de}"→"${a.para}"`).join(', ')}`);
            }
          }
        }
      } catch (eBaixas) {
        Logger.log(`   ⚠️ Erro ao atualizar IDs no Baixas_Historico: ${eBaixas.message}`);
      }
    }

    // Limpa MARCAR_FATURAR_USUARIO (col V, índice 21) para itens com reset de ciclo.
    // novaLinha cobre só 21 colunas (A–U) e não alcança col V — passo separado necessário.
    if (idsParaLimparUsuario.length > 0) {
      idsParaLimparUsuario.forEach(({ linha }) => {
        dbSheet.getRange(linha, MARCAR_FATURAR_USUARIO_COL + 1).setValue('');
      });
      Logger.log(`   🧹 MARCAR_FATURAR_USUARIO limpo para ${idsParaLimparUsuario.length} item(ns) com reset de ciclo`);
    }

    SpreadsheetApp.flush();
    if (novosValidados.length > 0 || updates.length > 0) {
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
    Logger.log(`   • ${updates.length} itens atualizados no Relatorio_DB (QTD. ABERTA preservada do DB)`);
    Logger.log(`   • ${autoExcluidos} itens marcados como Faturado (saíram do PEDIDOS)`);
    Logger.log(`   • ${avisos.length} aviso(s) de itens com baixa removidos de PEDIDOS`);
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

    // Grava aba de auditoria de duplicatas (sempre, para refletir estado atual)
    _gravarDuplicatasDebug_(duplicatasDebug);

    // Retorna contadores para o processo automático decidir se limpa cache
    return {
      novos: novosValidados.length,
      updates: updates.length,
      inativos: autoExcluidos,
      avisos: avisos.length,
      idsAtualizados: idsAtualizados.length
    };

  } catch (error) {
    Logger.log("\n❌ ERRO: " + error.message);
    throw error;
  }
}

// ====== ALERTAS DE FATURAMENTO ======
const ALERTAS_PROP_KEY = 'ALERTAS_FATURAMENTO';

/**
 * Lê a senha de controle da aba SENHA, célula A2.
 */
function _getSenha_() {
  try {
    const sheet = getSpreadsheet_().getSheetByName('SENHA');
    if (!sheet) return null;
    const val = sheet.getRange('A2').getValue();
    return val ? String(val).trim() : null;
  } catch (e) {
    Logger.log('⚠️ _getSenha_: ' + e.message);
    return null;
  }
}

/**
 * Registra um novo alerta (faturado_sem_baixa ou divergencia_qtd).
 * Cada alerta recebe um ID único baseado em timestamp + ID do item.
 * Evita duplicata por itemId + tipo.
 */
function _registrarAlertaFaturamento_(dados) {
  // Alertas de faturamento desativados — confirmação automática.
  Logger.log(`ℹ️ Alerta suprimido (${dados.tipo || 'faturado_sem_baixa'}): ID="${dados.id}"`);
}

/**
 * Retorna a lista de alertas pendentes para o HTML.
 * Auto-limpa alertas obsoletos conforme o tipo:
 *   faturado_sem_baixa → limpa se item não está mais como Faturado no DB
 *   divergencia_qtd    → limpa se item virou Faturado, saiu do DB,
 *                        ou DB QTD já caiu para <= pedidosQtd do alerta (baixa feita)
 */
function obterAlertasPendentes() {
  // Alertas desativados — limpa qualquer resíduo e retorna lista vazia.
  try {
    PropertiesService.getScriptProperties().deleteProperty(ALERTAS_PROP_KEY);
  } catch (_) {}
  return [];
}

/**
 * Valida a senha e, se correta, remove o alerta da lista.
 * Retorna { success, erro } para o HTML.
 */
function confirmarAlerta(alertaId, senhaDigitada) {
  try {
    const senhaCorreta = _getSenha_();
    if (!senhaCorreta) {
      return { success: false, erro: 'Aba SENHA ou célula A2 não encontrada. Configure a senha primeiro.' };
    }
    if (String(senhaDigitada).trim() !== senhaCorreta) {
      return { success: false, erro: 'Senha incorreta.' };
    }
    // Senha correta → remove o alerta
    const sp = PropertiesService.getScriptProperties();
    const lista = JSON.parse(sp.getProperty(ALERTAS_PROP_KEY) || '[]');
    const nova = lista.filter(a => a.alertaId !== alertaId);
    sp.setProperty(ALERTAS_PROP_KEY, JSON.stringify(nova));
    Logger.log(`✅ Alerta ${alertaId} confirmado e removido.`);
    return { success: true, restantes: nova.length };
  } catch (e) {
    Logger.log('⚠️ confirmarAlerta: ' + e.message);
    return { success: false, erro: e.message };
  }
}

/**
 * Remove todos os alertas de faturamento associados a um item específico.
 * Chamado quando a fonte muda QTD e o ciclo de faturamento é reiniciado.
 */
function _removerAlertasDoItem_(id) {
  try {
    const sp = PropertiesService.getScriptProperties();
    const lista = JSON.parse(sp.getProperty(ALERTAS_PROP_KEY) || '[]');
    const nova = lista.filter(a => String(a.id || '').trim() !== String(id || '').trim());
    if (nova.length !== lista.length) {
      sp.setProperty(ALERTAS_PROP_KEY, JSON.stringify(nova));
      Logger.log(`   🗑️ ${lista.length - nova.length} alerta(s) removido(s) para ID="${id}"`);
    }
  } catch (e) {
    Logger.log(`⚠️ _removerAlertasDoItem_: ${e.message}`);
  }
}

// ====== PURGA DE ITENS FINALIZADOS ======
/**
 * Remove do Relatorio_DB itens com status Faturado, Finalizado ou Excluido
 * cuja DATA_STATUS (coluna Q) seja anterior a hoje - DIAS_RETENCAO (padrão: 15 dias).
 *
 * Itens sem DATA_STATUS preenchida são ignorados com segurança (não há risco de
 * apagar algo que não sabemos quando foi alterado).
 *
 * Chamada automaticamente pelo processoAutomaticoCompleto().
 * Pode também ser executada manualmente pelo editor (sem UI).
 */
function purgarItensFinalizados() {
  const STATUS_FINAIS = new Set(['Faturado', 'Finalizado', 'Excluido']);
  const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('ℹ️ purgarItensFinalizados: DB vazio, nada a fazer.');
    return { purgados: 0 };
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = _getColumnIndexes_(headers);
  const statusCol = colMap['Status'];           // índice 0-based
  const dataStatusCol = colMap['DATA_STATUS'];  // índice 0-based

  if (statusCol === undefined || dataStatusCol === undefined) {
    Logger.log('⚠️ purgarItensFinalizados: colunas Status ou DATA_STATUS não encontradas — verifique os cabeçalhos do DB.');
    return { purgados: 0 };
  }

  const numCols = sheet.getLastColumn();
  const numRows = sheet.getLastRow() - 1; // sem cabeçalho
  const dados = sheet.getRange(2, 1, numRows, numCols).getValues();

  const limiteData = new Date();
  limiteData.setDate(limiteData.getDate() - DIAS_RETENCAO);

  // Coleta linhas para deletar de baixo pra cima (evita deslocamento de índice)
  const linhasParaDeletar = [];
  for (let i = dados.length - 1; i >= 0; i--) {
    const status = String(dados[i][statusCol] || '').trim();
    if (!STATUS_FINAIS.has(status)) continue;

    const dataStatus = dados[i][dataStatusCol];
    if (!dataStatus || !(dataStatus instanceof Date) || isNaN(dataStatus.getTime())) {
      // Sem data registrada → não apaga (segurança)
      continue;
    }

    if (dataStatus < limiteData) {
      linhasParaDeletar.push(i + 2); // +1 cabeçalho, +1 base-1
    }
  }

  if (linhasParaDeletar.length === 0) {
    Logger.log(`ℹ️ purgarItensFinalizados: nenhum item com mais de ${DIAS_RETENCAO} dias para purgar.`);
    return { purgados: 0 };
  }

  // Deleta de baixo pra cima para não deslocar índices
  linhasParaDeletar.forEach(linha => sheet.deleteRow(linha));
  SpreadsheetApp.flush();
  limparCache();
  Logger.log(`🗑️ purgarItensFinalizados: ${linhasParaDeletar.length} item(ns) com mais de ${DIAS_RETENCAO} dias purgado(s) do DB.`);
  return { purgados: linhasParaDeletar.length };
}

// ====== AUDITORIA DE DUPLICATAS ======
/**
 * Grava (ou limpa) a aba "Duplicatas_Debug" com os itens descartados na última sincronização.
 * Se não houver duplicatas, a aba fica com apenas o cabeçalho para indicar "tudo ok".
 * @param {Array[]} rows - Array de linhas [[timestamp, motivo, id, idExistente, cartela, cliente, pedido, oc, desc, tam]]
 */
function _gravarDuplicatasDebug_(rows) {
  const SHEET_NAME = 'Duplicatas_Debug';
  let sheet = getSpreadsheet_().getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = getSpreadsheet_().insertSheet(SHEET_NAME);
  }
  sheet.clearContents();

  const cabecalho = [
    'TIMESTAMP_EXEC', 'MOTIVO', 'ID_DESCARTADO', 'ID_EXISTENTE_NO_DB',
    'CARTELA', 'CLIENTE', 'PEDIDO', 'OC', 'DESC', 'TAMANHO'
  ];
  sheet.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);

  if (rows && rows.length > 0) {
    sheet.getRange(2, 1, rows.length, cabecalho.length).setValues(rows);
    Logger.log(`📋 Duplicatas_Debug: ${rows.length} item(ns) registrado(s)`);
  } else {
    Logger.log('📋 Duplicatas_Debug: nenhuma duplicata nesta sincronização');
  }
  SpreadsheetApp.flush();
}

// ====== COMPACTAR DB ======
/**
 * Remove linhas completamente vazias (sem ID na coluna A) do Relatorio_DB.
 * Deleta as linhas de baixo pra cima para não deslocar os índices durante a remoção.
 */
function compactarDB() {
  Logger.clear();
  Logger.log("=== COMPACTAR Relatorio_DB ===");

  const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
  if (!sheet) { Logger.log("❌ Aba não encontrada"); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log("ℹ️ DB vazio, nada a fazer"); return; }

  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const linhasVazias = [];

  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]).trim() === '') {
      linhasVazias.push(i + 2); // linha real na planilha (1-indexed, +1 cabeçalho, +1 0-index)
    }
  }

  if (linhasVazias.length === 0) {
    Logger.log("✅ Nenhuma linha vazia encontrada");
    return;
  }

  // Deleta linha a linha de baixo pra cima (índices já estão em ordem decrescente)
  linhasVazias.forEach(linha => sheet.deleteRow(linha));

  limparCache();
  Logger.log(`✅ ${linhasVazias.length} linha(s) vazia(s) removida(s)`);
}

// ====== CACHE ======
// Tamanho máximo por chave do CacheService (Apps Script limita a 100KB por chave)
const CACHE_CHUNK_SIZE = 90000;  // 90KB por chunk com margem de segurança
const CACHE_MAX_TOTAL  = 500000; // 500KB total (até ~5 chunks)

function limparCache() {
  try {
    const cache = CacheService.getScriptCache();
    // Remove chunks de dados (formato novo)
    const keysToRemove = ['timestamp_dados', 'dados_chunks_count', 'dados_completos'];
    const numChunksStr = cache.get('dados_chunks_count');
    if (numChunksStr) {
      const numChunks = parseInt(numChunksStr);
      for (let i = 0; i < numChunks; i++) keysToRemove.push(`dados_chunk_${i}`);
    }
    cache.removeAll(keysToRemove);
    Logger.log("🗑️ Cache limpo");
  } catch (e) {
    Logger.log("⚠️ Erro ao limpar cache: " + e.message);
  }
}

function obterDadosCache() {
  try {
    const cache = CacheService.getScriptCache();
    const timestamp = cache.get('timestamp_dados');
    if (!timestamp) return null;

    // Tenta formato novo (chunks)
    const numChunksStr = cache.get('dados_chunks_count');
    if (numChunksStr) {
      const numChunks = parseInt(numChunksStr);
      let dadosStr = '';
      for (let i = 0; i < numChunks; i++) {
        const chunk = cache.get(`dados_chunk_${i}`);
        if (!chunk) return null; // chunk expirou antes do timestamp
        dadosStr += chunk;
      }
      const dados = JSON.parse(dadosStr);
      const idade = Date.now() - parseInt(timestamp);
      Logger.log(`📦 Cache hit! ${numChunks} chunk(s), ${Math.floor(dadosStr.length/1024)}KB, Idade: ${Math.floor(idade/1000)}s`);
      return dados;
    }

    // Fallback: formato antigo (chave única)
    const dadosStr = cache.get('dados_completos');
    if (dadosStr) {
      const dados = JSON.parse(dadosStr);
      const idade = Date.now() - parseInt(timestamp);
      Logger.log(`📦 Cache hit (legado)! Idade: ${Math.floor(idade/1000)}s`);
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

    if (dadosStr.length > CACHE_MAX_TOTAL) {
      Logger.log(`⚠️ Dados muito grandes para cache (${Math.floor(dadosStr.length/1024)}KB > ${CACHE_MAX_TOTAL/1024}KB)`);
      return false;
    }

    const numChunks = Math.ceil(dadosStr.length / CACHE_CHUNK_SIZE);
    const cacheData = {
      'dados_chunks_count': numChunks.toString(),
      'timestamp_dados': Date.now().toString()
    };
    for (let i = 0; i < numChunks; i++) {
      cacheData[`dados_chunk_${i}`] = dadosStr.substring(i * CACHE_CHUNK_SIZE, (i + 1) * CACHE_CHUNK_SIZE);
    }
    cache.putAll(cacheData, CACHE_DURATION);
    Logger.log(`💾 Cache salvo: ${numChunks} chunk(s), ${Math.floor(dadosStr.length/1024)}KB, válido por ${CACHE_DURATION/60}min`);
    return true;
  } catch (e) {
    Logger.log("⚠️ Erro ao salvar cache: " + e.message);
    return false;
  }
}

// ====== SISTEMA WEB OTIMIZADO ======

// Cabeçalhos corretos do Relatorio_DB, na ordem exata em que são gravados por sincronizarDados()
const RELATORIO_DB_HEADERS = [
  'ID_UNICO', 'CARTELA', 'CLIENTE', 'PEDIDO', 'CÓD. CLIENTE',
  'CÓD. MARFIM', 'DESCRIÇÃO', 'TAMANHO', 'ORD. COMPRA', 'QTD. ABERTA',
  'CÓD. OS', 'DATA RECEB.', 'DT. ENTREGA', 'PRAZO', 'Status', 'MARCAR_FATURAR',
  'DATA_STATUS',            // Q - data em que o status foi alterado para Faturado/Finalizado/Excluido
  'POSICAO_FONTE',          // R - índice do item em DADOS_IMPORTADOS (preserva ordem original)
  'CODIGO_FIXO',            // S - UUID imutável por item
  'INFO_X',                 // T - campo da coluna X da fonte (informação adicional da OC)
  'LOTE',                   // U - número de lote da coluna Y da fonte
  'MARCAR_FATURAR_USUARIO'  // V - usuário que marcou o item para faturamento
];

/**
 * Garante que a aba Relatorio_DB existe e tem os cabeçalhos corretos na linha 1.
 * Chamada automaticamente por fetchAllDataUnified quando nenhum item é retornado.
 * NÃO sobrescreve cabeçalhos existentes para evitar perda de dados.
 */
function _garantirHeadersRelatorio_DB_() {
  try {
    let sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);

    // Cria a aba se não existir
    if (!sheet) {
      Logger.log(`📝 Criando aba ${DB_SHEET_NAME}...`);
      sheet = getSpreadsheet_().insertSheet(DB_SHEET_NAME);
    }

    // Verifica se a linha 1 está vazia ou sem ID_UNICO
    const primeiraLinha = sheet.getLastRow() >= 1
      ? sheet.getRange(1, 1, 1, Math.max(RELATORIO_DB_HEADERS.length, sheet.getLastColumn())).getValues()[0]
      : [];

    const temHeaderValido = primeiraLinha.some(h => String(h).trim() === 'ID_UNICO');

    if (!temHeaderValido) {
      Logger.log(`📝 Cabeçalhos ausentes ou incorretos — gravando cabeçalhos padrão na linha 1...`);
      Logger.log(`   Cabeçalhos existentes: [${primeiraLinha.filter(h => h).join(', ')}]`);
      sheet.getRange(1, 1, 1, RELATORIO_DB_HEADERS.length).setValues([RELATORIO_DB_HEADERS]);
      sheet.getRange(1, 1, 1, RELATORIO_DB_HEADERS.length).setFontWeight('bold').setBackground('#f0f2f5');
      sheet.setFrozenRows(1);
      SpreadsheetApp.flush();
      Logger.log(`✅ Cabeçalhos gravados: ${RELATORIO_DB_HEADERS.join(', ')}`);
      return true; // headers foram criados/corrigidos
    }

    // Estende cabeçalhos se o DB tem menos colunas que o esperado
    // (ex: nova coluna INFO_X adicionada a uma instalação existente)
    const colsExistentes = primeiraLinha.filter(h => String(h).trim() !== '').length;
    if (colsExistentes < RELATORIO_DB_HEADERS.length) {
      const faltam = RELATORIO_DB_HEADERS.slice(colsExistentes);
      sheet.getRange(1, colsExistentes + 1, 1, faltam.length).setValues([faltam]);
      sheet.getRange(1, colsExistentes + 1, 1, faltam.length).setFontWeight('bold').setBackground('#f0f2f5');
      SpreadsheetApp.flush();
      Logger.log(`📝 Cabeçalhos estendidos: adicionado(s) [${faltam.join(', ')}] a partir da col ${colsExistentes + 1}`);
      return true;
    }

    Logger.log(`✅ Cabeçalhos do ${DB_SHEET_NAME} OK (ID_UNICO encontrado)`);
    return false;
  } catch (e) {
    Logger.log(`❌ Erro ao verificar cabeçalhos: ${e.message}`);
    return false;
  }
}

function _readAllData_() {
  // Abre a planilha fresh a cada leitura para evitar cache de container do Apps Script
  const ss = SpreadsheetApp.openById("1qPJ8c7cq7qb86VJJ-iByeiaPnALOBcDPrPMeL75N2EI");
  const sheet = ss.getSheetByName(DB_SHEET_NAME);
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
    'DESCRIÇÃO': (() => { const d = getDisp('DESCRIÇÃO', 'N/A'); return d.replace(/\s*\[[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\]$/i, '').trim() || d; })(),
    'TAMANHO': getDisp('TAMANHO', 'N/A'),
    'CÓD. MARFIM': getDisp('CÓD. MARFIM', 'N/A'),
    'CÓD. OS': getDisp('CÓD. OS', 'N/A'),
    'ORD. COMPRA': getDisp('ORD. COMPRA', 'SEM OC'),
    CLIENTE: getDisp('CLIENTE', 'SEM CLIENTE'),
    PEDIDO: getDisp('PEDIDO', 'N/A'),

    // NÚMEROS/DATA cruas
    'QTD. ABERTA': qtdAberta,
    'QTD. ORIGINAL': calcularQtdOriginal(uniqueId, qtdAberta),
    'DT. ENTREGA': get('DT. ENTREGA', null),
    'DATA RECEB.': get('DATA RECEB.', null),
    // PRAZO em dias: positivo = dias restantes, negativo = dias de atraso
    'PRAZO': (() => {
      const dtEnt = get('DT. ENTREGA', null);
      if (!dtEnt) return null;
      const dtDate = dtEnt instanceof Date ? dtEnt : new Date(dtEnt);
      if (isNaN(dtDate.getTime())) return null;
      const hoje = new Date();
      hoje.setHours(0, 0, 0, 0);
      const dtNorm = new Date(dtDate.getTime());
      dtNorm.setHours(0, 0, 0, 0);
      return Math.round((dtNorm - hoje) / 86400000);
    })(),

    Status: getDisp('Status', 'Desconhecido'),
    MARCAR_FATURAR: getDisp('MARCAR_FATURAR', ''),
    MARCAR_FATURAR_USUARIO: getDisp('MARCAR_FATURAR_USUARIO', ''),
    INFO_X: getDisp('INFO_X', ''),
    LOTE:   getDisp('LOTE',   ''),
    // Posição original em DADOS_IMPORTADOS — lida por índice fixo (col R = índice 17 no DB)
    posicaoFonte: (typeof row[DB_POSICAO_FONTE_COL] === 'number' && !isNaN(row[DB_POSICAO_FONTE_COL]))
      ? row[DB_POSICAO_FONTE_COL]
      : 999999
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
        infoX: item.INFO_X || '',
        posicaoMin: item.posicaoFonte, // menor posição em DADOS_IMPORTADOS para ordenar os cards
        items: []
      };
    }
    // mantém a menor posição entre todos os itens do grupo
    if (item.posicaoFonte < byOC[oc].posicaoMin) {
      byOC[oc].posicaoMin = item.posicaoFonte;
    }
    // captura o primeiro valor não-vazio de INFO_X entre os itens da OC
    if (!byOC[oc].infoX && item.INFO_X) {
      byOC[oc].infoX = item.INFO_X;
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
    // Garante headers antes de ler — evita colMap vazio quando DB foi limpo e sincronização inseriu dados sem header
    _garantirHeadersRelatorio_DB_();
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

    // Diagnóstico: rows existem mas nenhum item foi retornado → provável problema de cabeçalhos
    if (rows.length > 0 && itemsWeb.length === 0) {
      Logger.log(`⚠️ ATENÇÃO: ${rows.length} linhas lidas mas NENHUM item convertido!`);
      Logger.log(`   Cabeçalhos encontrados: [${headers.filter(h => h).join(', ')}]`);
      Logger.log(`   Cabeçalhos esperados:   [${RELATORIO_DB_HEADERS.join(', ')}]`);
      Logger.log(`   Verifique se 'ID_UNICO' existe exatamente assim na linha 1 do ${DB_SHEET_NAME}`);

      // Tenta corrigir os cabeçalhos automaticamente se estiverem ausentes
      const corrigiu = _garantirHeadersRelatorio_DB_();
      if (corrigiu) {
        Logger.log(`   ✅ Cabeçalhos corrigidos automaticamente. Execute a sincronização para popular os dados.`);
      }
    }

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
    
    // Inclui e limpa avisos pendentes de itens parcializados removidos do PEDIDOS
    const spFetch = PropertiesService.getScriptProperties();
    const avisosPendentes = JSON.parse(spFetch.getProperty('AVISOS_PENDENTES') || '[]');
    if (avisosPendentes.length > 0) {
      spFetch.deleteProperty('AVISOS_PENDENTES');
      Logger.log(`   ⚠️ ${avisosPendentes.length} aviso(s) incluídos e limpos`);
    }
    result.avisosPendentes = avisosPendentes;

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
    const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
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

    // Lê QTD.ABERTA antes de faturar para registrar checkpoint se for faturamento parcial
    const qtdACol = colMap['QTD. ABERTA'];
    const qtdAbertaAtual = qtdACol !== undefined
      ? _toNumber_(sheet.getRange(linhaNum, qtdACol + 1).getValue())
      : 0;

    sheet.getRange(linhaNum, statusCol + 1).setValue("Faturado");
    sheet.getRange(linhaNum, DATA_STATUS_COL + 1).setValue(new Date()); // Q: data do status

    // Se ainda há saldo aberto (faturamento parcial), registra checkpoint para que o próximo
    // relatório de faturamento mostre apenas as baixas realizadas após este ponto.
    if (qtdAbertaAtual > 0 && uniqueId) {
      _registrarCheckpointFaturamento_(uniqueId, qtdAbertaAtual);
    }

    limparCache();
    Logger.log(`💰 ${uniqueId || 'sem-id'} → Faturado (linha ${linhaNum}) | QTD.ABERTA: ${qtdAbertaAtual}`);
    return { success: true, id: uniqueId || null, linha: linhaNum };
  } catch (e) {
    Logger.log(`❌ marcarFaturado: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

function excluirItem(uniqueId, planilhaLinha, _skipCache) {
  try {
    const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
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
    sheet.getRange(linhaNum, DATA_STATUS_COL + 1).setValue(new Date()); // Q: data do status
    if (!_skipCache) limparCache();
    Logger.log(`🗑️ ${uniqueId || 'sem-id'} → Excluido (linha ${linhaNum})`);
    return { success: true, id: uniqueId || null, linha: linhaNum };
  } catch (e) {
    Logger.log(`❌ excluirItem: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

function finalizarItem(uniqueId, planilhaLinha, _skipCache) {
  try {
    const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
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
    sheet.getRange(linhaNum, DATA_STATUS_COL + 1).setValue(new Date()); // Q: data do status
    if (!_skipCache) limparCache();
    Logger.log(`✅ ${uniqueId || 'sem-id'} → Finalizado (linha ${linhaNum})`);
    return { success: true, id: uniqueId || null, linha: linhaNum };
  } catch (e) {
    Logger.log(`❌ finalizarItem: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

function excluirMultiplosItens(items) {
  let ok = 0, fail = 0;
  const results = [];
  (items || []).forEach(it => {
    const linha = (it && it.planilhaLinha != null) ? it.planilhaLinha : (it ? it.linha : null);
    const id = (it && (it.uniqueId || it.id)) || null;
    const r = excluirItem(id, linha, true); // _skipCache=true: limpa uma só vez ao final
    results.push(r);
    r.success ? ok++ : fail++;
  });
  if (ok > 0) limparCache();
  return { success: fail === 0, processados: ok, falhas: fail, results };
}

function finalizarMultiplosItens(items) {
  let ok = 0, fail = 0;
  const results = [];
  (items || []).forEach(it => {
    const linha = (it && it.planilhaLinha != null) ? it.planilhaLinha : (it ? it.linha : null);
    const id = (it && (it.uniqueId || it.id)) || null;
    const r = finalizarItem(id, linha, true); // _skipCache=true: limpa uma só vez ao final
    results.push(r);
    r.success ? ok++ : fail++;
  });
  if (ok > 0) limparCache();
  return { success: fail === 0, processados: ok, falhas: fail, results };
}

// ====== FUNÇÕES PARA MARCAR ITENS PARA FATURAMENTO ======

function marcarParaFaturar(uniqueId, planilhaLinha, marcar, usuario) {
  try {
    const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
    const linhaNum = Number(planilhaLinha);

    if (!sheet) throw new Error("Aba DB não encontrada");
    if (!isFinite(linhaNum) || linhaNum < 2 || linhaNum > sheet.getLastRow()) {
      throw new Error(`Linha inválida: ${planilhaLinha}`);
    }

    // Lê cabeçalhos - força leitura de pelo menos 22 colunas (A-V)
    const numCols = Math.max(sheet.getLastColumn(), 22);
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const colMap = _getColumnIndexes_(headers);
    let marcarCol = colMap['MARCAR_FATURAR'];
    let usuarioCol = colMap['MARCAR_FATURAR_USUARIO'];

    Logger.log(`📋 DEBUG marcarParaFaturar - Colunas lidas: ${numCols}, Headers: ${headers.length}`);
    Logger.log(`📋 DEBUG - MARCAR_FATURAR encontrada no índice: ${marcarCol}`);
    Logger.log(`📋 DEBUG - MARCAR_FATURAR_USUARIO encontrada no índice: ${usuarioCol}`);

    if (marcarCol === undefined) {
      Logger.log("⚠️ Coluna 'MARCAR_FATURAR' não encontrada - criando automaticamente...");
      sheet.getRange(1, 16).setValue('MARCAR_FATURAR');
      marcarCol = 15;
      Logger.log("✅ Coluna 'MARCAR_FATURAR' criada na coluna P");
    }

    if (usuarioCol === undefined) {
      Logger.log("⚠️ Coluna 'MARCAR_FATURAR_USUARIO' não encontrada - criando automaticamente...");
      sheet.getRange(1, 22).setValue('MARCAR_FATURAR_USUARIO');
      usuarioCol = 21;
      Logger.log("✅ Coluna 'MARCAR_FATURAR_USUARIO' criada na coluna V");
    }

    // Ao desmarcar: valida que é o mesmo usuário que marcou
    if (!marcar) {
      const usuarioQueMarkou = String(sheet.getRange(linhaNum, usuarioCol + 1).getValue() || '').trim();
      const usuarioAtual = String(usuario || '').trim();
      if (usuarioQueMarkou && usuarioAtual && usuarioQueMarkou.toLowerCase() !== usuarioAtual.toLowerCase()) {
        Logger.log(`🚫 Desmarcação bloqueada: item marcado por "${usuarioQueMarkou}", tentativa de "${usuarioAtual}"`);
        return {
          success: false,
          bloqueado: true,
          error: `Este item foi marcado por "${usuarioQueMarkou}". Apenas esse usuário pode desmarcá-lo.`,
          id: uniqueId,
          linha: linhaNum
        };
      }
    }

    // Marca ou desmarca
    const valor = marcar ? "SIM" : "";
    const usuarioValor = marcar ? String(usuario || '').trim() : "";
    sheet.getRange(linhaNum, marcarCol + 1).setValue(valor);
    sheet.getRange(linhaNum, usuarioCol + 1).setValue(usuarioValor);

    SpreadsheetApp.flush();
    limparCache();

    Logger.log(`✓ ${uniqueId} → Marcado para faturar: ${marcar} por "${usuarioValor}" (linha ${linhaNum})`);
    return { success: true, id: uniqueId, linha: linhaNum, marcado: marcar, usuario: usuarioValor };
  } catch (e) {
    Logger.log(`❌ marcarParaFaturar: ${e.message}`);
    return { success: false, error: e.message, id: uniqueId || null, linha: planilhaLinha };
  }
}

// Retorna mapa {uniqueId → qtd} lendo direto a aba PEDIDOS coluna K (K4:K).
// Usado pelo Relatório Ponteiras para exibir a quantidade da fonte, não a quantidade
// reduzida por baixas parciais que fica no Relatorio_DB.
function obterQtdPedidos() {
  try {
    const sheet = getSpreadsheet_().getSheetByName(FONTE_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Aba PEDIDOS não encontrada', map: {} };

    const lastRow = sheet.getLastRow();
    if (lastRow < FONTE_DATA_START_ROW) return { success: true, map: {} };

    // Lê apenas colunas A (ID) e K (QTD) para evitar ler planilha inteira
    const numRows = lastRow - FONTE_DATA_START_ROW + 1;
    const colsAK = sheet.getRange(FONTE_DATA_START_ROW, 1, numRows, QTD_COL + 1).getValues();

    const map = {};
    colsAK.forEach(row => {
      const id = String(row[ID_COL] || '').trim();
      if (!id) return;
      map[id] = _toNumber_(row[QTD_COL]);
    });

    Logger.log(`📊 obterQtdPedidos: ${Object.keys(map).length} IDs mapeados`);
    return { success: true, map };
  } catch (e) {
    Logger.log(`❌ obterQtdPedidos: ${e.message}`);
    return { success: false, error: e.message, map: {} };
  }
}

function obterItensMarcadosParaFaturar() {
  Logger.log("🔍 INÍCIO obterItensMarcadosParaFaturar");

  try {
    const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
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

    // Força leitura de pelo menos 22 colunas (A-V) para incluir MARCAR_FATURAR_USUARIO
    const lastCol = Math.max(sheet.getLastColumn(), 22);
    Logger.log(`📊 Lendo ${lastCol} colunas (forçado mínimo 22)`);

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
          // QTD. ORIGINAL já reflete o valor real do DB antes das baixas do ciclo
          // (calculado em calcularQtdOriginal como QTD_ABERTA + baixas desde checkpoint).
          const qtdOriginal = item['QTD. ORIGINAL'] || 0;
          const qtdAberta   = item['QTD. ABERTA']   || 0;
          const saldo = qtdOriginal - qtdAberta;

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
            'PRAZO': item['PRAZO'],                         // Número de dias (positivo=a vencer, negativo=atrasado)
            'DT. ENTREGA': _fmtBR_(item['DT. ENTREGA']),  // Converte Date para string
            'DATA RECEB.': _fmtBR_(item['DATA RECEB.']),  // Converte Date para string
            Status: item.Status,
            MARCAR_FATURAR: item.MARCAR_FATURAR,
            MARCAR_FATURAR_USUARIO: item.MARCAR_FATURAR_USUARIO || '',
            INFO_X: item.INFO_X || '',
            LOTE:   item.LOTE   || '',
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

// ====== CONFIRMAR IMPRESSÃO: registra checkpoints para todos os itens do relatório ======
// Chamada pelo HTML no momento em que o usuário confirma a impressão do relatório de faturamento.
// Para cada item com QTD.ABERTA > 0, grava um CHECKPOINT no Baixas_Historico, redefinindo a
// base de cálculo do SALDO para os próximos faturamentos parciais.
function registrarCheckpointsFaturamento(items) {
  try {
    if (!Array.isArray(items) || items.length === 0) return { success: true, registrados: 0 };
    let registrados = 0;
    items.forEach(item => {
      const uniqueId  = String(item.uniqueId  || '').trim();
      const qtdAberta = Number(item['QTD. ABERTA'] || item.qtdAberta || 0);
      const saldo     = Number(item.SALDO || 0);
      // Só registra checkpoint se há baixas novas a faturar (SALDO > 0).
      // Evita zerar a base em re-impressões sem novas baixas.
      if (uniqueId && qtdAberta > 0 && saldo > 0) {
        _registrarCheckpointFaturamento_(uniqueId, qtdAberta);
        registrados++;
      }
    });
    Logger.log(`✅ registrarCheckpointsFaturamento: ${registrados} checkpoint(s) registrado(s)`);
    return { success: true, registrados };
  } catch (e) {
    Logger.log(`❌ registrarCheckpointsFaturamento: ${e.message}`);
    return { success: false, error: e.message };
  }
}

// ====== UTILITÁRIO: LIMPAR ALERTAS DE FATURAMENTO ======
/**
 * Remove todos os alertas pendentes de faturamento (faturado_sem_baixa e divergencia_qtd).
 * Execute manualmente pelo Apps Script quando quiser zerar os avisos.
 */
function limparTodosAlertas() {
  PropertiesService.getScriptProperties().deleteProperty('ALERTAS_FATURAMENTO');
  Logger.log('✅ Todos os alertas de faturamento foram limpos.');
}

// ====== UTILITÁRIO: CONFIRMAR TODOS OS ALERTAS (USO EM TESTES) ======
/**
 * Marca como "Faturado" todos os itens do Relatorio_DB que possuem MARCAR_FATURAR="SIM"
 * e em seguida limpa todos os alertas pendentes.
 * Use apenas para testes ou correções em lote — não exige senha.
 */
/**
 * Wrapper chamado pelo menu — exibe confirmação antes de executar.
 */
function confirmarTodosAlertasMenu() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '🧹 Confirmar todos os alertas de faturamento',
    'Isso irá:\n\n' +
    '• Marcar como "Faturado" todos os itens com MARCAR_FATURAR = SIM\n' +
    '• Limpar todos os alertas pendentes no Relatorio_DB\n\n' +
    'Use apenas para testes ou correções em lote. Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    Logger.log('ℹ️ confirmarTodosAlertasMenu: cancelado pelo usuário.');
    return;
  }
  confirmarTodosAlertas();
  ui.alert('✅ Concluído', 'Todos os alertas foram confirmados e os itens marcados como Faturado.', ui.ButtonSet.OK);
}

function confirmarTodosAlertas() {
  const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('⚠️ confirmarTodosAlertas: DB vazio ou não encontrado.');
    return;
  }

  const lastRow  = sheet.getLastRow();
  const lastCol  = Math.max(sheet.getLastColumn(), DATA_STATUS_COL + 1);
  const dados    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const agora    = new Date();
  let   marcados = 0;
  let   alertasDismissed = 0;

  dados.forEach((row, i) => {
    const marcar = String(row[MARCAR_FATURAR_COL] || '').trim().toUpperCase();
    if (marcar !== 'SIM') return;

    const linhaSheet = i + 2;
    const uniqueId   = String(row[ID_COL] || '').trim();
    const qtdAberta  = _toNumber_(row[DB_QTD_COL]);

    if (qtdAberta === 0) {
      // QTD zerada: baixa foi feita — seguro marcar como Faturado
      sheet.getRange(linhaSheet, STATUS_COL + 1).setValue('Faturado');
      sheet.getRange(linhaSheet, MARCAR_FATURAR_COL + 1).setValue('');
      sheet.getRange(linhaSheet, DATA_STATUS_COL + 1).setValue(agora);
      marcados++;
      Logger.log(`💰 Linha ${linhaSheet} → Faturado (ID="${uniqueId}") | QTD.ABERTA: 0`);
    } else {
      // QTD ainda aberta: apenas descarta o alerta, mantém Ativo
      sheet.getRange(linhaSheet, MARCAR_FATURAR_COL + 1).setValue('');
      alertasDismissed++;
      Logger.log(`⚠️ Linha ${linhaSheet} → alerta descartado, mantido Ativo (ID="${uniqueId}") | QTD.ABERTA: ${qtdAberta}`);
    }
  });

  PropertiesService.getScriptProperties().deleteProperty('ALERTAS_FATURAMENTO');
  limparCache();

  Logger.log(`✅ confirmarTodosAlertas: ${marcados} marcado(s) como Faturado, ${alertasDismissed} alerta(s) descartado(s) com QTD aberta.`);
}

/**
 * Corrige itens com Status=Faturado mas QTD.ABERTA > 0 (marcados incorretamente).
 * Reverte o status para Ativo e limpa MARCAR_FATURAR.
 * Execute manualmente pelo menu do Apps Script quando necessário.
 */
function corrigirFaturadosComSaldoAberto() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSpreadsheet_().getSheetByName(DB_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('DB vazio ou não encontrado.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), DATA_STATUS_COL + 1);
  const dados   = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let corrigidos = 0;

  dados.forEach((row, i) => {
    const status    = String(row[STATUS_COL] || '').trim();
    const qtdAberta = _toNumber_(row[DB_QTD_COL]);
    if (status !== 'Faturado' || qtdAberta === 0) return;

    const linhaSheet = i + 2;
    const uniqueId   = String(row[ID_COL] || '').trim();
    sheet.getRange(linhaSheet, STATUS_COL + 1).setValue('Ativo');
    sheet.getRange(linhaSheet, MARCAR_FATURAR_COL + 1).setValue('');
    sheet.getRange(linhaSheet, DATA_STATUS_COL + 1).setValue('');
    corrigidos++;
    Logger.log(`🔧 Linha ${linhaSheet} → revertido Faturado→Ativo (ID="${uniqueId}") | QTD.ABERTA: ${qtdAberta}`);
  });

  limparCache();
  ui.alert('✅ Correção concluída', `${corrigidos} item(ns) revertido(s) de Faturado → Ativo por ter QTD.ABERTA > 0.`, ui.ButtonSet.OK);
  Logger.log(`✅ corrigirFaturadosComSaldoAberto: ${corrigidos} item(ns) corrigido(s).`);
}

// ====== SISTEMA DE LOGIN COM NÍVEIS DE ACESSO ======
/**
 * Autentica usuário contra a aba CADASTRO.
 * Col A = usuário, Col B = senha, Col C = nível (TOTAL/PARCIAL).
 * Retorna { success, nivel, tempoSessao } ou { success: false, erro }.
 */
function autenticarLogin(usuario, senha) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName('CADASTRO');
    if (!sheet) {
      return { success: false, erro: 'Aba CADASTRO não encontrada. Crie a aba com os usuários.' };
    }

    // Lê tempo de sessão de E1 (padrão: 15 minutos)
    let tempoSessao = 15;
    try {
      const e1 = sheet.getRange('E1').getValue();
      const n = Number(e1);
      if (!isNaN(n) && n > 0) tempoSessao = n;
    } catch (e) {}

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      return { success: false, erro: 'Nenhum usuário cadastrado na aba CADASTRO.' };
    }

    const data = sheet.getRange(1, 1, lastRow, 3).getValues();
    const usuarioNorm = String(usuario || '').trim().toLowerCase();
    const senhaNorm = String(senha || '').trim();

    for (const row of data) {
      const u = String(row[0] || '').trim().toLowerCase();
      const s = String(row[1] || '').trim();
      const nivel = String(row[2] || '').trim().toUpperCase();
      if (u && u === usuarioNorm && s === senhaNorm) {
        return {
          success: true,
          nivel: nivel === 'PARCIAL' ? 'PARCIAL' : 'TOTAL',
          tempoSessao: tempoSessao
        };
      }
    }

    return { success: false, erro: 'Usuário ou senha incorretos.' };
  } catch (e) {
    Logger.log('❌ autenticarLogin: ' + e.message);
    return { success: false, erro: 'Erro ao autenticar: ' + e.message };
  }
}

/**
 * Retorna lista de todos os usuários cadastrados na aba CADASTRO.
 * Usado pelo filtro de usuários no relatório de faturamento.
 */
function listarUsuariosCadastrados() {
  try {
    const sheet = getSpreadsheet_().getSheetByName('CADASTRO');
    if (!sheet) return { success: false, usuarios: [] };
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return { success: true, usuarios: [] };
    const data = sheet.getRange(1, 1, lastRow, 1).getValues();
    const usuarios = data
      .map(row => String(row[0] || '').trim())
      .filter(u => u.length > 0);
    return { success: true, usuarios };
  } catch (e) {
    Logger.log('❌ listarUsuariosCadastrados: ' + e.message);
    return { success: false, usuarios: [] };
  }
}

/**
 * Retorna o nível atual de um usuário sem precisar da senha.
 * Usado para re-validação em tempo real a cada 30s.
 */
function obterNivelUsuario(usuario) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName('CADASTRO');
    if (!sheet) return { success: false };

    let tempoSessao = 15;
    try {
      const e1 = sheet.getRange('E1').getValue();
      const n = Number(e1);
      if (!isNaN(n) && n > 0) tempoSessao = n;
    } catch (e) {}

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return { success: false };

    const data = sheet.getRange(1, 1, lastRow, 3).getValues();
    const usuarioNorm = String(usuario || '').trim().toLowerCase();

    for (const row of data) {
      const u = String(row[0] || '').trim().toLowerCase();
      const nivel = String(row[2] || '').trim().toUpperCase();
      if (u && u === usuarioNorm) {
        return {
          success: true,
          nivel: nivel === 'PARCIAL' ? 'PARCIAL' : 'TOTAL',
          tempoSessao: tempoSessao
        };
      }
    }
    return { success: false };
  } catch (e) {
    Logger.log('❌ obterNivelUsuario: ' + e.message);
    return { success: false };
  }
}

/**
 * Registra login/logout/expiração na aba Log_Acessos.
 * Cria a aba e o cabeçalho automaticamente se não existirem.
 * tipo: 'LOGIN' | 'LOGOUT' | 'EXPIROU'
 */
function registrarAcesso(usuario, tipo) {
  try {
    const ss = getSpreadsheet_();
    const ABA = 'Log_Acessos';
    let sheet = ss.getSheetByName(ABA);

    if (!sheet) {
      sheet = ss.insertSheet(ABA);
      sheet.getRange(1, 1, 1, 3).setValues([['DATA/HORA', 'USUÁRIO', 'TIPO']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 140);
      sheet.setColumnWidth(3, 90);
    }

    const ts = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([ts, String(usuario || '').trim(), String(tipo || '').trim().toUpperCase()]);
  } catch (e) {
    Logger.log('❌ registrarAcesso: ' + e.message);
  }
}
