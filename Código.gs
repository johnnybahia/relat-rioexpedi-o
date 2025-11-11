
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

// ====== GERAR IDs ======

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

function gerarIdsFaltantes() {
  Logger.clear();
  Logger.log("=== GERANDO IDs COMPOSTOS ===");

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

// ====== SINCRONIZAÇÃO ======
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
      // Lê 15 colunas: A-O (ID até Status)
      // Status está na coluna O (índice 14 do array)
      dbData = dbSheet.getRange(2, 1, dbRows, 15).getValues();
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

    // 3) PROCESSAR
    Logger.log("\n🔄 3. PROCESSANDO");
    
    let novos = [];
    let updates = [];
    let marcaInativos = [];
    
    for (let [id, dbItem] of dbMap.entries()) {
      const statusAtual = dbItem.row[STATUS_COL];  // Coluna O (índice 14)
      if (statusAtual === "Excluido") continue;

      if (fonteMap.has(id)) {
        Logger.log(`   🔄 Match encontrado: ID="${id}" existe em PEDIDOS e Relatorio_DB`);
        const fonteRow = fonteMap.get(id);

        // Array de 15 elementos (índices 0-14)
        // Última posição (14) é o Status na coluna O
        const novaLinha = [
          fonteRow[ID_COL],      fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
          fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
          fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
          fonteRow[QTD_COL],     fonteRow[OS_COL],      fonteRow[DTREC_COL],
          fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   ""
        ];

        let mudou = false;
        // Compara as 14 primeiras colunas (0-13), excluindo Status
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
      } else {
        Logger.log(`   ❌ ID="${id}" existe em Relatorio_DB mas NÃO encontrado em PEDIDOS (com CARTELA preenchida)`);
        Logger.log(`      Status atual: "${statusAtual}", Linha: ${dbItem.linha}`);
        Logger.log(`      Motivo: ID não existe em PEDIDOS OU existe mas sem CARTELA preenchida`);
        if (statusAtual !== "Faturado" && statusAtual !== "Inativo") {
          Logger.log(`   ⚠️ Será marcado como Inativo`);
          marcaInativos.push({ linha: dbItem.linha, id: id, de: statusAtual });
        } else {
          Logger.log(`   ℹ️ Não será alterado (já é ${statusAtual})`);
        }
      }
      
      fonteMap.delete(id);
    }
    
    // Novos itens que estão em PEDIDOS mas não em Relatorio_DB
    for (let [id, fonteRow] of fonteMap.entries()) {
      Logger.log(`   🆕 Novo item: ID="${id}" está em PEDIDOS mas não em Relatorio_DB - será adicionado como Ativo`);
      Logger.log(`      CARTELA="${fonteRow[CARTELA_COL]}", CLIENTE="${fonteRow[CLIENTE_COL]}", OC="${fonteRow[OC_COL]}"`);

      // Array de 15 elementos, Status (índice 14) = "Ativo"
      const novaLinha = [
        fonteRow[ID_COL],      fonteRow[CARTELA_COL], fonteRow[CLIENTE_COL],
        fonteRow[PEDIDO_COL],  fonteRow[CODCLI_COL],  fonteRow[MARFIM_COL],
        fonteRow[DESC_COL],    fonteRow[TAM_COL],     fonteRow[OC_COL],
        fonteRow[QTD_COL],     fonteRow[OS_COL],      fonteRow[DTREC_COL],
        fonteRow[DTENT_COL],   fonteRow[PRAZO_COL],   "Ativo"
      ];
      novos.push(novaLinha);
    }
    
    Logger.log(`   🆕 Novos: ${novos.length}`);
    Logger.log(`   📝 Atualizar: ${updates.length}`);
    Logger.log(`   ⚠️ Marcar Inativo: ${marcaInativos.length}`);
    
    // 4) APLICAR
    Logger.log("\n💾 4. APLICANDO");
    if (novos.length > 0) {
      const proxLinha = dbSheet.getLastRow() + 1;
      dbSheet.getRange(proxLinha, 1, novos.length, 15).setValues(novos);
      Logger.log(`   ✅ ${novos.length} novos adicionados`);
    }
    if (updates.length > 0) {
      updates.forEach(u => {
        dbSheet.getRange(u.linha, 1, 1, 15).setValues([u.dados]);
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
    if (novos.length > 0 || updates.length > 0 || marcaInativos.length > 0) {
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
    Logger.log(`   • ${novos.length} novos itens adicionados ao Relatorio_DB como Ativo`);
    Logger.log(`   • ${updates.length} itens atualizados no Relatorio_DB`);
    Logger.log(`   • ${marcaInativos.length} itens marcados como Inativo (não encontrados em PEDIDOS)`);
    if (semId > 0) Logger.log(`   ⚠️ ${semId} linhas em PEDIDOS sem ID (ignoradas)`);
    if (semCartela > 0) Logger.log(`   ⚠️ ${semCartela} linhas em PEDIDOS sem CARTELA (ignoradas)`);
    Logger.log("=".repeat(70));
    
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

    Status: getDisp('Status', 'Desconhecido')
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
