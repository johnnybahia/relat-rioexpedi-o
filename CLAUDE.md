# CLAUDE.md — Sistema de Relatório de Pedidos / Expedição

> Lido automaticamente em toda sessão. Atualizar sempre que houver mudança arquitetural.
> Versão do sistema: **v15.6-SINCRONIZACAO** · Backend em Google Apps Script · Frontend em HTML+JS embutido no Apps Script
> Última revisão contra o código e dados reais: **28/09/2026**. Se este documento divergir do código, o código
> descreve o comportamento real — mas as **regras de negócio da seção 1.1 prevalecem**: se o código as viola, é o
> código que deve ser corrigido. Antes de propor mudanças, ler 1.1, 15.16–15.18 e 17 (plano pendente e ideias descartadas).

---

## 1. VISÃO GERAL

Sistema enterprise de gestão de pedidos de expedição em Google Sheets.  
Fluxo principal: **Fonte externa → DADOS_IMPORTADOS → PEDIDOS → Relatorio_DB → UI (index.html)**

Arquivos do projeto:
- `Código.gs` — todo o backend (~6600 linhas, único arquivo GAS)
- `index.html` — frontend servido via `HtmlService` (~6100 linhas)
- `RELATÓRIO DE PEDIDOS EXPEDIÇÃO CEARÁ - *.csv` — snapshots **antigos** das abas; não refletem o estado atual.
  Para diagnosticar dados, pedir ao usuário um export novo de DADOS_IMPORTADOS, PEDIDOS e Relatorio_DB.
- `executa1.py`, `OC_482500*.PDF` — extrator avulso de PDFs de pedidos que envia a outro web app;
  não faz parte deste sistema (o backend não tem `doPost`).

Planilha de origem (externa, somente leitura): ID `1GtYG4Ahy5XJyJjE37S27u8RyELdRkct8nDAVGIBRI-w`,
aba `RELATÓRIO GERAL DA PRODUÇÃO1` (constantes `SOURCE_ID`/`SOURCE_SHEET` em `importarDadosExternos`).

## 1.1 REGRAS DE NEGÓCIO (confirmadas pelo usuário em 28/09/2026 — prevalecem sobre o código)

1. **Rotina diária:** o usuário dá baixa (pode ser parcial — pedido de 1000, fatura 100, o saldo de 900 fica
   para outro dia), marca para faturar e imprime o relatório de faturamento (FAT-xxx). Quando a linha sai da
   planilha de origem, o sistema passa os itens **marcados** para `Faturado`. O faturado fica visível no HTML
   até a limpeza diária, para dar tempo de imprimir as etiquetas.
2. **Linha que some da origem = faturada ou cancelada** (num processo correto).
3. **Nenhum item vira `Faturado` sem usuário identificado:** só pela marcação do usuário (col V
   `MARCAR_FATURAR_USUARIO` preenchida) ou por confirmação explícita de um usuário na tela.
   QTD=0 sozinha não basta; inferência automática ("saiu da fonte") não basta.
   ⚠️ Em 28/09/2026 o código ainda viola esta regra em 5 caminhos — ver 15.16 e o plano da seção 17.
4. **A planilha de origem não pode ser alterada** (nenhuma coluna nova, nenhum código gravado nela).
   Portanto não existe identificador fixo por linha vindo da origem — ver 15.16.
5. **Limpeza diária:** todo dia útil, no horário da célula B2 da aba `CONFIGURAÇÕES`, apagar **todos** os
   `Faturado` e `Excluido` do Relatorio_DB. **Sem histórico de 15 dias.** `Finalizado` **não** é apagado
   (em 28/09 não havia nenhum: o botão "Finalizar" existe no código da tela, mas `handleFinalizar` nunca é chamado).
   ⚠️ Em 28/09/2026 o código ainda apaga só os com 15+ dias (`DIAS_RETENCAO`) — ver 15.18.
6. **Acionadores em produção:** `processoImportacao` a cada **15 min**; `processoAutomaticoCompleto` a cada
   **1 h**. O instalador do código cria 5 min e desfaz isso — ver seção 13.

---

## 2. ABAS DA PLANILHA

| Aba | Papel | Observação |
|---|---|---|
| `DADOS_IMPORTADOS` | Intermediária com dados da fonte externa | Atualizada via `importarDadosExternos()`. Célula H2 = timestamp (guard de sync) |
| `PEDIDOS` | Dados sincronizados e enriquecidos | IDs gerados aqui. Col A = ID_UNICO. Dados começam linha 4 |
| `Relatorio_DB` | Banco de dados final usado pela UI | 25 colunas (A-Y). Cabeçalhos em `RELATORIO_DB_HEADERS`; coluna nova entra sozinha via `_garantirHeadersRelatorio_DB_` |
| `CONFIGURAÇÕES` | Configuração editável pelo usuário | B2 = hora (0-23) da limpeza diária, padrão 11. A3 é só texto, gravado na criação da aba. **Não existe célula de dias** (15.18) |
| `CADASTRO` | Login | Col A usuário, B senha, C nível (TOTAL/PARCIAL); E1 = minutos de sessão |
| `Baixas_Historico` | Histórico de parcializações de QTD | Cabeçalho lido dinamicamente. Suporta coluna TIPO para CHECKPOINTs |
| `LOTE DILLY` | Mapeamento OC→Lotes para cliente Dilly | Consumido em FIFO durante sincronização |
| `original` | Dados originais para ordenação | Define posição dos itens dentro de uma OC |
| `Duplicatas_Debug` | Auditoria de itens descartados | Criada automaticamente |
| `SENHA` | Senha para confirmar alertas (cel A2) | Só usada por `confirmarAlerta()` — sistema de alertas desligado (15.8) |

---

## 3. COLUNAS — PEDIDOS (0-based)

```
0  A  ID_UNICO              ← gerado por sincronizarPedidosComFonte
1  B  CARTELA               (OK / vazio)
2  C  CLIENTE
3  D  CÓD. FILIAL           ← esta coluna NÃO existe no Relatorio_DB (gap)
4  E  PEDIDO
5  F  CÓD. CLIENTE          (normalizado para Dilly)
6  G  CÓD. MARFIM           (normalizado para Dilly)
7  H  DESCRIÇÃO + [UUID]    ← UUID embarcado como âncora de identidade
8  I  TAMANHO
9  J  ORD. COMPRA           ← displayValues preserva sufixos tipo "82249D"
10 K  QTD. ABERTA           ← autoridade de QTD (fonte)
11 L  CÓD. OS / LOTE        (Dilly: substituído pelo Lote da aba LOTE DILLY) ← ÚNICO campo que separa
                            linhas-irmãs; vazio ou "0" em 25% das linhas da fonte (15.16)
12 M  DATA RECEB.
13 N  DT. ENTREGA
14 O  PRAZO (texto)         ← contagem de dias: muda sozinho todo dia
15 P  TIMESTAMP_CRIACAO
16 Q  POSICAO_FONTE         (índice em DADOS_IMPORTADOS, fixo)
17 R  (reservado)
18 S  CODIGO_FIXO           (UUID imutável por item)
19 T  INFO_X                (col X da fonte)
20 U  LOTE                  (col Y da fonte) ← NÃO entra em ID, impressão digital nem identidade;
                            só é copiado. Preenchido em só ~15% das linhas; há valores repetidos
```

## 4. COLUNAS — Relatorio_DB (0-based)

```
0  A  ID_UNICO
1  B  CARTELA
2  C  CLIENTE
3  D  PEDIDO               ← DB_PEDIDO_COL=3 (sem col Filial → índices menores que PEDIDOS)
4  E  CÓD. CLIENTE
5  F  CÓD. MARFIM
6  G  DESCRIÇÃO
7  H  TAMANHO
8  I  ORD. COMPRA
9  J  QTD. ABERTA          ← DB_QTD_COL=9 (pode diferir de PEDIDOS por baixas)
10 K  CÓD. OS
11 L  DATA RECEB.
12 M  DT. ENTREGA
13 N  PRAZO
14 O  Status               ← Ativo | Inativo | Faturado | Finalizado | Excluido
                            Item zerado aguardando faturar = `Ativo` com QTD 0 (marcado ou não) — NÃO é Finalizado.
                            `Finalizado` só vem da ação manual `finalizarItem` (:5892); a tela não exibe o
                            botão (`handleFinalizar` nunca é chamado). O sync trata Finalizado como Faturado.
15 P  MARCAR_FATURAR       ← "SIM" | "" — marcação para emissão de NF
16 Q  DATA_STATUS          (data da última mudança de status)
17 R  POSICAO_FONTE
18 S  CODIGO_FIXO
19 T  INFO_X
20 U  LOTE
21 V  MARCAR_FATURAR_USUARIO ← JSON {"PERFIL":"USUARIO"}, ex.: {"BAIXA1":"EVELINE"}. É a prova de que
                             um usuário marcou (regra 1.1.3). Apagada no reset de ciclo
22 W  LOTE_EMISSAO         ← "FAT-001", "FAT-002"...
23 X  PERFIL_EMISSAO       ← perfil de baixa (BAIXA1…BAIXA5)
24 Y  SEQUENCIA            ← posição fixa do item dentro da OC (aba "original"); DB_SEQUENCIA_COL=24
```

## 5. COLUNAS — Baixas_Historico

```
ID_ITEM | DATA_HORA | QTD_BAIXADA | QTD_RESTANTE | QTD_ORIGINAL | USUARIO | TIPO
```
- `TIPO` = `"CHECKPOINT"` marca início de novo ciclo de faturamento
- `QTD_ORIGINAL` = `QTD_RESTANTE + QTD_BAIXADA` (valor antes desta baixa)
- Cabeçalho lido dinamicamente a cada operação — colunas podem estar em qualquer ordem

---

## 6. CONSTANTES CRÍTICAS

```javascript
FONTE_DATA_START_ROW = 4        // dados em PEDIDOS/DADOS_IMPORTADOS começam linha 4
DB_QTD_COL     = 9              // QTD. ABERTA no Relatorio_DB
MARCAR_FATURAR_COL = 15         // col P
MARCAR_FATURAR_USUARIO_COL = 21 // col V
LOTE_EMISSAO_COL = 22           // col W
STATUS_COL = 14                 // col O
DATA_STATUS_COL = 16            // col Q
PEDIDOS_LOTE_COL = DB_LOTE_COL = 20 // col U (LOTE, col Y da fonte)
DB_SEQUENCIA_COL = 24           // col Y
DIAS_RETENCAO = 15              // Faturado/Finalizado/Excluido com 15+ dias são apagados na limpeza diária
                                // (fixo no código; a regra 1.1.5 pede apagar todos — ver 15.18)
CONFIG_SHEET_NAME = 'CONFIGURAÇÕES', CONFIG_HORA_LIMPEZA_CELL = 'B2', CONFIG_HORA_LIMPEZA_PADRAO = 11
MIN_LINHAS_FONTE_PARA_FATURAR = 50 // travas do faturamento automático (7.3) — só cobrem esse caminho
MAX_FRACAO_SEM_FONTE = 0.5
TOLERANCIA_QUEDA_FONTE = 0.15
CACHE_DURATION = 600            // segundos (CacheService)
TZ = 'America/Fortaleza'
```

---

## 7. FLUXO DE SINCRONIZAÇÃO (passo a passo)

### 7.1 Importação (`processoImportacao` — produção: a cada 15 min)
1. Lê planilha externa via `openById()` + SOURCE_SHEET
2. Coluna J (CÓD. CLIENTE): usa `getDisplayValues()` para preservar sufixos "82249D", "14660U"
3. **Apaga (`clearContent`) e regrava `DADOS_IMPORTADOS` inteira** — é um espelho da origem; nada gravado
   nessa aba sobrevive à próxima importação
4. Atualiza H2 com timestamp (sinaliza ao próximo passo que há dados novos)
5. Agenda `processoAutomaticoCompleto` via `_agendarSincronizacao_(90)` para 90s depois (Código.gs:2281)
- ⚠️ **Não usa LockService.** Pode regravar a aba no meio de um sync (ver 13 e 15.17).
- Coluna A da origem ("Column 24") é uma numeração contínua 1…N, renumerada — **não serve como ID**.

### 7.2 Sync DADOS_IMPORTADOS → PEDIDOS (`sincronizarPedidosComFonte`)
- **Guard de H2**: se H2 igual ao último processado → aborta (sem retrabalho)
- Pré-carrega IDs do DB + fingerprints (evita colisões)
- Lê aba `original` para calcular `POSICAO_FONTE`
- Lê `LOTE DILLY` para enriquecimento de CÓD. OS (Dilly)
- Para cada linha em DADOS_IMPORTADOS (uma a uma, na ordem da aba — matching **guloso**):
  - Cria fingerprint: `CLIENTE|PEDIDO|PRODUTO|TAM|OC|OS|DATA_NORM`, onde **PRODUTO = DESCRIÇÃO** sem o
    sufixo `[uuid]` (`_chaveProduto_`, Código.gs:2780); CÓD. MARFIM só é usado se a DESCRIÇÃO estiver vazia.
    LOTE e QTD não entram.
  - Tentativas: `pedidosMap` (fingerprint, contra o PEDIDOS anterior) → `pedidosDillyMap` (Dilly, sem OS)
    → `dbFingerprintMap` (FIFO) → ID novo (próximo sufixo livre + UUID novo)
  - Entre candidatos com a mesma fingerprint (linhas-irmãs), desempata pela QTD mais próxima da QTD
    **anterior em PEDIDOS** (Código.gs:1777); empate → primeiro da lista. Ignora LOTE e o estado de
    baixa no DB → IDs podem migrar entre linhas-irmãs (15.16)
  - Normaliza MARFIM e CÓD. CLIENTE para Dilly
  - Para Dilly: substitui CÓD. OS pelo Lote da fila FIFO
- **Guarda anti-duplicata final** (ver 15.13): `idsAssignadosNestaRodada` detecta se o
  `idFinal` resolvido já foi usado por outra linha NESTA execução e renomeia
  (`-DUP{n}`) antes de escrever
- Escreve resultado em PEDIDOS

### 7.3 Sync PEDIDOS → Relatorio_DB (`sincronizarDados`)
- Garante headers do DB (`_garantirHeadersRelatorio_DB_`)
- Cria mapas: `fonteMap`, `dbMap`, impressões digitais, `codigoFixoMap`
- Lê `DADOS_IMPORTADOS` **de novo, direto da aba** (Código.gs:3115) para o índice de identidade `fonteIdent`
  (`CLIENTE|PEDIDO|PRODUTO|TAM`) e a contagem OC+OS — pode ser uma versão diferente da usada para montar
  PEDIDOS se uma importação rodar no meio (15.17)
- Para cada item do DB, tenta match por: **ID → UUID (CODIGO_FIXO) → fingerprint → troca de OC**
  (troca de OC só é aceita com exatamente 1 candidato livre, Código.gs:3564)
- **Executado em DUAS PASSADAS** (ver 15.12): Passada 1 resolve ID e UUID para
  todos os itens do DB antes de qualquer fingerprint rodar; Passada 2 resolve
  fingerprint e "não encontrado" só para quem sobrou sem match na Passada 1.
- Se o item do DB não casou com nada (saiu de PEDIDOS) — comportamento real em 28/09/2026 (commit `2ceb277`):
  - marcado pelo usuário (MARCAR_FATURAR=SIM **e** col V preenchida) → `Faturado` direto
    (⚠️ sem checar as travas de importação quebrada)
  - QTD=0 → `Faturado` silencioso (⚠️ sem usuário, sem travas — viola 1.1.3)
  - QTD>0 + `_motivoSaidaDaFonte_` + `autoFaturarSaidaFonte` → `Faturado` direto (⚠️ viola 1.1.3).
    `_motivoSaidaDaFonte_` = orçamento de excedente por identidade (linhas abertas no DB − linhas na
    fonte), consumido na ordem de iteração do DB — não sabe QUAL linha-irmã saiu.
    `autoFaturarSaidaFonte` = 4 travas: mínimo de linhas, fração máxima sem fonte, queda brusca da
    fonte, importação parcial (Código.gs:3243)
  - MARCAR_FATURAR=SIM sem usuário (legado de versões antigas) → mantém Ativo
  - senão → não faz nada ("aguarda reconsolidação")
  - `protecaoAtiva` (OC+OS) hoje **só aparece no log** — não decide nada
- O sync **não** grava mais MARCAR_FATURAR="SIM" automaticamente; alertas estão desligados (15.8)
- Adiciona itens novos de PEDIDOS que não existem no DB (travas contra duplicata por fingerprint e por
  identidade sem OC — ambas usam DESCRIÇÃO, Código.gs:3792/3825)
- Atualiza IDs no Baixas_Historico se IDs mudaram
- Grava cada linha alterada **inteira**, com os valores lidos no início do sync (Código.gs:3958) — ação
  de usuário feita durante o sync pode ser sobrescrita (15.17)
- A limpeza de faturados **não** roda aqui: é a ETAPA 3 do processo completo (7.4 e 15.18)

### 7.4 Processo Completo (`processoAutomaticoCompleto` — produção: a cada 1 h + 90s após cada importação)
1. Verifica pausa (`_sistemaPausado_()`)
2. Adquire `LockService` (`waitLock(30000)`; se não conseguir, aborta a execução)
3. ETAPA 0 `sincronizarPedidosComFonte(false)` (guard H2) → ETAPA 1 `verificarEGerarIDs` →
   ETAPA 2 `sincronizarDadosOtimizado` → ETAPA 3 limpeza de faturados (só se `_deveLimparFaturadosAgora_()`)
   → ETAPA 4 limpar cache
- **Não importa a origem** — quem importa é só `processoImportacao`.

---

## 8. SISTEMA DE BAIXAS (PARCIALIZAÇÕES)

### Propósito
Registrar entregas parciais sem descartar o item (que ainda tem saldo em aberto).

### Operações principais

| Função | O que faz |
|---|---|
| `aplicarBaixa(id, linha, qtd, user)` | Reduz QTD.ABERTA no DB + chama `registrarBaixa`. **Se `registrarBaixa` falhar, reverte o DB** |
| `registrarBaixa(id, baixada, restante, user)` | Adiciona linha no Baixas_Historico |
| `estornarBaixa(id, linhaDB, linhaHist, qtd)` | Remove linha do histórico + restaura QTD no DB |
| `editarUltimaBaixa(id, linhaDB, novaQtd, user)` | Edita a última entrada do histórico + recalcula QTD no DB |
| `obterHistoricoSessaoBaixas(id)` | Retorna baixas **após o último CHECKPOINT** (sessão atual) |
| `obterHistoricoBaixas(id)` | Retorna todas as baixas (exclui CHECKPOINTs) |

### Caches (module-level, limpas a cada operação)
- `_getSaldoEfetivoCache_()` — soma de `QTD_BAIXADA` desde o último CHECKPOINT por item
- `_getUltimaQtdOriginalCache_()` — `QTD_ORIGINAL` da última entrada não-CHECKPOINT após o último CHECKPOINT
- `calcularQtdOriginal(id, qtdAtual)` — retorna do cache ou fallback `qtdAtual`

### Checkpoints de Faturamento
- `TIPO="CHECKPOINT"` no Baixas_Historico
- Criado por `_registrarCheckpointFaturamento_(id, qtdAberta)`
- Quando ocorre:
  - Item faturado com QTD.ABERTA > 0 (faturamento parcial)
  - Reset de ciclo (QTD da fonte mudou durante ciclo de baixas)
- Efeito: próximas baixas só contam após este ponto

### Reset de Ciclo
```javascript
resetCiclo = temBaixas && pedidosQtd > 0
  && baselineId !== undefined && pedidosQtd !== baselineId
  && status não é Faturado/Finalizado
```
Efeito: registra CHECKPOINT, limpa MARCAR_FATURAR, MARCAR_FATURAR_USUARIO (col V) e LOTE_EMISSAO.
É o que acontece no faturamento parcial do dia a dia: a origem reduz a QTD depois da NF → novo ciclo.
Consequência: item marcado cuja QTD muda na origem perde a marcação (precisa ser marcado de novo).

---

## 9. SISTEMA DE FATURAMENTO

### Estados de MARCAR_FATURAR
- `""` (vazio) — item não marcado
- `"SIM"` + col V preenchida — marcado pelo usuário para emitir NF
- `"SIM"` sem col V — legado de versões antigas do sync; `limparMarcacoesSemUsuario()` apaga ao gerar relatório

### Quem define MARCAR_FATURAR="SIM"
- Só o **usuário**, via `marcarParaFaturar()` (checkbox ou fluxo "baixa total + marcar"). O sync não marca mais.

### Caminhos que gravam `Faturado` (auditados em 28/09/2026)
| Caminho | Tem usuário? |
|---|---|
| Sync: marcado pelo usuário + saiu da origem (Código.gs:3698) | Sim (col V) — mas sem as travas de importação |
| Sync: QTD=0 + saiu da origem (:3702) | **Não** |
| Sync: `_motivoSaidaDaFonte_` (:3705) | **Não** |
| Menu `faturarItensForaDaFonte()` (:5310) | **Não** (mesma lógica do sync, em lote) |
| Menu `confirmarTodosAlertas()` "(testes)" (:6431) | **Não confere** col V; ainda apaga LOTE_EMISSAO |
| `marcarFaturado()` (:5803) | **Não** — nenhuma tela chama, mas é pública (dá para chamar do navegador) |

Regra 1.1.3 exige que só o primeiro caminho (e uma confirmação explícita na tela) existam — plano na seção 17.

### Cálculo do SALDO no Relatório
```javascript
const qtdOriginal    = item['QTD. ORIGINAL'] || 0;  // de calcularQtdOriginal()
const qtdAberta      = item['QTD. ABERTA']   || 0;
const saldoCalculado = qtdOriginal - qtdAberta;
// Fallback: se não há baixas, qtdOriginal == qtdAberta → saldoCalculado == 0
// Neste caso usa QTD.ABERTA para não zerar o relatório
const saldo = saldoCalculado > 0 ? saldoCalculado : qtdAberta;
```
**Por que o fallback existe:** quando nenhuma baixa foi registrada (ex: usuário marcou sem dar baixa), `calcularQtdOriginal` retorna `qtdAberta` como fallback → saldo seria 0 sem esta correção.

### Fluxo de Geração do Relatório
1. Usuário clica "Gerar Relatório de Faturamento"
2. `obterItensMarcadosParaFaturar()` → lê todos com MARCAR_FATURAR="SIM"
3. Filtro por usuário → filtro por marca/INFO_X
4. `gerarNumeroLoteEmissao()` → "FAT-001", "FAT-002"...
5. `_gerarRelatorioFaturamentoPDF()` → abre janela de impressão
6. `registrarLoteEmissao()` → grava LOTE_EMISSAO (col W) nos itens

### Proteção Anti-Faturamento Indevido (OC+OS)
- Hoje **só é logada** (`protecaoAtiva`, Código.gs:3656) — não bloqueia nada. Quem decide o faturamento
  automático é o orçamento por identidade + as 4 travas (7.3).
- As 4 travas cobrem **só** o caminho `_motivoSaidaDaFonte_`: numa importação quebrada, itens marcados e
  itens com QTD=0 viram `Faturado` mesmo assim.

---

## 10. SISTEMA DE IDs

### Estrutura do ID
```
{CLIENTE}{FILIAL}{PEDIDO}{MARFIM}{TAMANHO}{OC}{OS}{DATA_yyyyMMdd}-{SUFIXO_NUMÉRICO}
```
Exemplo: `MARFIM MINAS7413040480114120CM179202520250917-1`

### Campos fora da string do ID ≠ campos que podem mudar à vontade
- Fora do ID: CARTELA, DESCRIÇÃO, QTD, LOTE (col U), DT. ENTREGA, PRAZO.
- ⚠️ **A DESCRIÇÃO não está no ID, mas é a chave de produto de TODAS as correspondências**
  (`_chaveProduto_`: impressão digital, identidade, recuperação pelo DB, trava de novos itens).
  Alterar a DESCRIÇÃO de um item na origem gera ID + UUID novos, deixa a linha antiga órfã (hoje →
  `Faturado` sem usuário, ou linha `Ativo` fantasma se já estava faturada) e insere uma duplicata.
  Qualquer diferença conta: espaço extra, maiúscula/minúscula, acento.
- Campos que identificam o item (mudar = item "novo"): CLIENTE, PEDIDO, DESCRIÇÃO, TAMANHO,
  ORD. COMPRA (troca de OC só é reconhecida com 1 candidato), CÓD. OS, DATA RECEB.
- Mudar LOTE, QTD, DT. ENTREGA, PRAZO ou CARTELA é seguro (só atualiza) — exceto entre linhas-irmãs,
  onde QTD é o critério de desempate e uma edição pode trocar IDs (15.16).
- Nem a origem (coluna A é renumerada) nem o CÓD. CLIENTE (é do artigo: 279 valores para 2.259 linhas)
  servem como ID de linha.

### UUID Imutável (CODIGO_FIXO)
- Gerado via `Utilities.getUuid()` na primeira sincronização
- Gravado em col S (PEDIDOS) e col 18 (DB)
- **Também embarcado na DESCRIÇÃO** como `[UUID]` — backup se colunas forem apagadas
- Permite recuperar histórico de baixas mesmo se ID mudar

### Recuperação de ID (ordem de prioridade)
1. Match exato de ID em PEDIDOS
2. Match por UUID (CODIGO_FIXO)
3. Match por fingerprint (impressão digital)
4. Geração de novo ID com sufixo numérico sequencial

---

## 11. REGRAS ESPECIAIS — DILLY

Detecção: `cliente.toUpperCase().includes('DILLY')`

### Normalização de CÓD. MARFIM e CÓD. CLIENTE
```javascript
// Se Dilly + marfim tem "-" + sufixo ≥ 2 chars:
// Substitui sufixo pelo número extraído do TAMANHO
"196338-120" + "110CM" → "196338-110"   // ✓ corrige
"7490-1"     + "100CM" → "7490-1"       // ✗ protegido (sufixo curto)
```
Operação idempotente — aplicar duas vezes = mesmo resultado.

### CÓD. OS via LOTE DILLY
- Aba `LOTE DILLY`: chave = `OC|CÓD_CLIENTE_BASE|TAMANHO_NUMÉRICO|QTD` → [Lotes] (FIFO)
- Durante sync: cada item Dilly consome (`shift()`) um Lote da fila
- Sobrescreve col L (CÓD. OS) com o Lote

### Fingerprint Dilly
- O produto na impressão digital é a DESCRIÇÃO (`_chaveProduto_`); `_normalizarMarfimDilly_` só entra
  quando a DESCRIÇÃO está vazia (reserva)
- Garante que "196338-120" em DADOS_IMPORTADOS bate com "196338-110" em PEDIDOS nesse caso de reserva

---

## 12. REGRAS ESPECIAIS — DAKOTA

Detecção: `cliente.toUpperCase().includes('DAKOTA')`

- No relatório de faturamento: cada filial Dakota = grupo separado (pelo nome do cliente)
- Demais clientes: agrupados por marca (INFO_X / col T)

---

## 13. TRIGGERS E AUTOMAÇÃO

Configuração em produção informada pelo usuário em 28/09/2026 (tela Acionadores):

| Trigger | Função | Frequência em produção |
|---|---|---|
| `processoImportacao` | Importa a origem, atualiza H2, agenda o processo completo | **15 min** (minutos :02/:17/:32/:47 em 28/09) |
| `processoAutomaticoCompleto` | ETAPAS 0–4 (7.4) | **1 h** (minuto :33 em 28/09) |
| Único, criado por `_agendarSincronizacao_` | `processoAutomaticoCompleto` | 90s após cada importação bem-sucedida. Não aparece na lista de Acionadores — conferir em "Execuções" se está rodando |

- ⚠️ O menu **"2. Ativar Geração Automática (a cada 5 min)"** (`instalarTriggerAutomatico`) apaga os
  acionadores e recria **os dois a cada 5 min** — desfaz a configuração de produção. Idem
  `instalarTriggerAutomaticoSilencioso`. Não rodar sem ajustar os intervalos depois.
- **Cota diária:** conta @gmail.com tem **90 min/dia** de execução total de acionadores (Workspace: 6 h).
  São ~96 importações + até ~120 processos completos por dia — todo trabalho adicionado ao sync
  precisa ser leve (em memória, sem leitura/escrita por item).
- **Choque de horário:** o acionador de hora em hora (:33) dispara ~1 min depois de uma importação (:32);
  como a importação não usa trava, o sync pode ler DADOS_IMPORTADOS no meio da regravação.
- **Limpeza de faturados:** roda na primeira execução do processo completo a partir da hora de
  `CONFIGURAÇÕES!B2`, só em dias úteis — não no minuto exato (Apps Script também não garante minuto exato).
- Onde este documento ou o código falam em "próximo ciclo", leia até 1 h (ou ~15 min se o sync pós-importação rodar).

**Guard de execução dupla:** `LockService.getScriptLock().waitLock(30000)` — só no processo completo e em
algumas ferramentas de menu; importação e ações de usuário **não** usam trava.  
**Pause do sistema:** `PropertiesService.SISTEMA_PAUSADO='true'` (menu IDs Personalizados)

---

## 14. UI / FRONTEND (index.html)

### Chamadas ao backend (`google.script.run`)
```javascript
fetchAllDataUnified(cacheBuster)              // carrega todos os dados
marcarParaFaturar(id, linha, marcar, user)   // marca/desmarca para NF
aplicarBaixa(id, linha, qtd, user)           // registra baixa parcial
obterHistoricoSessaoBaixas(id)               // histórico da sessão atual
estornarBaixa(id, linhaDB, linhaHist, qtd)  // estorna baixa específica
editarUltimaBaixa(id, linhaDB, novaQtd, user)
obterItensMarcadosParaFaturar()             // para gerar relatório
gerarNumeroLoteEmissao()                    // FAT-001, FAT-002...
registrarLoteEmissao(itensData)             // persiste lotes em col W
```

### Fluxo de marcação para faturamento (checkbox)
```
handleMarcarFaturar(id, linha, true)
  ├─ qtdAberta <= 0       → _executarMarcarFaturar() direto (sem baixa)
  ├─ qtdAberta < qtdOriginal  → confirm("Faturar completo ou manter parcial?")
  │    ├─ OK      → _aplicarBaixaTotalEMarcar() → baixa total + marca
  │    └─ Cancelar → _executarMarcarFaturar() (mantém saldo parcial)
  └─ qtdAberta == qtdOriginal → confirm("Nenhuma baixa — confirmar total?")
       └─ OK → _aplicarBaixaTotalEMarcar() → baixa total + marca
```

### Modais principais
- `baixaModal` — baixa parcial: mostra histórico + campo para nova entrada
- `baixaConfirmacaoModal` — após baixa parcial: opção TOTAL (zera) ou PARCIAL (mantém)
- `alertaFaturamentoModal` — desligado desde `ce75202` (06/05/2026): `display:none !important` (15.8)
- `filtroUsuarioFaturModal` — filtra itens por usuário antes de imprimir
- `filtroMarcaFaturModal` — filtra por marca/INFO_X antes de imprimir

### Acesso Parcial
- `document.body.classList.add('acesso-parcial')` oculta botões de ação via CSS
- **Não bloqueia no backend** — segurança real deve ser implementada no GAS se necessário
  (`obterNivelUsuario(usuario)` já existe para validar o nível no servidor)

### Linha do DB enviada pela tela
- `aplicarBaixa`, `marcarParaFaturar`, `excluirItem`, `finalizarItem` localizam o item pelo número de linha
  (`planilhaLinha`) que a tela mandou, **sem conferir o ID** — ver 15.17.

---

## 15. ARMADILHAS E DECISÕES NÃO ÓBVIAS

### 15.1 Gap de coluna D (PEDIDOS vs Relatorio_DB)
PEDIDOS tem col D (CÓD. FILIAL) que **não existe no DB**.  
Resultado: `PEDIDO_COL=4` em PEDIDOS mas `DB_PEDIDO_COL=3` no DB.  
Toda função de fingerprint tem parâmetro `isDbRow` para usar o índice correto.  
**Usar o índice errado cria "novos itens" fantasmas a cada sync.**

### 15.2 Coluna J de PEDIDOS precisa de displayValues
`getValues()` retorna número cru; sufixos como "82249D" são perdidos.  
`getDisplayValues()` é usado especificamente para a col J.  
Se alguém formatar essa coluna como número, os sufixos desaparecem silenciosamente.

### 15.3 SALDO=0 no relatório de faturamento
Causa: `calcularQtdOriginal()` retorna `qtdAberta` como fallback quando não há baixas → `SALDO = qtdAberta - qtdAberta = 0`.  
Correção aplicada (branch `claude/fix-invoice-zero-amount-8slzu`): `saldo = saldoCalculado > 0 ? saldoCalculado : qtdAberta`.  
**Cuidado ao modificar `obterItensMarcadosParaFaturar`:** o fallback para `qtdAberta` é intencional.

### 15.4 registrarBaixa não é verificada (histórico)
A versão atual `aplicarBaixa` **reverte o DB** se `registrarBaixa` falhar.  
Versões anteriores não verificavam — deixavam QTD zerada no DB sem registro no histórico.

### 15.5 MARCAR_FATURAR_USUARIO protege desmarcação
Só o usuário que marcou pode desmarcar.  
`marcarParaFaturar(id, linha, false, usuarioAtual)` valida `usuarioQueMarkou === usuarioAtual`.

### 15.6 LOTE_EMISSAO ausente do payload de obterItensMarcadosParaFaturar
O item serializado retornado por `obterItensMarcadosParaFaturar` não inclui `LOTE_EMISSAO`.  
O relatório HTML usa `item.LOTE_EMISSAO` para detectar reemissões — como não existe, **todas as emissões aparecem como "novas"** (nunca mostra badge de reemissão).  
Bug pré-existente; correção requer adicionar `LOTE_EMISSAO: item.LOTE_EMISSAO || ''` ao objeto serializado.

### 15.7 Proteção anti-faturamento é proporcional (hoje só logada — ver seção 9)
Se há 3 itens com OC+OS idênticos em DADOS_IMPORTADOS e 2 ativos no DB:  
- Slots = 2; depois de 2 matches, o 3º item **não é mais protegido** (faturamento liberado).  
Necessário entender que a contagem decrementa a cada match bem-sucedido.

### 15.8 Alertas de faturamento estão desativados
`_registrarAlertaFaturamento_` apenas loga, não persiste em PropertiesService.  
`obterAlertasPendentes()` limpa e retorna `[]` sempre.  
Histórico: o commit `ce75202` (06/05/2026) removeu o aviso no login (`verificarAlertasFaturamento`, com senha
da aba SENHA) para itens que saíam sem marcação, para "não interromper o fluxo". A lista ficava numa única
propriedade do script (limite de 9 KB por valor). **Não reativar** — o substituto planejado é o aviso da
seção 17, com a pendência gravada numa coluna do DB e sem bloquear a tela.

### 15.9 IDs com sufixo numérico podem mudar
Se itens forem reordenados em DADOS_IMPORTADOS, o sufixo de um item pode mudar na próxima sync (ex: `-1` vira `-2`).  
O **UUID (CODIGO_FIXO)** garante que o histórico de baixas seja preservado mesmo assim — **exceto entre
linhas-irmãs**: o UUID acompanha a vaga escolhida no matching, então migra junto com o ID (15.16).

### 15.10 Dilly: Lote é consumido em FIFO
Se LOTE DILLY tiver menos Lotes que itens Dilly, os últimos itens ficam sem CÓD. OS.  
Verificar aba LOTE DILLY se itens Dilly aparecerem com CÓD. OS vazio.

### 15.11 Dilly: IDs instáveis causam QTD reverter após baixa (CORRIGIDO)
**Causa raiz:** O fingerprint padrão inclui o campo OS. Em DADOS_IMPORTADOS o OS é o valor
original da fonte; em PEDIDOS o OS é o Lote (da aba LOTE DILLY). Os fingerprints nunca batem
→ `sincronizarPedidosComFonte()` gerava um novo ID a cada sync para itens Dilly.

**Efeito cascata:** IDs oscilanando entre sufixos `-1` e `-2`. O vínculo com `Baixas_Historico`
dependia do mecanismo `idsAtualizados` (que atualiza IDs no histórico após cada sync). Em
cenários com Lote instável (reordenação de DADOS_IMPORTADOS, mudança de QTD no LOTE DILLY key),
a fingerprint mudava e o item "some" do PEDIDOS → novo item adicionado ao DB com QTD da fonte.

**Correção (branch `claude/fix-invoice-zero-amount-8slzu`):** `pedidosDillyMap` — mapa
secundário em `sincronizarPedidosComFonte()` com fingerprint **sem OS** (`Cliente|Pedido|Marfim|Tam|OC|Data`).
Quando o match padrão falha para um item Dilly, tenta este mapa. O ID e UUID existentes são
reutilizados → IDs estáveis → `idsBaixados.has(id)` sempre TRUE → QTD preservada.

**Não modificar o fingerprint padrão** (`_criarImpressaoDigitalFromRow_`): afetaria matching
em `sincronizarDados()`. A correção é localizada apenas em `sincronizarPedidosComFonte()`.

### 15.12 Matching cascade fabricava duplicatas ativamente (CORRIGIDO — DUAS PASSADAS)
**Causa raiz:** `sincronizarDados()` resolvia ID → UUID → fingerprint num único loop,
item por item. Quando dois itens do DB são "gêmeos" de fingerprint (mesmo
CLIENTE/PEDIDO/MARFIM/TAM/OC/OS/DATA — comum em Dilly, onde o mesmo Lote pode valer
para várias linhas idênticas) e só um deles ainda tem correspondência real em PEDIDOS,
a ORDEM DE ITERAÇÃO decidia o resultado: se o item SEM UUID válido fosse processado
antes do item COM UUID válido, a TERCEIRA TENTATIVA (fingerprint) "roubava" a vaga em
PEDIDOS antes do dono legítimo confirmar pela SEGUNDA TENTATIVA (UUID/CODIGO_FIXO).
Como a busca por UUID consulta `fonteCodigoFixoMap` diretamente — sem checar se a vaga
em `fonteMap`/`fonteImpressoes` já foi consumida — o item legítimo TAMBÉM encontrava o
mesmo `novoId` de forma independente.

**Efeito:** dois itens do Relatorio_DB eram gravados com o MESMO ID na MESMA execução de
sync (`updates.push` duplicado, sem dedupe no `updates.forEach(...setValues...)` final nem
na "VALIDAÇÃO ANTI-DUPLICATA", que só valida o array `novos`, não `updates`). Duplicata
fabricada na hora, não apenas herdada de execuções anteriores. Também conflava o histórico
de baixas de dois itens diferentes sob um único ID (via `idsAtualizados`/`aliasMap`).

**Correção:** o loop único foi separado em duas passadas.
- **Passada 1** — resolve PRIMEIRA TENTATIVA (ID) e SEGUNDA TENTATIVA (UUID) para
  TODOS os itens de `dbEntriesParaProcessar`. Quem não casa é adiado para `pendentes`.
- **Passada 2** — resolve TERCEIRA TENTATIVA (fingerprint) e "NÃO ENCONTROU" só para
  quem sobrou em `pendentes`.

Isso garante que toda reivindicação legítima por UUID já esgotou sua vaga em
`fonteMap`/`fonteImpressoes` ANTES de qualquer fingerprint rodar — eliminando a corrida.
**Cuidado ao modificar a cascade:** não fundir as duas passadas de volta em um loop único
sem reintroduzir esta corrida; qualquer nova tentativa de match adicionada à Passada 1 deve
resolver para TODOS os itens antes de qualquer lógica de Passada 2 ler `fonteMap`/`fonteImpressoes`.

### 15.13 PEDIDOS sem guarda anti-duplicata final (FECHADO — PREVENTIVO)
**Gap estrutural (não um bug confirmado em dados reais):** `sincronizarPedidosComFonte()`
mantém duas estruturas de matching independentes sobre o mesmo pool de linhas de PEDIDOS —
`pedidosMap` (fingerprint COM OS, caminho principal) e `pedidosDillyMap` (fingerprint SEM OS,
fallback só usado para clientes Dilly quando o caminho principal não encontra nenhum
candidato). Cada linha de PEDIDOS é envolvida em DOIS wrapper objects independentes (um por
mapa), cada um com sua própria flag `usado`. Em tese, duas linhas DIFERENTES de
DADOS_IMPORTADOS podiam reivindicar a MESMA linha de PEDIDOS via tiers diferentes na mesma
execução, sem que nenhum dos dois tivesse visibilidade do outro.

Diferente da escrita em Relatorio_DB (`sincronizarDados()`, que tem a "VALIDAÇÃO
ANTI-DUPLICATA" como rede de segurança final — ver 15.12), a escrita de `novasPedidosData`
em PEDIDOS não tinha NENHUMA validação final contra `ID_UNICO` duplicado antes do
`setValues()`. O Set `idsUsados` existente é pré-carregado com todos os IDs do
Relatorio_DB — serve para evitar colisão ao GERAR um ID novo, mas não detecta duas linhas
da mesma rodada que ambas "legitimamente" reutilizaram o mesmo ID existente.

PEDIDOS é o ponto mais upstream do pipeline (reescrito do zero a cada ciclo a partir de
DADOS_IMPORTADOS) — todo o resto (Relatorio_DB, Baixas_Historico) referencia itens por
`ID_UNICO` originado aqui. Garantir unicidade neste ponto é a defesa de maior alcance
possível.

**Correção:** novo Set `idsAssignadosNestaRodada`, vazio no início de cada execução,
populado a cada linha processada. Imediatamente antes do `idsUsados.add(idFinal)` final,
verifica se este `idFinal` já foi assinalado nesta mesma rodada — se sim, renomeia para
`"{idFinal}-DUP{n}"` até ser único contra ambos os Sets, e zera `codigoFixo` (força geração
de um UUID novo, já que o UUID antigo pertence à linha que reivindicou o ID original e não
deve ser herdado pela linha renomeada).

**Não foi observada nenhuma ocorrência confirmada deste cenário em dados reais** — esta é
uma defesa estrutural preventiva, não a correção de um bug já manifestado (diferente do
Bug 3 em 15.12).

### 15.14 Dilly com OS vazio: slot único esgotado fabricava duplicatas em massa (CORRIGIDO — OC 488457)
**Sintoma:** Relatorio_DB com centenas de cópias do mesmo item — todas `Ativo`, cada uma com
ID e UUID próprios, mesma `POSICAO_FONTE`. Observado: OC 488457 com 424 linhas no DB para
40 na fonte (405 órfãs) e OC 489062 com 171 órfãs. Cliente DILLY NORDESTE em ambas.

**Causa raiz (em `sincronizarPedidosComFonte`):** quando a fonte tem OS **vazio** (col L de
DADOS_IMPORTADOS) e a fila LOTE DILLY se esgota (15.10), sobra em PEDIDOS uma linha com
CÓD. OS vazio. O fingerprint padrão (COM OS) dessa linha (`...|OC||data`) é idêntico ao de
TODAS as linhas-irmãs da fonte (que têm OS vazio), enquanto as demais linhas de PEDIDOS têm
fingerprint com o Lote (`...|OC|26021696|data`) que nunca bate com a fonte. Resultado: o
grupo inteiro disputa UM único slot no `pedidosMap`. O código antigo, no caso
"`matches` existe mas todos usados", ia **direto para `isNovo=true`** — sem tentar o
fallback `pedidosDillyMap` (sem OS) nem a recuperação `dbFingerprintMap`, que só existiam no
branch "nenhum match". As linhas perdedoras ganhavam **ID + UUID novos a cada sync** (sufixos
crescendo: -96, -97, -111... até -161).

**Efeito cascata (em `sincronizarDados`):** os IDs novos de PEDIDOS não existiam no DB;
linhas antigas do DB eram renomeadas por fingerprint até esgotarem os slots, e quando o
FIFO do LOTE DILLY embaralhava os Lotes entre linhas-irmãs (fingerprint com OS mudava),
linhas de PEDIDOS sem contraparte no DB eram adicionadas como "novas" — acumulando
duplicatas. As órfãs nunca eram limpas: `consumedFingerprints` é um Set de fingerprints
(não de slots), então após 1 match o `fpDisponiveisDB` deixava de proteger o grupo inteiro.

**Correção (branch `claude/oc-488457-duplication-2i9h4q`):**
1. O fluxo de matching foi reestruturado: a seleção de slot do fingerprint padrão foi
   separada do tratamento; o branch de fallback (Dilly sem OS → `dbFingerprintMap` →
   `isNovo`) agora roda **sempre que não há slot disponível** — tanto "nenhum match" quanto
   "todos usados". Validado com dados reais: 2660/2660 linhas passam a reutilizar ID
   (antes: 48 linhas geravam ID novo a cada sync, exatamente nas OCs 488457/489062).
2. `pedidosMap` e `pedidosDillyMap` agora compartilham os **mesmos wrapper objects** (flag
   `usado` única por linha física de PEDIDOS) — uma linha reivindicada por um caminho fica
   indisponível para o outro na mesma rodada (fecha na origem o cenário do 15.13; a guarda
   `-DUP` permanece como rede de segurança).
3. Nova função `limparDuplicatasOrfasDB()` (menu "🧹 Remover duplicatas órfãs do
   Relatorio_DB"): remove linhas que atendam TODOS os critérios — `Ativo`, sem
   MARCAR_FATURAR/USUARIO/LOTE_EMISSAO, ID e UUID ausentes de PEDIDOS, sem histórico em
   Baixas_Historico, e grupo (fingerprint SEM OS) com mais linhas ativas no DB que em
   PEDIDOS com quota > 0 (nunca toca itens que saíram legitimamente da fonte). Auditoria
   gravada na aba `Duplicatas_Removidas` (append; a `Duplicatas_Debug` é apagada a cada
   sync). Roda sob LockService.

**Ordem de aplicação:** publicar o código corrigido → rodar "🔄 Forçar resync agora"
(estabiliza os IDs em PEDIDOS) → rodar a limpeza de órfãs → o sync seguinte re-adiciona ao
DB os itens de PEDIDOS que ficaram sem correspondente, já com IDs estáveis.

### 15.15 Sentinela de duplicidade + ferramentas manuais (portado do sistema irmão Bahia)

As correções de 15.9–15.14 impedem os casos conhecidos de duplicação, mas não davam um
sinal automático caso um caminho novo escapasse — o problema só aparecia visualmente,
semanas depois, com centenas de linhas já acumuladas. Portado do projeto irmão
(`johnnybahia/EXPEDICAOBAHIA`) um conjunto de proteções que rodam **depois** de qualquer
duplicata já ter passado pelas travas de `sincronizarPedidosComFonte`/`sincronizarDados`.

**Sentinela automática** (`_verificarIntegridadeDuplicatas_` + `_registrarSentinelaDuplicatas_`,
chamada ao fim de `sincronizarDados()`, logo após `_gravarDuplicatasDebug_`): varre o
Relatorio_DB a cada sync e checa 5 condições —
1. `ID_UNICO` repetido;
2. `CÓDIGO_FIXO` (UUID) repetido;
3. mesmo item **aberto** em mais de uma ORD. COMPRA (mesma impressão digital sem OC);
4. linha aberta sem correspondência em PEDIDOS, quando outra linha do mesmo grupo é
   reconhecida (cobre o caso em que a contraparte já está Faturado — a verificação 3 não
   vê, porque só uma das linhas está aberta);
5. linhas abertas além do que `DADOS_IMPORTADOS` reconhece por identidade (mesmo orçamento
   de excedente que o sync usa para faturar automaticamente — ver 15.14 / `2ceb277`).

O resultado é gravado em `PropertiesService` (`ALERTA_DUPLICATAS`) e devolvido no payload
de `fetchAllDataUnified` como `result.alertaDuplicatas`; o frontend (`showAlertaDuplicidade`)
mostra isso no mesmo banner de avisos (`$avisosBanner`) já usado por `AVISOS_PENDENTES` —
**a duplicidade tem precedência**: se os dois avisos coincidirem no mesmo ciclo, o banner de
duplicidade sobrescreve o de avisos pendentes (decisão herdada do Bahia, não um bug). Quando
o próximo sync não encontra mais nada, a propriedade é apagada e o aviso some sozinho.

**Ferramentas de menu adicionadas** (todas leem o mesmo `_agruparItensDuplicadosDB_`, que
agrupa linhas do DB por impressão digital sem OC ou por UUID repetido):
- `verificarDuplicidadeAgora()` — roda a sentinela sob demanda e mostra o resultado num alert.
- `diagnosticarItensDuplicadosOC()` — só leitura; grava a aba `Duplicatas_OC_Diagnostico`
  com todo item em mais de uma OC, sinalizando qual cópia é a "provável correta" (a que
  não está `Faturado`) quando há uma órfã fantasma.
- `arquivarDuplicatasOrfas()` — marca `Excluido` (nunca apaga) as cópias órfãs que atendem
  simultaneamente: grupo de exatamente 2 linhas em OCs diferentes, uma `Faturado`/`Inativo`
  e a outra `Ativo`, o par OC+OS da órfã não existe mais em `DADOS_IMPORTADOS`, e a órfã não
  tem nenhuma baixa registrada em `Baixas_Historico`. Complementar a `limparDuplicatasOrfasDB()`
  (15.14): esta última compara contagem PEDIDOS vs. DB por fingerprint sem OS; a nova
  compara pares de linhas explicitamente em OCs diferentes.
- `faturarItensForaDaFonte()` — reforço manual do que o sync automático já decide sozinho a
  cada ciclo (mesma lógica de orçamento por identidade de 15.14/`2ceb277`); útil só para
  adiantar o resultado sem esperar o próximo processo completo. ⚠️ Fatura sem usuário — viola 1.1.3
  (plano: passar a só sinalizar, seção 17). Mesmas travas do sync:
  `MIN_LINHAS_FONTE_PARA_FATURAR` e `MAX_FRACAO_SEM_FONTE` abortam se a fonte parecer
  incompleta. Auditoria em aba própria `Itens_Fora_Da_Fonte`.

**Não portado do Bahia:** a coluna `INFO_Y` (Bahia usa a coluna Y da fonte para "conta e
ordem" — transferência entre filiais; no Ceará a coluna Y da fonte já é `LOTE`, usada para
outra coisa — não há dado de origem equivalente) e o rodapé de etiquetas NF (dados da
Marfim Bahia; o Ceará já tem os dados corretos da Marfim Ceará). A retenção de itens
Faturados ficou em `DIAS_RETENCAO = 15` (Bahia mudou para 1 dia). **Em 28/09/2026 o usuário decidiu que
o Ceará também deve apagar todos os faturados diariamente** — regra 1.1.5, pendente de implementação (15.18).

### 15.16 Incidente 28/09/2026: `Faturado` sem usuário + limite de identidade (linhas-irmãs) — ABERTO
**Sintoma:** 20 itens `Faturado` sem MARCAR_FATURAR_USUARIO e sem LOTE_EMISSAO, todos com QTD>0 (1.873 un.),
DATA_STATUS 22, 23 e 25/09. O usuário recebeu a lista com a evidência de cada item (CSV).
- **9 falsos (prova direta):** DASS ITAPIPOCA, pedido 143086, OC 16078984, 140/145/150CM (IDs `-18` a `-23`).
  O LOTE de cada um continua aberto na origem com a mesma QTD. Os lotes vivos estão sob os IDs `-2/-3/-12`
  (140CM) e `-2/-3/-15` (145CM e 150CM), que no DB mostram QTD 0 enquanto PEDIDOS mostra 130/120/113,
  420/390/430 e 10/10/19. Revisar o Baixas_Historico desses IDs antes de faturá-los.
- **11 com produto ausente da origem:** DASS ITAPIPOCA 142869/OC 16139161 (8), DASS ITAPIPOCA 142555 (1),
  DAKOTA MARANGUAPE 143198/OC 91286D (2). Nenhum veio de troca de DESCRIÇÃO.

**Mecanismo:** linhas-irmãs (mesma impressão digital, CÓD. OS vazio ou "0") são casadas de forma gulosa (7.2)
→ IDs e UUIDs migram entre lotes quando a origem inclui, remove ou edita irmãs → a linha que carregava a QTD
real perde a vaga → o orçamento de excedente (7.3) a fatura sem usuário. Mudar só a escolha de "qual irmã
faturar" **não resolve**: as irmãs com QTD 0 estão casadas por ID e nunca chegam à decisão.

**Limite estrutural (medido na origem de 28/09, 2.259 linhas):**
- CÓD. OS é o único separador: sem ele, 83% das linhas colidem; com ele, 12,7% (a impressão digital atual).
- 99% das colisões restantes têm OS vazio ou "0" (25% da origem).
- 40 grupos (192 linhas) são iguais até em LOTE e QTD — nenhuma chave tirada da origem os separa.
- LOTE preenchido em só 15% das linhas (340), com 9 valores repetidos.

**Risco latente (não manifestado em 28/09):** `dbFingerprintMap` (recuperação de ID em
`sincronizarPedidosComFonte`) e `fpVagasNovos` (trava de itens novos) incluem linhas `Faturado/Finalizado/Excluido`.
Uma linha nova que entra num grupo de irmãs pode herdar o ID de uma irmã já faturada e aparecer como `Faturado`,
ou seja, escondida. Em 28/09, 60 faturados tinham a mesma impressão digital de linhas vivas. Na mesma data:
0 linhas da origem sem item no DB e 0 linhas vivas escondidas como faturadas.

**Não rodar `corrigirFaturadosComSaldoAberto()` como correção:** reverte, mas o sync fatura de novo (loop).
**Prazo:** com a limpeza atual (15 dias) os 20 somem entre 07/10 e 12/10; com a limpeza diária (1.1.5), na
primeira execução. Tratar antes (seção 17, item 10).

### 15.17 Concorrência e número de linha — ABERTO
- `aplicarBaixa`, `marcarParaFaturar`, `excluirItem` e `finalizarItem` usam o número de linha enviado pela tela
  sem conferir o ID. Qualquer exclusão de linhas (limpeza diária, `limparDuplicatasOrfasDB`) move as linhas de
  baixo → quem abriu a tela antes da exclusão age sobre o **item errado** até recarregar.
- `processoImportacao` não usa trava e o processo completo lê DADOS_IMPORTADOS duas vezes (7.2 e 7.3): uma
  importação no meio faz as duas etapas verem versões diferentes da origem (choque típico :32 × :33, seção 13).
- O sync regrava linhas inteiras com o que leu no início, e as ações de usuário não usam trava: marcação ou baixa
  feita durante o sync pode ser sobrescrita; item desmarcado durante o sync pode ser faturado pela marcação antiga.

### 15.18 Limpeza de faturados — como está × como deve ser
- **Como está:** ETAPA 3 do processo completo; dias úteis; primeira execução a partir da hora de
  `CONFIGURAÇÕES!B2` (padrão 11); apaga `Faturado/Finalizado/Excluido` com `DATA_STATUS` de 15+ dias
  (`DIAS_RETENCAO`, fixo no código), linha a linha (`deleteRow` em loop); item sem DATA_STATUS nunca é apagado.
  Menu "🧹 Limpar Faturados agora" roda na hora. Em 28/09 nada tinha sido apagado (faturado mais antigo: 18/09).
- **Como deve ser (1.1.5):** apagar **todos** os `Faturado` e `Excluido` no horário de B2, sem idade mínima;
  `Finalizado` fica — plano, seção 17.
- A célula A3 da aba é texto fixo gravado na criação — não controla nada.

---

## 16. SEQUÊNCIA SEGURA PARA MUDANÇAS

Antes de qualquer alteração no código, verificar:

0. **Respeita a seção 1.1?** Nenhum caminho novo pode gravar `Faturado` sem usuário; nada pode ser gravado na origem.
1. **Afeta cálculo de QTD?** → Revisar `aplicarBaixa`, `registrarBaixa`, caches de saldo
2. **Afeta matching de itens?** → Revisar `_criarImpressaoDigital_`, `_chaveProduto_` (DESCRIÇÃO), gap de col D,
   normalização Dilly e linhas-irmãs (15.16)
3. **Afeta faturamento?** → Revisar `obterItensMarcadosParaFaturar`, cálculo de SALDO, fallback, tabela de
   caminhos que gravam `Faturado` (seção 9)
4. **Afeta sincronização?** → Revisar guard H2, travas de importação (7.3), reset de ciclo, concorrência (15.17)
5. **Afeta IDs?** → Verificar impacto no Baixas_Historico e recuperação por UUID
6. **Custa tempo de execução?** → Cota de 90 min/dia de acionadores (seção 13); nada de leitura/escrita por item
7. **Apaga linhas?** → Ações de usuário usam número de linha (15.17)

**Funções com maior superfície de impacto** (cuidado máximo):
- `sincronizarDados()` — toca todos os itens do DB
- `sincronizarPedidosComFonte()` — reescreve aba PEDIDOS inteira
- `calcularQtdOriginal()` / `_getSaldoEfetivoCache_()` — usadas no relatório de faturamento
- `_criarImpressaoDigital_()` — identidade de cada item

---

## 17. PLANO PENDENTE (discutido em 28/09/2026 — NADA IMPLEMENTADO; aguardando aprovação final do usuário)

Objetivo: cumprir 1.1.3 (nenhum `Faturado` sem usuário) e 1.1.5 (limpeza diária) sem mudar a rotina dos usuários.

1. **Regra única de faturamento no sync:** `Faturado` só se o item foi marcado pelo usuário (col V) e a linha saiu
   da origem — com as 4 travas de importação também nesse caminho e releitura da col V logo antes de gravar.
   Todo o resto que sai da origem (QTD>0, QTD=0 com ou sem baixa, "SIM" legado) → coluna nova Z
   `CONFERENCIA_SAIDA` = `PENDENTE dd/mm`; o status continua `Ativo`, fica fora do orçamento de excedente e não é
   reprocessado. Se o item voltar à origem, a pendência some sozinha.
2. **Aviso após o login** (não bloqueia; botão "Responder depois") + selo no item + contador que atualiza a cada
   recarga. Card com cliente, pedido, OC, descrição, tamanho, LOTE, QTD, baixas, data da saída e alerta de gêmeo
   ativo com o mesmo LOTE. Botões: **Faturado** → `Faturado` · **Cancelado** → `Excluido` · **Continua aberto**
   → mantém e não pergunta de novo · **Duplicata** (só quando há gêmeo) → `Excluido`.
   Função `confirmarSaidaFonte(id, linha, decisao, usuario)`: trava, reconfere se ainda está pendente, valida no
   servidor o nível TOTAL (aba CADASTRO), grava decisão + usuário + data e limpa o cache.
3. **Fechar os outros caminhos:** remover `marcarFaturado()` e `confirmarTodosAlertas()` (+ item de menu);
   `faturarItensForaDaFonte()` passa a só sinalizar `PENDENTE`.
4. **Identidade (`sincronizarPedidosComFonte`):**
   a. casar em duas passadas por grupo de irmãs — 1ª: linhas que não mudaram (comparar as colunas da origem **sem
      o PRAZO**, que muda todo dia) mantêm o ID; 2ª: as que mudaram disputam o resto por LOTE → QTD → estado no DB
      (QTD>0 antes de zerada por baixa) → ordem. Aplicar também em :1850 (Dilly), :1803/:1884 (`dbFingerprintMap`),
      :3551 (Passada 2) e na troca de OC (:3564, hoje exige candidato único);
   b. `dbFingerprintMap` e `fpVagasNovos` só com linhas abertas — linha nova nunca herda ID de item finalizado;
   c. reserva sem DESCRIÇÃO: reaproveitar o ID pela base do ID (cliente, filial, pedido, CÓD. MARFIM, tamanho, OC,
      OS, data). Em 28/09: 2.051 bases, nenhuma com mais de uma DESCRIÇÃO (não mistura produtos).
5. **Concorrência:** a importação usa a mesma trava do sync (pula a rodada se estiver ocupado); o sync relê as
   colunas do usuário (P, V, W, Z) antes de gravar.
6. **Conferência do ID** em baixa, estorno e edição de baixa, marcar, excluir e finalizar: se a linha não pertence
   mais àquele ID, procurar pelo ID; se não achar, pedir para recarregar a tela.
7. **Limpeza diária (1.1.5):** no horário de B2, dias úteis, apagar todos os `Faturado` e `Excluido` (confirmado
   pelo usuário; `Finalizado` não é apagado); nessa execução, limpar **antes** do sync; apagar em blocos
   contíguos de baixo para cima (`deleteRows`); remover `DIAS_RETENCAO`; reescrever o texto da A3.
   Consequência: "Cancelado" e "Duplicata" do aviso (→ `Excluido`) só podem ser desfeitos até a limpeza
   seguinte; depois disso, se a linha voltar à origem, entra como item novo, sem histórico.
8. **Sentinela, novos alertas:** `Faturado` sem usuário (deve ser sempre 0); mesmo LOTE numa linha ativa e noutra
   faturada ou pendente; `Ativo` com QTD 0 no DB e >0 em PEDIDOS.
9. **Modo simulação antes de ativar:** 3–5 dias úteis gravando numa aba de auditoria o que a lógica nova faria
   (faturar, sinalizar, apagar), sem alterar o DB. Ativar só com 0 "Faturado sem usuário" e avisos explicáveis.
10. **Reparo dos 20 itens (15.16)** antes de ligar a limpeza diária: os 11 ausentes voltam para `Ativo` + `PENDENTE`;
    os 9 gêmeos vão para uma aba de diagnóstico com os pares, para decisão manual.
11. Atualizar este documento depois de implementar.

**Garantias pretendidas:** nenhum `Faturado` sem usuário (por construção + sentinela); item que sai da origem sem
marcação nunca some calado (vai para o aviso); linha nova nunca herda item finalizado; cada aviso falso é
perguntado uma vez só. **Não garantível** (limite da origem, 1.1.4): saber qual linha é qual entre irmãs 100%
idênticas — o dano fica contido a um aviso, nunca a um faturamento.

### Ideias já avaliadas e DESCARTADAS (não repropor sem fato novo)
- Gravar código único na planilha de origem — proibido (1.1.4).
- Usar a coluna A da origem como ID — é numeração contínua 1…N, renumerada.
- Aba/"banco" de códigos com chave CÓD. CLIENTE + TAMANHO + DESCRIÇÃO (± PEDIDO, OC, QTD ou UUID nosso) — colide
  em 94% das linhas (83% com PEDIDO); QTD muda com baixas; incluir um UUID gerado por nós na chave é circular
  (a linha importada não traz o UUID).
- Mudar a fórmula do ID_UNICO — migra todos os IDs e o vínculo do Baixas_Historico sem resolver as linhas-irmãs.
- "Faturar primeiro as irmãs com QTD 0" — elas estão casadas por ID e nunca chegam à decisão.
- Reativar o sistema de alertas antigo (PropertiesService + senha) — desligado em `ce75202`; substituído pelo item 2.
- `corrigirFaturadosComSaldoAberto()` como correção dos 20 — o sync fatura de novo (loop).
- Só esconder os faturados na tela e apagar à noite mantendo 15 dias — o usuário quer apagar tudo diariamente (1.1.5).
