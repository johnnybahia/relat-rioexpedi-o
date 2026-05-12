# CLAUDE.md — Sistema de Relatório de Pedidos / Expedição

> Lido automaticamente em toda sessão. Atualizar sempre que houver mudança arquitetural.
> Versão do sistema: **v15.6-SINCRONIZACAO** · Backend em Google Apps Script · Frontend em HTML+JS embutido no Apps Script

---

## 1. VISÃO GERAL

Sistema enterprise de gestão de pedidos de expedição em Google Sheets.  
Fluxo principal: **Fonte externa → DADOS_IMPORTADOS → PEDIDOS → Relatorio_DB → UI (index.html)**

Arquivos do projeto:
- `Código.gs` — todo o backend (~4600 linhas, único arquivo GAS)
- `index.html` — frontend servido via `HtmlService`
- `RELATÓRIO DE PEDIDOS EXPEDIÇÃO - PEDIDOS.csv` — snapshot da aba PEDIDOS
- `RELATÓRIO DE PEDIDOS EXPEDIÇÃO - Relatorio_DB.csv` — snapshot do DB

---

## 2. ABAS DA PLANILHA

| Aba | Papel | Observação |
|---|---|---|
| `DADOS_IMPORTADOS` | Intermediária com dados da fonte externa | Atualizada via `importarDadosExternos()`. Célula H2 = timestamp (guard de sync) |
| `PEDIDOS` | Dados sincronizados e enriquecidos | IDs gerados aqui. Col A = ID_UNICO. Dados começam linha 4 |
| `Relatorio_DB` | Banco de dados final usado pela UI | 23 colunas (A-W). Mantém histórico de status |
| `Baixas_Historico` | Histórico de parcializações de QTD | Cabeçalho lido dinamicamente. Suporta coluna TIPO para CHECKPOINTs |
| `LOTE DILLY` | Mapeamento OC→Lotes para cliente Dilly | Consumido em FIFO durante sincronização |
| `original` | Dados originais para ordenação | Define posição dos itens dentro de uma OC |
| `Duplicatas_Debug` | Auditoria de itens descartados | Criada automaticamente |
| `SENHA` | Senha para confirmar alertas (cel A2) | Opcional; usado por `confirmarAlerta()` |

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
11 L  CÓD. OS / LOTE        (Dilly: substituído pelo Lote da aba LOTE DILLY)
12 M  DATA RECEB.
13 N  DT. ENTREGA
14 O  PRAZO (texto)
15 P  TIMESTAMP_CRIACAO
16 Q  POSICAO_FONTE         (índice em DADOS_IMPORTADOS, fixo)
17 R  (reservado)
18 S  CODIGO_FIXO           (UUID imutável por item)
19 T  INFO_X                (col X da fonte)
20 U  LOTE                  (col Y da fonte)
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
15 P  MARCAR_FATURAR       ← "SIM" | "" — marcação para emissão de NF
16 Q  DATA_STATUS          (data da última mudança de status)
17 R  POSICAO_FONTE
18 S  CODIGO_FIXO
19 T  INFO_X
20 U  LOTE
21 V  MARCAR_FATURAR_USUARIO
22 W  LOTE_EMISSAO         ← "FAT-001", "FAT-002"...
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
DIAS_RETENCAO = 15              // itens Finalizados/Excluídos purgados após 15 dias
CACHE_DURATION = 600            // segundos (CacheService)
TZ = 'America/Fortaleza'
```

---

## 7. FLUXO DE SINCRONIZAÇÃO (passo a passo)

### 7.1 Importação (a cada 5 min — `processoImportacao`)
1. Lê planilha externa via `openById()` + SOURCE_SHEET
2. Coluna J (CÓD. CLIENTE): usa `getDisplayValues()` para preservar sufixos "82249D", "14660U"
3. Copia para `DADOS_IMPORTADOS`
4. Atualiza H2 com timestamp (sinaliza ao próximo passo que há dados novos)
5. Agenda sync via `_agendarSincronizacao_()` para 90s depois

### 7.2 Sync DADOS_IMPORTADOS → PEDIDOS (`sincronizarPedidosComFonte`)
- **Guard de H2**: se H2 igual ao último processado → aborta (sem retrabalho)
- Pré-carrega IDs do DB + fingerprints (evita colisões)
- Lê aba `original` para calcular `POSICAO_FONTE`
- Lê `LOTE DILLY` para enriquecimento de CÓD. OS (Dilly)
- Para cada linha em DADOS_IMPORTADOS:
  - Cria fingerprint: `CLIENTE|PEDIDO|MARFIM|TAM|OC|OS|DATA_NORM`
  - Tenta reutilizar ID existente (por fingerprint → UUID → geração nova)
  - Normaliza MARFIM e CÓD. CLIENTE para Dilly
  - Para Dilly: substitui CÓD. OS pelo Lote da fila FIFO
- Escreve resultado em PEDIDOS

### 7.3 Sync PEDIDOS → Relatorio_DB (`sincronizarDados`)
- Garante headers do DB (`_garantirHeadersRelatorio_DB_`)
- Cria mapas: `fonteMap`, `dbMap`, impressões digitais, `codigoFixoMap`
- Lê `DADOS_IMPORTADOS` para proteção anti-faturamento indevido (OC+OS count)
- Para cada item do DB, tenta match por: **ID → UUID (CODIGO_FIXO) → fingerprint**
- Se item saiu de PEDIDOS:
  - QTD=0 ou MARCAR_FATURAR=SIM → marca `Faturado`, limpa MARCAR_FATURAR
  - QTD>0 sem marcação + sem proteção → `MARCAR_FATURAR="SIM"` + alerta
  - QTD>0 + proteção ativa → aguarda reconsolidação (não marca)
- Adiciona itens novos de PEDIDOS que não existem no DB
- Atualiza IDs no Baixas_Historico se IDs mudaram
- Purga itens finalizados (>15 dias)

### 7.4 Processo Completo (a cada 5 min — `processoAutomaticoCompleto`)
1. Verifica pausa (`_sistemaPausado_()`)
2. Adquire `LockService` (evita execuções simultâneas, timeout 30s)
3. Etapas: importação → geração de IDs → sincronização → purge → limpar cache

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
Efeito: registra CHECKPOINT, limpa MARCAR_FATURAR e LOTE_EMISSAO, remove alertas.

---

## 9. SISTEMA DE FATURAMENTO

### Estados de MARCAR_FATURAR
- `""` (vazio) — item não marcado
- `"SIM"` — marcado para emitir NF (pode ser por usuário OU pelo sync automático)

### Quem define MARCAR_FATURAR="SIM"
1. **Usuário** via checkbox na UI → `marcarParaFaturar()`
2. **Sync automático** → item saiu de PEDIDOS com QTD > 0 (saída inesperada)

### Cálculo do SALDO no Relatório
```javascript
const qtdOriginal    = item['QTD. ORIGINAL'] || 0;  // de calcularQtdOriginal()
const qtdAberta      = item['QTD. ABERTA']   || 0;
const saldoCalculado = qtdOriginal - qtdAberta;
// Fallback: se não há baixas, qtdOriginal == qtdAberta → saldoCalculado == 0
// Neste caso usa QTD.ABERTA para não zerar o relatório
const saldo = saldoCalculado > 0 ? saldoCalculado : qtdAberta;
```
**Por que o fallback existe:** quando nenhuma baixa foi registrada (ex: sync setou MARCAR_FATURAR automaticamente), `calcularQtdOriginal` retorna `qtdAberta` como fallback → saldo seria 0 sem esta correção.

### Fluxo de Geração do Relatório
1. Usuário clica "Gerar Relatório de Faturamento"
2. `obterItensMarcadosParaFaturar()` → lê todos com MARCAR_FATURAR="SIM"
3. Filtro por usuário → filtro por marca/INFO_X
4. `gerarNumeroLoteEmissao()` → "FAT-001", "FAT-002"...
5. `_gerarRelatorioFaturamentoPDF()` → abre janela de impressão
6. `registrarLoteEmissao()` → grava LOTE_EMISSAO (col W) nos itens

### Proteção Anti-Faturamento Indevido
- Ativa quando item some de PEDIDOS mas OC+OS ainda existe em DADOS_IMPORTADOS
- Significa: item pode ter sido reconsolidado com novo ID → aguardar
- **Lógica proporcional**: cada match consome 1 slot; múltiplos itens idênticos são liberados progressivamente

---

## 10. SISTEMA DE IDs

### Estrutura do ID
```
{CLIENTE}{FILIAL}{PEDIDO}{MARFIM}{TAMANHO}{OC}{OS}{DATA_yyyyMMdd}-{SUFIXO_NUMÉRICO}
```
Exemplo: `MARFIM MINAS7413040480114120CM179202520250917-1`

### Campos NÃO inclusos no ID (mutáveis)
- CARTELA (pode mudar sem alterar o item)
- DESCRIÇÃO (pode ter correções)

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
- `_criarImpressaoDigitalFromRow_` aplica `_normalizarMarfimDilly_` ao comparar
- Garante que "196338-120" em DADOS_IMPORTADOS bate com "196338-110" em PEDIDOS

---

## 12. REGRAS ESPECIAIS — DAKOTA

Detecção: `cliente.toUpperCase().includes('DAKOTA')`

- No relatório de faturamento: cada filial Dakota = grupo separado (pelo nome do cliente)
- Demais clientes: agrupados por marca (INFO_X / col T)

---

## 13. TRIGGERS E AUTOMAÇÃO

| Trigger | Função | Frequência |
|---|---|---|
| `processoImportacao` | Importa dados externos, atualiza H2, agenda sync | A cada 5 min |
| `processoAutomaticoCompleto` | Sync completo (PEDIDOS → DB → purge) | A cada 5 min |
| Agendado interno | `sincronizarPedidosComFonte` | 90s após importação |

**Guard de execução dupla:** `LockService.getScriptLock().waitLock(30000)`  
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
- `alertaFaturamentoModal` — confirmação de divergências (requer senha)
- `filtroUsuarioFaturModal` — filtra itens por usuário antes de imprimir
- `filtroMarcaFaturModal` — filtra por marca/INFO_X antes de imprimir

### Acesso Parcial
- `document.body.classList.add('acesso-parcial')` oculta botões de ação via CSS
- **Não bloqueia no backend** — segurança real deve ser implementada no GAS se necessário

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

### 15.7 Proteção anti-faturamento é proporcional
Se há 3 itens com OC+OS idênticos em DADOS_IMPORTADOS e 2 ativos no DB:  
- Slots = 2; depois de 2 matches, o 3º item **não é mais protegido** (faturamento liberado).  
Necessário entender que a contagem decrementa a cada match bem-sucedido.

### 15.8 Alertas de faturamento estão desativados
`_registrarAlertaFaturamento_` apenas loga, não persiste em PropertiesService.  
`obterAlertasPendentes()` limpa e retorna `[]` sempre.  
Para reativar: descomentar a lógica de JSON no código e garantir aba SENHA.

### 15.9 IDs com sufixo numérico podem mudar
Se itens forem reordenados em DADOS_IMPORTADOS, o sufixo de um item pode mudar na próxima sync (ex: `-1` vira `-2`).  
O **UUID (CODIGO_FIXO)** garante que o histórico de baixas seja preservado mesmo assim.

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

---

## 16. SEQUÊNCIA SEGURA PARA MUDANÇAS

Antes de qualquer alteração no código, verificar:

1. **Afeta cálculo de QTD?** → Revisar `aplicarBaixa`, `registrarBaixa`, caches de saldo
2. **Afeta matching de itens?** → Revisar `_criarImpressaoDigital_`, gap de col D, normalização Dilly
3. **Afeta faturamento?** → Revisar `obterItensMarcadosParaFaturar`, cálculo de SALDO, fallback
4. **Afeta sincronização?** → Revisar guard H2, proteção anti-faturamento, reset de ciclo
5. **Afeta IDs?** → Verificar impacto no Baixas_Historico e recuperação por UUID

**Funções com maior superfície de impacto** (cuidado máximo):
- `sincronizarDados()` — toca todos os itens do DB
- `sincronizarPedidosComFonte()` — reescreve aba PEDIDOS inteira
- `calcularQtdOriginal()` / `_getSaldoEfetivoCache_()` — usadas no relatório de faturamento
- `_criarImpressaoDigital_()` — identidade de cada item
