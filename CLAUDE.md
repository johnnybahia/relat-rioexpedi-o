# CLAUDE.md — Sistema de Relatório de Pedidos / Expedição

> Lido automaticamente em toda sessão. Atualizar sempre que houver mudança arquitetural.
> Versão do sistema: **v15.6-SINCRONIZACAO** · Backend em Google Apps Script · Frontend em HTML+JS embutido no Apps Script
> Última revisão contra o código e dados reais: **30/09/2026** (plano da seção 17 implementado). Se este documento
> divergir do código, o código descreve o comportamento real — mas as **regras de negócio da seção 1.1 prevalecem**:
> se o código as viola, é o código que deve ser corrigido. Antes de propor mudanças, ler 1.1, 15.16–15.19 e 17
> (estado da implantação e ideias descartadas).
>
> **Modo do sistema — `CONFIGURAÇÕES!B4`:** `SIMULACAO` (padrão) ou `ATIVO`. Ver 15.19 para o que cada modo liga.

---

## 1. VISÃO GERAL

Sistema enterprise de gestão de pedidos de expedição em Google Sheets.  
Fluxo principal: **Fonte externa → DADOS_IMPORTADOS → PEDIDOS → Relatorio_DB → UI (index.html)**

Arquivos do projeto:
- `Código.gs` — todo o backend (~7600 linhas, único arquivo GAS)
- `index.html` — frontend servido via `HtmlService` (~6350 linhas)
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
   ✅ Desde 30/09/2026 só existem esses dois caminhos (seção 9), nos dois modos. O que sai da origem sem
   marcação fica `Ativo` com a coluna Z = `PENDENTE` e vai para o aviso da tela (15.19).
4. **A planilha de origem não pode ser alterada** (nenhuma coluna nova, nenhum código gravado nela).
   Portanto não existe identificador fixo por linha vindo da origem — ver 15.16.
5. **Limpeza diária:** todo dia útil, no horário da célula B2 da aba `CONFIGURAÇÕES`, apagar **todos** os
   `Faturado` e `Excluido` do Relatorio_DB. **Sem histórico de 15 dias.** `Finalizado` **não** é apagado
   (em 28/09 não havia nenhum: o botão "Finalizar" existe no código da tela, mas `handleFinalizar` nunca é chamado).
   ✅ Implementado com `CONFIGURAÇÕES!B4 = ATIVO`. Em `SIMULACAO` continua a regra antiga de 15 dias e a aba
   `Auditoria_Sincronizacao` registra o que seria apagado. Nos dois modos, `Faturado` sem usuário **nunca** é
   apagado (fica até ser reparado — 15.18).
6. **Acionadores em produção:** `processoImportacao` a cada **15 min**; `processoAutomaticoCompleto` a cada
   **1 h**. Desde 30/09 os instaladores do código criam exatamente isso (seção 13).

---

## 2. ABAS DA PLANILHA

| Aba | Papel | Observação |
|---|---|---|
| `DADOS_IMPORTADOS` | Intermediária com dados da fonte externa | Atualizada via `importarDadosExternos()`. Célula H2 = timestamp (guard de sync) |
| `PEDIDOS` | Dados sincronizados e enriquecidos | IDs gerados aqui. Col A = ID_UNICO. Dados começam linha 4 |
| `Relatorio_DB` | Banco de dados final usado pela UI | 26 colunas (A-Z). Cabeçalhos em `RELATORIO_DB_HEADERS`; coluna nova entra sozinha via `_garantirHeadersRelatorio_DB_` (insere colunas físicas se faltar) |
| `CONFIGURAÇÕES` | Configuração editável pelo usuário | B2 = hora (0-23) da limpeza diária, padrão 11. **B4 = modo `SIMULACAO`/`ATIVO`** (vazio → `SIMULACAO`). A3/A4 são texto, regravados pelo código (`_atualizarTextosConfiguracoes_`). **Não existe célula de dias** |
| `Auditoria_Sincronizacao` | Registro do que o sistema fez ou faria (modo simulação) | DATA_HORA · MODO · TIPO · ID · DETALHE. Criada sozinha; guarda as últimas 5.000 linhas (tipos em 15.19) |
| `Reparo_Faturados` | Resultado do menu "🩹 Reparar faturados sem usuário" | Append; lista o que voltou para conferência e os pares com gêmea para decisão manual |
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
                            só é copiado. Preenchido em só ~15% das linhas; há valores repetidos.
                            No modo ATIVO é critério de desempate entre linhas-irmãs (15.16)
21 V  SEQ_ORIGINAL          (posição fixa do item na OC, aba "original") — PEDIDOS_SEQ_ORIGINAL_COL=21
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
                            `Finalizado` só vem da ação manual `finalizarItem`; a tela não exibe o
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
25 Z  CONFERENCIA_SAIDA    ← DB_CONFERENCIA_SAIDA_COL=25. "TIPO|quem|dd/MM/yyyy HH:mm|detalhe":
                             PENDENTE (saiu da origem sem marcação — aguarda usuário), ABERTO (usuário disse
                             "continua aberto" — não pergunta de novo), FATURADO/CANCELADO/DUPLICATA (decisão
                             registrada). Vazia = nada a conferir. Some sozinha se o item volta à origem.
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
DB_CONFERENCIA_SAIDA_COL = 25   // col Z
DIAS_RETENCAO = 15              // só no modo SIMULACAO (regra antiga da limpeza); ATIVO apaga tudo (15.18)
CONFIG_SHEET_NAME = 'CONFIGURAÇÕES', CONFIG_HORA_LIMPEZA_CELL = 'B2', CONFIG_HORA_LIMPEZA_PADRAO = 11
CONFIG_MODO_CELL = 'B4', CONFIG_MODO_PADRAO = 'SIMULACAO'   // lido por _modoAtivo_() (cache por execução)
AUDITORIA_SHEET_NAME = 'Auditoria_Sincronizacao', AUDITORIA_MAX_LINHAS = 5000
MIN_LINHAS_FONTE_PARA_FATURAR = 50 // travas de fonte confiável (7.3) — valem para faturar E para sinalizar
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
5. Agenda `processoAutomaticoCompleto` via `_agendarSincronizacao_(90)` para 90s depois
- Usa a mesma trava do sync (`LockService.getScriptLock().tryLock(20000)`): se o sync estiver rodando, a
  importação pula a rodada (a próxima vem em 15 min) — não regrava DADOS_IMPORTADOS no meio de um sync.
- Coluna A da origem ("Column 24") é uma numeração contínua 1…N, renumerada — **não serve como ID**.

### 7.2 Sync DADOS_IMPORTADOS → PEDIDOS (`sincronizarPedidosComFonte`)
- **Guard de H2**: se H2 igual ao último processado → aborta (sem retrabalho)
- Pré-carrega IDs do DB + fingerprints (evita colisões)
- Lê aba `original` para calcular `POSICAO_FONTE`
- Lê `LOTE DILLY` para enriquecimento de CÓD. OS (Dilly)
- Fingerprint: `CLIENTE|PEDIDO|PRODUTO|TAM|OC|OS|DATA_NORM`, onde **PRODUTO = DESCRIÇÃO** sem o sufixo
  `[uuid]` (`_chaveProduto_`); CÓD. MARFIM só é usado se a DESCRIÇÃO estiver vazia. LOTE e QTD não entram.
- **Qual linha de PEDIDOS (vaga) cada linha da origem reaproveita** depende do modo (`CONFIGURAÇÕES!B4`):
  - **ATIVO — planejador de linhas-irmãs** (`_planejarCasamentoIrmas_`, PASSO 2.7, antes do laço):
    agrupa origem e PEDIDOS pela fingerprint (Dilly: fingerprint sem OS) e decide o grupo inteiro de uma vez.
    Custo de cada par linha×vaga (`_custoIrma_`), do melhor para o pior: linha idêntica à versão anterior
    (todas as colunas da origem, **menos o PRAZO**) → menos colunas diferentes (QTD conta como igual se bater
    com a QTD do DB depois das baixas) → LOTE igual → OS igual (Dilly) → QTD mais próxima → estado no DB
    (aberto e intocado → ainda não está no DB → zerado/marcado/pendente → finalizado) → ordem.
    Reserva sem DESCRIÇÃO: linha que ficou sem vaga reaproveita uma sobra de PEDIDOS com a **mesma base de
    ID** (cliente, filial, pedido, CÓD. MARFIM, tamanho, OC, OS, data — `_idBaseFonte_`), nunca de item
    finalizado; auditoria `ID_POR_BASE`. Linhas quebradas (PASSO 1.5) reservam a sua vaga antes do planejador.
    Sem vaga → `dbFingerprintMapAbertos` (só itens abertos e nunca um ID que ainda está em PEDIDOS) → ID novo
    (pula IDs que ainda estão em PEDIDOS).
  - **SIMULACAO — casamento antigo, guloso**, linha a linha na ordem da aba: `pedidosMap` (desempate pela
    QTD mais próxima da QTD anterior; empate → primeira) → `pedidosDillyMap` → `dbFingerprintMap` (qualquer
    status, FIFO) → ID novo. O planejador roda ao lado só para comparar: cada linha que teria outro ID vai
    para a auditoria como `IDENTIDADE_SIMULADA` (mesmo conjunto não é regravado: hash em
    `ULTIMA_DIF_IDENTIDADE`). Os erros do casamento antigo estão em 15.16.
- Normaliza MARFIM e CÓD. CLIENTE para Dilly
- Para Dilly: substitui CÓD. OS pelo Lote da fila FIFO
- **Guarda anti-duplicata final** (ver 15.13): `idsAssignadosNestaRodada` detecta se o
  `idFinal` resolvido já foi usado por outra linha NESTA execução e renomeia
  (`-DUP{n}`) antes de escrever
- Escreve resultado em PEDIDOS

### 7.3 Sync PEDIDOS → Relatorio_DB (`sincronizarDados`)
- Garante headers do DB (`_garantirHeadersRelatorio_DB_`)
- Cria mapas: `fonteMap`, `dbMap`, impressões digitais, `codigoFixoMap`
- Lê `DADOS_IMPORTADOS` **de novo, direto da aba** para o índice de identidade `fonteIdent`
  (`CLIENTE|PEDIDO|PRODUTO|TAM`) e a contagem OC+OS. A importação usa a mesma trava, então não muda no meio.
- Para cada item do DB, tenta match por: **ID → UUID (CODIGO_FIXO) → fingerprint → troca de OC**
- **Executado em DUAS PASSADAS** (ver 15.12): Passada 1 resolve ID e UUID para
  todos os itens do DB antes de qualquer fingerprint rodar; Passada 2 resolve
  fingerprint e "não encontrado" só para quem sobrou sem match na Passada 1.
- Passada 2 no modo **ATIVO**: item `Faturado`/`Finalizado` **não** reivindica linha livre de PEDIDOS
  (nem por fingerprint nem por troca de OC — a linha livre é item novo e ficaria escondida); entre irmãs
  livres escolhe por LOTE → QTD (a do DB ou a do início do ciclo de baixas) → mais próxima
  (`_escolherVagaIrma_`); troca de OC com várias candidatas desempata por LOTE → QTD e, se os dados não
  separarem, só move item intocado (aberto, QTD>0, sem marcação) — auditoria `TROCA_OC_DESEMPATE`.
  No modo **SIMULACAO** fica a regra antiga (primeira vaga livre; troca de OC só com 1 candidata) e a
  auditoria registra `FINALIZADO_CAPTURA_SIMULADO`, `VAGA_IRMA_SIMULADA` e `TROCA_OC_SIMULADA`.
- Se o item do DB não casou com nada (saiu de PEDIDOS) — **regra única desde 30/09/2026, nos dois modos**:
  - marcado pelo usuário (MARCAR_FATURAR=SIM **e** col V preenchida) → `Faturado` + col Z
    `FATURADO|{usuário}` — só se a fonte for confiável nesta rodada; a marcação é relida logo antes de
    gravar (`_mesclarAlteracoesConcorrentes_`): desmarcado no meio do sync não fatura
  - col Z já `PENDENTE` ou `ABERTO` → não faz nada (aguarda o usuário / usuário disse que continua aberto)
  - `_motivoSaidaDaFonte_` + fonte confiável → col Z = `PENDENTE|sistema|data|motivo`, status **continua
    Ativo** (vale para QTD>0, QTD=0, com ou sem baixa, "SIM" legado sem usuário)
  - senão → não faz nada ("aguarda reconsolidação")
  - **fonte confiável** (`fonteConfiavel`) = as 4 travas de antes: mínimo de linhas, fração máxima sem
    fonte, queda brusca da fonte, importação parcial. Numa rodada suspeita nada é faturado nem sinalizado.
  - `_motivoSaidaDaFonte_` = orçamento de excedente por identidade (linhas abertas no DB − linhas na
    fonte). Linhas `PENDENTE`/`ABERTO` ficam fora de todas as contagens (identidade, OC+OS, excedente).
  - `protecaoAtiva` (OC+OS) **só aparece no log** — não decide nada
- Item que volta a casar (ID/UUID/fingerprint/troca de OC) com a col Z = `PENDENTE`/`ABERTO` tem a col Z
  apagada (`PENDENCIA_RESOLVIDA`) — o item voltou à origem
- O sync **não** grava MARCAR_FATURAR="SIM" nem `Faturado` sem usuário; o aviso antigo segue desligado (15.8)
- Adiciona itens novos de PEDIDOS que não existem no DB (travas contra duplicata por fingerprint e por
  identidade sem OC — ambas usam DESCRIÇÃO). Vagas de fingerprint (`_montarVagasFp_`): linha
  `PENDENTE`/`ABERTO` nunca segura vaga; no modo ATIVO item finalizado também não (em SIMULACAO segura, e a
  auditoria registra `NOVO_BARRADO_SIMULADO` quando só um finalizado barrou a linha nova)
- Atualiza IDs no Baixas_Historico se IDs mudaram
- Antes de gravar, `_mesclarAlteracoesConcorrentes_` relê as linhas do DB: se o ID da linha mudou, cancela;
  colunas que o usuário alterou durante o sync são preservadas (só as colunas que o sync quis mudar entram)
- A limpeza de faturados **não** roda aqui: roda no início do processo completo (7.4 e 15.18)
- Termina com a sentinela (15.15) e grava o buffer de auditoria (`_gravarAuditoria_`)

### 7.4 Processo Completo (`processoAutomaticoCompleto` — produção: a cada 1 h + 90s após cada importação)
1. Verifica pausa (`_sistemaPausado_()`)
2. Adquire `LockService` (`waitLock(30000)`; se não conseguir, aborta a execução)
3. Limpeza de faturados **antes** do sync (só se `_deveLimparFaturadosAgora_()` — dias úteis, a partir da
   hora de B2) → ETAPA 0 `sincronizarPedidosComFonte(false)` (guard H2) → ETAPA 1 `verificarEGerarIDs` →
   ETAPA 2 `sincronizarDadosOtimizado` → limpar cache
4. `finally`: grava a auditoria e solta a trava
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

### Caminhos que gravam `Faturado` (desde 30/09/2026 — os únicos que existem)
| Caminho | Usuário registrado em |
|---|---|
| Sync: marcado pelo usuário (P=SIM + col V) + saiu da origem + fonte confiável | col V (preservada) + col Z `FATURADO\|{col V}\|data\|marcado pelo usuário` |
| Tela: aviso de conferência → botão **Faturado** (`confirmarSaidaFonte`, só nível TOTAL, validado no servidor) | col Z `FATURADO\|{login}\|data\|confirmado no aviso` |

Removidos em 30/09: faturamento por QTD=0, faturamento por `_motivoSaidaDaFonte_` (agora → `PENDENTE`),
`marcarFaturado()`, `confirmarTodosAlertas()`/`confirmarTodosAlertasMenu()` e o item de menu "(testes)".
`faturarItensForaDaFonte()` (menu "📤 Sinalizar itens que já saíram da origem") só grava `PENDENTE`.
**Teste de regressão de qualquer mudança:** `Faturado` com col V vazia e col Z sem `FATURADO` só pode
existir se veio de antes de 30/09 (a sentinela conta: `faturadosSemUsuario`).

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
- **Só é logada** (`protecaoAtiva`) — não bloqueia nada. Quem decide se uma saída da origem é tratada é o
  orçamento por identidade + as 4 travas de fonte confiável (7.3).
- Desde 30/09 as 4 travas valem para **todos** os caminhos: numa importação quebrada nada é faturado (nem o
  que o usuário marcou) nem sinalizado — a rodada seguinte decide com a fonte completa.

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
  Qualquer diferença conta: espaço extra, maiúscula/minúscula, acento.
  - Modo **ATIVO**: a reserva pela base do ID (7.2) mantém ID, UUID e histórico quando só a DESCRIÇÃO muda
    (inclusive em todas as irmãs de um grupo ao mesmo tempo) — auditoria `ID_POR_BASE`.
  - Modo **SIMULACAO**: gera ID + UUID novos; a linha antiga vira `PENDENTE` (aviso, com gêmea pelo mesmo
    PEDIDO+LOTE → botão "Duplicata") e a nova entra como item novo.
- Campos que identificam o item (mudar = item "novo"): CLIENTE, PEDIDO, DESCRIÇÃO, TAMANHO,
  ORD. COMPRA (troca de OC só é reconhecida com 1 candidato), CÓD. OS, DATA RECEB.
- Mudar LOTE, QTD, DT. ENTREGA, PRAZO ou CARTELA é seguro (só atualiza). Entre linhas-irmãs, no modo
  SIMULACAO (casamento antigo) uma edição ou reordenação pode trocar IDs; no modo ATIVO o planejador
  (7.2) mantém cada lote com o seu ID (testado com os dados de 28/09 — 15.16).
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

- O menu **"2. Ativar Geração Automática (importação 15 min, sync 1 h)"** (`instalarTriggerAutomatico`) e
  `instalarTriggerAutomaticoSilencioso` apagam os acionadores e recriam exatamente a configuração de produção.
- **Cota diária:** conta @gmail.com tem **90 min/dia** de execução total de acionadores (Workspace: 6 h).
  São ~96 importações + até ~120 processos completos por dia — todo trabalho adicionado ao sync
  precisa ser leve (em memória, sem leitura/escrita por item).
- **Choque de horário:** o acionador de hora em hora (:33) dispara ~1 min depois de uma importação (:32).
  Desde 30/09 a importação usa a mesma trava do sync: um espera o outro (ou a importação pula a rodada).
- **Limpeza de faturados:** roda na primeira execução do processo completo a partir da hora de
  `CONFIGURAÇÕES!B2`, só em dias úteis — não no minuto exato (Apps Script também não garante minuto exato).
- Onde este documento ou o código falam em "próximo ciclo", leia até 1 h (ou ~15 min se o sync pós-importação rodar).

**Guard de execução dupla:** `LockService.getScriptLock()` — processo completo (`waitLock(30000)`), importação
(`tryLock(20000)`) e ferramentas de menu. Ações de usuário não pegam a trava; o sync protege o que elas
gravam relendo as linhas antes de escrever (`_mesclarAlteracoesConcorrentes_`). `confirmarSaidaFonte` usa a
trava do documento (respostas simultâneas ao aviso).  
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
registrarLoteEmissao(itensData)             // persiste lotes em col W — itensData: {planilhaLinha, uniqueId, loteId}
confirmarSaidaFonte(id, linha, decisao, usuario) // resposta ao aviso: FATURADO | CANCELADO | DUPLICATA | ABERTO
```

### Aviso de conferência (itens que saíram da origem sem marcação)
- `fetchAllDataUnified` devolve `pendentesSaida` (itens abertos com col Z = `PENDENTE`: id, linha, cliente,
  pedido, OC, descrição, tamanho, lote, QTD aberta/original, `saiuEm`, `motivo`, `gemeo` = outra linha aberta
  do mesmo PEDIDO com o mesmo LOTE) e `stats.pendentesSaida`.
- `#conferenciaSaidaModal` abre sozinho para nível TOTAL depois do login e quando aparece pendência nova numa
  recarga (não abre por cima de outra janela). Botões: **Faturado** (→ `Faturado`), **Cancelado** (→
  `Excluido`), **Continua aberto** (col Z `ABERTO`, não pergunta de novo), **Duplicata** (só com gêmea →
  `Excluido`). Pede confirmação antes de mudar o status. "Responder depois" fecha.
- Contador `#pendentes-saida-pill` no cabeçalho; selo "⏸️ SAIU DA ORIGEM — CONFERIR" na descrição do item
  (clicável) e "✔ conferido" para `ABERTO`. Nível PARCIAL não vê janela nem contador (o servidor também recusa).
- O usuário enviado é o login da sessão (aba CADASTRO), não o perfil BAIXA1…BAIXA5.

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
- Todas as ações por linha (`aplicarBaixa`, `estornarBaixa`, `editarUltimaBaixa`, `marcarParaFaturar`,
  `excluirItem`, `finalizarItem`, `registrarLoteEmissao`, `registrarCheckpointsFaturamento`,
  `confirmarSaidaFonte`) passam por `_resolverLinhaDoItem_(sheet, id, linha)`: confere se a linha ainda é
  daquele ID; se não for, procura o ID na coluna A; se não achar, devolve "recarregue a tela". Em baixa,
  estorno e edição a linha é resolvida **antes** de mexer no Baixas_Historico.

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
propriedade do script (limite de 9 KB por valor). **Não reativar** — foi substituído em 30/09/2026 pelo
aviso de conferência (seção 14), com a pendência gravada na coluna Z do DB e sem bloquear a tela.

### 15.9 IDs com sufixo numérico podem mudar
Se itens forem reordenados em DADOS_IMPORTADOS, o sufixo de um item pode mudar na próxima sync (ex: `-1` vira `-2`).  
O **UUID (CODIGO_FIXO)** garante que o histórico de baixas seja preservado mesmo assim — **exceto entre
linhas-irmãs**: o UUID acompanha a vaga escolhida no matching, então migra junto com o ID (15.16). No modo
ATIVO o planejador de irmãs escolhe a vaga pelos dados da linha (LOTE, QTD…), não pela ordem da aba.

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
- `faturarItensForaDaFonte()` (menu "📤 Sinalizar itens que já saíram da origem") — adianta o que o sync
  decide a cada ciclo (mesmo orçamento por identidade), mas desde 30/09 **só grava col Z = `PENDENTE`**
  (nunca `Faturado`); pula itens marcados pelo usuário e itens já `PENDENTE`/`ABERTO`. Mesmas travas do
  sync: `MIN_LINHAS_FONTE_PARA_FATURAR` e `MAX_FRACAO_SEM_FONTE` abortam se a fonte parecer incompleta.
  Auditoria na aba `Itens_Fora_Da_Fonte`.

**Checagens 6–8 da sentinela (desde 30/09):** `faturadosSemUsuario` (Faturado com col V vazia e col Z sem
`FATURADO` — deve ser 0 depois do reparo), `lotesSuspeitos` (mesmo PEDIDO+LOTE numa linha aberta e numa
faturada sem usuário), `ativosZeradosComSaldo` (aberto, sem marcação, QTD 0 no DB e >0 em PEDIDOS). Linhas
`PENDENTE`/`ABERTO` ficam fora das checagens 3–5 (já estão no aviso). O banner da tela mostra as três.

**Não portado do Bahia:** a coluna `INFO_Y` (Bahia usa a coluna Y da fonte para "conta e
ordem" — transferência entre filiais; no Ceará a coluna Y da fonte já é `LOTE`, usada para
outra coisa — não há dado de origem equivalente) e o rodapé de etiquetas NF (dados da
Marfim Bahia; o Ceará já tem os dados corretos da Marfim Ceará). A retenção de itens
Faturados ficou em `DIAS_RETENCAO = 15` (Bahia mudou para 1 dia). **Em 28/09/2026 o usuário decidiu que
o Ceará também deve apagar todos os faturados diariamente** — regra 1.1.5, pendente de implementação (15.18).

### 15.16 Incidente 28/09/2026: `Faturado` sem usuário + limite de identidade (linhas-irmãs) — CORRIGIDO (30/09)
**Situação em 30/09:** o sync não fatura mais sem usuário (nos dois modos); o planejador de irmãs, a recuperação
só com itens abertos e a Passada 2 que não deixa item finalizado capturar linha livre fecham o mecanismo abaixo
no modo **ATIVO** (7.2/7.3). Os 20 itens do incidente são tratados pelo menu "🩹 Reparar faturados sem usuário":
11 voltam para `Ativo` + `PENDENTE` (aviso) e 9 (os pares com gêmea viva) vão para a aba `Reparo_Faturados`
para decisão manual — conferir o Baixas_Historico dos IDs `-2/-3/-12` e `-2/-3/-15` antes de decidir.
Reproduzido em teste local com os CSVs de 28/09: no modo SIMULACAO a irmã do lote 139880 (marcada) toma o ID
da irmã 139881 e esta vira aviso falso; no modo ATIVO cada lote fica com o seu ID e só o marcado é faturado.

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

**Não rodar `corrigirFaturadosComSaldoAberto()` como correção:** reverte para Ativo sem pendência; o reparo
correto é o menu "🩹 Reparar faturados sem usuário" (o sync atual não refatura nada sozinho).
**Faturado sem usuário nunca é apagado pela limpeza** (15.18) — a evidência fica até alguém tratar.

**Limite que continua (1.1.4):** linhas-irmãs 100% idênticas (mesmo LOTE e QTD) não têm como ser distinguidas.
Com o planejador, a troca entre elas não muda nada visível (os dados são iguais) e a preferência pelo estado no DB
faz a irmã marcada/zerada ser a que "saiu" — o dano máximo é um aviso, nunca um faturamento.

### 15.17 Concorrência e número de linha — CORRIGIDO (30/09)
- Ações por linha conferem o ID (`_resolverLinhaDoItem_`, seção 14): exclusão de linhas não faz mais a tela
  agir sobre o item errado.
- `processoImportacao` usa a mesma trava do sync (7.1): o processo completo não vê duas versões da origem.
- `_mesclarAlteracoesConcorrentes_` relê as linhas antes de gravar: preserva o que o usuário mudou durante o sync
  e cancela o faturamento se a marcação sumiu nesse meio-tempo (auditoria `FATURAMENTO_CANCELADO`).

### 15.18 Limpeza de faturados
- Roda no **início** do processo completo (antes do sync), dias úteis, na primeira execução a partir da hora de
  `CONFIGURAÇÕES!B2` (padrão 11). Menu "🧹 Limpar Faturados agora" roda na hora.
- **ATIVO:** apaga **todos** os `Faturado` e `Excluido` (regra 1.1.5), em blocos contíguos de baixo para cima
  (`deleteRows`); `Finalizado` fica. Auditoria `LIMPEZA`.
- **SIMULACAO:** regra antiga (`Faturado/Finalizado/Excluido` com `DATA_STATUS` de 15+ dias) e auditoria
  `LIMPEZA_SIMULADA` com quantas linhas a regra nova apagaria.
- Nos dois modos, `Faturado` **sem usuário** (col V vazia e col Z sem `FATURADO`) não é apagado
  (auditoria `LIMPEZA_RETIDA`). `Excluido` por "Cancelado"/"Duplicata" do aviso some na limpeza seguinte:
  se a linha voltar à origem depois disso, entra como item novo.
- A3/A4 da aba são texto regravado pelo código — não controlam nada.

### 15.19 Modo SIMULACAO × ATIVO (`CONFIGURAÇÕES!B4`) e auditoria
| Muda com B4 | SIMULACAO (padrão) | ATIVO |
|---|---|---|
| Casamento de linhas-irmãs em PEDIDOS | antigo (guloso) + `IDENTIDADE_SIMULADA` | planejador (7.2) |
| Recuperação de ID pelo DB | qualquer status | só itens abertos, nunca ID que está em PEDIDOS |
| Reserva pela base do ID (DESCRIÇÃO corrigida) | não | sim (`ID_POR_BASE`) |
| Item finalizado captura linha livre (Passada 2) | sim + `FINALIZADO_CAPTURA_SIMULADO` | não |
| Escolha entre irmãs livres / troca de OC com várias candidatas | primeira / não aplica + `VAGA_IRMA_SIMULADA`, `TROCA_OC_SIMULADA` | LOTE → QTD (`TROCA_OC_DESEMPATE`) |
| Item finalizado segura vaga contra item novo | sim + `NOVO_BARRADO_SIMULADO` | não |
| Limpeza diária | 15 dias + `LIMPEZA_SIMULADA` | tudo (`LIMPEZA`, `LIMPEZA_RETIDA`) |

**Não depende de B4 (vale desde a publicação):** nenhum `Faturado` sem usuário; `PENDENTE` + aviso na tela;
`PENDENTE`/`ABERTO` fora das contagens e das vagas; trava na importação; conferência de ID por linha; releitura
antes de gravar; sentinela 6–8; reparo; remoção de `marcarFaturado`/`confirmarTodosAlertas`; instaladores 15 min/1 h.

Outros tipos na auditoria: `PENDENTE`, `PENDENCIA_RESOLVIDA`, `FATURADO_POR_MARCACAO`, `FATURAMENTO_CANCELADO`,
`CONFERENCIA_FATURADO|CANCELADO|DUPLICATA|ABERTO`, `REPARO_REABERTO`, `REPARO_MANUAL`.

**Como validar o modo novo antes de ligar:** com B4 = SIMULACAO, olhar `Auditoria_Sincronizacao` por 3–5 dias
úteis. Cada `IDENTIDADE_SIMULADA` mostra a linha da origem, o ID de hoje e o que o ATIVO daria — conferir alguns
pelo LOTE. Ligar ATIVO só com `faturadosSemUsuario` = 0 na sentinela (depois do reparo) e sem diferença
inexplicável. Teste local (Node + mock do Apps Script, 12 cenários + dados de 28/09) passou nos dois modos.

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
7. **Apaga linhas?** → Ações de usuário usam número de linha — conferir o ID com `_resolverLinhaDoItem_` (15.17)
8. **Muda identidade ou limpeza?** → Colocar atrás de `_modoAtivo_()` com auditoria no modo SIMULACAO (15.19) e
   rodar os cenários de teste (linhas-irmãs, DESCRIÇÃO corrigida, troca de OC, lote novo ao lado de faturado)

**Funções com maior superfície de impacto** (cuidado máximo):
- `sincronizarDados()` — toca todos os itens do DB
- `sincronizarPedidosComFonte()` — reescreve aba PEDIDOS inteira
- `calcularQtdOriginal()` / `_getSaldoEfetivoCache_()` — usadas no relatório de faturamento
- `_criarImpressaoDigital_()` — identidade de cada item
- `_planejarCasamentoIrmas_()` / `_custoIrma_()` — quem fica com qual ID entre linhas-irmãs (modo ATIVO)

---

## 17. IMPLANTAÇÃO DO PLANO DE 28/09 (implementado em 30/09/2026 — branch `claude/adoring-ride-lrmytw`)

Objetivo: cumprir 1.1.3 (nenhum `Faturado` sem usuário) e 1.1.5 (limpeza diária) sem mudar a rotina dos usuários.

| # | Item | Onde | Modo |
|---|---|---|---|
| 1 | Regra única de faturamento; saída sem marcação → col Z `PENDENTE` | `sincronizarDados` (7.3) | os dois |
| 2 | Aviso na tela + selo + contador; `confirmarSaidaFonte` (nível TOTAL no servidor) | index.html (14) | os dois |
| 3 | Removidos `marcarFaturado`, `confirmarTodosAlertas(+Menu)`; `faturarItensForaDaFonte` só sinaliza | seção 9 | os dois |
| 4a | Planejador de linhas-irmãs (+ Dilly, Passada 2, troca de OC) | `_planejarCasamentoIrmas_`, `_escolherVagaIrma_` | ATIVO |
| 4b | Recuperação de ID e vagas de fingerprint sem item finalizado | 7.2 / 7.3 | ATIVO |
| 4c | Reserva pela base do ID (DESCRIÇÃO corrigida) | `_idBaseFonte_` | ATIVO |
| 5 | Trava na importação; releitura antes de gravar | 7.1, `_mesclarAlteracoesConcorrentes_` | os dois |
| 6 | Conferência do ID nas ações por linha | `_resolverLinhaDoItem_` | os dois |
| 7 | Limpeza diária total, antes do sync, em blocos | `purgarItensFinalizados` (15.18) | ATIVO (SIMULACAO: 15 dias) |
| 8 | Sentinela 6–8 | `_verificarIntegridadeDuplicatas_` (15.15) | os dois |
| 9 | Modo simulação com auditoria | `_modoAtivo_`, `Auditoria_Sincronizacao` (15.19) | — |
| 10 | Reparo dos faturados sem usuário | menu 🩹 `repararFaturadosSemUsuario` | os dois |

**Sequência de implantação (usuário):**
1. Publicar `Código.gs` e `index.html` no Apps Script (nova versão da implantação web).
2. Abrir a planilha: a aba `CONFIGURAÇÕES` ganha A4/B4 (`SIMULACAO`); o Relatorio_DB ganha a coluna Z no
   primeiro sync. Conferir em Acionadores: importação 15 min e processo completo 1 h.
3. Menu "🩹 Reparar faturados sem usuário" (uma vez): os 11 voltam para o aviso; decidir à mão os 9 da aba
   `Reparo_Faturados` (conferir o Baixas_Historico dos IDs da gêmea antes).
4. 3–5 dias úteis em SIMULACAO lendo `Auditoria_Sincronizacao` (15.19).
5. `CONFIGURAÇÕES!B4 = ATIVO`. A primeira limpeza em ATIVO apaga de uma vez todos os Faturado/Excluido com
   usuário (em 28/09 seriam ~580 linhas) — avisar os usuários para imprimir etiquetas pendentes antes.
6. Voltar a SIMULACAO desfaz o modo novo na hora (não mexe em dados já gravados).

**Garantias (por construção + sentinela + teste local):** nenhum `Faturado` sem usuário; item que sai da origem
sem marcação nunca some calado (vai para o aviso); numa importação suspeita nada é faturado nem sinalizado; cada
aviso é perguntado uma vez só ("Continua aberto" não volta); com ATIVO, linha nova nunca herda item finalizado e
lote que não mudou não troca de ID. **Não garantível** (limite da origem, 1.1.4): saber qual linha é qual entre
irmãs 100% idênticas — o dano fica contido a um aviso, nunca a um faturamento.

**Pendências conhecidas:** decisão manual dos 9 pares do incidente; `corrigirFaturadosComSaldoAberto()` segue no
menu (reverte todo Faturado com QTD>0 para Ativo sem pendência — não usar; o sync depois os manda para o aviso);
a página não é responsiva no celular (tabelas largas; o aviso é).

### Ideias já avaliadas e DESCARTADAS (não repropor sem fato novo)
- Gravar código único na planilha de origem — proibido (1.1.4).
- Usar a coluna A da origem como ID — é numeração contínua 1…N, renumerada.
- Aba/"banco" de códigos com chave CÓD. CLIENTE + TAMANHO + DESCRIÇÃO (± PEDIDO, OC, QTD ou UUID nosso) — colide
  em 94% das linhas (83% com PEDIDO); QTD muda com baixas; incluir um UUID gerado por nós na chave é circular
  (a linha importada não traz o UUID).
- Mudar a fórmula do ID_UNICO — migra todos os IDs e o vínculo do Baixas_Historico sem resolver as linhas-irmãs.
- "Faturar primeiro as irmãs com QTD 0" — elas estão casadas por ID e nunca chegam à decisão.
- Reativar o sistema de alertas antigo (PropertiesService + senha) — desligado em `ce75202`; substituído pelo aviso da seção 14.
- `corrigirFaturadosComSaldoAberto()` como correção dos 20 — tira o registro sem pedir decisão; usar o reparo (menu 🩹).
- Só esconder os faturados na tela e apagar à noite mantendo 15 dias — o usuário quer apagar tudo diariamente (1.1.5).
