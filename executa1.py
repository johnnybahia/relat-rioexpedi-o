import pdfplumber
import re
import os
import shutil
import requests
import json
from datetime import datetime

# ================= CONFIGURAÇÃO =================
URL_WEBAPP = "https://script.google.com/macros/s/AKfycbzke-sTVigX4hkUkqLaTYRL0WDi_P-JAhc4PPsjwf0GuwSz_92lx43fQVM07XiiNBrjbA/exec"
PASTA_ENTRADA = './pedidos'
PASTA_LIDOS = './pedidos/lidos'
LOG_ENVIADOS = 'pedidos_enviados.json'
# =================================================

def carregar_log():
    if os.path.exists(LOG_ENVIADOS):
        try:
            with open(LOG_ENVIADOS, 'r') as f: return json.load(f)
        except: return []
    return []

def salvar_no_log(novos_pedidos):
    log = carregar_log()
    for p in novos_pedidos:
        item = {"oc": p['ordemCompra'], "cliente": p['cliente']}
        if item not in log: log.append(item)
    with open(LOG_ENVIADOS, 'w') as f:
        json.dump(log, f, indent=4)

def converter_data_curta(data_str):
    if not data_str:
        return datetime.now().strftime("%d/%m/%Y")
    parts = data_str.strip().split('/')
    if len(parts) == 3 and len(parts[2]) == 2:
        return f"{parts[0]}/{parts[1]}/20{parts[2]}"
    return data_str

def limpar_valor_monetario(texto):
    if not texto: return 0.0
    texto = texto.lower().replace('r$', '').replace('total', '').strip()
    if ',' in texto and '.' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    try: return float(texto)
    except: return 0.0

# RESTAURADO: Suas regras originais de Unidade
def identificar_unidade(texto):
    texto_upper = texto.upper()
    if re.search(r'\d+,\d+\s*(PR|PRS|PAR|PARES)\b', texto_upper): return "PAR"
    if re.search(r'\d+,\d+\s*(M|MTS|METRO|METROS)\b', texto_upper): return "METRO"
    if re.search(r'\b(PR|PRS|PAR|PARES)\b', texto_upper): return "PAR"
    if re.search(r'\b(M|MTS|METRO|METROS)\b', texto_upper): return "METRO"
    return "UNID"

# ================= FUNÇÕES DE LOCAL =================

# RESTAURADO: Suas regras originais de Localização com o Regex genérico
def extrair_local_dilly(texto):
    texto_upper = texto.upper()
    match_generico = re.search(r',\s*([A-Z\s]+)-[A-Z]{2}', texto_upper)
    if match_generico:
        cidade_encontrada = match_generico.group(1).strip()
        if len(cidade_encontrada) < 30 and "MARFIM" not in cidade_encontrada:
            return cidade_encontrada.title()
    if "BREJO" in texto_upper: return "Brejo Santo"
    if "MORADA" in texto_upper: return "Morada Nova"
    if "QUIXERAMOBIM" in texto_upper: return "Quixeramobim"
    return "N/D"

def extrair_local_dass(texto):
    texto_upper = texto.upper()
    match_cabecalho = re.search(r'DASS\s+(NE-\d{2})\s+([A-ZÀ-ÿ]+)', texto_upper)
    if match_cabecalho:
        codigo = match_cabecalho.group(1)
        cidade = match_cabecalho.group(2)
        return f"{cidade.title()} ({codigo})"
    cidades_encontradas = re.findall(r'CIDADE:\s*([^\n]+)', texto_upper)
    match_codigo = re.search(r'(NE-\d{2})', texto_upper)
    codigo_str = f" ({match_codigo.group(1)})" if match_codigo else ""
    for c in cidades_encontradas:
        c_limpa = c.replace("- BRAZIL", "").replace("BRAZIL", "").split("-")[0].strip().split("CEP")[0].strip()
        if "EUSEBIO" in c_limpa or "CRUZ DAS ALMAS" in c_limpa or "MARFIM" in c_limpa:
            continue
        if len(c_limpa) > 3:
            return f"{c_limpa.title()}{codigo_str}"
    return "N/D"

def extrair_local_aniger(texto):
    texto_upper = texto.upper()
    if re.search(r'QUIXERAMOBIM', texto_upper): return "Quixeramobim"
    if re.search(r'IVOTI', texto_upper): return "Ivoti"
    return "N/D"

# ================= PROCESSAMENTO POR CLIENTE =================

def extrair_lotes_dilly(texto_completo, ordem_compra):
    """Extrai cada item de cada lote da OC Dilly para envio à aba Lotes_OC."""
    lotes_itens = []

    lote_matches = list(re.finditer(r'Lote:\s*(\d+)', texto_completo))
    print(f"  [DEBUG] Lotes encontrados: {len(lote_matches)} {[m.group(1) for m in lote_matches]}")
    if not lote_matches:
        # Mostra trecho do texto para diagnóstico
        print(f"  [DEBUG] Primeiros 500 chars do texto:\n{texto_completo[:500]}")

    for i, lote_match in enumerate(lote_matches):
        numero_lote = lote_match.group(1)
        inicio = lote_match.start()
        fim = lote_matches[i + 1].start() if i + 1 < len(lote_matches) else len(texto_completo)
        secao = texto_completo[inicio:fim]

        linhas = secao.split('\n')
        print(f"  [DEBUG] Lote {numero_lote}: {len(linhas)} linhas. Primeiras 5:")
        for dbg in linhas[:5]:
            print(f"    | {repr(dbg)}")
        j = 0
        while j < len(linhas):
            linha = linhas[j].strip()
            # Item line: 6-digit code + item num + description + "PR" unit + qty (X,XX) + price (X,XX)
            m = re.match(r'^(\d{6})\s+\d+\s+(.+?)\s+PR\s+(\d+[,.]\d+)\s+\d+[,.]\d+', linha)
            if m:
                codigo = m.group(1)
                descricao = re.sub(r'\s+', ' ', m.group(2)).strip()
                qtd_str = m.group(3).replace(',', '.')
                try:
                    qtd = int(float(qtd_str))
                except ValueError:
                    qtd = 0

                tamanho = ""
                largura = ""
                k = j + 1
                while k < min(j + 6, len(linhas)):
                    prox = linhas[k].strip()
                    if re.match(r'^\d{6}\s+\d+', prox):
                        break
                    m_tam = re.match(r'Tamanho\s+(\S+)', prox)
                    if m_tam:
                        tamanho = m_tam.group(1)
                    m_lar = re.match(r'Largura\s+(\S+)', prox)
                    if m_lar:
                        largura = m_lar.group(1)
                    k += 1

                lotes_itens.append({
                    "oc": ordem_compra,
                    "lote": numero_lote,
                    "codigo": codigo,
                    "descricao": descricao,
                    "tamanho": tamanho,
                    "largura": largura,
                    "qtd": qtd
                })
            j += 1

    return lotes_itens


def processar_dilly(texto_completo, nome_arquivo):
    # Data de Emissão [cite: 4]
    match_emissao = re.search(r'Data Emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo)
    data_rec = match_emissao.group(1) if match_emissao else datetime.now().strftime("%d/%m/%Y")

    # Data de Entrega (Previsão) [cite: 24]
    match_entrega_tab = re.search(r'Previsão.*?(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
    data_ped = match_entrega_tab.group(1) if match_entrega_tab else data_rec

    # MELHORIA: Captura marca completa entre "Marca:" e "Ref.:"
    match_marca = re.search(r'Marca:\s*(.*?)\s*Ref\.:', texto_completo, re.IGNORECASE)
    marca = match_marca.group(1).strip() if match_marca else "DILLY"

    # MELHORIA: Captura Quantidade Total no rodapé
    match_qtd = re.search(r'Quantidade Total:\s*([\d\.,]+)', texto_completo)
    qtd = int(limpar_valor_monetario(match_qtd.group(1))) if match_qtd else 0

    # Valor Total [cite: 194]
    match_valor = re.search(r'Total\s*R\$([\d\.,]+)', texto_completo)
    valor = limpar_valor_monetario(match_valor.group(1)) if match_valor else 0.0

    # MELHORIA: Captura OC de 6 dígitos no topo [cite: 1]
    match_ordem = re.search(r'Ordem\s*Compra\s*(\d{6})', texto_completo, re.IGNORECASE)
    ordem_compra = match_ordem.group(1) if match_ordem else "N/D"

    valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    lotes = extrair_lotes_dilly(texto_completo, ordem_compra)

    return {
        "dataPedido": data_ped, "dataRecebimento": data_rec, "arquivo": nome_arquivo,
        "cliente": "DILLY SPORTS", "marca": marca, "local": extrair_local_dilly(texto_completo),
        "qtd": qtd, "unidade": identificar_unidade(texto_completo),
        "valor": valor_formatado, "ordemCompra": ordem_compra,
        "_lotes": lotes
    }

# (Funções processar_aniger, processar_dass e processar_dakota seguem a lógica original que você enviou)

def processar_aniger(texto_completo, nome_arquivo):
    match_emissao = re.search(r'Emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo)
    if not match_emissao:
        match_emissao = re.search(r'Emissão:.*?(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
    data_rec_str = match_emissao.group(1) if match_emissao else datetime.now().strftime("%d/%m/%Y")
    todas_datas = re.findall(r'(\d{2}/\d{2}/\d{4})', texto_completo)
    data_ped_str = data_rec_str
    try:
        data_rec_obj = datetime.strptime(data_rec_str, "%d/%m/%Y")
        for d_str in todas_datas:
            try:
                d_obj = datetime.strptime(d_str, "%d/%m/%Y")
                if (d_obj - data_rec_obj).days > 5:
                    data_ped_str = d_str
                    break
            except: continue
    except: pass
    marca = "ANIGER"
    if "NIKE" in texto_completo.upper(): marca = "NIKE (Aniger)"
    qtd = 0
    valor = 0.0
    match_totais = re.search(r'Totais\s+([\d\.,]+).*?([\d\.,]+)', texto_completo, re.DOTALL)
    if match_totais:
        qtd = int(limpar_valor_monetario(match_totais.group(1)))
        valor = limpar_valor_monetario(match_totais.group(2))
    match_ordem = re.search(r'Ordem\s+(?:de\s+)?[Cc]ompra[\s\S]{0,50}?(\d{6,})', texto_completo)
    ordem_compra = match_ordem.group(1) if match_ordem else "N/D"
    valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return {
        "dataPedido": data_ped_str, "dataRecebimento": data_rec_str, "arquivo": nome_arquivo,
        "cliente": "ANIGER", "marca": marca, "local": extrair_local_aniger(texto_completo),
        "qtd": qtd, "unidade": identificar_unidade(texto_completo),
        "valor": valor_formatado, "ordemCompra": ordem_compra
    }

def processar_dass(texto_completo, nome_arquivo):
    match_emissao = re.search(r'Data da emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.IGNORECASE)
    if match_emissao:
        data_rec = match_emissao.group(1)
    else:
        match_header = re.search(r'Hora.*?Data\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
        data_rec = match_header.group(1) if match_header else datetime.now().strftime("%d/%m/%Y")
    idx_inicio = texto_completo.find("Prev. Ent.")
    texto_busca = texto_completo[idx_inicio:] if idx_inicio != -1 else texto_completo
    match_entrega = re.search(r'\d{8}.*?(\d{2}/\d{2}/\d{4})', texto_busca, re.DOTALL)
    data_ped = match_entrega.group(1) if match_entrega else data_rec
    match_ordem = re.search(r'Ordem\s+(?:de\s+)?[Cc]ompra[\s\S]{0,50}?(\d{6,})', texto_completo)
    ordem_compra = match_ordem.group(1) if match_ordem else "N/D"
    match_marca = re.search(r'Marca:\s*([^\n]+)', texto_completo)
    marca = match_marca.group(1).strip() if match_marca else "N/D"
    valor = 0.0
    qtd = 0
    match_val = re.search(r'Total valor:\s*([\d\.,]+)', texto_completo)
    if match_val: valor = limpar_valor_monetario(match_val.group(1))
    match_qtd = re.search(r'Total peças:\s*([\d\.,]+)', texto_completo)
    if match_qtd: qtd = int(limpar_valor_monetario(match_qtd.group(1)))
    valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return {
        "dataPedido": data_ped, "dataRecebimento": data_rec, "arquivo": nome_arquivo,
        "cliente": "Grupo DASS", "marca": marca, "local": extrair_local_dass(texto_completo),
        "qtd": qtd, "unidade": identificar_unidade(texto_completo),
        "valor": valor_formatado, "ordemCompra": ordem_compra
    }

def processar_dakota(pages, nome_arquivo):
    pedidos = []
    compradores_conhecidos = ('saimon', 'ccarlos')
    for page in pages:
        tables = page.extract_tables()
        if not tables: continue
        for table in tables:
            for row in table:
                if not row: continue
                # Detecta deslocamento: página 3+ tem None na col 0
                offset = 1 if (row[0] is None or (row[0] and len(str(row[0])) > 20)) else 0
                if len(row) < 8 + offset: continue
                primeiro = str(row[0 + offset] or '').strip()
                if not primeiro.isdigit(): continue

                filial = ""; oc = ""; datas = []; unidade = "UNID"; qtd = 0
                for cell in row[offset:]:
                    val = str(cell or '').strip()
                    if not val or val == primeiro: continue
                    if not oc and re.match(r'^\d+[A-Za-z]$', val): oc = val
                    elif re.match(r'^\d{2}/\d{2}/\d{2}$', val): datas.append(val)
                    elif val.upper() in ('PR', 'MT'): unidade = 'PAR' if val.upper() == 'PR' else 'METRO'
                    elif val.lower() in compradores_conhecidos: continue
                    elif not filial and re.match(r'^[A-ZÀ-ÿa-zà-ÿ\s]+$', val) and len(val.strip()) >= 4: filial = val
                if not oc: continue
                for cell in row[offset:]:
                    val = str(cell or '').strip()
                    if val == primeiro or not val: continue
                    if re.match(r'^[\d\.,]+$', val):
                        try:
                            num = int(float(val.replace('.', '').replace(',', '.')))
                            if num > 0: qtd = num; break
                        except: continue
                emissao = converter_data_curta(datas[0]) if datas else datetime.now().strftime("%d/%m/%Y")
                entrega = converter_data_curta(datas[1]) if len(datas) >= 2 else emissao
                pedidos.append({
                    "dataPedido": entrega, "dataRecebimento": emissao, "arquivo": nome_arquivo,
                    "cliente": "DAKOTA", "marca": "DAKOTA (Todas)",
                    "local": filial.strip().title(), "qtd": qtd, "unidade": unidade,
                    "valor": "R$ 0,00", "ordemCompra": oc
                })
    return pedidos if pedidos else None
# ================= CONTROLADOR PRINCIPAL =================

def processar_pdf_inteligente(caminho_arquivo, nome_arquivo):
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() or ""
            texto_upper = texto_completo.upper()

            if "DILLY" in texto_upper:
                return [processar_dilly(texto_completo, nome_arquivo)]
            elif "ANIGER" in texto_upper:
                return [processar_aniger(texto_completo, nome_arquivo)]
            elif "DASS" in texto_upper or "01287588" in texto_completo:
                return [processar_dass(texto_completo, nome_arquivo)]
            elif "DAKOTA" in texto_upper:
                return processar_dakota(pdf.pages, nome_arquivo)
            else:
                return None
    except Exception as e:
        print(f"Erro ao abrir {nome_arquivo}: {e}")
        return []

def mover_arquivos_processados(lista_arquivos):
    if not os.path.exists(PASTA_LIDOS): os.makedirs(PASTA_LIDOS)
    print(f"\n📦 Movendo arquivos processados para: {PASTA_LIDOS}")
    for arquivo in set(lista_arquivos):
        try:
            shutil.move(os.path.join(PASTA_ENTRADA, arquivo), os.path.join(PASTA_LIDOS, arquivo))
        except: pass

# RESTAURADO: Tabela de visualização original no terminal
def main():
    if not os.path.exists(PASTA_ENTRADA):
        os.makedirs(PASTA_ENTRADA)
        print(f"Pasta criada. Coloque PDFs em '{PASTA_ENTRADA}'.")
        return

    log = carregar_log()
    ocs_enviadas = [item['oc'] for item in log]
    todos_pedidos_para_envio = []
    arquivos_para_mover = []
    arquivos = [f for f in os.listdir(PASTA_ENTRADA) if f.lower().endswith('.pdf')]

    print(f"📂 Processando {len(arquivos)} arquivos...")
    print("-" * 95)
    print(f"{'EMISSÃO':<12} | {'ENTREGA':<12} | {'OC':<12} | {'CLIENTE':<15} | {'MARCA':<15} | {'VALOR'}")
    print("-" * 95)

    for arq in arquivos:
        lista_pedidos = processar_pdf_inteligente(os.path.join(PASTA_ENTRADA, arq), arq)
        if lista_pedidos:
            for p in lista_pedidos:
                if p['ordemCompra'] in ocs_enviadas:
                    print(f"⏭️  OC {p['ordemCompra']} já enviada anteriormente. Pulando.")
                    continue
                todos_pedidos_para_envio.append(p)
                # RESTAURADO: Seu print detalhado original
                print(f"✅ {p['dataRecebimento']:<12} | {p['dataPedido']:<12} | {p['ordemCompra']:<12} | {p['cliente'][:15]:<15} | {p['marca'][:15]:<15} | {p['valor']}")
            arquivos_para_mover.append(arq)
        else:
            print(f"⚠️  Ignorado: {arq}")

    if todos_pedidos_para_envio:
        print("-" * 95)
        todos_lotes_para_envio = []
        for p in todos_pedidos_para_envio:
            todos_lotes_para_envio.extend(p.pop("_lotes", []))

        print(f"📤 Enviando {len(todos_pedidos_para_envio)} pedidos e {len(todos_lotes_para_envio)} itens de lote para Google Sheets...")
        try:
            payload = {"pedidos": todos_pedidos_para_envio}
            if todos_lotes_para_envio:
                payload["lotes"] = todos_lotes_para_envio
            response = requests.post(URL_WEBAPP, json=payload, timeout=30)
            if response.status_code == 200:
                salvar_no_log(todos_pedidos_para_envio)
                print(f"☁️  SUCESSO! Google recebeu os dados.")
                mover_arquivos_processados(arquivos_para_mover)
            else:
                print(f"❌ Erro HTTP {response.status_code}")
        except Exception as e:
            print(f"\n❌ Erro de conexão: {e}")
    else:
        print("\n⚠️  Nenhum pedido novo encontrado.")

    input("\nPressione ENTER para fechar...")

if __name__ == "__main__":
    main()