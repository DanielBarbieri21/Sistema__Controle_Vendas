import pandas as pd
from sqlalchemy import insert, select
from database import engine
from models import vendas, itens_venda
import os
import re
import warnings
from datetime import datetime

# Suprimir avisos do openpyxl sobre estilos padrão
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def encontrar_linha_cabecalho(caminho, sheet_name=0):
    """Encontra automaticamente a linha onde estão os cabeçalhos das colunas"""
    # Tentar diferentes linhas como cabeçalho (header)
    for header_row in range(0, 20):
        try:
            df = pd.read_excel(caminho, sheet_name=sheet_name, header=header_row, nrows=1)
            colunas = [str(col).strip() for col in df.columns]
            
            # Procurar por palavras-chave específicas nas colunas
            palavras_chave_obrigatorias = ["venda", "comprador", "data"]
            encontrados = sum(1 for col in colunas for palavra in palavras_chave_obrigatorias 
                           if palavra.lower() in col.lower())
            
            # Verificar se encontrou "N.º de venda" ou "Número da venda" que é muito específico
            tem_numero_venda = any("n.º" in col.lower() or "número" in col.lower() or "numero" in col.lower() 
                                  for col in colunas if "venda" in col.lower())
            
            if encontrados >= 2 or tem_numero_venda:  # Se encontrar pelo menos 2 palavras-chave ou número de venda
                return header_row
        except:
            continue
    return None

def importar_excel(caminho):
    # Validar se o arquivo existe
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    
    # Validar extensão do arquivo
    if not caminho.lower().endswith(('.xlsx', '.xls')):
        raise ValueError("O arquivo deve ser um arquivo Excel (.xlsx ou .xls)")
    
    # Tentar ler todas as abas do Excel
    try:
        excel_file = pd.ExcelFile(caminho)
        sheet_names = excel_file.sheet_names
    except Exception as e:
        raise ValueError(f"Erro ao ler o arquivo Excel: {str(e)}. Verifique se o arquivo não está corrompido ou aberto em outro programa.")
    
    df = None
    linha_cabecalho = None
    aba_usada = None
    
    # Primeiro, tentar header=5 que é o padrão do Mercado Livre (mais comum)
    for sheet_name in sheet_names:
        try:
            df_temp = pd.read_excel(caminho, sheet_name=sheet_name, header=5)
            if not df_temp.empty:
                colunas = [str(col).strip().lower() for col in df_temp.columns]
                # Verificar se tem as colunas essenciais do Mercado Livre
                tem_numero_venda = any("n.º" in col or "número" in col or "numero" in col 
                                      for col in colunas if "venda" in col)
                tem_data = any("data" in col and "venda" in col for col in colunas)
                tem_comprador = any("comprador" in col for col in colunas)
                
                if tem_numero_venda and tem_data and tem_comprador:
                    df = df_temp
                    linha_cabecalho = 5
                    aba_usada = sheet_name
                    break
        except:
            continue
    
    # Se não encontrou com header=5, tentar encontrar automaticamente
    if df is None or df.empty:
        for sheet_name in sheet_names:
            try:
                linha_cabecalho = encontrar_linha_cabecalho(caminho, sheet_name)
                
                if linha_cabecalho is not None:
                    df_temp = pd.read_excel(caminho, sheet_name=sheet_name, header=linha_cabecalho)
                    
                    # Verificar se encontrou colunas relevantes
                    colunas = [str(col).strip().lower() for col in df_temp.columns]
                    palavras_chave = ["venda", "comprador", "cliente", "item", "produto", "anúncio", "anuncio"]
                    encontrados = sum(1 for col in colunas for palavra in palavras_chave if palavra in col)
                    
                    if encontrados >= 2 and not df_temp.empty:
                        df = df_temp
                        aba_usada = sheet_name
                        break
            except Exception as e:
                continue
    
    # Se ainda não encontrou, tentar com header=4 (padrão original)
    if df is None or df.empty:
        try:
            df = pd.read_excel(caminho, header=4)
            linha_cabecalho = 4
        except:
            pass
    
    # Se ainda não encontrou, tentar sem especificar header
    if df is None or df.empty:
        try:
            df = pd.read_excel(caminho)
            linha_cabecalho = 0
        except Exception as e:
            raise ValueError(f"Erro ao ler o arquivo Excel: {str(e)}. Não foi possível encontrar os dados de vendas no arquivo.")
    
    # Verificar se o DataFrame está vazio
    if df.empty:
        raise ValueError("O arquivo Excel está vazio ou não contém dados válidos.")
    
    # Mapeamento direto para colunas conhecidas do Mercado Livre
    # Primeiro verificar se as colunas exatas existem
    mapeamento_direto = {
        "numero_venda": ["N.º de venda"],
        "data_venda": ["Data da venda"],
        "cliente_nome": ["Comprador"],
        "produto": ["Título do anúncio"],
        "cliente_cidade": ["Cidade"],
        "cliente_estado": ["Estado.1"],
        "status_venda": ["Descrição do status"],
        "quantidade": ["Unidades"],
        "preco_unitario": ["Preço unitário de venda do anúncio (BRL)"],
        "valor_total": ["Total (BRL)"]
    }
    
    # Padrões de busca para cada coluna (ordem: mais específico primeiro)
    # Usado como fallback se mapeamento direto não funcionar
    padroes_colunas = {
        "numero_venda": ["n.º de venda", "número da venda", "numero da venda", "nº venda", "n° venda", "id venda"],
        "data_venda": ["data da venda", "data venda"],
        "cliente_nome": ["comprador"],
        "produto": ["título do anúncio", "titulo do anuncio"],
        "cliente_cidade": ["cidade"],
        "cliente_estado": ["estado.1", "estado"],
        "status_venda": ["descrição do status", "descrição status"],
        "quantidade": ["unidades"],
        "preco_unitario": ["preço unitário de venda do anúncio", "preço unitário venda anúncio"],
        "valor_total": ["total (brl)", "total brl"]
    }
    
    # Verificar quais colunas existem no DataFrame
    # Filtrar colunas "Unnamed" e vazias
    colunas_originais = []
    for col in df.columns.tolist():
        col_str = str(col).strip()
        colunas_originais.append(col_str)
    
    colunas_validas = []
    for col in colunas_originais:
        col_lower = col.lower()
        # Ignorar colunas Unnamed, vazias ou NaN
        if (not col_lower.startswith('unnamed') 
            and col != '' 
            and col_lower != 'nan'
            and col != 'None'):
            try:
                if not pd.isna(col):
                    colunas_validas.append(col)
            except:
                colunas_validas.append(col)
    
    # Se não encontrou colunas válidas, usar todas (pode ser que os dados estejam em colunas sem nome)
    if not colunas_validas:
        colunas_validas = colunas_originais
    
    colunas_necessarias = ["numero_venda", "data_venda", "cliente_nome", "produto"]
    colunas_encontradas = {}
    
    # Primeiro: tentar mapeamento direto (nomes exatos do Mercado Livre)
    for col_nova, nomes_possiveis in mapeamento_direto.items():
        if col_nova not in colunas_encontradas:
            for nome in nomes_possiveis:
                if nome in colunas_originais:
                    colunas_encontradas[col_nova] = nome
                    break
    
    # Segundo: buscar correspondências exatas usando padrões (case-insensitive)
    for col_original in colunas_validas + colunas_originais:
        if col_original in colunas_encontradas.values():
            continue
        col_original_lower = str(col_original).lower().strip()
        
        for col_nova, padroes in padroes_colunas.items():
            if col_nova not in colunas_encontradas:
                for padrao in padroes:
                    padrao_lower = padrao.lower()
                    # Correspondência exata (case-insensitive)
                    if col_original_lower == padrao_lower:
                        colunas_encontradas[col_nova] = col_original
                        break
                if col_nova in colunas_encontradas:
                    break
    
    # Terceiro: buscar correspondências parciais (apenas para colunas críticas ainda não encontradas)
    for col_original in colunas_validas + colunas_originais:
        if col_original in colunas_encontradas.values():
            continue
        col_original_lower = str(col_original).lower().strip()
        
        for col_nova, padroes in padroes_colunas.items():
            if col_nova not in colunas_encontradas:
                for padrao in padroes:
                    padrao_lower = padrao.lower()
                    # Contém o padrão (com validações específicas)
                    if padrao_lower in col_original_lower:
                        # Para "comprador", só aceitar se for exatamente "comprador"
                        if col_nova == "cliente_nome":
                            if col_original_lower == "comprador":
                                colunas_encontradas[col_nova] = col_original
                                break
                        # Para "produto", só aceitar se contiver "título" E "anúncio"
                        elif col_nova == "produto":
                            if "título" in col_original_lower and "anúncio" in col_original_lower:
                                colunas_encontradas[col_nova] = col_original
                                break
                        else:
                            colunas_encontradas[col_nova] = col_original
                            break
                if col_nova in colunas_encontradas:
                    break
    
    # Verificar se todas as colunas necessárias foram encontradas
    colunas_faltando = [col for col in colunas_necessarias if col not in colunas_encontradas]
    if colunas_faltando:
        # Criar mensagem mais amigável
        nomes_amigaveis = {
            "numero_venda": "Número da venda",
            "data_venda": "Data da venda",
            "cliente_nome": "Nome do comprador/cliente",
            "produto": "Título do item/produto"
        }
        faltando_nomes = [nomes_amigaveis.get(col, col) for col in colunas_faltando]
        
        # Mostrar colunas válidas (sem Unnamed) primeiro, depois outras
        colunas_mostrar = colunas_validas[:15] if colunas_validas else colunas_originais[:15]
        if len(colunas_validas) > 15:
            colunas_mostrar.append(f"... e mais {len(colunas_validas) - 15} colunas válidas")
        elif len(colunas_originais) > 15 and not colunas_validas:
            colunas_mostrar.append(f"... e mais {len(colunas_originais) - 15} colunas")
        
        raise ValueError(
            f"Colunas obrigatórias não encontradas no Excel:\n{', '.join(faltando_nomes)}\n\n"
            f"Colunas encontradas no arquivo (primeiras 20):\n{', '.join(colunas_mostrar)}\n\n"
            f"O arquivo parece ter uma estrutura diferente do esperado.\n"
            f"Verifique se o arquivo contém dados de vendas e se as colunas têm nomes similares aos esperados."
        )
    
    # Renomear colunas
    df_renamed = df.rename(columns={v: k for k, v in colunas_encontradas.items()})
    
    # Garantir que todas as colunas necessárias existam (preencher com valores padrão se necessário)
    for col in ["cliente_cidade", "cliente_estado", "status_venda", "quantidade", "preco_unitario", "valor_total"]:
        if col not in df_renamed.columns:
            df_renamed[col] = "" if col in ["cliente_cidade", "cliente_estado", "status_venda"] else 0
    
    # Converter data - formato do Mercado Livre: "2 de janeiro de 2026 09:18 hs."
    def converter_data_ml(valor):
        """Converte formato de data do Mercado Livre para datetime"""
        if pd.isna(valor) or valor == "" or valor is None:
            return None
        
        # Se já é datetime, retornar
        if isinstance(valor, datetime) or pd.api.types.is_datetime64_any_dtype(type(valor)):
            return pd.to_datetime(valor)
        
        valor_str = str(valor).strip()
        
        # Mapeamento de meses em português
        meses_pt = {
            'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4,
            'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
            'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
        }
        
        # Tentar padrão do Mercado Livre: "2 de janeiro de 2026 09:18 hs."
        # Padrão: número + " de " + mês + " de " + ano + " " + hora + ":" + minuto + " hs."
        padrao_ml = r'(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})\s+(\d{1,2}):(\d{2})\s+hs\.'
        match = re.search(padrao_ml, valor_str, re.IGNORECASE)
        
        if match:
            dia = int(match.group(1))
            mes_nome = match.group(2).lower()
            ano = int(match.group(3))
            hora = int(match.group(4))
            minuto = int(match.group(5))
            
            if mes_nome in meses_pt:
                mes = meses_pt[mes_nome]
                try:
                    return pd.Timestamp(year=ano, month=mes, day=dia, hour=hora, minute=minuto)
                except:
                    pass
        
        # Tentar padrão alternativo sem hora: "2 de janeiro de 2026"
        padrao_ml_sem_hora = r'(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})'
        match = re.search(padrao_ml_sem_hora, valor_str, re.IGNORECASE)
        
        if match:
            dia = int(match.group(1))
            mes_nome = match.group(2).lower()
            ano = int(match.group(3))
            
            if mes_nome in meses_pt:
                mes = meses_pt[mes_nome]
                try:
                    return pd.Timestamp(year=ano, month=mes, day=dia)
                except:
                    pass
        
        # Se não conseguir, tentar conversão padrão do pandas
        try:
            return pd.to_datetime(valor_str, errors='coerce')
        except:
            return None
    
    # Aplicar conversão de data
    df_renamed["data_venda"] = df_renamed["data_venda"].apply(converter_data_ml)
    
    # Remover linhas com dados inválidos nas colunas obrigatórias
    df_renamed = df_renamed.dropna(subset=["numero_venda", "cliente_nome", "produto"])
    
    # Verificar se restaram dados após a limpeza
    if df_renamed.empty:
        raise ValueError("Nenhum dado válido encontrado após a limpeza. Verifique se as colunas obrigatórias estão preenchidas.")
    
    # Limpar e converter valores numéricos
    def limpar_numero(valor, padrao=0):
        if pd.isna(valor) or valor == "" or valor is None:
            return padrao
        try:
            # Remover caracteres não numéricos (exceto ponto e vírgula)
            if isinstance(valor, str):
                valor = valor.replace(",", ".").replace("R$", "").replace("$", "").strip()
            return float(valor)
        except (ValueError, TypeError):
            return padrao
    
    def limpar_inteiro(valor, padrao=0):
        if pd.isna(valor) or valor == "" or valor is None:
            return padrao
        try:
            return int(float(valor))
        except (ValueError, TypeError):
            return padrao
    
    # Aplicar limpeza
    df_renamed["quantidade"] = df_renamed["quantidade"].apply(limpar_inteiro)
    df_renamed["preco_unitario"] = df_renamed["preco_unitario"].apply(limpar_numero)
    df_renamed["valor_total"] = df_renamed["valor_total"].apply(limpar_numero)
    
    # Limpar strings
    def limpar_string(valor, padrao=""):
        if pd.isna(valor) or valor is None:
            return padrao
        return str(valor).strip()
    
    df_renamed["numero_venda"] = df_renamed["numero_venda"].apply(limpar_string)
    df_renamed["cliente_nome"] = df_renamed["cliente_nome"].apply(limpar_string)
    df_renamed["produto"] = df_renamed["produto"].apply(limpar_string)
    df_renamed["cliente_cidade"] = df_renamed["cliente_cidade"].apply(limpar_string)
    df_renamed["cliente_estado"] = df_renamed["cliente_estado"].apply(limpar_string)
    df_renamed["status_venda"] = df_renamed["status_venda"].apply(limpar_string)

    conn = engine.connect()
    trans = conn.begin()
    linhas_importadas = 0

    # Verificar duplicatas antes de inserir (mais eficiente)
    numeros_venda_existentes = set()
    resultado_existentes = conn.execute(
        select(vendas.c.numero_venda)
    ).fetchall()
    numeros_venda_existentes = {str(r[0]) for r in resultado_existentes}
    
    vendas_novas = 0
    vendas_duplicadas = 0
    itens_duplicados = 0
    
    try:
        for idx, row in df_renamed.iterrows():
            try:
                numero_venda = str(row["numero_venda"])
                
                # Verificar se a venda já existe
                if numero_venda in numeros_venda_existentes:
                    # Buscar ID da venda existente
                    result = conn.execute(
                        select(vendas.c.id).where(
                            vendas.c.numero_venda == numero_venda
                        )
                    ).first()
                    venda_id = result[0]
                    vendas_duplicadas += 1
                    
                    # Verificar se o item já existe para evitar duplicação
                    item_existe = conn.execute(
                        select(itens_venda.c.id).where(
                            (itens_venda.c.venda_id == venda_id) &
                            (itens_venda.c.produto == str(row["produto"]))
                        )
                    ).first()
                    
                    if item_existe:
                        itens_duplicados += 1
                        continue  # Pular item duplicado
                else:
                    # Nova venda - adicionar ao set
                    numeros_venda_existentes.add(numero_venda)
                    vendas_novas += 1
                    
                    # Inserir nova venda
                    # Converter data_venda para formato adequado ao banco
                    data_venda = row["data_venda"]
                    if pd.isna(data_venda) or data_venda is None:
                        data_venda = None
                    elif isinstance(data_venda, pd.Timestamp):
                        # Converter Timestamp para date do Python
                        data_venda = data_venda.date()
                    elif isinstance(data_venda, str) and data_venda.strip() == "":
                        data_venda = None
                    
                    venda_result = conn.execute(
                        insert(vendas).values(
                            numero_venda=str(row["numero_venda"]),
                            data_venda=data_venda,
                            cliente_nome=str(row["cliente_nome"]),
                            cliente_cidade=str(row.get("cliente_cidade", "")),
                            cliente_estado=str(row.get("cliente_estado", "")),
                            status_venda=str(row.get("status_venda", "")),
                            valor_total=float(row.get("valor_total", 0))
                        )
                    )
                    venda_id = venda_result.inserted_primary_key[0]

                # Inserir item da venda
                conn.execute(
                    insert(itens_venda).values(
                        venda_id=venda_id,
                        produto=str(row["produto"]),
                        quantidade=int(row.get("quantidade", 0)),
                        preco_unitario=float(row.get("preco_unitario", 0))
                    )
                )
                linhas_importadas += 1
            except Exception as e:
                # Log do erro mas continua com a próxima linha
                print(f"Erro ao processar linha {idx + 5}: {str(e)}")
                continue

        trans.commit()
        
        # Se nenhuma linha foi importada, verificar o motivo
        if linhas_importadas == 0:
            total_linhas = len(df_renamed)
            mensagem_erro = f"Nenhuma linha foi importada.\n\n"
            mensagem_erro += f"Total de linhas no arquivo: {total_linhas}\n"
            
            if vendas_duplicadas > 0:
                mensagem_erro += f"Vendas duplicadas encontradas: {vendas_duplicadas}\n"
            if itens_duplicados > 0:
                mensagem_erro += f"Itens duplicados ignorados: {itens_duplicados}\n"
            
            if vendas_duplicadas == total_linhas or (vendas_duplicadas + itens_duplicados) >= total_linhas:
                mensagem_erro += "\nTodas as linhas já existem no banco de dados.\n"
                mensagem_erro += "Nenhum dado novo para importar."
            else:
                mensagem_erro += "\nVerifique os dados do arquivo Excel."
            
            raise ValueError(mensagem_erro)
        
        # Mensagem informativa sobre duplicatas (se houver importações)
        if (vendas_duplicadas > 0 or itens_duplicados > 0) and linhas_importadas > 0:
            mensagem_duplicatas = []
            if vendas_duplicadas > 0:
                mensagem_duplicatas.append(f"{vendas_duplicadas} venda(s) já existente(s)")
            if itens_duplicados > 0:
                mensagem_duplicatas.append(f"{itens_duplicados} item(ns) duplicado(s) ignorado(s)")
            print(f"Aviso: {'; '.join(mensagem_duplicatas)}")
    except Exception as e:
        trans.rollback()
        raise e
    finally:
        conn.close()
    
    return linhas_importadas

