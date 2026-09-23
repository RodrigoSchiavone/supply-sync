import os
import json
import logging
import pandas as pd
from typing import List, Set
from config import ARQUIVO_LOJAS_BASE, PASTA_FINAL_ESTOQUE, PASTA_FINAL_VENDAS, CAMINHO_JSON_REPROCESSAR

def carregar_codigos_lojas_referencia() -> Set[str]:
    """Lê a primeira coluna (Código) da planilha base de lojas."""
    caminho_win = os.path.normpath(ARQUIVO_LOJAS_BASE)
    if not os.path.exists(caminho_win):
        logging.error(f"Arquivo base de lojas não encontrado em: {caminho_win}")
        return set()

    try:
        # Lê a primeira coluna 'Cod' (Coluna A)
        df_lojas = pd.read_excel(caminho_win, usecols=[0], engine="openpyxl")
        col_nome = df_lojas.columns[0]
        lojas = set(df_lojas[col_nome].dropna().astype(str).str.strip().tolist())
        logging.info(f"Carregadas {len(lojas)} lojas de referência da base.")
        return lojas
    except Exception as e:
        logging.error(f"Erro ao ler a planilha base de lojas ({caminho_win}): {e}")
        return set()

def extrair_lojas_do_arquivo(caminho_arquivo: str, coluna_idx: int) -> Set[str]:
    """Extrai os códigos únicos de loja de um arquivo gerado (ignora cabeçalhos e textos informativos)."""
    try:
        df = pd.read_excel(caminho_arquivo, usecols=[coluna_idx], engine="openpyxl")
        col_nome = df.columns[0]
        
        # Converte para string e limpa valores
        valores = df[col_nome].dropna().astype(str).str.strip().tolist()
        
        # Filtra apenas códigos numéricos válidos (ex: '800', '801')
        lojas_encontradas = {v for v in valores if v.isdigit()}
        return lojas_encontradas
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo {caminho_arquivo} para auditoria: {e}")
        return set()

def gerenciar_arquivos_pendentes_de_reprocessamento() -> List[str]:
    """Verifica se existem arquivos marcados anteriormente para remoção e reprocessamento."""
    if not os.path.exists(CAMINHO_JSON_REPROCESSAR):
        return []

    try:
        with open(CAMINHO_JSON_REPROCESSAR, "r", encoding="utf-8") as f:
            dados = json.load(f)
        
        arquivos_para_remover = dados.get("arquivos_incompletos", [])
        removidos = []

        for caminho in arquivos_para_remover:
            caminho_win = os.path.normpath(caminho)
            if os.path.exists(caminho_win):
                try:
                    os.remove(caminho_win)
                    removidos.append(caminho_win)
                    logging.info(f"Arquivo inconsistente removido para reprocessamento: {caminho_win}")
                except Exception as e:
                    logging.error(f"Não foi possível remover o arquivo {caminho_win}: {e}")
        
        # Limpa o arquivo JSON após remoção
        os.remove(CAMINHO_JSON_REPROCESSAR)
        return removidos
    except Exception as e:
        logging.error(f"Erro ao processar arquivo auxiliar de reprocessamento: {e}")
        return []

def auditar_arquivos_gerados():
    """Valida se todas as lojas estão presentes nos arquivos de Estoque e Vendas."""
    logging.info("\n--- INICIANDO AUDITORIA FINAL DE LOJAS ---")
    
    lojas_referencia = carregar_codigos_lojas_referencia()
    if not lojas_referencia:
        logging.warning("Auditoria ignorada: Lista de lojas de referência vazia ou inacessível.")
        return

    arquivos_com_falha = []

    # 1. Validação de Vendas (Coluna E = índice 4)[cite: 2]
    if os.path.exists(PASTA_FINAL_VENDAS):
        for f in os.listdir(PASTA_FINAL_VENDAS):
            if f.endswith(".xlsx") and not f.startswith("~$"):
                caminho_arq = os.path.join(PASTA_FINAL_VENDAS, f)
                lojas_presentes = extrair_lojas_do_arquivo(caminho_arq, coluna_idx=4)
                
                lojas_faltantes = lojas_referencia - lojas_presentes
                if lojas_faltantes:
                    msg = f"Inconsistência em Vendas [{f}]: Faltam {len(lojas_faltantes)} lojas ({', '.join(sorted(lojas_faltantes))})"
                    logging.warning(msg)
                    arquivos_com_falha.append(os.path.normpath(caminho_arq))
                else:
                    logging.info(f"Auditoria OK em Vendas [{f}]: Todas as {len(lojas_referencia)} lojas presentes.")

    # 2. Validação de Estoque (Coluna B = índice 1)[cite: 3]
    if os.path.exists(PASTA_FINAL_ESTOQUE):
        for f in os.listdir(PASTA_FINAL_ESTOQUE):
            if f.endswith(".xlsx") and not f.startswith("~$"):
                caminho_arq = os.path.join(PASTA_FINAL_ESTOQUE, f)
                lojas_presentes = extrair_lojas_do_arquivo(caminho_arq, coluna_idx=1)
                
                lojas_faltantes = lojas_referencia - lojas_presentes
                if lojas_faltantes:
                    msg = f"Inconsistência em Estoque [{f}]: Faltam {len(lojas_faltantes)} lojas ({', '.join(sorted(lojas_faltantes))})"
                    logging.warning(msg)
                    arquivos_com_falha.append(os.path.normpath(caminho_arq))
                else:
                    logging.info(f"Auditoria OK em Estoque [{f}]: Todas as {len(lojas_referencia)} lojas presentes.")

    # Registra no JSON auxiliar caso haja falhas
    if arquivos_com_falha:
        conteudo_json = {
            "descricao": "Arquivos marcados para exclusão e reprocessamento automático na próxima execução.",
            "arquivos_incompletos": arquivos_com_falha
        }
        with open(CAMINHO_JSON_REPROCESSAR, "w", encoding="utf-8") as f:
            json.dump(conteudo_json, f, indent=4, ensure_ascii=False)
        logging.warning(f"Foram encontrados {len(arquivos_com_falha)} arquivo(s) inconsistente(s). Registrados em: {CAMINHO_JSON_REPROCESSAR}")
    else:
        logging.info("Auditoria final concluída com sucesso! Todos os arquivos contêm a totalidade das lojas.")