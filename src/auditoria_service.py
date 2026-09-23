import os
import json
import logging
import pandas as pd
from typing import List, Set
from config import ARQUIVO_LOJAS_BASE, CAMINHO_JSON_REPROCESSAR

def carregar_codigos_lojas_referencia() -> Set[str]:
    """Lê a primeira coluna (Código) da planilha base de lojas."""
    caminho_win = os.path.normpath(ARQUIVO_LOJAS_BASE)
    if not os.path.exists(caminho_win):
        logging.error(f"Arquivo base de lojas não encontrado em: {caminho_win}")
        return set()

    try:
        df_lojas = pd.read_excel(caminho_win, usecols=[0], engine="openpyxl")
        col_nome = df_lojas.columns[0]
        lojas = set(df_lojas[col_nome].dropna().astype(str).str.strip().tolist())
        logging.info(f"Carregadas {len(lojas)} lojas de referência da base.")
        return lojas
    except Exception as e:
        logging.error(f"Erro ao ler a planilha base de lojas ({caminho_win}): {e}")
        return set()

def extrair_lojas_do_arquivo(caminho_arquivo: str, coluna_idx: int) -> Set[str]:
    """Extrai os códigos únicos de loja de um arquivo gerado."""
    try:
        df = pd.read_excel(caminho_arquivo, usecols=[coluna_idx], engine="openpyxl")
        col_nome = df.columns[0]
        valores = df[col_nome].dropna().astype(str).str.strip().tolist()
        
        # Filtra códigos numéricos válidos
        return {v for v in valores if v.isdigit()}
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo {caminho_arquivo} para auditoria: {e}")
        return set()

def gerenciar_arquivos_pendentes_de_reprocessamento() -> List[str]:
    """Verifica e remove arquivos marcados na execução anterior para reprocessamento."""
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
        
        os.remove(CAMINHO_JSON_REPROCESSAR)
        return removidos
    except Exception as e:
        logging.error(f"Erro ao processar arquivo de reprocessamento: {e}")
        return []

def auditar_arquivos_gerados(arquivos_gerados: List[dict]):
    """
    Audita APENAS os arquivos criados na execução atual.
    
    Estrutura esperada de 'arquivos_gerados':
    [
        {"caminho": ".../15-09-26.xlsx", "tipo": "vendas"},
        {"caminho": ".../15-09-26.xlsx", "tipo": "estoque"}
    ]
    """
    if not arquivos_gerados:
        logging.info("\n--- AUDITORIA: Nenhum arquivo novo foi gerado nesta execução. ---")
        return

    logging.info(f"\n--- INICIANDO AUDITORIA DE LOJAS ({len(arquivos_gerados)} arquivo(s) gerado(s) hoje) ---")
    
    lojas_referencia = carregar_codigos_lojas_referencia()
    if not lojas_referencia:
        logging.warning("Auditoria ignorada: Lista de lojas de referência vazia ou inacessível.")
        return

    arquivos_com_falha = []

    for item in arquivos_gerados:
        caminho_arq = os.path.normpath(item["caminho"])
        tipo = item["tipo"].lower()
        nome_arq = os.path.basename(caminho_arq)

        if not os.path.exists(caminho_arq):
            continue

        # Vendas = Coluna E (índice 4) | Estoque = Coluna B (índice 1)
        coluna_idx = 4 if tipo == "vendas" else 1
        
        lojas_presentes = extrair_lojas_do_arquivo(caminho_arq, coluna_idx=coluna_idx)
        lojas_faltantes = lojas_referencia - lojas_presentes

        if lojas_faltantes:
            msg = f"Inconsistência em {tipo.capitalize()} [{nome_arq}]: Faltam {len(lojas_faltantes)} lojas ({', '.join(sorted(lojas_faltantes))})"
            logging.warning(msg)
            arquivos_com_falha.append(caminho_arq)
        else:
            logging.info(f"Auditoria OK em {tipo.capitalize()} [{nome_arq}]: Todas as {len(lojas_referencia)} lojas presentes.")

    if arquivos_com_falha:
        conteudo_json = {
            "descricao": "Arquivos da execução atual marcados para exclusão e reprocessamento automático.",
            "arquivos_incompletos": arquivos_com_falha
        }
        with open(CAMINHO_JSON_REPROCESSAR, "w", encoding="utf-8") as f:
            json.dump(conteudo_json, f, indent=4, ensure_ascii=False)
        logging.warning(f"Total de {len(arquivos_com_falha)} arquivo(s) inconsistente(s) retido(s) para a próxima execução.")
    else:
        logging.info("Auditoria concluída: Todos os arquivos gerados hoje estão 100% completos.")