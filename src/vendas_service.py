import os
import logging
from datetime import datetime
from tqdm import tqdm
from config import PASTA_FINAL_VENDAS, ARQUIVO_VENDAS_ORIGINAL
from date_utils import obter_datas_faltantes
from excel_service import ExcelManager

def processar_vendas():
    logging.info("\n--- INICIANDO PROCESSAMENTO DE VENDAS ---")
    data_base_inicial = datetime(2025, 1, 1).date()
    datas_faltantes, _, dt_alvo = obter_datas_faltantes(PASTA_FINAL_VENDAS, data_inicio_padrao=data_base_inicial)

    if not datas_faltantes:
        logging.info(f"Vendas: Todos os arquivos até D-1 ({dt_alvo.strftime('%d/%m/%Y')}) já foram criados!")
        return

    logging.info(f"Pasta de destino: {PASTA_FINAL_VENDAS}")
    logging.info(f"Gerando {len(datas_faltantes)} arquivo(s) pendente(s) de Vendas...\n")

    modelo_path_win = os.path.normpath(ARQUIVO_VENDAS_ORIGINAL)
    if not os.path.exists(modelo_path_win):
        raise FileNotFoundError(f"Arquivo modelo original de vendas não encontrado em: {modelo_path_win}")

    with ExcelManager() as excel_mgr:
        wb_vendas = excel_mgr.excel.Workbooks.Open(Filename=modelo_path_win, UpdateLinks=0, ReadOnly=False)
        
        try:
            sheet = wb_vendas.ActiveSheet
            pt = sheet.PivotTables("Tabela dinâmica1")
            pf = pt.PivotFields("[Tempo].[Período].[Ano]")

            # Garante atualização síncrona da consulta OLAP
            try:
                pt.PivotCache.BackgroundQuery = False
            except Exception as e:
                logging.debug(f"Não foi possível ajustar BackgroundQuery: {e}")

            for dt_loop in tqdm(datas_faltantes, desc="Vendas", unit="arq"):
                # Monta a string MDX do nível de período de tempo
                string_mdx = (
                    f"[Tempo].[Período].[Ano].&[{dt_loop.strftime('%Y')}]."
                    f"&[{dt_loop.strftime('%m/%Y')}]."
                    f"&[{dt_loop.strftime('%Y-%m-%d')}T00:00:00]"
                )
                
                caminho_novo = os.path.join(PASTA_FINAL_VENDAS, f"{dt_loop.strftime('%d-%m-%Y')}.xlsx")
                caminho_novo_win = os.path.normpath(caminho_novo)

                # Aplicação do filtro no cubo de vendas
                try:
                    pt.CubeFields("[Tempo].[Período]").EnableMultiplePageItems = False
                    pf.VisibleItemsList = [string_mdx]
                except Exception:
                    # Fallback caso responda apenas por CurrentPageName
                    pf.CurrentPageName = string_mdx

                pt.Update()
                excel_mgr.aguardar_consultas_assincronas()

                # Salva uma cópia do relatório do dia
                wb_vendas.SaveCopyAs(Filename=caminho_novo_win)

        finally:
            wb_vendas.Close(SaveChanges=False)

    logging.info("Vendas concluído com sucesso!")
