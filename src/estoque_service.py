import os
import logging
import pywintypes
from datetime import datetime
from tqdm import tqdm
from config import PASTA_FINAL_ESTOQUE, ARQUIVO_ESTOQUE_ORIGINAL
from date_utils import obter_datas_faltantes
from excel_service import ExcelManager

def processar_estoque():
    logging.info("\n--- INICIANDO PROCESSAMENTO DE ESTOQUE ---")
    datas_faltantes, arquivo_modelo, dt_alvo = obter_datas_faltantes(PASTA_FINAL_ESTOQUE)

    if not datas_faltantes:
        logging.info(f"Estoque: Todos os arquivos até D-1 ({dt_alvo.strftime('%d/%m/%Y')}) já foram criados!")
        return []

    logging.info(f"Pasta de destino: {PASTA_FINAL_ESTOQUE}")
    logging.info(f"Gerando {len(datas_faltantes)} arquivo(s) pendente(s) de Estoque...\n")

    modelo_atual = arquivo_modelo if arquivo_modelo and os.path.exists(arquivo_modelo) else ARQUIVO_ESTOQUE_ORIGINAL

    if not os.path.exists(modelo_atual):
        raise FileNotFoundError(f"Arquivo modelo inicial de estoque não encontrado em: {modelo_atual}")

    try:
        with ExcelManager() as excel_mgr:
            for dt_loop in tqdm(datas_faltantes, desc="Estoque", unit="arq"):
                string_mdx = f"[Data Estoque].[Dia].&[{dt_loop.strftime('%Y-%m-%d')}T00:00:00]"
                caminho_novo = os.path.join(PASTA_FINAL_ESTOQUE, f"{dt_loop.strftime('%d-%m-%y')}.xlsx")

                modelo_win = os.path.normpath(modelo_atual)
                destino_win = os.path.normpath(caminho_novo)

                wb = excel_mgr.excel.Workbooks.Open(Filename=modelo_win, UpdateLinks=0, ReadOnly=False)
                try:
                    ws = wb.Worksheets("Planilha1")
                    pt = ws.PivotTables("Estoque")
                    pf = pt.PivotFields("[Data Estoque].[Dia].[Dia]")

                    pt.ManualUpdate = True
                    pf.VisibleItemsList = [string_mdx]
                    pt.ManualUpdate = False

                    wb.SaveAs(Filename=destino_win, FileFormat=51) # 51 = xlOpenXMLWorkbook (.xlsx)
                    modelo_atual = caminho_novo
                finally:
                    wb.Close(SaveChanges=False)

        logging.info("Estoque concluído com sucesso!")
        return [os.path.join(PASTA_FINAL_ESTOQUE, f"{dt.strftime('%d-%m-%y')}.xlsx") for dt in datas_faltantes]

    except (pywintypes.com_error, ConnectionError) as e:
        msg_erro = f"Falha de conexão com o servidor/cubo ao processar Estoque: {e}"
        logging.error(msg_erro)
        raise ConnectionError(msg_erro) from e
    except Exception as e:
        logging.error(f"Erro inesperado no processamento de Estoque: {e}")
        raise e