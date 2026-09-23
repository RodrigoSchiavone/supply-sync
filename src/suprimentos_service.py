import os
import logging
import pywintypes
from config import ARQUIVO_SUPRIMENTOS
from excel_service import ExcelManager

def processar_suprimentos():
    logging.info("\n--- INICIANDO ATUALIZAÇÃO DE SUPRIMENTOS ---")
    
    suprimentos_path_win = os.path.normpath(ARQUIVO_SUPRIMENTOS)
    if not os.path.exists(suprimentos_path_win):
        raise FileNotFoundError(f"Arquivo de Suprimentos não encontrado em: {suprimentos_path_win}")

    logging.info(f"Atualizando arquivo: {suprimentos_path_win}")

    try:
        with ExcelManager() as excel_mgr:
            wb = excel_mgr.excel.Workbooks.Open(Filename=suprimentos_path_win, UpdateLinks=0, ReadOnly=False)
            try:
                # Procura a tabela dinâmica 'suprimentos' nas abas da planilha
                pt_encontrada = False
                for sheet in wb.Worksheets:
                    try:
                        pt = sheet.PivotTables("suprimentos")
                        
                        # Garante atualização síncrona
                        try:
                            pt.PivotCache.BackgroundQuery = False
                        except Exception:
                            pass

                        pt.PivotCache.Refresh()
                        excel_mgr.aguardar_consultas_assincronas()
                        pt_encontrada = True
                        break
                    except Exception:
                        continue

                if not pt_encontrada:
                    # Fallback: Se não achar pelo nome exato, tenta dar RefreshAll na pasta de trabalho
                    logging.warning("Tabela 'suprimentos' não encontrada pelo nome específico. Executando RefreshAll...")
                    wb.RefreshAll()
                    excel_mgr.aguardar_consultas_assincronas()

                wb.Save()
                logging.info("Suprimentos atualizado e salvo com sucesso!")

            finally:
                wb.Close(SaveChanges=False)

    except (pywintypes.com_error, ConnectionError) as e:
        msg_erro = f"Falha de conexão com o servidor/cubo ao atualizar Suprimentos: {e}"
        logging.error(msg_erro)
        raise ConnectionError(msg_erro) from e
    except Exception as e:
        logging.error(f"Erro inesperado no processamento de Suprimentos: {e}")
        raise e