import os
import logging
from config import ARQUIVO_SUPRIMENTOS
from excel_service import ExcelManager

def processar_suprimentos():
    print("\n--- INICIANDO ATUALIZAÇÃO DE SUPRIMENTOS ---")
    
    suprimentos_path_win = os.path.normpath(ARQUIVO_SUPRIMENTOS)
    if not os.path.exists(suprimentos_path_win):
        raise FileNotFoundError(f"Arquivo de Suprimentos não encontrado em: {suprimentos_path_win}")

    print(f"Atualizando arquivo: {suprimentos_path_win}")

    with ExcelManager() as excel_mgr:
        wb = excel_mgr.excel.Workbooks.Open(Filename=suprimentos_path_win, UpdateLinks=0, ReadOnly=False)
        try:
            # Procura a tabela dinâmica 'Tabela dinâmica3' nas abas da planilha
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
                # Fallback: Se não achar pelo nome exato em abas específicas, tenta dar RefreshAll na pasta de trabalho
                logging.warning("Tabela suprimentos não encontrada pelo nome específico. Executando RefreshAll...")
                wb.RefreshAll()
                excel_mgr.aguardar_consultas_assincronas()

            wb.Save()
        finally:
            wb.Close(SaveChanges=False)

    print("Suprimentos atualizado com sucesso!")
