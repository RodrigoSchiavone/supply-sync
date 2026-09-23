import os
import win32com.client as win32

class ExcelManager:
    def __enter__(self):
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        self.excel.AskToUpdateLinks = False
        return self

    def processar_dia(self, modelo_path: str, caminho_destino: str, string_mdx: str):
        # Normaliza os caminhos para o formato nativo do Windows (\) exigido pela API COM do Excel
        modelo_path_win = os.path.normpath(modelo_path)
        caminho_destino_win = os.path.normpath(caminho_destino)

        wb = self.excel.Workbooks.Open(Filename=modelo_path_win, UpdateLinks=0, ReadOnly=False)
        try:
            ws = wb.Worksheets("Planilha1")
            pt = ws.PivotTables("Estoque")
            pf = pt.PivotFields("[Data Estoque].[Dia].[Dia]")

            pt.ManualUpdate = True
            pf.VisibleItemsList = [string_mdx]
            pt.ManualUpdate = False

            wb.SaveAs(Filename=caminho_destino_win, FileFormat=51) # 51 = xlOpenXMLWorkbook (.xlsx)
        finally:
            wb.Close(SaveChanges=False)

    def __exit__(self, exc_type, exc_val, exc_tb):
        if hasattr(self, 'excel'):
            self.excel.ScreenUpdating = True
            self.excel.DisplayAlerts = True
            self.excel.Quit()