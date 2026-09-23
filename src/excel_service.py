import os
import win32com.client as win32

class ExcelManager:
    def __enter__(self):
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        self.excel.AskToUpdateLinks = False
        
        try:
            self.excel.Visible = False
        except Exception:
            pass
            
        return self

    def aguardar_consultas_assincronas(self):
        """Aposta na sincronia de consultas assíncronas do Excel."""
        try:
            self.excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

    def __exit__(self, exc_type, exc_val, exc_tb):
        if hasattr(self, 'excel'):
            try:
                self.excel.DisplayAlerts = True
                self.excel.ScreenUpdating = True
                self.excel.Quit()
            except Exception:
                pass
            finally:
                del self.excel