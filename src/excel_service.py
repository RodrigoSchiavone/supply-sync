import os
import logging
import pywintypes
import win32com.client as win32

class ExcelManager:
    def __enter__(self):
        self.excel = win32.Dispatch("Excel.Application")
        
        # Suprime pop-ups de erro e solicitações de atualização do Excel
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        self.excel.AskToUpdateLinks = False
        
        try:
            self.excel.Visible = False
        except Exception:
            pass
            
        return self

    def aguardar_consultas_assincronas(self):
        """Aposta na sincronia de consultas assíncronas do Excel e trata eventuais erros de conexão."""
        try:
            self.excel.CalculateUntilAsyncQueriesDone()
        except pywintypes.com_error as e:
            logging.error(f"Erro COM/Conexão detectado ao aguardar atualização do Excel: {e}")
            raise ConnectionError("O servidor ou cubo Analysis Services recusou/perdeu a conexão.") from e
        except Exception as e:
            logging.warning(f"Aviso genérico ao aguardar consultas assíncronas: {e}")

    def __exit__(self, exc_type, exc_val, exc_tb):
        if hasattr(self, 'excel') and self.excel:
            try:
                self.excel.DisplayAlerts = True
                self.excel.ScreenUpdating = True
                self.excel.Quit()
            except Exception as e:
                logging.debug(f"Erro ao fechar aplicação Excel no exit: {e}")
            finally:
                del self.excel