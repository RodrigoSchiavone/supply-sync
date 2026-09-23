import os
import sys
from dotenv import load_dotenv

load_dotenv()

def obter_diretorio_base() -> str:
    """Retorna o diretório base (pasta do executável .exe ou do arquivo main.py)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DIR_BASE = obter_diretorio_base()

# Resoluções com fallback caso esteja vazio no .env
PASTA_FINAL = os.getenv("PASTA_FINAL")
if not PASTA_FINAL or not PASTA_FINAL.strip():
    PASTA_FINAL = os.path.join(DIR_BASE, "resultado")

ARQUIVO_ESTOQUE_ORIGINAL = os.getenv("ARQUIVO_ESTOQUE_ORIGINAL")
if not ARQUIVO_ESTOQUE_ORIGINAL or not ARQUIVO_ESTOQUE_ORIGINAL.strip():
    ARQUIVO_ESTOQUE_ORIGINAL = os.path.join(DIR_BASE, "estoque.xlsx")