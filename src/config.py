import os
import sys
from dotenv import load_dotenv

load_dotenv()

def obter_diretorio_base() -> str:
    """Retorna o diretório base do executável ou do arquivo fonte."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DIR_BASE = obter_diretorio_base()

def limpar_caminho(val: str) -> str:
    """Trata e limpa aspas ou prefixos indesejados lidos do .env."""
    if not val:
        return ""
    val = val.strip()
    if val.startswith(('r"', "r'")):
        val = val[1:]
    return val.strip('"' + "'")

# --- CAMINHOS DE TRABALHO ---
PASTA_FINAL_ESTOQUE = limpar_caminho(os.getenv("PASTA_FINAL_ESTOQUE")) or os.path.join(DIR_BASE, "resultado")
ARQUIVO_ESTOQUE_ORIGINAL = limpar_caminho(os.getenv("ARQUIVO_ESTOQUE_ORIGINAL")) or os.path.join(DIR_BASE, "estoque.xlsx")

PASTA_FINAL_VENDAS = limpar_caminho(os.getenv("PASTA_FINAL_VENDAS")) or os.path.join(DIR_BASE, "Vendas")
ARQUIVO_VENDAS_ORIGINAL = limpar_caminho(os.getenv("ARQUIVO_VENDAS_ORIGINAL")) or os.path.join(DIR_BASE, "vendas.xlsx")

ARQUIVO_SUPRIMENTOS = limpar_caminho(os.getenv("ARQUIVO_SUPRIMENTOS")) or os.path.join(DIR_BASE, "Compras.xlsx")
ARQUIVO_LOJAS_BASE = limpar_caminho(os.getenv("ARQUIVO_LOJAS_BASE")) or os.path.join(DIR_BASE, "Lojas.xlsx")

# --- ARQUIVOS LOCAIS DE LOG E AUXILIAR ---
CAMINHO_LOG = os.path.join(DIR_BASE, "execucao.log")
CAMINHO_JSON_REPROCESSAR = os.path.join(DIR_BASE, "reprocessar_arquivos.json")