import os
import sys
import logging
import traceback
from config import PASTA_FINAL_ESTOQUE, PASTA_FINAL_VENDAS
from estoque_service import processar_estoque
from vendas_service import processar_vendas

DEBUG_MODE = os.getenv("DEBUG", "False").lower() in ("true", "1", "t")

logging.basicConfig(
    level=logging.DEBUG if DEBUG_MODE else logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)

def run():
    # 1. Processamento da rotina de Estoque
    processar_estoque()
    
    # 2. Processamento da rotina de Vendas
    processar_vendas()

if __name__ == "__main__":
    try:
        run()
        print("\nProcessamento total concluído! Pressione ENTER para fechar a janela...")
        input()
    except Exception as e:
        print("\n" + "=" * 60)
        print(" OCORREU UM ERRO DURANTE A EXECUÇÃO:")
        print("=" * 60 + "\n")
        
        traceback.print_exc()
        
        print("\n" + "=" * 60)
        input("Pressione ENTER para sair...")
        sys.exit(1)