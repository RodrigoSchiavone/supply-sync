import os
import sys
import logging
import traceback
from config import CAMINHO_LOG
from estoque_service import processar_estoque
from vendas_service import processar_vendas
from suprimentos_service import processar_suprimentos
from auditoria_service import gerenciar_arquivos_pendentes_de_reprocessamento, auditar_arquivos_gerados

DEBUG_MODE = os.getenv("DEBUG", "False").lower() in ("true", "1", "t")

# Configura o sistema de Logs para gravar simultaneamente no arquivo execucao.log e no Terminal
log_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG if DEBUG_MODE else logging.INFO)

# Handler 1: Arquivo execucao.log
file_handler = logging.FileHandler(CAMINHO_LOG, encoding="utf-8")
file_handler.setFormatter(log_formatter)
root_logger.addHandler(file_handler)

# Handler 2: Console / Terminal
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)
root_logger.addHandler(console_handler)

def run():
    logging.info("==================================================")
    logging.info("INICIANDO ROTINA INTEGRADA DE ABASTECIMENTO")
    logging.info("==================================================")

    # 0. Limpa arquivos pré-marcados com falhas de loja na execução anterior
    gerenciar_arquivos_pendentes_de_reprocessamento()

    # 1. Processamento da rotina de Estoque
    processar_estoque()
    
    # 2. Processamento da rotina de Vendas
    processar_vendas()

    # 3. Processamento da rotina de Suprimentos
    processar_suprimentos()

    # 4. Auditoria Final de Presença de Lojas
    auditar_arquivos_gerados()

    logging.info("==================================================")
    logging.info("TODAS AS ROTINAS FORAM CONCLUÍDAS COM SUCESSO")
    logging.info("==================================================")

if __name__ == "__main__":
    try:
        run()
        print("\nProcessamento total concluído! Pressione ENTER para fechar a janela...")
        input()
    except Exception as e:
        logging.error("OCORREU UM ERRO CRÍTICO DURANTE A EXECUÇÃO:")
        logging.error(traceback.format_exc())
        
        print("\n" + "=" * 60)
        input("Pressione ENTER para sair...")
        sys.exit(1)