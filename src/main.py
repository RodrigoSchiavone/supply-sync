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

    # 0. Limpa falhas da execução anterior, se houver
    gerenciar_arquivos_pendentes_de_reprocessamento()

    # Guarda o histórico dos arquivos que foram efetivamente gerados HOJE
    novos_arquivos = []

    # 1. Estoque
    criados_estoque = processar_estoque() or []
    for arq in criados_estoque:
        novos_arquivos.append({"caminho": arq, "tipo": "estoque"})
    
    # 2. Vendas
    criados_vendas = processar_vendas() or []
    for arq in criados_vendas:
        novos_arquivos.append({"caminho": arq, "tipo": "vendas"})

    # 3. Suprimentos
    processar_suprimentos()

    # 4. Auditoria (Valida SOMENTE os arquivos novos criados)
    auditar_arquivos_gerados(novos_arquivos)

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