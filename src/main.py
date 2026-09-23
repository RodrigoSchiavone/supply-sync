import os
from tqdm import tqdm
from config import PASTA_FINAL, ARQUIVO_ESTOQUE_ORIGINAL
from date_utils import obter_datas_faltantes
from excel_service import ExcelManager

def run():
    datas_faltantes, arquivo_modelo, dt_alvo = obter_datas_faltantes(PASTA_FINAL)

    if not datas_faltantes:
        print(f"Todos os arquivos até D-1 ({dt_alvo.strftime('%d/%m/%Y')}) já foram criados!")
        return

    print(f"Pasta de destino: {PASTA_FINAL}")
    print(f"Gerando {len(datas_faltantes)} arquivo(s) pendente(s)...\n")

    modelo_atual = arquivo_modelo if arquivo_modelo and os.path.exists(arquivo_modelo) else ARQUIVO_ESTOQUE_ORIGINAL

    if not os.path.exists(modelo_atual):
        raise FileNotFoundError(f"Arquivo modelo inicial não encontrado em: {modelo_atual}")

    with ExcelManager() as excel_mgr:
        for dt_loop in tqdm(datas_faltantes, desc="Processando datas", unit="arq"):
            string_mdx = f"[Data Estoque].[Dia].&[{dt_loop.strftime('%Y-%m-%d')}T00:00:00]"
            caminho_novo = os.path.join(PASTA_FINAL, f"{dt_loop.strftime('%d-%m-%y')}.xlsx")

            excel_mgr.processar_dia(
                modelo_path=modelo_atual,
                caminho_destino=caminho_novo,
                string_mdx=string_mdx
            )
            modelo_atual = caminho_novo

    print("\nProcesso concluído com sucesso!")

if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"\nErro durante a execução: {e}")