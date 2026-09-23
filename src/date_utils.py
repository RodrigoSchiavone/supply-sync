import os
import re
from datetime import datetime, timedelta
from typing import List, Tuple, Optional

def converter_nome_para_data(nome_arquivo: str) -> Optional[datetime.date]:
    """Extrai a data do nome do arquivo no formato dd-mm-yy.xlsx."""
    match = re.search(r"^(\d{2})-(\d{2})-(\d{2})\.xlsx$", nome_arquivo, re.IGNORECASE)
    if match:
        dia, mes, ano = map(int, match.groups())
        try:
            return datetime(2000 + ano, mes, dia).date()
        except ValueError:
            return None
    return None

def obter_datas_faltantes(pasta_final: str) -> Tuple[List[datetime.date], Optional[str], datetime.date]:
    """Mapeia os arquivos na pasta de destino e retorna a lista de datas pendentes até D-1."""
    if not os.path.exists(pasta_final):
        os.makedirs(pasta_final)

    dt_max_existente = None
    arquivo_modelo = None

    for f in os.listdir(pasta_final):
        if f.endswith(".xlsx") and not f.startswith("~$"):
            dt_arq = converter_nome_para_data(f)
            if dt_arq:
                if dt_max_existente is None or dt_arq > dt_max_existente:
                    dt_max_existente = dt_arq
                    arquivo_modelo = os.path.join(pasta_final, f)

    dt_alvo = datetime.now().date() - timedelta(days=1)
    
    if dt_max_existente is None:
        dt_inicio = datetime.now().date() - timedelta(days=60)
    else:
        dt_inicio = dt_max_existente + timedelta(days=1)

    datas_faltantes = []
    curr = dt_inicio
    while curr <= dt_alvo:
        datas_faltantes.append(curr)
        curr += timedelta(days=1)

    return datas_faltantes, arquivo_modelo, dt_alvo
    