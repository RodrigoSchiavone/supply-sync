import os
import re
from datetime import datetime, timedelta
from typing import List, Tuple, Optional

def converter_nome_para_data(nome_arquivo: str) -> Optional[datetime.date]:
    """Converte nome de arquivo no formato dd-mm-yy.xlsx ou dd-mm-yyyy.xlsx para objeto date."""
    match = re.search(r"^(\d{2})-(\d{2})-(\d{2,4})\.xlsx$", nome_arquivo, re.IGNORECASE)
    if match:
        dia, mes, ano_raw = map(int, match.groups())
        ano = 2000 + ano_raw if ano_raw < 100 else ano_raw
        try:
            return datetime(ano, mes, dia).date()
        except ValueError:
            return None
    return None

def obter_datas_faltantes(
    pasta_final: str, 
    data_inicio_padrao: Optional[datetime.date] = None
) -> Tuple[List[datetime.date], Optional[str], datetime.date]:
    """Mapeia os arquivos existentes para calcular o intervalo de datas pendentes até D-1."""
    if not os.path.exists(pasta_final):
        os.makedirs(pasta_final, exist_ok=True)

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
        dt_inicio = data_inicio_padrao if data_inicio_padrao else (datetime.now().date() - timedelta(days=60))
    else:
        dt_inicio = dt_max_existente + timedelta(days=1)

    datas_faltantes = []
    curr = dt_inicio
    while curr <= dt_alvo:
        datas_faltantes.append(curr)
        curr += timedelta(days=1)

    return datas_faltantes, arquivo_modelo, dt_alvo
    