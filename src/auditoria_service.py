import os
import json
import logging
import requests
import pandas as pd
from datetime import datetime
from typing import List, Set, Dict
from config import ARQUIVO_LOJAS_BASE, CAMINHO_JSON_REPROCESSAR

# Arquivo para registrar o estado das notificações do dia
CAMINHO_JSON_STATUS_ALERTAS = os.path.join(os.path.dirname(CAMINHO_JSON_REPROCESSAR), "status_alertas.json")


def enviar_notificacao_whatsapp(mensagem: str):
    """Envia uma mensagem via HTTP para a lista de números do WhatsApp cadastrados no .env."""
    phones_str = os.getenv("WHATSAPP_PHONES", "")
    keys_str = os.getenv("WHATSAPP_API_KEYS", "")

    if not phones_str or not keys_str:
        logging.warning("WhatsApp: Mensagem não enviada. Configure WHATSAPP_PHONES e WHATSAPP_API_KEYS no .env")
        return

    telefones = [p.strip() for p in phones_str.split(",") if p.strip()]
    chaves = [k.strip() for k in keys_str.split(",") if k.strip()]

    if len(telefones) != len(chaves):
        logging.warning("WhatsApp: A quantidade de números e chaves de API no .env é diferente.")

    # Itera sobre cada par (telefone, chave) para realizar o envio
    for phone, key in zip(telefones, chaves):
        if not key or key == "SUA_API_KEY_AQUI":
            logging.warning(f"WhatsApp: Chave inválida para o número {phone}. Pulando...")
            continue

        try:
            url = f"https://api.callmebot.com/whatsapp.php?phone={phone}&text={requests.utils.quote(mensagem)}&apikey={key}"
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                logging.info(f"WhatsApp: Notificação enviada com sucesso para {phone}!")
            else:
                logging.error(f"WhatsApp: Falha ao enviar notificação para {phone} (HTTP {response.status_code})")
        except Exception as e:
            logging.error(f"WhatsApp: Erro ao conectar com o serviço para {phone}: {e}")


def carregar_lojas_referencia() -> Dict[str, str]:
    """
    Lê a planilha base de lojas.
    Retorna um dicionário com mapeamento {Código: Nome da Loja}.
    Assume Coluna A = Código e Coluna B = Nome da Loja.
    """
    caminho_win = os.path.normpath(ARQUIVO_LOJAS_BASE)
    if not os.path.exists(caminho_win):
        logging.error(f"Arquivo base de lojas não encontrado em: {caminho_win}")
        return {}

    try:
        df_lojas = pd.read_excel(caminho_win, usecols=[0, 1], engine="openpyxl")
        df_lojas.dropna(subset=[df_lojas.columns[0]], inplace=True)
        
        # Cria o dicionário {Código: Nome}
        mapeamento = {}
        for _, row in df_lojas.iterrows():
            cod = str(row.iloc[0]).strip()
            nome = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else "Nome não especificado"
            if cod.isdigit():
                mapeamento[cod] = nome
                
        logging.info(f"Carregadas {len(mapeamento)} lojas de referência da base.")
        return mapeamento
    except Exception as e:
        logging.error(f"Erro ao ler a planilha base de lojas ({caminho_win}): {e}")
        return {}


def extrair_lojas_do_arquivo(caminho_arquivo: str, coluna_idx: int) -> Set[str]:
    """Extrai os códigos únicos de loja de um arquivo gerado."""
    try:
        df = pd.read_excel(caminho_arquivo, usecols=[coluna_idx], engine="openpyxl")
        col_nome = df.columns[0]
        valores = df[col_nome].dropna().astype(str).str.strip().tolist()
        return {v for v in valores if v.isdigit()}
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo {caminho_arquivo} para auditoria: {e}")
        return set()


def gerenciar_arquivos_pendentes_de_reprocessamento() -> List[str]:
    """Verifica e remove arquivos marcados na execução anterior para reprocessamento."""
    if not os.path.exists(CAMINHO_JSON_REPROCESSAR):
        return []

    try:
        with open(CAMINHO_JSON_REPROCESSAR, "r", encoding="utf-8") as f:
            dados = json.load(f)
        
        arquivos_para_remover = dados.get("arquivos_incompletos", [])
        removidos = []

        for caminho in arquivos_para_remover:
            caminho_win = os.path.normpath(caminho)
            if os.path.exists(caminho_win):
                try:
                    os.remove(caminho_win)
                    removidos.append(caminho_win)
                    logging.info(f"Arquivo inconsistente removido para reprocessamento: {caminho_win}")
                except Exception as e:
                    logging.error(f"Não foi possível remover o arquivo {caminho_win}: {e}")
        
        os.remove(CAMINHO_JSON_REPROCESSAR)
        return removidos
    except Exception as e:
        logging.error(f"Erro ao processar arquivo de reprocessamento: {e}")
        return []


def auditar_arquivos_gerados(arquivos_gerados: List[dict]):
    """
    Audita APENAS os arquivos criados na execução atual e gerencia
    alertas inteligentes no WhatsApp apenas para alterações de estado.
    """
    if not arquivos_gerados:
        logging.info("\n--- AUDITORIA: Nenhum arquivo novo foi gerado nesta execução. ---")
        return

    logging.info(f"\n--- INICIANDO AUDITORIA DE LOJAS ({len(arquivos_gerados)} arquivo(s) gerado(s) hoje) ---")
    
    lojas_ref_map = carregar_lojas_referencia()
    if not lojas_ref_map:
        logging.warning("Auditoria ignorada: Lista de lojas de referência vazia ou inacessível.")
        return

    todas_lojas_codigos = set(lojas_ref_map.keys())
    arquivos_com_falha = []

    # Carrega histórico prévio de alertas do dia
    hoje_str = datetime.now().strftime("%Y-%m-%d")
    estado_alertas = {}
    if os.path.exists(CAMINHO_JSON_STATUS_ALERTAS):
        try:
            with open(CAMINHO_JSON_STATUS_ALERTAS, "r", encoding="utf-8") as f:
                estado_alertas = json.load(f)
        except Exception:
            estado_alertas = {}

    # Se mudou o dia, reseta o histórico de estado
    if estado_alertas.get("data") != hoje_str:
        estado_alertas = {"data": hoje_str, "falhas": {}}

    falhas_anteriores = estado_alertas.get("falhas", {})
    falhas_atuais = {}

    for item in arquivos_gerados:
        caminho_arq = os.path.normpath(item["caminho"])
        tipo = item["tipo"].lower()  # 'vendas' ou 'estoque'
        nome_arq = os.path.basename(caminho_arq)

        if not os.path.exists(caminho_arq):
            continue

        coluna_idx = 4 if tipo == "vendas" else 1
        lojas_presentes = extrair_lojas_do_arquivo(caminho_arq, coluna_idx=coluna_idx)
        lojas_faltantes = todas_lojas_codigos - lojas_presentes

        if lojas_faltantes:
            arquivos_com_falha.append(caminho_arq)
            falhas_atuais[tipo] = {
                "arquivo": nome_arq,
                "lojas": sorted(list(lojas_faltantes))
            }
            logging.warning(f"Inconsistência em {tipo.capitalize()} [{nome_arq}]: Faltam {len(lojas_faltantes)} loja(s).")
        else:
            logging.info(f"Auditoria OK em {tipo.capitalize()} [{nome_arq}]: Todas as lojas presentes.")

    # --- LÓGICA DE COMPARAÇÃO DE ESTADO E DISPARO DE MENSAGENS ---
    for tipo in ["vendas", "estoque"]:
        falha_ant = falhas_anteriores.get(tipo, {})
        falha_at = falhas_atuais.get(tipo, {})

        lojas_falha_ant = set(falha_ant.get("lojas", []))
        lojas_falha_at = set(falha_at.get("lojas", []))

        # 1. Novas lojas que passaram a falhar
        novas_falhas = lojas_falha_at - lojas_falha_ant
        if novas_falhas:
            detalhes = [f"- Cod {c}: {lojas_ref_map.get(c, 'Desconhecida')}" for c in sorted(novas_falhas)]
            msg = (
                f"⚠️ *ALERTA DE CARGA - {tipo.upper()}*\n"
                f"Arquivo: {falha_at.get('arquivo')}\n\n"
                f"Identificado que o cubo de *{tipo.upper()}* não carregou as seguintes lojas:\n"
                + "\n".join(detalhes)
            )
            enviar_notificacao_whatsapp(msg)

        # 2. Lojas que foram REGULARIZADAS na rodada atual
        regularizadas = lojas_falha_ant - lojas_falha_at
        if regularizadas:
            detalhes_reg = [f"- Cod {c}: {lojas_ref_map.get(c, 'Desconhecida')}" for c in sorted(regularizadas)]
            
            msg_reg = (
                f"✅ *REGULARIZAÇÃO DE CARGA - {tipo.upper()}*\n\n"
                f"O cubo de *{tipo.upper()}* foi atualizado e regularizado para as lojas:\n"
                + "\n".join(detalhes_reg)
            )
            
            # Se ainda sobrarem outras lojas falhando no mesmo cubo, inclui o alerta residual
            if lojas_falha_at:
                detalhes_residuais = [f"- Cod {c}: {lojas_ref_map.get(c, 'Desconhecida')}" for c in sorted(lojas_falha_at)]
                msg_reg += (
                    f"\n\n⚠️ *Atenção:* Permanece com divergência para as lojas:\n"
                    + "\n".join(detalhes_residuais)
                )
            
            enviar_notificacao_whatsapp(msg_reg)

    # Grava o novo estado atualizado no JSON local
    estado_alertas["falhas"] = falhas_atuais
    with open(CAMINHO_JSON_STATUS_ALERTAS, "w", encoding="utf-8") as f:
        json.dump(estado_alertas, f, indent=4, ensure_ascii=False)

    # Grava arquivo para reprocessamento de disco se houverem falhas
    if arquivos_com_falha:
        conteudo_json = {
            "descricao": "Arquivos marcados para exclusão e reprocessamento automático.",
            "arquivos_incompletos": arquivos_com_falha
        }
        with open(CAMINHO_JSON_REPROCESSAR, "w", encoding="utf-8") as f:
            json.dump(conteudo_json, f, indent=4, ensure_ascii=False)