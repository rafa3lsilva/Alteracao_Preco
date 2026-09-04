import os
import json
from datetime import datetime
import streamlit as st

ARQUIVO_HISTORICO = ".historico_alteracoes.json"

def obter_cliente_gsheets():
    """Tenta obter o cliente gsheets através de st.secrets ou arquivo local."""
    try:
        if hasattr(st, "secrets") and "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            from google.oauth2.service_account import Credentials
            import gspread
            
            cfg = st.secrets["connections"]["gsheets"]
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            creds_info = {
                "type": cfg.get("type", "service_account"),
                "project_id": cfg.get("project_id"),
                "private_key_id": cfg.get("private_key_id"),
                "private_key": cfg.get("private_key"),
                "client_email": cfg.get("client_email"),
                "client_id": cfg.get("client_id"),
                "auth_uri": cfg.get("auth_uri", "https://accounts.google.com/o/oauth2/auth"),
                "token_uri": cfg.get("token_uri", "https://oauth2.googleapis.com/token"),
                "auth_provider_x509_cert_url": cfg.get("auth_provider_x509_cert_url"),
                "client_x509_cert_url": cfg.get("client_x509_cert_url")
            }
            creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
            client = gspread.authorize(creds)
            spreadsheet_url = cfg.get("spreadsheet")
            sheet = client.open_by_url(spreadsheet_url)
            return sheet.sheet1
    except Exception as e:
        # Silenciosamente falha para modo local
        pass
    return None


def carregar_historico():
    """Carrega o histórico do Google Sheets (se configurado) ou do arquivo local."""
    ws = obter_cliente_gsheets()
    if ws:
        try:
            records = ws.get_all_records()
            historico = {}
            for row in records:
                d = str(row.get("data", "")).strip()
                k = str(row.get("chave_item", "")).strip()
                t = str(row.get("tipo", "")).strip()
                h = str(row.get("horario", "")).strip()
                
                if not d:
                    continue
                if d not in historico:
                    historico[d] = {"itens": [], "execucoes": []}
                
                if k and k not in historico[d]["itens"]:
                    historico[d]["itens"].append(k)
                    
            return historico
        except Exception:
            pass

    # Fallback local
    if os.path.exists(ARQUIVO_HISTORICO):
        try:
            with open(ARQUIVO_HISTORICO, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def salvar_historico(data_str, itens_processados, info_execucao=""):
    """Salva os itens no Google Sheets (se ativo) e no arquivo local."""
    ws = obter_cliente_gsheets()
    hora_agora = datetime.now().strftime("%H:%M:%S")

    # 1. Salvar no Google Sheets
    if ws:
        try:
            rows_to_append = []
            for it in itens_processados:
                rows_to_append.append([
                    data_str,
                    hora_agora,
                    info_execucao,
                    it.get('key', ''),
                    it.get('cod', ''),
                    it.get('descricao', ''),
                    it.get('hora', ''),
                    it.get('venda_atual', '')
                ])
            if rows_to_append:
                ws.append_rows(rows_to_append)
        except Exception as e:
            print(f"Erro ao gravar no Google Sheets: {e}")

    # 2. Salvar também localmente como backup
    historico = {}
    if os.path.exists(ARQUIVO_HISTORICO):
        try:
            with open(ARQUIVO_HISTORICO, 'r', encoding='utf-8') as f:
                historico = json.load(f)
        except Exception:
            historico = {}

    if data_str not in historico:
        historico[data_str] = {'itens': [], 'execucoes': []}
    
    if isinstance(historico[data_str], list):
        historico[data_str] = {'itens': historico[data_str], 'execucoes': []}

    if info_execucao == "primeiro":
        historico[data_str]['itens'] = [it['key'] for it in itens_processados]
    else:
        conjunto = set(historico[data_str]['itens'])
        for it in itens_processados:
            conjunto.add(it['key'])
        historico[data_str]['itens'] = list(conjunto)

    historico[data_str]['execucoes'].append({
        'horario': hora_agora,
        'qtd_itens': len(itens_processados),
        'tipo': info_execucao
    })

    try:
        with open(ARQUIVO_HISTORICO, 'w', encoding='utf-8') as f:
            json.dump(historico, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Aviso ao salvar histórico local: {e}")


def resetar_historico(data_especifica=None):
    """Reseta o histórico do dia no Google Sheets e localmente."""
    ws = obter_cliente_gsheets()
    hoje = data_especifica or datetime.now().strftime("%d/%m/%Y")
    
    # 1. Resetar no Google Sheets
    if ws:
        try:
            records = ws.get_all_records()
            # Mantém apenas as linhas que NÃO são de hoje
            cabecalho = ["data", "horario", "tipo", "chave_item", "codigo", "descricao", "hora", "preco"]
            novas_linhas = [cabecalho]
            for row in records:
                if str(row.get("data", "")).strip() != hoje:
                    novas_linhas.append([
                        row.get("data", ""),
                        row.get("horario", ""),
                        row.get("tipo", ""),
                        row.get("chave_item", ""),
                        row.get("codigo", ""),
                        row.get("descricao", ""),
                        row.get("hora", ""),
                        row.get("preco", "")
                    ])
            ws.clear()
            ws.update("A1", novas_linhas)
        except Exception as e:
            print(f"Erro ao resetar no Google Sheets: {e}")

    # 2. Resetar localmente
    if os.path.exists(ARQUIVO_HISTORICO):
        try:
            with open(ARQUIVO_HISTORICO, 'r', encoding='utf-8') as f:
                hist = json.load(f)
            if hoje in hist:
                del hist[hoje]
            with open(ARQUIVO_HISTORICO, 'w', encoding='utf-8') as f:
                json.dump(hist, f, ensure_ascii=False, indent=2)
        except Exception:
            try:
                os.remove(ARQUIVO_HISTORICO)
            except Exception:
                pass
