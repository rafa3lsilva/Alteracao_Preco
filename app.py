import io
import os
import json
from datetime import datetime
import pandas as pd
import streamlit as st

from gerar_relatorio_pdf import (
    extrair_dados_xlsx,
    verificar_conformidade_planilha,
    filtrar_dados,
    gerar_pdf,
    salvar_historico,
    carregar_historico,
    resetar_historico,
    ARQUIVO_HISTORICO,
    PALAVRAS_CHAVE_HORTIFRUTI
)

st.set_page_config(
    page_title="Gerador de Alteração de Preços",
    page_icon="🏷️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------- SIDEBAR -----------------
with st.sidebar:
    st.title("🏷️ Memória do Dia")
    
    historico = carregar_historico()
    hoje = datetime.now().strftime("%d/%m/%Y")
    info_hoje = historico.get(hoje, {})
    
    itens_hist = info_hoje.get('itens', []) if isinstance(info_hoje, dict) else (info_hoje if isinstance(info_hoje, list) else [])
    execucoes = info_hoje.get('execucoes', []) if isinstance(info_hoje, dict) else []
    
    if itens_hist:
        st.success(f"📌 **{len(itens_hist)} itens** registrados hoje ({hoje}).")
        if execucoes:
            st.caption("Últimas execuções:")
            for exc in execucoes[-3:]:
                tipo = exc.get('tipo', 'geral').upper()
                horario = exc.get('horario', '--')
                qtd = exc.get('qtd_itens', 0)
                st.markdown(f"- `{horario}` • **{tipo}** ({qtd} itens)")
    else:
        st.info(f"Nenhum histórico gravado para hoje ({hoje}).")

    st.markdown("---")
    if st.button("🗑️ Resetar Memória de Hoje", use_container_width=True, type="secondary"):
        resetar_historico()
        st.toast("Histórico resetado com sucesso!", icon="✅")
        st.rerun()

    st.markdown("---")
    with st.expander("📖 Como funciona a rotina", expanded=False):
        st.markdown("""
        - **1º Relatório (Manhã ~12h/13h):** Processa todos os itens e salva a base.
        - **2º Relatório (Tarde ~17h/18h):** Compara com a manhã e gera **apenas as novidades**.
        - **Consolidado:** Gera tudo do dia sem filtros.
        """)

# ----------------- MAIN CONTENT -----------------
st.title("🏷️ Relatório de Alteração de Preços")
st.caption("Gere relatórios em PDF formatados para conferência de gôndola, com controle automático de novidades (manhã/tarde).")

uploaded_file = st.file_uploader(
    "Carregue a planilha exportada do Varejofácil (.xlsx):",
    type=["xlsx"],
    help="Selecione o arquivo Excel exportado do Varejofácil (Listagem de Preços - Analítico)"
)

if uploaded_file is not None:
    # 1. Checagem de Conformidade
    file_bytes = uploaded_file.read()
    file_buffer = io.BytesIO(file_bytes)
    
    alertas = verificar_conformidade_planilha(file_buffer)
    if alertas:
        for al in alertas:
            st.warning(al)
    
    # 2. Extração dos Dados
    file_buffer.seek(0)
    try:
        dados = extrair_dados_xlsx(file_buffer, ignorar_secoes=PALAVRAS_CHAVE_HORTIFRUTI)
    except Exception as e:
        st.error(f"Erro ao processar a planilha: {e}")
        st.stop()

    total_secoes = len(dados['secoes'])
    total_itens_brutos = sum(len(itens) for itens in dados['secoes'].values())

    # Métricas nativas do Streamlit (perfeitas tanto no tema claro quanto escuro)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(label="Empresa", value=dados["empresa"][:20])
    m2.metric(label="Data do Arquivo", value=dados["periodo"])
    m3.metric(label="Seções Válidas", value=f"{total_secoes} seções")
    m4.metric(label="Total de Itens", value=f"{total_itens_brutos} itens")

    st.divider()

    # 3. Seleção de Modo
    st.subheader("🎯 Selecione o Tipo de Relatório Desejado:")
    
    modo_escolhido = st.radio(
        "Modo de geração:",
        options=["primeiro", "segundo", "todos"],
        format_func=lambda x: {
            "primeiro": "🌅 1º Relatório do Dia (Manhã) — Gera todos até agora e salva a base",
            "segundo": "🌇 2º Relatório do Dia (Tarde) — Gera APENAS as novas alterações da tarde",
            "todos": "📋 Relatório Consolidado — Gera todas as alterações do dia completo"
        }[x],
        label_visibility="collapsed"
    )

    # Filtragem dos dados de acordo com o modo
    secoes_filtradas, todos_itens_filtrados = filtrar_dados(dados, modo=modo_escolhido)
    total_filtrados = len(todos_itens_filtrados)

    st.write("")

    # Exibe informações sobre o resultado do filtro
    if modo_escolhido == "segundo":
        if total_filtrados == 0:
            st.warning("⚠️ **Nenhuma nova alteração encontrada para a tarde!**\n\nTodos os itens desta planilha já haviam sido impressos no 1º relatório da manhã.")
        else:
            st.success(f"✨ **{total_filtrados} novos itens encontrados** após o 1º relatório da manhã!")
    elif modo_escolhido == "primeiro":
        st.info(f"ℹ️ Serão incluídos **{total_filtrados} itens**. Ao baixar o PDF, estes itens ficarão salvos como a base da manhã.")
    else:
        st.info(f"ℹ️ Serão incluídos todos os **{total_filtrados} itens** encontrados na planilha.")

    # 4. Geração do PDF e Botão de Download
    if total_filtrados > 0:
        info_turno = {
            "primeiro": " - 1º TURNO (MANHÃ)",
            "segundo": " - 2º TURNO (NOVAS ALTERAÇÕES)",
            "todos": " - CONSOLIDADO DO DIA"
        }.get(modo_escolhido, "")

        pdf_buffer = io.BytesIO()
        sucesso = gerar_pdf(
            dados=dados,
            caminho_pdf_saida=pdf_buffer,
            secoes_filtradas=secoes_filtradas,
            info_turno=info_turno,
            forcar_quebra_secao=False
        )

        if sucesso:
            pdf_bytes = pdf_buffer.getvalue()
            nome_arquivo_pdf = {
                "primeiro": "1_RELATORIO_PRECOS.pdf",
                "segundo": "2_RELATORIO_PRECOS_NOVOS.pdf",
                "todos": "RELATORIO_ALTERACAO_PRECO.pdf"
            }.get(modo_escolhido, "RELATORIO_PRECOS.pdf")

            def on_download_click():
                if modo_escolhido in ["primeiro", "segundo"]:
                    data_salvar = dados['periodo'] if dados['periodo'] else datetime.now().strftime("%d/%m/%Y")
                    if " a " in data_salvar:
                        data_salvar = data_salvar.split(" a ")[-1].strip()
                    salvar_historico(data_salvar, todos_itens_filtrados, info_execucao=modo_escolhido)

            st.download_button(
                label=f"📥 Baixar Relatório em PDF ({total_filtrados} itens)",
                data=pdf_bytes,
                file_name=nome_arquivo_pdf,
                mime="application/pdf",
                on_click=on_download_click,
                type="primary",
                use_container_width=True
            )

    # 5. Tabela de Pré-visualização
    with st.expander(f"🔍 Pré-visualizar Tabela de Produtos ({total_filtrados} itens selecionados)", expanded=False):
        if total_filtrados > 0:
            linhas_tabela = []
            for sec, itens in secoes_filtradas.items():
                for it in itens:
                    linhas_tabela.append({
                        'Seção': sec,
                        'Código': it['cod'],
                        'Descrição': it['descricao'],
                        'Preço': it['venda_atual'],
                        'Hora': it['hora']
                    })
            df_preview = pd.DataFrame(linhas_tabela)
            st.dataframe(df_preview, use_container_width=True, hide_index=True)
        else:
            st.write("Nenhum item para exibir com os filtros atuais.")
else:
    # Boas-vindas nativas do Streamlit com alta legibilidade (100% adaptável a tema claro ou escuro)
    with st.container():
        st.info("👋 **Bem-vindo ao Gerador de Relatórios de Preço!**\n\nPara começar, faça o upload da planilha exportada do Varejofácil no campo acima.")
        
        col1, col2 = st.columns(2)
        with col1:
            with st.container(border=True):
                st.markdown("### 📋 Como Exportar no Varejofácil")
                st.markdown("""
                1. Acesse: **Home ➔ Venda ➔ Relatórios preço** *(Listagem de preços)*
                2. **Loja:** `1` *(SUPERMERCADO JEAN LTDA)*
                3. **Origem:** Selecionar `REAJUSTE INDIVIDUAL` e `NOTA FISCAL`
                4. **Tipo:** 🔘 **Produtos com Preços Alterados** *(Obrigatório)*
                5. **Período:** Data de Hoje *(Inicial e Final)*
                6. **Formato:** 🔘 **Analítico** *(Obrigatório)*
                7. **Quebra:** 🔘 **Seção** *(Obrigatório)*
                8. Clique no botão **Exportar** e envie o `.xlsx` aqui!
                """)
        with col2:
            with st.container(border=True):
                st.markdown("### ⚡ Vantagens da Rotina Diária")
                st.markdown("""
                - 🌅 **Pela Manhã (~12h/13h):** Gera o 1º lote e grava a base do dia.
                - 🌇 **À Tarde (~17h/18h):** Gera **apenas os itens novos**, sem repetir nada da manhã.
                - 🥬 **Hortifrúti:** Filtrado e removido automaticamente.
                - 📄 **Impressão Otimizada:** Aproveitamento contínuo de páginas.
                """)
