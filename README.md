# 🏷️ Gerador de Relatório de Alteração de Preços

Aplicação web desenvolvida com **Streamlit** e **ReportLab** para processamento, filtragem inteligente por turnos (manhã/tarde) e geração de relatórios de alteração de preços em formato PDF a partir de planilhas exportadas do **Varejofácil**.

---

## 🚀 Como Executar Localmente

1. Clone o repositório:
```bash
git clone <URL_DO_REPOSITORIO>
cd Alteracao_Preco
```

2. Crie e ative o ambiente virtual:
```bash
python -m venv .venv
# No Windows:
.venv\Scripts\activate
# No Linux/Mac:
source .venv/bin/activate
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

4. Execute o aplicativo:
```bash
streamlit run app.py
```

---

## 📋 Como Publicar no Streamlit Community Cloud

1. Suba o projeto para um repositório no **GitHub**.
2. Acesse [share.streamlit.io](https://share.streamlit.io).
3. Conecte sua conta do GitHub e selecione o repositório.
4. Defina o arquivo principal como: **`app.py`**.
5. Clique em **Deploy**!

---

## 📖 Documentação e Manual Operacional
Consulte o arquivo [TUTORIAL_SISTEMA.md](TUTORIAL_SISTEMA.md) para o passo a passo completo de exportação no Varejofácil e uso do sistema.
