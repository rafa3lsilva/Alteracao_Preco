# 📖 Manual Operacional: Exportação no Varejofácil e Geração de Relatórios

Este manual orienta o passo a passo para **exportar a planilha correta no Varejofácil** e **gerar os relatórios em PDF formatados** para conferência e troca de etiquetas na loja.

---

## PARTE 1: Como Exportar a Planilha no Varejofácil

### 📍 Caminho no Sistema
Acesse o menu:
`Home` ➔ `Venda` ➔ `Relatórios preço` (Tela: **Listagem de preços**)

---

### ⚙️ Configuração dos Filtros e Opções

Preencha os campos exatamente como indicado abaixo:

| Campo / Opção | Configuração Obrigatória | Observação |
| :--- | :--- | :--- |
| **1. Loja\*** | `1 - SUPERMERCADO JEAN LTDA` | Sua loja |
| **2. Origem** | `REAJUSTE INDIVIDUAL` e `NOTA FISCAL` | Origens de alteração |
| **3. Tipo\*** | 🔘 **Produtos com Preços Alterados** | ⚠️ **OBRIGATÓRIO**: Filtra apenas os itens com alteração |
| **4. Período Inicial / Final\*** | `Data de Hoje` (ex: 04/09/2026) | Data da conferência |
| **5. Formato\*** | 🔘 **Analítico** | ⚠️ **OBRIGATÓRIO**: Traz horários e códigos corretos |
| **6. Quebra\*** | 🔘 **Seção** | ⚠️ **OBRIGATÓRIO**: Agrupa por departamentos |
| **7. Preços** | 🔘 **Varejo** | Preço de gôndola |
| **8. Ordem / Direção** | 🔘 **Descrição** / 🔘 **Ascendente** | Ordem alfabética dos itens |
| **9. Exibir** | ☑️ Em Linha &nbsp; ☑️ Fora de Linha<br>☑️ Com Preço &nbsp; ☑️ Sem Preço | Deixar todas as caixas marcadas |
| **10. Seção** | *(Sem Hortifrúti)* | O sistema filtra Hortifrúti automaticamente |

---

### 💾 Exportando o Arquivo
1. No canto superior direito da tela do Varejofácil, clique no botão **`Exportar`**.
2. Salve o arquivo em formato **Excel (`.xlsx`)** na sua máquina (por exemplo, na pasta `Alteracao_Preco` ou na Área de Trabalho).

---

## PARTE 2: Como Gerar o Relatório (Opção 1 - Painel Web Streamlit)

A interface web permite carregar a planilha diretamente pelo navegador, ver o resumo em tempo real e baixar o PDF pronto.

### 1. Abrir a Interface
- Dê dois cliques no arquivo: **`INICIAR_PAINEL_WEB.bat`**.
- O navegador abrirá automaticamente a tela do gerador.

### 2. Rotina da Manhã (1º Relatório)
1. Exporte a planilha no Varejofácil por volta das **12h / 13h**.
2. No painel web, **arraste ou selecione a planilha `.xlsx`**.
3. Selecione a opção: **`🌅 1º Relatório (Manhã / Salva Base)`**.
4. Clique em **`📥 Baixar Relatório em PDF`**.
5. *O sistema salvará na memória os itens gerados pela manhã.*

### 3. Rotina da Tarde (2º Relatório - Apenas Novidades)
1. Exporte novamente a planilha atualizada no Varejofácil no fim da tarde (**17h / 18h**).
2. Arraste a nova planilha para o painel web.
3. Selecione a opção: **`🌇 2º Relatório (Tarde / Apenas Novos)`**.
4. O sistema compara automaticamente com a base da manhã:
   - Se houver novas alterações, exibirá a quantidade e liberará o botão **`📥 Baixar Relatório em PDF`** (apenas com as novidades da tarde).
   - Se nenhuma alteração nova foi feita, exibirá um aviso amigável: *"Nenhuma nova alteração encontrada para a tarde!"*.

### 4. Relatório Consolidado (Opcional)
- Para gerar um PDF com todas as alterações do dia completo em um único documento, selecione **`📋 Consolidado (Todos os Itens do Dia)`**.

### 5. Reset da Memória
- Na barra lateral esquerda do painel, há o botão **`🗑️ Resetar Memória de Hoje`** caso queira reiniciar os registros do dia.

---

## PARTE 3: Como Gerar Pelos Atalhos Rápidos (`.bat`)

Se preferir não usar o navegador, você também pode usar os atalhos rápidos diretamente na pasta:

1. Salve a planilha exportada na pasta com o nome **`ALTERACAO_PRECO.xlsx`**.
2. Dê dois cliques no executável desejado:
   - **`1_PRIMEIRO_RELATORIO.bat`**: Gera o PDF da manhã (`1_RELATORIO_PRECOS.pdf`) e salva a base.
   - **`2_SEGUNDO_RELATORIO.bat`**: Compara com a base e gera apenas as novidades da tarde (`2_RELATORIO_PRECOS_NOVOS.pdf`).
   - **`3_RELATORIO_COMPLETO.bat`**: Gera todas as alterações do dia (`RELATORIO_ALTERACAO_PRECO.pdf`).
   - **`4_RESETAR_HISTORICO.bat`**: Zera o histórico do dia.

---

## 💡 Dúvidas Frequentes

### Por que o 2º relatório da tarde não repete os produtos da manhã?
O sistema possui uma memória diária (`.historico_alteracoes.json`). Ao gerar o 1º relatório, ele registra o código de barras, preço e hora de cada produto. Na exportação da tarde, ele lê a nova planilha e descarta tudo o que já foi impresso pela manhã, mantendo apenas itens novos ou que sofreram reajuste posterior.

### E se eu esquecer e exportar no modo Sintético ou com Hortifrúti?
- **Modo Sintético**: O sistema exibirá um alerta avisando que o formato recomendado é o *Analítico*.
- **Hortifrúti**: O sistema detecta e descarta automaticamente qualquer produto ou seção de Hortifrúti, garantindo que o PDF saia apenas com os setores do supermercado.
