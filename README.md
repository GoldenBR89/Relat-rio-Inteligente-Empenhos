# 📦 Gestor Inteligente de Empenhos 

Este projeto automatiza o fluxo logístico de conferência de estoque e separação de pedidos da **Licitax**. Ele processa o **Balanço de Estoque** e cruza os dados com os **Pedidos de Venda**, gerando um relatório detalhado de separação e necessidades de compra.

---

## 🌟 Diferenciais do Projeto

* **Extração de Dados Complexos:** Identifica itens mesmo com formatações financeiras variadas no PDF (ex: quantidades como `177,000`).
* **Inteligência para Processadores:** Reconhece componentes com nomes extensos (ex: `PROC. R5 5600GT`) sem confundir o modelo com a quantidade.
* **Auto-Inclusão de Itens:** Se um pedido contiver um item fora do balanço, o sistema o adiciona ao relatório com saldo zero automaticamente.
* **Priorização Dinâmica:** Permite escolher critérios de desempate por volume de peças ou diversidade de SKUs.

## 🛠️ Tecnologias Utilizadas

* **Python 3.x**
* **Pandas:** Processamento de dados.
* **pdfplumber:** Extração técnica de texto de PDFs.
* **Tkinter:** Interface gráfica (GUI).
* **Openpyxl:** Formatação avançada de Excel (Zebra, alinhamentos e fórmulas).

## 📊 Estrutura do Relatório Gerado (.xlsx)

O programa gera um arquivo Excel com três abas:
1.  **Separação por Empenho:** Peças por pedido com efeito "zebra" para facilitar a leitura visual.
2.  **Resumo para Compras:** Consolidação do que precisa ser adquirido.
3.  **Estoque Inicial:** Espelho organizado e em ordem alfabética do estoque atual.

## 🚀 Como Executar

1. Instale as dependências necessárias:
   ```bash
   pip install pandas pdfplumber openpyxl

2. Execute o script pincipal:
   ```bash
   python calculadora_compras.py
