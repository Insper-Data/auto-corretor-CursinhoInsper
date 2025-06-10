# Processamento e Análise de Provas do Cursinho Insper

## 📥 Inputs Necessários

Antes de rodar os scripts, certifique-se de que os seguintes arquivos estejam presentes na pasta [`utils/`](utils):

- [`planilha_final.xlsx`](utils/planilha_final.xlsx): Planilha Excel onde os resultados serão gravados e analisados.
- [`respostas.csv`](utils/respostas.csv): Arquivo CSV contendo o gabarito e as respostas dos alunos.
- [`socioeconomica.xlsx`](utils/socioeconomica.xlsx): Planilha Excel com informações socioeconômicas dos alunos.

> **Importante:** Certifique-se de que todos esses arquivos estejam fechados em outros programas (como Excel) antes de rodar os scripts, para evitar erros de permissão.

---

## ▶️ Ordem de Execução dos Scripts

1. **Processar as Provas**
   - Execute [`processar_provas.py`](processar_provas.py) para processar as respostas dos alunos, calcular acertos gerais e por matéria, e atualizar a planilha de resultados.
   - Exemplo de execução:
     ```sh
     python processar_provas.py
     ```

2. **Gerar Análises e Gráficos**
   - Execute [`analise_resultados.py`](analise_resultados.py) para gerar análises estatísticas, gráficos de desempenho e análises socioeconômicas. Os resultados serão salvos na pasta `analises/`.
   - Exemplo de execução:
     ```sh
     python analise_resultados.py
     ```

3. **Gerar o Relatório Consolidado em PDF**
   - Execute [`gerar_relatorio.py`](gerar_relatorio.py) para compilar todas as análises e gráficos em um relatório PDF final.
   - Exemplo de execução:
     ```sh
     python gerar_relatorio.py
     ```

---

## 📤 Outputs Gerados

- **Planilha Atualizada:**  
  O arquivo [`utils/planilha_final.xlsx`](utils/planilha_final.xlsx) será atualizado com os resultados processados.

- **Pasta de Análises:**  
  Será criada uma pasta `analises/` contendo gráficos e relatórios intermediários gerados pelo script de análise.

- **Relatório Final em PDF:**  
  O arquivo `Relatorio_Final.pdf` será gerado na raiz do projeto, consolidando todas as análises e gráficos em um único documento.

---

## ℹ️ Observações

- Sempre feche as planilhas antes de rodar os scripts para evitar erros de leitura/escrita.
- Certifique-se de que todas as dependências listadas em [`requirements.txt`](requirements.txt) estejam instaladas:
  ```sh
  pip install -r requirements.txt
  ```