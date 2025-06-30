# Corretor Automático de Gabaritos - Cursinho Insper


Este projeto realiza a **correção automática de gabaritos** utilizando **visão computacional com OpenCV**. 
O sistema lê folhas escaneadas em PDF, detecta as respostas marcadas, e salva um CSV com todas as respostas dos alunos.

---
## Como usar?:

### Input:
- **PDFs escaneados**: Coloque os PDFs escaneados na pasta `imagens_pdf/`. O primeiro PDF deve ser o gabarito, e os demais são os PDFs dos alunos.

### 1. Clone o repositório

```bash
git clone https://github.com/Insper-Data/auto-corretor-CursinhoInsper.git
cd seu-repo
```

### 2. Crie o ambiente virtual

```bash 
python -m venv .venv
```

#### No windows

```bash
.venv\Scripts\activate
```

#### No macOS/Linux

```bash
source .venv/bin/activate
```

### 3. Instale as dependências
```bash
pip install -r requirements.txt
```

### 4. Instale o Poppler (requisito para ler PDFs)

O Poppler é necessário para que o script consiga converter PDFs em imagens.

#### Passos:

1. **Baixe o Poppler para Windows:**

   👉 [https://github.com/oschwartz10612/poppler-windows/releases/](https://github.com/oschwartz10612/poppler-windows/releases/)

   Baixe a versão mais recente do arquivo `.zip`.

2. **Extraia a pasta:**  
   Por exemplo, extraia para:  
    `C:\poppler`
3. **Adicione o Poppler ao PATH:**

- Abra o menu Iniciar e pesquise por **“variáveis de ambiente”**.
- Clique em **“Editar variáveis de ambiente do sistema”**.
- Na seção **“Variáveis do sistema”**, encontre a variável `Path` e clique em **Editar...**.
- Clique em **Novo** e adicione o caminho:
  ```
  C:\poppler\Library\bin
  ```
- Confirme clicando em **OK** em todas as janelas.

4. **Teste no terminal:**  
Abra o terminal (cmd ou PowerShell) e digite:
```bash
pdfinfo
``` 


## Como usar?

### 1. Coloque os arquivos PDF escaneados na pasta imagens_pdf
    - Primeiro PDF: gabarito.
    - Os demais são os PDFs dos alunos

### 2. Execute o script principal:

```bash
python corretor.py
```

### 3. O programa irá:
O programa irá:

- Converter os PDFs em imagens.

- Recortar automaticamente o bloco de questões.

- Detectar as respostas preenchidas.

- Salvar o arquivo respostas.csv, com o gabarito na primeira linha e as respostas dos alunos abaixo


⚠️ Sobre detecções incompletas
Caso alguma questão não tenha 5 bolinhas detectadas, o programa irá mostrar um aviso no terminal, como:
```less
⚠️ Questao 39: detectou 4 bolinhas (esperado: 5)
```

Esses casos podem acontecer por:

Impressão desalinhada ou distorcida.

Bolinhas muito preenchidas ou pouco nítidas.

**Ausência de âncoras** nas folhas de gabarito, o que dificultou a padronização da correção automática.




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