import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns # type: ignore
import os
from config import ARQUIVO_EXCEL, ARQUIVO_SOCIOECONOMICO, NOME_DA_PLANILHA_RESULTADOS, NOME_DA_PLANILHA_MATERIAS, NOME_DA_PLANILHA_SOCIO, PASTA_ANALISES

# --- Funções Auxiliares e de Análise (sem alterações) ---

def parse_questoes(texto_questoes):
    """Interpreta o texto da coluna 'Questões' (ex: '1-10, 15')."""
    questoes = set()
    if not isinstance(texto_questoes, str): return []
    partes = texto_questoes.split(',')
    for parte in partes:
        parte = parte.strip()
        if '-' in parte:
            inicio, fim = map(int, parte.split('-'))
            questoes.update(range(inicio, fim + 1))
        elif parte.isdigit():
            questoes.add(int(parte))
    return sorted(list(questoes))

def ler_mapeamento_materias(caminho_excel, nome_planilha):
    """Lê o mapeamento de matérias e questões da aba 'Materias' do Excel."""
    print("📚 Lendo mapeamento de matérias do Excel...")
    try:
        df_materias = pd.read_excel(caminho_excel, sheet_name=nome_planilha, header=4)
        mapeamento = {row['Matéria']: parse_questoes(row['Questões']) for _, row in df_materias.iterrows() if pd.notna(row['Matéria']) and pd.notna(row['Questões'])}
        if not mapeamento: raise ValueError("Nenhum mapeamento válido encontrado.")
        return mapeamento
    except Exception as e:
        print(f"❌ ERRO ao ler a aba 'Materias': {e}")
        return None

def analisar_estatisticas_gerais(df_resultados, pasta_saida):
    """Calcula estatísticas descritivas das notas e salva em um arquivo de texto."""
    print("📊 Gerando estatísticas descritivas...")
    # (Código interno sem alterações)
    notas = df_resultados['% de Acertos']
    caminho_arquivo = os.path.join(pasta_saida, 'estatisticas_gerais.txt')
    with open(caminho_arquivo, 'w', encoding='utf-8') as f:
        f.write(f"Média das Notas: {notas.mean():.2f}%\n")
        f.write(f"Mediana das Notas: {notas.median():.2f}%\n")
        f.write(f"Desvio Padrão: {notas.std():.2f}%\n")
        f.write(f"Nota Máxima: {notas.max():.2f}%\n")
        f.write(f"Nota Mínima: {notas.min():.2f}%\n")
    print(f"✅ Relatório de estatísticas salvo.")

def gerar_distribuicao_notas(df_resultados, pasta_saida):
    """Cria e salva um histograma da distribuição das notas finais."""
    print("📊 Gerando gráfico de distribuição de notas...")
    # (Código interno sem alterações)
    plt.figure(figsize=(10, 6)); sns.histplot(df_resultados['% de Acertos'], kde=True, bins=10, color='skyblue')
    plt.title('Distribuição das Notas Finais dos Alunos'); plt.xlabel('Porcentagem de Acertos (%)'); plt.ylabel('Número de Alunos'); plt.xlim(0, 100)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    caminho_arquivo = os.path.join(pasta_saida, 'distribuicao_notas.png'); plt.savefig(caminho_arquivo); plt.close()
    print(f"✅ Gráfico de distribuição salvo.")

def gerar_boxplot_por_materia(df_completo, mapeamento, pasta_saida):
    """
    Calcula a nota de cada aluno por matéria e gera um boxplot.
    Esta versão usa uma abordagem vetorizada para maior robustez e performance.
    """
    print("📊 Gerando boxplot de desempenho por matéria...")
    
    gabarito = df_completo.iloc[0]
    respostas_alunos = df_completo.iloc[1:]
    
    dados_boxplot = []
    
    # Itera sobre cada MATÉRIA, em vez de cada aluno.
    for materia, questoes in mapeamento.items():
        # Filtra as colunas (questões) que pertencem a esta matéria.
        colunas_materia = [q for q in questoes if q in respostas_alunos.columns]
        
        if not colunas_materia:
            continue
            
        # Seleciona as respostas de TODOS os alunos apenas para as questões da matéria atual.
        respostas_subset = respostas_alunos[colunas_materia]
        gabarito_subset = gabarito[colunas_materia]
        
        # Compara o DataFrame de respostas com a Série do gabarito.
        # O método .eq(gabarito_subset) compara cada linha de respostas com o gabarito.
        # .sum(axis=1) soma os acertos (True values) para cada aluno.
        acertos_por_aluno = respostas_subset.eq(gabarito_subset).sum(axis=1)
        
        # Calcula o percentual de acertos para cada aluno nesta matéria.
        percentuais = (acertos_por_aluno / len(colunas_materia)) * 100
        
        # Adiciona os resultados à lista de dados para o gráfico.
        for p in percentuais:
            dados_boxplot.append({'Matéria': materia, 'Percentual de Acertos': p})

    if not dados_boxplot:
        print("⚠️  Não foi possível gerar o gráfico por matéria. Verifique o mapeamento.")
        return
        
    df_plot = pd.DataFrame(dados_boxplot)

    # O código para gerar o gráfico continua o mesmo.
    plt.figure(figsize=(12, 7))
    sns.boxplot(x='Matéria', y='Percentual de Acertos', data=df_plot, palette='viridis', hue='Matéria', legend=False)
    sns.stripplot(x='Matéria', y='Percentual de Acertos', data=df_plot, color=".25", alpha=0.5)
    plt.title('Distribuição de Notas por Matéria')
    plt.xlabel('Matéria')
    plt.ylabel('Acertos (%)')
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    caminho_arquivo = os.path.join(pasta_saida, 'boxplot_desempenho_por_materia.png')
    plt.tight_layout()
    plt.savefig(caminho_arquivo)
    plt.close()
    
    print(f"✅ Gráfico de boxplot por matéria salvo.")

def analisar_dificuldade_questoes(df_completo, pasta_saida):
    """(Análise Extra) Identifica as questões mais fáceis e difíceis."""
    print("📊 Gerando análise de dificuldade das questões...")
    # (Código interno sem alterações)
    gabarito = df_completo.iloc[0]; respostas_alunos = df_completo.iloc[1:]
    acertos_por_questao = {q: ((respostas_alunos[q].values == gabarito[q]).sum() / len(respostas_alunos)) * 100 for q in respostas_alunos.columns if isinstance(q, int)}
    df_dificuldade = pd.Series(acertos_por_questao).sort_values()
    caminho_txt = os.path.join(pasta_saida, 'dificuldade_questoes.txt')
    with open(caminho_txt, 'w', encoding='utf-8') as f:
        f.write("As 5 questões MAIS DIFÍCEIS (menor % de acerto):\n")
        f.writelines(f"  - Questão {q}: {p:.1f}% de acerto\n" for q, p in df_dificuldade.head(5).items())
        f.write("\nAs 5 questões MAIS FÁCEIS (maior % de acerto):\n")
        f.writelines(f"  - Questão {q}: {p:.1f}% de acerto\n" for q, p in df_dificuldade.tail(5).iloc[::-1].items())
    plt.figure(figsize=(15, 7)); df_dificuldade.plot(kind='bar', color='coral')
    plt.title('Percentual de Acerto por Questão'); plt.xlabel('Número da Questão'); plt.ylabel('Acertos (%)')
    caminho_img = os.path.join(pasta_saida, 'dificuldade_questoes.png'); plt.tight_layout(); plt.savefig(caminho_img); plt.close()
    print(f"✅ Análise de dificuldade concluída.")

# --- NOVAS FUNÇÕES DE ANÁLISE SOCIOECONÔMICA ---

def analisar_desempenho_por_faixa_salarial(df_merged, pasta_saida):
    """Gera um boxplot do desempenho por faixa salarial."""
    print("📊 Gerando análise de desempenho por faixa salarial...")
    coluna_analise = 'FAIXA SALARIAL'
    if coluna_analise not in df_merged.columns:
        print(f"⚠️ Coluna '{coluna_analise}' não encontrada. Análise pulada.")
        return
        
    # Ordena as faixas salariais para melhor visualização no gráfico
    ordem_faixas = sorted(df_merged[coluna_analise].unique())

    plt.figure(figsize=(12, 7))
    sns.boxplot(x=coluna_analise, y='% de Acertos', data=df_merged, order=ordem_faixas, palette='plasma', hue=coluna_analise, legend=False)
    plt.title('Desempenho por Faixa Salarial Familiar', fontsize=16)
    plt.xlabel('Faixa Salarial', fontsize=12)
    plt.ylabel('Percentual de Acertos (%)', fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    caminho_arquivo = os.path.join(pasta_saida, 'desempenho_por_faixa_salarial.png')
    plt.tight_layout(); plt.savefig(caminho_arquivo); plt.close()
    print(f"✅ Gráfico por faixa salarial salvo.")

def analisar_desempenho_por_faixa_etaria(df_merged, pasta_saida):
    """Gera um boxplot do desempenho por faixa etária."""
    print("📊 Gerando análise de desempenho por faixa etária...")
    coluna_analise = 'Qual a sua idade?'
    if coluna_analise not in df_merged.columns:
        print(f"⚠️ Coluna '{coluna_analise}' não encontrada. Análise pulada.")
        return

    # Cria faixas etárias para agrupar os dados
    bins = [0, 17, 19, 21, 100]
    labels = ['Até 17', '18-19', '20-21', '22+']
    df_merged['Faixa Etária'] = pd.cut(df_merged[coluna_analise], bins=bins, labels=labels, right=True)

    plt.figure(figsize=(10, 6))
    sns.boxplot(x='Faixa Etária', y='% de Acertos', data=df_merged, palette='magma', hue='Faixa Etária', legend=False)
    plt.title('Desempenho por Faixa Etária', fontsize=16)
    plt.xlabel('Faixa Etária (anos)', fontsize=12)
    plt.ylabel('Percentual de Acertos (%)', fontsize=12)
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    caminho_arquivo = os.path.join(pasta_saida, 'desempenho_por_faixa_etaria.png')
    plt.tight_layout(); plt.savefig(caminho_arquivo); plt.close()
    print(f"✅ Gráfico por faixa etária salvo.")

def analisar_correlacao_distancia_nota(df_merged, pasta_saida):
    """Gera um gráfico de dispersão entre distância e nota."""
    print("📊 Gerando análise de correlação Distância x Nota...")
    coluna_analise = 'Quantos quilômetros aproximadamente são de distância da sua residência até o Insper (Rua Quatá, 200 - Vila Olímpia)? (Escreva apenas números)'
    if coluna_analise not in df_merged.columns:
        print(f"⚠️ Coluna '{coluna_analise}' não encontrada. Análise pulada.")
        return
    
    # Converte para numérico, tratando erros
    df_merged[coluna_analise] = pd.to_numeric(df_merged[coluna_analise], errors='coerce')
    df_plot = df_merged.dropna(subset=[coluna_analise, '% de Acertos'])

    plt.figure(figsize=(10, 6))
    sns.regplot(x=coluna_analise, y='% de Acertos', data=df_plot, scatter_kws={'alpha':0.6}, line_kws={'color':'red'})
    plt.title('Nota vs. Distância da Residência ao Insper', fontsize=16)
    plt.xlabel('Distância (km)', fontsize=12)
    plt.ylabel('Percentual de Acertos (%)', fontsize=12)
    plt.grid(True, linestyle='--', alpha=0.6)

    caminho_arquivo = os.path.join(pasta_saida, 'correlacao_distancia_nota.png')
    plt.tight_layout(); plt.savefig(caminho_arquivo); plt.close()
    print(f"✅ Gráfico de correlação Distância x Nota salvo.")

def analisar_desempenho_por_locomocao(df_merged, pasta_saida):
    """Gera um boxplot do desempenho por meio de locomoção."""
    print("📊 Gerando análise de desempenho por meio de locomoção...")
    coluna_analise = 'Qual meio de locomoção será usado para sua ida ao Insper?'
    if coluna_analise not in df_merged.columns:
        print(f"⚠️ Coluna '{coluna_analise}' não encontrada. Análise pulada.")
        return

    plt.figure(figsize=(12, 7))
    sns.boxplot(x=coluna_analise, y='% de Acertos', data=df_merged, palette='crest', hue=coluna_analise, legend=False)
    plt.title('Desempenho por Meio de Locomoção', fontsize=16)
    plt.xlabel('Meio de Locomoção', fontsize=12)
    plt.ylabel('Percentual de Acertos (%)', fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    caminho_arquivo = os.path.join(pasta_saida, 'desempenho_por_locomocao.png')
    plt.tight_layout(); plt.savefig(caminho_arquivo); plt.close()
    print(f"✅ Gráfico por meio de locomoção salvo.")

def analisar_desempenho_por_regiao(df_merged, pasta_saida):
    """
    Gera um boxplot do desempenho dos alunos por região geográfica,
    agrupando regiões menos comuns em 'Outros'.
    """
    print("📊 Gerando análise de desempenho por região geográfica (agrupada)...")
    coluna_analise = 'Distribuição geográfica'

    # Verifica se a coluna necessária existe no DataFrame
    if coluna_analise not in df_merged.columns:
        print(f"⚠️ Aviso: A coluna '{coluna_analise}' não foi encontrada. Análise pulada.")
        return
        
    # --- LÓGICA DE AGRUPAMENTO ---
    # 1. Define as categorias principais que queremos manter.
    zonas_principais = ['Zona Oeste', 'Zona Sul', 'Zona Norte', 'Zona Leste', 'Centro']
    
    # 2. Cria uma cópia da coluna para não alterar os dados originais.
    df_plot = df_merged.copy()
    
    # 3. Aplica a lógica: se a região não estiver na lista principal, classifica como 'Outros'.
    df_plot['Região Agrupada'] = df_plot[coluna_analise].apply(
        lambda x: x if pd.notna(x) and x in zonas_principais else 'Outros'
    )
    # --- FIM DA LÓGICA DE AGRUPAMENTO ---

    plt.figure(figsize=(12, 8))
    
    # Cria o boxplot usando a nova coluna 'Região Agrupada'
    sns.boxplot(
        x='Região Agrupada', 
        y='% de Acertos', 
        data=df_plot, 
        palette='coolwarm', 
        hue='Região Agrupada', 
        legend=False,
        order=zonas_principais + ['Outros'] # Define a ordem das colunas no gráfico
    )
    
    sns.stripplot(
        x='Região Agrupada', 
        y='% de Acertos', 
        data=df_plot, 
        color=".25", 
        alpha=0.6,
        order=zonas_principais + ['Outros']
    )

    plt.title('Desempenho por Região Geográfica', fontsize=16)
    plt.xlabel('Região', fontsize=12)
    plt.ylabel('Percentual de Acertos (%)', fontsize=12)
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    # Salva a imagem na pasta de análises
    caminho_arquivo = os.path.join(pasta_saida, 'desempenho_por_regiao.png')
    try:
        plt.tight_layout()
        plt.savefig(caminho_arquivo)
        plt.close()
        print(f"✅ Gráfico de desempenho por região (agrupada) salvo.")
    except Exception as e:
        print(f"❌ ERRO ao salvar o gráfico de desempenho por região: {e}")

# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    if not os.path.exists(PASTA_ANALISES):
        os.makedirs(PASTA_ANALISES)

    # --- CARREGAMENTO E UNIÃO DOS DADOS ---
    try:
        df_resultados = pd.read_excel(ARQUIVO_EXCEL, sheet_name=NOME_DA_PLANILHA_RESULTADOS, header=4)
        df_socio = pd.read_excel(ARQUIVO_SOCIOECONOMICO, sheet_name=NOME_DA_PLANILHA_SOCIO)

        # Une as duas tabelas usando o nome do aluno como chave
        df_merged = pd.merge(
            df_resultados,
            df_socio,
            left_on='Aluno/Questões',
            right_on='Nome Completo',
            how='inner' # 'inner' une apenas alunos presentes em ambas as planilhas
        )
        print("✅ Dados de resultados e socioeconômicos unidos com sucesso.")
        
        # Prepara DataFrames para as funções antigas
        df_alunos = df_resultados.drop(columns=['Aluno/Questões'])
        df_resultados_finais = df_resultados[df_resultados['Aluno/Questões'] != 'Gabarito'][['Acertos', '% de Acertos']]

    except FileNotFoundError as e:
        print(f"❌ ERRO: Arquivo não encontrado: {e.filename}. Verifique os nomes e caminhos.")
        exit()
    except PermissionError:
        print(f"❌ ERRO: O arquivo '{ARQUIVO_EXCEL}' ou '{ARQUIVO_SOCIOECONOMICO}' está aberto em outro programa. Feche-o e tente novamente.")
        exit()
    except Exception as e:
        print(f"❌ ERRO ao carregar ou unir os dados: {e}")
        exit()

    # --- EXECUÇÃO DAS ANÁLISES ---
    print("\n--- Iniciando Análises de Desempenho ---")
    mapeamento_materias = ler_mapeamento_materias(ARQUIVO_EXCEL, NOME_DA_PLANILHA_MATERIAS)
    if mapeamento_materias:
        gerar_boxplot_por_materia(df_alunos, mapeamento_materias, PASTA_ANALISES)
    
    analisar_estatisticas_gerais(df_resultados_finais, PASTA_ANALISES)
    gerar_distribuicao_notas(df_resultados_finais, PASTA_ANALISES)
    analisar_dificuldade_questoes(df_alunos, PASTA_ANALISES)
    print("✅ Análises de desempenho concluídas com sucesso.")

    print("\n--- Iniciando Análises Socioeconômicas ---")
    if not df_merged.empty:
        analisar_desempenho_por_faixa_salarial(df_merged, PASTA_ANALISES)
        analisar_desempenho_por_faixa_etaria(df_merged, PASTA_ANALISES)
        analisar_correlacao_distancia_nota(df_merged, PASTA_ANALISES)
        analisar_desempenho_por_locomocao(df_merged, PASTA_ANALISES)
        analisar_desempenho_por_regiao(df_merged, PASTA_ANALISES)
        print("✅ Análises socioeconômicas concluídas com sucesso.")
    else:
        print("⚠️ Nenhum aluno em comum encontrado entre as planilhas de resultados e socioeconômica. Análises socioeconômicas puladas.")

    print("\n🚀 Todas as análises foram concluídas com sucesso!")