import os
from fpdf import FPDF # type: ignore
from fpdf.enums import XPos, YPos # type: ignore
from datetime import datetime
from config import PASTA_ANALISES

# --- Classe para o Relatório em PDF (com API moderna) ---
class PDF(FPDF):
    def header(self):
        """Define o cabeçalho do PDF, que aparece em todas as páginas."""
        self.set_font('Helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de Análise de Provas do Cursinho Insper', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(10)

    def footer(self):
        """Define o rodapé do PDF, que aparece em todas as páginas."""
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')

    def chapter_title(self, title):
        """Cria um título de capítulo padronizado."""
        self.set_font('Helvetica', 'B', 14)
        self.cell(0, 10, title, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L')
        self.ln(5)

    def chapter_body(self, text_file_path=None, image_path=None, intro_text=""):
        """Adiciona o corpo do capítulo, com texto introdutório, lendo de arquivos de texto e/ou imagem."""
        if intro_text:
            self.set_font('Helvetica', '', 12)
            self.multi_cell(0, 10, intro_text)
            self.ln(5)

        if text_file_path and os.path.exists(text_file_path):
            with open(text_file_path, 'r', encoding='utf-8') as f:
                text = f.read()
            self.set_font('Courier', '', 12) # Fonte monoespaçada para texto de relatório
            self.multi_cell(0, 8, text)
            self.ln()

        if image_path and os.path.exists(image_path):
            # A centralização da imagem é feita calculando a posição x
            image_width = self.w - 40 # Largura da imagem
            x_position = (self.w - image_width) / 2
            self.image(image_path, x=x_position, w=image_width)
            self.ln(5)
            
# --- Função Principal para Gerar o Relatório ---
def criar_pdf_consolidado():
    """
    Lê os arquivos da pasta 'analises' e os compila em um único relatório PDF.
    """
    print("📄 Iniciando a criação do relatório em PDF (versão atualizada)...")
    
    if not os.path.exists(PASTA_ANALISES):
        print(f"❌ ERRO: A pasta '{PASTA_ANALISES}' não foi encontrada.")
        print("Por favor, execute o script 'analise_resultados.py' primeiro.")
        return

    pdf = PDF()
    
    # --- Página de Título ---
    pdf.add_page()
    pdf.set_font('Helvetica', 'B', 24)
    pdf.cell(0, 80, '', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 10, 'Relatório de Análise do Vestibulinho', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', 'I', 16)
    pdf.cell(0, 20, f"Gerado em: {datetime.now().strftime('%d/%m/%Y')}", align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    # --- Capítulo 1: Análise Geral ---
    pdf.add_page()
    pdf.chapter_title('1. Visão Geral do Desempenho da Turma')
    pdf.chapter_body(
        intro_text=(
            "Esta seção apresenta as estatísticas descritivas do desempenho geral da turma. "
            "Abaixo estão os principais indicadores de performance, seguidos por um histograma "
            "que ilustra a distribuição das notas finais (em porcentagem de acertos)."
        ),
        text_file_path=os.path.join(PASTA_ANALISES, 'estatisticas_gerais.txt'),
        image_path=os.path.join(PASTA_ANALISES, 'distribuicao_notas.png')
    )
    
    # --- Capítulo 2: Desempenho por Matéria ---
    pdf.add_page()
    pdf.chapter_title('2. Análise de Desempenho por Matéria')
    pdf.chapter_body(
        intro_text=(
            "O gráfico de boxplot a seguir mostra a distribuição das notas dos alunos em cada matéria. "
            "Cada 'caixa' representa onde 50% dos alunos se concentram (entre o 1º e o 3º quartil). "
            "A linha no meio da caixa é a mediana. Isso ajuda a identificar matérias com maior ou menor "
            "desempenho geral e a variabilidade dos resultados."
        ),
        image_path=os.path.join(PASTA_ANALISES, 'boxplot_desempenho_por_materia.png')
    )

    # --- Capítulo 3: Dificuldade das Questões ---
    pdf.add_page()
    pdf.chapter_title('3. Análise de Dificuldade das Questões')
    pdf.chapter_body(
        intro_text=(
            "Para identificar os pontos fortes e fracos do conteúdo, esta seção detalha o percentual "
            "de acerto para cada questão. O relatório destaca as 5 questões mais fáceis e as 5 mais "
            "difíceis. O gráfico de barras exibe o desempenho em todas as questões, ordenadas da "
            "mais difícil para a mais fácil."
        ),
        text_file_path=os.path.join(PASTA_ANALISES, 'dificuldade_questoes.txt'),
        image_path=os.path.join(PASTA_ANALISES, 'dificuldade_questoes.png')
    )

    # --- INÍCIO DAS NOVAS ANÁLISES SOCIOECONÔMICAS ---

    # --- Capítulo 4: Desempenho por Faixa Salarial ---
    pdf.add_page()
    pdf.chapter_title('4. Desempenho por Faixa Salarial Familiar')
    pdf.chapter_body(
        intro_text=(
            "Esta análise investiga se há uma relação entre a renda familiar dos alunos e seu desempenho na prova. "
            "O gráfico de boxplot abaixo compara a distribuição das notas para cada faixa salarial, permitindo "
            "observar tendências e a variabilidade dos resultados entre os diferentes grupos."
        ),
        image_path=os.path.join(PASTA_ANALISES, 'desempenho_por_faixa_salarial.png')
    )

    # --- Capítulo 5: Desempenho por Faixa Etária ---
    pdf.add_page()
    pdf.chapter_title('5. Desempenho por Faixa Etária')
    pdf.chapter_body(
        intro_text=(
            "A seguir, o desempenho dos alunos é agrupado por faixa etária. Este gráfico de boxplot "
            "ajuda a entender se a idade dos candidatos tem alguma correlação com as notas obtidas no vestibulinho."
        ),
        image_path=os.path.join(PASTA_ANALISES, 'desempenho_por_faixa_etaria.png')
    )

    # --- Capítulo 6: Desempenho por Região Geográfica ---
    pdf.add_page()
    pdf.chapter_title('6. Desempenho por Região Geográfica')
    pdf.chapter_body(
        intro_text=(
            "Analisamos aqui se a região de residência do aluno influencia seu desempenho. As principais "
            "zonas da cidade são comparadas, e as demais localidades foram agrupadas na categoria 'Outros' "
            "para uma visualização mais clara e focada."
        ),
        image_path=os.path.join(PASTA_ANALISES, 'desempenho_por_regiao.png')
    )

    # --- Capítulo 7: Desempenho por Meio de Locomoção ---
    pdf.add_page()
    pdf.chapter_title('7. Desempenho por Meio de Locomoção')
    pdf.chapter_body(
        intro_text=(
            "Esta seção explora a relação entre o meio de transporte utilizado pelo aluno para chegar ao "
            "cursinho e seu desempenho. O gráfico compara a distribuição de notas entre os diferentes meios de locomoção."
        ),
        image_path=os.path.join(PASTA_ANALISES, 'desempenho_por_locomocao.png')
    )

    # --- Capítulo 8: Correlação Nota vs. Distância ---
    pdf.add_page()
    pdf.chapter_title('8. Correlação entre Nota e Distância da Residência')
    pdf.chapter_body(
        intro_text=(
            "Este gráfico de dispersão investiga se existe uma correlação entre a distância (em km) que o aluno "
            "percorre de casa até o Insper e sua nota. A linha de regressão indica a tendência geral dos dados: "
            "uma linha plana, como a observada, sugere que não há uma correlação forte entre as duas variáveis."
        ),
        image_path=os.path.join(PASTA_ANALISES, 'correlacao_distancia_nota.png')
    )
    
    # --- FIM DAS NOVAS ANÁLISES ---

    caminho_pdf = os.path.join('Relatorio_Final.pdf')
    try:
        pdf.output(caminho_pdf)
        print(f"✅ Relatório em PDF gerado com sucesso!")
        print(f"   Arquivo salvo em: '{caminho_pdf}'")
    except Exception as e:
        print(f"❌ ERRO ao salvar o PDF: {e}")

# --- Ponto de Entrada do Script ---
if __name__ == "__main__":
    criar_pdf_consolidado()