import cv2
import numpy as np
import os
from pdf2image import convert_from_path
import shutil
import csv

entrada_pdf = "imagens_pdf"
saida_img = "imagens_gabaritos"
temp_dir = "temp"
csv_saida = "utils/respostas.csv"

os.makedirs(saida_img, exist_ok=True)
os.makedirs(temp_dir, exist_ok=True)

def remove_readonly(func, path, _):
    os.chmod(path, 0o777)
    func(path)


# -------- NOVA FUNÇÃO DE DETECÇÃO AUTOMÁTICA DE CÍRCULOS --------
def detectar_respostas(temp_folder):
    respostas = []
    alternativas = ["A", "B", "C", "D", "E"]

    for i in range(1, 61):
        caminho = os.path.join(temp_folder, f"questao_{i:02d}.jpg")
        imagem = cv2.imread(caminho)
        if imagem is None:
            respostas.append("")
            continue

        # Cortar para remover número da questão (ajuste conforme necessário)
        h, w = imagem.shape[:2]
        imagem = imagem[:, int(w * 0.1):-int(w * 0.02)]  # ← ajuste novo

        # Pré-processamento
        gray = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
        gray = cv2.medianBlur(gray, 5)
        _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)

        # Contornos
        contornos, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        bolinhas = []

        for cnt in contornos:
            area = cv2.contourArea(cnt)
            perimetro = cv2.arcLength(cnt, True)
            if perimetro == 0:
                continue
            circularidade = 4 * np.pi * (area / (perimetro ** 2))

            if 0.5 < circularidade < 1.2 and 3000 < area < 12000: # ← ajuste novo
                M = cv2.moments(cnt)
                if M["m00"] == 0:
                    continue
                cx = int(M["m10"] / M["m00"])
                cy = int(M["m01"] / M["m00"])
                bolinhas.append((cx, cy, cnt))

        if len(bolinhas) != 5:
            print(f"⚠️ Questão {i}: detectou {len(bolinhas)} bolinhas (esperado: 5)")
            respostas.append("Z")
            continue

        # Ordenar da esquerda pra direita
        bolinhas.sort(key=lambda x: x[0])

        # Calcular o preenchimento (quanto mais branco na máscara, mais marcada)
        preenchimentos = []
        for cx, cy, cnt in bolinhas:
            mask = np.zeros_like(gray)
            cv2.drawContours(mask, [cnt], -1, 255, -1)
            media = cv2.mean(thresh, mask=mask)[0]
            preenchimentos.append(media)

        idx = np.argmax(preenchimentos)  # maior preenchimento = mais branca = marcada
        respostas.append(alternativas[idx])



    #Salvas as questões com erro na pasta "erros"
    # Se não detectou 5 bolinhas, salva a imagem e o threshold na pasta de erros
    if len(bolinhas) != 5:
        erro_dir = os.path.join("erros", f"questao_{i:02d}")
        os.makedirs(erro_dir, exist_ok=True)
        cv2.imwrite(os.path.join(erro_dir, "imagem.jpg"), imagem)
        cv2.imwrite(os.path.join(erro_dir, "thresh.jpg"), thresh)

    return respostas


# ---------------------------------------------------------------


# função que processa a imagem do aluno, recorta a área de questões, aplica a máscara e salva na temp
def processar_imagem(path_img, temp_dir="temp"):
    image = cv2.imread(path_img)

    # Coordenadas do bloco onde estão as questões
    x1, x2 = 270, 3480
    y1, y2 = 1430, 4580
    bloco = image[y1:y2, x1:x2]

    # Máscara para isolar círculos pretos
    hsv = cv2.cvtColor(bloco, cv2.COLOR_BGR2HSV)
    lower_black = np.array([0, 0, 0])
    upper_black = np.array([180, 255, 80])
    mask_hsv = cv2.inRange(hsv, lower_black, upper_black)

    h, w = bloco.shape[:2]
    altura = h // 15
    largura = int(w / 4)

    # Limpa a pasta temporária
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir, onerror=remove_readonly)
    os.makedirs(temp_dir)

    margem = 10
    count = 1
    for coluna in range(4):           # ← percorre colunas primeiro
        for linha in range(15):       # ← depois percorre as linhas

            y_ini = linha * altura
            y_fim = (linha + 1) * altura
            x_ini = coluna * largura
            x_fim = (coluna + 1) * largura

            ry1 = max(0, y_ini - margem)
            ry2 = min(h, y_fim + margem)
            rx1 = x_ini if coluna == 0 else max(0, x_ini - margem)
            rx2 = min(w, x_fim + margem)

            recorte_img = bloco[ry1:ry2, rx1:rx2]

            if recorte_img.size == 0:
                print(f"❌ ERRO: questão {count:02d} está vazia! Coordenadas: {rx1}:{rx2}, {ry1}:{ry2}")
                continue

            cv2.imwrite(f"{temp_dir}/questao_{count:02d}.jpg", recorte_img)
            count += 1


    return detectar_respostas(temp_dir)



# ----------------------------
# Etapa 1 – Converter PDFs
# ----------------------------
imagens_convertidas = []

for arquivo in sorted(os.listdir(entrada_pdf)):
    if arquivo.endswith(".pdf"):
        nome_base = os.path.splitext(arquivo)[0]
        paginas = convert_from_path(
        os.path.join(entrada_pdf, arquivo),
        dpi=150,  # reduz o peso e acelera muito!
        poppler_path=r"C:\poppler\Library\bin"  # ajuste se necessário
        )

        for i, pagina in enumerate(paginas):
            img_nome = f"{nome_base}_p{i+1}.jpg"
            caminho_img = os.path.join(saida_img, img_nome)
            pagina.save(caminho_img, "JPEG")
            imagens_convertidas.append((nome_base, caminho_img))
        print(f"✅ {len(paginas)} páginas convertidas para imagens!")




# ----------------------------
# Etapa 2 – Separar gabarito
# ----------------------------

respostas_finais = []

gabarito = processar_imagem(imagens_convertidas[0][1])  # primeira imagem é o gabarito

for idx, (nome, caminho) in enumerate(imagens_convertidas[1:], start=1):
    print(f"\n🔍 Processando aluno {idx} → {caminho}")
    temp_aluno = os.path.join("debug_temp", f"aluno_{idx:02d}")
    respostas = processar_imagem(caminho, temp_dir=temp_aluno)

    if len(respostas) != 60:
        print(f"⚠️ Alerta: Aluno {idx} teve {len(respostas)} respostas detectadas (esperado: 60)")
    else:
        print(f"✅ Aluno {idx}: 60 respostas detectadas.")

    respostas_finais.append(respostas)




# ----------------------------
# Etapa 3 – Salvar CSV
# ----------------------------
colunas = [f"{i+1}" for i in range(60)]
with open(csv_saida, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(colunas)
    writer.writerow(gabarito)
    for resp in respostas_finais:
        writer.writerow(resp)

print(f"✅ Respostas salvas em '{csv_saida}' com sucesso!")