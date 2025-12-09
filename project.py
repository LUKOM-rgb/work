import pandas as pd
import fitz  # PyMuPDF
import requests
import os


# 1. Carregar a tabela (como já tinhas)
tabela = pd.read_excel("Book.xlsx")

# 2. Configurações
TEMPLATE_PDF = "teste.pdf" # O teu PDF original com o design
OUTPUT_FOLDER = "PDFs_Preenchidos"

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def download_image(url, filename):
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, stream=True, headers=headers)
        if response.status_code == 200:
            with open(filename, 'wb') as f:
                f.write(response.content)
            return True
    except:
        return False

# 3. O Ciclo para criar os ficheiros
for index, row in tabela.iterrows():
    print(f"A preencher PDF para a casa {row['ID_Casa']}...")
    
    # Abrir o PDF Template original
    doc = fitz.open(TEMPLATE_PDF)
    pagina = doc[0] # Vamos editar a página 1 (índice 0)
    
    # --- INSERIR TEXTO ---
    # As coordenadas são (X, Y). 
    
    # Escrever a Morada
    pagina.insert_text((185, 185), f"{row['Morada']}", fontsize=12, color=(0, 0, 0))
    
    # Escrever o Preço
    pagina.insert_text((185, 212), f"{row['Preco']} €", fontsize=14, color=(1, 0, 0)) # Vermelho

    # --- INSERIR IMAGEM ---
    link = row['Link_Imagem_Casa']
    if pd.notna(link):
        nome_img_temp = f"temp_{row['ID_Casa']}.jpg"
        
        if download_image(link, nome_img_temp):
            # No PyMuPDF, a imagem precisa de uma "Caixa" (Rect)
            # Rect(x_inicio, y_inicio, x_fim, y_fim)
            rect_imagem = fitz.Rect(30, 450, 350, 650)
            
            pagina.insert_image(rect_imagem, filename=nome_img_temp)
            
            # Fechar e apagar imagem temporária
            os.remove(nome_img_temp)

    # 4. Guardar como um NOVO ficheiro
    doc.save(f"{OUTPUT_FOLDER}/Casa_{row['ID_Casa']}.pdf")
    doc.close() 

print("Feito!")

apg = input("Quer apagar os ficheiros gerados? (sim/não): ").strip().lower()

# loop para apagar os ficheiros (isto só serve para os testes burros)
for index, row in tabela.iterrows():

    
    caminho_ficheiro = f"{OUTPUT_FOLDER}/Casa_{row['ID_Casa']}.pdf"

   
    if apg in ['sim', 's', 'yes', 'y']:
        try:
            os.remove(caminho_ficheiro)
            print(f"Ficheiro {caminho_ficheiro} apagado.") 
        except OSError as e:
            print(f"Erro ao apagar {caminho_ficheiro}: {e}")
    else:
        print("És cool")