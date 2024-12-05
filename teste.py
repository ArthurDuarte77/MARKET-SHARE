from PIL import Image

def calcular_porcentagem_preto(caminho_imagem):
    # Abre a imagem e converte para escala de cinza
    imagem = Image.open(caminho_imagem).convert('L')
    
    # Converte a imagem em uma lista de pixels
    pixels = list(imagem.getdata())
    total_pixels = len(pixels)

    # Define um limite para considerar um pixel como preto
    # 0 é preto absoluto, aumentando o limite se considerar tons próximos ao preto
    limiar_preto = 10

    # Conta quantos pixels são pretos (ou próximos do preto)
    pixels_pretos = sum(1 for pixel in pixels if pixel <= limiar_preto)

    # Calcula a porcentagem de pixels pretos
    porcentagem_preto = (pixels_pretos / total_pixels) * 100

    return porcentagem_preto

# Exemplo de uso
caminho_imagem = "branco.jpg"
porcentagem = calcular_porcentagem_preto(caminho_imagem)
print(f"A imagem contém {porcentagem:.2f}% de pixels pretos.")