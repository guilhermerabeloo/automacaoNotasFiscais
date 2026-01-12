import pyautogui
import re
import pdfplumber


def clicarEmImagem(imagem, i):
    ocorrencias = list(pyautogui.locateAllOnScreen(imagem))
    posicao = ocorrencias[i]
    left, top, width, height = posicao
    centro_x = left + width / 2
    centro_y = top + height / 2
    pyautogui.click(centro_x, centro_y)
    
def extrai_texto_de_pdf(path: str) -> list:
    textos = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                textos.append(t)
                
    return textos[0].split('\n')

def extrai_numero_nf(textos_da_nf: list) -> int:
    if any('SecretariaMunicipaldasFinanças' in texto for texto in textos_da_nf): # Verifica se a Nota Fiscal é de MEI
        try:
            indiceComNumeroNf = [i for i, item in enumerate(textos_da_nf) if 'NúmerodaNFS-e' in item][0] + 1 # Com base na posicao da tag, encontra o indice da lista onde tem o numero da nota 
            linhaComNumeroNf = textos_da_nf[indiceComNumeroNf]
            numeroNf = linhaComNumeroNf.split(' ')[0] # Extrai o numero da nota

            return numeroNf
        except Exception as err:
            raise Exception('Não foi possível extrair número da nota', err)
        
    elif any('SECRETARIA MUNICIPAL DAS FINANÇAS NFS-e' in texto for texto in textos_da_nf): # Verifica se a Nota Fiscal é de ME
        try:
            linhaComNumeroNf = [linha for linha in textos_da_nf if 'NOTA FISCAL ELETRÔNICA DE SERVIÇO - NFS-e' in linha][0] # Encontra o item na lista que contem o numero da nota fiscal
            numeroNf = re.search(r'\d+', linhaComNumeroNf).group() # Extrai apenas numeros para obter o numero da nota fiscal
            
            return numeroNf # Extrai o numero da nota
        except Exception as err:
            raise Exception('Não foi possível extrair número da nota', err)
        
    else: 
        raise Exception('PDF não é uma nota ou está em um padrão desconhecido')
    
    