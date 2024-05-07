import pyautogui

def clicarEmImagem(imagem):
    posicao = pyautogui.locateOnScreen(imagem)
    left, top, width, height = posicao
    centro_x = left + width / 2
    centro_y = top + height / 2
    pyautogui.click(centro_x, centro_y)