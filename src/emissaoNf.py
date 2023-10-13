from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import pyautogui

def emitirNf(informacoes, dataAtual, mes, ano): 
    # definicao de variaveis
    loginNfse = informacoes["loginNfse"]
    senhaNFse = informacoes["senhaNFse"]
    cnpjTomador = informacoes["cnpjTomador"]
    buscaTributoNacional = informacoes["buscaTributoNacional"]
    coordenadaTributoMunicipal = informacoes["coordenadaTributoMunicipal"]
    valorNotaFiscal = informacoes["valorNotaFiscal"]
    descricaoServico = informacoes["descricaoServico"]


    # Configurando o navegador
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('executable_path=C:\\Users\\guilherme.rabelo\\Documents\\RPA\\chromedriver.exe')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # ----------------- Emissão da nota no portal -------------------

    # Entrando na plataforma de emissão
    driver.get('https://www.nfse.gov.br/EmissorNacional/Login')
    driver.maximize_window()

    # Fazendo login 
    login = wait.until(EC.presence_of_element_located((By.ID, 'Inscricao')))
    login.send_keys(loginNfse)
    senha = wait.until(EC.presence_of_element_located((By.ID, 'Senha')))
    senha.send_keys(senhaNFse)
    driver.find_element(by='css selector', value='button.btn-primary').click()
    time.sleep(1)

    # Emissao da nota - Etapa 1
    btnNovaNf = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'btnAcesso')))
    btnNovaNf.click()

    competencia = wait.until(EC.presence_of_element_located((By.ID, 'DataCompetencia')))
    competencia.send_keys(dataAtual)
    competencia.send_keys(Keys.TAB)
    time.sleep(1)
    wait.until(EC.presence_of_element_located((By.ID, 'btnAvancar')))

    driver.execute_script("document.querySelectorAll('#Tomador_LocalDomicilio')[1].checked = true")
    driver.execute_script("document.getElementsByClassName('retratil')[1].style.display = ''")
    driver.execute_script("document.getElementById('pnlInscricaoBrasil').style.display = ''")
    time.sleep(1)

    tomador = driver.find_element(by='id', value='Tomador_Inscricao')
    tomador.send_keys(cnpjTomador)
    tomador.send_keys(Keys.TAB)
    time.sleep(1)
    driver.find_element(by='id', value='btnAvancar').click()
    time.sleep(2)

    # Emissao da nota - Etapa 2
    pyautogui.click(898,459)
    time.sleep(1)
    pyautogui.write('Fortaleza')
    time.sleep(1)
    pyautogui.click(747,702)
    time.sleep(1)
    pyautogui.click(194,639)
    time.sleep(1)
    pyautogui.write(buscaTributoNacional)
    time.sleep(1)
    pyautogui.click(257,604)

    driver.execute_script("document.getElementById('ServicoPrestado_HaExportacaoImunidadeNaoIncidencia').checked = true")
    driver.execute_script("document.getElementById('pnlMunIncid').style.display = ''")
    driver.execute_script("window.scrollBy(0, 500);")

    time.sleep(1)
    pyautogui.click(258,442)
    time.sleep(1)
    # pyautogui.click(358,569)
    pyautogui.click(coordenadaTributoMunicipal[0],coordenadaTributoMunicipal[1])
    time.sleep(1)

    driver.find_element(by='id', value='ServicoPrestado_Descricao').send_keys(f'{descricaoServico} {mes}/{ano}')
    time.sleep(1)
    driver.find_element(by='css selector', value='button.btn-primary').click()

    # Emissao da nota - Etapa 3
    campoValor = wait.until(EC.presence_of_element_located((By.ID, 'Valores_ValorServico')))
    campoValor.send_keys(valorNotaFiscal)
    campoValor.send_keys(Keys.TAB)
    time.sleep(1)

    driver.execute_script("document.querySelectorAll('#ValorTributos_TipoValorTributos')[2].checked = true")
    driver.find_element(by='css selector', value='button.btn-primary').click()

    # Emissão da nota - Etapa 4

    time.sleep(3)
    driver.quit()
    time.sleep(3)