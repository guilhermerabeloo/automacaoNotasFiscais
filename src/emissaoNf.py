from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from src.util import clicarEmImagem  
from pywinauto.application import Application
import json
import time
import pyautogui

def emitirNf(informacoes, dataAtual, mes, ano): 
    # definicao de variaveis
    loginNfse = informacoes["loginNfse"]
    senhaNFse = informacoes["senhaNFse"]
    razaoSocialTomador = informacoes["razaoSocial"]
    cnpjTomador = informacoes["cnpjTomador"]
    buscaTributoNacional = informacoes["buscaTributoNacional"]
    buscaTributoMunicipal = informacoes["buscaTributoMunicipal"]
    coordenadaTributoMunicipal = informacoes["coordenadaTributoMunicipal"]
    valorNotaFiscal = informacoes["valorNotaFiscal"]
    descricaoServico = informacoes["descricaoServico"]
    nomeArquivo = informacoes["arquivo"]

    with open('./config/config.json', 'r', encoding='utf-8') as file:
        config = json.load(file)
        localNotasFiscais = config['localNotasFiscais']

    # Configurando o navegador
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('executable_path=C:\\Users\\guilherme.rabelo\\Documents\\RPA\\chromedriver.exe')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--incognito')
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

    ## Obtendo numero da ultima nota
    driver.execute_script("""const quantidadeTabelas = document.querySelectorAll('table.table-striped').length
                             document.querySelectorAll('table.table-striped')[quantidadeTabelas-1].querySelector('tbody tr .td-opcoes a').click()
                        """)

    btnNovaNf = wait.until(EC.presence_of_element_located((By.ID, 'btnNovaNFSe')))
    numeroNota = int(driver.execute_script("return document.querySelectorAll('.form-group span.form-control-static')[3].innerText")) + 1

    if razaoSocialTomador == 'REGIO LOPES DE OLIVEIRA FILHO 07203962393':
        numeroNota = numeroNota-1
    elif razaoSocialTomador == '49.050.939 LEVY EMANUEL SANTIAGO PINTO':
        numeroNota = numeroNota-2

    btnNovaNf.click()
    # Emissao da nota - Etapa 1
    print('etapa 1 - Informações do tomador')

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
    time.sleep(3)

    # Emissao da nota - Etapa 2
    print('etapa 2 - Informações do serviço')
    clicarEmImagem('C:\\Users\\guilherme.rabelo\\Documents\\RPA_Python\\RPA_LancamentoNotasPJ\\assets\\btn_selectMunicipio.png', 1)
    time.sleep(1)
    pyautogui.write('Fortaleza')
    time.sleep(1)
    for _ in range(4):
        time.sleep(1)
        pyautogui.press('down')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    pyautogui.press('TAB')
    pyautogui.press('enter')
    time.sleep(10)
    pyautogui.write(buscaTributoNacional)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    driver.execute_script("document.getElementById('ServicoPrestado_HaExportacaoImunidadeNaoIncidencia').checked = true")
    driver.execute_script("document.getElementById('pnlMunIncid').style.display = ''")
    driver.execute_script("window.scrollBy(0, 500);")

    for _ in range(4):
        pyautogui.press('TAB')
        time.sleep(.5)
    time.sleep(1)

    if buscaTributoMunicipal=='print': 
        btnSelecionaCodigo = "C:\\Users\\guilherme.rabelo\\Documents\\RPA_Python\\RPA_LancamentoNotasPJ\\assets\\btn_selectCodigoMunicipio.png"
        clicarEmImagem(btnSelecionaCodigo,0)
        time.sleep(3)

        clicarEmImagem(coordenadaTributoMunicipal,0)
    else:
        pyautogui.press('enter')
        pyautogui.write(buscaTributoMunicipal)

    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(5)
    driver.execute_script(f"document.getElementById('ServicoPrestado_Descricao').value = '{descricaoServico} {mes}/{ano}'")
    time.sleep(1)
    driver.find_element(by='css selector', value='button.btn-primary').click()

    time.sleep(1)

    # Emissao da nota - Etapa 3
    print('etapa 3 - Valores')

    campoValor = wait.until(EC.presence_of_element_located((By.ID, 'Valores_ValorServico')))
    campoValor.send_keys(valorNotaFiscal)
    campoValor.send_keys(Keys.TAB)
    time.sleep(1)

    driver.execute_script("document.querySelectorAll('#ValorTributos_TipoValorTributos')[2].checked = true")
    driver.find_element(by='css selector', value='button.btn-primary').click()

    time.sleep(10)
    # Emissão da nota - Etapa 4
    print('etapa 4 - Conferência e emissão')
    time.sleep(30)
    btnEmitir = wait.until(EC.presence_of_element_located((By.ID, 'btnProsseguir')))
    btnEmitir.click()

    # Download da nota
    time.sleep(10)

    btnDownload = wait.until(EC.presence_of_element_located((By.ID, 'btnDownloadDANFSE')))
    btnDownload.click()

    Application(backend="win32").connect(title=f'Salvar como', timeout=60)
    time.sleep(3)

    pyautogui.write(f'{localNotasFiscais}\\{nomeArquivo}.pdf')
    time.sleep(1)
    pyautogui.hotkey('ALT', 'L')

    Application(backend="win32").connect(title=f'Confirmar Salvar como', timeout=60)
    time.sleep(.5)
    pyautogui.hotkey('ALT', 'S')

    time.sleep(3)
    driver.quit()
    print('Download realizado')
    time.sleep(3)
    return numeroNota