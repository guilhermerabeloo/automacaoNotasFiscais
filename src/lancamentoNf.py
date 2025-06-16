from selenium import webdriver
from pywinauto.application import Application
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import time
import json
import pyautogui

with open("./config/config.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    acessoZeev = sensitive_data["acessoZeev"]

login = acessoZeev["login"]
senha = acessoZeev["senha"]

def lancarNf(infoPj, dataAtual, numeroNf, caminhoNotaFiscal):
    # definicao de variaveis
    cnpj = infoPj["cnpj"]
    razaoSocial = infoPj["razaoSocial"]
    valorNotaFiscal = infoPj["valorNotaFiscal_zeev"]
    email = infoPj["email"]
    pedido = infoPj["pedido"]
    formaPagamento = infoPj["formaPagamento"]
    banco = infoPj["banco"]
    agencia = infoPj["agencia"]
    conta = infoPj["conta"]
    tipoDeConta = infoPj["tipoDeConta"]

    # Configurando o navegador
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('executable_path=C:\\Users\\guilherme.rabelo\\Documents\\RPA\\chromedriver.exe')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 60)

    # fazendo login
    driver.get('https://fornecedora.zeev.it/my/user-change')
    driver.maximize_window()

    wait.until(EC.presence_of_element_located((By.ID, 'mobileTabContent')))
    time.sleep(2)
    driver.find_element(by='id', value='username').send_keys(login)
    driver.find_element(by='id', value='password').send_keys(senha)
    driver.find_element(by='id', value='btnZeevLogin').click()
    # personificando pessoa e entrando na tela de lançamento
    elementoShadowDOM = wait.until(EC.presence_of_element_located((By.ID, 'userSearch')))
    buscaPessoaPersonificada = driver.execute_script('return arguments[0].shadowRoot.querySelector("#txtSearchUser")', elementoShadowDOM)
    time.sleep(1)
    buscaPessoaPersonificada.send_keys(email)
    buscaPessoaPersonificada.send_keys(Keys.ENTER)
    time.sleep(1)
    pessoaPersonificada = driver.execute_script('return arguments[0].shadowRoot.querySelector("tr.select-user")', elementoShadowDOM)
    pessoaPersonificada.click()
    time.sleep(1)
    driver.get('https://fornecedora.zeev.it/2.0/request?c=%2BfNoYh44G5jE5xy8Wkn6jd3D%2BKdqh2rEE4ZN4Rk1HXQJFtVGr6nxd3yFR3Vnqqz7k00drhB3Gg7qm1Guji%2B55g%3D%3D')    


    # Preenchendo formulário
    wait.until(EC.presence_of_all_elements_located((By.ID, 'containerRequest')))
    time.sleep(1)
    driver.execute_script("document.getElementById('formCamposAuxiliares').style.display = 'block'")

    form_empresa = driver.find_element(By.ID, 'inpempresaDeLancamento')
    select = Select(form_empresa)
    select.select_by_value('5')
    time.sleep(2)

    form_unidade = driver.find_element(By.ID, 'inpunidade')
    select = Select(form_unidade)
    select.select_by_value('13')

    form_departamento = driver.find_element(By.ID, 'inpdepartamento')
    select = Select(form_departamento)
    select.select_by_value('18')

    pyautogui.press('TAB')
    time.sleep(.01)
    pyautogui.press('DOWN')
    time.sleep(.01)
    pyautogui.press('SPACE')
    time.sleep(.01)

    for _ in range(8):
        pyautogui.press('TAB')
        time.sleep(.01)
    time.sleep(1)

    driver.find_element(By.ID, 'btnUploadnotaFiscal').click() 
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'modal-iframe')))
    time.sleep(2)
    
    iframeModal = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'modal-iframe')))
    driver.switch_to.frame(iframeModal)
    time.sleep(1)

    driver.find_element(By.CSS_SELECTOR, 'h4.modal-title').click()
    time.sleep(1)

    pyautogui.press('TAB')
    time.sleep(1)
    pyautogui.press('SPACE')

    Application(backend="win32").connect(title="Abrir", found_index=0, timeout=10)
    time.sleep(2)

    pyautogui.write(caminhoNotaFiscal)
    time.sleep(3)
    pyautogui.press('ENTER')
    time.sleep(1)

    for _ in range(4):
        pyautogui.press('TAB')
        time.sleep(1)
    pyautogui.press('SPACE')
    time.sleep(1)
    pyautogui.press('ENTER')

    time.sleep(5)
    for _ in range(4):
        pyautogui.press('TAB')
        time.sleep(.2)

    pyautogui.press('SPACE')

    time.sleep(5)
    driver.switch_to.default_content()

    form_emissao = driver.find_element(By.ID, 'inpdataDeEmissao')
    form_emissao.send_keys(dataAtual)

    driver.execute_script(f'document.getElementById("inpvalorDaNotaFiscal").value = "{valorNotaFiscal}"')

    form_cnpj = driver.find_element(By.ID, 'inppesquisaDeFornecedor')
    form_cnpj.send_keys(cnpj)

    driver.execute_script(f'document.getElementById("inprazaoSocialDoEmitente").value = "{razaoSocial}"')

    driver.execute_script(f'document.getElementById("inpnumeroDaNf").value = "{numeroNf}"')

    driver.execute_script(f'document.getElementById("inpverificacaoDeDuplicidade").value = "Validação OK - Nota ainda não lançada"')

    driver.execute_script(f'document.getElementById("inpnumeroOk").value = "ok"')

    driver.execute_script(f'document.getElementById("inpparcelasOk").value = "ok"')

    driver.execute_script(f'document.getElementById("inpcodUnidade").value = "13"')

    form_pedido = driver.find_element(By.ID, 'inppedido')
    form_pedido.send_keys(pedido)

    driver.execute_script(f'document.getElementById("inpvalidacaoDoPedido").value = "ok"')

    driver.execute_script(f'document.getElementById("inppedidosOk").value = "ok"')

    driver.execute_script(f'document.getElementById("inpvalorDaParcela").value = "{valorNotaFiscal}"')

    form_formaPagamento = driver.find_element(By.ID, 'inpformaDePagamento')
    select = Select(form_formaPagamento)
    select.select_by_value(formaPagamento)
    time.sleep(1)

    form_banco = driver.find_element(By.ID, 'inpinstituicaoBancaria')
    form_banco.send_keys(banco)

    form_agencia = driver.find_element(By.ID, 'inpagencia')
    form_agencia.send_keys(agencia)

    form_conta = driver.find_element(By.ID, 'inpcontaBancaria')
    form_conta.send_keys(conta)

    form_formaPagamento = driver.find_element(By.ID, 'inptipoDeConta')
    select = Select(form_formaPagamento)
    select.select_by_value(tipoDeConta)
    
    driver.execute_script(f'document.getElementById("inpvalorRateio").value = "{valorNotaFiscal}"')
    
    driver.execute_script(f'document.getElementById("inptotalR").value = "{valorNotaFiscal}"')
    
    driver.execute_script(f'document.getElementById("inpvalorRateioUnidade").value = "{valorNotaFiscal}"')
    
    driver.execute_script(f'document.getElementById("inptotalRUnidade").value = "{valorNotaFiscal}"')
    

    funcao = driver.execute_script('return document.querySelector("#controllers select")')
    if funcao:
        select = Select(funcao)
        select.select_by_index(1)

    time.sleep(30)

    btnConcluir = driver.find_element(By.ID, 'BtnSend')
    btnConcluir.click()
    time.sleep(15)
    driver.quit()
    time.sleep(2)