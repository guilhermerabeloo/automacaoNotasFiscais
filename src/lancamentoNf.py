from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC

import time
import json
import pyautogui

with open("../config/config.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    acessoZeev = sensitive_data["acessoZeev"]

login = acessoZeev["login"]
senha = acessoZeev["senha"]

def lancarNf(infoPj, dataAtual, numeroNf):
    # definicao de variaveis
    cnpj = infoPj["cnpj"]
    razaoSocial = infoPj["razaoSocial"]
    valorNotaFiscal = infoPj["valorNotaFiscal"]
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
    wait = WebDriverWait(driver, 20)

    # fazendo login
    driver.get('https://fornecedora.zeev.it/my/user-change')
    driver.maximize_window()

    wait = WebDriverWait(driver, 20)
    iframeLogin = wait.until(EC.presence_of_element_located((By.ID, 'ifrContent')))
    driver.switch_to.frame(iframeLogin)

    driver.find_element(by='id', value='login').send_keys(login)
    driver.find_element(by='id', value='password').send_keys(senha)
    driver.find_element(by='id', value='btnLogin').click()
    driver.switch_to.default_content()

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
    driver.get('https://fornecedora.zeev.it/2.0/request?c=8VTuGOamI%2B81yJLC5J%2Bw45MFU%2B%2FYJpiOIyKUMJStE0ezH4FfVONPE5fqfWFyR554NOQMPtok1fqhoLPrLKLVuA%3D%3D')    


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

    driver.execute_script("document.getElementById('inpcontaGerencial').value = '21028 - GASTOS COM PJ FIXO'")

    form_estoque = driver.find_element(By.ID, 'inpestoque')
    select = Select(form_estoque)
    select.select_by_value('24')

    form_urgencia = driver.find_element(By.ID, 'inpurgencia')
    select = Select(form_urgencia)
    select.select_by_value('1 - Normal')

    form_tipo = driver.find_element(By.ID, 'inptipo')
    select = Select(form_tipo)
    select.select_by_value('Servico')
    time.sleep(1)

    driver.execute_script("document.getElementById('inpnfsDePj-1').checked = true")
    driver.execute_script("window.scrollBy(0, 500);")
    time.sleep(1)

    driver.find_element(By.ID, 'btnUploadnotaFiscal').click()
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'iframe')))

    time.sleep(2)
    pyautogui.click(661,460)
    time.sleep(5)

    pyautogui.click(86,178)
    time.sleep(1)
    pyautogui.click(300,171)
    time.sleep(1)
    pyautogui.click(790,507)
    time.sleep(3)

    # pyautogui.click(1184,606)
    # time.sleep(.5)
    pyautogui.click(1555,748)
    time.sleep(5)

    form_emissao = driver.find_element(By.ID, 'inpdataDeEmissao')
    form_emissao.send_keys(dataAtual)

    form_valor = driver.find_element(By.ID, 'inpvalorDaNotaFiscal')
    form_valor.send_keys(float(valorNotaFiscal) / 10)

    form_cnpj = driver.find_element(By.ID, 'inppesquisaDeFornecedor')
    form_cnpj.send_keys(cnpj)

    driver.execute_script(f'document.getElementById("inprazaoSocialDoEmitente").value = "{razaoSocial}"')

    driver.execute_script(f'document.getElementById("inpnumeroDaNf").value = "{numeroNf}"')

    driver.execute_script(f'document.getElementById("inpverificacaoDeDuplicidade").value = "Validação OK - Nota ainda não lançada"')

    driver.execute_script(f'document.getElementById("inpnumeroOk").value = "ok"')

    form_pedido = driver.find_element(By.ID, 'inppedido')
    form_pedido.send_keys(pedido)

    driver.execute_script(f'document.getElementById("inpvalidacaoDoPedido").value = "ok"')

    driver.execute_script(f'document.getElementById("inppedidosOk").value = "ok"')

    form_dataPagamento = driver.find_element(By.ID, 'inpvencimentoDaParcela')
    form_dataPagamento.send_keys('05/12/2023') # atencao aqui

    form_valorParcela = driver.find_element(By.ID, 'inpvalorDaParcela')
    form_valorParcela.send_keys(float(valorNotaFiscal) / 10)

    form_formaPagamento = driver.find_element(By.ID, 'inpformaDePagamento')
    select = Select(form_formaPagamento)
    select.select_by_value(formaPagamento)
    time.sleep(1)

    form_valorParcela = driver.find_element(By.ID, 'inpinstituicaoBancaria')
    form_valorParcela.send_keys(banco)

    form_valorParcela = driver.find_element(By.ID, 'inpagencia')
    form_valorParcela.send_keys(agencia)

    form_valorParcela = driver.find_element(By.ID, 'inpcontaBancaria')
    form_valorParcela.send_keys(conta)

    form_formaPagamento = driver.find_element(By.ID, 'inptipoDeConta')
    select = Select(form_formaPagamento)
    select.select_by_value(tipoDeConta)

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