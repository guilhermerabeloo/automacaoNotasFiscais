import win32com.client as win32
import os

def envioDoEmail(enderecoEmail, nome, numeroDaNota, valorNota, emissaoNota, dataPagamento):
    pasta = 'C:\\Users\\guilherme.rabelo\\Downloads'

    arquivos = os.listdir(pasta)
    arquivos = [arquivo for arquivo in arquivos if os.path.isfile(os.path.join(pasta, arquivo))]
    ultimo_arquivo = max(arquivos, key=lambda arquivo: os.path.getmtime(os.path.join(pasta, arquivo)))
    caminhoNotaFiscal = f'{pasta}\\{ultimo_arquivo}'

    html_file_path = os.path.join(os.path.dirname(__file__), '../assets', 'confirmacaoLancamento.html')

    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    with open(html_file_path, 'r', encoding='utf-8') as file:
        html_body = file.read()
        html_body = html_body.format(nome=nome, numeroDaNota=numeroDaNota, valorNota=valorNota, emissaoNota=emissaoNota, dataPagamento=dataPagamento)

    email.To = enderecoEmail
    email.Subject = 'Nota fiscal de PJ lançada!'
    email.HTMLBody = html_body
    email.Attachments.Add(caminhoNotaFiscal)

    email.Send()
