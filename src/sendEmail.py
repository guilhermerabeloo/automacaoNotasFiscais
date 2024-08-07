import win32com.client as win32

def envioDoEmail(enderecoEmail, nome, numeroDaNota, valorNota, emissaoNota, dataPagamento, caminhoNotaFiscal):
    html_file_path = 'C:\\Users\\guilherme.rabelo\\Documents\\RPA_Python\\RPA_LancamentoNotasPJ\\assets\\confirmacaoLancamento.html'

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
