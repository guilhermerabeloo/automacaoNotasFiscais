from lancamentoNf import lancarNf
from emissaoNf import emitirNf
from sendEmail import envioDoEmail
import json
import datetime
import time

import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('../log/logsDeExecucao.log'), 
    ]
)

dataAtual = datetime.datetime.now().strftime('%d/%m/%Y')
mes = datetime.datetime.now().strftime('%m')
ano = datetime.datetime.now().strftime('%Y')

with open("../config/secrets.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    sensitive_data = sensitive_data["pjs"]

with open("../config/calendarioPagamento.json", "r", encoding="utf-8") as file:
    calendario = json.load(file)
    mesFormatado = f'{int(mes)}'
    dataPagamento = calendario[mesFormatado]

for pj in sensitive_data:
    nome = list(pj.keys())[0]
    informacoes = pj[nome]

    logging.info(f'Iniciando emissao da nota de {nome}')
    numeroNf = emitirNf(informacoes, dataAtual, mes, ano)
    logging.info(f'Finalizando emissao da nota de {nome}')

    logging.info(f'Iniciando lancamento de {nome}')
    lancarNf(informacoes, dataAtual, numeroNf)
    logging.info(f'Finalizando lancamento de {nome}')

    envioDoEmail(informacoes['email'], nome, numeroNf, informacoes['valorNotaFiscal'], dataAtual, dataPagamento)

    time.sleep(5)
logging.info(f'Fim da automacao')
