import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep  

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Ola {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. favor pagar no link http://www.pagar.com'

    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    webbrowser.opem(link_mensagem_whatsapp)