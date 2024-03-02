#Sessão de imports
import os
import pandas as pd
import win32com.client as win32
from datetime import datetime

#Leitura dos arquivos a serem consolidados.
caminho = "bases"
arquivos = os.listdir(caminho)
tabela_consolidada = pd.DataFrame()

for arquivo in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho,arquivo))
    tabela_vendas['Data de Venda'] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas['Data de Venda'], unit='d')
    tabela_consolidada = pd.concat([tabela_consolidada,tabela_vendas])

#ajuste e criação do arquivo consolidado
tabela_consolidada = tabela_consolidada.sort_values(by='Data de Venda')
tabela_consolidada = tabela_consolidada.reset_index(drop=True)
tabela_consolidada.to_excel('Vendas.xlsx', index=False)

#Processo de envio de email com arquivo consolidado
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.to = 'teste@gmail.com'
dt_atual = datetime.today().strftime("%d/%m/%y")
email.Subject = f'Relatório de Vendas {dt_atual}'
email.Body = f"""
Prezados,
Segue em anexo o Relatório de Vendas de {dt_atual} atualizado.
Atenciosamente
Rodrigo Reis
"""
email.Attachments.Add(os.path.join(os.getcwd(), 'Vendas.xlsx'))
email.Send()
