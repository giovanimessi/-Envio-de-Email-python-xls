#importa a base de dados
#visualizar base de dados

#pip install  pywin32
#envio de email

#utlizado panda biblioteca #instalar panda->pip install pandas e open pip install openpyxl

#faturamento por loja

#quantidade de produtos vendidos por loja

#tcket medio por produto em cada loja(faturamento dividido por quantidade de produto vendido por loja

#Envio do email com relatorio

#filtrar por coluna = tabela_venda[['ID Loja','Valor Final']]
#tabela_venda.grouby('ID Loja').sum()
#tabelas_vendas[['ID Loja','Valor Final']].grouby('ID Loja').sum()


import pandas as pd
#ler o arquivo xml
import win32com.client as win32


tabela_vendas = pd.read_excel('vendas.xlsx')


pd.set_option('display.max_columns', None)

faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
#quantidade de produtos vendidos por loja
print(faturamento)
print('-' * 50)
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()


print(quantidade)

print('-' * 50)
#ticket medio
ticket =  (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

print(ticket)


# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'pythonimpressionador@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Lira</p>
'''

mail.Send()

print('Email Enviado')










