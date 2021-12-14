#importa a base de dados
#visualizar base de dados

#utlizado panda biblioteca #instalar panda->pip install pandas e open pip install openpyxl

#faturamento por loja

#quantidade de produtos vendidos por loja

#tcket medio por produto em cada loja(faturamento dividido por quantidade de produto vendido por loja

#Envio do email com relatorio

import pandas as pd
#ler o arquivo xml


tabela_vendas = pd.read_excel('vendas.xlsx')

print(tabela_vendas)





