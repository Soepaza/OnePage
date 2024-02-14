import pandas as pd
import time

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

vendas = vendas.merge(lojas, on='ID Loja')
#print(vendas)

        #Criando um arquivo separado para cada shopping existente na base de dados.
dicionario_das_lojas = {}
for arquivo_separado in lojas['Loja']:
    dicionario_das_lojas[arquivo_separado] = vendas.loc[vendas['Loja'] == arquivo_separado, :]
    
# print(dicionario_das_lojas['Shopping Morumbi'])


        #atualizar a data todo dia, para enviar o indicador diario no email.
def data_atualizada():
    dia_enviar_no_email = vendas['Data'].max()
    return dia_enviar_no_email

print(data_atualizada())
