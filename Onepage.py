# Criar uma onepage que será atualizada todo dia  e enviada para o gerente com os indicadores das vendas por e-mail.
# a onepage sera enviada também para diretoria com os indicadores anuais.

# Mandar o email com a onepage para os gerentes
# Salvar o backup do dia
# Mandar o email para diretoria

import pandas as pd
import time

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

    # incluir nome da loja em vendas
vendas = vendas.merge(lojas, on='ID Loja')
# print(vendas)

    # Criando uma tabela para cada loja, colocando dentro de um dicionario usando o loc.
dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]
# print(dicionario_lojas['Shopping Recife'])

    # Definir dia do indicador


def dia_indicador():
    dia_indicador = vendas['Data'].max()
    dia_formatado = dia_indicador.strftime('%d/%m')
    return dia_formatado
# print(dia_indicador())

    #começar a criar o backup diaio de quando fazer o resumo na onepage