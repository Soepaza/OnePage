import pandas as pd
import time
import pathlib

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

vendas = vendas.merge(lojas, on='ID Loja')
# print(vendas)

# Criando um arquivo separado para cada shopping existente na base de dados.
dicionario_das_lojas = {}
for arquivo_separado in lojas['Loja']:
    dicionario_das_lojas[arquivo_separado] = vendas.loc[vendas['Loja']
                                                        == arquivo_separado, :]

# print(dicionario_das_lojas['Shopping Morumbi'])

    # atualizar a data todo dia, para enviar o indicador diario no email.


def data_atualizada_indicador():
    dia_enviar_no_email = vendas['Data'].max()
    return dia_enviar_no_email


#print(data_atualizada_indicador())

    # criar a pasta que vai receber os backups das lojas dentro do programa.
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

shoppings_existentes_napasta = caminho_backup.iterdir()   #iterdir retorna a lista de shoppings existentes
#print(shoppings_existentes_napasta)

lista_shoppings = []
for shopping in shoppings_existentes_napasta:
    lista_shoppings.append(shopping.name)
#print(lista_shoppings)

for shopping in dicionario_das_lojas:
    if shopping not in shoppings_existentes_napasta:
        nova_pasta = caminho_backup / shopping
        nova_pasta.mkdir()
