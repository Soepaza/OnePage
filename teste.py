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
dia_indicador = vendas['Loja'].min()
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

shoppings_existentes_napasta = caminho_backup.iterdir()   #iterdir retorna a lista de shoppings existentes
#print(shoppings_existentes_napasta)

lista_shoppings = []
for shopping in shoppings_existentes_napasta:
    lista_shoppings.append(shopping.name)
#print(lista_shoppings)

for shopping in dicionario_das_lojas:
    shopping_nome = shopping.strip()
    nova_pasta = caminho_backup / shopping_nome
    if shopping not in lista_shoppings and not nova_pasta.exists():
        nova_pasta.mkdir()

    #Cria o local do arquivo para salvar o backup diario
    #local_arquivo_salvo = "C:\Users\Home\...16_02_NomeLoja.xlsx"


    nome_shopping = "{}_{}_{}.xlsx".format(data_atualizada_indicador().month, data_atualizada_indicador().day, shopping)
    local_arquivo_salvo = caminho_backup / shopping / nome_shopping

    dicionario_das_lojas[shopping].to_excel(local_arquivo_salvo, index=False)


