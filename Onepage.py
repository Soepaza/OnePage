# Criar uma onepage que será atualizada todo dia  e enviada para o gerente com os indicadores das vendas por e-mail.
# a onepage sera enviada também para diretoria com os indicadores anuais.

# Mandar o email com a onepage para os gerentes
# Salvar o backup do dia
# Mandar o email para diretoria

import pandas as pd
import time
import pathlib

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
dia_indicador = vendas['Data'].max()
# print(dia_indicador())

    # começar a criar o backup diario de quando fazer o resumo na onepage
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
# percorre a lista de arquivos ja existentes
arquivos_existentes = caminho_backup.iterdir()

    # criando lista que contem todos shoppings ja existentes
lista_shoppings = [shopping.name.strip() for shopping in arquivos_existentes]
# print(lista_shoppings)

    #criando a pasta do shopping com a lista de shoppings existes(ou nao existentes)
for loja in dicionario_lojas:
    loja_formatada = loja.strip()
    try:
        if loja not in lista_shoppings:
            nova_pasta = caminho_backup / loja_formatada
            nova_pasta.mkdir()

            #criar o arquivo (#Onepage) dentro da pasta do shopping.
        nome_Onepage = "{}_{}_{}.xlsx".format(dia_indicador.month, dia_indicador.day, loja_formatada)
        local_arquivo = caminho_backup / loja_formatada / nome_Onepage #"C:/Users/Home/../17_02_Loja.xlsx"
            #mandar para o framework do python (dicionario)
        dicionario_lojas[loja].to_excel(local_arquivo)
    except FileExistsError:
        print('Continuando...')

    #calcular indicador 1 (faturamento)
loja = 'Norte Shopping'
vendas_loja = dicionario_lojas[loja]
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador,:]
faturamento_ano = vendas_loja['Valor Final'].sum()
faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    #calcular diversidade de produtos
qtd_produtos_ano = len(vendas_loja['Produto'].unique())

qtd_produtos_ano_dia = len(vendas_loja_dia['Produto'].unique())

    #calcular ticket medio
valor_venda = vendas_loja.groupby['Código Venda'].sum()
ticket_medio_ano = valor_venda['Valor Final'].mean()

valor_venda_dia = vendas_loja_dia.groupby['Código Venda'].sum()
ticket_medio_dia = vendas_loja_dia['Valor Final'].mean()