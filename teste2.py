import pandas as pd
import time
import pathlib

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

vendas = vendas.merge(lojas, on='ID Loja')

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]

dia_indicador = vendas['Data'].max()

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name.strip() for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    loja_formatada = loja.strip()

    if loja_formatada not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja_formatada
        nova_pasta.mkdir()

    # Salvar dentro da pasta
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja_formatada)
    local_arquivo = caminho_backup / loja_formatada / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)