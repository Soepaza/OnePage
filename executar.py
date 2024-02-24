import pandas as pd
import time
import pathlib
import win32com.client as win32

# ok
emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

vendas = vendas.merge(lojas, on='ID Loja')
# -----------
# criei dicionario na posicao da loja com todas as informacoes da venda daquela loja
dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

dia_indicador = vendas['Data'].max()

caminho_backup = pathlib.Path('Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]
# -----------criar pastas no 'Backup Arquivos Lojas'
for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    nome_arquivo = (f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx')
    local_arquivo = caminho_backup / loja / nome_arquivo

    dicionario_lojas[loja].to_excel(local_arquivo)
# -----------
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

loja = "Norte Shopping"
vendas_loja = dicionario_lojas[loja]
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

faturamento_ano = vendas_loja['Valor Final'].sum()
faturamento_dia = vendas_loja_dia['Valor Final'].sum()

qtd_produto_ano = len(vendas_loja['Produto'].unique())
qtd_produto_dia = len(vendas_loja_dia['Produto'].unique())

vendas_loja_sem_data = vendas_loja.drop(columns=['Data'])
vendas_loja_sem_data_dia = vendas_loja_dia.drop(columns=['Data'])

valor_venda = vendas_loja_sem_data.groupby('Código Venda').sum()
valor_venda_dia = vendas_loja_sem_data_dia.groupby('Código Venda').sum()

ticket_medio_ano = valor_venda['Valor Final'].mean()
ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
# ----------- checkpoint ok

outlook = win32.Dispatch('outlook.application')

nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
email = outlook.CreateItem(0)
email.To = emails.loc[emails['Loja'] == loja, 'Email'].values[0]
email.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'

attachment = pathlib.Path.cwd() / caminho_backup / loja / \
    f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
email.Attachments.Add(str(attachment))

email.Send()
