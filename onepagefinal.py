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

for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    qtd_produto_ano = len(vendas_loja['Produto'].unique())
    qtd_produto_dia = len(vendas_loja_dia['Produto'].unique())

    valor_venda = vendas_loja.groupby('Código Venda')['Valor Final'].sum()
    valor_venda_dia = vendas_loja_dia.groupby(
        'Código Venda')['Valor Final'].sum()

    ticket_medio_ano = valor_venda.mean()
    ticket_medio_dia = valor_venda_dia.mean()

# ----------- checkpoint ok

    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    email = outlook.CreateItem(0)
    email.To = emails.loc[emails['Loja'] == loja, 'Email'].values[0]
    email.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja} By Soepaza'

    def calcular_resultado(indicador_dia, meta_dia, indicador_ano, meta_ano):
        result_dia = '\U00002705' if indicador_dia >= meta_dia else '\U0000274C'
        result_ano = '\U00002705' if indicador_ano >= meta_ano else '\U0000274C'
        return result_dia, result_ano

    # Chamar a função para calcular os resultados
    result_faturamento_dia, result_faturamento_ano = calcular_resultado(
        faturamento_dia, meta_faturamento_dia, faturamento_ano, meta_faturamento_ano)
    result_diversidade_dia, result_diversidade_ano = calcular_resultado(
        qtd_produto_dia, meta_qtdeprodutos_dia, qtd_produto_ano, meta_qtdeprodutos_ano)
    result_ticket_dia, result_ticket_ano = calcular_resultado(
        ticket_medio_dia, meta_ticketmedio_dia, ticket_medio_ano, meta_ticketmedio_ano)

    email.HTMLBody = f'''
    <p>Bom dia <strong>{nome}</strong></p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da loja <strong>{loja}</strong> foi:</p>

    <table>
        <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
        </tr>
        <tr>
            <td style="text-align: center">Faturamento</td>
            <td style="text-align: center">R${faturamento_dia:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
            <td style="text-align: center">{result_faturamento_dia}</td>
        </tr>
        <tr>
            <td style="text-align: center">Diversidade de Produtos</td>
            <td style="text-align: center">{qtd_produto_dia}</td>
            <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
            <td style="text-align: center">{result_diversidade_dia}</td>
        </tr>
        <tr>
            <td style="text-align: center">Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
            <td style="text-align: center">{result_ticket_dia}</td>
        </tr>
        <br>
            <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
        </tr>
        <tr>
            <td style="text-align: center">Faturamento</td>
            <td style="text-align: center">R${faturamento_ano:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
            <td style="text-align: center">{result_faturamento_ano}</td>
        </tr>
        <tr>
            <td style="text-align: center">Diversidade de Produtos</td>
            <td style="text-align: center">{qtd_produto_ano}</td>
            <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
            <td style="text-align: center">{result_diversidade_ano}</td>
        </tr>
        <tr>
            <td style="text-align: center">Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
            <td style="text-align: center">{result_ticket_ano}</td>
        </tr>
        </br>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Soe</p>
    '''

    attachment = pathlib.Path.cwd() / caminho_backup / loja / \
        (f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx')
    email.Attachments.Add(str(attachment))

    email.Send()
# ----------- checkpoint ok
    faturamento_loja = vendas.groupby('Loja')['Valor Final'].sum()
    faturamento_loja_ano = faturamento_loja.sort_values(ascending=False)

    nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(
        dia_indicador.month, dia_indicador.day)
    faturamento_loja_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

    vendas_dia = vendas.loc[vendas['Data'] == dia_indicador, :]
    faturamento_loja_dia = vendas_dia.groupby('Loja')['Valor Final'].sum()
    faturamento_loja_dia = faturamento_loja_dia.sort_values(ascending=False)

    nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(
        dia_indicador.month, dia_indicador.day)
    faturamento_loja_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = emails.loc[emails['Loja'] == 'Diretoria', 'Email'].values[0]
    email.Subject = f'Ranking Dia e Ano {dia_indicador.day}/{dia_indicador.month}'
    email.Body = f'''
    Prezados, bom dia.

    Melhor loja do dia em Faturamento foi {faturamento_loja_dia.idxmax()} com o faturamento de: R${faturamento_loja_dia.max():.2f}
    Pior loja do dia em Faturamento foi {faturamento_loja_dia.idxmin()} com o faturamento de: R${faturamento_loja_dia.min():.2f}

    Melhor loja do ano em Faturamento foi {faturamento_loja_ano.idxmax()} com o faturamento de: R${faturamento_loja_ano.max():.2f}
    Pior loja do ano em Faturamento foi {faturamento_loja_ano.idxmin()} com o faturamento de: R${faturamento_loja_ano.min():.2f}

    Segue em anexo a atualização do ranking do ano e do dia de todas as lojas.

    Qualquer dúvida estou a disposição.
    Att.,
    Soe

    '''

    attachment = pathlib.Path.cwd() / caminho_backup / \
        f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
    email.Attachments.Add(str(attachment))

    attachment = pathlib.Path.cwd() / caminho_backup / \
        f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
    email.Attachments.Add(str(attachment))

    email.Send()
    print('E-mail da Loja {} enviado'.format(loja))