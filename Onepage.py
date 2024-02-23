# Criar uma onepage que será atualizada todo dia  e enviada para o gerente com os indicadores das vendas por e-mail.
# a onepage sera enviada também para diretoria com os indicadores anuais.

# Mandar o email com a onepage para os gerentes
# Salvar o backup do dia
# Mandar o email para diretoria

import pandas as pd
import time
import pathlib
import win32com.client as win32

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

# criando a pasta do shopping com a lista de shoppings existes(ou nao existentes)
for loja in dicionario_lojas:
    loja_formatada = loja.strip()
    try:
        if loja not in lista_shoppings:
            nova_pasta = caminho_backup / loja_formatada
            nova_pasta.mkdir()

            # criar o arquivo (#Onepage) dentro da pasta do shopping.
        nome_Onepage = "{}_{}_{}.xlsx".format(dia_indicador.month, dia_indicador.day, loja_formatada)
        local_arquivo = caminho_backup / loja_formatada / nome_Onepage  # "C:/Users/Home/../17_02_Loja.xlsx"
        # mandar para o framework do python (dicionario)
        dicionario_lojas[loja].to_excel(local_arquivo)
    except FileExistsError:
        print(f'Pasta para a Loja {loja_formatada} já existe. Continuando...')

    # calcular indicador 1 (faturamento)
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

    # calcular diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # Calcular ticket médio
    valor_venda = vendas_loja.groupby('Código Venda')['Valor Final'].sum()
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda')['Valor Final'].sum()

    ticket_medio_ano = valor_venda.mean()
    ticket_medio_dia = valor_venda_dia.mean()

    # Criar email para enviar os relatórios
    # criando instancia com outlook
    outlook = win32.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    if email is None:
        print("Erro ao criar o objeto de e-mail. Verifique sua configuração do Outlook.")
    # Adicione um possível código de saída ou solução aqui

    if not emails.loc[emails['Loja'] == loja].empty:
        nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
        email_to = emails.loc[emails['Loja'] == loja, 'E-mail'].values
        if email_to and email_to[0]:
            email.To = email_to[0]
            email.Subject = f'OnePage dia: {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
        
        # Restante do código para configuração do e-mail

        email.Send()
        print(f'Email da loja {loja} enviado.')
    else:
        print(f'E-mail da loja {loja} está vazio. Pulando para a próxima loja.')
else:
    print(f'Loja {loja} não encontrada no arquivo de e-mails. Pulando para a próxima loja.')

    if faturamento_dia >= meta_faturamento_dia and qtde_produtos_dia >= meta_qtdeprodutos_dia and ticket_medio_dia >= meta_ticketmedio_dia:
        cor_fat_dia = 'green'
        cor_qtd_dia = 'green'
        cor_ticket_dia = 'green'
    else:
        cor_fat_dia = 'red'
        cor_qtd_dia = 'red'
        cor_ticket_dia = 'red'
    try:

        email.HTMLBody = f'''
        <p>Bom dia, <strong>{nome}</strong> </p>

        <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>loja {loja}</strong> foi: </p>

        <table>
        <tr>
            <th>Indicador</th>
            <th>Valor dia</th>
            <th>Meta dia</th>
            <th>Cenário dia</th>
        </tr>
        <tr>
            <td>Faturamento</td>
            <td style="text-align:center">{faturamento_dia}</td>
            <td style="text-align:center">{meta_faturamento_dia}</td>
            <td style="text-align:center"><font color={cor_fat_dia}>◙</font></td> 
        </tr>
        <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align:center">{qtde_produtos_dia}</td>
            <td style="text-align:center">{meta_qtdeprodutos_dia}</td>
            <td style="text-align:center"><font color={cor_qtd_dia}>◙</font></td> 
        </tr>
        <tr>
            <td>Ticket Médio</td>
            <td style="text-align:center">{ticket_medio_dia}</td>
            <td style="text-align:center">{meta_ticketmedio_dia}</td>
            <td style="text-align:center"><font color={cor_ticket_dia}>◙</font></td> 
        </tr>
        
        </table>


        <p> Segue em anexo a planilha com todos os dados para mais detalhes. </p>
        <p> Qualquer dúvida estou a disposição.</p>
        <p> Att., Soe</p>
        '''
        attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja_formatada}.xlsx'
        email.Attachments.Add(str(attachment))
        email.Send()
        print(f'Email da loja {loja} enviado.')
    except Exception as e:
        print(f'Erro ao enviar e-mail para a loja {loja}: {str(e)}')
