# ### Passo 1 - Importar Arquivos e Bibliotecas

import pandas as pd
import win32com.client as win32
import pathlib

#importar base de dados
emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')
display(emails)
display(lojas)
display(vendas)

# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

#incluir o nome da loja em vendas
vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)

#criar uma tabela para cada uma das lojas
#loc para filtrar e a variavel loja é para cada loja 
#loc(linha, coluna) como queria todas colunas :
dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
display(dicionario_lojas['Rio Mar Recife'])
display(dicionario_lojas['Shopping Vila Velha']) 

#calcular o indicador do dia, o dia mais recente, o ultimo
dia_indicador = vendas['Data'].max()
print(dia_indicador)
print('{}/{}/{}'.format(dia_indicador.day, dia_indicador.month, dia_indicador.year))
#print(dia_indicador.day)

# ### Passo 3 - Salvar a planilha na pasta de backup

#identificar se a pasta já existe

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    #salvar dentro da pasta
    nome_arquivo = '{}_{}_{}_{}.xlsx'.format(dia_indicador.day, dia_indicador.month, dia_indicador.year, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)

# ### Passo 4 - Calcular o indicador para 1 loja

#definir metas 
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_ano = 500 
meta_ticketmedio_dia = 500 

#vou me basear por uma loja e depois criar um for pra fazer com todas
#a variavel vendas_loja_dia recebe a coluna de data .. loc para localizar(linha, coluna) como queria toda coluna :
for loja in dicionario_lojas:
    #loja = 'Norte Shopping'
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print('O faturamento do ano foi {:,}'.format(faturamento_ano))
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print('O faturamento do dia foi {:,}'.format(faturamento_dia))

    #diversidade de produtos
    #ver quantos produtos diferentes temos na coluna produto e para remover as duplicatas e como é uma coluna só, vamos usar o metodo.unique()
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    #print('Quantidade de produtos por ano {}'.format(qtde_produtos_ano))
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    #print('Quantidade de produto no dia {}'.format(qtde_produtos_dia))
    #ticket medio
    #agrupar todos valores da coluna de Código Venda
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    #display(valor_venda)
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    #print(ticket_medio_ano)
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    #print(ticket_medio_dia)
    #enviar o e-mail
    #metodo values[0] para pegar apenas um indice, assim para o email e nome
    outlook = win32.Dispatch('outlook.application')
    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = 'OnePage Data {}/{}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month, dia_indicador.year, loja)
    #mail.Body = 'Texto do E-mail'
    #ao invés de usar o método format, posso usar uma F-string.. OBS: para formatar uma variavel, basta colocar do lado (loja:.2)
    #mail.Subject = f'OnePage Data {dia_indicador.day}/{dia_indicador.month}/{dia_indicador.year} - Loja {loja}') 
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano>=meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano='red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia ='red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLbody = f'''
    <p>Bom dia {nome} </p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month}/{dia_indicador.year})</strong> da loja <strong>{loja}</strong> foi:</p>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia}</td>
        <td style="text-align: center">R${meta_faturamento_dia}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr> 
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano}</td>
        <td style="text-align: center">R${meta_faturamento_ano}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>  
    </table>
    <p>Se em anexo a planilha com todos os dados para mais detalhes.</p>
    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Schwanke</p>
    '''
    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.day}_{dia_indicador.month}_{dia_indicador.year}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))
    mail.Send()
    print('E-mail da loja {} enviado'.format(loja))

# ### Passo 5 - Criar o Ranking

#criar o ranking é listar todas as lojas por ordem de faturamento
#para isso vou agrupar a coluna de loja do arquivo vendas
#para ordenar vou usar o método.sort_values(by=coluna) e ascending=False é para deixar do maior valor para o menor, como se trata de um ranking
faturamento_lojas_ano = vendas.groupby('Loja')[['Loja','Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas_ano.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas_ano)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.day, dia_indicador.month)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja','Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
display(faturamento_loja_dia)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.day, dia_indicador.month)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

# ### Passo 6 - Enviar e-mail para diretoria

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}' 
mail.Body = f'''
Prezados, bom dia!

Melhor loja do dia em faturamento: Loja {faturamento_lojas_dia.index[0]} com faturamento de R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do dia em faturamento: Loja {faturamento_lojas_dia.index[-1]} com o faturamento de R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do ano em faturamento: Loja {faturamento_lojas_ano.index[0]} com o faturamento de R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do ano em faturamento: Loja {faturamento_lojas_ano.index[-1]} com o faturamento de R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo todos os rankins do dia e do ano de todas as lojas.
Qualquer dúvida estou à disposição.
Att.,
Schwanke
 

'''
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.day}_{dia_indicador.month}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.day}_{dia_indicador.month}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))
mail.Send()
print('E-mail da Diretoria enviado')

