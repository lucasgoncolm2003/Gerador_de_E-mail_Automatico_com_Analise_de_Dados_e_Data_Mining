import pandas as pd
import win32com.client as win32
tabela_vendas = pd.read_excel('Vendas.xlsx')
# pd.read_excel: Função Pandas de leitura de tabelas formatadas no Excel.
pd.set_option('display.max_columns', None)
# set_option('display.max_columns', None): mostra o Máximo de Colunas no display de Vendas.xlsx
# Pandas: Biblioteca de Integração de Python x Excel.
# Win32com.client: Integração de Python x APIs do Windows.

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
qtd_produto = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
# filtra a Tabela de Vendas em Colunas de acordo com as Colunas do Argumento.
# groupby: agrupa informação de acordo com o Argumento e soma as demais.
# colunas que atendam ao mesmo Argumento, assim, pode-se obter o Faturamento de cada Loja.

ticket_medio = (faturamento['Valor Final'] / qtd_produto['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
# to_frame(): transforma Argumento em uma Tabela.
# Operação de Colunas: variavel[Coluna] / variavel[Coluna].

outlook = win32.Dispatch('outlook.application')
# Integração Python x Outlook
mail = outlook.CreateItem(0)
# Criação de e-mail e elaboração de destinatário, assunto e corpo de texto
mail.To = 'seuemail@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezado(a),</p>
<p>Segue o relatório de vendas por loja</p>
<p><b>Faturamento</b>:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p><b>Quantidade Vendida</b>:</p>
{qtd_produto.to_html()}

<p><b>Ticket Médio dos Produtos</b>:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Em caso de dúvidas, à disposição para esclarecimentos.</p>
<p>At.te,</p>
<p><i>Lucas Gonçalves de Oliveira Martins.</i></p>
'''
# to_html: converte variável em um Formato de HTML.
# {:,.2f}: Formatação que acrescenta ponto e vírgula, com
# dois algarismos significativos após a vírgula (repres. os centavos).
# Três Aspas Simples: para textos com mais de uma linha.
mail.Send()
# Send(): envia o e-mail.
print("Confirmação: E-mail enviado")
