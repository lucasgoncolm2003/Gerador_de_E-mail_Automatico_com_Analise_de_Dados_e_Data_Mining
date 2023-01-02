# Gerador de E-mail Automático com Análise de Dados e Data Mining em Python - [Pandas &amp; Pywin32]
## Biblioteca Pandas
  A Biblioteca Pandas é importante para a Análise de Dados para o Python, nesse projeto, ela foi usada para fazer uma Leitura de uma Planilha do Excel: a Planilha Vendas.xlsx. A Pandas é usada para construir Estruturas de Dados, para Manipulação ou Limpeza de Dados. Além disso, engloba o Processamento Numérico e a Construção de Gráficos.
## Ticket Médio
   Nesse projeto, Duas Colunas foram criadas para fins de criar um Relatório: a Coluna de Faturamento, que é baseada na Soma Agrupada de Valores Finais por IDs das Lojas; a Coluna de Quantidade de Produtos, que é baseada na Soma Agrupada de Produtos por IDs das Lojas. A partir disso, pode-se encontrar o Ticket Médio de certa empresa: esse valor é baseado na Razão existente entre o Faturamento e a Quantidade de Produtos, ou seja, quanto se fatura por produto. Assim, uma nova Coluna é feita, associando o Ticket Médio com os IDs das Lojas.
## Biblioteca Pywin32
  Após isso, usa-se a Biblioteca Pywin32, uma Biblioteca que possui a função de manipular o Windows, seus Comandos e Aplicativos através do Python. Nesse caso, essa Biblioteca é usada para abrir o Outlook e enviar o E-mail, com um Destinatário Específico, com um Assunto Específico e com um Corpo de E-mail baseado em um Relatório com Faturamento, Quantidade Vendida e Ticket Médio dos Produtos (todos em relação aos IDs das Lojas). Após todo o e-mail estar desenvolvido, o Python envia o E-mail ao Destinatário e faz um Print da Confirmação no Terminal. Abaixo estão os prints de um modelo de E-mail de Relatório enviado automaticamente pelo mesmo código feito.
<div style="display: inline_block" align="center"><br>
  <img align="center" src="https://user-images.githubusercontent.com/112359793/210268114-699ef95c-19aa-4868-b933-1a00943559ff.png"/>
  <img align="center" src="https://user-images.githubusercontent.com/112359793/210268169-edb61894-2f27-42eb-866c-da9ab09e684e.png"/>
  <img align="center" src="https://user-images.githubusercontent.com/112359793/210268225-b021ebae-86d3-4e0a-a4a7-f4ac09d1ef19.png"/>
</div>
