# Separação de dados por Diretoria de Ensino (DE)

## Descrição

A Secretaria da Educação do Estado de São Paulo (SEDUC-SP) é possui 
unidades administrativas regionais, que são responsáveis por gerir as 
Unidades Escolares (UE) das regiões do Estado de São Paulo. Portanto, 
um desafio enfrentado pela Pasta é a comunicação com todas as Diretorias 
Regionais de Ensino (DE), de modo a manter uma comunicação centralizada, 
evitando ruído.

Neste processo, um dos problemas detectados é a separação de dados por
DE que se encontram em um único arquivo, seja em formato .csv ou .xlsx. 
Neste contexto, desenvolveu-se este programa para auxiliar o processo
de separação de arquivos por DE, a fim de encaminhar apenas os dados 
referentes à sua região.

## Como utilizar a aplicação?

1. Acesse o programa por meio do link: https://leofyada-de-heroku-programa-de-lsy865.streamlit.app. 
2. Clique no botão "Browse files" e selecione o arquivo que deseja
   utilizar em formato .csv ou .xlsx (a coluna do nome da DE deve
   ser a primeira da planilha importada);
3. Digite o final do nome de cada um dos arquivos. O início será 
   composto pelo nome da DE. Por exemplo, se digitar "-DE" o nome
   do arquivo de Adamantina será "Adamantina-DE".
   

