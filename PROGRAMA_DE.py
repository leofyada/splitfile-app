# Importando as bibliotecas que serão utilizadas
import base64
import streamlit as st
import pandas as pd
from zipfile import ZipFile
import openpyxl
import os
import io
from xlsxwriter import Workbook
import streamlit.components.v1 as components


# Criando um título para o aplicativo
st.title("Aplicativo para dividir dados por DE")
# Orientações de como utilizar a ferramenta (separar arquivos)
st.text("1-) Faça o upload do arquivo. "
        "\n\nIMPORTANTE: a coluna com a DE deve ser a primeira da planilha.")
# Atribuindo o arquivo a uma variável denominada "arquivo"
arquivo = st.file_uploader("Upload do arquivo", type=["csv", "xlsx"])
if arquivo is not None:
    # Realizando a leitura do arquivo excel e atribuindo à variável "arquivo_lido"
    arquivo_lido = pd.read_excel(arquivo, engine='openpyxl')
    # Renomeando a primeira coluna para "NM_DIRETORIA"
    arquivo_lido.rename(columns={list(arquivo_lido)[0]: "NM_DIRETORIA"}, inplace=True)
    # Criando um DataFrame com apenas as Diretorias de Ensino
    de = arquivo_lido.iloc[0:len(arquivo_lido), 0:1].drop_duplicates()
    # Criando um arquivo Zip
    zipObj = ZipFile("Diretorias", "w")
    ZipfileDotZip = "Diretorias"

    st.write("Todos os arquivos iniciarão com o nome da DE.")
    name = st.text_input("Digite o início do nome dos arquivos. "
                         "\n\n Exemplo: se digitar 'DE', o nome do arquivo da DE de ITU será DE-ITU.")


    # Criando um loop para separar os dados e criar arquivos
    for n in range(0, len(de), 1):
        base_de = arquivo_lido[arquivo_lido["NM_DIRETORIA"] == de.iloc[n, 0]]
        # criando arquivo csv
        base_de.to_excel(name+de.iloc[n, 0]+".xlsx", index=False)
        # adicionando arquivos a uma pasta
        zipObj.write(name+de.iloc[n, 0]+".xlsx")
    # fechando o arquivo Zip
    zipObj.close()


    with open(ZipfileDotZip, "rb") as f:
        bytes = f.read()
        b64 = base64.b64encode(bytes).decode()
        href = f"<a href=\"data:file/zip;base64,{b64}\" download='{ZipfileDotZip}.zip'>\
            Clique aqui para baixar os arquivos!\
        </a>"
    st.markdown(href, unsafe_allow_html=True)

# # Orientações de como utilizar a ferramenta (juntar arquivos)
# st.text("2-) Faça o upload da pasta em formato .zip com os arquivos que deseja juntar. "
#         "\n\nIMPORTANTE: os arquivos precisar ser em formato .csv ou .xlsx e precisam conter exatamente as mesmas colunas.")

# # Atribuindo o arquivo a uma variável denominada "arquivo"
# pasta = st.file_uploader("Upload da pasta", type=["zip"])

# if pasta is not None:
#     with ZipFile(pasta, 'r') as zObject:
#         zObject.extractall("/tmp")

# arquivo_completo = pd.DataFrame()

# for arquivo in os.listdir("/tmp"):
#     if arquivo.lower().endswith(('.xlsx', '.csv')):
#         arquivo_pandas = pd.read_excel("/tmp/"+arquivo, engine='openpyxl', sheet_name=0)
#         arquivo_pandas = arquivo_pandas[arquivo_pandas.filter(regex='^(?!Unnamed)').columns]
#         arquivo_completo = pd.concat([arquivo_completo, arquivo_pandas])

# arquivo_completo.drop(arquivo_completo.columns[len(arquivo_completo.columns)-1], axis=1, inplace=True)

# buffer = io.BytesIO()
# with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
#     arquivo_completo.to_excel(writer, sheet_name='Sheet1', index=False)

#     writer.close()

#     download = st.download_button(
#         label="Download do arquivo final",
#         data=buffer,
#         file_name='arquivo_completo.xlsx',
#         mime='application/vnd.ms-excel'
#     )

components.iframe("https://leonardo-yada.shinyapps.io/relatorios_personalizados/", height=300)

