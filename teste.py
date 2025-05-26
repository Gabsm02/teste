import streamlit as st
import pandas as pd


st.title('Hello world ')

# Carrega a planilha
arquivo = "Dashboard Executivo Celldowntime.xlsx"
df = pd.read_excel(arquivo, engine='openpyxl')

# Aplica filtro: mantém apenas registros com UF = "BA"
if "UF" in df.columns:
    df = df[df["UF"] == "BA"]

# Remove as colunas "CN" e "Selo" se existirem
for coluna in ["CN", "Selo"]:
    if coluna in df.columns:
        df.drop(columns=[coluna], inplace=True)

# Renomeia a coluna "Coordenação" para "Rota" se existir
if "COORDENAÇÃO" in df.columns:
    df.rename(columns={"COORDENAÇÃO": "ROTA"}, inplace=True)

# Salva o resultado em um novo arquivo
df.to_excel("planilha_limpa.xlsx", index=False)

print("Arquivo atualizado com sucesso!") 

# Carrega as planilhas
planilha1 = pd.read_excel('planilha_limpa.xlsx', engine='openpyxl')
planilha2 = pd.read_excel('Tabela de Rotas.xlsx', engine='openpyxl')

# Remove duplicatas para garantir correspondência única por município
rotas_unicas = planilha2[['MUNICÍPIO', 'ROTA']].drop_duplicates(subset='MUNICÍPIO')

# Remove a coluna 'ROTA' da planilha limpa, se já existir
if 'ROTA' in planilha1.columns:
    planilha1 = planilha1.drop(columns=['ROTA'])

# Faz o merge com base na coluna 'MUNICÍPIO'
planilha_final = pd.merge(planilha1, rotas_unicas, on='MUNICÍPIO', how='left')


# Define a nova ordem das colunas
nova_ordem = [
     'EVENTO', 'UF', 'MUNICÍPIO', 'SITE', 'ROTA', 'DATA CRIAÇÃO', 'GESTOR ATUAL',
    'PREVISÃO SOBRESSALENTE'
]

# Reordena as colunas
planilha_final = planilha_final[nova_ordem]


# Salva a planilha final
planilha_final.to_excel('planilha_atualizada.xlsx', index=False)

# Carrega os dados da planilha de backlog
backlog = pd.read_excel("Backlog_detalhes.xlsx", engine="openpyxl")

# Seleciona apenas as colunas desejadas
colunas_desejadas = ["EVENTO", "UF", "Ação", "Prazo", "Status", "Fabricante", "Modelo", "Cod_Modelo", "SS", "Afetação"]
backlog_filtrado = backlog[colunas_desejadas]

# Filtra apenas os registros com UF = "BA"
backlog_ba = backlog_filtrado[backlog_filtrado["UF"] == "BA"]

# Carrega a planilha base
planilha_limpa = pd.read_excel("planilha_atualizada.xlsx", engine="openpyxl")

# Também filtra a planilha base para UF = "BA"
planilha_limpa_ba = planilha_limpa[planilha_limpa["UF"] == "BA"]

# Faz o merge com base nas colunas "Evento" e "UF"
resultado = pd.merge(planilha_limpa_ba, backlog_ba, on=["EVENTO", "UF"], how="left")

# Renomeia a coluna "Coordenação" para "Rota" se existir
if "COORDENAÇÃO" in resultado.columns:
    df.rename(columns={"COORDENAÇÃO": "Rota"}, inplace=True)

if "GESTOR ATUAL" in resultado.columns and "Status" in resultado.columns:
    resultado.loc[resultado["GESTOR ATUAL"] == "ICOMON_O&M-BA_SE (Terceiro)", "Status"] = "Icomon Atuando"
    resultado.loc[resultado["GESTOR ATUAL"] == "O&M ACESSO VIVO_N1_BA/SE_Sobressalente", "Status"] = "Sobressalente"
    resultado.loc[resultado["GESTOR ATUAL"] == "O&M ACESSO VIVO_N1_BA/SE_Acesso", "Status"] = "Acesso"
    resultado.loc[resultado["GESTOR ATUAL"] == "O&M ACESSO VIVO_N1_BA/SE_Fatura_B", "Status"] = "Fatura B"
    resultado.loc[resultado["GESTOR ATUAL"] == "O&M ACESSO VIVO_N1_BA/SE", "Status"] = "Vivo Avaliar"
    resultado.loc[resultado["GESTOR ATUAL"] == "O&M ACESSO VIVO_N1_BA/SE_Falta_AC", "Status"] = "Vivo Avaliar"
    resultado.loc[resultado["GESTOR ATUAL"] == "BBN-SOBRESSALENTES", "Status"] = "Sobressalente"
    resultado.loc[resultado["GESTOR ATUAL"] == "O&M ACESSO VIVO_N1_BA/SE_Vandalismo", "Status"] = "Vivo Avaliar"
    resultado.loc[resultado["GESTOR ATUAL"] == "O&M ACESSO VIVO_N1_NE", "Status"] = "Vivo Avaliar"


# Salva o resultado em uma nova planilha
resultado.to_excel("planilha_final.xlsx", index=False)
resultado.to_csv('planilha_final.csv')

df=pd.read_excel('planilha_final.xlsx')
df_csv =pd.read_csv('planilha_final.csv')

st.dataframe(df_csv)


print("Planilha final criada com sucesso!")