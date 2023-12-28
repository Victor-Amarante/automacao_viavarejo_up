import os
import pandas as pd
import numpy as np
import openpyxl
import streamlit as st
from datetime import date, time, datetime, timedelta
from io import StringIO, BytesIO
import base64
from utils import bg_page
from janitor import clean_names

def generate_excel_download_link(df):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    hoje = date.today()
    towrite = BytesIO()
    df.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="BASE_TRATADA_VIAVAREJO_{hoje}.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

# coluna Baixa-ate colocar os horarios atraves da coluna na mesma base Prazo(SLA)
def diminuir_uma_hora(horario_atual):
    novo_horario = horario_atual - timedelta(hours=1)
    novo_horario = novo_horario.strftime("%H:%M")
    return novo_horario

def tratamento_automatico(df_base_geral, df_subtipo_tarefas, df_itapeva, df_impossibilidade):
    # excluir na coluna Escritorio tudo o que não for QCA
    df_base_geral = df_base_geral[df_base_geral["escritorio_"] == "QUEIROZ CAVALCANTI ADVOGADOS"]
    # adicionar 4 novas colunas (nomes no pop)
    colunas_para_adicionar = ["NUCLEO", "BAIXA_ATE", "ITAPEVA", "STATUS"]
    df_base_geral = df_base_geral.assign(**{coluna: "" for coluna in colunas_para_adicionar})
    # coluna Nucleo -> PROCV base geral com base subtipo tarefas para pegar os nucleos
    df_subtipo_tarefas = df_subtipo_tarefas.drop_duplicates(subset=["tipo"])
    procv1 = pd.merge(df_base_geral, df_subtipo_tarefas, how="left", left_on="sub_tipo", right_on="tipo", indicator=True)
    procv1["NUCLEO"] = procv1["nucleo"]
    procv1 = procv1.drop(['tipo', 'nucleo', '_merge'], axis=1)
    # coluna Baixa-ate colocar os horarios atraves da coluna na mesma base Prazo(SLA)
    procv1["prazo_sla_"] = pd.to_datetime(procv1["prazo_sla_"])
    procv1['BAIXA_ATE'] = procv1['prazo_sla_'].apply(diminuir_uma_hora)
    now = pd.Timestamp.now()
    cutoff_time = now.replace(hour=8, minute=0, second=0, microsecond=0)
    # Adiciona um "-" onde o tempo é anterior ao tempo de corte
    procv1['ATRASO'] = procv1['prazo_sla_'].apply(lambda x: '-' if x <= cutoff_time else '')
    # coluna Itapeva -> PROCV base geral com base itapeva para pegar o NPC
    df_itapeva = df_itapeva['npc']
    procv2 = pd.merge(procv1, df_itapeva, how="left", left_on="_processo_id", right_on="npc", indicator=True)
    procv2 = procv2.drop(['_merge'], axis=1)
    # coluna Status - > PROCV base geral com a base de impossibilidade e cancelamento
    df_impossibilidade = df_impossibilidade[["id_da_tarefa_", "status"]]
    procv3 = pd.merge(procv2, df_impossibilidade, how='left', left_on="id_da_tarefa_", right_on="id_da_tarefa_", indicator=True)
    procv3['STATUS'] = procv3['status_y']
    procv3 = procv3.drop(['status_y', '_merge'], axis=1)
    procv3.rename(columns={"status_x": "status", "npc": "NPC"}, inplace=True)
    procv3.loc[(procv3["STATUS"].isna()) & (procv3['ATRASO'] == '-'), 'STATUS'] = "ATRASO - DEVE SER JUSTIFICADO!"
    return procv3

st.set_page_config(
    page_title="Automação - VIAVAREJO",
    page_icon='qca_logo_2.png',
    layout="wide",
)
bg_page('bg_dark.png')
hide_menu = """
<style>
#MainMenu {
    visibility:visible;
}

footer {
    visibility:visible;
}

footer:before {
    content:'Desenvolvido pela Eficiência Jurídica - Controladoria Jurídica';
    display:block;
    position:relative;
    color:#6c6a76;
}
</style>
"""

st.markdown(hide_menu, unsafe_allow_html=True)

# Titulo da pagina
st.markdown('# Programa QCA - Via Varejo')
st.markdown("## Tratamento Automático da Planilha")
st.markdown('#### Para que a automação funcione, insira abaixo as bases necessárias para os PROCVs.')

st.divider()
st.warning("Todas as bases devem estar no formato Excel (.xlsx)")   # informar ao usuario que todos os arquivos precisam ser arquivos excel
arquivo1 = st.file_uploader("Importe a BASE PRINCIPAL", type="xlsx")    # base principal
arquivo2 = st.file_uploader("Importe a BASE SUBTIPO TAREFAS", type="xlsx")  # base subTipo tarefas
arquivo3 = st.file_uploader("Importe a BASE ITAPEVA", type="xlsx")  # base itapeva
arquivo4 = st.file_uploader("Importe a BASE IMPOSSIBILIDADE E CANCELAMENTO", type="xlsx")   # base impossibilidade

if (arquivo1 is not None) and (arquivo2 is not None) and (arquivo3 is not None) and (arquivo4 is not None):
    df_base_geral = pd.read_excel(arquivo1, engine="openpyxl")
    df_base_geral = clean_names(df_base_geral)
    
    df_subtipo_tarefas = pd.read_excel(arquivo2, engine="openpyxl")
    df_subtipo_tarefas = clean_names(df_subtipo_tarefas)
    
    df_itapeva = pd.read_excel(arquivo3, engine="openpyxl")
    df_itapeva = clean_names(df_itapeva)
    
    df_impossibilidade = pd.read_excel(arquivo4, engine="openpyxl")
    df_impossibilidade = clean_names(df_impossibilidade)
    
    st.success("As bases foram carregadas!")
    
    botao_tratamento_automatico = st.button("Iniciar o procedimento")
    if botao_tratamento_automatico:
        base_tratada = tratamento_automatico(df_base_geral=df_base_geral, df_subtipo_tarefas=df_subtipo_tarefas,
                                            df_itapeva=df_itapeva, df_impossibilidade=df_impossibilidade)
        st.success("O processo foi finalizado! A base está disponível para download abaixo")
        generate_excel_download_link(base_tratada)
        
