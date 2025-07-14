import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Meu App",
    layout="wide",  # permite usar toda a largura da tela
    initial_sidebar_state="auto"  # ou "expanded" / "collapsed"
)


# Função simples para proteção por senha
def check_password():
    def password_entered():
        if st.session_state["password"] == "ThisIsCadastro":
            st.session_state["password_correct"] = True
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Digite a senha para acessar o relatório:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Digite a senha para acessar o relatório:", type="password", on_change=password_entered, key="password")
        st.error("Senha incorreta")
        return False
    else:
        return True

if not check_password():
    st.stop()

st.title("Relatório de Migração para o Plames Ideal")

# Carregar planilha (altere o caminho para sua planilha)
# Pode ser CSV ou Excel
@st.cache_data
def load_data():
    return pd.read_excel("Migracao_valsa.xlsx")
df = load_data()

# KPI: Quantidade total de registros
total_registros = len(df)
st.metric(label="Quantidade total de registros", value=total_registros)

# Mostrar dados
st.dataframe(df)

# Função para gerar Excel para download
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatório')
        writer.close()
    processed_data = output.getvalue()
    return processed_data

# Download em Excel
excel_data = to_excel(df)
st.download_button(label='📥 Baixar relatório Excel',
                   data=excel_data,
                   file_name='relatorio.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Download em CSV
csv = df.to_csv(index=False).encode('utf-8')
st.download_button(label='📥 Baixar relatório CSV',
                   data=csv,
                   file_name='relatorio.csv',
                   mime='text/csv')
