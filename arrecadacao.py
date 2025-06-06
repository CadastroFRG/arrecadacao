import streamlit as st
import pandas as pd
import os

# Configurar variáveis de ambiente para locale antes de importar locale
os.environ['LC_ALL'] = 'en_US.UTF-8'
os.environ['LANG'] = 'en_US.UTF-8'

import locale

# Tentar setar locale para 'en_US.UTF-8'; se falhar, usar locale básica 'C'
try:
    locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_ALL, 'C')

import yagmail
from fpdf import FPDF
from datetime import datetime
from docx import Document
from docx.shared import Pt
import re


# Configurar locale para formatação de moeda
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except locale.Error:
        st.warning("Locale pt_BR não encontrado. A formatação de moeda pode usar '.' como separador decimal.")

DATA_PATH = "dados_formulario.csv"
EMAIL_REMETENTE = "brunomelo@frg.com.br" # ATUALIZE COM SEU E-MAIL
EMAIL_SENHA = "Trocar@123" # ATUALIZE COM SUA SENHA DE APP DO GMAIL
# --- ETAPAS ATUALIZADAS ---
ETAPAS = ["Aguardando Resposta", "Respondido", "Relação de Crédito", "Desconto de quitação de deficit", "Termo de Portabilidade", "Carta de Portabilidade"]
WORD_TEMPLATE_PATH = "template_quitacao.docx"
WORD_TEMPLATE_PORT_PATH = "template_termo_de_portabilidade.docx"
# --- NOVO TEMPLATE ---
WORD_TEMPLATE_CARTA_PATH = "template_carta.docx" # Certifique-se de que este arquivo existe e é .docx

def carregar_dados():
    colunas_necessarias = [
        "Nome", "Matricula", "CPF", "Email", "Comentário", "Área", "Etapa",
        "Dados Adicionais", "Creditar", "Banco", "Conta", "Agencia", "NomeAgencia",
        "ValorRS", "TipoEntidade", "Patrocinadora", "Plano", "QtdeCotas", "ValorCota",
        "DataValorCota", "MesAnoRelacao", "DataPagamento",
        "NRefDoc", "Rua", "Complemento", "Bairro", "CEP", "Cidade", "UF",
        "MesCalculoCotaDoc", "Deficit2014", "Deficit2022",
        # NOVAS COLUNAS PARA TERMO DE PORTABILIDADE
        "Data_admissao", "Data_desligamento", "Data_inscricao",
        "plano_de_beneficio", "CNPB", "plano_receptor", "cnpj_plano_receptor",
        "endereco_plano_receptor", "cep_plano_receptor", "cidade_plano_receptor",
        "contato_plano_receptor", "telefone_plano_receptor", "email_plano_receptor",
        "banco_plano_receptor", "agencia_plano_receptor", "conta_plano_receptor",
        "Parcela_Participante", "Parcela_Patrocinadora", "Total_acumulado",
        "Regime_de_tributacao", "Recursos_portados", "debito", "total_a_ser_portado",
        "Data_base_portabilidade",
        # --- NOVAS COLUNAS PARA CARTA DE PORTABILIDADE ---
        "Data_de_Transferencia_Carta", "Banco_Carta", "Agencia_Carta", "Conta_Corrente_Carta" 
    ]
    if os.path.exists(DATA_PATH):
        try:
            df = pd.read_csv(DATA_PATH)
            for col in colunas_necessarias:
                if col not in df.columns:
                    df[col] = pd.NA
            return df
        except pd.errors.EmptyDataError:
            return pd.DataFrame(columns=colunas_necessarias)
        except Exception as e:
            st.error(f"Erro ao carregar o arquivo CSV: {e}")
            return pd.DataFrame(columns=colunas_necessarias)
    else:
        return pd.DataFrame(columns=colunas_necessarias)

def salvar_dados(novo_dado):
    df = carregar_dados()
    novo_dado_df = pd.DataFrame([novo_dado])
    for col in df.columns:
        if col not in novo_dado_df.columns:
            novo_dado_df[col] = pd.NA
    novo_dado_df = novo_dado_df[df.columns]
    df = pd.concat([df, novo_dado_df], ignore_index=True)
    df.to_csv(DATA_PATH, index=False)

def atualizar_etapa(nome, nova_etapa):
    df = carregar_dados()
    if not df.empty and "Nome" in df.columns:
        df.loc[df["Nome"] == nome, "Etapa"] = nova_etapa
        df.to_csv(DATA_PATH, index=False)

def salvar_dados_completos(nome, dados_dict):
    df = carregar_dados()
    if not df.empty and "Nome" in df.columns:
        idx_list = df[df["Nome"] == nome].index
        if not idx_list.empty:
            idx = idx_list[0] # Pega o primeiro índice se houver múltiplos (não deveria)
            for chave, valor in dados_dict.items():
                if chave in df.columns:
                    df.loc[idx, chave] = valor
                else:
                    st.warning(f"Tentativa de salvar coluna inexistente: {chave}")
            df.to_csv(DATA_PATH, index=False)
            return df.loc[idx].to_dict()
    return {}

EMAILS_POR_AREA = {"RH": "rh@empresa.com", "Financeiro": "financeiro@empresa.com", "Seguridade": "seguridade@empresa.com"} # Exemplo, adicione mais se necessário

def enviar_email(email_pessoal, nome, area):
    try:
        destinatario = EMAILS_POR_AREA.get(area)
        if not destinatario:
            st.warning(f"⚠️ Nenhum e-mail configurado para a área: {area}")
            return
        # Usar uma senha de aplicativo se o 2FA do Gmail estiver ativado
        # Certifique-se de que EMAIL_REMETENTE e EMAIL_SENHA estão configurados corretamente
        yag = yagmail.SMTP(EMAIL_REMETENTE, EMAIL_SENHA)
        assunto = f"Novo cadastro aguardando resposta - {nome}"
        conteudo = f"Olá equipe de {area},\n\nUm novo formulário foi preenchido por {nome} ({email_pessoal}).\n\nPor favor, acesse o sistema.\n\nAtt,\nSistema Streamlit"
        yag.send(to=destinatario, subject=assunto, contents=conteudo)
        st.info(f"E-mail de notificação enviado para {destinatario}.")
    except Exception as e:
        st.error(f"❌ Erro ao enviar e-mail: {e}.")


def gerar_pdf_relacao_credito(dados):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", 'B', size=12)
    pdf.cell(0, 10, "REAL GRANDEZA", ln=True, align='C')
    pdf.set_font("Arial", 'B', size=10)
    pdf.cell(0, 5, "FUNDAÇÃO DE PREVIDENCIAE ASSISTÊNCIA SOCIAL", ln=True, align='C')
    pdf.ln(5)
    mes_ano_relacao = dados.get('MesAnoRelacao', datetime.now().strftime("%b/%y").lower())
    current_y_for_relation = pdf.get_y()
    pdf.set_font("Arial", size=10)
    pdf.set_xy(150, current_y_for_relation)
    pdf.multi_cell(50, 5, f"Relação nº 158\n{mes_ano_relacao}", align='R')
    pdf.set_xy(10, current_y_for_relation + 5)
    pdf.cell(0, 5, f"GBP/AMX {mes_ano_relacao}", ln=False)
    pdf.set_y(current_y_for_relation + 10)
    pdf.ln(5)
    pdf.set_font("Arial", '', size=10)
    pdf.cell(0, 7, "DIRETORIA DE SEGURIDADE - DS", ln=True)
    pdf.cell(0, 7, "GERÊNCIA DE ESTATÍSTICA E ATUÁRIA - GEA", ln=True)
    pdf.ln(5)
    pdf.set_font("Arial", 'B', size=12)
    pdf.cell(0, 7, "PORTABILIDADE", ln=True, align='C')
    pdf.ln(5)
    pdf.set_font("Arial", size=10)
    pdf.cell(30, 7, "Creditar:")
    x_before_cod_banco = pdf.get_x()
    pdf.set_x(120)
    pdf.cell(25, 7, "Cód. do Banco:")
    pdf.set_font("Arial", 'B', size=10)
    pdf.cell(0, 7, str(dados.get('Banco', '')), ln=True)
    pdf.set_x(x_before_cod_banco -20) # Deve ser a mesma X da célula "Creditar:"
    pdf.set_font("Arial", 'B', size=10)
    pdf.cell(0, 7, str(dados.get('Creditar', 'Banco Bradesco')), ln=True)
    pdf.set_font("Arial", size=10)
    pdf.cell(15, 7, "Nome:")
    pdf.set_font("Arial", 'B', size=10)
    pdf.cell(0, 7, "Real Grandeza", ln=True)
    pdf.set_font("Arial", size=10)
    col_width_conta = 30; col_width_cod_ag = 30; col_width_nome_ag = 60; col_width_valor = 0
    pdf.cell(col_width_conta, 7, f"Conta: {dados.get('Conta', '')}")
    pdf.cell(col_width_cod_ag, 7, f"Cód. Agência: {dados.get('Agencia', '')}")
    pdf.cell(col_width_nome_ag, 7, f"Nome da Agência: {dados.get('NomeAgencia', '')}")
    pdf.set_font("Arial", 'B', size=10)
    pdf.cell(col_width_valor, 7, f"Valor em R$: {dados.get('ValorRS', '')}", ln=True)
    pdf.set_font("Arial", size=10)
    pdf.cell(35, 7, "Tipo de Entidade:"); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('TipoEntidade', 'Fechada')), ln=True); pdf.set_font("Arial", size=10)
    pdf.cell(35, 7, "PATROCINADORA:"); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('Patrocinadora', 'FURNAS')), ln=True); pdf.set_font("Arial", size=10)
    pdf.cell(35, 7, "PLANO:"); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('Plano', 'CONTRIBUIÇÃO DEFINIDA - CD')), ln=True); pdf.set_font("Arial", size=10)
    pdf.cell(150, 7, "Total", align='R'); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('ValorRS', '')), ln=True); pdf.set_font("Arial", size=10)
    pdf.ln(3)
    pdf.cell(0, 7, f"Para pagamento dia: {dados.get('DataPagamento', '03/jun/2025')}", ln=True)
    pdf.ln(7)
    pdf.set_font("Arial", 'B', size=11); pdf.cell(0, 7, "Identificação do Participante", ln=True, align='C'); pdf.set_font("Arial", size=10)
    pdf.cell(20, 7, "Nome:"); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('Nome', '')), ln=True); pdf.set_font("Arial", size=10)
    pdf.cell(20, 7, "Matrícula:"); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('Matricula', '')), ln=True); pdf.set_font("Arial", size=10)
    pdf.cell(20, 7, "C.P.F.:"); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('CPF', '')), ln=True); pdf.set_font("Arial", size=10)
    pdf.cell(30, 7, "Qtde. de Cotas:"); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('QtdeCotas', '')), ln=True); pdf.set_font("Arial", size=10)
    data_valor_cota_pdf = dados.get('DataValorCota', '30/04/2025')
    pdf.cell(55, 7, f"Valor da Cota ({data_valor_cota_pdf}):"); pdf.set_font("Arial", 'B', size=10); pdf.cell(0, 7, str(dados.get('ValorCota', '')), ln=True); pdf.set_font("Arial", size=10)
    pdf.ln(10)
    pdf.set_font("Arial", 'I', size=9); pdf.cell(0, 7, "Patrícia Melo e Souza", ln=True, align='C'); pdf.cell(0, 5, "Diretora de Seguridade", ln=True, align='C')
    output_filename = f"relacao_credito_{dados.get('Nome', 'Desconhecido').replace(' ', '_')}.pdf"
    pdf.output(output_filename, 'F')
    return output_filename

def formatar_moeda_para_exibicao(valor_numerico):
    try:
        # Tenta formatar como moeda pt_BR, que usa vírgula para decimal e ponto para milhar
        return locale.currency(float(valor_numerico), grouping=True, symbol=None)
    except (ValueError, TypeError):
        return "0,00"

def desformatar_string_para_float(valor_str):
    if valor_str is None or str(valor_str).strip() == "" or str(valor_str).lower() == 'nan':
        return 0.0
    try:
        # Remove separadores de milhar pt-BR (.), depois substitui vírgula decimal pt-BR (,) por ponto (.)
        return float(str(valor_str).replace('.', '').replace(',', '.'))
    except ValueError:
        # Tenta tratar como se já fosse um float em formato de string com ponto decimal
        try:
            return float(valor_str)
        except ValueError:
            st.warning(f"Não foi possível converter '{valor_str}' para número. Usando 0.0.")
            return 0.0

# --- FUNÇÃO PARA SUBSTITUIÇÃO MAIS ROBUSTA (ajustada para manter estilo) ---
def replace_placeholders_in_document(doc, substitutions):
    """
    Substitui placeholders em parágrafos e células de tabelas do documento DOCX.
    Esta função tenta ser mais robusta para placeholders que podem estar divididos em runs,
    e tenta preservar o estilo da primeira run.
    """
    # Helper para processar runs e preservar estilo
    def process_paragraph_runs(p, key, value):
        full_text = "".join([run.text for run in p.runs])
        if key in full_text:
            new_full_text = full_text.replace(key, value)
            
            # Se houver runs, tente manter o estilo da primeira
            if p.runs:
                first_run_style = p.runs[0].style # Guarda o estilo da primeira run
                first_run_font = p.runs[0].font # Guarda as propriedades da fonte
                
                # Limpar todas as runs existentes removendo os elementos XML
                # Cria uma lista para evitar problemas de modificação durante iteração
                for run in list(p.runs): 
                    p.runs[0]._element.getparent().remove(run._element) 
                
                # Adicionar uma nova run e aplicar o estilo da primeira run
                new_run = p.add_run(new_full_text)
                new_run.style = first_run_style
                new_run.font.name = first_run_font.name
                new_run.font.size = first_run_font.size
                new_run.font.bold = first_run_font.bold
                new_run.font.italic = first_run_font.italic
                new_run.font.underline = first_run_font.underline
            else: # Se não houver runs (parágrafo vazio), crie uma nova
                p.add_run(new_full_text)

    # Para parágrafos no corpo principal
    for p in doc.paragraphs:
        for key, value in substitutions.items():
            process_paragraph_runs(p, key, value)

    # Para tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in substitutions.items():
                        process_paragraph_runs(p, key, value)
    
    # Para cabeçalhos e rodapés (se houver)
    for section in doc.sections:
        # Cabeçalhos
        if section.header:
            for p in section.header.paragraphs:
                for key, value in substitutions.items():
                    process_paragraph_runs(p, key, value)
            for table in section.header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for key, value in substitutions.items():
                                process_paragraph_runs(p, key, value)

        # Rodapés
        if section.footer:
            for p in section.footer.paragraphs:
                for key, value in substitutions.items():
                    process_paragraph_runs(p, key, value)
            for table in section.footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for key, value in substitutions.items():
                                process_paragraph_runs(p, key, value)


def gerar_documento_quitacao(dados_completos):
    if not os.path.exists(WORD_TEMPLATE_PATH):
        st.error(f"Template Word '{WORD_TEMPLATE_PATH}' não encontrado!")
        return None, None
    try:
        doc = Document(WORD_TEMPLATE_PATH)
    except Exception as e:
        st.error(f"Erro ao carregar o template Word '{WORD_TEMPLATE_PATH}': {e}")
        return None, None

    try:
        qtde_cotas = desformatar_string_para_float(dados_completos.get('QtdeCotas', '0'))
        valor_cota = desformatar_string_para_float(dados_completos.get('ValorCota', '0'))
        total_reserva_poupanca_rs = qtde_cotas * valor_cota

        deficit_2014_val = desformatar_string_para_float(dados_completos.get('Deficit2014', '0'))
        deficit_2022_val = desformatar_string_para_float(dados_completos.get('Deficit2022', '0'))
        debito_total_deficit_rs = deficit_2014_val + deficit_2022_val

        # Lógica para {{DESCRICAO_DEFICIT}}
        anos_com_deficit = []
        if deficit_2014_val > 0:
            anos_com_deficit.append("2014")
        if deficit_2022_val > 0:
            anos_com_deficit.append("2022")
        placeholder_descricao_deficit = " e ".join(anos_com_deficit) if anos_com_deficit else ""

        valor_a_receber_rs = total_reserva_poupanca_rs - debito_total_deficit_rs
    except Exception as e:
        st.error(f"Erro nos cálculos para o documento: {e}. Verifique os valores de cotas e déficits.")
        st.write("Valores usados nos cálculos (após conversão):")
        st.write(f"     Qtde Cotas: {qtde_cotas if 'qtde_cotas' in locals() else 'Erro'}")
        st.write(f"     Valor Cota: {valor_cota if 'valor_cota' in locals() else 'Erro'}")
        st.write(f"     Déficit 2014: {deficit_2014_val if 'deficit_2014_val' in locals() else 'Erro'}")
        st.write(f"     Déficit 2022: {deficit_2022_val if 'deficit_2022_val' in locals() else 'Erro'}")
        return None, None

    substituicoes = {
        "{{N_REF}}": str(dados_completos.get('NRefDoc', '')),
        "{{NOME_PARTICIPANTE}}": str(dados_completos.get('Nome', '')),
        "{{ENDERECO_RUA}}": str(dados_completos.get('Rua', '')),
        "{{ENDERECO_COMPLEMENTO}}": str(dados_completos.get('Complemento', '')),
        "{{ENDERECO_BAIRRO}}": str(dados_completos.get('Bairro', '')),
        "{{ENDERECO_CEP}}": str(dados_completos.get('CEP', '')),
        "{{ENDERECO_CIDADE_UF}}": f"{dados_completos.get('Cidade', '')} - {dados_completos.get('UF', '')}",
        "{{ASSUNTO_MATRICULA}}": str(dados_completos.get('Matricula', '')),
        "{{ASSUNTO_PLANO}}": str(dados_completos.get('Plano', '')),
        "{{ASSUNTO_EMPRESA}}": str(dados_completos.get('Patrocinadora', '')),
        "{{DATA_PAGAMENTO_CREDITO}}": str(dados_completos.get('DataPagamento', '')),
        "{{MES_CALCULO_COTA}}": str(dados_completos.get('MesCalculoCotaDoc', '')),
        "{{SALDO_RESERVA_COTAS}}": formatar_moeda_para_exibicao(qtde_cotas),
        "{{VALOR_DA_COTA_RS}}": formatar_moeda_para_exibicao(valor_cota),
        "{{TOTAL_RESERVA_POUPANCA_RS}}": formatar_moeda_para_exibicao(total_reserva_poupanca_rs),
        "{{DEBITO_TOTAL_DEFICIT_RS}}": formatar_moeda_para_exibicao(debito_total_deficit_rs),
        "{{DESCRICAO_DEFICIT}}": placeholder_descricao_deficit, # Atualizado
        "{{VALOR_A_RECEBER_RS}}": formatar_moeda_para_exibicao(valor_a_receber_rs)
    }

    # --- LINHAS DE DEBUG ADICIONADAS PARA QUITAÇÃO ---
    st.write("---")
    st.write("### Debugging: Dicionário de Substituições para Quitação")
    st.json(substituicoes) # Usa st.json para uma visualização mais legível
    st.write("---")
    # --------------------------------------------------

    # Chama a nova função de substituição
    replace_placeholders_in_document(doc, substituicoes)

    nome_base = str(dados_completos.get('Nome', 'Desconhecido')).replace(' ', '_').replace('/', '_')
    output_docx_path = f"quitacao_deficit_{nome_base}.docx"
    output_pdf_path = f"quitacao_deficit_{nome_base}.pdf"

    try:
        doc.save(output_docx_path)
        st.info(f"Arquivo DOCX '{output_docx_path}' gerado.")
        try:
            convert(output_docx_path, output_pdf_path)
            st.info(f"Arquivo PDF '{output_pdf_path}' gerado.")
            return output_pdf_path, output_docx_path
        except Exception as e_pdf:
            st.warning(f"DOCX gerado, mas falha ao converter para PDF: {e_pdf}")
            st.info("Verifique se o Microsoft Word ou LibreOffice está instalado e se o erro 'pywintypes' foi resolvido (veja instruções).")
            return None, output_docx_path
    except Exception as e_docx:
        st.error(f"Erro ao salvar o documento DOCX: {e_docx}")
        return None, None

def gerar_documento_portabilidade(dados_completos):
    if not os.path.exists(WORD_TEMPLATE_PORT_PATH):
        st.error(f"Template Word de Portabilidade '{WORD_TEMPLATE_PORT_PATH}' não encontrado!")
        return None, None
    try:
        doc = Document(WORD_TEMPLATE_PORT_PATH)
    except Exception as e:
        st.error(f"Erro ao carregar o template Word de Portabilidade '{WORD_TEMPLATE_PORT_PATH}': {e}")
        return None, None

    try:
        # Cálculos de valores (ajuste se a lógica for diferente)
        parcela_participante_val = desformatar_string_para_float(dados_completos.get('Parcela_Participante', '0'))
        parcela_patrocinadora_val = desformatar_string_para_float(dados_completos.get('Parcela_Patrocinadora', '0'))
        debito_val = desformatar_string_para_float(dados_completos.get('debito', '0'))

        total_acumulado_val = parcela_participante_val + parcela_patrocinadora_val
        valor_total_a_ser_portado_val = total_acumulado_val - debito_val

    except Exception as e:
        st.error(f"Erro nos cálculos para o Termo de Portabilidade: {e}. Verifique os valores monetários.")
        return None, None
    
    substituicoes = {
        "{{NOME_PARTICIPANTE}}": str(dados_completos.get('Nome', '')),
        "{{CPF}}": str(dados_completos.get('CPF', '')),
        "{{Matricula}}": str(dados_completos.get('Matricula', '')),
        "{{ENDERECO_COMPLEMENTO}}": str(dados_completos.get('Complemento', '')), # Reutilizado
        "{{ENDERECO_RUA}}": str(dados_completos.get('Rua', '')), # Reutilizado
        "{{ENDERECO_BAIRRO}}": str(dados_completos.get('Bairro', '')), # Reutilizado
        "{{ENDERECO_CIDADE_UF}}": f"{dados_completos.get('Cidade', '')} - {dados_completos.get('UF', '')}", # Reutilizado
        "{{ENDERECO_CEP}}": str(dados_completos.get('CEP', '')), # Reutilizado
        "{{ASSUNTO_EMPRESA}}": str(dados_completos.get('Patrocinadora', '')), # Reutilizado
        "{{Data_admissao}}": str(dados_completos.get('Data_admissao', '')),
        "{{Data_desligamento}}": str(dados_completos.get('Data_desligamento', '')),
        "{{Data_inscricao}}": str(dados_completos.get('Data_inscricao', '')),
        "{{plano_de_beneficio}}": str(dados_completos.get('plano_de_beneficio', '')),
        "{{CNPB}}": str(dados_completos.get('CNPB', '')),
        "{{plano_receptor}}": str(dados_completos.get('plano_receptor', '')),
        "{{cnpj_plano_receptor}}": str(dados_completos.get('cnpj_plano_receptor', '')),
        "{{endereco_plano_receptor}}": str(dados_completos.get('endereco_plano_receptor', '')),
        "{{cep_plano_receptor}}": str(dados_completos.get('cep_plano_receptor', '')),
        "{{cidade_plano_receptor}}": str(dados_completos.get('cidade_plano_receptor', '')),
        "{{contato_plano_receptor}}": str(dados_completos.get('contato_plano_receptor', '')),
        "{{telefone_plano_receptor}}": str(dados_completos.get('telefone_plano_receptor', '')),
        "{{email_plano_receptor}}": str(dados_completos.get('email_plano_receptor', '')),
        "{{banco_plano_receptor}}": str(dados_completos.get('banco_plano_receptor', '')),
        "{{agencia_plano_receptor}}": str(dados_completos.get("agencia_plano_receptor", '')),
        "{{conta_plano_receptor}}": str(dados_completos.get("conta_plano_receptor", '')),
        "{{Parcela_Participante}}": formatar_moeda_para_exibicao(parcela_participante_val),
        "{{Parcela_Patrocinadora}}": formatar_moeda_para_exibicao(parcela_patrocinadora_val),
        "{{Total_acumulado}}": formatar_moeda_para_exibicao(total_acumulado_val),
        "{{Regime_de_tributacao}}": str(dados_completos.get('Regime_de_tributacao', '')),
        "{{Recursos_portados}}": formatar_moeda_para_exibicao(desformatar_string_para_float(dados_completos.get('Recursos_portados', '0'))),
        "{{debito}}": formatar_moeda_para_exibicao(debito_val),
        "{{total_a_ser_portado}}": formatar_moeda_para_exibicao(valor_total_a_ser_portado_val),
        "{{Data_base}}": str(dados_completos.get('Data_base_portabilidade', ''))
    }

    # --- LINHAS DE DEBUG ADICIONADAS PARA PORTABILIDADE ---
    st.write("---")
    st.write("### Debugging: Dicionário de Substituições para Termo de Portabilidade")
    st.json(substituicoes) # Usa st.json para uma visualização mais legível
    st.write("---")
    # -------------------------------------------------------

    # Chama a nova função de substituição para o Termo de Portabilidade
    replace_placeholders_in_document(doc, substituicoes)

    nome_base = str(dados_completos.get('Nome', 'Desconhecido')).replace(' ', '_').replace('/', '_')
    output_docx_path = f"termo_portabilidade_{nome_base}.docx"
    output_pdf_path = f"termo_portabilidade_{nome_base}.pdf"

    try:
        doc.save(output_docx_path)
        st.info(f"Arquivo DOCX '{output_docx_path}' gerado.")
        try:
            convert(output_docx_path, output_pdf_path)
            st.info(f"Arquivo PDF '{output_pdf_path}' gerado.")
            return output_pdf_path, output_docx_path
        except Exception as e_pdf:
            st.warning(f"DOCX gerado, mas falha ao converter para PDF: {e_pdf}")
            st.info("Verifique se o Microsoft Word ou LibreOffice está instalado e se o erro 'pywintypes' foi resolvido (veja instruções).")
            return None, output_docx_path
    except Exception as e_docx:
        st.error(f"Erro ao salvar o documento DOCX: {e_docx}")
        return None, None

# --- NOVA FUNÇÃO: GERAR CARTA DE PORTABILIDADE ENTRE PLANOS ---
def gerar_documento_carta_portabilidade(dados_completos):
    if not os.path.exists(WORD_TEMPLATE_CARTA_PATH):
        st.error(f"Template Word da Carta de Portabilidade '{WORD_TEMPLATE_CARTA_PATH}' não encontrado! Por favor, converta seu template .doc para .docx.")
        return None, None
    try:
        doc = Document(WORD_TEMPLATE_CARTA_PATH)
    except Exception as e:
        st.error(f"Erro ao carregar o template Word da Carta de Portabilidade '{WORD_TEMPLATE_CARTA_PATH}': {e}")
        return None, None

    # Obter dados já existentes do CSV ou preenchidos anteriormente
    # Usando .get() para garantir que não haja erro se a chave não existir
    data_transferencia = str(dados_completos.get('Data_de_Transferencia_Carta', ''))
    banco_carta = str(dados_completos.get('Banco_Carta', ''))
    agencia_carta = str(dados_completos.get('Agencia_Carta', ''))
    conta_corrente_carta = str(dados_completos.get('Conta_Corrente_Carta', ''))

    substituicoes = {
        # Dados do participante (já existem)
        "{{NOME_PARTICIPANTE}}": str(dados_completos.get('Nome', '')),
        # CORREÇÃO AQUI: Garante que 'Complemento' é string antes de chamar replace
        "{{ENDERECO_COMPLEMENTO}}": str(dados_completos.get('Complemento', '')).replace('nan', ''), 
        "{{ENDERECO_RUA}}": str(dados_completos.get('Rua', '')),
        "{{ENDERECO_BAIRRO}}": str(dados_completos.get('Bairro', '')),
        "{{ENDERECO_CEP}}": str(dados_completos.get('CEP', '')),
        "{{ENDERECO_CIDADE_UF}}": f"{dados_completos.get('Cidade', '')} - {dados_completos.get('UF', '')}",
        "{{ASSUNTO_PLANO}}": str(dados_completos.get('Plano', '')), # Plano original
        # Dados específicos da carta de portabilidade (inputs do usuário)
        "{{DATA_DE_TRANSFERENCIA}}": data_transferencia,
        "{{BANCO}}": banco_carta,
        "{{AGENCIA}}": agencia_carta,
        "{{CONTA_CORRENTE}}": conta_corrente_carta,
        "{{N_Ref}}": str(dados_completos.get('NRefDoc', '')), # Reutiliza NRefDoc se quiser
        # Data atual para o cabeçalho da carta
        "{{DATA_ATUAL_CARTA}}": datetime.now().strftime("%d de %B de %Y").replace('maio', 'maio'), # Ajuste de mês para português
        # Lembre-se de ajustar 'maio' para o mês atual, se precisar de flexibilidade para todos os meses
        # Ex: .replace('January', 'janeiro').replace('February', 'fevereiro')...
    }
    
    # --- LINHAS DE DEBUG ADICIONADAS PARA CARTA DE PORTABILIDADE ---
    st.write("---")
    st.write("### Debugging: Dicionário de Substituições para Carta de Portabilidade")
    st.json(substituicoes) # Usa st.json para uma visualização mais legível
    st.write("---")
    # ---------------------------------------------------------------

    replace_placeholders_in_document(doc, substituicoes)

    nome_base = str(dados_completos.get('Nome', 'Desconhecido')).replace(' ', '_').replace('/', '_')
    output_docx_path = f"carta_portabilidade_{nome_base}.docx"
    output_pdf_path = f"carta_portabilidade_{nome_base}.pdf"

    try:
        doc.save(output_docx_path)
        st.info(f"Arquivo DOCX '{output_docx_path}' gerado.")
        try:
            convert(output_docx_path, output_pdf_path)
            st.info(f"Arquivo PDF '{output_pdf_path}' gerado.")
            return output_pdf_path, output_docx_path
        except Exception as e_pdf:
            st.warning(f"DOCX gerado, mas falha ao converter para PDF: {e_pdf}")
            st.info("Verifique se o Microsoft Word ou LibreOffice está instalado e se o erro 'pywintypes' foi resolvido (veja instruções).")
            return None, output_docx_path
    except Exception as e_docx:
        st.error(f"Erro ao salvar o documento DOCX: {e_docx}")
        return None, None


# --- STREAMLIT UI ---
# ATUALIZAR AS ABAS AQUI
st.set_page_config(layout="wide", page_title="Gestão de Formulários FRG")
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📥 Formulário Inicial", "📊 Kanban", "📝 Relação de Crédito", "📉 Desconto de Déficit", "📄 Termo de Portabilidade", "📧 Carta de Portabilidade"])


with tab1:
    st.header("📥 Preencha o Formulário Inicial")
    with st.form("form_inicial_tab1"):
        nome_t1 = st.text_input("Nome Completo", key="nome_t1")
        matricula_t1 = st.text_input("Matrícula", key="mat_t1")
        cpf_t1 = st.text_input("CPF", key="cpf_t1")
        email_t1 = st.text_input("Email Contato", key="email_t1")
        comentario_t1 = st.text_area("Comentário", key="com_t1")
        area_t1 = st.selectbox("Área", list(EMAILS_POR_AREA.keys()), key="area_t1_sb")
        enviado_t1 = st.form_submit_button("🚀 Enviar")
        if enviado_t1:
            if nome_t1 and email_t1 and cpf_t1:
                novo_dado = {col: pd.NA for col in carregar_dados().columns}
                novo_dado.update({
                    "Nome": nome_t1, "Matricula": matricula_t1, "CPF": cpf_t1, "Email": email_t1,
                    "Comentário": comentario_t1, "Área": area_t1, "Etapa": "Aguardando Resposta",
                    "MesAnoRelacao": datetime.now().strftime("%b/%y").lower(),
                    "Deficit2014": "0,00", "Deficit2022": "0,00",
                    "Parcela_Participante": "0,00", "Parcela_Patrocinadora": "0,00",
                    "Total_acumulado": "0,00", "Recursos_portados": "0,00", "debito": "0,00",
                    "total_a_ser_portado": "0,00"
                })
                salvar_dados(novo_dado)
                st.success(f"✅ Dados de {nome_t1} salvos!")
                st.rerun()
            else:
                st.warning("⚠️ Preencha Nome, CPF e Email.")

with tab2: # KANBAN
    st.header("📌 Painel Kanban")
    # Certificar que todas as etapas são colunas
    colunas_kanban = st.columns(len(ETAPAS))
    df_kanban = carregar_dados() # Recarrega para garantir dados atualizados

    for i, etapa_k in enumerate(ETAPAS):
        with colunas_kanban[i]:
            # Filtrar por etapa, garantindo que "Etapa" exista
            etapa_df_k = df_kanban[df_kanban["Etapa"] == etapa_k] if "Etapa" in df_kanban.columns else pd.DataFrame()
            st.subheader(f"{etapa_k} ({len(etapa_df_k)})")
            
            # Ordenar por nome para consistência
            etapa_df_k = etapa_df_k.sort_values(by="Nome", ascending=True)

            for idx_k, row_k in etapa_df_k.iterrows():
                key_base_k = f"{row_k.get('Nome','key')}_{idx_k}_{etapa_k.replace(' ','_')}"
                with st.expander(f"{row_k.get('Nome','Sem Nome')} ({row_k.get('Area','N/A')})", expanded=False):
                    st.caption(f"Matrícula: {row_k.get('Matricula', 'N/A')} | CPF: {row_k.get('CPF', 'N/A')}")
                    
                    # Botões de transição
                    if etapa_k == "Aguardando Resposta":
                        if st.button("✅ Respondido", key=f"resp_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Respondido"); st.rerun()
                    elif etapa_k == "Respondido":
                        if st.button("➡️ Relação Crédito", key=f"rel_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Relação de Crédito"); st.rerun()
                        if st.button("➡️ Termo Portabilidade", key=f"port_k_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Termo de Portabilidade"); st.rerun()
                        if st.button("➡️ Carta de Portabilidade", key=f"carta_k_{key_base_k}"): # NOVO BOTÃO
                            atualizar_etapa(row_k["Nome"], "Carta de Portabilidade"); st.rerun()
                    elif etapa_k == "Relação de Crédito":
                        if st.button("➡️ Desconto Déficit", key=f"desc_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Desconto de quitação de deficit"); st.rerun()
                        if st.button("⏪ Voltar para Respondido", key=f"volt_resp_rel_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Respondido"); st.rerun()
                    elif etapa_k == "Desconto de quitação de deficit":
                        st.info("Preencher na Aba 'Desconto de Déficit'")
                        if st.button("➡️ Termo Portabilidade", key=f"next_port_from_desc_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Termo de Portabilidade"); st.rerun()
                        if st.button("⏪ Voltar para Relação de Crédito", key=f"volt_rel_desc_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Relação de Crédito"); st.rerun()
                    elif etapa_k == "Termo de Portabilidade":
                        st.info("Preencher na Aba 'Termo de Portabilidade'")
                        if st.button("➡️ Carta de Portabilidade", key=f"port_to_carta_{key_base_k}"): # Transição para Carta
                            atualizar_etapa(row_k["Nome"], "Carta de Portabilidade"); st.rerun()
                        if st.button("⏪ Voltar para Respondido", key=f"volt_resp_port_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Respondido"); st.rerun()
                    elif etapa_k == "Carta de Portabilidade": # NOVA ETAPA NO KANBAN
                        st.info("Preencher na Aba 'Carta de Portabilidade'")
                        if st.button("⏪ Voltar para Termo Portabilidade", key=f"volt_term_carta_{key_base_k}"):
                            atualizar_etapa(row_k["Nome"], "Termo de Portabilidade"); st.rerun()


with tab3: # RELAÇÃO DE CRÉDITO
    st.header("📝 Detalhes da Relação de Crédito")
    df_rel = carregar_dados()
    fase_df_rel = df_rel[df_rel["Etapa"] == "Relação de Crédito"] if "Etapa" in df_rel.columns else pd.DataFrame()
    if 'pdf_file_rc' not in st.session_state: st.session_state.pdf_file_rc = None
    if 'pdf_label_rc' not in st.session_state: st.session_state.pdf_label_rc = ""

    if not fase_df_rel.empty:
        nomes_rel = fase_df_rel["Nome"].unique()
        pessoa_rel_key = 'sel_pessoa_rel_tab3'
        if pessoa_rel_key not in st.session_state or st.session_state[pessoa_rel_key] not in nomes_rel:
            st.session_state[pessoa_rel_key] = nomes_rel[0]
        pessoa_rel = st.selectbox("Pessoa:", nomes_rel, key=pessoa_rel_key)
        dados_pessoa_rel = fase_df_rel[fase_df_rel["Nome"] == pessoa_rel].iloc[0].to_dict()

        with st.form(f"form_rel_{pessoa_rel.replace(' ','_')}"):
            st.subheader(f"Dados para {pessoa_rel}")
            c1, c2 = st.columns(2)
            # Coluna 1
            dados_adicionais_val = str(dados_pessoa_rel.get("Dados Adicionais", ""))
            creditar_val = str(dados_pessoa_rel.get("Creditar", "Banco Bradesco S.A."))
            banco_val = str(dados_pessoa_rel.get("Banco", "237"))
            conta_val = str(dados_pessoa_rel.get("Conta", ""))
            agencia_val = str(dados_pessoa_rel.get("Agencia", ""))
            nome_agencia_val = str(dados_pessoa_rel.get("NomeAgencia", ""))
            valor_rs_val = str(dados_pessoa_rel.get("ValorRS", "0,00"))
            with c1:
                dado_extra_rc_t3 = st.text_area("Info Adicionais (Portabilidade)", value=dados_adicionais_val, key=f"dado_extra_rc_t3_{pessoa_rel}")
                creditar_rc_t3 = st.text_input("Banco a Creditar", value=creditar_val, key=f"creditar_rc_t3_{pessoa_rel}")
                banco_rc_t3 = st.text_input("Cód. Banco", value=banco_val, key=f"banco_rc_t3_{pessoa_rel}")
                conta_rc_t3 = st.text_input("Conta", value=conta_val, key=f"conta_rc_t3_{pessoa_rel}")
            # Coluna 2
            agencia_rc_t3 = st.text_input("Cód. Agência", value=agencia_val, key=f"agencia_rc_t3_{pessoa_rel}")
            nome_agencia_rc_t3 = st.text_input("Nome Agência", value=nome_agencia_val, key=f"nome_agencia_rc_t3_{pessoa_rel}")
            valor_rs_rc_t3 = st.text_input("Valor Total R$", value=valor_rs_val, key=f"valor_rs_rc_t3_{pessoa_rel}")
            tipo_entidade_val = str(dados_pessoa_rel.get("TipoEntidade", "Fechada"))
            patrocinadora_val = str(dados_pessoa_rel.get("Patrocinadora", "FURNAS"))
            plano_val = str(dados_pessoa_rel.get("Plano", "CONTRIBUIÇÃO DEFINIDA - CD"))
            qtde_cotas_val = str(dados_pessoa_rel.get("QtdeCotas", "0,00"))
            valor_cota_val = str(dados_pessoa_rel.get("ValorCota", "0,00"))
            data_vc_val = str(dados_pessoa_rel.get("DataValorCota", "dd/mm/aaaa"))
            with c2:
                tipo_entidade_rc_t3 = st.text_input("Tipo Entidade", value=tipo_entidade_val, key=f"tipo_entidade_rc_t3_{pessoa_rel}")
                patrocinadora_rc_t3 = st.text_input("Patrocinadora", value=patrocinadora_val, key=f"patrocinadora_rc_t3_{pessoa_rel}")
                plano_rc_t3 = st.text_input("Plano", value=plano_val, key=f"plano_rc_t3_{pessoa_rel}")
                qtde_cotas_rc_t3 = st.text_input("Qtde Cotas", value=qtde_cotas_val, help="Ex: 12345,67", key=f"qtde_cotas_rc_t3_{pessoa_rel}")
                valor_cota_rc_t3 = st.text_input("Valor Cota", value=valor_cota_val, help="Ex: 12,345678", key=f"valor_cota_rc_t3_{pessoa_rel}")
                data_vc_rc_t3 = st.text_input("Data Base Cota", value=data_vc_val, key=f"data_vc_rc_t3_{pessoa_rel}")

            mes_ano_rel_val = str(dados_pessoa_rel.get("MesAnoRelacao", datetime.now().strftime("%b/%y").lower()))
            data_pag_val = str(dados_pessoa_rel.get("DataPagamento", "dd/mm/aaaa"))
            mes_ano_rel_rc_t3 = st.text_input("Mês/Ano Relação", value=mes_ano_rel_val, key=f"mes_ano_rel_rc_t3_{pessoa_rel}")
            data_pag_rc_t3 = st.text_input("Data Pagamento PDF", value=data_pag_val, key=f"data_pag_rc_t3_{pessoa_rel}")
            
            submitted_rc_t3 = st.form_submit_button("💾 Salvar e Gerar PDF (Relação)")
            if submitted_rc_t3:
                dados_save_rc = {
                    "Dados Adicionais": dado_extra_rc_t3, "Creditar": creditar_rc_t3, "Banco": banco_rc_t3,
                    "Conta": conta_rc_t3, "Agencia": agencia_rc_t3, "NomeAgencia": nome_agencia_rc_t3,
                    "ValorRS": valor_rs_rc_t3, "TipoEntidade": tipo_entidade_rc_t3, "Patrocinadora": patrocinadora_rc_t3,
                    "Plano": plano_rc_t3, "QtdeCotas": qtde_cotas_rc_t3, "ValorCota": valor_cota_rc_t3,
                    "DataValorCota": data_vc_rc_t3, "MesAnoRelacao": mes_ano_rel_rc_t3, "DataPagamento": data_pag_rc_t3
                }
                dados_att_rc = salvar_dados_completos(pessoa_rel, dados_save_rc)
                if dados_att_rc:
                    pdf_path_rc = gerar_pdf_relacao_credito(dados_att_rc)
                    st.success(f"✅ Dados salvos para {pessoa_rel}!")
                    if pdf_path_rc and os.path.exists(pdf_path_rc):
                        st.session_state.pdf_file_rc = pdf_path_rc
                        st.session_state.pdf_label_rc = f"📥 Baixar PDF Relação ({pessoa_rel})"
                else: st.error("Falha ao salvar dados.")
    else: st.info("Nenhuma pessoa em 'Relação de Crédito'.")

    if st.session_state.get('pdf_file_rc') and os.path.exists(st.session_state.pdf_file_rc):
        with open(st.session_state.pdf_file_rc, "rb") as f:
            st.download_button(st.session_state.pdf_label_rc, f, os.path.basename(st.session_state.pdf_file_rc), "application/pdf", key="dl_pdf_rc_btn", on_click=lambda: setattr(st.session_state, 'pdf_file_rc', None))


with tab4: # DESCONTO DE DÉFICIT
    st.header("📉 Desconto de Déficit e Documento de Quitação")
    df_desc = carregar_dados()
    fase_df_desc = df_desc[df_desc["Etapa"] == "Desconto de quitação de deficit"] if "Etapa" in df_desc.columns else pd.DataFrame()

    if 'pdf_file_quit' not in st.session_state: st.session_state.pdf_file_quit = None
    if 'docx_file_quit' not in st.session_state: st.session_state.docx_file_quit = None
    if 'label_quit' not in st.session_state: st.session_state.label_quit = ""

    if not fase_df_desc.empty:
        nomes_desc = fase_df_desc["Nome"].unique()
        pessoa_desc_key = 'sel_pessoa_desc_tab4'
        if pessoa_desc_key not in st.session_state or st.session_state[pessoa_desc_key] not in nomes_desc:
            st.session_state[pessoa_desc_key] = nomes_desc[0]
        pessoa_desc = st.selectbox("Pessoa:", nomes_desc, key=pessoa_desc_key)
        dados_pessoa_desc = fase_df_desc[fase_df_desc["Nome"] == pessoa_desc].iloc[0].to_dict()

        with st.form(f"form_desc_{pessoa_desc.replace(' ','_')}"):
            st.subheader(f"Dados para Documento de Quitação: {pessoa_desc}")
            n_ref_t4 = st.text_input("N.Ref (Doc):", value=str(dados_pessoa_desc.get("NRefDoc", "")), key=f"nref_t4_{pessoa_desc}")
            c1_t4, c2_t4 = st.columns(2)
            with c1_t4:
                rua_t4 = st.text_input("Rua:", value=str(dados_pessoa_desc.get("Rua", "")), key=f"rua_t4_{pessoa_desc}")
                comp_t4 = st.text_input("Complemento:", value=str(dados_pessoa_desc.get("Complemento", "")), key=f"comp_t4_{pessoa_desc}")
                bairro_t4 = st.text_input("Bairro:", value=str(dados_pessoa_desc.get("Bairro", "")), key=f"bairro_t4_{pessoa_desc}")
            with c2_t4:
                cep_t4 = st.text_input("CEP:", value=str(dados_pessoa_desc.get("CEP", "")), key=f"cep_t4_{pessoa_desc}")
                cidade_t4 = st.text_input("Cidade:", value=str(dados_pessoa_desc.get("Cidade", "")), key=f"cidade_t4_{pessoa_desc}")
                uf_t4 = st.text_input("UF:", value=str(dados_pessoa_desc.get("UF", "")), max_chars=2, key=f"uf_t4_{pessoa_desc}")
            
            mes_calc_t4 = st.text_input("Mês Cálculo Cota (Doc.) (Ex: abril/2025):", value=str(dados_pessoa_desc.get("MesCalculoCotaDoc", "")), key=f"mes_calc_t4_{pessoa_desc}")
            
            st.markdown("**Valores de Déficit (Use somente números. Para decimais, use vírgula, ex: 20000 ou 2000,50):**")
            def14_t4 = st.text_input("Déficit 2014 (R$):", value=str(dados_pessoa_desc.get("Deficit2014", "0,00")), key=f"def14_t4_{pessoa_desc}")
            def22_t4 = st.text_input("Déficit 2022 (R$):", value=str(dados_pessoa_desc.get("Deficit2022", "0,00")), key=f"def22_t4_{pessoa_desc}")

            submitted_desc_t4 = st.form_submit_button("💾 Salvar e Gerar Documentos (Quitação)")
            if submitted_desc_t4:
                val_def14_float_debug = desformatar_string_para_float(def14_t4)
                val_def22_float_debug = desformatar_string_para_float(def22_t4)
                
                dados_save_desc = {
                    "NRefDoc": n_ref_t4, "Rua": rua_t4, "Complemento": comp_t4, "Bairro": bairro_t4,
                    "CEP": cep_t4, "Cidade": cidade_t4, "UF": uf_t4, "MesCalculoCotaDoc": mes_calc_t4,
                    "Deficit2014": formatar_moeda_para_exibicao(val_def14_float_debug), # Salva formatado
                    "Deficit2022": formatar_moeda_para_exibicao(val_def22_float_debug)  # Salva formatado
                }
                dados_att_desc = salvar_dados_completos(pessoa_desc, dados_save_desc)
                if dados_att_desc:
                    pdf_p, docx_p = gerar_documento_quitacao(dados_att_desc)
                    st.success(f"✅ Dados salvos para {pessoa_desc}!")
                    st.session_state.label_quit = f"Docs Quitação ({pessoa_desc})"
                    st.session_state.pdf_file_quit = pdf_p if pdf_p and os.path.exists(pdf_p) else None
                    st.session_state.docx_file_quit = docx_p if docx_p and os.path.exists(docx_p) else None
                    if not pdf_p and docx_p: st.warning("PDF não gerado, DOCX disponível.")
                    elif not pdf_p and not docx_p: st.error("Falha ao gerar docs.")
                else: st.error("Falha ao salvar dados.")
    else: st.info("Nenhuma pessoa em 'Desconto de quitação de deficit'.")

    if st.session_state.get('pdf_file_quit') and os.path.exists(st.session_state.pdf_file_quit):
        with open(st.session_state.pdf_file_quit, "rb") as f:
            st.download_button("📥 Baixar PDF Quitação", f, os.path.basename(st.session_state.pdf_file_quit), "application/pdf", key="dl_pdf_quit_btn", on_click=lambda: setattr(st.session_state, 'pdf_file_quit', None))
    if st.session_state.get('docx_file_quit') and os.path.exists(st.session_state.docx_file_quit):
        with open(st.session_state.docx_file_quit, "rb") as f:
            st.download_button("📥 Baixar DOCX Quitação", f, os.path.basename(st.session_state.docx_file_quit), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_docx_quit_btn", on_click=lambda: setattr(st.session_state, 'docx_file_quit', None))
        # Limpar label após botões serem exibidos e clicados
        # Este on_click lambda irá limpar o estado quando o botão for clicado, mas
        # você precisa recarregar a página ou fazer algo para o botão desaparecer.
        # Uma alternativa é usar um callback de submit que limpa o estado.

with tab5: # TERMO DE PORTABILIDADE - NOVA ABA
    st.header("📄 Termo de Portabilidade")
    df_port = carregar_dados()
    fase_df_port = df_port[df_port["Etapa"] == "Termo de Portabilidade"] if "Etapa" in df_port.columns else pd.DataFrame()

    if 'pdf_file_port' not in st.session_state: st.session_state.pdf_file_port = None
    if 'docx_file_port' not in st.session_state: st.session_state.docx_file_port = None
    if 'label_port' not in st.session_state: st.session_state.label_port = ""

    if not fase_df_port.empty:
        nomes_port = fase_df_port["Nome"].unique()
        pessoa_port_key = 'sel_pessoa_port_tab5'
        if pessoa_port_key not in st.session_state or st.session_state[pessoa_port_key] not in nomes_port:
            st.session_state[pessoa_port_key] = nomes_port[0]
        pessoa_port = st.selectbox("Pessoa:", nomes_port, key=pessoa_port_key)
        dados_pessoa_port = fase_df_port[fase_df_port["Nome"] == pessoa_port].iloc[0].to_dict()

        with st.form(f"form_port_{pessoa_port.replace(' ','_')}"):
            st.subheader(f"Dados para Termo de Portabilidade: {pessoa_port}")

            st.markdown("### Dados do Participante")
            c_part1, c_part2 = st.columns(2)
            with c_part1:
                data_admissao_t5 = st.text_input("Data de Admissão (dd/mm/aaaa):", value=str(dados_pessoa_port.get("Data_admissao", "")), key=f"data_adm_t5_{pessoa_port}")
                data_desligamento_t5 = st.text_input("Data de Desligamento (dd/mm/aaaa):", value=str(dados_pessoa_port.get("Data_desligamento", "")), key=f"data_des_t5_{pessoa_port}")
            with c_part2:
                data_inscricao_t5 = st.text_input("Data de Inscrição no Plano (dd/mm/aaaa):", value=str(dados_pessoa_port.get("Data_inscricao", "")), key=f"data_ins_t5_{pessoa_port}")
                # Reuso de campos de endereço já existentes
                st.info(f"Endereço: {str(dados_pessoa_port.get('Rua', '')).replace('nan','')}, {str(dados_pessoa_port.get('Complemento', '')).replace('nan', '')} - {str(dados_pessoa_port.get('Bairro', '')).replace('nan', '')}, {str(dados_pessoa_port.get('Cidade', '')).replace('nan', '')} - {str(dados_pessoa_port.get('UF', '')).replace('nan', '')}, CEP: {str(dados_pessoa_port.get('CEP', '')).replace('nan', '')}") # Garantindo strings
            
            st.markdown("### Dados do Plano Receptor")
            c_pr1, c_pr2 = st.columns(2)
            with c_pr1:
                plano_beneficio_t5 = st.text_input("Nome do Plano de Benefício Receptor:", value=str(dados_pessoa_port.get("plano_de_beneficio", "")), key=f"plano_ben_t5_{pessoa_port}")
                cnpb_t5 = st.text_input("CNPB (Plano Receptor):", value=str(dados_pessoa_port.get("CNPB", "")), key=f"cnpb_t5_{pessoa_port}")
                plano_receptor_t5 = st.text_input("Nome do Plano Receptor:", value=str(dados_pessoa_port.get("plano_receptor", "")), key=f"plano_rec_t5_{pessoa_port}")
                cnpj_plano_receptor_t5 = st.text_input("CNPJ do Plano Receptor:", value=str(dados_pessoa_port.get("cnpj_plano_receptor", "")), key=f"cnpj_pr_t5_{pessoa_port}")
                endereco_plano_receptor_t5 = st.text_input("Endereço do Plano Receptor:", value=str(dados_pessoa_port.get("endereco_plano_receptor", "")), key=f"end_pr_t5_{pessoa_port}")
                cep_plano_receptor_t5 = st.text_input("CEP do Plano Receptor:", value=str(dados_pessoa_port.get("cep_plano_receptor", "")), key=f"cep_pr_t5_{pessoa_port}")
            with c_pr2:
                cidade_plano_receptor_t5 = st.text_input("Cidade-UF do Plano Receptor:", value=str(dados_pessoa_port.get("cidade_plano_receptor", "")), key=f"cidade_pr_t5_{pessoa_port}")
                contato_plano_receptor_t5 = st.text_input("Contato do Plano Receptor:", value=str(dados_pessoa_port.get("contato_plano_receptor", "")), key=f"cont_pr_t5_{pessoa_port}")
                telefone_plano_receptor_t5 = st.text_input("Telefone do Plano Receptor:", value=str(dados_pessoa_port.get("telefone_plano_receptor", "")), key=f"tel_pr_t5_{pessoa_port}")
                email_plano_receptor_t5 = st.text_input("Email do Plano Receptor:", value=str(dados_pessoa_port.get("email_plano_receptor", "")), key=f"email_pr_t5_{pessoa_port}")
                banco_plano_receptor_t5 = st.text_input("Nome - N.º do Banco:", value=str(dados_pessoa_port.get("banco_plano_receptor", "")), key=f"banco_pr_t5_{pessoa_port}")
                agencia_plano_receptor_t5 = st.text_input("Agência - N.º / Nome / Cidade / UF:", value=str(dados_pessoa_port.get("agencia_plano_receptor", "")), key=f"ag_pr_t5_{pessoa_port}")
                conta_plano_receptor_t5 = st.text_input("Conta Corrente:", value=str(dados_pessoa_port.get("conta_plano_receptor", "")), key=f"conta_pr_t5_{pessoa_port}")

            st.markdown("### Dados da Portabilidade (Valores em R$)")
            st.info("Use somente números. Para decimais, use vírgula, ex: 12345,67")
            c_port_val1, c_port_val2 = st.columns(2)
            with c_port_val1:
                parcela_participante_t5 = st.text_input("Direito Acumulado - Parcela Participante:", value=str(dados_pessoa_port.get("Parcela_Participante", "0,00")), key=f"pp_t5_{pessoa_port}")
                parcela_patrocinadora_t5 = st.text_input("Direito Acumulado - Parcela Patrocinadora:", value=str(dados_pessoa_port.get("Parcela_Patrocinadora", "0,00")), key=f"ppa_t5_{pessoa_port}")
                regime_tributacao_t5 = st.text_input("Regime de Tributação:", value=str(dados_pessoa_port.get("Regime_de_tributacao", "")), key=f"reg_trib_t5_{pessoa_port}")
                recursos_portados_t5 = st.text_input("Recursos Portados de Entidades Fechadas:", value=str(dados_pessoa_port.get("Recursos_portados", "0,00")), key=f"rec_port_t5_{pessoa_port}")
            with c_port_val2:
                debito_t5 = st.text_input("Débitos junto à Real Grandeza:", value=str(dados_pessoa_port.get("debito", "0,00")), key=f"debito_t5_{pessoa_port}")
                data_base_portabilidade_t5 = st.text_input("Data Base (dd/mm/aaaa):", value=str(dados_pessoa_port.get("Data_base_portabilidade", "")), key=f"data_base_port_t5_{pessoa_port}")

            submitted_port_t5 = st.form_submit_button("💾 Salvar e Gerar Documentos (Portabilidade)")
            if submitted_port_t5:
                # Converter para float para cálculos e depois formatar para salvar no CSV
                val_pp_float = desformatar_string_para_float(parcela_participante_t5)
                val_ppa_float = desformatar_string_para_float(parcela_patrocinadora_t5)
                val_rec_port_float = desformatar_string_para_float(recursos_portados_t5)
                val_debito_float = desformatar_string_para_float(debito_t5)

                # Cálculos para o CSV (salva o resultado dos cálculos, não os inputs puros)
                total_acumulado_calc = val_pp_float + val_ppa_float
                total_a_ser_portado_calc = total_acumulado_calc - val_debito_float

                dados_save_port = {
                    "Data_admissao": data_admissao_t5,
                    "Data_desligamento": data_desligamento_t5,
                    "Data_inscricao": data_inscricao_t5,
                    "plano_de_beneficio": plano_beneficio_t5,
                    "CNPB": cnpb_t5,
                    "plano_receptor": plano_receptor_t5,
                    "cnpj_plano_receptor": cnpj_plano_receptor_t5,
                    "endereco_plano_receptor": endereco_plano_receptor_t5,
                    "cep_plano_receptor": cep_plano_receptor_t5,
                    "cidade_plano_receptor": cidade_plano_receptor_t5,
                    "contato_plano_receptor": contato_plano_receptor_t5,
                    "telefone_plano_receptor": telefone_plano_receptor_t5,
                    "email_plano_receptor": email_plano_receptor_t5,
                    "banco_plano_receptor": banco_plano_receptor_t5,
                    "agencia_plano_receptor": agencia_plano_receptor_t5,
                    "conta_plano_receptor": conta_plano_receptor_t5,
                    "Parcela_Participante": formatar_moeda_para_exibicao(val_pp_float),
                    "Parcela_Patrocinadora": formatar_moeda_para_exibicao(val_ppa_float),
                    "Total_acumulado": formatar_moeda_para_exibicao(total_acumulado_calc),
                    "Regime_de_tributacao": regime_tributacao_t5,
                    "Recursos_portados": formatar_moeda_para_exibicao(val_rec_port_float),
                    "debito": formatar_moeda_para_exibicao(val_debito_float),
                    "total_a_ser_portado": formatar_moeda_para_exibicao(total_a_ser_portado_calc),
                    "Data_base_portabilidade": data_base_portabilidade_t5
                }
                dados_att_port = salvar_dados_completos(pessoa_port, dados_save_port)
                if dados_att_port:
                    pdf_p_port, docx_p_port = gerar_documento_portabilidade(dados_att_port)
                    st.success(f"✅ Dados salvos para {pessoa_port}!")
                    st.session_state.label_port = f"Docs Portabilidade ({pessoa_port})"
                    st.session_state.pdf_file_port = pdf_p_port if pdf_p_port and os.path.exists(pdf_p_port) else None
                    st.session_state.docx_file_port = docx_p_port if docx_p_port and os.path.exists(docx_p_port) else None
                    if not pdf_p_port and docx_p_port: st.warning("PDF não gerado, DOCX disponível.")
                    elif not pdf_p_port and not docx_p_port: st.error("Falha ao gerar docs.")
                else: st.error("Falha ao salvar dados.")
    else: st.info("Nenhuma pessoa em 'Termo de Portabilidade'.")

    if st.session_state.get('label_port'):
        st.subheader(st.session_state.label_port)
        if st.session_state.get('pdf_file_port') and os.path.exists(st.session_state.pdf_file_port):
            with open(st.session_state.pdf_file_port, "rb") as f:
                st.download_button("📥 Baixar PDF Portabilidade", f, os.path.basename(st.session_state.pdf_file_port), "application/pdf", key="dl_pdf_port_btn", on_click=lambda: setattr(st.session_state, 'pdf_file_port', None))
        if st.session_state.get('docx_file_port') and os.path.exists(st.session_state.docx_file_port):
            with open(st.session_state.docx_file_port, "rb") as f:
                st.download_button("📥 Baixar DOCX Portabilidade", f, os.path.basename(st.session_state.docx_file_port), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_docx_port_btn", on_click=lambda: setattr(st.session_state, 'docx_file_port', None))
        if st.session_state.pdf_file_port is None and st.session_state.docx_file_port is None:
               st.session_state.label_port = ""

with tab6: # CARTA DE PORTABILIDADE ENTRE PLANOS (NOVA ABA)
    st.header("📧 Carta de Portabilidade entre Planos")
    df_carta = carregar_dados()
    fase_df_carta = df_carta[df_carta["Etapa"] == "Carta de Portabilidade"] if "Etapa" in df_carta.columns else pd.DataFrame()

    if 'pdf_file_carta' not in st.session_state: st.session_state.pdf_file_carta = None
    if 'docx_file_carta' not in st.session_state: st.session_state.docx_file_carta = None
    if 'label_carta' not in st.session_state: st.session_state.label_carta = ""

    if not fase_df_carta.empty:
        nomes_carta = fase_df_carta["Nome"].unique()
        pessoa_carta_key = 'sel_pessoa_carta_tab6'
        if pessoa_carta_key not in st.session_state or st.session_state[pessoa_carta_key] not in nomes_carta:
            st.session_state[pessoa_carta_key] = nomes_carta[0]
        pessoa_carta = st.selectbox("Pessoa:", nomes_carta, key=pessoa_carta_key)
        dados_pessoa_carta = fase_df_carta[fase_df_carta["Nome"] == pessoa_carta].iloc[0].to_dict()

        with st.form(f"form_carta_{pessoa_carta.replace(' ','_')}"):
            st.subheader(f"Dados para a Carta de Portabilidade: {pessoa_carta}")

            # Exibir dados já existentes do participante
            st.markdown("### Dados do Participante (Pré-preenchidos)")
            st.info(f"Nome: {dados_pessoa_carta.get('Nome', 'N/A')}")
            # --- CORREÇÃO DO ERRO 'float' object has no attribute 'replace' AQUI ---
            st.info(f"Endereço: {str(dados_pessoa_carta.get('Complemento', '')).replace('nan', '')} {str(dados_pessoa_carta.get('Rua', 'N/A')).replace('nan', '')}, {str(dados_pessoa_carta.get('Bairro', 'N/A')).replace('nan', '')}")
            st.info(f"CEP: {str(dados_pessoa_carta.get('CEP', 'N/A')).replace('nan', '')} - {str(dados_pessoa_carta.get('Cidade', 'N/A')).replace('nan', '')} - {str(dados_pessoa_carta.get('UF', 'N/A')).replace('nan', '')}")
            # --- FIM DA CORREÇÃO ---
            st.info(f"Plano de Origem: {dados_pessoa_carta.get('Plano', 'N/A')}")
            st.info(f"N.Ref (Doc): {dados_pessoa_carta.get('NRefDoc', 'N/A')}") # Mostrar N.Ref se for usado

            st.markdown("### Informações Específicas da Transferência (Preencher)")
            # Campos para input do usuário
            data_transferencia_t6 = st.text_input("Data de Transferência (dd/mm/aaaa):", value=str(dados_pessoa_carta.get("Data_de_Transferencia_Carta", "")), key=f"data_transf_t6_{pessoa_carta}")
            banco_t6 = st.text_input("Banco (para o Plano FRGPrev):", value=str(dados_pessoa_carta.get("Banco_Carta", "")), key=f"banco_t6_{pessoa_carta}")
            agencia_t6 = st.text_input("Agência (do Plano FRGPrev):", value=str(dados_pessoa_carta.get("Agencia_Carta", "")), key=f"agencia_t6_{pessoa_carta}")
            conta_corrente_t6 = st.text_input("Conta Corrente (do Plano FRGPrev):", value=str(dados_pessoa_carta.get("Conta_Corrente_Carta", "")), key=f"cc_t6_{pessoa_carta}")

            submitted_carta_t6 = st.form_submit_button("💾 Salvar e Gerar Documentos (Carta de Portabilidade)")
            if submitted_carta_t6:
                dados_save_carta = {
                    "Data_de_Transferencia_Carta": data_transferencia_t6,
                    "Banco_Carta": banco_t6,
                    "Agencia_Carta": agencia_t6,
                    "Conta_Corrente_Carta": conta_corrente_t6
                }
                # Salvar os dados e obter o dicionário completo e atualizado
                dados_att_carta = salvar_dados_completos(pessoa_carta, dados_save_carta)
                
                if dados_att_carta:
                    pdf_p_carta, docx_p_carta = gerar_documento_carta_portabilidade(dados_att_carta)
                    st.success(f"✅ Dados salvos e documentos gerados para {pessoa_carta}!")
                    st.session_state.label_carta = f"Docs Carta de Portabilidade ({pessoa_carta})"
                    st.session_state.pdf_file_carta = pdf_p_carta if pdf_p_carta and os.path.exists(pdf_p_carta) else None
                    st.session_state.docx_file_carta = docx_p_carta if docx_p_carta and os.path.exists(docx_p_carta) else None
                    if not pdf_p_carta and docx_p_carta: st.warning("PDF não gerado, DOCX disponível. Verifique a instalação do Word/LibreOffice.")
                    elif not pdf_p_carta and not docx_p_carta: st.error("Falha ao gerar documentos.")
                else:
                    st.error("Falha ao salvar dados para a Carta de Portabilidade.")
    else: st.info("Nenhuma pessoa na etapa 'Carta de Portabilidade'.")

    if st.session_state.get('label_carta'):
        st.subheader(st.session_state.label_carta)
        if st.session_state.get('pdf_file_carta') and os.path.exists(st.session_state.pdf_file_carta):
            with open(st.session_state.pdf_file_carta, "rb") as f:
                st.download_button("📥 Baixar PDF Carta", f, os.path.basename(st.session_state.pdf_file_carta), "application/pdf", key="dl_pdf_carta_btn", on_click=lambda: setattr(st.session_state, 'pdf_file_carta', None))
        if st.session_state.get('docx_file_carta') and os.path.exists(st.session_state.docx_file_carta):
            with open(st.session_state.docx_file_carta, "rb") as f:
                st.download_button("📥 Baixar DOCX Carta", f, os.path.basename(st.session_state.docx_file_carta), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_docx_carta_btn", on_click=lambda: setattr(st.session_state, 'docx_file_carta', None))
        if st.session_state.pdf_file_carta is None and st.session_state.docx_file_carta is None:
               st.session_state.label_carta = ""


st.sidebar.header("📊 Todos os Dados")
df_todos_sb = carregar_dados()
if st.sidebar.checkbox("Mostrar tabela de dados", True, key="cb_dados_sb"):
    if df_todos_sb.empty: st.sidebar.info("Nenhum dado.")
    else:
        st.sidebar.dataframe(df_todos_sb)
        # LINHA CORRIGIDA PARA O ENCODING AQUI
        csv_sb = df_todos_sb.to_csv(index=False).encode('utf-8-sig')
        st.sidebar.download_button("⬇️ Baixar CSV", csv_sb, "dados_completos.csv", "text/csv", key="dl_csv_sb_btn")