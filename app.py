import streamlit as st
import gspread
import qrcode
import io
import os
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from urllib.parse import quote, unquote
from gspread.exceptions import APIError, SpreadsheetNotFound

# Configurações do Google Sheets
ESCOPO = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

credenciais = ServiceAccountCredentials.from_json_keyfile_dict(
    st.secrets["gcp_service_account"], 
    ESCOPO
)
CLIENTE = gspread.authorize(credenciais)

ID_PLANILHA = '1Zye1EfKONPvGOFd-wYYy7w8UxWi4S0MPmG-2zLWR9WI'

COLUNAS = [
    "NOME", "STATUS", "OS", "CTP", "IMPRESSORA", "MODELO",
    "DATA", "VALOR", "QTD. CHAPA", "C", "M", "Y", "K", "P", "TIPO DE IMP.", "CONFIRMADOR"
]

def acessar_planilha():
    try:
        planilha = CLIENTE.open_by_key(ID_PLANILHA)
        aba = planilha.sheet1  # Acessa a primeira aba da planilha
        return aba
    except SpreadsheetNotFound:
        st.error("Planilha não encontrada! Verifique o ID da planilha.")
        return None
    except APIError as e:
        st.error(f"Erro ao acessar a planilha: {str(e)}")
        return None

def buscar_dados_os(numero_os):
    try:
        aba = acessar_planilha()
        if not aba:
            st.error("Não foi possível acessar a planilha.")
            return None
        
        # Busca todas as linhas da planilha
        dados = aba.get_all_values()
        if not dados or len(dados) < 2:
            st.error("A planilha está vazia ou não tem dados além do cabeçalho.")
            return None
        
        # Normaliza o número da OS para evitar problemas de formatação
        numero_os = str(numero_os).strip()
        
        # Encontra a linha correspondente ao número da OS
        for i, linha in enumerate(dados[1:], start=2):  # Começa da linha 2 (após cabeçalho)
            os_planilha = str(linha[COLUNAS.index("OS")]).strip()
            if os_planilha == numero_os:
                return {
                    "linha": i,
                    "NOME": linha[COLUNAS.index("NOME")],
                    "STATUS": linha[COLUNAS.index("STATUS")],
                    "OS": linha[COLUNAS.index("OS")],
                    "CTP": linha[COLUNAS.index("CTP")],
                    "IMPRESSORA": linha[COLUNAS.index("IMPRESSORA")],
                    "MODELO": linha[COLUNAS.index("MODELO")],
                    "DATA": linha[COLUNAS.index("DATA")],
                    "VALOR": linha[COLUNAS.index("VALOR")],
                    "QTD. CHAPA": linha[COLUNAS.index("QTD. CHAPA")],
                    "C": linha[COLUNAS.index("C")],
                    "M": linha[COLUNAS.index("M")],
                    "Y": linha[COLUNAS.index("Y")],
                    "K": linha[COLUNAS.index("K")],
                    "P": linha[COLUNAS.index("P")],
                    "TIPO DE IMP.": linha[COLUNAS.index("TIPO DE IMP.")],
                    "CONFIRMADOR": linha[COLUNAS.index("CONFIRMADOR")]
                }
        st.error(f"OS '{numero_os}' não encontrada na planilha. Verifique se o número está correto.")
        return None
    except Exception as e:
        st.error(f"Erro ao buscar dados: {str(e)}")
        return None

def gerar_qrcode(numero_os):
    url = f"https://chapa-saida-nicopel.streamlit.app/?os={quote(numero_os)}"
    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(url)
    qr.make(fit=True)
    
    img = qr.make_image(fill='black', back_color='white')
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return buffer.getvalue()

def pagina_principal():
    st.title("🔍 Consulta de Ordem de Serviço")
    numero_os = st.text_input("Digite o número da OS")
    
    if numero_os:
        dados = buscar_dados_os(numero_os)
        if dados:
            st.query_params["os"] = numero_os
            st.rerun()
        else:
            st.warning("OS não encontrada! Verifique o número digitado.")
    
    st.subheader("Ou escaneie o QR Code")
    qr_numero = st.text_input("Número da OS para gerar QR Code")
    if qr_numero:
        qr_image = gerar_qrcode(qr_numero)
        st.image(qr_image, caption=f"QR Code para OS {qr_numero}")

def pagina_confirmacao(numero_os):
    try:
        dados = buscar_dados_os(numero_os)
        if not dados:
            return

        st.title(f"✅ Confirmação de Saída - OS {numero_os}")
        st.write(f"**Produto:** {dados['NOME']}")
        
        with st.form(key='confirmar_saida'):
            nome_confirmador = st.text_input("👤 Seu nome para confirmação")
            
            if st.form_submit_button("Confirmar Saída"):
                aba = acessar_planilha()
                if aba:
                    updates = [
                        (COLUNAS.index("STATUS") + 1, "SAIDA"),
                        (COLUNAS.index("CONFIRMADOR") + 1, nome_confirmador),
                        (COLUNAS.index("DATA") + 1, datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                    ]
                    
                    for col, value in updates:
                        aba.update_cell(dados['linha'], col, value)
                    
                    st.success("Saída confirmada com sucesso!")
                    st.balloons()
                    st.query_params.clear()
                    st.rerun()
                else:
                    st.error("Falha na conexão com a planilha!")
    except APIError as e:
        st.error(f"Erro na API: {e.response.json().get('error', {}).get('message', 'Erro desconhecido')}")
    except Exception as e:
        st.error(f"Erro na confirmação: {str(e)}")

def pagina_detalhes(numero_os):
    try:
        dados = buscar_dados_os(numero_os)
        if not dados:
            return

        st.title(f"📋 Detalhes da OS {numero_os}")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Informações Principais")
            st.metric("Produto", dados['NOME'])
            st.metric("Status", dados['STATUS'])
            st.metric("Data", dados['DATA'])
            st.metric("Confirmado por", dados['CONFIRMADOR'])
        
        with col2:
            st.subheader("Detalhes Técnicos")
            st.metric("Impressora", dados['IMPRESSORA'])
            st.metric("Modelo", dados['MODELO'])
            st.metric("Qtd. Chapas", dados['QTD. CHAPA'])
            st.metric("Tipo de Impressão", dados['TIPO DE IMP.'])
        
        st.subheader("Especificações de Cor")
        cols = st.columns(4)
        cols[0].metric("C", dados['C'])
        cols[1].metric("M", dados['M'])
        cols[2].metric("Y", dados['Y'])
        cols[3].metric("K", dados['K'])
        
        st.subheader("Dados Complementares")
        st.write(f"**CTP:** {dados['CTP']}")
        st.write(f"**Valor:** R$ {dados['VALOR']}")
        st.write(f"**P:** {dados['P']}")
    except Exception as e:
        st.error(f"Erro ao carregar detalhes: {str(e)}")

# Controle de navegação principal
if "os" in st.query_params:
    numero_os = unquote(st.query_params.get("os", [""])[0])
    if numero_os:
        dados = buscar_dados_os(numero_os)
        if dados:
            if dados.get("STATUS") == "SAIDA":
                pagina_detalhes(numero_os)
            else:
                pagina_confirmacao(numero_os)
        else:
            st.error("OS não encontrada na base de dados!")
    else:
        st.error("Parâmetro OS inválido na URL!")
else:
    pagina_principal()
