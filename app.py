import streamlit as st
import gspread
import qrcode
import io
import os
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
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
    "NOME", "STATUS","INVENTARIO 20/05","OS", "CTP", "IMPRESSORA", "MODELO",
    "DATA", "VALOR", "QTD. CHAPA", "C", "M", "Y", "K", "P", "TIPO DE IMP.", "CONFIRMADOR", "DATA DA CONFIRMAÇÃO"
]

FUSO_BRASILIA = pytz.timezone("America/Sao_Paulo")



def acessar_planilha():
    try:
        planilha = CLIENTE.open_by_key(ID_PLANILHA)
        return planilha.sheet1
    except SpreadsheetNotFound:
        st.error("Planilha não encontrada! Verifique o ID.")
        return None
    except APIError as e:
        st.error(f"Erro na API do Google: {e.response.status_code}")
        return None
    except Exception as e:
        st.error(f"Erro inesperado: {str(e)}")
        return None

def buscar_dados_os(numero_os):
    try:
        aba = acessar_planilha()
        if not aba:
            return None
            
        # Converter para string e remover espaços
        numero_os = str(numero_os).strip()
        
        # Procurar em toda a coluna OS (coluna 3)
        celula = aba.find(numero_os, in_column=3)
        
        if not celula:
            st.error(f"OS {numero_os} não encontrada na linha!")
            return None

        valores = aba.row_values(celula.row)
        
        # Preencher valores faltantes
        valores += [''] * (len(COLUNAS) - len(valores))
        
        return {
            **dict(zip(COLUNAS, valores)),
            'linha': celula.row
        }
    except Exception as e:
        st.error(f"Erro na busca: {str(e)}")
        return None

def gerar_qrcode(numero_os):
    try:
        # Codificar parâmetros corretamente
        params = quote(str(numero_os))
        url = f"https://chapa-saida-nicopel.streamlit.app/?os={params}"
        
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(url)
        qr.make(fit=True)
        
        img = qr.make_image(fill_color="black", back_color="white")
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        return img_byte_arr
    except Exception as e:
        st.error(f"Erro ao gerar QR Code: {str(e)}")
        return None
    
def pagina_principal():
    st.title("📤 Sistema de Registro de Saída")
    
    numero_os = st.text_input("🔢 Número da OS", key="os_input")
    
    if st.button("Gerar QR Code", key="gerar_btn"):
        if numero_os:
            with st.spinner("Processando..."):
                dados = buscar_dados_os(numero_os)
                if dados:
                    qr_bytes = gerar_qrcode(numero_os)
                    if qr_bytes:
                        st.session_state.qr_data = {
                            'bytes': qr_bytes,
                            'nome_arquivo': f"OS_{numero_os}_{dados['NOME'].replace(' ', '_')}.png"
                        }
                else:
                    st.session_state.qr_data = None
        else:
            st.warning("Digite o número da OS primeiro!")

    if 'qr_data' in st.session_state and st.session_state.qr_data:
        col1, col2 = st.columns(2)
        with col1:
            st.image(st.session_state.qr_data['bytes'], caption="QR Code para Confirmação")
        with col2:
            st.download_button(
                label="⬇️ Baixar QR Code",
                data=st.session_state.qr_data['bytes'],
                file_name=st.session_state.qr_data['nome_arquivo'],
                mime="image/png"
            )

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
                    aba.update_cell(dados['linha'], COLUNAS.index("STATUS") + 1, "SAIDA")
                    aba.update_cell(dados['linha'], COLUNAS.index("CONFIRMADOR") + 1, nome_confirmador)
                    aba.update_cell(dados['linha'], COLUNAS.index("DATA DA CONFIRMAÇÃO") + 1, datetime.now(FUSO_BRASILIA).strftime("%d/%m/%Y %H:%M:%S"))
                    st.success("Saída confirmada com sucesso!")
                    st.balloons()
                else:
                    st.error("Falha na conexão com a planilha!")
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
            st.metric("Data de Confirmação", dados['DATA DA CONFIRMAÇÃO'])
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
    except Exception as e:
        st.error(f"Erro ao carregar detalhes: {str(e)}")

# ... (mantenha as funções pagina_principal, pagina_confirmacao e pagina_detalhes do código anterior)

# Controle de navegação principal
if "os" in st.query_params:
    numero_os = unquote(st.query_params.get("os", ""))
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

# Limpeza de arquivos temporários
if os.path.exists("temp_qrcodes"):
    for file in os.listdir("temp_qrcodes"):
        if file.endswith(".png"):
            os.remove(os.path.join("temp_qrcodes", file))
