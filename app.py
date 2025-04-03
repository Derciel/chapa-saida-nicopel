import streamlit as st
import gspread
import qrcode
import io
import os
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from urllib.parse import quote, unquote
from gspread.exceptions import APIError, SpreadsheetNotFound

# Configurações do Google Sheets (Atualizado)
ESCOPO = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Carregar credenciais corretamente via Secrets (Correção crucial)
credenciais = ServiceAccountCredentials.from_json_keyfile_dict(
    st.secrets["gcp_service_account"], 
    ESCOPO
)
CLIENTE = gspread.authorize(credenciais)  # Nome da variável corrigido

ID_PLANILHA = '1Zye1EfKONPvGOFd-wYYy7w8UxWi4S0MPmG-2zLWR9WI'

COLUNAS = [
    "NOME", "STATUS", "OS", "CTP", "IMPRESSORA", "MODELO",
    "DATA", "VALOR", "QTD. CHAPA", "C", "M", "Y", "K", "P", "TIPO DE IMP.", "CONFIRMADOR"
]

def acessar_planilha():
    try:
        planilha = CLIENTE.open_by_key(ID_PLANILHA)
        return planilha.sheet1
    except SpreadsheetNotFound:
        st.error("Planilha não encontrada! Verifique o ID.")
        return None
    except APIError as e:
        error_msg = e.response.json().get('error', {}).get('message', 'Erro desconhecido')
        st.error(f"Erro na API do Google: {error_msg}")
        return None
    except Exception as e:
        st.error(f"Erro inesperado: {str(e)}")
        return None

def buscar_dados_os(numero_os):
    try:
        aba = acessar_planilha()
        if not aba:
            return None
            
        # Normalização avançada
        numero_os = (
            str(numero_os)
            .strip()
            .upper()
            .replace("OS", "")
            .replace("#", "")
            .strip()
        )
        
        # Remover zeros à esquerda se for numérico
        if numero_os.isdigit():
            numero_os = str(int(numero_os))
            
        # Busca com tolerância
        celula = None
        try:
            celula = aba.find(numero_os, in_column=3)
        except gspread.exceptions.CellNotFound:
            pass
            
        if not celula:
            # Tentar busca parcial
            todas_oss = aba.col_values(3)
            matches = [os for os in todas_oss if numero_os in os]
            
            if matches:
                celula = aba.find(matches[0], in_column=3)
            else:
                st.error(f"OS {numero_os} não encontrada na planilha!")
                return None

        
def gerar_qrcode(numero_os):
    try:
        
        APP_URL = "https://chapa-saida-nicopel.streamlit.app/"
        params = quote(str(numero_os))
        url = f"{APP_URL}?os={params}"
        
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

# ... (mantenha as funções pagina_principal, pagina_confirmacao e pagina_detalhes)
def pagina_principal():
    st.title("📤 Sistema de Registro de Saída de Chapas")
    
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
            
        # Verificar se já teve saída
        if dados.get("STATUS") == "SAIDA":
            data_saida = dados['DATA'] if dados['DATA'] else 'data não registrada'
            confirmador = dados['CONFIRMADOR'] if dados['CONFIRMADOR'] else 'usuário desconhecido'
            st.error(f"⚠️ ATENÇÃO: OS {numero_os} já teve saída confirmada!")
            st.write(f"**Data/Hora:** {data_saida}")
            st.write(f"**Confirmado por:** {confirmador}")
            st.write("---")
            if st.button("↩️ Voltar para o Início"):
                st.query_params.clear()
            return
            
        st.title(f"✅ Confirmação de Saída - OS {numero_os}")
        st.write(f"**Produto:** {dados['NOME']}")
        
        with st.form(key='confirmar_saida'):
            nome_confirmador = st.text_input("👤 Seu nome para confirmação", placeholder="Digite seu nome completo")
            
            if st.form_submit_button("✔️ Confirmar Saída"):
                if not nome_confirmador.strip():
                    st.warning("Por favor, digite seu nome para confirmar!")
                    return
                
                aba = acessar_planilha()
                if aba:
                    updates = [
                        (COLUNAS.index("STATUS") + 1, "SAIDA"),
                        (COLUNAS.index("CONFIRMADOR") + 1, nome_confirmador.strip()),
                        (COLUNAS.index("DATA") + 1, datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                    ]
                    
                    for col, value in updates:
                        aba.update_cell(dados['linha'], col, value)
                        
                    st.success("✅ Saída confirmada com sucesso!")
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
    numero_os = unquote(st.query_params.get("os", [""])[0])  # Formato novo
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

