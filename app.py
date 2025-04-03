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
            
        # Converter para string e remover espaços
        numero_os = str(numero_os).strip()
        
        # Procurar em toda a coluna OS (coluna 3)
        celula = aba.find(numero_os, in_column=3)
        
        if not celula:
            st.error(f"OS {numero_os} não encontrada na linha!")
            return None

        valores = aba.row_values(celula.row)
        
        # Preencher valores faltantes (CORREÇÃO AQUI)
        valores += [''] * (len(COLUNAS) - len(valores))  # <--- Parêntese adicionado
        
        return {
            **dict(zip(COLUNAS, valores)),
            'linha': celula.row
        }
    except Exception as e:
        st.error(f"Erro na busca: {str(e)}")
        return None
        
def gerar_qrcode(numero_os):
    try:
        # URL corrigida para deploy (substitua pelo seu domínio)
        APP_URL = "https://chapa-saida-nicopel.streamlit.app/"  # 👈 ALTERAR PARA SUA URL!
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
                    # Atualização mais robusta
                    updates = [
                        (COLUNAS.index("STATUS") + 1, "SAIDA"),
                        (COLUNAS.index("CONFIRMADOR") + 1, nome_confirmador),
                        (COLUNAS.index("DATA") + 1, datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                    ]
                    
                    for col, value in updates:
                        aba.update_cell(dados['linha'], col, value)
                        
                    st.success("Saída confirmada com sucesso!")
                    st.balloons()
                    st.experimental_rerun()
                else:
                    st.error("Falha na conexão com a planilha!")
    except APIError as e:
        st.error(f"Erro na API: {e.response.json().get('error', {}).get('message', 'Erro desconhecido')}")
    except Exception as e:
        st.error(f"Erro na confirmação: {str(e)}")

# Controle de navegação principal (Atualizado)
query_params = st.experimental_get_query_params()
if "os" in query_params:
    numero_os = unquote(query_params["os"][0])
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

# Removido a limpeza de arquivos temporários (não necessário com BytesIO)
