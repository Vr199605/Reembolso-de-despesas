import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
import os
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal de Reembolso - Globus", layout="wide")

# --- REGRAS DA POLÍTICA ---
LIMITES = {
    "REFEIÇÃO VIAGEM (em R$)": 150.0,
    "ESTACIONAMENTO (em R$)": 70.0,
    "BEBIDA ALCOÓLICA (em R$)": 50.0
}

CATEGORIAS = [
    "ESTACIONAMENTO (em R$)", 
    "PEDÁGIO (em R$)", 
    "KM¹ (em qtde)", 
    "REPRESENTAÇÃO (em R$)", 
    "TAXI / UBER (em R$)", 
    "REFEIÇÃO VIAGEM (em R$)", 
    "OUTROS* (em R$)"
]
VALOR_KM = 1.37
ARQUIVO_EXCEL = "base_reembolsos.xlsx"
URL_APP = "https://reembolso-de-despesas.streamlit.app/"

# --- FUNÇÃO PARA GARANTIR QUE O EXCEL EXISTA ---
def garantir_excel():
    if not os.path.exists(ARQUIVO_EXCEL):
        # Cria um DataFrame vazio com as colunas necessárias se o arquivo não existir
        df_vazio = pd.DataFrame(columns=[
            "ID", "Colaborador", "Data_Solic", "Data_Item", 
            "Status", "Categoria", "Valor", "Motivo", 
            "Comentario_Admin", "Caminho_Arquivo"
        ])
        df_vazio.to_excel(ARQUIVO_EXCEL, index=False)

def carregar_dados_do_excel():
    garantir_excel()
    try:
        df = pd.read_excel(ARQUIVO_EXCEL)
        if df.empty:
            return []
        
        solicitacoes = []
        for id_solic in df['ID'].unique():
            rows = df[df['ID'] == id_solic]
            primeira = rows.iloc[0]
            detalhes = []
            for _, r in rows.iterrows():
                detalhes.append({
                    "categoria": r['Categoria'],
                    "valor": r['Valor'],
                    "motivo": r['Motivo'],
                    "data": str(r['Data_Item'])
                })
            solicitacoes.append({
                "id": int(id_solic),
                "Colaborador": primeira['Colaborador'],
                "Data": primeira['Data_Solic'],
                "Detalhes": detalhes,
                "Status": primeira['Status'],
                "CaminhoArquivo": primeira['Caminho_Arquivo'],
                "Comentario": primeira.get('Comentario_Admin', '')
            })
        return solicitacoes
    except:
        return []

# Inicialização do Banco de Dados na Sessão
if 'db_global' not in st.session_state:
    st.session_state.db_global = carregar_dados_do_excel()

# --- FUNÇÕES DE SISTEMA ---
def atualizar_excel():
    todos_itens = []
    for solic in st.session_state.db_global:
        for item in solic['Detalhes']:
            todos_itens.append({
                "ID": solic['id'],
                "Colaborador": solic['Colaborador'],
                "Data_Solic": solic['Data'],
                "Data_Item": item.get('data'),
                "Status": solic['Status'],
                "Categoria": item['categoria'],
                "Valor": item['valor'],
                "Motivo": item['motivo'],
                "Comentario_Admin": solic.get('Comentario', ''),
                "Caminho_Arquivo": solic['CaminhoArquivo']
            })
    df = pd.DataFrame(todos_itens)
    df.to_excel(ARQUIVO_EXCEL, index=False)

def enviar_aviso_ao_gabriel(solicitacao):
    destinatario = "victormoreiraicnv@gmail.com"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"📩 Nova Solicitação de Reembolso: {solicitacao['Colaborador']}"
    corpo = f"Olá Gabriel Coelho,\n\nNova solicitação de {solicitacao['Colaborador']} (ID #{solicitacao['id']}).\nLink para acesso: {URL_APP}"
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)
        server.quit()
    except: pass

def salvar_arquivos_locais(files):
    if not os.path.exists("comprovantes"): os.makedirs("comprovantes")
    paths = []
    for file in files:
        path = os.path.join("comprovantes", f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.name}")
        with open(path, "wb") as f: f.write(file.getbuffer())
        paths.append(path)
    return ";".join(paths)

# --- INICIALIZAÇÃO DE UI ---
if 'items_reembolso' not in st.session_state: 
    st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()}]
if 'nome_colab' not in st.session_state:
    st.session_state.nome_colab = ""

# --- INTERFACE ---
aba_guia, aba_colab, aba_admin = st.tabs(["📖 Guia de Preenchimento", "🚀 Solicitar Reembolso", "🔑 Verificação e Aprovação (Gabriel)"])

with aba_guia:
    st.title("📖 Como solicitar seu reembolso")
    st.markdown("Siga o passo a passo para garantir que sua solicitação seja processada rapidamente.")
    st.write("Caso ainda tenha alguma dúvida segue abaixo o manual de politicas de viagens e reembolso.")
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    if os.path.exists(caminho_manual):
        with open(caminho_manual, "rb") as f:
            st.download_button("📥 BAIXAR MANUAL DE REEMBOLSO (PDF)", f, file_name="manual_reembolso.pdf")
    st.info("💡 Ao clicar em enviar, o Gabriel Coelho será notificado com o link para análise.")

with aba_colab:
    st.header("Formulário de Reembolso - Globus")
    st.session_state.nome_colab = st.text_input("Nome Completo", value=st.session_state.nome_colab, placeholder="Digite seu nome aqui...")
    st.markdown("---")
    
    for i, item in enumerate(st.session_state.items_reembolso):
        c1, c2, c3, c4, c5 = st.columns([1.2, 1.8, 1.2, 1.8, 0.4])
        dt_val = item.get('data', datetime.now())
        if isinstance(dt_val, str): dt_val = datetime.strptime(dt_val.split()[0], "%Y-%m-%d")
            
        item['data'] = c1.date_input(f"Data {i}", value=dt_val, key=f"d_{i}")
        item['categoria'] = c2.selectbox(f"Categoria {i}", CATEGORIAS, key=f"c_{i}")
        
        if item['categoria'] == "KM¹ (em qtde)":
            km = c3.number_input("Qtd KM", min_value=0, key=f"v_{i}", value=0)
            item['valor'] = round(km * VALOR_KM, 2)
            c3.markdown(f"<div style='background-color: #f0f2f6; padding: 10px; border-radius: 5px; text-align: center; font-weight: bold;'>R$ {item['valor']:.2f}</div>", unsafe_allow_html=True)
        else:
            item['valor'] = c3.number_input(f"Valor R$ {i}", min_value=0.0, format="%.2f", key=f"v_{i}", value=0.0)
            
        item['motivo'] = c4.text_input(f"Motivo {i}", key=f"m_{i}", placeholder="Obrigatório", label_visibility="collapsed")
        if c5.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()

    if st.button("➕ Adicionar Outro Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()})
        st.rerun()
        
    arquivos = st.file_uploader("Anexar Comprovantes (Obrigatório)", accept_multiple_files=True, key="up_colab")
    
    if st.button("ENVIAR PARA VERIFICAÇÃO"):
        if st.session_state.nome_colab and arquivos and any(it['valor'] > 0 for it in st.session_state.items_reembolso):
            caminhos = salvar_arquivos_locais(arquivos)
            detalhes = []
            for it in st.session_state.items_reembolso:
                d = it.copy()
                d['data'] = d['data'].strftime("%Y-%m-%d") if not isinstance(d['data'], str) else d['data']
                detalhes.append(d)
                
            nova_solic = {
                "id": len(st.session_state.db_global) + 1,
                "Colaborador": st.session_state.nome_colab,
                "Data": datetime.now().strftime("%Y-%m-%d"),
                "Detalhes": detalhes,
                "Status": "Em Verificação",
                "CaminhoArquivo": caminhos
            }
            
            st.session_state.db_global.append(nova_solic)
            atualizar_excel()
            enviar_aviso_ao_gabriel(nova_solic)
            
            # --- RESET TOTAL ---
            st.session_state.nome_colab = ""
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()}]
            st.success("✅ Sua solicitação foi enviada para análise com sucesso!")
            st.balloons()
            st.rerun()

with aba_admin:
    st.title("⌛ Solicitações Pendentes")
    if st.text_input("Senha", type="password") == "globus2026":
        st.session_state.db_global = carregar_dados_do_excel()
        pendentes = [s for s in st.session_state.db_global if s['Status'] == "Em Verificação"]
        
        if not pendentes:
            st.info("Não há solicitações aguardando aprovação.")
            if st.button("🔄 Atualizar"): st.rerun()
        
        for idx, solic in enumerate(pendentes):
            with st.expander(f"📌 ID #{solic['id']} - {solic['Colaborador']}"):
                for i_it, it in enumerate(solic['Detalhes']):
                    ec1, ec2, ec3, ec4 = st.columns(4)
                    it['data'] = ec1.text_input(f"Data_{idx}_{i_it}", it['data'])
                    it['categoria'] = ec2.selectbox(f"Cat_{idx}_{i_it}", CATEGORIAS, index=CATEGORIAS.index(it['categoria']))
                    it['valor'] = ec3.number_input(f"Val_{idx}_{i_it}", value=float(it['valor']))
                    it['motivo'] = ec4.text_input(f"Mot_{idx}_{i_it}", it['motivo'])
                
                dec = st.radio(f"Decisão #{solic['id']}", ["Aprovado", "Reprovado"], horizontal=True)
                if st.button(f"Finalizar #{solic['id']}", key=f"btn_{idx}"):
                    solic['Status'] = dec
                    atualizar_excel()
                    st.success("Processado!")
                    st.rerun()
