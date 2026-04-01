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

# --- INOVAÇÃO: MEMÓRIA COMPARTILHADA ---
@st.cache_resource
def iniciar_db_global():
    return []

db_global = iniciar_db_global()

# --- FUNÇÕES DE AUXÍLIO ---
def formatar_moeda(valor):
    if valor is None: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- FUNÇÕES DE SISTEMA ---

def atualizar_excel():
    todos_itens = []
    for solic in db_global:
        for item in solic['Detalhes']:
            todos_itens.append({
                "ID": solic['id'],
                "Colaborador": solic['Colaborador'],
                "Data_Item": item.get('data', solic['Data']),
                "Status": solic['Status'],
                "Categoria": item['categoria'],
                "Valor": item['valor'],
                "Motivo": item['motivo'],
                "Comentario_Admin": solic.get('Comentario', ''),
                "Caminho_Arquivo": solic['CaminhoArquivo']
            })
    if todos_itens:
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
    corpo = f"Olá Gabriel Coelho,\n\nNova solicitação de {solicitacao['Colaborador']} (ID #{solicitacao['id']}).\nLink: {URL_APP}"
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

def enviar_email_automatico(dados, arquivo_pdf, caminhos_arquivos):
    destinatario = "gabriel.coelho@globusseguros.com.br"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"[{dados['Status'].upper()}] Reembolso: {dados['Colaborador']}"
    corpo = f"Solicitação finalizada.\nLink: {URL_APP}"
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        if os.path.exists(arquivo_pdf):
            with open(arquivo_pdf, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(arquivo_pdf))
                msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

def salvar_arquivos_locais(files):
    if not os.path.exists("comprovantes"): os.makedirs("comprovantes")
    paths = []
    for file in files:
        path = os.path.join("comprovantes", f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.name}")
        with open(path, "wb") as f: f.write(file.getbuffer())
        paths.append(path)
    return ";".join(paths)

def gerar_relatorio_pdf(dados, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"RELATÓRIO - {dados['Colaborador']}", styles['Title'])]
    doc.build(elements)

# --- INICIALIZAÇÃO ---
if 'items_reembolso' not in st.session_state: 
    st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()}]
if 'nome_colab' not in st.session_state:
    st.session_state.nome_colab = ""

# --- INTERFACE ---
aba_guia, aba_colab, aba_admin = st.tabs(["📖 Guia de Preenchimento", "🚀 Solicitar Reembolso", "🔑 Verificação e Aprovação (Gabriel)"])

with aba_guia:
    st.title("📖 Guia de Reembolso Globus")
    st.markdown("Siga as instruções para solicitar seu reembolso.")
    st.write("Caso ainda tenha alguma dúvida segue abaixo o manual de politicas de viagens e reembolso.")
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    if os.path.exists(caminho_manual):
        with open(caminho_manual, "rb") as f:
            st.download_button("📥 BAIXAR MANUAL DE REEMBOLSO (PDF)", f, file_name="manual_reembolso.pdf")

with aba_colab:
    st.header("Formulário de Reembolso - Globus")
    st.session_state.nome_colab = st.text_input("Nome Completo", value=st.session_state.nome_colab, placeholder="Digite seu nome completo")
    st.markdown("---")
    
    for i, item in enumerate(st.session_state.items_reembolso):
        c1, c2, c3, c4, c5 = st.columns([1.2, 1.8, 1.2, 1.8, 0.4])
        dt_val = item.get('data', datetime.now())
        if isinstance(dt_val, str): dt_val = datetime.strptime(dt_val, "%d/%m/%Y")
            
        item['data'] = c1.date_input(f"Data {i}", value=dt_val, key=f"d_{i}")
        item['categoria'] = c2.selectbox(f"Categoria {i}", CATEGORIAS, key=f"c_{i}")
        
        if item['categoria'] == "KM¹ (em qtde)":
            km = c3.number_input("Qtd KM", min_value=0, key=f"v_{i}", value=0)
            item['valor'] = round(km * VALOR_KM, 2)
            c3.markdown(f"**{formatar_moeda(item['valor'])}**")
        else:
            # Valor aparece como 0.00 (limpo) para não confundir
            item['valor'] = c3.number_input(f"Valor R$ {i}", min_value=0.0, format="%.2f", key=f"v_{i}", value=0.0)
            
        item['motivo'] = c4.text_input(f"Motivo {i}", key=f"m_{i}", placeholder="Motivo obrigatório")
        
        if c5.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()
            
        if item['valor'] > 0 and item['categoria'] in LIMITES and item['valor'] > LIMITES[item['categoria']]:
            st.warning(f"Limite excedido! Máximo: {formatar_moeda(LIMITES[item['categoria']])}")

    if st.button("➕ Adicionar Outro Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()})
        st.rerun()
        
    arquivos = st.file_uploader("Anexar Comprovantes (Obrigatório)", accept_multiple_files=True, key="up_colab")
    
    if st.button("ENVIAR PARA VERIFICAÇÃO"):
        if st.session_state.nome_colab and arquivos and any(it['valor'] > 0 for it in st.session_state.items_reembolso):
            caminhos = salvar_arquivos_locais(arquivos)
            detalhes_finais = []
            for it in st.session_state.items_reembolso:
                d = it.copy()
                if not isinstance(d['data'], str): d['data'] = d['data'].strftime("%d/%m/%Y")
                detalhes_finais.append(d)
                
            nova_solic = {
                "id": len(db_global) + 1,
                "Colaborador": st.session_state.nome_colab,
                "Data": datetime.now().strftime("%d/%m/%Y"),
                "Detalhes": detalhes_finais,
                "Status": "Em Verificação",
                "CaminhoArquivo": caminhos,
                "Comentario": ""
            }
            
            db_global.append(nova_solic)
            atualizar_excel()
            enviar_aviso_ao_gabriel(nova_solic)
            
            # --- RESET TOTAL ---
            st.session_state.nome_colab = ""
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()}]
            
            st.success("✅ Enviado para análise!")
            st.balloons()
            st.rerun()
        else:
            st.error("Verifique os campos obrigatórios e os valores.")

with aba_admin:
    st.title("⌛ Solicitações Pendentes")
    if st.text_input("Senha", type="password") == "globus2026":
        pendentes = [s for s in db_global if s['Status'] == "Em Verificação"]
        for idx, solic in enumerate(pendentes):
            with st.expander(f"Solicitação #{solic['id']} - {solic['Colaborador']}"):
                for i_it, it in enumerate(solic['Detalhes']):
                    ec1, ec2, ec3, ec4 = st.columns(4)
                    it['data'] = ec1.text_input("Data", it['data'], key=f"ad_{idx}_{i_it}")
                    it['categoria'] = ec2.selectbox("Cat", CATEGORIAS, index=CATEGORIAS.index(it['categoria']), key=f"ac_{idx}_{i_it}")
                    it['valor'] = ec3.number_input("Valor", value=float(it['valor']), key=f"av_{idx}_{i_it}")
                    it['motivo'] = ec4.text_input("Motivo", it['motivo'], key=f"am_{idx}_{i_it}")
                
                if st.button(f"Aprovar #{solic['id']}", key=f"ap_{idx}"):
                    solic['Status'] = "Aprovado"
                    atualizar_excel()
                    st.rerun()
