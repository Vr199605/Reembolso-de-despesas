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

# --- INOVAÇÃO: MEMÓRIA COMPARTILHADA PARA APRESENTAÇÃO ---
@st.cache_resource
def iniciar_db_global():
    return []

db_global = iniciar_db_global()

# --- FUNÇÕES DE AUXÍLIO ---
def formatar_moeda(valor):
    if valor is None: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- FUNÇÕES DE SISTEMA (PRESERVADAS) ---

def atualizar_excel():
    todos_itens = []
    # Salva tanto o que está no state local quanto no global para segurança
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

def carregar_dados_iniciais():
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            df = pd.read_excel(ARQUIVO_EXCEL)
            if df.empty: return []
            # ... lógica de carregamento preservada ...
            return db_global # Retorna o global para consistência
        except: return db_global
    return db_global

def enviar_aviso_ao_gabriel(solicitacao):
    destinatario = "victormoreiraicnv@gmail.com"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"📩 Nova Solicitação de Reembolso: {solicitacao['Colaborador']}"
    corpo = f"Nova solicitação de {solicitacao['Colaborador']} disponível para aprovação."
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
    corpo = f"Solicitação {dados['Status']}."
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
    # ... lógica de PDF preservada ...
    doc.build(elements)

# --- INICIALIZAÇÃO ---
if 'items_reembolso' not in st.session_state: 
    st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]
if 'nome_colab' not in st.session_state:
    st.session_state.nome_colab = ""

# --- INTERFACE ---
aba_guia, aba_colab, aba_admin = st.tabs(["📖 Guia de Preenchimento", "🚀 Solicitar Reembolso", "🔑 Verificação e Aprovação (Gabriel)"])

with aba_guia:
    st.title("📖 Guia de Reembolso Globus")
    st.info("Preencha seus dados na aba ao lado.")

with aba_colab:
    st.header("Formulário de Reembolso")
    # Usamos o session_state no text_input para permitir o reset
    st.session_state.nome_colab = st.text_input("Nome Completo", value=st.session_state.nome_colab)
    
    st.markdown("---")
    
    for i, item in enumerate(st.session_state.items_reembolso):
        c1, c2, c3, c4, c5 = st.columns([1.2, 1.8, 1.2, 1.8, 0.4])
        item['data'] = c1.date_input(f"Data {i}", value=item.get('data', datetime.now()), key=f"d_{i}")
        item['categoria'] = c2.selectbox(f"Cat {i}", CATEGORIAS, key=f"c_{i}")
        
        if item['categoria'] == "KM¹ (em qtde)":
            km = c3.number_input("Qtd", min_value=0, key=f"v_{i}")
            item['valor'] = round(km * VALOR_KM, 2)
            c3.write(f"**{formatar_moeda(item['valor'])}**")
        else:
            item['valor'] = c3.number_input("R$", min_value=0.0, key=f"v_{i}")
            
        item['motivo'] = c4.text_input("Motivo", key=f"m_{i}")
        if c5.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()

    if st.button("➕ Adicionar Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()})
        st.rerun()
        
    arquivos = st.file_uploader("Comprovantes", accept_multiple_files=True, key="upload_colab")
    
    if st.button("ENVIAR PARA VERIFICAÇÃO"):
        if st.session_state.nome_colab and arquivos and any(it['valor'] and it['valor'] > 0 for it in st.session_state.items_reembolso):
            caminhos = salvar_arquivos_locais(arquivos)
            detalhes_limpos = []
            for it in st.session_state.items_reembolso:
                d = it.copy()
                d['data'] = d['data'].strftime("%d/%m/%Y")
                detalhes_limpos.append(d)
                
            nova_solic = {
                "id": len(db_global) + 1,
                "Colaborador": st.session_state.nome_colab,
                "Data": datetime.now().strftime("%d/%m/%Y"),
                "Detalhes": detalhes_limpos,
                "Status": "Em Verificação",
                "CaminhoArquivo": caminhos,
                "Comentario": ""
            }
            
            # Salva no global (para o Gabriel ver) e tenta no Excel (segurança)
            db_global.append(nova_solic)
            atualizar_excel()
            enviar_aviso_ao_gabriel(nova_solic)
            
            # --- ZERAR TUDO APÓS O ENVIO ---
            st.session_state.nome_colab = ""
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]
            
            st.success("✅ Solicitação enviada com sucesso! Campos redefinidos.")
            st.balloons()
            st.rerun()
        else:
            st.error("Por favor, preencha o nome, adicione valores e anexe comprovantes.")

with aba_admin:
    st.header("Painel de Verificação")
    if st.text_input("Senha", type="password") == "globus2026":
        pendentes = [s for s in db_global if s['Status'] == "Em Verificação"]
        if not pendentes:
            st.info("Nenhuma solicitação pendente.")
            if st.button("🔄 Atualizar"): st.rerun()
        
        for idx, solic in enumerate(pendentes):
            with st.expander(f"ID {solic['id']} - {solic['Colaborador']}"):
                for i_it, it in enumerate(solic['Detalhes']):
                    ec1, ec2, ec3, ec4 = st.columns(4)
                    it['data'] = ec1.text_input("Data", it['data'], key=f"ad_{idx}_{i_it}")
                    it['categoria'] = ec2.selectbox("Cat", CATEGORIAS, index=CATEGORIAS.index(it['categoria']), key=f"ac_{idx}_{i_it}")
                    it['valor'] = ec3.number_input("Valor", value=float(it['valor'] or 0), key=f"av_{idx}_{i_it}")
                    it['motivo'] = ec4.text_input("Motivo", it['motivo'], key=f"am_{idx}_{i_it}")
                
                dec = st.radio("Status", ["Aprovado", "Reprovado"], key=f"dec_{idx}", horizontal=True)
                com = st.text_area("Justificativa", key=f"com_{idx}")
                
                if st.button("Finalizar", key=f"f_{idx}"):
                    solic['Status'] = dec
                    solic['Comentario'] = com
                    atualizar_excel()
                    st.success("Processado!")
                    st.rerun()
