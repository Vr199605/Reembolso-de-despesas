import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import os
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal de Reembolso - Globus", layout="wide")

# --- REGRAS DA POLÍTICA (PDF) ---
LIMITES = {
    "REFEIÇÃO VIAGEM (em R$)": 100.0, 
    "REPRESENTAÇÃO (em R$)": 100.0, 
    "ESTACIONAMENTO (em R$)": 70.0, 
    "BEBIDA ALCOÓLICA (em R$)": 50.0
}
CATEGORIAS = ["ESTACIONAMENTO (em R$)", "PEDÁGIO (em R$)", "KM¹ (em qtde)", "KM² (em R$)", "REPRESENTAÇÃO (em R$)", "TAXI / UBER (em R$)", "REFEIÇÃO VIAGEM (em R$)", "OUTROS* (em R$)"]
VALOR_KM = 1.37

# --- FUNÇÕES DE SISTEMA ---

def enviar_email_automatico(dados, arquivo_pdf, arquivo_comprovante):
    destinatario = "gabriel.coelho@globusseguros.com.br"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    status_formatado = dados['Status'].upper()
    msg['Subject'] = f"[{status_formatado}] Reembolso: {dados['Colaborador']} - ID {dados['id']}"

    corpo = f"Olá Gabriel Coelho,\n\nSolicitação de reembolso processada.\n\nColaborador: {dados['Colaborador']}\nStatus: {status_formatado}\nData: {dados['Data']}"
    if dados['Status'] == "Reprovado": corpo += f"\nMOTIVO: {dados.get('Comentario', 'Não informado')}"
    msg.attach(MIMEText(corpo, 'plain'))

    try:
        for arq in [arquivo_pdf, arquivo_comprovante]:
            with open(arq, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(arq))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arq)}"'
                msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erro e-mail: {e}")
        return False

def salvar_arquivo_local(file):
    if not os.path.exists("comprovantes"): os.makedirs("comprovantes")
    path = os.path.join("comprovantes", file.name)
    with open(path, "wb") as f: f.write(file.getbuffer())
    return path

def gerar_relatorio_pdf(dados, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=letter)
    styles, elements = getSampleStyleSheet(), []
    elements.append(Paragraph(f"Relatório de Reembolso - Globus", styles['Title']))
    info = [["Colaborador:", dados['Colaborador']], ["Data:", dados['Data']], ["Status:", dados['Status']], ["Obs:", dados.get('Comentario', '-')]]
    t_info = Table(info, colWidths=[150, 300])
    t_info.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey)]))
    elements.append(t_info)
    elements.append(Spacer(1, 20))
    data_despesas = [["Categoria", "Valor", "Motivo"]]
    for item in dados['Detalhes']: data_despesas.append([item['categoria'], f"R$ {item['valor']:.2f}", item['motivo']])
    t_desp = Table(data_despesas, colWidths=[150, 100, 250])
    t_desp.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('GRID', (0,0), (-1,-1), 0.5, colors.grey)]))
    elements.append(t_desp)
    doc.build(elements)

# --- INTERFACE ---
if 'db' not in st.session_state: st.session_state.db = []
if 'items_reembolso' not in st.session_state: st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""}]

aba_colab, aba_christian, aba_admin = st.tabs(["🚀 Solicitar Reembolso", "⚖️ Portal de Aprovação", "🔑 Admin (Gabriel)"])

with aba_colab:
    st.header("Nova Solicitação")
    nome = st.text_input("Nome Completo")
    data_solic = st.date_input("Data da Despesa", format="DD/MM/YYYY")
    
    for i, item in enumerate(st.session_state.items_reembolso):
        col1, col2, col3, col4 = st.columns([2, 1, 2, 0.5])
        item['categoria'] = col1.selectbox(f"Cat {i}", CATEGORIAS, key=f"cat_{i}")
        if item['categoria'] == "KM¹ (em qtde)":
            qtd_km = col2.number_input("Qtde KM", step=1, key=f"km_{i}")
            item['valor'] = qtd_km * VALOR_KM
            col2.caption(f"Total: R$ {item['valor']:.2f}")
        else:
            item['valor'] = col2.number_input(f"Valor", step=0.01, key=f"val_{i}")
            if item['categoria'] in LIMITES and item['valor'] > LIMITES[item['categoria']]:
                col2.warning(f"Limite: R${LIMITES[item['categoria']]}")
        item['motivo'] = col3.text_input(f"Motivo", key=f"mot_{i}")
        if col4.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()

    if st.button("➕ Adicionar Outra Categoria"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""})
        st.rerun()

    arquivo = st.file_uploader("Comprovante", type=['pdf', 'png', 'jpg'])
    if st.button("Enviar Solicitação"):
        if nome and arquivo:
            path = salvar_arquivo_local(arquivo)
            st.session_state.db.append({"id": len(st.session_state.db)+1, "Colaborador": nome, "Data": data_solic.strftime("%d/%m/%Y"), "Detalhes": st.session_state.items_reembolso.copy(), "Status": "Pendente", "CaminhoArquivo": path})
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""}]
            st.success("Enviado!")

with aba_christian:
    st.header("Painel Christian Wellisch")
    pendentes = [s for s in st.session_state.db if s['Status'] == "Pendente"]
    for solic in pendentes:
        with st.expander(f"Solicitação #{solic['id']} - {solic['Colaborador']}"):
            col_inf, col_img = st.columns(2)
            with col_inf:
                st.table(pd.DataFrame(solic['Detalhes']))
                decisao = st.radio("Decisão", ["Aprovado", "Reprovado"], key=f"d_{solic['id']}")
                motivo = st.text_area("Comentário", key=f"c_{solic['id']}")
                if st.button("Confirmar Aprovação", key=f"b_{solic['id']}"):
                    solic['Status'], solic['Comentario'] = decisao, motivo
                    nome_pdf = f"Relatorio_ID_{solic['id']}.pdf"
                    gerar_relatorio_pdf(solic, nome_pdf)
                    enviar_email_automatico(solic, nome_pdf, solic['CaminhoArquivo'])
                    st.rerun()
            with col_img:
                if solic['CaminhoArquivo'].lower().endswith('pdf'): st.write("Arquivo PDF Anexado")
                else: st.image(solic['CaminhoArquivo'], use_container_width=True)

with aba_admin:
    st.header("Área Restrita - Gabriel Coelho")
    pw = st.text_input("Senha Admin", type="password")
    if pw == "globus2026":
        filtro = st.selectbox("Filtrar Colaborador", ["Todos"] + list(set(s['Colaborador'] for s in st.session_state.db)))
        dados_v = st.session_state.db if filtro == "Todos" else [s for s in st.session_state.db if s['Colaborador'] == filtro]
        
        for i, solic in enumerate(dados_v):
            with st.expander(f"EDITAR: ID {solic['id']} - {solic['Colaborador']} ({solic['Status']})"):
                solic['Colaborador'] = st.text_input("Nome", solic['Colaborador'], key=f"adm_n_{i}")
                solic['Status'] = st.selectbox("Status", ["Pendente", "Aprovado", "Reprovado"], index=["Pendente", "Aprovado", "Reprovado"].index(solic['Status']), key=f"adm_s_{i}")
                if st.button("Salvar Alterações", key=f"adm_b_{i}"): st.success("Atualizado!"); st.rerun()
