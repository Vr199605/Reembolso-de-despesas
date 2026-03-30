import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
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

# --- REGRAS DA POLÍTICA (Baseado no PDF) ---
LIMITES = {
    "REFEIÇÃO VIAGEM (em R$)": 150.0,  # Limite atualizado [cite: 56, 65]
    "ESTACIONAMENTO (em R$)": 70.0,   # Limite por visita [cite: 82]
    "BEBIDA ALCOÓLICA (em R$)": 50.0   # Apenas em representação [cite: 58]
}

# Categorias atualizadas (KM² removido)
CATEGORIAS = [
    "ESTACIONAMENTO (em R$)", 
    "PEDÁGIO (em R$)", 
    "KM¹ (em qtde)", 
    "REPRESENTAÇÃO (em R$)", 
    "TAXI / UBER (em R$)", 
    "REFEIÇÃO VIAGEM (em R$)", 
    "OUTROS* (em R$)"
]
VALOR_KM = 1.37 # Fator de conversão automático [cite: 74, 75]

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

    corpo = f"Olá Gabriel Coelho,\n\nUma solicitação de reembolso foi processada no Portal Globus.\n\nColaborador: {dados['Colaborador']}\nStatus: {status_formatado}\nData: {dados['Data']}\n"
    if dados['Status'] == "Reprovado":
        corpo += f"\nMOTIVO DA REPROVAÇÃO: {dados.get('Comentario', 'Não informado')}"
    
    msg.attach(MIMEText(corpo, 'plain'))

    try:
        for arq in [arquivo_pdf, arquivo_comprovante]:
            if os.path.exists(arq):
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
        st.error(f"Erro no envio do e-mail: {e}")
        return False

def salvar_arquivo_local(file):
    if not os.path.exists("comprovantes"): os.makedirs("comprovantes")
    path = os.path.join("comprovantes", file.name)
    with open(path, "wb") as f: f.write(file.getbuffer())
    return path

def gerar_relatorio_pdf(dados, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    styles = getSampleStyleSheet()
    elements = []
    title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], fontSize=18, textColor=colors.HexColor("#1f4e79"), alignment=1, spaceAfter=20)
    
    elements.append(Paragraph("RELATÓRIO DE REEMBOLSO OFICIAL - GLOBUS", title_style))
    elements.append(Spacer(1, 10))

    info_data = [
        [Paragraph("<b>ID SOLICITAÇÃO:</b>", styles['Normal']), f"#{dados['id']}", Paragraph("<b>DATA:</b>", styles['Normal']), dados['Data']],
        [Paragraph("<b>COLABORADOR:</b>", styles['Normal']), dados['Colaborador'], Paragraph("<b>STATUS:</b>", styles['Normal']), dados['Status']],
        [Paragraph("<b>APROVADOR:</b>", styles['Normal']), "CHRISTIAN WELLISCH", "", ""]
    ]
    
    t_info = Table(info_data, colWidths=[1.2*inch, 2.5*inch, 1*inch, 1.8*inch])
    t_info.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke), ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke)]))
    elements.append(t_info)
    elements.append(Spacer(1, 20))

    despesas_data = [["CATEGORIA", "VALOR (R$)", "JUSTIFICATIVA / MOTIVO"]]
    total_geral = 0
    for item in dados['Detalhes']:
        despesas_data.append([item['categoria'], f"{item['valor']:,.2f}", item['motivo']])
        total_geral += item['valor']
    despesas_data.append(["TOTAL A REEMBOLSAR", f"{total_geral:,.2f}", ""])

    t_desp = Table(despesas_data, colWidths=[2.2*inch, 1.2*inch, 3.1*inch])
    t_desp.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'), ('ALIGN', (1,1), (1,-1), 'RIGHT')]))
    elements.append(t_desp)

    if dados.get('Comentario'):
        elements.append(Spacer(1, 20))
        elements.append(Paragraph(f"<b>OBSERVAÇÕES:</b>", styles['Normal']))
        elements.append(Paragraph(dados['Comentario'], styles['Normal']))
    doc.build(elements)

# --- INTERFACE ---
if 'db' not in st.session_state: st.session_state.db = []
if 'items_reembolso' not in st.session_state: st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""}]

aba_colab, aba_christian, aba_admin = st.tabs(["🚀 Solicitar Reembolso", "⚖️ Portal de Aprovação", "🔑 Admin (Gabriel)"])

with aba_colab:
    st.header("Nova Solicitação")
    nome = st.text_input("Nome Completo")
    data_solic = st.date_input("Data da Despesa", format="DD/MM/YYYY")
    
    st.markdown("---")
    for i, item in enumerate(st.session_state.items_reembolso):
        col1, col2, col3, col4 = st.columns([2, 1.5, 2, 0.5])
        item['categoria'] = col1.selectbox(f"Categoria {i+1}", CATEGORIAS, key=f"cat_{i}")
        
        if item['categoria'] == "KM¹ (em qtde)":
            qtd_km = col2.number_input("Quantidade de KM", min_value=0, step=1, key=f"km_{i}")
            item['valor'] = round(qtd_km * VALOR_KM, 2)
            col2.markdown(f"<h3 style='color: #1f4e79; margin:0;'>R$ {item['valor']:.2f}</h3>", unsafe_allow_html=True)
        else:
            item['valor'] = col2.number_input(f"Valor R$", min_value=0.0, step=0.01, key=f"val_{i}")
            if item['categoria'] in LIMITES and item['valor'] > LIMITES[item['categoria']]:
                col2.warning(f"Limite: R${LIMITES[item['categoria']]}")
        
        item['motivo'] = col3.text_input(f"Motivo / Justificativa", key=f"mot_{i}")
        if col4.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()

    if st.button("➕ Adicionar Outro Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""})
        st.rerun()

    arquivo = st.file_uploader("Anexar Comprovante (Obrigatório)", type=['pdf', 'png', 'jpg'])
    
    if st.button("Enviar para Aprovação"):
        if nome and arquivo and any(it['valor'] > 0 for it in st.session_state.items_reembolso):
            path = salvar_arquivo_local(arquivo)
            st.session_state.db.append({
                "id": len(st.session_state.db)+1, "Colaborador": nome, "Data": data_solic.strftime("%d/%m/%Y"), 
                "Detalhes": st.session_state.items_reembolso.copy(), "Status": "Pendente", "CaminhoArquivo": path
            })
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""}]
            st.success("Enviado com sucesso!")

with aba_christian:
    st.header("Portal de Aprovação - Christian Wellisch")
    senha_ch = st.text_input("Senha de Acesso (Christian)", type="password")
    if senha_ch == "maldivas2026":
        pendentes = [s for s in st.session_state.db if s['Status'] == "Pendente"]
        for solic in pendentes:
            with st.expander(f"SOLICITAÇÃO #{solic['id']} - {solic['Colaborador']}"):
                c_inf, c_img = st.columns([1, 1])
                with c_inf:
                    st.table(pd.DataFrame(solic['Detalhes']))
                    decisao = st.radio("Decisão", ["Aprovado", "Reprovado"], key=f"d_{solic['id']}", horizontal=True)
                    motivo_rep = st.text_area("Justificativa", key=f"c_{solic['id']}")
                    if st.button("Confirmar e Enviar", key=f"b_{solic['id']}"):
                        solic['Status'] = decisao
                        solic['Comentario'] = motivo_rep
                        nome_pdf = f"Relatorio_ID_{solic['id']}.pdf"
                        gerar_relatorio_pdf(solic, nome_pdf)
                        if enviar_email_automatico(solic, nome_pdf, solic['CaminhoArquivo']):
                            st.success("Enviado para Gabriel Coelho.")
                            st.rerun()
                with c_img:
                    st.write("**Comprovante:**")
                    if solic['CaminhoArquivo'].lower().endswith('pdf'):
                        with open(solic['CaminhoArquivo'], "rb") as f:
                            b64 = base64.b64encode(f.read()).decode('utf-8')
                            st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="400"></iframe>', unsafe_allow_html=True)
                    else: st.image(solic['CaminhoArquivo'])
    elif senha_ch != "": st.error("Acesso Negado.")

with aba_admin:
    st.header("Painel Administrativo - Gabriel Coelho")
    senha_adm = st.text_input("Senha de Acesso (Admin)", type="password")
    if senha_adm == "globus2026":
        for idx, solic in enumerate(st.session_state.db):
            with st.expander(f"EDITAR ID {solic['id']} - {solic['Colaborador']}"):
                c1, c2 = st.columns(2)
                solic['Colaborador'] = c1.text_input("Nome", solic['Colaborador'], key=f"adm_n_{idx}")
                solic['Status'] = c2.selectbox("Status", ["Pendente", "Aprovado", "Reprovado"], index=["Pendente", "Aprovado", "Reprovado"].index(solic['Status']), key=f"adm_s_{idx}")
                for i_item, item in enumerate(solic['Detalhes']):
                    ec1, ec2, ec3 = st.columns([2, 1, 2])
                    item['categoria'] = ec1.selectbox(f"Cat {i_item+1}", CATEGORIAS, index=CATEGORIAS.index(item['categoria']) if item['categoria'] in CATEGORIAS else 0, key=f"adm_cat_{idx}_{i_item}")
                    item['valor'] = ec2.number_input(f"R$", value=float(item['valor']), key=f"adm_v_{idx}_{i_item}")
                    item['motivo'] = ec3.text_input(f"Motivo", value=item['motivo'], key=f"adm_m_{idx}_{i_item}")
                if st.button("Salvar", key=f"adm_btn_{idx}"): st.rerun()
