import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
import os
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

# --- INOVAÇÃO: BANCO DE DADOS EM MEMÓRIA COMPARTILHADA (Solução para Nuvem) ---
@st.cache_resource
def iniciar_banco_dados_global():
    # Isso cria uma lista que todas as abas e usuários conseguem enxergar simultaneamente
    return []

db_global = iniciar_banco_dados_global()

# --- FUNÇÕES DE AUXÍLIO ---
def formatar_moeda(valor):
    if valor is None: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- FUNÇÕES DE SISTEMA ---

def enviar_aviso_ao_gabriel(solicitacao):
    destinatario = "victormoreiraicnv@gmail.com"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"📩 Nova Solicitação de Reembolso: {solicitacao['Colaborador']}"
    corpo = f"Olá Gabriel Coelho,\n\nUma nova solicitação foi enviada por {solicitacao['Colaborador']}.\n\nID: {solicitacao['id']}"
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
    status_formatado = dados['Status'].upper()
    msg['Subject'] = f"[{status_formatado}] Reembolso: {dados['Colaborador']} - ID {dados['id']}"
    corpo = f"Olá,\n\nA solicitação de {dados['Colaborador']} foi {status_formatado}.\nJustificativa: {dados.get('Comentario', 'N/A')}"
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        if os.path.exists(arquivo_pdf):
            with open(arquivo_pdf, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(arquivo_pdf))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo_pdf)}"'
                msg.attach(part)
        for caminho in caminhos_arquivos.split(";"):
            if os.path.exists(caminho):
                with open(caminho, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(caminho))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho)}"'
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
    doc = SimpleDocTemplate(nome_arquivo, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=30)
    styles = getSampleStyleSheet()
    elements = []
    cor_primaria = colors.HexColor("#1f4e79")
    style_header = ParagraphStyle('Header', parent=styles['Normal'], fontSize=20, textColor=cor_primaria, alignment=1, spaceAfter=10, fontName='Helvetica-Bold')
    style_label = ParagraphStyle('Label', parent=styles['Normal'], fontSize=9, textColor=colors.grey, fontName='Helvetica-Bold')
    style_value = ParagraphStyle('Value', parent=styles['Normal'], fontSize=11, textColor=colors.black)
    elements.append(Paragraph("RELATÓRIO DE REEMBOLSO", style_header))
    elements.append(HRFlowable(width="100%", thickness=2, color=cor_primaria, spaceAfter=20))
    info_data = [
        [Paragraph("COLABORADOR", style_label), Paragraph("ID SOLICITAÇÃO", style_label), Paragraph("STATUS", style_label)],
        [Paragraph(dados['Colaborador'], style_value), Paragraph(f"#{dados['id']}", style_value), Paragraph(dados['Status'].upper(), style_value)],
        [Spacer(1, 10), Spacer(1, 10), Spacer(1, 10)],
        [Paragraph("DATA DE EMISSÃO", style_label), Paragraph("APROVADOR", style_label), ""] ,
        [Paragraph(datetime.now().strftime("%d/%m/%Y"), style_value), Paragraph("GABRIEL COELHO", style_value), ""]
    ]
    t_info = Table(info_data, colWidths=[2.5*inch, 2*inch, 2*inch])
    t_info.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOTTOMPADDING', (0,0), (-1,-1), 8)]))
    elements.append(t_info)
    elements.append(Spacer(1, 20))
    despesas_data = [["DATA", "CATEGORIA", "MOTIVO / JUSTIFICATIVA", "VALOR"]]
    total_geral = 0
    for item in dados['Detalhes']:
        despesas_data.append([item.get('data', dados['Data']), item['categoria'], Paragraph(item['motivo'], styles['Normal']), formatar_moeda(item['valor'])])
        total_geral += item['valor']
    t_desp = Table(despesas_data, colWidths=[0.9*inch, 1.8*inch, 3.2*inch, 1.1*inch])
    t_desp.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), cor_primaria), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    elements.append(t_desp)
    elements.append(Spacer(1, 15))
    elements.append(Paragraph(f"TOTAL GERAL: {formatar_moeda(total_geral)}", style_value))
    doc.build(elements)

# --- INICIALIZAÇÃO DE DADOS LOCAIS ---
if 'items_reembolso' not in st.session_state: 
    st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]

# --- INTERFACE ---
aba_guia, aba_colab, aba_admin = st.tabs(["📖 Guia de Preenchimento", "🚀 Solicitar Reembolso", "🔑 Verificação e Aprovação (Gabriel)"])

with aba_guia:
    st.title("📖 Guia de Reembolso")
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    if os.path.exists(caminho_manual):
        with open(caminho_manual, "rb") as f:
            st.download_button("📥 BAIXAR MANUAL DE REEMBOLSO (PDF)", f, file_name="manual_reembolso.pdf")
    st.info("💡 Suas solicitações aparecem instantaneamente para o Gabriel Coelho.")

with aba_colab:
    st.header("Solicitar Reembolso")
    nome = st.text_input("Nome Completo")
    for i, item in enumerate(st.session_state.items_reembolso):
        c1, c2, c3, c4 = st.columns([1, 1.5, 1, 2])
        item['data'] = c1.date_input(f"Data {i}", value=item.get('data', datetime.now()), key=f"d_{i}").strftime("%d/%m/%Y")
        item['categoria'] = c2.selectbox(f"Categoria {i}", CATEGORIAS, key=f"c_{i}")
        if item['categoria'] == "KM¹ (em qtde)":
            km = c3.number_input("Qtd", min_value=0, key=f"v_{i}")
            item['valor'] = round(km * VALOR_KM, 2)
            c3.write(f"{formatar_moeda(item['valor'])}")
        else:
            item['valor'] = c3.number_input("R$", min_value=0.0, key=f"v_{i}")
        item['motivo'] = c4.text_input("Motivo", key=f"m_{i}")
    
    if st.button("➕ Adicionar Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()})
        st.rerun()
    
    arquivos = st.file_uploader("Anexar Comprovantes", accept_multiple_files=True)
    if st.button("ENVIAR SOLICITAÇÃO"):
        if nome and arquivos and any(it['valor'] > 0 for it in st.session_state.items_reembolso):
            caminhos = salvar_arquivos_locais(arquivos)
            nova = {
                "id": len(db_global) + 1, 
                "Colaborador": nome, 
                "Data": datetime.now().strftime("%d/%m/%Y"), 
                "Detalhes": st.session_state.items_reembolso.copy(), 
                "Status": "Em Verificação", 
                "CaminhoArquivo": caminhos, 
                "Comentario": ""
            }
            db_global.append(nova) # Adiciona à memória compartilhada
            enviar_aviso_ao_gabriel(nova)
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]
            st.success("Solicitação enviada com sucesso!")
            st.rerun()

with aba_admin:
    st.header("Painel do Gabriel Coelho")
    senha = st.text_input("Senha de Acesso", type="password")
    if senha == "globus2026":
        st.subheader("⏳ Solicitações Pendentes")
        # Lê diretamente da memória global compartilhada
        pendentes = [s for s in db_global if s['Status'] == "Em Verificação"]
        
        if not pendentes:
            st.info("Não há solicitações aguardando aprovação no momento.")
        
        for idx, solic in enumerate(pendentes):
            with st.expander(f"ID {solic['id']} - {solic['Colaborador']}"):
                col_e, col_v = st.columns([2, 1])
                with col_e:
                    st.write("📝 **Ajustes:**")
                    for i_it, it in enumerate(solic['Detalhes']):
                        ec1, ec2, ec3, ec4 = st.columns(4)
                        it['data'] = ec1.text_input("Data", it['data'], key=f"ad_{idx}_{i_it}")
                        it['categoria'] = ec2.selectbox("Cat", CATEGORIAS, index=CATEGORIAS.index(it['categoria']), key=f"ac_{idx}_{i_it}")
                        it['valor'] = ec3.number_input("Valor", value=float(it['valor'] or 0), key=f"av_{idx}_{i_it}")
                        it['motivo'] = ec4.text_input("Motivo", it['motivo'], key=f"am_{idx}_{i_it}")
                    
                    dec = st.radio("Decisão", ["Aprovado", "Reprovado"], key=f"dec_{idx}", horizontal=True)
                    com = st.text_area("Justificativa Financeiro", key=f"com_{idx}")
                    
                    if st.button("Finalizar Processamento", key=f"btn_{idx}"):
                        solic['Status'] = dec
                        solic['Comentario'] = com
                        nome_pdf = f"Relatorio_ID_{solic['id']}.pdf"
                        gerar_relatorio_pdf(solic, nome_pdf)
                        enviar_email_automatico(solic, nome_pdf, solic['CaminhoArquivo'])
                        st.success(f"ID {solic['id']} processado e e-mail enviado!")
                        st.rerun()
                
                with col_v:
                    st.write("📂 **Comprovantes:**")
                    for path in solic['CaminhoArquivo'].split(";"):
                        if os.path.exists(path):
                            with open(path, "rb") as f:
                                st.download_button(label=f"Ver {os.path.basename(path)}", data=f, file_name=os.path.basename(path), key=f"dl_{path}")
    elif senha != "":
        st.error("Senha incorreta.")
