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

# --- FUNÇÕES DE AUXÍLIO ---
def formatar_moeda(valor):
    if valor is None: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- FUNÇÕES DE SISTEMA ---

def enviar_aviso_ao_gabriel(solicitacao):
    """Envia e-mail para o Gabriel avisando que um colaborador enviou uma solicitação"""
    destinatario = "gabriel.coelho@globusseguros.com.br"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"📩 Nova Solicitação de Reembolso: {solicitacao['Colaborador']}"

    corpo = f"""
    Olá Gabriel Coelho,
    
    Um colaborador acabou de enviar uma nova solicitação de reembolso no portal.
    
    DETALHES:
    - Colaborador: {solicitacao['Colaborador']}
    - Data do Envio: {solicitacao['Data']}
    
    Por favor, acesse o portal para verificar, ajustar e aprovar a solicitação:
    https://reembolsodespesas.streamlit.app/
    A senha para acesso é: globus2026
    """
    msg.attach(MIMEText(corpo, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        return False

def enviar_email_automatico(dados, arquivo_pdf, arquivo_comprovante):
    destinatario = "gabriel.coelho@globusseguros.com.br"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    status_formatado = dados['Status'].upper()
    msg['Subject'] = f"[{status_formatado}] Reembolso: {dados['Colaborador']} - ID {dados['id']}"

    corpo = f"Olá Gabriel Coelho,\n\nUma solicitação de reembolso foi finalizada por você no Portal Globus.\n\nColaborador: {dados['Colaborador']}\nStatus: {status_formatado}\n"
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
        [Paragraph("<b>APROVADOR:</b>", styles['Normal']), "GABRIEL COELHO", "", ""]
    ]
    t_info = Table(info_data, colWidths=[1.2*inch, 2.5*inch, 1*inch, 1.8*inch])
    t_info.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke)]))
    elements.append(t_info)
    elements.append(Spacer(1, 20))
    despesas_data = [["CATEGORIA", "VALOR (R$)", "JUSTIFICATIVA / MOTIVO"]]
    total_geral = 0
    for item in dados['Detalhes']:
        despesas_data.append([item['categoria'], formatar_moeda(item['valor']), item['motivo']])
        total_geral += item['valor']
    despesas_data.append(["TOTAL A REEMBOLSAR", formatar_moeda(total_geral), ""])
    t_desp = Table(despesas_data, colWidths=[2.2*inch, 1.2*inch, 3.1*inch])
    t_desp.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('ALIGN', (1,1), (1,-1), 'RIGHT')]))
    elements.append(t_desp)
    if dados.get('Comentario'):
        elements.append(Spacer(1, 20))
        elements.append(Paragraph(f"<b>OBSERVAÇÕES:</b>", styles['Normal']))
        elements.append(Paragraph(dados['Comentario'], styles['Normal']))
    doc.build(elements)

# --- INTERFACE ---
if 'db' not in st.session_state: st.session_state.db = []
if 'items_reembolso' not in st.session_state: st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": ""}]

aba_colab, aba_admin = st.tabs(["🚀 Solicitar Reembolso", "🔑 Verificação e Aprovação (Gabriel)"])

with aba_colab:
    st.header("Nova Solicitação")
    nome = st.text_input("Nome Completo")
    data_solic = st.date_input("Data da Despesa", format="DD/MM/YYYY")
    st.markdown("---")
    for i, item in enumerate(st.session_state.items_reembolso):
        col1, col2, col3, col4 = st.columns([2, 1.5, 2, 0.5])
        item['categoria'] = col1.selectbox(f"Categoria {i+1}", CATEGORIAS, key=f"cat_{i}")
        if item['categoria'] == "KM¹ (em qtde)":
            qtd_km = col2.number_input("Quantidade de KM", min_value=0, step=1, value=None, key=f"km_{i}")
            valor_calc = round((qtd_km if qtd_km else 0) * VALOR_KM, 2)
            item['valor'] = valor_calc
            col2.markdown(f"<h3 style='color: #1f4e79; margin:0;'>{formatar_moeda(valor_calc)}</h3>", unsafe_allow_html=True)
        else:
            item['valor'] = col2.number_input(f"Valor R$", min_value=0.0, step=0.01, format="%.2f", value=None, key=f"val_{i}")
            if item['valor'] and item['categoria'] in LIMITES and item['valor'] > LIMITES[item['categoria']]:
                col2.warning(f"Limite: {formatar_moeda(LIMITES[item['categoria']])}")
        item['motivo'] = col3.text_input(f"Motivo / Justificativa", key=f"mot_{i}")
        if col4.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()
    if st.button("➕ Adicionar Outro Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": None, "motivo": ""})
        st.rerun()
    arquivo = st.file_uploader("Anexar Comprovante (Obrigatório)", type=['pdf', 'png', 'jpg'])
    if st.button("Enviar para Verificação"):
        if nome and arquivo and any(it['valor'] and it['valor'] > 0 for it in st.session_state.items_reembolso):
            path = salvar_arquivo_local(arquivo)
            detalhes_limpos = [it.copy() for it in st.session_state.items_reembolso]
            nova_solic = {
                "id": len(st.session_state.db)+1, "Colaborador": nome, "Data": data_solic.strftime("%d/%m/%Y"), 
                "Detalhes": detalhes_limpos, "Status": "Em Verificação", "CaminhoArquivo": path
            }
            st.session_state.db.append(nova_solic)
            enviar_aviso_ao_gabriel(nova_solic)
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": ""}]
            st.success("Enviado! Gabriel Coelho recebeu um e-mail para verificar.")
        else:
            st.error("Preencha todos os campos.")

with aba_admin:
    st.header("Painel de Controle - Gabriel Coelho")
    senha_adm = st.text_input("Senha de Acesso", type="password")
    if senha_adm == "globus2026":
        verificar = [s for s in st.session_state.db if s['Status'] == "Em Verificação"]
        if not verificar: st.info("Não há solicitações pendentes para sua aprovação.")
        for idx, solic in enumerate(verificar):
            with st.expander(f"ID {solic['id']} - {solic['Colaborador']}"):
                c_edit, c_view = st.columns([1.5, 1])
                with c_edit:
                    solic['Colaborador'] = st.text_input("Nome do Colaborador", solic['Colaborador'], key=f"adm_n_{idx}")
                    for i_item, item in enumerate(solic['Detalhes']):
                        ec1, ec2, ec3 = st.columns([2, 1.2, 2])
                        item['categoria'] = ec1.selectbox(f"Cat {i_item+1}", CATEGORIAS, index=CATEGORIAS.index(item['categoria']), key=f"adm_cat_{idx}_{i_item}")
                        item['valor'] = ec2.number_input(f"Valor", value=float(item['valor'] or 0), format="%.2f", key=f"adm_v_{idx}_{i_item}")
                        item['motivo'] = ec3.text_input(f"Motivo", value=item['motivo'], key=f"adm_m_{idx}_{i_item}")
                    
                    st.markdown("---")
                    decisao = st.radio("Sua Decisão", ["Aprovado", "Reprovado"], key=f"dec_{idx}", horizontal=True)
                    motivo_final = st.text_area("Justificativa / Comentário Interno", key=f"com_{idx}")
                    
                    if st.button("FINALIZAR E ENVIAR RELATÓRIO", key=f"fin_{idx}"):
                        solic['Status'] = decisao
                        solic['Comentario'] = motivo_final
                        nome_pdf = f"Relatorio_ID_{solic['id']}.pdf"
                        gerar_relatorio_pdf(solic, nome_pdf)
                        enviar_email_automatico(solic, nome_pdf, solic['CaminhoArquivo'])
                        st.success(f"Solicitação #{solic['id']} finalizada com sucesso!")
                        st.rerun()
                with c_view:
                    with open(solic['CaminhoArquivo'], "rb") as f:
                        st.download_button(label="📂 Baixar Comprovante", data=f, file_name=os.path.basename(solic['CaminhoArquivo']), key=f"dl_{idx}")
    elif senha_adm != "": st.error("Senha incorreta.")
