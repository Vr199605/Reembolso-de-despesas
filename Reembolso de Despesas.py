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

# --- FUNÇÕES DE SISTEMA ---

def enviar_email_automatico(dados, arquivo_pdf, arquivo_comprovante):
    destinatario = "gabriel.coelho@globusseguros.com.br"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"  # <--- COLOQUE A SENHA DE 16 DÍGITOS AQUI

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    status_formatado = dados['Status'].upper()
    msg['Subject'] = f"[{status_formatado}] Reembolso: {dados['Colaborador']} - ID {dados['id']}"

    corpo = f"""
    Olá Gabriel Coelho,
    
    Uma solicitação de reembolso foi processada por CHRISTIAN WELLISCH.
    
    Colaborador: {dados['Colaborador']}
    Status: {status_formatado}
    Data da Solicitação: {dados['Data']}
    """
    if dados['Status'] == "Reprovado":
        corpo += f"\nMOTIVO DA REPROVAÇÃO: {dados.get('Comentario', 'Não informado')}"
    
    msg.attach(MIMEText(corpo, 'plain'))

    # Anexar Relatório
    try:
        with open(arquivo_pdf, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(arquivo_pdf))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo_pdf)}"'
            msg.attach(part)

        # Anexar Comprovante
        with open(arquivo_comprovante, "rb") as f:
            part_comp = MIMEApplication(f.read(), Name=os.path.basename(arquivo_comprovante))
            part_comp['Content-Disposition'] = f'attachment; filename="COMPROVANTE_{os.path.basename(arquivo_comprovante)}"'
            msg.attach(part_comp)

        # LOGICA DE ENVIO REAL (O que faltava)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")
        return False

def salvar_arquivo_local(file):
    if not os.path.exists("comprovantes"): 
        os.makedirs("comprovantes")
    # Limpar nome do arquivo para evitar erros de path
    safe_filename = "".join([c for c in file.name if c.isalnum() or c in ('.','_')]).strip()
    path = os.path.join("comprovantes", safe_filename)
    with open(path, "wb") as f: 
        f.write(file.getbuffer())
    return path

# (Mantenha a função gerar_relatorio_pdf como está, ela está correta)
def gerar_relatorio_pdf(dados, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph(f"Relatório de Reembolso - Globus", styles['Title']))
    elements.append(Spacer(1, 12))
    
    info = [
        ["Colaborador:", dados['Colaborador']],
        ["Data Solicitação:", dados['Data']],
        ["Aprovador:", "CHRISTIAN WELLISCH"],
        ["Status Final:", dados['Status']],
        ["Justificativa:", dados.get('Comentario', 'Aprovado conforme política')]
    ]
    t_info = Table(info, colWidths=[150, 300])
    t_info.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold')]))
    elements.append(t_info)
    elements.append(Spacer(1, 20))

    data_despesas = [["Categoria", "Valor", "Motivo"]]
    for cat, det in dados['Detalhes'].items():
        data_despesas.append([cat, f"R$ {det['valor']:.2f}", det['motivo']])

    t_despesas = Table(data_despesas, colWidths=[150, 100, 250])
    t_despesas.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('ALIGN', (0,0), (-1,-1), 'CENTER')
    ]))
    elements.append(t_despesas)
    doc.build(elements)

# --- REGRAS E INTERFACE ---
LIMITES = {"REFEIÇÃO VIAGEM (em R$)": 100.0, "REPRESENTAÇÃO (em R$)": 100.0, "ESTACIONAMENTO (em R$)": 70.0, "BEBIDA ALCOÓLICA (em R$)": 50.0}
CATEGORIAS = ["ESTACIONAMENTO (em R$)", "PEDÁGIO (em R$)", "KM¹ (em qtde)", "KM² (em R$)", "REPRESENTAÇÃO (em R$)", "TAXI / UBER (em R$)", "REFEIÇÃO VIAGEM (em R$)", "OUTROS* (em R$)"]

if 'db' not in st.session_state: st.session_state.db = []

aba_colab, aba_christian = st.tabs(["🚀 Solicitar Reembolso", "⚖️ Portal de Aprovação"])

with aba_colab:
    st.header("Nova Solicitação")
    nome = st.text_input("Nome Completo")
    data_solic = st.date_input("Data da Despesa")
    selecionadas = st.multiselect("Categorias", CATEGORIAS)
    
    detalhes_form = {}
    for cat in selecionadas:
        c_v, c_m = st.columns([1, 2])
        v = c_v.number_input(f"Valor {cat}", step=0.01, key=f"v_{cat}")
        m = c_m.text_input(f"Motivo {cat}", key=f"m_{cat}")
        detalhes_form[cat] = {"valor": v, "motivo": m}
    
    arquivo = st.file_uploader("Comprovante", type=['pdf', 'png', 'jpg'])
    
    if st.button("Enviar Solicitação"):
        if nome and arquivo and detalhes_form:
            path = salvar_arquivo_local(arquivo)
            st.session_state.db.append({
                "id": len(st.session_state.db) + 1,
                "Colaborador": nome,
                "Data": str(data_solic),
                "Detalhes": detalhes_form,
                "Status": "Pendente",
                "CaminhoArquivo": path
            })
            st.success("Solicitação registrada!")
        else:
            st.error("Preencha tudo!")

with aba_christian:
    st.header("Painel Christian")
    pendentes = [s for s in st.session_state.db if s['Status'] == "Pendente"]
    
    for solic in pendentes:
        with st.expander(f"Solicitação #{solic['id']} - {solic['Colaborador']}"):
            decisao = st.radio("Decisão", ["Aprovado", "Reprovado"], key=f"d_{solic['id']}")
            motivo = st.text_area("Comentário", key=f"c_{solic['id']}")
            
            if st.button("Confirmar", key=f"b_{solic['id']}"):
                solic['Status'] = decisao
                solic['Comentario'] = motivo
                
                nome_pdf = f"Relatorio_ID_{solic['id']}.pdf"
                gerar_relatorio_pdf(solic, nome_pdf)
                
                if enviar_email_automatico(solic, nome_pdf, solic['CaminhoArquivo']):
                    st.success("E-mail enviado!")
                    st.rerun()