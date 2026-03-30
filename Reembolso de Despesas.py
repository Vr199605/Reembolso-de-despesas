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

        # LOGICA DE ENVIO REAL
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
    safe_filename = "".join([c for c in file.name if c.isalnum() or c in ('.','_')]).strip()
    path = os.path.join("comprovantes", safe_filename)
    with open(path, "wb") as f: 
        f.write(file.getbuffer())
    return path

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
    # Alterado para iterar sobre lista de dicionários para suportar duplicatas
    for item in dados['Detalhes']:
        data_despesas.append([item['categoria'], f"R$ {item['valor']:.2f}", item['motivo']])

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
if 'items_reembolso' not in st.session_state: st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""}]

aba_colab, aba_christian = st.tabs(["🚀 Solicitar Reembolso", "⚖️ Portal de Aprovação"])

with aba_colab:
    st.header("Nova Solicitação")
    nome = st.text_input("Nome Completo")
    data_solic = st.date_input("Data da Despesa", format="DD/MM/YYYY")
    
    st.markdown("### Itens de Despesa")
    for i, item in enumerate(st.session_state.items_reembolso):
        col1, col2, col3, col4 = st.columns([2, 1, 2, 0.5])
        item['categoria'] = col1.selectbox(f"Categoria {i+1}", CATEGORIAS, key=f"cat_{i}")
        item['valor'] = col2.number_input(f"Valor {i+1}", step=0.01, key=f"val_{i}")
        item['motivo'] = col3.text_input(f"Motivo {i+1}", key=f"mot_{i}")
        if col4.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()

    if st.button("➕ Adicionar Outra Categoria"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""})
        st.rerun()

    arquivo = st.file_uploader("Comprovante", type=['pdf', 'png', 'jpg'])
    
    if st.button("Enviar Solicitação"):
        if nome and arquivo and any(it['valor'] > 0 for it in st.session_state.items_reembolso):
            path = salvar_arquivo_local(arquivo)
            # Formatação da data para PT-BR
            data_br = data_solic.strftime("%d/%m/%Y")
            
            st.session_state.db.append({
                "id": len(st.session_state.db) + 1,
                "Colaborador": nome,
                "Data": data_br,
                "Detalhes": st.session_state.items_reembolso.copy(),
                "Status": "Pendente",
                "CaminhoArquivo": path
            })
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": ""}]
            st.success("Solicitação registrada com sucesso!")
        else:
            st.error("Preencha todos os campos obrigatórios.")

with aba_christian:
    st.header("Painel Christian")
    pendentes = [s for s in st.session_state.db if s['Status'] == "Pendente"]
    
    for solic in pendentes:
        with st.expander(f"Solicitação #{solic['id']} - {solic['Colaborador']} ({solic['Data']})"):
            st.write("**Resumo das Despesas:**")
            df_resumo = pd.DataFrame(solic['Detalhes'])
            st.table(df_resumo)
            
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
