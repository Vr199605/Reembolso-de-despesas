import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
import io
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ReportLab para o layout
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

# Configuração da página
st.set_page_config(page_title="Sistema de Reembolso", layout="wide")

# --- FUNÇÃO GERAR PDF (REPORTLAB) ---
def gerar_pdf(nome, data_sol, dados_tabela, total):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    elements = []
    styles = getSampleStyleSheet()
    
    elements.append(Paragraph(f"Relatório de Reembolso de Despesas", styles['Title']))
    elements.append(Spacer(1, 24))
    
    elements.append(Paragraph(f"<b>Colaborador:</b> {nome}", styles['Normal']))
    elements.append(Paragraph(f"<b>Data da Solicitação:</b> {data_sol}", styles['Normal']))
    elements.append(Spacer(1, 20))

    data = [["Data", "Categoria", "Motivo", "Valor (R$)"]]
    for item in dados_tabela:
        data.append([item['Data'], item['Categoria'], item['Motivo'], f"{item['Valor Total']:.2f}"])
    
    t = Table(data, colWidths=[80, 150, 220, 85])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('WORDWRAP', (0, 0), (-1, -1), True),
    ]))
    elements.append(t)
    
    elements.append(Spacer(1, 25))
    elements.append(Paragraph(f"<b>VALOR TOTAL A REEMBOLSAR: R$ {total:.2f}</b>", styles['Heading2']))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- FUNÇÕES DE E-MAIL ---
def enviar_email_com_pdf(destinatario, assunto, corpo, pdf_buffer=None):
    seu_email = "victormoreiraicnv@gmail.com"
    senha_app = "odym ioqm ybew ejnn" # Coloque sua senha de app aqui
    msg = MIMEMultipart()
    msg['From'] = seu_email
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'plain'))
    
    if pdf_buffer:
        part = MIMEApplication(pdf_buffer.read(), Name="reembolso_aprovado.pdf")
        part['Content-Disposition'] = 'attachment; filename="reembolso_aprovado.pdf"'
        msg.attach(part)
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(seu_email, senha_app)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

# --- INTERFACE ---
aba_guia, aba_solicitacao, aba_aprovacao = st.tabs(["📖 Guia Passo a Passo", "📋 Solicitação", "🔐 Aprovação (Gabriel Coelho)"])

with aba_guia:
    st.header("📖 Guia de Preenchimento de Reembolso")
    st.info("Siga os passos abaixo para garantir que sua solicitação seja processada sem erros.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        ### 1️⃣ Identificação
        Preencha seu **nome completo** e a **data atual** da solicitação. 
        
        ### 2️⃣ Seleção de Categorias
        Escolha todas as categorias que compõem sua despesa. Você pode selecionar várias ao mesmo tempo.
        
        ### 3️⃣ Detalhamento por Item
        Para cada categoria selecionada, informe:
        * **Data da despesa**: O dia em que o gasto ocorreu.
        * **Valor/Quantidade**: Se for KM, informe a distância. O sistema calcula o valor automaticamente.
        * **Motivo**: Descreva brevemente a finalidade do gasto (campo obrigatório).
        """)
    
    with col2:
        st.markdown("""
        ### 4️⃣ Comprovantes
        O upload de arquivos é **obrigatório**. Aceitamos PDF e imagens (PNG/JPG). Certifique-se de que o arquivo está legível.
        
        ### 5️⃣ Envio
        Clique em **Enviar Solicitação**. O sistema notificará o responsável e você receberá o aviso de sucesso na tela.

        ### 6️⃣ Data de Pagamento
        Após a análise e aprovação da solicitação, o pagamento será realizado em **D+5** (cinco dias úteis após a data da análise).
        """)
    
    st.markdown("---")
    st.subheader("❓ Ainda tem dúvidas?")
    st.write("Acesse o manual completo das políticas de viagens e reembolso no botão abaixo:")
    
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    
    try:
        with open(caminho_manual, "rb") as f:
            st.download_button(
                label="📥 Baixar Manual de Reembolso (PDF)",
                data=f,
                file_name="manual_reembolso.pdf",
                mime="application/pdf"
            )
    except FileNotFoundError:
        st.error("Arquivo 'manual_reembolso.pdf' não encontrado na pasta 'documentos'.")

with aba_solicitacao:
    st.title("🚀 Solicitação de Reembolso de Despesas")
    st.markdown("---")
    col_perfil1, col_perfil2 = st.columns(2)
    with col_perfil1: nome = st.text_input("Nome completo", key="nome_user")
    with col_perfil2: data_solicitacao = st.date_input("Data de solicitação", format="DD/MM/YYYY", key="data_sol")

    categorias_disponiveis = ["ESTACIONAMENTO (em R$)", "PEDÁGIO (em qtde)", "KM (em qtde)", "REPRESENTAÇÃO (em R$)", "TAXI / UBER (em R$)", "REFEIÇÃO VIAGEM (em R$)", "OUTROS* (em R$)"]
    selecionadas = st.multiselect("Selecione as categorias:", categorias_disponiveis)
    dados_despesas = []

    if selecionadas:
        for cat in selecionadas:
            with st.container():
                c1, c2, c3, c4 = st.columns([2, 2, 2, 4])
                c1.markdown(f"**{cat}**")
                d_desp = c2.date_input(f"Data", format="DD/MM/YYYY", key=f"d_{cat}")
                if "KM (em qtde)" in cat:
                    q_km = c3.number_input("Qtde KM", min_value=0.0, step=0.1, value=None, key=f"v_{cat}")
                    v_fin = (q_km * 1.37) if q_km else 0.0
                    if q_km: c3.info(f"R$ {v_fin:.2f}")
                else:
                    v_fin = c3.number_input("Valor R$", min_value=0.0, step=0.01, value=None, key=f"v_{cat}")
                    if "REFEIÇÃO" in cat: 
                        c3.markdown("**Limite até R$ 150**")
                    elif "ESTACIONAMENTO" in cat: 
                        c3.markdown("**Limite até R$ 70**")
                    v_fin = v_fin if v_fin else 0.0
                mot = c4.text_input("Motivo *", key=f"m_{cat}")
                dados_despesas.append({"Data": d_desp.strftime('%d/%m/%Y'), "Categoria": cat, "Valor Total": v_fin, "Motivo": mot})
        
        st.subheader("Anexar Comprovantes")
        arq = st.file_uploader("Upload (Obrigatório) *", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

        if st.button("Enviar Solicitação"):
            if any(not d["Motivo"].strip() for d in dados_despesas) or not arq:
                st.error("Preencha todos os motivos e anexe os arquivos.")
            else:
                st.session_state['solicitacao'] = {"nome": nome, "data": data_solicitacao.strftime('%d/%m/%Y'), "itens": dados_despesas}
                enviar_email_com_pdf("victormoreiraicnv@gmail.com", f"Solicitação: {nome}", f"O colaborador {nome} enviou uma solicitação de reembolso.")
                st.success("Enviado! Gabriel Coelho foi notificado.")
                st.warning("Você só receberá dentro do limite das politicas.")

with aba_aprovacao:
    st.title("🔐 Área de Verificação")
    if st.text_input("Senha", type="password") == "12345":
        if 'solicitacao' in st.session_state:
            sol = st.session_state['solicitacao']
            st.subheader(f"Ajuste de Solicitação: {sol['nome']}")
            
            dados_ajustados = []
            for i, item in enumerate(sol['itens']):
                with st.container():
                    c1, c2, c3, c4 = st.columns([2, 2, 2, 4])
                    c1.markdown(f"**{item['Categoria']}**")
                    adj_data = c2.text_input("Data", value=item['Data'], key=f"adj_d_{i}")
                    adj_val = c3.number_input("Valor R$", value=item['Valor Total'], key=f"adj_v_{i}")
                    adj_mot = c4.text_input("Motivo", value=item['Motivo'], key=f"adj_m_{i}")
                    dados_ajustados.append({"Data": adj_data, "Categoria": item['Categoria'], "Valor Total": adj_val, "Motivo": adj_mot})
            
            total_adj = sum(d["Valor Total"] for d in dados_ajustados)
            st.metric("Total Final", f"R$ {total_adj:.2f}")

            if st.button("✅ Aprovar e Enviar PDF"):
                pdf = gerar_pdf(sol['nome'], sol['data'], dados_ajustados, total_adj)
                enviar_email_com_pdf("victormoreiraicnv@gmail.com", f"APROVADO - {sol['nome']}", "Seguem os dados aprovados.", pdf)
                st.success("Aprovado e PDF enviado!")
            if st.button("❌ Reprovar"): st.error("Reprovada.")
        else: st.info("Sem pendências.")
