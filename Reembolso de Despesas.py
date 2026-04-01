import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
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

# --- PARÂMETROS FIXOS ---
URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT91PST4WLbGEWCEbIzaslCurwyOeJYMJHaTTcbQsgX0LrnVU_p_A5gnlfTjIQJs7KxyKTTREmSAJJE/pub?output=csv"
VALOR_KM = 1.37
LIMITES = {
    "REFEIÇÃO VIAGEM (em R$)": 150.0,
    "ESTACIONAMENTO (em R$)": 70.0,
    "BEBIDA ALCOÓLICA (em R$)": 50.0
}
CATEGORIAS = [
    "ESTACIONAMENTO (em R$)", "PEDÁGIO (em R$)", "KM¹ (em qtde)", 
    "REPRESENTAÇÃO (em R$)", "TAXI / UBER (em R$)", 
    "REFEIÇÃO VIAGEM (em R$)", "OUTROS* (em R$)"
]

# --- FUNÇÕES DE INTERFACE ---
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- PROCESSAMENTO DE DADOS ---
def processar_dataframe(df):
    db_recuperado = []
    try:
        df.columns = [c.strip() for c in df.columns]
        for solic_id in df['ID'].dropna().unique():
            df_solic = df[df['ID'] == solic_id]
            primeira_linha = df_solic.iloc[0]
            detalhes = []
            for _, row in df_solic.iterrows():
                detalhes.append({
                    "categoria": str(row.get('Categoria', 'Outros')),
                    "valor": float(row.get('Valor', 0)),
                    "motivo": str(row.get('Motivo', '')),
                    "data": str(row.get('Data Despesa', primeira_linha.get('Data Envio', '01/04/2026')))
                })
            db_recuperado.append({
                "id": int(solic_id),
                "Colaborador": str(primeira_linha.get('Colaborador', 'Desconhecido')),
                "Data": str(primeira_linha.get('Data Envio', '01/04/2026')),
                "Status": str(primeira_linha.get('Status', 'Em Verificação')),
                "Detalhes": detalhes,
                "CaminhoArquivo": str(primeira_linha.get('Comprovante', '')),
                "Comentario": str(primeira_linha.get('Comentário Admin', '')) if pd.notna(primeira_linha.get('Comentário Admin', '')) else ""
            })
        return db_recuperado
    except:
        return []

def carregar_dados_nuvem():
    try:
        df = pd.read_csv(URL_GOOGLE_SHEETS).dropna(how='all')
        return processar_dataframe(df)
    except:
        return []

# --- COMUNICAÇÃO (E-MAIL) ---
def enviar_email(assunto, corpo, destinatario, anexos=[]):
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = remetente, destinatario, assunto
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        for caminho in anexos:
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
    except:
        return False

# --- GERAÇÃO DE PDF ---
def gerar_pdf(dados, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = [Paragraph("RELATÓRIO DE REEMBOLSO - GLOBUS", styles['Title']), Spacer(1, 12)]
    
    info = [[f"ID: #{dados['id']}", f"Colaborador: {dados['Colaborador']}"], [f"Status: {dados['Status']}", f"Data: {dados['Data']}"]]
    elements.append(Table(info, colWidths=[2*inch, 4*inch]))
    elements.append(Spacer(1, 24))
    
    t_data = [["Data", "Categoria", "Valor", "Motivo"]]
    total = 0
    for it in dados['Detalhes']:
        t_data.append([it['data'], it['categoria'], formatar_moeda(it['valor']), it['motivo']])
        total += it['valor']
    
    t_data.append(["", "TOTAL", formatar_moeda(total), ""])
    table = Table(t_data, colWidths=[1*inch, 2*inch, 1.2*inch, 2.3*inch])
    table.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.grey), ('GRID', (0,0), (-1,-1), 1, colors.black)]))
    elements.append(table)
    doc.build(elements)

# --- APP STREAMLIT ---
if 'db' not in st.session_state:
    st.session_state.db = carregar_dados_nuvem()

tab1, tab2 = st.tabs(["🚀 Solicitar", "🔑 Admin"])

with tab1:
    st.header("Solicitação de Reembolso")
    nome = st.text_input("Nome Completo")
    
    if 'rows' not in st.session_state: st.session_state.rows = 1
    
    items = []
    for i in range(st.session_state.rows):
        c1, c2, c3, c4 = st.columns([1, 1.5, 1, 2])
        d_data = c1.date_input(f"Data", key=f"d_{i}").strftime("%d/%m/%Y")
        d_cat = c2.selectbox(f"Categoria", CATEGORIAS, key=f"c_{i}")
        if d_cat == "KM¹ (em qtde)":
            km = c3.number_input("Qtd KM", min_value=0, key=f"v_{i}")
            d_val = round(km * VALOR_KM, 2)
            c3.write(f"R$ {d_val}")
        else:
            d_val = c3.number_input("Valor R$", min_value=0.0, key=f"v_{i}")
        d_mot = c4.text_input("Motivo", key=f"m_{i}")
        items.append({"categoria": d_cat, "valor": d_val, "motivo": d_mot, "data": d_data})

    if st.button("➕ Adicionar Item"):
        st.session_state.rows += 1
        st.rerun()

    arq = st.file_uploader("Comprovante (PDF/Imagem)")
    
    if st.button("Enviar"):
        if nome and arq and all(x['motivo'] for x in items):
            novo_id = len(st.session_state.db) + 1
            # Simulação de salvamento (em um cenário real, você faria append no Google Sheets via API)
            st.success(f"Solicitação #{novo_id} enviada com sucesso!")
            enviar_email(f"Novo Reembolso: {nome}", f"O colaborador {nome} enviou uma nova solicitação.", "gabriel.coelho@globusseguros.com.br")
        else:
            st.error("Preencha todos os campos obrigatórios.")

with tab2:
    st.header("Painel Gabriel Coelho")
    senha = st.text_input("Senha", type="password")
    if senha == "globus2026":
        if st.button("🔄 Sincronizar Google Sheets"):
            st.session_state.db = carregar_dados_nuvem()
        
        pendentes = [s for s in st.session_state.db if s['Status'] == "Em Verificação"]
        for p in pendentes:
            with st.expander(f"ID {p['id']} - {p['Colaborador']}"):
                st.write(p['Detalhes'])
                obs = st.text_area("Observação", key=f"obs_{p['id']}")
                col_a, col_r = st.columns(2)
                if col_a.button("✅ Aprovar", key=f"ap_{p['id']}"):
                    p['Status'] = "Aprovado"
                    st.rerun()
                if col_r.button("❌ Reprovar", key=f"rp_{p['id']}"):
                    p['Status'] = "Reprovado"
                    st.rerun()
