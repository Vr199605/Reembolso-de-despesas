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

# LINK DA SUA PLANILHA GOOGLE (FORMATO CSV)
URL_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTchCpRAYj7uoqFQURSWWtdWAWBqS89-4qH9DIamqGQ2IykWHvLT_I-jSPrsyY-v_Zy2gDVJtkc-qnQ/pub?output=csv"
ARQUIVO_LOCAL_BACKUP = "base_reembolsos.xlsx"

# --- FUNÇÕES DE AUXÍLIO ---
def formatar_moeda(valor):
    if valor is None: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- FUNÇÕES DE SISTEMA ---

def atualizar_excel():
    """Mantém um backup local em Excel das alterações feitas no Session State"""
    todos_itens = []
    for solic in st.session_state.db:
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
    df = pd.DataFrame(todos_itens)
    df.to_excel(ARQUIVO_LOCAL_BACKUP, index=False)

def carregar_dados_nuvem():
    """Tenta carregar os dados do link do Google Sheets, se falhar, tenta o local"""
    try:
        # Lê diretamente do link CSV que você forneceu
        df = pd.read_csv(URL_GOOGLE_SHEETS)
        db_recuperado = []
        for solic_id in df['ID'].unique():
            df_solic = df[df['ID'] == solic_id]
            primeira_linha = df_solic.iloc[0]
            detalhes = []
            for _, row in df_solic.iterrows():
                detalhes.append({
                    "categoria": row['Categoria'],
                    "valor": row['Valor'],
                    "motivo": row['Motivo'],
                    "data": row.get('Data_Item', primeira_linha['Data'])
                })
            
            db_recuperado.append({
                "id": int(solic_id),
                "Colaborador": primeira_linha['Colaborador'],
                "Data": primeira_linha['Data'],
                "Status": primeira_linha['Status'],
                "Detalhes": detalhes,
                "CaminhoArquivo": primeira_linha['Caminho_Arquivo'],
                "Comentario": str(primeira_linha['Comentario_Admin']) if pd.notna(primeira_linha['Comentario_Admin']) else ""
            })
        return db_recuperado
    except Exception as e:
        # Se der erro no link (internet ou formato), tenta carregar o arquivo local
        if os.path.exists(ARQUIVO_LOCAL_BACKUP):
            df = pd.read_excel(ARQUIVO_LOCAL_BACKUP)
            # ... (mesma lógica de processamento do excel)
            return carregar_dados_locais_fallback(df)
        return []

def carregar_dados_locais_fallback(df):
    db_recuperado = []
    for solic_id in df['ID'].unique():
        df_solic = df[df['ID'] == solic_id]
        primeira_linha = df_solic.iloc[0]
        detalhes = []
        for _, row in df_solic.iterrows():
            detalhes.append({
                "categoria": row['Categoria'],
                "valor": row['Valor'],
                "motivo": row['Motivo'],
                "data": row.get('Data_Item', primeira_linha['Data'])
            })
        db_recuperado.append({
            "id": int(solic_id),
            "Colaborador": primeira_linha['Colaborador'],
            "Data": primeira_linha['Data'],
            "Status": primeira_linha['Status'],
            "Detalhes": detalhes,
            "CaminhoArquivo": primeira_linha['Caminho_Arquivo'],
            "Comentario": primeira_linha['Comentario_Admin']
        })
    return db_recuperado

def enviar_aviso_ao_gabriel(solicitacao):
    destinatario = "gabriel.coelho@globusseguros.com.br"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"📩 Nova Solicitação de Reembolso: {solicitacao['Colaborador']}"
    corpo = f"Olá Gabriel Coelho,\n\nUm colaborador enviou uma solicitação de reembolso.\nID: {solicitacao['id']}\nColaborador: {solicitacao['Colaborador']}\n\nAcesse: https://reembolsodespesas.streamlit.app/"
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587); server.starttls()
        server.login(remetente, senha); server.send_message(msg); server.quit()
        return True
    except: return False

def enviar_email_automatico(dados, arquivo_pdf, arquivo_comprovante):
    destinatario = "gabriel.coelho@globusseguros.com.br"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"[{dados['Status'].upper()}] Reembolso: {dados['Colaborador']} - ID {dados['id']}"
    corpo = f"Solicitação de reembolso finalizada.\nColaborador: {dados['Colaborador']}\nStatus: {dados['Status']}\nObs: {dados.get('Comentario', '')}"
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        for arq in [arquivo_pdf, arquivo_comprovante]:
            if os.path.exists(arq):
                with open(arq, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(arq))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arq)}"'
                    msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587); server.starttls()
        server.login(remetente, senha); server.send_message(msg); server.quit()
        return True
    except: return False

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
    info_data = [[Paragraph("<b>ID:</b>", styles['Normal']), f"#{dados['id']}", Paragraph("<b>COLABORADOR:</b>", styles['Normal']), dados['Colaborador']], [Paragraph("<b>STATUS:</b>", styles['Normal']), dados['Status'], Paragraph("<b>APROVADOR:</b>", styles['Normal']), "GABRIEL COELHO"]]
    t_info = Table(info_data, colWidths=[1.2*inch, 2.5*inch, 1*inch, 1.8*inch])
    t_info.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke)]))
    elements.append(t_info); elements.append(Spacer(1, 20))
    despesas_data = [["DATA", "CATEGORIA", "VALOR (R$)", "MOTIVO"]]
    total_geral = 0
    for item in dados['Detalhes']:
        despesas_data.append([item.get('data', dados['Data']), item['categoria'], formatar_moeda(item['valor']), item['motivo']])
        total_geral += item['valor']
    despesas_data.append(["", "TOTAL A REEMBOLSAR", formatar_moeda(total_geral), ""])
    t_desp = Table(despesas_data, colWidths=[0.9*inch, 2.0*inch, 1.1*inch, 2.5*inch])
    t_desp.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('ALIGN', (2,1), (2,-1), 'RIGHT')]))
    elements.append(t_desp)
    if dados.get('Comentario'):
        elements.append(Spacer(1, 20)); elements.append(Paragraph(f"<b>OBSERVAÇÕES:</b>", styles['Normal'])); elements.append(Paragraph(dados['Comentario'], styles['Normal']))
    doc.build(elements)

def gerar_relatorio_mensal_pdf(lista_solicitacoes, mes_ano, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=30)
    styles = getSampleStyleSheet(); elements = []
    title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], fontSize=20, textColor=colors.HexColor("#1f4e79"), alignment=1, spaceAfter=10)
    elements.append(Paragraph("FECHAMENTO MENSAL DE REEMBOLSOS", title_style))
    elements.append(Paragraph(f"Período: {mes_ano} | Empresa: Globus Seguros", styles['Normal']))
    data_table = [["COLABORADOR", "CATEGORIA", "DATA", "VALOR (R$)"]]
    total_periodo = 0
    for s in lista_solicitacoes:
        for item in s['Detalhes']:
            data_table.append([s['Colaborador'], item['categoria'], item.get('data', s['Data']), formatar_moeda(item['valor'])])
            total_periodo += item['valor']
    t = Table(data_table, colWidths=[2*inch, 2.5*inch, 1*inch, 1.5*inch])
    t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('GRID', (0,0), (-1,-1), 0.5, colors.grey)]))
    elements.append(t); elements.append(Spacer(1, 25))
    elements.append(Paragraph(f"<b>TOTAL GERAL DO MÊS: {formatar_moeda(total_periodo)}</b>", styles['Normal']))
    doc.build(elements)

# --- INICIALIZAÇÃO DE DADOS ---
if 'db' not in st.session_state: 
    st.session_state.db = carregar_dados_nuvem()

if 'items_reembolso' not in st.session_state: 
    st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]

aba_colab, aba_admin = st.tabs(["🚀 Solicitar Reembolso", "🔑 Verificação e Aprovação (Gabriel)"])

with aba_colab:
    st.header("Nova Solicitação")
    nome = st.text_input("Nome Completo")
    st.info("💡 Observação: Em caso de Almoço com Cliente, favor utilizar a categoria 'OUTROS* (em R$)'")
    st.markdown("---")
    for i, item in enumerate(st.session_state.items_reembolso):
        col_data, col_cat, col_val, col_mot, col_del = st.columns([1.2, 1.8, 1.2, 1.8, 0.4])
        item['data'] = col_data.date_input(f"Data {i+1}", value=item.get('data', datetime.now()), format="DD/MM/YYYY", key=f"date_{i}")
        item['categoria'] = col_cat.selectbox(f"Categoria {i+1}", CATEGORIAS, key=f"cat_{i}")
        if item['categoria'] == "KM¹ (em qtde)":
            qtd_km = col_val.number_input("Qtd KM", min_value=0, step=1, value=None, key=f"km_{i}")
            valor_calc = round((qtd_km if qtd_km else 0) * VALOR_KM, 2)
            item['valor'] = valor_calc
            col_val.markdown(f"<p style='color: #1f4e79; font-weight: bold; margin:0;'>{formatar_moeda(valor_calc)}</p>", unsafe_allow_html=True)
        else:
            item['valor'] = col_val.number_input(f"Valor R$", min_value=0.0, step=0.01, format="%.2f", value=None, key=f"val_{i}")
            if item['valor'] and item['categoria'] in LIMITES and item['valor'] > LIMITES[item['categoria']]:
                st.warning(f"Limite para {item['categoria']} é {formatar_moeda(LIMITES[item['categoria']])}.")
        item['motivo'] = col_mot.text_input(f"Motivo (Obrigatório)", key=f"mot_{i}")
        if col_del.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i); st.rerun()
    if st.button("➕ Adicionar Outro Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}); st.rerun()
    arquivo = st.file_uploader("Anexar Comprovante (Obrigatório)", type=['pdf', 'png', 'jpg'])
    if st.button("Enviar para Verificação"):
        todos_motivos = all(it['motivo'].strip() != "" for it in st.session_state.items_reembolso)
        if nome and arquivo and any(it['valor'] and it['valor'] > 0 for it in st.session_state.items_reembolso) and todos_motivos:
            path = salvar_arquivo_local(arquivo)
            detalhes_limpos = []
            for it in st.session_state.items_reembolso:
                d = it.copy(); d['data'] = d['data'].strftime("%d/%m/%Y"); detalhes_limpos.append(d)
            nova_solic = {"id": len(st.session_state.db) + 1, "Colaborador": nome, "Data": datetime.now().strftime("%d/%m/%Y"), "Detalhes": detalhes_limpos, "Status": "Em Verificação", "CaminhoArquivo": path, "Comentario": ""}
            st.session_state.db.append(nova_solic); atualizar_excel(); enviar_aviso_ao_gabriel(nova_solic)
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]
            st.success("Enviado! Gabriel recebeu um e-mail."); st.rerun()
        else: st.error("Preencha todos os campos e anexe o comprovante.")

with aba_admin:
    st.header("Painel de Controle - Gabriel Coelho")
    senha_adm = st.text_input("Senha de Acesso", type="password")
    if senha_adm == "globus2026":
        # RECARREGA DO LINK GOOGLE AO LOGAR
        if st.button("🔄 Sincronizar com Planilha Google"):
            st.session_state.db = carregar_dados_nuvem()
            st.success("Dados sincronizados com a nuvem!")

        st.subheader("📊 Relatórios e Fechamento Mensal")
        col_m1, col_m2 = st.columns([1, 2])
        mes_ref = col_m1.selectbox("Selecione o Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
        ano_ref = col_m2.selectbox("Ano", [2025, 2026, 2027])
        
        meses_map = {"Janeiro":"01", "Fevereiro":"02", "Março":"03", "Abril":"04", "Maio":"05", "Junho":"06", "Julho":"07", "Agosto":"08", "Setembro":"09", "Outubro":"10", "Novembro":"11", "Dezembro":"12"}
        filtro_mes_ano = f"{meses_map[mes_ref]}/{ano_ref}"
        
        solicitacoes_mes = [s for s in st.session_state.db if s['Status'] == "Aprovado" and any(filtro_mes_ano in it.get('data', s['Data']) for it in s['Detalhes'])]
        
        if st.button("📄 GERAR PDF DE FECHAMENTO MENSAL"):
            if solicitacoes_mes:
                nome_pdf_mensal = f"Fechamento_{mes_ref}_{ano_ref}.pdf"
                gerar_relatorio_mensal_pdf(solicitacoes_mes, f"{mes_ref}/{ano_ref}", nome_pdf_mensal)
                with open(nome_pdf_mensal, "rb") as f: st.download_button("📥 Baixar Relatório", f, file_name=nome_pdf_mensal)
            else: st.warning("Não existem despesas 'Aprovadas' para este período.")
        
        st.markdown("---")
        st.subheader("⏳ Solicitações Pendentes")
        verificar = [s for s in st.session_state.db if s['Status'] == "Em Verificação"]
        if not verificar: st.info("Não há solicitações pendentes.")
        for idx, solic in enumerate(verificar):
            with st.expander(f"ID {solic['id']} - {solic['Colaborador']}"):
                c_edit, c_view = st.columns([1.5, 1])
                with c_edit:
                    for i_item, item in enumerate(solic['Detalhes']):
                        ec0, ec1, ec2, ec3 = st.columns([1, 1.5, 1, 1.5])
                        item['data'] = ec0.text_input(f"Data", value=item.get('data', solic['Data']), key=f"adm_d_{idx}_{i_item}")
                        item['categoria'] = ec1.selectbox(f"Cat", CATEGORIAS, index=CATEGORIAS.index(item['categoria']), key=f"adm_cat_{idx}_{i_item}")
                        item['valor'] = ec2.number_input(f"Valor", value=float(item['valor'] or 0), format="%.2f", key=f"adm_v_{idx}_{i_item}")
                        item['motivo'] = ec3.text_input(f"Motivo", value=item['motivo'], key=f"adm_m_{idx}_{i_item}")
                    decisao = st.radio("Sua Decisão", ["Aprovado", "Reprovado"], key=f"dec_{idx}", horizontal=True)
                    motivo_final = st.text_area("Justificativa", key=f"com_{idx}")
                    if st.button("FINALIZAR", key=f"fin_{idx}"):
                        solic['Status'] = decisao; solic['Comentario'] = motivo_final; atualizar_excel()
                        nome_pdf = f"Relatorio_ID_{solic['id']}.pdf"
                        gerar_relatorio_pdf(solic, nome_pdf); enviar_email_automatico(solic, nome_pdf, solic['CaminhoArquivo'])
                        st.success(f"Solicitação #{solic['id']} finalizada!"); st.rerun()
                with c_view:
                    if os.path.exists(solic['CaminhoArquivo']):
                        with open(solic['CaminhoArquivo'], "rb") as f: st.download_button(label="📂 Baixar Comprovante", data=f, file_name=os.path.basename(solic['CaminhoArquivo']), key=f"dl_{idx}")
    elif senha_adm != "": st.error("Senha incorreta.")
