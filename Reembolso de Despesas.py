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
URL_APP = "https://reembolso-de-despesas.streamlit.app/"

# --- INOVAÇÃO: MEMÓRIA COMPARTILHADA ---
@st.cache_resource
def iniciar_db_global():
    return []

db_global = iniciar_db_global()

# --- FUNÇÕES DE AUXÍLIO ---
def formatar_moeda(valor):
    if valor is None: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- FUNÇÕES DE SISTEMA ---

def atualizar_excel():
    todos_itens = []
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

def enviar_aviso_ao_gabriel(solicitacao):
    destinatario = "victormoreiraicnv@gmail.com"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"📩 Nova Solicitação de Reembolso: {solicitacao['Colaborador']}"
    
    corpo = f"""
    Olá Gabriel Coelho,
    
    Um colaborador enviou uma nova solicitação de reembolso.
    
    Colaborador: {solicitacao['Colaborador']}
    ID da Solicitação: #{solicitacao['id']}
    
    Para visualizar e aprovar as solicitações pendentes, acesse o link abaixo:
    {URL_APP}
    """
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
    corpo = f"Olá Gabriel Coelho,\n\nUma solicitação foi finalizada.\n\nColaborador: {dados['Colaborador']}\nStatus: {status_formatado}\nLink: {URL_APP}"
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        if os.path.exists(arquivo_pdf):
            with open(arquivo_pdf, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(arquivo_pdf))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo_pdf)}"'
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
    elements.append(Paragraph(f"RELATÓRIO DE REEMBOLSO - {dados['Colaborador']}", style_header))
    elements.append(HRFlowable(width="100%", thickness=2, color=cor_primaria, spaceAfter=20))
    doc.build(elements)

# --- INICIALIZAÇÃO ---
if 'items_reembolso' not in st.session_state: 
    st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()}]
if 'nome_colab' not in st.session_state:
    st.session_state.nome_colab = ""

# --- INTERFACE ---
aba_guia, aba_colab, aba_admin = st.tabs(["📖 Guia de Preenchimento", "🚀 Solicitar Reembolso", "🔑 Verificação e Aprovação (Gabriel)"])

with aba_guia:
    st.title("📖 Como solicitar seu reembolso")
    st.markdown("""
    Bem-vindo ao **Portal de Reembolsos Globus**. Siga o passo a passo abaixo para garantir que sua solicitação seja processada rapidamente.
    
    ---
    ### 1️⃣ Identificação
    Na aba **'Solicitar Reembolso'**, comece preenchendo seu **Nome Completo**. Isso é fundamental para a organização dos pagamentos.

    ### 2️⃣ Adicionando Despesas
    Você pode adicionar várias despesas em uma única solicitação:
    * **Data:** Selecione a data exata em que o gasto ocorreu.
    * **Categoria:** Escolha o tipo de despesa (ex: Estacionamento, Uber, Pedágio).
    * **Valor:** Insira o valor conforme o comprovante.
    * *Nota para KM:* Ao selecionar **KM¹**, insira a quantidade rodada e o sistema calculará automaticamente o valor (R$ 1,37/km).
    * **Motivo:** Descreva brevemente o motivo do gasto.

    ### 3️⃣ Comprovantes
    **Nenhuma despesa é aprovada sem comprovante.** Você pode selecionar múltiplos arquivos de uma vez.
    ---
    """)
    
    st.write("Caso ainda tenha alguma dúvida segue abaixo o manual de politicas de viagens e reembolso.")
    
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    if os.path.exists(caminho_manual):
        with open(caminho_manual, "rb") as f:
            st.download_button("📥 BAIXAR MANUAL DE REEMBOLSO (PDF)", f, file_name="manual_reembolso.pdf")
    
    st.info("💡 Assim que você clicar em 'Enviar', o Gabriel Coelho receberá uma notificação imediata com o link para análise.")

with aba_colab:
    st.header("Formulário de Reembolso - Globus")
    st.session_state.nome_colab = st.text_input("Nome Completo", value=st.session_state.nome_colab, placeholder="Digite seu nome aqui...")
    st.markdown("---")
    
    for i, item in enumerate(st.session_state.items_reembolso):
        with st.container():
            c1, c2, c3, c4, c5 = st.columns([1.2, 1.8, 1.2, 1.8, 0.4])
            
            dt_val = item.get('data', datetime.now())
            if isinstance(dt_val, str):
                dt_val = datetime.strptime(dt_val, "%d/%m/%Y")
                
            item['data'] = c1.date_input(f"Data {i}", value=dt_val, key=f"d_{i}")
            item['categoria'] = c2.selectbox(f"Categoria {i}", CATEGORIAS, key=f"c_{i}")
            
            if item['categoria'] == "KM¹ (em qtde)":
                km = c3.number_input("Qtd KM", min_value=0, key=f"v_{i}", value=0)
                item['valor'] = round(km * VALOR_KM, 2)
                c3.markdown(f"<div style='background-color: #f0f2f6; padding: 10px; border-radius: 5px; text-align: center; font-weight: bold;'>{formatar_moeda(item['valor'])}</div>", unsafe_allow_html=True)
            else:
                # Valor inicial 0.00 para não confundir o colaborador
                item['valor'] = c3.number_input(f"Valor R$ {i}", min_value=0.0, format="%.2f", key=f"v_{i}", value=0.0)
            
            item['motivo'] = c4.text_input(f"Motivo {i}", key=f"m_{i}", placeholder="Descreva o motivo...", label_visibility="collapsed")
            
            if c5.button("🗑️", key=f"del_{i}"):
                st.session_state.items_reembolso.pop(i)
                st.rerun()
            
            if item['valor'] > 0 and item['categoria'] in LIMITES and item['valor'] > LIMITES[item['categoria']]:
                st.markdown(f"<div style='background-color: #fff3cd; color: #856404; padding: 8px; border-radius: 5px; margin-top: -10px; margin-bottom: 10px; font-size: 14px;'>⚠️ Limite permitido para esta categoria: {formatar_moeda(LIMITES[item['categoria']])}</div>", unsafe_allow_html=True)

    if st.button("➕ Adicionar Outro Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()})
        st.rerun()
        
    st.markdown("---")
    arquivos = st.file_uploader("Anexar Comprovantes (Obrigatório)", accept_multiple_files=True, key="up_colab")
    
    if st.button("ENVIAR PARA VERIFICAÇÃO"):
        if st.session_state.nome_colab and arquivos and any(it['valor'] > 0 for it in st.session_state.items_reembolso):
            with st.spinner("Enviando solicitação..."):
                caminhos = salvar_arquivos_locais(arquivos)
                detalhes_finais = []
                for it in st.session_state.items_reembolso:
                    d = it.copy()
                    if not isinstance(d['data'], str):
                        d['data'] = d['data'].strftime("%d/%m/%Y")
                    detalhes_finais.append(d)
                    
                nova_solic = {
                    "id": len(db_global) + 1,
                    "Colaborador": st.session_state.nome_colab,
                    "Data": datetime.now().strftime("%d/%m/%Y"),
                    "Detalhes": detalhes_finais,
                    "Status": "Em Verificação",
                    "CaminhoArquivo": caminhos,
                    "Comentario": ""
                }
                
                db_global.append(nova_solic)
                atualizar_excel()
                enviar_aviso_ao_gabriel(nova_solic)
                
                # --- RESET TOTAL DOS CAMPOS ---
                st.session_state.nome_colab = ""
                st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": 0.0, "motivo": "", "data": datetime.now()}]
                
                st.success("✅ Sua solicitação foi enviada para análise com sucesso!")
                st.balloons()
                st.rerun()
        else:
            st.error("Por favor, preencha seu nome, insira valores válidos e anexe os comprovantes.")

with aba_admin:
    st.title("⌛ Solicitações Pendentes")
    senha_adm = st.text_input("Acesso Restrito - Digite a Senha", type="password")
    if senha_adm == "globus2026":
        pendentes = [s for s in db_global if s['Status'] == "Em Verificação"]
        
        if not pendentes:
            st.info("Não há solicitações aguardando aprovação.")
            if st.button("🔄 Atualizar Lista"): st.rerun()
        
        for idx, solic in enumerate(pendentes):
            with st.expander(f"📌 ID #{solic['id']} - {solic['Colaborador']}"):
                c_edit, c_view = st.columns([2, 1])
                with c_edit:
                    for i_it, it in enumerate(solic['Detalhes']):
                        ec1, ec2, ec3, ec4 = st.columns(4)
                        it['data'] = ec1.text_input(f"Data_{idx}_{i_it}", it['data'])
                        it['categoria'] = ec2.selectbox(f"Cat_{idx}_{i_it}", CATEGORIAS, index=CATEGORIAS.index(it['categoria']))
                        it['valor'] = ec3.number_input(f"Val_{idx}_{i_it}", value=float(it['valor'] or 0))
                        it['motivo'] = ec4.text_input(f"Mot_{idx}_{i_it}", it['motivo'])
                    
                    dec = st.radio(f"Decisão Final #{solic['id']}", ["Aprovado", "Reprovado"], horizontal=True)
                    mot_fin = st.text_area(f"Observações Financeiro #{solic['id']}")
                    
                    if st.button(f"Confirmar Processamento #{solic['id']}", key=f"fin_btn_{idx}"):
                        solic['Status'] = dec
                        solic['Comentario'] = mot_fin
                        atualizar_excel()
                        pdf_nome = f"Relatorio_{solic['id']}.pdf"
                        gerar_relatorio_pdf(solic, pdf_nome)
                        enviar_email_automatico(solic, pdf_nome, solic['CaminhoArquivo'])
                        st.success(f"Solicitação #{solic['id']} processada!")
                        st.rerun()
                
                with c_view:
                    st.write("📂 **Comprovantes Anexados:**")
                    for path in solic['CaminhoArquivo'].split(";"):
                        if os.path.exists(path):
                            with open(path, "rb") as f:
                                st.download_button(label=f"Visualizar {os.path.basename(path)}", data=f, file_name=os.path.basename(path), key=f"dl_{path}")
    elif senha_adm != "":
        st.error("Senha incorreta. Tente novamente.")
