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

# --- FUNÇÕES DE AUXÍLIO ---
def formatar_moeda(valor):
    if valor is None: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- FUNÇÕES DE SISTEMA ---
def atualizar_excel():
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
    df.to_excel(ARQUIVO_EXCEL, index=False)

def carregar_dados_iniciais():
    if os.path.exists(ARQUIVO_EXCEL):
        try:
            df = pd.read_excel(ARQUIVO_EXCEL)
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
        except: return []
    return []

def enviar_aviso_ao_gabriel(solicitacao):
    destinatario = "victormoreiraicnv@gmail.com"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"📩 Nova Solicitação de Reembolso: {solicitacao['Colaborador']}"
    # Link atualizado para incluir contexto visual no e-mail
    corpo = f"Olá Gabriel Coelho,\n\nUma nova solicitação foi enviada e está aguardando sua conferência e possíveis ajustes.\n\nColaborador: {solicitacao['Colaborador']}\nID: {solicitacao['id']}\nLink para aprovação: https://reembolsodespesas.streamlit.app/"
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

def enviar_email_automatico(dados, arquivo_pdf, arquivo_comprovante):
    destinatario = "victormoreiraicnv@gmail.com"
    remetente = "victormoreiraicnv@gmail.com"
    senha = "odym ioqm ybew ejnn"
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = f"[{dados['Status'].upper()}] Reembolso Globus - ID {dados['id']}"
    corpo = f"Olá,\n\nSua solicitação de reembolso ID {dados['id']} foi {dados['Status']}.\n\nConfira os detalhes no PDF anexo."
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
    except: return False

def salvar_arquivo_local(file):
    if not os.path.exists("comprovantes"): os.makedirs("comprovantes")
    path = os.path.join("comprovantes", file.name)
    with open(path, "wb") as f: f.write(file.getbuffer())
    return path

def gerar_relatorio_pdf(dados, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=50)
    styles = getSampleStyleSheet()
    elements = []
    
    azul_globus = colors.HexColor("#1f4e79")
    
    style_tit = ParagraphStyle('Tit', parent=styles['Title'], fontSize=22, textColor=azul_globus, spaceAfter=20, fontName='Helvetica-Bold')
    style_label = ParagraphStyle('Lab', parent=styles['Normal'], fontSize=10, fontName='Helvetica-Bold', textColor=azul_globus)
    style_header_tab = ParagraphStyle('HTab', parent=styles['Normal'], fontSize=10, fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=1)
    style_cell = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=9, leading=12)

    elements.append(Paragraph("RELATÓRIO DE REEMBOLSO", style_tit))
    elements.append(Paragraph("<b>GLOBUS SEGUROS</b>", ParagraphStyle('Sub', parent=styles['Normal'], fontSize=12, alignment=1, spaceAfter=30)))
    
    info = [
        [Paragraph("ID SOLICITAÇÃO", style_label), f"#{dados['id']}", Paragraph("DATA EMISSÃO", style_label), datetime.now().strftime('%d/%m/%Y')],
        [Paragraph("COLABORADOR", style_label), dados['Colaborador'], Paragraph("STATUS", style_label), dados['Status'].upper()]
    ]
    t_info = Table(info, colWidths=[1.5*inch, 2.2*inch, 1.5*inch, 2*inch])
    t_info.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.8, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 8),
    ]))
    elements.append(t_info)
    elements.append(Spacer(1, 25))

    items_data = [[
        Paragraph("DATA", style_header_tab), 
        Paragraph("CATEGORIA", style_header_tab), 
        Paragraph("MOTIVO / JUSTIFICATIVA", style_header_tab), 
        Paragraph("VALOR", style_header_tab)
    ]]
    
    total_reembolso = 0
    for it in dados['Detalhes']:
        items_data.append([
            Paragraph(it.get('data', dados['Data']), style_cell),
            Paragraph(it['categoria'], style_cell),
            Paragraph(it['motivo'], style_cell),
            Paragraph(formatar_moeda(it['valor']), ParagraphStyle('Val', parent=style_cell, alignment=2))
        ])
        total_reembolso += it['valor']
    
    items_data.append([
        "", "", 
        Paragraph("<b>TOTAL A REEMBOLSAR</b>", ParagraphStyle('TotL', parent=style_cell, alignment=2)), 
        Paragraph(f"<b>{formatar_moeda(total_reembolso)}</b>", ParagraphStyle('TotV', parent=style_cell, alignment=2))
    ])
    
    t_items = Table(items_data, colWidths=[0.9*inch, 1.8*inch, 3.3*inch, 1.2*inch])
    t_items.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), azul_globus),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#f2f2f2")),
    ]))
    elements.append(t_items)
    doc.build(elements)

# --- INICIALIZAÇÃO DE ESTADO ---
if 'db' not in st.session_state: st.session_state.db = carregar_dados_iniciais()
if 'items_reembolso' not in st.session_state: 
    st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]

# --- INTERFACE ---
aba_guia, aba_colab, aba_admin = st.tabs(["📖 Guia de Preenchimento", "🚀 Solicitar Reembolso", "🔑 Verificação e Aprovação (Gabriel)"])

with aba_guia:
    st.markdown("""
    <style>
        .guia-container {
            padding: 20px;
            border-radius: 10px;
            border-left: 6px solid #1f4e79;
            margin-bottom: 25px;
        }
        .passo-titulo {
            color: #1f4e79;
            font-weight: bold;
            font-size: 1.3em;
            margin-top: 20px;
        }
        .importante-box {
            background-color: rgba(31, 78, 121, 0.1);
            padding: 15px;
            border-radius: 8px;
            border: 1px solid #1f4e79;
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="guia-container">', unsafe_allow_html=True)
    st.title("📘 Guia Oficial de Reembolso Globus")
    st.write("Bem-vindo ao portal. Este guia foi desenhado para que sua experiência seja fluida e seu reembolso pago com agilidade.")
    
    st.markdown('<p class="passo-titulo">1️⃣ Identificação do Colaborador</p>', unsafe_allow_html=True)
    st.write("O preenchimento do seu **Nome Completo** é o primeiro passo. Sem ele, o sistema não consegue vincular os comprovantes à sua folha de pagamento.")
    
    st.markdown('<p class="passo-titulo">2️⃣ Detalhamento das Despesas (Obrigatório)</p>', unsafe_allow_html=True)
    st.markdown("""
    Cada item deve ser inserido individualmente para auditoria:
    * **Data:** Deve ser a mesma data impressa no comprovante fiscal.
    * **Categoria:** Escolha a que melhor descreve o gasto. Se for **KM**, apenas digite a quantidade; o sistema aplica o valor de **R$ 1,37/km** automaticamente.
    * **Motivo/Justificativa:** Seja específico (ex: 'Almoço com cliente da Corretora X'). Este campo é **obrigatório**.
    """)

    st.markdown('<p class="passo-titulo">3️⃣ Anexo de Comprovantes</p>', unsafe_allow_html=True)
    st.write("Você deve anexar um arquivo (Foto ou PDF) que contenha todos os comprovantes listados. Certifique-se de que a imagem não esteja cortada ou desfocada.")

    st.markdown('<p class="passo-titulo">4️⃣ Fluxo de Aprovação e Pagamento</p>', unsafe_allow_html=True)
    st.markdown('<div class="importante-box">', unsafe_allow_html=True)
    st.write("🚀 **Envio:** Após clicar em enviar, os dados vão para análise do Gabriel Coelho.")
    st.write("⏳ **Análise:** O status 'Em Verificação' indica que seus dados estão sendo conferidos.")
    st.write("💰 **Pagamento:** Uma vez aprovado, o prazo para crédito em conta é de **D+5 (cinco dias úteis)**.")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    # TEXTO ADICIONADO CONFORME PEDIDO
    st.write("caso ainda tenha alguma dúvida segue abaixo o manual de politicas de viagens e reembolso.")
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    if os.path.exists(caminho_manual):
        with open(caminho_manual, "rb") as f:
            st.download_button("📥 Baixar Manual em PDF", f, file_name="manual_politicas_globus.pdf")

with aba_colab:
    st.header("Solicitar Novo Reembolso")
    
    if st.button("🔄 Resetar Tudo", key="btn_reset_colab"):
        st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]
        if 'nome_input' in st.session_state: st.session_state.nome_input = ""
        st.rerun()

    nome_usuario = st.text_input("Seu Nome Completo", key="nome_input")
    st.markdown("---")
    
    for i, item in enumerate(st.session_state.items_reembolso):
        c1, c2, c3, c4, c5 = st.columns([1.2, 1.8, 1.2, 1.8, 0.4])
        item['data'] = c1.date_input(f"Data {i+1}", value=item.get('data', datetime.now()), key=f"dt_{i}")
        item['categoria'] = c2.selectbox(f"Categoria {i+1}", CATEGORIAS, key=f"ct_{i}")
        
        if item['categoria'] == "KM¹ (em qtde)":
            km_qtd = c3.number_input("Qtd KM", min_value=0, step=1, value=None, key=f"kmq_{i}")
            item['valor'] = round((km_qtd or 0) * VALOR_KM, 2)
            c3.info(f"R$ {item['valor']}")
        else:
            item['valor'] = c3.number_input(f"Valor R$", min_value=0.0, format="%.2f", value=None, key=f"vl_{i}")
            
            # VALIDAÇÃO DE LIMITES DE VALORES
            if item['categoria'] == "ESTACIONAMENTO (em R$)" and item['valor'] is not None and item['valor'] > 70:
                c3.warning("Limite: R$ 70")
            if item['categoria'] == "REFEIÇÃO VIAGEM (em R$)" and item['valor'] is not None and item['valor'] > 150:
                c3.warning("Limite: R$ 150")
        
        item['motivo'] = c4.text_input(f"Motivo {i+1}", key=f"mt_{i}", placeholder="Obrigatório")
        
        if c5.button("🗑️", key=f"rm_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()

    if st.button("➕ Adicionar Mais Itens"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()})
        st.rerun()

    arquivo_anexo = st.file_uploader("Anexe o Comprovante", type=['pdf', 'png', 'jpg'], key="file_colab")
    
    if st.button("Enviar para Verificação", type="primary"):
        validos = all(it['motivo'].strip() != "" and it['valor'] is not None for it in st.session_state.items_reembolso)
        if nome_usuario and arquivo_anexo and validos:
            caminho = salvar_arquivo_local(arquivo_anexo)
            detalhes_finais = [it.copy() for it in st.session_state.items_reembolso]
            for dfinal in detalhes_finais: dfinal['data'] = dfinal['data'].strftime("%d/%m/%Y")
            
            nova_solic = {
                "id": len(st.session_state.db)+1,
                "Colaborador": nome_usuario,
                "Data": datetime.now().strftime("%d/%m/%Y"),
                "Detalhes": detalhes_finais,
                "Status": "Em Verificação",
                "CaminhoArquivo": caminho,
                "Comentario": ""
            }
            
            st.session_state.db.append(nova_solic)
            atualizar_excel()
            enviar_aviso_ao_gabriel(nova_solic)
            
            # MENSAGEM DE SUCESSO CONFORME SOLICITADO
            st.success("foi enviado para verificação")
            st.balloons()
            
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]
            st.rerun()
        else:
            st.error("⚠️ Preencha todos os campos obrigatórios (Nome, Valor, Motivo) e anexe o comprovante.")

with aba_admin:
    st.header("Área Administrativa")
    if st.button("🔄 Atualizar Painel", key="reset_admin_tab"): st.rerun()
    
    acesso = st.text_input("Chave de Segurança", type="password")
    if acesso == "globus2026":
        pendentes = [s for s in st.session_state.db if s['Status'] == "Em Verificação"]
        
        if not pendentes:
            st.info("Tudo em dia! Nenhuma solicitação pendente.")
            
        for idx, solic in enumerate(pendentes):
            with st.expander(f"📦 ID #{solic['id']} - {solic['Colaborador']}"):
                st.write(f"**Data da Solicitação:** {solic['Data']}")
                
                # ÁREA DE AJUSTES PELO GABRIEL
                st.markdown("---")
                st.subheader("📝 Revisão e Ajustes")
                
                novos_detalhes = []
                for i_detalhe, item in enumerate(solic['Detalhes']):
                    col1, col2, col3, col4 = st.columns([1, 1.5, 1, 2])
                    nova_data = col1.text_input(f"Data item {i_detalhe+1}", value=item['data'], key=f"adj_dt_{idx}_{i_detalhe}")
                    nova_cat = col2.selectbox(f"Categoria item {i_detalhe+1}", CATEGORIAS, index=CATEGORIAS.index(item['categoria']) if item['categoria'] in CATEGORIAS else 0, key=f"adj_cat_{idx}_{i_detalhe}")
                    novo_val = col3.number_input(f"Valor item {i_detalhe+1}", value=float(item['valor']), key=f"adj_val_{idx}_{i_detalhe}")
                    novo_mot = col4.text_input(f"Motivo item {i_detalhe+1}", value=item['motivo'], key=f"adj_mot_{idx}_{i_detalhe}")
                    novos_detalhes.append({"data": nova_data, "categoria": nova_cat, "valor": novo_val, "motivo": novo_mot})
                
                st.markdown("---")
                c_inf, c_img = st.columns([1, 1])
                with c_inf:
                    status_sel = st.radio("Ação", ["Aprovado", "Reprovado"], key=f"rad_{idx}")
                    obs_adm = st.text_area("Justificativa / Observação (Enviada ao colaborador)", key=f"obs_adm_{idx}")
                    
                    if st.button("Confirmar Decisão", key=f"btn_adm_{idx}"):
                        # Salva os ajustes feitos pelo Gabriel antes de processar
                        solic['Detalhes'] = novos_detalhes
                        solic['Status'] = status_sel
                        solic['Comentario'] = obs_adm
                        atualizar_excel()
                        
                        pdf_path = f"Relatorio_Final_ID_{solic['id']}.pdf"
                        gerar_relatorio_pdf(solic, pdf_path)
                        enviar_email_automatico(solic, pdf_path, solic['CaminhoArquivo'])
                        
                        st.success("Processado com sucesso!")
                        st.rerun()
                with c_img:
                    if os.path.exists(solic['CaminhoArquivo']):
                        st.write("📄 **Comprovante:**")
                        with open(solic['CaminhoArquivo'], "rb") as f:
                            st.download_button("Visualizar Arquivo", f, file_name=os.path.basename(solic['CaminhoArquivo']), key=f"dl_adm_{idx}")
    elif acesso != "":
        st.error("Chave incorreta.")
