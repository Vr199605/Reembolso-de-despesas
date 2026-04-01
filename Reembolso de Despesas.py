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

# --- FUNÇÕES DE AUXÍLIO ---
def formatar_moeda(valor):
    if valor is None: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- FUNÇÕES DE SISTEMA ---

def atualizar_excel():
    """Salva o estado atual do db no arquivo Excel para permanência de dados"""
    todos_itens = []
    if 'db' in st.session_state:
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
            if df.empty:
                return []
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
        except:
            return []
    else:
        colunas = ["ID", "Colaborador", "Data_Item", "Status", "Categoria", "Valor", "Motivo", "Comentario_Admin", "Caminho_Arquivo"]
        df_vazio = pd.DataFrame(columns=colunas)
        df_vazio.to_excel(ARQUIVO_EXCEL, index=False)
        return []

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
    
    Um colaborador acabou de enviar uma nova solicitação de reembolso no portal.
    
    DETALHES:
    - Colaborador: {solicitacao['Colaborador']}
    - Data do Envio: {datetime.now().strftime('%d/%m/%Y')}
    
    Por favor, acesse o portal para verificar, ajustar e aprovar a solicitação na aba de Verificação:
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
    except:
        return False

def enviar_email_automatico(dados, arquivo_pdf, caminhos_arquivos):
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
    except:
        return False

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
    t_desp.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), cor_primaria), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('ALIGN', (-1,0), (-1,-1), 'RIGHT'), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (-1,0), 10), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8)]))
    elements.append(t_desp)
    total_data = [["", "TOTAL A RECEBER:", formatar_moeda(total_geral)]]
    t_total = Table(total_data, colWidths=[2.7*inch, 3.2*inch, 1.1*inch])
    t_total.setStyle(TableStyle([('ALIGN', (1,0), (1,0), 'RIGHT'), ('ALIGN', (2,0), (2,0), 'RIGHT'), ('FONTNAME', (1,0), (2,0), 'Helvetica-Bold'), ('FONTSIZE', (1,0), (2,0), 12), ('TEXTCOLOR', (2,0), (2,0), cor_primaria), ('TOPPADDING', (0,0), (-1,-1), 10)]))
    elements.append(t_total)
    if dados.get('Comentario'):
        elements.append(Spacer(1, 30))
        elements.append(Paragraph("OBSERVAÇÕES DO FINANCEIRO", style_label))
        elements.append(HRFlowable(width="30%", thickness=1, color=colors.lightgrey, align='LEFT'))
        elements.append(Spacer(1, 5))
        elements.append(Paragraph(dados['Comentario'], styles['Normal']))
    doc.build(elements)

# --- INICIALIZAÇÃO DE DADOS ---
if 'db' not in st.session_state: 
    st.session_state.db = carregar_dados_iniciais()
if 'items_reembolso' not in st.session_state: 
    st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]

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
    * **Motivo:** Descreva brevemente o motivo do gasto (ex: 'Visita ao cliente X'). **Este campo é obrigatório.**

    ### 3️⃣ Comprovantes
    **Nenhuma despesa é aprovada sem comprovante.**
    * Você pode selecionar múltiplos arquivos de uma vez.

    ### 4️⃣ Limites da Política
    Fique atento aos limites automáticos do sistema:
    * **Refeição Viagem:** Até R$ 150,00
    * **Estacionamento:** Até R$ 70,00
    ---
    """)
    
    # --- DOWNLOAD DO MANUAL ---
    caminho_manual = os.path.join("documentos", "manual_reembolso.pdf")
    if os.path.exists(caminho_manual):
        with open(caminho_manual, "rb") as f:
            st.download_button("📥 BAIXAR MANUAL DE REEMBOLSO (PDF)", f, file_name="manual_reembolso.pdf")
    
    st.info("💡 Assim que você clicar em 'Enviar', o Gabriel Coelho receberá uma notificação imediata para análise.")

with aba_colab:
    st.header("Formulário de Reembolso - Globus")
    nome = st.text_input("Nome Completo")
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
                st.warning(f"Item {i+1}: Limite para {item['categoria']} é {formatar_moeda(LIMITES[item['categoria']])}.")
        
        item['motivo'] = col_mot.text_input(f"Motivo (Obrigatório)", key=f"mot_{i}")
        
        if col_del.button("🗑️", key=f"del_{i}"):
            st.session_state.items_reembolso.pop(i)
            st.rerun()
            
    col_reset, col_add = st.columns([1, 4])
    if col_reset.button("🔄 Resetar Formulário"):
        st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]
        st.rerun()

    if col_add.button("➕ Adicionar Outro Item"):
        st.session_state.items_reembolso.append({"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()})
        st.rerun()
        
    arquivos = st.file_uploader("Anexar Comprovantes (Obrigatório)", type=['pdf', 'png', 'jpg'], accept_multiple_files=True)
    
    if st.button("Enviar para Verificação"):
        todos_motivos_preenchidos = all(it['motivo'].strip() != "" for it in st.session_state.items_reembolso)
        if nome and arquivos and any(it['valor'] and it['valor'] > 0 for it in st.session_state.items_reembolso) and todos_motivos_preenchidos:
            caminhos = salvar_arquivos_locais(arquivos)
            detalhes_limpos = []
            for it in st.session_state.items_reembolso:
                d = it.copy()
                d['data'] = d['data'].strftime("%d/%m/%Y")
                detalhes_limpos.append(d)
                
            db_atual = carregar_dados_iniciais()
            nova_solic = {
                "id": len(db_atual) + 1,
                "Colaborador": nome, 
                "Data": datetime.now().strftime("%d/%m/%Y"), 
                "Detalhes": detalhes_limpos, 
                "Status": "Em Verificação", 
                "CaminhoArquivo": caminhos, 
                "Comentario": ""
            }
            db_atual.append(nova_solic)
            st.session_state.db = db_atual
            atualizar_excel()
            
            enviar_aviso_ao_gabriel(nova_solic)
            st.session_state.items_reembolso = [{"categoria": CATEGORIAS[0], "valor": None, "motivo": "", "data": datetime.now()}]
            st.success("Enviado com sucesso!")
            st.rerun()
        else:
            st.error("Preencha todos os campos corretamente.")

with aba_admin:
    st.header("Painel de Controle - Gabriel Coelho")
    senha_adm = st.text_input("Senha de Acesso", type="password")
    if senha_adm == "globus2026":
        # RECARREGA OS DADOS PARA GARANTIR QUE OS NOVOS ENVIOS APAREÇAM
        st.session_state.db = carregar_dados_iniciais()
        
        st.subheader("⏳ Solicitações Pendentes")
        # GARANTE QUE FILTRAMOS DO ESTADO MAIS RECENTE
        verificar = [s for s in st.session_state.db if s['Status'] == "Em Verificação"]
        
        if not verificar:
            st.info("Não há solicitações aguardando aprovação.")

        for idx, solic in enumerate(verificar):
            with st.expander(f"ID {solic['id']} - {solic['Colaborador']}"):
                c_edit, c_view = st.columns([1.5, 1])
                with c_edit:
                    st.write("📝 **Ajustar e Processar:**")
                    # LOOP PELOS DETALHES PARA PERMITIR AJUSTE
                    for i_item, item in enumerate(solic['Detalhes']):
                        ec0, ec1, ec2, ec3 = st.columns([1, 1.5, 1, 1.5])
                        item['data'] = ec0.text_input(f"Data", value=item.get('data', solic['Data']), key=f"adm_d_{idx}_{i_item}")
                        item['categoria'] = ec1.selectbox(f"Cat", CATEGORIAS, index=CATEGORIAS.index(item['categoria']), key=f"adm_cat_{idx}_{i_item}")
                        item['valor'] = ec2.number_input(f"Valor", value=float(item['valor'] or 0), key=f"adm_v_{idx}_{i_item}")
                        item['motivo'] = ec3.text_input(f"Motivo", value=item['motivo'], key=f"adm_m_{idx}_{i_item}")
                    
                    decisao = st.radio("Sua Decisão", ["Aprovado", "Reprovado"], key=f"dec_{idx}", horizontal=True)
                    motivo_final = st.text_area("Justificativa", key=f"com_{idx}")
                    
                    if st.button("FINALIZAR E ENVIAR E-MAIL", key=f"fin_{idx}"):
                        solic['Status'] = decisao
                        solic['Comentario'] = motivo_final
                        atualizar_excel()
                        nome_pdf = f"Relatorio_ID_{solic['id']}.pdf"
                        gerar_relatorio_pdf(solic, nome_pdf)
                        
                        enviar_email_automatico(solic, nome_pdf, solic['CaminhoArquivo'])
                        
                        st.success(f"Solicitação #{solic['id']} finalizada e enviada por e-mail!")
                        st.rerun()
                
                with c_view:
                    st.write("📂 **Comprovantes:**")
                    for path in solic['CaminhoArquivo'].split(";"):
                        if os.path.exists(path):
                            with open(path, "rb") as f:
                                st.download_button(label=f"Baixar {os.path.basename(path)}", data=f, file_name=os.path.basename(path), key=f"dl_{path}")
    elif senha_adm != "": 
        st.error("Senha incorreta.")
