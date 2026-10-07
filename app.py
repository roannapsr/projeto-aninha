import os
import io
import re
from datetime import datetime
import streamlit as st
from PIL import Image
from dotenv import load_dotenv
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from google import genai
from google.genai import types

# ----------------------------------------------------
# 1. Configurações Iniciais da Página
# ----------------------------------------------------
load_dotenv()

if os.path.exists("aninha.jpeg"):
    icone_aba = Image.open("aninha.jpeg")
elif os.path.exists("aninha.png"):
    icone_aba = Image.open("aninha.png")
else:
    icone_aba = "👩‍⚕️"

st.set_page_config(
    page_title="Dra. Aninha - Perícia Médica Previdenciária",
    page_icon=icone_aba,
    layout="wide"
)

st.markdown(
    """
    <style>
    section[data-testid="stSidebar"] button[kind="primary"] {
        background-color: #0d6efd !important;
        border-color: #0d6efd !important;
        color: white !important;
    }
    section[data-testid="stSidebar"] button[kind="primary"]:hover {
        background-color: #0b5ed7 !important;
        border-color: #0a58ca !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ----------------------------------------------------
# 2. Inicialização do Cliente Gemini e Instrução Oficial
# ----------------------------------------------------
api_key = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}
hoje = datetime.now()
DATA_ATUAL_EXTENSO = f"Recife, {hoje.day:02d} de {MESES_PT[hoje.month]} de {hoje.year}"

SYSTEM_INSTRUCTION = f"""
Você é a Dra. Aninha, médica perita judicial previdenciária e assistente técnica pericial de alto nível.
Sua missão é atuar em auxílio à Dra. Ana Paula da Costa Henriques (CRM-PE 11.395), Perita Judicial Federal da Seção Judiciária de Pernambuco (19ª Vara - TRF5).

COMPORTAMENTO EM CONVERSAS E ANÁLISES (CHAT COMUM):
- Quando a médica enviar dados do periciando, resumos ou fizer perguntas, atue como colega perita consultora.
- Responda ESTRITAMENTE ao que foi solicitado na demanda atual. Não desvie de assunto nem responda quesitos ou tópicos não solicitados.
- Mantenha títulos e destaques sempre em negrito para facilitar a leitura.

AO FIXAR MARCOS TEMPORAIS (DID e DII):
- Foque EXCLUSIVAMENTE na análise da DID (Data de Início da Doença) e da DII (Data de Início da Incapacidade).
- NÃO responda quesitos do juízo e NÃO elabore a minuta de laudo completo.
- Fundamente tecnicamente a correlação clínico-documental com base na Lei 8.213/91.

AO RESPONDER AOS QUESITOS DO JUÍZO:
- Identifique cada quesito em negrito com seu tema (ex: **Quesito 1 - Identificação:**, **Quesito 7 - Incapacidade laborativa:**, etc.) e forneça a resposta técnica pericial fundamentada logo abaixo.
- Conclua obrigatoriamente TODOS os 19 quesitos do Quadro I do Juízo da 19ª Vara, do 1 ao 19 sem interrupção.

QUANDO FOR SOLICITADO O LAUDO PERICIAL OFICIAL:
- NUNCA inicie sua resposta com saudações, introduções ou conversas prévias.
- Comece diretamente em: PODER JUDICIÁRIO DA UNIÃO.
- Mantenha apenas os títulos das seções numeradas em negrito (ex: **1. PREÂMBULO**).
- Os campos do preâmbulo e dos dados do periciado devem ser texto regular (ex: Processo nº: ..., Autor: ..., Nome: ...).
- Conclua obrigatoriamente TODOS os 19 quesitos.
- Finalize com:
ANA PAULA DA COSTA HENRIQUES
Médica Perita - CRM-PE 11.395
{DATA_ATUAL_EXTENSO}
Assinatura Eletrônica

ESTRUTURA OBRIGATÓRIA DO LAUDO OFICIAL:

PODER JUDICIÁRIO DA UNIÃO
TRIBUNAL REGIONAL FEDERAL DA 5ª REGIÃO
JUSTIÇA FEDERAL DE PRIMEIRA INSTÂNCIA
SEÇÃO JUDICIÁRIA DE PERNAMBUCO - 19ª VARA

LAUDO DE EXAME MÉDICO-PERICIAL

**1. PREÂMBULO**
Processo nº:
Ação:
Órgão Julgador:
Autor:
Réu:
Data da Perícia:

**2. PERÍCIA MÉDICA**
Eu, Ana Paula da Costa Henriques, médico (CRM-PE 11.395), perito judicial nos autos do Processo abaixo discriminado, tendo realizado os levantamentos e pesquisas julgadas necessárias, venho apresentar o meu LAUDO PERICIAL.

**3. DADOS DO(A) PERICIADO(A)**
Nome: 
Identidade civil: 
CPF: 
Data do nascimento: 
Idade: 
Sexo biológico: 
Identidade de gênero: 
Escolaridade: 
Estado civil: 

**4. HISTÓRICO LABORAL DO(A) PERICIADO(A)**
Ocupação habitual: 
Descrição da atividade: 
Tempo de exercício da ocupação habitual: 
Data declarada de afastamento do trabalho: 
Experiência laboral anterior: 
Formação técnico-profissional: 
Reabilitação profissional: 

**5. HISTÓRICO**
Da análise da petição inicial e dos documentos apresentados, depreende-se que o(a) periciado(a) estaria acometido pela(s) seguinte(s) patologia(s) - CID: 

**6. HISTÓRICO DA DOENÇA ATUAL**

**7. EXAME CLÍNICO**
Sinais Vitais: PA: [X] mmHg. FC: [X] bpm.
Exame Geral: 
Aparelho Cardiorrespiratório: 
Exame Dermatológico: 
Aparelho Locomotor: 

**8. DOCUMENTOS AVALIADOS**
(Relação cronológica e análise crítica dos atestados, laudos e exames)

**9. CONCLUSÃO PERICIAL**
Data do Início da Doença (DID): 
Data do Início da Incapacidade (DII): 
O conjunto de patologias referenciado nos documentos médicos e administrativos juntados aos autos foi integralmente considerado no âmbito desta perícia, conduzida segundo protocolo técnico-científico próprio da medicina pericial.

Patologia / Condição Clínica | CID-10 | Enquadramento Pericial e Fundamentação Técnica

Conclusões diagnósticas: 
Esclarecimentos: 
Incapacidade: 

**10. QUESITOS DO JUÍZO**
**QUADRO I – QUESITAÇÃO PADRÃO**
**1. Identificação e Exame Clínico:**
**2. Vínculo com a Perita:**
**3. Profissão Declarada / Última Atividade:**
**4. Profissões Anteriores:**
**5. Correspondência entre Queixas e Exame Físico:**
**6. Doença, Sequela ou Deficiência (Exames Comprobatórios):**
**7. Incapacidade Laborativa e Natureza (Temporária ou Definitiva):**
**8. Intermitência da Incapacidade:**
**9. Afastamento Superior a 15 Dias:**
**10. Prazo Estimado para Tratamento/Recuperação:**
**11. Recuperação em até Dois Anos:**
**12. Evolução do Quadro (Progressivo, Regressivo ou Estável):**
**13. Extensão da Incapacidade (Total ou Parcial):**
**14. Possibilidade de Exercício da Atividade Habitual ou Readaptação:**
**15. Agravamento ou Progressão de Doença Pré-existente:**
**16. Atividades da Vida Diária / Independência:**
**17. Atos da Vida Civil:**
**18. Doenças Graves Especificadas em Lei (Art. 151 Lei 8.213/91):**
**19. Uso Contínuo de Medicação e Disponibilidade no SUS:**
**QUADRO II – QUESITOS ESPECÍFICOS PARA AUXÍLIO-DOENÇA E APOSENTADORIA POR INVALIDEZ:** Não se aplicam
**QUADRO III – QUESITOS ESPECÍFICOS PARA MENORES DE 16 ANOS:** Não se aplicam
**QUESITOS IV – QUESITOS ESPECÍFICOS PARA PORTADORES DE HIV/AIDS:** Não se aplicam
**QUADRO V – CONSIDERAÇÕES / AVALIAÇÃO FUNCIONAL:** Ver conclusão pericial

**11. ANEXOS (Fotos e Laudo da perícia trazidas pelo periciando(a))**

ANA PAULA DA COSTA HENRIQUES
Médica Perita - CRM-PE 11.395
{DATA_ATUAL_EXTENSO}
Assinatura Eletrônica
"""

# ----------------------------------------------------
# 3. Funções de Apoio e Gerador DOCX (Arial, 14 Título, 11 Corpo)
# ----------------------------------------------------
def extrair_nome_arquivo_laudo(texto_laudo: str) -> str:
    """Extrai o nome do periciado para gerar um arquivo no formato 'Laudo de [Nome].docx'."""
    match = re.search(r"(?:Nome|\*\*Nome\*\*)\s*:\s*([^\n\r]+)", str(texto_laudo or ""), re.IGNORECASE)
    if match:
        nome_bruto = match.group(1).replace("*", "").strip()
        nome_limpo = re.sub(r'[\\/*?:"<>|]', "", nome_bruto).strip()
        if nome_limpo:
            return f"Laudo de {nome_limpo}.docx"
    return "Laudo_Pericial_Dra_Aninha.docx"

def adicionar_paragrafo_com_negrito(paragrafo, texto: str, tamanho=11, bold_padrao=False):
    """Insere runs com fonte Arial e negrito condicional."""
    partes = re.split(r"(\*\*.*?\*\*)", texto)
    for p in partes:
        if not p:
            continue
        if p.startswith("**") and p.endswith("**"):
            run = paragrafo.add_run(p[2:-2])
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(tamanho)
        else:
            run = paragrafo.add_run(p)
            run.bold = bold_padrao
            run.font.name = "Arial"
            run.font.size = Pt(tamanho)

def gerar_docx_do_laudo(texto_laudo: str) -> io.BytesIO:
    doc = docx.Document()

    # Define Arial como a fonte padrão no estilo Normal do documento
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Arial"
    style_normal.font.size = Pt(11)

    # Margens padrão (2.0 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # 1. Brasão Proporcional e Discreto (0.7 polegadas)
    brasao_arquivo = None
    for nome_b in ["brasao.png", "brasao.jpg", "brasao.jpeg", "logo.png"]:
        if os.path.exists(nome_b):
            brasao_arquivo = nome_b
            break

    if brasao_arquivo:
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(8)
        run_logo = p_logo.add_run()
        run_logo.add_picture(brasao_arquivo, width=Inches(0.72))

    texto_seguro = str(texto_laudo or "")
    texto_seguro = texto_seguro.replace("###", "").replace("##", "")

    linhas = texto_seguro.split("\n")
    inicio_real = 0
    for idx, l in enumerate(linhas):
        l_upper = l.strip().upper().replace("*", "")
        if "PODER JUDICIÁRIO DA UNIÃO" in l_upper or "LAUDO DE EXAME MÉDICO-PERICIAL" in l_upper or "1. PREÂMBULO" in l_upper:
            inicio_real = idx
            break

    linhas_oficiais = linhas[inicio_real:]

    # Remove qualquer variação de assinatura gerada no corpo para padronizar no final
    linhas_corpo = []
    for l in linhas_oficiais:
        l_check = l.strip().replace("*", "")
        if any(termo in l_check for termo in [
            "ANA PAULA DA COSTA HENRIQUES",
            "Médica Perita - CRM-PE 11.395",
            "Assinatura Eletrônica"
        ]) or re.search(r"Recife,\s+\d{1,2}\s+de\s+[a-zA-ZçÇ]+\s+de\s+\d{4}", l_check):
            continue
        linhas_corpo.append(l)

    i = 0
    while i < len(linhas_corpo):
        linha = linhas_corpo[i].strip()
        if not linha:
            i += 1
            continue

        linha_sem_md = linha.replace("*", "").strip()

        # 2. Cabeçalho Institucional Centralizado (Arial, Tam 11 + Negrito)
        if any(h in linha_sem_md.upper() for h in [
            "PODER JUDICIÁRIO DA UNIÃO",
            "TRIBUNAL REGIONAL FEDERAL DA 5ª REGIÃO",
            "JUSTIÇA FEDERAL DE PRIMEIRA INSTÂNCIA",
            "SEÇÃO JUDICIÁRIA DE PERNAMBUCO - 19ª VARA"
        ]):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(linha_sem_md)
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(11)
            i += 1
            continue

        # 3. Título Principal em Destaque Alinhado à Esquerda (Arial, Tam 14 + Negrito)
        if "LAUDO DE EXAME MÉDICO-PERICIAL" in linha_sem_md.upper():
            p_titulo = doc.add_paragraph()
            p_titulo.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_titulo.paragraph_format.space_before = Pt(24)
            p_titulo.paragraph_format.space_after = Pt(16)
            run_titulo = p_titulo.add_run("LAUDO DE EXAME MÉDICO-PERICIAL")
            run_titulo.bold = True
            run_titulo.font.name = "Arial"
            run_titulo.font.size = Pt(14)
            i += 1
            continue

        # Tabela pericial (Arial, Tam 11)
        if "|" in linha:
            linhas_tabela = []
            while i < len(linhas_corpo) and "|" in linhas_corpo[i]:
                l_tab = linhas_corpo[i].strip()
                if not re.match(r"^\|?[\s\-:|]+\|?$", l_tab):
                    colunas = [c.strip().replace("**", "") for c in l_tab.strip("|").split("|")]
                    if any(colunas):
                        linhas_tabela.append(colunas)
                i += 1

            if linhas_tabela:
                num_cols = max(len(r) for r in linhas_tabela)
                table = doc.add_table(rows=len(linhas_tabela), cols=num_cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.style = "Table Grid"

                for r_idx, row_data in enumerate(linhas_tabela):
                    for c_idx, cell_value in enumerate(row_data):
                        if c_idx < num_cols:
                            cell = table.cell(r_idx, c_idx)
                            cell.text = cell_value
                            for p_c in cell.paragraphs:
                                for r_c in p_c.runs:
                                    r_c.font.name = "Arial"
                                    r_c.font.size = Pt(11)
                                    if r_idx == 0:
                                        r_c.bold = True
                p_espaco = doc.add_paragraph()
                p_espaco.paragraph_format.space_before = Pt(4)
            continue

        # Parágrafos e Seções (Arial, Tam 11)
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15

        eh_secao = any(linha_sem_md.startswith(f"{n}.") for n in range(1, 13)) or linha_sem_md.startswith("QUADRO")

        if eh_secao:
            # Título da seção (1. PREÂMBULO) tam 11 + negrito
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            adicionar_paragrafo_com_negrito(p, linha, tamanho=11, bold_padrao=True)
        else:
            # Corpo do laudo tam 11
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(3)
            linha_formatada = linha
            if any(campo in linha for campo in ["Processo nº:", "Ação:", "Órgão Julgador:", "Autor:", "Réu:", "Data da Perícia:"]):
                linha_formatada = linha.replace("**", "")
            adicionar_paragrafo_com_negrito(p, linha_formatada, tamanho=11, bold_padrao=False)

        i += 1

    # 4. Encerramento Centralizado Padronizado com Tamanhos Específicos
    # Nome da médica: Arial 11 + negrito
    p_ass1 = doc.add_paragraph()
    p_ass1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ass1.paragraph_format.space_before = Pt(32)
    p_ass1.paragraph_format.space_after = Pt(2)
    run_ass1 = p_ass1.add_run("ANA PAULA DA COSTA HENRIQUES")
    run_ass1.bold = True
    run_ass1.font.name = "Arial"
    run_ass1.font.size = Pt(11)

    # Função e CRM: Arial 11
    p_ass2 = doc.add_paragraph()
    p_ass2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ass2.paragraph_format.space_before = Pt(0)
    p_ass2.paragraph_format.space_after = Pt(4)
    run_ass2 = p_ass2.add_run("Médica Perita - CRM-PE 11.395")
    run_ass2.font.name = "Arial"
    run_ass2.font.size = Pt(11)

    # Local e data: Arial 10
    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_data.paragraph_format.space_before = Pt(0)
    p_data.paragraph_format.space_after = Pt(18)
    run_data = p_data.add_run(DATA_ATUAL_EXTENSO)
    run_data.font.name = "Arial"
    run_data.font.size = Pt(10)

    # Assinatura Eletrônica: Arial 9
    p_eletr = doc.add_paragraph()
    p_eletr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eletr.paragraph_format.space_before = Pt(0)
    p_eletr.paragraph_format.space_after = Pt(0)
    run_eletr = p_eletr.add_run("Assinatura Eletrônica")
    run_eletr.font.name = "Arial"
    run_eletr.font.size = Pt(9)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ----------------------------------------------------
# 4. Gerenciamento de Estado (Session State)
# ----------------------------------------------------
if "messages" not in st.session_state:
    st.session_state.messages = []

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if "ultimo_laudo_gerado" not in st.session_state:
    st.session_state.ultimo_laudo_gerado = ""

# ----------------------------------------------------
# 5. Barra Lateral (Sidebar)
# ----------------------------------------------------
with st.sidebar:
    col_v1, col_img, col_v2 = st.columns([1, 2, 1])
    with col_img:
        if os.path.exists("aninha.jpeg"):
            st.image("aninha.jpeg", width=140)
        elif os.path.exists("aninha.png"):
            st.image("aninha.png", width=140)
        else:
            st.markdown("<h1 style='text-align: center;'>👩‍⚕️</h1>", unsafe_allow_html=True)

    if st.button("🔄 Iniciar Novo Caso Pericial", use_container_width=True):
        st.session_state.messages = []
        st.session_state.ultimo_laudo_gerado = ""
        st.session_state.uploader_key += 1
        st.rerun()

    st.markdown("---")

    st.markdown("### 📁 Anexar Documentos")
    arquivos_anexos = st.file_uploader(
        "Envie relatórios, exames ou autos (Word, PDF, Imagens):",
        type=["pdf", "docx", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.uploader_key}",
        help="Selecione os documentos para análise pericial técnica."
    )

    if arquivos_anexos:
        st.info(f"📎 {len(arquivos_anexos)} documento(s) anexado(s).")

    st.markdown("---")

    st.markdown("### ⚡ Ações Rápidas")
    btn_laudo = st.button("🚀 Gerar Laudo Completo", type="primary", use_container_width=True)
    btn_dii = st.button("🗓️ Fixar DII/DID", use_container_width=True)
    btn_quesitos = st.button("📋 Responder Quesitos", use_container_width=True)

    if st.session_state.ultimo_laudo_gerado:
        st.markdown("---")
        st.markdown("### 💾 Exportar Laudo")
        docx_buffer = gerar_docx_do_laudo(st.session_state.ultimo_laudo_gerado)
        nome_arquivo_doc = extrair_nome_arquivo_laudo(st.session_state.ultimo_laudo_gerado)
        st.download_button(
            label="📥 Baixar Laudo em Word (.docx)",
            data=docx_buffer,
            file_name=nome_arquivo_doc,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.
