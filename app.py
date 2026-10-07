import os
import io
import re
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

SYSTEM_INSTRUCTION = """
Você é a Dra. Aninha, médica perita judicial previdenciária e assistente técnica pericial de alto nível.
Sua missão é atuar em auxílio à Dra. Ana Paula da Costa Henriques (CRM-PE 11.395), Perita Judicial Federal da Seção Judiciária de Pernambuco (19ª Vara - TRF5).

COMPORTAMENTO EM CONVERSAS E ANÁLISES (CHAT COMUM):
- Quando a médica enviar dados do periciando, resumos ou fizer perguntas, atue como colega perita consultora.
- Analise criticamente os fatos, aponte a correlação biomecânica com a profissão habitual, discuta a existência ou não de incapacidade laborativa (temporária, total/parcial), sugira marcos de DID e DII e discuta a fundamentação conforme a Lei 8.213/91.
- Responda de forma clara, direta e técnica no chat. NÃO gere a minuta formal de 11 tópicos a menos que seja explicitamente solicitado o laudo completo pelo botão oficial.
- Mantenha títulos e destaques sempre em negrito para facilitar a leitura.

AO RESPONDER AOS QUESITOS DO JUÍZO:
- Identifique cada quesito em negrito com seu tema (ex: **Quesito 1 - Identificação:**, **Quesito 7 - Incapacidade laborativa:**, etc.) e forneça a resposta técnica pericial fundamentada logo abaixo.
- Conclua obrigatoriamente TODOS os 19 quesitos do Quadro I do Juízo da 19ª Vara, do 1 ao 19 sem interrupção.

QUANDO FOR SOLICITADO O LAUDO PERICIAL OFICIAL:
- NUNCA inicie sua resposta com saudações, introduções ou conversas prévias.
- SUA RESPOSTA DEVE COMEÇAR DIRETAMENTE NA PRIMEIRA LINHA DO CABEÇALHO OFICIAL: "PODER JUDICIÁRIO DA UNIÃO".
- Utilize negrito nos títulos de seções, rótulos de campos e identificadores dos quesitos.
- Conclua obrigatoriamente TODOS os 19 quesitos e encerre com a data e assinatura da Dra. Ana Paula da Costa Henriques.

ESTRUTURA OBRIGATÓRIA DO LAUDO OFICIAL:

PODER JUDICIÁRIO DA UNIÃO
TRIBUNAL REGIONAL FEDERAL DA 5ª REGIÃO
JUSTIÇA FEDERAL DE PRIMEIRA INSTÂNCIA
SEÇÃO JUDICIÁRIA DE PERNAMBUCO - 19ª VARA

LAUDO DE EXAME MÉDICO-PERICIAL

**1. PREÂMBULO**
**Processo nº:**
**Ação:**
**Órgão Julgador:**
**Autor:**
**Réu:**
**Data da Perícia:**

**2. PERÍCIA MÉDICA**
Eu, Ana Paula da Costa Henriques, médico (CRM-PE 11.395), perito judicial nos autos do Processo abaixo discriminado, tendo realizado os levantamentos e pesquisas julgadas necessárias, venho apresentar o meu LAUDO PERICIAL.

**3. DADOS DO(A) PERICIADO(A)**
**Nome:** 
**Identidade civil:** 
**CPF:** 
**Data do nascimento:** 
**Idade:** 
**Sexo biológico:** 
**Identidade de gênero:** 
**Escolaridade:** 
**Estado civil:** 

**4. HISTÓRICO LABORAL DO(A) PERICIADO(A)**
**Ocupação habitual:** 
**Descrição da atividade:** 
**Tempo de exercício da ocupação habitual:** 
**Data declarada de afastamento do trabalho:** 
**Experiência laboral anterior:** 
**Formação técnico-profissional:** 
**Reabilitação profissional:** 

**5. HISTÓRICO**
Da análise da petição inicial e dos documentos apresentados, depreende-se que o(a) periciado(a) estaria acometido pela(s) seguinte(s) patologia(s) - CID: 

**6. HISTÓRICO DA DOENÇA ATUAL**

**7. EXAME CLÍNICO**
**Sinais Vitais:** PA: [X] mmHg. FC: [X] bpm.
**Exame Geral:** 
**Aparelho Cardiorrespiratório:** 
**Exame Dermatológico:** 
**Aparelho Locomotor:** 

**8. DOCUMENTOS AVALIADOS**
(Relação cronológica e análise crítica dos atestados, laudos e exames)

**9. CONCLUSÃO PERICIAL**
**Data do Início da Doença (DID):** 
**Data do Início da Incapacidade (DII):** 
O conjunto de patologias referenciado nos documentos médicos e administrativos juntados aos autos foi integralmente considerado no âmbito desta perícia, conduzida segundo protocolo técnico-científico próprio da medicina pericial. O procedimento adotado observou as seguintes etapas: identificação da demanda judicial e de seu objeto; levantamento e qualificação do periciado; coleta dos dados específicos da perícia; mapeamento das enfermidades alegadas e de sua evolução clínica; realização de anamnese e exame físico e/ou mental; apreciação crítica de atestados, laudos, exames complementares e documentos administrativos; cotejo com perícias anteriores eventualmente existentes; síntese integrativa de todos os elementos colhidos; e elaboração das respostas técnicas aos quesitos formulados pelo Juízo e pelas partes.

Patologia / Condição Clínica | CID-10 | Enquadramento Pericial e Fundamentação Técnica

**Conclusões diagnósticas:** 
**Esclarecimentos:** 
**Incapacidade:** 

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

Recife, [Data da Perícia].

**ANA PAULA DA COSTA HENRIQUES**
Médica Perita - CRM-PE 11.395
Assinatura Eletrônica
"""

# ----------------------------------------------------
# 3. Função para Gerar Arquivo .DOCX Limpo e com Negrito Real
# ----------------------------------------------------
def adicionar_paragrafo_com_negrito(paragrafo, texto: str, tamanho=10, bold_padrao=False):
    partes = re.split(r"(\*\*.*?\*\*)", texto)
    for p in partes:
        if not p:
            continue
        if p.startswith("**") and p.endswith("**"):
            run = paragrafo.add_run(p[2:-2])
            run.bold = True
            run.font.size = Pt(tamanho)
        else:
            run = paragrafo.add_run(p)
            run.bold = bold_padrao
            run.font.size = Pt(tamanho)

def gerar_docx_do_laudo(texto_laudo: str) -> io.BytesIO:
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    brasao_arquivo = None
    for nome_b in ["brasao.png", "brasao.jpg", "brasao.jpeg", "logo.png"]:
        if os.path.exists(nome_b):
            brasao_arquivo = nome_b
            break

    if brasao_arquivo:
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(2)
        run_logo = p_logo.add_run()
        run_logo.add_picture(brasao_arquivo, width=Inches(1.1))

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

    i = 0
    while i < len(linhas_oficiais):
        linha = linhas_oficiais[i].strip()
        if not linha:
            i += 1
            continue

        linha_sem_md = linha.replace("*", "").strip()

        # Cabeçalho Oficial
        if any(h in linha_sem_md.upper() for h in [
            "PODER JUDICIÁRIO DA UNIÃO",
            "TRIBUNAL REGIONAL FEDERAL DA 5ª REGIÃO",
            "JUSTIÇA FEDERAL DE PRIMEIRA INSTÂNCIA",
            "SEÇÃO JUDICIÁRIA DE PERNAMBUCO - 19ª VARA",
            "LAUDO DE EXAME MÉDICO-PERICIAL"
        ]):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(linha_sem_md)
            run.bold = True
            run.font.size = Pt(11)
            i += 1
            continue

        # Assinatura Oficial
        if any(ass in linha_sem_md for ass in ["ANA PAULA DA COSTA HENRIQUES", "Médica Perita - CRM-PE 11.395", "Assinatura Eletrônica"]):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(linha_sem_md)
            run.bold = True
            run.font.size = Pt(10)
            i += 1
            continue

        # Tabela pericial
        if "|" in linha:
            linhas_tabela = []
            while i < len(linhas_oficiais) and "|" in linhas_oficiais[i]:
                l_tab = linhas_oficiais[i].strip()
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
                            if r_idx == 0:
                                for p_c in cell.paragraphs:
                                    for r_c in p_c.runs:
                                        r_c.bold = True
                p_espaco = doc.add_paragraph()
                p_espaco.paragraph_format.space_before = Pt(4)
            continue

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(3)

        eh_secao = any(linha_sem_md.startswith(f"{n}.") for n in range(1, 13)) or linha_sem_md.startswith("QUADRO")

        if eh_secao:
            p.paragraph_format.space_before = Pt(8)
            adicionar_paragrafo_com_negrito(p, linha, tamanho=10.5, bold_padrao=True)
        else:
            adicionar_paragrafo_com_negrito(p, linha, tamanho=10, bold_padrao=False)

        i += 1

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
        st.download_button(
            label="📥 Baixar Laudo em Word (.docx)",
            data=docx_buffer,
            file_name="Laudo_Pericial_Dra_Aninha.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# ----------------------------------------------------
# 6. Captura de Entrada e Gerenciamento do Fluxo
# ----------------------------------------------------
col_header_avatar, col_header_text = st.columns([1, 8], vertical_alignment="center")

with col_header_avatar:
    if os.path.exists("aninha.jpeg"):
        st.image("aninha.jpeg", width=75)
    elif os.path.exists("aninha.png"):
        st.image("aninha.png", width=75)
    else:
        st.markdown("## 👩‍⚕️")

with col_header_text:
    st.markdown("## Dra. Aninha — Assistente de Perícia Previdenciária")
    st.caption("Padrão Oficial TRF5 / 19ª Vara — Dra. Ana Paula da Costa Henriques (CRM-PE 11.395)")

st.markdown("---")

prompt_usuario = st.chat_input("Digite dados do periciando, exame clínico ou orientações...")

prompt_acionado = None
eh_pedido_laudo = False

if prompt_usuario:
    prompt_acionado = prompt_usuario
    eh_pedido_laudo = False
elif btn_laudo:
    eh_pedido_laudo = True
    prompt_acionado = (
        "Elabore a minuta completa do LAUDO DE EXAME MÉDICO-PERICIAL oficial da 19ª Vara / TRF5, "
        "com base estritamente nos documentos anexados e nos fatos informados na discussão. "
        "Comece DIRETAMENTE pelo cabeçalho institucional (PODER JUDICIÁRIO DA UNIÃO), sem mensagens prévias. "
        "Mantenha todos os títulos, rótulos e números de quesitos em negrito. "
        "Preencha todos os 11 itens oficiais, fundamentando DID e DII, o quadro comparativo de patologias, "
        "respondendo obrigatoriamente a todos os 19 quesitos do Juízo do 1 ao 19 com fundamentação técnica "
        "e finalizando com o encerramento formal da Dra. Ana Paula da Costa Henriques (CRM-PE 11.395)."
    )
elif btn_dii:
    prompt_acionado = (
        "Com base nos autos clínicos e relatórios apresentados, proceda à análise rigorosa "
        "dos marcos temporais conforme a Lei 8.213/91. Fixe e justifique detalhadamente a "
        "DID (Data de Início da Doença) e a DII (Data de Início da Incapacidade), apontando os "
        "documentos probatórios que sustentam cada marco."
    )
elif btn_quesitos:
    prompt_acionado = (
        "Com base nos documentos médicos e no histórico do caso, responda aos 19 quesitos padrão do Juízo "
        "(Quadro I da 19ª Vara). Apresente cada item no formato: '**Quesito [Número] - [Tema]:**' seguido da resposta "
        "técnica pericial fundamentada. Responda a todos os 19 quesitos, do 1 ao 19, de forma completa."
    )

if prompt_acionado:
    texto_card_usuario = "🚀 **Solicitação:** Elaborar e formatar o Laudo Médico-Pericial oficial completo." if eh_pedido_laudo else prompt_acionado
    st.session_state.messages.append({"role": "user", "content": texto_card_usuario})

if not st.session_state.messages:
    with st.chat_message("assistant", avatar="aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️"):
        st.markdown(
            "Olá, Doutora! Sou a **Dra. Aninha**, sua assistente técnica de Perícia Médica Judicial.\n\n"
            "Envie os dados do caso no chat ou anexe os documentos na barra lateral para analisarmos. "
            "Quando desejar formalizar a minuta oficial, basta clicar em **🚀 Gerar Laudo Completo**!"
        )

for msg in st.session_state.messages:
    avatar_icon = ("aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️") if msg["role"] == "assistant" else None
    with st.chat_message(msg["role"], avatar=avatar_icon):
        st.markdown(msg["content"])
        if msg.get("is_laudo_card", False) and st.session_state.ultimo_laudo_gerado:
            doc_buf_msg = gerar_docx_do_laudo(st.session_state.ultimo_laudo_gerado)
            st.download_button(
                label="📥 Baixar Laudo em Word (.docx)",
                data=doc_buf_msg,
                file_name="Laudo_Pericial_Dra_Aninha.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_history_{msg.get('key_id', 0)}"
            )

# ----------------------------------------------------
# 7. Execução da IA com Limite Estável (8192) e Stream Seguro
# ----------------------------------------------------
if prompt_acionado:
    if not api_key:
        st.error("Chave GEMINI_API_KEY não configurada no ambiente (.env ou Secrets)!")
    else:
        contents = []
        texto_documentos_docx = ""

        if arquivos_anexos:
            for arq in arquivos_anexos:
                nome_baixo = arq.name.lower()
                dados_arquivo = arq.getvalue()

                if nome_baixo.endswith(".docx"):
                    try:
                        doc = docx.Document(io.BytesIO(dados_arquivo))
                        paragrafos = [p.text for p in doc.paragraphs if p.text.strip()]
                        texto_extraido = "\n".join(paragrafos)
                        texto_documentos_docx += f"\n\n--- DOCUMENTO WORD ANEXADO: {arq.name} ---\n{texto_extraido}\n"
                    except Exception as e_docx:
                        st.warning(f"Não foi possível ler o texto do arquivo {arq.name}: {e_docx}")
                elif nome_baixo.endswith(".pdf"):
                    contents.append(
                        types.Part.from_bytes(
                            data=dados_arquivo,
                            mime_type="application/pdf"
                        )
                    )
                elif nome_baixo.endswith((".png", ".jpg", ".jpeg")):
                    mime = "image/png" if nome_baixo.endswith(".png") else "image/jpeg"
                    contents.append(
                        types.Part.from_bytes(
                            data=dados_arquivo,
                            mime_type=mime
                        )
                    )

        historico_texto = "\n--- HISTÓRICO DA DISCUSSÃO PERICIAL ---\n"
        for m in st.session_state.messages[:-1]:
            papel = "MÉDICO" if m["role"] == "user" else "DRA. ANINHA"
            historico_texto += f"{papel}: {m['content']}\n"

        if texto_documentos_docx:
            historico_texto += f"\nCONTEÚDO EXTRAÍDO DOS ARQUIVOS WORD:\n{texto_documentos_docx}\n"

        if eh_pedido_laudo:
            demanda_final = (
                f"{prompt_acionado}\n\n[INSTRUÇÃO CRUCIAL: Elabore o laudo COMPLETO de ponta a ponta sem cortes. "
                "Comece em 'PODER JUDICIÁRIO DA UNIÃO' e responda a todos os 19 quesitos na íntegra até o encerramento com a assinatura.]"
            )
        else:
            demanda_final = prompt_acionado

        historico_texto += f"\nNOVA DEMANDA:\n{demanda_final}"
        contents.append(historico_texto)

        # Configuração estável com max_output_tokens=8192 (limite máximo suportado) e thinking=0
        config_ia = types.GenerateContentConfig(
            system_instruction=SYSTEM_INSTRUCTION,
            max_output_tokens=8192,
            thinking_config=types.ThinkingConfig(thinking_budget=0)
        )

        with st.chat_message("assistant", avatar="aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️"):
            status_box = st.empty()
            try:
                # 1. BOTÃO GERAR LAUDO: Coleta via stream em segundo plano para máxima confiabilidade
                if eh_pedido_laudo:
                    status_box.info("⏳ Só um momento, Dra Aninha está gerando o laudo solicitado...")
                    
                    stream_laudo = client.models.generate_content_stream(
                        model="gemini-3.6-flash",
                        contents=contents,
                        config=config_ia
                    )
                    
                    partes_laudo = []
                    for chunk in stream_laudo:
                        if chunk.text:
                            partes_laudo.append(chunk.text)
                    
                    texto_laudo = "".join(partes_laudo).strip()
                    status_box.empty()

                    if not texto_laudo:
                        st.error("Não foi possível gerar a minuta do laudo. Tente novamente.")
                    else:
                        st.session_state.ultimo_laudo_gerado = texto_laudo
                        doc_buf = gerar_docx_do_laudo(texto_laudo)

                        msg_sucesso = (
                            "✅ **Laudo Pericial Oficial elaborado com sucesso!**\n\n"
                            "A minuta oficial do TRF5 (19ª Vara) com os 11 tópicos padronizados, análise cronológica, "
                            "fundamentação de DID/DII e quesitos foi formatada no arquivo Word abaixo."
                        )
                        st.markdown(msg_sucesso)
                        st.download_button(
                            label="📥 Baixar Laudo Oficial em Word (.docx)",
                            data=doc_buf,
                            file_name="Laudo_Pericial_Dra_Aninha.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_now_{len(st.session_state.messages)}"
                        )
                        st.session_state.messages.append({
                            "role": "assistant",
                            "content": msg_sucesso,
                            "is_laudo_card": True,
                            "key_id": len(st.session_state.messages)
                        })

                # 2. CHAT / DISCUSSÃO: Streaming em tempo real
                else:
                    status_box.info("⏳ Só um momento, Dra Aninha responderá em breve...")
                    response_stream = client.models.generate_content_stream(
                        model="gemini-3.6-flash",
                        contents=contents,
                        config=config_ia
                    )

                    def stream_com_limpeza(stream):
                        limpou = False
                        for chunk in stream:
                            if chunk.text:
                                if not limpou:
                                    status_box.empty()
                                    limpou = True
                                yield chunk.text

                    resposta_completa = st.write_stream(stream_com_limpeza(response_stream))
                    status_box.empty()
                    st.session_state.messages.append({"role": "assistant", "content": resposta_completa})

            except Exception as err:
                status_box.empty()
                st.error(f"Erro na resposta da Dra. Aninha: {err}")
