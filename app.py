import os
import io
import time
import streamlit as st
from PIL import Image
from dotenv import load_dotenv
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types

# ----------------------------------------------------
# 1. Configurações Iniciais da Página
# ----------------------------------------------------
load_dotenv()

# Ícone oficial da aba com a foto da Aninha
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

# Estilização CSS: Botão primário azul na barra lateral
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

Ao elaborar o LAUDO PERICIAL, siga RIGOROSAMENTE e de forma padronizada a seguinte estrutura oficial:

PODER JUDICIÁRIO DA UNIÃO
TRIBUNAL REGIONAL FEDERAL DA 5ª REGIÃO
JUSTIÇA FEDERAL DE PRIMEIRA INSTÂNCIA
SEÇÃO JUDICIÁRIA DE PERNAMBUCO - 19ª VARA
LAUDO DE EXAME MÉDICO-PERICIAL

1. PREÂMBULO (Processo nº, Ação, Órgão Julgador, Autor, Réu, Data da Perícia)
2. PERÍCIA MÉDICA (Eu, Ana Paula da Costa Henriques, médica perita judicial...)
3. DADOS DO(A) PERICIADO(A) (Nome, RG, CPF, Nascimento, Idade, Sexo biológico, Gênero, Escolaridade, Estado civil)
4. HISTÓRICO LABORAL DO(A) PERICIADO(A) (Ocupação habitual, descrição detalhada da atividade e esforço biomecânico, tempo de exercício, data de afastamento, experiência anterior, reabilitação)
5. HISTÓRICO / PATOLOGIAS ALEGADAS (CID alegados na inicial e documentos)
6. HISTÓRICO DA DOENÇA ATUAL (HDA - cronologia dos sintomas, tratamentos realizados, uso de analgésicos/fisioterapia)
7. EXAME CLÍNICO (Sinais vitais, Exame Geral, Aparelho Cardiorrespiratório, Exame Dermatológico, Aparelho Locomotor / Testes ortopédicos / neurológicos específicos)
8. DOCUMENTOS AVALIADOS (Relação cronológica e análise crítica de atestados, laudos, exames complementares de imagem/laboratoriais)
9. CONCLUSÃO PERICIAL:
   - Data do Início da Doença (DID) fundamentada
   - Data do Início da Incapacidade (DII) fundamentada
   - Síntese técnico-científica e fundamentação pericial
   - Tabela / Quadro pericial: Patologia | CID-10 | Enquadramento Pericial e Fundamentação Técnica
   - Conclusões diagnósticas e Esclarecimentos
   - Incapacidade (Ausente / Temporária / Permanente / Parcial / Total / Não se aplica)
10. QUESITOS DO JUÍZO:
   - Responda objetiva e fundamentadamente aos quesitos de 1 a 19 do QUADRO I (Quesitação Padrão do Juízo).
   - Indique 'Não se aplicam' para os Quadros II, III e IV se for o caso.
   - Quadro V: 'Ver conclusão pericial'.
11. ENCERRAMENTO E ASSINATURA:
   Recife, [Data atual].
   ANA PAULA DA COSTA HENRIQUES
   Médica Perita - CRM-PE 11.395

Diretrizes Técnicas:
- Fundamente com base na Lei 8.213/91 e na biomecânica da ocupação habitual.
- Nunca afirme incapacidade sem evidência de limitação funcional demonstrada documentalmente ou no exame físico.
"""

# ----------------------------------------------------
# 3. Função para Gerar Arquivo .DOCX
# ----------------------------------------------------
def gerar_docx_do_laudo(texto_laudo: str) -> io.BytesIO:
    doc = docx.Document()

    # Configuração das margens (padrão 2cm)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Cabeçalho Principal Centralizado
    p_cabecalho = doc.add_paragraph()
    p_cabecalho.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_cab = p_cabecalho.add_run(
        "PODER JUDICIÁRIO DA UNIÃO\n"
        "TRIBUNAL REGIONAL FEDERAL DA 5ª REGIÃO\n"
        "JUSTIÇA FEDERAL DE PRIMEIRA INSTÂNCIA\n"
        "SEÇÃO JUDICIÁRIA DE PERNAMBUCO - 19ª VARA\n\n"
        "LAUDO DE EXAME MÉDICO-PERICIAL\n"
    )
    run_cab.bold = True
    run_cab.font.size = Pt(11)

    # Processamento do texto linha a linha
    linhas = texto_laudo.split("\n")
    for linha in linhas:
        l_strip = linha.strip()
        if not l_strip:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)

        # Detecta se é título de seção (ex: 1. PREÂMBULO, 9. CONCLUSÃO)
        eh_secao = any(l_strip.startswith(f"{i}.") for i in range(1, 13)) or l_strip.startswith("QUADRO")

        if eh_secao or l_strip.startswith("#"):
            run = p.add_run(l_strip.replace("#", "").strip())
            run.bold = True
            run.font.size = Pt(11)
            p.paragraph_format.space_before = Pt(8)
        else:
            run = p.add_run(l_strip)
            run.font.size = Pt(10)

    # Rodapé / Assinatura
    p_rodape = doc.add_paragraph()
    p_rodape.paragraph_format.space_before = Pt(20)
    p_rodape.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_rod = p_rodape.add_run(
        "\n_________________________________________________\n"
        "ANA PAULA DA COSTA HENRIQUES\n"
        "Médica Perita - CRM-PE 11.395"
    )
    run_rod.bold = True
    run_rod.font.size = Pt(10)

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

    # Reiniciar caso
    if st.button("🔄 Iniciar Novo Caso Pericial", use_container_width=True):
        st.session_state.messages = []
        st.session_state.ultimo_laudo_gerado = ""
        st.session_state.uploader_key += 1
        st.rerun()

    st.markdown("---")

    # Upload com suporte a DOCX, PDF e Imagens
    st.markdown("### 📁 Anexar Documentos")
    arquivos_anexos = st.file_uploader(
        "Envie relatórios, exames ou autos (Word, PDF, Imagens):",
        type=["pdf", "docx", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.uploader_key}",
        help="Selecione os documentos do caso pericial."
    )

    if arquivos_anexos:
        st.info(f"📎 {len(arquivos_anexos)} documento(s) anexado(s).")

    st.markdown("---")

    # Ações Rápidas
    st.markdown("### ⚡ Ações Rápidas")
    btn_laudo = st.button("🚀 Gerar Laudo Completo", type="primary", use_container_width=True)
    btn_dii = st.button("🗓️ Fixar DII/DID", use_container_width=True)
    btn_quesitos = st.button("📋 Responder Quesitos", use_container_width=True)

    # Exibe o botão de baixar Word se houver laudo gerado
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
# 6. Interface Principal
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

if not st.session_state.messages:
    with st.chat_message("assistant", avatar="aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️"):
        st.markdown(
            "Olá, Doutora! Sou a **Dra. Aninha**, sua assistente técnica de Perícia Médica Judicial.\n\n"
            "Estou configurada com o padrão oficial do seu **Laudo Pericial (TRF5 - 19ª Vara)**. "
            "Pode ditar os dados do caso, ou anexar os laudos e autos em **Word (.docx), PDF ou Imagem** "
            "na barra lateral e clicar em **🚀 Gerar Laudo Completo** para redigir o laudo e disponibilizar o arquivo `.docx` para download!"
        )

for msg in st.session_state.messages:
    avatar_icon = ("aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️") if msg["role"] == "assistant" else None
    with st.chat_message(msg["role"], avatar=avatar_icon):
        st.markdown(msg["content"])

# ----------------------------------------------------
# 7. Processamento e Streaming
# ----------------------------------------------------
prompt_usuario = st.chat_input("Digite dados do periciando, exame clínico ou orientações...")

prompt_acionado = None
eh_pedido_laudo = False

if prompt_usuario:
    prompt_acionado = prompt_usuario
elif btn_laudo:
    eh_pedido_laudo = True
    prompt_acionado = (
        "Com base em todos os documentos anexados e nos fatos informados, elabore a minuta completa "
        "do LAUDO DE EXAME MÉDICO-PERICIAL oficial, preenchendo todos os 11 itens: Preâmbulo, Perícia Médica, "
        "Dados do Periciado, Histórico Laboral com esforço biomecânico, Histórico de Patologias (CID), "
        "Histórico da Doença Atual (HDA), Exame Clínico, Documentos Avaliados em ordem cronológica, "
        "Conclusão Pericial com fundamentação técnica de DID/DII e tabela comparativa, respostas a todos os "
        "19 quesitos do Quadro I do Juízo, e o encerramento com assinatura da Dra. Ana Paula da Costa Henriques."
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
        "Com base nos documentos médicos e no histórico do caso, responda de forma técnica, "
        "precisa e conclusiva a todos os 19 quesitos padrão do Quadro I do Juízo da 19ª Vara."
    )

if prompt_acionado:
    if not api_key:
        st.error("Chave GEMINI_API_KEY não configurada no ambiente (.env ou Secrets)!")
    else:
        st.session_state.messages.append({"role": "user", "content": prompt_acionado})
        with st.chat_message("user"):
            st.markdown(prompt_acionado)

        contents = []
        texto_documentos_docx = ""

        # Processamento inteligente de anexos
        if arquivos_anexos:
            for arq in arquivos_anexos:
                nome_baixo = arq.name.lower()
                dados_arquivo = arq.getvalue()

                # Se for Word (.docx), extrai o texto diretamente
                if nome_baixo.endswith(".docx"):
                    try:
                        doc = docx.Document(io.BytesIO(dados_arquivo))
                        paragrafos = [p.text for p in doc.paragraphs if p.text.strip()]
                        texto_extraido = "\n".join(paragrafos)
                        texto_documentos_docx += f"\n\n--- DOCUMENTO WORD ANEXADO: {arq.name} ---\n{texto_extraido}\n"
                    except Exception as e_docx:
                        st.warning(f"Não foi possível ler o texto do arquivo {arq.name}: {e_docx}")

                # Se for PDF
                elif nome_baixo.endswith(".pdf"):
                    contents.append(
                        types.Part.from_bytes(
                            data=dados_arquivo,
                            mime_type="application/pdf"
                        )
                    )
                # Se for imagem
                elif nome_baixo.endswith((".png", ".jpg", ".jpeg")):
                    mime = "image/png" if nome_baixo.endswith(".png") else "image/jpeg"
                    contents.append(
                        types.Part.from_bytes(
                            data=dados_arquivo,
                            mime_type=mime
                        )
                    )

        # Contexto da conversa
        historico_texto = "\n--- HISTÓRICO DA DISCUSSÃO PERICIAL ---\n"
        for m in st.session_state.messages[:-1]:
            papel = "MÉDICO" if m["role"] == "user" else "DRA. ANINHA"
            historico_texto += f"{papel}: {m['content']}\n"

        if texto_documentos_docx:
            historico_texto += f"\nCONTEÚDO EXTRAÍDO DOS ARQUIVOS WORD:\n{texto_documentos_docx}\n"

        historico_texto += f"\nNOVA DEMANDA:\n{prompt_acionado}"
        contents.append(historico_texto)

        config_ia = types.GenerateContentConfig(
            system_instruction=SYSTEM_INSTRUCTION
        )

        with st.chat_message("assistant", avatar="aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️"):
            try:
                response_stream = client.models.generate_content_stream(
                    model="gemini-3.6-flash",
                    contents=contents,
                    config=config_ia
                )
                resposta_completa = st.write_stream(
                    chunk.text for chunk in response_stream if chunk.text
                )
                st.session_state.messages.append({"role": "assistant", "content": resposta_completa})

                # Se foi gerado um laudo (ou se a resposta contém a estrutura pericial), disponibiliza para download
                if eh_pedido_laudo or "LAUDO DE EXAME MÉDICO-PERICIAL" in resposta_completa or "PREÂMBULO" in resposta_completa:
                    st.session_state.ultimo_laudo_gerado = resposta_completa
                    doc_buf = gerar_docx_do_laudo(resposta_completa)
                    st.download_button(
                        label="📥 Baixar Este Laudo em Word (.docx)",
                        data=doc_buf,
                        file_name="Laudo_Pericial_Dra_Aninha.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{len(st.session_state.messages)}"
                    )
            except Exception as err:
                st.error(f"Erro na resposta da Dra. Aninha: {err}")
