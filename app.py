import os
import time
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import types

# ----------------------------------------------------
# 1. Configurações Iniciais da Página
# ----------------------------------------------------
load_dotenv()

st.set_page_config(
    page_title="Dra. Aninha - Perícia Médica Previdenciária",
    page_icon="⚖️",
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
# 2. Inicialização do Cliente Gemini e Instrução do Sistema
# ----------------------------------------------------
api_key = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

SYSTEM_INSTRUCTION = """
Você é a Dra. Aninha, médica perita judicial previdenciária e assistente técnica pericial experiente.
Sua missão é auxiliar peritos e operadores do direito na elaboração de laudos periciais judiciais,
análise de incapacidade laborativa (temporária ou permanente, parcial ou total), fixação técnica e
fundamentada de marcos temporais clínicos (DID - Data de Início da Doença e DII - Data de Início da Incapacidade)
e respostas aos quesitos judiciais e das partes.

Diretrizes:
1. Mantenha tom estritamente técnico, formal, pericial e fundamentado na literatura médica e na legislação previdenciária (Lei 8.213/91).
2. Sempre correlacione os achados clínicos e os exames de imagem/laboratoriais com a atividade profissional desempenhada pelo periciando.
3. Não presuma incapacidade sem respaldo em elementos comprobatórios de limitação funcional para a função habitual.
"""

# ----------------------------------------------------
# 3. Gerenciamento do Estado da Sessão (Session State)
# ----------------------------------------------------
if "messages" not in st.session_state:
    st.session_state.messages = []

# ----------------------------------------------------
# 4. Barra Lateral (Sidebar)
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
        st.rerun()

    st.markdown("---")

    st.markdown("### 📁 Anexar Documentos")
    arquivos_anexos = st.file_uploader(
        "Envie relatórios, exames ou autos (PDF ou Imagens):",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        help="Selecione um ou mais laudos/exames para análise técnica pericial."
    )

    if arquivos_anexos:
        st.info(f"📎 {len(arquivos_anexos)} documento(s) carregado(s).")

    st.markdown("---")

    st.markdown("### ⚡ Ações Rápidas")
    btn_laudo = st.button("🚀 Gerar Laudo", type="primary", use_container_width=True)
    btn_dii = st.button("🗓️ Fixar DII/DID", use_container_width=True)
    btn_quesitos = st.button("⚖️ Responder Quesitos", use_container_width=True)

# ----------------------------------------------------
# 5. Interface Principal (Header e Chat)
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
    st.caption("Análise pericial técnica judicial, incapacidade laborativa, DII/DID e quesitos oficiais.")

st.markdown("---")

if not st.session_state.messages:
    with st.chat_message("assistant", avatar="aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️"):
        st.markdown(
            "Olá, Doutor(a)! Sou a **Aninha**, sua assistente técnica em Perícia Previdenciária.\n\n"
            "Você pode me ditar o histórico clínico, informar a profissão do periciando ou "
            "**anexar relatórios e exames (PDF ou imagem)** na barra lateral para analisarmos "
            "juntos a capacidade laborativa, fixação de DID/DII e respostas aos quesitos!"
        )

for msg in st.session_state.messages:
    avatar_icon = ("aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️") if msg["role"] == "assistant" else None
    with st.chat_message(msg["role"], avatar=avatar_icon):
        st.markdown(msg["content"])

# ----------------------------------------------------
# 6. Processamento de Entradas (Chat ou Botões de Ação)
# ----------------------------------------------------
prompt_usuario = st.chat_input("Digite detalhes do periciando, exames, perguntas ou orientações...")

prompt_acionado = None
if prompt_usuario:
    prompt_acionado = prompt_usuario
elif btn_laudo:
    prompt_acionado = (
        "Com base em todo o caso pericial em discussão e nos documentos anexados, "
        "elabore a minuta estruturada do Laudo Médico Pericial Judicial completo "
        "(identificação, histórico clínico-ocupacional, análise crítica dos exames/laudos, "
        "discussão técnica detalhada da capacidade laborativa, fixação fundamentada da DID e DII, "
        "e conclusão pericial formal)."
    )
elif btn_dii:
    prompt_acionado = (
        "Com base nos autos clínicos e relatórios apresentados, proceda à análise rigorosa "
        "dos marcos temporais com base na Lei 8.213/91. Fixe e justifique detalhadamente a "
        "DID (Data de Início da Doença) e a DII (Data de Início da Incapacidade), apontando os "
        "documentos probatórios que sustentam cada marco."
    )
elif btn_quesitos:
    prompt_acionado = (
        "Com base nos documentos médicos e no histórico do caso, responda de forma técnica, "
        "precisa e conclusiva aos quesitos periciais apresentados (do Juízo e das partes)."
    )

if prompt_acionado:
    if not api_key:
        st.error("Chave GEMINI_API_KEY não configurada no ambiente (.env ou Secrets)!")
    else:
        st.session_state.messages.append({"role": "user", "content": prompt_acionado})
        with st.chat_message("user"):
            st.markdown(prompt_acionado)

        contents = []

        if arquivos_anexos:
            for arq in arquivos_anexos:
                tipo_mime = arq.type
                if not tipo_mime:
                    if arq.name.lower().endswith(".pdf"):
                        tipo_mime = "application/pdf"
                    elif arq.name.lower().endswith((".jpg", ".jpeg")):
                        tipo_mime = "image/jpeg"
                    elif arq.name.lower().endswith(".png"):
                        tipo_mime = "image/png"
                    else:
                        tipo_mime = "application/octet-stream"

                contents.append(
                    types.Part.from_bytes(
                        data=arq.getvalue(),
                        mime_type=tipo_mime
                    )
                )

        historico_texto = "\n--- HISTÓRICO DA DISCUSSÃO PERICIAL ---\n"
        for m in st.session_state.messages[:-1]:
            papel = "MÉDICO/USUÁRIO" if m["role"] == "user" else "DRA. ANINHA"
            historico_texto += f"{papel}: {m['content']}\n"
        historico_texto += f"\nNOVA DEMANDA:\n{prompt_acionado}"
        contents.append(historico_texto)

        config_rapida = types.GenerateContentConfig(
            system_instruction=SYSTEM_INSTRUCTION,
            temperature=0.2
        )

        with st.chat_message("assistant", avatar="aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️"):
            # Loop de tolerância a falhas para contornar instabilidade temporária (503)
            max_tentativas = 4
            resposta_sucesso = False

            for tentativa in range(1, max_tentativas + 1):
                try:
                    response_stream = client.models.generate_content_stream(
                        model="gemini-3.6-flash",
                        contents=contents,
                        config=config_rapida
                    )
                    resposta_completa = st.write_stream(
                        chunk.text for chunk in response_stream if chunk.text
                    )
                    st.session_state.messages.append({"role": "assistant", "content": resposta_completa})
                    resposta_sucesso = True
                    break
                except Exception as err:
                    erro_str = str(err)
                    if "503" in erro_str and tentativa < max_tentativas:
                        time.sleep(2.5 * tentativa)
                        continue
                    else:
                        st.error(f"Erro na resposta: {err}")
                        break
