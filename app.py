import os
import time
import docx
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

# Caminho do avatar da Aninha
AVATAR_ANINHA = "aninha.jpeg" if os.path.exists("aninha.jpeg") else "🩺"

st.set_page_config(
    page_title="Dra. Aninha - Perícia Previdenciária",
    page_icon=AVATAR_ANINHA,
    layout="wide",
    initial_sidebar_state="auto"
)
st.markdown("""
<style>
    /* No computador (telas maiores que 768px): esconde a setinha e trava a barra */
    @media (min-width: 769px) {
        [data-testid="stSidebarCollapseButton"] {
            display: none !important;
        }
        [data-testid="stSidebarHeader"] {
            display: none !important;
        }
    }

    /* No celular: permite recolher e dá espaço para o botão de fechar */
    @media (max-width: 768px) {
        [data-testid="stSidebarHeader"] {
            display: flex !important;
        }
    }

    [data-testid="stSidebarUserContent"] {
        padding-top: 0.8rem !important;
    }
</style>
""", unsafe_allow_html=True)
# Cliente Gemini
@st.cache_resource
def get_client():
    return genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

client = get_client()

def extrair_texto_docx(arquivo):
    doc = docx.Document(arquivo)
    texto_completo = []
    for paragrafo in doc.paragraphs:
        if paragrafo.text.strip():
            texto_completo.append(paragrafo.text)
    return "\n".join(texto_completo)

# Prompt especializado para perícia médica previdenciária
SYSTEM_INSTRUCTION = """
Você é a 'Dra. Aninha', assistente pericial de inteligência artificial de altíssimo nível, especializada em Perícia Médica Previdenciária e Judicial (Justiça Federal - Juizados Especiais Federais e Varas Previdenciárias).
Seu propósito é prestar suporte técnico direto à médica perita na análise de documentos, formulação de conclusões periciais e resposta aos quesitos judiciais e das partes.

DIRETRIZES TÉCNICAS MANDATÓRIAS:
1. DISTINÇÃO FUNDAMENTAL: Doença/Lesão (CID) NÃO equivale a incapacidade laborativa. A incapacidade só existe quando há perda ou redução substancial do desempenho das funções exigidas pela atividade profissional habitual declarada.
2. CLASSIFICAÇÃO DA INCAPACIDADE:
   - Temporária (com prognóstico de recuperação clínica/cirúrgica) vs. Permanente (consolidada/irreversível).
   - Total (inviabiliza qualquer ato da profissão) vs. Parcial (permite tarefas sem sobrecarga ao segmento acometido).
   - Uniprofissional vs. Multiprofissional.
3. MARCOS TEMPORAIS CRÍTICOS:
   - DID (Data de Início da Doença): marco clínico inicial dos sintomas ou diagnóstico comprovado em documento idôneo.
   - DII (Data de Início da Incapacidade): momento em que a patologia tornou-se incapacitante para a função laboral, embasada em exames de imagem, prontuários de internação, relatórios circunstanciados ou concessão administrativa anterior.
4. NEXO TÉCNICO E CONCAUSALIDADE: Identifique se há nexo causal direto, concausalidade (agravamento pelo trabalho habitual) ou patologia eminentemente degenerativa/constitucional sem nexo.
5. RESPOSTAS A QUESITOS:
   - Responda com linguagem forense culta, técnica, concisa e peremptória.
   - Não emita juízos de valor de mérito jurídico privativos do Magistrado.
6. ANÁLISE DE DOCUMENTOS ANEXADOS:
   - Destaque a data, o método diagnóstico, a assinatura médica com CRM e se o exame evidencia limitação funcional ou apenas alteração anatômica degenerativa compatível com a faixa etária.
"""

# Mensagem inicial padrão
MENSAGEM_INICIAL = {
    "role": "assistant",
    "content": "Olá, Doutora! Sou a **Aninha**, sua assistente técnica em Perícia Previdenciária.\n\nVocê pode me ditar o histórico clínico, informar a profissão do periciando ou **anexar relatórios e exames (PDF ou imagem)** na barra lateral para analisarmos juntos a capacidade laborativa, fixação de DID/DII e respostas aos quesitos!"
}

# Inicialização do histórico na sessão
if "messages" not in st.session_state:
    st.session_state.messages = [MENSAGEM_INICIAL]

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# Barra Lateral (Avatar centralizado, Upload e Ações Rápidas)
with st.sidebar:
    if os.path.exists("aninha.jpeg"):
        col_esq, col_centro, col_dir = st.columns([1, 2, 1])
        with col_centro:
            st.image("aninha.jpeg", use_container_width=True)
    else:
        st.markdown("<h1 style='text-align: center;'>🩺</h1>", unsafe_allow_html=True)
    
    st.markdown("### 📁 Anexar Documentos Médicos")
    st.caption("Envie laudos, atestados, exames de imagem ou petições.")
    
    uploaded_files = st.file_uploader(
        "Selecione arquivos (PDF, PNG, JPG, JPEG):",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.uploader_key}"
    )
    
    # Confirmação visual de upload na própria barra lateral
    if uploaded_files:
        qtd = len(uploaded_files)
        msg_doc = f"✅ {qtd} documento{'s' if qtd > 1 else ''} anexado{'s' if qtd > 1 else ''} com sucesso!"
        st.success(msg_doc)
    
    st.divider()
    st.markdown("### ⚡ Ações Rápidas de Perícia")
    
    if st.button("🔄 Iniciar Nova Análise", use_container_width=True, type="secondary"):
        st.session_state.messages = [
            {
                "role": "assistant",
                "content": "Novo caso pericial iniciado! Pode inserir os dados do periciando ou anexar novos documentos."
            }
        ]
        st.session_state.uploader_key += 1  # Limpa todos os arquivos anexados
        st.rerun()

    btn_dii = st.button("⏱️ Fixar DII / DID", use_container_width=True)
    btn_quesitos = st.button("❓ Responder Quesitos", use_container_width=True)
    btn_laudo = st.button("📄 Gerar Laudo Pericial", use_container_width=True, type="primary")

# Cabeçalho da página principal com miniatura da Aninha
col_avatar, col_header = st.columns([0.8, 9], vertical_alignment="center")

with col_avatar:
    if os.path.exists("aninha.jpeg"):
        st.image("aninha.jpeg", width=65)
    else:
        st.markdown("### 🩺")

with col_header:
    st.title("Dra. Aninha — Assistente de Perícia Previdenciária")
    st.caption("Análise pericial técnica judicial, incapacidade laborativa, DII/DID e quesitos oficiais.")

# Exibe histórico do chat
for msg in st.session_state.messages:
    avatar = "👤" if msg["role"] == "user" else AVATAR_ANINHA
    with st.chat_message(msg["role"], avatar=avatar):
        st.markdown(msg["content"])

# Processa cliques nos botões rápidos
prompt_acao = None
if btn_dii:
    prompt_acao = "Com base em todo o histórico e documentos que analisamos até aqui, estruture uma proposta técnica fundamentada para a fixação de DID (Data de Início da Doença) e DII (Data de Início da Incapacidade), justificando em quais documentos médicos e achados nos apoiamos."
elif btn_quesitos:
    prompt_acao = "Elabore as respostas técnicas fundamentadas para os quesitos judiciais unificados da TNU/Justiça Federal sobre capacidade/incapacidade laborativa para este caso."
elif btn_laudo:
    prompt_acao = "Com base em todo o caso pericial em discussão, elabore a minuta estruturada do Laudo Médico Pericial Judicial (incluindo identificação, histórico clínico-ocupacional, análise dos exames, discussão técnica sobre a capacidade laborativa, fixação de DII/DID e conclusão pericial)."

user_input = st.chat_input("Digite detalhes do periciando, exames, perguntas ou orientações...")
prompt_final = prompt_acao if prompt_acao else user_input

if prompt_final:
    st.session_state.messages.append({"role": "user", "content": prompt_final})
    with st.chat_message("user", avatar="👤"):
        st.markdown(prompt_final)

    with st.chat_message("assistant", avatar=AVATAR_ANINHA):
        with st.spinner("Dra. Aninha está analisando os elementos periciais..."):
            contents = []
            
            # Leitura de arquivos anexados
            if uploaded_files:
                for f in uploaded_files:
                    bytes_data = f.read()
                    mime = f.type

                    if mime == "application/pdf":
                        part = types.Part.from_bytes(data=bytes_data, mime_type="application/pdf")
                        contents.append(part)
                    elif mime in ["image/png", "image/jpeg", "image/jpg"]:
                        part = types.Part.from_bytes(data=bytes_data, mime_type=mime)
                        contents.append(part)
                    elif f.name.endswith(".docx") or mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                        texto_docx = extrair_texto_docx(f)
                        if texto_docx.strip():
                            contents.append(f"\n[Conteúdo do arquivo anexado: {f.name}]\n{texto_docx}\n")

            # 2. Histórico de contexto da conversa (alinhado com contents = [])
            conversa_contexto = "HISTÓRICO PERICIAL DA SESSÃO ATÉ O MOMENTO:\n"
            for m in st.session_state.messages[:-1]:
                autor = "Médica Perita" if m["role"] == "user" else "Dra. Aninha"
                conversa_contexto += f"{autor}: {m['content']}\n"

            conversa_contexto += f"\nNOVA DEMANDA DA MÉDICA:\n{prompt_final}"
            contents.append(conversa_contexto)

            # 3. Chamada do modelo Gemini (alinhado com contents = [])
            config_rapida = types.GenerateContentConfig(
                system_instruction=SYSTEM_INSTRUCTION,
                temperature=0.2
            )

           # Modelos alternativos caso haja sobrecarga momentânea (503)
            modelos_disponiveis = ["gemini-3.6-flash", "gemini-2.0-flash"]

            sucesso = False
            for modelo_nome in modelos_disponiveis:
                for tentativa in range(2):
                    try:
                        response = client.models.generate_content(
                            model=modelo_nome,
                            contents=contents,
                            config=config_rapida
                        )
                        resposta_texto = response.text
                        st.markdown(resposta_texto)
                        st.session_state.messages.append({"role": "assistant", "content": resposta_texto})
                        sucesso = True
                        break
                    except Exception as err:
                        if "503" in str(err) or "NOT_FOUND" in str(err):
                            time.sleep(2)
                            continue
                        break
                if sucesso:
                    break

            if not sucesso:
                st.error("Servidores do Google temporariamente indisponíveis. Por favor, tente novamente em instantes.")
