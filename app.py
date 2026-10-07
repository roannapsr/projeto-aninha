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
- Analise criticamente os fatos, aponte a correlação biomecânica com a profissão habitual, discuta a existência ou não de incapacidade laborativa (temporária, total/parcial), sugere marcos de DID e DII e discuta a fundamentação conforme a Lei 8.213/91.
- Responda de forma clara, direta e técnica no chat. NÃO gere a minuta formal de 11 tópicos a menos que seja explicitamente solicitado o laudo completo.

QUANDO FOR SOLICITADO O LAUDO PERICIAL OFICIAL:
- NUNCA inicie sua resposta com saudações, introduções ou conversas prévias.
- SUA RESPOSTA DEVE COMEÇAR DIRETAMENTE NA PRIMEIRA LINHA DO CABEÇALHO OFICIAL: "PODER JUDICIÁRIO DA UNIÃO".
- NÃO UTILIZE MARCADORES DE ASTERISCOS (**) OU HASHTAGS (###). Entregue texto formal limpo.
- Conclua obrigatoriamente TODOS os 19 quesitos do Juízo e encerre com a data e assinatura da Dra. Ana Paula da Costa Henriques.

ESTRUTURA OBRIGATÓRIA DO LAUDO OFICIAL:

PODER JUDICIÁRIO DA UNIÃO
TRIBUNAL REGIONAL FEDERAL DA 5ª REGIÃO
JUSTIÇA FEDERAL DE PRIMEIRA INSTÂNCIA
SEÇÃO JUDICIÁRIA DE PERNAMBUCO - 19ª VARA

LAUDO DE EXAME MÉDICO-PERICIAL

1. PREÂMBULO
Processo nº:
Ação:
Órgão Julgador:
Autor:
Réu:
Data da Perícia:

2. PERÍCIA MÉDICA
Eu, Ana Paula da Costa Henriques, médico (CRM-PE 11.395), perito judicial nos autos do Processo abaixo discriminado, tendo realizado os levantamentos e pesquisas julgadas necessárias, venho apresentar o meu LAUDO PERICIAL.

3. DADOS DO(A) PERICIADO(A)
Nome: 
Identidade civil: 
CPF: 
Data do nascimento: 
Idade: 
Sexo biológico: 
Identidade de gênero: 
Escolaridade: 
Estado civil: 

4. HISTÓRICO LABORAL DO(A) PERICIADO(A)
Ocupação habitual: 
Descrição da atividade: 
Tempo de exercício da ocupação habitual: 
Data declarada de afastamento do trabalho: 
Experiência laboral anterior: 
Formação técnico-profissional: 
Reabilitação profissional: 

5. HISTÓRICO
Da análise da petição inicial e dos documentos apresentados, depreende-se que o(a) periciado(a) estaria acometido pela(s) seguinte(s) patologia(s) - CID: 

6. HISTÓRICO DA DOENÇA ATUAL

7. EXAME CLÍNICO
Sinais Vitais: PA: [X] mmHg. FC: [X] bpm.
Exame Geral: 
Aparelho Cardiorrespiratório: 
Exame Dermatológico: 
Aparelho Locomotor: 

8. DOCUMENTOS AVALIADOS
(Relação cronológica e análise crítica dos atestados, laudos e exames)

9. CONCLUSÃO PERICIAL
Data do Início da Doença (DID): 
Data do Início da Incapacidade (DII): 
O conjunto de patologias referenciado nos documentos médicos e administrativos juntados aos autos foi integralmente considerado no âmbito desta perícia, conduzida segundo protocolo técnico-científico próprio da medicina pericial. O procedimento adotado observou as seguintes etapas: identificação da demanda judicial e de seu objeto; levantamento e qualificação do periciado; coleta dos dados específicos da perícia; mapeamento das enfermidades alegadas e de sua evolução clínica; realização de anamnese e exame físico e/ou mental; apreciação crítica de atestados, laudos, exames complementares e documentos administrativos; cotejo com perícias anteriores eventualmente existentes; síntese integrativa de todos os elementos colhidos; e elaboração das respostas técnicas aos quesitos formulados pelo Juízo e pelas partes.

Patologia / Condição Clínica | CID-10 | Enquadramento Pericial e Fundamentação Técnica

Conclusões diagnósticas: 
Esclarecimentos: 
Incapacidade: 

10. QUESITOS DO JUÍZO
QUADRO I – QUESITAÇÃO PADRÃO
1. O Autor foi devidamente identificado por meio de documento original com foto e submetido a exame clínico completo?
2. O periciando é ou já foi paciente do ilustre perito?
3. Qual a profissão declarada pelo periciando? Caso esteja desempregado, qual a última atividade exercida pelo periciando?
4. Quais profissões o periciando declara já ter desempenhado?
5. Os dados objetivos do exame físico estão em correspondência com as queixas apresentadas?
6. O periciando é portador de alguma doença, sequela ou deficiência? Quais? Indicar exames em que se baseia.
7. A doença, deficiência física ou mental, anomalia ou lesão de que o periciando é portador incapacita para o exercício de atividade laborativa? Quais elementos levaram à convicção pericial? Tal incapacidade é temporária ou definitiva?
8. É possível afirmar se a incapacidade do periciando foi intermitente? Com base em que se afirma isso?
9. Caso exista apenas incapacidade temporária, a doença ou sequela incapacita a parte autora para seu trabalho ou atividade habitual por mais de 15 dias?
10. Caso a incapacidade seja temporária, qual o prazo ideal para tratamento, ainda que por estimativa, durante o qual o periciando não poderia trabalhar na sua atividade habitual?
11. O tempo estimado de recuperação é de dois anos, contados a partir da data de início da incapacidade/impedimento?
12. O quadro do periciado é progressivo, regressivo ou estável? Justifique.
13. Tal incapacidade inviabiliza o exercício de toda atividade laborativa (incapacidade total) ou apenas de algumas atividades laborativas (incapacidade parcial)?
14. Esclareça o perito, caso a incapacidade seja parcial, se o periciando pode exercer a atividade que habitualmente executa/executou, indicando, caso negativo, as atividades que poderá desempenhar, levando em conta o grau de escolaridade, idade e as condições sócio-econômicas.
15. A incapacidade decorreu de progressão ou agravamento de doença ou lesão da qual o periciado já era portador? Justifique.
16. Há incapacidade para o desempenho das atividades da vida independente? Ou seja, o periciando é capaz para realizar as atividades da vida diária (banhar-se, vestir-se, pentear-se, comer, passear etc.) independentemente da ajuda de terceiros?
17. O periciando é incapaz para o desempenho dos atos da vida civil? Parcial ou totalmente?
18. O periciando está acometido de tuberculose ativa, hanseníase, alienação mental, neoplasia maligna, cegueira, paralisia irreversível incapacitante, cardiopatia grave, doença de Parkinson, espondiloartrose anquilosante, nefropatia grave, estado avançado de doença Paget, AIDS, contaminação por radiação e/ou hepatopatia grave?
19. É necessário que o periciando faça uso constante de medicação? Em caso affirmativo, a medicação é fornecida pelo Sistema Único de Saúde – SUS?
QUADRO II – QUESITOS ESPECÍFICOS PARA AUXÍLIO-DOENÇA E/OU APOSENTADORIA POR INVALIDEZ – Não se aplicam
QUADRO III – QUESITOS ESPECÍFICOS PARA PERICIANDO MENORES DE 16 ANOS – Não se aplicam
QUESITIOS IV – QUESITOS ESPECÍFICOS PARA PORTADORES DA SÍNDROME DA IMUNODEFICIÊNCIA ADQUIRIDA – Não se aplicam
QUADRO V – CONSIDERAÇÕES (AVALIAÇÃO FUNCIONAL POR DOMÍNIOS)
Preste o senhor Perito os esclarecimentos adicionais que considerar necessários. Os esclarecimentos devem ser elaborados de forma clara e com linguagem acessível aos leigos (juiz, advogados e partes).
Ver conclusão pericial

11. ANEXOS (Fotos e Laudo da perícia trazidas pelo periciando(a))

Recife, [Data da Perícia].

ANA PAULA DA COSTA HENRIQUES
Médica Perita - CRM-PE 11.395
Assinatura Eletrônica
"""

# ----------------------------------------------------
# 3. Função para Gerar Arquivo .DOCX Limpo e Oficial
# ----------------------------------------------------
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

    texto_limpo = texto_laudo.replace("**", "").replace("###", "").replace("##", "")

    linhas = texto_limpo.split("\n")
    inicio_real = 0
    for idx, l in enumerate(linhas):
        l_upper = l.strip().upper()
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

        if any(h in linha.upper() for h in [
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
            run = p.add_run(linha)
            run.bold = True
            run.font.size = Pt(11)
            i += 1
            continue

        if any(ass in linha for ass in ["ANA PAULA DA COSTA HENRIQUES", "Médica Perita - CRM-PE 11.395", "Assinatura Eletrônica"]):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(linha)
            run.bold = True
            run.font.size = Pt(10)
            i += 1
            continue

        if "|" in linha:
            linhas_tabela = []
            while i < len(linhas_oficiais) and "|" in linhas_oficiais[i]:
                l_tab = linhas_oficiais[i].strip()
                if not re.match(r"^\|?[\s\-:|]+\|?$", l_tab):
                    colunas = [c.strip() for c in l_tab.strip("|").split("|")]
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

        eh_secao = any(linha.startswith(f"{n}.") for n in range(1, 13)) or linha.startswith("QUADRO")

        if eh_secao:
            p.paragraph_format.space_before = Pt(8)
            run = p.add_run(linha)
            run.bold = True
            run.font.size = Pt(10.5)
        else:
            if ":" in linha and len(linha.split(":", 1)[0]) < 45:
                rotulo, valor = linha.split(":", 1)
                run_r = p.add_run(rotulo + ":")
                run_r.bold = True
                run_r.font.size = Pt(10)
                run_v = p.add_run(valor)
                run_v.font.size = Pt(10)
            else:
                run = p.add_run(linha)
                run.font.size = Pt(10)

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
# 7. Processamento e Execução
# ----------------------------------------------------
prompt_usuario = st.chat_input("Digite dados do periciando, exame clínico ou orientações...")

prompt_acionado = None
eh_pedido_laudo = False

if prompt_usuario:
    # No chat: SEMPRE é análise pericial e discussão clínica (não gera laudo direto)
    prompt_acionado = prompt_usuario
    eh_pedido_laudo = False
elif btn_laudo:
    # Apenas o botão dispara a geração do laudo em arquivo
    eh_pedido_laudo = True
    prompt_acionado = (
        "Elabore a minuta completa do LAUDO DE EXAME MÉDICO-PERICIAL oficial da 19ª Vara / TRF5, "
        "com base estritamente nos documentos anexados e nos fatos informados na discussão. "
        "Comece DIRETAMENTE pelo cabeçalho institucional (PODER JUDICIÁRIO DA UNIÃO), sem mensagens prévias. "
        "Preencha todos os 11 itens oficiais sem marcadores de asteriscos (**), fundamentando DID e DII, "
        "o quadro comparativo de patologias, respondendo integralmente aos 19 quesitos do Juízo "
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
        "Com base nos documentos médicos e no histórico do caso, responda de forma técnica, "
        "precisa e conclusiva aos 19 quesitos padrão do Quadro I do Juízo da 19ª Vara."
    )

if prompt_acionado:
    if not api_key:
        st.error("Chave GEMINI_API_KEY não configurada no ambiente (.env ou Secrets)!")
    else:
        texto_exibicao_usuario = "🚀 **Solicitação:** Elaborar e formatar o Laudo Médico-Pericial oficial completo." if eh_pedido_laudo else prompt_acionado
        st.session_state.messages.append({"role": "user", "content": texto_exibicao_usuario})
        with st.chat_message("user"):
            st.markdown(texto_exibicao_usuario)

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
                f"{prompt_acionado}\n\n[INSTRUÇÃO CRUCIAL: Elabore o laudo COMPLETO de ponta a ponta sem abreviações. "
                "Comece em 'PODER JUDICIÁRIO DA UNIÃO' e responda a todos os quesitos até o encerramento com a assinatura.]"
            )
        else:
            demanda_final = prompt_acionado

        historico_texto += f"\nNOVA DEMANDA:\n{demanda_final}"
        contents.append(historico_texto)

        config_ia = types.GenerateContentConfig(
            system_instruction=SYSTEM_INSTRUCTION,
            max_output_tokens=8192
        )

        with st.chat_message("assistant", avatar="aninha.jpeg" if os.path.exists("aninha.jpeg") else "👩‍⚕️"):
            status_box = st.empty()
            try:
                # 1. EXCLUSIVO DO BOTÃO GERAR LAUDO: Gera em segundo plano e só entrega o download
                if eh_pedido_laudo:
                    status_box.info("⏳ Só um momento, Dra Aninha está analisando as informações...")
                    response = client.models.generate_content(
                        model="gemini-3.6-flash",
                        contents=contents,
                        config=config_ia
                    )
                    texto_laudo = response.text
                    status_box.empty()

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

                # 2. CHAT / ANÁLISE / DEMAIS AÇÕES: Análise técnica em texto no chat com streaming
                else:
                    status_box.info("⏳ Só um momento, Dra Aninha está analisando as informações...")
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
