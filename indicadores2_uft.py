# ============================================================
# INDICADORES ACADÊMICOS · UFT · VERSÃO 7
# Universidade Federal do Tocantins
#
# Calcula por CAMPUS · CURSO · ANO:
#   1. Índice de Evasão
#   2. Índice de Retenção
#   3. Taxa de Sucesso
#
# Layout e aparência: iguais às versões 2/3.
# Mudanças da v7 em relação à v3 (somente contagem):
#   A. DEDUPLICAÇÃO SEGURA — uma linha por ENTRADA:
#        MATRICULA × COD_CURSO × CAMPUS × ANO_INGRESSO × PERIODO_INGRESSO
#      • cópias exatas ..................................... removidas
#      • mesma entrada repetida com dados diferentes ....... fica 1 linha
#      • mesma matrícula em cursos diferentes .............. mantidas
#      • reingresso no mesmo curso em outro ano/semestre ... mantidas
#        (todas as entradas são contadas, cada uma na sua coorte)
#      • MATRICULA comparada exatamente: "…x" ≠ "…X"
#        (na base 16.11.03a são alunos diferentes)
#   B. DECLINANTES FORA DO CÁLCULO — não entram no numerador nem no
#      denominador (total) de nenhum índice.
#   C. CORREÇÕES DE CÁLCULO encontradas na v3:
#      1. Retenção: a v3 usava prazo fixo de 5 anos para TODOS os cursos
#         (DURAÇÃO vem vazia para 100% dos ativos). Agora o prazo é por
#         curso (PRAZOS_OFICIAIS ou estimado dos formados) e o tempo é
#         contado em semestres até o último semestre registrado na base.
#      2. Falecimento, Habilitação, encerramentos de mobilidade/
#         intercâmbio/convênio e "Em análise" NÃO são evasão: saem do
#         cálculo, como os declinantes.
#   D. TRANSIÇÃO UFNT (Araguaína e Tocantinópolis, 2023): os alunos que
#      migraram formam a categoria MIGRADO. Eles CONTINUAM no total (eram
#      alunos da UFT até migrar), mas não são evasão nem sucesso.
#      Tirá-los do total criava 100% de evasão (só sobravam os que já
#      tinham saído antes da migração — viés de sobrevivência).
#   E. Aba "Entenda os Cálculos": fórmulas e metodologia para leigos.
#   F. Logo da UFT (arquivo logo_uft.png na mesma pasta) no cabeçalho.
#      3. Mapa de calor: a v3 fazia média simples das taxas anuais e
#         mostrava 0 onde não havia curso. Agora: taxa agregada por
#         campus × curso, célula em branco quando o curso não existe.
#      4. Rankings: mínimo de 20 alunos (igual ao gráfico de dispersão),
#         para um curso com 1 aluno não aparecer com 100%.
#      5. DURAÇÃO-ANO com vírgula ("4,5") agora é lida corretamente.
# ============================================================

# ============================================================
# CORREÇÃO EVENT LOOP WINDOWS
# ============================================================

import asyncio
import sys

if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# ============================================================
# IMPORTS
# ============================================================

import xml.sax
import io
import base64
from html import escape
import os
import re
import unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

# Logo da UFT: arquivo "logo_uft.png" na mesma pasta do script.
# Se o arquivo não existir, o painel usa o emoji 📊 (nada quebra).
DIRETORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
CAMINHO_LOGO = os.path.join(DIRETORIO_SCRIPT, "logo_uft.png")
LOGO_B64 = None
if os.path.isfile(CAMINHO_LOGO):
    with open(CAMINHO_LOGO, "rb") as arquivo_logo:
        LOGO_B64 = base64.b64encode(arquivo_logo.read()).decode()

st.set_page_config(
    page_title="Indicadores Acadêmicos · UFT · v7",
    page_icon=CAMINHO_LOGO if LOGO_B64 else "📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# PALETA INSTITUCIONAL
# ============================================================

COR_PRIMARIA   = "#0D2B4E"
COR_SECUNDARIA = "#1A4A7A"
COR_ACENTO     = "#1E6FBF"
COR_DESTAQUE   = "#C8A84B"
COR_FUNDO      = "#F4F6F9"
COR_CARD       = "#FFFFFF"
COR_BORDA      = "#DDE3EC"
COR_TEXTO      = "#0D2B4E"
COR_TEXTO_LEVE = "#6B7A99"
COR_ALERTA     = "#E07B00"
COR_SUCESSO    = "#1B7A4A"
COR_EVASAO     = "#C0392B"
COR_RETENCAO   = "#E07B00"

SEQUENCIA_AZUL = [
    "#C8D8EE", "#93B5D8", "#5E92C2",
    "#2F6FAC", "#0D4C8B", "#0A3366"
]

# Nomes que não podem ser recuperados apenas da coluna CURSO do ODS.
# No relatório 16.11.03a, estas habilitações aparecem genericamente como
# "Curso de Letras"; o código é a chave que distingue cada curso.
NOMES_CURSO_POR_CODIGO = {
    "44M600NC": "Curso de Letras - Matutino (Núcleo Comum)",
    "44M601P": "Curso de Letras - Língua Portuguesa e Literaturas - Matutino",
    "44M602I": "Curso de Letras - Língua Inglesa e Literaturas - Matutino",
    "44N600NC": "Curso de Letras - Noturno (Núcleo Comum)",
    "44N601P": "Curso de Letras - Língua Portuguesa e Literaturas - Noturno",
    "44N602I": "Curso de Letras - Língua Inglesa e Literaturas - Noturno",
    "51M600L": "Curso de Letras - Libras - Licenciatura",
}


def nome_curso_sem_campus(codigo, curso, campus):
    """Retorna o nome completo do curso, sem repetir a cidade do campus."""
    codigo = "" if pd.isna(codigo) else str(codigo).strip()
    curso = "" if pd.isna(curso) else str(curso).strip()
    campus = "" if pd.isna(campus) else str(campus).strip()

    if codigo in NOMES_CURSO_POR_CODIGO:
        return NOMES_CURSO_POR_CODIGO[codigo]

    if campus:
        # Remove somente a cidade usada como sufixo, inclusive quando ela vem
        # antes de um complemento entre parênteses, sem apagar outros termos.
        padrao = rf"\s*-\s*{re.escape(campus)}(?=\s*\(|\s*$)"
        curso = re.sub(padrao, "", curso, flags=re.IGNORECASE)

    return re.sub(r"\s{2,}", " ", curso).strip(" -")


def separar_rotulo_curso(rotulo):
    """Separa o identificador interno no formato 'código — nome'."""
    partes = str(rotulo).split(" — ", maxsplit=1)
    return (partes[0], partes[1]) if len(partes) == 2 else ("", partes[0])


def formatar_opcao_curso(rotulo):
    """Exibe primeiro o nome, deixando o código como informação complementar."""
    codigo, nome = separar_rotulo_curso(rotulo)
    return f"{nome} ({codigo})" if codigo else nome


def chave_ordenacao_curso(rotulo):
    """Ordena alfabeticamente pelo nome, ignorando acentos e pontuação."""
    codigo, nome = separar_rotulo_curso(rotulo)
    nome_sem_acentos = "".join(
        caractere
        for caractere in unicodedata.normalize("NFKD", nome)
        if not unicodedata.combining(caractere)
    ).casefold()
    nome_normalizado = re.sub(r"[^a-z0-9]+", " ", nome_sem_acentos).strip()
    return nome_normalizado, codigo.casefold()

# ============================================================
# CSS GLOBAL
# ============================================================

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lora:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body {{
    font-family: 'DM Sans', sans-serif;
    color: {COR_TEXTO};
}}
.main        {{ background-color: {COR_FUNDO}; padding-top: 0.5rem; }}
[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, {COR_PRIMARIA} 0%, {COR_SECUNDARIA} 100%);
}}
[data-testid="stSidebar"] * {{ color: white !important; }}
[data-testid="stSidebar"] label {{
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}}
[data-testid="stSidebar"] [data-baseweb="select"] > div {{
    background-color: white !important;
    border-color: rgba(255,255,255,0.55) !important;
}}
[data-testid="stSidebar"] [data-baseweb="select"] *,
[data-testid="stSidebar"] [data-baseweb="select"] svg {{
    color: {COR_TEXTO} !important;
    fill: {COR_TEXTO} !important;
}}
[data-testid="stSidebar"] [data-baseweb="tag"] * {{
    color: white !important;
}}
[data-testid="stSidebar"] [data-testid="stPills"] button {{
    background-color: rgba(255,255,255,0.12) !important;
    border: 1px solid rgba(255,255,255,0.45) !important;
}}
[data-testid="stSidebar"] [data-testid="stPills"] button * {{
    color: white !important;
}}
[data-testid="stSidebar"] [data-testid="stPills"] button[aria-pressed="true"] {{
    background-color: white !important;
}}
[data-testid="stSidebar"] [data-testid="stPills"] button[aria-pressed="true"] * {{
    color: {COR_TEXTO} !important;
}}
.chart-card {{
    background: {COR_CARD};
    border: 1px solid {COR_BORDA};
    border-radius: 14px;
    padding: 18px 18px 8px;
    box-shadow: 0 2px 10px rgba(13,43,78,0.05);
    margin-bottom: 12px;
}}
.kpi-wrapper {{
    display: flex;
    gap: 10px;
    margin-bottom: 20px;
    flex-wrap: wrap;
}}
hr {{ border: none; border-top: 1px solid {COR_BORDA}; margin: 20px 0; }}

/* v7 · abas "Painel" e "Entenda os Cálculos" na paleta institucional */
[data-testid="stTabs"] [role="tablist"] {{
    gap: 10px; border-bottom: 3px solid {COR_DESTAQUE}; margin-bottom: 20px;
}}
[data-testid="stTab"] {{
    padding: 12px 26px !important; border-radius: 12px 12px 0 0;
    background: white; border: 1px solid {COR_BORDA}; border-bottom: none;
}}
[data-testid="stTab"] p {{ font-size: 1rem !important; font-weight: 700; color: {COR_SECUNDARIA} !important; }}
[data-testid="stTab"]:hover {{ background: {COR_FUNDO}; }}
[data-testid="stTab"][aria-selected="true"] {{
    background: linear-gradient(135deg, {COR_PRIMARIA} 0%, {COR_SECUNDARIA} 100%);
    border-color: {COR_PRIMARIA};
}}
[data-testid="stTab"][aria-selected="true"] p {{ color: white !important; }}
[data-testid="stTab"] .react-aria-SelectionIndicator {{ display: none; }}
.explica {{
    background: {COR_CARD}; border: 1px solid {COR_BORDA}; border-radius: 14px;
    padding: 16px 20px; margin-bottom: 14px; line-height: 1.6; font-size: 0.93rem;
}}
.explica b {{ color: {COR_PRIMARIA}; }}
</style>
""",
    unsafe_allow_html=True
)

# ============================================================
# HELPERS DE INTERFACE
# ============================================================

def chart_container(fig):
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)


def secao(titulo):
    st.markdown(
        f"""
<div style="
    font-family: Georgia, serif;
    font-size: 1rem;
    font-weight: 600;
    color: {COR_PRIMARIA};
    padding-bottom: 6px;
    border-bottom: 2px solid {COR_DESTAQUE};
    margin-bottom: 14px;
">
{titulo}
</div>
""",
        unsafe_allow_html=True
    )


def kpi_card(icone, valor, label, cor, subtitulo="", ajuda=""):
    """
    Cartão de indicador (mesmo visual da versão original).
    `ajuda`: texto explicativo exibido ao passar o mouse (marcado com ⓘ).
    Linhas do texto de ajuda podem ser separadas por "\n".
    """
    sub = (
        f'<div style="font-size:0.65rem;color:{COR_TEXTO_LEVE};margin-top:2px;">{subtitulo}</div>'
        if subtitulo else ""
    )
    dica = ""
    marca = ""
    if ajuda:
        dica = ' title="' + "&#10;".join(escape(l, quote=True) for l in ajuda.split("\n")) + '"'
        marca = ' <span style="opacity:0.55;">ⓘ</span>'
    return (
        f'<div{dica} style="background:white;border:1px solid {COR_BORDA};border-radius:14px;'
        f'padding:16px 14px;flex:1;min-width:125px;position:relative;{"cursor:help;" if ajuda else ""}">'
        f'<div style="position:absolute;top:0;left:0;right:0;height:4px;background:{cor};'
        f'border-radius:14px 14px 0 0;"></div>'
        f'<div style="font-size:1.4rem;">{icone}</div>'
        f'<div style="font-family:\'Lora\',serif;font-size:1.9rem;font-weight:700;color:{COR_PRIMARIA};">{valor}</div>'
        f'<div style="font-size:0.68rem;font-weight:700;letter-spacing:0.07em;text-transform:uppercase;'
        f'color:{COR_TEXTO_LEVE};">{label}{marca}</div>'
        f'{sub}</div>'
    )


def selo_logo(altura_px: int) -> str:
    """Logo da UFT sobre um selo branco arredondado (fica nítida no fundo azul)."""
    if not LOGO_B64:
        return f'<div style="font-size:{altura_px * 0.6:.0f}px;">📊</div>'
    return (
        f'<div style="background:white;border-radius:{altura_px * 0.18:.0f}px;padding:{altura_px * 0.1:.0f}px;'
        f'display:flex;align-items:center;justify-content:center;flex-shrink:0;'
        f'box-shadow:0 4px 14px rgba(0,0,0,0.18);">'
        f'<img src="data:image/png;base64,{LOGO_B64}" alt="Logo UFT" '
        f'style="height:{altura_px}px;width:auto;display:block;"></div>'
    )


LAYOUT_BASE = dict(
    paper_bgcolor="white",
    plot_bgcolor="white",
    font=dict(family="DM Sans", color=COR_TEXTO),
    margin=dict(t=30, b=20, l=10, r=10)
)

# ============================================================
# LEITURA DO ARQUIVO ODS COM PARSER SAX
# (Evita carregar >500 MB de XML diretamente em memória)
# ============================================================

class ODSHandler(xml.sax.ContentHandler):
    """
    Lê o content.xml de um arquivo ODS linha a linha via SAX,
    convertendo cada <table:table-row> em uma lista de strings
    e acumulando tudo em self.rows.
    """

    def __init__(self):
        self.in_row        = False
        self.in_cell       = False
        self.in_text       = False
        self.current_row   = []
        self.current_text  = ""
        self.repeat_count  = 1
        self.rows          = []          # acumula TODAS as linhas

    # ---- SAX callbacks ----

    def startElement(self, name, attrs):
        if name == "table:table-row":
            self.in_row      = True
            self.current_row = []

        elif name == "table:table-cell":
            self.in_cell      = True
            self.current_text = ""
            rep = attrs.get("table:number-columns-repeated", "1")
            try:
                rc = int(rep)
                self.repeat_count = rc if rc <= 50 else 1
            except ValueError:
                self.repeat_count = 1

        elif name == "text:p":
            self.in_text = True

    def characters(self, content):
        if self.in_text:
            self.current_text += content

    def endElement(self, name):
        if name == "text:p":
            self.in_text = False

        elif name == "table:table-cell":
            self.in_cell = False
            for _ in range(self.repeat_count):
                self.current_row.append(self.current_text)
            self.current_text = ""
            self.repeat_count = 1

        elif name == "table:table-row":
            self.in_row = False
            row = self.current_row[:]
            # Remove células vazias à direita
            while row and row[-1] == "":
                row.pop()
            if row:
                self.rows.append(row)


@st.cache_data(show_spinner="🔄 Extraindo dados do arquivo ODS…")
def carregar_ods(caminho: str, versao_arquivo=None) -> pd.DataFrame:
    """
    Abre o arquivo ODS, faz o parse SAX do content.xml interno
    e retorna um DataFrame já limpo.

    Parâmetros
    ----------
    caminho : str
        Caminho para o arquivo .ods (ex.: 'dados_uft.ods')
    versao_arquivo : tuple, opcional
        Assinatura (data de modificação e tamanho) usada para invalidar o cache
        quando o conteúdo do arquivo for substituído mantendo o mesmo nome.

    Retorna
    -------
    pd.DataFrame
        DataFrame com as colunas do arquivo e apenas
        os campus reconhecidos da UFT.
    """

    import zipfile

    # 1. Abre o zip e lê content.xml
    with zipfile.ZipFile(caminho, "r") as zf:
        with zf.open("content.xml") as xml_file:
            handler = ODSHandler()
            parser  = xml.sax.make_parser()
            parser.setContentHandler(handler)
            parser.parse(xml_file)

    if not handler.rows:
        raise ValueError("Nenhuma linha encontrada no arquivo ODS.")

    # 2. Monta DataFrame: primeira linha = cabeçalho
    header = handler.rows[0]
    data   = handler.rows[1:]

    # Garante que todas as linhas tenham o mesmo comprimento
    n = len(header)
    data = [
        (row + [""] * n)[:n]
        for row in data
        if len(row) >= max(5, n // 3)   # descarta linhas muito curtas
    ]

    df = pd.DataFrame(data, columns=header)

    # 3. Renomeia colunas duplicadas automaticamente
    cols     = list(df.columns)
    seen     = {}
    new_cols = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            new_cols.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            new_cols.append(c)
    df.columns = new_cols

    # 4. Valida as colunas que identificam univocamente cada curso
    colunas_obrigatorias = {"COD_CURSO", "CURSO", "CAMPUS"}
    colunas_ausentes = sorted(colunas_obrigatorias - set(df.columns))
    if colunas_ausentes:
        raise ValueError(
            "Colunas obrigatórias ausentes no ODS: " + ", ".join(colunas_ausentes)
        )

    # 5. Limpa espaços antes de validar os valores de campus e situação
    # v7: inclui o tipo texto do pandas 3 ("str"/"string"); com apenas "object"
    # a limpeza não rodaria e espaços extras esconderiam duplicatas.
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = df[col].astype(object).str.strip()

    # 6. Mantém apenas campus válidos da UFT
    CAMPUS_UFT = {
        "Araguaína", "Arraias", "Gurupi",
        "Miracema", "Palmas", "Porto Nacional", "Tocantinópolis"
    }
    if "CAMPUS" in df.columns:
        df = df[df["CAMPUS"].isin(CAMPUS_UFT)].copy()

    # 7. Conversão de tipos
    for col in ["ANO_INGRESSO", "ANO_EVASAO", "DURAÇÃO-SEM", "DURAÇÃO-ANO"]:
        if col in df.columns:
            # v7: aceita vírgula decimal ("4,5"); antes virava vazio.
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", ".", regex=False),
                errors="coerce",
            )

    # 8. Remove dados pessoais (LGPD)
    for col in ["NOME", "CPF", "CPF_NUM"]:
        if col in df.columns:
            df.drop(columns=col, inplace=True)

    # 9. Corrige nomes incompletos e elimina a cidade repetida no nome.
    # O código continua em sua própria coluna e passa a integrar todos os
    # agrupamentos, evitando fundir habilitações que têm o mesmo texto em CURSO.
    df["COD_CURSO"] = df["COD_CURSO"].astype("string").str.strip()
    df["CURSO"] = [
        nome_curso_sem_campus(codigo, curso, campus)
        for codigo, curso, campus in zip(
            df["COD_CURSO"], df["CURSO"], df["CAMPUS"]
        )
    ]

    return df.reset_index(drop=True)

# ============================================================
# CLASSIFICAÇÃO DE SITUAÇÕES
# ============================================================

# ──────────────────────────────────────────────────────────────
# DEFINIÇÕES CONCEITUAIS
# ──────────────────────────────────────────────────────────────
#
# EVASÃO: aluno saiu do curso SEM concluir.
#   Inclui: Desistência, Desvinculado, Matrícula Cancelada, Jubilado,
#           Transferência Interna/Externa/Ex-offício,
#           Reopção de Curso, Troca de Turno.
#
# MIGRADO PARA A UFNT: "Transição UFNT" — continua no total, mas não é
#   evasão nem sucesso (categoria própria).
#
# FORA DO CÁLCULO (não entram no total nem em nenhum índice):
#   Declinante, Falecimento, Habilitação,
#   encerramentos de intercâmbio/mobilidade/aluno especial/convênio/
#   apostilamento e "Em análise".
#
# RETENÇÃO: aluno ATIVO cujo tempo no curso, em semestres, supera
#   o prazo do curso + tolerância.
#   tempo = 2 × (ano_ref − ano_ingresso) + (sem_ref − sem_ingresso) + 1
#   ano_ref/sem_ref = último semestre registrado na base.
#   prazo = PRAZOS_OFICIAIS[código] ou, se ausente, estimado dos formados
#           do próprio curso (ver estimar_prazos).
#
# SUCESSO: aluno concluiu o curso com êxito.
#   Inclui: Formado.
#
# Todas as taxas são calculadas por COORTE de ingresso:
#   denominador = total ingressantes daquele (campus, curso, ano).
# ──────────────────────────────────────────────────────────────

SITUACOES_EVASAO = {
    "Desistência",
    "Desvinculado",
    "Matrícula Cancelada",
    "Jubilado",
    "Transferência Interna",
    "Transferência Externa",
    "Transferência Ex-offício",
    "Reopção de Curso",
    "Troca de Turno",
}

SITUACOES_SUCESSO = {"Formado"}

SITUACOES_ATIVO  = {"Vinculado", "Reingresso Administrativo"}

# v7: alunos de Araguaína e Tocantinópolis transferidos para a UFNT (2023).
# Categoria própria: ficam no total, mas não são evasão nem sucesso.
SITUACOES_MIGRACAO = {"Transição UFNT"}
CAMPI_UFNT = {"Araguaína", "Tocantinópolis"}
ANO_TRANSICAO_UFNT = 2023


def rotulo_campus(campus: str) -> str:
    """Nome do campus para os gráficos, sinalizando os que viraram UFNT."""
    return f"{campus} (UFNT {ANO_TRANSICAO_UFNT})" if campus in CAMPI_UFNT else campus

# v7: situações retiradas da base ANTES de qualquer cálculo — não entram
# no total (denominador) nem em nenhum numerador.
SITUACOES_FORA_DO_CALCULO = {
    "Declinante",                              # não efetivou o curso
    "Falecimento",                             # não é evasão (critério Inep)
    "Habilitação",                             # segue na habilitação (outra entrada)
    "Encerramento Intercâmbio Internacional",  # vínculos temporários,
    "Encerramento Mobilidade Acadêmica",       # não são alunos regulares
    "Encerramento Aluno Especial",             # do curso
    "Encerramento de Convênio",
    "Encerramento/Apostilamento",
    "Em análise",                              # situação indefinida
}

# Prazo de integralização (em SEMESTRES) por código de curso, conforme PPC.
# Preencha quando tiver o valor oficial: ele substitui a estimativa.
# Exemplo: "23I500G": 12,   # Medicina · Palmas
PRAZOS_OFICIAIS = {}

# Estimativa do prazo quando o curso não está em PRAZOS_OFICIAIS:
# menor duração (DURAÇÃO-SEM) que concentra pelo menos 10% dos formados
# que ingressaram por processo seletivo (PS/SISU). Formados que chegam por
# transferência têm créditos aproveitados e distorceriam o valor.
MIN_FORMADOS_PARA_ESTIMAR = 20
PRAZO_MIN_SEM, PRAZO_MAX_SEM = 6, 12        # limites plausíveis de graduação
INGRESSO_REGULAR = r"^(PS |SISU|Processo Seletivo)"


def classifica_situacao(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona coluna SITUACAO_GRUPO com valores:
        EVASAO | SUCESSO | ATIVO | MIGRADO | OUTRO
    """
    df = df.copy()

    cond_evasao  = df["FORMAEVASAO"].isin(SITUACOES_EVASAO)
    cond_sucesso = df["FORMAEVASAO"].isin(SITUACOES_SUCESSO)
    cond_ativo   = df["FORMAEVASAO"].isin(SITUACOES_ATIVO)
    cond_migrado = df["FORMAEVASAO"].isin(SITUACOES_MIGRACAO)

    df["SITUACAO_GRUPO"] = "OUTRO"
    df.loc[cond_ativo,   "SITUACAO_GRUPO"] = "ATIVO"
    df.loc[cond_migrado, "SITUACAO_GRUPO"] = "MIGRADO"
    df.loc[cond_evasao,  "SITUACAO_GRUPO"] = "EVASAO"
    df.loc[cond_sucesso, "SITUACAO_GRUPO"] = "SUCESSO"

    return df


# ============================================================
# V7 · DEDUPLICAÇÃO SEGURA E SITUAÇÕES FORA DO CÁLCULO
# ============================================================

# Uma ENTRADA = o mesmo aluno (matrícula), no mesmo curso e campus,
# com o mesmo ano e semestre de ingresso. Todas as entradas são contadas:
# matrícula em cursos diferentes e reingresso em outro ano/semestre
# são entradas diferentes.
CHAVE_ENTRADA = ["MATRICULA", "COD_CURSO", "CAMPUS", "ANO_INGRESSO", "PERIODO_INGRESSO"]

# Desempate quando a MESMA entrada aparece em mais de uma linha
# (menor número = maior prioridade).
PRIORIDADE_GRUPO = {"SUCESSO": 0, "ATIVO": 1, "MIGRADO": 2, "EVASAO": 3, "OUTRO": 4}


def resolver_mesma_entrada(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mantém UMA linha inteira por entrada. Entre linhas da mesma entrada fica:
      1) a linha com situação Formado;
      2) senão, a de saída mais recente (sem saída = ainda ativa);
      3) senão, a de maior prioridade (Formado > Ativo > Migrado > Evasão > Outro);
      4) senão, a primeira linha do arquivo.
    Linhas de cursos ou ingressos diferentes nunca competem entre si.
    Requer a coluna SITUACAO_GRUPO (classifica_situacao).
    """
    df = df.copy()
    df["_ORDEM"] = range(len(df))
    df["_FORMADO"] = df["SITUACAO_GRUPO"].eq("SUCESSO")
    per_saida = pd.to_numeric(df["PERIODO_EVASAO"], errors="coerce").fillna(0)
    # Saída como número ordenável (ano*10 + semestre); vazia = infinito (ativa).
    df["_SAIDA"] = (df["ANO_EVASAO"] * 10 + per_saida).fillna(float("inf"))
    df["_PRIO"] = df["SITUACAO_GRUPO"].map(PRIORIDADE_GRUPO).fillna(9)

    ordenado = df.sort_values(
        CHAVE_ENTRADA + ["_FORMADO", "_SAIDA", "_PRIO", "_ORDEM"],
        ascending=[True] * len(CHAVE_ENTRADA) + [False, False, True, True],
    )
    return (
        ordenado.loc[~ordenado.duplicated(CHAVE_ENTRADA, keep="first")]
        .sort_values("_ORDEM")
        .drop(columns=["_ORDEM", "_FORMADO", "_SAIDA", "_PRIO"])
    )


def consolidar_entradas(df: pd.DataFrame):
    """
    1. remove cópias exatas (todas as colunas iguais);
    2. deixa UMA linha por entrada (resolver_mesma_entrada);
    3. retira as situações fora do cálculo (SITUACOES_FORA_DO_CALCULO)
       e qualquer situação não mapeada;
    4. confere que não sobrou entrada repetida e que a conta de linhas fecha.
    Retorna (df_consolidado, funil) — funil = contagens de cada etapa,
    usadas na aba "Entenda os Cálculos".
    A matrícula nunca é convertida para maiúsculas/minúsculas:
    na base 16.11.03a, "…x" e "…X" são alunos diferentes.
    """
    lidas = len(df)

    df = df.drop_duplicates()                                   # 1
    sem_copias = len(df)

    df = resolver_mesma_entrada(classifica_situacao(df))        # 2
    entradas = len(df)

    mapeadas = SITUACOES_EVASAO | SITUACOES_SUCESSO | SITUACOES_ATIVO | SITUACOES_MIGRACAO
    fora = df["FORMAEVASAO"].isin(SITUACOES_FORA_DO_CALCULO) | ~df["FORMAEVASAO"].isin(mapeadas)
    fora_por_situacao = df.loc[fora, "FORMAEVASAO"].value_counts().to_dict()
    df = df.loc[~fora]                                          # 3

    # 4 · verificações: o painel para com erro se algo não fechar.
    if df.duplicated(CHAVE_ENTRADA).any():
        raise ValueError("Deduplicação incompleta: ainda há entradas repetidas.")
    if lidas - (lidas - sem_copias) - (sem_copias - entradas) - int(fora.sum()) != len(df):
        raise ValueError("A conciliação de linhas não fecha; revise a consolidação.")

    funil = {
        "lidas": lidas,
        "copias": lidas - sem_copias,
        "repetidas": sem_copias - entradas,
        "fora": fora_por_situacao,
        "usadas": len(df),
    }
    return df.drop(columns=["SITUACAO_GRUPO"]).reset_index(drop=True), funil


@st.cache_data(show_spinner="🔎 Conferindo duplicatas…")
def carregar_consolidado(caminho: str, versao_arquivo=None):
    """Lê o ODS e aplica a consolidação da v7. Cacheado. → (df, funil)"""
    return consolidar_entradas(carregar_ods(caminho, versao_arquivo))


def estimar_prazos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Define o prazo (em semestres) de cada curso.

    1. PRAZOS_OFICIAIS[código], se informado  → fonte "Oficial (PPC)"
    2. senão, estimado dos formados do curso    → fonte "Estimado dos formados"
       = menor DURAÇÃO-SEM com ≥ 10% dos formados de ingresso regular
         (PS/SISU), exigindo ≥ MIN_FORMADOS_PARA_ESTIMAR formados e
         resultado entre PRAZO_MIN_SEM e PRAZO_MAX_SEM.
    3. senão, o prazo de um curso EQUIVALENTE no mesmo campus, quando
       houver um único valor possível:
         a) mesmo nome de curso (ex.: código novo de um curso antigo);
         b) mesma família de código — 2 primeiros dígitos
            (ex.: Letras Núcleo Comum 44M600NC ↔ habilitações 44M601P).
    4. senão → prazo vazio (NaN), fonte "Sem prazo": o ativo não é
       classificado como retido (não há como saber).

    Retorna DataFrame: CAMPUS, COD_CURSO, PRAZO_SEM, PRAZO_FONTE, FORMADOS_USADOS
    """
    cursos = df[["CAMPUS", "COD_CURSO", "CURSO"]].drop_duplicates(["CAMPUS", "COD_CURSO"])
    formados = df[
        df["FORMAEVASAO"].isin(SITUACOES_SUCESSO)
        & df["FORMAINGRESSO"].astype(str).str.match(INGRESSO_REGULAR)
        & df["DURAÇÃO-SEM"].notna()
    ]

    def primeiro_pico(duracoes):
        contagem = duracoes.value_counts().sort_index()
        candidatos = contagem[contagem >= 0.10 * len(duracoes)]
        return candidatos.index.min() if len(candidatos) else float("nan")

    est = (
        formados.groupby(["CAMPUS", "COD_CURSO"])["DURAÇÃO-SEM"]
        .agg(FORMADOS_USADOS="size", ESTIMADO=primeiro_pico)
        .reset_index()
    )
    prazos = cursos.merge(est, on=["CAMPUS", "COD_CURSO"], how="left")
    prazos["FORMADOS_USADOS"] = prazos["FORMADOS_USADOS"].fillna(0).astype(int)
    valido = (
        prazos["FORMADOS_USADOS"].ge(MIN_FORMADOS_PARA_ESTIMAR)
        & prazos["ESTIMADO"].between(PRAZO_MIN_SEM, PRAZO_MAX_SEM)
    )
    prazos["PRAZO_SEM"] = prazos["ESTIMADO"].where(valido)
    prazos["PRAZO_FONTE"] = valido.map({True: "Estimado dos formados", False: "Sem prazo"})

    # Curso equivalente no mesmo campus (só quando o valor é único).
    prazos["FAMILIA"] = prazos["COD_CURSO"].astype(str).str[:2]
    com_prazo = prazos[prazos["PRAZO_SEM"].notna()]
    for chave, rotulo in [("CURSO", "mesmo nome"), ("FAMILIA", "mesma família de código")]:
        unicos = com_prazo.groupby(["CAMPUS", chave])["PRAZO_SEM"].agg(["nunique", "first"])
        unicos = unicos[unicos["nunique"] == 1]["first"]
        sem = prazos["PRAZO_SEM"].isna()
        achado = pd.Series(
            [unicos.get((c, k)) for c, k in zip(prazos["CAMPUS"], prazos[chave])],
            index=prazos.index,
        )
        usar = sem & achado.notna()
        prazos.loc[usar, "PRAZO_SEM"] = achado[usar]
        prazos.loc[usar, "PRAZO_FONTE"] = f"Curso equivalente ({rotulo})"

    oficial = prazos["COD_CURSO"].map(PRAZOS_OFICIAIS)
    prazos.loc[oficial.notna(), "PRAZO_SEM"] = oficial
    prazos.loc[oficial.notna(), "PRAZO_FONTE"] = "Oficial (PPC)"
    return prazos.drop(columns=["ESTIMADO", "FAMILIA", "CURSO"])


def semestre_referencia(df: pd.DataFrame) -> tuple[int, int]:
    """Último semestre registrado na base (ingresso ou saída) → (ano, semestre)."""
    per_ing = pd.to_numeric(df["PERIODO_INGRESSO"], errors="coerce").fillna(1)
    per_sai = pd.to_numeric(df["PERIODO_EVASAO"], errors="coerce").fillna(1)
    codigo = max(
        (df["ANO_INGRESSO"] * 10 + per_ing).max(),
        (df["ANO_EVASAO"] * 10 + per_sai).max(),
    )
    return int(codigo // 10), int(codigo % 10)


def detecta_retencao(
    df: pd.DataFrame,
    anos_tolerancia: int = 2,
    referencia: tuple[int, int] = (2026, 1),
) -> pd.DataFrame:
    """
    Marca como RETIDO os alunos ATIVOS cujo tempo no curso supera
    o prazo do curso + tolerância. Tudo em semestres.

        tempo  = 2 × (ano_ref − ANO_INGRESSO) + (sem_ref − PERIODO_INGRESSO) + 1
        RETIDO se tempo > PRAZO_SEM + 2 × anos_tolerancia

    Ativos de cursos sem prazo (PRAZO_SEM vazio) continuam ATIVO.
    """
    df = df.copy()
    df["SITUACAO_GRUPO_FINAL"] = df["SITUACAO_GRUPO"]
    ano_ref, sem_ref = referencia

    per_ing = pd.to_numeric(df["PERIODO_INGRESSO"], errors="coerce").fillna(1)
    df["TEMPO_SEM"] = 2 * (ano_ref - df["ANO_INGRESSO"]) + (sem_ref - per_ing) + 1

    retido = (
        df["SITUACAO_GRUPO"].eq("ATIVO")
        & df["PRAZO_SEM"].notna()
        & df["TEMPO_SEM"].gt(df["PRAZO_SEM"] + 2 * anos_tolerancia)
    )
    df.loc[retido, "SITUACAO_GRUPO_FINAL"] = "RETIDO"
    return df


@st.cache_data(show_spinner=False)
def preparar_base(
    caminho: str,
    anos_tolerancia: int,
    versao_arquivo,
) -> pd.DataFrame:
    """Carrega, consolida (v7), aplica prazos e classifica a base."""
    df, _ = carregar_consolidado(caminho, versao_arquivo)
    prazos = estimar_prazos(df)
    referencia = semestre_referencia(df)
    df = df.merge(prazos[["CAMPUS", "COD_CURSO", "PRAZO_SEM", "PRAZO_FONTE"]],
                  on=["CAMPUS", "COD_CURSO"], how="left", validate="many_to_one")
    df = classifica_situacao(df)
    return detecta_retencao(df, anos_tolerancia=anos_tolerancia, referencia=referencia)

# ============================================================
# CÁLCULO DOS ÍNDICES
# ============================================================

def calcular_indicadores(
    df: pd.DataFrame,
    ano_col: str = "ANO_INGRESSO"
) -> pd.DataFrame:
    """
    Calcula os índices por coorte (CAMPUS × COD_CURSO × CURSO × ANO_INGRESSO).

    Fórmulas
    --------
    ÍNDICE DE EVASÃO   = (evadidos / total_ingressantes) × 100
    ÍNDICE DE RETENÇÃO = (retidos  / total_ingressantes) × 100
    TAXA DE SUCESSO    = (formados / total_ingressantes) × 100

    Retorna
    -------
    pd.DataFrame com colunas:
        CAMPUS, COD_CURSO, CURSO, ANO, TOTAL, EVADIDOS, RETIDOS, FORMADOS,
        ATIVOS, MIGRADOS, IDX_EVASAO, IDX_RETENCAO, TAXA_SUCESSO
    """
    # Permite reutilizar a classificação já calculada e cacheada pela interface.
    # A classificação só é feita aqui quando a função é usada isoladamente.
    if "SITUACAO_GRUPO_FINAL" not in df.columns:
        df = detecta_retencao(classifica_situacao(df))

    chaves = ["CAMPUS", "COD_CURSO", "CURSO", ano_col]
    contagens = (
        df.groupby(chaves, dropna=False, observed=True)["SITUACAO_GRUPO_FINAL"]
        .value_counts()
        .unstack(fill_value=0)
    )

    resultado = pd.DataFrame(index=contagens.index)
    resultado["TOTAL"] = contagens.sum(axis=1).astype(int)
    for coluna_saida, situacao in {
        "EVADIDOS": "EVASAO",
        "RETIDOS": "RETIDO",
        "FORMADOS": "SUCESSO",
        "ATIVOS": "ATIVO",
        "MIGRADOS": "MIGRADO",
    }.items():
        resultado[coluna_saida] = (
            contagens[situacao].astype(int)
            if situacao in contagens.columns
            else 0
        )

    resultado["IDX_EVASAO"] = (
        resultado["EVADIDOS"] / resultado["TOTAL"] * 100
    ).round(2)
    resultado["IDX_RETENCAO"] = (
        resultado["RETIDOS"] / resultado["TOTAL"] * 100
    ).round(2)
    resultado["TAXA_SUCESSO"] = (
        resultado["FORMADOS"] / resultado["TOTAL"] * 100
    ).round(2)

    resultado = resultado.reset_index().rename(columns={ano_col: "ANO"})
    resultado["ANO"] = resultado["ANO"].astype(int)

    return resultado.sort_values(
        ["CAMPUS", "CURSO", "COD_CURSO", "ANO"]
    ).reset_index(drop=True)

# ============================================================
# ABA "ENTENDA OS CÁLCULOS" (explicação para o usuário leigo)
# ============================================================

# Categorias na ordem em que aparecem na explicação.
CATEGORIAS_DIDATICAS = [
    ("SUCESSO", "🎓", "Formado", COR_SUCESSO,
     "Concluiu o curso.",
     "Formado"),
    ("EVASAO", "🚪", "Evadido", COR_EVASAO,
     "Saiu do curso sem concluir.",
     "Desistência, Desvinculado, Matrícula Cancelada, Jubilado, Transferência "
     "Interna/Externa/Ex-offício, Reopção de Curso, Troca de Turno"),
    ("RETIDO", "⏳", "Retido", COR_RETENCAO,
     "Ainda matriculado, mas já passou do prazo do curso + tolerância.",
     "Vinculado ou Reingresso Administrativo, com tempo acima do limite"),
    ("ATIVO", "📚", "Ativo no prazo", COR_ACENTO,
     "Ainda matriculado e dentro do prazo.",
     "Vinculado ou Reingresso Administrativo, com tempo dentro do limite"),
    ("MIGRADO", "🔀", "Migrado para a UFNT", COR_TEXTO_LEVE,
     f"Era aluno de Araguaína ou Tocantinópolis e foi transferido para a UFNT em {ANO_TRANSICAO_UFNT}.",
     "Transição UFNT"),
]

EXPLICA_FORA = {
    "Declinante": "aprovados que não chegaram a cursar",
    "Falecimento": "falecimento não é evasão",
    "Habilitação": "o aluno segue na habilitação escolhida (outra entrada, já contada)",
    "Encerramento Intercâmbio Internacional": "vínculo temporário, não é aluno regular do curso",
    "Encerramento Mobilidade Acadêmica": "vínculo temporário, não é aluno regular do curso",
    "Encerramento Aluno Especial": "vínculo temporário, não é aluno regular do curso",
    "Encerramento de Convênio": "vínculo temporário, não é aluno regular do curso",
    "Encerramento/Apostilamento": "registro administrativo, não é trajetória de aluno",
    "Em análise": "situação ainda não definida",
}


def _caixa(html: str) -> None:
    """Caixa branca de texto explicativo (HTML em uma linha, sem indentação)."""
    st.markdown(f'<div class="explica">{html}</div>', unsafe_allow_html=True)


def mostrar_metodologia(df_filtrado, funil, tolerancia, referencia) -> None:
    """
    Aba didática: explica em linguagem simples de onde vem cada número.
    Todos os exemplos usam os dados do filtro atual da barra lateral.
    """
    total = len(df_filtrado)
    contagem = df_filtrado["SITUACAO_GRUPO_FINAL"].value_counts()
    n = {chave: int(contagem.get(chave, 0)) for chave, *_ in CATEGORIAS_DIDATICAS}

    def pct(valor):
        return round(valor / total * 100, 1) if total else 0.0

    ano_ref, sem_ref = referencia

    _caixa(
        "Este painel responde a uma pergunta simples: <b>o que aconteceu com os alunos que "
        "entraram em cada curso da UFT?</b> Cada aluno é acompanhado até a situação registrada "
        "no relatório 16.11.03a e colocado em <b>uma</b> de cinco situações. Os índices mostram "
        "quanto cada situação representa do total.<br><br>"
        "💡 <b>Todos os exemplos desta página usam os números do filtro escolhido agora na barra "
        "lateral.</b> Mude o filtro e as contas mudam junto."
    )

    # ---------------------------------------------------------------- 1
    secao("1️⃣ As cinco situações possíveis de um aluno")
    cartoes = []
    for chave, icone, nome, cor, significado, sistema in CATEGORIAS_DIDATICAS:
        cartoes.append(
            f'<div style="background:white;border:1px solid {COR_BORDA};border-left:6px solid {cor};'
            f'border-radius:12px;padding:14px 16px;flex:1;min-width:200px;">'
            f'<div style="font-size:1.3rem;">{icone} <b style="color:{COR_PRIMARIA};font-size:1rem;">{nome}</b></div>'
            f'<div style="font-family:\'Lora\',serif;font-size:1.5rem;font-weight:700;color:{cor};margin:4px 0;">'
            f'{n[chave]:,} <span style="font-size:0.9rem;color:{COR_TEXTO_LEVE};">({pct(n[chave])}%)</span></div>'
            f'<div style="font-size:0.85rem;color:{COR_TEXTO};">{significado}</div>'
            f'<div style="font-size:0.72rem;color:{COR_TEXTO_LEVE};margin-top:6px;">'
            f'<b>No sistema:</b> {sistema}</div></div>'
        )
    st.markdown('<div class="kpi-wrapper">' + "".join(cartoes) + "</div>", unsafe_allow_html=True)

    # Barra 100%: mostra visualmente como o total se divide.
    fig = go.Figure()
    for chave, icone, nome, cor, *_ in CATEGORIAS_DIDATICAS:
        if n[chave]:
            fig.add_trace(go.Bar(
                y=["Total"], x=[pct(n[chave])], orientation="h", name=nome,
                marker_color=cor, text=f"{pct(n[chave])}%", textposition="inside",
                insidetextanchor="middle",
                hovertemplate=f"{nome}: {n[chave]:,} alunos ({pct(n[chave])}%)<extra></extra>",
            ))
    fig.update_layout(**LAYOUT_BASE)
    fig.update_layout(barmode="stack", height=170, showlegend=True,
                      legend=dict(orientation="h", y=-0.35, traceorder="normal"),
                      xaxis=dict(range=[0, 100], ticksuffix="%"), yaxis=dict(visible=False))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    soma = " + ".join(f"{n[c]:,}" for c, *_ in CATEGORIAS_DIDATICAS)
    st.caption(f"As cinco situações somam o total: {soma} = {total:,} alunos (100%). "
               "Nenhum aluno fica de fora e nenhum é contado duas vezes.")

    # ---------------------------------------------------------------- 2
    secao("2️⃣ As fórmulas")
    _caixa(
        "Todo índice do painel é uma <b>proporção</b>: pega-se uma parte dos alunos, divide-se "
        "pelo total e multiplica-se por 100. O resultado responde à pergunta "
        "<b>“de cada 100 alunos que entraram, quantos…?”</b>"
    )
    c1, c2 = st.columns(2)
    with c1:
        st.latex(r"\text{Taxa de Sucesso} = \frac{\text{Formados}}{\text{Total de alunos}} \times 100")
        st.latex(r"\text{Índice de Evasão} = \frac{\text{Evadidos}}{\text{Total de alunos}} \times 100")
    with c2:
        st.latex(r"\text{Índice de Retenção} = \frac{\text{Retidos}}{\text{Total de alunos}} \times 100")
        st.latex(r"\text{Migração UFNT} = \frac{\text{Migrados}}{\text{Total de alunos}} \times 100")

    st.markdown("**Aplicando ao filtro atual:**")
    exemplos = pd.DataFrame([
        {"Índice": "🎓 Taxa de Sucesso", "Conta": f"{n['SUCESSO']:,} ÷ {total:,} × 100", "Resultado": f"{pct(n['SUCESSO'])}%"},
        {"Índice": "🚪 Índice de Evasão", "Conta": f"{n['EVASAO']:,} ÷ {total:,} × 100", "Resultado": f"{pct(n['EVASAO'])}%"},
        {"Índice": "⏳ Índice de Retenção", "Conta": f"{n['RETIDO']:,} ÷ {total:,} × 100", "Resultado": f"{pct(n['RETIDO'])}%"},
        {"Índice": "🔀 Migração UFNT", "Conta": f"{n['MIGRADO']:,} ÷ {total:,} × 100", "Resultado": f"{pct(n['MIGRADO'])}%"},
    ])
    st.dataframe(exemplos, use_container_width=True, hide_index=True)
    _caixa(
        f"📣 <b>Em linguagem do dia a dia:</b> de cada 100 alunos deste filtro, "
        f"<b>{pct(n['SUCESSO'])}</b> se formaram, <b>{pct(n['EVASAO'])}</b> saíram sem concluir, "
        f"<b>{pct(n['RETIDO'])}</b> estão atrasados, <b>{pct(n['ATIVO'])}</b> estão cursando no prazo"
        + (f" e <b>{pct(n['MIGRADO'])}</b> foram transferidos para a UFNT." if n["MIGRADO"] else ".")
        + "<br><br>Os gráficos por ano, campus e curso usam <b>a mesma fórmula</b>, só que aplicada "
        "a cada grupo separadamente (por exemplo, só os alunos que entraram em 2018)."
    )

    # ---------------------------------------------------------------- 3
    secao("3️⃣ Quem entra na conta")
    _caixa(
        "A unidade de contagem é a <b>entrada</b>: <b>um aluno, em um curso, em um ano/semestre de "
        "ingresso</b>. Na prática:<br>"
        "• O aluno que <b>trocou de curso</b> aparece nos dois cursos, porque teve uma história em cada um "
        "(saiu de um, entrou no outro).<br>"
        "• O aluno que <b>saiu e voltou</b> ao mesmo curso em outro ano aparece nas duas turmas.<br>"
        "• <b>Linhas repetidas</b> do relatório (a mesma entrada registrada duas vezes) são contadas "
        "<b>uma única vez</b>. Quando as linhas repetidas discordam, vale a que tem “Formado” ou, "
        "não havendo, a situação mais recente."
    )
    passos = [
        ("📄", "Linhas no relatório 16.11.03a", f"{funil['lidas']:,}", COR_PRIMARIA),
        ("➖", "Cópias idênticas (mesma linha duas vezes)", f"− {funil['copias']:,}", COR_TEXTO_LEVE),
        ("➖", "Repetições da mesma entrada com dados diferentes", f"− {funil['repetidas']:,}", COR_TEXTO_LEVE),
    ]
    for situacao, qtd in sorted(funil["fora"].items(), key=lambda x: -x[1]):
        motivo = EXPLICA_FORA.get(situacao, "situação não prevista nas regras")
        passos.append(("➖", f"{situacao} <span style='color:{COR_TEXTO_LEVE};'>— {motivo}</span>",
                       f"− {qtd:,}", COR_TEXTO_LEVE))
    passos.append(("✅", "<b>Alunos considerados nos cálculos</b>", f"<b>{funil['usadas']:,}</b>", COR_SUCESSO))
    linhas = "".join(
        f'<div style="display:flex;justify-content:space-between;gap:12px;padding:7px 4px;'
        f'border-bottom:1px solid {COR_BORDA};font-size:0.88rem;">'
        f'<span>{icone} {texto}</span><span style="color:{cor};white-space:nowrap;">{valor}</span></div>'
        for icone, texto, valor, cor in passos
    )
    _caixa(linhas)
    st.caption("Esta conta é da base inteira, antes dos filtros da barra lateral.")

    # ---------------------------------------------------------------- 4
    secao("4️⃣ Como a retenção é medida")
    limite_extra = 2 * tolerancia
    tempo_ex = 2 * (ano_ref - 2019) + (sem_ref - 1) + 1
    limite_ex = 8 + limite_extra
    _caixa(
        "Um aluno é <b>retido</b> quando ainda está matriculado, mas já passou do tempo esperado para se formar. "
        "São três passos:<br><br>"
        "<b>① Prazo do curso</b> (em semestres). Vem do PPC quando informado no código "
        "(<code>PRAZOS_OFICIAIS</code>). Quando não está informado, é estimado pelos próprios formados do curso: "
        "é a menor duração em que se formaram pelo menos 10% dos alunos que entraram por processo seletivo. "
        "Por exemplo, Medicina fica com 12 semestres, Direito com 10 e Administração com 8.<br>"
        f"<b>② Tempo no curso</b>: semestres desde o ingresso até <b>{ano_ref}.{sem_ref}</b>, o último "
        "semestre registrado na base, contando o semestre de entrada.<br>"
        f"<b>③ Regra</b>: retido se o tempo for maior que <b>prazo + tolerância</b>. A tolerância está em "
        f"<b>{tolerancia} ano(s) = {limite_extra} semestre(s)</b> e pode ser mudada na barra lateral."
    )
    st.latex(r"\text{Tempo} = 2 \times (\text{ano}_{ref} - \text{ano}_{ingresso}) + "
             r"(\text{semestre}_{ref} - \text{semestre}_{ingresso}) + 1")
    _caixa(
        f"🧮 <b>Exemplo:</b> aluno que entrou em <b>2019.1</b> num curso de <b>8 semestres</b>, ainda matriculado.<br>"
        f"Tempo = 2 × ({ano_ref} − 2019) + ({sem_ref} − 1) + 1 = <b>{tempo_ex} semestres</b>.<br>"
        f"Limite = 8 + {limite_extra} = <b>{limite_ex} semestres</b>.<br>"
        + (f"Como {tempo_ex} &gt; {limite_ex}, ele é <b>retido</b>." if tempo_ex > limite_ex
           else f"Como {tempo_ex} ≤ {limite_ex}, ele está <b>no prazo</b>.")
    )

    # ---------------------------------------------------------------- 5
    secao("5️⃣ Araguaína e Tocantinópolis: a transição para a UFNT")
    _caixa(
        f"Em {ANO_TRANSICAO_UFNT}, os campi de <b>Araguaína</b> e <b>Tocantinópolis</b> passaram a formar a "
        "<b>Universidade Federal do Norte do Tocantins (UFNT)</b>, e os alunos desses campi foram "
        "transferidos automaticamente.<br><br>"
        "Esses alunos <b>não abandonaram o curso</b>, então não contam como evasão. Eles também <b>não se formaram "
        "na UFT</b>, então não contam como sucesso. Por isso formam uma situação própria, "
        "<b>Migrado para a UFNT</b>, e <b>continuam no total</b>, porque eram alunos da UFT até a transição."
    )
    ufnt = df_filtrado[df_filtrado["CAMPUS"].isin(CAMPI_UFNT)]
    if len(ufnt):
        tot_u = len(ufnt)
        mig_u = int(ufnt["SITUACAO_GRUPO_FINAL"].eq("MIGRADO").sum())
        ev_u = int(ufnt["SITUACAO_GRUPO_FINAL"].eq("EVASAO").sum())
        errado = round(ev_u / (tot_u - mig_u) * 100, 1) if tot_u > mig_u else 100.0
        certo = round(ev_u / tot_u * 100, 1)
        col_e, col_c = st.columns(2)
        with col_e:
            _caixa(
                f"❌ <b>Se tirássemos os migrados do total</b><br>"
                f"{ev_u:,} ÷ ({tot_u:,} − {mig_u:,}) × 100 = "
                f"<b style='color:{COR_EVASAO};font-size:1.3rem;'>{errado}%</b> de evasão<br>"
                f"<span style='font-size:0.8rem;color:{COR_TEXTO_LEVE};'>Sobrariam quase só os alunos que já "
                "tinham saído antes da transição, e a evasão ficaria inflada (em alguns cursos, 100%).</span>"
            )
        with col_c:
            _caixa(
                f"✅ <b>Mantendo os migrados no total (usado no painel)</b><br>"
                f"{ev_u:,} ÷ {tot_u:,} × 100 = "
                f"<b style='color:{COR_SUCESSO};font-size:1.3rem;'>{certo}%</b> de evasão<br>"
                f"<span style='font-size:0.8rem;color:{COR_TEXTO_LEVE};'>Mostra a evasão real enquanto esses "
                "alunos estavam na UFT.</span>"
            )
        st.caption(
            "Números de Araguaína e Tocantinópolis dentro do filtro atual. Nas turmas que entraram de 2022 "
            "a 2024 nesses campi, quase todos os alunos migraram. Por isso a taxa de sucesso dessas turmas "
            "é baixa: elas não tiveram tempo de se formar na UFT."
        )
    else:
        st.caption("O filtro atual não inclui Araguaína nem Tocantinópolis.")

    # ---------------------------------------------------------------- 6
    secao("6️⃣ Como ler cada parte do painel")
    _caixa(
        "📊 <b>Velocímetros</b>: os três índices do filtro inteiro. Quanto mais cheio, maior o percentual.<br>"
        "📈 <b>Evolução por ano de ingresso</b>: cada ponto é uma turma (os alunos que entraram naquele ano). "
        "Turmas recentes ainda estão cursando, por isso têm pouco sucesso e pouca evasão, e isso é esperado. "
        "A linha tracejada marca a transição para a UFNT.<br>"
        "🏛️ <b>Comparativo por campus</b>: o mesmo índice calculado para cada campus. Os campi marcados "
        f"“(UFNT {ANO_TRANSICAO_UFNT})” têm alunos migrados no total.<br>"
        "🔥 <b>Mapa de calor</b>: evasão de cada curso em cada campus. Quanto mais vermelho, maior a evasão. "
        "O traço (–) significa que o curso não existe naquele campus, e não evasão zero. Passe o mouse "
        "para ver o total de alunos e de migrados.<br>"
        "📋 <b>Top cursos</b> e 🎯 <b>dispersão</b>: só entram cursos com pelo menos 20 alunos no filtro, para que "
        "um curso com 1 ou 2 alunos não apareça com 0% ou 100%.<br>"
        "ⓘ <b>Cartões do topo</b>: passe o mouse sobre eles para ver a conta de cada número."
    )

    # ---------------------------------------------------------------- 7
    secao("⚠️ Cuidados ao interpretar")
    _caixa(
        "• <b>Evasão do curso não é evasão da universidade.</b> Quem trocou de curso (transferência interna, "
        "reopção) conta como evasão do curso de origem, mesmo continuando na UFT.<br>"
        "• <b>É um retrato da data do relatório.</b> A situação de cada aluno é a registrada no arquivo "
        "16.11.03a. Alunos ativos ainda podem se formar ou evadir.<br>"
        "• <b>Turmas recentes estão incompletas.</b> Compare turmas antigas com antigas e recentes com recentes.<br>"
        "• <b>Prazos estimados</b> são uma aproximação. Com os prazos oficiais dos PPCs preenchidos no código, "
        "a retenção fica exata.<br>"
        "• Esta é uma <b>metodologia interna da UFT</b>, inspirada nos indicadores de trajetória do Inep, mas "
        "que não substitui os indicadores oficiais (Inep/TCU)."
    )

    with st.expander("📖 Glossário: situação no sistema → situação no painel"):
        glossario = [{"Situação no sistema (FORMAEVASAO)": s_, "No painel": "🎓 Formado"} for s_ in sorted(SITUACOES_SUCESSO)]
        glossario += [{"Situação no sistema (FORMAEVASAO)": s_, "No painel": "🚪 Evadido"} for s_ in sorted(SITUACOES_EVASAO)]
        glossario += [{"Situação no sistema (FORMAEVASAO)": s_, "No painel": "📚 Ativo (no prazo) ou ⏳ Retido"} for s_ in sorted(SITUACOES_ATIVO)]
        glossario += [{"Situação no sistema (FORMAEVASAO)": s_, "No painel": "🔀 Migrado para a UFNT"} for s_ in sorted(SITUACOES_MIGRACAO)]
        glossario += [{"Situação no sistema (FORMAEVASAO)": s_, "No painel": f"⛔ Fora da conta — {EXPLICA_FORA.get(s_, '')}"}
                      for s_ in sorted(SITUACOES_FORA_DO_CALCULO)]
        st.dataframe(pd.DataFrame(glossario), use_container_width=True, hide_index=True)


# ============================================================
# INTERFACE · SIDEBAR
# ============================================================

st.sidebar.markdown(
    '<div style="padding:0.8rem 0 1.2rem;display:flex;align-items:center;gap:12px;">'
    + selo_logo(46)
    + '<div><div style="font-family:\'Lora\',serif;font-size:1.1rem;font-weight:700;">UFT · Indicadores</div>'
    '<div style="font-size:0.75rem;opacity:0.7;margin-top:0.2rem;">Evasão · Retenção · Sucesso</div></div>'
    '</div>',
    unsafe_allow_html=True
)

tolerancia = st.sidebar.slider(
    "Tolerância p/ retenção (anos além da duração prevista)",
    min_value=0,
    max_value=5,
    value=2,
    help=(
        "Alunos ATIVOS com tempo cursado > prazo do curso + "
        "este valor serão classificados como RETIDOS."
    )
)

# ============================================================
# HEADER
# ============================================================

st.markdown(
    f"""
<div style="
    background: linear-gradient(135deg, {COR_PRIMARIA} 0%, {COR_SECUNDARIA} 60%, {COR_ACENTO} 100%);
    border-radius: 16px;
    padding: 28px 36px;
    margin-bottom: 22px;
    display: flex;
    align-items: center;
    gap: 20px;
    box-shadow: 0 8px 32px rgba(13,43,78,0.22);
">
{selo_logo(78)}
<div style="flex: 1;">
    <div style="font-family:'Lora',serif;font-size:1.6rem;font-weight:700;color:white;">
        Indicadores Acadêmicos - PROGRAD
    </div>
    <div style="font-size:0.88rem;color:rgba(255,255,255,0.72);margin-top:4px;">
        Universidade Federal do Tocantins · Evasão · Retenção · Sucesso
    </div>
</div>
</div>
""",
    unsafe_allow_html=True
)

# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================

# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================

# Caminho resolvido a partir da pasta do próprio script. Assim, o painel pode
# ser iniciado de outro diretório sem deixar de encontrar o arquivo ODS.
CAMINHO_ODS = os.path.join(DIRETORIO_SCRIPT, "16.11.03a.ods")

if not os.path.isfile(CAMINHO_ODS):
    st.error(
        f"❌ Arquivo não encontrado: **{CAMINHO_ODS}**\n\n"
        "Verifique se o arquivo está na mesma pasta do script "
        "e se o nome está correto."
    )
    st.stop()

try:
    versao_arquivo = (
        os.path.getmtime(CAMINHO_ODS),
        os.path.getsize(CAMINHO_ODS),
    )
    df_raw_final = preparar_base(
        CAMINHO_ODS,
        tolerancia,
        versao_arquivo,
    )
    _, funil_base = carregar_consolidado(CAMINHO_ODS, versao_arquivo)

except Exception as e:
    st.error(f"❌ Erro ao carregar arquivo: {e}")
    st.stop()

# ============================================================
# FILTROS DINÂMICOS NA SIDEBAR
# ============================================================

st.sidebar.markdown("### 1. Selecione os campi")
st.sidebar.caption(
    "Todos estão marcados inicialmente. Desmarque “Todos os campi” para escolher uma combinação."
)

campus_lista = sorted(df_raw_final["CAMPUS"].dropna().unique())


def marcar_todos_os_campi():
    if st.session_state.get("todos_os_campi", False):
        for campus in campus_lista:
            st.session_state[f"campus::{campus}"] = True


todos_os_campi = st.sidebar.checkbox(
    "Todos os campi",
    value=True,
    key="todos_os_campi",
    on_change=marcar_todos_os_campi,
)

campus_sel = []
for campus in campus_lista:
    chave_campus = f"campus::{campus}"
    if chave_campus not in st.session_state:
        st.session_state[chave_campus] = True

    campus_marcado = st.sidebar.checkbox(
        campus,
        key=chave_campus,
        disabled=todos_os_campi,
    )
    if todos_os_campi or campus_marcado:
        campus_sel.append(campus)

st.sidebar.caption(f"{len(campus_sel)} de {len(campus_lista)} campi selecionados.")

if not campus_sel:
    st.sidebar.warning("Selecione pelo menos um campus.")
    st.stop()

mask_campus = df_raw_final["CAMPUS"].isin(campus_sel)

df_raw_final["CURSO_FILTRO"] = (
    df_raw_final["COD_CURSO"].fillna("Sem código").astype(str)
    + " — "
    + df_raw_final["CURSO"].fillna("Sem nome").astype(str)
)
cursos_disp = sorted(
    df_raw_final.loc[mask_campus, "CURSO_FILTRO"].unique(),
    key=chave_ordenacao_curso,
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 2. Selecione os cursos")
st.sidebar.caption(
    "Mantenha todos marcados ou desative a opção abaixo para pesquisar e escolher vários cursos."
)

selecionar_todos_cursos = st.sidebar.checkbox(
    "Todos os cursos disponíveis",
    value=True,
    key="todos_os_cursos",
    help="Desmarque para escolher uma combinação personalizada de cursos.",
)

# Remove da memória do widget cursos que deixaram de pertencer aos campi
# escolhidos, evitando seleções invisíveis quando o campus é alterado.
chave_cursos = "cursos_personalizados"
if chave_cursos in st.session_state:
    st.session_state[chave_cursos] = [
        curso
        for curso in st.session_state[chave_cursos]
        if curso in cursos_disp
    ]

if selecionar_todos_cursos:
    curso_sel = cursos_disp
    st.sidebar.caption(f"Todos os {len(cursos_disp)} cursos disponíveis estão selecionados.")
else:
    curso_sel = st.sidebar.multiselect(
        "Pesquisar e selecionar cursos",
        cursos_disp,
        key=chave_cursos,
        format_func=formatar_opcao_curso,
        placeholder="Digite o código ou parte do nome",
        help="Pesquise e selecione quantos cursos desejar.",
    )

    if not curso_sel:
        st.sidebar.warning("Selecione pelo menos um curso.")
        st.stop()

    st.sidebar.markdown(f"**Cursos selecionados: {len(curso_sel)}**")
    for curso_escolhido in curso_sel:
        codigo_selecionado, nome_selecionado = separar_rotulo_curso(curso_escolhido)
        st.sidebar.markdown(
            f"- {nome_selecionado}  \n  Código: **{codigo_selecionado}**"
        )

mask_curso = df_raw_final["CURSO_FILTRO"].isin(curso_sel)

anos_disp    = sorted(
    df_raw_final["ANO_INGRESSO"].dropna().astype(int).unique()
)
ano_min, ano_max = int(min(anos_disp)), int(max(anos_disp))
ano_range    = st.sidebar.slider(
    "Período de ingresso", ano_min, ano_max, (ano_min, ano_max)
)

# Aplica filtros
mask = (
    mask_campus &
    mask_curso &
    df_raw_final["ANO_INGRESSO"].between(ano_range[0], ano_range[1])
)
df_filtrado = df_raw_final[mask].copy()

if df_filtrado.empty:
    st.warning(
        "Nenhum registro foi encontrado para a combinação de campi, cursos "
        "e período selecionada. Ajuste os filtros para continuar."
    )
    st.stop()

# ============================================================
# CALCULA INDICADORES
# ============================================================

indicadores = calcular_indicadores(df_filtrado)

total_alunos = len(df_filtrado)                        # entradas (1 por aluno/curso/ingresso)
total_cursos = df_filtrado["COD_CURSO"].nunique()
total_campus = df_filtrado["CAMPUS"].nunique()

evadidos_total  = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "EVASAO").sum()
retidos_total   = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "RETIDO").sum()
formados_total  = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "SUCESSO").sum()
ativos_total    = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "ATIVO").sum()
migrados_total  = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "MIGRADO").sum()

idx_evasao_geral   = round(evadidos_total / max(total_alunos, 1) * 100, 1)
idx_retencao_geral = round(retidos_total  / max(total_alunos, 1) * 100, 1)
taxa_sucesso_geral = round(formados_total / max(total_alunos, 1) * 100, 1)
idx_migracao_geral = round(migrados_total / max(total_alunos, 1) * 100, 1)

# ============================================================
# ABAS · PAINEL  |  ENTENDA OS CÁLCULOS
# ============================================================

aba_painel, aba_metodo = st.tabs(["📊 Painel de Indicadores", "📘 Entenda os Cálculos"])

with aba_painel:
    # ============================================================
    # KPIs GLOBAIS
    # ============================================================

    def _conta(parte, total, pct):
        return f"= {parte:,} ÷ {total:,} × 100 = {pct}%"


    cartoes = [
        kpi_card("👥", f"{total_alunos:,}", "TOTAL ALUNOS", COR_PRIMARIA, ajuda=(
            "Alunos considerados no filtro atual.\n"
            "Cada aluno conta 1 vez por curso e por ingresso.\n"
            "Não entram: declinantes, falecimentos e encerramentos\n"
            "de mobilidade/intercâmbio. Detalhes na aba 📘.")),
        kpi_card("🎓", f"{formados_total:,}", "FORMADOS", COR_SUCESSO, f"Sucesso: {taxa_sucesso_geral}%", ajuda=(
            "Taxa de Sucesso = Formados ÷ Total × 100\n"
            + _conta(formados_total, total_alunos, taxa_sucesso_geral))),
        kpi_card("🚪", f"{evadidos_total:,}", "EVADIDOS", COR_EVASAO, f"Evasão: {idx_evasao_geral}%", ajuda=(
            "Saíram do curso sem concluir (desistência, desvínculo,\n"
            "cancelamento, jubilamento, transferência, reopção…).\n"
            "Índice de Evasão = Evadidos ÷ Total × 100\n"
            + _conta(evadidos_total, total_alunos, idx_evasao_geral))),
        kpi_card("⏳", f"{retidos_total:,}", "RETIDOS", COR_RETENCAO, f"Retenção: {idx_retencao_geral}%", ajuda=(
            f"Ativos há mais tempo que o prazo do curso + {tolerancia} ano(s).\n"
            "Índice de Retenção = Retidos ÷ Total × 100\n"
            + _conta(retidos_total, total_alunos, idx_retencao_geral))),
        kpi_card("📚", f"{ativos_total:,}", "ATIVOS (no prazo)", COR_ACENTO, ajuda=(
            "Alunos ainda matriculados e dentro do prazo do curso\n"
            "(ou em cursos novos, ainda sem prazo definido).")),
    ]
    if migrados_total:
        cartoes.append(kpi_card("🔀", f"{migrados_total:,}", "MIGRADOS UFNT", COR_TEXTO_LEVE, f"Migração: {idx_migracao_geral}%", ajuda=(
            f"Alunos de Araguaína e Tocantinópolis transferidos para a UFNT em {ANO_TRANSICAO_UFNT}.\n"
            "Continuam no total, mas NÃO contam como evasão nem como sucesso.\n"
            + _conta(migrados_total, total_alunos, idx_migracao_geral))))
    cartoes += [
        kpi_card("🏛️", f"{total_campus}", "CAMPUS", COR_PRIMARIA),
        kpi_card("📋", f"{total_cursos}", "CURSOS", COR_SECUNDARIA),
    ]
    st.markdown('<div class="kpi-wrapper">' + "".join(cartoes) + "</div>", unsafe_allow_html=True)

    # ============================================================
    # GAUGE TRIPLO
    # ============================================================

    secao("📊 Visão Geral dos Índices")

    c_g1, c_g2, c_g3 = st.columns(3)

    def gauge(valor, titulo, cor_bar, cor_steps):
        fig = go.Figure(
            go.Indicator(
                mode="gauge+number",
                value=valor,
                number={"suffix": "%"},
                title={"text": titulo, "font": {"size": 13}},
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {"color": cor_bar},
                    "steps": [
                        {"range": [0, 33],  "color": cor_steps[0]},
                        {"range": [33, 66], "color": cor_steps[1]},
                        {"range": [66, 100],"color": cor_steps[2]},
                    ],
                }
            )
        )
        fig.update_layout(**LAYOUT_BASE, height=230)
        return fig

    with c_g1:
        chart_container(gauge(
            taxa_sucesso_geral,
            "Taxa de Sucesso",
            COR_SUCESSO,
            ["#EEF2F8", "#B8DFC8", "#1B7A4A"]
        ))

    with c_g2:
        chart_container(gauge(
            idx_evasao_geral,
            "Índice de Evasão",
            COR_EVASAO,
            ["#EEF2F8", "#F7C5BE", "#C0392B"]
        ))

    with c_g3:
        chart_container(gauge(
            idx_retencao_geral,
            "Índice de Retenção",
            COR_RETENCAO,
            ["#EEF2F8", "#FAE4C0", "#E07B00"]
        ))

    # ============================================================
    # EVOLUÇÃO TEMPORAL DOS ÍNDICES
    # ============================================================

    secao("📈 Evolução Temporal dos Índices por Ano de Ingresso")

    # Calcula por ano de ingresso
    registros_ano = []
    for ano, grp in df_filtrado.groupby("ANO_INGRESSO"):
        tot = len(grp)
        ev  = (grp["SITUACAO_GRUPO_FINAL"] == "EVASAO").sum()
        rt  = (grp["SITUACAO_GRUPO_FINAL"] == "RETIDO").sum()
        fm  = (grp["SITUACAO_GRUPO_FINAL"] == "SUCESSO").sum()
        registros_ano.append({
            "ANO"          : int(ano),
            "IDX_EVASAO"   : round(ev / tot * 100, 2),
            "IDX_RETENCAO" : round(rt / tot * 100, 2),
            "TAXA_SUCESSO" : round(fm / tot * 100, 2),
            "TOTAL"        : tot,
        })

    df_evolucao = pd.DataFrame(registros_ano).sort_values("ANO")

    fig_evolucao = go.Figure()
    fig_evolucao.add_trace(go.Scatter(
        x=df_evolucao["ANO"], y=df_evolucao["TAXA_SUCESSO"],
        mode="lines+markers", name="Taxa de Sucesso",
        line=dict(color=COR_SUCESSO, width=2.5),
        marker=dict(size=6)
    ))
    fig_evolucao.add_trace(go.Scatter(
        x=df_evolucao["ANO"], y=df_evolucao["IDX_EVASAO"],
        mode="lines+markers", name="Índice de Evasão",
        line=dict(color=COR_EVASAO, width=2.5),
        marker=dict(size=6)
    ))
    fig_evolucao.add_trace(go.Scatter(
        x=df_evolucao["ANO"], y=df_evolucao["IDX_RETENCAO"],
        mode="lines+markers", name="Índice de Retenção",
        line=dict(color=COR_RETENCAO, width=2.5, dash="dot"),
        marker=dict(size=6)
    ))
    fig_evolucao.update_layout(
        **LAYOUT_BASE,
        height=380,
        yaxis_title="(%)",
        xaxis_title="Ano de Ingresso",
        legend=dict(orientation="h", y=1.1)
    )
    # v7: marca a transição para a UFNT quando Araguaína/Tocantinópolis estão no filtro.
    if set(campus_sel) & CAMPI_UFNT and ano_range[0] <= ANO_TRANSICAO_UFNT <= ano_range[1]:
        fig_evolucao.add_vline(
            x=ANO_TRANSICAO_UFNT,
            line_dash="dash",
            line_color=COR_TEXTO_LEVE,
            annotation_text=f"Transição UFNT ({ANO_TRANSICAO_UFNT})",
            annotation_position="top left",
            annotation_font=dict(size=11, color=COR_TEXTO_LEVE),
        )
    chart_container(fig_evolucao)

    # ============================================================
    # COMPARATIVO POR CAMPUS
    # ============================================================

    secao("🏛️ Comparativo por Campus")

    registros_campus = []
    for campus, grp in df_filtrado.groupby("CAMPUS"):
        tot = len(grp)
        registros_campus.append({
            "CAMPUS"       : campus,
            "IDX_EVASAO"   : round((grp["SITUACAO_GRUPO_FINAL"] == "EVASAO").sum() / tot * 100, 2),
            "IDX_RETENCAO" : round((grp["SITUACAO_GRUPO_FINAL"] == "RETIDO").sum() / tot * 100, 2),
            "TAXA_SUCESSO" : round((grp["SITUACAO_GRUPO_FINAL"] == "SUCESSO").sum() / tot * 100, 2),
            "TOTAL"        : tot,
        })

    df_campus = pd.DataFrame(registros_campus)
    df_campus["CAMPUS"] = df_campus["CAMPUS"].map(rotulo_campus)   # v7: "Araguaína (UFNT 2023)"

    c_camp1, c_camp2, c_camp3 = st.columns(3)

    def bar_campus(col_y, titulo, cor):
        df_ord = df_campus.sort_values(col_y, ascending=True)
        fig    = px.bar(
            df_ord, x=col_y, y="CAMPUS",
            orientation="h", text=col_y,
            height=300, color_discrete_sequence=[cor]
        )
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
        fig.update_xaxes(range=[0, max(df_ord[col_y].max() * 1.25, 1)])   # espaço para o rótulo
        fig.update_layout(**LAYOUT_BASE, xaxis_title="%", yaxis_title="")
        fig.update_layout(title=dict(text=titulo, font=dict(size=13)))
        return fig

    with c_camp1:
        chart_container(bar_campus("TAXA_SUCESSO",  "Taxa de Sucesso (%)",    COR_SUCESSO))
    with c_camp2:
        chart_container(bar_campus("IDX_EVASAO",    "Índice de Evasão (%)",   COR_EVASAO))
    with c_camp3:
        chart_container(bar_campus("IDX_RETENCAO",  "Índice de Retenção (%)", COR_RETENCAO))

    # ============================================================
    # HEATMAP POR CAMPUS × CURSO
    # ============================================================

    secao("🔥 Mapa de Calor · Evasão por Campus e Curso")

    # v7: mesmo visual da versão original (grade completa, fundo cinza-claro,
    # um valor em cada célula), mas com a taxa correta: soma dos evadidos ÷
    # soma dos alunos do curso no campus (e não a média das taxas anuais).
    # Onde o curso não existe no campus, a célula fica cinza-claro com "–"
    # (a versão original mostrava 0, como se houvesse curso sem evasão).
    heat = (
        indicadores
        .assign(CAMPUS=indicadores["CAMPUS"].map(rotulo_campus))
        .groupby(["CAMPUS", "CURSO"])[["EVADIDOS", "TOTAL", "MIGRADOS"]]
        .sum()
        .reset_index()
    )
    heat["IDX_EVASAO"] = (heat["EVADIDOS"] / heat["TOTAL"] * 100).round(1)
    pivot_evasao = heat.pivot(index="CURSO", columns="CAMPUS", values="IDX_EVASAO")
    # Informações extras para o hover: total de alunos e migrados para a UFNT.
    extras_heat = [
        heat.pivot(index="CURSO", columns="CAMPUS", values=col)
        .reindex_like(pivot_evasao)
        .map(lambda v: "–" if pd.isna(v) else f"{int(v):,}")
        .to_numpy()
        for col in ["TOTAL", "MIGRADOS"]
    ]

    texto_heat = pivot_evasao.map(lambda v: "–" if pd.isna(v) else f"{v:g}")

    fig_heat = px.imshow(
        pivot_evasao.fillna(0),          # célula sem curso pinta com a cor de fundo
        aspect="auto",
        color_continuous_scale=[
            [0,   "#EEF2F8"],
            [0.3, "#F7C5BE"],
            [0.7, "#E07070"],
            [1,   "#C0392B"]
        ],
        height=max(400, len(pivot_evasao) * 22),
        labels=dict(x="Campus", y="CURSO", color="Evasão (%)")
    )
    fig_heat.update_traces(
        text=texto_heat.to_numpy(),
        texttemplate="%{text}",
        customdata=np.dstack(extras_heat),
        hovertemplate=(
            "Campus: %{x}<br>Curso: %{y}<br>Evasão (%): %{text}"
            "<br>Alunos: %{customdata[0]} · Migrados UFNT: %{customdata[1]}<extra></extra>"
        ),
    )
    fig_heat.update_layout(**LAYOUT_BASE)
    fig_heat.update_xaxes(showgrid=False)
    fig_heat.update_yaxes(showgrid=False, automargin=True)   # nomes inteiros
    chart_container(fig_heat)

    # ============================================================
    # TOP CURSOS: EVASÃO × SUCESSO
    # ============================================================

    secao("📋 Top Cursos — Evasão vs. Sucesso")

    tab_cursos = (
        df_filtrado
        .groupby(["COD_CURSO", "CURSO"])["SITUACAO_GRUPO_FINAL"]
        .value_counts()
        .unstack(fill_value=0)
        .assign(
            TOTAL=lambda d: d.sum(axis=1)
        )
    )
    for col in ["EVASAO", "SUCESSO", "RETIDO", "ATIVO", "MIGRADO"]:
        if col not in tab_cursos.columns:
            tab_cursos[col] = 0

    tab_cursos["IDX_EVASAO"]   = (tab_cursos["EVASAO"]  / tab_cursos["TOTAL"] * 100).round(1)
    tab_cursos["TAXA_SUCESSO"]  = (tab_cursos["SUCESSO"] / tab_cursos["TOTAL"] * 100).round(1)
    tab_cursos["IDX_RETENCAO"]  = (tab_cursos["RETIDO"]  / tab_cursos["TOTAL"] * 100).round(1)
    tab_cursos = tab_cursos.reset_index()

    # v7: mesmo mínimo do gráfico de dispersão, para um curso com poucos
    # alunos não liderar o ranking com 0% ou 100%.
    MIN_ALUNOS_RANKING = 20
    tab_ranking = tab_cursos[tab_cursos["TOTAL"] >= MIN_ALUNOS_RANKING]

    c_top1, c_top2 = st.columns(2)

    with c_top1:
        secao("⬆️ Maiores Índices de Evasão")
        top_evasao = tab_ranking.nlargest(10, "IDX_EVASAO")[
            ["COD_CURSO", "CURSO", "TOTAL", "EVASAO", "IDX_EVASAO"]
        ].rename(columns={
            "COD_CURSO": "Código do Curso",
            "TOTAL": "Total", "EVASAO": "Evadidos",
            "IDX_EVASAO": "Evasão (%)"
        })
        st.dataframe(top_evasao, use_container_width=True, hide_index=True)

    with c_top2:
        secao("⬆️ Maiores Taxas de Sucesso")
        top_sucesso = tab_ranking.nlargest(10, "TAXA_SUCESSO")[
            ["COD_CURSO", "CURSO", "TOTAL", "SUCESSO", "TAXA_SUCESSO"]
        ].rename(columns={
            "COD_CURSO": "Código do Curso",
            "TOTAL": "Total", "SUCESSO": "Formados",
            "TAXA_SUCESSO": "Sucesso (%)"
        })
        st.dataframe(top_sucesso, use_container_width=True, hide_index=True)

    # ============================================================
    # SCATTER: EVASÃO × SUCESSO × RETENÇÃO por CURSO
    # ============================================================

    secao("🎯 Dispersão: Evasão × Sucesso (por Curso)")

    scatter_df = tab_ranking.copy()   # mínimo 20 alunos

    fig_scatter = px.scatter(
        scatter_df,
        x="IDX_EVASAO",
        y="TAXA_SUCESSO",
        size="TOTAL",
        color="IDX_RETENCAO",
        hover_name="CURSO",
        hover_data={"COD_CURSO": True, "TOTAL": True, "IDX_RETENCAO": True},
        color_continuous_scale=[
            [0,   COR_SUCESSO],
            [0.5, COR_DESTAQUE],
            [1,   COR_EVASAO]
        ],
        labels={
            "IDX_EVASAO"   : "Índice de Evasão (%)",
            "TAXA_SUCESSO"  : "Taxa de Sucesso (%)",
            "IDX_RETENCAO"  : "Retenção (%)",
            "TOTAL"         : "Nº alunos",
            "COD_CURSO"     : "Código do Curso",
        },
        height=480,
    )
    fig_scatter.update_layout(**LAYOUT_BASE)
    chart_container(fig_scatter)

    # ============================================================
    # TABELA DE INDICADORES COMPLETA (COM DOWNLOAD)
    # ============================================================

    secao("📄 Tabela Completa de Indicadores (Campus × Curso × Ano)")

    colunas_exib = {
        "CAMPUS"       : "Campus",
        "COD_CURSO"    : "Código do Curso",
        "CURSO"        : "Curso",
        "ANO"          : "Ano Ingresso",
        "TOTAL"        : "Total",
        "FORMADOS"     : "Formados",
        "EVADIDOS"     : "Evadidos",
        "RETIDOS"      : "Retidos",
        "ATIVOS"       : "Ativos",
        "MIGRADOS"     : "Migrados UFNT",
        "TAXA_SUCESSO" : "Sucesso (%)",
        "IDX_EVASAO"   : "Evasão (%)",
        "IDX_RETENCAO" : "Retenção (%)",
    }
    tabela_final = indicadores.rename(columns=colunas_exib)[list(colunas_exib.values())]

    st.dataframe(
        tabela_final,
        use_container_width=True,
        height=460,
        hide_index=True
    )

    # ---- Download Excel ----
    # Gera um único arquivo .xlsx com duas abas:
    #   Aba 1 "Indicadores"  → tabela de indicadores calculados
    #   Aba 2 "Base Filtrada" → registros individuais filtrados

    @st.cache_data(show_spinner=False)
    def gerar_excel(df_indicadores: pd.DataFrame, df_base: pd.DataFrame) -> bytes:
        """
        Escreve dois DataFrames em abas distintas de um arquivo Excel
        e retorna os bytes prontos para download.

        Parâmetros
        ----------
        df_indicadores : DataFrame com os indicadores calculados
        df_base        : DataFrame com os registros individuais filtrados

        Retorna
        -------
        bytes do arquivo .xlsx
        """
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

            # --- Aba 1: Indicadores ---
            df_indicadores.to_excel(
                writer,
                sheet_name="Indicadores",
                index=False
            )

            ws_ind = writer.sheets["Indicadores"]

            # Largura automática das colunas
            for col_cells in ws_ind.columns:
                max_len = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in col_cells
                )
                ws_ind.column_dimensions[
                    col_cells[0].column_letter
                ].width = min(max_len + 4, 50)

            # --- Aba 2: Base Filtrada ---
            df_base_clean = df_base.drop(
                columns=["SITUACAO_GRUPO", "SITUACAO_GRUPO_FINAL", "CURSO_FILTRO"],
                errors="ignore"
            )
            df_base_clean.to_excel(
                writer,
                sheet_name="Base Filtrada",
                index=False
            )

            ws_base = writer.sheets["Base Filtrada"]

            for col_cells in ws_base.columns:
                max_len = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in col_cells
                )
                ws_base.column_dimensions[
                    col_cells[0].column_letter
                ].width = min(max_len + 4, 60)

        buffer.seek(0)
        return buffer.read()


    @st.cache_data(show_spinner=False)
    def gerar_excel_indicadores(df_indicadores: pd.DataFrame) -> bytes:
        """Gera a planilha resumida somente quando o download é solicitado."""
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_indicadores.to_excel(writer, sheet_name="Indicadores", index=False)
            ws = writer.sheets["Indicadores"]
            for col_cells in ws.columns:
                max_len = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in col_cells
                )
                ws.column_dimensions[col_cells[0].column_letter].width = min(
                    max_len + 4, 50
                )
        buffer.seek(0)
        return buffer.read()


    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        st.download_button(
            label="⬇️ Baixar Indicadores (.xlsx)",
            data=lambda: gerar_excel_indicadores(tabela_final),
            file_name="indicadores_uft.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click="ignore",
        )

    with col_dl2:
        st.download_button(
            label="⬇️ Baixar Relatório Completo (.xlsx)",
            data=lambda: gerar_excel(tabela_final, df_filtrado),
            file_name="relatorio_indicadores_uft.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Arquivo com duas abas: Indicadores e Base Filtrada",
            on_click="ignore",
        )

with aba_metodo:
    mostrar_metodologia(
        df_filtrado=df_filtrado,
        funil=funil_base,
        tolerancia=tolerancia,
        referencia=semestre_referencia(df_raw_final),
    )

# ============================================================
# RODAPÉ
# ============================================================

st.markdown(
    f"""
<hr>
<div style="
    text-align: center;
    font-size: 0.72rem;
    color: {COR_TEXTO_LEVE};
    padding: 8px 0 16px;
">
    Painel de Indicadores Acadêmicos · UFT · Dados protegidos por LGPD<br>
    Cálculos por coorte de ingresso (ANO_INGRESSO) ·
    Evasão = saída sem conclusão ·
    Retenção = ativo além do prazo do curso + {tolerancia} ano(s) ·
    Sucesso = Formado · Migrados UFNT = no total, fora de evasão e sucesso
</div>
""",
    unsafe_allow_html=True
)
