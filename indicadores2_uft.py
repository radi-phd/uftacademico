# ============================================================
# INDICADORES ACADÊMICOS · UFT
# Universidade Federal do Tocantins
#
# Calcula por CAMPUS · CURSO · ANO:
#   1. Índice de Evasão
#   2. Índice de Retenção
#   3. Taxa de Sucesso
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
import os
import re
import unicodedata
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Indicadores Acadêmicos · UFT",
    page_icon="📊",
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


def kpi_card(icone, valor, label, cor, subtitulo=""):
    sub = (
        f'<div style="font-size:0.65rem;color:{COR_TEXTO_LEVE};margin-top:2px;">'
        f'{subtitulo}</div>'
        if subtitulo else ""
    )
    return f"""
<div style="
    background: white;
    border: 1px solid {COR_BORDA};
    border-radius: 14px;
    padding: 16px 14px;
    flex: 1;
    min-width: 170px;
    position: relative;
">
<div style="
    position: absolute; top: 0; left: 0; right: 0;
    height: 4px; background: {cor};
    border-radius: 14px 14px 0 0;
"></div>
<div style="font-size: 1.4rem;">{icone}</div>
<div style="
    font-family: 'Lora', serif;
    font-size: 1.9rem;
    font-weight: 700;
    color: {COR_PRIMARIA};
">{valor}</div>
<div style="
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.07em;
    text-transform: uppercase;
    color: {COR_TEXTO_LEVE};
">{label}</div>
{sub}
</div>
"""


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
    for col in df.select_dtypes("object").columns:
        df[col] = df[col].str.strip()

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
            df[col] = pd.to_numeric(df[col], errors="coerce")

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
# EVASÃO: aluno saiu SEM concluir o curso de forma positiva.
#   Inclui: Desistência, Desvinculado, Matrícula Cancelada,
#           Jubilado, Declinante, Falecimento,
#           Transferência Interna/Externa/Ex-offício,
#           Reopção de Curso, Troca de Turno,
#           Transição UFNT, encerramentos especiais.
#
# RETENÇÃO: aluno ainda está vinculado / em curso, mas
#   ultrapassou a duração padrão prevista (detecção via
#   ANO_INGRESSO vs. ANO_EVASAO e DURAÇÃO-ANO).
#   Aqui usamos: situação = "Vinculado" (ativo, não formou),
#   e para coorte definimos retenção como alunos que
#   ingressaram há mais de (DURAÇÃO-ANO + 2) anos sem formar.
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
    "Declinante",
    "Falecimento",
    "Transferência Interna",
    "Transferência Externa",
    "Transferência Ex-offício",
    "Reopção de Curso",
    "Troca de Turno",
    "Transição UFNT",
    "Encerramento Intercâmbio Internacional",
    "Encerramento Mobilidade Acadêmica",
    "Encerramento Aluno Especial",
    "Encerramento de Convênio",
    "Encerramento/Apostilamento",
    "Habilitação",
    "Em análise",
}

SITUACOES_SUCESSO = {"Formado"}

SITUACOES_ATIVO  = {"Vinculado", "Reingresso Administrativo"}


def classifica_situacao(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona coluna SITUACAO_GRUPO com valores:
        EVASAO | SUCESSO | ATIVO | OUTRO
    """
    df = df.copy()

    cond_evasao  = df["FORMAEVASAO"].isin(SITUACOES_EVASAO)
    cond_sucesso = df["FORMAEVASAO"].isin(SITUACOES_SUCESSO)
    cond_ativo   = df["FORMAEVASAO"].isin(SITUACOES_ATIVO)

    df["SITUACAO_GRUPO"] = "OUTRO"
    df.loc[cond_ativo,   "SITUACAO_GRUPO"] = "ATIVO"
    df.loc[cond_evasao,  "SITUACAO_GRUPO"] = "EVASAO"
    df.loc[cond_sucesso, "SITUACAO_GRUPO"] = "SUCESSO"

    return df


def detecta_retencao(
    df: pd.DataFrame,
    anos_tolerancia: int = 2,
    ano_referencia: int | None = None,
) -> pd.DataFrame:
    """
    Marca como RETIDO alunos ATIVOS cujo tempo já cursado
    supera a duração prevista + tolerância.

    Parâmetros
    ----------
    df                : DataFrame com SITUACAO_GRUPO já definido
    anos_tolerancia   : anos extras além da duração mínima (padrão = 2)
    ano_referencia    : ano usado para vínculos ativos; por padrão, o ano atual

    Lógica
    ------
    Para alunos com SITUACAO_GRUPO == 'ATIVO':
        ano_referencia = ANO_EVASAO se disponível, senão ano atual
        tempo_cursado  = ano_referencia - ANO_INGRESSO
        se tempo_cursado > DURAÇÃO-ANO + tolerância → RETIDO
    """
    df = df.copy()
    ano_atual = ano_referencia or pd.Timestamp.today().year
    df["SITUACAO_GRUPO_FINAL"] = df["SITUACAO_GRUPO"]

    mascara_ativo = df["SITUACAO_GRUPO"] == "ATIVO"
    if mascara_ativo.sum() == 0:
        return df

    sub = df.loc[mascara_ativo].copy()

    ano_ref     = sub["ANO_EVASAO"].fillna(ano_atual)
    tempo       = ano_ref - sub["ANO_INGRESSO"]
    duracao_ref = sub["DURAÇÃO-ANO"].fillna(5)   # padrão 5 anos se não informado

    # Seleciona os índices originais dos alunos retidos
    retido_idx = sub.index[tempo > (duracao_ref + anos_tolerancia)]

    df.loc[retido_idx, "SITUACAO_GRUPO_FINAL"] = "RETIDO"

    return df


@st.cache_data(show_spinner=False)
def preparar_base(
    caminho: str,
    anos_tolerancia: int,
    versao_arquivo,
    ano_referencia: int,
) -> pd.DataFrame:
    """Carrega e classifica a base uma única vez por valor de tolerância."""
    df = carregar_ods(caminho, versao_arquivo)
    df = classifica_situacao(df)
    return detecta_retencao(
        df,
        anos_tolerancia=anos_tolerancia,
        ano_referencia=ano_referencia,
    )

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
        ATIVOS, IDX_EVASAO, IDX_RETENCAO, TAXA_SUCESSO
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
# INTERFACE · SIDEBAR
# ============================================================

st.sidebar.markdown(
    """
<div style="padding: 0.8rem 0 1.2rem;">
    <div style="font-family:'Lora',serif;font-size:1.1rem;font-weight:700;">
        📊 UFT · Indicadores
    </div>
    <div style="font-size:0.75rem;opacity:0.7;margin-top:0.2rem;">
        Evasão · Retenção · Sucesso
    </div>
</div>
""",
    unsafe_allow_html=True
)

tolerancia = st.sidebar.slider(
    "Tolerância p/ retenção (anos além da duração prevista)",
    min_value=0,
    max_value=5,
    value=2,
    help=(
        "Alunos ATIVOS com tempo cursado > duração prevista + "
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
<div style="font-size: 2.8rem;">📊</div>
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
DIRETORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
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
        pd.Timestamp.today().year,
    )

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

total_alunos = len(df_filtrado)
total_cursos = df_filtrado["COD_CURSO"].nunique()
total_campus = df_filtrado["CAMPUS"].nunique()

evadidos_total  = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "EVASAO").sum()
retidos_total   = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "RETIDO").sum()
formados_total  = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "SUCESSO").sum()
ativos_total    = (df_filtrado["SITUACAO_GRUPO_FINAL"] == "ATIVO").sum()

idx_evasao_geral   = round(evadidos_total / max(total_alunos, 1) * 100, 1)
idx_retencao_geral = round(retidos_total  / max(total_alunos, 1) * 100, 1)
taxa_sucesso_geral = round(formados_total / max(total_alunos, 1) * 100, 1)

# ============================================================
# KPIs GLOBAIS
# ============================================================

st.markdown(
    f"""
<div class="kpi-wrapper">
{kpi_card("👥", f"{total_alunos:,}", "TOTAL ALUNOS", COR_PRIMARIA)}
{kpi_card("🎓", f"{formados_total:,}", "FORMADOS", COR_SUCESSO, f"Sucesso: {taxa_sucesso_geral}%")}
{kpi_card("🚪", f"{evadidos_total:,}", "EVADIDOS", COR_EVASAO, f"Evasão: {idx_evasao_geral}%")}
{kpi_card("⏳", f"{retidos_total:,}", "RETIDOS", COR_RETENCAO, f"Retenção: {idx_retencao_geral}%")}
{kpi_card("📚", f"{ativos_total:,}", "ATIVOS (no prazo)", COR_ACENTO)}
{kpi_card("🏛️", f"{total_campus}", "CAMPUS", COR_PRIMARIA)}
{kpi_card("📋", f"{total_cursos}", "CURSOS", COR_SECUNDARIA)}
</div>
""",
    unsafe_allow_html=True
)

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

c_camp1, c_camp2, c_camp3 = st.columns(3)

def bar_campus(col_y, titulo, cor):
    df_ord = df_campus.sort_values(col_y, ascending=True)
    fig    = px.bar(
        df_ord, x=col_y, y="CAMPUS",
        orientation="h", text=col_y,
        height=300, color_discrete_sequence=[cor]
    )
    fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
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

pivot_evasao = (
    indicadores
    .groupby(["CAMPUS", "CURSO"])["IDX_EVASAO"]
    .mean()
    .round(1)
    .reset_index()
    .pivot(index="CURSO", columns="CAMPUS", values="IDX_EVASAO")
    .fillna(0)
)

fig_heat = px.imshow(
    pivot_evasao,
    text_auto=True,
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
fig_heat.update_layout(**LAYOUT_BASE)
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
for col in ["EVASAO", "SUCESSO", "RETIDO", "ATIVO"]:
    if col not in tab_cursos.columns:
        tab_cursos[col] = 0

tab_cursos["IDX_EVASAO"]   = (tab_cursos["EVASAO"]  / tab_cursos["TOTAL"] * 100).round(1)
tab_cursos["TAXA_SUCESSO"]  = (tab_cursos["SUCESSO"] / tab_cursos["TOTAL"] * 100).round(1)
tab_cursos["IDX_RETENCAO"]  = (tab_cursos["RETIDO"]  / tab_cursos["TOTAL"] * 100).round(1)
tab_cursos = tab_cursos.reset_index()

c_top1, c_top2 = st.columns(2)

with c_top1:
    secao("⬆️ Maiores Índices de Evasão")
    top_evasao = tab_cursos.nlargest(10, "IDX_EVASAO")[
        ["COD_CURSO", "CURSO", "TOTAL", "EVASAO", "IDX_EVASAO"]
    ].rename(columns={
        "COD_CURSO": "Código do Curso",
        "TOTAL": "Total", "EVASAO": "Evadidos",
        "IDX_EVASAO": "Evasão (%)"
    })
    st.dataframe(top_evasao, use_container_width=True, hide_index=True)

with c_top2:
    secao("⬆️ Maiores Taxas de Sucesso")
    top_sucesso = tab_cursos.nlargest(10, "TAXA_SUCESSO")[
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

scatter_df = tab_cursos[tab_cursos["TOTAL"] >= 20].copy()   # mínimo 20 alunos

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
    Retenção = ativo além do prazo + {tolerancia} ano(s) ·
    Sucesso = Formado
</div>
""",
    unsafe_allow_html=True
)
