# ============================================================
# INDICADORES ACADÊMICOS · UFT · VERSÃO 11 (arquivo único)
# Universidade Federal do Tocantins
#
# Calcula por CAMPUS · CURSO · ANO:
#   1. Índice de Evasão
#   2. Índice de Retenção
#   3. Taxa de Sucesso
#
# Execução:  streamlit run indicadores11_uft.py
# Na mesma pasta: 16.11.03a.ods e logo_uft.png (opcional).
#
# A v11 integra duas linhas do painel. A metodologia é a mesma da v7/v10.
#   • Da v10: base única compartilhada (st.cache_resource), leitura enxuta
#     do ODS, retenção aplicada só no filtro, resultados em cache por
#     filtro, Excel com xlsxwriter, barra lateral nova (campus em botões,
#     busca de cursos, "Limpar"), resumo dos filtros e rodapé institucional.
#   • Da versão modular (v7 modular): avisos inteligentes, momento da
#     evasão, formas de evasão e perfil por turno.
#   • Novos na v11: comparação com a UFT nos cartões, faixa das turmas em
#     andamento na evolução, curva acumulada no momento da evasão, pico do
#     momento sem a barra "13º ou mais", cartões dos gráficos sem a caixa
#     vazia, rankings com barra de percentual.
#
# Este arquivo é gerado a partir do projeto modular (app.py + src/), em
# que cada seção abaixo é um arquivo. Os títulos das seções seguem os
# nomes dos módulos de lá.
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

import os
import re
import unicodedata
import pandas as pd
import base64
import plotly.graph_objects as go
from html import escape
import streamlit as st
import xml.sax
import zipfile
import numpy as np
from dataclasses import dataclass
import io
import xlsxwriter
import plotly.express as px


# ============================================================
# CONFIGURAÇÃO CENTRAL · caminhos, paleta e regras de negócio
# ============================================================
# Tudo o que é "valor ajustável" do painel fica aqui. Para mudar uma cor,
# um prazo oficial (PPC) ou uma regra de classificação, edite este arquivo.


# ------------------------------------------------------------
# CAMINHOS
# ------------------------------------------------------------
# Pasta do script: o ODS e a logo ficam junto dele.
RAIZ_PROJETO = os.path.dirname(os.path.abspath(__file__))

# Logo da UFT (logo_uft.png). Se não existir, o painel usa o emoji 📊.
CAMINHO_LOGO = os.path.join(RAIZ_PROJETO, "logo_uft.png")

# Relatório 16.11.03a (ODS), na mesma pasta do script.
CAMINHO_ODS = os.path.join(RAIZ_PROJETO, "16.11.03a.ods")

# ------------------------------------------------------------
# PÁGINA
# ------------------------------------------------------------
TITULO_PAGINA = "Indicadores Acadêmicos · UFT · v11"

# ------------------------------------------------------------
# PALETA INSTITUCIONAL
# (as cores também viram variáveis CSS: --cor-primaria, --cor-fundo, ...)
# ------------------------------------------------------------
COR_PRIMARIA   = "#28476b"
COR_SECUNDARIA = "#2b507a"
COR_ACENTO     = "#7491b3"
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

# ------------------------------------------------------------
# CURSOS
# ------------------------------------------------------------
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

# Campi reconhecidos da UFT: linhas de outros campi são descartadas na leitura.
CAMPUS_UFT = {
    "Araguaína", "Arraias", "Gurupi",
    "Miracema", "Palmas", "Porto Nacional", "Tocantinópolis"
}

# ------------------------------------------------------------
# CLASSIFICAÇÃO DE SITUAÇÕES
# ------------------------------------------------------------

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

# Uma ENTRADA = o mesmo aluno (matrícula), no mesmo curso e campus,
# com o mesmo ano e semestre de ingresso. Todas as entradas são contadas:
# matrícula em cursos diferentes e reingresso em outro ano/semestre
# são entradas diferentes.
CHAVE_ENTRADA = ["MATRICULA", "COD_CURSO", "CAMPUS", "ANO_INGRESSO", "PERIODO_INGRESSO"]

# Desempate quando a MESMA entrada aparece em mais de uma linha
# (menor número = maior prioridade).
PRIORIDADE_GRUPO = {"SUCESSO": 0, "ATIVO": 1, "MIGRADO": 2, "EVASAO": 3, "OUTRO": 4}

# ------------------------------------------------------------
# RANKINGS E GRÁFICOS
# ------------------------------------------------------------
# v7: mesmo mínimo do gráfico de dispersão, para um curso com poucos
# alunos não liderar o ranking com 0% ou 100%.
MIN_ALUNOS_RANKING = 20

# ------------------------------------------------------------
# ANÁLISES DA EVASÃO (momento, forma e perfil)
# ------------------------------------------------------------
# Cor das barras que NÃO são o destaque (a maior fica em COR_EVASAO).
COR_EVASAO_CLARA = "#F7C5BE"

# Momento: períodos mostrados individualmente; acima disso vira "13º ou mais".
PERIODO_MAX_GRAFICO = 12

# Forma: quantas formas aparecem separadas; o restante é somado em "Outros".
MAX_FORMAS_EVASAO = 6

# Turno: nome da coluna no ODS (a primeira que existir é usada). Se o seu
# relatório tiver o turno em outra coluna, acrescente o nome aqui.
COLUNAS_TURNO = ["TURNO", "TURNO_CURSO"]

# Sem coluna de turno, o turno é deduzido do 3º caractere do código do curso
# (ex.: 44M601P → M = Matutino; 44N600NC → N = Noturno). Confira com a base real.
TURNO_POR_LETRA_DO_CODIGO = {
    "M": "Matutino",
    "V": "Vespertino",
    "N": "Noturno",
    "I": "Integral",
}
ORDEM_TURNOS = ["Matutino", "Vespertino", "Noturno", "Integral"]
TURNO_NAO_IDENTIFICADO = "Não identificado"

# ------------------------------------------------------------
# LEITURA DO ODS · base enxuta (v10)
# ------------------------------------------------------------
# Só estas colunas ficam na memória. As demais (ETNIA, DT_NASCIMENTO,
# NOME, CPF…) são descartadas já na leitura, linha a linha.
COLUNAS_OBRIGATORIAS = [
    "MATRICULA", "COD_CURSO", "CURSO", "CAMPUS",
    "ANO_INGRESSO", "PERIODO_INGRESSO", "FORMAINGRESSO",
    "FORMAEVASAO", "ANO_EVASAO", "PERIODO_EVASAO", "DURAÇÃO-SEM",
]
COLUNAS_OPCIONAIS = COLUNAS_TURNO + ["DURAÇÃO-ANO"]

# Dados pessoais: nunca entram na base nem na comparação de cópias (LGPD).
COLUNAS_PESSOAIS = {"NOME", "CPF", "CPF_NUM"}

# Textos repetidos guardados como "category" (ocupam bem menos memória).
COLUNAS_CATEGORIA = [
    "CAMPUS", "COD_CURSO", "CURSO", "FORMAINGRESSO", "FORMAEVASAO",
    "PERIODO_INGRESSO", "PERIODO_EVASAO", "PRAZO_FONTE", "SITUACAO_GRUPO",
] + COLUNAS_TURNO

# ------------------------------------------------------------
# FILTROS E CACHE (v10)
# ------------------------------------------------------------
TOLERANCIA_PADRAO = 2          # anos além do prazo do curso
TOLERANCIA_MAXIMA = 5

# Resultados por combinação de filtros, compartilhados entre usuários.
CACHE_FILTROS_MAX = 64         # combinações guardadas
CACHE_FILTROS_TTL = 3600       # segundos

# ------------------------------------------------------------
# APRIMORAMENTOS v11
# ------------------------------------------------------------
# Evolução: turmas com pelo menos este % de alunos ainda matriculados
# (ativos + retidos) ganham uma faixa cinza: os índices ainda vão mudar.
LIMIAR_TURMA_EM_ANDAMENTO = 50

# ------------------------------------------------------------
# EXPORTAÇÃO (aba "Base Filtrada" do Excel)
# ------------------------------------------------------------
# Sem MATRICULA, ETNIA nem DT_NASCIMENTO: o painel é público (LGPD).
COLUNAS_EXPORTACAO = {
    "CAMPUS": "CAMPUS",
    "COD_CURSO": "COD_CURSO",
    "CURSO": "CURSO",
    "TURNO": "TURNO",
    "ANO_INGRESSO": "ANO_INGRESSO",
    "PERIODO_INGRESSO": "PERIODO_INGRESSO",
    "FORMAINGRESSO": "FORMAINGRESSO",
    "FORMAEVASAO": "FORMAEVASAO",
    "ANO_EVASAO": "ANO_EVASAO",
    "PERIODO_EVASAO": "PERIODO_EVASAO",
    "DURAÇÃO-SEM": "DURAÇÃO-SEM",
    "PRAZO_SEM": "PRAZO_SEM",
    "PRAZO_FONTE": "PRAZO_FONTE",
    "TEMPO_SEM": "TEMPO_SEM",
    "SITUACAO_GRUPO_FINAL": "SITUACAO_NO_PAINEL",
}


# ============================================================
# RÓTULOS · funções puras de texto para cursos e campi
# ============================================================
# Sem dependência do Streamlit: podem ser testadas isoladamente.


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


def chave_ordenacao_texto(texto):
    """Ordena alfabeticamente, ignorando acentos e pontuação."""
    sem_acentos = "".join(
        caractere
        for caractere in unicodedata.normalize("NFKD", str(texto))
        if not unicodedata.combining(caractere)
    ).casefold()
    return re.sub(r"[^a-z0-9]+", " ", sem_acentos).strip()


def rotulo_campus(campus: str) -> str:
    """Nome do campus para os gráficos, sinalizando os que viraram UFNT."""
    return f"{campus} (UFNT {ANO_TRANSICAO_UFNT})" if campus in CAMPI_UFNT else campus


# ============================================================
# LOGO DA UFT
# ============================================================
# Arquivo esperado: logo_uft.png, na mesma pasta do script.
# Se não existir, o painel usa o emoji 📊 (nada quebra).


LOGO_B64 = None
if os.path.isfile(CAMINHO_LOGO):
    with open(CAMINHO_LOGO, "rb") as arquivo_logo:
        LOGO_B64 = base64.b64encode(arquivo_logo.read()).decode()


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


# ============================================================
# TEMA DOS GRÁFICOS PLOTLY
# ============================================================
# Layout padrão aplicado a todas as figuras: fig.update_layout(**LAYOUT_BASE)


LAYOUT_BASE = dict(
    paper_bgcolor="white",
    plot_bgcolor="white",
    font=dict(family="DM Sans", color=COR_TEXTO),
    margin=dict(t=30, b=20, l=10, r=10)
)


def figura_barras_destaque(
    rotulos, valores, percentuais=None, horizontal=False,
    altura=340, titulo_x="", titulo_y="Evasões", fora_do_destaque=(),
):
    """
    Barras de evasão: a maior em COR_EVASAO e as demais em COR_EVASAO_CLARA,
    para o pico saltar aos olhos. `percentuais` aparece no hover.
    Em `horizontal=True` a primeira categoria fica no topo.
    `fora_do_destaque`: rótulos que nunca são o pico (ex.: "13º ou mais",
    que soma vários períodos).
    """
    rotulos, valores = list(rotulos), [int(v) for v in valores]
    maximo = max(valores) if valores else 0
    disputam = [v for r, v in zip(rotulos, valores) if r not in fora_do_destaque]
    pico = max(disputam) if disputam else maximo
    cores = [
        COR_EVASAO if v == pico and v > 0 and r not in fora_do_destaque else COR_EVASAO_CLARA
        for r, v in zip(rotulos, valores)
    ]
    detalhe = list(percentuais) if percentuais is not None else [0.0] * len(valores)
    x, y = (valores, rotulos) if horizontal else (rotulos, valores)

    fig = go.Figure(go.Bar(
        x=x, y=y, orientation="h" if horizontal else "v",
        marker_color=cores, text=[f"{v:,}" for v in valores],
        textposition="outside", cliponaxis=False, customdata=detalhe,
        hovertemplate=(
            "%{y}<br>" if horizontal else "%{x}<br>"
        ) + ("%{x:,}" if horizontal else "%{y:,}") + " evasões (%{customdata:.1f}%)<extra></extra>",
    ))
    fig.update_layout(**LAYOUT_BASE, height=altura, showlegend=False)
    folga = [0, maximo * 1.18 if maximo else 1]
    if horizontal:
        fig.update_xaxes(range=folga, title=titulo_y)
        fig.update_yaxes(autorange="reversed", title=titulo_x, type="category")
    else:
        fig.update_yaxes(range=folga, title=titulo_y)
        fig.update_xaxes(title=titulo_x, type="category")
    return fig


# ============================================================
# COMPONENTES VISUAIS REUTILIZÁVEIS
# ============================================================


PLOTLY_CONFIG = {"displaylogo": False}


def chart_container(fig):
    """
    Gráfico dentro de um cartão com borda.
    v11: container nativo. A versão anterior abria um <div> num markdown e o
    fechava em outro; o Streamlit não aninha HTML entre chamadas, então
    sobrava um cartão branco vazio acima de cada gráfico.
    """
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)


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


def comparacao_uft(valor, referencia, maior_e_melhor) -> str:
    """
    v11: linha "▲ 2.3 p.p. vs UFT (35.0%)" para os cartões do topo.
    Verde quando o filtro está melhor que a UFT no mesmo período, vermelho
    quando está pior (para evasão e retenção, menor é melhor).
    """
    diferenca = round(valor - referencia, 1)
    if diferenca == 0:
        return f'<div class="kpi-comparacao">= UFT ({referencia}%)</div>'
    melhor = (diferenca > 0) == maior_e_melhor
    seta = "▲" if diferenca > 0 else "▼"
    return (
        f'<div class="kpi-comparacao"><span class="{"melhor" if melhor else "pior"}">'
        f'{seta} {abs(diferenca)} p.p.</span> vs UFT ({referencia}%)</div>'
    )


def kpi_card(icone, valor, label, cor, subtitulo="", ajuda="", comparacao=""):
    """
    Cartão de indicador (mesmo visual da versão original).
    `ajuda`: texto explicativo exibido ao passar o mouse (marcado com ⓘ).
    Linhas do texto de ajuda podem ser separadas por "\n".
    `comparacao`: HTML de comparacao_uft(), exibido abaixo do subtítulo.
    """
    sub = (
        f'<div style="font-size:0.65rem;color:{COR_TEXTO_LEVE};margin-top:2px;">{subtitulo}</div>'
        if subtitulo else ""
    ) + comparacao
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


def caixa_explicativa(html: str) -> None:
    """Caixa branca de texto explicativo (HTML em uma linha, sem indentação)."""
    st.markdown(f'<div class="explica">{html}</div>', unsafe_allow_html=True)


# ============================================================
# ESTILOS · CSS do painel (no projeto modular: assets/css/*.css)
# ============================================================
# As cores da paleta ficam disponíveis no CSS como variáveis:
#   var(--cor-primaria), var(--cor-fundo), var(--cor-borda) ...

_VARIAVEIS_CSS = {
    "--cor-primaria": COR_PRIMARIA,
    "--cor-secundaria": COR_SECUNDARIA,
    "--cor-acento": COR_ACENTO,
    "--cor-destaque": COR_DESTAQUE,
    "--cor-fundo": COR_FUNDO,
    "--cor-card": COR_CARD,
    "--cor-borda": COR_BORDA,
    "--cor-texto": COR_TEXTO,
    "--cor-texto-leve": COR_TEXTO_LEVE,
    "--cor-alerta": COR_ALERTA,
    "--cor-sucesso": COR_SUCESSO,
    "--cor-evasao": COR_EVASAO,
    "--cor-retencao": COR_RETENCAO,
}

CSS_FONTES = '''
/* Fontes do painel. O @import precisa vir antes de qualquer outra regra. */
@import url('https://fonts.googleapis.com/css2?family=Lora:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap');
'''

CSS_BASE = '''
/* Estilo geral da página. As cores vêm das variáveis (:root) geradas
   a partir da paleta definida em src/config.py. */
html, body {
    font-family: 'DM Sans', sans-serif;
    color: var(--cor-texto);
}
.main { background-color: var(--cor-fundo); padding-top: 0.5rem; }
hr { border: none; border-top: 1px solid var(--cor-borda); margin: 20px 0; }
'''

CSS_SIDEBAR = '''
/* Barra lateral (v10: campus em botões, busca de cursos, passos numerados) */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, var(--cor-primaria) 0%, var(--cor-secundaria) 100%);
}
[data-testid="stSidebar"] * { color: white !important; }
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {
    font-size: 0.95rem;
    font-weight: 700;
}
[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background-color: white !important;
    border-color: rgba(255,255,255,0.55) !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] *,
[data-testid="stSidebar"] [data-baseweb="select"] svg {
    color: var(--cor-texto) !important;
    fill: var(--cor-texto) !important;
}
/* Etiquetas dos cursos escolhidos: cor escura para o texto branco ter contraste */
[data-testid="stSidebar"] [data-baseweb="tag"] {
    background-color: var(--cor-secundaria) !important;
}
[data-testid="stSidebar"] [data-baseweb="tag"] *,
[data-testid="stSidebar"] [data-baseweb="tag"] svg {
    color: white !important;
    fill: white !important;
}
/* Caixa de busca de cursos (st.multiselect nas versões novas do Streamlit) */
[data-testid="stSidebar"] [data-testid="stMultiSelect"] [role="group"] {
    background-color: white !important;
}
[data-testid="stSidebar"] [data-testid="stMultiSelect"] [role="group"] *:not([data-tag]):not([data-tag] *) {
    color: var(--cor-texto) !important;   /* ícones usam currentColor */
}
[data-testid="stSidebar"] [data-testid="stMultiSelect"] input::placeholder {
    color: var(--cor-texto-leve) !important;
    opacity: 1;
}
[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-tag] {
    background-color: var(--cor-secundaria) !important;
}
[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-tag] * {
    color: white !important;
}
/* Botões de campus (st.pills: "stButtonGroup" nas versões novas, "stPills" nas antigas) */
[data-testid="stSidebar"] [data-testid="stButtonGroup"] button,
[data-testid="stSidebar"] [data-testid="stPills"] button {
    background-color: rgba(255,255,255,0.10) !important;
    border: 1px solid rgba(255,255,255,0.45) !important;
}
[data-testid="stSidebar"] [data-testid="stButtonGroup"] button:hover,
[data-testid="stSidebar"] [data-testid="stPills"] button:hover {
    background-color: rgba(255,255,255,0.22) !important;
}
[data-testid="stSidebar"] [data-testid="stButtonGroup"] button *,
[data-testid="stSidebar"] [data-testid="stPills"] button * {
    color: white !important;
}
[data-testid="stSidebar"] [data-testid="stButtonGroup"] button[aria-pressed="true"],
[data-testid="stSidebar"] [data-testid="stPills"] button[aria-pressed="true"],
[data-testid="stSidebar"] [data-testid="stPills"] button[kind="pillsActive"] {
    background-color: var(--cor-destaque) !important;
    border-color: var(--cor-destaque) !important;
}
[data-testid="stSidebar"] [data-testid="stButtonGroup"] button[aria-pressed="true"] *,
[data-testid="stSidebar"] [data-testid="stPills"] button[aria-pressed="true"] *,
[data-testid="stSidebar"] [data-testid="stPills"] button[kind="pillsActive"] * {
    color: var(--cor-texto) !important;
    font-weight: 700;
}
[data-testid="stSidebar"] .stButton button {
    background-color: rgba(255,255,255,0.12) !important;
    border: 1px solid rgba(255,255,255,0.55) !important;
}
[data-testid="stSidebar"] .stButton button:hover {
    background-color: rgba(255,255,255,0.24) !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] details {
    border-color: rgba(255,255,255,0.35) !important;
}
.filtro-passo {
    display: flex; align-items: center; gap: 8px;
    margin: 18px 0 2px; font-size: 0.72rem; font-weight: 700;
    letter-spacing: 0.08em; text-transform: uppercase; opacity: 0.75;
}
.filtro-passo span {
    display: inline-flex; align-items: center; justify-content: center;
    width: 20px; height: 20px; border-radius: 50%;
    background: var(--cor-destaque); color: var(--cor-texto) !important;
    font-size: 0.7rem;
}
'''

CSS_COMPONENTES = '''
/* Cartões, KPIs e caixas de texto reutilizáveis */
/* Cartão dos gráficos: st.container(border=True) em chart_container() */
[data-testid="stVerticalBlockBorderWrapper"]:has(.stPlotlyChart) {
    background: var(--cor-card);
    border-color: var(--cor-borda) !important;
    border-radius: 14px;
    box-shadow: 0 2px 10px rgba(13,43,78,0.05);
}
.kpi-wrapper {
    display: flex;
    gap: 10px;
    margin-bottom: 20px;
    flex-wrap: wrap;
}
.explica {
    background: var(--cor-card); border: 1px solid var(--cor-borda); border-radius: 14px;
    padding: 16px 20px; margin-bottom: 14px; line-height: 1.6; font-size: 0.93rem;
}
.explica b { color: var(--cor-primaria); }

/* Avisos inteligentes (insights automáticos) */
.insights-wrapper {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(270px, 1fr));
    gap: 10px;
    margin-bottom: 8px;
}
.insight {
    display: flex; gap: 12px; align-items: flex-start;
    background: var(--cor-card); border: 1px solid var(--cor-borda);
    border-left: 6px solid var(--cor-destaque); border-radius: 12px;
    padding: 14px 16px; line-height: 1.5; font-size: 0.92rem;
}
.insight-icone { font-size: 1.5rem; line-height: 1; }
.insight b { color: var(--cor-primaria); }

/* Resumo dos filtros ativos, logo abaixo do cabeçalho (v10) */
.resumo-filtros {
    display: flex; flex-wrap: wrap; gap: 8px; align-items: center;
    margin: -8px 0 18px; font-size: 0.82rem; color: var(--cor-texto-leve);
}
.resumo-filtros .chip {
    background: white; border: 1px solid var(--cor-borda); border-radius: 999px;
    padding: 3px 12px; color: var(--cor-primaria); font-weight: 600;
}

/* Comparação com a UFT nos cartões do topo (v11) */
.kpi-comparacao { font-size: 0.68rem; color: var(--cor-texto-leve); margin-top: 3px; }
.kpi-comparacao .melhor { color: var(--cor-sucesso); font-weight: 700; }
.kpi-comparacao .pior { color: var(--cor-evasao); font-weight: 700; }

/* Glossário em tabela HTML: textos longos quebram a linha (v10) */
.glossario { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
.glossario th {
    text-align: left; color: var(--cor-primaria);
    border-bottom: 2px solid var(--cor-destaque); padding: 7px 10px;
}
.glossario td { padding: 7px 10px; border-bottom: 1px solid var(--cor-borda); vertical-align: top; }
.glossario td:first-child { white-space: nowrap; }

/* Rodapé institucional (v10) */
.rodape {
    margin-top: 36px; padding: 18px 0 22px; border-top: 2px solid var(--cor-destaque);
    text-align: center; line-height: 1.7;
}
.rodape-autoria { font-size: 0.85rem; color: var(--cor-primaria); }
.rodape-fonte { font-size: 0.72rem; color: var(--cor-texto-leve); margin-top: 6px; }
'''

CSS_ABAS = '''
/* Abas "Painel" e "Entenda os Cálculos" na paleta institucional */
[data-testid="stTabs"] [role="tablist"] {
    gap: 10px; border-bottom: 3px solid var(--cor-destaque); margin-bottom: 20px;
}
[data-testid="stTab"] {
    padding: 12px 26px !important; border-radius: 12px 12px 0 0;
    background: white; border: 1px solid var(--cor-borda); border-bottom: none;
}
[data-testid="stTab"] p { font-size: 1rem !important; font-weight: 700; color: var(--cor-secundaria) !important; }
[data-testid="stTab"]:hover { background: var(--cor-fundo); }
[data-testid="stTab"][aria-selected="true"] {
    background: linear-gradient(135deg, var(--cor-primaria) 0%, var(--cor-secundaria) 100%);
    border-color: var(--cor-primaria);
}
[data-testid="stTab"][aria-selected="true"] p { color: white !important; }
[data-testid="stTab"] .react-aria-SelectionIndicator { display: none; }
'''

def montar_css() -> str:
    """Junta o CSS e as variáveis de cor. O @import das fontes vem primeiro."""
    variaveis = ":root {\n" + "".join(f"    {k}: {v};\n" for k, v in _VARIAVEIS_CSS.items()) + "}\n"
    return (CSS_FONTES + variaveis + CSS_BASE + CSS_SIDEBAR
            + CSS_COMPONENTES + CSS_ABAS)


def aplicar_estilos() -> None:
    """Injeta o CSS global do painel."""
    st.markdown(f"\n<style>\n{montar_css()}</style>\n", unsafe_allow_html=True)


# ============================================================
# LAYOUT GERAL · página, marca, cabeçalho, resumo dos filtros e rodapé
# ============================================================


def configurar_pagina() -> None:
    """Deve ser o PRIMEIRO comando Streamlit executado pelo app."""
    st.set_page_config(
        page_title=TITULO_PAGINA,
        page_icon=CAMINHO_LOGO if LOGO_B64 else "📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )


def marca_sidebar() -> None:
    """Logo e título no topo da barra lateral."""
    st.sidebar.markdown(
        '<div style="padding:0.8rem 0 0.6rem;display:flex;align-items:center;gap:12px;">'
        + selo_logo(46)
        + '<div><div style="font-family:\'Lora\',serif;font-size:1.1rem;font-weight:700;">UFT · Indicadores</div>'
        '<div style="font-size:0.75rem;opacity:0.7;margin-top:0.2rem;">Evasão · Retenção · Sucesso</div></div>'
        '</div>',
        unsafe_allow_html=True
    )


def cabecalho() -> None:
    """Faixa azul com o título do painel."""
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


def resumo_filtros(filtros, base) -> None:
    """Chips com os filtros ativos: o usuário vê o que está olhando sem abrir a barra."""
    if not filtros.campi:
        txt_campus = "Todos os campi"
    elif len(filtros.campi) <= 3:
        txt_campus = ", ".join(filtros.campi)
    else:
        txt_campus = f"{len(filtros.campi)} campi"
    if not filtros.cursos:
        txt_curso = "Todos os cursos"
    elif len(filtros.cursos) == 1:
        txt_curso = base.nome_curso(filtros.cursos[0])
    else:
        txt_curso = f"{len(filtros.cursos)} cursos"
    st.markdown(
        '<div class="resumo-filtros">Mostrando:'
        f'<span class="chip">🏛️ {escape(txt_campus)}</span>'
        f'<span class="chip">📋 {escape(txt_curso)}</span>'
        f'<span class="chip">📅 {filtros.anos[0]}–{filtros.anos[1]}</span>'
        f'<span class="chip">⏳ tolerância {filtros.tolerancia} ano(s)</span></div>',
        unsafe_allow_html=True,
    )


def rodape(referencia: tuple) -> None:
    """Autoria institucional e fonte dos dados (definições: aba 📘)."""
    ano_ref, sem_ref = referencia
    st.markdown(
        '<div class="rodape">'
        '<div class="rodape-autoria">'
        'Elaborado pela <b>Coordenação de Avaliação e Educação (CAE)</b><br>'
        'Pró-Reitoria de Graduação (PROGRAD) · Universidade Federal do Tocantins</div>'
        f'<div class="rodape-fonte">Fonte: relatório 16.11.03a, com dados até {ano_ref}.{sem_ref} · '
        'Dados protegidos pela LGPD</div>'
        '</div>',
        unsafe_allow_html=True
    )


# ============================================================
# LEITURA DO ARQUIVO ODS COM PARSER SAX
# (Evita carregar >500 MB de XML diretamente em memória)
# ============================================================
# v10: guarda só as colunas usadas nos cálculos (COLUNAS_OBRIGATORIAS e
# COLUNAS_OPCIONAIS) e uma "impressão digital" de cada linha inteira, sem
# dados pessoais, para achar cópias exatas sem manter tudo na memória.


class ODSHandler(xml.sax.ContentHandler):
    """
    Lê o content.xml de um arquivo ODS linha a linha via SAX.

    A 1ª linha é o cabeçalho. Das demais, guarda apenas as colunas
    desejadas (self.rows) e uma impressão digital da linha inteira sem
    dados pessoais (self.hashes).
    """

    def __init__(self, colunas_desejadas):
        super().__init__()
        self.colunas_desejadas = colunas_desejadas
        self.in_text = False
        self.current_row = []
        self.current_text = ""
        self.repeat_count = 1
        self.header = None
        self.colunas = []
        self.idx_guardar = []
        self.idx_hash = []
        self.min_len = 0
        self.rows = []
        self.hashes = []

    # ---- SAX callbacks ----

    def startElement(self, name, attrs):
        if name == "table:table-row":
            self.current_row = []

        elif name == "table:table-cell":
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
            self.current_row.extend([self.current_text] * self.repeat_count)
            self.current_text = ""
            self.repeat_count = 1

        elif name == "table:table-row":
            row = self.current_row
            self.current_row = []
            # Remove células vazias à direita
            while row and row[-1] == "":
                row.pop()
            if not row:
                return
            if self.header is None:
                self._definir_cabecalho(row)
                return
            if len(row) < self.min_len:        # descarta linhas muito curtas
                return
            n = len(self.header)
            row = [c.strip() for c in (row + [""] * n)[:n]]
            self.rows.append(tuple(row[i] for i in self.idx_guardar))
            self.hashes.append(hash(tuple(row[i] for i in self.idx_hash)))

    def _definir_cabecalho(self, header):
        # Renomeia colunas duplicadas (COL, COL.1, COL.2…)
        vistos = {}
        nomes = []
        for c in (h.strip() for h in header):
            if c in vistos:
                vistos[c] += 1
                nomes.append(f"{c}.{vistos[c]}")
            else:
                vistos[c] = 0
                nomes.append(c)
        self.header = nomes
        self.min_len = max(5, len(nomes) // 3)
        self.colunas = [c for c in nomes if c in self.colunas_desejadas]
        self.idx_guardar = [nomes.index(c) for c in self.colunas]
        self.idx_hash = [i for i, c in enumerate(nomes) if c not in COLUNAS_PESSOAIS]


def carregar_ods(caminho: str) -> tuple[pd.DataFrame, pd.Series]:
    """
    Abre o arquivo ODS e retorna (DataFrame enxuto, impressão digital de
    cada linha), já restrito aos campi da UFT e com os tipos convertidos.
    Sem cache próprio: quem guarda o resultado é src.dados.base.
    """
    with zipfile.ZipFile(caminho, "r") as zf:
        with zf.open("content.xml") as xml_file:
            handler = ODSHandler(set(COLUNAS_OBRIGATORIAS + COLUNAS_OPCIONAIS))
            parser = xml.sax.make_parser()
            parser.setContentHandler(handler)
            parser.parse(xml_file)

    if not handler.rows:
        raise ValueError("Nenhuma linha encontrada no arquivo ODS.")

    colunas_ausentes = sorted(set(COLUNAS_OBRIGATORIAS) - set(handler.colunas))
    if colunas_ausentes:
        raise ValueError(
            "Colunas obrigatórias ausentes no ODS: " + ", ".join(colunas_ausentes)
        )

    df = pd.DataFrame.from_records(handler.rows, columns=handler.colunas)
    impressao = pd.Series(handler.hashes, dtype="int64")
    del handler

    # Mantém apenas campus válidos da UFT
    manter = df["CAMPUS"].isin(CAMPUS_UFT).to_numpy()
    df = df.loc[manter].reset_index(drop=True)
    impressao = impressao.loc[manter].reset_index(drop=True)

    # Conversão de tipos (aceita vírgula decimal: "4,5")
    for col in ["ANO_INGRESSO", "ANO_EVASAO", "DURAÇÃO-SEM", "DURAÇÃO-ANO"]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", ".", regex=False),
                errors="coerce",
            )

    # Corrige nomes incompletos e elimina a cidade repetida no nome.
    # Calculado uma vez por combinação distinta (e não por linha).
    trios = df[["COD_CURSO", "CURSO", "CAMPUS"]].drop_duplicates()
    trios["CURSO_NOVO"] = [
        nome_curso_sem_campus(codigo, curso, campus)
        for codigo, curso, campus in trios.itertuples(index=False)
    ]
    df["CURSO"] = df.merge(
        trios, on=["COD_CURSO", "CURSO", "CAMPUS"], how="left"
    )["CURSO_NOVO"].to_numpy()

    return df, impressao


# ============================================================
# CLASSIFICAÇÃO DE SITUAÇÕES + DEDUPLICAÇÃO SEGURA (v7)
# ============================================================
# Uma ENTRADA = o mesmo aluno (matrícula), no mesmo curso e campus,
# com o mesmo ano e semestre de ingresso. Todas as entradas são contadas.


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


def consolidar_entradas(df: pd.DataFrame, impressao: pd.Series):
    """
    1. remove cópias exatas (todas as colunas iguais, exceto dados pessoais,
       comparadas pela impressão digital lida do ODS — ver leitor_ods);
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

    df = df.loc[~impressao.duplicated().to_numpy()]             # 1
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


# ============================================================
# PRAZOS DOS CURSOS E DETECÇÃO DE RETENÇÃO
# ============================================================


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


def tempo_no_curso(df: pd.DataFrame, referencia: tuple[int, int]) -> pd.Series:
    """
    Semestres desde o ingresso até o semestre de referência, contando o
    semestre de entrada:

        tempo = 2 × (ano_ref − ANO_INGRESSO) + (sem_ref − PERIODO_INGRESSO) + 1

    Não depende da tolerância, por isso é calculado uma vez na base.
    """
    ano_ref, sem_ref = referencia
    per_ing = pd.to_numeric(df["PERIODO_INGRESSO"], errors="coerce").fillna(1)
    return 2 * (ano_ref - df["ANO_INGRESSO"]) + (sem_ref - per_ing) + 1


def situacao_final(df: pd.DataFrame, anos_tolerancia: int) -> np.ndarray:
    """
    SITUACAO_GRUPO com a retenção aplicada (v10: só nas linhas do filtro,
    uma comparação vetorizada, sem recarregar a base):

        RETIDO se ATIVO e TEMPO_SEM > PRAZO_SEM + 2 × anos_tolerancia

    Ativos de cursos sem prazo (PRAZO_SEM vazio) continuam ATIVO.
    """
    retido = (
        df["SITUACAO_GRUPO"].eq("ATIVO")
        & df["PRAZO_SEM"].notna()
        & df["TEMPO_SEM"].gt(df["PRAZO_SEM"] + 2 * anos_tolerancia)
    )
    return np.where(retido.to_numpy(), "RETIDO", df["SITUACAO_GRUPO"].astype(str).to_numpy())


# ============================================================
# BASE TRATADA · UMA CÓPIA COMPARTILHADA POR TODOS OS USUÁRIOS (v10)
# ============================================================
# leitura → consolidação → prazos → classificação, uma vez por servidor.
# st.cache_resource NÃO copia o objeto a cada acesso (st.cache_data
# copiava a base inteira a cada clique de cada usuário). Por isso a base
# é SOMENTE LEITURA: filtrar_base sempre devolve um DataFrame novo.


@dataclass(frozen=True)
class BaseIndicadores:
    """Base pronta para os filtros. SOMENTE LEITURA: nunca altere `df`."""
    df: pd.DataFrame
    funil: dict
    referencia: tuple
    campi: tuple
    cursos: pd.DataFrame     # catálogo: COD_CURSO, CURSO, CAMPI, ROTULO
    ano_min: int
    ano_max: int

    def nome_curso(self, codigo: str) -> str:
        linha = self.cursos.loc[self.cursos["COD_CURSO"] == codigo, "CURSO"]
        return linha.iat[0] if len(linha) else codigo


def versao_do_arquivo(caminho: str) -> tuple:
    """Data de modificação e tamanho: muda quando o ODS é substituído."""
    return os.path.getmtime(caminho), os.path.getsize(caminho)


def _catalogo_cursos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Um item por código de curso, com o nome mais frequente e o(s) campus em
    que é ofertado. Ordenado pelo nome, ignorando acentos.
    """
    cat = (
        df.groupby(["COD_CURSO", "CURSO", "CAMPUS"], observed=True)
        .size().rename("N").reset_index()
        .astype({"COD_CURSO": str, "CURSO": str, "CAMPUS": str})
    )
    nomes = cat.sort_values("N", ascending=False).drop_duplicates("COD_CURSO")
    campi_por_codigo = cat.groupby("COD_CURSO")["CAMPUS"].agg(lambda s: tuple(sorted(set(s))))
    cursos = nomes[["COD_CURSO", "CURSO"]].copy()
    cursos["CAMPI"] = cursos["COD_CURSO"].map(campi_por_codigo)
    cursos["ROTULO"] = [
        f"{nome} · {', '.join(campi)} ({codigo})"
        for codigo, nome, campi in zip(cursos["COD_CURSO"], cursos["CURSO"], cursos["CAMPI"])
    ]
    ordem = sorted(range(len(cursos)), key=lambda i: (
        chave_ordenacao_texto(cursos["CURSO"].iat[i]),
        chave_ordenacao_texto(", ".join(cursos["CAMPI"].iat[i])),
        cursos["COD_CURSO"].iat[i].casefold(),
    ))
    return cursos.iloc[ordem].reset_index(drop=True)


@st.cache_resource(
    show_spinner="🔄 Preparando a base de dados (acontece só na primeira abertura)…",
    max_entries=1,
)
def carregar_base(caminho: str, versao_arquivo) -> BaseIndicadores:
    """
    Lê o ODS, consolida, estima prazos e classifica — uma única vez por
    servidor. `versao_arquivo` invalida o cache quando o ODS é trocado.
    A regra de retenção (que depende da tolerância) fica para filtrar_base.
    """
    df, impressao = carregar_ods(caminho)
    df, funil = consolidar_entradas(df, impressao)
    del impressao

    prazos = estimar_prazos(df)
    referencia = semestre_referencia(df)
    df = df.merge(prazos[["CAMPUS", "COD_CURSO", "PRAZO_SEM", "PRAZO_FONTE"]],
                  on=["CAMPUS", "COD_CURSO"], how="left", validate="many_to_one")
    df = classifica_situacao(df)
    df["TEMPO_SEM"] = tempo_no_curso(df, referencia)

    # MATRICULA só serve para a deduplicação: sai daqui (memória + LGPD).
    df = df.drop(columns=["MATRICULA"])
    for col in COLUNAS_CATEGORIA:
        if col in df.columns:
            df[col] = df[col].astype("category")
    for col in ["ANO_EVASAO", "DURAÇÃO-SEM", "DURAÇÃO-ANO", "PRAZO_SEM", "TEMPO_SEM"]:
        if col in df.columns:
            df[col] = df[col].astype("float32")
    df = df.reset_index(drop=True)

    anos = df["ANO_INGRESSO"].dropna()
    return BaseIndicadores(
        df=df,
        funil=funil,
        referencia=referencia,
        campi=tuple(sorted(df["CAMPUS"].dropna().astype(str).unique())),
        cursos=_catalogo_cursos(df),
        ano_min=int(anos.min()),
        ano_max=int(anos.max()),
    )


def filtrar_base(base: BaseIndicadores, campi: tuple, cursos: tuple,
                 anos: tuple, tolerancia: int) -> pd.DataFrame:
    """
    Recorte da base conforme os filtros + regra de retenção.
    `campi`/`cursos` vazios = todos. Retorna um DataFrame NOVO, com a
    coluna SITUACAO_GRUPO_FINAL (a base compartilhada não é alterada).
    """
    df = base.df
    mask = df["ANO_INGRESSO"].between(anos[0], anos[1])
    if campi:
        mask = mask & df["CAMPUS"].isin(campi)
    if cursos:
        mask = mask & df["COD_CURSO"].isin(cursos)
    recorte = df.loc[mask]
    return recorte.assign(SITUACAO_GRUPO_FINAL=situacao_final(recorte, tolerancia))


# ============================================================
# CÁLCULO DOS ÍNDICES (evasão, retenção, sucesso)
# ============================================================
# Funções puras (sem Streamlit). Recebem o recorte filtrado, que já traz
# a coluna SITUACAO_GRUPO_FINAL (src.dados.base.filtrar_base).


GRUPOS_SAIDA = {
    "EVADIDOS": "EVASAO",
    "RETIDOS": "RETIDO",
    "FORMADOS": "SUCESSO",
    "ATIVOS": "ATIVO",
    "MIGRADOS": "MIGRADO",
}


def contar_por(df: pd.DataFrame, chaves: list) -> pd.DataFrame:
    """
    Conta as situações por grupo e calcula os índices:
        ÍNDICE DE EVASÃO   = (evadidos / total) × 100
        ÍNDICE DE RETENÇÃO = (retidos  / total) × 100
        TAXA DE SUCESSO    = (formados / total) × 100
    Colunas: chaves…, TOTAL, EVADIDOS, RETIDOS, FORMADOS, ATIVOS, MIGRADOS,
             IDX_EVASAO, IDX_RETENCAO, TAXA_SUCESSO
    """
    contagens = (
        df.groupby(chaves + ["SITUACAO_GRUPO_FINAL"], observed=True, dropna=False)
        .size()
        .unstack(fill_value=0)
    )
    resultado = pd.DataFrame(index=contagens.index)
    resultado["TOTAL"] = contagens.sum(axis=1).astype(int)
    for coluna_saida, situacao in GRUPOS_SAIDA.items():
        resultado[coluna_saida] = (
            contagens[situacao].astype(int) if situacao in contagens.columns else 0
        )
    resultado["IDX_EVASAO"] = (resultado["EVADIDOS"] / resultado["TOTAL"] * 100).round(2)
    resultado["IDX_RETENCAO"] = (resultado["RETIDOS"] / resultado["TOTAL"] * 100).round(2)
    resultado["TAXA_SUCESSO"] = (resultado["FORMADOS"] / resultado["TOTAL"] * 100).round(2)
    resultado = resultado.reset_index()
    # Categorias viram texto: os resultados são pequenos e seguem para gráficos.
    for c in chaves:
        if isinstance(resultado[c].dtype, pd.CategoricalDtype):
            resultado[c] = resultado[c].astype(str)
    return resultado


def calcular_indicadores(df: pd.DataFrame) -> pd.DataFrame:
    """
    Índices por coorte (CAMPUS × COD_CURSO × CURSO × ANO_INGRESSO).

    Colunas: CAMPUS, COD_CURSO, CURSO, ANO, TOTAL, EVADIDOS, RETIDOS,
    FORMADOS, ATIVOS, MIGRADOS, IDX_EVASAO, IDX_RETENCAO, TAXA_SUCESSO
    """
    resultado = contar_por(df, ["CAMPUS", "COD_CURSO", "CURSO", "ANO_INGRESSO"])
    resultado = resultado.rename(columns={"ANO_INGRESSO": "ANO"})
    resultado["ANO"] = resultado["ANO"].astype(int)
    return resultado.sort_values(
        ["CAMPUS", "CURSO", "COD_CURSO", "ANO"]
    ).reset_index(drop=True)


# ------------------------------------------------------------
# Resumo geral do filtro atual (alimenta KPIs e velocímetros)
# ------------------------------------------------------------

@dataclass(frozen=True)
class ResumoGeral:
    total_alunos: int
    total_cursos: int
    total_campus: int
    evadidos_total: int
    retidos_total: int
    formados_total: int
    ativos_total: int
    migrados_total: int
    idx_evasao_geral: float
    idx_retencao_geral: float
    taxa_sucesso_geral: float
    idx_migracao_geral: float

    @property
    def contagem(self) -> dict:
        """Alunos por situação, nas chaves usadas pela metodologia."""
        return {
            "SUCESSO": self.formados_total,
            "EVASAO": self.evadidos_total,
            "RETIDO": self.retidos_total,
            "ATIVO": self.ativos_total,
            "MIGRADO": self.migrados_total,
        }


def resumo_geral(df_filtrado: pd.DataFrame) -> ResumoGeral:
    """Totais e índices gerais do filtro (1 linha = 1 entrada aluno/curso/ingresso)."""
    total_alunos = len(df_filtrado)
    contagem = pd.Series(df_filtrado["SITUACAO_GRUPO_FINAL"]).value_counts()
    n = {g: int(contagem.get(g, 0)) for g in GRUPOS_SAIDA.values()}

    def pct(parte):
        return round(parte / max(total_alunos, 1) * 100, 1)

    return ResumoGeral(
        total_alunos=total_alunos,
        total_cursos=int(df_filtrado["COD_CURSO"].nunique()),
        total_campus=int(df_filtrado["CAMPUS"].nunique()),
        evadidos_total=n["EVASAO"],
        retidos_total=n["RETIDO"],
        formados_total=n["SUCESSO"],
        ativos_total=n["ATIVO"],
        migrados_total=n["MIGRADO"],
        idx_evasao_geral=pct(n["EVASAO"]),
        idx_retencao_geral=pct(n["RETIDO"]),
        taxa_sucesso_geral=pct(n["SUCESSO"]),
        idx_migracao_geral=pct(n["MIGRADO"]),
    )


# ------------------------------------------------------------
# Tabela por curso (alimenta ranking e dispersão)
# ------------------------------------------------------------

def calcular_tabela_cursos(df_filtrado: pd.DataFrame) -> pd.DataFrame:
    """Contagens e índices por curso (COD_CURSO × CURSO), com 1 casa decimal."""
    return contar_por(df_filtrado, ["COD_CURSO", "CURSO"]).round(
        {"IDX_EVASAO": 1, "TAXA_SUCESSO": 1, "IDX_RETENCAO": 1}
    )


def filtrar_ranking(tab_cursos: pd.DataFrame) -> pd.DataFrame:
    """Só cursos com alunos suficientes (MIN_ALUNOS_RANKING) entram em rankings e dispersão."""
    return tab_cursos[tab_cursos["TOTAL"] >= MIN_ALUNOS_RANKING]


# ============================================================
# ANÁLISES DA EVASÃO · momento, forma e perfil (turno)
# ============================================================
# Funções puras (sem Streamlit). Todas recebem o df_filtrado, que já traz
# a coluna SITUACAO_GRUPO_FINAL, e consideram só os alunos EVADIDOS.


def evadidos(df: pd.DataFrame) -> pd.DataFrame:
    """Linhas dos alunos classificados como EVASÃO (mesmo critério dos KPIs)."""
    return df[df["SITUACAO_GRUPO_FINAL"] == "EVASAO"]


# ------------------------------------------------------------
# 2 · MOMENTO DA EVASÃO
# ------------------------------------------------------------

def periodo_da_evasao(df: pd.DataFrame):
    """
    Período do curso em que cada aluno evadiu, em semestres, contando o
    semestre de ingresso (mesma regra da retenção):

        período = 2 × (ano_saída − ano_ingresso) + (sem_saída − sem_ingresso) + 1

    Quem sai no semestre em que entrou está no 1º período.
    Retorna (períodos válidos como int, info com as exclusões).
    """
    ev = evadidos(df)
    # astype(str): na base compartilhada os períodos são "category".
    per_ing = pd.to_numeric(ev["PERIODO_INGRESSO"].astype(str), errors="coerce").fillna(1)
    per_sai = pd.to_numeric(ev["PERIODO_EVASAO"].astype(str), errors="coerce").fillna(1)
    periodo = 2 * (ev["ANO_EVASAO"].astype(float) - ev["ANO_INGRESSO"]) + (per_sai - per_ing) + 1

    sem_data = int(periodo.isna().sum())
    com_data = periodo.dropna()
    inconsistentes = int((com_data < 1).sum())          # saída antes do ingresso
    validos = com_data[com_data >= 1].astype(int)
    return validos, {"total": len(ev), "sem_data": sem_data, "inconsistentes": inconsistentes}


def evasao_por_periodo(df: pd.DataFrame):
    """
    Evasões por período do curso. Colunas: PERIODO, ROTULO, EVASOES,
    PERCENTUAL e ACUMULADO (% das evasões que já ocorreram até o período).
    Períodos acima de PERIODO_MAX_GRAFICO são somados em "Nº ou mais".
    Retorna (DataFrame, info). info["metade_ate"] = rótulo do período em
    que o acumulado chega a 50% (v11).
    """
    validos, info = periodo_da_evasao(df)
    colunas = ["PERIODO", "ROTULO", "EVASOES", "PERCENTUAL", "ACUMULADO"]
    info["metade_ate"] = None
    if validos.empty:
        return pd.DataFrame(columns=colunas), info

    limite = min(int(validos.max()), PERIODO_MAX_GRAFICO)
    contagem = validos[validos <= limite].value_counts().reindex(range(1, limite + 1), fill_value=0)
    linhas = [(p, f"{p}º", int(n)) for p, n in contagem.items()]
    acima = int((validos > limite).sum())
    if acima:
        linhas.append((limite + 1, f"{limite + 1}º ou mais", acima))

    resultado = pd.DataFrame(linhas, columns=["PERIODO", "ROTULO", "EVASOES"])
    resultado["PERCENTUAL"] = (resultado["EVASOES"] / len(validos) * 100).round(1)
    resultado["ACUMULADO"] = (resultado["EVASOES"].cumsum() / len(validos) * 100).round(1)
    info["metade_ate"] = resultado.loc[resultado["ACUMULADO"] >= 50, "ROTULO"].iloc[0]
    return resultado[colunas], info


# ------------------------------------------------------------
# 3 · FORMAS DE EVASÃO
# ------------------------------------------------------------

def contagem_formas(df: pd.DataFrame) -> pd.Series:
    """Evasões por forma (FORMAEVASAO), da maior para a menor."""
    # astype(str): numa coluna "category", value_counts lista também as
    # situações com zero casos (Formado, Vinculado…).
    return evadidos(df)["FORMAEVASAO"].astype(str).value_counts()


def evasao_por_forma(df: pd.DataFrame):
    """
    Evasões por forma. Colunas: FORMA, EVASOES, PERCENTUAL.
    Mostra as MAX_FORMAS_EVASAO maiores e soma o resto em "Outros".
    Retorna (DataFrame, lista das formas agrupadas em "Outros").
    """
    cont = contagem_formas(df)
    agrupadas = []
    if len(cont) > MAX_FORMAS_EVASAO:
        agrupadas = list(cont.index[MAX_FORMAS_EVASAO:])
        outros = int(cont.iloc[MAX_FORMAS_EVASAO:].sum())
        cont = pd.concat([cont.iloc[:MAX_FORMAS_EVASAO], pd.Series({"Outros": outros})])

    resultado = pd.DataFrame({"FORMA": cont.index.astype(str), "EVASOES": cont.to_numpy().astype(int)})
    total = resultado["EVASOES"].sum()
    resultado["PERCENTUAL"] = (resultado["EVASOES"] / total * 100).round(1) if total else 0.0
    return resultado, agrupadas


# ------------------------------------------------------------
# 4 · PERFIL DOS EVADIDOS (turno)
# ------------------------------------------------------------

_TURNOS_EQUIVALENTES = {
    "MATUTINO": "Matutino", "MANHA": "Matutino",
    "VESPERTINO": "Vespertino", "TARDE": "Vespertino",
    "NOTURNO": "Noturno", "NOITE": "Noturno",
    "INTEGRAL": "Integral",
}


def _normaliza_turno(valor) -> str:
    """Padroniza o texto do turno (sem acento/caixa) para as quatro categorias."""
    if pd.isna(valor) or str(valor).strip() == "":
        return TURNO_NAO_IDENTIFICADO
    texto = "".join(
        c for c in unicodedata.normalize("NFKD", str(valor)) if not unicodedata.combining(c)
    ).strip().upper()
    return _TURNOS_EQUIVALENTES.get(texto, str(valor).strip().title())


def turno_dos_alunos(df: pd.DataFrame):
    """
    Turno de cada linha de `df`. Usa a coluna de COLUNAS_TURNO, se existir
    na base; senão deduz da letra do código do curso.
    Retorna (Series de turnos, origem: texto que descreve a fonte).
    """
    for coluna in COLUNAS_TURNO:
        if coluna in df.columns:
            return df[coluna].astype(object).map(_normaliza_turno), f"coluna {coluna}"
    letra = df["COD_CURSO"].astype(str).str.strip().str[2:3].str.upper()
    turnos = letra.map(TURNO_POR_LETRA_DO_CODIGO).fillna(TURNO_NAO_IDENTIFICADO)
    return turnos, "código do curso"


def evasao_por_turno(df: pd.DataFrame):
    """
    Evasões por turno. Colunas: TURNO, EVASOES, PERCENTUAL (sobre todas as evasões).
    Os quatro turnos principais aparecem sempre; outros valores e "Não identificado"
    só aparecem se houver alunos. Retorna (DataFrame, origem do turno).
    """
    ev = evadidos(df)
    turnos, origem = turno_dos_alunos(ev)
    cont = turnos.value_counts()

    extras = sorted(t for t in cont.index if t not in ORDEM_TURNOS and t != TURNO_NAO_IDENTIFICADO)
    ordem = ORDEM_TURNOS + extras + ([TURNO_NAO_IDENTIFICADO] if cont.get(TURNO_NAO_IDENTIFICADO, 0) else [])

    resultado = pd.DataFrame({"TURNO": ordem, "EVASOES": [int(cont.get(t, 0)) for t in ordem]})
    total = len(ev)
    resultado["PERCENTUAL"] = (resultado["EVASOES"] / total * 100).round(1) if total else 0.0
    return resultado, origem


# ============================================================
# AVISOS INTELIGENTES · resumo automático dos resultados
# ============================================================
# Cada aviso é uma frase gerada a partir dos números já calculados para o
# filtro atual (src.dados.resultados). Para criar um aviso novo: escreva
# uma função _aviso_xxx(...) que devolva um Insight (ou None, se não se
# aplicar) e acrescente-a na lista de gerar_insights().


@dataclass(frozen=True)
class Insight:
    icone: str
    texto: str          # pode conter <b>…</b>; valores vindos da base já vêm escapados


def _pct(parte, total) -> float:
    return round(parte / total * 100, 1) if total else 0.0


def _aviso_curso(por_campus_curso: pd.DataFrame):
    """por_campus_curso: CAMPUS, COD_CURSO, CURSO, EVADIDOS, TOTAL."""
    if por_campus_curso.empty or por_campus_curso["EVADIDOS"].max() == 0:
        return None
    topo = por_campus_curso.loc[por_campus_curso["EVADIDOS"].idxmax()]
    n, alunos_curso = int(topo["EVADIDOS"]), int(topo["TOTAL"])
    nome = escape(str(topo["CURSO"]))
    if por_campus_curso["CAMPUS"].nunique() > 1:          # só cita o campus se houver mais de um
        nome += f" ({escape(str(topo['CAMPUS']))})"
    return Insight(
        "🏫",
        f"O curso com maior número de evasões é <b>{nome}</b>: {n:,} evasões, "
        f"{_pct(n, alunos_curso)}% dos {alunos_curso:,} alunos dele no filtro.",
    )


def _aviso_periodo(por_periodo: pd.DataFrame, info: dict):
    if por_periodo.empty:
        return None
    # v11: o pico é procurado entre os períodos individuais. A barra
    # "Nº ou mais" soma vários semestres e venceria sempre na comparação.
    individuais = por_periodo[~por_periodo["ROTULO"].astype(str).str.endswith("ou mais")]
    candidatos = individuais if individuais["EVASOES"].sum() else por_periodo
    topo = candidatos.loc[candidatos["EVASOES"].idxmax()]
    rotulo = str(topo["ROTULO"])
    if rotulo.endswith("ou mais"):
        onde = f"<b>depois do {int(topo['PERIODO']) - 1}º período</b>"
    else:
        onde = f"no <b>{rotulo} período</b>"
    # v11: além do pico, quando metade das evasões já aconteceu.
    metade = info.get("metade_ate")
    complemento = (
        f" Metade das evasões acontece até o <b>{metade} período</b>."
        if metade and not str(metade).endswith("ou mais") else ""
    )
    return Insight(
        "⏱️",
        f"O maior número de evasões ocorre {onde}: {int(topo['EVASOES']):,} evasões "
        f"({topo['PERCENTUAL']}% das que têm data de saída).{complemento}",
    )


def _aviso_forma(formas: pd.Series):
    if formas.empty:
        return None
    n = int(formas.iloc[0])
    return Insight(
        "🚪",
        f"<b>{escape(str(formas.index[0]))}</b> é a principal forma de evasão: "
        f"{n:,} casos ({_pct(n, int(formas.sum()))}% das evasões).",
    )


def _aviso_turno(por_turno: pd.DataFrame, origem: str):
    identificados = por_turno[por_turno["TURNO"] != TURNO_NAO_IDENTIFICADO]
    if identificados.empty or identificados["EVASOES"].max() == 0:
        return None
    topo = identificados.loc[identificados["EVASOES"].idxmax()]
    nota = " Turno deduzido do código do curso." if origem == "código do curso" else ""
    return Insight(
        "🌙",
        f"O turno <b>{escape(str(topo['TURNO']).lower())}</b> concentra a maior quantidade de "
        f"alunos evadidos: {int(topo['EVASOES']):,} ({topo['PERCENTUAL']}%).{nota}",
    )


def gerar_insights(por_campus_curso, por_periodo, info_periodo, formas,
                   por_turno, origem_turno) -> list[Insight]:
    """Lista de avisos do filtro atual (na ordem em que aparecem na tela)."""
    if formas.empty:
        return [Insight("✅", "Nenhuma evasão registrada no filtro atual.")]

    avisos = [
        _aviso_curso(por_campus_curso),
        _aviso_periodo(por_periodo, info_periodo),
        _aviso_forma(formas),
        _aviso_turno(por_turno, origem_turno),
    ]
    return [a for a in avisos if a is not None]


# ============================================================
# RESULTADOS DO PAINEL · tudo o que a tela mostra, por filtro (v10)
# ============================================================
# Os números de cada combinação de filtros são calculados uma vez e
# reaproveitados por todos os usuários (cache limitado: CACHE_FILTROS_MAX
# combinações, expiram em CACHE_FILTROS_TTL segundos). O resultado é
# pequeno (tabelas agregadas), ao contrário da base, que não é copiada.


@dataclass(frozen=True)
class ResultadosPainel:
    resumo: ResumoGeral
    indicadores: pd.DataFrame        # campus × curso × ano
    por_ano: pd.DataFrame
    por_campus: pd.DataFrame
    por_curso: pd.DataFrame          # ranking e dispersão
    ufnt: dict                       # total, migrados, evadidos (Araguaína + Tocantinópolis)
    periodo: pd.DataFrame            # momento da evasão
    periodo_info: dict
    formas: pd.DataFrame             # formas de evasão (com "Outros")
    formas_agrupadas: list
    turno: pd.DataFrame              # perfil dos evadidos
    turno_origem: str
    insights: list[Insight]


@st.cache_data(
    max_entries=CACHE_FILTROS_MAX,
    ttl=CACHE_FILTROS_TTL,
    show_spinner="📊 Calculando indicadores…",
)
def calcular_painel(versao_arquivo, campi: tuple, cursos: tuple,
                    anos: tuple, tolerancia: int):
    """
    Todos os números do painel para uma combinação de filtros.
    Retorna None quando o filtro não tem alunos.
    """
    base = carregar_base(CAMINHO_ODS, versao_arquivo)
    df = filtrar_base(base, campi, cursos, anos, tolerancia)
    if df.empty:
        return None

    indicadores = calcular_indicadores(df)
    periodo, periodo_info = evasao_por_periodo(df)
    formas, formas_agrupadas = evasao_por_forma(df)
    turno, turno_origem = evasao_por_turno(df)
    por_campus_curso = (
        indicadores.groupby(["CAMPUS", "COD_CURSO", "CURSO"], as_index=False)
        [["EVADIDOS", "TOTAL"]].sum()
    )
    ufnt = df[df["CAMPUS"].isin(CAMPI_UFNT)]["SITUACAO_GRUPO_FINAL"]

    return ResultadosPainel(
        resumo=resumo_geral(df),
        indicadores=indicadores,
        por_ano=contar_por(df, ["ANO_INGRESSO"]).sort_values("ANO_INGRESSO"),
        por_campus=contar_por(df, ["CAMPUS"]),
        por_curso=calcular_tabela_cursos(df),
        ufnt={
            "total": len(ufnt),
            "migrados": int((ufnt == "MIGRADO").sum()),
            "evadidos": int((ufnt == "EVASAO").sum()),
        },
        periodo=periodo,
        periodo_info=periodo_info,
        formas=formas,
        formas_agrupadas=formas_agrupadas,
        turno=turno,
        turno_origem=turno_origem,
        insights=gerar_insights(
            por_campus_curso, periodo, periodo_info,
            contagem_formas(df), turno, turno_origem,
        ),
    )


# ============================================================
# EXPORTAÇÃO · geração das planilhas Excel para download
# ============================================================
# v10: o Excel é escrito linha a linha pelo xlsxwriter em modo
# "constant_memory" (cada linha vai para o disco assim que é escrita).
# Com openpyxl, cada download do relatório completo ocupava ~425 MB.
# Os arquivos são gerados só no clique e NÃO ficam em cache.


def _escrever_aba(workbook, df: pd.DataFrame, nome: str, largura_max: int) -> None:
    ws = workbook.add_worksheet(nome)
    negrito = workbook.add_format({"bold": True})
    # Largura pelo maior texto da coluna (amostra das primeiras linhas)
    amostra = df.head(2000)
    for i, col in enumerate(df.columns):
        maior = max([len(str(col))] + [len(str(v)) for v in amostra[col].tolist()])
        ws.set_column(i, i, min(maior + 4, largura_max))
    ws.write_row(0, 0, [str(c) for c in df.columns], negrito)
    ws.freeze_panes(1, 0)
    valores = df.astype(object).where(df.notna(), None)   # vazio → célula em branco
    for n_linha, linha in enumerate(valores.itertuples(index=False, name=None), start=1):
        ws.write_row(n_linha, 0, linha)


def _gerar_xlsx(abas) -> bytes:
    """abas: lista de (DataFrame, nome da aba, largura máxima)."""
    buffer = io.BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"constant_memory": True})
    for df, nome, largura_max in abas:
        _escrever_aba(workbook, df, nome, largura_max)
    workbook.close()
    return buffer.getvalue()


def gerar_excel_indicadores(df_indicadores: pd.DataFrame) -> bytes:
    """Planilha resumida (aba "Indicadores")."""
    return _gerar_xlsx([(df_indicadores, "Indicadores", 50)])


def gerar_excel_completo(df_indicadores: pd.DataFrame, filtros: tuple, versao_arquivo) -> bytes:
    """Duas abas: "Indicadores" e "Base Filtrada" (registros sem dados pessoais)."""
    base = carregar_base(CAMINHO_ODS, versao_arquivo)
    df_base = filtrar_base(base, *filtros)
    colunas = [c for c in COLUNAS_EXPORTACAO if c in df_base.columns]
    df_base = df_base[colunas].rename(columns=COLUNAS_EXPORTACAO)
    # Categorias → texto simples (o Excel não tem tipo "categoria")
    df_base = df_base.astype({c: str for c in df_base.columns
                              if isinstance(df_base[c].dtype, pd.CategoricalDtype)})
    return _gerar_xlsx([
        (df_indicadores, "Indicadores", 50),
        (df_base, "Base Filtrada", 60),
    ])


# ============================================================
# CARGA DA BASE · valida o arquivo, prepara a base e trata erros
# ============================================================


def carregar_dados() -> tuple[BaseIndicadores, tuple]:
    """
    Devolve (base compartilhada, versão do arquivo).
    Mostra o erro na tela e interrompe a execução se algo falhar.
    """
    if not os.path.isfile(CAMINHO_ODS):
        st.error(
            f"❌ Arquivo não encontrado: **{CAMINHO_ODS}**\n\n"
            "Verifique se o arquivo está na pasta data/ do projeto "
            "e se o nome está correto (16.11.03a.ods)."
        )
        st.stop()

    try:
        versao_arquivo = versao_do_arquivo(CAMINHO_ODS)
        base = carregar_base(CAMINHO_ODS, versao_arquivo)
    except Exception as e:
        st.error(f"❌ Erro ao carregar arquivo: {e}")
        st.stop()

    return base, versao_arquivo


# ============================================================
# BARRA LATERAL · filtros de campus, curso, ano e retenção (v10)
# ============================================================
# Campus em botões clicáveis, cursos numa busca única ("vazio = todos"),
# período, ajuste de retenção recolhido e botão "Limpar filtros".


@dataclass(frozen=True)
class Filtros:
    campi: tuple          # vazio = todos
    cursos: tuple         # vazio = todos (códigos)
    anos: tuple           # (início, fim)
    tolerancia: int

    def chave(self) -> tuple:
        """Argumentos de calcular_painel / filtrar_base, nesta ordem."""
        return self.campi, self.cursos, self.anos, self.tolerancia

    @property
    def restringe(self) -> bool:
        """True quando o filtro olha só parte da UFT (campus ou curso)."""
        return bool(self.campi or self.cursos)


def _passo(numero, titulo):
    """Título numerado de cada filtro."""
    st.sidebar.markdown(
        f'<div class="filtro-passo"><span>{numero}</span>{titulo}</div>',
        unsafe_allow_html=True,
    )


def filtros_sidebar(base: BaseIndicadores) -> Filtros:
    """Desenha os filtros na barra lateral e devolve a seleção atual."""
    padrao = {
        "f_campi": [],
        "f_cursos": [],
        "f_anos": (base.ano_min, base.ano_max),
        "f_tolerancia": TOLERANCIA_PADRAO,
    }
    # O valor inicial de cada filtro vem só daqui (os controles não recebem
    # "value="), para o botão Limpar poder reescrevê-los sem conflito.
    for chave, valor in padrao.items():
        st.session_state.setdefault(chave, valor)

    def limpar_filtros():
        # Grava o padrão em cada controle (em vez de apagar a chave): assim a
        # tela também é atualizada, e não só os números.
        for chave, valor in padrao.items():
            st.session_state[chave] = list(valor) if isinstance(valor, list) else valor

    col_titulo, col_limpar = st.sidebar.columns([3, 2], vertical_alignment="center")
    col_titulo.markdown("#### 🔎 Filtros")
    col_limpar.button("↺ Limpar", on_click=limpar_filtros, width="stretch",
                      help="Volta a mostrar todos os campi, cursos e anos.")

    # ---- 1. Campus --------------------------------------------------------
    _passo(1, "Campus")
    campi_sel = st.sidebar.pills(
        "Campus", base.campi, selection_mode="multi",
        key="f_campi", label_visibility="collapsed",
    )
    campi_sel = tuple(sorted(campi_sel or ()))
    st.sidebar.caption(
        f"✔ {len(campi_sel)} de {len(base.campi)} selecionado(s). Clique de novo para desmarcar."
        if campi_sel else "Todos os campi. Clique para escolher um ou mais."
    )

    # ---- 2. Curso ---------------------------------------------------------
    _passo(2, "Curso")
    cursos_disp = base.cursos
    if campi_sel:
        cursos_disp = cursos_disp[
            cursos_disp["CAMPI"].map(lambda campi: bool(set(campi) & set(campi_sel)))
        ]
    rotulos_curso = dict(zip(cursos_disp["COD_CURSO"], cursos_disp["ROTULO"]))

    # Ao trocar de campus, remove cursos que deixaram de estar disponíveis.
    cursos_validos = [c for c in st.session_state["f_cursos"] if c in rotulos_curso]
    if cursos_validos != st.session_state["f_cursos"]:
        st.session_state["f_cursos"] = cursos_validos

    cursos_sel = st.sidebar.multiselect(
        "Curso", list(rotulos_curso), key="f_cursos",
        format_func=rotulos_curso.get,
        placeholder=f"Todos os {len(rotulos_curso)} cursos · digite para buscar",
        label_visibility="collapsed",
    )
    cursos_sel = tuple(sorted(cursos_sel))
    st.sidebar.caption(
        f"✔ {len(cursos_sel)} curso(s) selecionado(s)." if cursos_sel
        else "Deixe vazio para ver todos. Busque por nome ou código."
    )

    # ---- 3. Período -------------------------------------------------------
    _passo(3, "Ano de ingresso")
    anos_sel = st.sidebar.slider(
        "Ano de ingresso", base.ano_min, base.ano_max,
        key="f_anos", label_visibility="collapsed",
    )

    # ---- 4. Retenção (ajuste avançado) ------------------------------------
    st.sidebar.markdown("")
    with st.sidebar.expander("⚙️ Ajuste da retenção"):
        tolerancia = st.slider(
            "Tolerância (anos além do prazo do curso)",
            min_value=0, max_value=TOLERANCIA_MAXIMA, key="f_tolerancia",
            help=(
                "Alunos ATIVOS com tempo cursado > prazo do curso + "
                "este valor são classificados como RETIDOS."
            ),
        )

    return Filtros(campi_sel, cursos_sel, tuple(anos_sel), int(tolerancia))


# ============================================================
# PAINEL · avisos inteligentes (resumo automático)
# ============================================================


def render_avisos(insights) -> None:
    """Cartões com as conclusões automáticas do filtro atual."""
    secao("💡 Avisos Inteligentes")
    cartoes = "".join(
        f'<div class="insight"><div class="insight-icone">{i.icone}</div>'
        f'<div>{i.texto}</div></div>'
        for i in insights
    )
    st.markdown(f'<div class="insights-wrapper">{cartoes}</div>', unsafe_allow_html=True)
    st.caption("Gerados automaticamente a partir do filtro atual da barra lateral.")


# ============================================================
# PAINEL · cartões de indicadores (KPIs globais)
# ============================================================


def render_kpis(resumo, tolerancia: int, resumo_uft=None) -> None:
    """
    Linha de cartões com totais e índices do filtro atual.
    `resumo_uft`: resumo da UFT inteira no mesmo período e tolerância (v11).
    Quando informado, os cartões de índice mostram a diferença em p.p.
    """
    r = resumo

    def _conta(parte, total, pct):
        return f"= {parte:,} ÷ {total:,} × 100 = {pct}%"

    def _comparar(campo, maior_e_melhor):
        if resumo_uft is None:
            return ""
        return comparacao_uft(getattr(r, campo), getattr(resumo_uft, campo), maior_e_melhor)

    cartoes = [
        kpi_card("👥", f"{r.total_alunos:,}", "TOTAL ALUNOS", COR_PRIMARIA, ajuda=(
            "Alunos considerados no filtro atual.\n"
            "Cada aluno conta 1 vez por curso e por ingresso.\n"
            "Não entram: declinantes, falecimentos e encerramentos\n"
            "de mobilidade/intercâmbio. Detalhes na aba 📘.")),
        kpi_card("🎓", f"{r.formados_total:,}", "FORMADOS", COR_SUCESSO, f"Sucesso: {r.taxa_sucesso_geral}%", ajuda=(
            "Taxa de Sucesso = Formados ÷ Total × 100\n"
            + _conta(r.formados_total, r.total_alunos, r.taxa_sucesso_geral)),
            comparacao=_comparar("taxa_sucesso_geral", True)),
        kpi_card("🚪", f"{r.evadidos_total:,}", "EVADIDOS", COR_EVASAO, f"Evasão: {r.idx_evasao_geral}%", ajuda=(
            "Saíram do curso sem concluir (desistência, desvínculo,\n"
            "cancelamento, jubilamento, transferência, reopção…).\n"
            "Índice de Evasão = Evadidos ÷ Total × 100\n"
            + _conta(r.evadidos_total, r.total_alunos, r.idx_evasao_geral)),
            comparacao=_comparar("idx_evasao_geral", False)),
        kpi_card("⏳", f"{r.retidos_total:,}", "RETIDOS", COR_RETENCAO, f"Retenção: {r.idx_retencao_geral}%", ajuda=(
            f"Ativos há mais tempo que o prazo do curso + {tolerancia} ano(s).\n"
            "Índice de Retenção = Retidos ÷ Total × 100\n"
            + _conta(r.retidos_total, r.total_alunos, r.idx_retencao_geral)),
            comparacao=_comparar("idx_retencao_geral", False)),
        kpi_card("📚", f"{r.ativos_total:,}", "ATIVOS (no prazo)", COR_ACENTO, ajuda=(
            "Alunos ainda matriculados e dentro do prazo do curso\n"
            "(ou em cursos novos, ainda sem prazo definido).")),
    ]
    if r.migrados_total:
        cartoes.append(kpi_card("🔀", f"{r.migrados_total:,}", "MIGRADOS UFNT", COR_TEXTO_LEVE, f"Migração: {r.idx_migracao_geral}%", ajuda=(
            f"Alunos de Araguaína e Tocantinópolis transferidos para a UFNT em {ANO_TRANSICAO_UFNT}.\n"
            "Continuam no total, mas NÃO contam como evasão nem como sucesso.\n"
            + _conta(r.migrados_total, r.total_alunos, r.idx_migracao_geral))))
    cartoes += [
        kpi_card("🏛️", f"{r.total_campus}", "CAMPUS", COR_PRIMARIA),
        kpi_card("📋", f"{r.total_cursos}", "CURSOS", COR_SECUNDARIA),
    ]
    st.markdown('<div class="kpi-wrapper">' + "".join(cartoes) + "</div>", unsafe_allow_html=True)
    if resumo_uft is not None:
        st.caption(
            "▲▼ Diferença em pontos percentuais (p.p.) para a UFT inteira no mesmo período de "
            "ingresso e com a mesma tolerância. Verde = melhor que a UFT; vermelho = pior."
        )


# ============================================================
# PAINEL · velocímetros (gauge) dos três índices
# ============================================================


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


def render_velocimetros(taxa_sucesso_geral, idx_evasao_geral, idx_retencao_geral) -> None:
    secao("📊 Visão Geral dos Índices")

    c_g1, c_g2, c_g3 = st.columns(3)

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
# PAINEL · evolução dos índices por ano de ingresso
# ============================================================


def render_evolucao(por_ano, campi_no_filtro, anos) -> None:
    """`por_ano`: contar_por(df, ["ANO_INGRESSO"]). `campi_no_filtro`: campi considerados."""
    secao("📈 Evolução Temporal dos Índices por Ano de Ingresso")

    df_evolucao = por_ano.assign(
        MATRICULADOS=lambda d: ((d["ATIVOS"] + d["RETIDOS"]) / d["TOTAL"] * 100).round(1)
    )
    detalhe = df_evolucao[["TOTAL", "MATRICULADOS"]].to_numpy()

    fig_evolucao = go.Figure()

    # v11: faixa cinza nas turmas em que muitos alunos ainda estão matriculados.
    # Nelas o sucesso e a evasão ainda vão subir: não compare com turmas antigas.
    em_andamento = df_evolucao.loc[
        df_evolucao["MATRICULADOS"] >= LIMIAR_TURMA_EM_ANDAMENTO, "ANO_INGRESSO"
    ]
    for ano in em_andamento:
        fig_evolucao.add_vrect(
            x0=ano - 0.5, x1=ano + 0.5, fillcolor=COR_TEXTO_LEVE,
            opacity=0.12, line_width=0, layer="below",
        )
    if len(em_andamento):
        fig_evolucao.add_annotation(
            x=em_andamento.max(), y=1.0, yref="paper", yanchor="bottom",
            text="Turmas em andamento", showarrow=False, xanchor="right",
            font=dict(size=11, color=COR_TEXTO_LEVE),
        )

    for coluna, nome, cor, traco in [
        ("TAXA_SUCESSO", "Taxa de Sucesso", COR_SUCESSO, None),
        ("IDX_EVASAO", "Índice de Evasão", COR_EVASAO, None),
        ("IDX_RETENCAO", "Índice de Retenção", COR_RETENCAO, "dot"),
    ]:
        fig_evolucao.add_trace(go.Scatter(
            x=df_evolucao["ANO_INGRESSO"], y=df_evolucao[coluna],
            mode="lines+markers", name=nome,
            line=dict(color=cor, width=2.5, dash=traco),
            marker=dict(size=6),
            customdata=detalhe,
            hovertemplate=(
                f"{nome}: %{{y:.1f}}%<br>Turma de %{{x}}: %{{customdata[0]:,}} alunos, "
                "%{customdata[1]}% ainda matriculados<extra></extra>"
            ),
        ))
    fig_evolucao.update_layout(
        **LAYOUT_BASE,
        height=380,
        yaxis_title="(%)",
        xaxis_title="Ano de Ingresso",
        legend=dict(orientation="h", y=1.12),
        hovermode="x unified",
    )
    # Marca a transição para a UFNT quando Araguaína/Tocantinópolis estão no filtro.
    if set(campi_no_filtro) & CAMPI_UFNT and anos[0] <= ANO_TRANSICAO_UFNT <= anos[1]:
        fig_evolucao.add_vline(
            x=ANO_TRANSICAO_UFNT,
            line_dash="dash",
            line_color=COR_TEXTO_LEVE,
            annotation_text=f"Transição UFNT ({ANO_TRANSICAO_UFNT})",
            annotation_position="top left",
            annotation_font=dict(size=11, color=COR_TEXTO_LEVE),
        )
    chart_container(fig_evolucao)
    if len(em_andamento):
        st.caption(
            f"Faixa cinza: turmas com {LIMIAR_TURMA_EM_ANDAMENTO}% ou mais dos alunos ainda matriculados. "
            "Os índices delas ainda vão mudar; compare turmas antigas com antigas."
        )


# ============================================================
# PAINEL · comparativo entre campi
# ============================================================


def bar_campus(df_campus, col_y, titulo, cor):
    df_ord = df_campus.sort_values(col_y, ascending=True)
    fig    = px.bar(
        df_ord, x=col_y, y="CAMPUS",
        orientation="h", text=col_y,
        height=300, color_discrete_sequence=[cor],
        hover_data={"TOTAL": ":,"}, labels={"TOTAL": "Alunos"},
    )
    fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
    fig.update_xaxes(range=[0, max(df_ord[col_y].max() * 1.25, 1)])   # espaço para o rótulo
    fig.update_layout(**LAYOUT_BASE, xaxis_title="%", yaxis_title="")
    fig.update_layout(title=dict(text=titulo, font=dict(size=13)))
    return fig


def render_comparativo_campus(por_campus) -> None:
    """`por_campus`: contar_por(df, ["CAMPUS"])."""
    secao("🏛️ Comparativo por Campus")

    df_campus = por_campus.assign(
        CAMPUS=lambda d: d["CAMPUS"].map(rotulo_campus)   # "Araguaína (UFNT 2023)"
    )

    c_camp1, c_camp2, c_camp3 = st.columns(3)

    with c_camp1:
        chart_container(bar_campus(df_campus, "TAXA_SUCESSO",  "Taxa de Sucesso (%)",    COR_SUCESSO))
    with c_camp2:
        chart_container(bar_campus(df_campus, "IDX_EVASAO",    "Índice de Evasão (%)",   COR_EVASAO))
    with c_camp3:
        chart_container(bar_campus(df_campus, "IDX_RETENCAO",  "Índice de Retenção (%)", COR_RETENCAO))


# ============================================================
# PAINEL · mapa de calor de evasão por campus × curso
# ============================================================


def render_heatmap(indicadores) -> None:
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
# PAINEL · momento da evasão (em qual período os alunos saem)
# ============================================================


def render_momento_evasao(dados, info) -> None:
    """`dados`, `info`: resultado de src.dados.analises.evasao_por_periodo."""
    secao("⏱️ Momento da Evasão · em qual período os alunos mais saem?")

    if dados.empty:
        st.info("Não há evasões com data de saída no filtro atual.")
        return

    fig = figura_barras_destaque(
        dados["ROTULO"], dados["EVASOES"], dados["PERCENTUAL"],
        titulo_x="Período do curso (semestres desde o ingresso)",
        fora_do_destaque=[r for r in dados["ROTULO"] if str(r).endswith("ou mais")],
    )
    # v11: curva do acumulado (eixo da direita) — "até o Nº período, X% das
    # evasões já aconteceram".
    fig.add_trace(go.Scatter(
        x=list(dados["ROTULO"]), y=list(dados["ACUMULADO"]), yaxis="y2",
        mode="lines+markers", name="% acumulado",
        line=dict(color=COR_PRIMARIA, width=2), marker=dict(size=5),
        hovertemplate="Até o %{x}: %{y:.1f}% das evasões<extra></extra>",
    ))
    fig.update_layout(yaxis2=dict(
        overlaying="y", side="right", range=[0, 105], ticksuffix="%",
        showgrid=False, title="Acumulado",
    ))
    chart_container(fig)

    nota = (
        "Barras: evasões em cada período (o período com mais evasões em destaque; a última barra "
        "soma vários semestres e não entra nessa comparação). Linha: percentual acumulado. "
        "Período = semestres entre o ingresso e a saída, contando o semestre de entrada "
        "(quem sai no semestre em que entrou está no 1º período)."
    )
    if info.get("metade_ate"):
        nota += f" Metade das evasões ocorre até o {info['metade_ate']} período."
    excluidas = []
    if info["sem_data"]:
        excluidas.append(f"{info['sem_data']:,} sem ano de saída")
    if info["inconsistentes"]:
        excluidas.append(f"{info['inconsistentes']:,} com saída anterior ao ingresso")
    if excluidas:
        nota += " Ficaram de fora: " + " e ".join(excluidas) + "."
    st.caption(nota)


# ============================================================
# PAINEL · formas de evasão (desistência, cancelamento, transferência…)
# ============================================================


def render_formas_evasao(dados, agrupadas) -> None:
    """`dados`, `agrupadas`: resultado de src.dados.analises.evasao_por_forma."""
    secao("🚪 Formas de Evasão · qual é a principal?")

    if dados.empty:
        st.info("Não há evasões no filtro atual.")
        return

    chart_container(figura_barras_destaque(
        dados["FORMA"], dados["EVASOES"], dados["PERCENTUAL"],
        horizontal=True, altura=max(300, 48 * len(dados) + 80), titulo_y="Evasões",
    ))
    if agrupadas:
        st.caption("Em “Outros”: " + ", ".join(agrupadas) + ".")


# ============================================================
# PAINEL · perfil dos alunos evadidos (turno)
# ============================================================


def render_perfil_evadidos(dados, origem) -> None:
    """`dados`, `origem`: resultado de src.dados.analises.evasao_por_turno."""
    secao("👥 Perfil dos Evadidos · por turno")

    if dados["EVASOES"].sum() == 0:
        st.info("Não há evasões no filtro atual.")
        return

    chart_container(figura_barras_destaque(
        dados["TURNO"], dados["EVASOES"], dados["PERCENTUAL"], titulo_x="Turno",
        altura=max(300, 80 + 48 * len(dados)),
    ))
    if origem == "código do curso":
        st.caption(
            "Turno deduzido da letra do código do curso (M, V, N, I), pois a base não tem "
            "coluna de turno. “Não identificado” = código fora desse padrão."
        )
    else:
        st.caption(
            f"Turno lido da {origem} do relatório. "
            "“Não identificado” = turno em branco no relatório."
        )


# ============================================================
# PAINEL · rankings de cursos (evasão e sucesso)
# ============================================================


def _top10(tab_ranking, coluna_pct, coluna_qtd, rotulo_qtd, rotulo_pct):
    """
    Os 10 maiores por `coluna_pct`. v11: o código vai junto do nome
    ("Curso (código)"), para sobrar espaço para a barra de percentual.
    """
    top = tab_ranking.nlargest(10, coluna_pct)
    return top.assign(
        Curso=top["CURSO"] + " (" + top["COD_CURSO"] + ")"
    )[["Curso", "TOTAL", coluna_qtd, coluna_pct]].rename(columns={
        "TOTAL": "Total", coluna_qtd: rotulo_qtd, coluna_pct: rotulo_pct,
    })


def _colunas(rotulo_pct):
    """v11: percentual com barra de 0 a 100% dentro da célula."""
    return {
        "Curso": st.column_config.TextColumn(width="medium"),
        rotulo_pct: st.column_config.ProgressColumn(
            rotulo_pct, min_value=0, max_value=100, format="%.1f%%"),
    }


def render_ranking(tab_ranking) -> None:
    secao("📋 Top Cursos — Evasão vs. Sucesso")
    st.caption(f"Só entram cursos com pelo menos {MIN_ALUNOS_RANKING} alunos no filtro.")

    c_top1, c_top2 = st.columns(2)

    with c_top1:
        secao("⬆️ Maiores Índices de Evasão")
        st.dataframe(
            _top10(tab_ranking, "IDX_EVASAO", "EVADIDOS", "Evadidos", "Evasão (%)"),
            width="stretch", hide_index=True, column_config=_colunas("Evasão (%)"),
        )

    with c_top2:
        secao("⬆️ Maiores Taxas de Sucesso")
        st.dataframe(
            _top10(tab_ranking, "TAXA_SUCESSO", "FORMADOS", "Formados", "Sucesso (%)"),
            width="stretch", hide_index=True, column_config=_colunas("Sucesso (%)"),
        )


# ============================================================
# PAINEL · dispersão evasão × sucesso por curso
# ============================================================


def render_dispersao(tab_ranking) -> None:
    secao("🎯 Dispersão: Evasão × Sucesso (por Curso)")

    if tab_ranking.empty:
        st.info(f"Nenhum curso do filtro tem pelo menos {MIN_ALUNOS_RANKING} alunos.")
        return

    fig_scatter = px.scatter(
        tab_ranking,                     # mínimo MIN_ALUNOS_RANKING alunos
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
# PAINEL · tabela completa de indicadores + downloads
# ============================================================


MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def render_tabela(indicadores, filtros, versao_arquivo) -> None:
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
        width="stretch",
        height=460,
        hide_index=True,
        column_config={"Ano Ingresso": st.column_config.NumberColumn(format="%d")},
    )

    # Os arquivos só são gerados no clique (callable) e não ficam em cache:
    # nada de uma planilha guardada na memória por combinação de filtros.
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        st.download_button(
            label="⬇️ Baixar Indicadores (.xlsx)",
            data=lambda: gerar_excel_indicadores(tabela_final),
            file_name="indicadores_uft.xlsx",
            mime=MIME_XLSX,
            on_click="ignore",
        )

    with col_dl2:
        st.download_button(
            label="⬇️ Baixar Relatório Completo (.xlsx)",
            data=lambda: gerar_excel_completo(tabela_final, filtros.chave(), versao_arquivo),
            file_name="relatorio_indicadores_uft.xlsx",
            mime=MIME_XLSX,
            help="Duas abas: Indicadores e Base Filtrada (sem dados pessoais).",
            on_click="ignore",
        )


# ============================================================
# ABA "PAINEL DE INDICADORES"
# ============================================================
# Cada seção da aba é um arquivo desta pasta com uma função render_*.
# Todas recebem números já calculados (src.dados.resultados), nunca a base.
# Para criar uma seção nova: calcule o que ela precisa em calcular_painel,
# crie o arquivo aqui, importe-o abaixo e chame-o em render_painel().


def render_painel(res, res_uft, filtros, campi_no_filtro, versao_arquivo) -> None:
    """
    Desenha todas as seções da aba Painel, de cima para baixo.
    `res`: ResultadosPainel do filtro. `res_uft`: UFT inteira no mesmo
    período (None quando o filtro já é a UFT inteira).
    """
    resumo = res.resumo
    render_kpis(resumo, filtros.tolerancia, res_uft.resumo if res_uft else None)
    render_avisos(res.insights)
    render_velocimetros(
        resumo.taxa_sucesso_geral,
        resumo.idx_evasao_geral,
        resumo.idx_retencao_geral,
    )
    render_evolucao(res.por_ano, campi_no_filtro, filtros.anos)
    render_comparativo_campus(res.por_campus)
    render_heatmap(res.indicadores)

    render_momento_evasao(res.periodo, res.periodo_info)
    col_forma, col_turno = st.columns(2)
    with col_forma:
        render_formas_evasao(res.formas, res.formas_agrupadas)
    with col_turno:
        render_perfil_evadidos(res.turno, res.turno_origem)

    tab_ranking = filtrar_ranking(res.por_curso)
    render_ranking(tab_ranking)
    render_dispersao(tab_ranking)

    render_tabela(res.indicadores, filtros, versao_arquivo)


# ============================================================
# METODOLOGIA · textos e categorias usados na aba "Entenda os Cálculos"
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


# ============================================================
# METODOLOGIA · uma função por seção da aba "Entenda os Cálculos"
# ============================================================


def introducao() -> None:
    """Texto de abertura da aba."""
    caixa_explicativa(
        "Este painel responde a uma pergunta simples: <b>o que aconteceu com os alunos que "
        "entraram em cada curso da UFT?</b> Cada aluno é acompanhado até a situação registrada "
        "no relatório 16.11.03a e colocado em <b>uma</b> de cinco situações. Os índices mostram "
        "quanto cada situação representa do total.<br><br>"
        "💡 <b>Todos os exemplos desta página usam os números do filtro escolhido agora na barra "
        "lateral.</b> Mude o filtro e as contas mudam junto."
    )


def secao_situacoes(n, total, pct) -> None:
    """1️⃣ As cinco situações possíveis de um aluno"""
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
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
    soma = " + ".join(f"{n[c]:,}" for c, *_ in CATEGORIAS_DIDATICAS)
    st.caption(f"As cinco situações somam o total: {soma} = {total:,} alunos (100%). "
               "Nenhum aluno fica de fora e nenhum é contado duas vezes.")


def secao_formulas(n, total, pct) -> None:
    """2️⃣ As fórmulas"""
    secao("2️⃣ As fórmulas")
    caixa_explicativa(
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
    st.dataframe(exemplos, width="stretch", hide_index=True)
    caixa_explicativa(
        f"📣 <b>Em linguagem do dia a dia:</b> de cada 100 alunos deste filtro, "
        f"<b>{pct(n['SUCESSO'])}</b> se formaram, <b>{pct(n['EVASAO'])}</b> saíram sem concluir, "
        f"<b>{pct(n['RETIDO'])}</b> estão atrasados, <b>{pct(n['ATIVO'])}</b> estão cursando no prazo"
        + (f" e <b>{pct(n['MIGRADO'])}</b> foram transferidos para a UFNT." if n["MIGRADO"] else ".")
        + "<br><br>Os gráficos por ano, campus e curso usam <b>a mesma fórmula</b>, só que aplicada "
        "a cada grupo separadamente (por exemplo, só os alunos que entraram em 2018)."
    )


def secao_quem_entra(funil) -> None:
    """3️⃣ Quem entra na conta"""
    secao("3️⃣ Quem entra na conta")
    caixa_explicativa(
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
    caixa_explicativa(linhas)
    st.caption("Esta conta é da base inteira, antes dos filtros da barra lateral.")


def secao_retencao(tolerancia, ano_ref, sem_ref) -> None:
    """4️⃣ Como a retenção é medida"""
    secao("4️⃣ Como a retenção é medida")
    limite_extra = 2 * tolerancia
    tempo_ex = 2 * (ano_ref - 2019) + (sem_ref - 1) + 1
    limite_ex = 8 + limite_extra
    caixa_explicativa(
        "Um aluno é <b>retido</b> quando ainda está matriculado, mas já passou do tempo esperado para se formar. "
        "São três passos:<br><br>"
        "<b>① Prazo do curso</b> (em semestres). Vem do PPC quando informado no código "
        "(<code>PRAZOS_OFICIAIS</code>). Quando não está informado, é estimado pelos próprios formados do curso: "
        "é a menor duração em que se formaram pelo menos 10% dos alunos que entraram por processo seletivo. "
        "Por exemplo, Medicina fica com 12 semestres, Direito com 10 e Administração com 8.<br>"
        f"<b>② Tempo no curso</b>: semestres desde o ingresso até <b>{ano_ref}.{sem_ref}</b>, o último "
        "semestre registrado na base, contando o semestre de entrada.<br>"
        f"<b>③ Regra</b>: retido se o tempo for maior que <b>prazo + tolerância</b>. A tolerância está em "
        f"<b>{tolerancia} ano(s) = {limite_extra} semestre(s)</b> e pode ser mudada na barra lateral "
        "(item “⚙️ Ajuste da retenção”)."
    )
    st.latex(r"\text{Tempo} = 2 \times (\text{ano}_{ref} - \text{ano}_{ingresso}) + "
             r"(\text{semestre}_{ref} - \text{semestre}_{ingresso}) + 1")
    caixa_explicativa(
        f"🧮 <b>Exemplo:</b> aluno que entrou em <b>2019.1</b> num curso de <b>8 semestres</b>, ainda matriculado.<br>"
        f"Tempo = 2 × ({ano_ref} − 2019) + ({sem_ref} − 1) + 1 = <b>{tempo_ex} semestres</b>.<br>"
        f"Limite = 8 + {limite_extra} = <b>{limite_ex} semestres</b>.<br>"
        + (f"Como {tempo_ex} &gt; {limite_ex}, ele é <b>retido</b>." if tempo_ex > limite_ex
           else f"Como {tempo_ex} ≤ {limite_ex}, ele está <b>no prazo</b>.")
    )


def secao_ufnt(ufnt) -> None:
    """5️⃣ Araguaína e Tocantinópolis: a transição para a UFNT"""
    secao("5️⃣ Araguaína e Tocantinópolis: a transição para a UFNT")
    caixa_explicativa(
        f"Em {ANO_TRANSICAO_UFNT}, os campi de <b>Araguaína</b> e <b>Tocantinópolis</b> passaram a formar a "
        "<b>Universidade Federal do Norte do Tocantins (UFNT)</b>, e os alunos desses campi foram "
        "transferidos automaticamente.<br><br>"
        "Esses alunos <b>não abandonaram o curso</b>, então não contam como evasão. Eles também <b>não se formaram "
        "na UFT</b>, então não contam como sucesso. Por isso formam uma situação própria, "
        "<b>Migrado para a UFNT</b>, e <b>continuam no total</b>, porque eram alunos da UFT até a transição."
    )
    if ufnt["total"]:
        tot_u, mig_u, ev_u = ufnt["total"], ufnt["migrados"], ufnt["evadidos"]
        errado = round(ev_u / (tot_u - mig_u) * 100, 1) if tot_u > mig_u else 100.0
        certo = round(ev_u / tot_u * 100, 1)
        col_e, col_c = st.columns(2)
        with col_e:
            caixa_explicativa(
                f"❌ <b>Se tirássemos os migrados do total</b><br>"
                f"{ev_u:,} ÷ ({tot_u:,} − {mig_u:,}) × 100 = "
                f"<b style='color:{COR_EVASAO};font-size:1.3rem;'>{errado}%</b> de evasão<br>"
                f"<span style='font-size:0.8rem;color:{COR_TEXTO_LEVE};'>Sobrariam quase só os alunos que já "
                "tinham saído antes da transição, e a evasão ficaria inflada (em alguns cursos, 100%).</span>"
            )
        with col_c:
            caixa_explicativa(
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


def secao_ler_painel() -> None:
    """6️⃣ Como ler cada parte do painel"""
    secao("6️⃣ Como ler cada parte do painel")
    caixa_explicativa(
        "📊 <b>Velocímetros</b>: os três índices do filtro inteiro. Quanto mais cheio, maior o percentual.<br>"
        "🏷️ <b>Resumo dos filtros</b> (logo abaixo do título): mostra o que está sendo olhado. "
        "O botão <b>↺ Limpar</b> da barra lateral volta tudo ao padrão.<br>"
        "▲▼ <b>Comparação com a UFT</b>: quando o filtro olha só parte da UFT (um campus ou curso), os "
        "cartões do topo mostram a diferença, em pontos percentuais, para a UFT inteira no mesmo período. "
        "Verde = melhor que a UFT; vermelho = pior.<br>"
        "📈 <b>Evolução por ano de ingresso</b>: cada ponto é uma turma (os alunos que entraram naquele ano). "
        "Turmas recentes ainda estão cursando, por isso têm pouco sucesso e pouca evasão, e isso é esperado. "
        f"A <b>faixa cinza</b> marca as turmas com {LIMIAR_TURMA_EM_ANDAMENTO}% ou mais dos alunos ainda "
        "matriculados. A linha tracejada marca a transição para a UFNT.<br>"
        "🏛️ <b>Comparativo por campus</b>: o mesmo índice calculado para cada campus. Os campi marcados "
        f"“(UFNT {ANO_TRANSICAO_UFNT})” têm alunos migrados no total.<br>"
        "🔥 <b>Mapa de calor</b>: evasão de cada curso em cada campus. Quanto mais vermelho, maior a evasão. "
        "O traço (–) significa que o curso não existe naquele campus, e não evasão zero. Passe o mouse "
        "para ver o total de alunos e de migrados.<br>"
        "📋 <b>Top cursos</b> e 🎯 <b>dispersão</b>: só entram cursos com pelo menos 20 alunos no filtro, para que "
        "um curso com 1 ou 2 alunos não apareça com 0% ou 100%.<br>"
        "💡 <b>Avisos inteligentes</b>: frases geradas sozinhas com os principais resultados do filtro "
        "(curso, período, forma e turno com mais evasões).<br>"
        "⏱️ <b>Momento</b>, 🚪 <b>formas</b> e 👥 <b>perfil</b> da evasão: veja a seção 7 desta página.<br>"
        "ⓘ <b>Cartões do topo</b>: passe o mouse sobre eles para ver a conta de cada número."
    )


def secao_analises_evasao() -> None:
    """7️⃣ Momento, forma e perfil da evasão"""
    secao("7️⃣ Momento, forma e perfil da evasão")
    caixa_explicativa(
        "Estes três gráficos olham <b>só para os alunos evadidos</b> do filtro atual.<br><br>"
        "⏱️ <b>Momento (período).</b> É o número de semestres entre o ingresso e a saída, contando o "
        "semestre de entrada: quem sai no mesmo semestre em que entrou está no <b>1º período</b>. "
        "É contado pelo calendário, então um aluno que trancou o curso por um tempo "
        "aparece em um período maior do que o da grade. Evasões sem ano de saída, ou com saída anterior "
        "ao ingresso, ficam fora desse gráfico (o número aparece embaixo dele). A <b>linha</b> mostra o "
        "percentual acumulado: no ponto em que ela passa de 50%, metade das evasões já aconteceu.<br><br>"
        "🚪 <b>Formas.</b> É a situação registrada no sistema (Desistência, Matrícula Cancelada, Transferência…). "
        "Aparecem as mais frequentes e o restante é somado em “Outros”.<br><br>"
        "👥 <b>Perfil (turno).</b> Vem da coluna de turno do relatório, quando existir; se não existir, "
        "é deduzido da letra do código do curso (M, V, N, I). Quando é deduzido, o painel avisa.<br><br>"
        "💡 <b>Avisos inteligentes</b> usam os mesmos números: apontam o item de maior contagem em cada "
        "análise. Um curso grande tende a ter mais evasões em número absoluto, por isso o aviso do curso "
        "mostra também o percentual sobre os alunos dele. O pico do momento da evasão é procurado entre "
        "os períodos individuais: a barra “13º ou mais” soma vários semestres e não entra nessa disputa."
    )


def secao_cuidados() -> None:
    """⚠️ Cuidados ao interpretar (+ glossário)"""
    secao("⚠️ Cuidados ao interpretar")
    caixa_explicativa(
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

    # v10: sempre aberto e em tabela HTML — as explicações longas quebram a
    # linha em vez de ficarem cortadas numa tabela do Streamlit.
    secao("📖 Glossário: situação no sistema → situação no painel")
    glossario = [(s_, "🎓 Formado") for s_ in sorted(SITUACOES_SUCESSO)]
    glossario += [(s_, "🚪 Evadido") for s_ in sorted(SITUACOES_EVASAO)]
    glossario += [(s_, "📚 Ativo (no prazo) ou ⏳ Retido") for s_ in sorted(SITUACOES_ATIVO)]
    glossario += [(s_, "🔀 Migrado para a UFNT") for s_ in sorted(SITUACOES_MIGRACAO)]
    glossario += [(s_, f"⛔ Fora da conta — {EXPLICA_FORA.get(s_, '')}")
                  for s_ in sorted(SITUACOES_FORA_DO_CALCULO)]
    linhas = "".join(
        f"<tr><td>{escape(sistema)}</td><td>{escape(no_painel)}</td></tr>"
        for sistema, no_painel in glossario
    )
    caixa_explicativa(
        '<table class="glossario"><thead><tr><th>Situação no sistema (FORMAEVASAO)</th>'
        f"<th>No painel</th></tr></thead><tbody>{linhas}</tbody></table>"
    )


# ============================================================
# ABA "ENTENDA OS CÁLCULOS" (explicação para o usuário leigo)
# ============================================================
# Todos os exemplos usam os dados do filtro atual da barra lateral.
# Para acrescentar uma seção: crie a função em secoes.py e chame-a abaixo.


def mostrar_metodologia(resultados, funil, tolerancia, referencia) -> None:
    """
    Aba didática: explica em linguagem simples de onde vem cada número.
    Todos os exemplos usam os dados do filtro atual da barra lateral
    (`resultados`: ResultadosPainel já calculado e em cache).
    """
    total = resultados.resumo.total_alunos
    contagem = resultados.resumo.contagem
    n = {chave: contagem[chave] for chave, *_ in CATEGORIAS_DIDATICAS}

    def pct(valor):
        return round(valor / total * 100, 1) if total else 0.0

    ano_ref, sem_ref = referencia

    introducao()
    secao_situacoes(n, total, pct)
    secao_formulas(n, total, pct)
    secao_quem_entra(funil)
    secao_retencao(tolerancia, ano_ref, sem_ref)
    secao_ufnt(resultados.ufnt)
    secao_ler_painel()
    secao_analises_evasao()
    secao_cuidados()


# ============================================================
# EXECUÇÃO DO PAINEL
# ============================================================

# 1. Página e estilo
configurar_pagina()
aplicar_estilos()

# 2. Base: uma única cópia na memória, compartilhada por todos os usuários
base, versao_arquivo = carregar_dados()

# 3. Barra lateral e cabeçalho
marca_sidebar()
filtros = filtros_sidebar(base)
cabecalho()
resumo_filtros(filtros, base)

# 4. Números do filtro atual (em cache por combinação de filtros)
resultados = calcular_painel(versao_arquivo, *filtros.chave())
if resultados is None:
    st.warning(
        "Nenhum aluno encontrado para esta combinação de campus, curso e ano. "
        "Ajuste os filtros ou clique em **↺ Limpar** na barra lateral."
    )
    st.stop()

# UFT inteira no mesmo período, para comparar (só quando o filtro é parcial).
resultados_uft = (
    calcular_painel(versao_arquivo, (), (), filtros.anos, filtros.tolerancia)
    if filtros.restringe else None
)

# 5. Abas
aba_painel, aba_metodo = st.tabs(["📊 Painel de Indicadores", "📘 Entenda os Cálculos"])

with aba_painel:
    render_painel(
        resultados,
        resultados_uft,
        filtros,
        campi_no_filtro=filtros.campi or base.campi,
        versao_arquivo=versao_arquivo,
    )

with aba_metodo:
    mostrar_metodologia(
        resultados=resultados,
        funil=base.funil,
        tolerancia=filtros.tolerancia,
        referencia=base.referencia,
    )

# 6. Rodapé
rodape(base.referencia)
