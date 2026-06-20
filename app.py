# -*- coding: utf-8 -*-
"""
Dashboard SUS - Internações e Procedimentos
IESB - Centro Universitário
Made by Erick Jost
"""

import dash
from dash import dcc, html, dash_table
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# =========================================================
# 1. INICIALIZAÇÃO DO APP
# =========================================================
app = dash.Dash(__name__)
app.title = "Dashboard SUS - Internações e Procedimentos"
server = app.server

# =========================================================
# 2. CARREGAMENTO E TRATAMENTO DOS DADOS
# =========================================================
ARQUIVO_DADOS = "CIA014_SUS_Prova_2.xlsx"

df = pd.read_excel(ARQUIVO_DADOS)

# Colunas de quantidade e valor por tipo de procedimento (qtd_01..qtd_08 / vl_02..vl_08)
COLUNAS_QTD = [c for c in df.columns if c.lower().startswith("qtd_") and c.upper() != "QTD_TOTAL"]
COLUNAS_VL = [c for c in df.columns if c.lower().startswith("vl_") and c.upper() != "VL_TOTAL"]


def converter_numero_brasileiro(serie):
    """Converte colunas que podem vir como texto no padrão BR (1.234,56) para float.
    Funciona também se a coluna já vier numérica (Excel já interpretou como número)."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    serie_str = serie.astype(str).str.strip()
    serie_str = serie_str.str.replace(".", "", regex=False)  # remove separador de milhar
    serie_str = serie_str.str.replace(",", ".", regex=False)  # vírgula decimal -> ponto
    return pd.to_numeric(serie_str, errors="coerce")


# Garante que as colunas numéricas estão de fato numéricas (caso venham como texto/objeto
# no padrão brasileiro, com vírgula decimal, como é comum em exportações do TabNet/Excel)
colunas_numericas = COLUNAS_QTD + COLUNAS_VL + ["QTD_Total", "VL_Total"]
for col in colunas_numericas:
    if col in df.columns:
        df[col] = converter_numero_brasileiro(df[col])

# Latitude/Longitude também podem vir com vírgula decimal
for col in ["LATITUDE", "LONGITUDE"]:
    if col in df.columns:
        df[col] = converter_numero_brasileiro(df[col])

# Preenche ausências de quantidade/valor com 0 (linhas sem nenhum procedimento no mês)
df[COLUNAS_QTD + COLUNAS_VL + ["QTD_Total", "VL_Total"]] = df[
    COLUNAS_QTD + COLUNAS_VL + ["QTD_Total", "VL_Total"]
].fillna(0)

# Remove municípios sem coordenadas válidas (necessárias para o mapa)
df_mapa = df.dropna(subset=["LATITUDE", "LONGITUDE"]).copy()
df_mapa = df_mapa[(df_mapa["LATITUDE"] != 0) | (df_mapa["LONGITUDE"] != 0)]

# Paleta de cores fixa por região (mantém consistência em todos os gráficos)
CORES_REGIAO = {
    "Norte": "#2E8B57",
    "Nordeste": "#E07B39",
    "Centro-Oeste": "#D4AC0D",
    "Sudeste": "#2471A3",
    "Sul": "#8E44AD",
}

# =========================================================
# 2b. ESTILOS CSS (substituem as classes do dash_bootstrap_components)
# =========================================================
FONTE = "Helvetica, Arial, sans-serif"
COR_TEXTO_MUTED = "#6c757d"
COR_BORDA = "#e2e2e2"

ESTILO_PAGINA = {
    "fontFamily": FONTE,
    "backgroundColor": "#f8f9fa",
    "padding": "24px 32px 40px 32px",
    "maxWidth": "1300px",
    "margin": "0 auto",
}

ESTILO_TEXTO_INSTITUCIONAL = {
    "color": COR_TEXTO_MUTED,
    "fontSize": "13px",
    "margin": "0",
}

ESTILO_TITULO_PRINCIPAL = {
    "fontWeight": "700",
    "margin": "4px 0 4px 0",
}

ESTILO_SUBTITULO = {
    "color": COR_TEXTO_MUTED,
    "margin": "0 0 12px 0",
}

ESTILO_LINHA_FLEX = {
    "display": "flex",
    "flexWrap": "wrap",
    "gap": "16px",
}

ESTILO_CARD = {
    "backgroundColor": "white",
    "border": f"1px solid {COR_BORDA}",
    "borderRadius": "10px",
    "padding": "20px 24px",
    "boxShadow": "0 1px 3px rgba(0,0,0,0.06)",
    "flex": "1 1 320px",
}

ESTILO_CARD_TITULO = {
    "textTransform": "uppercase",
    "color": COR_TEXTO_MUTED,
    "fontSize": "12px",
    "fontWeight": "700",
    "letterSpacing": "0.04em",
    "margin": "0 0 6px 0",
}

ESTILO_CARD_VALOR = {
    "fontWeight": "700",
    "fontSize": "32px",
    "margin": "0",
}

ESTILO_CARD_SUBVALOR = {
    "color": COR_TEXTO_MUTED,
    "fontSize": "13px",
}

ESTILO_SECAO = {
    "backgroundColor": "white",
    "border": f"1px solid {COR_BORDA}",
    "borderRadius": "10px",
    "padding": "20px 24px",
    "marginBottom": "20px",
    "boxShadow": "0 1px 3px rgba(0,0,0,0.06)",
}

ESTILO_CABECALHO_SECAO_LABEL = {
    "color": COR_TEXTO_MUTED,
    "fontSize": "12px",
    "margin": "0",
}

ESTILO_CABECALHO_SECAO_TITULO = {
    "fontWeight": "700",
    "margin": "2px 0 14px 0",
}

ESTILO_RODAPE_SECAO = {
    "textAlign": "right",
    "color": COR_TEXTO_MUTED,
    "fontSize": "12px",
    "fontStyle": "italic",
    "marginTop": "10px",
}

ESTILO_GRID_2_COL = {
    "display": "flex",
    "flexWrap": "wrap",
    "gap": "16px",
}

ESTILO_COL_METADE = {
    "flex": "1 1 460px",
    "minWidth": "0",
}

# =========================================================
# 3. FUNÇÕES AUXILIARES DE FORMATAÇÃO
# =========================================================
def formatar_numero_resumido(valor):
    """Formata números grandes em formato resumido (ex: 440M, 21B)."""
    valor = float(valor)
    if abs(valor) >= 1_000_000_000:
        return f"{valor / 1_000_000_000:.2f}B".replace(".", ",")
    if abs(valor) >= 1_000_000:
        return f"{valor / 1_000_000:.1f}M".replace(".", ",")
    if abs(valor) >= 1_000:
        return f"{valor / 1_000:.1f}K".replace(".", ",")
    return f"{valor:.0f}"


def formatar_moeda(valor):
    """Formata valor monetário no padrão brasileiro: R$ 1.234.567,89"""
    valor = float(valor)
    texto = f"{valor:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def formatar_inteiro_br(valor):
    """Formata inteiro no padrão brasileiro: 1.234.567"""
    return f"{int(round(valor)):,}".replace(",", ".")


# =========================================================
# 4. INDICADORES GLOBAIS (KPIs)
# =========================================================
total_internacoes = df["QTD_Total"].sum()
total_investido = df["VL_Total"].sum()

kpi_internacoes_fmt = formatar_numero_resumido(total_internacoes)
kpi_investido_fmt = formatar_numero_resumido(total_investido)

# =========================================================
# 5. TABELA DE DADOS POR UF
# =========================================================
tabela_uf = (
    df.groupby(["UF", "UF_Nome"], as_index=False)
    .agg(Qtd_Total_Procedimentos=("QTD_Total", "sum"), Valor_Total_Procedimentos=("VL_Total", "sum"))
    .sort_values("UF")
)

valor_total_geral = tabela_uf["Valor_Total_Procedimentos"].sum()
tabela_uf["Pct_Valor_Gasto"] = (
    (tabela_uf["Valor_Total_Procedimentos"] / valor_total_geral * 100) if valor_total_geral else 0
)

# Linha de TOTAL no rodapé
linha_total = pd.DataFrame(
    {
        "UF": ["TOTAL"],
        "UF_Nome": [""],
        "Qtd_Total_Procedimentos": [tabela_uf["Qtd_Total_Procedimentos"].sum()],
        "Valor_Total_Procedimentos": [tabela_uf["Valor_Total_Procedimentos"].sum()],
        "Pct_Valor_Gasto": [100.0 if valor_total_geral else 0],
    }
)
tabela_uf_final = pd.concat([tabela_uf, linha_total], ignore_index=True)

# Colunas formatadas para exibição
tabela_uf_final["Qtd Total de Procedimentos"] = tabela_uf_final["Qtd_Total_Procedimentos"].apply(formatar_inteiro_br)
tabela_uf_final["Valor Total dos Procedimentos"] = tabela_uf_final["Valor_Total_Procedimentos"].apply(formatar_moeda)
tabela_uf_final["% Valor Gasto"] = tabela_uf_final["Pct_Valor_Gasto"].apply(lambda v: f"{v:.2f}%".replace(".", ","))

tabela_uf_exibicao = tabela_uf_final[
    ["UF", "UF_Nome", "Qtd Total de Procedimentos", "Valor Total dos Procedimentos", "% Valor Gasto"]
].rename(columns={"UF": "UF", "UF_Nome": "Nome da UF"})

# =========================================================
# 6. MAPA DE BOLHAS GEOGRÁFICO
# =========================================================
df_mapa_agrupado = (
    df_mapa.groupby(["Codigo_Municipio", "Nome_Municipio", "UF", "Regiao_Nome", "LATITUDE", "LONGITUDE"], as_index=False)
    .agg(QTD_Total=("QTD_Total", "sum"), VL_Total=("VL_Total", "sum"))
)
df_mapa_agrupado = df_mapa_agrupado[df_mapa_agrupado["VL_Total"] > 0]

fig_mapa = px.scatter_mapbox(
    df_mapa_agrupado,
    lat="LATITUDE",
    lon="LONGITUDE",
    size="VL_Total",
    color="Regiao_Nome",
    color_discrete_map=CORES_REGIAO,
    hover_name="Nome_Municipio",
    hover_data={
        "UF": True,
        "VL_Total": ":,.2f",
        "QTD_Total": ":,.0f",
        "LATITUDE": False,
        "LONGITUDE": False,
        "Regiao_Nome": True,
    },
    size_max=38,
    zoom=3.3,
    center={"lat": -14.2, "lon": -51.9},
    mapbox_style="open-street-map",
    labels={
        "Regiao_Nome": "Região",
        "VL_Total": "Valor Total (R$)",
        "QTD_Total": "Qtd. Total",
    },
    height=600,
)
fig_mapa.update_layout(
    margin={"r": 0, "t": 10, "l": 0, "b": 0},
    legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0),
    paper_bgcolor="white",
)

# =========================================================
# 7. GRÁFICOS DA REGIÃO CENTRO-OESTE
# =========================================================
df_co = df[df["Regiao_Nome"] == "Centro-Oeste"].copy()

co_por_uf = df_co.groupby(["UF", "UF_Nome"], as_index=False).agg(
    QTD_Total=("QTD_Total", "sum"), VL_Total=("VL_Total", "sum")
)
co_por_uf["QTD_Total_Milhoes"] = co_por_uf["QTD_Total"] / 1_000_000

ORDEM_UF_CO = ["DF", "GO", "MT", "MS"]
co_por_uf["UF"] = pd.Categorical(co_por_uf["UF"], categories=ORDEM_UF_CO, ordered=True)
co_por_uf = co_por_uf.sort_values("UF")

# Gráfico 1: Barras - Quantidade total de procedimentos (em milhões) por UF
fig_barras_co = px.bar(
    co_por_uf,
    x="UF",
    y="QTD_Total_Milhoes",
    text=co_por_uf["QTD_Total_Milhoes"].apply(lambda v: f"{v:.2f}M".replace(".", ",")),
    color="UF",
    color_discrete_sequence=["#D4AC0D", "#C8920D", "#B97A0C", "#A8650C"],
    labels={"QTD_Total_Milhoes": "Qtd. Total de Procedimentos (Milhões)", "UF": "UF"},
)
fig_barras_co.update_traces(textposition="outside", showlegend=False)
fig_barras_co.update_layout(
    title="Quantidade Total de Procedimentos por UF (Centro-Oeste)",
    yaxis_title="Qtd. Total (Milhões)",
    xaxis_title="UF",
    plot_bgcolor="white",
    paper_bgcolor="white",
    margin={"t": 60, "b": 30, "l": 40, "r": 20},
)

# Gráfico 2: Donut - Distribuição do valor total por UF (Centro-Oeste)
valor_consolidado_co = co_por_uf["VL_Total"].sum()
valor_consolidado_co_fmt = formatar_numero_resumido(valor_consolidado_co)

fig_donut_co = go.Figure(
    data=[
        go.Pie(
            labels=co_por_uf["UF"],
            values=co_por_uf["VL_Total"],
            hole=0.6,
            marker=dict(colors=["#D4AC0D", "#C8920D", "#B97A0C", "#A8650C"]),
            textinfo="label+percent",
            hovertemplate="UF: %{label}<br>Valor: R$ %{value:,.2f}<extra></extra>",
        )
    ]
)
fig_donut_co.update_layout(
    title="Distribuição do Valor Total dos Procedimentos por UF (Centro-Oeste)",
    annotations=[
        dict(
            text=f"<b>{valor_consolidado_co_fmt}</b><br>Total",
            x=0.5,
            y=0.5,
            font_size=20,
            showarrow=False,
        )
    ],
    paper_bgcolor="white",
    margin={"t": 60, "b": 30, "l": 20, "r": 20},
)

# =========================================================
# 8. COMPONENTES DE LAYOUT (cabeçalho / rodapé institucional)
# =========================================================
def cabecalho_secao(titulo):
    return html.Div(
        [
            html.Span("IESB - Centro Universitário", style=ESTILO_CABECALHO_SECAO_LABEL),
            html.H4(titulo, style=ESTILO_CABECALHO_SECAO_TITULO),
        ]
    )


def rodape_secao():
    return html.Div("Made by Erick Jost", style=ESTILO_RODAPE_SECAO)


def cartao_kpi(titulo, valor_formatado, valor_bruto_formatado, cor):
    return html.Div(
        [
            html.P(titulo, style=ESTILO_CARD_TITULO),
            html.H2(valor_formatado, style={**ESTILO_CARD_VALOR, "color": cor}),
            html.Small(valor_bruto_formatado, style=ESTILO_CARD_SUBVALOR),
        ],
        style=ESTILO_CARD,
    )


# =========================================================
# 9. LAYOUT DO APP
# =========================================================
app.layout = html.Div(
    [
        # Cabeçalho principal
        html.Div(
            [
                html.P("IESB - Centro Universitário", style=ESTILO_TEXTO_INSTITUCIONAL),
                html.H1("Dashboard SUS - Internações e Procedimentos", style=ESTILO_TITULO_PRINCIPAL),
                html.P(
                    "Análise de quantidade de procedimentos e investimentos do SUS por município, UF e região",
                    style=ESTILO_SUBTITULO,
                ),
                html.Hr(style={"borderColor": COR_BORDA}),
            ],
            style={"marginBottom": "16px"},
        ),
        # KPIs
        html.Div(
            [
                cartao_kpi(
                    "Quantidade Total de Internações",
                    kpi_internacoes_fmt,
                    formatar_inteiro_br(total_internacoes),
                    "#2471A3",
                ),
                cartao_kpi(
                    "Total de Investimentos Realizados",
                    kpi_investido_fmt,
                    formatar_moeda(total_investido),
                    "#2E8B57",
                ),
            ],
            style={**ESTILO_LINHA_FLEX, "marginBottom": "24px"},
        ),
        # Tabela de dados por UF
        html.Div(
            [
                cabecalho_secao("Procedimentos e Investimentos por Unidade da Federação"),
                dash_table.DataTable(
                    data=tabela_uf_exibicao.to_dict("records"),
                    columns=[{"name": c, "id": c} for c in tabela_uf_exibicao.columns],
                    style_table={"overflowX": "auto"},
                    style_cell={
                        "textAlign": "center",
                        "padding": "8px",
                        "fontFamily": FONTE,
                        "fontSize": "14px",
                    },
                    style_header={
                        "backgroundColor": "#2471A3",
                        "color": "white",
                        "fontWeight": "bold",
                    },
                    style_data_conditional=[
                        {
                            "if": {"filter_query": '{UF} = "TOTAL"'},
                            "backgroundColor": "#EAF2F8",
                            "fontWeight": "bold",
                        },
                        {
                            "if": {"row_index": "odd"},
                            "backgroundColor": "#F8F9FA",
                        },
                    ],
                    page_size=15,
                    sort_action="native",
                ),
                rodape_secao(),
            ],
            style=ESTILO_SECAO,
        ),
        # Mapa de bolhas
        html.Div(
            [
                cabecalho_secao("Distribuição Geográfica dos Investimentos por Município"),
                dcc.Graph(figure=fig_mapa, config={"scrollZoom": True}),
                rodape_secao(),
            ],
            style=ESTILO_SECAO,
        ),
        # Gráficos da região Centro-Oeste
        html.Div(
            [
                cabecalho_secao("Detalhamento da Região Centro-Oeste (DF, GO, MT, MS)"),
                html.Div(
                    [
                        html.Div(dcc.Graph(figure=fig_barras_co), style=ESTILO_COL_METADE),
                        html.Div(dcc.Graph(figure=fig_donut_co), style=ESTILO_COL_METADE),
                    ],
                    style=ESTILO_GRID_2_COL,
                ),
                rodape_secao(),
            ],
            style=ESTILO_SECAO,
        ),
        # Rodapé geral
        html.Div(
            [
                html.Hr(style={"borderColor": COR_BORDA}),
                html.P(
                    "IESB - Centro Universitário  |  Made by Erick Jost",
                    style={"textAlign": "center", "color": COR_TEXTO_MUTED, "fontSize": "13px"},
                ),
            ],
            style={"marginTop": "12px"},
        ),
    ],
    style=ESTILO_PAGINA,
)

# =========================================================
# 10. EXECUÇÃO DO APP
# =========================================================
if __name__ == "__main__":
    app.run(debug=True)