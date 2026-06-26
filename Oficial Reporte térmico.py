import os, io, math
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader
from PIL import Image as PILImage
PILImage.MAX_IMAGE_PIXELS = None   # evita DecompressionBombError em imagens grandes

# ─────────────────────────────────────────────────────────────
# CONFIGURAÇÕES GLOBAIS
# ─────────────────────────────────────────────────────────────
PATH = (
    "C:/Users/raul.schmidt/OneDrive - Matrix comercializadora de energia elétrica LTDA"
    "/Documentos/Gás/Inteligencia/Estudo termelétrica/IPDO e REPDOE Consolidado v.2.xlsx"
)
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PDF  = os.path.join(
    OUTPUT_DIR,
    f"Report_Termoeletrica_{datetime.today().strftime('%d_%m_%Y')}.pdf"
)

# Slide 16:9
SLIDE_W = 33.87 * cm
SLIDE_H = 19.05 * cm

# Cabeçalho dos slides de conteúdo
HEADER_H  = 1.6  * cm
FOOTER_H  = 1.05 * cm
SIDEBAR_W = 0.30 * cm

# Qualidade
DPI = 220

# Cores
C_ORANGE  = colors.HexColor("#FF4500")
C_DARK    = colors.HexColor("#002060")
C_GREY    = colors.HexColor("#595959")
C_LGREY   = colors.HexColor("#E0E0E0")
C_WHITE   = colors.white
C_BLACK   = colors.black

# Laranja claro para gradiente da capa
C_ORANGE2 = colors.HexColor("#FF6A33")   # laranja mais claro p/ highlight radial
C_ORANGE3 = colors.HexColor("#CC3700")   # laranja mais escuro p/ bordas

MESES_PT = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",
            7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}
MESES_PT_LONGO = {1:"janeiro",2:"fevereiro",3:"março",4:"abril",
                  5:"maio",6:"junho",7:"julho",8:"agosto",
                  9:"setembro",10:"outubro",11:"novembro",12:"dezembro"}

CORES_MALHA = {"NTS":"#002060","TAG":"#FF4500","TBG":"#A5A5A5"}
CORES_TIPO_USINA = {
    "Nuclear":"#C5D9C1","Gas Natural OffGrid":"#B7DDE8",
    "Gas Natural Ongrid":"#E6B8D7","Carvão":"#FCD5B4",
    "Vapor":"#FFFF99","Óleo":"#D9D9D9","Biomassa":"#7FD1AE","Diesel":"#5B9BD5",
}
CORES_REPDOE = {
    "(UC) Unit commitment":"#B7DDE8","Inflex.":"#C5D9C1",
    "Ordem de Mérito":"#E6B8D7","Razão Elétrica":"#FCD5B4",
    "Garantia Energética":"#FFFF99","GE SUB GSUB":"#D9D9D9",
    "Exportação":"#7FD1AE","Recomposição de Reserva":"#5B9BD5",
}
CORES_USINAS = [
    "#808080","#FF4500","#6A0DAD","#F4A261","#1F77B4","#2CA02C",
    "#17BECF","#FFD700","#8BC34A","#002060","#E63946","#9467BD","#00B894",
]
MOTIVOS_REPDOE = list(CORES_REPDOE.keys())

# ─────────────────────────────────────────────────────────────
# MATPLOTLIB — estilo global (Helvetica-like)
# ─────────────────────────────────────────────────────────────
plt.rcParams.update({
    "font.family":       "DejaVu Sans",
    "font.size":         9,
    "axes.titlesize":    11,
    "axes.titleweight":  "bold",
    "axes.labelsize":    9,
    "xtick.labelsize":   8,
    "ytick.labelsize":   8,
    "legend.fontsize":   8,
    "figure.dpi":        DPI,
    "savefig.dpi":       DPI,
    "figure.facecolor":  "white",
    "axes.facecolor":    "white",
    "axes.grid":         False,
})

PLOT_W = (SLIDE_W - 1.0*cm - 0.7*cm) / cm   # cm disponível para o gráfico
PLOT_H = (SLIDE_H - HEADER_H - FOOTER_H - 0.5*cm) / cm
FIGSIZE = (PLOT_W / 2.54, PLOT_H / 2.54)

# ─────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────────────────────
def janela_pivot(pivot):
    hoje = pd.Timestamp.today().normalize()
    ini  = hoje - pd.Timedelta(days=8)
    fim  = min(hoje, pivot.index.max())
    return pivot.loc[(pivot.index >= ini) & (pivot.index <= fim)]

def lbl_c(d): return f"{d.day}/{MESES_PT[d.month]}"
def lbl_l(d): return f"{d.day}/{MESES_PT_LONGO[d.month]}"
def lbl_n(d): return f"{d.day:02d}/{d.month:02d}/{d.year}"

def fmt(v, mn=0):
    return f"{int(v):,}".replace(",", ".") if v >= mn else ""

def borda(ax):
    for s in ["top","right"]:   ax.spines[s].set_visible(False)
    for s in ["left","bottom"]:
        ax.spines[s].set_visible(True); ax.spines[s].set_color("#cccccc")
    ax.tick_params(axis="x", length=3, color="#cccccc")
    ax.tick_params(axis="y", length=3, color="#cccccc")

def lbl_dentro(ax, mn=200, fs=8):
    for ct in ax.containers:
        ax.bar_label(ct, labels=[fmt(v,mn) for v in ct.datavalues],
                     label_type="center", fontsize=fs, color="black", fontweight="bold")

def lbl_topo(ax, total, fs=9):
    for i,v in enumerate(total.values):
        ax.text(i, v*1.02, fmt(v), ha="center", va="bottom", fontsize=fs, fontweight="bold")

def lbl_dir(ax, total, fs=8):
    for i,v in enumerate(total.values):
        ax.text(v*1.03, i, fmt(v), va="center", fontsize=fs, fontweight="bold")

def leg(ax, titulo, ncol=4):
    ax.legend(title=titulo, bbox_to_anchor=(0.5,-0.13), loc="upper center",
              ncol=ncol, frameon=False, title_fontsize=8)

def ler_ipdo():
    df = pd.read_excel(PATH, sheet_name="IPDO")
    df.columns = df.columns.str.strip()
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    return df

def ler_repdoe():
    df = pd.read_excel(PATH, sheet_name="REPDOE Completo")
    df.columns = df.columns.astype(str).str.strip()
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    return df

def conv(df, campo):
    if df[campo].dtype == object:
        df[campo] = (df[campo].astype(str)
                     .str.replace(".", "", regex=False)
                     .str.replace(",", ".", regex=False))
    df[campo] = pd.to_numeric(df[campo], errors="coerce")
    return df

def fig2buf(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=DPI, bbox_inches="tight", facecolor="white")
    buf.seek(0); plt.close(fig); return buf

# ─────────────────────────────────────────────────────────────
# GRÁFICOS
# ─────────────────────────────────────────────────────────────
def g_capacidade_malha():
    df = ler_ipdo(); df = conv(df, "Capacidade Verificada (m³)")
    pivot = pd.pivot_table(df, index="Data", columns="Malha",
                           values="Capacidade Verificada (m³)", aggfunc="sum", fill_value=0)
    pivot.index = pd.to_datetime(pivot.index)
    ultima = pivot.index.max()
    pivot = pivot.loc[(pivot.index >= ultima-pd.Timedelta(days=8)) & (pivot.index <= ultima)]
    pivot = pivot[[c for c in ["NTS","TAG","TBG"] if c in pivot.columns]].sort_index()
    total = pivot.sum(axis=1); lp = total + total.max()*0.18
    labels = [lbl_l(d) for d in pivot.index]
    fig, ax = plt.subplots(figsize=FIGSIZE)
    pivot.plot(kind="bar", stacked=True, ax=ax,
               color=[CORES_MALHA[c] for c in pivot.columns], edgecolor="none")
    ax2 = ax.twinx()
    ax2.set_ylim(0, lp.max()*1.06)
    ax2.plot(np.arange(len(total)), lp.values, color="#222", marker="o", lw=2, ls="--")
    ax2.set_yticks([]); ax2.set_ylabel("")
    for ct in ax.containers:
        ax.bar_label(ct, label_type="center", fmt="%.0f", fontsize=8, color="white", fontweight="bold")
    for i,v in enumerate(total.values):
        ax2.text(i, lp.iloc[i]*1.01, f"{v:.0f}", ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax.set_xticklabels(labels, rotation=30, ha="right"); borda(ax)
    leg(ax, "Malha", ncol=3); ax.set_xlabel(""); ax.set_ylabel("m³")
    plt.tight_layout(rect=[0,0.09,1,1]); return fig2buf(fig)

def g_prog_verif():
    df = ler_ipdo(); df = df[df["Tipo Usina"]=="Gas Natural Ongrid"].copy()
    for c in ["Média Diária Programada","Média Diária Verificada"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    pivot = pd.pivot_table(df, index="Data",
                           values=["Média Diária Programada","Média Diária Verificada"],
                           aggfunc="sum", fill_value=0)
    pivot.index = pd.to_datetime(pivot.index)
    pivot = janela_pivot(pivot).sort_index()
    labels = [lbl_l(d) for d in pivot.index]
    fig, ax = plt.subplots(figsize=FIGSIZE)
    pivot.plot(kind="bar", ax=ax,
               color={"Média Diária Programada":"#002060","Média Diária Verificada":"#FF4500"}, width=0.7)
    for ct in ax.containers:
        ax.bar_label(ct, labels=[fmt(v,0) if v>0 else "" for v in ct.datavalues],
                     fontsize=8, color="black", padding=3)
    ax.set_xticklabels(labels, rotation=30, ha="right"); borda(ax)
    ax.legend(["Programada","Verificada"], frameon=False,
              bbox_to_anchor=(0.5,-0.13), loc="upper center", ncol=2)
    ax.set_xlabel(""); ax.set_ylabel("MW")
    plt.tight_layout(rect=[0,0.09,1,1]); return fig2buf(fig)

def g_tipo_usina():
    df = ler_ipdo()
    df["Média Diária Verificada"] = pd.to_numeric(df["Média Diária Verificada"], errors="coerce")
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    pivot = pd.pivot_table(df, index="Data", columns="Tipo Usina",
                           values="Média Diária Verificada", aggfunc="sum", fill_value=0)
    pivot.index = pd.to_datetime(pivot.index)
    pivot = janela_pivot(pivot).sort_index(ascending=False)
    cols = [c for c in CORES_TIPO_USINA if c in pivot.columns]; pivot = pivot[cols]
    labels = [lbl_c(d) for d in pivot.index]
    fig, ax = plt.subplots(figsize=FIGSIZE)
    pivot.plot(kind="barh", stacked=True, ax=ax,
               color=[CORES_TIPO_USINA[c] for c in pivot.columns], edgecolor="none", width=0.7)
    lbl_dentro(ax, mn=200); total = pivot.sum(axis=1)
    lbl_dir(ax, total); ax.set_xlim(0, total.max()*1.18)
    ax.set_yticklabels(labels); borda(ax)
    leg(ax, "Tipo Usina", ncol=4); ax.set_xlabel("MW"); ax.set_ylabel("")
    plt.tight_layout(rect=[0,0.09,1,1]); return fig2buf(fig)

def g_repdoe(titulo, filtro=None):
    df = ler_repdoe()
    for col in MOTIVOS_REPDOE:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["TiPo_limpo"] = (df["TiPo"].astype(str).str.lower().str.strip()
                        .str.replace("-","",regex=False).str.replace(" ","",regex=False))
    if filtro: df = df[df["TiPo_limpo"]==filtro].copy()
    pivot = pd.pivot_table(df, index="Data", values=MOTIVOS_REPDOE, aggfunc="sum", fill_value=0)
    pivot = janela_pivot(pivot).sort_index(ascending=False)
    cols = [c for c in CORES_REPDOE if c in pivot.columns]
    pivot = pivot[pivot[cols].sum().sort_values(ascending=False).index.tolist()]
    labels = [lbl_c(d) for d in pivot.index]
    fig, ax = plt.subplots(figsize=FIGSIZE)
    pivot.plot(kind="barh", stacked=True, ax=ax,
               color=[CORES_REPDOE[c] for c in pivot.columns], edgecolor="none")
    lbl_dentro(ax, mn=200, fs=8); total = pivot.sum(axis=1)
    lbl_dir(ax, total); ax.set_xlim(0, total.max()*1.18)
    ax.set_yticklabels(labels); borda(ax)
    leg(ax, "Motivo", ncol=4); ax.set_xlabel("MW"); ax.set_ylabel("")
    plt.tight_layout(rect=[0,0.09,1,1]); return fig2buf(fig)

def g_dono_nome():
    df = ler_repdoe()
    df["Consumo"] = pd.to_numeric(df["Consumo"], errors="coerce").fillna(0)
    df["TOTAL"]   = pd.to_numeric(df["TOTAL"],   errors="coerce").fillna(0)
    df = df[df["Consumo"]>0].copy()
    df["Dono_Nome"] = df["Dono"].astype(str).str.strip()+" - "+df["Nome IPDO"].astype(str).str.strip()
    pivot = pd.pivot_table(df, index="Data", columns="Dono_Nome",
                           values="TOTAL", aggfunc="sum", fill_value=0)
    pivot = janela_pivot(pivot).sort_index()
    pivot = pivot[pivot.sum().sort_values(ascending=False).index.tolist()]
    labels = [lbl_n(d) for d in pivot.index]
    fig, ax = plt.subplots(figsize=FIGSIZE)
    pivot.plot(kind="bar", stacked=True, ax=ax, edgecolor="none")
    ax.set_xticklabels(labels, rotation=0)
    lbl_dentro(ax, mn=20, fs=8); total = pivot.sum(axis=1)
    lbl_topo(ax, total); ax.set_ylim(0, total.max()*1.15)
    borda(ax); leg(ax, "Dono / Nome IPDO", ncol=4)
    ax.set_xlabel(""); ax.set_ylabel("MW")
    plt.tight_layout(rect=[0,0.09,1,1]); return fig2buf(fig)

def g_exp_malha_nome():
    """
    Geração (TOTAL) por Malha + Nome IPDO — Gas Natural On-Grid,
    somente linhas com Consumo > 0 e exp == 1.
    Igual à tabela dinâmica: NTS (Cubatão, Seropédica, Termorio, Termomacaé)
    e TBG (Araucária, Canoas).
    """
    df = ler_repdoe()

    # Campos numéricos
    df["TOTAL"]   = pd.to_numeric(df["TOTAL"],   errors="coerce").fillna(0)
    df["Consumo"] = pd.to_numeric(df["Consumo"], errors="coerce").fillna(0)
    df["exp"]     = pd.to_numeric(df["exp"],     errors="coerce").fillna(0)

    # Filtros: tipo Gas Natural On-Grid + consumo > 0 + exp == 1
    df["TiPo_limpo"] = (df["TiPo"].astype(str).str.lower().str.strip()
                        .str.replace("-","",regex=False)
                        .str.replace(" ","",regex=False))
    df = df[
        (df["TiPo_limpo"] == "gasnaturalongrid") &
        (df["Consumo"] > 0) &
        (df["exp"] == 1)
    ].copy()

    # Coluna chave: Malha + Nome IPDO  (ex: "NTS - Termomacaé", "TBG - Canoas")
    df["Malha_Nome"] = (df["Malha"].astype(str).str.strip()
                        + " - "
                        + df["Nome IPDO"].astype(str).str.strip())

    pivot = pd.pivot_table(df, index="Data", columns="Malha_Nome",
                           values="TOTAL", aggfunc="sum", fill_value=0)
    pivot = janela_pivot(pivot).sort_index()

    # Ordena colunas pelo total decrescente
    pivot = pivot[pivot.sum().sort_values(ascending=False).index.tolist()]

    # Cores: NTS em tons azuis/escuros, TBG em tons cinza/verde
    CORES_EXP = [
        "#002060","#FF4500","#6A0DAD","#F4A261","#1F77B4","#2CA02C",
        "#17BECF","#FFD700","#8BC34A","#A5A5A5","#E63946","#9467BD","#00B894",
    ]

    labels = [lbl_n(d) for d in pivot.index]

    fig, ax = plt.subplots(figsize=FIGSIZE)
    pivot.plot(kind="bar", stacked=True, ax=ax,
               color=CORES_EXP[:len(pivot.columns)], edgecolor="none")
    ax.set_xticklabels(labels, rotation=0)
    lbl_dentro(ax, mn=20, fs=8)
    total = pivot.sum(axis=1)
    lbl_topo(ax, total)
    ax.set_ylim(0, total.max() * 1.15)
    borda(ax)
    leg(ax, "Malha / Nome IPDO", ncol=4)
    ax.set_xlabel(""); ax.set_ylabel("MW")
    plt.tight_layout(rect=[0,0.09,1,1])
    return fig2buf(fig)


def g_consumo_dono():
    df = ler_repdoe()
    df["TOTAL"]   = pd.to_numeric(df["TOTAL"],   errors="coerce").fillna(0)
    df["Consumo"] = pd.to_numeric(df["Consumo"], errors="coerce").fillna(0)
    df = df[df["TOTAL"]>0].copy()
    df = df[df["Dono"].notna() &
            (df["Dono"].astype(str).str.strip()!="0") &
            (df["Dono"].astype(str).str.strip()!="")].copy()
    pivot = pd.pivot_table(df, index="Data", columns="Dono",
                           values="Consumo", aggfunc="sum", fill_value=0)
    pivot = janela_pivot(pivot).sort_index()
    pivot = pivot[pivot.sum().sort_values(ascending=False).index.tolist()]
    labels = [lbl_c(d) for d in pivot.index]
    fig, ax = plt.subplots(figsize=FIGSIZE)
    pivot.plot(kind="bar", stacked=True, ax=ax, edgecolor="none")
    ax.set_xticklabels(labels, rotation=0)
    lbl_dentro(ax, mn=200, fs=9); total = pivot.sum(axis=1)
    lbl_topo(ax, total); ax.set_ylim(0, total.max()*1.15)
    borda(ax); leg(ax, "Dono", ncol=6)
    ax.set_xlabel(""); ax.set_ylabel("m³")
    plt.tight_layout(rect=[0,0.09,1,1]); return fig2buf(fig)

def g_malha_usina(malha):
    df = ler_ipdo()
    df["Capacidade Verificada (m³)"] = pd.to_numeric(
        df["Capacidade Verificada (m³)"], errors="coerce").fillna(0)
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[df["Malha"].astype(str).str.strip()==malha].copy()
    pivot = pd.pivot_table(df, index="Data", columns="Usinas",
                           values="Capacidade Verificada (m³)", aggfunc="sum", fill_value=0)
    pivot = janela_pivot(pivot).sort_index()
    pivot = pivot[pivot.sum().sort_values(ascending=False).index.tolist()]
    labels = [lbl_c(d) for d in pivot.index]
    fig, ax = plt.subplots(figsize=FIGSIZE)
    pivot.plot(kind="bar", stacked=True, ax=ax,
               color=CORES_USINAS[:len(pivot.columns)], edgecolor="none")
    ax.set_xticklabels(labels, rotation=0)
    lbl_dentro(ax, mn=200, fs=8); total = pivot.sum(axis=1)
    lbl_topo(ax, total); ax.set_ylim(0, total.max()*1.15)
    borda(ax); leg(ax, "Usinas", ncol=5)
    ax.set_xlabel(""); ax.set_ylabel("m³")
    plt.tight_layout(rect=[0,0.09,1,1]); return fig2buf(fig)

# ─────────────────────────────────────────────────────────────
# DASHBOARD — 4 gráficos por subsistema (2x2) em um só slide
# ─────────────────────────────────────────────────────────────

SUBSISTEMAS     = ["SE/CO", "NE", "SUL", "N"]
CORES_MOTIVOS   = {
    "Inflex.":                 "#C5D9C1",
    "Ordem de Mérito":         "#E6B8D7",
    "Razão Elétrica":          "#FCD5B4",
    "Garantia Energética":     "#FFFF99",
    "GE SUB GSUB":             "#D9D9D9",
    "Exportação":              "#7FD1AE",
    "Recomposição de Reserva": "#5B9BD5",
    "(UC) Unit commitment":    "#B7DDE8",
}
MOTIVOS_DASH = list(CORES_MOTIVOS.keys())

# Figsize menor para caber 4 gráficos no slide
FIGSIZE_DASH = (PLOT_W / 2.54 / 2 - 0.4,
                (PLOT_H / 2.54) / 2 - 0.5)


def _pivot_subsistema(df, regiao):
    """Filtra e pivota REPDOE por subsistema, retornando pivot diário dos motivos."""
    # normaliza coluna Região
    col_reg = next((c for c in df.columns if c.strip().lower() == "região"), None)
    if col_reg is None:
        col_reg = next((c for c in df.columns if "regi" in c.lower()), None)

    dff = df.copy()
    if col_reg:
        dff = dff[dff[col_reg].astype(str).str.strip().str.upper() == regiao.upper()].copy()

    for m in MOTIVOS_DASH:
        if m in dff.columns:
            dff[m] = pd.to_numeric(dff[m], errors="coerce").fillna(0)
        else:
            dff[m] = 0

    cols_presentes = [m for m in MOTIVOS_DASH if m in dff.columns]
    pivot = pd.pivot_table(dff, index="Data", values=cols_presentes,
                           aggfunc="sum", fill_value=0)
    pivot = janela_pivot(pivot).sort_index()

    # Ordena colunas por volume total
    ordem = pivot.sum().sort_values(ascending=False).index.tolist()
    return pivot[ordem]


def _mini_grafico(df_raw, regiao):
    """
    Gera um mini gráfico de barras empilhadas para um subsistema.
    Retorna buffer PNG.
    """
    pivot = _pivot_subsistema(df_raw, regiao)

    fig, ax = plt.subplots(figsize=FIGSIZE_DASH)

    if pivot.empty or pivot.sum().sum() == 0:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center",
                transform=ax.transAxes, fontsize=10, color="#aaa")
        ax.set_title(regiao, fontsize=10, fontweight="bold", color="#002060")
        borda(ax)
        plt.tight_layout()
        return fig2buf(fig)

    cores = [CORES_MOTIVOS[c] for c in pivot.columns if c in CORES_MOTIVOS]
    pivot.plot(kind="bar", stacked=True, ax=ax, color=cores,
               edgecolor="none", width=0.75)

    # Labels dentro apenas se segmento grande
    for ct in ax.containers:
        ax.bar_label(ct,
                     labels=[fmt(v, 150) for v in ct.datavalues],
                     label_type="center", fontsize=6,
                     color="black", fontweight="bold")

    # Total acima de cada barra
    total = pivot.sum(axis=1)
    for i, v in enumerate(total.values):
        ax.text(i, v * 1.02, fmt(v), ha="center", va="bottom",
                fontsize=7, fontweight="bold")

    ax.set_ylim(0, total.max() * 1.18)
    ax.set_title(regiao, fontsize=10, fontweight="bold", color="#002060", pad=4)
    ax.set_xticklabels([lbl_c(d) for d in pivot.index],
                       rotation=30, ha="right", fontsize=7)
    ax.set_xlabel(""); ax.set_ylabel("MW", fontsize=7)
    ax.get_legend().remove()
    borda(ax)
    plt.tight_layout()
    return fig2buf(fig)


def g_dashboard_subsistemas():
    """
    Retorna um buffer PNG com 4 mini gráficos (2x2) — um por subsistema.
    Inclui legenda compartilhada na parte inferior.
    """
    df = ler_repdoe()
    for m in MOTIVOS_DASH:
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors="coerce").fillna(0)

    fig = plt.figure(figsize=(18, 9.5), facecolor="white", dpi=110)

    gs = fig.add_gridspec(3, 2,
                          height_ratios=[1, 1, 0.16],
                          hspace=0.55, wspace=0.25,
                          left=0.05, right=0.97,
                          top=0.95, bottom=0.02)

    posicoes = [(0,0), (0,1), (1,0), (1,1)]

    for idx, regiao in enumerate(SUBSISTEMAS):
        r, c_idx = posicoes[idx]
        ax = fig.add_subplot(gs[r, c_idx])

        pivot = _pivot_subsistema(df, regiao)

        if pivot.empty or pivot.sum().sum() == 0:
            ax.text(0.5, 0.5, "Sem dados", ha="center", va="center",
                    transform=ax.transAxes, fontsize=10, color="#aaa")
            ax.set_title(regiao, fontsize=12, fontweight="bold", color="#002060")
            ax.set_xlim(0, 1); ax.set_ylim(0, 1)
            borda(ax)
            continue

        cores = [CORES_MOTIVOS.get(col, "#cccccc") for col in pivot.columns]
        pivot.plot(kind="bar", stacked=True, ax=ax, color=cores,
                   edgecolor="none", width=0.78, legend=False)

        # Labels dentro dos segmentos
        for ct in ax.containers:
            ax.bar_label(ct,
                         labels=[fmt(v, 100) for v in ct.datavalues],
                         label_type="center", fontsize=6.5,
                         color="black", fontweight="bold")

        # Total acima de cada barra
        total = pivot.sum(axis=1)
        for i, v in enumerate(total.values):
            ax.text(i, v * 1.015, fmt(v), ha="center", va="bottom",
                    fontsize=7.5, fontweight="bold")

        ax.set_ylim(0, total.max() * 1.18)
        ax.set_title(regiao, fontsize=12, fontweight="bold",
                     color="#002060", pad=5)
        ax.set_xticklabels([lbl_c(d) for d in pivot.index],
                           rotation=30, ha="right", fontsize=7.5)
        ax.set_xlabel("")
        ax.set_ylabel("MW", fontsize=8)
        borda(ax)

    # ── Legenda compartilhada ──────────────────────────────
    ax_leg = fig.add_subplot(gs[2, :])
    ax_leg.axis("off")
    patches = [
        mpatches.Patch(color=CORES_MOTIVOS[m], label=m)
        for m in MOTIVOS_DASH
        if m in CORES_MOTIVOS
    ]
    ax_leg.legend(handles=patches, loc="center", ncol=4,
                  frameon=False, fontsize=9,
                  title="Motivos de Despacho", title_fontsize=10)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=110, facecolor="white")
    buf.seek(0)
    plt.close(fig)
    return buf


# ─────────────────────────────────────────────────────────────
# SLIDE ESPECIAL — dashboard 2x2 (sem gráfico único, layout próprio)
# ─────────────────────────────────────────────────────────────
def _desenha_slide_dashboard(c, buf_png, data_str, num_pag, total_pag):
    """Slide com cabeçalho laranja e o dashboard 2x2 ocupando toda a área útil."""
    W, H = SLIDE_W, SLIDE_H

    # fundo branco
    c.setFillColor(C_WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    # cabeçalho laranja
    c.setFillColor(C_ORANGE)
    c.rect(0, H - HEADER_H, W, HEADER_H, fill=1, stroke=0)

    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 12.5)
    c.drawString(1.0*cm, H - HEADER_H + 0.50*cm,
                 "Despacho Térmico por Subsistema – Motivos (MW)")
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(W - 0.8*cm, H - HEADER_H + 0.50*cm, "Matrix Energia")

    # barra lateral laranja
    c.setFillColor(C_ORANGE)
    c.rect(0, 0, SIDEBAR_W, H - HEADER_H, fill=1, stroke=0)

    # linha separadora
    c.setStrokeColor(C_LGREY)
    c.setLineWidth(0.4)
    c.line(SIDEBAR_W, H - HEADER_H, W, H - HEADER_H)

    # imagem do dashboard
    img_x = SIDEBAR_W + 0.4*cm
    img_y = FOOTER_H + 0.2*cm
    img_w = W - img_x - 0.4*cm
    img_h = H - HEADER_H - FOOTER_H - 0.3*cm
    c.drawImage(ImageReader(buf_png), img_x, img_y,
                width=img_w, height=img_h,
                preserveAspectRatio=True, anchor="nw")

    # rodapé
    c.setStrokeColor(C_ORANGE)
    c.setLineWidth(1.2)
    c.line(SIDEBAR_W + 0.3*cm, FOOTER_H - 0.12*cm,
           W - 0.5*cm,         FOOTER_H - 0.12*cm)
    c.setFont("Helvetica", 7)
    c.setFillColor(C_GREY)
    c.drawString(SIDEBAR_W + 0.5*cm, 0.35*cm,
                 f"Matrix Energia  |  Report Termelétrico – Trading Gas  |  {data_str}")
    c.drawRightString(W - 0.6*cm, 0.35*cm, f"{num_pag} / {total_pag}")


# ─────────────────────────────────────────────────────────────
# SLIDES DUPLOS — Usinas despachando por subsistema
# ─────────────────────────────────────────────────────────────

def _pivot_usinas_regiao(df_raw, regiao):
    """
    Retorna pivot diário de TOTAL por Usina, para o subsistema (Região),
    apenas usinas com TOTAL > 0 no período e nome válido.
    """
    df = df_raw.copy()

    # Filtra por Região
    col_reg = next((c for c in df.columns if "regi" in c.lower()), None)
    if col_reg:
        df = df[df[col_reg].astype(str).str.strip().str.upper() == regiao.upper()].copy()

    df["TOTAL"] = pd.to_numeric(df["TOTAL"], errors="coerce").fillna(0)

    # Usa exatamente a coluna "Usina"
    col_usina = next((c for c in df.columns if c.strip() == "Usina"), None)
    if col_usina is None:
        # fallback: primeira coluna que contenha "usina" case-insensitive
        col_usina = next((c for c in df.columns if "usina" in c.lower()), None)
    if col_usina is None:
        return pd.DataFrame()

    # Remove linhas onde Usina é nula, "0", vazia ou numérica
    df = df[
        df[col_usina].notna() &
        (df[col_usina].astype(str).str.strip() != "") &
        (df[col_usina].astype(str).str.strip() != "0") &
        (~df[col_usina].astype(str).str.strip().str.match(r"^\d+$"))
    ].copy()

    pivot = pd.pivot_table(df, index="Data", columns=col_usina,
                           values="TOTAL", aggfunc="sum", fill_value=0)
    pivot = janela_pivot(pivot).sort_index()

    # Remove usinas com TOTAL == 0 em todos os dias do período
    pivot = pivot.loc[:, pivot.sum() > 0]

    # Ordena por volume total decrescente
    pivot = pivot[pivot.sum().sort_values(ascending=False).index.tolist()]
    return pivot


def _grafico_usinas(df_raw, regiao, ax):
    """
    Plota no ax o gráfico de barras empilhadas por usina (TOTAL > 0)
    para o subsistema indicado.
    """
    pivot = _pivot_usinas_regiao(df_raw, regiao)

    if pivot.empty or pivot.sum().sum() == 0:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center",
                transform=ax.transAxes, fontsize=11, color="#aaa")
        ax.set_title(regiao, fontsize=13, fontweight="bold", color="#002060")
        borda(ax)
        return

    # Paleta de cores para as usinas
    PALETA = [
        "#002060","#FF4500","#6A0DAD","#F4A261","#1F77B4","#2CA02C",
        "#17BECF","#FFD700","#8BC34A","#A5A5A5","#E63946","#9467BD",
        "#00B894","#D62728","#FF7F0E","#BCBD22","#7F7F7F","#17BECF",
    ]
    cores = [PALETA[i % len(PALETA)] for i in range(len(pivot.columns))]

    pivot.plot(kind="bar", stacked=True, ax=ax,
               color=cores, edgecolor="none", width=0.78, legend=False)

    # Labels dentro — só mostra se segmento >= 50 MW
    for ct in ax.containers:
        ax.bar_label(ct,
                     labels=[fmt(v, 50) for v in ct.datavalues],
                     label_type="center", fontsize=6.5,
                     color="white", fontweight="bold")

    # Total acima de cada barra
    total = pivot.sum(axis=1)
    for i, v in enumerate(total.values):
        ax.text(i, v * 1.015, fmt(v), ha="center", va="bottom",
                fontsize=8, fontweight="bold")

    ax.set_ylim(0, total.max() * 1.18)
    ax.set_title(regiao, fontsize=13, fontweight="bold", color="#002060", pad=6)
    ax.set_xticklabels([lbl_c(d) for d in pivot.index],
                       rotation=30, ha="right", fontsize=8)
    ax.set_xlabel("")
    ax.set_ylabel("MW", fontsize=8)
    borda(ax)

    # Legenda das usinas abaixo do gráfico
    handles = [mpatches.Patch(color=cores[i], label=col)
               for i, col in enumerate(pivot.columns)]
    ax.legend(handles=handles,
              bbox_to_anchor=(0.5, -0.22), loc="upper center",
              ncol=max(1, len(pivot.columns) // 2),
              frameon=False, fontsize=7, title="Usinas", title_fontsize=8)


def g_usinas_subsistema(regiao):
    """
    Gera PNG com UM gráfico de usinas despachando (TOTAL > 0)
    para o subsistema indicado, ocupando o slide inteiro.
    """
    df = ler_repdoe()
    df["TOTAL"] = pd.to_numeric(df["TOTAL"], errors="coerce").fillna(0)

    pivot = _pivot_usinas_regiao(df, regiao)

    fig, ax = plt.subplots(figsize=(18, 8.5), facecolor="white", dpi=110)
    fig.subplots_adjust(left=0.05, right=0.97, top=0.93, bottom=0.22)

    if pivot.empty or pivot.sum().sum() == 0:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center",
                transform=ax.transAxes, fontsize=14, color="#aaa")
        ax.set_title(regiao, fontsize=14, fontweight="bold", color="#002060")
        borda(ax)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=110, facecolor="white")
        buf.seek(0); plt.close(fig); return buf

    PALETA = [
        "#002060","#FF4500","#6A0DAD","#F4A261","#1F77B4","#2CA02C",
        "#17BECF","#FFD700","#8BC34A","#A5A5A5","#E63946","#9467BD",
        "#00B894","#D62728","#FF7F0E","#BCBD22","#7F7F7F","#E377C2",
    ]
    cores = [PALETA[i % len(PALETA)] for i in range(len(pivot.columns))]

    pivot.plot(kind="bar", stacked=True, ax=ax,
               color=cores, edgecolor="none", width=0.78, legend=False)

    # Labels dentro — só mostra se segmento >= 50 MW
    for ct in ax.containers:
        ax.bar_label(ct,
                     labels=[fmt(v, 50) for v in ct.datavalues],
                     label_type="center", fontsize=7,
                     color="white", fontweight="bold")

    # Total acima de cada barra
    total = pivot.sum(axis=1)
    for i, v in enumerate(total.values):
        ax.text(i, v * 1.012, fmt(v), ha="center", va="bottom",
                fontsize=9, fontweight="bold")

    ax.set_ylim(0, total.max() * 1.15)
    ax.set_xticklabels([lbl_c(d) for d in pivot.index],
                       rotation=30, ha="right", fontsize=9)
    ax.set_xlabel("")
    ax.set_ylabel("MW", fontsize=9)
    borda(ax)

    # Legenda com nomes das usinas
    handles = [mpatches.Patch(color=cores[i], label=col)
               for i, col in enumerate(pivot.columns)]
    ax.legend(handles=handles,
              bbox_to_anchor=(0.5, -0.18), loc="upper center",
              ncol=min(6, len(pivot.columns)),
              frameon=False, fontsize=8,
              title="Usinas", title_fontsize=9)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=110, facecolor="white")
    buf.seek(0); plt.close(fig)
    return buf



def _capa_bg_png(w_px, h_px):
    """
    Gera um PNG em memória que imita o fundo laranja com highlight
    claro no centro-direita (igual ao PPT original).
    """
    img_arr = np.zeros((h_px, w_px, 3), dtype=np.uint8)

    base = np.array([255, 69, 0])        # #FF4500
    hi   = np.array([255, 130, 80])      # highlight mais claro
    dk   = np.array([180, 40, 0])        # canto mais escuro

    cx, cy = w_px * 0.72, h_px * 0.38   # centro do highlight (canto direito)

    for y in range(h_px):
        for x in range(w_px):
            dx = (x - cx) / w_px
            dy = (y - cy) / h_px
            dist = math.sqrt(dx*dx + dy*dy)
            t = min(dist / 0.55, 1.0)
            # mistura base → highlight no centro, base→dark nas bordas
            if dist < 0.55:
                color = hi * (1-t) + base * t
            else:
                t2 = min((dist-0.55)/0.45, 1.0)
                color = base * (1-t2) + dk * t2
            img_arr[y, x] = color.astype(np.uint8)

    pil = PILImage.fromarray(img_arr, "RGB")
    buf = io.BytesIO()
    pil.save(buf, format="PNG")
    buf.seek(0)
    return buf


def _desenha_capa(c, data_str):
    W, H = SLIDE_W, SLIDE_H

    # ── fundo com gradiente radial ────────────────────────
    bg_buf = _capa_bg_png(1280, 720)
    c.drawImage(ImageReader(bg_buf), 0, 0, width=W, height=H)

    # ── linha branca vertical (igual ao PPT) ──────────────
    c.setStrokeColor(C_WHITE)
    c.setLineWidth(2.2)
    linha_x = 2.55 * cm
    c.line(linha_x, H*0.27, linha_x, H*0.73)

    # ── "Matrix Energia" — Helvetica-Bold grande ──────────
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 34)
    c.drawString(3.1 * cm, H * 0.52, "MTX")

    # ── subtítulo ─────────────────────────────────────────
    c.setFont("Helvetica", 14)
    c.drawString(3.1 * cm, H * 0.42, "Report Termelétrico - Trading Gas")

    # ── data ──────────────────────────────────────────────
    c.setFont("Helvetica-Bold", 13)
    c.drawString(3.1 * cm, H * 0.33, data_str)

    # ── texto "matrix" lado direito ──────────────────────
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 42)
    c.drawCentredString(W * 0.78, H * 0.47, "matrix")





# ─────────────────────────────────────────────────────────────
# SLIDE DE CONTEÚDO
# ─────────────────────────────────────────────────────────────
def _desenha_slide(c, titulo, buf_png, data_str, num_pag, total_pag):
    W, H = SLIDE_W, SLIDE_H

    # fundo branco
    c.setFillColor(C_WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    # ── cabeçalho laranja ─────────────────────────────────
    c.setFillColor(C_ORANGE)
    c.rect(0, H - HEADER_H, W, HEADER_H, fill=1, stroke=0)

    # título na faixa
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 12.5)
    c.drawString(1.0*cm, H - HEADER_H + 0.50*cm, titulo)

    # "Matrix Energia" lado direito da faixa
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(W - 0.8*cm, H - HEADER_H + 0.50*cm, "MTX")

    # ── barra lateral laranja fina ────────────────────────
    c.setFillColor(C_ORANGE)
    c.rect(0, 0, SIDEBAR_W, H - HEADER_H, fill=1, stroke=0)

    # ── linha separadora fina abaixo do cabeçalho ─────────
    c.setStrokeColor(C_LGREY)
    c.setLineWidth(0.4)
    c.line(SIDEBAR_W, H - HEADER_H, W, H - HEADER_H)

    # ── gráfico ───────────────────────────────────────────
    img_x = SIDEBAR_W + 0.5*cm
    img_y = FOOTER_H + 0.3*cm
    img_w = W - img_x - 0.5*cm
    img_h = H - HEADER_H - FOOTER_H - 0.5*cm
    c.drawImage(ImageReader(buf_png), img_x, img_y,
                width=img_w, height=img_h,
                preserveAspectRatio=True, anchor="nw")

    # ── linha laranja no rodapé ───────────────────────────
    c.setStrokeColor(C_ORANGE)
    c.setLineWidth(1.2)
    c.line(SIDEBAR_W + 0.3*cm, FOOTER_H - 0.12*cm,
           W - 0.5*cm,         FOOTER_H - 0.12*cm)

    # texto do rodapé
    c.setFont("Helvetica", 7)
    c.setFillColor(C_GREY)
    c.drawString(SIDEBAR_W + 0.5*cm, 0.35*cm,
                 f"MTX  |  Report Termelétrico – Trading Gas  |  {data_str}")
    c.drawRightString(W - 0.6*cm, 0.35*cm, f"{num_pag} / {total_pag}")


# ─────────────────────────────────────────────────────────────
# SEPARADOR DE SEÇÃO
# ─────────────────────────────────────────────────────────────
def _desenha_separador(c, titulo_secao, data_str, num_pag, total_pag):
    """
    Slide separador de seção — fundo laranja com gradiente (igual à capa),
    linha branca vertical e título da seção centralizado.
    """
    W, H = SLIDE_W, SLIDE_H

    # fundo laranja com gradiente
    bg_buf = _capa_bg_png(1280, 720)
    c.drawImage(ImageReader(bg_buf), 0, 0, width=W, height=H)

    # linha branca vertical à esquerda
    c.setStrokeColor(C_WHITE)
    c.setLineWidth(2.2)
    c.line(2.55*cm, H*0.20, 2.55*cm, H*0.80)

    # título da seção — grande, centralizado verticalmente
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 32)
    c.drawString(3.2*cm, H*0.52, titulo_secao)

    # linha decorativa branca abaixo do título
    titulo_w = len(titulo_secao) * 17   # estimativa em pontos
    c.setStrokeColor(C_WHITE)
    c.setLineWidth(1.5)
    c.line(3.2*cm, H*0.47, 3.2*cm + titulo_w, H*0.47)

    # rodapé discreto
    c.setFont("Helvetica", 7)
    c.setFillColor(C_WHITE)
    c.setFillAlpha(0.6)
    c.drawString(3.2*cm, 0.45*cm,
                 f"MTX  |  Report Termelétrico – Trading Gas  |  {data_str}")
    c.drawRightString(W - 0.6*cm, 0.45*cm, f"{num_pag} / {total_pag}")
    c.setFillAlpha(1.0)   # restaura opacidade


# ─────────────────────────────────────────────────────────────
# GERAÇÃO DO PDF
# ─────────────────────────────────────────────────────────────
def _data_ptbr():
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]
    d = datetime.today()
    return f"{d.day:02d} de {meses[d.month-1]} de {d.year}"


def gerar_pdf(graficos, output_path):
    data_str  = _data_ptbr()
    total_pag = len(graficos) + 7   # capa + dashboard + 4 subsistemas + separador + slides normais

    c = rl_canvas.Canvas(output_path, pagesize=(SLIDE_W, SLIDE_H))

    # ── Pág 1: Capa ───────────────────────────────────────
    _desenha_capa(c, data_str)
    c.showPage()

    # ── Pág 2: Dashboard 2x2 subsistemas ──────────────────
    print("  → Gerando dashboard por subsistema...")
    buf_dash = g_dashboard_subsistemas()
    _desenha_slide_dashboard(c, buf_dash, data_str, 2, total_pag)
    c.showPage()

    # ── Págs 3-6: Um slide por subsistema (usinas) ────────
    for i, reg in enumerate(["SE/CO", "SUL", "NE", "N"], start=3):
        print(f"  → Gerando slide usinas {reg}...")
        buf = g_usinas_subsistema(reg)
        _desenha_slide(c,
                       f"Usinas Despachando – {reg} (MW)",
                       buf, data_str, i, total_pag)
        c.showPage()

    # ── Pág 7: Separador de seção ─────────────────────────
    _desenha_separador(c,
                       "Estudo Térmicas – Gas Natural On-Grid",
                       data_str, 7, total_pag)
    c.showPage()

    # ── Pág 8+: Slides normais ────────────────────────────
    for i, item in enumerate(graficos, start=8):
        _desenha_slide(c, item["titulo"], item["buf"],
                       data_str, i, total_pag)
        c.showPage()

    c.save()
    print(f"\n✅  PDF gerado:\n   {output_path}\n")


# ─────────────────────────────────────────────────────────────
# EXECUÇÃO
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("⏳ Lendo dados e gerando gráficos...")

    graficos = [
        {"titulo": "Geração Térmica – Repdoe (MW) – Todos os Tipos de Usina",
         "buf": g_repdoe("Motivos de Despacho - Todos os Tipos de Usina", None)},
        {"titulo": "Geração Térmica – Repdoe (MW) – Gas Natural On-Grid",
         "buf": g_repdoe("Motivos de Despacho - Gas Natural On-Grid", "gasnaturalongrid")},
        {"titulo": "Consumo GN – Repdoe (km³) – Gas Natural On-Grid – Por Player",
         "buf": g_consumo_dono()},
        {"titulo": "Geração Térmica – Repdoe (MW) – Gas Natural On-Grid – Por Player",
         "buf": g_dono_nome()},
        {"titulo": "Geração Térmica – Gas Natural On-Grid ",
         "buf": g_exp_malha_nome()},
        {"titulo": "Geração Térmica Verificada por Tipo de Usina (MW)",
         "buf": g_tipo_usina()},
        {"titulo": "Média da Geração Programada x Verificada (Gas Natural On-Grid)",
         "buf": g_prog_verif()},
        {"titulo": "Consumo de Gás Natural por Malha – NTS, TAG e TBG",
         "buf": g_capacidade_malha()},
        {"titulo": "Consumo de Gás Natural por Malha e Player – NTS",
         "buf": g_malha_usina("NTS")},
        {"titulo": "Consumo de Gás Natural por Malha e Player – TAG",
         "buf": g_malha_usina("TAG")},
        {"titulo": "Consumo de Gás Natural por Malha e Player – TBG",
         "buf": g_malha_usina("TBG")},
    ]
