from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Gadget Sales Analytics", page_icon="📱", layout="wide")

EXCEL_PATH = Path(__file__).parent / "data_penjualan_gadget.xlsx"
SHEET_NAME = "mobile_sales_data"

NAVY = "#2B2B3A"
CORAL = "#F5B7A5"
CORAL_DARK = "#E8896E"
MINT = "#9EE29A"
SLATE = "#6C7A89"
BG = "#4B3832"   
WHITE = "#FFFFFF"
SIDEBAR = NAVY
CARD = WHITE
BORDER = "#DADAB8"
BLACK = "#000000"
TEXT = NAVY
MUTED = SLATE
INK = "#111111"         
INK_SOFT = "#2B2B2B"    
PAGE_TEXT = "#FFFFFF"   
PALETTE = [NAVY, CORAL, MINT, SLATE, CORAL_DARK, "#C8EFC5"]
PRODUCT_COLORS = [NAVY, CORAL_DARK]

st.markdown(
    f"""
    <style>
    .stApp, [data-testid="stAppViewContainer"] {{ background-color: {BG}; color: {PAGE_TEXT}; }}
    [data-testid="stHeader"] {{ background: transparent; }}
    .stApp h1, .stApp h2, .stApp h3 {{ color: {PAGE_TEXT} !important; }}
    [data-testid="stCaptionContainer"] p {{ color: #E5E5E5 !important; }}
    [data-testid="stSidebar"] {{ background-color: {SIDEBAR}; }}
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p {{ color: {WHITE} !important; }}
    
    .stTabs [data-baseweb="tab-list"] {{ gap: 6px; border-bottom: 2px solid {CORAL}; }}

    /* Mengatur tab yang TIDAK AKTIF agar teksnya hitam */
    .stTabs [data-baseweb="tab"] {{ 
        background-color: #E0E0E0 !important; /* Beri background terang agar teks hitam terbaca */
        padding: 8px 18px; 
        border-radius: 8px 8px 0 0; 
    }}
    .stTabs [data-baseweb="tab"] * {{ 
        color: {INK} !important; /* teks tab tidak aktif: hitam */
    }}
    
    /* Mengatur tab yang SEDANG AKTIF */
    .stTabs [aria-selected="true"] {{ 
        background: {CORAL} !important; 
        font-weight: 600; 
    }}
    .stTabs [aria-selected="true"] * {{ 
        color: {INK} !important; /* teks tab aktif: hitam di atas coral */
    }}

    /* Tombol download: latar terang, teks hitam */
    [data-testid="stDownloadButton"] button {{ background-color: {CORAL} !important; border: 1px solid {WHITE}; }}
    [data-testid="stDownloadButton"] button * {{ color: {INK} !important; font-weight: 600; }}

    /* Peringatan / error: teks gelap di atas kotak terang */
    [data-testid="stAlert"] * {{ color: {INK} !important; }}

    hr {{ border-color: {BORDER}; }}

    .kpi {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-left: 8px solid var(--accent);
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 14px;
        box-shadow: 0 1px 3px rgba(43,43,58,0.08);
    }}
    .kpi .label {{ color: {INK_SOFT}; font-size: 0.85rem; letter-spacing: .04em; text-transform: uppercase; font-weight: 600; }}
    .kpi .value {{ color: {INK}; font-size: 1.85rem; font-weight: 700; line-height: 1.3; }}
    .kpi .sub   {{ color: {INK_SOFT}; font-size: 0.8rem; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# fungsi bantu
def rupiah(x: float) -> str:
    sign = "-" if x < 0 else ""
    return f"{sign}Rp {abs(x):,.0f}".replace(",", ".")


def rupiah_short(x: float) -> str:
    a = abs(x)
    sign = "-" if x < 0 else ""
    if a >= 1e9:
        return f"{sign}Rp {a / 1e9:,.2f} M".replace(",", "X").replace(".", ",").replace("X", ".")
    if a >= 1e6:
        return f"{sign}Rp {a / 1e6:,.1f} jt".replace(",", "X").replace(".", ",").replace("X", ".")
    return rupiah(x)


def angka(x: float) -> str:
    return f"{x:,.0f}".replace(",", ".")


def kpi_card(col, icon: str, label: str, value: str, sub: str, accent: str):
    col.markdown(
        f"""
        <div class="kpi" style="--accent:{accent}">
            <div class="label">{icon} {label}</div>
            <div class="value">{value}</div>
            <div class="sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def style_fig(fig, height=400):
    grid = "rgba(43,43,58,0.12)"
    fig.update_layout(
        paper_bgcolor=CARD,
        plot_bgcolor=CARD,
        font=dict(color=INK, size=13),
        title=dict(x=0.02, xanchor="left", font=dict(color=INK, size=16)),
        legend=dict(font=dict(color=INK), bgcolor="rgba(0,0,0,0)"),
        height=height,
        margin=dict(l=40, r=30, t=70, b=50),
        hoverlabel=dict(bgcolor=WHITE, font_color=INK, bordercolor=NAVY),
    )
    axis = dict(tickfont=dict(color=INK), title_font=dict(color=INK),
                gridcolor=grid, linecolor=NAVY, zerolinecolor=grid,
                automargin=True, title_standoff=15)
    fig.update_xaxes(**axis)
    fig.update_yaxes(**axis)
    return fig


def show(fig, height=400):
    st.plotly_chart(style_fig(fig, height), theme=None, config={"displayModeBar": False})


# Load data

@st.cache_data(show_spinner="Memuat data...")
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    df.columns = df.columns.str.strip()

    df["Price"] = pd.to_numeric(df["Price"], errors="coerce")
    df["Quantity Sold"] = pd.to_numeric(df["Quantity Sold"], errors="coerce")
    df["Dispatch Date"] = pd.to_datetime(df["Dispatch Date"], errors="coerce")
    df["Inward Date"] = pd.to_datetime(df["Inward Date"], errors="coerce")
    df = df.dropna(subset=["Price", "Quantity Sold", "Dispatch Date"]).copy()

    df["Total_Revenue"] = df["Price"] * df["Quantity Sold"]
    df["Bulan"] = df["Dispatch Date"].dt.to_period("M").dt.to_timestamp()
    df["Lead_Time"] = (df["Dispatch Date"] - df["Inward Date"]).dt.days
    return df


if not EXCEL_PATH.exists():
    st.error(f"File data tidak ditemukan: `{EXCEL_PATH}`. Letakkan di folder yang sama dengan file ini.")
    st.stop()

df = load_data(str(EXCEL_PATH))

# Sidebar: filter

st.sidebar.header("🔎 Filter")

min_date = df["Dispatch Date"].min().date()
max_date = df["Dispatch Date"].max().date()

start_date, end_date = st.sidebar.slider(
    "Periode Dispatch (Shipping)",
    min_value=min_date,
    max_value=max_date,
    value=(min_date, max_date),
    format="DD MMM YYYY",
)


def multi(label: str, col: str):
    opts = sorted(df[col].dropna().unique())
    return st.sidebar.multiselect(label, opts, default=opts)


product_filter = multi("Kategori Produk", "Product")
brand_filter = multi("Brand", "Brand")
region_filter = multi("Wilayah (Region)", "Region")

mask = (
    (df["Dispatch Date"] >= pd.Timestamp(start_date))
    & (df["Dispatch Date"] < pd.Timestamp(end_date) + pd.Timedelta(days=1))
    & df["Product"].isin(product_filter)
    & df["Brand"].isin(brand_filter)
    & df["Region"].isin(region_filter)
)
filtered = df[mask].copy()

# Header

st.title("Gadget Sales Analytics")

if filtered.empty:
    st.warning("Tidak ada data untuk kombinasi filter yang dipilih.")
    st.stop()

# KPI (6 indikator, 2 baris)

total_trx = len(filtered)
unit_terjual = int(filtered["Quantity Sold"].sum())
total_revenue = filtered["Total_Revenue"].sum()
avg_price = total_revenue / unit_terjual
avg_rev_trx = total_revenue / total_trx
avg_lead = filtered["Lead_Time"].mean()

top_brand = filtered.groupby("Brand")["Total_Revenue"].sum().idxmax()
top_region = filtered.groupby("Region")["Total_Revenue"].sum().idxmax()
n_brand = filtered["Brand"].nunique()

c1, c2, c3 = st.columns(3)
kpi_card(c1, "🧾", "Total Transaksi", angka(total_trx), f"{n_brand} brand aktif", NAVY)
kpi_card(c2, "📦", "Unit Terjual", angka(unit_terjual),
         f"≈ {unit_terjual / total_trx:,.1f} unit/transaksi".replace(".", ","), CORAL)
kpi_card(c3, "💰", "Total Revenue", rupiah_short(total_revenue), rupiah(total_revenue), MINT)

c4, c5, c6 = st.columns(3)
kpi_card(c4, "🏷️", "Harga Rata-rata/Unit", rupiah(avg_price), "Revenue ÷ unit terjual", SLATE)
kpi_card(c5, "🧮", "Rata-rata Revenue/Transaksi", rupiah_short(avg_rev_trx),
         f"Brand teratas: {top_brand}", CORAL)
kpi_card(c6, "⏱️", "Rata-rata Lead Time", f"{avg_lead:,.1f} hari".replace(".", ","),
         f"Region teratas: {top_region}", MINT)

# Visualisasi

tab1, tab2, tab3, tab4 = st.tabs(["📈 Ringkasan", "🏷️ Brand & Harga", "📊 Distribusi", "📋 Data"])

with tab1:
    left, right = st.columns([3, 2])

    with left:
        monthly = (
            filtered.groupby(["Bulan", "Product"], as_index=False)["Total_Revenue"].sum()
            .sort_values("Bulan")
        )
        monthly["Revenue (juta)"] = monthly["Total_Revenue"] / 1e6
        fig = px.line(monthly, x="Bulan", y="Revenue (juta)", color="Product", markers=True,
                      title="Tren Revenue Bulanan", color_discrete_sequence=PRODUCT_COLORS)
        fig.update_traces(line=dict(width=3), marker=dict(size=7),
                          hovertemplate="%{x|%b %Y}<br>Rp %{y:,.1f} juta<extra>%{fullData.name}</extra>")
        fig.update_layout(xaxis_title="", yaxis_title="Revenue (Rp juta)", legend_title_text="")
        show(fig, 420)

    with right:
        reg_sales = filtered.groupby("Region", as_index=False)["Total_Revenue"].sum()
        fig = px.pie(reg_sales, names="Region", values="Total_Revenue", hole=0.55,
                     title="Kontribusi Revenue per Region", color_discrete_sequence=PALETTE)
        fig.update_traces(textinfo="percent+label", textposition="outside",
                          textfont=dict(color=INK, size=13),
                          marker=dict(line=dict(color=WHITE, width=3)))
        fig.update_layout(showlegend=False)
        show(fig, 420)

with tab2:
    left, right = st.columns(2)

    with left:
        brand_sales = (
            filtered.groupby("Brand", as_index=False)["Total_Revenue"].sum()
            .sort_values("Total_Revenue", ascending=False).head(10)
        )
        brand_sales["Revenue (juta)"] = brand_sales["Total_Revenue"] / 1e6
        brand_sales["Label"] = brand_sales["Total_Revenue"].apply(rupiah_short)
        fig = px.bar(brand_sales, x="Brand", y="Revenue (juta)", text="Label",
                     title="Top 10 Brand Berdasarkan Revenue")
        fig.update_traces(marker_color=CORAL, textposition="outside", cliponaxis=False,
                          textfont=dict(color=INK),
                          hovertemplate="%{x}<br>Rp %{y:,.1f} juta<extra></extra>")
        fig.update_layout(xaxis_title="", yaxis_title="Revenue (Rp juta)")
        fig.update_xaxes(tickangle=-35)
        show(fig, 440)

    with right:
        brand_prod = (
            filtered.groupby(["Brand", "Product"], as_index=False)
            .agg(Revenue=("Total_Revenue", "sum"), Unit=("Quantity Sold", "sum"),
                 Transaksi=("Price", "size"))
        )
        brand_prod["Harga Rata-rata/Unit"] = brand_prod["Revenue"] / brand_prod["Unit"]
        fig = px.scatter(brand_prod, x="Harga Rata-rata/Unit", y="Unit", color="Product",
                         size="Transaksi", hover_name="Brand", size_max=30, opacity=0.85,
                         title="Harga Rata-rata vs Unit Terjual (per Brand)",
                         color_discrete_sequence=PRODUCT_COLORS)
        fig.update_layout(xaxis_title="Harga rata-rata per unit (Rp)",
                          yaxis_title="Total unit terjual", legend_title_text="")
        show(fig, 440)

with tab3:
    left, right = st.columns(2)

    with left:
        fig = px.histogram(filtered, x="Price", color="Product", nbins=30, barmode="overlay",
                           opacity=0.75, title="Distribusi Harga Produk",
                           color_discrete_sequence=PRODUCT_COLORS)
        fig.update_layout(xaxis_title="Harga (Rp)", yaxis_title="Jumlah transaksi",
                          legend_title_text="")
        show(fig, 440)

    with right:
        heat = filtered.pivot_table(index="Region", columns="Brand", values="Total_Revenue",
                                    aggfunc="sum", fill_value=0) / 1e6
        fig = px.imshow(heat, aspect="auto", text_auto=".0f",
                        color_continuous_scale=[[0, "#F7FDF6"], [0.5, MINT], [1, "#4CAF50"]],
                        title="Heatmap Revenue: Region × Brand (Dalam Juta)")
        fig.update_traces(textfont=dict(color=INK))
        fig.update_layout(xaxis_title="", yaxis_title="", coloraxis_colorbar=dict(title="Juta"))
        fig.update_xaxes(tickangle=-45)
        show(fig, 440)

with tab4:
    kolom_tabel = ["Dispatch Date", "Product", "Brand", "Product Code", "Price", "Quantity Sold",
                   "Total_Revenue", "Region", "Customer Name", "Customer Location", "Lead_Time"]
    display_df = filtered.sort_values("Dispatch Date", ascending=False)[kolom_tabel]

    st.dataframe(display_df.head(1000), hide_index=True)
    st.caption(f"Menampilkan 1.000 transaksi terbaru dari {angka(len(filtered))} hasil filter. "
               "Unduh CSV untuk data lengkap.")
    st.download_button(
        "⬇️ Download Data Hasil Filter (CSV)",
        data=display_df.to_csv(index=False).encode("utf-8"),
        file_name="data_penjualan_gadget_filtered.csv",
        mime="text/csv",
    )