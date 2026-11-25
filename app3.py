import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# ============================================================
# DASHBOARD SPK TIRIZ
# Framework: Streamlit
# Jalankan: python -m streamlit run app.py
# ============================================================

st.set_page_config(
    page_title="Dashboard SPK Tiriz",
    page_icon="🥛",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Palet warna
NAVY = "#17365D"
BLUE = "#2457A6"
GREEN = "#15966B"
AMBER = "#D49B27"
RED = "#B42318"
MUTED = "#6B7280"

# ============================================================
# CSS / TEMA VISUAL
# ============================================================


st.markdown("""
<style>

/* =====================================
   1. WARNA DASAR
===================================== */
:root {
    --navy: #17365D;
    --blue: #2457A6;
    --green: #15966B;
    --amber: #D49B27;
    --red: #B42318;
    --muted: #6B7280;
}

.stApp {
    background-color: #F5F7FA;
}

/* =====================================
   2. SIDEBAR
===================================== */
section[data-testid="stSidebar"] {
    background-color: #17365D !important;
}

/* Judul dan teks sidebar */
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label {
    color: #FFFFFF !important;
}

/* =====================================
   3. FILE UPLOADER
===================================== */

/* Judul widget upload tetap putih */
section[data-testid="stSidebar"]
[data-testid="stFileUploader"] > label {
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
}

/* Kotak upload berwarna putih */
section[data-testid="stSidebar"]
[data-testid="stFileUploaderDropzone"] {
    background-color: #FFFFFF !important;
    border: 1px dashed #AAB8C8 !important;
    border-radius: 10px !important;
}

/* Semua teks di dalam kotak upload */
section[data-testid="stSidebar"]
[data-testid="stFileUploaderDropzone"] * {
    color: #17365D !important;
    -webkit-text-fill-color: #17365D !important;
}

/* Teks instruksi upload */
section[data-testid="stSidebar"]
[data-testid="stFileUploaderDropzoneInstructions"],
section[data-testid="stSidebar"]
[data-testid="stFileUploaderDropzoneInstructions"] * {
    color: #17365D !important;
    -webkit-text-fill-color: #17365D !important;
}

/* Informasi ukuran file */
section[data-testid="stSidebar"]
[data-testid="stFileUploader"] small,
section[data-testid="stSidebar"]
[data-testid="stFileUploader"] small * {
    color: #17365D !important;
    -webkit-text-fill-color: #17365D !important;
}

/* Tombol Browse files */
section[data-testid="stSidebar"]
[data-testid="stFileUploaderDropzone"] button {
    background-color: #2457A6 !important;
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px !important;
}

/* =====================================
   4. JUDUL DASHBOARD
===================================== */
h1, h2, h3 {
    color: #17365D;
    font-weight: 700;
}

/* =====================================
   5. TEKS UTAMA
===================================== */
.stApp p {
    color: #374151;
}

/* =====================================
   6. METRIC CARD
===================================== */
[data-testid="stMetric"] {
    background-color: #FFFFFF;
    padding: 18px;
    border-radius: 12px;
    border: 1px solid #E5E7EB;
    box-shadow: 0 2px 6px rgba(0,0,0,0.04);
}

/* =====================================
   7. TOMBOL
===================================== */
.stButton > button {
    background-color: #2457A6;
    color: #FFFFFF;
    border: none;
    border-radius: 8px;
    padding: 8px 18px;
    font-weight: 600;
}

.stButton > button:hover {
    background-color: #17365D;
    color: #FFFFFF;
    border: none;
}

/* =====================================
   8. INPUT DAN SELECTBOX
===================================== */
.stTextInput input,
.stNumberInput input,
.stSelectbox div[data-baseweb="select"] {
    border-radius: 8px;
}

/* =====================================
   9. TABEL
===================================== */
[data-testid="stDataFrame"] {
    border: 1px solid #E5E7EB;
    border-radius: 8px;
}

/* =====================================
   10. ALERT
===================================== */
[data-testid="stAlert"] {
    border-radius: 8px;
}

/* =====================================
   11. FOOTER
===================================== */
footer {
    visibility: hidden;
}

</style>
""", unsafe_allow_html=True)



# ============================================================
# LOKASI FILE EXCEL DAN FUNGSI PEMBACAAN DATA
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = BASE_DIR / "Database_ModelBase_SPK_Tiriz_Revisi_Dosen_v3.xlsx"


@st.cache_data
def load_workbook_data(file_bytes=None):
    """Membaca data dari empat sheet Excel yang digunakan dashboard."""
    source = file_bytes if file_bytes is not None else DEFAULT_FILE

    movement = pd.read_excel(source, sheet_name="PERGERAKAN_STOK")
    usage = pd.read_excel(source, sheet_name="TRANSAKSI_PEMAKAIAN")
    purchases = pd.read_excel(source, sheet_name="TRANSAKSI_PEMBELIAN")
    parameters = pd.read_excel(source, sheet_name="MASTER_PARAMETER")

    movement["Tanggal"] = pd.to_datetime(movement["Tanggal"])
    usage["Tanggal"] = pd.to_datetime(usage["Tanggal"])
    purchases["Tanggal_Pembelian"] = pd.to_datetime(
        purchases["Tanggal_Pembelian"]
    )

    # Hitung ulang biaya agar tidak bergantung pada formula Excel.
    purchases["Total_Biaya_Rp"] = (
        purchases["Jumlah_Dibeli_L"] * purchases["Harga_per_L_Rp"]
    )

    parameter_dict = dict(
        zip(parameters["Parameter"], parameters["Nilai"])
    )

    return movement, usage, purchases, parameter_dict


def rupiah(value):
    """Format angka menjadi Rupiah."""
    return "Rp" + f"{float(value):,.0f}".replace(",", ".")


def litres(value):
    """Format angka liter dengan koma desimal Indonesia."""
    return (
        f"{float(value):,.1f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


# ============================================================
# SIDEBAR DAN PEMBACAAN EXCEL
# ============================================================

st.sidebar.markdown("## 🥛 SPK Tiriz")
st.sidebar.caption(
    "Sistem Pendukung Keputusan Pembelian Bahan Baku Susu"
)

uploaded_file = st.sidebar.file_uploader(
    "Unggah file Excel database",
    type=["xlsx"],
    help=(
        "Secara default aplikasi membaca file Excel v3 yang berada "
        "satu folder dengan app.py."
    ),
)

try:
    if uploaded_file is not None:
        movement, usage, purchases, params = load_workbook_data(
            uploaded_file.getvalue()
        )
    else:
        if not DEFAULT_FILE.exists():
            st.error(
                "File Excel database tidak ditemukan. Letakkan file "
                "'Database_ModelBase_SPK_Tiriz_Revisi_Dosen_v3.xlsx' "
                "di folder yang sama dengan app.py atau unggah melalui sidebar."
            )
            st.stop()

        movement, usage, purchases, params = load_workbook_data()

except ImportError as error:
    st.error(
        "Library untuk membaca Excel belum terpasang. "
        "Jalankan: python -m pip install openpyxl"
    )
    st.exception(error)
    st.stop()

except Exception as error:
    st.error(f"Gagal membaca file Excel: {error}")
    st.info(
        "Pastikan Excel memiliki sheet: PERGERAKAN_STOK, "
        "TRANSAKSI_PEMAKAIAN, TRANSAKSI_PEMBELIAN, dan MASTER_PARAMETER."
    )
    st.exception(error)
    st.stop()

# ============================================================
# HEADER DASHBOARD + LOGO TIRIZ
# Letakkan logo_tiriz.png satu folder dengan app.py.
# ============================================================

import base64

logo_path = BASE_DIR / "logo_tiriz.jpg"

if logo_path.exists():
    logo_b64 = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
    logo_html = (
        f'<img class="brand-logo" src="data:image/png;base64,{logo_b64}" '
        'alt="Logo Tiriz FTIP Unpad">'
    )
else:
    logo_html = (
        '<div class="brand-logo logo-fallback">'
        '<span>LOGO TIRIZ</span>'
        '</div>'
    )

st.markdown(
    f"""
    <div class="brand-header">
        {logo_html}
        <div class="brand-copy">
            <h1>Dashboard SPK Tiriz</h1>
            <p>Sistem Pendukung Keputusan Pembelian Bahan Baku Susu</p>
            <p>FTIP – Universitas Padjadjaran</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="warning-box">
        <b>DATA SIMULASI.</b> Transaksi pembelian, pemakaian, stok,
        stockout, dan kedaluwarsa pada file Excel merupakan data dummy
        untuk keperluan model. Data ini bukan catatan aktual UMKM Tiriz.
        Kapasitas 70 L dan batas anggaran Rp1.000.000 merupakan asumsi
        yang dapat diubah.
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# RINGKASAN OPERASIONAL
# ============================================================

st.markdown("### Ringkasan operasional")

total_usage = usage["Kebutuhan_Aktual_Dummy_L"].sum()
total_purchase = purchases["Jumlah_Dibeli_L"].sum()
stockout_days = int((movement["Stockout_L"] > 0).sum())
total_shortage = movement["Stockout_L"].sum()
expired_days = int((movement["Susu_Expired_L"] > 0).sum())
total_expired = movement["Susu_Expired_L"].sum()

m1, m2, m3, m4 = st.columns(4)

m1.metric(
    "Total pemakaian",
    f"{litres(total_usage)} L",
)

m2.metric(
    "Total pembelian",
    f"{litres(total_purchase)} L",
)

m3.metric(
    "Hari stockout",
    f"{stockout_days} hari",
    f"{litres(total_shortage)} L tidak terpenuhi",
)

m4.metric(
    "Susu kedaluwarsa",
    f"{litres(total_expired)} L",
    f"{expired_days} hari terjadi kedaluwarsa",
)

# ============================================================
# KALKULATOR KEPUTUSAN PEMBELIAN
# ============================================================

st.markdown("### Kalkulator keputusan pembelian")
st.caption(
    "Ubah parameter di bawah ini. Rekomendasi dan estimasi biaya "
    "akan dihitung ulang secara langsung."
)

last_three = usage["Kebutuhan_Aktual_Dummy_L"].tail(3)

default_demand = (
    float(last_three.mean()) if len(last_three) else 25.0
)
default_period = int(params.get("Periode pembelian", 2))
default_capacity = float(params.get("Kapasitas penyimpanan", 70))
default_budget = float(
    params.get("Batas anggaran pembelian", 1_000_000)
)
default_price = (
    float(purchases["Harga_per_L_Rp"].iloc[-1])
    if not purchases.empty else 14_000.0
)
default_stock = (
    float(movement["Stok_Akhir_L"].iloc[-1])
    if not movement.empty else 0.0
)

with st.container(border=True):
    c1, c2, c3 = st.columns(3)

    with c1:
        D = st.number_input(
            "D — Forecast kebutuhan harian (L/hari)",
            min_value=0.0,
            value=round(default_demand, 1),
            step=0.5,
        )
        P = st.number_input(
            "P — Periode pembelian (hari)",
            min_value=1,
            value=default_period,
            step=1,
        )

    with c2:
        S = st.number_input(
            "S — Stok layak pakai (L)",
            min_value=0.0,
            value=round(default_stock, 1),
            step=1.0,
        )
        C = st.number_input(
            "C — Kapasitas penyimpanan (L)",
            min_value=0.0,
            value=default_capacity,
            step=1.0,
        )

    with c3:
        H = st.number_input(
            "H — Harga susu (Rp/L)",
            min_value=0,
            value=int(default_price),
            step=500,
        )
        budget = st.number_input(
            "Batas anggaran owner (Rp)",
            min_value=0,
            value=int(default_budget),
            step=50_000,
        )

# Rumus Model Base sesuai revisi dosen
K = D * P
Q = max(0.0, K - S)
X = max(0.0, C - S)
recommendation = min(Q, X)
B = recommendation * H
remaining_need = max(0.0, Q - recommendation)

r1, r2, r3, r4, r5 = st.columns(5)

r1.metric("K — Kebutuhan siklus", f"{litres(K)} L")
r2.metric("Q — Kebutuhan beli", f"{litres(Q)} L")
r3.metric("X — Sisa kapasitas", f"{litres(X)} L")
r4.metric("Rekomendasi beli", f"{litres(recommendation)} L")
r5.metric("B — Estimasi biaya", rupiah(B))

if Q == 0:
    st.success(
        "Stok layak pakai sudah cukup untuk memenuhi kebutuhan selama "
        "satu siklus. Rekomendasi: tidak perlu membeli."
    )
elif Q > X:
    st.warning(
        f"Kapasitas membatasi pembelian. Kebutuhan pembelian "
        f"{litres(Q)} L, tetapi sisa kapasitas hanya {litres(X)} L. "
        f"Masih ada kebutuhan {litres(remaining_need)} L "
        "yang perlu diputuskan owner."
    )
else:
    st.success(
        f"Rekomendasi pembelian {litres(recommendation)} L dapat "
        "memenuhi kebutuhan siklus dan masih sesuai kapasitas penyimpanan."
    )

if B > budget:
    st.error(
        f"Estimasi biaya {rupiah(B)} melebihi batas anggaran "
        f"{rupiah(budget)}. Model memberi peringatan; jumlah beli "
        "tidak otomatis dikurangi. Keputusan akhir mengikuti kebijakan owner."
    )
else:
    st.info(
        f"Estimasi biaya {rupiah(B)} masih berada dalam batas "
        f"anggaran {rupiah(budget)}."
    )

with st.expander("Lihat rumus Model Base"):
    st.latex(r"K = D \times P")
    st.latex(r"Q = \max(0, K-S)")
    st.latex(r"X = \max(0, C-S)")
    st.latex(r"Q_{\text{rekomendasi}} = \min(Q,X)")
    st.latex(r"B = Q_{\text{rekomendasi}} \times H")

    st.caption(
        "Harga memengaruhi estimasi biaya, bukan jumlah liter rekomendasi. "
        "Batas anggaran menjadi peringatan keputusan owner."
    )

# ============================================================
# GRAFIK PERGERAKAN STOK DAN PEMAKAIAN
# ============================================================

st.markdown("### Pergerakan stok dan pemakaian")

period_choice = st.radio(
    "Rentang grafik",
    ["Seluruh periode", "14 hari terakhir"],
    horizontal=True,
    label_visibility="collapsed",
)

chart_movement = movement.copy()

if period_choice == "14 hari terakhir":
    chart_movement = chart_movement.tail(14)

left, right = st.columns(2)

with left:
    fig_stock = px.line(
        chart_movement,
        x="Tanggal",
        y=["Kebutuhan_L", "Stok_Akhir_L"],
        markers=True,
        labels={
            "value": "Liter",
            "Tanggal": "Tanggal",
            "variable": "Data",
            "Kebutuhan_L": "Kebutuhan",
            "Stok_Akhir_L": "Stok akhir",
        },
        color_discrete_sequence=[BLUE, GREEN],
        title="Kebutuhan dan stok akhir harian",
    )

    fig_stock.update_layout(
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="#344054"),
        legend_title_text="",
        margin=dict(l=10, r=10, t=55, b=10),
        yaxis_title="Liter",
        xaxis_title="Tanggal",
    )

    fig_stock.update_xaxes(
        showgrid=False,
        tickfont=dict(color="#667085"),
    )
    fig_stock.update_yaxes(
        gridcolor="#E5E9F0",
        tickfont=dict(color="#667085"),
    )

    st.plotly_chart(fig_stock, use_container_width=True)

with right:
    fig_events = px.bar(
        chart_movement,
        x="Tanggal",
        y=["Susu_Masuk_L", "Susu_Expired_L"],
        barmode="group",
        labels={
            "value": "Liter",
            "Tanggal": "Tanggal",
            "variable": "Data",
            "Susu_Masuk_L": "Susu masuk",
            "Susu_Expired_L": "Susu kedaluwarsa",
        },
        color_discrete_sequence=["#6B8FC9", AMBER],
        title="Pembelian masuk dan susu kedaluwarsa",
    )

    fig_events.update_layout(
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="#344054"),
        legend_title_text="",
        margin=dict(l=10, r=10, t=55, b=10),
        yaxis_title="Liter",
        xaxis_title="Tanggal",
    )

    fig_events.update_xaxes(
        showgrid=False,
        tickfont=dict(color="#667085"),
    )
    fig_events.update_yaxes(
        gridcolor="#E5E9F0",
        tickfont=dict(color="#667085"),
    )

    st.plotly_chart(fig_events, use_container_width=True)

# ============================================================
# RIWAYAT OPERASIONAL
# ============================================================

st.markdown("### Riwayat operasional")

tab1, tab2, tab3 = st.tabs(
    ["Pergerakan stok", "Pembelian", "Pemakaian"]
)

with tab1:
    stock_view = movement.copy()
    stock_view["Tanggal"] = stock_view["Tanggal"].dt.strftime(
        "%d-%m-%Y"
    )

    st.dataframe(
        stock_view[
            [
                "Tanggal",
                "Stok_Awal_L",
                "Susu_Masuk_L",
                "Kebutuhan_L",
                "Susu_Terpenuhi_L",
                "Stockout_L",
                "Susu_Expired_L",
                "Stok_Akhir_L",
                "Status",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Unduh data pergerakan stok (CSV)",
        stock_view.to_csv(index=False).encode("utf-8-sig"),
        file_name="pergerakan_stok_tiriz.csv",
        mime="text/csv",
    )

with tab2:
    purchase_view = purchases.copy()
    purchase_view["Tanggal_Pembelian"] = (
        purchase_view["Tanggal_Pembelian"].dt.strftime("%d-%m-%Y")
    )

    st.dataframe(
        purchase_view,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Unduh data pembelian (CSV)",
        purchase_view.to_csv(index=False).encode("utf-8-sig"),
        file_name="transaksi_pembelian_tiriz.csv",
        mime="text/csv",
    )

with tab3:
    usage_view = usage.copy()
    usage_view["Tanggal"] = usage_view["Tanggal"].dt.strftime(
        "%d-%m-%Y"
    )

    st.dataframe(
        usage_view,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Unduh data pemakaian (CSV)",
        usage_view.to_csv(index=False).encode("utf-8-sig"),
        file_name="transaksi_pemakaian_tiriz.csv",
        mime="text/csv",
    )

# ============================================================
# FOOTER
# ============================================================

st.markdown("---")
st.caption(
    "Dashboard prototipe SPK Tiriz | Data simulasi dari Excel revisi dosen | "
    "Perhitungan interaktif berjalan di aplikasi, belum tersambung ke database online."
)
