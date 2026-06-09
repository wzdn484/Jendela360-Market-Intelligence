#!/usr/bin/env python3
"""Professional SaaS dashboard for SPEEDHOME rental market analysis & Jendela360 benchmarking."""

from __future__ import annotations

import base64
import io
import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# Path files
BASE_DIR = os.path.dirname(__file__)
DATA_FILE = Path(BASE_DIR) / "data_properti.csv"
LOGO_FILE = Path(BASE_DIR) / "logo.jpg"

DISPLAY_COLUMNS = {
    "judul_listing": "Judul Listing",
    "nama_property_area": "Nama Property / Area",
    "tipe_kamar_jumlah_kamar_tidur": "Tipe Kamar / Kamar Tidur",
    "harga_per_bulan_rm": "Harga / Bulan (RM)",
    "harga_per_tahun_rm": "Harga / Tahun (RM)",
    "ukuran_unit_sqft": "Ukuran (sqft)",
    "status_furnitur": "Status Furnitur",
    "link_listing": "Link SPEEDHOME",
}

# --- Standard Colors ---
JENDELA_BLUE = "#00ADEF"  
DISTINCT_COLORS = ["#00ADEF", "#F58220", "#8DC63F", "#E03C31", "#8C4799"]


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------


@st.cache_data(show_spinner=False)
def load_data(csv_path: str) -> pd.DataFrame:
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    for column in ("harga_per_bulan_rm", "harga_per_tahun_rm", "ukuran_unit_sqft"):
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")
    return df


def format_rm(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"RM {value:,.0f}"


def compute_mode_price(series: pd.Series) -> float | None:
    clean = series.dropna()
    if clean.empty:
        return None
    modes = clean.mode()
    return float(modes.iloc[0]) if not modes.empty else None


def filter_dataframe(df: pd.DataFrame, query: str) -> pd.DataFrame:
    if not query.strip():
        return df

    needle = query.strip().casefold()
    search_cols = [col for col in ("nama_property_area", "judul_listing") if col in df.columns]
    if not search_cols:
        return df

    mask = pd.Series(False, index=df.index)
    for col in search_cols:
        mask = mask | df[col].astype(str).str.casefold().str.contains(needle, na=False)
    return df[mask]


def price_summary(df: pd.DataFrame) -> dict[str, float | int | None]:
    prices = df["harga_per_bulan_rm"].dropna() if "harga_per_bulan_rm" in df.columns else pd.Series(dtype=float)
    sqft = df["ukuran_unit_sqft"].dropna() if "ukuran_unit_sqft" in df.columns else pd.Series(dtype=float)

    avg_price = float(prices.mean()) if not prices.empty else None
    median_price = float(prices.median()) if not prices.empty else None
    mode_price = compute_mode_price(prices)
    fair_price = (
        (avg_price + median_price) / 2
        if avg_price is not None and median_price is not None
        else None
    )
    avg_sqft = float(sqft.mean()) if not sqft.empty else None

    return {
        "total_units": len(df),
        "avg_price": avg_price,
        "median_price": median_price,
        "mode_price": mode_price,
        "fair_price": fair_price,
        "avg_sqft": avg_sqft,
    }


def compute_price_per_sqft(df: pd.DataFrame) -> float | None:
    valid_df = df.dropna(subset=["harga_per_bulan_rm", "ukuran_unit_sqft"])
    if valid_df.empty:
        return None
    valid_df["price_per_sqft"] = valid_df["harga_per_bulan_rm"] / valid_df["ukuran_unit_sqft"]
    return float(valid_df["price_per_sqft"].median())


def build_display_frame(df: pd.DataFrame) -> pd.DataFrame:
    visible = [col for col in DISPLAY_COLUMNS if col in df.columns]
    display = df[visible].rename(columns=DISPLAY_COLUMNS).copy()

    for col in ("Harga / Bulan (RM)", "Harga / Tahun (RM)", "Ukuran (sqft)"):
        if col in display.columns:
            display[col] = display[col].apply(
                lambda v: "" if pd.isna(v) else f"{int(v):,}"
            )

    return display


def chart_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    plot_df = df.dropna(subset=["ukuran_unit_sqft", "harga_per_bulan_rm"]).copy()
    plot_df["Ukuran (sqft)"] = plot_df["ukuran_unit_sqft"]
    plot_df["Harga / Bulan (RM)"] = plot_df["harga_per_bulan_rm"]
    plot_df["Property"] = plot_df.get("nama_property_area", plot_df.get("judul_listing", ""))
    plot_df["Status Furnitur"] = plot_df.get("status_furnitur", "Unknown").fillna("Unknown")
    return plot_df


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Listings")
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# UI components
# ---------------------------------------------------------------------------


def inject_styles() -> None:
    st.markdown(
        f"""
        <style>
            #MainMenu {{visibility: hidden;}}
            footer {{visibility: hidden;}}
            header[data-testid="stHeader"] {{
                background: transparent;
            }}
            .block-container {{
                padding-top: 1rem;
                padding-bottom: 3rem;
                max-width: 1400px;
            }}
            .section-label {{
                font-size: 0.85rem;
                font-weight: 700;
                letter-spacing: 0.1em;
                text-transform: uppercase;
                color: {JENDELA_BLUE};
                margin: 0 0 1rem 0;
                border-left: 3px solid {JENDELA_BLUE};
                padding-left: 0.75rem;
            }}
            div[data-testid="stMetricValue"] {{
                color: {JENDELA_BLUE} !important;
                font-weight: 800;
                font-size: 1.6rem !important; 
                white-space: nowrap;
            }}
            .explanation-text {{
                font-size: 0.9rem;
                color: #94a3b8;
                line-height: 1.5;
                margin-bottom: 1rem;
            }}
            /* Adaptive Glassmorphic Logo Container for Dark/Light Themes */
            .adaptive-logo-box {{
                background: rgba(255, 255, 255, 0.92);
                padding: 10px 16px;
                border-radius: 10px;
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
                display: inline-flex;
                align-items: center;
                justify-content: center;
                max-width: 100%;
            }}
            .adaptive-logo-box img {{
                max-height: 45px;
                width: auto;
                object-fit: contain;
            }}
            /* Business insights custom subtext positioning */
            .fee-breakdown {{
                font-size: 0.9rem;
                color: #94a3b8;
                margin-top: -5px;
                margin-bottom: 12px;
                line-height: 1.6;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    col1, col2 = st.columns([2, 6])
    
    with col1:
        if LOGO_FILE.exists():
            try:
                with open(LOGO_FILE, "rb") as f:
                    encoded_image = base64.b64encode(f.read()).decode()
                st.markdown(
                    f'<div class="adaptive-logo-box">'
                    f'  <img src="data:image/jpeg;base64,{encoded_image}" alt="Jendela360 Logo">'
                    f'</div>',
                    unsafe_allow_html=True
                )
            except Exception:
                st.markdown(f"<div class='adaptive-logo-box'><h2 style='color:{JENDELA_BLUE}; margin:0; font-weight:800;'>Jendela360</h2></div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='adaptive-logo-box'><h2 style='color:{JENDELA_BLUE}; margin:0; font-weight:800;'>Jendela360</h2></div>", unsafe_allow_html=True)
            
    with col2:
        st.markdown(
            """
            <h1 style='margin-bottom: 0; padding-bottom: 0;'>SPEEDHOME Market Analytics</h1>
            <p style='color: #94a3b8; font-size: 1.1rem; margin-top: 5px;'>
                Dasbor analisis harga sewa dan pemetaan properti pasar kompetitor untuk tim representatif.
            </p>
            """,
            unsafe_allow_html=True,
        )
    st.markdown("<hr style='margin-top: 0.5rem; margin-bottom: 2rem; border-color: #334155;'/>", unsafe_allow_html=True)


def render_summary_metrics(summary: dict[str, float | int | None]) -> None:
    st.markdown('<p class="section-label">Ringkasan Metrik Pasar</p>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5, c6 = st.columns(6)

    c1.metric("Total Listing", f"{summary['total_units']:,}")
    c2.metric("Rata-rata Harga", format_rm(summary["avg_price"]))
    c3.metric("Median Harga", format_rm(summary["median_price"]))
    c4.metric("Modus Harga", format_rm(summary["mode_price"]))
    c5.metric("Harga Wajar (Avg/Med)", format_rm(summary["fair_price"]))
    c6.metric(
        "Rata-rata Ukuran",
        f"{summary['avg_sqft']:,.0f} sqft" if summary["avg_sqft"] is not None else "—",
    )


def render_pricing_advisor(df: pd.DataFrame) -> None:
    with st.expander("Pricing Advisor & Simulasi Cicilan Jendela360", expanded=True):
        st.markdown('<p class="section-label">Alat Estimasi Untuk Pemilik & Penyewa</p>', unsafe_allow_html=True)
        
        st.markdown(
            '<div class="explanation-text">'
            'Gunakan kalkulator ini untuk merekomendasikan harga sewa kepada <strong>Pemilik Properti</strong> berdasarkan median harga kompetitor, '
            'sekaligus mensimulasikan skema tagihan <strong>Cicilan 12 Bulan</strong> andalan Jendela360 untuk kemudahan calon <strong>Penyewa</strong>.'
            '</div>',
            unsafe_allow_html=True
        )

        median_price_per_sqft = compute_price_per_sqft(df)

        # Dua kolom input agar layout tetap rapi dan compact
        col_in1, col_in2 = st.columns(2)
        with col_in1:
            unit_size = st.number_input(
                "Masukkan Ukuran Unit Klien (sqft)",
                min_value=100,
                max_value=5000,
                value=800,
                step=50,
            )
        with col_in2:
            success_fee_pct = st.selectbox(
                "Success Fee Jendela360 (%)",
                options=[5, 8, 10, 12, 15],
                index=2,
                format_func=lambda x: f"{x}%",
                help="Persentase potongan komisi revenue sharing dari pemilik properti untuk Jendela360."
            )

        st.divider()

        if median_price_per_sqft is None:
            st.warning("Data ukuran & harga tidak cukup untuk memberikan rekomendasi.")
        else:
            # Kalkulasi Harga Sewa
            recommended_monthly_price = unit_size * median_price_per_sqft
            recommended_yearly_price = recommended_monthly_price * 12
            
            # Kalkulasi Success Fee Jendela360
            success_fee_monthly = recommended_monthly_price * (success_fee_pct / 100)
            net_owner_monthly = recommended_monthly_price - success_fee_monthly
            success_fee_yearly = success_fee_monthly * 12

            # Simulasi cicilan Jendela360 (Harga Tahunan dibagi 12 bulan flat)
            cicilan_per_bulan = recommended_yearly_price / 12

            col_owner, col_tenant = st.columns(2)
            
            with col_owner:
                st.markdown("##### 🏢 Untuk Pemilik (Owner)")
                st.metric(label="Rekomendasi Sewa Bersih (Net)", value=f"RM {net_owner_monthly:,.0f}")
                
                # Breakdown Transparan yang Terlihat Sangat Profesional
                st.markdown(
                    f"""
                    <div class="fee-breakdown">
                        • Sewa Gross: <b>RM {recommended_monthly_price:,.0f}/bln</b><br>
                        • Komisi J360 ({success_fee_pct}%): <span style='color: #ef4444; font-weight: 600;'>-RM {success_fee_monthly:,.0f}/bln</span><br>
                        • Proyeksi Profit J360 (Annual): <span style='color: #22c55e; font-weight: 600;'>RM {success_fee_yearly:,.0f}/thn</span>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
                st.caption(f"*Berdasarkan median pasar area ini: RM {median_price_per_sqft:.2f} / sqft.*")
                
            with col_tenant:
                st.markdown("##### 💳 Untuk Penyewa (Tenant)")
                st.metric(label="Cicilan 12 Bulan (J360)", value=f"RM {cicilan_per_bulan:,.0f}/bln")
                st.caption("✔️ Tanpa bayar full di depan.<br>✔️ Solusi ringan bagi penyewa.", unsafe_allow_html=True)


def render_positioning_matrix(df: pd.DataFrame) -> None:
    plot_df = chart_dataframe(df)

    st.markdown(
        '<p class="section-label">Pemetaan Harga vs. Ukuran Properti</p>',
        unsafe_allow_html=True,
    )

    if plot_df.empty:
        st.info("Tidak cukup data untuk membuat grafik sebar.")
        return

    fig = px.scatter(
        plot_df,
        x="Ukuran (sqft)",
        y="Harga / Bulan (RM)",
        color="Status Furnitur",
        hover_name="Property",
        color_discrete_sequence=DISTINCT_COLORS, 
        hover_data={
            "Ukuran (sqft)": ":,.0f",
            "Harga / Bulan (RM)": ":,.0f",
            "Status Furnitur": True,
        },
        title="Distribusi Harga Sewa Berdasarkan Ukuran Unit",
        height=480,
    )
    
    fig.update_traces(marker=dict(size=12, opacity=0.85, line=dict(width=1, color="white")))
    
    fig.update_layout(
        title=dict(font=dict(size=16), x=0),
        legend=dict(
            orientation="h", 
            yanchor="top", 
            y=-0.18, 
            xanchor="center", 
            x=0.5,
            title=dict(text="")
        ),
        margin=dict(l=10, r=10, t=50, b=60),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)", zeroline=False),
        yaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)", zeroline=False),
    )

    st.plotly_chart(fig, use_container_width=True)


def render_business_insights(df: pd.DataFrame) -> None:
    left, right = st.columns([1, 1.3], gap="large")

    with left:
        render_pricing_advisor(df)

    with right:
        render_positioning_matrix(df)


def render_listings_table(df: pd.DataFrame) -> None:
    st.markdown('<p class="section-label">Daftar Properti (Listings)</p>', unsafe_allow_html=True)

    if df.empty:
        st.warning("Tidak ada data yang cocok dengan filter.")
        return

    display_df = build_display_frame(df)
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            DISPLAY_COLUMNS["link_listing"]: st.column_config.LinkColumn(
                DISPLAY_COLUMNS["link_listing"],
                display_text="Buka Web SPEEDHOME",
                validate=r"^https?://",
            )
        },
    )


def render_export_buttons(df: pd.DataFrame) -> None:
    if df.empty:
        return

    st.markdown('<p class="section-label">Unduh Laporan Data</p>', unsafe_allow_html=True)
    export_df = df.rename(columns=DISPLAY_COLUMNS)
    col_csv, col_xlsx = st.columns(2)

    with col_csv:
        st.download_button(
            label="Unduh Format CSV",
            data=to_csv_bytes(export_df),
            file_name="Laporan_Pasar_Speedhome.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_xlsx:
        st.download_button(
            label="Unduh Format Excel",
            data=to_excel_bytes(export_df),
            file_name="Laporan_Pasar_Speedhome.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    st.set_page_config(
        layout="wide",
        page_title="SPEEDHOME Market Analytics",
        page_icon="📊",
        initial_sidebar_state="expanded",
    )
    inject_styles()
    render_header()

    if not DATA_FILE.exists():
        st.error(
            f"File Data `{DATA_FILE.name}` tidak ditemukan. "
            "Jalankan script scraping terlebih dahulu untuk mengambil data."
        )
        st.stop()

    df = load_data(str(DATA_FILE))

    with st.sidebar:
        st.markdown("### Pencarian & Filter")
        search_query = st.text_input(
            "Cari Area atau Nama Property",
            placeholder="Contoh: Mont Kiara, Sentul...",
        )
        st.divider()
        st.markdown("**Status Dataset**")
        st.write(f"Total Baris Data: **{len(df):,}**")
        st.write(f"Nama File: `{DATA_FILE.name}`")
        st.divider()
        st.info("Catatan: Data diperbarui otomatis jika file CSV hasil scraping diperbarui.")

    filtered = filter_dataframe(df, search_query)
    summary = price_summary(filtered)

    if filtered.empty:
        st.warning("Pencarian tidak ditemukan. Silakan coba kata kunci lain.")
    else:
        render_summary_metrics(summary)

    st.divider()

    if not filtered.empty:
        render_business_insights(filtered)
        st.divider()
        render_listings_table(filtered)
        st.divider()
        render_export_buttons(filtered)


if __name__ == "__main__":
    main()