"""
app.py
Aplikasi Streamlit: GENERATE LABEL BARANG
Otomatisasi pembuatan label barang dari file Excel ke dokumen Word (.docx)
berdasarkan template referensi 'Label New Bold.docx'.
"""

import io
import os
import re
import streamlit as st
import pandas as pd

from label_generator import (
    read_excel_data,
    filter_by_range,
    filter_by_manual_list,
    generate_word_document,
)

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="GENERATE LABEL BARANG",
    page_icon="🏷️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Custom CSS untuk tampilan bersih dan profesional
st.markdown(
    """
    <style>
    .main-header {
        text-align: center;
        padding: 1.5rem 0 1rem 0;
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        border-radius: 12px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        color: white !important;
        font-weight: 700;
        letter-spacing: 1px;
        margin-bottom: 0.3rem;
    }
    .main-header p {
        color: #e0e6ed;
        font-size: 1.05rem;
        margin-bottom: 0;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.6rem 1.5rem;
        transition: all 0.2s ease;
    }
    .generate-btn button {
        background-color: #0284c7 !important;
        color: white !important;
        font-size: 1.1rem !important;
        border: none !important;
        box-shadow: 0 4px 10px rgba(2, 132, 199, 0.3);
    }
    .generate-btn button:hover {
        background-color: #0369a1 !important;
        box-shadow: 0 6px 14px rgba(2, 132, 199, 0.4);
    }
    .download-btn button {
        background-color: #16a34a !important;
        color: white !important;
        font-size: 1.15rem !important;
        font-weight: 700 !important;
        border: none !important;
        box-shadow: 0 4px 10px rgba(22, 163, 74, 0.3);
    }
    .download-btn button:hover {
        background-color: #15803d !important;
        box-shadow: 0 6px 14px rgba(22, 163, 74, 0.4);
    }
    .item-hint {
        font-size: 0.9rem;
        color: #1e293b;
        margin-top: 6px;
        padding: 6px 10px;
        background-color: #f1f5f9;
        border-left: 4px solid #0284c7;
        border-radius: 4px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Header Aplikasi
st.markdown(
    """
    <div class="main-header">
        <h1>🏷️ GENERATE LABEL BARANG</h1>
        <p>Aplikasi otomatisasi pembuatan label barang dari file Excel ke dokumen Word (.docx) siap cetak</p>
    </div>
    """,
    unsafe_allow_html=True,
)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates", "Label New Bold.docx")

# Inisialisasi session state
if "excel_data" not in st.session_state:
    st.session_state.excel_data = None
if "uploaded_file_name" not in st.session_state:
    st.session_state.uploaded_file_name = None
if "generated_doc" not in st.session_state:
    st.session_state.generated_doc = None
if "generated_file_name" not in st.session_state:
    st.session_state.generated_file_name = None
if "skip_hidden" not in st.session_state:
    st.session_state.skip_hidden = True

# Bagian 1: Upload File Excel
st.subheader("1. 📂 Upload File Excel")

col_upload, col_sample = st.columns([3, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "Pilih file Excel daftar barang (.xlsx)",
        type=["xlsx"],
        help="Upload file Excel dengan kolom No, List Item, dan Spec.",
    )

with col_sample:
    sample_file_path = os.path.join(os.path.dirname(__file__), "sample_data", "LIST BARANG.xlsx")
    if os.path.exists(sample_file_path):
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if st.button("📁 Gunakan Sample Data", help="Gunakan file referensi LIST BARANG.xlsx bawaan"):
            with open(sample_file_path, "rb") as f:
                sample_bytes = io.BytesIO(f.read())
                df, err = read_excel_data(sample_bytes, skip_hidden_rows=st.session_state.skip_hidden)
                if err:
                    st.error(err)
                else:
                    st.session_state.excel_data = df
                    st.session_state.uploaded_file_name = "LIST BARANG.xlsx (Sample)"
                    st.session_state.generated_doc = None
                    st.rerun()

# Opsi baris tersembunyi
col_opt, _ = st.columns([2, 2])
with col_opt:
    skip_hidden_toggle = st.checkbox(
        "Abaikan baris tersembunyi (hidden rows) di Excel",
        value=st.session_state.skip_hidden,
        help="Jika dicentang, baris yang di-hide/disembunyikan di Excel tidak akan dimuat.",
    )
    if skip_hidden_toggle != st.session_state.skip_hidden:
        st.session_state.skip_hidden = skip_hidden_toggle
        if st.session_state.uploaded_file_name:
            if "Sample" in st.session_state.uploaded_file_name:
                with open(sample_file_path, "rb") as f:
                    df, err = read_excel_data(io.BytesIO(f.read()), skip_hidden_rows=st.session_state.skip_hidden)
                    st.session_state.excel_data = df
                    st.rerun()

# Proses file yang diupload user
if uploaded_file is not None:
    if st.session_state.uploaded_file_name != uploaded_file.name:
        file_bytes = io.BytesIO(uploaded_file.getvalue())
        df, err = read_excel_data(file_bytes, skip_hidden_rows=st.session_state.skip_hidden)
        if err:
            st.error(f"❌ {err}")
            st.session_state.excel_data = None
            st.session_state.uploaded_file_name = None
        else:
            st.session_state.excel_data = df
            st.session_state.uploaded_file_name = uploaded_file.name
            st.session_state.generated_doc = None
            st.rerun()

# Jika belum ada data
if st.session_state.excel_data is None:
    st.info("ℹ️ Silakan upload file Excel daftar barang (.xlsx) untuk memulai.")
    st.stop()

# Informasi File yang Berhasil Dibaca
df_all = st.session_state.excel_data
total_items = len(df_all)
min_no = int(df_all["No"].min())
max_no = int(df_all["No"].max())

col_info1, col_info2 = st.columns(2)
with col_info1:
    st.success(f"📄 **File:** `{st.session_state.uploaded_file_name}`")
with col_info2:
    st.success(f"📦 **Jumlah Data Barang Terbaca:** `{total_items}` item (No. `{min_no}` s/d `{max_no}`)")

# Tampilkan tabel seluruh barang
with st.expander("📋 Lihat Seluruh Daftar Barang dari Excel", expanded=False):
    st.dataframe(
        df_all[["No", "List Item", "Spec"]],
        use_container_width=True,
        hide_index=True,
    )

st.markdown("---")

# Bagian 2: Pemilihan Nomor Barang
st.subheader("2. 🎯 Pemilihan Nomor Barang")

method = st.radio(
    "Pilih metode penentuan nomor barang:",
    options=["Metode A: Rentang Nomor", "Metode B: Pilih Nomor Manual"],
    horizontal=True,
)

selected_df = pd.DataFrame()
missing_numbers = []
filter_error = None
selection_label = ""
suggested_filename = "Label_Barang.docx"

if method == "Metode A: Rentang Nomor":
    col_range1, col_range2 = st.columns(2)
    with col_range1:
        # Nilai default nomor awal sesuai nomor pertama yang ada di Excel
        start_val = min_no
        start_no = st.number_input(
            "Nomor Awal (sesuai kolom No di Excel):",
            min_value=min_no,
            max_value=max_no,
            value=start_val,
            step=1,
            help="Masukkan nomor barang pertama yang ingin dicetak.",
        )
        # Tampilkan nama barang awal secara real-time
        matching_start = df_all[df_all["No"] == int(start_no)]
        if not matching_start.empty:
            item_name_start = matching_start.iloc[0]["List Item"]
            spec_start = matching_start.iloc[0]["Spec"]
            desc_start = f"{item_name_start} ({spec_start})" if spec_start and item_name_start else (item_name_start if item_name_start else spec_start)
            st.markdown(f"<div class='item-hint'>📍 <b>Barang Awal:</b> [No. {int(start_no)}] {desc_start}</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='item-hint'>⚠️ Nomor {int(start_no)} tidak ditemukan di Excel</div>", unsafe_allow_html=True)

    with col_range2:
        end_default = min(max_no, int(start_no) + 10)
        end_no = st.number_input(
            "Nomor Akhir (sesuai kolom No di Excel):",
            min_value=min_no,
            max_value=max_no,
            value=end_default,
            step=1,
            help="Masukkan nomor barang terakhir yang ingin dicetak.",
        )
        matching_end = df_all[df_all["No"] == int(end_no)]
        if not matching_end.empty:
            item_name_end = matching_end.iloc[-1]["List Item"]
            spec_end = matching_end.iloc[-1]["Spec"]
            desc_end = f"{item_name_end} ({spec_end})" if spec_end and item_name_end else (item_name_end if item_name_end else spec_end)
            st.markdown(f"<div class='item-hint'>📍 <b>Barang Akhir:</b> [No. {int(end_no)}] {desc_end}</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='item-hint'>⚠️ Nomor {int(end_no)} tidak ditemukan di Excel</div>", unsafe_allow_html=True)

    # Validasi rentang nomor
    selected_df, missing_numbers, filter_error = filter_by_range(df_all, int(start_no), int(end_no))
    selection_label = f"{int(start_no)}–{int(end_no)}"
    suggested_filename = f"Label_Barang_{int(start_no)}-{int(end_no)}.docx"

else:
    # Metode B: Pilih Nomor Manual
    st.markdown(
        "Masukkan nomor barang yang ingin dicetak (pisahkan dengan koma atau spasi), "
        "atau pilih langsung dari daftar dropdown:"
    )

    col_input, col_preset = st.columns([3, 1])
    with col_input:
        manual_input_str = st.text_input(
            "Ketik nomor barang manual:",
            placeholder="Contoh: 90, 91, 92, 101, 105",
            help="Gunakan tanda koma (,) atau spasi untuk memisahkan beberapa nomor barang.",
        )
    with col_preset:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        quick_select = st.button("Isi Contoh (90, 91, 92, 93, 94)")
        if quick_select:
            manual_input_str = "90, 91, 92, 93, 94"

    all_available_nos = df_all["No"].drop_duplicates().tolist()
    multiselect_nos = st.multiselect(
        "Atau pilih nomor barang dari daftar dropdown:",
        options=all_available_nos,
        default=[],
        help="Anda dapat memilih beberapa nomor barang secara langsung.",
    )

    parsed_nos = []
    if manual_input_str.strip():
        tokens = re.findall(r"\d+", manual_input_str)
        parsed_nos.extend([int(t) for t in tokens])

    if multiselect_nos:
        parsed_nos.extend(multiselect_nos)

    seen = set()
    unique_selected_nos = []
    for n in parsed_nos:
        if n not in seen:
            seen.add(n)
            unique_selected_nos.append(n)

    if not unique_selected_nos:
        st.warning("⚠️ Silakan masukkan atau pilih setidaknya satu nomor barang.")
    else:
        selected_df, missing_numbers, filter_error = filter_by_manual_list(df_all, unique_selected_nos)
        selection_label = ", ".join(map(str, unique_selected_nos))
        suggested_filename = "Label_Barang_Custom.docx"

# Tampilkan error validasi jika ada
if filter_error:
    st.error(f"❌ {filter_error}")
    st.stop()

# Tampilkan peringatan jika ada nomor yang tidak ditemukan
if missing_numbers:
    st.warning(
        f"⚠️ Perhatian: Nomor berikut tidak ditemukan dalam data Excel: "
        f"`{', '.join(map(str, missing_numbers))}`"
    )

if selected_df.empty:
    st.info("ℹ️ Tidak ada barang yang dipilih untuk dibuatkan label.")
    st.stop()

st.markdown("---")

# Bagian 3: Preview Sebelum Generate
st.subheader("3. 👁️ Preview Data Sebelum Generate")

count_label = len(selected_df)
pages_needed = (count_label + 29) // 30

col_prev_info1, col_prev_info2, col_prev_info3 = st.columns(3)
with col_prev_info1:
    st.markdown(f"**Nomor yang dipilih:** `{selection_label}`")
with col_prev_info2:
    st.markdown(f"**Jumlah label:** `{count_label}` label")
with col_prev_info3:
    st.markdown(f"**Estimasi Halaman Word:** `{pages_needed}` halaman (30 label/hlm)")

# Tampilkan tabel preview dengan tinggi dinamis agar SEMUA baris terlihat tanpa terpotong
table_height = min(900, max(160, (count_label + 1) * 36 + 10))
st.dataframe(
    selected_df[["No", "List Item", "Spec"]],
    use_container_width=True,
    hide_index=True,
    height=table_height,
)

st.markdown("---")

# Bagian 4: Tombol Generate dan Download Word
st.subheader("4. 🖨️ Generate Dokumen Word")

if not os.path.exists(TEMPLATE_PATH):
    st.error(
        f"❌ File template referensi tidak ditemukan di `{TEMPLATE_PATH}`. "
        "Pastikan file 'Label New Bold.docx' berada di folder 'templates/'."
    )
    st.stop()

col_btn_generate, col_btn_download = st.columns([1, 1])

with col_btn_generate:
    st.markdown('<div class="generate-btn">', unsafe_allow_html=True)
    generate_clicked = st.button("🚀 GENERATE WORD", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

if generate_clicked:
    with st.spinner("Sedang memproses dan menyusun label ke dokumen Word..."):
        try:
            items_list = selected_df.to_dict(orient="records")
            doc_stream = generate_word_document(items_list, template_path=TEMPLATE_PATH)
            st.session_state.generated_doc = doc_stream.getvalue()
            st.session_state.generated_file_name = suggested_filename
            st.success(f"✅ Dokumen Word berhasil dibuat: **{suggested_filename}** ({count_label} label)")
        except Exception as e:
            st.error(f"❌ Gagal membuat dokumen Word. Detail: {str(e)}")

# Tampilkan tombol download jika dokumen sudah digenerate
if st.session_state.generated_doc is not None:
    with col_btn_download:
        st.markdown('<div class="download-btn">', unsafe_allow_html=True)
        st.download_button(
            label=f"📥 DOWNLOAD WORD ({st.session_state.generated_file_name})",
            data=st.session_state.generated_doc,
            file_name=st.session_state.generated_file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
