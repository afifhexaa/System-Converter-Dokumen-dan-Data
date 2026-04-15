import streamlit as st
import os
import io
from PIL import Image
from PyPDF2 import PdfMerger
import fitz  # PyMuPDF for PDF compression
import pandas as pd
from docx import Document
import tempfile
from docx2pdf import convert as docx2pdf_convert
from streamlit_sortables import sort_items
import chardet  # For encoding detection

# ==========================
# Helper Functions
# ==========================
def compress_pdf(input_file, quality_percent):
    input_bytes = input_file.read()
    doc = fitz.open(stream=input_bytes, filetype="pdf")
    output_pdf = fitz.open()

    scale = quality_percent / 100.0
    matrix = fitz.Matrix(scale, scale)

    for page in doc:
        pix = page.get_pixmap(matrix=matrix)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=int(quality_percent))

        img_pdf = fitz.open()
        rect = fitz.Rect(0, 0, pix.width, pix.height)
        page_img = img_pdf.new_page(width=pix.width, height=pix.height)
        page_img.insert_image(rect, stream=img_byte_arr.getvalue())

        output_pdf.insert_pdf(img_pdf)
        img_pdf.close()

    compressed_stream = io.BytesIO()
    output_pdf.save(compressed_stream)
    output_pdf.close()
    compressed_stream.seek(0)
    return compressed_stream

def compress_image(image_file, quality_percent):
    img = Image.open(image_file)
    img_format = img.format
    output = io.BytesIO()
    img.save(output, format=img_format, optimize=True, quality=int(quality_percent))
    output.seek(0)
    return output

def convert_images_to_pdf(image_files, image_order):
    image_list = [Image.open(image_files[i]) for i in image_order]
    pdf_bytes = io.BytesIO()
    image_list[0].save(pdf_bytes, format="PDF", save_all=True, append_images=image_list[1:])
    return pdf_bytes

def read_table_file(input_file):
    if input_file.name.endswith(('.xlsx', '.xls')):
        return pd.read_excel(input_file)
    else:
        raw_bytes = input_file.read()
        result = chardet.detect(raw_bytes)
        encoding = result['encoding'] or 'utf-8'
        try:
            return pd.read_csv(io.StringIO(raw_bytes.decode(encoding)))
        except Exception as e:
            st.error(f"Gagal membaca file CSV. Error: {e}")
            st.stop()

def convert_table(input_file, target_format):
    df = read_table_file(input_file)
    st.subheader("Preview Data (5 baris pertama)")
    st.dataframe(df.head())
    output = io.BytesIO()
    if target_format == 'csv':
        df.to_csv(output, index=False)
    elif target_format == 'xlsx':
        df.to_excel(output, index=False, engine='openpyxl')
    elif target_format == 'xls':
        df.to_excel(output, index=False, engine='xlwt')
    output.seek(0)
    return output

def convert_docx_to_pdf(docx_file):
    with tempfile.TemporaryDirectory() as tmpdirname:
        input_path = os.path.join(tmpdirname, "input.docx")
        output_path = os.path.join(tmpdirname, "output.pdf")
        with open(input_path, "wb") as f:
            f.write(docx_file.read())
        docx2pdf_convert(input_path, output_path)
        with open(output_path, "rb") as f:
            return io.BytesIO(f.read())

# ==========================
# Streamlit UI
# ==========================
st.set_page_config(page_title="File Toolkit", layout="wide")
st.title("📦 File Toolkit: Compress & Convert")

menu = st.sidebar.selectbox("Pilih Fitur", [
    "📉 Kompresi File PDF", 
    "🖼️ Kompresi Gambar",
    "🖼️ Konversi Gambar ke PDF", 
    "📊 Konversi Tabel (Excel/CSV)",
    "📄 Konversi Word ke PDF",
    "ℹ️ Tentang Aplikasi"
])

# ============= Kompres PDF =============
if menu == "📉 Kompresi File PDF":
    st.header("Kompresi PDF")
    pdf_files = st.file_uploader("Unggah file PDF", type="pdf", accept_multiple_files=True)
    quality = st.slider("Persentase Kompresi", min_value=5, max_value=95, value=50)

    if pdf_files:
        for i, pdf_file in enumerate(pdf_files):
            original_size = len(pdf_file.getvalue()) / 1024
            st.write(f"**{pdf_file.name}** - Ukuran asli: {original_size:.2f} KB")

            with st.spinner(f"Mengompresi {pdf_file.name}..."):
                compressed = compress_pdf(pdf_file, quality)
            compressed_size = len(compressed.getvalue()) / 1024
            st.write(f"Ukuran setelah kompresi: {compressed_size:.2f} KB")

            file_name = st.text_input(f"Nama file output untuk {pdf_file.name} (tanpa .pdf)", value=f"compressed_{i}")
            st.download_button("Unduh File", data=compressed, file_name=f"{file_name}.pdf", key=f"download_pdf_{i}")

# ============= Kompres Gambar =============
elif menu == "🖼️ Kompresi Gambar":
    st.header("Kompresi Gambar")
    image_files = st.file_uploader("Unggah Gambar", type=["jpg", "jpeg", "png", "tiff"], accept_multiple_files=True)
    quality = st.slider("Persentase Kompresi Gambar", min_value=5, max_value=95, value=50)

    if image_files:
        for i, img_file in enumerate(image_files):
            original_size = len(img_file.getvalue()) / 1024
            st.write(f"**{img_file.name}** - Ukuran asli: {original_size:.2f} KB")

            with st.spinner(f"Mengompresi {img_file.name}..."):
                compressed = compress_image(img_file, quality)
            compressed_size = len(compressed.getvalue()) / 1024
            st.write(f"Ukuran setelah kompresi: {compressed_size:.2f} KB")

            file_name = st.text_input(f"Nama file output untuk {img_file.name} (tanpa ekstensi)", value=f"compressed_image_{i}")
            ext = img_file.name.split(".")[-1]
            st.download_button("Unduh Gambar", data=compressed, file_name=f"{file_name}.{ext}", key=f"download_img_{i}")

# ============= Gambar ke PDF =============
elif menu == "🖼️ Konversi Gambar ke PDF":
    st.header("Gambar ke PDF")
    uploaded_images = st.file_uploader("Unggah Gambar", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if uploaded_images:
        st.write("**Urutkan gambar (drag untuk atur posisi):**")
        st.markdown("Gunakan urutan untuk menentukan nomor halaman di PDF (Halaman 1 = gambar pertama)")

        previews = []
        for img_file in uploaded_images:
            previews.append(f"🖼️ {img_file.name}")

        sorted_labels = sort_items(previews, direction="horizontal")
        idx_order = [previews.index(label) for label in sorted_labels]

        st.write("**Preview Urutan Halaman:**")
        cols = st.columns(len(idx_order))
        for display_num, (idx, col) in enumerate(zip(idx_order, cols), start=1):
            with col:
                st.image(uploaded_images[idx], caption=f"Halaman {display_num}\n{uploaded_images[idx].name}", width=100)

        pdf_bytes = convert_images_to_pdf(uploaded_images, idx_order)
        file_name = st.text_input("Nama file output (tanpa .pdf)", value="output_pdf")
        st.download_button("Unduh PDF", data=pdf_bytes, file_name=f"{file_name}.pdf")

# ============= Konversi Tabel =============
elif menu == "📊 Konversi Tabel (Excel/CSV)":
    st.header("Konversi Excel/CSV")
    data_file = st.file_uploader("Unggah file Excel atau CSV", type=["xlsx", "xls", "csv"])
    target_format = st.selectbox("Konversi ke format", options=["csv", "xlsx", "xls"])

    if data_file:
        df = read_table_file(data_file)

        st.subheader("📋 Ringkasan Data")
        st.markdown(f"Jumlah record: **{len(df)}**")
        st.markdown(f"Jumlah kolom: **{len(df.columns)}**")
        st.markdown("**Daftar Kolom:**")
        st.write(list(df.columns))

        st.subheader("📌 Kolom dengan Nilai Kosong")
        null_counts = df.isnull().sum()
        empty_cols = null_counts[null_counts > 0]
        if not empty_cols.empty:
            st.write(empty_cols)
        else:
            st.info("Tidak ada kolom dengan data kosong.")

        st.subheader("Preview Data (5 baris pertama)")
        st.dataframe(df.head())

        output = io.BytesIO()
        if target_format == 'csv':
            df.to_csv(output, index=False)
        elif target_format == 'xlsx':
            df.to_excel(output, index=False, engine='openpyxl')
        elif target_format == 'xls':
            df.to_excel(output, index=False, engine='xlwt')
        output.seek(0)

        file_name = st.text_input("Nama file output", value="converted")
        st.download_button("Unduh File", data=output, file_name=f"{file_name}.{target_format}")

# ============= Konversi Word ke PDF =============
elif menu == "📄 Konversi Word ke PDF":
    st.header("Konversi Word ke PDF")
    word_file = st.file_uploader("Unggah file Word", type=["docx"])

    if word_file:
        with st.spinner("Mengonversi ke PDF..."):
            output = convert_docx_to_pdf(word_file)
        file_name = st.text_input("Nama file output", value="converted_pdf")
        st.download_button("Unduh File", data=output, file_name=f"{file_name}.pdf")

# ============= Tentang =============
elif menu == "ℹ️ Tentang Aplikasi":
    st.markdown("""
    ### 📦 File Toolkit
    Aplikasi ini memungkinkan Anda:
    - Mengompresi file PDF dengan tingkat yang dapat disesuaikan (bisa banyak file sekaligus)
    - Mengompresi gambar JPG, PNG, TIFF (multi-file)
    - Menggabungkan gambar menjadi satu PDF dengan urutan drag & drop
    - Mengkonversi file Excel/CSV antar format
    - Mengonversi dokumen Word (DOCX) ke PDF
    
    Made By M Afif:)


    """)

    # #### 🔧 Teknologi:
    # - Streamlit
    # - Pillow, pandas, PyMuPDF, PyPDF2, docx2pdf
    # - streamlit-sortables untuk drag & drop
    # - Engine Excel: openpyxl, xlwt

    # Dibuat untuk efisiensi, kecepatan, dan kemudahan penggunaan.