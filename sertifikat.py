import streamlit as st
import pandas as pd
import fitz  # PyMuPDF untuk manipulasi PDF
import os  # Menambahkan import os untuk mengelola file dan folder
from PIL import Image
import io
import zipfile

def add_text_to_pdf(template_pdf, output_pdf, start_number, excel_path, atribut_positions, atribut_font_size, atribut_font_type, atribut_font_color, nomor_position, nomor_font_size, nomor_font_type, nomor_font_color, is_bold_atribut, is_bold_nomor):
    df = pd.read_excel(excel_path)  # Baca file Excel
    df.columns = df.columns.str.strip()  # Pastikan tidak ada spasi di header kolom
    
    if not all(atribut in df.columns for atribut in atribut_positions.keys()):
        st.error("Format Excel salah! Pastikan kolom yang dibutuhkan ada.")
        return []

    pdf_files = []
    
    # Loop untuk setiap baris (sertifikat) di data Excel
    for idx, row in df.iterrows():
        nomor_sertifikat = f"{start_number}.{str(idx+1).zfill(2)}"
        
        # Buka template PDF setiap kali untuk memastikan tidak ada data sebelumnya yang tersisa
        doc = fitz.open(template_pdf)
        page = doc[0]  # Mengambil halaman pertama template
        
        # Menambahkan nomor sertifikat dengan Bold atau tidak
        nomor_font = nomor_font_type + "B" if is_bold_nomor else nomor_font_type  # Gunakan Bold jika diinginkan
        page.insert_text((nomor_position[0], nomor_position[1]), f"{nomor_sertifikat}", 
                         fontsize=nomor_font_size, 
                         fontname=nomor_font, 
                         color=nomor_font_color)
        
        # Menambahkan atribut lainnya (misalnya, Nama, NIP, dll)
        for atribut, position in atribut_positions.items():
            text = row.get(atribut, "")
            
            # Menambahkan atribut ke PDF
            atribut_font = atribut_font_type[atribut] + "B" if is_bold_atribut[atribut] else atribut_font_type[atribut]  # Gunakan Bold jika diinginkan
            page.insert_text((position[0], position[1]), f"{text}", 
                             fontsize=atribut_font_size[atribut], 
                             fontname=atribut_font, 
                             color=atribut_font_color[atribut])
        
        # Simpan hasil PDF untuk sertifikat ini
        output_filename = f"{output_pdf}/{row['Nama']}_sertifikat.pdf"
        doc.save(output_filename)
        
        pdf_files.append(output_filename)

    return pdf_files



# Fungsi untuk menampilkan pratinjau PDF sebagai gambar
def preview_pdf_as_image(pdf_path):
    # Menggunakan PyMuPDF untuk render halaman PDF sebagai gambar
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap()  # Mendapatkan gambar dari halaman pertama
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)  # Mengonversi pixmap ke gambar
    return img

# Fungsi utama
def generate_certificates(template_pdf, excel_path, output_dir, start_number, atribut_positions, atribut_font_size, atribut_font_type, atribut_font_color, nomor_position, nomor_font_size, nomor_font_type, nomor_font_color, is_bold_atribut, is_bold_nomor):
    # Pastikan folder output ada
    os.makedirs(output_dir, exist_ok=True)  # Membuat folder jika belum ada
    
    pdf_files = add_text_to_pdf(template_pdf, output_dir, start_number, excel_path, atribut_positions, atribut_font_size, atribut_font_type, atribut_font_color, nomor_position, nomor_font_size, nomor_font_type, nomor_font_color, is_bold_atribut, is_bold_nomor)
    
    return pdf_files



# Fungsi untuk membuat ZIP file dari daftar file PDF
def create_zip(pdf_files, zip_filename):
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for pdf in pdf_files:
            zipf.write(pdf, os.path.basename(pdf))  # Menambahkan file PDF ke dalam ZIP
    return zip_filename

# Streamlit UI
st.title("Generate Sertifikat")

# Form untuk memilih atribut yang dibutuhkan
atribut_form = st.sidebar.form(key="atribut_form")
atribut_list = st.sidebar.multiselect("Pilih Atribut untuk Sertifikat", 
                                     ["Nama", "NIP", "Jabatan", "Tanggal", "Alamat"], 
                                     default=["Nama", "NIP"]) 
atribut_form.form_submit_button("Pilih Atribut")

# Upload template PDF dan Excel
uploaded_pdf = st.file_uploader("Upload Template Sertifikat (PDF)", type=["pdf"])
uploaded_excel = st.file_uploader("Upload Daftar Peserta (Excel)", type=["xlsx"])
start_number = st.text_input("Nomor Sertifikat Awal", "")

# Input untuk mengatur posisi dan font untuk nomor sertifikat
st.sidebar.subheader("Pengaturan Tata Letak Nomor Sertifikat")
nomor_x_position = st.sidebar.slider("Posisi X (Horizontal) Nomor", min_value=0, max_value=500, value=100)
nomor_y_position = st.sidebar.slider("Posisi Y (Vertical) Nomor", min_value=0, max_value=500, value=150)
nomor_font_size = st.sidebar.slider("Ukuran Font Nomor", min_value=8, max_value=30, value=12)
nomor_font_type = st.sidebar.selectbox("Jenis Font Nomor", ["helv", "times", "courier", "helvetica", "arial"])
nomor_font_color = st.sidebar.color_picker("Pilih Warna Font Nomor", "#000000")
is_bold_nomor = st.sidebar.checkbox("Bold Nomor Sertifikat", value=False)

# Input untuk mengatur posisi dan font untuk atribut lainnya
atribut_positions = {}
atribut_font_size = {}
atribut_font_type = {}
atribut_font_color = {}
is_bold_atribut = {}

for atribut in atribut_list:
    st.sidebar.subheader(f"Pengaturan Tata Letak {atribut}")
    atribut_x_position = st.sidebar.slider(f"Posisi X (Horizontal) {atribut}", min_value=0, max_value=500, value=100)
    atribut_y_position = st.sidebar.slider(f"Posisi Y (Vertical) {atribut}", min_value=0, max_value=500, value=150)
    atribut_font_size[atribut] = st.sidebar.slider(f"Ukuran Font {atribut}", min_value=8, max_value=30, value=12)
    atribut_font_type[atribut] = st.sidebar.selectbox(f"Jenis Font {atribut}", ["helv", "times", "courier", "helvetica", "arial"])
    atribut_font_color[atribut] = st.sidebar.color_picker(f"Pilih Warna Font {atribut}", "#000000")
    is_bold_atribut[atribut] = st.sidebar.checkbox(f"Bold {atribut}", value=False)
    
    atribut_positions[atribut] = (atribut_x_position, atribut_y_position)

# Menampilkan dan memproses sertifikat jika template dan excel sudah diupload
if uploaded_pdf and uploaded_excel and start_number:
    output_dir = "generated_certificates"
    
    # Pastikan folder output ada
    os.makedirs(output_dir, exist_ok=True)  # Membuat folder jika belum ada
    
    # Simpan file template PDF yang diunggah
    template_pdf = os.path.join(output_dir, "template.pdf")
    with open(template_pdf, "wb") as f:
        f.write(uploaded_pdf.read())
    
    excel_path = os.path.join(output_dir, "peserta.xlsx")
    with open(excel_path, "wb") as f:
        f.write(uploaded_excel.read())

    # Convert warna dari Hex ke RGB dengan nilai di antara 0 hingga 1
    atribut_font_color_rgb = {atribut: tuple(int(color.lstrip('#')[i:i+2], 16) / 255.0 for i in (0, 2, 4)) 
                              for atribut, color in atribut_font_color.items()}
    
    nomor_font_color_rgb = tuple(int(nomor_font_color.lstrip('#')[i:i+2], 16) / 255.0 for i in (0, 2, 4))

    # Generate sertifikat
    pdf_files = generate_certificates(template_pdf, excel_path, output_dir, start_number,
                                      atribut_positions, atribut_font_size, 
                                      atribut_font_type, atribut_font_color_rgb,
                                      (nomor_x_position, nomor_y_position), nomor_font_size, 
                                      nomor_font_type, nomor_font_color_rgb, 
                                      is_bold_atribut, is_bold_nomor)

    if pdf_files:
        st.success(f"Sertifikat berhasil dibuat sebanyak {len(pdf_files)}.")

        # Menampilkan pratinjau sertifikat
        st.subheader("Preview Sertifikat")
        preview_index = st.slider("Pilih Sertifikat untuk Preview", 0, len(pdf_files)-1, 0)
        pdf_view = pdf_files[preview_index]
        
        # Menampilkan pratinjau sertifikat
        img = preview_pdf_as_image(pdf_view)
        st.image(img, caption="Preview Sertifikat", use_column_width=True)

        # Menambahkan tombol untuk mengunduh semua sertifikat dalam bentuk ZIP
        zip_filename = os.path.join(output_dir, "sertifikat.zip")
        zip_file_path = create_zip(pdf_files, zip_filename)
        with open(zip_file_path, 'rb') as zip_file:
            st.download_button("Unduh Semua Sertifikat (ZIP)", zip_file, file_name="sertifikat.zip")
