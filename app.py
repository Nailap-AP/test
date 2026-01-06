# nanote_app.py
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import io
from docx import Document
from docx.shared import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
import base64
import tempfile
import os

# Konfigurasi halaman
st.set_page_config(
    page_title="NaNote - Catatan Praktikum & Kalkulator PSA",
    page_icon="🔬",
    layout="wide"
)

# CSS kustom
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.8rem;
        color: #1E40AF;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .section-box {
        background-color: #F0F9FF;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #3B82F6;
        margin-bottom: 1.5rem;
    }
    .result-box {
        background-color: #F0FDF4;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #10B981;
        margin-bottom: 1.5rem;
    }
    .stButton button {
        background-color: #3B82F6;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 0.5rem 2rem;
    }
    .stButton button:hover {
        background-color: #2563EB;
    }
</style>
""", unsafe_allow_html=True)

# Fungsi untuk menghitung PSA
def hitung_psa(pdi_values, vol_values, diameter_values):
    """
    Menghitung hasil PSA dari data PDI, %vol, dan diameter
    
    Parameters:
    pdi_values: list of PDI values
    vol_values: list of volume percentage values
    diameter_values: list of diameter values (nm)
    
    Returns:
    DataFrame dengan hasil perhitungan
    """
    if len(pdi_values) != len(vol_values) or len(vol_values) != len(diameter_values):
        st.error("Jumlah data PDI, %vol, dan diameter harus sama!")
        return None
    
    # Buat DataFrame untuk perhitungan
    df = pd.DataFrame({
        'No': range(1, len(pdi_values) + 1),
        'PDI': pdi_values,
        '%Vol': vol_values,
        'Diameter (nm)': diameter_values
    })
    
    # Hitung berat berdasarkan %vol (asumsi densitas sama)
    df['Berat Relatif'] = df['%Vol'] / 100
    
    # Hitung diameter rata-rata berbobot
    df['Diameter x Berat'] = df['Diameter (nm)'] * df['Berat Relatif']
    diameter_rata_rata = df['Diameter x Berat'].sum()
    
    # Hitung PDI rata-rata berbobot
    df['PDI x Berat'] = df['PDI'] * df['Berat Relatif']
    pdi_rata_rata = df['PDI x Berat'].sum()
    
    # Hitung distribusi ukuran
    df['Kategori'] = pd.cut(df['Diameter (nm)'], 
                           bins=[0, 10, 50, 100, 500, 1000, np.inf],
                           labels=['<10 nm', '10-50 nm', '50-100 nm', 
                                  '100-500 nm', '500-1000 nm', '>1000 nm'])
    
    distribusi = df.groupby('Kategori')['%Vol'].sum().reset_index()
    
    # Tentukan kualitas nanomaterial berdasarkan PDI rata-rata
    if pdi_rata_rata < 0.1:
        kualitas = "Sangat Baik (Monodispers)"
    elif pdi_rata_rata < 0.2:
        kualitas = "Baik"
    elif pdi_rata_rata < 0.3:
        kualitas = "Cukup"
    else:
        kualitas = "Buruk (Polydispers)"
    
    # Hasil ringkasan
    hasil = {
        'Diameter Rata-rata (nm)': round(diameter_rata_rata, 2),
        'PDI Rata-rata': round(pdi_rata_rata, 3),
        'Kualitas Nanomaterial': kualitas,
        'Distribusi Ukuran': distribusi,
        'Data Lengkap': df
    }
    
    return hasil

# Fungsi untuk membuat grafik distribusi
def buat_grafik_distribusi(df):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
    
    # Grafik 1: Scatter plot diameter vs %vol
    ax1.scatter(df['Diameter (nm)'], df['%Vol'], alpha=0.6, color='blue')
    ax1.set_xlabel('Diameter (nm)')
    ax1.set_ylabel('% Volume')
    ax1.set_title('Distribusi Ukuran Partikel')
    ax1.grid(True, alpha=0.3)
    
    # Grafik 2: Bar chart distribusi kategori
    kategori_counts = df['Kategori'].value_counts().sort_index()
    ax2.bar(kategori_counts.index.astype(str), kategori_counts.values, color='green', alpha=0.7)
    ax2.set_xlabel('Kategori Ukuran')
    ax2.set_ylabel('Jumlah Partikel')
    ax2.set_title('Distribusi Kategori Ukuran')
    ax2.tick_params(axis='x', rotation=45)
    ax2.grid(True, alpha=0.3)
    
    plt.tight_layout()
    return fig

# Fungsi untuk membuat file Word dari catatan
def buat_docx(judul, isi_catatan, data_psa=None):
    doc = Document()
    
    # Judul
    doc.add_heading(judul, 0)
    doc.add_paragraph(f"Dibuat pada: {datetime.now().strftime('%d %B %Y, %H:%M:%S')}")
    doc.add_paragraph(" ")
    
    # Isi catatan
    doc.add_heading('Catatan Praktikum', level=1)
    doc.add_paragraph(isi_catatan)
    
    # Jika ada data PSA
    if data_psa:
        doc.add_heading('Data PSA', level=1)
        doc.add_paragraph(f"Diameter Rata-rata: {data_psa['Diameter Rata-rata (nm)']} nm")
        doc.add_paragraph(f"PDI Rata-rata: {data_psa['PDI Rata-rata']}")
        doc.add_paragraph(f"Kualitas: {data_psa['Kualitas Nanomaterial']}")
    
    doc.add_heading('Metadata', level=1)
    doc.add_paragraph(f"Aplikasi: NaNote v1.0")
    doc.add_paragraph(f"Generator: Streamlit Web App")
    
    return doc

# Fungsi untuk membuat file PDF dari hasil PSA
def buat_pdf(hasil_psa, grafik_path):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []
    
    # Judul
    title = Paragraph("Laporan Hasil PSA Nanomaterial", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1, 12))
    
    # Tanggal
    date_str = Paragraph(f"Tanggal: {datetime.now().strftime('%d %B %Y')}", styles['Normal'])
    elements.append(date_str)
    elements.append(Spacer(1, 24))
    
    # Hasil perhitungan
    elements.append(Paragraph("Hasil Analisis PSA:", styles['Heading2']))
    elements.append(Spacer(1, 12))
    
    data_hasil = [
        ["Parameter", "Nilai"],
        ["Diameter Rata-rata", f"{hasil_psa['Diameter Rata-rata (nm)']} nm"],
        ["PDI Rata-rata", f"{hasil_psa['PDI Rata-rata']}"],
        ["Kualitas Nanomaterial", hasil_psa['Kualitas Nanomaterial']]
    ]
    
    table = Table(data_hasil, colWidths=[200, 200])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table)
    elements.append(Spacer(1, 24))
    
    # Grafik
    if os.path.exists(grafik_path):
        elements.append(Paragraph("Grafik Distribusi Ukuran:", styles['Heading2']))
        elements.append(Spacer(1, 12))
        img = Image(grafik_path, width=400, height=200)
        elements.append(img)
        elements.append(Spacer(1, 24))
    
    # Data lengkap
    elements.append(Paragraph("Data Lengkap Pengukuran:", styles['Heading2']))
    elements.append(Spacer(1, 12))
    
    # Buat tabel data
    df = hasil_psa['Data Lengkap']
    data_table = [df.columns.tolist()] + df.values.tolist()
    
    pdf_table = Table(data_table, colWidths=[50, 80, 80, 100, 100, 80])
    pdf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    elements.append(pdf_table)
    
    # Build PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer

# Sidebar
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/1995/1995463.png", width=100)
    st.title("NaNote")
    st.markdown("**Aplikasi Catatan Praktikum & Kalkulator PSA**")
    
    st.markdown("---")
    st.markdown("### Menu")
    menu = st.radio("Pilih Modul:", ["🏠 Beranda", "📝 Catatan Praktikum", "🧮 Kalkulator PSA", "📊 Hasil & Ekspor"])
    
    st.markdown("---")
    st.markdown("### Panduan Singkat")
    with st.expander("Cara menggunakan"):
        st.markdown("""
        1. **Catatan Praktikum**: Buat catatan praktikum dan simpan sebagai .docx
        2. **Kalkulator PSA**: Input data PDI, %vol, dan diameter
        3. **Hasil & Ekspor**: Lihat hasil dan ekspor sebagai PDF
        """)
    
    st.markdown("---")
    st.markdown("### Informasi")
    st.markdown(f"""
    **Versi:** 1.0  
    **Terakhir diperbarui:** {datetime.now().strftime('%d %B %Y')}  
    **Untuk:** Praktikum Nanomaterial
    """)

# Konten utama berdasarkan menu
if menu == "🏠 Beranda":
    st.markdown('<h1 class="main-header">🔬 NaNote - Catatan Praktikum & Kalkulator PSA</h1>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class="section-box">
            <h3>📝 Catatan Praktikum</h3>
            <p>Buat dan simpan catatan praktikum Anda dalam format Microsoft Word (.docx).</p>
            <p>Fitur:</p>
            <ul>
                <li>Editor teks lengkap</li>
                <li>Simpan sebagai .docx</li>
                <li>Template otomatis</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="section-box">
            <h3>🧮 Kalkulator PSA</h3>
            <p>Hitung hasil Particle Size Analysis dari data PDI, %vol, dan diameter.</p>
            <p>Fitur:</p>
            <ul>
                <li>Input data multiple</li>
                <li>Perhitungan otomatis</li>
                <li>Analisis kualitas nanomaterial</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="section-box">
            <h3>📊 Hasil & Ekspor</h3>
            <p>Lihat hasil analisis dan ekspor dalam format PDF.</p>
            <p>Fitur:</p>
            <ul>
                <li>Visualisasi grafik</li>
                <li>Ekspor PDF profesional</li>
                <li>Tabel data lengkap</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### 🎯 Tujuan Aplikasi")
    st.markdown("""
    NaNote dirancang untuk membantu praktikan dalam:
    1. **Mencatat** hasil praktikum nanomaterial secara digital
    2. **Menganalisis** data PSA (Particle Size Analysis)
    3. **Mengekspor** hasil dalam format standar (Word dan PDF)
    4. **Menyimpan** data secara terstruktur untuk referensi masa depan
    """)
    
    st.markdown("### ⚙️ Cara Memulai")
    st.markdown("""
    1. Pilih **"Catatan Praktikum"** dari menu untuk mulai mencatat
    2. Gunakan **"Kalkulator PSA"** untuk menganalisis data nanomaterial
    3. Ekspor hasil Anda melalui **"Hasil & Ekspor"**
    """)

elif menu == "📝 Catatan Praktikum":
    st.markdown('<h1 class="main-header">📝 Catatan Praktikum</h1>', unsafe_allow_html=True)
    
    with st.form("catatan_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            judul = st.text_input("Judul Catatan:", placeholder="Contoh: Praktikum Sintesis Nanopartikel SiO2")
            nama_praktikan = st.text_input("Nama Praktikan:", placeholder="Nama lengkap Anda")
            mata_praktikum = st.text_input("Mata Praktikum:", placeholder="Contoh: Nanomaterial dan Aplikasinya")
        
        with col2:
            tanggal = st.date_input("Tanggal Praktikum:", datetime.now())
            kelompok = st.text_input("Kelompok:", placeholder="Contoh: Kelompok 5")
            asisten = st.text_input("Asisten Lab:", placeholder="Nama asisten laboratorium")
        
        st.markdown("### Isi Catatan")
        isi_catatan = st.text_area(
            "Detail Praktikum:",
            height=300,
            placeholder="Tuliskan detail praktikum Anda di sini:\n\n1. Tujuan\n2. Alat dan Bahan\n3. Prosedur\n4. Hasil Pengamatan\n5. Analisis Data\n6. Kesimpulan"
        )
        
        # Template catatan
        with st.expander("Gunakan Template"):
            template_pilihan = st.selectbox(
                "Pilih Template:",
                ["Pilih template...", "Sintesis Nanopartikel", "Karakterisasi PSA", "Analisis XRD", "Template Kosong"]
            )
            
            if template_pilihan == "Sintesis Nanopartikel":
                isi_catatan = """1. TUJUAN
Mensintesis nanopartikel [jenis] dengan metode [metode] dan menganalisis karakteristiknya.

2. ALAT DAN BAHAN
- [Daftar alat]
- [Daftar bahan]

3. PROSEDUR KERJA
a. Persiapan Larutan
b. Sintesis Nanopartikel
c. Pencucian dan Pengeringan
d. Karakterisasi

4. HASIL PENGAMATAN
- Warna produk: 
- Bentuk fisik:
- Hasil karakterisasi:

5. ANALISIS DATA
[Analisis data yang diperoleh]

6. KESIMPULAN
[Kesimpulan praktikum]"""
        
        submit_catatan = st.form_submit_button("💾 Simpan Catatan", use_container_width=True)
    
    if submit_catatan and judul and isi_catatan:
        # Buat dokumen Word
        doc = buat_docx(judul, isi_catatan)
        
        # Simpan ke buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # Tampilkan preview
        with st.expander("Preview Catatan", expanded=True):
            st.markdown(f"**Judul:** {judul}")
            st.markdown(f"**Praktikan:** {nama_praktikan} | **Tanggal:** {tanggal.strftime('%d %B %Y')}")
            st.markdown("---")
            st.markdown(isi_catatan)
        
        # Tombol download
        st.download_button(
            label="⬇️ Download sebagai Word (.docx)",
            data=buffer,
            file_name=f"Catatan_{judul.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        st.success("✅ Catatan berhasil dibuat! Silakan download file Word di atas.")
        
        # Simpan ke session state untuk digunakan di modul lain
        st.session_state['catatan'] = {
            'judul': judul,
            'isi': isi_catatan,
            'praktikan': nama_praktikan,
            'tanggal': tanggal
        }

elif menu == "🧮 Kalkulator PSA":
    st.markdown('<h1 class="main-header">🧮 Kalkulator Particle Size Analysis</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="section-box">
    <h4>📋 Panduan Input Data</h4>
    <p>Masukkan data PSA Anda dalam tabel di bawah. Untuk setiap sampel, isikan:</p>
    <ul>
        <li><strong>PDI</strong>: Polydispersity Index (0.0 - 1.0)</li>
        <li><strong>%Vol</strong>: Persentase volume (%)</li>
        <li><strong>Diameter</strong>: Diameter partikel (nm)</li>
    </ul>
    <p><em>Catatan: PDI < 0.1 menunjukkan distribusi monodispers yang sangat baik</em></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Input data dengan tabel dinamis
    st.markdown("### 📊 Input Data PSA")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        jumlah_data = st.number_input("Jumlah Sampel:", min_value=1, max_value=20, value=5, step=1)
    
    with col2:
        if st.button("🔄 Reset Tabel", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    
    # Inisialisasi data di session state
    if 'data_psa' not in st.session_state:
        st.session_state['data_psa'] = pd.DataFrame({
            'No': range(1, jumlah_data + 1),
            'PDI': [0.15] * jumlah_data,
            '%Vol': [20.0] * jumlah_data,
            'Diameter (nm)': [50.0] * jumlah_data
        })
    
    # Buat editor tabel
    edited_df = st.data_editor(
        st.session_state['data_psa'],
        column_config={
            "No": st.column_config.NumberColumn(
                "No",
                disabled=True
            ),
            "PDI": st.column_config.NumberColumn(
                "PDI",
                min_value=0.0,
                max_value=1.0,
                step=0.01,
                format="%.3f"
            ),
            "%Vol": st.column_config.NumberColumn(
                "%Vol",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                format="%.1f"
            ),
            "Diameter (nm)": st.column_config.NumberColumn(
                "Diameter (nm)",
                min_value=0.1,
                step=0.1,
                format="%.1f"
            )
        },
        hide_index=True,
        num_rows="dynamic"
    )
    
    # Tombol hitung
    if st.button("🚀 Hitung PSA", type="primary", use_container_width=True):
        with st.spinner("Menghitung hasil PSA..."):
            # Ambil data dari tabel
            pdi_values = edited_df['PDI'].tolist()
            vol_values = edited_df['%Vol'].tolist()
            diameter_values = edited_df['Diameter (nm)'].tolist()
            
            # Hitung PSA
            hasil = hitung_psa(pdi_values, vol_values, diameter_values)
            
            if hasil:
                # Simpan hasil ke session state
                st.session_state['hasil_psa'] = hasil
                st.session_state['data_input'] = edited_df
                
                # Tampilkan hasil
                st.markdown('<div class="result-box">', unsafe_allow_html=True)
                st.markdown("### 📈 Hasil Analisis PSA")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric(
                        "Diameter Rata-rata",
                        f"{hasil['Diameter Rata-rata (nm)']} nm",
                        delta=None
                    )
                
                with col2:
                    pdi_value = hasil['PDI Rata-rata']
                    if pdi_value < 0.1:
                        delta_color = "normal"
                        arrow = "↓"
                    elif pdi_value < 0.2:
                        delta_color = "normal"
                        arrow = "↘"
                    else:
                        delta_color = "inverse"
                        arrow = "↑"
                    
                    st.metric(
                        "PDI Rata-rata",
                        f"{pdi_value:.3f}",
                        delta=arrow,
                        delta_color=delta_color
                    )
                
                with col3:
                    kualitas = hasil['Kualitas Nanomaterial']
                    if "Sangat Baik" in kualitas:
                        icon = "✅"
                    elif "Baik" in kualitas:
                        icon = "👍"
                    elif "Cukup" in kualitas:
                        icon = "⚠️"
                    else:
                        icon = "❌"
                    
                    st.metric(
                        "Kualitas",
                        f"{icon} {kualitas}",
                        delta=None
                    )
                
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Tampilkan grafik
                st.markdown("### 📊 Visualisasi Data")
                fig = buat_grafik_distribusi(hasil['Data Lengkap'])
                st.pyplot(fig)
                
                # Simpan grafik ke file sementara
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmpfile:
                    fig.savefig(tmpfile.name, dpi=300, bbox_inches='tight')
                    st.session_state['grafik_path'] = tmpfile.name
                
                # Tampilkan distribusi ukuran
                st.markdown("### 📋 Distribusi Ukuran Partikel")
                st.dataframe(hasil['Distribusi Ukuran'], use_container_width=True)
                
                st.success("✅ Perhitungan PSA selesai! Silakan buka menu 'Hasil & Ekspor' untuk menyimpan hasil.")

elif menu == "📊 Hasil & Ekspor":
    st.markdown('<h1 class="main-header">📊 Hasil & Ekspor Data</h1>', unsafe_allow_html=True)
    
    if 'hasil_psa' not in st.session_state:
        st.warning("⚠️ Belum ada data PSA. Silakan gunakan Kalkulator PSA terlebih dahulu.")
        st.info("Pergi ke menu **🧮 Kalkulator PSA** untuk menghitung data PSA Anda.")
    else:
        hasil = st.session_state['hasil_psa']
        
        # Tampilkan ringkasan
        st.markdown('<div class="result-box">', unsafe_allow_html=True)
        st.markdown("### 📋 Ringkasan Hasil PSA")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"""
            **Diameter Rata-rata:** {hasil['Diameter Rata-rata (nm)']} nm  
            **PDI Rata-rata:** {hasil['PDI Rata-rata']}  
            **Kualitas:** {hasil['Kualitas Nanomaterial']}
            """)
        
        with col2:
            st.markdown(f"""
            **Jumlah Sampel:** {len(hasil['Data Lengkap'])}  
            **Total %Vol:** {hasil['Data Lengkap']['%Vol'].sum():.1f}%  
            **Rentang Diameter:** {hasil['Data Lengkap']['Diameter (nm)'].min():.1f} - {hasil['Data Lengkap']['Diameter (nm)'].max():.1f} nm
            """)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Tab untuk data detail
        tab1, tab2, tab3 = st.tabs(["📄 Data Lengkap", "📈 Grafik", "🔄 Ekspor Data"])
        
        with tab1:
            st.markdown("### 📊 Data Lengkap Pengukuran")
            st.dataframe(hasil['Data Lengkap'], use_container_width=True)
            
            # Statistik deskriptif
            st.markdown("### 📊 Statistik Deskriptif")
            desc_stats = hasil['Data Lengkap'][['PDI', '%Vol', 'Diameter (nm)']].describe()
            st.dataframe(desc_stats, use_container_width=True)
        
        with tab2:
            st.markdown("### 📈 Visualisasi Distribusi")
            if 'grafik_path' in st.session_state:
                fig = buat_grafik_distribusi(hasil['Data Lengkap'])
                st.pyplot(fig)
            else:
                st.info("Menampilkan grafik...")
                fig = buat_grafik_distribusi(hasil['Data Lengkap'])
                st.pyplot(fig)
        
        with tab3:
            st.markdown("### 💾 Ekspor Hasil Analisis")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 📄 Ekspor sebagai PDF")
                st.markdown("Hasil PSA dalam format PDF yang rapi, termasuk:")
                st.markdown("""
                - Ringkasan hasil
                - Grafik distribusi
                - Tabel data lengkap
                - Metadata analisis
                """)
                
                if st.button("🖨️ Generate PDF Report", use_container_width=True):
                    with st.spinner("Membuat PDF..."):
                        if 'grafik_path' in st.session_state:
                            pdf_buffer = buat_pdf(hasil, st.session_state['grafik_path'])
                            
                            st.download_button(
                                label="⬇️ Download PDF Report",
                                data=pdf_buffer,
                                file_name=f"PSA_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                            
                            # Preview PDF
                            base64_pdf = base64.b64encode(pdf_buffer.read()).decode('utf-8')
                            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>'
                            st.markdown(pdf_display, unsafe_allow_html=True)
            
            with col2:
                st.markdown("#### 📊 Ekspor Data Mentah")
                st.markdown("Ekspor data dalam format yang bisa dibuka di Excel atau software analisis lainnya.")
                
                # Ekspor sebagai CSV
                csv = hasil['Data Lengkap'].to_csv(index=False)
                st.download_button(
                    label="⬇️ Download sebagai CSV",
                    data=csv,
                    file_name=f"PSA_Data_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                
                # Ekspor sebagai Excel
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    hasil['Data Lengkap'].to_excel(writer, sheet_name='Data PSA', index=False)
                    hasil['Distribusi Ukuran'].to_excel(writer, sheet_name='Distribusi', index=False)
                    
                    # Buat sheet ringkasan
                    ringkasan_data = {
                        'Parameter': ['Diameter Rata-rata (nm)', 'PDI Rata-rata', 'Kualitas Nanomaterial', 
                                     'Jumlah Sampel', 'Total %Vol', 'Tanggal Analisis'],
                        'Nilai': [hasil['Diameter Rata-rata (nm)'], hasil['PDI Rata-rata'], 
                                 hasil['Kualitas Nanomaterial'], len(hasil['Data Lengkap']),
                                 hasil['Data Lengkap']['%Vol'].sum(), datetime.now().strftime('%d/%m/%Y %H:%M')]
                    }
                    pd.DataFrame(ringkasan_data).to_excel(writer, sheet_name='Ringkasan', index=False)
                
                excel_buffer.seek(0)
                
                st.download_button(
                    label="⬇️ Download sebagai Excel",
                    data=excel_buffer,
                    file_name=f"PSA_Full_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        # Informasi tambahan
        st.markdown("---")
        st.markdown("### 📝 Catatan Interpretasi")
        with st.expander("Penjelasan Parameter PSA"):
            st.markdown("""
            **1. Diameter Rata-rata:**
            - Ukuran rata-rata partikel dalam nanometer (nm)
            - Menentukan sifat fisika-kimia nanomaterial
            
            **2. PDI (Polydispersity Index):**
            - **< 0.1**: Sangat baik, distribusi monodispers
            - **0.1 - 0.2**: Baik, distribusi seragam
            - **0.2 - 0.3**: Cukup, agak polydispers
            - **> 0.3**: Buruk, sangat polydispers
            
            **3. %Vol:**
            - Persentase volume setiap populasi partikel
            - Total harus mendekati 100%
            
            **4. Kualitas Nanomaterial:**
            - Berdasarkan PDI dan keseragaman distribusi
            - Semakin monodispers, semakin baik untuk aplikasi
            """)

# Footer
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🔬 <strong>NaNote v1.0</strong> - Aplikasi Catatan Praktikum & Kalkulator PSA</p>
        <p>Dikembangkan untuk membantu praktikan nanomaterial | © 2026</p>
    </div>
    """,
    unsafe_allow_html=True
)
