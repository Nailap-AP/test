# NaNote - Aplikasi Catatan Praktikum & Kalkulator PSA

Aplikasi web untuk mencatat hasil praktikum dan mengkalkulasi hasil PSA (Particle Size Analysis) untuk nanomaterial.

## Fitur Utama

### 📝 Catatan Praktikum
- Buat catatan praktikum dengan editor teks lengkap
- Simpan catatan dalam format Microsoft Word (.docx)
- Template otomatis untuk berbagai jenis praktikum
- Metadata lengkap (nama praktikan, tanggal, kelompok, dll)

### 🧮 Kalkulator PSA
- Input data PDI, %vol, dan diameter untuk multiple sampel
- Perhitungan otomatis diameter rata-rata berbobot
- Analisis PDI rata-rata
- Klasifikasi kualitas nanomaterial
- Visualisasi data dengan grafik distribusi

### 📊 Hasil & Ekspor
- Tampilan hasil analisis yang komprehensif
- Ekspor hasil sebagai PDF report profesional
- Ekspor data mentah dalam format CSV dan Excel
- Grafik distribusi yang informatif

## Teknologi yang Digunakan
- **Streamlit** - Framework web aplikasi
- **Python 3.9+** - Bahasa pemrograman utama
- **Pandas & NumPy** - Analisis data
- **Matplotlib** - Visualisasi grafik
- **python-docx** - Generasi file Word
- **ReportLab** - Generasi file PDF

### Jalankan Lokal
```bash
# Clone repository
git clone https://github.com/username/nanote.git
cd nanote

# Install dependencies
pip install -r requirements.txt

# Run aplikasi
streamlit run app.py
