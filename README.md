# GENERATE LABEL BARANG OTOMATIS (EXCEL KE WORD)

Aplikasi berbasis web lokal menggunakan **Python** dan **Streamlit** untuk melakukan otomatisasi pembuatan label barang dari file Excel (`.xlsx`) menjadi dokumen Microsoft Word (`.docx`) siap cetak, dengan tata letak dan desain yang presisi mengikuti template referensi `Label New Bold.docx`.

---

## 📌 Fitur Utama

1. **Upload File Excel Sederhana**:
   - Membaca file Excel daftar barang secara otomatis berdasarkan nama header (`No`, `List Item`, `Spec`), bukan posisi kolom.
   - Kolom `UOM` otomatis diabaikan dari dokumen Word.
   - Mengabaikan baris kosong dan baris kategori (misal kategori `PERLENGKAPAN KERJA`).
   - Menampilkan total barang dan tabel pratinjau seluruh data.

2. **Dua Metode Pemilihan Nomor Barang**:
   - **Metode A (Rentang Nomor)**: Memilih barang berdasarkan rentang nomor awal s/d nomor akhir (misal No. 101 s/d 120).
   - **Metode B (Pilih Nomor Manual)**: Memasukkan nomor-nomor barang tertentu melalui input teks dipisah koma (misal `101, 105, 110, 115, 120`) atau memilih dari dropdown multi-select.

3. **Validasi & Error Handling Lengkap**:
   - Memastikan nomor awal tidak lebih besar dari nomor akhir.
   - Memberikan notifikasi peringatan jika ada nomor yang diminta tetapi tidak ada dalam file Excel.
   - Mengabaikan baris kategori secara otomatis.
   - Pesan error ramah pengguna dalam Bahasa Indonesia.

4. **Pratinjau Data Sebelum Generate**:
   - Menampilkan ringkasan nomor yang dipilih, jumlah label, dan estimasi halaman Word.
   - Menampilkan tabel data yang akan diproses persis sesuai data asli Excel tanpa pengubahan.

5. **Format Word Sesuai Referensi (`Label New Bold.docx`)**:
   - Ukuran kertas **Letter Landscape** (11 x 8.5 inci) dengan margin presisi.
   - Tabel 10 kolom dan 3 baris label (diselingi 2 baris spacer) dengan kapasitas **30 label per halaman**.
   - Orientasi tulisan vertikal (`tbRl`), perataan tengah (*center*), jenis font **Nyala**, tebal (**Bold**).
   - Ukuran font: Nama Barang **22 pt**, Spesifikasi **20 pt**.
   - Nama barang berada di bagian atas dan spesifikasi berada tepat di bawahnya.
   - Jika spesifikasi kosong, hanya menampilkan nama barang tanpa teks tambahan.
   - **Otomatis multi-halaman**: Jika jumlah label lebih dari 30, tabel secara otomatis dilanjutkan ke halaman berikutnya secara rapi tanpa merusak tata letak.

6. **Generate & Download Langsung**:
   - Penamaan file otomatis:
     - Rentang: `Label_Barang_101-120.docx`
     - Manual: `Label_Barang_Custom.docx`
   - File hasil unduhan dapat langsung dibuka dan diedit di Microsoft Word.

---

## 📁 Struktur Project

```text
generate-label-barang/
│
├── app.py                     # Antarmuka web lokal Streamlit
├── label_generator.py         # Modul inti pemrosesan Excel, validasi & pembuatan dokumen Word
├── requirements.txt           # Daftar dependensi library Python
├── README.md                  # Dokumentasi lengkap penggunaan aplikasi
├── templates/
│   └── Label New Bold.docx    # Template referensi format dan desain label
└── sample_data/
    └── LIST BARANG.xlsx       # Contoh file Excel referensi untuk pengujian
```

---

## 🚀 Cara Instalasi dan Menjalankan Aplikasi

### 1. Prasyarat
Pastikan Anda sudah menginstal **Python 3.9** atau yang lebih baru di komputer Anda.

### 2. Instalasi Dependensi
Buka terminal / PowerShell di direktori project ini, lalu jalankan:
```bash
pip install -r requirements.txt
```

### 3. Menjalankan Aplikasi Streamlit
Jalankan perintah berikut:
```bash
streamlit run app.py
```
Aplikasi akan otomatis terbuka di browser default pada alamat:
`http://localhost:8501`

---

## 🔄 Alur Penggunaan

```text
[1. UPLOAD EXCEL] ➔ [2. PILIH NOMOR BARANG] ➔ [3. PREVIEW DATA] ➔ [4. GENERATE WORD] ➔ [5. DOWNLOAD WORD]
```

1. **Upload Excel**:
   - Unggah file `.xlsx` Anda atau klik tombol **"Gunakan Sample Data"** untuk langsung mencoba file referensi bawaan.
2. **Pilih Nomor Barang**:
   - Pilih **Metode A** untuk rentang nomor (contoh: 101 sampai 120), atau
   - Pilih **Metode B** untuk nomor manual (contoh: 101, 105, 110, 115, 120).
3. **Periksa Preview**:
   - Pastikan data barang dan spesifikasi yang muncul di tabel sudah sesuai keinginan.
4. **Tekan Tombol GENERATE WORD**:
   - Sistem akan menyusun dokumen Word sesuai format `Label New Bold.docx`.
5. **Download Dokumen**:
   - Klik tombol **DOWNLOAD WORD** untuk menyimpan file `.docx`.

---

## 🧪 Hasil Pengujian Referensi

Telah diuji dengan data referensi `LIST BARANG.xlsx` dan template `Label New Bold.docx`:
- **Pembacaan Excel**: Berhasil membaca 368 item barang valid, mengabaikan header dan baris kategori `A. PERLENGKAPAN KERJA`.
- **Rentang Nomor 101–120**: Menghasilkan tepat 20 label pada baris 1 dan baris 2 tabel, dengan font Nyala Bold, orientasi vertikal, dan posisi spesifikasi di bawah nama barang.
- **Label Lebih dari 30**: Berhasil membagi secara otomatis ke halaman kedua dengan layout tabel identik.
- **Barang Tanpa Spec**: Ditampilkan nama barang saja tanpa baris kosong berlebih atau teks buatan.
