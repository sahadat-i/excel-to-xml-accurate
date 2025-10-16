# Project Excel to Accurate XML Converter
A Streamlit-based tool to preview, validate, and convert Excel data into XML format—designed for accounting workflows and Accurate 5 integration.

🚀 Fitur Utama
- ✅ Preview Data dengan Pagination
Tampilkan data Excel secara interaktif, lengkap dengan kolom nomor urut dan navigasi halaman.
- 🔍 Validasi Kolom AKUN BANK
Deteksi baris yang belum memilih akun bank ("" atau "PILIH AKUN BANK") sebelum proses XML.
- 🧾 Generate XML Otomatis
Konversi data valid ke format XML sesuai struktur Accurate, dengan fallback nilai default untuk kolom seperti RATE.
- 📤 Download XML (opsional, bisa ditambahkan)
Simpan hasil XML ke file untuk di-upload ke Accurate atau sistem lain.

📁 Struktur File
├── app.py               # Main Streamlit app
├── requirements.txt     # Daftar dependensi
└── README.md            # Dokumentasi



📄 Format Excel yang Didukung
- Sheet: sistem hanya mendeteksi Sheet1
- Header dimulai dari baris ke-2 (baris pertama berisi TRANSACTIONID di sel B1)
- Kolom wajib:
- ID
- NO INVOICE
- AKUN BANK
- NO AKUN
- TOTAL BAYAR / TOTAL TERIMA

⚙️ Cara Menjalankan
- Install dependensi
pip install -r requirements.txt
- Jalankan aplikasi
streamlit run app.py
- Upload file Excel (.xlsx)
Ikuti instruksi di UI untuk preview dan validasi data.

🧠 Catatan Teknis
- Kolom "RATE" otomatis diisi 1 jika kosong
- Kolom "DESCRIPTION" dikosongkan jika NaN
- Tombol Generate XML hanya aktif jika semua baris valid
- Gunakan st.data_editor(..., hide_index=True) untuk tampilan bersih

📌 Rencana Pengembangan
- [ ] Tombol download XML
- [ ] Validasi kolom tambahan (e.g. PAYEE, NAMA BANK)
- [ ] Filter dan pencarian invoice
- [ ] Multi-sheet support

## 🚀 Preview Halaman Utama

![Preview Halaman Utama](https://github.com/user-attachments/assets/b9f50d97-d4bb-4659-aae3-825e6eca89bc)


## 🚀 Preview Halaman Jenis Transaksi Penerimaan

![Preview Halaman Penerimaan](https://github.com/user-attachments/assets/bfd7fc75-91fe-4056-865c-a2a620bae0fe)

## 🚀 Preview Halaman Jenis Transaksi Pembayaran

![Preview Halaman Pembayaran](https://github.com/user-attachments/assets/72e8c2cf-c6bf-4f52-ba47-96ede7193e3e)




