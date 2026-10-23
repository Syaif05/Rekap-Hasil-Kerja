# 🚀 Rekap Hasil Kerja - Asisten Input Follow Up Shipment

Aplikasi web modern, ringan, dan siap pakai untuk mempermudah pelaporan hasil kerja follow up shipment toko secara otomatis, akurat, dan super cepat.

Dirancang agar **bisa langsung digunakan oleh siapa saja** di tim kerja tanpa perlu instalasi rumit.

---

## ✨ Fitur Unggulan

1. **Auto-Grouping Kendala Otomatis**:
   - Semua shipment dengan kendala yang sama langsung digabungkan dalam satu kartu batch.
   - Nomor shipment otomatis digabung menggunakan pemisah koma:
     ```text
     S-261021-AGJFYMH,S-261021-AGRGPRM,S-261021-AGJSVJH,S-261021-AGQGJNR
     ```
2. **📋 Tempel / Paste Langsung dari Spreadsheet (Google Sheets / Excel)**:
   - Tidak perlu repot unduh dan unggah file `.xlsx`.
   - Cukup salin (*Copy*) tabel langsung dari Google Sheets atau Microsoft Excel, lalu klik tombol **Tempel / Paste Spreadsheet**.
   - Otomatis mendeteksi baris, kolom, dan pemisah (*tab/comma*) secara cerdas.
3. **📁 Dukungan Upload File Excel**:
   - Bisa juga mengunggah file `.xlsx` / `.xls` kapan saja dari komputer atau HP.
4. **Urutan Salin 1-ke-6 Sesuai Form Portal**:
   1. **Nomor Shipment** (dipisahkan koma)
   2. **Jumlah Personil Toko** (angka saja)
   3. **Nama Pejabat Toko** (otomatis menggunakan Nama AC jika nama pejabat kosong)
   4. **Status Equipment Motor?** (`Bagus` / `Rusak`)
   5. **Hasil Follow Up** (Kategori grouping)
   6. **Catatan Tambahan (Kendala)**
5. **Tombol Salin Cepat 1-Klik**:
   - Salin langsung dengan sekali klik ke clipboard beserta feedback visual dan audio chime lembut.
   - Tombol **"Salin Semua Field"** untuk menyalin seluruh baris dalam format tab-separated.
6. **Rekap Semua Catatan Kendala**:
   - Di bagian bawah, seluruh catatan tambahan kendala dari semua batch otomatis dirangkum dalam satu paragraf dipisahkan koma, siap disalin sekali klik.
7. **Floating Progress Bar Real-Time**:
   - Kapsul melayang di bagian bawah dengan persentase real-time, efek animasi spring bounce, dan pendaran warna dinamis.
8. **Dukungan Mode Terang & Gelap (Dark Mode)**:
   - Nyaman digunakan saat shift siang maupun shift malam.

---

## 🌐 Cara Mengakses / Menjalankan

### Cara 1: Buka Website di Vercel (Online Gratis)
Aplikasi ini di-hosting di Vercel sehingga seluruh tim dapat langsung mengaksesnya melalui link web tanpa perlu menginstal apa pun di komputer atau HP masing-masing.

### Cara 2: Menjalankan Secara Lokal (Opsional)
```bash
git clone https://github.com/Syaif05/Rekap-Hasil-Kerja.git
cd Rekap-Hasil-Kerja
npm start
```
Buka browser dan buka `http://localhost:3000`.

---

## 🚀 Panduan Hosting ke Vercel (Gratis)

1. Buka [Vercel](https://vercel.com) dan login dengan akun GitHub Anda (`Syaif05`).
2. Klik tombol **"Add New..."** ➔ **"Project"**.
3. Pilih repositori **`Rekap-Hasil-Kerja`** dari daftar repositori GitHub Anda.
4. Biarkan pengaturan default:
   - Framework Preset: **Other**
   - Root Directory: `./`
   - Build Command: `echo 'Static build ready'`
   - Output Directory: `./`
5. Klik **"Deploy"**.
6. Dalam hitungan detik, link website publik (contoh: `https://rekap-hasil-kerja.vercel.app`) siap dibagikan ke seluruh rekan kerja Anda!

---

## 🛠️ Teknologi yang Digunakan
- **HTML5 & Vanilla CSS**: Desain modern white capsule, cepat, responsif di desktop & mobile tanpa ketergantungan library berat.
- **JavaScript (ES6+)**: Logika pemrosesan data, auto-grouping engine, parsing spreadsheet clipboard.
- **SheetJS (xlsx.full.min.js)**: Pengurai file Excel offline di sisi klien (*client-side*).

---

© 2026 Syaifulloh. Dibuat untuk efisiensi kerja tim follow up shipment.
