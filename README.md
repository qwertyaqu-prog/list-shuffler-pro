# List Shuffler Pro

**List Shuffler Pro** adalah aplikasi berbasis desktop (GUI) menggunakan Python dan Tkinter yang berfungsi untuk merapikan, membagi, dan mengacak baris teks (list) lalu mengekspornya langsung ke dalam format file Excel (`.xlsx`).

Aplikasi ini sangat berguna untuk membagi data nama, produk, atau data teks lainnya ke dalam beberapa kolom secara instan, baik dengan mempertahankan urutan asli maupun diacak secara random.

## ✨ Fitur Utama
* **Input Real-time Counter:** Menghitung jumlah item/baris yang dimasukkan secara langsung.
* **Dual Mode Pembagian:**
  * **Mode Jumlah Kolom:** Menentukan target berapa kolom Excel yang ingin dibuat[cite: 1].
  * **Mode Item per Kolom:** Menentukan batasan maksimal baris data per kolomnya[cite: 1].
* **Fitur Shuffle (Pengacakan):** Mengacak urutan item menggunakan algoritma pseudo-random bawaan Python sebelum diekspor[cite: 1].
* **Ekspor Instan ke Excel:** Menyimpan hasil pembagian data langsung menjadi file spreadsheet `.xlsx`[cite: 1].
* **Desain Minimalis:** Antarmuka responsif dan bersih dengan tema `clam` dari TTK Tkinter[cite: 1].

## 🚀 Cara Instalasi & Menjalankan

### Prasyarat
Pastikan Anda sudah menginstal **Python 3.x** di komputer Anda.

### 1. Clone Repositori
Buka Terminal atau Command Prompt (CMD), lalu jalankan perintah ini:
```bash
git clone https://github.com/qwertyaqu-prog/list-shuffler-pro.git

cd list-shuffler-pro
```

### 2. Instal Library yang Dibutuhkan

Aplikasi ini membutuhkan library openpyxl untuk memproses file Excel[cite: 1]. Instal melalui terminal/CMD:
```Bash
pip install openpyxl
```

### 3. Jalankan Aplikasi

Jalankan script utama aplikasi dengan perintah berikut:
```Bash
python list-shuffler-pro.py
```

🛠️ Teknologi yang Digunakan

    Python 3

    Tkinter & TTK (Antarmuka Grafis / GUI)[cite: 1]

    OpenPyXL (Manipulasi Spreadsheet Excel)[cite: 1]

### 📝 Lisensi

Project ini dilisensikan di bawah MIT License.
