# 📰 Text Preprocessing — Kata Kunci & Hasil Preprocessing

Script Python untuk melakukan **preprocessing teks berita** dari file `input.txt`.  
Menghasilkan file Excel (`output_preprocessing.xlsx`) berisi **kata kunci** dan **tabel hasil preprocessing** lengkap dengan keterangan stopword dan stemming.

---

## 📁 Struktur Folder

```
sorting/
├── input.txt                  ← File teks input (isi artikel berita)
├── stopword.txt               ← Daftar stopword Bahasa Indonesia
├── preprocess.py              ← Script utama preprocessing
├── requirements.txt           ← Daftar dependencies
├── outputs/
│   └── output_preprocessing.xlsx  ← Hasil output Excel
└── .venv/                     ← Virtual environment (dibuat manual)
```

---

## ⚙️ Instalasi

### 1. Pastikan Python sudah terinstall
```bash
python --version
# Minimal Python 3.8+
```

### 2. Buat Virtual Environment
```bash
python -m venv .venv
```

### 3. Aktifkan Virtual Environment

**Windows (PowerShell):**
```powershell
.venv\Scripts\Activate.ps1
```

**Windows (Command Prompt):**
```cmd
.venv\Scripts\activate.bat
```

**Linux / macOS:**
```bash
source .venv/bin/activate
```

> Tanda `(.venv)` akan muncul di depan terminal jika berhasil aktif.

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

Dependencies yang diinstall:
| Package | Versi | Kegunaan |
|---------|-------|----------|
| `pandas` | 3.0.1 | Manipulasi data & Excel |
| `openpyxl` | 3.1.5 | Tulis file `.xlsx` dengan styling |
| `Sastrawi` | 1.0.1 | Stemming Bahasa Indonesia |

---

## 📝 Format Input (`input.txt`)

```
Baris 1  → Judul berita
Baris 2  → Nama penulis / sumber (misal: Nama - detikNews)
Baris 3+ → Isi paragraf artikel
```

Contoh:
```
Kapal Feri Filipina Tenggelam, 29 Orang Hilang
Ikhwanul Khabibi - detikNews

Jakarta - Sebuah kapal feri di Filipina yang membawa 84 orang...
```

---

## ▶️ Menjalankan Script

Pastikan virtual environment **sudah aktif**, lalu jalankan:

```bash
python preprocess.py
```

Output akan tersimpan di:
```
outputs/output_preprocessing.xlsx
```

---

## 📊 Format Output Excel

File Excel berisi satu sheet dengan susunan:

1. **Header** — Judul & sumber berita
2. **Isi Artikel** — Teks lengkap
3. **KATA KUNCI** — Daftar kata yang paling sering muncul (non-stopword)
4. **HASIL PREPROCESSING** — Tabel dengan kolom:

| Kolom | Keterangan |
|-------|-----------|
| No. | Posisi kemunculan pertama kata di teks |
| Kata | Kata asli (setelah case folding) |
| Frekuensi | Berapa kali kata muncul di artikel |
| Keterangan | `Stopword` / `Stemming (asli→hasil)` / *(kosong jika kata biasa)* |

---

## 🔄 Menonaktifkan Virtual Environment

Setelah selesai:
```bash
deactivate
```
