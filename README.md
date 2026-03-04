# Text Preprocessing — Kata Kunci & Hasil Preprocessing

Script Python untuk melakukan **preprocessing teks berita** dari file `input.txt`.  
Menghasilkan file Excel (`output_preprocessing.xlsx`) berisi **kata kunci** dan **tabel hasil preprocessing** lengkap dengan keterangan stopword dan stemming.

---

## Struktur Folder

```
sorting/
├── input.txt                  <- File teks input (isi artikel berita)
├── stopword.txt               <- Daftar stopword Bahasa Indonesia
├── preprocess.py              <- Script utama preprocessing
├── tf-idf.py                  <- Script perhitungan TF-IDF & Cosine Similarity
├── input_tf-idf.txt           <- Input query dan dokumen untuk TF-IDF
├── requirements.txt           <- Daftar dependencies
├── outputs/
│   ├── output_preprocessing.xlsx  <- Hasil output preprocessing
│   └── output_tf-idf.xlsx         <- Hasil output TF-IDF
└── .venv/                     <- Virtual environment (dibuat manual)
```

---

## Instalasi

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

## Format Input

### `input.txt` (untuk preprocess.py)
```
Baris 1  -> Judul berita
Baris 2  -> Nama penulis / sumber (misal: Nama - detikNews)
Baris 3+ -> Isi paragraf artikel
```

### `input_tf-idf.txt` (untuk tf-idf.py)
```
Q  = teks query
D1 = teks dokumen 1
D2 = teks dokumen 2
...
```

---

## Menjalankan Script

Pastikan virtual environment **sudah aktif**, lalu jalankan:

```bash
# Preprocessing teks berita
python preprocess.py

# TF-IDF dan Cosine Similarity
python tf-idf.py
```

Output akan tersimpan di folder `outputs/`.

---

## Format Output Excel

### output_preprocessing.xlsx
File Excel berisi satu sheet dengan susunan:

1. **Header** — Judul & sumber berita
2. **Isi Artikel** — Teks lengkap
3. **KATA KUNCI** — Daftar kata yang paling sering muncul (non-stopword)
4. **HASIL PREPROCESSING** — Tabel dengan kolom:

| Kolom | Keterangan |
|-------|-----------|
| No. | Posisi kemunculan pertama kata di teks |
| Kata | Kata dasar (setelah stemming) |
| Frekuensi | Berapa kali kata muncul di artikel |
| Keterangan | `Stopword` / `Stemming (asli->hasil)` / *(kosong jika kata biasa)* |

### output_tf-idf.xlsx
File Excel berisi tiga sheet:

| Sheet | Isi |
|-------|-----|
| TF-IDF | Tabel TF (normalisasi), IDF, dan Wdt per term |
| Dokumen | Info dokumen asli beserta term yang mewakili |
| Cosine Similarity | Perhitungan WD, Panjang Vektor, dan ranking dokumen |

---

## Menonaktifkan Virtual Environment

Setelah selesai:
```bash
deactivate
```
