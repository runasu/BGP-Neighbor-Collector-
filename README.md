# 📘 BGP Neighbor Collector & Excel Parser (SecureCRT + Python)

Automasi pengambilan data **BGP neighbors** dari perangkat Junos via SecureCRT, kemudian memproses hasilnya menjadi file **Excel yang informatif dan siap presentasi**.  
Proyek ini menggabungkan **VBScript (SecureCRT)** sebagai collector dan **Python** sebagai parser & Excel report generator.

> ⚠️ *Catatan:* Semua kredensial, hostname, dan informasi sensitif **tidak disertakan** dan harus dikonfigurasi manual di lingkungan masing‑masing.

***

## 🚀 Fitur Utama

### 🔐 SecureCRT Script (Collector)

*   Login otomatis via SSH ke daftar perangkat.
*   Mengeksekusi perintah:
    *   `set cli screen-length 0`
    *   `set cli timestamp`
    *   `show bgp neighbor | no-more`
*   Output setiap perangkat disimpan sebagai:
    *   **Log gabungan** (central log)
    *   **Log per-host** (per device)
*   Struktur log dioptimalkan untuk parser Python (menggunakan marker `>>>` dan header `HOST:`).

***

### 📊 Python Parser (Excel Generator)

Parser menggunakan log gabungan/per-host untuk membangun laporan Excel dengan tampilan profesional:

#### ✔️ Fitur parsing:

*   Deteksi fleksibel baris `HOST:` dengan/ tanpa IP dan timestamp.
*   Robust terhadap variasi output Junos.
*   Fallback otomatis:
    *   Jika tidak ditemukan header host, parser tetap mengekstrak data dari seluruh log.
*   Deteksi & ekstraksi otomatis:
    *   Peer IP
    *   Description
    *   IPv4/IPv6 Received & Advertised
    *   Total Received/Advertised
    *   State

#### ✔️ Fitur tampilan Excel (menggunakan `xlsxwriter`):

*   **Header berwarna** (warna corporate gaya Microsoft).
*   **Freeze panes** (header + kolom pertama).
*   **Autofit kolom otomatis** berdasarkan lebar konten.
*   **Format angka** `#,##0`.
*   **Conditional Formatting:**
    *   State: Established → Hijau, Idle/Active/Connect → Merah.
    *   Kolom Total: Data bars.
    *   Kolom Delta (Total Received − Total Advertised): 3‑color scale + icon set (panah).
*   **Table Style** dengan banded rows.
*   **Sheet Summary**:
    *   Rekap agregat per perangkat.
    *   Baris Grand Total otomatis.
*   **Sheet Dashboard** (Grafik otomatis):
    *   Top‑10 Total Received per perangkat.
    *   Donut chart distribusi BGP State.
    *   Stacked chart Received vs Advertised (Top‑10).

***

## 📁 Struktur File

    /BGP-Collector/
    │
    ├── securecrt/
    │   └── bgp_collector.vbs        # Script SecureCRT (collector)
    │
    ├── parser/
    │   ├── parse_bgp_to_excel_v2_8.py
    │   └── requirements.txt         # pandas, xlsxwriter, openpyxl
    │
    └── output/
        ├── bgp_peers_all.txt        # Log gabungan
        ├── logs/YYYYMMDD/           # Log per-host
        └── BGP_Report.xlsx          # Hasil Excel

***

## 🛠️ Cara Penggunaan

### 1️⃣ Jalankan Script SecureCRT

*   Edit file `bgp_collector.vbs` dan sesuaikan:
    *   Folder target
    *   Lokasi parser
    *   Daftar perangkat
*   Jalankan dari menu **Script → Run** di SecureCRT.
*   Script akan membuat log gabungan + folder log per-perangkat otomatis.

***

### 2️⃣ Jalankan Parser Python

Pastikan dependensi:

```bash
pip install pandas xlsxwriter openpyxl
```

Jalankan parser:

```bash
python parse_bgp_to_excel_v2_8.py \
  --log "bgp_peers_all.txt" \
  --out "BGP_Report.xlsx" \
  --mode per-host \
  --engine xlsxwriter
```

***

## 🧰 Requirement

*   **Python 3.8+**
*   Library:
    *   `pandas`
    *   `xlsxwriter`
    *   `openpyxl`
*   **SecureCRT** (untuk eksekusi script VBS)

***

## 🔒 Keamanan

Proyek ini **tidak** menyimpan:

*   username/password
*   hostname
*   IP address real

Semua kredensial dikeluarkan dari repo dan harus diisi oleh pengguna sesuai lingkungan masing‑masing.

***

