# WA Automation Server

Eksekutor blast WhatsApp terjadwal — bagian dari sistem **WA Blast Automation**.  
Repo ini berjalan di background via **Windows Task Scheduler**, membaca job dari **Supabase**, membuka file Excel dari SharePoint/OneDrive, lalu mengirim screenshot range ke grup WhatsApp.

> **Repo ini adalah pasangan dari:**  
> `WA automation` (Node.js + Express + dashboard web) — keduanya harus berjalan bersama.

---

## 🧩 Gambaran Sistem

```
┌─────────────────────────────────────┐     ┌──────────────────────────────────────────┐
│   WA automation  (repo dashboard)   │     │  WA automation server  (repo ini)        │
│                                     │     │                                          │
│  ✅ Node.js + Express               │     │  ✅ Python + win32com (Excel COM)        │
│  ✅ Frontend: Vite + Tailwind CSS   │◄───►│  ✅ Eksekutor blast WA harian            │
│  ✅ Supabase Auth & Database        │     │  ✅ Windows Task Scheduler               │
│  ✅ Baileys WA session manager      │     │  ✅ Baca Excel via SharePoint/OneDrive   │
│  ✅ Form input automation           │     │  ✅ Kirim pesan ke grup WA               │
└─────────────────────────────────────┘     └──────────────────────────────────────────┘
              ↕ keduanya membaca & menulis ke
                   ┌─────────────────────┐
                   │   Supabase Database  │
                   │  (form_request,      │
                   │   wa_accounts,       │
                   │   wa_groups, dsb.)   │
                   └─────────────────────┘
```

---

## ⚙️ Cara Kerja

```
Windows Task Scheduler
│
│  Trigger: setiap jam (mis. 06:00, 07:00, 08:00, ...)
▼
blast_worker.py
│
├─ 1. Baca .env → ambil SUPABASE_URL & SUPABASE_SERVICE_ROLE_KEY
│
├─ 2. Query Supabase REST API → cari form_request yang:
│      - status = 'active'
│      - tanggal_blast mengandung tanggal hari ini
│      - jam_blast dalam window [jam, jam + 15 menit]
│      - last_run_date ≠ hari ini (belum jalan hari ini)
│
├─ 3. Untuk setiap job ditemukan:
│      - Buka file Excel dari SharePoint via COM (ms-excel: protocol)
│      - RefreshAll data external
│      - Screenshot range sel yang ditentukan di sheet_tables
│      - Kirim ke Baileys Local Server → diteruskan ke grup WA
│      - Update last_run_date = hari ini
│
└─ 4. Selesai — tunggu trigger Task Scheduler berikutnya
```

---

## 📋 Prasyarat

Sebelum setup, pastikan sudah ada:

| Kebutuhan | Keterangan |
|---|---|
| **Python >= 3.9** | [python.org](https://www.python.org/downloads/) — pastikan di-checklist "Add to PATH" saat install |
| **Microsoft Excel** | Harus terinstall di Windows (bukan Excel Online) |
| **Login Microsoft** | Akun Microsoft/Office harus sudah login di komputer (untuk buka file SharePoint) |
| **Repo `WA automation` berjalan** | Node.js server harus aktif (default port 4000 untuk Baileys) |
| **Akun Supabase** | [supabase.com](https://supabase.com/) — gratis, sudah setup dari repo dashboard |

---

## 🚀 Setup & Instalasi

### 1. Clone repo ini

```bash
git clone <URL_REPO_INI>
cd "WA automation server"
```

### 2. Install dependencies Python

> **Rekomendasi:** gunakan virtual environment agar tidak bentrok dengan instalasi Python lain.

```bash
# Buat virtual environment (opsional tapi disarankan)
python -m venv venv

# Aktifkan venv (Windows)
venv\Scripts\activate

# Install semua dependencies
pip install requests pywin32 python-dotenv playwright

# Install browser Playwright (opsional, hanya kalau dipakai)
# playwright install chromium
```

### 3. Buat file `.env`

Salin template lalu isi nilainya:

```bash
copy .env.example .env
```

Lalu buka `.env` dan isi:

```env
# URL project Supabase kamu
SUPABASE_URL=https://xxxxxxxxxxxxxxxx.supabase.co

# Service Role Key — ambil dari Supabase Dashboard
# Project Settings > API > service_role (secret)
SUPABASE_SERVICE_ROLE_KEY=eyJhbGci...

# URL Node.js server WA automation yang mengelola sesi Baileys
# Pastikan server ini sudah berjalan sebelum blast_worker.py dijalankan
BAILEYS_LOCAL_URL=http://localhost:4000
```

**Cara ambil key Supabase:**
1. Buka [Supabase Dashboard](https://supabase.com/dashboard)
2. Pilih project kamu
3. Pergi ke **Project Settings → API**
4. Salin `URL` dan `service_role` key (bukan `anon` key!)

> [!CAUTION]
> `SUPABASE_SERVICE_ROLE_KEY` memberi akses **penuh** ke database tanpa RLS. Jangan pernah share, commit ke Git, atau expose key ini.

### 4. Edit path di `run_blast.bat`

Buka file `run_blast.bat` dan sesuaikan `REPO_DIR` dengan lokasi repo ini di komputermu:

```bat
set REPO_DIR=C:\Users\fresn\OneDrive\Documents\INSP-Project\WA automation server
```

> [!NOTE]
> Kalau kamu pakai virtual environment (venv), uncomment baris aktivasi venv di dalam `run_blast.bat`:
> ```bat
> call "%REPO_DIR%\venv\Scripts\activate.bat"
> ```

### 5. Test jalankan manual

Pastikan server WA automation (Node.js, port 4000) sudah berjalan, lalu:

```bash
# Aktifkan venv dulu kalau pakai
venv\Scripts\activate

# Test jalankan worker
python blast_worker.py
```

Cek output di terminal dan di file `worker.log`. Jika ada job yang due hari ini, job akan diproses. Jika tidak ada job due, output normal: *"Gak ada job due, worker selesai."*

**Untuk testing paksa semua job jalan (abaikan jam/tanggal):**
```bash
# Tambahkan sementara di .env:
FORCE_RUN_ALL=1
```
> [!WARNING]
> Hapus atau set kembali `FORCE_RUN_ALL=0` setelah testing! Jangan biarkan aktif di production.

---

## ⏰ Setup Windows Task Scheduler

Ini adalah langkah final agar `blast_worker.py` berjalan otomatis setiap jam.

### Cara via GUI (Task Scheduler):

1. Buka **Task Scheduler** (`taskschd.msc`)
2. Klik **Create Basic Task...**
3. Isi nama, misal: `WA Blast Worker`
4. **Trigger:** pilih **Daily** → centang **Repeat task every 1 hour** selama `Indefinitely`
5. **Action:** pilih **Start a program**
   - Program/script: browse ke `run_blast.bat` (path lengkap)
   - Start in: path folder repo ini (tanpa nama file)
6. Finish

### Cara via Command Prompt (satu perintah):

Buka CMD sebagai **Administrator**, lalu jalankan:

```cmd
schtasks /create /tn "WA Blast Worker" /tr "\"C:\Users\fresn\OneDrive\Documents\INSP-Project\WA automation server\run_blast.bat\"" /sc hourly /mo 1 /f
```

Ganti path sesuai lokasi `run_blast.bat` di komputermu.

### Verifikasi task terdaftar:

```cmd
schtasks /query /tn "WA Blast Worker"
```

### Jalankan manual via Task Scheduler:

```cmd
schtasks /run /tn "WA Blast Worker"
```

---

## 🗂️ Struktur Folder

```
WA automation server/
├── blast_worker.py     # Script utama: query Supabase → eksekusi blast WA
├── run_blast.bat       # File BAT untuk Windows Task Scheduler
├── .env                # Konfigurasi lokal (tidak di-commit ke Git)
├── .env.example        # Template .env — salin & isi
├── worker.log          # Log output setiap kali worker jalan (auto-generated)
├── chrome_profile/     # Profil Chrome Playwright (auto-generated, jangan dihapus)
└── .gitignore
```

---

## 🗄️ Tabel Database Supabase yang Dibaca

| Tabel | Kolom yang dipakai | Keterangan |
|---|---|---|
| `form_request` | `id`, `nama_automation`, `link`, `jam_blast`, `tanggal_blast`, `caption`, `last_run_date`, `status` | Data job automation |
| `request_sheets` | `sheet_name` | Sheet Excel yang dipilih |
| `sheet_tables` | `cell_range` | Range sel yang di-screenshot (mis. `C1:E9`) |
| `request_groups` | — | Grup WA tujuan per automation |
| `wa_groups` | `group_jid`, `group_name`, `wa_account_id` | Identitas grup WhatsApp |


## 🔗 Repositori Terkait

| Repo | Deskripsi |
|---|---|
| [`WA automation`](https://github.com/fres12/WA-Automation-Server) | Dashboard web (Node.js + Vite), input form automation, manajemen sesi WA |
| `WA automation server` *(repo ini)* | Eksekutor blast harian, berjalan via Windows Task Scheduler |

