# Skill Jurnal Perkuliahan 

[![Antigravity Skill](https://img.shields.io/badge/Antigravity-Skill-10b981.svg)](https://github.com/mariofahmi/skilljurnalperkuliahan)
[![Template](https://img.shields.io/badge/Template-UNIROW%20FKIP-059669.svg)](https://unirow.ac.id)
[![Kurikulum](https://img.shields.io/badge/Kurikulum-OBE%202026-047857.svg)](https://unirow.ac.id)
[![Evaluasi](https://img.shields.io/badge/Rencana-16%20Pertemuan-06b6d4.svg)](https://mariofahmi.github.io/skilljurnalperkuliahan/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Official skill and automation agent in **Google Antigravity & Agentic AI** for generating, previewing, and exporting official **Jurnal Perkuliahan (16 Pertemuan)** 

🚀 **Web Simulator & Preview Portal:** [https://mariofahmi.github.io/skilljurnalperkuliahan/](https://mariofahmi.github.io/skilljurnalperkuliahan/)

---

## 🌟 Fitur Utama

1. **Format Resmi Presisi 1 Halaman A4 Sempurna:**
   - Mempertahankan integritas KOP surat resmi 3-kolom simetris berlogo UNIROW FKIP dengan garis ganda (*double divider*).
   - Tabel 6 kolom standar: `No | Hari | Tanggal | Jam | Makul | Pokok Bahasan`.
   - Blok tanda tangan simetris 2-kolom: Kiri Ketua Program Studi (**Mario Fahmi Syahrial, M.Pd.**) dan Kanan Dosen Pengampu.
2. **Kepatuhan Format Blank Resmi:**
   - **Semester/Kelas:** Bagian kelas dibiarkan kosong dengan format titik-titik `[Semester] / ..........`.
   - **Tahun Akademik:** Format kosong bertitik `20... / 20...`.
   - **Tanggal Pengesahan:** Dikosongkan `Tuban, .........................`.
   - **Jadwal Harian:** Kolom Hari, Tanggal, Jam tabel kosong secara default (`""`) untuk diisi saat perkuliahan berlangsung.
3. **Ekspor Serbaguna:**
   - Ekspor dokumen Word (`.docx`) ber-XML resmi siap cetak dan arsip akreditasi.
   - Ekspor cetak PDF A4 dengan isolasi CSS `@media print`.
   - Ekspor format Markdown (`.md`) dan salin ke clipboard.
4. **Katalog 61 Mata Kuliah Terintegrasi:**
   - Seluruh kurikulum PPKn Semester 1 sampai 8 lengkap dengan 16 rincian pokok bahasan mingguan siap pakai.

---

## 💻 Instalasi di Google Antigravity

Jalankan perintah clone berikut di terminal workspace Antigravity Anda:

```bash
git clone https://github.com/mariofahmi/skilljurnalperkuliahan.git .agents/skills/jurnal-perkuliahan-unirow
```

Atau letakkan file `jurnal-perkuliahan-unirow.skill` di workspace Anda.

---

## 🚀 Penggunaan AI Slash Command

Cukup ketik salah satu slash command berikut di chat prompt Antigravity:

```markdown
/jurnal-perkuliahan @RPS_Mata_Kuliah.docx
/jurnal @Kontrak_Kuliah.docx
/jurnal-unirow @data.json
```

AI akan secara otomatis membaca identitas mata kuliah, dosen pengampu, serta memetakan ke-16 pokok bahasan mingguan langsung ke template resmi tanpa perlu input data berulang.

---

## 🛠️ Eksekusi Terminal Python

Dokumen juga dapat digenerate secara langsung via terminal CLI:

```bash
# Sintaks dari file RPS (.docx)
py scripts/generate_jurnal.py "RPS_Mata_Kuliah.docx" "Jurnal_Perkuliahan.docx"

# Sintaks dari data JSON
py scripts/generate_jurnal.py "data.json" "Jurnal_Perkuliahan.docx"
```

---

## 📝 Format Data Input (`data.json`)

```json
{
  "prodi": "Pendidikan Pancasila dan Kewarganegaraan (PPKn)",
  "dosen": "Dr. Usep Supriatna, M.Pd.",
  "kaprodi": "Mario Fahmi Syahrial, M.Pd.",
  "mk": "Ilmu Budaya Dasar",
  "semester_kelas": "1 / ..........",
  "tahun_akademik": "20... / 20...",
  "tanggal_penutup": ".........................",
  "baris": [
    {
      "no": 1,
      "hari": "",
      "tanggal": "",
      "jam": "",
      "pokok_bahasan": "Pengantar Ilmu Budaya Dasar: Hakikat, Ruang Lingkup, dan Urgensi bagi Sarjana PPKn"
    },
    {
      "no": 8,
      "hari": "",
      "tanggal": "",
      "jam": "",
      "pokok_bahasan": "UJIAN TENGAH SEMESTER (UTS)"
    },
    {
      "no": 16,
      "hari": "",
      "tanggal": "",
      "jam": "",
      "pokok_bahasan": "UJIAN AKHIR SEMESTER (UAS)"
    }
  ]
}
```

---

## 🏛️ Hak Cipta & Pengembang

Dikembangkan oleh **Mario Fahmi** untuk Program Studi Pendidikan Pancasila dan Kewarganegaraan (PPKn) & Fakultas Keguruan dan Ilmu Pendidikan (FKIP), Universitas PGRI Ronggolawe (UNIROW) Tuban. Lisensi di bawah [MIT License](LICENSE).
