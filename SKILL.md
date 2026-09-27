---
name: jurnal-perkuliahan-unirow
description: >
  Membuat dokumen JURNAL PERKULIAHAN (catatan materi/pokok bahasan per
  pertemuan) Universitas PGRI Ronggolawe (UNIROW) Tuban FKIP, sebagai file
  .docx berisi kop UNIROW, identitas mata kuliah, dan tabel No/Hari/
  Tanggal/Jam/Makul/Pokok Bahasan (satu baris per pertemuan, tanpa kolom
  absen/checklist). Gunakan setiap kali pengguna minta "jurnal perkuliahan",
  "jurnal mengajar", "buatkan jurnal kuliah" untuk UNIROW/Ronggolawe
  Tuban/FKIP — baik diisi manual per baris, maupun WAJIB dipakai ketika
  pengguna hanya mengunggah Kontrak Kuliah dan/atau RPS lalu minta
  dibuatkan jurnalnya — skill membaca sendiri dokumen tersebut untuk
  mengambil Prodi/Dosen/Mata kuliah dan menyusun baris pokok bahasan
  mingguan dari rencana pembelajaran, tanpa pengguna perlu mengetik ulang
  data manual.
---

# Jurnal Perkuliahan Universitas PGRI Ronggolawe (UNIROW) Tuban

## PENTING — asal-usul template ini

Skill ini **tidak** dibangun dari file `.docx` resmi kampus yang lengkap —
sumbernya hanya screenshot terpotong dari baris header sebuah tabel
(`No | Hari | Tanggal | Jam | Makul | Pokok Bahasan | ...`). Kolom lain di
luar keenam ini (misalnya kolom absen/checklist pertemuan yang sempat
terlihat di screenshot) **sengaja tidak disertakan** sesuai permintaan
pengguna — tabel final hanya: `No, Hari, Tanggal, Jam, Makul, Pokok
Bahasan`, satu baris per pertemuan.

Karena tidak ada file asli untuk kop surat, judul dokumen, identitas mata
kuliah, dan blok tanda tangan, semua bagian itu **dirancang sendiri**
mengikuti gaya kop UNIROW yang konsisten dengan skill lain (`uts-unirow`,
`uas-unirow`). Jika pengguna punya file resmi kampus untuk dokumen ini,
**minta diunggah** dan sesuaikan `scripts/generate_jurnal.js` mengikuti
tata letak aslinya — jangan asumsikan rancangan saat ini sudah final tanpa
verifikasi dari pengguna.

## Cara kerja

Dokumen ini dibuat **dari nol** dengan `docx` (npm, docx-js) melalui
`scripts/generate_jurnal.js` — bukan hasil edit dari template lama, karena
tidak ada file dasar yang bisa diedit. Skrip menghasilkan:

- Kop UNIROW (logo `assets/logo_unirow.png` + nama kampus + alamat),
  judul "JURNAL PERKULIAHAN" + tahun akademik.
- Blok identitas: Fakultas (selalu FKIP), Program Studi, Dosen Pengampu,
  Semester/Kelas.
- Tabel utama A4 portrait: `No | Hari | Tanggal | Jam | Makul | Pokok
  Bahasan` — satu baris per pertemuan (default 16 baris/satu semester).
- Blok tanda tangan penutup ("Mengetahui, Ketua Program Studi" dan
  "Dosen Pengampu").

## Langkah pembuatan

### 1. Baca dokumen yang diunggah pengguna dulu, sebelum bertanya

Jika pengguna mengunggah **Kontrak Kuliah** dan/atau **RPS**, gunakan skill
`file-reading`/`docx` (`extract-text <file>`) untuk mengambil:

- Prodi, Dosen Pengampu, Mata Kuliah — dari Kontrak Kuliah atau RPS.
- **Baris Pokok Bahasan mingguan** — dari rencana 16 minggu di RPS (materi
  tiap minggu langsung dipetakan ke `pokok_bahasan` baris yang sesuai;
  minggu UTS/UAS diisi "UTS"/"UAS" apa adanya jika RPS menandainya begitu).
- `Hari`, `Tanggal`, dan `Jam` biasanya **tidak ada** di RPS/Kontrak Kuliah (jadwal
  ditentukan terpisah oleh kampus) — kosongkan (default `""`) karena dokumen ini
  merupakan jurnal berjalan yang diisi manual oleh dosen saat perkuliahan berlangsung.
- **Aturan Pengosongan Khusus (WAJIB):**
  - **Kelas:** Bagian kelas **jangan diisi** huruf/kode kelas tertentu. Format resmi: `[Semester] / ..........`.
  - **Tanggal & Tahun:** Bagian tanggal dan tahun **jangan diisi** nilai spesifik:
    - Tahun Akademik: `20... / 20...` (dibiarkan titik-titik untuk tahunnya).
    - Tanggal Pengesahan: `Tuban, .........................` (dibiarkan titik-titik tanpa tanggal/tahun).

Sebutkan singkat field mana yang diambil otomatis dari dokumen mana.

### 2. Parameter Fleksibel

- Jika pengguna tidak memberikan data spesifik tahun/tanggal/kelas, selalu gunakan format titik-titik di atas.
- Konfirmasi apakah jurnal diisi penuh 16 baris sekaligus (jika RPS
  lengkap), atau hanya sebagian — default: isi pokok bahasan yang datanya tersedia,
  kosongkan kolom hari/tanggal/jam.

### 3. Susun data JSON

```json
{
  "prodi": "...",
  "dosen": "...",
  "semester_kelas": "5 / ..........",
  "tahun_akademik": "20... / 20...",
  "jumlah_pertemuan": 16,
  "tanggal_penutup": ".........................",
  "baris": [
    {"no": 1, "hari": "", "tanggal": "", "jam": "",
     "makul": "Reading Comprehension", "pokok_bahasan": "Kontrak kuliah & pengantar"}
  ]
}
```

Baris yang tidak dicantumkan di `baris` otomatis dibuat kosong (nomor tetap
terisi 1..jumlah_pertemuan). Lihat contoh lengkap di
`references/contoh_data.json`.

### 4. Jalankan generator

**Opsi A (Utama / Rekomendasi Antigravity — Python):**
```bash
py -3 .agents/skills/jurnal-perkuliahan-unirow/scripts/generate_jurnal.py data.json output.docx
# ATAU langsung ekstrak otomatis dari berkas RPS:
py -3 .agents/skills/jurnal-perkuliahan-unirow/scripts/generate_jurnal.py path/ke/RPS.docx output.docx --hari "Senin" --jam "08:00-09:40"
```

**Opsi B (Node.js):**
```bash
node scripts/generate_jurnal.js data.json output.docx
```

(Path logo default `assets/logo_unirow.png` relatif terhadap skill ini; override dengan opsi `--logo /path/ke/logo.png` jika perlu.)

### 5. Verifikasi visual sebelum diserahkan

```bash
python3 /mnt/skills/public/docx/scripts/office/soffice.py --headless \
  --convert-to pdf /mnt/user-data/outputs/Jurnal_<Nama_MK>.docx
pdftoppm -jpeg -r 120 /mnt/user-data/outputs/Jurnal_<Nama_MK>.pdf /tmp/preview
```

`view` setiap halaman dan periksa: kop & identitas tampil benar, semua
baris pertemuan bernomor urut, teks Pokok Bahasan tidak terpotong, dan
blok tanda tangan di akhir tidak terpotong antar halaman (jika dokumen
meluber ke halaman berikutnya, itu wajar untuk jumlah baris besar — cukup
pastikan blok tanda tangan utuh dalam satu halaman, tambahkan `PageBreak`
sebelum blok tanda tangan di skrip jika perlu).

### 6. Serahkan file

Pastikan file akhir ada di `/mnt/user-data/outputs/` dengan nama jelas,
lalu panggil `present_files`.

## Aturan Konten Penting

- **Bagian kelas pada baris Semester/Kelas jangan diisi kaku** (misalnya `/ A`), melainkan selalu dikosongkan dengan titik-titik isian bebas: `[Semester] / ..........` agar dokumen fleksibel digunakan untuk kelas manapun (A, B, Pagi, Sore, dsb.) baik saat diketik maupun ditulis manual dengan pulpen.
- Fakultas selalu FKIP (Fakultas Keguruan dan Ilmu Pendidikan).
- Jangan mengarang nama dosen, prodi, hari, atau tanggal yang tidak
  disebutkan pengguna atau tidak ada di dokumen — tanyakan atau kosongkan.
- Jumlah pertemuan default 16 (satu semester); jika pengguna sebut lain
  (mis. semester pendek), ubah `jumlah_pertemuan` di data JSON.
- **Tidak ada kolom absen/checklist pertemuan** — sesuai permintaan
  eksplisit pengguna, tabel hanya berisi enam kolom di atas. Jangan
  menambahkannya kembali kecuali diminta ulang.
- Desain tata letak dirancang rapi, proporsional, dan padat sehingga seluruh
  16 baris pertemuan beserta KOP resmi dan blok tanda tangan muat utuh
  dalam **1 lembar halaman (Single-Page Fit)**.

## Spesifikasi Halaman (referensi presisi)

| Aspek | Nilai Standar Presisi |
|---|---|
| Orientasi | A4 Portrait (11906 x 16838 DXA) |
| Margin | Top 850 DXA (0.59 in), Bottom 850 DXA (0.59 in), Left 1080 DXA (0.75 in), Right 1080 DXA (0.75 in) |
| Lebar Konten | 9746 DXA |
| Tipografi | Times New Roman di seluruh dokumen |
| KOP Surat | 3-Kolom Simetris (Logo Kiri 1200 DXA, Teks Tengah 7346 DXA, Spacer Kanan 1200 DXA) |
| Garis KOP | Garis ganda resmi (*double line border* 12 dxa / 1.5 pt) |
| Blok Identitas | Tabular laser-aligned (Label 1900 DXA, Titik Dua 200 DXA, Nilai 7646 DXA) |
| Kolom Tabel Jurnal | No (450), Hari (950), Tanggal (1250), Jam (1050), Makul (1950), Pokok Bahasan (4096 DXA) |
| Shading Header | `#E2E2E2` lembut dan kontras |
| Blok Tanda Tangan | 2-Kolom (4873 DXA tiap sisi), nama bergaris bawah tebal (Kaprodi & Dosen Pengampu) |

