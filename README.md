# METT-4 KSDAE — Versi Web

Versi web dari **Management Effectiveness Tracking Tool (METT-4) KSDAE v4.4** untuk penilaian efektivitas pengelolaan kawasan konservasi di Indonesia. Aplikasi ini dibangun dari file `METT 4 _Indonesia_KSDAE_VER4.4.xlsm` dan dijalankan dari **GitHub Pages**.

Ada dua mode:

* **Mode lokal** (bawaan): data tersimpan di browser masing-masing dan di file .json. Tidak perlu server.
* **Mode online**: semua orang yang membuka link mengisi **satu penilaian yang sama** (mis. 3 kelompok sekaligus). Data tersimpan di Google Sheet milik admin lewat Google Apps Script. Pemasangan sekitar 15 menit; ikuti **[docs/PANDUAN-ONLINE.md](docs/PANDUAN-ONLINE.md)**.

> **Prinsip utama: file Excel adalah acuan.** Teks pertanyaan, pilihan jawaban, panduan, daftar dropdown, **rumus**, dan aturan skor tidak ditulis ulang. Semuanya diambil otomatis dari file Excel oleh `tools/extract.py`. Rumus Excel asli (1.907 rumus) dihitung oleh mesin rumus di browser. Hasilnya diuji sama dengan Excel (lihat [Pengujian](#pengujian)).

## Fitur

| Lembar di Excel | Di versi web |
|---|---|
| Tentang, Pengenalan METT 4 | Halaman awal & pengantar |
| Atribut Kawasan | Formulir dengan dropdown, multi-pilih, dan validasi yang sama |
| Penilaian Ancaman | 62 sub-ancaman dalam 15 kategori, termasuk panduan tingkat ancaman |
| Daftar Pertanyaan & Skor | Daftar 42 butir, skor, skor maks., status kelengkapan, kolom skor sebelumnya |
| 1. … 38. (+ 7A, 31A, 35A, 36A) | Satu halaman per butir: "tidak relevan", pilihan A–D / checkbox, bukti, aksi, sumber informasi |
| Detail nilai utama / spesies / habitat | Tabel isian dengan dropdown asli |
| Dashboard | Skor per elemen, grafik radar/batang/donat, tabel kondisi |
| Aksi meningkatkan Pengelolaan | Rencana aksi (kapan, siapa, anggaran) |
| Print BA, Print_2 | Berita Acara & Rencana Aksi siap cetak / simpan PDF |
| Makro ekspor | Unduh XLSX/CSV dengan kolom yang sama + **Rekap Multi-Kawasan** |

Fitur lain:

* **Simpan otomatis** di browser, plus **Simpan file / Buka** (.json) untuk berpindah perangkat atau mengirim ke pusat.
* **Mode online:** isian bersama banyak kelompok, sinkron otomatis per kolom, tetap jalan saat sinyal putus, dan **Riwayat Perubahan** dengan tombol *Kembalikan*.
* **Isi skor tahun lalu** dari file penilaian sebelumnya.
* Tetap **bisa dipakai tanpa internet** setelah halaman pernah dibuka sekali.
* **Privasi:** pada mode lokal tidak ada data yang dikirim ke server mana pun. Pada mode online, data hanya dikirim ke Google Sheet milik admin.

## Menjalankan di GitHub Pages

1. Buat repository baru di GitHub (mis. `mett4-ksdae-web`), **Public**.
2. Unggah seluruh isi folder ini ke repository (lewat tombol *Add file → Upload files*, atau `git push`).
3. Buka **Settings → Pages**. Pada *Build and deployment*, pilih **Source: Deploy from a branch**, **Branch: `main`**, folder **`/ (root)`**, lalu **Save**.
4. Tunggu 1–2 menit. Aplikasi tersedia di `https://<nama-akun>.github.io/mett4-ksdae-web/`.

> Repository **privat** juga bisa memakai GitHub Pages, tetapi hanya untuk paket berbayar tertentu (GitHub Pro/Team/Enterprise). Aturan paket bisa berubah, jadi periksa dokumentasi GitHub sebelum memutuskan.

Untuk mencoba di komputer sendiri:

```bash
python3 -m http.server 8000
# buka http://localhost:8000
```

(Membuka `index.html` langsung dari folder juga berjalan, kecuali fitur offline.)

## Alur pemakaian

1. **Atribut Kawasan** → isi data kawasan, terutama *Nama Kawasan*, *Nilai penting 1–5*, *Tahun* dan *Tanggal Penilaian*.
2. **Penilaian Ancaman** → klik **Isi** pada ancaman yang ada, pilih tingkat ancaman dan kerusakannya.
3. **Pertanyaan 1–38** → pilih jawaban, isi bukti & aksi. Kotak nomor di menu kiri berwarna hijau (lengkap), kuning (perlu dilengkapi), atau abu-abu (tidak relevan).
4. **Dashboard** → lihat hasil.
5. **Aksi meningkatkan Pengelolaan** → lengkapi kapan, siapa, dan anggaran.
6. **Print BA / Print_2** → *Cetak / Simpan PDF*.
7. **Kelola Data & Ekspor** → *Simpan ke file*, *Unduh XLSX*, dan *Tandai FINAL*.

Untuk rekap tingkat UPT/Direktorat: kumpulkan file `.json` dari tiap kawasan, buka **Rekap Multi-Kawasan**, lalu tambahkan semua file → *Unduh XLSX gabungan*. Hasilnya setara file `METT4_exported_data.xlsm`, ditambah sheet Ringkasan.

## Struktur repository

```
index.html              kerangka aplikasi
css/app.css             tampilan & gaya cetak
data/workbook.js        ISI FILE EXCEL (otomatis dari tools/extract.py — jangan diedit manual)
js/engine.js            mesin rumus Excel (SUM, IF, COUNTIF, CONCAT, TEXT, referensi, tabel)
js/model.js             data penilaian + logika makro VBA + validasi + ekspor
js/charts.js            grafik SVG dari definisi grafik Excel
js/xlsx.js              penulis/pembaca XLSX tanpa pustaka luar
js/sync.js              sinkronisasi mode online (per kolom, antrean offline)
js/config.js            PENGATURAN: isi SYNC_URL untuk mode online
js/app.js               halaman & interaksi
backend/Code.gs         backend Google Apps Script untuk mode online
sw.js, manifest…        mode offline (PWA)
assets/img/             logo & sampul dari file Excel
tools/extract.py        Excel (.xlsm) → data/workbook.js
tools/vba_reader.py     pembaca kode makro VBA dari .xlsm
tests/                  pengujian kesetaraan dengan Excel
docs/TEMUAN.md          kejanggalan pada file asli & keputusan penanganannya
docs/PANDUAN-ONLINE.md  cara memasang mode online (Google Sheet + Apps Script)
```

## Memperbarui bila ada versi Excel baru

Bila tim merevisi file Excel (teks pertanyaan, rumus, opsi dropdown), cukup jalankan:

```bash
pip install openpyxl
python3 tools/extract.py "METT 4 _Indonesia_KSDAE_VER4.5.xlsm" data/workbook.js
node tests/check_excel_cache.js && node tests/test_logic.js
```

Lalu commit `data/workbook.js`. Ubah juga nama cache di `sw.js` (mis. `v1.0.1`) agar pengguna offline mendapat versi baru.

Perubahan teks, opsi, dan rumus akan **otomatis** ikut. Perubahan **struktur** (menambah butir pertanyaan, memindahkan baris isian, sheet baru) mungkin memerlukan penyesuaian kecil di `js/app.js` atau `js/model.js`.

## Pengujian

| Uji | Perintah | Hasil saat rilis |
|---|---|---|
| Mesin web vs nilai yang dihitung Microsoft Excel (tersimpan di file) | `node tests/check_excel_cache.js` | 1.907 rumus, 0 beda |
| Logika makro, validasi, ekspor | `node tests/test_logic.js` | 240 lulus |
| 6 skenario isian: mesin web vs file Excel asli yang dihitung LibreOffice | `node tests/make_scenarios.js tests/out` lalu `python3 tests/compare_with_excel.py "<file .xlsm>" tests/out` | 11.442 perbandingan, 0 beda (2 perbedaan LibreOffice–Excel dijelaskan di TEMUAN.md) |
| Antarmuka di browser (Playwright) | `node tests/e2e.js http://localhost:8000/` | lulus |
| Backend `Code.gs` (dijalankan di atas tiruan layanan Google) | `node tests/test_sync_backend.js` | 20 lulus |
| Mode online: 3 browser mengisi bersamaan, konflik, putus-sambung, riwayat | `node tests/sync_server.js 8770 8771` lalu `node tests/e2e_online.js` | 26 lulus |

Uji mesin rumus, logika, dan backend juga berjalan otomatis di GitHub Actions (`.github/workflows/tests.yml`) setiap ada perubahan.

## Temuan pada file asli

Beberapa rumus dan pemetaan ekspor di file Excel mengandung kejanggalan. Contohnya: rentang COUNTIF yang bergeser, kolom ekspor yang membaca sel keliru, dan satu rumus `#REF!`. Semuanya **dipertahankan apa adanya** agar angka sama dengan Excel. Satu-satunya pengecualian adalah makro skor butir **7a-c** yang di Excel selalu error sehingga skornya tidak pernah terhitung. Rincian dan alasannya ada di [docs/TEMUAN.md](docs/TEMUAN.md).

## Kredit

METT-4: Stolton, S., Hockings, M. and Dudley, N. (2020). *Management Effectiveness Tracking Tool. Reporting Progress at Protected Area Sites: Fourth Edition.* Adaptasi Indonesia dan file Excel METT-4 KSDAE v4.4 oleh Kementerian Kehutanan, Ditjen KSDAE. Kredit lengkap ada di halaman **Tentang** aplikasi (diambil dari file Excel).

Lisensi kode aplikasi web ini belum ditetapkan. Konfirmasi dengan pemilik tool sebelum repository dipublikasikan atau dipakai pihak lain.
