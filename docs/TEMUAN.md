# Temuan pada File Excel Asli (METT-4 KSDAE v4.4)

Dokumen ini mencatat kejanggalan yang ditemukan saat file `METT 4 _Indonesia_KSDAE_VER4.4.xlsm` dipindahkan ke versi web.

Aturan yang dipakai: **file Excel adalah acuan utama.** Semua kejanggalan rumus, teks, dan ekspor di bawah **dipertahankan apa adanya**, sehingga versi web menghasilkan angka yang sama persis dengan Excel. Hanya ada **satu pengecualian** (bagian B): kesalahan pada kode makro yang membuat Excel berhenti dengan pesan error.

Kalau tim pemilik tool memutuskan untuk memperbaiki salah satu temuan, perbaiki dulu di file Excel. Setelah itu jalankan ulang `tools/extract.py` (lihat README). Versi web akan otomatis mengikuti.

Kolom **Lokasi** memakai alamat sel Excel (`Sheet!Sel`).

---

## A. Kejanggalan yang dipertahankan

### A1. Rumus perhitungan dan tampilan

| # | Lokasi | Temuan | Dampak |
|---|---|---|---|
| 1 | `Reports Data!C24` dan `C32` | Untuk level "Rendah atau tidak ada informasi", rumusnya menghitung rentang `E2:E63` / `F2:F63`. Level lain memakai `E3:E64` / `F3:F64`. | Baris "15. Lainnya" (baris 64) tidak ikut dihitung untuk level tersebut di grafik donat Tingkat Ancaman dan Keparahan. |
| 2 | `Print BA!Y27:Y31` (kolom "Pop process") | Mengambil kolom C (Ukuran Populasi) pada Detail Penilaian Spesies. Seharusnya kolom D (Kelahiran, kematian, struktur umur). | Tabel spesies di Berita Acara menampilkan Ukuran Populasi dua kali. |
| 3 | `Dashboard!Z11` | Rumus persen total "sebelumnya" merujuk `W11`, padahal sel itu kosong (tidak berisi rumus jumlah). | Persen total penilaian sebelumnya selalu tampil 0%. |
| 4 | `Dashboard!U36:AA36` | Baris ke-6 tabel spesies membaca baris 17 Detail Penilaian Spesies. Baris itu berada di luar tabel isian (baris 12–16). | Baris ke-6 di Dashboard selalu berisi 0. |
| 5 | `Reports Data!C17` | Kategori "14. Lainnya" menggabungkan baris 63 dan 64 (14. Lainnya dan 15. Lainnya). | Hanya pengelompokan. Tidak salah hitung. |
| 6 | Semua rumus yang merujuk sel kosong | Perilaku standar Excel: rumus `=A1` dengan A1 kosong menghasilkan **0**. | Sebelum data diisi, nama kawasan dan nama orang di Dashboard dan Berita Acara tampil "0", dan tanggal kosong tampil "00/01/1900". |

### A2. Data ekspor (sheet tersembunyi yang disalin makro ke `METT4_exported_data.xlsm`)

#### PA attributes data sheet

* Kolom **Value 1** (`BA2`) membaca `Atribut Kawasan!B29` (tahun kunjungan). Seharusnya B31 (Nilai penting 1). Kolom **Value 2** membaca B31. Akibatnya Nilai penting 2 (B32) tidak pernah terekspor.
* Kolom **Manager No** (`CE2`) membaca B47 (total peserta). Seharusnya B48. Jumlah Manajer/Pejabat tidak terekspor.
* Baris 51 (Kementerian/Lembaga/Pemda) dan baris 61 (Nomor kontak) tidak punya kolom ekspor.
* Kolom **Project** (`CU2`) berisi rumus rusak `='Atribut Kawasan'!#REF!`, jadi selalu bernilai `#REF!`.
* Beberapa judul kolom tertukar antara isian dan catatan: U/V (WH criteria), AG/AH (Authority), CW/CX (Responsible Name), DA/DB (Email).
* Makro hanya menyalin kolom **A sampai CZ**. Akibatnya Email, Tahun Penilaian, dan Tanggal Penilaian (kolom DA–DF) tidak ikut masuk file rekap.

#### Question_data_dump (Assessment data)

* `Q29` sumber informasi ke-2 membaca `'29.'!B4`. Seharusnya B24.
* `Q31A`, `Q35A`, `Q36A`: kolom Justifikasi (J) dan Aksi (A) membaca B16 dan B20. Padahal isiannya ada di B13 dan B17. Kedua kolom ini selalu kosong (0).
* `Q31A`, `Q35A`, `Q36A`: teks rincian sumber informasi membaca B24–B29. Padahal isiannya ada di B21–B26. Teksnya bergeser atau hilang di ekspor.
* `Q35ANA` membaca `Daftar Pertanyaan & Skor!D44` (skor maksimum). Seharusnya A44 (status tidak relevan).
* `Q37J` membaca `'4.'!B15`, yaitu justifikasi butir 4. Seharusnya justifikasi butir 37.

Versi web tetap menyimpan isian di sel yang benar-benar diketik pengguna (B13, B17, B21–B26, dst.). Akibatnya halaman Aksi, Rencana Aksi, dan Berita Acara tetap lengkap. Hanya **file ekspor** yang mengikuti rumus asli di atas.

### A3. Teks dan nilai bawaan

| # | Lokasi | Temuan |
|---|---|---|
| 1 | `Detail penilaian nilai utama!A5:A15` | Teks panduan memakai istilah "Buruk / Membaik / Memburuk". Pilihan dropdown berisi "Kurang / Meningkat / Menurun". |
| 2 | `Atribut Kawasan!B4` | Berisi "Jakarta". Kelihatannya sisa data contoh. Tetap menjadi nilai awal; pengguna perlu menggantinya. |
| 3 | `'1.'!B3` … `'38.'!B3`, `Tentang!D29:D30` | Kolom penjelasan berisi teks petunjuk sebagai **nilai sel**, bukan placeholder. Akibatnya ada dua hal: (a) aturan VBA "Penjelasan harus diisi" tidak pernah terpicu karena sel tidak kosong; (b) teks petunjuk ikut terekspor bila tidak dihapus. Di versi web, teks ini otomatis terpilih (blok) saat kolom pertama kali diklik, jadi langsung tertimpa ketikan. |
| 4 | `Print BA!B59` | Teks berisi "penilaian penilaian" (kata ganda). |
| 5 | Grafik ancaman pada Print BA (`chart8`) | Judulnya "Skor METT per Elemen Pengelolaan", padahal isinya grafik ancaman. |

---

## B. Satu-satunya perbaikan: skor butir 7a-c

**Lokasi:** kode VBA sheet `7A.` (modul `Sheet1`, prosedur `Worksheet_Deactivate`).

**Temuan:** makro menghitung skor dari `Check Box 29`, `Check Box 30`, dan `Check Box 31`. Ketiga kontrol itu **tidak ada** di sheet 7A. Checkbox pernyataan 7a–7c bernama `Check Box 8`, `9`, dan `10`. Nama inilah yang dipakai makro "tidak relevan" pada sheet yang sama. Sheet 31A, 35A, dan 36A memakai pola yang sama dan sudah benar (8, 9, 10).

**Perilaku di Excel:** saat pengguna meninggalkan sheet 7A, VBA berhenti dengan error ("The item with the specified name wasn't found"). Skor 7a-c **tidak pernah tertulis dan tetap 0**, padahal skor maksimumnya tetap dihitung 3.

**Perilaku versi web:** skor 7a-c = jumlah checkbox 8–10 yang dicentang (0–3), sama dengan 31A, 35A, dan 36A. Halaman 7A menampilkan catatan teknis tentang hal ini.

**Alasan:** ini bukan rumus atau angka. Ini makro yang gagal berjalan, dan maksud perancangnya jelas dari tiga sheet sejenis. Meniru error justru membuat satu butir penilaian tidak bisa diisi sama sekali.

> Kalau tim memutuskan perilaku Excel harus ditiru persis (7a-c selalu 0), beri tahu pengembang. Cukup ubah satu baris di `js/model.js` (bagian `buildSchema`, cabang `names.length === 0`).

---

## C. Perbedaan perilaku yang tidak terhindarkan (bukan rumus)

| Excel | Versi web |
|---|---|
| Skor ditulis VBA ke sel C6 saat pengguna **pindah sheet**. | Skor langsung diperbarui saat jawaban dipilih. Hasil angkanya sama. |
| Bila isian wajib kosong, Excel menahan pengguna di sheet itu dan menampilkan satu pesan saja (`If … ElseIf`). | Semua pesan ditampilkan sekaligus dan kotak butir ditandai kuning. Pengguna tetap bisa berpindah halaman, tetapi tombol **Tandai FINAL** terkunci sampai semua lengkap. |
| Tombol makro "Eksport Data …" menyalin baris ke `METT4_exported_data.xlsm` yang sedang terbuka. | Tombol unduh XLSX/CSV (kolom sama), ditambah halaman **Rekap Multi-Kawasan** yang menggabungkan banyak file. |
| Beberapa tombol navigasi menuju nama sheet lama yang sudah tidak ada ("METT 4 questions+scores", "METT 4 introduction", "Protected area attributes"), sehingga tidak berfungsi. | Diarahkan ke sheet yang dimaksud (Daftar Pertanyaan & Skor, Pengenalan METT 4, Atribut Kawasan). |
| `TEXT(...,"dddd")` dan `"mmmm"` mengikuti bahasa Windows/Excel pengguna. | Nama hari dan bulan selalu dalam bahasa Indonesia (seperti Excel berlokal Indonesia). Angka memakai koma desimal dan titik ribuan. |
| Ketikan "1500000" atau "17/08/2025" otomatis diubah Excel menjadi angka/tanggal. | Ditiru untuk isian satu baris. Nomor yang diawali 0 (mis. nomor HP) tetap disimpan sebagai teks. |
| Checkbox "Informasi tidak mencukupi" di ketiga lembar Detail tidak tertaut ke sel mana pun. | Disimpan di file penilaian (.json), tidak ikut ekspor. Sama seperti Excel. |
| Sheet tersembunyi (Q&A yang masih draf, Drop down lists, Data lookups, dll.) | Tetap tidak ditampilkan. Isinya tetap dipakai oleh rumus. |
| Daftar dropdown Excel menampilkan baris kosong pada sumber daftar. | Baris kosong tidak ditampilkan. |
| Grafik digambar oleh Excel. | Digambar ulang dari definisi grafik asli (jenis, seri, rentang, warna). Tampilannya mirip, tidak identik piksel per piksel. |

---

## D. Cara temuan ini diperiksa

* `tests/check_excel_cache.js` membandingkan **1.907 rumus** hasil mesin web dengan nilai yang terakhir dihitung Microsoft Excel (tersimpan di dalam file). Hasil: **0 beda**.
* `tests/compare_with_excel.py` mengisi file Excel asli dengan 6 skenario (kosong, semua tidak relevan, skor maksimal, dan 3 acak), menghitungnya dengan LibreOffice, lalu membandingkan semua 1.907 sel rumus per skenario (11.442 perbandingan). Hasil: **0 beda**, selain 2 sel yang perbedaannya berasal dari LibreOffice, bukan dari aplikasi web:
  * `PA attributes data sheet!CU2`: Excel menghasilkan `#REF!`, LibreOffice `#VALUE!`. Versi web = Excel.
  * `Print BA!B14` saat tanggal kosong: Excel menulis "00 Januari 1900", LibreOffice "30 Desember 1899". Versi web = Excel.
* `tests/test_logic.js` (240 uji) memeriksa logika makro: skor A–D = 0–3 untuk ke-38 butir, skor tambahan, "tidak relevan", pesan wajib, multi-pilih, validasi data, dan ekspor.
