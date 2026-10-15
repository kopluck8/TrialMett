# Panduan Mode Online (Isian Bersama 3 Kelompok)

Dengan mode online, **semua orang yang membuka link aplikasi mengisi satu penilaian yang sama**. Isian dari kelompok 1, 2, dan 3 langsung tergabung dan terlihat oleh semua (diperbarui otomatis setiap ±6 detik). Datanya disimpan di **Google Sheet milik admin**.

```
 Kelompok 1 ─┐
 Kelompok 2 ─┼──►  Aplikasi (GitHub Pages)  ──►  Google Apps Script  ──►  Google Sheet admin
 Kelompok 3 ─┘         (link yang dibagikan)          (backend/Code.gs)       sheet "data" + "log"
```

Waktu pemasangan: sekitar 15 menit, cukup sekali. Yang dibutuhkan hanya akun Google (pribadi atau kantor).

---

## Langkah 1 — Siapkan Google Sheet dan backend

1. Buka <https://sheets.new> untuk membuat Google Sheet baru. Beri nama, misalnya **`METT4 – TN Baluran 2025 (data bersama)`**.
2. Menu **Ekstensi → Apps Script**.
3. Hapus isi `Code.gs` yang ada. Salin **seluruh isi** file [`backend/Code.gs`](../backend/Code.gs) dari repository ini, tempel, lalu klik **Simpan** (ikon disket).
4. Di bilah atas, pilih fungsi **`setup`**, lalu klik **Jalankan**.
   * Google akan meminta izin: **Tinjau izin → pilih akun → Lanjutan → Buka (tidak aman) → Izinkan**.
   * Peringatan "aplikasi belum diverifikasi" ini wajar untuk skrip buatan sendiri.
   * Setelah selesai, Google Sheet Anda punya dua sheet baru: **`data`** dan **`log`**.

## Langkah 2 — Terbitkan sebagai aplikasi web

1. Di editor Apps Script: **Terapkan (Deploy) → Deployment baru**.
2. Ikon roda gigi → **Aplikasi web**.
3. Isi:
   * Deskripsi: `METT4 v1`
   * **Jalankan sebagai: Saya** (akun Anda)
   * **Yang memiliki akses: Siapa saja** (*Anyone*)
4. Klik **Terapkan**, lalu salin **URL aplikasi web** (berakhiran `/exec`).

> Pilihan "Siapa saja" diperlukan agar anggota kelompok tidak perlu login. Data tetap tersimpan di akun Anda. Mereka hanya bisa membaca dan mengubah isian penilaian lewat aplikasi, tidak bisa membuka Google Sheet Anda secara langsung.

## Langkah 3 — Hubungkan aplikasi

1. Di repository GitHub, buka file **`js/config.js`** → ikon pensil (Edit).
2. Isi URL tadi:

   ```js
   window.METT_CONFIG = {
     SYNC_URL: 'https://script.google.com/macros/s/AKfycb.../exec',
     POLL_SECONDS: 6,
   };
   ```

3. **Commit changes**. Tunggu 1–2 menit hingga GitHub Pages diperbarui.
4. Buka link aplikasi. Di pojok kanan atas akan muncul **● Tersinkron hh.mm.ss**. Aplikasi meminta nama pengisi; isi misalnya `Ani – Kelompok 2`. Nama ini dicatat di Riwayat.

Bagikan link aplikasi ke ketiga kelompok. Selesai.

---

## Cara kerja saat dipakai bersama

* **Yang dikirim hanya isian yang berubah**, per kolom. Kelompok yang mengisi bagian berbeda (mis. K1 Atribut Kawasan, K2 butir 1–20, K3 Ancaman) tidak saling menimpa.
* **Bila dua orang mengubah kolom yang sama**, yang terakhir terkirim yang berlaku. Pilihan A–D satu butir dikirim sebagai satu paket, jadi satu butir tidak mungkin punya dua jawaban.
* **Yang sedang mengetik tidak terganggu.** Kolom yang sedang diketik tidak ditimpa oleh pembaruan dari orang lain.
* **Sinyal putus?** Isian tetap tersimpan di perangkat dan dikirim otomatis saat tersambung lagi. Status di kanan atas menjadi *Offline / Gagal tersambung — n perubahan menunggu*.
* **Salah hapus atau tertimpa?** Buka menu **Data → Riwayat Perubahan**, cari baris yang dimaksud, lalu klik **Kembalikan**. Semua perubahan tercatat: waktu, nama pengisi, nilai lama, dan nilai baru.
* **Skor, dashboard, Berita Acara** dihitung di perangkat masing-masing dari data bersama yang sama, dengan rumus Excel yang sama. Hasilnya identik di semua perangkat.

### Saran pembagian kerja

Pilihan Anda adalah *satu formulir bersama, bebas*: semua orang boleh mengisi bagian mana saja. Agar tidak terjadi saling timpa, sepakati di awal sesi siapa mengerjakan bagian mana. Contoh:

| Kelompok | Bagian |
|---|---|
| 1 | Atribut Kawasan, Penilaian Ancaman, butir 1–13 (Perencanaan & Input) |
| 2 | butir 14–31A (Proses) |
| 3 | butir 32–38 + Detail nilai/spesies/habitat (Output & Outcome), Aksi |

Kotak nomor di menu kiri menunjukkan butir yang sudah lengkap (hijau), perlu dilengkapi (kuning), atau tidak relevan (abu-abu), untuk semua kelompok sekaligus.

---

## Tugas admin

| Kapan | Apa |
|---|---|
| Akhir setiap sesi | **Data → Simpan cadangan (.json)**. Cadangan ini bisa diunggah kembali bila diperlukan. |
| Selesai penilaian | Lengkapi semua butir → **Tandai sebagai FINAL** → cetak Berita Acara → simpan cadangan → unduh XLSX. |
| Penilaian kawasan/tahun berikutnya | Buat Google Sheet **baru** (ulangi Langkah 1–3) lalu ganti `SYNC_URL`. Sheet lama tetap menjadi arsip. |
| Mengubah `Code.gs` | **Terapkan → Kelola deployment → Edit (pensil) → Versi: Versi baru → Terapkan.** Cara ini mempertahankan URL. Bila membuat *deployment baru*, URL berubah dan `js/config.js` harus diperbarui. |
| Melihat data mentah | Buka Google Sheet. Sheet `data` berisi satu baris per isian (alamat sel sama dengan Excel). Sheet `log` berisi seluruh riwayat. Google Sheets juga menyimpan *Histori versi* sendiri (File → Histori versi). |

Untuk kembali ke mode lokal (tanpa server), kosongkan `SYNC_URL` di `js/config.js`.

---

## Risiko dan batasan (harap dibaca)

1. **Siapa pun yang mendapat link bisa mengubah atau menghapus isian.** Ini sesuai pilihan "siapa pun yang punya link". Repository GitHub bersifat publik, sehingga URL backend di `js/config.js` juga bisa dilihat siapa saja.
   * Mitigasi yang sudah ada: Riwayat + tombol Kembalikan, cadangan .json, dan Histori versi Google Sheets.
   * Bila nanti dirasa perlu, aplikasi dapat ditambah **kode akses** sederhana tanpa mengubah data yang sudah ada.
2. **Bukan sinkron instan.** Perubahan orang lain muncul dalam beberapa detik (default 6 detik), bukan seketika seperti Google Docs.
3. **Kuota Google Apps Script.** Akun Google punya batas jumlah eksekusi skrip yang berjalan bersamaan dan batas harian. Untuk 3 kelompok (±10–20 perangkat) dengan sinkron setiap 6 detik, beban ini diperkirakan masih wajar. [Medium confidence — berdasarkan pola penggunaan Apps Script pada umumnya; angka kuota pastinya tidak saya cek.] *This needs verification:* lihat halaman kuota resmi Apps Script. Bila status sering *Gagal tersambung*, naikkan `POLL_SECONDS` menjadi 10–15.
4. **Pengujian.** Kode `backend/Code.gs` sudah diuji di komputer pengembang dengan meniru layanan Google. Pengujian mencakup 3 browser mengisi bersamaan, putus-sambung koneksi, dan pengembalian dari riwayat. Namun kode ini **belum dijalankan di server Google yang sebenarnya**.
   * Lakukan uji coba singkat dengan 2–3 perangkat sebelum sesi penilaian resmi.
   * Bila status kanan atas tetap merah, buka menu **Data**: pesan galatnya tampil di sana.
5. Satu Google Sheet = **satu penilaian (satu kawasan)**.
