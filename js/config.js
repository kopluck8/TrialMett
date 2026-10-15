/*
 * Konfigurasi METT-4 Web
 * ----------------------
 * SYNC_URL kosong  → mode lokal: data tersimpan di browser masing-masing + file .json.
 * SYNC_URL diisi   → mode ONLINE: semua orang yang membuka link mengisi satu penilaian
 *                    yang sama, tersimpan di Google Sheet (lihat docs/PANDUAN-ONLINE.md).
 *
 * Isi dengan URL aplikasi web Google Apps Script yang berakhiran /exec, contoh:
 *   SYNC_URL: 'https://script.google.com/macros/s/AKfycb.../exec',
 */
window.METT_CONFIG = {
  SYNC_URL: 'https://script.google.com/macros/s/AKfycbxJuLMkYR723V4T1XF4h5MZny6zb-B6UMW4dtF9rjshAPMXDeInmNgQfS8fj7LqPLZpSA/exec',
  POLL_SECONDS: 6, // seberapa sering (detik) aplikasi mengambil isian terbaru dari pengguna lain
};
