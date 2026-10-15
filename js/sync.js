/*
 * METT-4 Web — Sinkronisasi data bersama (Google Apps Script / Google Sheets)
 * ------------------------------------------------------------------
 * Setiap isian adalah satu "sel" (alamat sel Excel yang sama). Yang dikirim
 * ke server hanya sel yang berubah, sehingga beberapa orang/kelompok bisa
 * mengisi bagian berbeda pada saat bersamaan tanpa saling menimpa.
 *
 * Aturan bila dua orang mengubah SEL YANG SAMA: perubahan yang terakhir
 * terkirim yang berlaku; nilai lama tetap tercatat di sheet "log" sehingga
 * dapat dikembalikan dari halaman Riwayat.
 *
 * Isian tetap tersimpan di browser saat sinyal putus dan dikirim otomatis
 * ketika koneksi kembali.
 */
(function (root) {
  'use strict';
  const enc = (v) => JSON.stringify(v === undefined ? null : v);
  const LS = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* penuh / diblokir */ } },
  };
  function hash(s) { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return (h >>> 0).toString(36); }

  class Sync {
    /**
     * @param model  MettModel.Model
     * @param url    URL aplikasi web Apps Script (…/exec)
     * @param opts   {pollMs, pushDelayMs, fetch}
     */
    constructor(model, url, opts) {
      opts = opts || {};
      this.model = model;
      this.url = url;
      this.pollMs = opts.pollMs || 6000;
      this.pushDelayMs = opts.pushDelayMs || 700;
      this.fetch = opts.fetch || ((...a) => root.fetch(...a));
      this.key = 'mett4.sync.' + hash(url);
      this.synced = {};       // kunci -> nilai server terakhir yang diketahui (JSON)
      this.lastRev = 0;
      this.status = 'connecting';
      this.error = null;
      this.lastOk = null;
      this.pushing = null;
      this.pulling = null;
      this.listeners = [];
      this.clientId = LS.get('mett4.client') || ('c' + Math.random().toString(36).slice(2, 10));
      LS.set('mett4.client', this.clientId);
      this.loaded = false;
    }
    get by() { return LS.get('mett4.by') || ''; }
    set by(v) { LS.set('mett4.by', v || ''); }
    onStatus(fn) { this.listeners.push(fn); }
    setStatus(s, err) { this.status = s; this.error = err || null; for (const fn of this.listeners) fn(this); }

    // ---------------------------------------------------------- cache lokal
    loadLocal() {
      try {
        const o = JSON.parse(LS.get(this.key) || 'null');
        if (o) { this.synced = o.synced || {}; this.lastRev = o.lastRev || 0; return o.state || null; }
      } catch (e) { /* abaikan */ }
      return null;
    }
    saveLocal() {
      LS.set(this.key, JSON.stringify({ synced: this.synced, lastRev: this.lastRev, state: JSON.parse(this.model.toFile()) }));
    }

    // ---------------------------------------------------------- selisih
    dirtyKeys() {
      const local = this.model.syncMap();
      const out = [];
      for (const k of Object.keys(local)) if (enc(local[k]) !== (this.synced[k] === undefined ? enc(null) : this.synced[k])) out.push(k);
      // sel yang belum pernah ada di server dan bernilai kosong/false tidak perlu dikirim (mengurangi catatan riwayat)
      // (kecuali anggota kelompok jawaban, yang harus selalu dikirim utuh agar tetap konsisten)
      return this.model.expandSyncGroups(out).filter((k) => this.model.inSyncGroup(k) || !(this.synced[k] === undefined && !local[k]));
    }
    pendingCount() { return this.dirtyKeys().length; }

    // ---------------------------------------------------------- jaringan
    async request(method, params, body) {
      const q = new URLSearchParams(Object.assign({ _: Date.now() }, params || {})).toString();
      const res = await this.fetch(this.url + (this.url.includes('?') ? '&' : '?') + q, method === 'POST'
        ? { method: 'POST', body: JSON.stringify(body), redirect: 'follow' }   // text/plain: tanpa preflight CORS
        : { method: 'GET', redirect: 'follow' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const j = await res.json();
      if (!j.ok) throw new Error(j.error || 'Server menolak permintaan');
      return j;
    }

    async start() {
      const st = this.loadLocal();
      if (st) this.model.load(st);
      this.loaded = true;
      await this.pull();
      await this.push();
      this.timer = setInterval(() => { if (!root.document || root.document.visibilityState !== 'hidden') this.tick(); }, this.pollMs);
      if (root.addEventListener) {
        root.addEventListener('online', () => this.tick());
        if (root.document) root.document.addEventListener('visibilitychange', () => { if (root.document.visibilityState === 'visible') this.tick(); });
      }
    }
    stop() { clearInterval(this.timer); }
    async tick() { await this.push(); await this.pull(); }
    schedulePush() {
      this.saveLocal();
      this.setStatus(this.status === 'offline' || this.status === 'error' ? this.status : 'pending');
      clearTimeout(this._pt);
      this._pt = setTimeout(() => this.push(), this.pushDelayMs);
    }

    async pull() {
      if (this.pulling) return this.pulling;
      this.pulling = (async () => {
        try {
          const j = await this.request('GET', { action: 'pull', since: this.lastRev });
          const local = this.model.syncMap();
          const dirty = new Set(this.dirtyKeys()); // sudah diperluas per kelompok (mis. 4 opsi satu butir)
          const apply = [];
          for (const c of j.changes) {
            const remote = enc(c.v);
            const has = Object.prototype.hasOwnProperty.call(local, c.k);
            const loc = has ? enc(local[c.k]) : enc(null);
            // ada isian lokal yang belum terkirim → jangan ditimpa (akan dikirim dan berlaku sebagai yang terakhir)
            if (!dirty.has(c.k) && loc !== remote) apply.push({ k: c.k, v: c.v });
            this.synced[c.k] = remote;
          }
          this.lastRev = Math.max(this.lastRev, j.rev || 0);
          if (apply.length) this.model.applyRemote(apply);
          this.saveLocal();
          this.lastOk = new Date();
          this.setStatus(this.dirtyKeys().length ? 'pending' : 'ok');
          return apply;
        } catch (e) {
          this.setStatus(root.navigator && root.navigator.onLine === false ? 'offline' : 'error', e.message);
          return [];
        } finally { this.pulling = null; }
      })();
      return this.pulling;
    }

    async push() {
      if (this.pushing) return this.pushing;
      const keys = this.dirtyKeys();
      if (!keys.length) return 0;
      const local = this.model.syncMap();
      const changes = keys.map((k) => ({ k, v: k in local ? local[k] : null }));
      this.pushing = (async () => {
        this.setStatus('syncing');
        try {
          for (let i = 0; i < changes.length; i += 2000) {
            const part = changes.slice(i, i + 2000);
            await this.request('POST', {}, { action: 'push', by: this.by || 'Tanpa nama', client: this.clientId, changes: part });
            for (const c of part) this.synced[c.k] = enc(c.v);
          }
          this.saveLocal();
          this.lastOk = new Date();
          this.setStatus(this.dirtyKeys().length ? 'pending' : 'ok');
          return changes.length;
        } catch (e) {
          this.setStatus(root.navigator && root.navigator.onLine === false ? 'offline' : 'error', e.message);
          return 0;
        } finally { this.pushing = null; }
      })();
      const n = await this.pushing;
      if (n && this.dirtyKeys().length) this.schedulePush(); // ada ketikan baru selama pengiriman
      return n;
    }

    async log(limit) { return (await this.request('GET', { action: 'log', limit: limit || 300 })).log; }
  }

  root.MettSync = { Sync };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.MettSync;
})(typeof window !== 'undefined' ? window : globalThis);
