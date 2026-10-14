"""
Pembaca kode VBA dari vbaProject.bin (format OLE / Compound File + kompresi MS-OVBA).
Implementasi mandiri tanpa dependensi, cukup untuk membaca source modul VBA.
"""
import struct


def _read_ole(d):
    assert d[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "bukan file OLE"
    ssz = 1 << struct.unpack_from("<H", d, 30)[0]
    mssz = 1 << struct.unpack_from("<H", d, 32)[0]
    nfat = struct.unpack_from("<I", d, 44)[0]
    dirstart = struct.unpack_from("<I", d, 48)[0]
    cutoff = struct.unpack_from("<I", d, 56)[0]
    minifatstart = struct.unpack_from("<I", d, 60)[0]
    difstart = struct.unpack_from("<I", d, 68)[0]
    ndif = struct.unpack_from("<I", d, 72)[0]

    def sec(i):
        return d[512 + i * ssz: 512 + (i + 1) * ssz]

    fatsecs = list(struct.unpack_from("<109I", d, 76))
    s = difstart
    for _ in range(ndif):
        vals = struct.unpack("<%dI" % (ssz // 4), sec(s))
        fatsecs += vals[:-1]
        s = vals[-1]
    fat = []
    for f in fatsecs[:nfat]:
        fat += struct.unpack("<%dI" % (ssz // 4), sec(f))

    def stream(start, size=None):
        out, n = [], start
        while n < 0xFFFFFFF0:
            out.append(sec(n))
            n = fat[n]
        b = b"".join(out)
        return b if size is None else b[:size]

    dirb = stream(dirstart)
    entries = []
    for i in range(len(dirb) // 128):
        e = dirb[i * 128:(i + 1) * 128]
        nl = struct.unpack_from("<H", e, 64)[0]
        entries.append(dict(name=e[:max(nl - 2, 0)].decode("utf-16le", "ignore"),
                            l=struct.unpack_from("<I", e, 68)[0], r=struct.unpack_from("<I", e, 72)[0],
                            c=struct.unpack_from("<I", e, 76)[0], st=struct.unpack_from("<I", e, 116)[0],
                            sz=struct.unpack_from("<I", e, 120)[0]))
    root = entries[0]
    ministream = stream(root["st"], root["sz"])
    minifat = []
    if minifatstart < 0xFFFFFFF0:
        mb = stream(minifatstart)
        minifat = list(struct.unpack("<%dI" % (len(mb) // 4), mb))

    def ministr(start, size):
        out = b""
        while start < 0xFFFFFFF0:
            out += ministream[start * mssz:(start + 1) * mssz]
            start = minifat[start]
        return out[:size]

    paths = {}

    def walk(idx, prefix):
        if idx == 0xFFFFFFFF:
            return
        e = entries[idx]
        walk(e["l"], prefix)
        walk(e["r"], prefix)
        p = prefix + [e["name"]]
        paths["/".join(p)] = e
        if e["c"] != 0xFFFFFFFF:
            walk(e["c"], p)

    walk(root["c"], [])

    def get(p):
        e = paths[p]
        return ministr(e["st"], e["sz"]) if e["sz"] < cutoff else stream(e["st"], e["sz"])

    return paths, get


def _decompress(c):
    assert c[0] == 1
    out = bytearray()
    i = 1
    while i < len(c):
        hdr = struct.unpack_from("<H", c, i)[0]
        i += 2
        size = (hdr & 0xFFF) + 3
        end = min(i - 2 + size, len(c))
        start = len(out)
        if not (hdr >> 15) & 1:
            out += c[i:i + 4096]
            i += 4096
            continue
        while i < end:
            fb = c[i]
            i += 1
            for bit in range(8):
                if i >= end:
                    break
                if not (fb >> bit) & 1:
                    out.append(c[i])
                    i += 1
                else:
                    tok = struct.unpack_from("<H", c, i)[0]
                    i += 2
                    dl = len(out) - start
                    bc = max((dl - 1).bit_length(), 4)
                    ln = (tok & (0xFFFF >> bc)) + 3
                    off = (tok >> (16 - bc)) + 1
                    for _ in range(ln):
                        out.append(out[-off])
    return bytes(out)


def read_vba_modules(vba_bin):
    """{nama_modul: source_code} dari isi vbaProject.bin."""
    paths, get = _read_ole(vba_bin)
    dirs = _decompress(get("VBA/dir"))
    i, mods, cur = 0, [], {}
    while i < len(dirs):
        rid, sz = struct.unpack_from("<HI", dirs, i)
        i += 6
        if rid == 9:
            sz = 6
        data = dirs[i:i + sz]
        i += sz
        if rid == 0x19:
            cur = {"name": data.decode("latin1")}
        elif rid == 0x1A:
            cur["stream"] = data.decode("latin1")
        elif rid == 0x31:
            cur["off"] = struct.unpack("<I", data)[0]
        elif rid == 0x2B:
            mods.append(cur)
    out = {}
    for m in mods:
        raw = get("VBA/" + m["stream"])
        out[m["name"]] = _decompress(raw[m["off"]:]).decode("cp1252", "replace")
    return out


if __name__ == "__main__":
    import sys
    import zipfile
    z = zipfile.ZipFile(sys.argv[1])
    for k, v in read_vba_modules(z.read("xl/vbaProject.bin")).items():
        print("=" * 20, k)
        print(v)
