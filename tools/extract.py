#!/usr/bin/env python3
"""
Ekstraktor METT-4 KSDAE  (Excel .xlsm  ->  data/workbook.js)

Membaca file Excel METT-4 asli dan menyalin SELURUH isinya apa adanya ke
data/workbook.js agar dapat dipakai aplikasi web:

  * semua nilai sel (teks pertanyaan, opsi jawaban, panduan, daftar dropdown, dst.)
  * semua rumus Excel, disalin verbatim (teks rumusnya, bukan hasil terjemahan)
  * format angka (persen, tanggal, dll.)
  * kontrol form (checkbox / option button): label, sel tertaut, posisi, makro
  * aturan VBA per sheet (skor pilihan A-D, skor checkbox tambahan, isian wajib,
    multi-pilih dropdown) -- diurai otomatis dari kode VBA di file
  * validasi data (dropdown) dan sumber daftarnya
  * definisi Tabel Excel (untuk structured reference seperti Table2[[#This Row],...])
  * definisi grafik (jenis, judul, seri, rentang data, warna)

Pemakaian:
    python3 tools/extract.py "METT 4 _Indonesia_KSDAE_VER4.4.xlsm" data/workbook.js

Butuh: Python 3.9+, openpyxl.  Tidak ada nilai yang diubah, dibulatkan, atau diterjemahkan.
"""
import datetime
import json
import os
import re
import sys
import zipfile
import warnings
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_reader import read_vba_modules  # noqa: E402

warnings.filterwarnings("ignore")

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

EXCEL_EPOCH = datetime.datetime(1899, 12, 30)


def to_serial(dt):
    if isinstance(dt, datetime.datetime):
        d = dt - EXCEL_EPOCH
    elif isinstance(dt, datetime.date):
        d = datetime.datetime(dt.year, dt.month, dt.day) - EXCEL_EPOCH
    elif isinstance(dt, datetime.time):
        return (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86400
    else:
        return dt
    return d.days + d.seconds / 86400


def norm_path(base, target):
    if target.startswith("/"):
        return target[1:]
    parts = base.split("/")[:-1] + target.split("/")
    out = []
    for p in parts:
        if p == "..":
            out.pop()
        elif p and p != ".":
            out.append(p)
    return "/".join(out)


def read_rels(z, path):
    d = os.path.dirname(path)
    rp = f"{d}/_rels/{os.path.basename(path)}.rels"
    if rp not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rp))
    out = {}
    for r in root.findall("rel:Relationship", NS):
        tgt = r.get("Target")
        keep_raw = tgt.startswith("#") or r.get("TargetMode") == "External"
        out[r.get("Id")] = (tgt if keep_raw else norm_path(path, tgt), r.get("Type").split("/")[-1])
    return out


def text_of(el):
    return "".join(t.text or "" for t in el.iter() if t.tag.endswith("}t"))


# --------------------------------------------------------------------------- VML
def parse_vml(raw):
    """Kembalikan {shapeId(int): {...}} dari file VML (label, anchor, link, makro)."""
    s = raw.decode("utf-8", "replace")
    shapes = {}
    for m in re.finditer(r"<v:shape\b(.*?)</v:shape>", s, re.S):
        body = m.group(0)
        idm = re.search(r'o:spid="_x0000_s(\d+)"', body) or re.search(r'\bid="_x0000_s(\d+)"', body)
        if not idm:
            continue
        sid = int(idm.group(1))
        tb = re.search(r"<v:textbox.*?</v:textbox>", body, re.S)
        label = ""
        if tb:
            t = re.sub(r"<br\s*/?>", "\n", tb.group(0))
            t = re.sub(r"<[^>]+>", "", t)
            t = (t.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                 .replace("&quot;", '"').replace("&#10;", "\n").replace("&nbsp;", " "))
            label = re.sub(r"[ \t\r]+", " ", t).strip()
        def tag(name):
            mm = re.search(rf"<x:{name}>(.*?)</x:{name}>", body, re.S)
            return mm.group(1).strip() if mm else None
        anchor = tag("Anchor")
        if anchor:
            anchor = [int(x) for x in anchor.replace("\n", "").split(",")]
        shapes[sid] = {
            "label": label,
            "anchor": anchor,
            "objectType": (re.search(r'ObjectType="([^"]+)"', body) or [None, None])[1],
            "macro": tag("FmlaMacro"),
            "link": tag("FmlaLink"),
            "checked": tag("Checked"),
        }
    return shapes


# --------------------------------------------------------------------------- VBA
def parse_question_vba(code):
    """Urai Worksheet_Deactivate pada sheet pertanyaan menjadi aturan skor & isian wajib."""
    if "Worksheet_Deactivate" not in code:
        return None
    rule = {"kind": None, "na": None, "naReasonCell": None, "options": [], "scoreBoxes": [],
            "scoreCell": None, "required": []}
    m = re.search(r'Shapes\("(Check Box \d+)"\)\.OLEFormat\.Object\.Value = 1 And Cells\((\d+), (\d+)\)\.Value = ""\s*Then\s*'
                  r'Me\.Activate\s*MsgBox "([^"]*)"', code)
    if m:
        rule["na"] = m.group(1)
        rule["naReasonCell"] = get_column_letter(int(m.group(3))) + m.group(2)
        rule["required"].append({"when": "na", "cell": rule["naReasonCell"], "msg": m.group(4)})
    opts = re.findall(r'Shapes\("(Option Button \d+)"\)\.OLEFormat\.Object\.Value = 1 Then\s*Cells\((\d+), (\d+)\)\.Value = (\d+)', code)
    if opts:
        rule["kind"] = "radio"
        for name, r, c, v in opts:
            rule["options"].append({"name": name, "score": int(v)})
            rule["scoreCell"] = get_column_letter(int(c)) + r
    boxes = re.findall(r'Shapes\("(Check Box \d+)"\)\.OLEFormat\.Object\.Value = 1 Then\s*score = score \+ 1', code)
    if boxes:
        rule["kind"] = "additive"
        rule["scoreBoxes"] = boxes
        mm = re.search(r"Cells\((\d+), (\d+)\)\.Value = score", code)
        rule["scoreCell"] = get_column_letter(int(mm.group(2))) + mm.group(1)
    for r, c, msg in re.findall(r'If Cells\((\d+), (\d+)\)\.Value = "" Then\s*Me\.Activate\s*MsgBox "([^"]*)"', code):
        rule["required"].append({"when": "answered", "cell": get_column_letter(int(c)) + r, "msg": msg})
    return rule


def parse_multiselect_vba(code):
    """Worksheet_Change 'multi-pilih dropdown' -> daftar alamat sel/rentang."""
    if "Worksheet_Change" not in code or "Oldvalue & \", \" & Newvalue" not in code:
        return []
    out = []
    for a in re.findall(r'Target\.Address = "([^"]+)"', code):
        out.append(a.replace("$", ""))
    for a in re.findall(r'Intersect\(Target, Range\("([^"]+)"\)\)', code):
        out.append(a.replace("$", ""))
    return out


def parse_na_macros(modules):
    """Makro checkbox 'tidak relevan' (Module5/7/8): checkbox -> kontrol yang dinonaktifkan."""
    out = {}
    for name, code in modules.items():
        for sub, body in re.findall(r"Sub (\w+)\(\)(.*?)End Sub", code, re.S):
            m = re.search(r'If ActiveSheet\.Shapes\("(Check Box \d+)"\)\.OLEFormat\.Object\.Value <> 1 Then(.*?)Else(.*)', body, re.S)
            if not m:
                continue
            enabled = re.findall(r'Shapes\("([^"]+)"\)\.OLEFormat\.Object\.Enabled = True', m.group(2))
            disabled = re.findall(r'Shapes\("([^"]+)"\)\.OLEFormat\.Object\.Enabled = False', m.group(3))
            cleared = re.findall(r'Shapes\("([^"]+)"\)\.OLEFormat\.Object\.Value = False', m.group(3))
            setc = re.findall(r"Cells\((\d+), (\d+)\)\.Value = (\d+)", m.group(3))
            out[sub] = {"checkbox": m.group(1), "enableWhenOff": enabled, "disableWhenOn": disabled,
                        "clearWhenOn": cleared,
                        "setWhenOn": [{"cell": get_column_letter(int(c)) + r, "value": int(v)} for r, c, v in setc]}
    return out


# --------------------------------------------------------------------------- charts
def parse_chart(raw):
    root = ET.fromstring(raw)
    chart = root.find("c:chart", NS)
    out = {"title": None, "types": [], "series": [], "legend": None}
    t = chart.find("c:title", NS)
    if t is not None:
        tx = t.find(".//c:tx", NS)
        if tx is not None and tx.find("c:strRef", NS) is not None:
            out["titleRef"] = tx.find("c:strRef/c:f", NS).text
        out["title"] = text_of(t).strip() or None
    out["autoTitleDeleted"] = (chart.find("c:autoTitleDeleted", NS) is not None and
                               chart.find("c:autoTitleDeleted", NS).get("val") == "1")
    plot = chart.find("c:plotArea", NS)
    for ct in plot:
        tag = ct.tag.split("}")[1]
        if not tag.endswith("Chart"):
            continue
        info = {"type": tag}
        for k in ("barDir", "grouping", "radarStyle", "holeSize", "firstSliceAng", "gapWidth", "overlap", "varyColors"):
            e = ct.find(f"c:{k}", NS)
            if e is not None:
                info[k] = e.get("val")
        out["types"].append(info)
        for ser in ct.findall("c:ser", NS):
            s = {"chartType": tag}
            tx = ser.find("c:tx", NS)
            if tx is not None:
                f = tx.find(".//c:f", NS)
                s["nameRef"] = f.text if f is not None else None
                lit = tx.find("c:v", NS)
                s["name"] = (lit.text if lit is not None else text_of(tx).strip()) or None
            for k, key in (("cat", "catRef"), ("val", "valRef")):
                e = ser.find(f"c:{k}", NS)
                if e is not None:
                    f = e.find(".//c:f", NS)
                    s[key] = f.text if f is not None else None
            sp = ser.find("c:spPr", NS)
            col = None
            if sp is not None:
                c1 = sp.find("a:solidFill/a:srgbClr", NS)
                c2 = sp.find("a:ln/a:solidFill/a:srgbClr", NS)
                col = (c1 if c1 is not None else c2)
                col = col.get("val") if col is not None else None
                if col is None:
                    sc = sp.find(".//a:schemeClr", NS)
                    col = ("scheme:" + sc.get("val")) if sc is not None else None
            s["color"] = col
            pts = []
            for dpt in ser.findall("c:dPt", NS):
                idx = dpt.find("c:idx", NS).get("val")
                cc = dpt.find(".//a:solidFill/a:srgbClr", NS)
                sc = dpt.find(".//a:solidFill/a:schemeClr", NS)
                pts.append({"idx": int(idx), "color": cc.get("val") if cc is not None else
                            (("scheme:" + sc.get("val")) if sc is not None else None)})
            if pts:
                s["pointColors"] = pts
            lbl = ser.find("c:dLbls", NS)
            if lbl is not None:
                sv = lbl.find("c:showVal", NS)
                sp_ = lbl.find("c:showPercent", NS)
                s["showVal"] = sv is not None and sv.get("val") == "1"
                s["showPercent"] = sp_ is not None and sp_.get("val") == "1"
            out["series"].append(s)
    for ax in plot.findall("c:valAx", NS):
        sc = ax.find("c:scaling", NS)
        mx = sc.find("c:max", NS) if sc is not None else None
        mn = sc.find("c:min", NS) if sc is not None else None
        out.setdefault("valAxis", {})
        if mx is not None:
            out["valAxis"]["max"] = float(mx.get("val"))
        if mn is not None:
            out["valAxis"]["min"] = float(mn.get("val"))
        nf = ax.find("c:numFmt", NS)
        if nf is not None:
            out["valAxis"]["numFmt"] = nf.get("formatCode")
    lg = chart.find("c:legend", NS)
    if lg is not None:
        pos = lg.find("c:legendPos", NS)
        out["legend"] = pos.get("val") if pos is not None else "r"
    return out


# --------------------------------------------------------------------------- main
def cell_value(c):
    v = c.value
    if v is None:
        return None
    if hasattr(v, "text") and type(v).__name__ == "ArrayFormula":
        return {"f": v.text}
    if isinstance(v, str) and v.startswith("="):
        return {"f": v}
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return {"v": to_serial(v), "date": True}
    if type(v).__name__ == "CellRichText":
        return {"v": str(v)}
    return {"v": v}


def main(src, dst):
    z = zipfile.ZipFile(src)
    wb = openpyxl.load_workbook(src, keep_vba=False)
    wb_cached = openpyxl.load_workbook(src, data_only=True)
    wbxml = ET.fromstring(z.read("xl/workbook.xml"))
    wbrels = read_rels(z, "xl/workbook.xml")
    modules = read_vba_modules(z.read("xl/vbaProject.bin")) if "xl/vbaProject.bin" in z.namelist() else {}
    na_macros = parse_na_macros(modules)

    out = {"meta": {"source": os.path.basename(src),
                    "extractedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "note": "Dihasilkan otomatis oleh tools/extract.py. Jangan diedit manual."},
           "sheets": [], "tables": {}, "charts": [], "definedNames": {}, "naMacros": na_macros}

    for dn in wbxml.findall("m:definedNames/m:definedName", NS):
        if dn.get("localSheetId") is None and not dn.get("name").startswith("_xl"):
            out["definedNames"][dn.get("name")] = dn.text

    sheet_elems = wbxml.findall("m:sheets/m:sheet", NS)
    cached_values = {}
    for se in sheet_elems:
        name = se.get("name")
        path = wbrels[se.get(f"{{{NS['r']}}}id")][0]
        raw = z.read(path)
        sx = ET.fromstring(raw)
        ws = wb[name]
        wsc = wb_cached[name]
        pr = sx.find("m:sheetPr", NS)
        codename = pr.get("codeName") if pr is not None else None
        sv = sx.find("m:sheetViews/m:sheetView", NS)
        sheet = {"name": name, "state": se.get("state", "visible"), "codeName": codename,
                 "showZeros": not (sv is not None and sv.get("showZeros") == "0"),
                 "cells": {}, "merged": [str(r) for r in ws.merged_cells.ranges],
                 "controls": [], "validations": [], "vba": None, "multiSelect": [],
                 "hiddenRows": [], "colWidths": {}}
        for row in ws.iter_rows():
            for c in row:
                cv = cell_value(c)
                if cv is None:
                    continue
                if c.number_format and c.number_format != "General":
                    cv["nf"] = c.number_format
                sheet["cells"][c.coordinate] = cv
                if "f" in cv:
                    cached = wsc[c.coordinate].value
                    if isinstance(cached, (datetime.datetime, datetime.date, datetime.time)):
                        cached = to_serial(cached)
                    cached_values.setdefault(name, {})[c.coordinate] = cached
        # sel tidak terkunci (= sel isian pengguna pada sheet terproteksi), hyperlink, komentar
        sheet["protected"] = bool(ws.protection.sheet)
        sheet["unlocked"] = []
        sheet["links"] = {}
        sheet["comments"] = {}
        for row in ws.iter_rows(max_row=min(ws.max_row, 400)):
            for c in row:
                if c.protection is not None and c.protection.locked is False:
                    sheet["unlocked"].append(c.coordinate)
                if c.hyperlink is not None:
                    sheet["links"][c.coordinate] = c.hyperlink.location or c.hyperlink.target
                if c.comment is not None:
                    sheet["comments"][c.coordinate] = c.comment.text
        for idx, dim in ws.row_dimensions.items():
            if dim.hidden:
                sheet["hiddenRows"].append(idx)
        for col, dim in ws.column_dimensions.items():
            if dim.width:
                sheet["colWidths"][col] = round(dim.width, 2)

        # ---- validations (standard + x14)
        for dv in ws.data_validations.dataValidation:
            sheet["validations"].append({"sqref": str(dv.sqref), "type": dv.type, "formula1": dv.formula1,
                                         "formula2": dv.formula2, "operator": dv.operator,
                                         "errorStyle": dv.errorStyle or "stop",
                                         "showError": bool(dv.showErrorMessage),
                                         "prompt": dv.prompt, "error": dv.error})
        for x in sx.iter(f"{{{NS['x14']}}}dataValidation"):
            f = x.find("x14:formula1/xm:f", NS)
            sq = x.find("xm:sqref", NS)
            sheet["validations"].append({"sqref": sq.text if sq is not None else "", "type": x.get("type"),
                                         "formula1": f.text if f is not None else None,
                                         "prompt": x.get("prompt")})

        # ---- controls (form controls)
        rels = read_rels(z, path)
        vml = {}
        for rid, (target, typ) in rels.items():
            if typ == "vmlDrawing":
                vml.update(parse_vml(z.read(target)))
        for ctl in sx.iter(f"{{{NS['m']}}}control"):
            sid = int(ctl.get("shapeId"))
            rid = ctl.get(f"{{{NS['r']}}}id")
            cp = {}
            if rid in rels and rels[rid][1] == "ctrlProp":
                cpx = ET.fromstring(z.read(rels[rid][0]))
                cp = dict(cpx.attrib)
            v = vml.get(sid, {})
            link = cp.get("fmlaLink") or v.get("link")
            anchor = v.get("anchor")
            sheet["controls"].append({
                "name": ctl.get("name"), "shapeId": sid,
                "type": cp.get("objectType") or v.get("objectType"),
                "label": v.get("label", ""),
                "link": link.replace("&amp;", "&") if link else None,
                "macro": (cp.get("fmlaMacro") or v.get("macro") or "").replace("[0]!", "") or None,
                "firstButton": cp.get("firstButton") == "1",
                "checked": cp.get("checked"),
                "col": anchor[0] if anchor else None, "row": anchor[2] + 1 if anchor else None,
                "anchor": anchor,
            })
        sheet["controls"].sort(key=lambda c: (c["row"] or 0, c["col"] or 0))

        # ---- VBA rules
        code = modules.get(codename, "") if codename else ""
        sheet["vba"] = parse_question_vba(code)
        sheet["multiSelect"] = parse_multiselect_vba(code)

        # ---- charts & images
        for rid, (target, typ) in rels.items():
            if typ != "drawing":
                continue
            dx = ET.fromstring(z.read(target))
            drels = read_rels(z, target)
            for anc in list(dx):
                gfs = anc.findall(".//c:chart", NS)
                # teks judul di dalam grup (kotak teks di atas grafik), urut sesuai dokumen
                texts = []
                for sp in anc.iter(f"{{{NS['xdr']}}}sp"):
                    tt = "".join(t.text or "" for t in sp.iter(f"{{{NS['a']}}}t")).strip()
                    if tt:
                        texts.append(re.sub(r"\s+", " ", tt))
                frm = anc.find("xdr:from", NS)
                to = anc.find("xdr:to", NS)
                pos = None
                if frm is not None:
                    pos = {"fromCol": int(frm.find("xdr:col", NS).text), "fromRow": int(frm.find("xdr:row", NS).text),
                           "toCol": int(to.find("xdr:col", NS).text) if to is not None else None,
                           "toRow": int(to.find("xdr:row", NS).text) if to is not None else None}
                for gi, gf in enumerate(gfs):
                    cpath = drels[gf.get(f"{{{NS['r']}}}id")][0]
                    ch = parse_chart(z.read(cpath))
                    ch.update({"sheet": name, "file": os.path.basename(cpath), "pos": pos,
                               "groupTitle": texts[gi] if gi < len(texts) else None})
                    out["charts"].append(ch)
                if not gfs and texts and anc.find(".//xdr:sp", NS) is not None and pos:
                    hl = anc.find(".//a:hlinkClick", NS)
                    target = None
                    if hl is not None:
                        rid = hl.get(f"{{{NS['r']}}}id")
                        target = drels.get(rid, (None,))[0] if rid in drels else None
                    sheet.setdefault("shapes", []).append({"text": " ".join(texts), "pos": pos, "link": target})

        # ---- tables
        for rid, (target, typ) in rels.items():
            if typ == "table":
                tx = ET.fromstring(z.read(target))
                out["tables"][tx.get("name")] = {
                    "sheet": name, "ref": tx.get("ref"),
                    "columns": [tc.get("name") for tc in tx.findall("m:tableColumns/m:tableColumn", NS)],
                }
        out["sheets"].append(sheet)

    js = "/* AUTO-GENERATED by tools/extract.py dari file Excel METT-4 asli. JANGAN DIEDIT MANUAL. */\n"
    js += "window.METT_WB = " + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";\n"
    with open(dst, "w", encoding="utf-8") as fh:
        fh.write(js)
    # nilai cache Excel (hasil hitung terakhir Excel) -- dipakai untuk pengujian
    cache_path = os.path.join(os.path.dirname(dst), "..", "tests", "fixtures", "excel_cached_values.json")
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as fh:
        json.dump(cached_values, fh, ensure_ascii=False, indent=0, default=str)
    n_f = sum(1 for s in out["sheets"] for c in s["cells"].values() if "f" in c)
    print(f"OK: {len(out['sheets'])} sheet, {n_f} rumus, {sum(len(s['controls']) for s in out['sheets'])} kontrol, "
          f"{len(out['charts'])} grafik, {len(out['tables'])} tabel -> {dst}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
