"""Baca isi .pptx sebagai markdown TANPA python-pptx (hanya pustaka standar Python).

Pengganti ringan slide_text.py bila python-pptx belum terpasang. Per slide: judul, teks (urut atas ke bawah),
tabel, kategori/nilai grafik, alt text gambar, dan catatan pembicara, di bawah penanda `<!-- Slide number: N -->`.

Pemakaian: python3 scripts/read_pptx.py presentasi.pptx
"""

import posixpath
import re
import sys
import zipfile

try:  # defusedxml bila ada (berkas dari pihak luar); bila tidak, ElementTree standar (tidak memuat entitas luar)
    from defusedxml import ElementTree as ET
except ImportError:  # pragma: no cover
    import xml.etree.ElementTree as ET

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = "{%s}id" % NS["r"]
MAX_PART = 50 * 1024 * 1024  # tolak bagian XML raksasa (zip bomb)


def read(z, name):
    info = z.getinfo(name)
    if info.file_size > MAX_PART:
        raise ValueError(f"{name} terlalu besar ({info.file_size} byte)")
    return z.read(name)


def rels(z, part):
    d, f = posixpath.split(part)
    name = posixpath.join(d, "_rels", f + ".rels")
    if name not in z.namelist():
        return {}
    out = {}
    for r in ET.fromstring(read(z, name)).findall("rel:Relationship", NS):
        if r.get("TargetMode") == "External":
            continue
        t = r.get("Target", "")
        # Target absolut ("/ppt/charts/chart1.xml", gaya pptxgenjs) relatif ke akar paket; selain itu relatif ke folder bagian.
        full = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
        out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], full)
    return out


def one_line(s):
    return " ".join(str(s).split())


def para_text(p):
    return "".join(t.text or "" for t in p.iter("{%s}t" % NS["a"]))


def shape_pos(el):
    off = el.find(".//a:off", NS)
    return (int(off.get("y", 0)), int(off.get("x", 0))) if off is not None else (0, 0)


def chart_lines(z, part):
    root = ET.fromstring(read(z, part))
    out = ["### Grafik"]
    for ser in root.iter("{%s}ser" % NS["c"]):
        name = one_line(" ".join(v.text or "" for v in ser.findall("c:tx//c:v", NS))) or "seri"
        cats = [v.text or "" for v in ser.findall("c:cat//c:pt/c:v", NS)]
        vals = [v.text or "" for v in ser.findall("c:val//c:pt/c:v", NS)]
        out.append(f"- {name}: " + ", ".join(f"{c}={v}" for c, v in zip(cats, vals)) if cats else f"- {name}: " + ", ".join(vals))
    return out


def slide_lines(z, part):
    root = ET.fromstring(read(z, part))
    rel = rels(z, part)
    items = []
    tree = root.find("p:cSld/p:spTree", NS)
    for el in tree.iter():
        tag = el.tag.split("}")[-1]
        if tag == "sp":
            ph = el.find("p:nvSpPr/p:nvPr/p:ph", NS)
            paras = [one_line(para_text(p)) for p in el.findall("p:txBody/a:p", NS)]
            paras = [p for p in paras if p]
            if not paras or (ph is not None and ph.get("type") in ("sldNum", "dt", "ftr")):
                continue
            if ph is not None and ph.get("type") in ("title", "ctrTitle"):
                items.append((shape_pos(el), ["# " + " ".join(paras)]))
            else:
                items.append((shape_pos(el), paras))
        elif tag == "graphicFrame":
            tbl = el.find(".//a:tbl", NS)
            if tbl is not None:
                rows = [[one_line(" ".join(para_text(p) for p in tc.findall(".//a:p", NS))).replace("|", "\\|") for tc in tr.findall("a:tc", NS)]
                        for tr in tbl.findall("a:tr", NS)]
                lines = ["| " + " | ".join(r) + " |" for r in rows]
                if lines:
                    lines.insert(1, "|" + " --- |" * len(rows[0]))
                items.append((shape_pos(el), lines))
            chart = el.find(".//c:chart", NS)
            if chart is not None and chart.get(R_ID) in rel:
                items.append((shape_pos(el), chart_lines(z, rel[chart.get(R_ID)][1])))
        elif tag == "pic":
            alt = el.find("p:nvPicPr/p:cNvPr", NS)
            if alt is not None and alt.get("descr"):
                items.append((shape_pos(el), [f"![{one_line(alt.get('descr'))}]()"]))
    items.sort(key=lambda it: it[0])
    # Kotak satu baris yang sejajar (y sama, mis. nomor · teks · keterangan agenda) digabung menjadi satu baris.
    out, last_y = [], None
    for (y, _x), lines in items:
        if len(lines) == 1 and last_y is not None and abs(y - last_y) < 20000 and out and not out[-1].startswith(("#", "|")):
            out[-1] += " · " + lines[0]
        else:
            out += lines
        last_y = y if len(lines) == 1 else None
    for typ, target in rel.values():
        if typ == "notesSlide" and target in z.namelist():
            notes = ET.fromstring(read(z, target))
            for sp in notes.iter("{%s}sp" % NS["p"]):
                ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
                if ph is not None and ph.get("type") == "body":
                    txt = " ".join(one_line(para_text(p)) for p in sp.findall("p:txBody/a:p", NS)).strip()
                    if txt:
                        out += ["", "### Catatan pembicara", txt]
    return out


def main(path):
    with zipfile.ZipFile(path) as z:
        pres = ET.fromstring(read(z, "ppt/presentation.xml"))
        prel = rels(z, "ppt/presentation.xml")
        for n, sid in enumerate(pres.findall("p:sldIdLst/p:sldId", NS), 1):
            part = prel[sid.get(R_ID)][1]
            print(f"<!-- Slide number: {n} -->")
            print("\n".join(slide_lines(z, part)))
            print()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
