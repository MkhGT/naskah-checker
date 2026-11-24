"""docx_edit.py - penyuntingan DOCX tingkat-run untuk pembangkit varian.

Semua penyuntingan teks memakai offset karakter pada teks paragraf (gabungan run, termasuk run di dalam
hyperlink), sehingga format di sekitar bagian yang diubah (miring, superskrip, ukuran) tetap terjaga.
"""
import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run


# Run dan teks
def inline_runs(par):
    for item in par.iter_inner_content():
        if isinstance(item, Run):
            yield item
        else:
            yield from item.runs


def run_spans(par):
    """[(run, mulai, akhir)] pada teks gabungan paragraf."""
    out, pos = [], 0
    for r in inline_runs(par):
        n = len(r.text)
        out.append((r, pos, pos + n))
        pos += n
    return out


def par_text(par):
    return "".join(r.text for r in inline_runs(par))


def _child_len(child):
    """Panjang teks yang disumbang satu anak run (memakai logika Run.text milik python-docx)."""
    tmp = OxmlElement("w:r")
    tmp.append(copy.deepcopy(child))
    return len(tmp.text)


def _new_run_like(run, text):
    r = OxmlElement("w:r")
    rpr = run._r.find(qn("w:rPr"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def split_at(par, k):
    """Pastikan ada batas run pada offset karakter k. Menangani run campuran (teks, tab, jeda, beberapa w:t).
    False bila k jatuh di tengah elemen non-teks yang tak terbagi."""
    for r, a, b in run_spans(par):
        if not (a < k < b):
            continue
        rel = k - a
        kids = [ch for ch in r._r if ch.tag != qn("w:rPr")]
        acc, left, right = 0, [], []
        for i, ch in enumerate(kids):
            n = _child_len(ch)
            if acc + n <= rel:
                left.append(ch)
            elif acc >= rel:
                right.append(ch)
            else:                                                   # k jatuh di dalam anak ini
                if ch.tag != qn("w:t"):
                    return False
                text = ch.text or ""
                lt, rt = copy.deepcopy(ch), copy.deepcopy(ch)
                lt.text, rt.text = text[:rel - acc], text[rel - acc:]
                for t in (lt, rt):
                    t.set(qn("xml:space"), "preserve")
                left.append(lt)
                right.append(rt)
            acc += n
        new = OxmlElement("w:r")
        rpr = r._r.find(qn("w:rPr"))
        if rpr is not None:
            new.append(copy.deepcopy(rpr))
        for ch in kids:
            r._r.remove(ch)
        for ch in left:
            r._r.append(ch)
        for ch in right:
            new.append(copy.deepcopy(ch))
        r._r.addnext(new)
        return True
    return True


def replace_range(par, start, end, new):
    """Ganti teks [start, end) dengan `new`; start == end berarti menyisipkan. Format sekitar dipertahankan."""
    if not (split_at(par, start) and split_at(par, end)):
        raise ValueError("run tidak dapat dipecah pada offset yang diminta")
    spans = run_spans(par)
    if start == end:
        if not new:
            return
        left = next((r for r, a, b in spans if b == start and b > a), None)
        if left is not None:
            left._r.addnext(_new_run_like(left, new))
            return
        right = next((r for r, a, b in spans if a == start and b > a), None)
        if right is not None:
            right._r.addprevious(_new_run_like(right, new))
            return
        par.add_run(new)
        return
    inside = [r for r, a, b in spans if a >= start and b <= end and b > a]
    if not inside:
        raise ValueError("rentang tidak memuat run")
    if new:
        inside[0].text = new
        rest = inside[1:]
    else:
        rest = inside
    for r in rest:
        r._r.getparent().remove(r._r)


def set_italic_range(par, start, end, value):
    if not (split_at(par, start) and split_at(par, end)):
        raise ValueError("run tidak dapat dipecah")
    for r, a, b in run_spans(par):
        if a >= start and b <= end and b > a:
            r.italic = value


# Paragraf / elemen
def delete_paragraph(par):
    """Hapus paragraf; bila satu-satunya paragraf di sel tabel, cukup dikosongkan (sel wajib punya paragraf)."""
    el = par._p
    parent = el.getparent()
    if parent.tag == qn("w:tc") and len(parent.findall(qn("w:p"))) == 1:
        for k in list(el):
            if k.tag != qn("w:pPr"):
                el.remove(k)
        return
    parent.remove(el)


def clone_after(par):
    new = copy.deepcopy(par._p)
    par._p.addnext(new)
    return Paragraph(new, par._parent)


def sentence_spans(text):
    return [(m.start(), m.end()) for m in re.finditer(r"[^.!?]+[.!?]?", text) if m.group(0).strip()]


# Kolom, field nomor halaman
def set_columns(sect_pr, n):
    cols = sect_pr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        anchor = next((a for a in (sect_pr.find(qn(t)) for t in
                                   ("w:pgNumType", "w:lnNumType", "w:pgBorders", "w:paperSrc", "w:pgMar", "w:pgSz"))
                       if a is not None), None)
        if anchor is not None:
            anchor.addnext(cols)
        else:
            sect_pr.append(cols)
    cols.set(qn("w:num"), str(n))
    cols.set(qn("w:space"), "708")


def add_page_field(par):
    run = par.add_run()
    for kind, text in (("begin", None), (None, " PAGE "), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def replace_page_fields(container_el, text="1"):
    """Ganti field PAGE (kompleks maupun fldSimple) dengan teks biasa. Mengembalikan jumlah field diganti."""
    count = 0
    for p in list(container_el.iter(qn("w:p"))):
        for f in list(p.findall(qn("w:fldSimple"))):
            if re.search(r"\bPAGE\b", f.get(qn("w:instr")) or ""):
                run = OxmlElement("w:r")
                t = OxmlElement("w:t"); t.text = text; run.append(t)
                f.addprevious(run); p.remove(f); count += 1
        runs = [c for c in p if c.tag == qn("w:r")]
        i = 0
        while i < len(runs):
            fc = runs[i].find(qn("w:fldChar"))
            if fc is not None and fc.get(qn("w:fldCharType")) == "begin":
                depth, j = 0, i
                while j < len(runs):
                    c = runs[j].find(qn("w:fldChar"))
                    if c is not None:
                        depth += 1 if c.get(qn("w:fldCharType")) == "begin" else 0
                        depth -= 1 if c.get(qn("w:fldCharType")) == "end" else 0
                    if depth == 0:
                        break
                    j += 1
                seg = runs[i:j + 1]
                instr = "".join(t.text or "" for r in seg for t in r.iter(qn("w:instrText")))
                if re.search(r"\bPAGE\b", instr):
                    new = OxmlElement("w:r")
                    rpr = seg[-1].find(qn("w:rPr")) if seg[-1].find(qn("w:rPr")) is not None else seg[0].find(qn("w:rPr"))
                    if rpr is not None:
                        new.append(copy.deepcopy(rpr))
                    t = OxmlElement("w:t"); t.text = text; new.append(t)
                    seg[0].addprevious(new)
                    for r in seg:
                        p.remove(r)
                    count += 1
                i = j + 1
            else:
                i += 1
    return count


def strip_non_field_text(container_el):
    """Buang semua teks yang BUKAN bagian dari field (mis. nomor halaman) pada seluruh paragraf di dalam
    `container_el`, termasuk paragraf di dalam tabel. Struktur (tabel, sel, paragraf, field) dipertahankan.
    Mengembalikan jumlah elemen yang dibuang."""
    removed = 0
    for p in container_el.iter(qn("w:p")):
        depth = 0
        for ch in list(p):
            if ch.tag == qn("w:pPr") or ch.tag == qn("w:fldSimple"):
                continue
            if ch.tag == qn("w:r"):
                fc = ch.find(qn("w:fldChar"))
                if fc is not None:
                    kind = fc.get(qn("w:fldCharType"))
                    depth += 1 if kind == "begin" else 0
                    depth -= 1 if kind == "end" else 0
                    continue
                if depth > 0:
                    continue
            p.remove(ch)
            removed += 1
    return removed


# Jumlah halaman (metadata docProps/app.xml) dan pengukuran dengan LibreOffice
def soffice_cmd(soffice=None):
    cand = soffice or os.environ.get("SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice")
    if not cand:
        return None
    return [sys.executable, str(cand)] if str(cand).endswith(".py") else [str(cand)]


def count_pdf_pages(pdf):
    try:
        from pypdf import PdfReader
        return len(PdfReader(str(pdf)).pages)
    except ImportError:
        pass
    if shutil.which("pdfinfo"):
        out = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout
        m = re.search(r"Pages:\s+(\d+)", out)
        if m:
            return int(m.group(1))
    found = re.findall(rb"/Type\s*/Page[^s]", Path(pdf).read_bytes())
    return len(found) or None


def measure_pages(docx_path, soffice=None, timeout=180):
    """Jumlah halaman hasil render LibreOffice (pendekatan; Word bisa berbeda sedikit). None bila tidak tersedia."""
    cmd = soffice_cmd(soffice)
    if not cmd:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "m.docx"
        shutil.copy(docx_path, src)
        try:
            subprocess.run(cmd + ["--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)],
                           capture_output=True, timeout=timeout, cwd=tmp)
        except (subprocess.TimeoutExpired, OSError):
            return None
        pdf = Path(tmp) / "m.pdf"
        return count_pdf_pages(pdf) if pdf.exists() else None


def read_pages_meta(docx_path):
    try:
        with zipfile.ZipFile(docx_path) as z:
            m = re.search(r"<Pages>(\d+)</Pages>", z.read("docProps/app.xml").decode("utf8", "ignore"))
            return int(m.group(1)) if m else None
    except Exception:
        return None


def patch_pages(docx_path, pages):
    """Tulis <Pages> pada docProps/app.xml (meniru yang dilakukan Word saat berkas disimpan)."""
    docx_path = str(docx_path)
    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/app.xml":
                xml = data.decode("utf8")
                if re.search(r"<Pages>\d+</Pages>", xml):
                    xml = re.sub(r"<Pages>\d+</Pages>", f"<Pages>{pages}</Pages>", xml)
                else:
                    xml = xml.replace("</Properties>", f"<Pages>{pages}</Pages></Properties>")
                data = xml.encode("utf8")
            zout.writestr(item, data)
    os.replace(tmp, docx_path)


def normalize_zip(docx_path, date_time=(2020, 1, 1, 0, 0, 0)):
    """Seragamkan cap waktu entri ZIP agar berkas identik byte-per-byte untuk isi yang sama (determinisme)."""
    docx_path = str(docx_path)
    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            info = zipfile.ZipInfo(item.filename, date_time=date_time)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = item.external_attr
            zout.writestr(info, zin.read(item.filename))
    os.replace(tmp, docx_path)
