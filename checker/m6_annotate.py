"""M6 - Anotasi Komentar: menempatkan temuan sebagai komentar Word pada paragraf terkait.

Dokumen asli tidak diubah; hasilnya disimpan sebagai salinan berkomentar.
"""
from collections import OrderedDict

from docx import Document

from .m2_parser import iter_paragraphs

AUTHOR, INITIALS = "Pemeriksa Naskah", "PN"
LABEL = {"error": "KESALAHAN", "warning": "PERINGATAN", "info": "SARAN"}


def annotate(src_path, dst_path, results, parsed):
    doc = Document(src_path)
    pmap = {idx: p for idx, p, _ in iter_paragraphs(doc)}

    # satukan temuan per paragraf; temuan tanpa lokasi digantung pada judul
    fallback = parsed.title.idx if parsed.title else next((p.idx for p in parsed.paras if p.text), None)
    grouped, general = OrderedDict(), []
    for r in results:
        for f in r.findings:
            if not f.comment:
                continue
            target = f.para_index if f.para_index in pmap else fallback
            (grouped.setdefault(target, [])).append((r, f, f.para_index is None))
    placed, skipped = 0, []
    for idx, items in grouped.items():
        par = pmap.get(idx)
        runs = par.runs if par is not None else []
        if not runs:
            skipped += [f for _, f, _ in items]
            continue
        items.sort(key=lambda t: ({"error": 0, "warning": 1, "info": 2}[t[1].severity], t[0].id))
        lines = []
        for r, f, is_general in items:
            tag = "[Umum] " if is_general else ""
            line = f"{tag}[{r.code}] {LABEL[f.severity]}: {f.message}"
            if f.suggestion:
                line += f" Saran: {f.suggestion}"
            lines.append(line)
        c = doc.add_comment(runs=[runs[0], runs[-1]], text=lines[0], author=AUTHOR, initials=INITIALS)
        for extra in lines[1:]:
            c.add_paragraph(extra)
        placed += 1
    doc.save(dst_path)
    return {"comments": placed, "unplaced": len(skipped)}
