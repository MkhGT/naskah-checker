"""M2 - Parser DOCX: membaca paragraf, heading, tabel, gambar, section, header, footer.

Semua paragraf (badan, sel tabel, isi sdt) diberi indeks berurutan oleh
`iter_paragraphs`. M6 memakai iterator yang sama untuk menempatkan komentar,
sehingga indeks pada temuan selalu menunjuk paragraf yang sama.
"""
import re
import zipfile
from dataclasses import dataclass, field
from typing import List, Optional

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

ALIGN = {0: "left", 1: "center", 2: "right", 3: "justify"}

SECTION_ALIASES = [
    ("pendahuluan", r"pendahuluan|introduction"),
    ("metodologi", r"(?:metodologi|metode)(?: penelitian)?|methods?|methodology"),
    ("hasil_pembahasan", r"hasil(?: dan pembahasan)?|pembahasan|results?(?: and discussion)?"),
    ("kesimpulan", r"kesimpulan|simpulan|conclusions?"),
    ("ucapan", r"ucapan terima ?kasih|acknowledg\w+"),
    ("referensi", r"referensi|daftar pustaka|references"),
]
CAPTION_RE = re.compile(r"^(Gambar|Gbr\.?|Figure|Fig\.?|Tabel|Table)\s*(\d+)", re.I)
CITE_RE = re.compile(r"\[(\d{1,3}(?:\s*[,\u2013\u2014-]\s*\d{1,3})*)\]")
# kata/simbol di depan kurung siku yang menandakan notasi matematika (rentang, himpunan), bukan sitasi
MATH_CUE_RE = re.compile(r"(?:rentang|interval|selang|range|domain|himpunan|bernilai)\b[^\[\]]{0,25}$|[\u2208=<>\u2264\u2265]\s*$", re.I)
DOI_RE = re.compile(r"10\.\d{4,9}/[^\s\"\u201d>]+", re.I)
SUP = "\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079"      # superskrip unicode 0-9
SUP_TO_DIGIT = str.maketrans(SUP, "0123456789")
DIGIT_TO_SUP = str.maketrans("0123456789", SUP)
MARKER_RE = re.compile(rf"[{SUP}]+(?:[\u02d2,\u201a][{SUP}]+)*")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")


# Struktur data
@dataclass
class ParaInfo:
    idx: int
    text: str
    style: str
    role: str = "other"
    level: int = 0                       # level heading
    font: Optional[str] = None
    size_pt: Optional[float] = None
    bold: bool = False
    italic: bool = False                 # ada run miring
    all_caps: bool = False
    alignment: str = "left"
    first_indent_cm: Optional[float] = None
    left_indent_cm: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    section_index: int = 0               # indeks section Word (0 = section pertama)
    line_spacing: object = 1.0          # float (kelipatan) atau str "N pt" (tepat/minimal)
    table_index: Optional[int] = None
    table_row: Optional[int] = None      # nomor baris (mulai 0) bila paragraf berada di tabel
    section: Optional[str] = None        # kunci bagian level-1 terdekat
    numbered: bool = False
    has_drawing: bool = False
    has_break: bool = False              # berisi jeda halaman/kolom
    fake_heading: bool = False
    words: int = 0
    raw_text: str = ""                   # teks asli gabungan run (offset sama dengan italic_ranges)
    italic_ranges: list = field(default_factory=list)   # [(mulai, akhir)] karakter bercetak miring


@dataclass
class ParsedDoc:
    page: dict = field(default_factory=dict)
    headers_footers: dict = field(default_factory=dict)   # wilayah -> {text, defined, page_field, content}
    hf_flags: dict = field(default_factory=dict)           # different_first, odd_even
    authors: Optional[ParaInfo] = None
    author_names: List[dict] = field(default_factory=list)      # {name, markers}
    affiliations: List[ParaInfo] = field(default_factory=list)
    affiliation_markers: List[int] = field(default_factory=list)
    affiliation_count: int = 0
    corresponding: Optional[dict] = None                   # {idx, text, email, inline, lines}
    paras: List[ParaInfo] = field(default_factory=list)
    headings: List[ParaInfo] = field(default_factory=list)
    title: Optional[ParaInfo] = None
    abstract_text: str = ""
    abstract_idx: Optional[int] = None
    abstract_idxs: List[int] = field(default_factory=list)       # semua paragraf isi abstrak
    keywords: List[str] = field(default_factory=list)
    keywords_idx: Optional[int] = None
    keywords_idxs: List[int] = field(default_factory=list)       # semua paragraf kata kunci
    tables: List[dict] = field(default_factory=list)
    figures: List[ParaInfo] = field(default_factory=list)
    captions: List[dict] = field(default_factory=list)
    references: List[dict] = field(default_factory=list)
    citations: List[dict] = field(default_factory=list)   # {idx, nums}
    pages_meta: Optional[int] = None
    total_words: int = 0
    # letak elemen identitas artikel: kunci -> [{label, idx, table_index}]; table_index None = bukan di tabel.
    # kunci: informasi (Informasi/Riwayat Artikel), abstrak, kata_kunci, korespondensi
    identity: dict = field(default_factory=dict)

    def para(self, idx):
        return self.paras[idx] if idx is not None and 0 <= idx < len(self.paras) else None

    def section_paras(self, key, roles=("body", "list")):
        return [p for p in self.paras if p.section == key and p.role in roles]


# Iterasi dokumen
def iter_blocks(doc):
    def walk(parent_el):
        for child in parent_el.iterchildren():
            if child.tag == qn("w:p"):
                yield "p", Paragraph(child, doc._body)
            elif child.tag == qn("w:tbl"):
                yield "tbl", Table(child, doc._body)
            elif child.tag == qn("w:sdt"):
                content = child.find(qn("w:sdtContent"))
                if content is not None:
                    yield from walk(content)
    yield from walk(doc.element.body)


def iter_paragraphs(doc):
    """Hasilkan (idx, Paragraph, table_index|None) berurutan. Dipakai M2 dan M6."""
    idx, tcount = 0, 0
    for kind, obj in iter_blocks(doc):
        if kind == "p":
            yield idx, obj, None
            idx += 1
        else:
            ti, tcount, seen = tcount, tcount + 1, []
            for row in obj.rows:
                for cell in row.cells:
                    if any(cell._tc is x for x in seen):
                        continue
                    seen.append(cell._tc)
                    for p in cell.paragraphs:
                        yield idx, p, ti
                        idx += 1


# Utilitas format efektif
def norm(text):
    return re.sub(r"\s+", " ", (text or "").replace("\u00a0", " ")).strip()


def _chain(style):
    while style is not None:
        yield style
        style = style.base_style


def _main_run(par):
    runs = [r for r in par.runs if r.text.strip()]
    return max(runs, key=lambda r: len(r.text)) if runs else None


def _font(par, attr, run=None):
    if run is not None:
        v = getattr(run.font, attr)
        if v is not None:
            return v
    for st in _chain(par.style):
        v = getattr(st.font, attr)
        if v is not None:
            return v
    return None


def _pf(par, attr):
    v = getattr(par.paragraph_format, attr)
    if v is not None:
        return v
    for st in _chain(par.style):
        v = getattr(st.paragraph_format, attr)
        if v is not None:
            return v
    return None


def _run_italic(par, run):
    """Miring efektif: run -> style karakter -> style paragraf."""
    if run.font.italic is not None:
        return bool(run.font.italic)
    for st in _chain(run.style):
        if st.font.italic is not None:
            return bool(st.font.italic)
    return bool(_font(par, "italic"))


def _text_and_italics(par):
    """Teks paragraf dari run (termasuk run di dalam hyperlink) dan rentang karakter yang miring."""
    parts, ranges, pos = [], [], 0
    for item in par.iter_inner_content():
        for r in ([item] if isinstance(item, Run) else item.runs):
            t = r.text
            if not t:
                continue
            if _run_italic(par, r):
                if ranges and ranges[-1][1] == pos:
                    ranges[-1] = (ranges[-1][0], pos + len(t))
                else:
                    ranges.append((pos, pos + len(t)))
            parts.append(t)
            pos += len(t)
    return "".join(parts), ranges


def _doc_default_size(doc):
    v = doc.styles.element.xpath("w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")
    return int(v[0]) / 2 if v else 10.0


def _doc_default_line(doc):
    """Spasi baris bawaan dokumen (docDefaults), mis. 278/240 = 1,16."""
    sp = doc.styles.element.xpath("w:docDefaults/w:pPrDefault/w:pPr/w:spacing")
    if sp and sp[0].get(qn("w:line")) and sp[0].get(qn("w:lineRule")) in (None, "auto"):
        return round(int(sp[0].get(qn("w:line"))) / 240, 2)
    return None


def _doc_default_spacing(doc):
    """(sebelum, sesudah) paragraf bawaan dokumen dalam pt; 0 bila tidak ada."""
    sp = doc.styles.element.xpath("w:docDefaults/w:pPrDefault/w:pPr/w:spacing")

    def tw(attr):
        v = sp[0].get(qn(attr)) if sp else None
        return round(int(v) / 20, 1) if v else 0.0
    return tw("w:before"), tw("w:after")


def _ind_attrs(ind):
    out = {}
    for k, v in ind.attrib.items():
        name = k.split("}")[-1]
        if name in ("left", "start"):
            out["left"] = int(v)
        elif name in ("hanging", "firstLine"):
            out[name] = int(v)
    return out


def _merge_ind(base, new):
    """Atribut indentasi digabung per atribut; hanging dan firstLine saling meniadakan."""
    m = dict(base)
    if "hanging" in new:
        m.pop("firstLine", None)
    if "firstLine" in new:
        m.pop("hanging", None)
    m.update(new)
    return m


def _num_ref(par):
    """(numId, ilvl) efektif: numPr langsung, lalu style; (None, '0') bila tidak bernomor."""
    num_id = ilvl = None
    sources = [par._p.find(qn("w:pPr") + "/" + qn("w:numPr"))]
    sources += [st.element.find(qn("w:pPr") + "/" + qn("w:numPr")) for st in _chain(par.style)]
    for np_ in sources:
        if np_ is None:
            continue
        n, l = np_.find(qn("w:numId")), np_.find(qn("w:ilvl"))
        if num_id is None and n is not None:
            num_id = n.get(qn("w:val"))
        if ilvl is None and l is not None:
            ilvl = l.get(qn("w:val"))
    return (None if num_id == "0" else num_id), (ilvl or "0")


def _lvl_ind(numbering, num_id, ilvl):
    try:
        a = numbering.xpath(f'w:num[@w:numId="{num_id}"]/w:abstractNumId/@w:val')
        if not a:
            return {}
        lv = numbering.xpath(f'w:abstractNum[@w:abstractNumId="{a[0]}"]/w:lvl[@w:ilvl="{ilvl}"]')
        ind = lv[0].find(qn("w:pPr") + "/" + qn("w:ind")) if lv else None
        return _ind_attrs(ind) if ind is not None else {}
    except Exception:
        return {}


def _effective_indent(par, numbering, cache):
    """(kiri_cm, baris_pertama_cm) efektif: definisi penomoran < style (turun-temurun) < format langsung.
    baris_pertama negatif = gantung (hanging). None bila tidak ditetapkan di mana pun."""
    merged = {}
    num_id, ilvl = _num_ref(par)
    if numbering is not None and num_id:
        if (num_id, ilvl) not in cache:
            cache[(num_id, ilvl)] = _lvl_ind(numbering, num_id, ilvl)
        merged = _merge_ind(merged, cache[(num_id, ilvl)])
    for st in reversed(list(_chain(par.style))):
        ind = st.element.find(qn("w:pPr") + "/" + qn("w:ind"))
        if ind is not None:
            merged = _merge_ind(merged, _ind_attrs(ind))
    ind = par._p.find(qn("w:pPr") + "/" + qn("w:ind"))
    if ind is not None:
        merged = _merge_ind(merged, _ind_attrs(ind))
    cm_ = lambda tw: round(tw * 2.54 / 1440, 2)
    left = cm_(merged["left"]) if "left" in merged else None
    if "firstLine" in merged:
        first = cm_(merged["firstLine"])
    elif "hanging" in merged:
        first = -cm_(merged["hanging"])
    else:
        first = None
    return left, first


def _is_numbered(par):
    direct = par._p.xpath("w:pPr/w:numPr/w:numId/@w:val")
    if direct:
        return direct[0] != "0"
    for st in _chain(par.style):
        v = st.element.xpath("w:pPr/w:numPr/w:numId/@w:val")
        if v:
            return v[0] != "0"
    return False


def canonical_section(title):
    t = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", title or "")
    t = re.sub(r"\((?:optional|opsional)\)", "", t, flags=re.I)
    t = norm(t).lower().rstrip(".:")
    for key, rx in SECTION_ALIASES:
        if re.fullmatch(rx, t):
            return key
    return None


def _is_citation_bracket(text, m):
    """Heuristik: bukan sitasi bila nilai 0 ([0,1]), indeks variabel (x[1]), atau didahului kata/simbol rentang."""
    start = m.start()
    if start > 0 and (text[start - 1].isalnum() or text[start - 1] == "_"):    # x[1], A[2]
        return False
    is_pair = bool(re.search(r"[,\u2013\u2014-]", m.group(1)))                  # notasi interval selalu [a,b]
    if is_pair and MATH_CUE_RE.search(text[max(0, start - 40):start]):          # rentang [1,5], x \u2208 [1,2]
        return False
    return not any(int(n) == 0 for n in re.findall(r"\d+", m.group(1)))        # IEEE mulai dari [1]


def extract_citations(text):
    """Nomor sitasi IEEE dari teks: [1], [1], [3], [2]-[5], [4]-[6] -> daftar int.
    Kurung siku bernotasi matematika/indeks diabaikan (lihat _is_citation_bracket)."""
    nums, last_end, last_num = [], None, None
    for m in CITE_RE.finditer(text):
        if not _is_citation_bracket(text, m):
            continue
        got = []
        for part in re.split(r"\s*,\s*", m.group(1)):
            rng = re.split(r"\s*[\u2013\u2014-]\s*", part)
            if len(rng) == 2 and rng[0].isdigit() and rng[1].isdigit() and int(rng[0]) <= int(rng[1]):
                got += list(range(int(rng[0]), int(rng[1]) + 1))
            elif part.isdigit():
                got.append(int(part))
        if got and last_end is not None and last_num is not None \
                and re.fullmatch(r"\s*[\u2013\u2014-]\s*", text[last_end:m.start()]) and got[0] > last_num:
            nums += list(range(last_num + 1, got[0]))    # [5]-[8] -> 6,7
        nums += got
        last_end, last_num = m.end(), (got[-1] if got else None)
    return nums


# Halaman, header/footer
def _sec_dict(s):
    cols = s._sectPr.find(qn("w:cols"))
    num = cols.get(qn("w:num")) if cols is not None else None
    w, h = round(s.page_width.cm, 1), round(s.page_height.cm, 1)
    paper = "A4" if (w, h) == (21.0, 29.7) else ("Letter" if (w, h) == (21.6, 27.9) else f"{w}x{h} cm")
    return {
        "paper": paper, "width_cm": w, "height_cm": h,
        "orientation": "landscape" if s.page_width > s.page_height else "portrait",
        "top_cm": round(s.top_margin.cm, 2), "bottom_cm": round(s.bottom_margin.cm, 2),
        "left_cm": round(s.left_margin.cm, 2), "right_cm": round(s.right_margin.cm, 2),
        "columns": int(num) if num else 1,
        "start_type": str(s.start_type).split(".")[-1].split(" ")[0] if s.start_type is not None else None,
    }


def _page_info(doc):
    secs = [_sec_dict(s) for s in doc.sections]
    return {**secs[0], "sections": len(secs), "sections_info": secs}


def _el_text(root):
    mc_fallback = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
    out = []
    for el in root.iter(qn("w:t")):
        if any(a.tag == mc_fallback for a in el.iterancestors()):
            continue
        out.append(el.text or "")
    return norm("".join(out))


def _part_text(part):
    return _el_text(part._element)


def _hf_layout(root):
    """Letak isi header/footer terhadap tabel: (punya_tabel, ada_isi_di_luar_tabel).
    'Isi' = teks atau field nomor halaman pada paragraf tingkat-atas (bukan di dalam w:tbl). Paragraf kosong
    penutup yang selalu ada setelah tabel (syarat Word) tidak dihitung."""
    state = {"table": False, "outside": False}

    def walk(parent):
        for c in parent.iterchildren():
            if c.tag == qn("w:tbl"):
                state["table"] = True
            elif c.tag == qn("w:p"):
                if _el_text(c) or _has_page_field(c):
                    state["outside"] = True
            elif c.tag == qn("w:sdt"):
                content = c.find(qn("w:sdtContent"))
                if content is not None:
                    walk(content)
    walk(root)
    return state["table"], state["outside"]


def _has_page_field(el):
    for t in el.iter(qn("w:instrText")):
        if re.search(r"\bPAGE\b", t.text or ""):
            return True
    return any(re.search(r"\bPAGE\b", f.get(qn("w:instr")) or "") for f in el.iter(qn("w:fldSimple")))


def _headers_footers(doc):
    """Enam wilayah pada section pertama: header/footer halaman pertama, ganjil (default), genap."""
    s = doc.sections[0]
    parts = {"header_first": s.first_page_header, "header_odd": s.header, "header_even": s.even_page_header,
             "footer_first": s.first_page_footer, "footer_odd": s.footer, "footer_even": s.even_page_footer}
    out = {}
    for key, part in parts.items():
        defined = not part.is_linked_to_previous     # baca sebelum menyentuh isi
        text = _part_text(part) if defined else ""
        has_table, outside = _hf_layout(part._element) if defined else (False, False)
        out[key] = {"text": text, "defined": defined,
                    "page_field": _has_page_field(part._element) if defined else False,
                    "content": re.sub(r"[\s\d]+$", "", text),     # teks di luar nomor halaman
                    "has_table": has_table,                          # ada tabel pada bagian ini
                    "in_table": has_table and not outside}           # seluruh isi berada di dalam tabel
    flags = {"different_first": bool(s.different_first_page_header_footer),
             "odd_even": bool(doc.settings.odd_and_even_pages_header_footer)}
    return out, flags


def _pages_meta(path):
    try:
        with zipfile.ZipFile(path) as z:
            m = re.search(r"<Pages>(\d+)</Pages>", z.read("docProps/app.xml").decode("utf8", "ignore"))
            return int(m.group(1)) if m else None
    except Exception:
        return None


# Parser utama
def parse(path):
    doc = Document(path)
    hf, hf_flags = _headers_footers(doc)
    pd = ParsedDoc(page=_page_info(doc), headers_footers=hf, hf_flags=hf_flags,
                   pages_meta=_pages_meta(path))
    default_size = _doc_default_size(doc)
    default_line = _doc_default_line(doc) or 1.0
    default_before, default_after = _doc_default_spacing(doc)
    try:
        numbering = doc.part.numbering_part.element
    except Exception:
        numbering = None
    ind_cache, sec_i = {}, 0

    raw = list(iter_paragraphs(doc))
    table_ranges = {}
    first_heading_idx = None

    # ---- lintasan 1: deskripsi tiap paragraf, deteksi heading -------------
    for idx, p, ti in raw:
        text = norm(p.text)
        run = _main_run(p)
        size = _font(p, "size", run)
        al = p.alignment if p.alignment is not None else _pf(p, "alignment")
        ls = _pf(p, "line_spacing")
        sb, sa = _pf(p, "space_before"), _pf(p, "space_after")
        ind = _effective_indent(p, numbering, ind_cache)
        info = ParaInfo(
            idx=idx, text=text, style=p.style.name,
            font=_font(p, "name", run),
            size_pt=round(size.pt, 1) if size is not None else default_size,
            bold=bool(_font(p, "bold", run)), all_caps=bool(_font(p, "all_caps", run)),
            italic=any(r.italic or (r.italic is None and _font(p, "italic", r)) for r in p.runs if r.text.strip()),
            alignment=ALIGN.get(int(al), "left") if al is not None else "left",
            left_indent_cm=ind[0], first_indent_cm=ind[1],
            space_before_pt=round(sb.pt, 1) if sb is not None else default_before,
            space_after_pt=round(sa.pt, 1) if sa is not None else default_after,
            section_index=sec_i,
            line_spacing=(default_line if ls is None else round(ls, 2) if isinstance(ls, float)
                          else f"{round(ls.pt, 1)} pt"),
            table_index=ti, numbered=_is_numbered(p),
            has_drawing=bool(p._p.xpath(".//w:drawing")), words=len(text.split()),
            has_break=bool(p._p.xpath(".//w:br[@w:type='page' or @w:type='column']")),
        )
        if ti is not None:
            tr = p._p.xpath("ancestor::w:tr[1]")
            info.table_row = len(tr[0].xpath("preceding-sibling::w:tr")) if tr else None
        if ti is None and text:
            info.raw_text, info.italic_ranges = _text_and_italics(p)
        if ti is not None:
            r = table_ranges.setdefault(ti, [idx, idx])
            r[1] = idx
        # heading
        lvl_m = re.match(r"Heading\s*(\d)", p.style.name, re.I)
        outline = p._p.xpath("w:pPr/w:outlineLvl/@w:val")
        if ti is None and text and (lvl_m or outline):
            info.role, info.level = "heading", int(lvl_m.group(1)) if lvl_m else int(outline[0]) + 1
        elif ti is None and text and len(text) <= 60 and canonical_section(text) and info.bold \
                and p.style.name != "Title":
            info.role, info.level, info.fake_heading = "heading", 1, True
        pd.paras.append(info)
        if p._p.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:      # paragraf ini mengakhiri sebuah section
            sec_i += 1

    # ---- lintasan 2: konteks bagian, peran paragraf -----------------------
    current, num_counter = None, [0, 0, 0]
    for info in pd.paras:
        if info.role == "heading":
            if first_heading_idx is None:
                first_heading_idx = info.idx
            if info.level == 1:
                current = canonical_section(info.text) or "lainnya"
            pd.headings.append(info)
        info.section = current if first_heading_idx is not None and info.idx >= first_heading_idx else None

    tables_front = {ti for ti, (a, _) in table_ranges.items()
                    if first_heading_idx is None or a < first_heading_idx}
    for ti, (a, b) in sorted(table_ranges.items()):
        pd.tables.append({"index": ti, "first_idx": a, "last_idx": b, "front": ti in tables_front})
    # perataan & ukuran tabel dari objek Table
    t_objs = [o for k, o in iter_blocks(doc) if k == "tbl"]
    for t in pd.tables:
        if t["index"] < len(t_objs):
            tb = t_objs[t["index"]]
            t["alignment"] = tb.alignment.name.lower() if tb.alignment is not None else "left"
            t["rows"], t["cols"] = len(tb.rows), len(tb.columns)

    for info in pd.paras:
        if info.role == "heading" or not info.text and not info.has_drawing:
            continue
        if info.table_index is not None:
            info.role = "table"
            continue
        if info.has_drawing:
            info.role = "figure"
            pd.figures.append(info)
            continue
        if info.style == "Title":
            info.role = "title"
        elif info.style == "Caption" or CAPTION_RE.match(info.text):
            info.role = "caption"
        elif info.section == "referensi":
            info.role = "reference"
        elif info.section is None:
            info.role = "front"
        elif info.numbered:
            info.role = "list"
        else:
            info.role = "body"

    # ---- judul -------------------------------------------------------------
    pd.title = next((p for p in pd.paras if p.role == "title"), None)
    if pd.title is None:
        pd.title = next((p for p in pd.paras if p.role == "front" and p.text
                         and ((p.size_pt or 0) >= 13 or p.alignment == "center")), None)
        if pd.title:
            pd.title.role = "title"

    # ---- caption -----------------------------------------------------------
    for p in pd.paras:
        if p.role == "caption":
            m = CAPTION_RE.match(p.text)
            if m:
                kind = "tabel" if m.group(1).lower().startswith(("tabel", "table")) else "gambar"
                pd.captions.append({"kind": kind, "number": int(m.group(2)), "idx": p.idx,
                                    "alignment": p.alignment, "size_pt": p.size_pt, "text": p.text,
                                    "well_formed": bool(re.match(r"^(Gambar|Tabel)\s+\d+\.\s+\S", p.text))})

    _extract_front(doc, pd, raw, table_ranges)
    _extract_authors(pd, raw)
    _extract_identity(pd)
    _extract_references(pd, doc, raw)

    # ---- sitasi dalam teks ---------------------------------------------------
    front_idx = {p.idx for p in pd.paras if p.table_index in tables_front}
    for p in pd.paras:
        if p.role in ("body", "list", "table") and p.idx not in front_idx:
            nums = extract_citations(p.text)
            if nums:
                pd.citations.append({"idx": p.idx, "nums": nums})

    pd.total_words = sum(p.words for p in pd.paras if p.role not in ("reference",))
    return pd


# Abstrak & kata kunci (kotak tabel atau paragraf biasa)
def _split_keywords(chunks):
    out = []
    for c in chunks:
        out += [k.strip(" .") for k in re.split(r"[;,]", c) if k.strip(" .")]
    return out


def _keywords_from_table(pd, raw, tb):
    seen = []
    for row in tb.rows:
        for c in row.cells:
            if any(c._tc is x for x in seen):
                continue
            seen.append(c._tc)
            paras = [p for p in c.paragraphs if norm(p.text)]
            lab = re.match(r"(?i)^(kata kunci|keywords?)\s*:?\s*(.*)$", norm(paras[0].text)) if paras else None
            if lab:
                chunks = ([lab.group(2)] if lab.group(2) else []) + [norm(p.text) for p in paras[1:]]
                pd.keywords = _split_keywords(chunks)
                pd.keywords_idx = _find_idx(raw, paras[0])
                pd.keywords_idxs = [i for i in (_find_idx(raw, q) for q in paras) if i is not None]


def _extract_front(doc, pd, raw, table_ranges):
    # (a) pola tabel: sel header "ABSTRAK" + sel "Kata Kunci:"
    abs_col = None
    for kind, tb in iter_blocks(doc):
        if kind != "tbl":
            continue
        hdr = [norm(c.text).lower() for c in tb.rows[0].cells]
        for ci, h in enumerate(hdr):
            if h in ("abstrak", "abstract"):
                abs_col = ci
        if abs_col is None:
            continue
        best, seen = ("", None), []
        for row in tb.rows[1:]:
            if abs_col < len(row.cells):
                c = row.cells[abs_col]
                if any(c._tc is x for x in seen):
                    continue
                seen.append(c._tc)
                if len(norm(c.text)) > len(best[0]):
                    best = (norm(c.text), c)
        if best[0]:
            pd.abstract_text = best[0]
            first = next((p for p in best[1].paragraphs if norm(p.text)), None)
            pd.abstract_idx = _find_idx(raw, first)
            pd.abstract_idxs = [i for i in (_find_idx(raw, q) for q in best[1].paragraphs if norm(q.text))
                                if i is not None]
        _keywords_from_table(pd, raw, tb)          # kata kunci: sel yang diawali "Kata Kunci"
        break

    # kata kunci di tabel lain (mis. kotak identitas terpecah menjadi dua tabel)
    if not pd.keywords:
        for kind, tb in iter_blocks(doc):
            if kind == "tbl":
                _keywords_from_table(pd, raw, tb)
                if pd.keywords:
                    break

    # (b) pola paragraf biasa
    if not pd.abstract_text:
        for i, p in enumerate(pd.paras):
            if p.role in ("front", "other", "body") and re.match(r"(?i)^(abstrak|abstract)\b[\s:.\u2014-]*", p.text):
                rest = re.sub(r"(?i)^(abstrak|abstract)\b[\s:.\u2014-]*", "", p.text)
                chunks = [rest] if rest else []
                idxs = [p.idx] if rest else []
                for q in pd.paras[i + 1:]:
                    if q.role == "heading" or re.match(r"(?i)^(kata kunci|keywords?)\b", q.text):
                        break
                    if q.text:
                        chunks.append(q.text)
                        idxs.append(q.idx)
                pd.abstract_text, pd.abstract_idx = norm(" ".join(chunks)), p.idx
                pd.abstract_idxs = idxs
                break
    if not pd.keywords:
        for i, p in enumerate(pd.paras):
            if p.table_index is not None:
                continue
            m = re.match(r"(?i)^(kata kunci|keywords?)\s*[:\u2014-]\s*(.*)$", p.text)
            if m and m.group(2).strip():                       # "Kata kunci: a; b; c" dalam satu paragraf
                pd.keywords, pd.keywords_idx = _split_keywords([m.group(2)]), p.idx
                pd.keywords_idxs = [p.idx]
                break
            if m and p.section is None:                        # label sendiri, kata kunci di paragraf-paragraf berikutnya
                chunks, idxs = [], [p.idx]                     # (susunan template, tetapi diketik di luar tabel)
                for q in pd.paras[i + 1:i + 11]:
                    if q.role == "heading" or q.table_index is not None or re.search(r":\s*$", q.text) \
                            or re.match(r"(?i)^(penulis korespondensi|corresponding author|abstrak|abstract|"
                                        r"informasi artikel|riwayat artikel)\b", q.text):
                        break
                    if q.text:
                        chunks.append(q.text)
                        idxs.append(q.idx)
                if chunks:
                    pd.keywords, pd.keywords_idx = _split_keywords(chunks), p.idx
                    pd.keywords_idxs = idxs
                    break


def _find_idx(raw, par):
    if par is None:
        return None
    for idx, p, _ in raw:
        if p._p is par._p:
            return idx
    return None


# Letak elemen identitas artikel (di tabel atau paragraf biasa)
INFO_LABEL_RE = re.compile(r"(?i)^(informasi artikel|article information|article info)\b")
HISTORY_LABEL_RE = re.compile(r"(?i)^(riwayat artikel|article history)\b")


def _extract_identity(pd):
    """Catat paragraf dan indeks tabel tiap elemen identitas agar M3 dapat memeriksa apakah elemen itu
    berada di dalam tabel (template: Informasi/Riwayat Artikel, Abstrak, Kata Kunci, Penulis Korespondensi)."""
    ident = {"informasi": [], "abstrak": [], "kata_kunci": [], "korespondensi": []}

    def put(key, label, idx):
        par = pd.para(idx)
        if par is not None:
            ident[key].append({"label": label, "idx": idx, "table_index": par.table_index})

    for label, rx in (("Informasi Artikel", INFO_LABEL_RE), ("Riwayat Artikel", HISTORY_LABEL_RE)):
        hit = next((p for p in pd.paras if p.section is None and p.text and rx.match(p.text)), None)
        if hit is not None:
            put("informasi", label, hit.idx)
    put("abstrak", "Abstrak", pd.abstract_idx)
    put("kata_kunci", "Kata Kunci", pd.keywords_idx)
    if pd.corresponding:
        put("korespondensi", "Penulis Korespondensi", pd.corresponding["idx"])
    pd.identity = ident


# Daftar pustaka
def _extract_references(pd, doc, raw):
    para_map = {idx: p for idx, p, _ in raw}
    auto = 0
    for p in pd.paras:
        if p.role != "reference" or not p.text:
            continue
        m = re.match(r"^\[(\d+)\]\s*(.*)$", p.text)
        if m:
            num, body = int(m.group(1)), m.group(2)
        elif p.numbered:
            auto += 1
            num, body = auto, p.text
        else:
            num, body = None, p.text
        doi_m = DOI_RE.search(body)
        doi = doi_m.group(0).rstrip(".,;)") if doi_m else None
        no_doi = DOI_RE.sub("", body)
        years = [int(y) for y in re.findall(r"\b((?:19|20)\d{2})\b", no_doi)]
        tm = re.search(r"[\u201c\"](.+?)[\u201d\"]", body)
        title = re.sub(r"[,.]\s*$", "", tm.group(1)).strip() if tm else None
        pd.references.append({
            "idx": p.idx, "num": num, "text": body, "doi": doi,
            "year": years[-1] if years else None, "title": title,
            "has_quoted_title": bool(tm), "italic": p.italic, "size_pt": p.size_pt,
            "journal_like": bool(re.search(r"\bvol\.", body, re.I) and re.search(r"\b(no\.|pp\.)", body, re.I)),
            "numbered": num is not None,
        })


# Penulis, afiliasi, penulis korespondensi
def _with_superscripts(par):
    """Teks paragraf; angka pada run superskrip diubah ke superskrip unicode agar seragam."""
    out = []
    for r in par.runs:
        out.append(r.text.translate(DIGIT_TO_SUP) if r.font.superscript else r.text)
    return norm("".join(out))


def _marker_sets(text):
    """'Nama\u00b9, Nama\u00b2\u02d2\u00b3' -> himpunan angka per kelompok penanda."""
    groups = []
    for m in MARKER_RE.finditer(text):
        nums = {int(re.sub(r"\D", "", part.translate(SUP_TO_DIGIT)))
                for part in re.split(r"[\u02d2,\u201a]", m.group(0)) if part}
        groups.append(nums)
    return groups


def _extract_authors(pd, raw):
    para_map = {idx: p for idx, p, _ in raw}
    stop = {pd.abstract_idx, pd.keywords_idx}
    start = pd.title.idx if pd.title else -1
    block = []
    for p in pd.paras:
        if p.role != "front" or not p.text or p.idx <= start:
            continue
        if p.idx in stop or re.match(r"(?i)^(abstrak|abstract|kata kunci|keywords?)\b", p.text):
            break
        if re.match(r"(?i)^(penulis korespondensi|corresponding author)\b", p.text):
            continue
        block.append(p)
    if block:
        pd.authors, pd.affiliations = block[0], block[1:]
        text = _with_superscripts(para_map[pd.authors.idx])
        text = re.sub(r"\(\s*\d+\s*pt\s*\)", "", text)                  # petunjuk "(10 pt)" bawaan template
        for seg in re.split(r"\s*(?:,\s+|;|\s&\s|\s(?:and|dan)\s)\s*", text):
            name = norm(MARKER_RE.sub("", seg)).strip(" *,")
            if name:
                markers = sorted(set().union(*_marker_sets(seg)) if _marker_sets(seg) else [])
                pd.author_names.append({"name": name, "markers": markers})
        aff_text = " ".join(_with_superscripts(para_map[p.idx]) for p in pd.affiliations)
        groups = _marker_sets(aff_text)
        pd.affiliation_markers = sorted(set().union(*groups)) if groups else []
        pd.affiliation_count = len(groups) if groups else len(pd.affiliations)

    # penulis korespondensi: paragraf berlabel + (isi pada paragraf yang sama atau berikutnya)
    label_re = re.compile(r"(?i)^(penulis korespondensi|corresponding author|informasi artikel|riwayat artikel|"
                          r"kata kunci|keywords?|abstrak|abstract)\b")
    for i, p in enumerate(pd.paras):
        m = re.match(r"(?i)^(penulis korespondensi|corresponding author)\s*:?\s*(.*)$", p.text)
        if m and p.text:
            chunks = [m.group(2)] if m.group(2) else []
            lines = []
            for q in pd.paras[i + 1:i + (4 if m.group(2) else 6)]:
                if q.table_index != p.table_index or q.role == "heading" or re.search(r":\s*$", q.text) \
                        or label_re.match(q.text):
                    break
                if q.text:
                    chunks.append(q.text)
                    lines.append({"idx": q.idx, "text": q.text})
            content = norm(" ".join(chunks))
            em = EMAIL_RE.search(content)
            pd.corresponding = {"idx": p.idx, "text": content, "email": em.group(0) if em else None,
                                "inline": bool(m.group(2)), "lines": lines}
            break
