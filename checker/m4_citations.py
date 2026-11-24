"""M4 - Sitasi dan Daftar Pustaka: format IEEE, jumlah referensi, kecocokan sitasi."""
import re
from datetime import date

from .models import CriterionResult


def _close(a, b, tol):
    return a is not None and abs(a - b) <= tol


def run(pd, crit, options=None):
    """Kembalikan CriterionResult kriteria 14 (M5 menambahkan temuan Crossref kemudian)."""
    prm = crit["params"]
    res = CriterionResult(crit["id"], crit["kode"], crit["nama"], "M4", crit["severity"],
                          group=crit.get("kelompok", ""))
    refs = pd.references
    head = next((h for h in pd.headings if h.section == "referensi" and h.level == 1), None)
    head_idx = head.idx if head else None

    if not refs:
        res.add("Daftar pustaka tidak ditemukan atau kosong.", para=head_idx,
                suggestion="Tambahkan bagian Referensi dengan entri bergaya IEEE: [1] A. Penulis, \u201cJudul,\u201d Jurnal, vol. x, no. x, pp. x\u2013x, 2025, doi: ...")
        res.note = "0 referensi"
        return res.finalize()

    n = len(refs)
    res.note = f"{n} referensi, {len({c for cit in pd.citations for c in cit['nums']})} dirujuk di teks"

    # 1. jumlah minimal
    if n < prm["min_references"]:
        res.add(f"Daftar pustaka berisi {n} referensi; minimal {prm['min_references']}.", para=head_idx,
                suggestion="Tambahkan referensi relevan dari jurnal bereputasi.")

    # 2. penomoran IEEE berurutan
    nums = [r["num"] for r in refs]
    for r in refs:
        if r["num"] is None:
            res.add("Entri referensi tidak bernomor format IEEE '[n]'.", para=r["idx"],
                    suggestion="Awali dengan [1], [2], ... sesuai urutan kemunculan di teks.")
    numbered = [x for x in nums if x is not None]
    if numbered and numbered != list(range(1, len(numbered) + 1)):
        res.add(f"Nomor daftar pustaka tidak berurutan 1..{len(numbered)}: {numbered[:12]}.", para=head_idx)

    # 3. kecocokan sitasi <-> daftar pustaka
    listed = {r["num"] for r in refs if r["num"] is not None}
    cited_order, bad_cites = [], {}
    for cit in pd.citations:
        for k in cit["nums"]:
            if k not in cited_order:
                cited_order.append(k)
            if k not in listed:
                bad_cites.setdefault(cit["idx"], set()).add(k)
    for idx, ks in sorted(bad_cites.items()):
        res.add("Sitasi " + ", ".join(f"[{k}]" for k in sorted(ks)) + " tidak ada di daftar pustaka.",
                para=idx, severity="error",
                suggestion="Perbaiki nomor sitasi atau tambahkan referensinya. Jika ini bukan sitasi "
                           "(mis. rentang nilai), tulis tanpa kurung siku, contoh: 'antara 0 dan 1'.")
    uncited = [r for r in refs if r["num"] is not None and r["num"] not in cited_order]
    for r in uncited[:5]:
        res.add(f"Referensi [{r['num']}] tidak pernah disitasi di dalam teks.", para=r["idx"],
                severity="warning", suggestion="Sitasi di teks atau hapus dari daftar pustaka.")
    if len(uncited) > 5:
        res.add(f"{len(uncited) - 5} referensi lain juga tidak disitasi.", severity="warning", comment=False)
    in_order = [k for k in cited_order if k in listed]
    if in_order and in_order != sorted(in_order):
        first_bad = next(k for a, k in zip(in_order, in_order[1:]) if k < a)
        res.add("Sitasi IEEE harus bernomor menurut urutan kemunculan pertama di teks "
                f"(pelanggaran pertama di sekitar [{first_bad}]).", severity="warning",
                para=pd.citations[0]["idx"] if pd.citations else None, comment=False)

    # 4. format entri IEEE
    for r in refs:
        miss = []
        if not r["has_quoted_title"]:
            miss.append("judul dalam tanda kutip")
        if r["year"] is None:
            miss.append("tahun")
        if r["journal_like"] and not r["italic"]:
            miss.append("nama jurnal dicetak miring")
        if miss:
            res.add(f"Entri [{r['num'] or '?'}] belum sesuai format IEEE: kurang " + ", ".join(miss) + ".",
                    para=r["idx"], severity="warning",
                    suggestion="Format: [n] A. Penulis, \u201cJudul,\u201d Nama Jurnal, vol., no., pp., tahun, doi.")

    # 5. ukuran huruf
    bad_size = [r for r in refs if not _close(r["size_pt"], prm["reference_size_pt"], prm["size_tolerance"])]
    for r in bad_size[:3]:
        res.add(f"Ukuran huruf entri [{r['num'] or '?'}] {r['size_pt']} pt; seharusnya {prm['reference_size_pt']} pt.",
                para=r["idx"], severity="warning")
    if len(bad_size) > 3:
        res.add(f"{len(bad_size) - 3} entri lain juga tidak berukuran {prm['reference_size_pt']} pt.",
                severity="warning", comment=False)

    # 6. komposisi: 80% jurnal, DOI jelas, terbitan <= 5 tahun
    journals = [r for r in refs if r["journal_like"]]
    share = len(journals) / n
    if share < prm["journal_share_min"]:
        res.add(f"Hanya {share:.0%} referensi terdeteksi sebagai artikel jurnal; target minimal "
                f"{prm['journal_share_min']:.0%}.", para=head_idx, severity="warning",
                suggestion="Perbanyak artikel jurnal bereputasi (vol., no., pp., dan DOI jelas).")
    no_doi = [r for r in journals if not r["doi"]]
    for r in no_doi[:3]:
        res.add(f"Artikel jurnal [{r['num'] or '?'}] belum mencantumkan DOI.", para=r["idx"], severity="warning",
                suggestion="Tambahkan 'doi: 10.xxxx/...'.")
    cutoff = date.today().year - prm["recent_years"]
    old = [r for r in journals if r["year"] and r["year"] < cutoff]
    for r in old[:3]:
        res.add(f"Artikel jurnal [{r['num'] or '?'}] terbit {r['year']}; batas terbitan {prm['recent_years']} tahun terakhir (\u2265 {cutoff}).",
                para=r["idx"], severity="warning")

    # 7. duplikasi
    seen = {}
    for r in refs:
        key = (r["doi"] or re.sub(r"\W+", "", r["text"].lower())[:80])
        if key in seen:
            res.add(f"Referensi [{r['num'] or '?'}] duplikat dengan [{seen[key]}].", para=r["idx"], severity="warning")
        seen.setdefault(key, r["num"] or "?")
    return res.finalize()
