"""
Trax CDC Export Report extractor
--------------------------------
Reads a scanned export-shipment bundle (SAD, Commercial Invoice, Packing List, ...)
and appends ONE row to the CDC Export Report Excel (sheet "Customs sheet 2026").

Fields = the 13 red-marked items on the scan:
 1 INV.NO  2 Customer  3 Customs Decl. no  4 Date  5 HTS code  6 Description
 7 Qty  8 Total price  9 Destination  10 Rate KHR  11 NW  12 GW  13 Kind of fabric
Column T (Amount/KHR) stays the template formula =P*S.

Every value is read from several documents and majority-voted; mismatches are reported.
Usage:  python extractor.py shipment.pdf --xlsx CDC_Export_Report_2026.xlsx
"""
import os
# One thread per tesseract process: running several processes AND their internal OpenMP threads
# oversubscribes the CPU and makes OCR much slower. Must be set before tesseract starts.
os.environ.setdefault("OMP_THREAD_LIMIT", "1")

import argparse, re, shutil, subprocess, sys, tempfile, threading
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import numpy as np
import openpyxl
import pytesseract
from PIL import Image

# ============================ CONFIG (edit here) ============================
SHEET_NAME = None            # None = first sheet ("Customs sheet 2026")
FIRST_DATA_ROW = 6
OCR_DPI = 250                # 200 = ~35% faster, slightly less accurate on small print
OCR_WORKERS = max(1, min(8, (os.cpu_count() or 2)))   # pages OCR'd in parallel
COLUMNS = {                  # field -> Excel column
    "inv_no": "C", "customer": "D", "customs_no": "F", "customs_date": "G",
    "hts": "I", "description": "M", "qty": "N", "amount": "P",
    "destination": "R", "rate": "S", "nw": "U", "gw": "V", "fabric": "X",
}
EXTRA_COLUMNS = {"onboard_date": "E", "procedure": "J"}   # only with --extras
FORMULA_COLUMN = ("T", "=P{r}*S{r}")                       # Amount/KHR
CUSTOMER_MAP = {"ADIDAS": "ADIDAS", "PUMA": "PUMA", "NIKE": "NIKE",
                "UNDER ARMOUR": "UNDER ARMOUR", "NEW BALANCE": "NEW BALANCE"}
COUNTRY_MAP = {"UK": "United Kingdom", "U.K.": "United Kingdom", "GB": "United Kingdom",
               "US": "United States", "USA": "United States", "DE": "Germany",
               "FR": "France", "NL": "Netherlands", "JP": "Japan", "CN": "China",
               "HK": "Hong Kong", "TW": "Taiwan", "AU": "Australia", "CA": "Canada", "KR": "South Korea"}
DUP_E_POLICY = "group"       # "group": same E no. only allowed directly under its own group (several invoices, one declaration)
                             # "skip" : refuse any E no. already in the sheet   |  "allow": no E-number check
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
ERROR_LOG = APP_DIR / "skipped_files.csv"   # every skipped PDF is listed here
ERROR_DIR = APP_DIR / "error_files"         # ...and a COPY of it is put here (None = no copy)
ERROR_RESULT_LOG = APP_DIR / "skipped_results.csv"   # what was extracted from every skipped PDF (to check by eye)
# ===========================================================================

TESSERACT_PATHS = [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                   r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                   str(Path.home() / "AppData/Local/Programs/Tesseract-OCR/tesseract.exe"),
                   str(Path.home() / "AppData/Local/Tesseract-OCR/tesseract.exe")]


def setup_tesseract():
    """Find tesseract: TESSERACT_CMD env var, PATH, then the usual Windows folders."""
    import os
    cmd = os.environ.get("TESSERACT_CMD")
    if cmd and Path(cmd).exists():
        pytesseract.pytesseract.tesseract_cmd = cmd
        return
    if shutil.which("tesseract"):
        return
    for p in TESSERACT_PATHS:
        if Path(p).exists():
            pytesseract.pytesseract.tesseract_cmd = p
            return
    raise RuntimeError("Tesseract OCR not found. Install the UB-Mannheim Windows build, or set "
                       "TESSERACT_CMD to the full path of tesseract.exe, then restart.")


# ----------------------------- PDF -> images -----------------------------
def render_pages(pdf, dpi=OCR_DPI):
    """Yield PIL images, one per page (rotation applied)."""
    try:
        import fitz  # PyMuPDF (pip install pymupdf) - no poppler needed
        with fitz.open(str(pdf)) as doc:
            for page in doc:
                pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB)
                yield Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        return
    except ImportError:
        pass
    if not shutil.which("pdftoppm"):
        raise RuntimeError("Install PyMuPDF (pip install pymupdf) or poppler (pdftoppm).")
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), f"{tmp}/p"], check=True)
        for f in sorted(Path(tmp).glob("p-*.png")):
            yield Image.open(f).convert("RGB").copy()


def remove_red_ink(img):
    """Whiten red pen marks / circles / numbers so OCR sees the printed text."""
    a = np.array(img)                                   # uint8 copy (no slow int64 conversion)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    mask = (r > 120) & (r.astype(np.int16) - np.maximum(g, b) > 50)
    a[mask] = 255
    return Image.fromarray(a)


def page_count(pdf):
    try:
        import fitz
        with fitz.open(str(pdf)) as doc:
            return len(doc)
    except ImportError:
        out = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout
        m = re.search(r"Pages:\s+(\d+)", out)
        return int(m[1]) if m else 0


def _ocr_page(img, i, pdf, debug_dir, keep):
    """Worker: clean one page image and OCR it (tesseract runs as a separate process,
    so threads really do run in parallel)."""
    clean = remove_red_ink(img)
    if keep is not None:
        keep[i - 1] = clean
    text = pytesseract.image_to_string(clean, config="--psm 6")
    if debug_dir:
        Path(debug_dir).mkdir(parents=True, exist_ok=True)
        (Path(debug_dir) / f"{Path(pdf).stem}_p{i}.txt").write_text(text, encoding="utf-8")
    return text


_OCR_CACHE = {}          # (path, size, mtime) -> page texts: Cut PDF pages then Extract = OCR only once


def _ocr_key(pdf):
    st = Path(pdf).stat()
    return (str(Path(pdf).resolve()), st.st_size, st.st_mtime_ns)


def ocr_pdf(pdf, debug_dir=None, keep=None, progress=None):
    """OCR all pages in parallel (one tesseract process per CPU core). Page order is preserved.
    The text is remembered, so 'Cut PDF pages' followed by 'Extract' does not OCR twice."""
    try:
        key = _ocr_key(pdf)
    except OSError:
        key = None
    if keep is None and key in _OCR_CACHE:
        if progress:
            progress(len(_OCR_CACHE[key]), len(_OCR_CACHE[key]))
        return list(_OCR_CACHE[key])
    pages = _ocr_pdf_uncached(pdf, debug_dir, keep, progress)
    if key is not None and keep is None:
        _OCR_CACHE[key] = list(pages)
    return pages


def _ocr_pdf_uncached(pdf, debug_dir=None, keep=None, progress=None):
    setup_tesseract()
    total = page_count(pdf)
    keep_map = {} if keep is not None else None
    results = [None] * total
    gate = threading.Semaphore(OCR_WORKERS * 2)      # limit page images held in memory
    done = 0

    def finished(fut):
        gate.release()

    with ThreadPoolExecutor(max_workers=OCR_WORKERS) as pool:
        futs = []
        for i, img in enumerate(render_pages(pdf), 1):   # render in this thread, OCR in the pool
            gate.acquire()
            f = pool.submit(_ocr_page, img, i, pdf, debug_dir, keep_map)
            f.add_done_callback(finished)
            futs.append(f)
        for f in as_completed(futs):
            f.result()                                   # re-raise OCR errors
            done += 1
            if progress:
                progress(done, len(futs))
    pages = [f.result() for f in futs]
    if keep is not None:
        keep.extend(keep_map[i] for i in sorted(keep_map))
    return pages


# ------------------------------- parsing --------------------------------
def classify(text):
    t = text.upper()
    if "SINGLE ADMIN" in t or "OFFICE OF LODGEMENT" in t:
        return "sad"
    if "INFORMATION PAGE" in t:
        return "info"
    if re.search(r"COMMERCIAL\s*INV", t):
        return "invoice"
    if "PACKING LIST" in t:
        return "pl"
    return None


def num(s):
    return float(s.replace(",", ""))


def vote(cands):
    """cands = [(value, source)] in priority order -> (best_value, sources, all_distinct)"""
    cands = [c for c in cands if c[0] not in (None, "")]
    if not cands:
        return None, [], []
    counts = Counter(v for v, _ in cands)
    top = max(counts.values())
    best = next(v for v, _ in cands if counts[v] == top)      # priority breaks ties
    return best, [s for v, s in cands if v == best], list(counts)


def first(pattern, text, flags=re.I | re.M, group=1):
    m = re.search(pattern, text, flags)
    return m.group(group) if m else None


INV_NO_PATTERNS = (
    r"[I|1l]NV\s*[+/&]\s*P(?:KL|L)\s*:\s*([A-Z$]{1,2}\d{6}(?:\s?-\s?\d{1,2})?)",          # "INV+PL: SC260456" / "INV&PKL: SC260461"
    r"INV\.?\s*N[O0K]\.?\s*[:\-]?\s*([A-Z$]{1,2}\d{6}(?:\s?-\s?\d{1,2})?)")             # "INV.NO. SC260461-1"  ("$C" = OCR misread of "SC")


def _inv_no(text, patterns=INV_NO_PATTERNS):
    for p in patterns:
        v = first(p, text)
        if v:
            return v.upper().replace("$", "S").replace(" ", "")
    return None


def _weight(kind, text):
    """TOTAL N.W. / G.W. from an invoice. Tolerates OCR damage: 'TOTALN.W: 330 KGS' (no space,
    lost decimal point) -> 3.30. Invoices always print 2 decimals, so '1130' is read as 11.30."""
    m = re.search(rf"TOTAL\s*{kind}\.?\s*W\.?\s*:*\s*([\d,]+(?:\.\d{{1,2}})?)", text, re.I)
    if not m:
        return None
    v = m[1].replace(",", "")
    if "." not in v and len(v) >= 3:
        v = v[:-2] + "." + v[-2:]
    return float(v)


def _invoice_groups(pages):
    """One text block per COMMERCIAL INVOICE number found in the PDF (header page + the
    header-less continuation pages that follow it). Pages with the same number are merged."""
    return [t for t, _ in _invoice_groups_idx(pages)]


def _invoice_groups_idx(pages):
    """Same as _invoice_groups but returns [(text, [page indexes])] so Gemini can be shown
    exactly the pages of ONE invoice."""
    heads = [i for i, t in enumerate(pages) if classify(t) == "invoice" and is_invoice_header(t)]
    groups, gpages = {}, {}
    for n, i in enumerate(heads):
        end = heads[n + 1] if n + 1 < len(heads) else len(pages)
        parts, j = [pages[i]], i + 1
        while j < end and classify(pages[j]) is None and not is_invoice_header(pages[j]):
            parts.append(pages[j])
            j += 1
        key = _inv_no(pages[i], INV_NO_PATTERNS[1:]) or f"?page{i + 1}"
        groups[key] = groups.get(key, "") + "\n" + "\n".join(parts)
        gpages.setdefault(key, []).extend(range(i, j))
    return [(groups[k], gpages[k]) for k in groups]


def _fill_missing_amount(recs):
    """Several invoices under one customs declaration: the SAD shows the declaration TOTAL. If exactly
    ONE invoice amount could not be read, it must be total - (the others). Marked 'sad-balance' so
    it is reported as WARN (verify), not silently trusted."""
    total = recs[0].get("_sad_total")
    miss = [r for r in recs if r["amount"]["value"] in (None, "")]
    if total and len(miss) == 1 and all(r["amount"]["value"] for r in recs if r is not miss[0]):
        v = round(total - sum(r["amount"]["value"] for r in recs if r is not miss[0]), 2)
        if v > 0:
            miss[0]["amount"] = {"value": v, "sources": ["sad-balance"], "distinct": [v]}


def _extract_bundle(pages, extra=None):
    """One invoice in the PDF -> one result dict (as before).
    Several invoices (one customs E number, e.g. SC260461, SC260461-1 ... -5) -> the first
    invoice's result dict with the others in res["_rest"]; each becomes its own Excel row."""
    docs = {}
    for t in pages:
        k = classify(t)
        if k and k not in docs:
            docs[k] = t
    gi = _invoice_groups_idx(pages)
    groups = [t for t, _ in gi]
    first_idx = {}
    for i, t in enumerate(pages):
        k = classify(t)
        if k and k not in first_idx:
            first_idx[k] = i
    shared = [first_idx[k] for k in ("sad", "info") if k in first_idx]
    if len(groups) >= 2:
        info, sad = docs.get("info", ""), docs.get("sad", "")
        recs = [_extract_one(g, "", info, sad, extra if n == 0 else None, multi=True)
                for n, g in enumerate(groups)]
        for r, (_, ip) in zip(recs, gi):      # pages Gemini must see to repair THIS invoice
            r["_pages"] = sorted(set(shared + ip))
        _fill_missing_amount(recs)
        recs[0]["_rest"] = recs[1:]
        return recs[0], sorted(docs)
    inv, pl, info, sad = (docs.get(k, "") for k in ("invoice", "pl", "info", "sad"))
    if len(groups) == 1:          # invoice header page + its continuation pages (more description/HTS lines)
        inv = groups[0]
    r = _extract_one(inv, pl, info, sad, extra)
    r["_pages"] = sorted(set(first_idx.values()) | set(gi[0][1] if gi else []))
    return r, sorted(docs)



def extract(pages, extra=None):
    """One PDF may hold SEVERAL customs declarations (E numbers). Loop from E number to E number,
    read each declaration on its own, and chain all invoices (in order) into one result:
    res = first invoice, group_of(res) = every invoice of every E number (one Excel row each)."""
    eg = [g for g in split_by_e(pages) if g["inv"]]
    if len(eg) >= 2:
        subs, flat = [], []
        for n, g in enumerate(eg):
            idx = sorted([g["sad"]] + g["inv"])
            r, d = _extract_bundle([pages[i] for i in idx], extra if n == 0 else None)
            for rec in group_of(r):
                rec["_pages"] = [idx[j] for j in rec.get("_pages", [])]
            r["_e_label"] = g["e"] or f"UNKNOWN{n + 1}"
            subs.append((r, d))
            flat += group_of(r)
        res = subs[0][0]
        res["_all"], res["_subs"] = flat, subs
        return res, sorted({k for _, d in subs for k in d})
    return _extract_bundle(pages, extra)



HTS_LINE = re.compile(r"^[ \t]*(?P<desc>[^\n]*?)\s*\(?\s*HTS\.?\s*CODE\s*[:\-~]?\s*(?P<hts>\d{6,10})(?!\d)", re.I | re.M)


def _clean_desc(d):
    d = d.replace("|", "").replace("\u2019", "'").replace("`", "'").replace("\u00b4", "'")
    return re.sub(r"\s+", " ", d).strip(" (:-").upper()


def _multi_items(text):
    """Every '<DESCRIPTION> (HTS.CODE : 6109902000)' line of one document, in order of first
    appearance, duplicates removed. Returns (descriptions, hts_codes)."""
    descs, codes = [], []
    for m in HTS_LINE.finditer(text):
        d = _clean_desc(m["desc"])
        if d and d not in descs:
            descs.append(d)
        if m["hts"] not in codes:
            codes.append(m["hts"])
    return descs, codes


_ITEM_LINE = re.compile(r"^\s*\d{10}\b(?P<rest>.*)$", re.M)         # "4502411820 AA10007527 LF1980 2 18.33 36.66"
_NUM2 = re.compile(r"(?<![\d.])\d[\d,]*\.\d{2}(?!\d)")


def _line_sums(text):
    """Add up the item lines of ONE invoice (PO no. ... qty price amount). Used when the GRAND TOTAL line is
    unreadable, and to cross-check it. Tolerates OCR damage inside a line (garbled price or qty).
    Returns (qty or None if any line's qty is unreadable, amount, number of lines)."""
    qty, amt, n, qty_ok = 0, 0.0, 0, True
    for m in _ITEM_LINE.finditer(text):
        nums = [num(x) for x in _NUM2.findall(m["rest"])]
        if not nums:
            continue
        hit = None
        for pr, a in zip(nums, nums[1:]):                       # consecutive (price, amount) with amount = k x price
            k = round(a / pr) if pr else 0
            if k >= 1 and abs(a - k * pr) <= 0.011 * k + 0.005:
                hit = (k, a)
                break
        if hit:
            qty += hit[0]; amt += hit[1]
        else:                                                   # price unreadable: amount = last number, qty = small integer
            amt += nums[-1]
            ints = [t for t in m["rest"].split()[2:] if t.isdigit() and len(t) <= 3]
            if ints:
                qty += int(ints[0])
            else:
                qty_ok = False
        n += 1
    return (qty if qty_ok else None), round(amt, 2), n


def _extract_one(inv, pl, info, sad, extra=None, multi=False):
    """multi=True: `inv` is ONE invoice of a multi-invoice declaration. The SAD page then holds
    the totals / item 1 of the whole declaration, so it is only used for the shared fields
    (customs no./date, rate, destination ...), never to vote on this invoice's amount or HTS."""
    C = {k: [] for k in ["inv_no", "customer", "customs_no", "customs_date", "hts", "description",
                         "qty", "amount", "destination", "rate", "nw", "gw", "fabric",
                         "onboard_date", "procedure", "mode"]}
    allt = "\n".join(t for t in (sad, inv, pl, info) if t)

    # 1 INV.NO
    for src, t in ((("invoice", inv),) if multi else
                   (("info", info), ("pl", pl), ("invoice", inv), ("sad", sad))):
        v = _inv_no(t)
        if v:
            C["inv_no"].append((v, src))

    # 2 Customer (consignee keyword -> short name)
    for src, t in (("invoice", inv), ("pl", pl), ("info", info), ("sad", sad)):
        for key, name in CUSTOMER_MAP.items():
            if key in t.upper():
                C["customer"].append((name, src))
                break

    # 3/4 Customs declaration number + date
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+(E\s?\d{4,6})\b", info)
    if m:
        C["customs_no"].append((m[2].replace(" ", ""), "info"))
        C["customs_date"].append((datetime.strptime(m[1], "%d/%m/%Y"), "info"))
    m = re.search(r"\b(E\s?\d{5})\s+(\d{2}/\d{2}/\d{4})", sad)
    if m:
        C["customs_no"].append((m[1].replace(" ", ""), "sad"))
        C["customs_date"].append((datetime.strptime(m[2], "%d/%m/%Y"), "sad"))
    v = first(r"Customs Declaration No\.?:?\s*(E\d{5})", allt)
    if v:
        C["customs_no"].append((v, "cargo"))

    # 5 HTS code
    # several descriptions / HTS codes in one invoice -> "A/B/C" and "6104630000/6109902000"
    for src, t in (("pl", pl), ("invoice", inv)):
        _, codes = _multi_items(t)
        if codes:
            C["hts"].append(("/".join(codes), src))
    m = re.search(r"\b(6\d{7})\s*\|?\s*(00)\b", sad)
    if m and not multi:
        sad_code = m[1] + m[2]
        joined = [v for v, _ in C["hts"]]
        # SAD item 1 shows only ONE code: it confirms a multi-code value when it is one of them
        hit = next((v for v in joined if any(sad_code == c or (len(c) >= 6 and sad_code.startswith(c))
                                             for c in v.split("/"))), None)
        C["hts"].append((hit or sad_code, "sad"))

    # 6 Description + 13 Fabric
    for src, t in (("pl", pl), ("invoice", inv)):
        descs, _ = _multi_items(t)
        if descs:
            C["description"].append(("/".join(descs), src))
        v = first(r"\(\s*(KNITTED|WOVEN)\s*\)", t) or first(r"\b(KNITTED|WOVEN)\b", t)
        if v:
            C["fabric"].append((v.upper(), src))

    # 7 Qty / 11 NW / 12 GW  (packing list grand total row)
    m = re.search(r"GRAND TOTAL:?\s*([\d,]+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", pl, re.I)
    if m:
        C["qty"].append((int(num(m[1])), "pl"))
        C["nw"].append((num(m[2]), "pl"))
        C["gw"].append((num(m[3]), "pl"))
    for src, t in (("invoice", inv), ("pl", pl)):
        v = _weight("N", t)
        if v is not None:
            C["nw"].append((v, src))
        v = _weight("G", t)
        if v is not None:
            C["gw"].append((v, src))
    # 8 Total price (+ qty/price cross-check) from the invoice item line
    m = re.search(r"([\d,]{3,})\s+(\d+\.\d{2})\s+([\d,]+\.\d{2})\s*$", inv, re.M)
    unit_price = None
    if m:
        C["qty"].append((int(num(m[1])), "invoice"))
        unit_price = num(m[2])
        C["amount"].append((num(m[3]), "invoice"))
    m = re.search(r"GRAND TOTAL:?\s*(?:([\d,]+)\s+)?([\d,]+\.\d{2})", inv, re.I)
    if m:
        C["amount"].append((num(m[2]), "invoice"))
        if m[1]:
            C["qty"].append((int(num(m[1])), "invoice"))

    lq, la, ln = _line_sums(inv)
    if ln and not any(s == "invoice" for _, s in C["amount"]):          # GRAND TOTAL not read -> sum of item lines
        C["amount"].append((la, "lines"))
        if lq is not None and not any(s == "invoice" for _, s in C["qty"]):
            C["qty"].append((lq, "lines"))

    # SAD: amount / rate / qty / procedure / flight date
    v = first(r"USD\s*[\|\[\(\{]?\s*([\d,]+\.\d{2})", sad)
    sad_total = num(v) if v else None
    if v and not multi:
        C["amount"].append((num(v), "sad"))
    m = re.search(r"(\d)[,.\s]?(\d{3})[.,]\s?(\d{4})\b", sad)     # 10 Exchange rate (tolerant of OCR)
    if m:
        C["rate"].append((float(f"{m[1]}{m[2]}.{m[3]}"), "sad"))
    amt = sad_total if multi else vote(C["amount"])[0]                              # safety net: customs value (KHR) / USD amount
    if amt:
        for kv in re.findall(r"\b(\d{1,3}(?:,\d{3}){2,3})\b", sad):
            q = num(kv) / amt
            if 3000 < q < 6000 and abs(q - round(q)) < 0.01:
                C["rate"].append((float(round(q)), "sad-derived"))
                break
    v = first(r"NAR\s*([\d,]+)", sad)
    if v and not C["qty"] and not multi:      # SAD page 1 shows item 1 only: use it just when no invoice/packing-list total
        C["qty"].append((int(num(v)), "sad"))
    v = first(r"\b(\d{4})\s*\|?\s*000(?!\d)", sad)
    if v:
        C["procedure"].append((int(v), "sad"))
    v = first(r"AK?\w*\s*/\s*(\d{2}\.\d{2}\.\d{4})", sad)
    if v:
        C["onboard_date"].append((datetime.strptime(v, "%d.%m.%Y"), "sad"))
    else:                       # SAD box 18/21: "8K0803 24/10/2026" = flight no. + departure date
        m = re.search(r"\b(?=[A-Z0-9]?[A-Z])[A-Z0-9]{2}\d{3,4}\s*[/|]?\s*(\d{2})[/-](\d{2})[/-](\d{4})\b", sad)
        if m:
            try:
                C["onboard_date"].append((datetime(int(m[3]), int(m[2]), int(m[1])), "sad"))
            except ValueError:
                pass

    # 9 Destination
    for src, t in (("invoice", inv), ("pl", pl)):
        m = (re.search(r"^BY\s+(AIR|SEA)\s+FROM\b.*\bTO\s+([A-Z][A-Z .]*?)\s*$", t, re.M)   # "... TO GERMANY"
             or re.search(r"^BY\s+(AIR|SEA)\s+FROM\b.*\s(\S{2,4})\s*$", t, re.M))       # "... TO UK"
        if m:
            C["destination"].append((COUNTRY_MAP.get(m[2].upper().strip("!|"), m[2].title()), src))
            C["mode"].append((m[1].upper(), src))
    sad_dest = re.sub(r"Taiwan,?\s*Province\s+of\s+China", "Taiwan", sad, flags=re.I)
    for name in set(COUNTRY_MAP.values()) | {"Taiwan"}:
        if re.search(rf"Destination.*?\n.*?\b{name}\b", sad_dest, re.S):
            C["destination"].append((name, "sad"))

    for f, v in (extra or {}).items():          # Gemini values: lowest priority voter
        if f in C and v not in (None, ""):
            C[f].append((v, "gemini"))

    res = {}
    for k, cands in C.items():
        best, srcs, distinct = vote(cands)
        res[k] = {"value": best, "sources": srcs, "distinct": distinct}
    res["_unit_price"] = unit_price
    res["_lines"] = (lq, la, ln)
    res["_sad_total"] = sad_total
    return res


def _validate_one(res, docs):
    issues = []
    for need in ("sad", "invoice", "pl", "info"):
        if need not in docs:
            issues.append(("WARN", f"document not found in PDF: {need}"))
    for k, d in res.items():
        if k.startswith("_") or k == "mode":
            continue
        if k in COLUMNS and d["value"] in (None, ""):
            issues.append(("ERROR", f"{k}: not found"))
        elif len(d["distinct"]) > 1:
            issues.append(("WARN", f"{k}: documents disagree {d['distinct']} -> using {d['value']!r}"))
        elif k in COLUMNS and d["sources"] == ["lines"]:
            issues.append(("WARN", f"{k}: GRAND TOTAL not readable - summed from the item lines; verify"))
        elif k in COLUMNS and d["sources"] == ["sad-balance"]:
            issues.append(("WARN", f"{k}: not read from the invoice - worked out as SAD total minus the other invoices; verify"))
        elif k in COLUMNS and d["sources"] == ["gemini"]:
            issues.append(("WARN", f"{k}: value comes from Gemini only - verify against the document"))
        elif k in COLUMNS and len(d["sources"]) == 1:
            issues.append(("INFO", f"{k}: confirmed by one document only ({d['sources'][0]})"))
    g = lambda k: res[k]["value"]
    if g("inv_no") and not re.fullmatch(r"[A-Z]{1,2}\d{6}(-\d{1,2})?", g("inv_no")):
        issues.append(("WARN", f"inv_no format looks odd: {g('inv_no')}"))
    if len(res["mode"]["distinct"]) > 1:
        issues.append(("WARN", "invoice and packing list disagree on transport mode "
                       f"{res['mode']['distinct']} (AWB/SAD show AIR) - check source documents"))
    want = "62" if g("fabric") == "WOVEN" else "61"          # 61 = knitted, 62 = woven garments
    for code in (g("hts") or "").split("/"):
        if code and not code.startswith(want):
            issues.append(("WARN", f"HTS {code} unexpected for {(g('fabric') or 'knit').lower()} garments"))
    if g("qty") and res["_unit_price"] and g("amount"):
        if abs(g("qty") * res["_unit_price"] - g("amount")) > 0.5:
            issues.append(("WARN", f"qty x unit price != amount ({g('qty')} x {res['_unit_price']})"))
    if g("nw") and g("gw") and g("nw") > g("gw"):
        issues.append(("ERROR", "net weight is greater than gross weight"))
    return issues


def group_of(res):
    """[first invoice, other invoices of the same declaration ...]"""
    return res.get("_all") or ([res] + list(res.get("_rest") or []))


def validate(res, docs):
    """Single invoice: as before. Several invoices under one E number: every invoice is checked,
    messages are tagged [INV.NO], and the invoice amounts must add up to the SAD total."""
    if res.get("_subs"):                       # several E numbers in this file: check each on its own
        out, miss = [], False
        for r, d in res["_subs"]:
            rc = {k: v for k, v in r.items() if k not in ("_all", "_subs")}
            for l, m in validate(rc, d):
                if m.startswith("document not found"):
                    miss = True
                elif l != "INFO":
                    out.append((l, f"[{r['_e_label']}] {m}"))
        if miss:
            out.insert(0, ("WARN", "document not found in PDF: " + ", ".join(
                x for x in ("sad", "invoice", "pl", "info") if x not in docs)))
        recs = group_of(res)
        nos = [r["inv_no"]["value"] for r in recs if r["inv_no"]["value"]]
        dups = sorted({n for n in nos if nos.count(n) > 1})
        if dups:
            out.append(("ERROR", f"same INV.NO found more than once in the file: {dups}"))
        out.insert(0, ("INFO", f"{len(res['_subs'])} customs numbers, {len(recs)} invoices -> {len(recs)} Excel rows "
                               f"(table below shows the first invoice only)"))
        return out
    issues = _validate_one(res, docs)
    recs = group_of(res)
    if len(recs) == 1:
        return issues
    full = ("sad", "invoice", "pl", "info")
    out = [(l, m) for l, m in issues if m.startswith("document not found")]
    for r, d in zip(recs, [docs] + [full] * (len(recs) - 1)):
        tag = r["inv_no"]["value"] or "?"
        out += [(l, f"[{tag}] {m}") for l, m in _validate_one(r, d)
                if l != "INFO" and not m.startswith("document not found")]
    nos = [r["inv_no"]["value"] for r in recs if r["inv_no"]["value"]]
    dups = sorted({n for n in nos if nos.count(n) > 1})
    if dups:
        out.append(("ERROR", f"same INV.NO read on more than one invoice page: {dups} - check OCR / edit"))
    amts = [r["amount"]["value"] for r in recs]
    total, sad_total = sum(a for a in amts if a), res.get("_sad_total")
    if sad_total and all(a is not None for a in amts):
        if abs(total - sad_total) > 0.5:
            out.append(("WARN", f"invoice amounts add up to {total:,.2f} but the customs declaration "
                                f"shows {sad_total:,.2f} - an invoice may be missing or misread"))
        else:
            out.append(("INFO", f"invoice amounts add up to the customs declaration total {sad_total:,.2f}"))
    lines = "; ".join(f"{r['inv_no']['value']} qty {r['qty']['value']} USD {r['amount']['value']} "
                      f"NW {r['nw']['value']} GW {r['gw']['value']}" for r in recs)
    out.insert(0, ("INFO", f"{len(recs)} invoices under one customs no. -> {len(recs)} Excel rows "
                           f"(table below shows the first only): {lines}"))
    return out


# ----------------------------- trim PDF pages ----------------------------
def is_invoice_header(text, top_lines=15):
    """True if the page is a Commercial Invoice (header near the top, not a packing list)."""
    head = "\n".join(text.upper().splitlines()[:top_lines])
    return bool(re.search(r"COMMERCIAL\s*INV", head)) and "PACKING LIST" not in head


def _sad_e(text):
    """Customs (E) number printed on a SAD first page, e.g. 'E56006 23/10/2026' -> 'E56006'."""
    m = re.search(r"\b(E\s?\d{5})\s+(\d{2}/\d{2}/\d{4})", text)
    return m[1].replace(" ", "") if m else None


def _is_invoice_continuation(text):
    """Page after an invoice header page that still belongs to the same invoice
    (more item lines / HTS lines / GRAND TOTAL / TOTAL N.W.), not a packing list or other sheet."""
    if classify(text) is not None:
        return False
    return bool(re.search(r"HTS\.?\s*CODE|PRICE\s*/?\s*USD|\bAMOUNT\b|GRAND\s*TOTAL|TOTAL\s*[NG]\.?\s*W", text, re.I))


def split_by_e(pages):
    """Walk the bundle page by page, from E number to E number.
    Every SAD first page with a NEW customs E number starts a new declaration; the COMMERCIAL
    INVOICE pages (header page + its continuation pages) that follow belong to it until the next
    E number.  Returns [{'e': 'E56006', 'sad': page_idx, 'inv': [page_idx, ...]}, ...] (0-based)."""
    groups, by_e, cur, early = [], {}, None, []
    for i, t in enumerate(pages):
        k = classify(t)
        if k == "sad":
            e = _sad_e(t)
            if e and e in by_e:                       # E number seen before -> back to that group
                cur = by_e[e]
            elif e or cur is None:                    # new E number (or first SAD, E unreadable)
                cur = {"e": e, "sad": i, "inv": [], "_last": None}
                if e:
                    by_e[e] = cur
                groups.append(cur)
            # else: 2nd page of the SAD that is already open -> ignored
        elif is_invoice_header(t):
            if cur is None:
                early.append(i)                       # invoice before any SAD: give to first declaration
            else:
                cur["inv"].append(i)
                cur["_last"] = i
        elif cur is not None and cur["_last"] == i - 1 and _is_invoice_continuation(t):
            cur["inv"].append(i)
            cur["_last"] = i
    if early and groups:
        groups[0]["inv"] = early + groups[0]["inv"]
    for g in groups:
        g.pop("_last", None)
    return groups


def pick_pages(pages):
    """Pages to keep for the FIRST declaration (kept for backward compatibility)."""
    groups = split_by_e(pages)
    return sorted([groups[0]["sad"]] + groups[0]["inv"]) if groups else []


def write_pages(pdf, keep, out):
    """Copy the chosen pages (0-based) of pdf into a new file. The original is never touched."""
    try:
        import fitz
        with fitz.open(str(pdf)) as doc:
            doc.select(keep)
            doc.save(str(out), garbage=3, deflate=True)
        return
    except ImportError:
        pass
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        raise RuntimeError("Install PyMuPDF (pip install pymupdf) to cut PDF pages.")
    r, w = PdfReader(str(pdf)), PdfWriter()
    for i in keep:
        w.add_page(r.pages[i])
    with open(out, "wb") as f:
        w.write(f)


SPLIT_FILES = False      # True = ALSO save one extra PDF per E number
WRITE_TRIMMED_PDF = False   # False = "Cut PDF pages" writes NO pdf: result is shown in the dashboard only
WRITE_REPORT_CSV = False    # True  = also save <name>_split_report.csv


def trim_pdf_all(pdf, out=None, debug_dir=None, progress=None, log=print):
    """OCR the bundle, loop from E number to E number and keep, for every customs declaration,
    the SAD first page + its COMMERCIAL INVOICE pages - all in ONE file <name>_trimmed.pdf
    (E number order as in the original). Each E number is then test-read (no extra OCR) so you
    see which ones need a patch. Returns a list of dicts: e, pages, status (OK/CHECK/PATCH), notes."""
    pages = ocr_pdf(pdf, debug_dir, None, progress)
    groups = split_by_e(pages)
    if not groups:
        raise RuntimeError("Customs declaration (SAD) page not found - nothing cut")
    pdf = Path(pdf)
    dst = Path(out) if out else pdf.with_name(pdf.stem + "_trimmed.pdf")
    if WRITE_TRIMMED_PDF and dst.resolve() == pdf.resolve():
        raise RuntimeError("Output would overwrite the original PDF")
    log(f"{len(groups)} customs declaration(s) found: {', '.join(g['e'] or '?' for g in groups)}")
    results, keep_all = [], []
    for n, g in enumerate(groups, 1):
        keep = sorted([g["sad"]] + g["inv"])
        label = g["e"] or f"UNKNOWN{n}"
        r = {"e": label, "pages": [i + 1 for i in keep], "status": "PATCH", "notes": ""}
        results.append(r)
        if not g["inv"]:
            r["notes"] = "no COMMERCIAL INVOICE page found after this E number - not kept"
            continue
        keep_all += keep
        try:                                   # test-read from the OCR text we already have
            res, docs = extract([pages[i] for i in keep])
            issues = validate(res, docs)
        except Exception as e:
            issues = [("ERROR", f"{type(e).__name__}: {e}")]
        errs = [m for l, m in issues if l == "ERROR"]
        warns = [m for l, m in issues if l == "WARN" and not m.startswith("document not found")]
        r["status"] = "PATCH" if errs else ("CHECK" if warns else "OK")
        r["notes"] = error_reason(errs) if errs else "; ".join(warns[:3])
        if WRITE_TRIMMED_PDF and SPLIT_FILES:
            write_pages(pdf, keep, dst.with_name(dst.stem.removesuffix("_trimmed") + f"_{label}_trimmed" + dst.suffix))
    if not keep_all:
        raise RuntimeError("No COMMERCIAL INVOICE page found - nothing cut")
    keep_all = sorted(set(keep_all))
    if WRITE_TRIMMED_PDF:
        write_pages(pdf, keep_all, dst)
        log(f"Kept pages {[i + 1 for i in keep_all]} of {len(pages)} -> {dst.name} (one file, {len(groups)} E number(s))")
    else:
        dst = pdf          # no new file: Extract reads the original PDF and uses the same page cut
        log(f"Cut done: keeping pages {[i + 1 for i in keep_all]} of {len(pages)} (no PDF written - press 1. Extract)")
    for r in results:
        log(f"  [{r['status']:<5}] {r['e']:<8} pages {r['pages']}" + (f"  {r['notes']}" if r["notes"] else ""))
    bad = [r["e"] for r in results if r["status"] == "PATCH"]
    chk = [r["e"] for r in results if r["status"] == "CHECK"]
    log(f"SUMMARY: {len(results) - len(bad) - len(chk)} OK, {len(chk)} to check, {len(bad)} need a patch"
        + (f": {', '.join(bad)}" if bad else ""))
    try:
        if not WRITE_REPORT_CSV:
            raise OSError
        import csv
        rep = pdf.with_name(pdf.stem + "_split_report.csv")
        with open(rep, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["E number", "status", "pages (original PDF)", "notes"])
            for r in results:
                w.writerow([r["e"], r["status"], " ".join(map(str, r["pages"])), r["notes"]])
        log(f"Report saved: {rep.name}")
    except OSError as e:
        if WRITE_REPORT_CSV:
            log(f"(could not write report: {e})")
    trim_pdf_all.out_file = dst
    return results


def trim_pdf(pdf, out=None, debug_dir=None, progress=None, log=print):
    """Cut the bundle into ONE trimmed PDF (all E numbers). Returns (file, kept_pages_1_based)."""
    results = trim_pdf_all(pdf, out, debug_dir, progress, log)
    trim_pdf.last_results = results
    kept = sorted({p for r in results if r["status"] != "PATCH" or r["pages"] for p in r["pages"]})
    return trim_pdf_all.out_file, kept


# ------------------------------- Excel ----------------------------------
def _norm(v):
    return re.sub(r"\s", "", str(v)).upper() if v else ""


def plan_row(ws, res, force=False):
    """Decide which row the new shipment goes to, or raise RuntimeError (= skip this file)."""
    inv, e_new = res["inv_no"]["value"], _norm(res["customs_no"]["value"])
    last, e_rows = None, {}
    for r in range(FIRST_DATA_ROW, max(ws.max_row, FIRST_DATA_ROW) + 1):
        v = ws[f"{COLUMNS['inv_no']}{r}"].value
        if not v:
            continue
        if _norm(v) == _norm(inv) and not force:
            raise RuntimeError(f"Invoice {inv} already exists in row {r}. Use --force to add anyway.")
        last = r
        e = _norm(ws[f"{COLUMNS['customs_no']}{r}"].value)
        if e:
            e_rows.setdefault(e, []).append(r)
    if last is None:
        return FIRST_DATA_ROW
    e_last = _norm(ws[f"{COLUMNS['customs_no']}{last}"].value)
    if e_new and e_new in e_rows and not force and DUP_E_POLICY != "allow":
        rows = ", ".join(map(str, e_rows[e_new]))
        shown = res["customs_no"]["value"]
        if DUP_E_POLICY == "skip":
            raise RuntimeError(f"Duplicate customs no. {shown}: already in row {rows}. Use --force to add anyway.")
        if e_new != e_last:
            raise RuntimeError(f"Duplicate customs no. {shown}: already in row {rows}, but the last group is "
                               f"{ws[f'{COLUMNS['customs_no']}{last}'].value}. Writing here would split the group. "
                               f"Use --force to add anyway.")
    return last + 1 if (e_new and e_new == e_last) else last + 2       # new E no. -> one blank row


def _check_invoices(ws, recs):
    """Refuse the whole declaration if any of its invoice numbers is already in the sheet."""
    have = {}
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        v = ws[f"{COLUMNS['inv_no']}{r}"].value
        if v:
            have.setdefault(_norm(v), r)
    hit = [(rec["inv_no"]["value"], have[_norm(rec["inv_no"]["value"])])
           for rec in recs if _norm(rec["inv_no"]["value"]) in have]
    if hit:
        raise RuntimeError("Invoice " + ", ".join(f"{n} already exists in row {r}" for n, r in hit)
                           + ". Delete the old row(s) or use --force to add anyway.")


def precheck(xlsx, res, force=False):
    """Same duplicate checks as writing, without writing. Returns an error text or None."""
    try:
        ws = openpyxl.load_workbook(xlsx, read_only=False).worksheets[0] if not SHEET_NAME else \
            openpyxl.load_workbook(xlsx)[SHEET_NAME]
        recs = group_of(res)
        if len(recs) > 1 and not force:
            _check_invoices(ws, recs)
        plan_row(ws, recs[0], force)
    except RuntimeError as e:
        return str(e)
    return None


def fill_workbook(xlsx, res, out, extras=False, force=False):
    wb = openpyxl.load_workbook(xlsx)
    ws = wb[SHEET_NAME] if SHEET_NAME else wb.worksheets[0]
    recs = group_of(res)
    if len(recs) > 1 and not force:
        _check_invoices(ws, recs)
    cols = dict(COLUMNS, **(EXTRA_COLUMNS if extras else {}))
    fc, ft = FORMULA_COLUMN
    rows = []
    for rec in recs:          # same E number -> plan_row puts each invoice directly under the previous one
        row = plan_row(ws, rec, force)
        for field, col in cols.items():
            val = rec[field]["value"]
            if val is not None:
                ws[f"{col}{row}"] = val
        if ws[f"{fc}{row}"].value is None:
            ws[f"{fc}{row}"] = ft.format(r=row)
        rows.append(row)
    res["_rows"] = rows
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)              # nothing is saved if any invoice was refused above
    return rows[0]


def backup(path):
    b = Path(path).parent / "backup"
    b.mkdir(exist_ok=True)
    dst = b / f"{Path(path).stem}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, dst)
    return dst


GEMINI_MAX_CALLS = 12        # cost guard: max Gemini requests per PDF (one per invoice that needs help)


def _apply_gemini(rec, extra):
    """Add Gemini's answers to ONE already-voted invoice record (same rule as before:
    Gemini is the lowest-priority voter - it fills gaps and breaks ties, it never overrules OCR)."""
    for f, v in extra.items():
        if f not in COLUMNS or v in (None, "") or f not in rec:
            continue
        d = rec[f]
        if d["value"] in (None, ""):
            rec[f] = {"value": v, "sources": ["gemini"], "distinct": [v]}
        elif v in d["distinct"]:
            if d["value"] != v and len(d["distinct"]) > 1 and d["sources"] and len(d["sources"]) == 1:
                d["value"], d["sources"] = v, ["gemini"]      # documents disagreed, Gemini sides with another one
            elif d["value"] == v:
                d["sources"] = d["sources"] + ["gemini"]
        else:
            d["distinct"] = d["distinct"] + [v]


def analyze(pdf, use_gemini=False, debug_dir=None, log=print, progress=None):
    """OCR -> vote -> (optional Gemini for missing/disputed fields, PER INVOICE) -> validate."""
    imgs = []
    pages = ocr_pdf(pdf, debug_dir, imgs if use_gemini else None, progress)
    res, docs = extract(pages)
    gem_notes = []
    if use_gemini:
        recs = group_of(res)               # every invoice of every customs number, not only the first
        jobs = []
        for k, rec in enumerate(recs, 1):
            wanted = [f for f in COLUMNS if rec[f]["value"] in (None, "") or len(rec[f]["distinct"]) > 1]
            if wanted:
                jobs.append((rec, wanted, k))
        if not jobs:
            log("Gemini: not needed, OCR found every field")
        else:
            try:
                import gemini_fallback as gf
            except Exception as e:
                gf = None
                gem_notes.append(f"Gemini skipped ({type(e).__name__}: {e})")
            for n, (rec, wanted, k) in enumerate(jobs):
                if gf is None:
                    break
                tag = rec["inv_no"]["value"] or f"invoice {k}"
                if n >= GEMINI_MAX_CALLS:
                    gem_notes.append(f"Gemini call limit ({GEMINI_MAX_CALLS}) reached - [{tag}] not checked")
                    continue
                idx = [i for i in rec.get("_pages", []) if i < len(imgs) and imgs[i] is not None]
                if not idx:
                    gem_notes.append(f"Gemini skipped [{tag}]: no page images")
                    continue
                try:
                    log(f"Gemini: [{tag}] sending {len(idx)} page(s) to Google for fields {wanted}")
                    hint = rec["inv_no"]["value"] if len(recs) > 1 else None
                    _apply_gemini(rec, gf.ask([imgs[i] for i in idx], wanted, hint=hint))
                except Exception as e:       # no key / no internet / API error: keep the OCR result
                    gem_notes.append(f"Gemini skipped [{tag}] ({type(e).__name__}: {e}); OCR could not settle {wanted}")
                    if isinstance(e, RuntimeError) and "key" in str(e).lower():
                        break                # no point trying the other invoices
    issues = validate(res, docs)
    for note in reversed(gem_notes):
        issues.insert(0, ("WARN", note))
    return res, docs, issues


def set_manual(res, field, text):
    """Let the user type a value (GUI double-click). Raises ValueError if it cannot be parsed."""
    import gemini_fallback as gf
    v = gf._clean(field, text)
    if v is None:
        raise ValueError(f"'{text}' is not a valid value for {field}")
    res[field] = {"value": v, "sources": ["manual"], "distinct": [v]}


def save_skipped_result(pdf, res, issues):
    """Write what WAS extracted from a skipped PDF (value, source, problem) to skipped_results.csv,
    one line per field per invoice, so the result can be checked even though nothing went to Excel."""
    import csv
    cols = dict(COLUMNS, **EXTRA_COLUMNS)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new = not ERROR_RESULT_LOG.exists()
    with open(ERROR_RESULT_LOG, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["time", "file", "record", "col", "field", "value", "sources", "status"])
        recs = group_of(res) if res else []
        for rec in recs:
            tag = rec.get("_e_label") or rec["inv_no"]["value"] or "?"
            for field, col in cols.items():
                d = rec.get(field)
                if not d:
                    continue
                v = d["value"]
                v = v.strftime("%d/%b/%Y") if isinstance(v, datetime) else v
                status = "NOT FOUND" if v in (None, "") else ("DISAGREE " + str(d["distinct"]) if len(d["distinct"]) > 1 else "")
                w.writerow([stamp, Path(pdf).name, tag, col, field, "" if v is None else v,
                            ",".join(d["sources"] or []), status])
        for lvl, msg in (issues or []):
            if lvl != "INFO":
                w.writerow([stamp, Path(pdf).name, "", "", f"[{lvl}]", msg, "", ""])


def record_skip(pdf, reason, log=print, save=True, res=None, issues=None):
    """A PDF could not be processed: show a console message, list it in skipped_files.csv,
    and copy the PDF into error_files/ so it can be checked later. (The original is never moved.)
    If res/issues are given, the partial extraction result is saved to skipped_results.csv too."""
    import csv
    where = ""
    if save:
        try:
            new = not ERROR_LOG.exists()
            with open(ERROR_LOG, "a", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                if new:
                    w.writerow(["time", "file", "reason"])
                w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), str(pdf), reason])
            where = f"  (listed in {ERROR_LOG.name}"
            if ERROR_DIR and Path(pdf).is_file():
                ERROR_DIR.mkdir(exist_ok=True)
                shutil.copy2(pdf, ERROR_DIR / Path(pdf).name)
                where += f", copy in {ERROR_DIR.name}/"
            if res is not None:
                save_skipped_result(pdf, res, issues)
                where += f", extracted values in {ERROR_RESULT_LOG.name}"
            where += ")"
        except OSError as e:
            where = f"  (could not write error file: {e})"
    log(f"SKIPPED: {Path(pdf).name} -> {reason}{where}")


def friendly_error(e):
    """Short, readable reason for a failed file."""
    name = type(e).__name__
    if isinstance(e, PermissionError):
        return "Excel file is open or locked - close it and run again"
    if name in ("CalledProcessError", "FileDataError", "EmptyFileError", "PDFPageCountError"):
        return "PDF cannot be opened (damaged, empty or not a real PDF)"
    return str(e) if isinstance(e, RuntimeError) else f"{name}: {e}"


def error_reason(errors):
    """Collapse '<field>: not found' lines into one short sentence."""
    missing = [m.split(":")[0] for m in errors if m.endswith("not found")]
    other = [m for m in errors if not m.endswith("not found")]
    parts = []
    if missing:
        shown = ", ".join(missing[:5]) + (" ..." if len(missing) > 5 else "")
        parts.append(f"{len(missing)} required field(s) not found ({shown})"
                     + (" - is this a shipment bundle?" if len(missing) > 8 else ""))
    return "; ".join(parts + other)


def process(pdf, xlsx, out=None, extras=False, force=False, dry=False, debug_dir=None, use_gemini=False):
    res, docs, issues = analyze(pdf, use_gemini, debug_dir)
    row = None
    errors = [m for lvl, m in issues if lvl == "ERROR"]
    if errors:
        reason = error_reason(errors)
        gem = [m for lvl, m in issues if lvl == "WARN" and m.startswith("Gemini")]
        if gem:
            reason += " | " + " | ".join(gem[:2])      # e.g. wrong API key / model retired / quota
        record_skip(pdf, reason, save=not dry, res=res, issues=issues)
    elif not dry:
        out = out or xlsx
        if Path(out).resolve() == Path(xlsx).resolve():
            backup(xlsx)
        row = fill_workbook(xlsx, res, out, extras, force)
    return res, issues, row


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf", nargs="+", help="PDF file(s) or a folder containing PDFs")
    ap.add_argument("--trim-only", action="store_true",
                    help="only cut the PDF: keep customs declaration page 1 + COMMERCIAL INVOICE page(s), "
                         "save as <name>_trimmed.pdf (no Excel needed)")
    ap.add_argument("--xlsx", help="CDC Export Report workbook (required unless --trim-only)")
    ap.add_argument("--out", help="save here instead of updating --xlsx in place (a backup is kept otherwise)")
    ap.add_argument("--extras", action="store_true", help="also fill unmarked columns E (flight date) and J (procedure)")
    ap.add_argument("--force", action="store_true", help="allow duplicate INV.NO / E number")
    ap.add_argument("--dup-e", choices=["group", "skip", "allow"], default=None,
                    help="repeated customs (E) number: group=only under its own group (default), skip=never, allow=no check")
    ap.add_argument("--dry-run", action="store_true", help="extract and check only, do not write Excel")
    ap.add_argument("--gemini", action="store_true",
                    help="OPT-IN: send pages to Google Gemini when OCR misses/disputes a field (needs API key)")
    ap.add_argument("--debug", action="store_true", help="save OCR text to ./ocr_debug")
    a = ap.parse_args()
    global DUP_E_POLICY
    DUP_E_POLICY = a.dup_e or DUP_E_POLICY
    files = []
    for p in a.pdf:
        pp = Path(p)
        files += [f for f in sorted(pp.glob("*.pdf")) if not f.stem.endswith("_trimmed")] if pp.is_dir() else [pp]
    if not a.trim_only and not a.xlsx:
        ap.error("--xlsx is required unless --trim-only is used")
    t_start = datetime.now()
    print(f"START   {t_start:%d/%m/%Y %H:%M:%S}   ({len(files)} file(s))")
    if a.trim_only:
        ok = bad = 0
        for pdf in files:
            print(f"\n=== {pdf}")
            try:
                trim_pdf(pdf, debug_dir="ocr_debug" if a.debug else None)
                ok += 1
            except Exception as e:
                if "Tesseract OCR not found" in str(e):
                    print("STOP:", e)
                    break
                record_skip(pdf, friendly_error(e), save=True)
                bad += 1
        print(f"\nSUMMARY: {ok} trimmed, {bad} skipped")
        return
    src, done, skipped = a.xlsx, 0, 0
    for pdf in files:
        print(f"\n=== {pdf}")
        try:
            res, issues, row = process(pdf, src, a.out, a.extras, a.force, a.dry_run,
                                       "ocr_debug" if a.debug else None, a.gemini)
        except Exception as e:                       # duplicate INV.NO, Excel open, damaged PDF, OCR crash...
            msg = friendly_error(e)
            if "Tesseract OCR not found" in msg:     # setup problem, not a file problem
                print("STOP:", msg)
                break
            record_skip(pdf, msg, save=not a.dry_run)
            skipped += 1
            continue
        recs = group_of(res)
        for n, rec in enumerate(recs, 1):
            if len(recs) > 1:
                print(f"  --- invoice {n}/{len(recs)}")
            for f, c in {**COLUMNS, **EXTRA_COLUMNS}.items():
                v = rec[f]["value"]
                print(f"  {c:>2}  {f:<13} {v.strftime('%d/%b/%y') if isinstance(v, datetime) else v}")
        for lvl, msg in issues:
            print(f"  [{lvl}] {msg}")
        if row:
            done += 1
            rows = res.get("_rows") or [row]
            print(f"  -> written to row {rows[0]}" if len(rows) == 1
                  else f"  -> {len(rows)} invoices written to rows {rows[0]}-{rows[-1]}")
        elif any(l == "ERROR" for l, _ in issues):
            skipped += 1                            # record_skip already printed the SKIPPED line
        else:
            print("  -> checked only, nothing written")
        if a.out and row:
            src = a.out        # next PDF appends to the same output file
    t_end = datetime.now()
    print(f"\nFINISH  {t_end:%d/%m/%Y %H:%M:%S}   (took {(t_end - t_start).total_seconds():.0f} sec)")
    print(f"SUMMARY: {done} written, {skipped} skipped" + (f"  -> see {ERROR_LOG}" if skipped and not a.dry_run else ""))


if __name__ == "__main__":
    main()
