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
               "HK": "Hong Kong", "AU": "Australia", "CA": "Canada", "KR": "South Korea"}
DUP_E_POLICY = "group"       # "group": same E no. only allowed directly under its own group (several invoices, one declaration)
                             # "skip" : refuse any E no. already in the sheet   |  "allow": no E-number check
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
ERROR_LOG = APP_DIR / "skipped_files.csv"   # every skipped PDF is listed here
ERROR_DIR = APP_DIR / "error_files"         # ...and a COPY of it is put here (None = no copy)
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


def ocr_pdf(pdf, debug_dir=None, keep=None, progress=None):
    """OCR all pages in parallel (one tesseract process per CPU core). Page order is preserved."""
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
    if "COMMERCIAL INVOICE" in t:
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


def extract(pages, extra=None):
    docs = {}
    for t in pages:
        k = classify(t)
        if k and k not in docs:
            docs[k] = t
    C = {k: [] for k in ["inv_no", "customer", "customs_no", "customs_date", "hts", "description",
                         "qty", "amount", "destination", "rate", "nw", "gw", "fabric",
                         "onboard_date", "procedure", "mode"]}
    inv, pl, info, sad = (docs.get(k, "") for k in ("invoice", "pl", "info", "sad"))
    allt = "\n".join(docs.values())

    # 1 INV.NO
    for src, t in (("info", info), ("pl", pl), ("invoice", inv)):
        for p in (r"INV\+PL:\s*([A-Z]\d{6})", r"INV\.?\s*N[O0K]\.?\s*[:\-]?\s*([A-Z]\d{6})"):
            v = first(p, t)
            if v:
                C["inv_no"].append((v.upper(), src))
                break

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
    for src, t in (("pl", pl), ("invoice", inv)):
        v = first(r"HTS\.?\s*CODE\s*[:\-~]?\s*(\d{10})", t)
        if v:
            C["hts"].append((v, src))
    m = re.search(r"\b(6\d{7})\s*\|?\s*(00)\b", sad)
    if m:
        C["hts"].append((m[1] + m[2], "sad"))

    # 6 Description + 13 Fabric
    for src, t in (("pl", pl), ("invoice", inv)):
        v = first(r"^\s*(.+?)\s*\(\s*HTS", t)
        if v:
            C["description"].append((v.strip().upper().replace("|", "").strip(), src))
        v = first(r"\(\s*(KNITTED|WOVEN)\s*\)", t)
        if v:
            C["fabric"].append((v.upper(), src))

    # 7 Qty / 11 NW / 12 GW  (packing list grand total row)
    m = re.search(r"GRAND TOTAL:?\s*([\d,]+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", pl, re.I)
    if m:
        C["qty"].append((int(num(m[1])), "pl"))
        C["nw"].append((num(m[2]), "pl"))
        C["gw"].append((num(m[3]), "pl"))
    for src, t in (("invoice", inv), ("pl", pl)):
        v = first(r"TOTAL N\.?W\.?:?\s*([\d,]+\.\d{2})", t)
        if v:
            C["nw"].append((num(v), src))
        v = first(r"TOTAL G\.?W\.?:?\s*([\d,]+\.\d{2})", t)
        if v:
            C["gw"].append((num(v), src))
    # 8 Total price (+ qty/price cross-check) from the invoice item line
    m = re.search(r"([\d,]{3,})\s+(\d+\.\d{2})\s+([\d,]+\.\d{2})\s*$", inv, re.M)
    unit_price = None
    if m:
        C["qty"].append((int(num(m[1])), "invoice"))
        unit_price = num(m[2])
        C["amount"].append((num(m[3]), "invoice"))
    v = first(r"GRAND TOTAL:?\s*(?:[\d,]+\s+)?([\d,]+\.\d{2})", inv)
    if v:
        C["amount"].append((num(v), "invoice"))

    # SAD: amount / rate / qty / procedure / flight date
    v = first(r"USD\s*[\|\[\(]?\s*([\d,]+\.\d{2})", sad)
    if v:
        C["amount"].append((num(v), "sad"))
    m = re.search(r"(\d)[,.\s]?(\d{3})[.,]\s?(\d{4})\b", sad)     # 10 Exchange rate (tolerant of OCR)
    if m:
        C["rate"].append((float(f"{m[1]}{m[2]}.{m[3]}"), "sad"))
    amt = vote(C["amount"])[0]                              # safety net: customs value (KHR) / USD amount
    if amt:
        for kv in re.findall(r"\b(\d{1,3}(?:,\d{3}){2,3})\b", sad):
            q = num(kv) / amt
            if 3000 < q < 6000 and abs(q - round(q)) < 0.01:
                C["rate"].append((float(round(q)), "sad-derived"))
                break
    v = first(r"NAR\s*([\d,]+)", sad)
    if v:
        C["qty"].append((int(num(v)), "sad"))
    v = first(r"\b(\d{4})\s*\|?\s*000(?!\d)", sad)
    if v:
        C["procedure"].append((int(v), "sad"))
    v = first(r"AK?\w*\s*/\s*(\d{2}\.\d{2}\.\d{4})", sad)
    if v:
        C["onboard_date"].append((datetime.strptime(v, "%d.%m.%Y"), "sad"))

    # 9 Destination
    for src, t in (("invoice", inv), ("pl", pl)):
        m = re.search(r"^BY\s+(AIR|SEA)\s+FROM\b.*\s(\S{2,4})\s*$", t, re.M)
        if m:
            C["destination"].append((COUNTRY_MAP.get(m[2].upper().strip("!|"), m[2].title()), src))
            C["mode"].append((m[1].upper(), src))
    for name in set(COUNTRY_MAP.values()):
        if re.search(rf"Destination.*?\n.*?\b{name}\b", sad, re.S):
            C["destination"].append((name, "sad"))

    for f, v in (extra or {}).items():          # Gemini values: lowest priority voter
        if f in C and v not in (None, ""):
            C[f].append((v, "gemini"))

    res = {}
    for k, cands in C.items():
        best, srcs, distinct = vote(cands)
        res[k] = {"value": best, "sources": srcs, "distinct": distinct}
    res["_unit_price"] = unit_price
    return res, sorted(docs)


def validate(res, docs):
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
        elif k in COLUMNS and d["sources"] == ["gemini"]:
            issues.append(("WARN", f"{k}: value comes from Gemini only - verify against the document"))
        elif k in COLUMNS and len(d["sources"]) == 1:
            issues.append(("INFO", f"{k}: confirmed by one document only ({d['sources'][0]})"))
    g = lambda k: res[k]["value"]
    if g("inv_no") and not re.fullmatch(r"[A-Z]\d{6}", g("inv_no")):
        issues.append(("WARN", f"inv_no format looks odd: {g('inv_no')}"))
    if len(res["mode"]["distinct"]) > 1:
        issues.append(("WARN", "invoice and packing list disagree on transport mode "
                       f"{res['mode']['distinct']} (AWB/SAD show AIR) - check source documents"))
    if g("hts") and not g("hts").startswith("61"):
        issues.append(("WARN", f"HTS {g('hts')} unexpected for knit garments"))
    if g("qty") and res["_unit_price"] and g("amount"):
        if abs(g("qty") * res["_unit_price"] - g("amount")) > 0.5:
            issues.append(("WARN", f"qty x unit price != amount ({g('qty')} x {res['_unit_price']})"))
    if g("nw") and g("gw") and g("nw") > g("gw"):
        issues.append(("ERROR", "net weight is greater than gross weight"))
    return issues


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


def precheck(xlsx, res, force=False):
    """Same duplicate checks as writing, without writing. Returns an error text or None."""
    try:
        ws = openpyxl.load_workbook(xlsx, read_only=False).worksheets[0] if not SHEET_NAME else \
            openpyxl.load_workbook(xlsx)[SHEET_NAME]
        plan_row(ws, res, force)
    except RuntimeError as e:
        return str(e)
    return None


def fill_workbook(xlsx, res, out, extras=False, force=False):
    wb = openpyxl.load_workbook(xlsx)
    ws = wb[SHEET_NAME] if SHEET_NAME else wb.worksheets[0]
    row = plan_row(ws, res, force)
    cols = dict(COLUMNS, **(EXTRA_COLUMNS if extras else {}))
    for field, col in cols.items():
        val = res[field]["value"]
        if val is not None:
            ws[f"{col}{row}"] = val
    fc, ft = FORMULA_COLUMN
    if ws[f"{fc}{row}"].value is None:
        ws[f"{fc}{row}"] = ft.format(r=row)
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return row


def backup(path):
    b = Path(path).parent / "backup"
    b.mkdir(exist_ok=True)
    dst = b / f"{Path(path).stem}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, dst)
    return dst


def analyze(pdf, use_gemini=False, debug_dir=None, log=print, progress=None):
    """OCR -> vote -> (optional Gemini for missing/disputed fields) -> validate."""
    imgs = []
    pages = ocr_pdf(pdf, debug_dir, imgs if use_gemini else None, progress)
    res, docs = extract(pages)
    gem_note = None
    if use_gemini:
        wanted = [f for f in COLUMNS if res[f]["value"] in (None, "") or len(res[f]["distinct"]) > 1]
        if not wanted:
            log("Gemini: not needed, OCR found every field")
        else:
            try:
                import gemini_fallback as gf
                sel, seen = [], set()
                for t, im in zip(pages, imgs):
                    k = classify(t)
                    if k and k not in seen:
                        seen.add(k)
                        sel.append(im)
                log(f"Gemini: sending {len(sel)} pages to Google for fields {wanted}")
                res, docs = extract(pages, gf.ask(sel, wanted))
            except Exception as e:           # no key / no internet / API error: keep OCR result
                gem_note = (f"Gemini skipped ({type(e).__name__}: {e}). "
                            f"OCR could not settle these fields: {wanted}")
    issues = validate(res, docs)
    if gem_note:
        issues.insert(0, ("WARN", gem_note))
    return res, docs, issues


def set_manual(res, field, text):
    """Let the user type a value (GUI double-click). Raises ValueError if it cannot be parsed."""
    import gemini_fallback as gf
    v = gf._clean(field, text)
    if v is None:
        raise ValueError(f"'{text}' is not a valid value for {field}")
    res[field] = {"value": v, "sources": ["manual"], "distinct": [v]}


def record_skip(pdf, reason, log=print, save=True):
    """A PDF could not be processed: show a console message, list it in skipped_files.csv,
    and copy the PDF into error_files/ so it can be checked later. (The original is never moved.)"""
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
        record_skip(pdf, error_reason(errors), save=not dry)
    elif not dry:
        out = out or xlsx
        if Path(out).resolve() == Path(xlsx).resolve():
            backup(xlsx)
        row = fill_workbook(xlsx, res, out, extras, force)
    return res, issues, row


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf", nargs="+", help="PDF file(s) or a folder containing PDFs")
    ap.add_argument("--xlsx", required=True, help="CDC Export Report workbook")
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
        files += sorted(pp.glob("*.pdf")) if pp.is_dir() else [pp]
    t_start = datetime.now()
    print(f"START   {t_start:%d/%m/%Y %H:%M:%S}   ({len(files)} file(s))")
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
        for f, c in {**COLUMNS, **EXTRA_COLUMNS}.items():
            v = res[f]["value"]
            print(f"  {c:>2}  {f:<13} {v.strftime('%d/%b/%y') if isinstance(v, datetime) else v}")
        for lvl, msg in issues:
            print(f"  [{lvl}] {msg}")
        if row:
            done += 1
            print(f"  -> written to row {row}")
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
