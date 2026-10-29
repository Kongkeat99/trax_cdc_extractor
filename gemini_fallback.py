"""
Optional Gemini fallback (OPT-IN). Only used with --gemini or the GUI checkbox,
and only when OCR left a field empty or the documents disagreed.
Pages are sent to Google's Gemini API - check company policy before enabling.

Setup:  pip install google-genai
        set GEMINI_API_KEY=your_key      (or put the key alone in gemini_key.txt next to this file)
        optional: set GEMINI_MODEL=<model name>   (tried first; model names change over time)
If a model has been retired (404 / NOT_FOUND) the next one in MODEL_FALLBACKS is tried automatically.
"""
import io
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")
MODEL_FALLBACKS = ["gemini-3.6-flash", "gemini-2.5-flash"]   # tried in order if MODEL is gone
MAX_SIDE = 2000
RETRIES = 3                                                  # per model, for 429 / 5xx / timeouts

HELP = {
    "inv_no": "commercial invoice number, e.g. C266151 or SC260461-1 (keep any -1, -2 suffix)",
    "customer": "short buyer/brand name from the consignee, e.g. ADIDAS",
    "customs_no": "customs declaration (manifest) registration no, letter E + digits, e.g. E55338",
    "customs_date": "customs declaration registration date, format DD/MM/YYYY",
    "hts": "HTS code(s) exactly as printed on the invoice/packing list (6 to 10 digits); if several, "
           "join with '/' in order of appearance, e.g. 6104630000/6109902000 or 610990/610463",
    "description": "garment description, e.g. UNISEX T-SHIRT (no HTS code, no fabric); if several "
                   "different items, join with '/' in order of appearance",
    "qty": "total quantity in pieces (integer)",
    "amount": "total invoice amount in USD (number)",
    "destination": "country of destination in English, e.g. United Kingdom",
    "rate": "exchange rate KHR per USD from the customs declaration (box 23)",
    "nw": "total net weight in kg (number)",
    "gw": "total gross weight in kg (number)",
    "fabric": "KNITTED or WOVEN",
}


def api_key():
    k = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    app_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    f = app_dir / "gemini_key.txt"
    if not k and f.exists():
        k = f.read_text().strip()
    if not k:
        raise RuntimeError("Gemini API key not found. Set GEMINI_API_KEY or create gemini_key.txt.")
    return k


def _jpeg(img):
    img = img.copy()
    img.thumbnail((MAX_SIDE, MAX_SIDE))
    buf = io.BytesIO()
    img.convert("RGB").save(buf, "JPEG", quality=85)
    return buf.getvalue()


def _codes(s, lo=6, hi=10):
    """'6104630000 / 6109902000' -> '6104630000/6109902000'. None unless EVERY part is 6-10 digits."""
    out = []
    for part in re.split(r"[/,;|\n]+", s):
        d = re.sub(r"\D", "", part)
        if not d:
            continue
        if not lo <= len(d) <= hi:
            return None
        if d not in out:
            out.append(d)
    return "/".join(out) or None


def _clean(field, v):
    """Convert Gemini's JSON value to the same type the OCR parser produces
    (None = unreadable / not valid, so a wrong guess never reaches the Excel sheet)."""
    if v in (None, "", "null"):
        return None
    s = str(v).strip()
    try:
        if field in ("customs_date", "onboard_date"):
            return datetime.strptime(s, "%d/%m/%Y")
        if field == "procedure":
            n = int(s)
            return n if n > 0 else None
        if field == "qty":
            n = int(float(s.replace(",", "")))
            return n if n > 0 else None
        if field in ("amount", "rate", "nw", "gw"):
            x = float(s.replace(",", ""))
            return x if x > 0 else None
        if field == "hts":
            return _codes(s)
        if field == "customs_no":
            e = re.sub(r"\s", "", s).upper()
            return e if re.fullmatch(r"E\d{4,6}", e) else None
        if field == "inv_no":
            i = re.sub(r"\s", "", s).upper()
            return i if re.fullmatch(r"[A-Z]{1,2}\d{6}(-\d{1,2})?", i) else None
        if field == "fabric":
            u = s.upper()
            return "KNITTED" if u.startswith("KNIT") else "WOVEN" if u.startswith("WOVEN") else None
        if field == "description":
            return re.sub(r"\s*/\s*", "/", re.sub(r"\s+", " ", s)).upper()
        if field in ("customer",):
            return s.upper()
        return s
    except ValueError:
        return None


def _json(text):
    """Gemini normally returns clean JSON; tolerate ```json fences, a list wrapper or empty text."""
    if not text:
        raise ValueError("Gemini returned an empty answer")
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip())
    d = json.loads(t)
    if isinstance(d, list):
        d = next((x for x in d if isinstance(x, dict)), {})
    if not isinstance(d, dict):
        raise ValueError("Gemini answer is not a JSON object")
    return d


def _status(e):
    for a in ("code", "status_code"):
        c = getattr(e, a, None)
        if isinstance(c, int):
            return c
    m = re.match(r"\s*(\d{3})\b", str(e))
    return int(m[1]) if m else None


def _generate(client, types, model, parts):
    last = None
    for attempt in range(RETRIES):
        try:
            return client.models.generate_content(
                model=model, contents=parts,
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0))
        except Exception as e:
            last, c = e, _status(e)
            net = isinstance(e, (TimeoutError, ConnectionError)) or type(e).__module__.startswith(("httpx", "httpcore"))
            if c in (429, 500, 502, 503, 504) or net:
                if attempt < RETRIES - 1:
                    time.sleep(2 * 2 ** attempt)         # 2s, 4s
                continue
            raise
    raise last


def ask(images, fields, hint=None):
    """images: PIL pages of ONE invoice (SAD, invoice, packing list...). fields: field keys wanted.
    hint: invoice number to focus on when the pages also contain other invoices."""
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key())
    spec = "\n".join(f'- "{f}": {HELP[f]}' for f in fields)
    focus = (f"Read ONLY the commercial invoice (and its packing list) numbered {hint}; ignore other invoices. "
             if hint else "")
    prompt = (
        "These are pages of one export shipment bundle from a garment factory in Cambodia "
        "(customs declaration, commercial invoice, packing list). Red pen marks may appear; ignore them "
        "and read the printed values. " + focus + "Return ONE JSON object with exactly these keys, "
        "using null when a value is not clearly readable (never guess):\n" + spec
    )
    parts = [types.Part.from_bytes(data=_jpeg(i), mime_type="image/jpeg") for i in images] + [prompt]
    err = None
    for model in dict.fromkeys([MODEL] + MODEL_FALLBACKS):          # unique, in order
        try:
            resp = _generate(client, types, model, parts)
            break
        except Exception as e:
            err = e
            if _status(e) == 404 or "NOT_FOUND" in str(e):           # model retired / renamed -> next one
                continue
            raise
    else:
        raise RuntimeError(f"no Gemini model available (last error: {err}). Set GEMINI_MODEL to a current model name.")
    data = _json(resp.text)
    return {f: _clean(f, data.get(f)) for f in fields}
