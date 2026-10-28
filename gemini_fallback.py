"""
Optional Gemini fallback (OPT-IN). Only used with --gemini or the GUI checkbox,
and only when OCR left a field empty or the documents disagreed.
Pages are sent to Google's Gemini API - check company policy before enabling.

Setup:  pip install google-genai
        set GEMINI_API_KEY=your_key      (or put the key alone in gemini_key.txt next to this file)
        optional: set GEMINI_MODEL=<model name>   (default below; model names change over time)
"""
import io
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
MAX_SIDE = 2000

HELP = {
    "inv_no": "commercial invoice number, e.g. C266151",
    "customer": "short buyer/brand name from the consignee, e.g. ADIDAS",
    "customs_no": "customs declaration (manifest) registration no, letter E + digits, e.g. E55338",
    "customs_date": "customs declaration registration date, format DD/MM/YYYY",
    "hts": "10-digit HTS code printed on the invoice/packing list",
    "description": "garment description, e.g. UNISEX T-SHIRT (no HTS code, no fabric)",
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


def _clean(field, v):
    """Convert Gemini's JSON value to the same type the OCR parser produces."""
    if v in (None, "", "null"):
        return None
    s = str(v).strip()
    try:
        if field == "customs_date":
            return datetime.strptime(s, "%d/%m/%Y")
        if field in ("qty",):
            return int(float(s.replace(",", "")))
        if field in ("amount", "rate", "nw", "gw"):
            return float(s.replace(",", ""))
        if field == "hts":
            d = re.sub(r"\D", "", s)
            return d if len(d) == 10 else None
        if field == "customs_no":
            return re.sub(r"\s", "", s).upper()
        if field in ("inv_no", "description", "fabric", "customer"):
            return s.upper()
        return s
    except ValueError:
        return None


def ask(images, fields):
    """images: PIL pages (SAD, invoice, packing list...). fields: list of field keys wanted."""
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key())
    spec = "\n".join(f'- "{f}": {HELP[f]}' for f in fields)
    prompt = (
        "These are pages of one export shipment bundle from a garment factory in Cambodia "
        "(customs declaration, commercial invoice, packing list). Red pen marks may appear; ignore them "
        "and read the printed values. Return ONE JSON object with exactly these keys, using null when a "
        "value is not clearly readable (never guess):\n" + spec
    )
    parts = [types.Part.from_bytes(data=_jpeg(i), mime_type="image/jpeg") for i in images] + [prompt]
    resp = client.models.generate_content(
        model=MODEL, contents=parts,
        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0))
    data = json.loads(resp.text)
    return {f: _clean(f, data.get(f)) for f in fields}
