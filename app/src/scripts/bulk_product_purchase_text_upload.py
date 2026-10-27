#!/usr/bin/env python3
"""
Bulk Product Purchase Order Text upload (MM02, Purchase Order Text view)
- API_PRODUCT_SRV / A_ProductPurchaseText, SAP DEV.

Columns: Product, Language, LongText

Each Product + Language pair becomes ONE text. If the same pair appears on
several rows, the rows are joined in order with a line break, so a long text
can be split across rows. Line breaks inside a cell (Alt+Enter) are kept.

Upsert: the script first checks whether the text already exists in SAP. If it
does, it is replaced with PATCH (If-Match falls back to "*" when SAP returns
no ETag); otherwise it is created with POST. Set UPSERT = False to always POST
(SAP then rejects texts that already exist).

    python bulk_product_purchase_text_upload.py --template xlsx
    python bulk_product_purchase_text_upload.py -f purchase_texts.xlsx --dry-run
    python bulk_product_purchase_text_upload.py -f purchase_texts.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sap_common import load_metadata, rows_of, run_cli, service_url, validate_payload  # noqa: E402

TITLE = "Material PO texts"
SLUG = "product_po_text"
SERVICE = "API_PRODUCT_SRV"
ENTITY_SET = "A_ProductPurchaseText"
ENTITY_TYPE = "A_ProductPurchaseTextType"
BASE_URL = service_url(SERVICE)          # DEV: my440263
DOC_FIELD = ""                            # no document number; Group is shown instead
YEAR_FIELD = None
GROUP_AS_DOC = True                      # response has no new number; show Product / Language
TCODE = "MM02"
ICON = "📝"
COLOR = "#0891B2"
GROUP_LABEL = "Product / Language"
ITEMS_LABEL = "Texts"
DESCRIPTION = ("Create or replace the purchase order text (long text) on the material master. "
               "One text per Product + Language; extra rows for the same pair are appended "
               "as new lines.")

UPSERT = True                            # True: create if missing, replace if it exists


def count_items(p: dict) -> int:
    return 1


COLUMNS = ["Product", "Language", "LongText"]
COLUMN_NOTES = {
    "Product": "Required. Material number (max 40).",
    "Language": "Required. 2-letter language code, e.g. EN or DE.",
    "LongText": "Required. The purchase order text. Use Alt+Enter for line breaks, or repeat "
                "Product + Language on more rows to continue the text.",
}
SAMPLE_ROWS = [
    {"Product": "RM019", "Language": "EN",
     "LongText": "Supplier must deliver in original packaging.\n"
                 "Include certificate of analysis with every batch."},
    {"Product": "RM019", "Language": "EN", "LongText": "Deliver to gate 2 between 9:00 and 17:00."},
    {"Product": "FG228", "Language": "EN", "LongText": "Drawing rev. C applies. Tolerance +/- 0.05 mm."},
]


def build_documents(df) -> list[dict]:
    md = load_metadata(SERVICE)
    groups: dict[tuple, dict] = {}
    loose: list[dict] = []

    for row in rows_of(df):
        product = row.get("Product", "").strip()
        language = row.get("Language", "").strip().upper()
        text = row.get("LongText", "").replace("\r\n", "\n").replace("\r", "\n").rstrip()

        if not product and not language and not text.strip():
            continue  # fully blank row

        errs = []
        if not product:
            errs.append("Product is blank")
        if not language:
            errs.append("Language is blank (use a 2-letter code, e.g. EN)")
        elif len(language) != 2 or not language.isalpha():
            errs.append(f"Language '{language}' should be a 2-letter code, e.g. EN or DE")
        if not text.strip():
            errs.append("LongText is blank")

        if not (product and language):
            loose.append({"ref": f"LINE{row['_line']}", "lines": str(row["_line"]),
                          "errors": [f"line {row['_line']}: {e}" for e in errs], "payload": {}})
            continue
        g = groups.setdefault((product, language), {"parts": [], "lines": [], "errors": []})
        if text.strip():
            g["parts"].append(text)
        g["lines"].append(row["_line"])
        g["errors"] += [f"line {row['_line']}: {e}" for e in errs]

    docs = []
    for (product, language), g in groups.items():
        payload = {"Product": product, "Language": language, "LongText": "\n".join(g["parts"])}
        docs.append({"ref": f"{product} / {language}", "lines": ",".join(map(str, g["lines"])),
                     "errors": g["errors"] + validate_payload(md, ENTITY_TYPE, payload),
                     "payload": payload})
    return docs + loose


def _key_path(product: str, language: str) -> str:
    def k(v):
        return quote(str(v).replace("'", "''"), safe="")
    return f"{ENTITY_SET}(Product='{k(product)}',Language='{k(language)}')"


def _ok(resp) -> bool:
    return 200 <= resp.status_code < 300


def post_document(client, payload: dict):
    """Upsert: PATCH the text if it exists, otherwise POST it. Returns (response, note)."""
    if UPSERT:
        path = _key_path(payload["Product"], payload["Language"])
        check = client.get(path)
        if check.status_code == 200:
            etag = check.headers.get("ETag")
            if not etag:
                try:
                    etag = check.json().get("d", {}).get("__metadata", {}).get("etag")
                except ValueError:
                    etag = None
            resp = client.patch(path, {"LongText": payload["LongText"]}, etag or "*")
            return resp, ("Updated (text already existed)" if _ok(resp) else "")
    resp = client.post(ENTITY_SET, payload)
    return resp, ("Created" if _ok(resp) else "")


if __name__ == "__main__":
    run_cli(sys.modules[__name__])