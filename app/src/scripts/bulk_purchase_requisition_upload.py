#!/usr/bin/env python3
"""
Bulk Purchase Requisition upload (ME51N) - API_PURCHASEREQUISITION_2 (OData V4)
/ PurchaseReqn with _PurchaseRequisitionItem, SAP DEV.

One row = one PR item. Rows with the same PR_Ref form one purchase requisition
(header fields are taken from the first row of each PR_Ref). Item numbers are
optional: leave PurchaseRequisitionItem blank and SAP numbers the items 10, 20 ...

This is an OData V4 service, so the payload differs from the V2 scripts:
dates are plain "YYYY-MM-DD", quantities and prices are JSON numbers, and the
items are a plain list (no {"results": [...]}). The PR number is read from the
V4 response body (no "d" wrapper).

Not covered: account assignment (cost center etc.) - that needs the separate
account-assignment entity. Use items without AccountAssignmentCategory.

    python bulk_purchase_requisition_upload.py --template xlsx
    python bulk_purchase_requisition_upload.py -f purchase_requisitions.xlsx --dry-run
    python bulk_purchase_requisition_upload.py -f purchase_requisitions.xlsx
"""
from __future__ import annotations

import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sap_common import DEV_HOST, parse_date, rows_of, run_cli, sap_bool  # noqa: E402

TITLE = "Purchase requisitions"
SLUG = "purchase_requisition"
SERVICE = "API_PURCHASEREQUISITION_2"
ENTITY_SET = "PurchaseReqn"
BASE_URL = (f"{DEV_HOST}/sap/opu/odata4/sap/api_purchaserequisition_2/srvd_a2x/sap/"
            f"purchaserequisition/0001")       # DEV: my440263 (OData V4)
DOC_FIELD = "PurchaseRequisition"
YEAR_FIELD = None
TCODE = "ME51N"
ICON = "📋"
COLOR = "#4F46E5"
GROUP_LABEL = "PR_Ref"
ITEMS_LABEL = "Items"
DESCRIPTION = ("Create purchase requisitions with one or more items (OData V4). One row per "
               "item, grouped by PR_Ref; SAP numbers the items when PurchaseRequisitionItem "
               "is blank.")

ITEM_NAV = "_PurchaseRequisitionItem"


def count_items(p: dict) -> int:
    return len(p.get(ITEM_NAV, []))


# ---------------------------------------------------------------------------
# Columns. Field -> (kind, max length). kind: str | date | dec | bool | item
# ---------------------------------------------------------------------------
HEADER_FIELDS = {
    "PurchaseRequisitionType": ("str", 4),
    "PurReqnDescription": ("str", 40),
    "PurReqnHeaderNote": ("str", None),
    "SourceDetermination": ("bool", None),
}
ITEM_FIELDS = {
    "PurchaseRequisitionItem": ("item", 5),
    "PurchaseRequisitionItemText": ("str", 40),
    "Material": ("str", 40),
    "MaterialGroup": ("str", 9),
    "RequestedQuantity": ("dec", None),
    "BaseUnit": ("str", 3),
    "BaseUnitISOCode": ("str", 3),
    "Plant": ("str", 4),
    "StorageLocation": ("str", 4),
    "PurchasingGroup": ("str", 3),
    "PurchasingOrganization": ("str", 4),
    "DeliveryDate": ("date", None),
    "PurchaseRequisitionPrice": ("dec", None),
    "PurReqnPriceQuantity": ("dec", None),
    "PurReqnItemCurrency": ("str", 5),
    "FixedSupplier": ("str", 10),
    "RequisitionerName": ("str", 12),
}
HEADER_DEFAULTS = {"SourceDetermination": False, "PurReqnDoOnlyValidation": False}

COLUMNS = ["PR_Ref"] + list(HEADER_FIELDS) + list(ITEM_FIELDS)

COLUMN_NOTES = {
    "PR_Ref": "Required. Your own key - rows with the same value form one purchase requisition.",
    "PurchaseRequisitionType": "Required (first row of the PR). Document type, e.g. ZDOM or NB.",
    "PurReqnDescription": "Optional header description (max 40).",
    "PurReqnHeaderNote": "Optional header note (free text).",
    "SourceDetermination": "Optional. TRUE = let SAP determine the source of supply. "
                           "Default FALSE.",
    "PurchaseRequisitionItem": "Optional item number, e.g. 10. Leave blank and SAP numbers "
                               "the items.",
    "PurchaseRequisitionItemText": "Short text (max 40). Required when Material is blank.",
    "Material": "Material number. Leave blank for a text item (then MaterialGroup is needed).",
    "MaterialGroup": "Material group. Required for text items without Material.",
    "RequestedQuantity": "Required. Quantity, e.g. 10.",
    "BaseUnit": "Required (or BaseUnitISOCode). Unit, e.g. KG, PC.",
    "BaseUnitISOCode": "Optional ISO unit, e.g. KGM for KG, PCE for PC.",
    "Plant": "Required.",
    "StorageLocation": "Optional.",
    "PurchasingGroup": "Required, e.g. 002.",
    "PurchasingOrganization": "Optional. Leave blank to let SAP determine it.",
    "DeliveryDate": "Required. YYYY-MM-DD or DD.MM.YYYY.",
    "PurchaseRequisitionPrice": "Optional valuation price (otherwise taken from the material).",
    "PurReqnPriceQuantity": "Optional price unit, e.g. 1.",
    "PurReqnItemCurrency": "Optional currency of the price, e.g. INR.",
    "FixedSupplier": "Optional fixed supplier number.",
    "RequisitionerName": "Optional requisitioner (max 12).",
}

SAMPLE_ROWS = [
    {"PR_Ref": "PR1", "PurchaseRequisitionType": "ZDOM", "PurReqnDescription": "Bulk PR Test",
     "PurReqnHeaderNote": "Created using the bulk upload", "SourceDetermination": "FALSE",
     "PurchaseRequisitionItemText": "Test Material TG10-1", "Material": "RM019",
     "RequestedQuantity": "10", "BaseUnit": "KG", "BaseUnitISOCode": "KGM", "Plant": "01TG",
     "PurchasingGroup": "002", "DeliveryDate": "2026-11-15"},
    {"PR_Ref": "PR1", "PurchaseRequisitionItemText": "Test Material TG10-2", "Material": "RM019",
     "RequestedQuantity": "10", "BaseUnit": "KG", "BaseUnitISOCode": "KGM", "Plant": "01TG",
     "PurchasingGroup": "002", "DeliveryDate": "2026-11-15"},
]


# ---------------------------------------------------------------------------
# Conversion (V4 JSON types)
# ---------------------------------------------------------------------------
def _number(v: str):
    s = v.replace(" ", "")
    s = s.replace(",", ".") if ("," in s and "." not in s) else s.replace(",", "")
    try:
        d = Decimal(s)
    except InvalidOperation:
        raise ValueError(f"'{v}' is not a number")
    return int(d) if d == d.to_integral_value() else float(d)


def _convert(field: str, kind: str, maxlen, value: str):
    if kind == "date":
        return parse_date(value).isoformat()
    if kind == "bool":
        return sap_bool(value)
    if kind == "dec":
        return _number(value)
    if kind == "item":
        if not value.isdigit():
            raise ValueError(f"'{value}' must be a number, e.g. 10")
        value = str(int(value))
    if maxlen and len(value) > maxlen:
        raise ValueError(f"'{value}' is longer than {maxlen} characters")
    return value


def _build(row: dict, fields: dict, errors: list[str]) -> dict:
    out = {}
    for f, (kind, maxlen) in fields.items():
        v = row.get(f, "").strip()
        if not v:
            continue
        try:
            out[f] = _convert(f, kind, maxlen, v)
        except ValueError as e:
            errors.append(f"{f}: {e}")
    return out


# ---------------------------------------------------------------------------
# Build documents
# ---------------------------------------------------------------------------
def build_documents(df) -> list[dict]:
    prs: dict[str, dict] = {}

    for row in rows_of(df):
        errs: list[str] = []
        ref = row.get("PR_Ref", "").strip() or f"LINE{row['_line']}"
        if ref not in prs:
            if not row.get("PurchaseRequisitionType"):
                errs.append(f"PurchaseRequisitionType is required on the first row of {ref}")
            header = {**HEADER_DEFAULTS, **_build(row, HEADER_FIELDS, errs)}
            header[ITEM_NAV] = []
            prs[ref] = {"ref": ref, "lines": [], "errors": [], "payload": header}
        pr = prs[ref]

        for f in ("RequestedQuantity", "Plant", "PurchasingGroup", "DeliveryDate"):
            if not row.get(f):
                errs.append(f"{f} is required")
        if not row.get("BaseUnit") and not row.get("BaseUnitISOCode"):
            errs.append("BaseUnit (or BaseUnitISOCode) is required")
        if not row.get("Material"):
            if not row.get("PurchaseRequisitionItemText"):
                errs.append("fill Material, or PurchaseRequisitionItemText for a text item")
            if not row.get("MaterialGroup"):
                errs.append("MaterialGroup is required for a text item without Material")
        item = _build(row, ITEM_FIELDS, errs)
        if isinstance(item.get("RequestedQuantity"), (int, float)) and item["RequestedQuantity"] <= 0:
            errs.append("RequestedQuantity must be greater than 0")
        pr["payload"][ITEM_NAV].append(item)
        pr["lines"].append(row["_line"])
        pr["errors"] += [f"line {row['_line']}: {e}" for e in errs]

    docs = []
    for pr in prs.values():
        numbers = [i["PurchaseRequisitionItem"] for i in pr["payload"][ITEM_NAV]
                   if "PurchaseRequisitionItem" in i]
        if numbers and len(numbers) != len(pr["payload"][ITEM_NAV]):
            pr["errors"].append("fill PurchaseRequisitionItem on every row of this PR, or on none")
        if len(set(numbers)) != len(numbers):
            pr["errors"].append("PurchaseRequisitionItem numbers repeat within this PR")
        docs.append({"ref": pr["ref"], "lines": ",".join(map(str, pr["lines"])),
                     "errors": pr["errors"], "payload": pr["payload"]})
    return docs


# ---------------------------------------------------------------------------
# Post (OData V4: number comes back in the body, not under "d")
# ---------------------------------------------------------------------------
class _V4Response:
    """Wraps the V4 response so the shared poster finds the PR number under "d"."""

    def __init__(self, resp):
        self._resp = resp
        self.status_code = resp.status_code
        self.headers = resp.headers
        self.text = resp.text

    def json(self):
        body = self._resp.json()
        return body if "d" in body else {"d": body}


def post_document(client, payload: dict):
    """POST one PR. Returns (response, note) like the other scripts' hooks."""
    resp = client.post(ENTITY_SET, payload)
    if 200 <= resp.status_code < 300:
        try:
            items = len(resp.json().get(ITEM_NAV, []))
            note = f"Created with {items} item(s)" if items else "Created"
        except ValueError:
            note = "Created"
        return _V4Response(resp), note
    return resp, ""                      # errors: sap_common.extract_error reads V2 and V4


if __name__ == "__main__":
    run_cli(sys.modules[__name__])
