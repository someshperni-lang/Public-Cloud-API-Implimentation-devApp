#!/usr/bin/env python3
"""
Bulk Subcontracting Purchase Requisition upload (ME51N, item category L)
- API_PURCHASEREQUISITION_2 (OData V4) / PurchaseReqn, SAP DEV.

Same template layout as the subcontracting PO upload: an Excel workbook with
sheets Header / Items / ScheduleLines / Components, linked by
PR_Ref -> PurchaseRequisitionItem -> ScheduleLine. Every item is sent with
item category L (internal key 3).

How the four sheets reach SAP (the PR API is not the PO API)
------------------------------------------------------------
- Header        -> PurchaseReqn header.
- Items         -> PR items.
- ScheduleLines -> a PR item has ONE delivery date and quantity (no schedule
                   lines in the API). One schedule line gives the item its
                   DeliveryDate and quantity. Several schedule lines make one PR
                   item each (item 10 with lines 1, 2 -> items 10, 11), like
                   entering several delivery dates in ME51N.
- Components    -> NOT sent: the PR API has no component entity, SAP takes the
                   components from the material's BOM. Component rows are
                   linked and checked, and every result says how many were not
                   sent, so nothing is dropped silently. Enter or change
                   components on the subcontracting PO if they differ from the BOM.

Input layouts
-------------
1. The workbook above (--template xlsx).
2. One flat sheet (CSV or Excel), like the subcontracting PO: one row per
   component (or per schedule line / item without components), with PR_Ref,
   item and schedule-line columns repeated. Header / item / schedule-line values
   are read from the first row of each PR / item / schedule line. Rows without
   PurchaseRequisitionItem are one item each (the older one-row-per-item sheets
   still work).

    python bulk_subcontracting_pr_upload.py --template xlsx
    python bulk_subcontracting_pr_upload.py -f Subcontracting_PR_Upload.xlsx --dry-run
    python bulk_subcontracting_pr_upload.py -f Subcontracting_PR_Upload.xlsx
"""
from __future__ import annotations

import io
import re
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import bulk_purchase_requisition_upload as pr  # noqa: E402
from sap_common import clean, read_table, run_cli  # noqa: E402

TITLE = "Subcontracting purchase requisitions"
SLUG = "subcontracting_pr"
SERVICE = pr.SERVICE
ENTITY_SET = pr.ENTITY_SET
BASE_URL = pr.BASE_URL                    # DEV: my440263 (OData V4)
DOC_FIELD = pr.DOC_FIELD
YEAR_FIELD = None
TCODE = "ME51N"
ICON = "🏭"
COLOR = "#7C3AED"
GROUP_LABEL = "PR_Ref"
ITEMS_LABEL = "Items"
DESCRIPTION = ("Subcontracting purchase requisitions (item category L, OData V4) from the same "
               "Header / Items / ScheduleLines / Components workbook as the subcontracting PO. "
               "Several schedule lines become one PR item each; components come from the BOM "
               "(component rows are checked but cannot be sent with a PR).")

count_items = pr.count_items

# ---------------------------------------------------------------------------
# Columns
# ---------------------------------------------------------------------------
HEADER_COLUMNS = list(pr.HEADER_FIELDS)
ITEM_COLUMNS = ["PurchaseRequisitionItem", "PurchaseRequisitionItemText", "Material",
                "RequestedQuantity", "BaseUnit", "BaseUnitISOCode", "Plant", "StorageLocation",
                "PurchasingGroup", "PurchasingOrganization", "Subcontractor", "FixedSupplier",
                "Supplier", "PurchasingInfoRecord", "PurchaseContract", "PurchaseContractItem",
                "PurchaseRequisitionPrice", "PurReqnPriceQuantity", "PurReqnItemCurrency",
                "TaxCode", "Batch", "RequisitionerName", "RequirementTracking",
                "GoodsReceiptIsExpected", "InvoiceIsExpected", "ItemText", "ItemTextLanguage"]
# flat names (same as the subcontracting PO flat sheet)
SCHED_COLUMNS = ["ScheduleLine", "DeliveryDate", "ScheduleLineQuantity"]
COMP_COLUMNS = ["ComponentReservationItem", "ComponentMaterial", "ComponentRequirementDate",
                "ComponentQuantity", "ComponentUnit", "ComponentPlant"]
COLUMNS = ["PR_Ref"] + HEADER_COLUMNS + ITEM_COLUMNS + SCHED_COLUMNS + COMP_COLUMNS

SHEETS = {
    "Header": ["PR_Ref"] + HEADER_COLUMNS,
    "Items": ["PR_Ref"] + ITEM_COLUMNS,
    "ScheduleLines": ["PR_Ref", "PurchaseRequisitionItem", "ScheduleLine",
                      "ScheduleLineDeliveryDate", "ScheduleLineOrderQuantity"],
    "Components": ["PR_Ref", "PurchaseRequisitionItem", "ScheduleLine", "ReservationItem",
                   "Material", "RequirementDate", "QuantityInEntryUnit", "EntryUnit", "Plant"],
}
SCHED_SHEET_TO_FLAT = {"ScheduleLine": "ScheduleLine", "ScheduleLineDeliveryDate": "DeliveryDate",
                       "ScheduleLineOrderQuantity": "ScheduleLineQuantity"}
COMP_SHEET_TO_FLAT = {"ReservationItem": "ComponentReservationItem",
                      "Material": "ComponentMaterial", "RequirementDate": "ComponentRequirementDate",
                      "QuantityInEntryUnit": "ComponentQuantity", "EntryUnit": "ComponentUnit",
                      "Plant": "ComponentPlant"}

COLUMN_NOTES = {c: pr.COLUMN_NOTES.get(c, "") for c in COLUMNS}
COLUMN_NOTES.update({
    "PurchaseRequisitionType": "Required (first row of the PR), e.g. ZDOM or your "
                               "subcontracting PR type.",
    "PurchaseRequisitionItem": "Item number, e.g. 10 - links Items, ScheduleLines and Components. "
                               "Blank in a flat sheet = one item per row, numbered by SAP.",
    "Material": "Required. The finished / semi-finished material the subcontractor makes. "
                "It needs a BOM in the plant (SAP takes the components from it).",
    "RequestedQuantity": "Item quantity. Optional when the schedule lines give the quantities; "
                         "if filled, it must equal their total.",
    "BaseUnit": "Required (or BaseUnitISOCode), e.g. PC.",
    "Plant": "Required. Receiving plant.",
    "Subcontractor": "Recommended. The subcontractor (supplier) who gets the components.",
    "FixedSupplier": "Optional fixed supplier, usually the same as Subcontractor.",
    "PurchasingInfoRecord": "Optional subcontracting info record (price for the work).",
    "PurchaseRequisitionPrice": "Optional valuation price for the subcontracting work.",
    "ScheduleLine": "Schedule line, e.g. 1. Several lines on one item = one PR item per line.",
    "DeliveryDate": "Required. Delivery date of the schedule line.",
    "ScheduleLineQuantity": "Quantity of the schedule line. Blank = RequestedQuantity "
                            "(only with one schedule line).",
    "ScheduleLineDeliveryDate": "Required. Delivery date of the schedule line.",
    "ScheduleLineOrderQuantity": "Quantity of the schedule line. Blank = the item's "
                                 "RequestedQuantity (only with one schedule line).",
    "ComponentReservationItem": "Component line number, e.g. 1. Reference only.",
    "ComponentMaterial": "Component material. Reference only - not sent, SAP uses the BOM.",
    "ComponentRequirementDate": "Component requirement date. Reference only.",
    "ComponentQuantity": "Component quantity. Reference only.",
    "ComponentUnit": "Component unit, e.g. KG. Reference only.",
    "ComponentPlant": "Component plant. Reference only.",
    "ReservationItem": "Component line number, e.g. 1. Reference only.",
    "RequirementDate": "Component requirement date. Reference only.",
    "QuantityInEntryUnit": "Component quantity. Reference only.",
    "EntryUnit": "Component unit, e.g. KG. Reference only.",
})

SAMPLE_ROWS = [      # flat: one row per component
    {"PR_Ref": "SPR1", "PurchaseRequisitionType": "ZDOM", "PurReqnDescription": "Subcontracting",
     "PurchaseRequisitionItem": "10", "PurchaseRequisitionItemText": "Machining of FG228",
     "Material": "FG228", "RequestedQuantity": "100", "BaseUnit": "PC", "Plant": "01TG",
     "PurchasingGroup": "002", "Subcontractor": "18300010", "ScheduleLine": "1",
     "DeliveryDate": "2026-12-15", "ScheduleLineQuantity": "60",
     "ComponentReservationItem": "1", "ComponentMaterial": "RM019",
     "ComponentRequirementDate": "2026-12-10", "ComponentQuantity": "60",
     "ComponentUnit": "KG", "ComponentPlant": "01TG"},
    {"PR_Ref": "SPR1", "PurchaseRequisitionItem": "10", "ScheduleLine": "2",
     "DeliveryDate": "2026-12-30", "ScheduleLineQuantity": "40",
     "ComponentReservationItem": "1", "ComponentMaterial": "RM019",
     "ComponentRequirementDate": "2026-12-24", "ComponentQuantity": "40",
     "ComponentUnit": "KG", "ComponentPlant": "01TG"},
    {"PR_Ref": "SPR1", "PurchaseRequisitionItem": "20", "Material": "FG229",
     "RequestedQuantity": "50", "BaseUnit": "PC", "Plant": "01TG", "PurchasingGroup": "002",
     "Subcontractor": "18300010", "ItemText": "Return the scrap with the finished goods.",
     "ScheduleLine": "1", "DeliveryDate": "2026-12-20", "ScheduleLineQuantity": "50",
     "ComponentReservationItem": "1", "ComponentMaterial": "RM020",
     "ComponentRequirementDate": "2026-12-15", "ComponentQuantity": "25",
     "ComponentUnit": "KG", "ComponentPlant": "01TG"},
]


# ---------------------------------------------------------------------------
# 4-sheet workbook <-> flat rows (same approach as the subcontracting PO)
# ---------------------------------------------------------------------------
def _num(v) -> str:
    v = str(v or "").strip()
    return str(int(v)) if v.isdigit() else v


def _sheet(xl, name) -> list[dict]:
    if name not in xl.sheet_names:
        return []
    df = pd.read_excel(xl, sheet_name=name, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip() for c in df.columns]
    rows = [{k: clean(v) for k, v in r.items()} for r in df.to_dict(orient="records")]
    return [r for r in rows if any(r.values())]


def workbook_to_flat(raw: bytes) -> pd.DataFrame:
    """Header/Items/ScheduleLines/Components -> one row per component."""
    xl = pd.ExcelFile(io.BytesIO(raw))
    headers, items = _sheet(xl, "Header"), _sheet(xl, "Items")
    scheds, comps = _sheet(xl, "ScheduleLines"), _sheet(xl, "Components")

    def key_item(r):
        return (r.get("PR_Ref", ""), _num(r.get("PurchaseRequisitionItem")))

    def key_sched(r):
        return key_item(r) + (_num(r.get("ScheduleLine")) or "1",)

    out = []
    for h in headers:
        hrow = {c: h.get(c, "") for c in SHEETS["Header"]}
        first = True
        pr_items = [i for i in items if i.get("PR_Ref") == h.get("PR_Ref")]
        if not pr_items:
            out.append(hrow)
            continue
        for it in pr_items:
            irow = {**(hrow if first else {"PR_Ref": hrow["PR_Ref"]}),
                    **{c: it.get(c, "") for c in ITEM_COLUMNS}}
            first = False
            it_scheds = [s for s in scheds if key_item(s) == key_item(it)]
            if not it_scheds:
                out.append(irow)
                continue
            for s in it_scheds:
                srow = {**irow, **{flat: s.get(col, "") for col, flat in SCHED_SHEET_TO_FLAT.items()}}
                s_comps = [c for c in comps if key_sched(c) == key_sched(s)]
                if not s_comps:
                    out.append(srow)
                for c in s_comps:
                    out.append({**srow, **{flat: c.get(col, "")
                                           for col, flat in COMP_SHEET_TO_FLAT.items()}})
    # rows pointing at a PR / item / schedule line that does not exist: keep them as
    # rows, so build_documents reports them (missing header / item data) instead of
    # dropping them
    known_pr = {h.get("PR_Ref") for h in headers}
    known_item = {key_item(i) for i in items}
    known_sched = {key_sched(s) for s in scheds}
    for it in items:
        if it.get("PR_Ref") not in known_pr:
            out.append({"PR_Ref": it.get("PR_Ref", ""), **{c: it.get(c, "") for c in ITEM_COLUMNS}})
    for s in scheds:
        if key_item(s) not in known_item:
            out.append({"PR_Ref": s.get("PR_Ref", ""),
                        "PurchaseRequisitionItem": s.get("PurchaseRequisitionItem", ""),
                        **{flat: s.get(col, "") for col, flat in SCHED_SHEET_TO_FLAT.items()}})
    for c in comps:
        if key_sched(c) not in known_sched:
            out.append({"PR_Ref": c.get("PR_Ref", ""),
                        "PurchaseRequisitionItem": c.get("PurchaseRequisitionItem", ""),
                        "ScheduleLine": c.get("ScheduleLine", ""),
                        **{flat: c.get(col, "") for col, flat in COMP_SHEET_TO_FLAT.items()}})
    return pd.DataFrame(out, columns=COLUMNS).fillna("")


def read_upload(source, filename: str) -> pd.DataFrame:
    """Accept the 4-sheet workbook or a flat CSV / Excel sheet."""
    if filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raw = source if isinstance(source, (bytes, bytearray)) else Path(source).read_bytes()
        names = pd.ExcelFile(io.BytesIO(raw)).sheet_names
        if "Header" in names and "Items" in names:
            return workbook_to_flat(raw)
        return read_table(raw, filename)
    return read_table(source, filename)


def template_xlsx_bytes() -> bytes:
    """4-sheet workbook template - same layout and style as the subcontracting PO one."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    sample = {name: [] for name in SHEETS}
    seen_h, seen_i, seen_s = set(), set(), set()
    for r in SAMPLE_ROWS:
        ki = (r["PR_Ref"], r.get("PurchaseRequisitionItem", ""))
        ks = ki + (r.get("ScheduleLine", ""),)
        if r["PR_Ref"] not in seen_h:
            seen_h.add(r["PR_Ref"])
            sample["Header"].append({c: r.get(c, "") for c in SHEETS["Header"]})
        if ki not in seen_i:
            seen_i.add(ki)
            sample["Items"].append({c: r.get(c, "") for c in SHEETS["Items"]})
        if ks not in seen_s:
            seen_s.add(ks)
            sample["ScheduleLines"].append({
                "PR_Ref": r["PR_Ref"], "PurchaseRequisitionItem": ki[1],
                "ScheduleLine": r.get("ScheduleLine", ""),
                "ScheduleLineDeliveryDate": r.get("DeliveryDate", ""),
                "ScheduleLineOrderQuantity": r.get("ScheduleLineQuantity", "")})
        if r.get("ComponentMaterial"):
            sample["Components"].append({
                "PR_Ref": r["PR_Ref"], "PurchaseRequisitionItem": ki[1],
                "ScheduleLine": r.get("ScheduleLine", ""),
                **{col: r.get(flat, "") for col, flat in COMP_SHEET_TO_FLAT.items()}})

    head_font = Font(name="Arial", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="5B21B6")
    wb = Workbook()
    info = wb.active
    info.title = "Instructions"
    lines = [
        ("Subcontracting PR upload template (API_PURCHASEREQUISITION_2, OData V4, DEV)", True),
        ("Fill Header, Items, ScheduleLines and Components. The example rows (row 2 onwards) "
         "show the layout - replace them with your own data.", False),
        ("Header: one row per PR. PR_Ref is your own key.", False),
        ("Items: linked to Header by PR_Ref. Material is the finished good. Item category is "
         "always sent as 3 (subcontracting).", False),
        ("ScheduleLines: linked to Items by PR_Ref + PurchaseRequisitionItem. A PR item has one "
         "delivery date, so several schedule lines become one PR item each (item 10 with lines "
         "1 and 2 is sent as items 10 and 11).", False),
        ("Components: linked to schedule lines by PR_Ref + PurchaseRequisitionItem + "
         "ScheduleLine. One row per component material. Reference only - the PR API cannot "
         "receive components, SAP takes them from the material's BOM. The results show how "
         "many component rows were not sent.", False),
        ("Dates: YYYY-MM-DD or DD.MM.YYYY. Cells are text-formatted so leading zeros are kept.",
         False),
    ]
    for i, (text, bold) in enumerate(lines, 1):
        c = info.cell(row=i, column=1, value=text)
        c.font = Font(name="Arial", bold=bold, size=13 if bold else 10)
        c.alignment = Alignment(wrap_text=True)
    info.column_dimensions["A"].width = 110
    for name, cols in SHEETS.items():
        ws = wb.create_sheet(name)
        for j, col in enumerate(cols, 1):
            h = ws.cell(row=1, column=j, value=col)
            h.font, h.fill = head_font, head_fill
            ws.column_dimensions[h.column_letter].width = max(12, len(col) + 2)
            for k, ex in enumerate(sample[name], 2):
                ws.cell(row=k, column=j, value=ex.get(col, "")).font = Font(name="Arial")
            for r in range(2, 301):
                ws.cell(row=r, column=j).number_format = "@"
        ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Payloads
# ---------------------------------------------------------------------------
_COMPONENTS_NOT_SENT: dict[int, int] = {}    # id(payload) -> component rows not sent


def _dec(v: str):
    s = str(v or "").replace(" ", "")
    s = s.replace(",", ".") if ("," in s and "." not in s) else s.replace(",", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def build_documents(df) -> list[dict]:
    """Collapse rows to one row per PR item / schedule line, then build with the PR builder."""
    rows = [{k: str(v).strip() for k, v in r.items()}
            for r in df.fillna("").to_dict("records")]
    groups: dict[tuple, dict] = {}          # (ref, item, sched) -> collapsed row
    order: list[tuple] = []
    item_rows: dict[tuple, dict] = {}       # (ref, item) -> first row with item data
    extra: dict[str, list[str]] = {}
    comps: dict[str, int] = {}
    pr_has_items: dict[str, bool] = {}

    for i, r in enumerate(rows):
        line = i + 2
        ref = r.get("PR_Ref", "") or f"LINE{line}"
        has_item = any(r.get(c) for c in ITEM_COLUMNS + SCHED_COLUMNS[1:] + COMP_COLUMNS)
        pr_has_items.setdefault(ref, False)
        if not has_item and not r.get("PurchaseRequisitionItem"):
            groups.setdefault((ref, "", "header"), {"row": r, "lines": [line], "header": True})
            if (ref, "", "header") not in order:
                order.append((ref, "", "header"))
            continue
        pr_has_items[ref] = True
        item = _num(r.get("PurchaseRequisitionItem")) or f"row{line}"
        sched = _num(r.get("ScheduleLine")) or "1"
        key = (ref, item, sched)
        if key not in groups:
            base = item_rows.setdefault((ref, item), r)
            merged = {**{c: base.get(c, "") for c in ITEM_COLUMNS}, **{
                c: r.get(c, "") for c in ["PR_Ref"] + HEADER_COLUMNS}}
            merged.update({c: r.get(c, "") for c in SCHED_COLUMNS})
            groups[key] = {"row": merged, "lines": [line], "header": False}
            order.append(key)
        else:
            groups[key]["lines"].append(line)
        if r.get("ComponentMaterial"):
            comps[ref] = comps.get(ref, 0) + 1
            for c, label in (("ComponentQuantity", "quantity"), ("ComponentUnit", "unit")):
                if not r.get(c):
                    extra.setdefault(ref, []).append(
                        f"line {line}: component {r['ComponentMaterial']} has no {label}")
        elif any(r.get(c) for c in COMP_COLUMNS):
            extra.setdefault(ref, []).append(f"line {line}: component row without Material")

    # schedule lines -> PR items
    out_rows, out_lines = [], []
    by_item: dict[tuple, list[tuple]] = {}
    for key in order:
        if not groups[key]["header"]:
            by_item.setdefault(key[:2], []).append(key)
    header_rows = {k[0]: groups[k]["row"] for k in order if groups[k]["header"]}
    emitted_header: set[str] = set()
    for key in order:
        g = groups[key]
        if g["header"]:
            if not pr_has_items.get(key[0]):
                out_rows.append(g["row"])
                out_lines.append(g["lines"])
            continue
        ref, item, sched = key
        if not any(g["row"].get(c) for c in ITEM_COLUMNS if c != "PurchaseRequisitionItem"):
            what = "component" if any(rows[ln - 2].get("ComponentMaterial")
                                      for ln in g["lines"]) else "schedule line"
            extra.setdefault(ref, []).append(
                f"line {g['lines'][0]}: {what} for item {item} - item {item} is not on the "
                f"Items sheet")
            continue
        scheds = by_item[(ref, item)]
        idx = scheds.index(key)
        row = dict(g["row"])
        if ref not in emitted_header and ref in header_rows:
            for c in HEADER_COLUMNS:
                row[c] = row.get(c) or header_rows[ref].get(c, "")
        emitted_header.add(ref)
        item_qty = row.get("RequestedQuantity", "")
        if row.get("ScheduleLineQuantity"):
            row["RequestedQuantity"] = row["ScheduleLineQuantity"]
        elif len(scheds) > 1:
            extra.setdefault(ref, []).append(
                f"item {item}: schedule line {sched} needs a quantity (the item has "
                f"{len(scheds)} schedule lines)")
        if idx == 0 and len(scheds) > 1 and item_qty:
            total = sum((_dec(groups[k]["row"].get("ScheduleLineQuantity")) or 0) for k in scheds)
            if _dec(item_qty) is not None and total != _dec(item_qty):
                extra.setdefault(ref, []).append(
                    f"item {item}: schedule line quantities add up to {total}, "
                    f"RequestedQuantity is {item_qty}")
        if item.isdigit() and idx > 0:
            row["PurchaseRequisitionItem"] = str(int(item) + idx)
        elif not item.isdigit():
            row["PurchaseRequisitionItem"] = ""
        out_rows.append(row)
        out_lines.append(g["lines"])

    pr_columns = pr.COLUMNS + [c for c in ITEM_COLUMNS if c not in pr.COLUMNS]  # + Subcontractor
    flat = pd.DataFrame([{c: r.get(c, "") for c in pr_columns} for r in out_rows],
                        columns=pr_columns).fillna("")
    docs = pr.build_prs(flat, [pr.SUBCONTRACTING], pr.SUBCONTRACTING) if out_rows else []

    # line numbers back to the uploaded rows
    def fix(text):
        return re.sub(r"line (\d+):", lambda m: f"line {out_lines[int(m.group(1)) - 2][0]}:"
                      if 0 <= int(m.group(1)) - 2 < len(out_lines) else m.group(0), text)

    by_ref_lines: dict[str, list[int]] = {}
    for r, ls in zip(out_rows, out_lines):
        by_ref_lines.setdefault(r.get("PR_Ref", ""), []).extend(ls)
    for d in docs:
        d["errors"] = [fix(e) for e in d["errors"]] + extra.get(d["ref"], [])
        if not pr_has_items.get(d["ref"], True):
            d["payload"][pr.ITEM_NAV] = []
            d["errors"] = [f"PR {d['ref']} has no items - add its rows on the Items sheet "
                           f"(or item columns in the flat sheet)"]
        lines = sorted(set(by_ref_lines.get(d["ref"], [])))
        if lines:
            d["lines"] = ",".join(map(str, lines))
        if comps.get(d["ref"]):
            _COMPONENTS_NOT_SENT[id(d["payload"])] = comps[d["ref"]]
    return docs


def post_document(client, payload: dict):
    """POST one PR (shared poster) and say how many component rows were not sent."""
    resp, note = pr.post_document(client, payload)
    n = _COMPONENTS_NOT_SENT.get(id(payload), 0)
    if n and note:
        note += (f"; {n} component row(s) not sent - the PR API has no components, "
                 f"SAP uses the BOM")
    return resp, note


if __name__ == "__main__":
    run_cli(sys.modules[__name__])