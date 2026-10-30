#!/usr/bin/env python3
"""
Bulk Purchase Requisition upload (ME51N) - API_PURCHASEREQUISITION_2 (OData V4)
/ PurchaseReqn with _PurchaseRequisitionItem, SAP DEV.

One row = one PR item. Rows with the same PR_Ref form one purchase requisition
(header fields come from the first row of each PR_Ref). Item numbers are
optional: leave PurchaseRequisitionItem blank and SAP numbers the items.

All PR kinds from one sheet - set PRCategory on each row:

  STANDARD        stock material                      (blank = STANDARD)
  IMPORT          same as STANDARD; use your import PR type, foreign supplier/currency
  CONSUMABLE      account-assigned (K/F/P/A...) - material or text item
  SERVICE         service (product type 2), account-assigned, performance period
  SUBCONTRACTING  item category L (3) - components come from the material's BOM
  STOCK_TRANSFER  item category U (7) - needs SupplyingPlant
  CONSIGNMENT     item category K (2)
  LIMIT           item category B (1) - limit amounts + period, account-assigned

A PR can mix categories. PurchasingDocumentItemCategory and ProductTypeCode
can also be filled directly; they override what PRCategory sets.

OData V4: dates are "YYYY-MM-DD", amounts are JSON numbers, items and account
assignments are plain lists. The PR number is read from the V4 response.

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
DESCRIPTION = ("Standard, import, consumable, service, subcontracting, stock transfer, "
               "consignment and limit PRs (OData V4). One row per item, grouped by PR_Ref; "
               "PRCategory on each row sets the item type.")

ITEM_NAV = "_PurchaseRequisitionItem"
ACCT_NAV = "_PurchaseReqnAcctAssgmt"
TEXT_NAV = "_PurchaseReqnItemText"
ITEM_TEXT_TYPE = "B01"                    # PR item text


def count_items(p: dict) -> int:
    return len(p.get(ITEM_NAV, []))


# ---------------------------------------------------------------------------
# Item categories (same letters -> internal keys as the PO script)
# ---------------------------------------------------------------------------
ITEM_CATEGORY_MAP = {"": "0", "0": "0", "B": "1", "1": "1", "K": "2", "2": "2", "L": "3",
                     "3": "3", "U": "7", "7": "7", "D": "9", "9": "9", "5": "5"}

# PRCategory -> what it sets and what it needs
CATEGORIES = {
    "STANDARD": {"item_cat": "0", "product_type": "1"},
    "IMPORT": {"item_cat": "0", "product_type": "1"},
    "CONSUMABLE": {"item_cat": "0", "product_type": "1", "acct": True},
    "SERVICE": {"item_cat": "0", "product_type": "2", "acct": True, "period": True},
    "SUBCONTRACTING": {"item_cat": "3", "product_type": "1", "material": True},
    "STOCK_TRANSFER": {"item_cat": "7", "product_type": "1", "material": True,
                       "need": ["SupplyingPlant"]},
    "CONSIGNMENT": {"item_cat": "2", "product_type": "1", "material": True},
    "LIMIT": {"item_cat": "1", "product_type": "1", "acct": True, "period": True,
              "need": ["OverallLimitAmount"], "no_qty": True},
}
CATEGORY_ALIASES = {"": "STANDARD", "STOCK": "STANDARD", "NORMAL": "STANDARD",
                    "STO": "STOCK_TRANSFER", "STOCK TRANSFER": "STOCK_TRANSFER",
                    "TRANSFER": "STOCK_TRANSFER", "SUBCON": "SUBCONTRACTING",
                    "CONSUMPTION": "CONSUMABLE", "CONSUMABLES": "CONSUMABLE",
                    "SERVICES": "SERVICE", "IMPORTS": "IMPORT"}

# ---------------------------------------------------------------------------
# Fields: name -> (kind, max length). kind: str | date | dec | bool | item
# Lengths follow the API schema (PurchaseReqn / item / account assignment).
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
    "AccountAssignmentCategory": ("str", 1),
    "PurchasingDocumentItemCategory": ("str", 1),
    "ProductTypeCode": ("str", 2),
    "FixedSupplier": ("str", 10),
    "Supplier": ("str", 10),
    "SupplyingPlant": ("str", 4),
    "Subcontractor": ("str", 10),
    "PurchasingInfoRecord": ("str", 10),
    "PurchaseContract": ("str", 10),
    "PurchaseContractItem": ("item", 5),
    "ServicePerformer": ("str", 10),
    "PerformancePeriodStartDate": ("date", None),
    "PerformancePeriodEndDate": ("date", None),
    "ExpectedOverallLimitAmount": ("dec", None),
    "OverallLimitAmount": ("dec", None),
    "TaxCode": ("str", 2),
    "GoodsReceiptIsExpected": ("bool", None),
    "InvoiceIsExpected": ("bool", None),
    "GoodsReceiptIsNonValuated": ("bool", None),
    "Batch": ("str", 10),
    "RequisitionerName": ("str", 12),
    "RequirementTracking": ("str", 10),
    "MaterialGoodsReceiptDuration": ("dec", None),
    "PurReqnItemExternalReference": ("str", 70),
}
ACCT_FIELDS = {          # column name -> API field (one assignment per item)
    "CostCenter": ("CostCenter", "str", 10),
    "GLAccount": ("GLAccount", "str", 10),
    "OrderID": ("OrderID", "str", 12),
    "WBSElement": ("WBSElementExternalID", "str", 24),
    "MasterFixedAsset": ("MasterFixedAsset", "str", 12),
    "FixedAsset": ("FixedAsset", "str", 4),
    "ProfitCenter": ("ProfitCenter", "str", 10),
    "SalesOrder": ("SalesOrder", "str", 10),
    "SalesOrderItem": ("SalesOrderItem", "item", 6),
    "BusinessArea": ("BusinessArea", "str", 4),
    "FundsCenter": ("FundsCenter", "str", 16),
    "GoodsRecipientName": ("GoodsRecipientName", "str", 12),
    "UnloadingPointName": ("UnloadingPointName", "str", 25),
}
# Which account-assignment object each category letter needs
ACCT_OBJECT = {"K": ["CostCenter"], "F": ["OrderID"], "P": ["WBSElement"],
               "A": ["MasterFixedAsset"], "C": ["SalesOrder"], "E": ["SalesOrder"]}

HEADER_DEFAULTS = {"SourceDetermination": False, "PurReqnDoOnlyValidation": False}

COLUMNS = (["PR_Ref", "PRCategory"] + list(HEADER_FIELDS) + list(ITEM_FIELDS)
           + list(ACCT_FIELDS) + ["ItemText", "ItemTextLanguage"])

COLUMN_NOTES = {
    "PR_Ref": "Required. Your own key - rows with the same value form one purchase requisition.",
    "PRCategory": "STANDARD (blank), IMPORT, CONSUMABLE, SERVICE, SUBCONTRACTING, "
                  "STOCK_TRANSFER, CONSIGNMENT or LIMIT. Sets item category / product type "
                  "and the checks below.",
    "PurchaseRequisitionType": "Required (first row of the PR), e.g. ZDOM; your import type "
                               "for IMPORT PRs.",
    "PurReqnDescription": "Optional header description (max 40).",
    "PurReqnHeaderNote": "Optional header note (free text).",
    "SourceDetermination": "Optional. TRUE = SAP determines the source of supply. Default FALSE.",
    "PurchaseRequisitionItem": "Optional item number, e.g. 10. Blank = SAP numbers the items.",
    "PurchaseRequisitionItemText": "Short text (max 40). Required when Material is blank.",
    "Material": "Material / service product. Blank = text item (then MaterialGroup is needed). "
                "Required for SUBCONTRACTING, STOCK_TRANSFER and CONSIGNMENT.",
    "MaterialGroup": "Required for text items without Material.",
    "RequestedQuantity": "Required (not for LIMIT). Quantity, e.g. 10.",
    "BaseUnit": "Required (or BaseUnitISOCode), e.g. KG, PC, EA; for services e.g. AU or H.",
    "BaseUnitISOCode": "Optional ISO unit, e.g. KGM for KG, PCE for PC.",
    "Plant": "Required. For STOCK_TRANSFER: the receiving plant.",
    "StorageLocation": "Optional.",
    "PurchasingGroup": "Required, e.g. 002.",
    "PurchasingOrganization": "Optional. Blank = SAP determines it.",
    "DeliveryDate": "Required. YYYY-MM-DD or DD.MM.YYYY. SERVICE / LIMIT: defaults to "
                    "PerformancePeriodStartDate.",
    "PurchaseRequisitionPrice": "Valuation price. Needed for text items, services and "
                                "materials without a price.",
    "PurReqnPriceQuantity": "Price unit, e.g. 1.",
    "PurReqnItemCurrency": "Currency of the price, e.g. INR; for IMPORT e.g. USD.",
    "AccountAssignmentCategory": "K cost center, F order, P project (WBS), A asset, C/E sales "
                                 "order. Required for CONSUMABLE, SERVICE and LIMIT.",
    "PurchasingDocumentItemCategory": "Optional override. Letters converted like ME51N: "
                                      "B limit, K consignment, L subcontracting, U stock "
                                      "transfer, D service; or internal keys 0/1/2/3/5/7/9.",
    "ProductTypeCode": "Optional override: 1 = material, 2 = service.",
    "FixedSupplier": "Optional fixed supplier (IMPORT: the foreign supplier).",
    "Supplier": "Optional desired supplier.",
    "SupplyingPlant": "STOCK_TRANSFER: the plant that ships the goods.",
    "Subcontractor": "SUBCONTRACTING: optional subcontractor (supplier).",
    "PurchasingInfoRecord": "Optional source of supply: info record.",
    "PurchaseContract": "Optional source of supply: contract.",
    "PurchaseContractItem": "Contract item, e.g. 10.",
    "ServicePerformer": "SERVICE: optional service performer.",
    "PerformancePeriodStartDate": "SERVICE / LIMIT: required start of the period.",
    "PerformancePeriodEndDate": "SERVICE / LIMIT: required end of the period.",
    "ExpectedOverallLimitAmount": "LIMIT: expected value.",
    "OverallLimitAmount": "LIMIT: required overall limit.",
    "TaxCode": "Optional tax code, e.g. G4.",
    "GoodsReceiptIsExpected": "Optional TRUE / FALSE (services often FALSE).",
    "InvoiceIsExpected": "Optional TRUE / FALSE.",
    "GoodsReceiptIsNonValuated": "Optional TRUE / FALSE.",
    "Batch": "Optional batch.",
    "RequisitionerName": "Optional requisitioner (max 12).",
    "RequirementTracking": "Optional tracking number.",
    "MaterialGoodsReceiptDuration": "Optional GR processing days.",
    "PurReqnItemExternalReference": "Optional external reference.",
    "CostCenter": "Account assignment K (also a CO object for F/P when required).",
    "GLAccount": "Account assignment: G/L account (usually needed for CONSUMABLE / SERVICE).",
    "OrderID": "Account assignment F: internal / production order.",
    "WBSElement": "Account assignment P: WBS element, e.g. P-1000-01.",
    "MasterFixedAsset": "Account assignment A: asset number.",
    "FixedAsset": "Account assignment A: asset sub-number, e.g. 0.",
    "ProfitCenter": "Optional profit center.",
    "SalesOrder": "Account assignment C/E: sales order.",
    "SalesOrderItem": "Account assignment C/E: sales order item, e.g. 10.",
    "BusinessArea": "Optional business area.",
    "FundsCenter": "Optional funds center.",
    "GoodsRecipientName": "Optional goods recipient (max 12).",
    "UnloadingPointName": "Optional unloading point (max 25).",
    "ItemText": "Optional item long text (sent as item text B01). Alt+Enter for new lines.",
    "ItemTextLanguage": "Language of ItemText. Default EN.",
}

_BASE = {"Plant": "01TG", "PurchasingGroup": "002", "DeliveryDate": "2026-11-15"}
SAMPLE_ROWS = [
    {"PR_Ref": "PR1", "PRCategory": "STANDARD", "PurchaseRequisitionType": "ZDOM",
     "PurReqnDescription": "Stock material", "PurchaseRequisitionItemText": "Test Material TG10-1",
     "Material": "RM019", "RequestedQuantity": "10", "BaseUnit": "KG", "BaseUnitISOCode": "KGM",
     **_BASE},
    {"PR_Ref": "PR1", "PRCategory": "STANDARD", "PurchaseRequisitionItemText": "Test Material TG10-2",
     "Material": "RM019", "RequestedQuantity": "10", "BaseUnit": "KG", "BaseUnitISOCode": "KGM",
     "ItemText": "Deliver in original packaging.", **_BASE},
    {"PR_Ref": "PR2", "PRCategory": "CONSUMABLE", "PurchaseRequisitionType": "ZDOM",
     "PurReqnDescription": "Office consumables", "PurchaseRequisitionItemText": "A4 paper",
     "MaterialGroup": "L002", "RequestedQuantity": "20", "BaseUnit": "PC",
     "PurchaseRequisitionPrice": "250", "PurReqnPriceQuantity": "1", "PurReqnItemCurrency": "INR",
     "AccountAssignmentCategory": "K", "CostCenter": "01TG1001", "GLAccount": "61002000", **_BASE},
    {"PR_Ref": "PR3", "PRCategory": "SERVICE", "PurchaseRequisitionType": "ZDOM",
     "PurReqnDescription": "Maintenance service", "PurchaseRequisitionItemText": "AMC for compressor",
     "MaterialGroup": "L002", "RequestedQuantity": "1", "BaseUnit": "AU",
     "PurchaseRequisitionPrice": "50000", "PurReqnPriceQuantity": "1", "PurReqnItemCurrency": "INR",
     "AccountAssignmentCategory": "K", "CostCenter": "01TG1001", "GLAccount": "61002000",
     "PerformancePeriodStartDate": "2026-11-01", "PerformancePeriodEndDate": "2027-10-31",
     "Plant": "01TG", "PurchasingGroup": "002"},
    {"PR_Ref": "PR4", "PRCategory": "SUBCONTRACTING", "PurchaseRequisitionType": "ZDOM",
     "PurReqnDescription": "Subcontracting", "Material": "FG228", "RequestedQuantity": "100",
     "BaseUnit": "PC", "Subcontractor": "18300010", **_BASE},
    {"PR_Ref": "PR5", "PRCategory": "STOCK_TRANSFER", "PurchaseRequisitionType": "ZDOM",
     "PurReqnDescription": "Stock transfer", "Material": "RM019", "RequestedQuantity": "50",
     "BaseUnit": "KG", "SupplyingPlant": "01DR", **_BASE},
    {"PR_Ref": "PR6", "PRCategory": "IMPORT", "PurchaseRequisitionType": "ZIMP",
     "PurReqnDescription": "Import material", "Material": "RM019", "RequestedQuantity": "500",
     "BaseUnit": "KG", "FixedSupplier": "18300020", "PurReqnItemCurrency": "USD", **_BASE},
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


def _convert(kind: str, maxlen, value: str):
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
    for col, spec in fields.items():
        api, kind, maxlen = spec if len(spec) == 3 else (col, *spec)
        v = row.get(col, "").strip()
        if not v:
            continue
        try:
            out[api] = _convert(kind, maxlen, v)
        except ValueError as e:
            errors.append(f"{col}: {e}")
    return out


def _category(row: dict, errs: list[str]) -> str:
    raw = row.get("PRCategory", "").strip().upper().replace("-", "_")
    cat = CATEGORY_ALIASES.get(raw, CATEGORY_ALIASES.get(raw.replace("_", " "), raw))
    if cat not in CATEGORIES:
        errs.append(f"PRCategory '{row.get('PRCategory')}' is not one of "
                    f"{', '.join(CATEGORIES)}")
        return "STANDARD"
    return cat


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

        cat = _category(row, errs)
        rule = CATEGORIES[cat]
        item = _build(row, ITEM_FIELDS, errs)

        # item category / product type from PRCategory unless given explicitly
        given = row.get("PurchasingDocumentItemCategory", "").strip().upper()
        if given:
            internal = ITEM_CATEGORY_MAP.get(given)
            if internal is None:
                errs.append(f"PurchasingDocumentItemCategory '{given}' is not supported "
                            f"(use B, K, L, U, D or 0/1/2/3/5/7/9)")
            else:
                item["PurchasingDocumentItemCategory"] = internal
        elif rule["item_cat"] != "0":                 # "0" (standard) is SAP's default
            item["PurchasingDocumentItemCategory"] = rule["item_cat"]
        if not row.get("ProductTypeCode") and rule["product_type"] != "1":   # 1 = default
            item["ProductTypeCode"] = rule["product_type"]

        # service / limit: delivery date defaults to the period start
        if rule.get("period"):
            for f in ("PerformancePeriodStartDate", "PerformancePeriodEndDate"):
                if not row.get(f):
                    errs.append(f"{f} is required for {cat} items")
            if "DeliveryDate" not in item and "PerformancePeriodStartDate" in item:
                item["DeliveryDate"] = item["PerformancePeriodStartDate"]
            s, e = item.get("PerformancePeriodStartDate"), item.get("PerformancePeriodEndDate")
            if s and e and e < s:
                errs.append("PerformancePeriodEndDate is before PerformancePeriodStartDate")

        # common required fields
        need = ["Plant", "PurchasingGroup"] + rule.get("need", [])
        if "DeliveryDate" not in item:
            need.append("DeliveryDate")
        if not rule.get("no_qty"):
            need.append("RequestedQuantity")
            if not row.get("BaseUnit") and not row.get("BaseUnitISOCode"):
                errs.append("BaseUnit (or BaseUnitISOCode) is required")
        for f in need:
            if not row.get(f) and f not in item:
                errs.append(f"{f} is required for {cat} items")
        if rule.get("material") and not row.get("Material"):
            errs.append(f"Material is required for {cat} items")
        if not row.get("Material"):
            if not row.get("PurchaseRequisitionItemText"):
                errs.append("fill Material, or PurchaseRequisitionItemText for a text item")
            if not row.get("MaterialGroup"):
                errs.append("MaterialGroup is required for a text item without Material")
        q = item.get("RequestedQuantity")
        if isinstance(q, (int, float)) and q <= 0:
            errs.append("RequestedQuantity must be greater than 0")
        if cat == "STOCK_TRANSFER" and row.get("SupplyingPlant") == row.get("Plant"):
            errs.append("SupplyingPlant must differ from Plant")

        # account assignment
        acct_cat = item.get("AccountAssignmentCategory", "").upper()
        acct = _build(row, ACCT_FIELDS, errs)
        if rule.get("acct") and not acct_cat:
            errs.append(f"AccountAssignmentCategory is required for {cat} items (e.g. K)")
        if acct_cat:
            item["AccountAssignmentCategory"] = acct_cat
            for col in ACCT_OBJECT.get(acct_cat, []):
                if not row.get(col):
                    errs.append(f"{col} is required for account assignment {acct_cat}")
            if acct:
                item[ACCT_NAV] = [acct]
        elif acct:
            errs.append("account assignment columns are filled but AccountAssignmentCategory "
                        "is empty")

        # item long text
        if row.get("ItemText"):
            text = row["ItemText"].replace("\r\n", "\n").replace("\r", "\n")
            lang = (row.get("ItemTextLanguage") or "EN").strip().upper()
            if len(lang) != 2:
                errs.append(f"ItemTextLanguage '{lang}' should be a 2-letter code, e.g. EN")
            item[TEXT_NAV] = [{"TextObjectType": ITEM_TEXT_TYPE, "Language": lang,
                               "PlainLongText": text}]

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