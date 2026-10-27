#!/usr/bin/env python3
"""
Bulk STANDARD Purchase Order upload (ME21N) - API_PURCHASEORDER_PROCESS_SRV, SAP DEV.

Standard and account-assigned (K, F, P, A ...) purchase orders from one flat
sheet. Subcontracting POs (item category L / 3) have their own script:
bulk_subcontracting_po_upload.py - rows with category L or 3 are rejected here.

  PO_Ref            -> one purchase order (header columns read from its first row)
  PurchaseOrderItem -> one item per PO_Ref + item number
  ScheduleLine      -> one schedule line per item + schedule line (defaults to 0001)
  AccountAssignmentCategory + CostCenter/GLAccount/WBSElement/OrderID/Asset
                    -> account assignment on the item

    python bulk_purchase_order_upload.py --template xlsx
    python bulk_purchase_order_upload.py -f purchase_orders.xlsx --dry-run
    python bulk_purchase_order_upload.py -f purchase_orders.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sap_common import (EntityBuilder, load_metadata, pad, rows_of, run_cli,  # noqa: E402
                        service_url, validate_payload)

TITLE = "Standard purchase orders"
SLUG = "purchase_order"
SERVICE = "API_PURCHASEORDER_PROCESS_SRV"
ENTITY_SET = "A_PurchaseOrder"
HEADER_TYPE = "A_PurchaseOrderType"
ITEM_TYPE = "A_PurchaseOrderItemType"
SCHED_TYPE = "A_PurchaseOrderScheduleLineType"
ACCT_TYPE = "A_PurOrdAccountAssignmentType"
BASE_URL = service_url(SERVICE)          # DEV: my440263
DOC_FIELD = "PurchaseOrder"
YEAR_FIELD = None
TCODE = "ME21N"
ICON = "🛒"
COLOR = "#2563EB"
GROUP_LABEL = "PO_Ref"
ITEMS_LABEL = "Items"
DESCRIPTION = ("Standard and account-assigned purchase orders (cost center, order, project, "
               "asset). One flat sheet, grouped by PO_Ref → item → schedule line.")

# External (ME21N) item category letter -> INTERNAL key the API expects.
# Sending the letter (e.g. "L") fails with "Use a supported purchase order item category".
ITEM_CATEGORY_MAP = {
    "": "0",    # standard
    "0": "0",
    "B": "1",   # limit
    "K": "2",   # consignment
    "L": "3",   # subcontracting (own upload)
    "U": "7",   # stock transfer
    "D": "9",   # service
}
VALID_INTERNAL_ITEM_CATEGORIES = {"0", "1", "2", "3", "5", "7", "9"}
SUBCONTRACTING_CATEGORY = "3"


def count_items(p: dict) -> int:
    return len(p.get("to_PurchaseOrderItem", {}).get("results", []))


HEADER_FIELDS = ["CompanyCode", "PurchaseOrderType", "Supplier", "PurchasingOrganization",
                 "PurchasingGroup", "DocumentCurrency", "PurchaseOrderDate", "PaymentTerms",
                 "IncotermsClassification", "IncotermsLocation1", "ExchangeRate",
                 "ExchangeRateIsFixed", "CorrespncExternalReference", "Language"]
ITEM_FIELDS = ["PurchaseOrderItem", "Material", "PurchaseOrderItemText", "MaterialGroup", "Plant",
               "StorageLocation", "OrderQuantity", "PurchaseOrderQuantityUnit", "NetPriceAmount",
               "NetPriceQuantity", "TaxCode", "PurchaseOrderItemCategory",
               "AccountAssignmentCategory", "GoodsReceiptIsExpected", "InvoiceIsExpected",
               "InvoiceIsGoodsReceiptBased", "PurchasingInfoRecord", "PurchaseRequisition",
               "PurchaseRequisitionItem", "PurchaseContract", "PurchaseContractItem",
               "RequisitionerName", "Batch", "IsReturnsItem", "PurchasingItemIsFreeOfCharge"]
SCHED_MAP = {"ScheduleLine": "ScheduleLine", "ScheduleLineDeliveryDate": "DeliveryDate",
             "ScheduleLineOrderQuantity": "ScheduleLineQuantity"}
ACCT_MAP = {"AccountAssignmentNumber": "AccountAssignmentNumber", "CostCenter": "CostCenter",
            "GLAccount": "GLAccount", "WBSElement": "WBSElement", "OrderID": "OrderID",
            "MasterFixedAsset": "MasterFixedAsset", "FixedAsset": "FixedAsset",
            "ProfitCenter": "AcctProfitCenter", "Quantity": "AcctQuantity"}

COLUMNS = (["PO_Ref"] + HEADER_FIELDS + ITEM_FIELDS + list(SCHED_MAP.values())
           + ["AccountAssignmentNumber", "CostCenter", "GLAccount", "WBSElement", "OrderID",
              "MasterFixedAsset", "FixedAsset", "AcctProfitCenter", "AcctQuantity"])

COLUMN_NOTES = {
    "PO_Ref": "Required. Your own key - rows with the same value form one PO.",
    "CompanyCode": "Required (first row of the PO).",
    "PurchaseOrderType": "Required, e.g. NB.",
    "Supplier": "Required.",
    "PurchasingOrganization": "Required.",
    "PurchasingGroup": "Required, e.g. 001.",
    "DocumentCurrency": "Required, e.g. INR.",
    "PurchaseOrderItem": "Item number, e.g. 10 (padded to 00010). Defaults to 10, 20, 30...",
    "Material": "Material, or leave blank and fill PurchaseOrderItemText + MaterialGroup.",
    "Plant": "Required.",
    "OrderQuantity": "Required.",
    "PurchaseOrderQuantityUnit": "Required, e.g. PC.",
    "NetPriceAmount": "Net price.",
    "TaxCode": "Tax code, e.g. G4.",
    "PurchaseOrderItemCategory": "Blank = standard. ME21N letters are converted: B = limit (1), "
                                 "K = consignment (2), U = stock transfer (7), D = service (9); "
                                 "internal keys 0/1/2/5/7/9 also work. "
                                 "Subcontracting (L/3) goes in the subcontracting upload.",
    "AccountAssignmentCategory": "K = cost center, F = order, P = project, A = asset.",
    "GoodsReceiptIsExpected": "TRUE / FALSE.",
    "InvoiceIsExpected": "TRUE / FALSE.",
    "ScheduleLine": "Defaults to 0001.",
    "DeliveryDate": "Required. Delivery date of the schedule line.",
    "ScheduleLineQuantity": "Defaults to OrderQuantity.",
    "AccountAssignmentNumber": "Defaults to 01 when an account assignment is filled.",
    "CostCenter": "Account assignment K.",
    "GLAccount": "G/L account for account-assigned items.",
    "OrderID": "Account assignment F.",
    "WBSElement": "Account assignment P.",
    "MasterFixedAsset": "Account assignment A.",
}

SAMPLE_ROWS = [
    {"PO_Ref": "PO1", "CompanyCode": "1810", "PurchaseOrderType": "NB", "Supplier": "18300010",
     "PurchasingOrganization": "1810", "PurchasingGroup": "001", "DocumentCurrency": "INR",
     "PurchaseOrderItem": "10", "Material": "RM019", "Plant": "1810", "OrderQuantity": "50",
     "PurchaseOrderQuantityUnit": "KG", "NetPriceAmount": "120.00", "NetPriceQuantity": "1",
     "TaxCode": "G4", "GoodsReceiptIsExpected": "TRUE", "InvoiceIsExpected": "TRUE",
     "DeliveryDate": "2026-10-20"},
    {"PO_Ref": "PO1", "PurchaseOrderItem": "20", "Material": "RM020", "Plant": "1810",
     "OrderQuantity": "20", "PurchaseOrderQuantityUnit": "KG", "NetPriceAmount": "80.00",
     "NetPriceQuantity": "1", "TaxCode": "G4", "GoodsReceiptIsExpected": "TRUE",
     "InvoiceIsExpected": "TRUE", "DeliveryDate": "2026-10-22"},
    {"PO_Ref": "PO2", "CompanyCode": "1810", "PurchaseOrderType": "NB", "Supplier": "18300010",
     "PurchasingOrganization": "1810", "PurchasingGroup": "001", "DocumentCurrency": "INR",
     "PurchaseOrderItem": "10", "PurchaseOrderItemText": "Office chairs", "MaterialGroup": "L001",
     "Plant": "1810", "OrderQuantity": "5", "PurchaseOrderQuantityUnit": "EA",
     "NetPriceAmount": "4500.00", "NetPriceQuantity": "1", "TaxCode": "G4",
     "AccountAssignmentCategory": "K", "GoodsReceiptIsExpected": "TRUE",
     "InvoiceIsExpected": "TRUE", "DeliveryDate": "2026-10-25", "CostCenter": "18101101",
     "GLAccount": "61002000"},
]


def build_documents(df) -> list[dict]:
    md = load_metadata(SERVICE)
    hb, ib = EntityBuilder(md, HEADER_TYPE), EntityBuilder(md, ITEM_TYPE)
    sb, ab = EntityBuilder(md, SCHED_TYPE), EntityBuilder(md, ACCT_TYPE)
    pos: dict[str, dict] = {}

    for row in rows_of(df):
        errs: list[str] = []
        ref = row.get("PO_Ref") or f"LINE{row['_line']}"
        if ref not in pos:
            for f in ("CompanyCode", "PurchaseOrderType", "Supplier", "PurchasingOrganization",
                      "PurchasingGroup", "DocumentCurrency"):
                if not row.get(f):
                    errs.append(f"{f} is required on the first row of {ref}")
            header = hb.build(row, HEADER_FIELDS, errs)
            header["to_PurchaseOrderItem"] = {"results": []}
            pos[ref] = {"ref": ref, "lines": [], "errors": [], "payload": header, "items": {}}
        po = pos[ref]

        item_no = pad("PurchaseOrderItem", row.get("PurchaseOrderItem")
                      or str(10 * (len(po["items"]) + 1)))
        if item_no not in po["items"]:
            cat_in = row.get("PurchaseOrderItemCategory", "").strip().upper()
            cat = ITEM_CATEGORY_MAP.get(cat_in, cat_in)
            if cat == SUBCONTRACTING_CATEGORY:
                errs.append(f"item {item_no}: category {cat_in} is "
                            f"subcontracting - use the Subcontracting PO upload")
            elif cat not in VALID_INTERNAL_ITEM_CATEGORIES:
                errs.append(f"item {item_no}: unknown PurchaseOrderItemCategory '{cat_in}'")
            for f in ("Plant", "OrderQuantity", "PurchaseOrderQuantityUnit", "DeliveryDate"):
                if not row.get(f):
                    errs.append(f"{f} is required for item {item_no}")
            if not row.get("Material") and not row.get("PurchaseOrderItemText"):
                errs.append(f"item {item_no}: fill Material or PurchaseOrderItemText")
            item = ib.build(dict(row, PurchaseOrderItem=item_no,
                                 PurchaseOrderItemCategory=cat if cat_in else ""),
                            ITEM_FIELDS, errs)
            item["to_ScheduleLine"] = {"results": []}
            if row.get("AccountAssignmentCategory"):
                acct = ab.build(row, ACCT_MAP, errs, defaults={"AccountAssignmentNumber": "01"})
                if len(acct) > 1:
                    item["to_AccountAssignment"] = {"results": [acct]}
                else:
                    errs.append(f"item {item_no}: AccountAssignmentCategory "
                                f"{row['AccountAssignmentCategory']} needs CostCenter/"
                                f"OrderID/WBSElement/asset")
            po["payload"]["to_PurchaseOrderItem"]["results"].append(item)
            po["items"][item_no] = {"item": item, "scheds": set(), "row": row}
        it = po["items"][item_no]

        sched_no = pad("ScheduleLine", row.get("ScheduleLine") or "1")
        if sched_no not in it["scheds"]:
            sl = sb.build(dict(row, ScheduleLine=sched_no), SCHED_MAP, errs,
                          defaults={"ScheduleLineOrderQuantity": row.get("OrderQuantity")
                                    or it["row"].get("OrderQuantity", "")})
            it["item"]["to_ScheduleLine"]["results"].append(sl)
            it["scheds"].add(sched_no)

        po["lines"].append(row["_line"])
        po["errors"] += [f"line {row['_line']}: {e}" for e in errs]

    return [{"ref": po["ref"], "lines": ",".join(map(str, po["lines"])),
             "errors": po["errors"] + validate_payload(md, HEADER_TYPE, po["payload"]),
             "payload": po["payload"]} for po in pos.values()]


if __name__ == "__main__":
    run_cli(sys.modules[__name__])