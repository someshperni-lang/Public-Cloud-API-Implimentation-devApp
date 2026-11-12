#!/usr/bin/env python3
"""
Bulk STANDARD Purchase Order upload (ME21N) - API_PURCHASEORDER_PROCESS_SRV, SAP DEV.

Standard and account-assigned (K, F, P, A ...) purchase orders from one flat
sheet with the full header / item / schedule line / account assignment field
set. Subcontracting POs (item category L / 3, with components) have their own
script: bulk_subcontracting_po_upload.py - category L / 3 and the Component*
columns are rejected here.

  PORefID           -> one purchase order (header columns read from its first row)
  PurchaseOrderItem -> one item per PORefID + item number
  ScheduleLine      -> one schedule line per item + schedule line (defaults to 0001)
  AccountAssignmentNumber -> one account assignment per item + number; repeat the
                       item row with 02, 03 ... (and Quantity, MultipleAcctAssgmt-
                       Distribution 1 = by quantity / 2 = by percent) to split it

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
GROUP_LABEL = "PORefID"
ITEMS_LABEL = "Items"
DESCRIPTION = ("Standard and account-assigned purchase orders (cost center, order, project, "
               "asset) with the full field set. One flat sheet, grouped by PORefID → item → "
               "schedule line / account assignment.")

def count_items(p: dict) -> int:
    return len(p.get("to_PurchaseOrderItem", {}).get("results", []))


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

# ---------------------------------------------------------------------------
# Columns - same names and order as the upload template
# ---------------------------------------------------------------------------
GROUP_COLUMN = "PORefID"
HEADER_FIELDS = [
    "CompanyCode", "PurchaseOrderType", "Supplier", "Language", "PaymentTerms",
    "CashDiscount1Days", "CashDiscount2Days", "NetPaymentDays", "CashDiscount1Percent",
    "CashDiscount2Percent", "PurchasingOrganization", "PurchasingGroup", "PurchaseOrderDate",
    "DocumentCurrency", "ExchangeRate", "ExchangeRateIsFixed", "ValidityStartDate",
    "ValidityEndDate", "SupplierQuotationExternalID", "PurchasingCollectiveNumber",
    "SupplierRespSalesPersonName", "SupplierPhoneNumber", "SupplyingSupplier", "SupplyingPlant",
    "IncotermsClassification", "CorrespncExternalReference", "CorrespncInternalReference",
    "InvoicingParty", "IncotermsVersion", "IncotermsLocation1", "IncotermsLocation2",
    "AddressCityName", "AddressFaxNumber", "AddressHouseNumber", "AddressName",
    "AddressPostalCode", "AddressStreetName", "AddressPhoneNumber", "AddressRegion",
    "AddressCountry", "AddressCorrespondenceLanguage"]
ITEM_FIELDS = [
    "PurchaseOrderItem", "PurchaseOrderItemText", "Plant", "StorageLocation", "MaterialGroup",
    "SupplierMaterialNumber", "OrderQuantity", "PurchaseOrderQuantityUnit", "OrderPriceUnit",
    "OrderPriceUnitToOrderUnitNmrtr", "OrdPriceUnitToOrderUnitDnmntr", "NetPriceAmount",
    "NetPriceQuantity", "TaxCode", "ShippingInstruction", "TaxDeterminationDate", "TaxCountry",
    "PriceIsToBePrinted", "OverdelivTolrtdLmtRatioInPct", "UnlimitedOverdeliveryIsAllowed",
    "UnderdelivTolrtdLmtRatioInPct", "ValuationType", "PurchaseOrderItemCategory",
    "AccountAssignmentCategory", "MultipleAcctAssgmtDistribution", "PartialInvoiceDistribution",
    "GoodsReceiptIsExpected", "GoodsReceiptIsNonValuated", "InvoiceIsExpected",
    "InvoiceIsGoodsReceiptBased", "PurchaseContract", "PurchaseContractItem", "Customer",
    "Subcontractor", "SupplierIsSubcontractor", "ItemNetWeight", "ItemWeightUnit",
    "TaxJurisdiction", "PricingDateControl", "ItemVolume", "ItemVolumeUnit",
    "SupplierConfirmationControlKey", "IncotermsTransferLocation", "EvaldRcptSettlmtIsAllowed",
    "PurchaseRequisition", "PurchaseRequisitionItem", "IsReturnsItem", "RequisitionerName",
    "ServicePackage", "Material", "InternationalArticleNumber", "ManufacturerMaterial",
    "ServicePerformer", "ProductType", "Batch", "PurchasingItemIsFreeOfCharge",
    "DeliveryAddressName", "DeliveryAddressCityName", "DeliveryAddressPostalCode",
    "DeliveryAddressRegion", "DeliveryAddressCountry", "ConsumptionTaxCtrlCode"]
ACCT_FIELDS = [
    "AccountAssignmentNumber", "Quantity", "GLAccount", "CostCenter", "ControllingArea",
    "ProfitCenter", "WBSElement", "OrderID", "MasterFixedAsset", "FixedAsset",
    "GoodsRecipientName", "UnloadingPointName"]
SCHED_FIELDS = ["ScheduleLine", "ScheduleLineDeliveryDate", "ScheduleLineOrderQuantity"]
COMPONENT_COLUMNS = ["ComponentMaterial", "ComponentQuantity", "ComponentUnit", "ComponentPlant"]

COLUMNS = ([GROUP_COLUMN] + HEADER_FIELDS + ITEM_FIELDS + ACCT_FIELDS + SCHED_FIELDS
           + COMPONENT_COLUMNS)

# account-assignment object each category needs (at least one of)
ACCT_OBJECT = {"K": ["CostCenter"], "F": ["OrderID"], "P": ["WBSElement"],
               "A": ["MasterFixedAsset"]}

_NOTES = {
    "PORefID": "Required. Your own key - rows with the same value form one PO.",
    "CompanyCode": "Required (first row of the PO).",
    "PurchaseOrderType": "Required, e.g. NB.",
    "Supplier": "Required. Supplier number.",
    "Language": "Optional, e.g. EN.",
    "PaymentTerms": "Optional, e.g. 0001. Blank = taken from the supplier.",
    "PurchasingOrganization": "Required.",
    "PurchasingGroup": "Required, e.g. 001.",
    "PurchaseOrderDate": "Optional. Blank = today.",
    "DocumentCurrency": "Required, e.g. INR.",
    "ExchangeRate": "Optional, for foreign-currency POs.",
    "SupplyingPlant": "Header supplying plant (stock transfer POs).",
    "AddressName": "Header address fields: optional one-time supplier address.",
    "PurchaseOrderItem": "Item number, e.g. 10 (padded to 00010). Defaults to 10, 20, 30...",
    "PurchaseOrderItemText": "Short text. Required when Material is blank (text item).",
    "Plant": "Required.",
    "MaterialGroup": "Required for text items without Material.",
    "OrderQuantity": "Required.",
    "PurchaseOrderQuantityUnit": "Required, e.g. PC.",
    "OrderPriceUnit": "Optional order price unit when it differs from the order unit.",
    "NetPriceAmount": "Net price.",
    "NetPriceQuantity": "Price unit, e.g. 1.",
    "TaxCode": "Tax code, e.g. G4.",
    "PurchaseOrderItemCategory": "Blank = standard. ME21N letters are converted: B = limit (1), "
                                 "K = consignment (2), U = stock transfer (7), D = service (9); "
                                 "internal keys 0/1/2/5/7/9 also work. "
                                 "Subcontracting (L/3) goes in the subcontracting upload.",
    "AccountAssignmentCategory": "K = cost center, F = order, P = project, A = asset.",
    "MultipleAcctAssgmtDistribution": "Several account assignments on one item: 1 = by "
                                      "quantity, 2 = by percentage.",
    "PartialInvoiceDistribution": "Several account assignments: how partial invoices are split.",
    "GoodsReceiptIsExpected": "TRUE / FALSE.",
    "InvoiceIsExpected": "TRUE / FALSE.",
    "InvoiceIsGoodsReceiptBased": "TRUE / FALSE (GR-based invoice verification).",
    "Subcontractor": "Not for subcontracting POs - use the Subcontracting PO upload.",
    "PurchaseRequisition": "Optional PR this item is created from.",
    "PurchaseRequisitionItem": "Item of that PR, e.g. 10.",
    "Material": "Material, or leave blank and fill PurchaseOrderItemText + MaterialGroup.",
    "ProductType": "Optional: 1 = material, 2 = service.",
    "AccountAssignmentNumber": "Defaults to 01. Repeat the item row with 02, 03 ... to add "
                               "more account assignments to the same item.",
    "Quantity": "Account assignment quantity (when splitting by quantity).",
    "GLAccount": "G/L account for account-assigned items.",
    "CostCenter": "Account assignment K.",
    "OrderID": "Account assignment F.",
    "WBSElement": "Account assignment P.",
    "MasterFixedAsset": "Account assignment A.",
    "FixedAsset": "Asset sub-number, e.g. 0.",
    "ScheduleLine": "Defaults to 0001. Repeat the item row with 2, 3 ... for more delivery dates.",
    "ScheduleLineDeliveryDate": "Required. Delivery date of the schedule line.",
    "ScheduleLineOrderQuantity": "Defaults to OrderQuantity.",
}
for _c in COMPONENT_COLUMNS:
    _NOTES[_c] = "Subcontracting only - leave blank here and use the Subcontracting PO upload."


def _auto_note(col: str) -> str:
    """Note from the DEV metadata label for columns without a hand-written note."""
    md = load_metadata(SERVICE)
    for et in (HEADER_TYPE, ITEM_TYPE, ACCT_TYPE, SCHED_TYPE):
        p = md.field(et, col) if md else None
        if p:
            kind = {"Edm.DateTime": " Date: YYYY-MM-DD.", "Edm.Boolean": " TRUE / FALSE.",
                    "Edm.Decimal": " Number."}.get(p.get("Type", ""), "")
            return f"Optional. {p.get('quickinfo') or p.get('label') or col}.{kind}"
    return "Optional."


COLUMN_NOTES = {c: _NOTES.get(c) or _auto_note(c) for c in COLUMNS}

SAMPLE_ROWS = [
    {"PORefID": "PO1", "CompanyCode": "1810", "PurchaseOrderType": "NB", "Supplier": "18300010",
     "PurchasingOrganization": "1810", "PurchasingGroup": "001", "DocumentCurrency": "INR",
     "PurchaseOrderItem": "10", "Material": "RM019", "Plant": "1810", "OrderQuantity": "50",
     "PurchaseOrderQuantityUnit": "KG", "NetPriceAmount": "120.00", "NetPriceQuantity": "1",
     "TaxCode": "G4", "GoodsReceiptIsExpected": "TRUE", "InvoiceIsExpected": "TRUE",
     "ScheduleLineDeliveryDate": "2026-12-20"},
    {"PORefID": "PO1", "PurchaseOrderItem": "20", "Material": "RM020", "Plant": "1810",
     "OrderQuantity": "20", "PurchaseOrderQuantityUnit": "KG", "NetPriceAmount": "80.00",
     "NetPriceQuantity": "1", "TaxCode": "G4", "GoodsReceiptIsExpected": "TRUE",
     "InvoiceIsExpected": "TRUE", "ScheduleLineDeliveryDate": "2026-12-22"},
    {"PORefID": "PO2", "CompanyCode": "1810", "PurchaseOrderType": "NB", "Supplier": "18300010",
     "PurchasingOrganization": "1810", "PurchasingGroup": "001", "DocumentCurrency": "INR",
     "PurchaseOrderItem": "10", "PurchaseOrderItemText": "Office chairs", "MaterialGroup": "L001",
     "Plant": "1810", "OrderQuantity": "6", "PurchaseOrderQuantityUnit": "EA",
     "NetPriceAmount": "4500.00", "NetPriceQuantity": "1", "TaxCode": "G4",
     "AccountAssignmentCategory": "K", "MultipleAcctAssgmtDistribution": "1",
     "GoodsReceiptIsExpected": "TRUE", "InvoiceIsExpected": "TRUE",
     "AccountAssignmentNumber": "01", "Quantity": "4", "CostCenter": "18101101",
     "GLAccount": "61002000", "ScheduleLineDeliveryDate": "2026-12-25"},
    {"PORefID": "PO2", "PurchaseOrderItem": "10", "AccountAssignmentNumber": "02",
     "Quantity": "2", "CostCenter": "18101102", "GLAccount": "61002000"},
]


def build_documents(df) -> list[dict]:
    md = load_metadata(SERVICE)
    hb, ib = EntityBuilder(md, HEADER_TYPE), EntityBuilder(md, ITEM_TYPE)
    sb, ab = EntityBuilder(md, SCHED_TYPE), EntityBuilder(md, ACCT_TYPE)
    pos: dict[str, dict] = {}

    for row in rows_of(df):
        errs: list[str] = []
        ref = row.get(GROUP_COLUMN) or f"LINE{row['_line']}"
        if ref not in pos:
            for f in ("CompanyCode", "PurchaseOrderType", "Supplier", "PurchasingOrganization",
                      "PurchasingGroup", "DocumentCurrency"):
                if not row.get(f):
                    errs.append(f"{f} is required on the first row of {ref}")
            header = hb.build(row, HEADER_FIELDS, errs)
            header["to_PurchaseOrderItem"] = {"results": []}
            pos[ref] = {"ref": ref, "lines": [], "errors": [], "payload": header, "items": {}}
        po = pos[ref]

        if any(row.get(c) for c in COMPONENT_COLUMNS):
            errs.append("components are for subcontracting POs - use the Subcontracting PO "
                        "upload")

        item_no = pad("PurchaseOrderItem", row.get("PurchaseOrderItem")
                      or str(10 * (len(po["items"]) + 1)))
        new_item = item_no not in po["items"]
        if new_item:
            cat_in = row.get("PurchaseOrderItemCategory", "").strip().upper()
            cat = ITEM_CATEGORY_MAP.get(cat_in, cat_in)
            if cat == SUBCONTRACTING_CATEGORY:
                errs.append(f"item {item_no}: category {cat_in} is "
                            f"subcontracting - use the Subcontracting PO upload")
            elif cat not in VALID_INTERNAL_ITEM_CATEGORIES:
                errs.append(f"item {item_no}: unknown PurchaseOrderItemCategory '{cat_in}'")
            for f in ("Plant", "OrderQuantity", "PurchaseOrderQuantityUnit",
                      "ScheduleLineDeliveryDate"):
                if not row.get(f):
                    errs.append(f"{f} is required for item {item_no}")
            if not row.get("Material"):
                if not row.get("PurchaseOrderItemText"):
                    errs.append(f"item {item_no}: fill Material or PurchaseOrderItemText")
                if not row.get("MaterialGroup"):
                    errs.append(f"item {item_no}: MaterialGroup is required for a text item")
            item = ib.build(dict(row, PurchaseOrderItem=item_no,
                                 PurchaseOrderItemCategory=cat if cat_in else ""),
                            ITEM_FIELDS, errs)
            item["to_ScheduleLine"] = {"results": []}
            po["payload"]["to_PurchaseOrderItem"]["results"].append(item)
            po["items"][item_no] = {"item": item, "scheds": set(), "accts": set(), "row": row}
        it = po["items"][item_no]
        acct_cat = it["row"].get("AccountAssignmentCategory", "").strip().upper()

        # schedule lines: only rows that carry schedule data (or the item's first row)
        if new_item or any(row.get(c) for c in SCHED_FIELDS):
            sched_no = pad("ScheduleLine", row.get("ScheduleLine") or
                           str(len(it["scheds"]) + 1 if not new_item else 1))
            if sched_no in it["scheds"]:
                errs.append(f"item {item_no}: schedule line {sched_no} appears twice")
            else:
                if not new_item and not row.get("ScheduleLineDeliveryDate"):
                    errs.append(f"item {item_no}: ScheduleLineDeliveryDate is required for "
                                f"schedule line {sched_no}")
                sl = sb.build(dict(row, ScheduleLine=sched_no), SCHED_FIELDS, errs,
                              defaults={"ScheduleLineOrderQuantity":
                                        it["row"].get("OrderQuantity", "") if new_item else ""})
                if not new_item and "ScheduleLineOrderQuantity" not in sl:
                    errs.append(f"item {item_no}: ScheduleLineOrderQuantity is required for "
                                f"schedule line {sched_no}")
                it["item"]["to_ScheduleLine"]["results"].append(sl)
                it["scheds"].add(sched_no)

        # account assignments: one per item + AccountAssignmentNumber
        has_acct = any(row.get(c) for c in ACCT_FIELDS if c != "AccountAssignmentNumber")
        if has_acct and not acct_cat:
            errs.append(f"item {item_no}: account assignment columns are filled but "
                        f"AccountAssignmentCategory is empty")
        elif acct_cat and (has_acct or new_item):
            acct_no = pad("AccountAssignmentNumber", row.get("AccountAssignmentNumber")
                          or str(len(it["accts"]) + 1))
            if acct_no in it["accts"]:
                errs.append(f"item {item_no}: account assignment {acct_no} appears twice")
            else:
                need = ACCT_OBJECT.get(acct_cat, [])
                if need and not any(row.get(c) for c in need):
                    errs.append(f"item {item_no}: account assignment {acct_cat} needs "
                                f"{' / '.join(need)}")
                acct = ab.build(dict(row, AccountAssignmentNumber=acct_no), ACCT_FIELDS, errs)
                it["item"].setdefault("to_AccountAssignment", {"results": []})[
                    "results"].append(acct)
                it["accts"].add(acct_no)

        po["lines"].append(row["_line"])
        po["errors"] += [f"line {row['_line']}: {e}" for e in errs]

    docs = []
    for po in pos.values():
        for no, it in po["items"].items():
            accts = it["item"].get("to_AccountAssignment", {}).get("results", [])
            if len(accts) > 1 and not it["row"].get("MultipleAcctAssgmtDistribution"):
                po["errors"].append(f"item {no}: {len(accts)} account assignments - set "
                                    f"MultipleAcctAssgmtDistribution (1 = quantity, "
                                    f"2 = percentage)")
        docs.append({"ref": po["ref"], "lines": ",".join(map(str, po["lines"])),
                     "errors": po["errors"] + validate_payload(md, HEADER_TYPE, po["payload"]),
                     "payload": po["payload"]})
    return docs


if __name__ == "__main__":
    run_cli(sys.modules[__name__])