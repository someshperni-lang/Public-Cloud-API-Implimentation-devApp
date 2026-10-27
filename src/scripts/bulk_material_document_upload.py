#!/usr/bin/env python3
"""
Bulk Goods Movement upload (MIGO) - API_MATERIAL_DOCUMENT_SRV, SAP DEV system.

One row = one material document item. Rows are grouped into one document when
they share DocGroup (or PO / production order / reservation / delivery),
GoodsMovementCode, GoodsMovementType, DocumentDate and PostingDate.

GoodsMovementCode and GoodsMovementRefDocType are derived from the movement
type (MOVEMENT_TYPE_RULES) unless you fill them in the row yourself.

    python bulk_material_document_upload.py --template xlsx
    python bulk_material_document_upload.py -f goods_movements.xlsx --dry-run
    python bulk_material_document_upload.py -f goods_movements.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sap_common import (EntityBuilder, load_metadata, service_url, rows_of, run_cli,  # noqa: E402
                        validate_payload)

TITLE = "Goods movements"
SLUG = "material_document"
SERVICE = "API_MATERIAL_DOCUMENT_SRV"
ENTITY_SET = "A_MaterialDocumentHeader"
HEADER_TYPE = "A_MaterialDocumentHeaderType"
ITEM_TYPE = "A_MaterialDocumentItemType"
BASE_URL = service_url(SERVICE)          # DEV: my440263
DOC_FIELD = "MaterialDocument"
YEAR_FIELD = "MaterialDocumentYear"
TCODE = "MIGO"
ICON = "📦"
COLOR = "#0D9488"
GROUP_LABEL = "DocGroup / PO"
ITEMS_LABEL = "Items"


def count_items(p: dict) -> int:
    return len(p.get("to_MaterialDocumentItem", {}).get("results", []))
DESCRIPTION = ("Post goods receipts, goods issues, transfer postings and other receipts. "
               "One row per item; the movement code is derived from the movement type.")

# movement type -> (GoodsMovementCode, GoodsMovementRefDocType, required fields)
MOVEMENT_TYPE_RULES = {
    # 01 - GR for purchase order
    "101": ("01", "B", ["PurchaseOrder", "PurchaseOrderItem"]),
    "102": ("01", "B", ["PurchaseOrder", "PurchaseOrderItem"]),
    "122": ("01", "B", ["PurchaseOrder", "PurchaseOrderItem"]),
    "123": ("01", "B", ["PurchaseOrder", "PurchaseOrderItem"]),
    "161": ("01", "B", ["PurchaseOrder", "PurchaseOrderItem"]),
    "162": ("01", "B", ["PurchaseOrder", "PurchaseOrderItem"]),
    # 03 - goods issue
    "201": ("03", "", ["Material", "CostCenter"]),
    "202": ("03", "", ["Material", "CostCenter"]),
    "221": ("03", "", ["Material", "WBSElement"]),
    "222": ("03", "", ["Material", "WBSElement"]),
    "231": ("03", "", ["Material", "SalesOrder", "SalesOrderItem"]),
    "232": ("03", "", ["Material", "SalesOrder", "SalesOrderItem"]),
    "241": ("03", "", ["Material", "MasterFixedAsset"]),
    "242": ("03", "", ["Material", "MasterFixedAsset"]),
    "261": ("03", "", ["Material", "ManufacturingOrder"]),
    "262": ("03", "", ["Material", "ManufacturingOrder"]),
    "551": ("03", "", ["Material"]),
    "552": ("03", "", ["Material"]),
    # 04 - transfer posting
    "301": ("04", "", ["Material", "IssuingOrReceivingPlant"]),
    "302": ("04", "", ["Material", "IssuingOrReceivingPlant"]),
    "309": ("04", "", ["Material", "IssgOrRcvgMaterial"]),
    "310": ("04", "", ["Material", "IssgOrRcvgMaterial"]),
    "311": ("04", "", ["Material", "IssuingOrReceivingStorageLoc"]),
    "312": ("04", "", ["Material", "IssuingOrReceivingStorageLoc"]),
    "321": ("04", "", ["Material"]),
    "322": ("04", "", ["Material"]),
    "343": ("04", "", ["Material"]),
    "344": ("04", "", ["Material"]),
    "541": ("04", "", ["Material", "Supplier"]),
    "542": ("04", "", ["Material", "Supplier"]),
    # 05 - other goods receipts
    "501": ("05", "", ["Material"]),
    "502": ("05", "", ["Material"]),
    "511": ("05", "", ["Material", "Supplier"]),
    "512": ("05", "", ["Material", "Supplier"]),
    "521": ("05", "", ["Material"]),
    "522": ("05", "", ["Material"]),
    "561": ("05", "", ["Material"]),
    "562": ("05", "", ["Material"]),
}

HEADER_FIELDS = {"DocumentDate": "DocumentDate", "PostingDate": "PostingDate",
                 "MaterialDocumentHeaderText": "HeaderText",
                 "ReferenceDocument": "ReferenceDocument", "GoodsMovementCode": "GoodsMovementCode"}

ITEM_FIELDS = ["Material", "Plant", "StorageLocation", "Batch", "GoodsMovementType",
               "InventoryStockType", "InventorySpecialStockType", "QuantityInEntryUnit",
               "EntryUnit", "PurchaseOrder", "PurchaseOrderItem", "GoodsMovementRefDocType",
               "ManufacturingOrder", "ManufacturingOrderItem", "Reservation", "ReservationItem",
               "Delivery", "DeliveryItem", "CostCenter", "GLAccount", "WBSElement", "SalesOrder",
               "SalesOrderItem", "MasterFixedAsset", "FixedAsset", "Supplier", "Customer",
               "IssuingOrReceivingPlant", "IssuingOrReceivingStorageLoc", "IssgOrRcvgMaterial",
               "IssgOrRcvgBatch", "GoodsMovementReasonCode", "MaterialDocumentItemText",
               "GoodsRecipientName", "UnloadingPointName", "ManufactureDate",
               "ShelfLifeExpirationDate", "IsCompletelyDelivered"]

COLUMNS = (["DocGroup", "GoodsMovementType", "GoodsMovementCode", "DocumentDate", "PostingDate",
            "HeaderText", "ReferenceDocument"]
           + [f for f in ITEM_FIELDS if f != "GoodsMovementType"] + ["SerialNumbers"])

COLUMN_NOTES = {
    "DocGroup": "Optional. Rows with the same value go into one material document.",
    "GoodsMovementType": "Required. 101, 201, 261, 311, 541, 501 ... (see MOVEMENT_TYPE_RULES).",
    "GoodsMovementCode": "Optional. Leave blank - derived from movement type (01/02/03/04/05).",
    "DocumentDate": "Required. YYYY-MM-DD or DD.MM.YYYY.",
    "PostingDate": "Required. YYYY-MM-DD or DD.MM.YYYY.",
    "HeaderText": "Optional header text (max 25).",
    "ReferenceDocument": "Optional. Delivery note / bill of lading (max 16).",
    "Material": "Required except for 101 against a PO (taken from the PO item).",
    "Plant": "Required.",
    "StorageLocation": "Usually required.",
    "QuantityInEntryUnit": "Required. Quantity.",
    "EntryUnit": "Required. Unit of entry, e.g. PC, KG.",
    "PurchaseOrder": "101/102/122/161: purchase order number.",
    "PurchaseOrderItem": "101/102/122/161: PO item, e.g. 10 (padded to 00010).",
    "GoodsMovementRefDocType": "Optional. Derived: B for PO receipts, F for production order receipts.",
    "ManufacturingOrder": "261/262 and 101 against a production order.",
    "CostCenter": "201/202.",
    "WBSElement": "221/222.",
    "SalesOrder": "231/232.",
    "MasterFixedAsset": "241/242.",
    "Supplier": "541/542, 511/512.",
    "IssuingOrReceivingPlant": "301/302: receiving plant.",
    "IssuingOrReceivingStorageLoc": "311/312: receiving storage location.",
    "IssgOrRcvgMaterial": "309/310: receiving material.",
    "GoodsMovementReasonCode": "Required by some movement types (e.g. 122 return).",
    "SerialNumbers": "Optional. Several serial numbers separated by ;",
}

SAMPLE_ROWS = [
    {"GoodsMovementType": "101", "DocumentDate": "2026-10-08", "PostingDate": "2026-10-08",
     "HeaderText": "GR for PO", "ReferenceDocument": "DN-1001", "Plant": "1810",
     "StorageLocation": "181A", "QuantityInEntryUnit": "10", "EntryUnit": "KG",
     "PurchaseOrder": "4500000001", "PurchaseOrderItem": "10"},
    {"GoodsMovementType": "201", "DocumentDate": "2026-10-08", "PostingDate": "2026-10-08",
     "HeaderText": "GI to cost center", "Material": "RM019", "Plant": "1810",
     "StorageLocation": "181A", "QuantityInEntryUnit": "2", "EntryUnit": "KG",
     "CostCenter": "18101101"},
    {"DocGroup": "SLOC-1", "GoodsMovementType": "311", "DocumentDate": "2026-10-08",
     "PostingDate": "2026-10-08", "HeaderText": "SLoc to SLoc", "Material": "RM019",
     "Plant": "1810", "StorageLocation": "181A", "QuantityInEntryUnit": "3", "EntryUnit": "KG",
     "IssuingOrReceivingStorageLoc": "181B"},
]


def _rule(row):
    mvt = row.get("GoodsMovementType", "")
    code, ref, req = MOVEMENT_TYPE_RULES.get(mvt, ("", "", ["Material"]))
    if mvt in ("101", "102") and not row.get("PurchaseOrder") and row.get("ManufacturingOrder"):
        code, ref, req = "02", "F", ["ManufacturingOrder"]
    return code, ref, req


def build_documents(df) -> list[dict]:
    md = load_metadata(SERVICE)
    hb, ib = EntityBuilder(md, HEADER_TYPE), EntityBuilder(md, ITEM_TYPE)
    groups: dict[str, dict] = {}

    for row in rows_of(df):
        errs: list[str] = []
        mvt = row.get("GoodsMovementType", "")
        code, ref, required = _rule(row)
        code = row.get("GoodsMovementCode") or code
        ref = row.get("GoodsMovementRefDocType") or ref
        if not mvt:
            errs.append("GoodsMovementType is empty")
        elif not code:
            errs.append(f"Movement type {mvt} is not in MOVEMENT_TYPE_RULES - fill GoodsMovementCode")
        for f in required + ["Plant", "QuantityInEntryUnit", "EntryUnit", "DocumentDate", "PostingDate"]:
            if not row.get(f):
                errs.append(f"{f} is required for movement type {mvt}")

        item = ib.build(row, ITEM_FIELDS, errs,
                        defaults={"GoodsMovementRefDocType": ref} if ref else None)
        serials = [s.strip() for s in row.get("SerialNumbers", "").replace(",", ";").split(";")
                   if s.strip()]
        if serials:
            item["to_SerialNumbers"] = {"results": [{"SerialNumber": s} for s in serials]}

        grp_ref = (row.get("DocGroup") or row.get("PurchaseOrder") or row.get("ManufacturingOrder")
                   or row.get("Reservation") or row.get("Delivery") or f"LINE{row['_line']}")
        key = "|".join([grp_ref, code, mvt, row.get("DocumentDate", ""), row.get("PostingDate", "")])
        if key not in groups:
            hrow = dict(row, GoodsMovementCode=code)
            header = hb.build(hrow, HEADER_FIELDS, errs)
            groups[key] = {"ref": f"{grp_ref}-{mvt}", "lines": [], "errors": [],
                           "payload": {**header, "to_MaterialDocumentItem": {"results": []}}}
        g = groups[key]
        g["lines"].append(row["_line"])
        g["errors"] += [f"line {row['_line']}: {e}" for e in errs]
        g["payload"]["to_MaterialDocumentItem"]["results"].append(item)

    docs = []
    for g in groups.values():
        g["errors"] += validate_payload(md, HEADER_TYPE, g["payload"])
        g["lines"] = ",".join(map(str, g["lines"]))
        docs.append(g)
    return docs


if __name__ == "__main__":
    run_cli(sys.modules[__name__])
