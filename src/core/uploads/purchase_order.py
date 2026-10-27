"""
Purchase orders (ME21N) - API_PURCHASEORDER_PROCESS_SRV -> A_PurchaseOrder.

Standard, account-assigned (K, F, P, A ...) and subcontracting (item
category 3) POs from one flat sheet: PO_Ref -> item -> schedule line ->
subcontracting components.
"""
from __future__ import annotations

from ..sap_common import EntityBuilder, load_metadata, pad, rows_of, validate_payload

TITLE = "Purchase orders"
SLUG = "purchase_order"
SERVICE = "API_PURCHASEORDER_PROCESS_SRV"
ENTITY_SET = "A_PurchaseOrder"
HEADER_TYPE = "A_PurchaseOrderType"
ITEM_TYPE = "A_PurchaseOrderItemType"
SCHED_TYPE = "A_PurchaseOrderScheduleLineType"
COMP_TYPE = "A_POSubcontractingComponentType"
ACCT_TYPE = "A_PurOrdAccountAssignmentType"
RESULT_KEYS = ["PurchaseOrder"]
DESCRIPTION = ("Standard, account-assigned and subcontracting POs. "
               "One flat sheet, grouped by PO_Ref, item and schedule line.")

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
COMP_MAP = {"Material": "ComponentMaterial", "RequirementDate": "ComponentRequirementDate",
            "QuantityInEntryUnit": "ComponentQuantity", "EntryUnit": "ComponentUnit",
            "Plant": "ComponentPlant"}

COLUMNS = (["PO_Ref"] + HEADER_FIELDS + ITEM_FIELDS + list(SCHED_MAP.values())
           + ["AccountAssignmentNumber", "CostCenter", "GLAccount", "WBSElement", "OrderID",
              "MasterFixedAsset", "FixedAsset", "AcctProfitCenter", "AcctQuantity"]
           + ["ComponentMaterial", "ComponentRequirementDate", "ComponentQuantity",
              "ComponentUnit", "ComponentPlant"])

COLUMN_NOTES = {
    "PO_Ref": "Required. Your own key - rows with the same value form one PO.",
    "CompanyCode": "Required (first row of the PO).",
    "PurchaseOrderType": "Required, e.g. NB (standard) or ZSUB (subcontracting).",
    "Supplier": "Required.",
    "PurchasingOrganization": "Required.",
    "PurchasingGroup": "Required, e.g. 001.",
    "DocumentCurrency": "Required, e.g. INR.",
    "PurchaseOrderItem": "Item number, e.g. 10. Defaults to 10, 20, 30...",
    "Material": "Material, or leave blank and fill PurchaseOrderItemText + MaterialGroup.",
    "Plant": "Required.",
    "OrderQuantity": "Required.",
    "PurchaseOrderQuantityUnit": "Required, e.g. PC.",
    "NetPriceAmount": "Net price.",
    "TaxCode": "Tax code, e.g. G4.",
    "PurchaseOrderItemCategory": "Blank = standard, 3 = subcontracting, 2 = consignment.",
    "AccountAssignmentCategory": "K = cost center, F = order, P = project, A = asset.",
    "GoodsReceiptIsExpected": "TRUE / FALSE.",
    "InvoiceIsExpected": "TRUE / FALSE.",
    "ScheduleLine": "Defaults to 0001.",
    "DeliveryDate": "Required. Delivery date of the schedule line.",
    "ScheduleLineQuantity": "Defaults to OrderQuantity.",
    "AccountAssignmentNumber": "Defaults to 01 when an account assignment is filled.",
    "CostCenter": "Account assignment K.",
    "GLAccount": "G/L account for account-assigned items.",
    "ComponentMaterial": "Subcontracting only: component to provide. One row per component.",
    "ComponentRequirementDate": "Subcontracting: defaults to DeliveryDate.",
    "ComponentQuantity": "Subcontracting: component quantity.",
    "ComponentUnit": "Subcontracting: component unit, e.g. KG.",
    "ComponentPlant": "Subcontracting: defaults to item Plant.",
}

SAMPLE_ROWS = [
    {"PO_Ref": "PO1", "CompanyCode": "1810", "PurchaseOrderType": "NB", "Supplier": "18300010",
     "PurchasingOrganization": "1810", "PurchasingGroup": "001", "DocumentCurrency": "INR",
     "PurchaseOrderItem": "10", "Material": "RM019", "Plant": "1810", "OrderQuantity": "50",
     "PurchaseOrderQuantityUnit": "KG", "NetPriceAmount": "120.00", "NetPriceQuantity": "1",
     "TaxCode": "G4", "GoodsReceiptIsExpected": "TRUE", "InvoiceIsExpected": "TRUE",
     "DeliveryDate": "2026-10-20"},
    {"PO_Ref": "PO2", "CompanyCode": "1810", "PurchaseOrderType": "ZSUB", "Supplier": "18300010",
     "PurchasingOrganization": "1810", "PurchasingGroup": "001", "DocumentCurrency": "INR",
     "PurchaseOrderItem": "10", "Material": "FG228", "Plant": "1810", "OrderQuantity": "100",
     "PurchaseOrderQuantityUnit": "PC", "NetPriceAmount": "250.00", "NetPriceQuantity": "1",
     "TaxCode": "G4", "PurchaseOrderItemCategory": "3", "GoodsReceiptIsExpected": "TRUE",
     "InvoiceIsExpected": "TRUE", "ScheduleLine": "1", "DeliveryDate": "2026-10-20",
     "ScheduleLineQuantity": "100", "ComponentMaterial": "RM019", "ComponentQuantity": "100",
     "ComponentUnit": "KG"},
    {"PO_Ref": "PO3", "CompanyCode": "1810", "PurchaseOrderType": "NB", "Supplier": "18300010",
     "PurchasingOrganization": "1810", "PurchasingGroup": "001", "DocumentCurrency": "INR",
     "PurchaseOrderItem": "10", "PurchaseOrderItemText": "Office chairs", "MaterialGroup": "L001",
     "Plant": "1810", "OrderQuantity": "5", "PurchaseOrderQuantityUnit": "EA",
     "NetPriceAmount": "4500.00", "NetPriceQuantity": "1", "TaxCode": "G4",
     "AccountAssignmentCategory": "K", "GoodsReceiptIsExpected": "TRUE",
     "InvoiceIsExpected": "TRUE", "DeliveryDate": "2026-10-25", "CostCenter": "18101101",
     "GLAccount": "61002000"},
]


def build_documents(table) -> list[dict]:
    md = load_metadata(SERVICE)
    hb, ib = EntityBuilder(md, HEADER_TYPE), EntityBuilder(md, ITEM_TYPE)
    sb, ab = EntityBuilder(md, SCHED_TYPE), EntityBuilder(md, ACCT_TYPE)
    cb = EntityBuilder(md, COMP_TYPE)
    pos: dict[str, dict] = {}

    for row in rows_of(table):
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
            for f in ("Plant", "OrderQuantity", "PurchaseOrderQuantityUnit", "DeliveryDate"):
                if not row.get(f):
                    errs.append(f"{f} is required for item {item_no}")
            if not row.get("Material") and not row.get("PurchaseOrderItemText"):
                errs.append(f"item {item_no}: fill Material or PurchaseOrderItemText")
            item = ib.build(dict(row, PurchaseOrderItem=item_no), ITEM_FIELDS, errs)
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
            po["items"][item_no] = {"item": item, "scheds": {}, "row": row}
        it = po["items"][item_no]

        sched_no = pad("ScheduleLine", row.get("ScheduleLine") or "1")
        if sched_no not in it["scheds"]:
            sl = sb.build(dict(row, ScheduleLine=sched_no), SCHED_MAP, errs,
                          defaults={"ScheduleLineOrderQuantity": row.get("OrderQuantity")
                                    or it["row"].get("OrderQuantity", "")})
            it["item"]["to_ScheduleLine"]["results"].append(sl)
            it["scheds"][sched_no] = sl
        sl = it["scheds"][sched_no]

        if row.get("ComponentMaterial"):
            for f in ("ComponentQuantity", "ComponentUnit"):
                if not row.get(f):
                    errs.append(f"{f} is required for component {row['ComponentMaterial']}")
            if it["item"].get("PurchaseOrderItemCategory", "") != "3":
                errs.append(f"item {item_no}: components need PurchaseOrderItemCategory 3")
            comp = cb.build(row, COMP_MAP, errs, defaults={
                "RequirementDate": row.get("DeliveryDate") or it["row"].get("DeliveryDate", ""),
                "Plant": row.get("Plant") or it["row"].get("Plant", "")})
            comps = sl.setdefault("to_SubcontractingComponent", {"results": []})["results"]
            comps.append({"ReservationItem": str(len(comps) + 1).zfill(4), **comp})

        po["lines"].append(row["_line"])
        po["errors"] += [f"line {row['_line']}: {e}" for e in errs]

    docs = []
    for po in pos.values():
        p = po["payload"]
        docs.append({"ref": po["ref"], "lines": ",".join(map(str, po["lines"])),
                     "errors": po["errors"] + validate_payload(md, HEADER_TYPE, p),
                     "payload": p})
    return docs
