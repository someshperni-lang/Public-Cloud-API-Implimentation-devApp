#!/usr/bin/env python3
"""
Bulk Subcontracting Purchase Requisition upload (ME51N, item category L)
- API_PURCHASEREQUISITION_2 (OData V4) / PurchaseReqn, SAP DEV.

One row = one subcontracting PR item. Rows with the same PR_Ref form one
purchase requisition. Every item is sent with item category L (internal key 3).
The PR API takes no component list: SAP takes the components from the
material's BOM, so the material needs a BOM in the plant.

Standard, consumable, service and other PRs have their own script
(bulk_purchase_requisition_upload.py); this one reuses its builder and poster,
so both create PRs the same way.

    python bulk_subcontracting_pr_upload.py --template xlsx
    python bulk_subcontracting_pr_upload.py -f subcontracting_prs.xlsx --dry-run
    python bulk_subcontracting_pr_upload.py -f subcontracting_prs.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import bulk_purchase_requisition_upload as pr  # noqa: E402
from sap_common import run_cli  # noqa: E402

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
DESCRIPTION = ("Subcontracting purchase requisitions (item category L, OData V4). One row per "
               "item, grouped by PR_Ref; SAP takes the components from the material's BOM.")

count_items = pr.count_items
post_document = pr.post_document

ITEM_COLUMNS = ["PurchaseRequisitionItem", "PurchaseRequisitionItemText", "Material",
                "RequestedQuantity", "BaseUnit", "BaseUnitISOCode", "Plant", "StorageLocation",
                "PurchasingGroup", "PurchasingOrganization", "DeliveryDate", "Subcontractor",
                "FixedSupplier", "Supplier", "PurchasingInfoRecord", "PurchaseContract",
                "PurchaseContractItem", "PurchaseRequisitionPrice", "PurReqnPriceQuantity",
                "PurReqnItemCurrency", "TaxCode", "Batch", "RequisitionerName",
                "RequirementTracking", "GoodsReceiptIsExpected", "InvoiceIsExpected"]
COLUMNS = ["PR_Ref"] + list(pr.HEADER_FIELDS) + ITEM_COLUMNS + ["ItemText", "ItemTextLanguage"]

COLUMN_NOTES = {c: pr.COLUMN_NOTES.get(c, "") for c in COLUMNS}
COLUMN_NOTES.update({
    "PurchaseRequisitionType": "Required (first row of the PR), e.g. ZDOM or your "
                               "subcontracting PR type.",
    "Material": "Required. The finished / semi-finished material the subcontractor makes. "
                "It needs a BOM in the plant (components come from the BOM).",
    "RequestedQuantity": "Required. Quantity to be made, e.g. 100.",
    "BaseUnit": "Required (or BaseUnitISOCode), e.g. PC.",
    "Plant": "Required. Receiving plant.",
    "Subcontractor": "Recommended. The subcontractor (supplier) who gets the components.",
    "FixedSupplier": "Optional fixed supplier, usually the same as Subcontractor.",
    "PurchasingInfoRecord": "Optional subcontracting info record (price for the work).",
    "PurchaseRequisitionPrice": "Optional valuation price for the subcontracting work.",
})

SAMPLE_ROWS = [
    {"PR_Ref": "SPR1", "PurchaseRequisitionType": "ZDOM",
     "PurReqnDescription": "Subcontracting", "PurchaseRequisitionItemText": "Machining of FG228",
     "Material": "FG228", "RequestedQuantity": "100", "BaseUnit": "PC", "Plant": "01TG",
     "PurchasingGroup": "002", "DeliveryDate": "2026-11-30", "Subcontractor": "18300010"},
    {"PR_Ref": "SPR1", "Material": "FG229", "RequestedQuantity": "50", "BaseUnit": "PC",
     "Plant": "01TG", "PurchasingGroup": "002", "DeliveryDate": "2026-12-05",
     "Subcontractor": "18300010", "ItemText": "Return the scrap with the finished goods."},
]


def build_documents(df) -> list[dict]:
    return pr.build_prs(df, [pr.SUBCONTRACTING], pr.SUBCONTRACTING)


if __name__ == "__main__":
    run_cli(sys.modules[__name__])
