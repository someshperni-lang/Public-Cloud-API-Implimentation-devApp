#!/usr/bin/env python3
"""
Bulk Supplier Invoice upload (MIRO) - API_SUPPLIERINVOICE_PROCESS_SRV, SAP DEV system.

One row = one invoice line. Rows with the same InvoiceRef form one invoice;
header columns are read from the first row of each invoice.

ItemType decides what a row is:
  PO   - line referencing a purchase order item (to_SuplrInvcItemPurOrdRef)
  GL   - direct G/L account line              (to_SupplierInvoiceItemGLAcct)
  TAX  - manual tax line, only needed when TaxIsCalculatedAutomatically = FALSE
         (to_SupplierInvoiceTax)

    python bulk_supplier_invoice_upload.py --template xlsx
    python bulk_supplier_invoice_upload.py -f invoices.xlsx --dry-run
    python bulk_supplier_invoice_upload.py -f invoices.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sap_common import (EntityBuilder, load_metadata, service_url, rows_of, run_cli,  # noqa: E402
                        validate_payload)

TITLE = "Supplier invoices"
SLUG = "supplier_invoice"
SERVICE = "API_SUPPLIERINVOICE_PROCESS_SRV"
ENTITY_SET = "A_SupplierInvoice"
HEADER_TYPE = "A_SupplierInvoiceType"
PO_TYPE = "A_SuplrInvcItemPurOrdRefType"
GL_TYPE = "A_SupplierInvoiceItemGLAcctType"
TAX_TYPE = "A_SupplierInvoiceTaxType"
BASE_URL = service_url(SERVICE)          # DEV: my440263
DOC_FIELD = "SupplierInvoice"
YEAR_FIELD = "FiscalYear"
TCODE = "MIRO"
ICON = "🧾"
COLOR = "#EA580C"
GROUP_LABEL = "InvoiceRef"
ITEMS_LABEL = "Lines"


def count_items(p: dict) -> int:
    return len(p.get("to_SuplrInvcItemPurOrdRef", {}).get("results", [])) + len(p.get("to_SupplierInvoiceItemGLAcct", {}).get("results", []))
DESCRIPTION = ("Create supplier invoices with PO-reference lines, G/L lines and optional "
               "manual tax lines. One row per invoice line, grouped by InvoiceRef.")

HEADER_FIELDS = ["CompanyCode", "DocumentDate", "PostingDate", "InvoiceReceiptDate",
                 "SupplierInvoiceIDByInvcgParty", "InvoicingParty", "DocumentCurrency",
                 "InvoiceGrossAmount", "TaxIsCalculatedAutomatically", "PaymentTerms",
                 "DueCalculationBaseDate", "DocumentHeaderText", "AccountingDocumentType",
                 "SupplierInvoiceIsCreditMemo", "SupplierInvoiceStatus", "PaymentBlockingReason",
                 "PaymentMethod", "AssignmentReference", "SupplierPostingLineItemText",
                 "BusinessPlace", "BusinessSectionCode", "IN_GSTPartner", "IN_GSTPlaceOfSupply",
                 "IN_InvoiceReferenceNumber", "TaxDeterminationDate"]

# sheet column -> SAP field, per line type
PO_MAP = {"SupplierInvoiceItem": "InvoiceItem", "PurchaseOrder": "PurchaseOrder",
          "PurchaseOrderItem": "PurchaseOrderItem", "Plant": "Plant",
          "ReferenceDocument": "ReferenceDocument",
          "ReferenceDocumentFiscalYear": "ReferenceDocumentFiscalYear",
          "ReferenceDocumentItem": "ReferenceDocumentItem", "TaxCode": "TaxCode",
          "DocumentCurrency": "DocumentCurrency", "SupplierInvoiceItemAmount": "ItemAmount",
          "PurchaseOrderQuantityUnit": "QuantityUnit",
          "QuantityInPurchaseOrderUnit": "Quantity", "SupplierInvoiceItemText": "ItemText",
          "IsSubsequentDebitCredit": "IsSubsequentDebitCredit",
          "IsFinallyInvoiced": "IsFinallyInvoiced", "IN_HSNOrSACCode": "HSNOrSACCode",
          "ServiceEntrySheet": "ServiceEntrySheet", "ServiceEntrySheetItem": "ServiceEntrySheetItem"}
GL_MAP = {"SupplierInvoiceItem": "InvoiceItem", "CompanyCode": "CompanyCode",
          "GLAccount": "GLAccount", "CostCenter": "CostCenter", "ProfitCenter": "ProfitCenter",
          "InternalOrder": "InternalOrder", "WBSElement": "WBSElement",
          "DebitCreditCode": "DebitCreditCode", "DocumentCurrency": "DocumentCurrency",
          "SupplierInvoiceItemAmount": "ItemAmount", "TaxCode": "TaxCode",
          "SupplierInvoiceItemText": "ItemText", "AssignmentReference": "AssignmentReference",
          "Quantity": "Quantity", "QuantityUnit": "QuantityUnit"}
TAX_MAP = {"TaxCode": "TaxCode", "DocumentCurrency": "DocumentCurrency",
           "TaxAmount": "TaxAmount", "TaxBaseAmountInTransCrcy": "TaxBaseAmount"}

ITEM_COLUMNS = ["ItemType", "InvoiceItem", "PurchaseOrder", "PurchaseOrderItem", "Plant",
                "ReferenceDocument", "ReferenceDocumentFiscalYear", "ReferenceDocumentItem",
                "Quantity", "QuantityUnit", "GLAccount", "CostCenter", "ProfitCenter",
                "InternalOrder", "WBSElement", "DebitCreditCode", "ItemAmount", "TaxCode",
                "ItemText", "HSNOrSACCode", "IsSubsequentDebitCredit", "IsFinallyInvoiced",
                "ServiceEntrySheet", "ServiceEntrySheetItem", "TaxAmount", "TaxBaseAmount"]
COLUMNS = ["InvoiceRef"] + HEADER_FIELDS + ITEM_COLUMNS

COLUMN_NOTES = {
    "InvoiceRef": "Required. Your own key - all rows with the same value form one invoice.",
    "CompanyCode": "Required (first row of the invoice).",
    "DocumentDate": "Required. Invoice date.",
    "PostingDate": "Required.",
    "SupplierInvoiceIDByInvcgParty": "Supplier's invoice number (reference).",
    "InvoicingParty": "Required. Supplier number.",
    "DocumentCurrency": "Required, e.g. INR. Copied to every line.",
    "InvoiceGrossAmount": "Required. Gross amount including tax.",
    "TaxIsCalculatedAutomatically": "TRUE = SAP calculates tax; FALSE = add TAX rows.",
    "SupplierInvoiceIsCreditMemo": "X for a credit memo, blank for an invoice.",
    "SupplierInvoiceStatus": "Leave blank to post. Use the status your system expects to park.",
    "BusinessPlace": "India: GST business place.",
    "IN_GSTPartner": "India: GST partner (usually the supplier).",
    "IN_GSTPlaceOfSupply": "India: place of supply (state code, e.g. 36).",
    "ItemType": "Required. PO, GL or TAX.",
    "InvoiceItem": "Optional line number; numbered 1,2,3 per invoice when blank.",
    "PurchaseOrder": "PO lines: purchase order.",
    "PurchaseOrderItem": "PO lines: PO item, e.g. 10.",
    "ReferenceDocument": "PO lines with GR-based IV: material document of the GR.",
    "ReferenceDocumentFiscalYear": "PO lines with GR-based IV: GR document year.",
    "ReferenceDocumentItem": "PO lines with GR-based IV: GR document item, e.g. 1.",
    "Quantity": "PO lines: invoiced quantity. GL lines: optional quantity.",
    "QuantityUnit": "Unit of the quantity, e.g. PC.",
    "GLAccount": "GL lines: G/L account.",
    "CostCenter": "GL lines: cost center (or ProfitCenter / InternalOrder / WBSElement).",
    "DebitCreditCode": "GL lines: S = debit, H = credit. Defaults to S.",
    "ItemAmount": "PO and GL lines: net line amount.",
    "TaxCode": "PO / GL / TAX lines: tax code, e.g. G4.",
    "HSNOrSACCode": "PO lines, India: HSN / SAC code.",
    "TaxAmount": "TAX lines: tax amount.",
    "TaxBaseAmount": "TAX lines: base amount.",
}

SAMPLE_ROWS = [
    {"InvoiceRef": "INV1", "CompanyCode": "1810", "DocumentDate": "2026-10-08",
     "PostingDate": "2026-10-08", "SupplierInvoiceIDByInvcgParty": "SUP-INV-001",
     "InvoicingParty": "18300010", "DocumentCurrency": "INR", "InvoiceGrossAmount": "2950.00",
     "TaxIsCalculatedAutomatically": "TRUE", "ItemType": "PO", "PurchaseOrder": "4500000001",
     "PurchaseOrderItem": "10", "Plant": "1810", "ReferenceDocument": "5000000001",
     "ReferenceDocumentFiscalYear": "2026", "ReferenceDocumentItem": "1", "Quantity": "10",
     "QuantityUnit": "PC", "ItemAmount": "2500.00", "TaxCode": "G4"},
    {"InvoiceRef": "INV2", "CompanyCode": "1810", "DocumentDate": "2026-10-08",
     "PostingDate": "2026-10-08", "SupplierInvoiceIDByInvcgParty": "SUP-INV-002",
     "InvoicingParty": "18300010", "DocumentCurrency": "INR", "InvoiceGrossAmount": "1180.00",
     "TaxIsCalculatedAutomatically": "TRUE", "ItemType": "GL", "GLAccount": "61002000",
     "CostCenter": "18101101", "DebitCreditCode": "S", "ItemAmount": "1000.00",
     "TaxCode": "G4", "ItemText": "Freight charges"},
]


def build_documents(df) -> list[dict]:
    md = load_metadata(SERVICE)
    hb = EntityBuilder(md, HEADER_TYPE)
    builders = {"PO": (EntityBuilder(md, PO_TYPE), PO_MAP, "to_SuplrInvcItemPurOrdRef"),
                "GL": (EntityBuilder(md, GL_TYPE), GL_MAP, "to_SupplierInvoiceItemGLAcct"),
                "TAX": (EntityBuilder(md, TAX_TYPE), TAX_MAP, "to_SupplierInvoiceTax")}
    invoices: dict[str, dict] = {}

    for row in rows_of(df):
        errs: list[str] = []
        ref = row.get("InvoiceRef") or f"LINE{row['_line']}"
        if ref not in invoices:
            for f in ("CompanyCode", "DocumentDate", "PostingDate", "InvoicingParty",
                      "DocumentCurrency", "InvoiceGrossAmount"):
                if not row.get(f):
                    errs.append(f"{f} is required on the first row of invoice {ref}")
            header = hb.build(row, HEADER_FIELDS, errs)
            invoices[ref] = {"ref": ref, "lines": [], "errors": [], "payload": header,
                             "currency": row.get("DocumentCurrency", ""),
                             "company": row.get("CompanyCode", ""), "counter": {"PO": 0, "GL": 0}}
        inv = invoices[ref]
        itype = row.get("ItemType", "").upper()
        if itype not in builders:
            errs.append(f"ItemType '{row.get('ItemType', '')}' must be PO, GL or TAX")
        else:
            b, mapping, nav = builders[itype]
            need = {"PO": ["PurchaseOrder", "PurchaseOrderItem", "ItemAmount"],
                    "GL": ["GLAccount", "ItemAmount"],
                    "TAX": ["TaxCode", "TaxAmount"]}[itype]
            for f in need:
                if not row.get(f):
                    errs.append(f"{f} is required for {itype} lines")
            defaults = {"DocumentCurrency": inv["currency"]}
            if itype in ("PO", "GL"):
                inv["counter"][itype] += 1
                n = inv["counter"][itype]
                width = 6 if itype == "PO" else 4
                defaults["SupplierInvoiceItem"] = str(n).zfill(width)
                if itype == "GL":
                    defaults.update({"CompanyCode": inv["company"], "DebitCreditCode": "S"})
            line = b.build(row, mapping, errs, defaults=defaults)
            if "SupplierInvoiceItem" in line:
                width = 6 if itype == "PO" else 4
                line["SupplierInvoiceItem"] = line["SupplierInvoiceItem"].zfill(width)
            inv["payload"].setdefault(nav, {"results": []})["results"].append(line)
        inv["lines"].append(row["_line"])
        inv["errors"] += [f"line {row['_line']}: {e}" for e in errs]

    docs = []
    for inv in invoices.values():
        p = inv["payload"]
        if not (p.get("to_SuplrInvcItemPurOrdRef") or p.get("to_SupplierInvoiceItemGLAcct")):
            inv["errors"].append("invoice has no PO or GL lines")
        if p.get("TaxIsCalculatedAutomatically") is False and not p.get("to_SupplierInvoiceTax"):
            inv["errors"].append("TaxIsCalculatedAutomatically is FALSE but there are no TAX rows")
        inv["errors"] += validate_payload(md, HEADER_TYPE, p)
        docs.append({"ref": inv["ref"], "lines": ",".join(map(str, inv["lines"])),
                     "errors": inv["errors"], "payload": p})
    return docs


if __name__ == "__main__":
    run_cli(sys.modules[__name__])
