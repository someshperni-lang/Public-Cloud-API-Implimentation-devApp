#!/usr/bin/env python3
"""
Bulk Supplier Invoice upload (MIRO) - API_SUPPLIERINVOICE_PROCESS_SRV, SAP DEV system.

Field set and grouping follow the reference script (InvoiceRefID version):
one row = one PO-reference invoice line. Rows with the same InvoiceRefID form
one invoice (deep insert into A_SupplierInvoice); header columns are read
from the first row of each invoice, line columns from every row, nested under
to_SuplrInvcItemPurOrdRef.

- CompanyCode is sent on the header only (PO lines use Plant).
- DocumentCurrency of a line = ItemDocumentCurrency, else the header currency.
- RetentionDueDate and TaxDeterminationDate fill both the header (first row)
  and each line (its own row), as in the reference script.
- Dates: DD.MM.YYYY (e.g. 16.04.2026) or YYYY-MM-DD.

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
BASE_URL = service_url(SERVICE)          # DEV: my440263
DOC_FIELD = "SupplierInvoice"
YEAR_FIELD = "FiscalYear"
TCODE = "MIRO"
ICON = "🧾"
COLOR = "#EA580C"
GROUP_LABEL = "InvoiceRefID"
ITEMS_LABEL = "Lines"
ITEM_NAV = "to_SuplrInvcItemPurOrdRef"
DESCRIPTION = ("Create supplier invoices against purchase orders. One row per PO invoice "
               "line, grouped by InvoiceRefID; header fields come from each invoice's first row.")


def count_items(p: dict) -> int:
    return len(p.get(ITEM_NAV, {}).get("results", []))


# ---------------------------------------------------------------------------
# Fields - same names and order as the reference script
# ---------------------------------------------------------------------------
GROUP_COLUMN = "InvoiceRefID"
HEADER_FIELDS = [
    "CompanyCode", "DocumentDate", "PostingDate", "SupplierInvoiceIDByInvcgParty",
    "InvoicingParty", "DocumentCurrency", "InvoiceGrossAmount", "UnplannedDeliveryCost",
    "DocumentHeaderText", "ManualCashDiscount", "PaymentTerms", "DueCalculationBaseDate",
    "CashDiscount1Percent", "CashDiscount1Days", "CashDiscount2Percent", "CashDiscount2Days",
    "NetPaymentDays", "PaymentBlockingReason", "AccountingDocumentType", "SupplierInvoiceStatus",
    "PaymentMethod", "PaymentReference", "InvoiceReference", "InvoiceReferenceFiscalYear",
    "TaxIsCalculatedAutomatically", "BusinessPlace", "BusinessArea",
    "SuplrInvcIsCapitalGoodsRelated", "SupplierInvoiceIsCreditMemo", "TaxDeterminationDate",
    "TaxReportingDate", "TaxFulfillmentDate", "InvoiceReceiptDate", "SupplierVATRegistration",
    "RetentionDueDate", "HouseBank", "HouseBankAccount", "IN_GSTPartner", "IN_GSTPlaceOfSupply",
    "IN_InvoiceReferenceNumber"]
ITEM_FIELDS = [
    "SupplierInvoiceItem", "PurchaseOrder", "PurchaseOrderItem", "Plant", "ReferenceDocument",
    "ReferenceDocumentFiscalYear", "ReferenceDocumentItem", "IsSubsequentDebitCredit", "TaxCode",
    "TaxJurisdiction", "SupplierInvoiceItemAmount", "PurchaseOrderQuantityUnit",
    "QuantityInPurchaseOrderUnit", "PurchaseOrderPriceUnit", "QtyInPurchaseOrderPriceUnit",
    "SupplierInvoiceItemText", "FreightSupplier", "IsNotCashDiscountLiable",
    "RetentionAmountInDocCurrency", "RetentionPercentage", "RetentionDueDate",
    "ServiceEntrySheet", "ServiceEntrySheetItem", "TaxCountry", "IsFinallyInvoiced",
    "TaxDeterminationDate", "IN_HSNOrSACCode", "IN_CustomDutyAssessableValue"]
ITEM_CURRENCY_COLUMN = "ItemDocumentCurrency"

# sheet columns: header fields, then line fields not already listed, then line currency
COLUMNS = ([GROUP_COLUMN] + HEADER_FIELDS
           + [f for f in ITEM_FIELDS if f not in HEADER_FIELDS] + [ITEM_CURRENCY_COLUMN])

HEADER_REQUIRED = ["CompanyCode", "DocumentDate", "PostingDate", "InvoicingParty",
                   "DocumentCurrency", "InvoiceGrossAmount"]
ITEM_REQUIRED = ["PurchaseOrder", "PurchaseOrderItem", "SupplierInvoiceItemAmount"]

_NOTES = {
    "InvoiceRefID": "Required. Your own key - all rows with the same value form one invoice.",
    "CompanyCode": "Required (first row of the invoice). Header only - lines use Plant.",
    "DocumentDate": "Required. Invoice date, DD.MM.YYYY or YYYY-MM-DD.",
    "PostingDate": "Required. DD.MM.YYYY or YYYY-MM-DD.",
    "SupplierInvoiceIDByInvcgParty": "Supplier's invoice number (reference).",
    "InvoicingParty": "Required. Supplier number.",
    "DocumentCurrency": "Required, e.g. INR. Used on every line unless ItemDocumentCurrency "
                        "is filled.",
    "InvoiceGrossAmount": "Required. Gross amount including tax.",
    "UnplannedDeliveryCost": "Optional unplanned delivery costs (e.g. freight not on the PO).",
    "ManualCashDiscount": "Optional cash discount amount entered manually.",
    "PaymentTerms": "Optional, e.g. 0001. Blank = from the supplier / PO.",
    "CashDiscount1Percent": "Optional. Overrides the payment terms: discount 1 percentage.",
    "CashDiscount1Days": "Optional. Days for discount 1.",
    "CashDiscount2Percent": "Optional. Discount 2 percentage.",
    "CashDiscount2Days": "Optional. Days for discount 2.",
    "NetPaymentDays": "Optional. Days until net payment.",
    "PaymentBlockingReason": "Optional payment block, e.g. A.",
    "AccountingDocumentType": "Optional, e.g. RE. Blank = system default.",
    "SupplierInvoiceStatus": "Leave blank to post. Use the status your system expects to park.",
    "PaymentReference": "Optional payment reference.",
    "InvoiceReference": "Credit memos / subsequent documents: the original invoice number.",
    "InvoiceReferenceFiscalYear": "Fiscal year of InvoiceReference.",
    "TaxIsCalculatedAutomatically": "TRUE = SAP calculates tax.",
    "BusinessPlace": "India: GST business place.",
    "SuplrInvcIsCapitalGoodsRelated": "TRUE / FALSE: capital goods invoice.",
    "SupplierInvoiceIsCreditMemo": "X for a credit memo, blank for an invoice.",
    "TaxDeterminationDate": "Optional tax date - used on the header (first row) and on each "
                            "line (its own row).",
    "RetentionDueDate": "Optional retention due date - header (first row) and each line.",
    "HouseBank": "Optional house bank for payment.",
    "HouseBankAccount": "Optional house bank account.",
    "IN_GSTPartner": "India: GST partner (usually the supplier).",
    "IN_GSTPlaceOfSupply": "India: place of supply (state code, e.g. 36).",
    "IN_InvoiceReferenceNumber": "India: e-invoice reference number (IRN).",
    "SupplierInvoiceItem": "Optional line number; numbered 1, 2, 3 per invoice when blank.",
    "PurchaseOrder": "Required. Purchase order.",
    "PurchaseOrderItem": "Required. PO item, e.g. 10.",
    "Plant": "Plant of the PO item.",
    "ReferenceDocument": "GR-based invoice verification: material document of the GR.",
    "ReferenceDocumentFiscalYear": "GR-based IV: GR document year.",
    "ReferenceDocumentItem": "GR-based IV: GR document item, e.g. 1.",
    "IsSubsequentDebitCredit": "TRUE for a subsequent debit / credit.",
    "TaxCode": "Tax code, e.g. G4.",
    "SupplierInvoiceItemAmount": "Required. Net line amount.",
    "PurchaseOrderQuantityUnit": "Unit of the invoiced quantity, e.g. PC.",
    "QuantityInPurchaseOrderUnit": "Invoiced quantity.",
    "PurchaseOrderPriceUnit": "Optional order price unit when it differs from the order unit.",
    "QtyInPurchaseOrderPriceUnit": "Quantity in the order price unit.",
    "SupplierInvoiceItemText": "Optional line text.",
    "FreightSupplier": "Optional freight supplier (planned delivery costs).",
    "IsNotCashDiscountLiable": "TRUE = line is not liable for cash discount.",
    "RetentionAmountInDocCurrency": "Optional retention amount.",
    "RetentionPercentage": "Optional retention percentage.",
    "ServiceEntrySheet": "Service invoices: service entry sheet.",
    "ServiceEntrySheetItem": "Service entry sheet item.",
    "IsFinallyInvoiced": "TRUE = final invoice for the PO item.",
    "IN_HSNOrSACCode": "India: HSN / SAC code.",
    "IN_CustomDutyAssessableValue": "India imports: assessable value for customs duty.",
    "ItemDocumentCurrency": "Optional line currency. Blank = header DocumentCurrency.",
}


def _auto_note(col: str) -> str:
    md = load_metadata(SERVICE)
    for et in (HEADER_TYPE, PO_TYPE):
        p = md.field(et, col) if md else None
        if p:
            kind = {"Edm.DateTime": " Date.", "Edm.Boolean": " TRUE / FALSE.",
                    "Edm.Decimal": " Number."}.get(p.get("Type", ""), "")
            return f"Optional. {p.get('quickinfo') or p.get('label') or col}.{kind}"
    return "Optional."


COLUMN_NOTES = {c: _NOTES.get(c) or _auto_note(c) for c in COLUMNS}

SAMPLE_ROWS = [
    {"InvoiceRefID": "INV1", "CompanyCode": "1810", "DocumentDate": "08.12.2026",
     "PostingDate": "08.12.2026", "SupplierInvoiceIDByInvcgParty": "SUP-INV-001",
     "InvoicingParty": "18300010", "DocumentCurrency": "INR", "InvoiceGrossAmount": "2950.00",
     "TaxIsCalculatedAutomatically": "TRUE", "PurchaseOrder": "4500000001",
     "PurchaseOrderItem": "10", "Plant": "1810", "ReferenceDocument": "5000000001",
     "ReferenceDocumentFiscalYear": "2026", "ReferenceDocumentItem": "1",
     "QuantityInPurchaseOrderUnit": "10", "PurchaseOrderQuantityUnit": "PC",
     "SupplierInvoiceItemAmount": "2500.00", "TaxCode": "G4"},
    {"InvoiceRefID": "INV2", "CompanyCode": "1810", "DocumentDate": "08.12.2026",
     "PostingDate": "08.12.2026", "SupplierInvoiceIDByInvcgParty": "SUP-INV-002",
     "InvoicingParty": "18300010", "DocumentCurrency": "INR", "InvoiceGrossAmount": "5900.00",
     "TaxIsCalculatedAutomatically": "TRUE", "PaymentTerms": "0001",
     "PurchaseOrder": "4500000002", "PurchaseOrderItem": "10", "Plant": "1810",
     "QuantityInPurchaseOrderUnit": "20", "PurchaseOrderQuantityUnit": "KG",
     "SupplierInvoiceItemAmount": "3000.00", "TaxCode": "G4", "IN_HSNOrSACCode": "39269099"},
    {"InvoiceRefID": "INV2", "PurchaseOrder": "4500000002", "PurchaseOrderItem": "20",
     "Plant": "1810", "QuantityInPurchaseOrderUnit": "10", "PurchaseOrderQuantityUnit": "KG",
     "SupplierInvoiceItemAmount": "2000.00", "TaxCode": "G4", "IN_HSNOrSACCode": "39269099",
     "IsFinallyInvoiced": "TRUE"},
]


def build_documents(df) -> list[dict]:
    md = load_metadata(SERVICE)
    hb, ib = EntityBuilder(md, HEADER_TYPE), EntityBuilder(md, PO_TYPE)
    invoices: dict[str, dict] = {}

    for row in rows_of(df):
        errs: list[str] = []
        ref = row.get(GROUP_COLUMN) or f"LINE{row['_line']}"
        if ref not in invoices:
            for f in HEADER_REQUIRED:
                if not row.get(f):
                    errs.append(f"{f} is required on the first row of invoice {ref}")
            header = hb.build(row, HEADER_FIELDS, errs)
            header[ITEM_NAV] = {"results": []}
            invoices[ref] = {"ref": ref, "lines": [], "errors": [], "payload": header,
                             "currency": row.get("DocumentCurrency", ""), "items": set()}
        inv = invoices[ref]

        for f in ITEM_REQUIRED:
            if not row.get(f):
                errs.append(f"{f} is required on every line")
        item_no = str(row.get("SupplierInvoiceItem")
                      or len(inv["payload"][ITEM_NAV]["results"]) + 1)
        if item_no.isdigit():
            item_no = item_no.zfill(6)
        if item_no in inv["items"]:
            errs.append(f"SupplierInvoiceItem {item_no} appears twice in invoice {ref}")
        inv["items"].add(item_no)
        line = ib.build(dict(row, SupplierInvoiceItem=item_no), ITEM_FIELDS, errs)
        currency = row.get(ITEM_CURRENCY_COLUMN) or inv["currency"]
        if currency:
            line["DocumentCurrency"] = currency
        inv["payload"][ITEM_NAV]["results"].append(line)

        inv["lines"].append(row["_line"])
        inv["errors"] += [f"line {row['_line']}: {e}" for e in errs]

    docs = []
    for inv in invoices.values():
        inv["errors"] += validate_payload(md, HEADER_TYPE, inv["payload"])
        docs.append({"ref": inv["ref"], "lines": ",".join(map(str, inv["lines"])),
                     "errors": inv["errors"], "payload": inv["payload"]})
    return docs


if __name__ == "__main__":
    run_cli(sys.modules[__name__])