#!/usr/bin/env python3
"""
Bulk Product (material master) creation - API_PRODUCT_SRV, SAP DEV system.

One flat sheet. Rows with the same Product build one deep-insert payload:

  Product + basic data + description            -> A_Product, to_Description
  Plant (+ MRP / purchasing fields)             -> to_Plant
  StorageLocation                               -> to_Plant/to_StorageLocation
  LotSizingProcedure, PlannedDeliveryDays ...   -> to_Plant/to_ProductSupplyPlanning
  LoadingGroup                                  -> to_Plant/to_PlantSales
  ValuationArea + ValuationClass + prices       -> to_Valuation
  SalesOrg + DistributionChannel                -> to_SalesDelivery
  SalesTaxCountry + TaxCategory + TaxClass      -> to_ProductSalesTax
  AlternativeUnit + Numerator / Denominator     -> to_ProductUnitsOfMeasure

Repeat the Product on extra rows to add more plants, storage locations,
valuation areas, sales areas, tax rows or alternative units.

    python bulk_product_upload.py --template xlsx
    python bulk_product_upload.py -f products.xlsx --dry-run
    python bulk_product_upload.py -f products.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sap_common import (EntityBuilder, load_metadata, service_url, rows_of, run_cli,  # noqa: E402
                        validate_payload)

TITLE = "Products"
SLUG = "product"
SERVICE = "API_PRODUCT_SRV"
ENTITY_SET = "A_Product"
BASE_URL = service_url(SERVICE)          # DEV: my440263
DOC_FIELD = "Product"
YEAR_FIELD = None
TCODE = "MM01"
ICON = "🏷️"
COLOR = "#7C3AED"
GROUP_LABEL = "Product"
ITEMS_LABEL = "Plants"


def count_items(p: dict) -> int:
    return len(p.get("to_Plant", {}).get("results", []))
DESCRIPTION = ("Create products with description, plant, storage location, MRP, valuation, "
               "sales and tax data in one deep insert. Repeat the Product to add more views.")

T = {"product": "A_ProductType", "desc": "A_ProductDescriptionType", "plant": "A_ProductPlantType",
     "sloc": "A_ProductStorageLocationType", "mrp": "A_ProductSupplyPlanningType",
     "psales": "A_ProductPlantSalesType", "val": "A_ProductValuationType",
     "sales": "A_ProductSalesDeliveryType", "tax": "A_ProductSalesTaxType",
     "uom": "A_ProductUnitsOfMeasureType"}

PRODUCT_MAP = {f: f for f in ["Product", "ProductType", "IndustrySector", "BaseUnit", "ProductGroup",
                              "Division", "GrossWeight", "NetWeight", "WeightUnit",
                              "ItemCategoryGroup", "ProductOldID", "IsBatchManagementRequired",
                              "ProductHierarchy", "CountryOfOrigin", "SerialNumberProfile"]}
DESC_MAP = {"Language": "Language", "ProductDescription": "ProductDescription"}
PLANT_MAP = {"Plant": "Plant", "PurchasingGroup": "PurchasingGroup", "ProfitCenter": "ProfitCenter",
             "MRPType": "MRPType", "MRPResponsible": "MRPResponsible",
             "ProcurementType": "ProcurementType", "AvailabilityCheckType": "AvailabilityCheckType",
             "GoodsReceiptDuration": "GoodsReceiptDuration",
             "IsBatchManagementRequired": "PlantBatchManaged",
             "SerialNumberProfile": "PlantSerialNumberProfile"}
SLOC_MAP = {"StorageLocation": "StorageLocation", "WarehouseStorageBin": "StorageBin"}
MRP_MAP = {"LotSizingProcedure": "LotSizingProcedure", "MinimumLotSizeQuantity": "MinimumLotSize",
           "SafetyStockQuantity": "SafetyStock", "PlannedDeliveryDurationInDays": "PlannedDeliveryDays",
           "ReorderThresholdQuantity": "ReorderPoint", "MRPGroup": "MRPGroup"}
PSALES_MAP = {"LoadingGroup": "LoadingGroup"}
VAL_MAP = {"ValuationArea": "ValuationArea", "ValuationClass": "ValuationClass",
           "PriceDeterminationControl": "PriceDeterminationControl",
           "InventoryValuationProcedure": "PriceControl", "StandardPrice": "StandardPrice",
           "MovingAveragePrice": "MovingAveragePrice", "PriceUnitQty": "PriceUnit",
           "Currency": "Currency"}
SALES_MAP = {"ProductSalesOrg": "SalesOrg", "ProductDistributionChnl": "DistributionChannel",
             "SupplyingPlant": "DeliveringPlant", "ItemCategoryGroup": "SalesItemCategoryGroup",
             "AccountDetnProductGroup": "AccountAssignmentGroup",
             "MinimumOrderQuantity": "MinimumOrderQuantity", "SalesMeasureUnit": "SalesUnit"}
TAX_MAP = {"Country": "SalesTaxCountry", "TaxCategory": "TaxCategory",
           "TaxClassification": "TaxClassification"}
UOM_MAP = {"AlternativeUnit": "AlternativeUnit", "QuantityNumerator": "AltUnitNumerator",
           "QuantityDenominator": "AltUnitDenominator"}


def _cols(*maps):
    out = []
    for m in maps:
        for c in m.values():
            if c not in out:
                out.append(c)
    return out


COLUMNS = _cols(PRODUCT_MAP, DESC_MAP, PLANT_MAP, SLOC_MAP, MRP_MAP, PSALES_MAP, VAL_MAP,
                SALES_MAP, TAX_MAP, UOM_MAP)

COLUMN_NOTES = {
    "Product": "Required. Product / material number (max 40). Repeat on extra rows.",
    "ProductType": "Required, e.g. ROH, HALB, FERT, HAWA.",
    "IndustrySector": "Required, e.g. M.",
    "BaseUnit": "Required, e.g. PC, KG.",
    "ProductGroup": "Material group.",
    "Division": "e.g. 00.",
    "ItemCategoryGroup": "Basic-data item category group, e.g. NORM.",
    "IsBatchManagementRequired": "TRUE / FALSE (product level).",
    "Language": "Description language, default EN.",
    "ProductDescription": "Required. Description (max 40).",
    "Plant": "Plant view. Leave blank for basic data only.",
    "PurchasingGroup": "Plant: purchasing group.",
    "MRPType": "Plant: MRP type, e.g. PD, ND.",
    "MRPResponsible": "Plant: MRP controller, e.g. 001.",
    "ProcurementType": "Plant: E = in-house, F = external, X = both.",
    "AvailabilityCheckType": "Plant: availability check, e.g. 02.",
    "PlantBatchManaged": "Plant: batch management TRUE / FALSE.",
    "StorageLocation": "Storage location view.",
    "LotSizingProcedure": "Supply planning: lot size, e.g. EX.",
    "PlannedDeliveryDays": "Supply planning: planned delivery time in days.",
    "LoadingGroup": "Plant sales: loading group, e.g. 0001.",
    "ValuationArea": "Accounting view. Usually the plant.",
    "ValuationClass": "Accounting: valuation class, e.g. 3000.",
    "PriceControl": "Accounting: S = standard, V = moving average.",
    "StandardPrice": "Accounting: standard price (PriceControl S).",
    "MovingAveragePrice": "Accounting: moving average price (PriceControl V).",
    "PriceUnit": "Accounting: price unit, default 1.",
    "Currency": "Accounting: currency, e.g. INR.",
    "SalesOrg": "Sales view: sales organisation.",
    "DistributionChannel": "Sales view: distribution channel, e.g. 10.",
    "DeliveringPlant": "Sales view: delivering plant.",
    "SalesItemCategoryGroup": "Sales view: item category group, e.g. NORM.",
    "AccountAssignmentGroup": "Sales view: account assignment group, e.g. 01.",
    "SalesTaxCountry": "Sales tax: country, e.g. IN.",
    "TaxCategory": "Sales tax: tax category, e.g. JOIG.",
    "TaxClassification": "Sales tax: classification, e.g. 0 or 1.",
    "AlternativeUnit": "Alternative unit of measure, e.g. BOX.",
    "AltUnitNumerator": "How many base units ...",
    "AltUnitDenominator": "... per this many alternative units.",
}

SAMPLE_ROWS = [
    {"Product": "RM-DEV-001", "ProductType": "ROH", "IndustrySector": "M", "BaseUnit": "KG",
     "ProductGroup": "L001", "Division": "00", "Language": "EN",
     "ProductDescription": "Raw material DEV test", "Plant": "1810", "PurchasingGroup": "001",
     "MRPType": "PD", "MRPResponsible": "001", "ProcurementType": "F",
     "AvailabilityCheckType": "02", "StorageLocation": "181A", "LotSizingProcedure": "EX",
     "PlannedDeliveryDays": "5", "ValuationArea": "1810", "ValuationClass": "3000",
     "PriceControl": "V", "MovingAveragePrice": "120.00", "PriceUnit": "1", "Currency": "INR"},
    {"Product": "RM-DEV-001", "Plant": "1810", "StorageLocation": "181B"},
    {"Product": "FG-DEV-001", "ProductType": "FERT", "IndustrySector": "M", "BaseUnit": "PC",
     "ProductGroup": "L001", "Division": "00", "ItemCategoryGroup": "NORM", "Language": "EN",
     "ProductDescription": "Finished good DEV test", "Plant": "1810", "MRPType": "PD",
     "MRPResponsible": "001", "ProcurementType": "E", "AvailabilityCheckType": "02",
     "StorageLocation": "181A", "LotSizingProcedure": "EX", "LoadingGroup": "0001",
     "ValuationArea": "1810", "ValuationClass": "7920", "PriceControl": "S",
     "StandardPrice": "500.00", "PriceUnit": "1", "Currency": "INR", "SalesOrg": "1810",
     "DistributionChannel": "10", "DeliveringPlant": "1810", "SalesItemCategoryGroup": "NORM",
     "AccountAssignmentGroup": "01", "SalesTaxCountry": "IN", "TaxCategory": "JOIG",
     "TaxClassification": "1"},
]


def _upsert(lst: list, key: tuple, keyfn, item: dict):
    for x in lst:
        if keyfn(x) == key:
            for k, v in item.items():
                x.setdefault(k, v)
            return x
    lst.append(item)
    return item


def build_documents(df) -> list[dict]:
    md = load_metadata(SERVICE)
    B = {k: EntityBuilder(md, v) for k, v in T.items()}
    products: dict[str, dict] = {}

    for row in rows_of(df):
        errs: list[str] = []
        prod = row.get("Product", "")
        if not prod:
            errs.append("Product is empty")
            prod = f"LINE{row['_line']}"
        if prod not in products:
            for f in ("ProductType", "IndustrySector", "BaseUnit", "ProductDescription"):
                if not row.get(f):
                    errs.append(f"{f} is required on the first row of {prod}")
            payload = B["product"].build(row, PRODUCT_MAP, errs)
            payload["Product"] = row.get("Product", "")
            products[prod] = {"ref": prod, "lines": [], "errors": [], "payload": payload}
        doc = products[prod]
        p = doc["payload"]
        base = {"Product": row.get("Product", "")}

        if row.get("ProductDescription"):
            d = B["desc"].build(row, DESC_MAP, errs, defaults={"Language": "EN"})
            _upsert(p.setdefault("to_Description", {"results": []})["results"],
                    (d.get("Language"),), lambda x: (x.get("Language"),), {**base, **d})

        if row.get("Plant"):
            pl = B["plant"].build(row, PLANT_MAP, errs)
            plant = _upsert(p.setdefault("to_Plant", {"results": []})["results"],
                            (pl["Plant"],), lambda x: (x.get("Plant"),), {**base, **pl})
            pbase = {**base, "Plant": pl["Plant"]}
            if row.get("StorageLocation"):
                sl = B["sloc"].build(row, SLOC_MAP, errs)
                _upsert(plant.setdefault("to_StorageLocation", {"results": []})["results"],
                        (sl["StorageLocation"],), lambda x: (x.get("StorageLocation"),),
                        {**pbase, **sl})
            mrp = B["mrp"].build(row, MRP_MAP, errs)
            if mrp:
                plant.setdefault("to_ProductSupplyPlanning", dict(pbase)).update(mrp)
            ps = B["psales"].build(row, PSALES_MAP, errs)
            if ps:
                plant.setdefault("to_PlantSales", dict(pbase)).update(ps)
        elif any(row.get(c) for c in ("StorageLocation", "MRPType", "LotSizingProcedure")):
            errs.append("plant-level columns are filled but Plant is empty")

        if row.get("ValuationArea") or row.get("ValuationClass"):
            v = B["val"].build(row, VAL_MAP, errs, defaults={"ValuationArea": row.get("Plant", ""),
                                                             "PriceUnitQty": "1"})
            if not v.get("ValuationArea"):
                errs.append("ValuationArea is empty (and no Plant to default from)")
            if v.get("InventoryValuationProcedure") == "S" and "StandardPrice" not in v:
                errs.append("PriceControl S needs StandardPrice")
            if v.get("InventoryValuationProcedure") == "V" and "MovingAveragePrice" not in v:
                errs.append("PriceControl V needs MovingAveragePrice")
            _upsert(p.setdefault("to_Valuation", {"results": []})["results"],
                    (v.get("ValuationArea"),), lambda x: (x.get("ValuationArea"),), {**base, **v})

        if row.get("SalesOrg"):
            s = B["sales"].build(row, SALES_MAP, errs)
            if "ProductDistributionChnl" not in s:
                errs.append("DistributionChannel is required with SalesOrg")
            _upsert(p.setdefault("to_SalesDelivery", {"results": []})["results"],
                    (s.get("ProductSalesOrg"), s.get("ProductDistributionChnl")),
                    lambda x: (x.get("ProductSalesOrg"), x.get("ProductDistributionChnl")),
                    {**base, **s})

        if row.get("SalesTaxCountry"):
            t = B["tax"].build(row, TAX_MAP, errs)
            _upsert(p.setdefault("to_ProductSalesTax", {"results": []})["results"],
                    (t.get("Country"), t.get("TaxCategory")),
                    lambda x: (x.get("Country"), x.get("TaxCategory")), {**base, **t})

        if row.get("AlternativeUnit"):
            u = B["uom"].build(row, UOM_MAP, errs)
            _upsert(p.setdefault("to_ProductUnitsOfMeasure", {"results": []})["results"],
                    (u.get("AlternativeUnit"),), lambda x: (x.get("AlternativeUnit"),),
                    {**base, **u})

        doc["lines"].append(row["_line"])
        doc["errors"] += [f"line {row['_line']}: {e}" for e in errs]

    docs = []
    for d in products.values():
        d["errors"] += validate_payload(md, T["product"], d["payload"])
        d["lines"] = ",".join(map(str, d["lines"]))
        docs.append(d)
    return docs


if __name__ == "__main__":
    run_cli(sys.modules[__name__])
