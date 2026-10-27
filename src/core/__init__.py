"""The four DEV upload types, in the order shown in the app."""
from . import material_document, supplier_invoice, purchase_order, product

UPLOADS = [material_document, supplier_invoice, purchase_order, product]
BY_SLUG = {m.SLUG: m for m in UPLOADS}
