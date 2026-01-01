"""
compare_sku_exports.py
Compare exported PDF SKUs and Excel SKUs to find mismatches.
"""

import csv
from pathlib import Path

def load_pdf_skus(path):
    skus = set()
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter='\t')
        for row in reader:
            if row.get("HasImage", "False").lower() == "true":
                # Use the Normalized column for matching
                norm = row.get("Normalized", "").strip()
                if norm:
                    skus.add(norm)
    return skus

def load_excel_skus(path):
    skus = set()
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter='\t')
        for row in reader:
            skus.add(row["Normalized"].strip())
    return skus

def main():
    pdf_export = Path("data/pdf_skus_export.txt")
    excel_export = Path("data/excel_skus_export.txt")
    if not pdf_export.exists() or not excel_export.exists():
        print("Export files not found. Please run the main automation first.")
        return

    pdf_skus = load_pdf_skus(pdf_export)
    excel_skus = load_excel_skus(excel_export)

    missing_in_pdf = sorted(excel_skus - pdf_skus)
    extra_in_pdf = sorted(pdf_skus - excel_skus)

    print("\n=== SKU Comparison Report ===")
    print(f"Total SKUs in Excel: {len(excel_skus)}")
    print(f"Total SKUs in PDF (with images): {len(pdf_skus)}")
    print(f"SKUs in Excel but missing in PDF: {len(missing_in_pdf)}")
    print(f"SKUs in PDF but not in Excel: {len(extra_in_pdf)}")

    if missing_in_pdf:
        print("\nSKUs in Excel but missing in PDF (first 20):")
        for sku in missing_in_pdf[:20]:
            print(f"  {sku}")
        if len(missing_in_pdf) > 20:
            print(f"  ... and {len(missing_in_pdf) - 20} more")

    if extra_in_pdf:
        print("\nSKUs in PDF but not in Excel (first 20):")
        for sku in extra_in_pdf[:20]:
            print(f"  {sku}")
        if len(extra_in_pdf) > 20:
            print(f"  ... and {len(extra_in_pdf) - 20} more")

if __name__ == "__main__":
    main()
