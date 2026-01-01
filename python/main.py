"""
Main Orchestrator Module
Command-line interface for Excel automation
"""

import argparse
import sys
from pathlib import Path
from tkinter import Tk, filedialog

from .config import Config
from .pdf_parser import KohlerPDFParser
from .image_extractor import ImageExtractor
from .excel_handler import ExcelHandler
from .summary_builder import SummaryBuilder

class ExcelAutomation:
    """Main automation orchestrator"""

    def __init__(self, config: Config = None):
        """
        Initialize automation
        
        Args:
            config: Optional configuration object
        """
        self.config = config or Config()

    def fill_images_from_pdf(self, excel_path: str, pdf_path: str = None) -> bool:
        """
        Fill images from PDF to Excel file

        Args:
            excel_path: Path to Excel file
            pdf_path: Path to Pdf file (if None, will prompt user)
        Returns:
            True if successful
        """
        print("=" * 60)
        print(" FILL IMAGES FROM PDF")
        print("=" * 60)

        # Validate Excel file
        if not Path(excel_path).exists():
            print(f" Error: Excel file '{excel_path}' does not exist.")
            return False
        
        # Get PDF path if not provided
        if not pdf_path:
            pdf_path = self._select_pdf_file()
            if not pdf_path:
                print(" No PDF file selected. Exiting.")
                return False
            
        # Validate PDF file
        if not Path(pdf_path).exists():
            print(f" Error: PDF file '{pdf_path}' does not exist.")
            return False
        
        try:
            # Step 1: Parse PDF to get SKU-image mappings
            print("\n STEP 1: Parsing PDF for SKU-image mappings...")
            parser = KohlerPDFParser(pdf_path, self.config)
            sku_map = parser.parse()

            if not sku_map:
                print(" No SKUs found in PDF")
                return False
            # Export normalized PDF SKUs to file
            self._export_pdf_skus(sku_map)
            
            # Step 2: Extract images from PDF
            print("\n STEP 2: Extracting images from PDF...")
            extractor = ImageExtractor(pdf_path, self.config)
            image_paths = extractor.extract_images(sku_map)

            if not image_paths:
                print(" No images extracted from PDF")
                return False
            
            # Step 3: Open Excel and scan for SKUs
            print("\n STEP 3: Scanning Excel for SKUs...")
            excel = ExcelHandler(excel_path, self.config)
            excel.open()

            # Scan Excel for SKU codes
            excel.scan_skus()
            # Export normalized Excel SKUs to file
            self._export_excel_skus(excel.sheet_sku_map)

            # Step 4: Insert images into Excel
            print("\n STEP 4: Inserting images into Excel...")
            inserted_count = excel.insert_images(image_paths)
            # Show comparison of missing SKUs
            self._print_missing_sku_report(excel.sheet_sku_map, image_paths)
            # Step 5: Save Excel file
            excel.save()
            excel.close()

            # Step 6: Cleaup temporary images
            print("\n STEP 5: Cleaning up temporary images...")
            extractor.cleanup()

            print("\n" + "=" * 60)
            print(f" Completed: Inserted {inserted_count} images into '{excel_path}'")
            print("=" * 60)
            return True
        except Exception as e:
            print(f" Error during processing: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _export_pdf_skus(self, sku_map):
        """Export normalized PDF SKUs and image info to a text file"""
        out_path = "data/pdf_skus_export.txt"
        try:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write("SKU\tNormalized\tHasImage\tImagePage\tImageBBox\n")
                for sku, data in sku_map.items():
                    norm = self._normalize_sku(sku)
                    has_img = str(data.get("has_image", False))
                    page = str(data.get("image_page", ""))
                    bbox = str(data.get("image_bbox", ""))
                    f.write(f"{sku}\t{norm}\t{has_img}\t{page}\t{bbox}\n")
            print(f" Exported PDF SKUs to {out_path}")
        except Exception as e:
            print(f" Error exporting PDF SKUs: {e}")

    def _export_excel_skus(self, sheet_sku_map):
        """Export normalized Excel SKUs (with sheet info) to a text file"""
        out_path = "data/excel_skus_export.txt"
        try:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write("Sheet\tRow\tSKU\tNormalized\n")
                for sheet, skus in sheet_sku_map.items():
                    for row, sku_raw, sku_norm in skus:
                        f.write(f"{sheet}\t{row}\t{sku_raw}\t{sku_norm}\n")
            print(f" Exported Excel SKUs to {out_path}")
        except Exception as e:
            print(f" Error exporting Excel SKUs: {e}")

    def _normalize_sku(self, sku):
        from .pdf_parser import normalize_sku
        return normalize_sku(sku)
        
    def create_summary(self, excel_path: str) -> bool:
        """
        Create or update summary sheet in Excel file

        Args:
            excel_path: Path to Excel file
        Returns:
            True if successful
        """
        print("=" * 60)
        print(" CREATE/UPDATE SUMMARY SHEET")
        print("=" * 60)

        # Validate Excel file
        if not Path(excel_path).exists():
            print(f" Error: Excel file '{excel_path}' does not exist.")
            return False
        
        try:
            # Step 1: Open Excel file
            print("\n STEP 1: Opening Excel file...")
            excel = ExcelHandler(excel_path, self.config)
            excel.open()

            # Step 2: Build summary sheet
            print("\n STEP 2: Building summary sheet...")
            summary_builder = SummaryBuilder(excel.workbook, self.config)
            success = summary_builder.build_summary()

            if not success:
                print(" No summary data created.")
                excel.close()
                return False

            # Step 3: Save and close Excel file
            print("\n STEP 3: Saving and closing Excel file...")
            excel.save()
            excel.close()

            print("\n" + "=" * 60)
            print(f" Completed: Summary sheet updated in '{excel_path}'")
            print("=" * 60)

            return True
        
        except Exception as e:
            print(f" Error during processing: {e}")
            import traceback
            traceback.print_exc()
            return False
        
    def _print_missing_sku_report(self, sheet_sku_map, image_paths):
        """Print report of SKUs missing images"""
        print("\n" + "=" * 60)
        print("📊 MISSING IMAGE REPORT")
        print("=" * 60)
        
        all_excel_skus = set()
        missing_skus = []
        
        for sheet_name, skus in sheet_sku_map.items():
            for _, sku_raw, sku_norm in skus:
                all_excel_skus.add((sku_norm, sku_raw))
                if sku_norm not in image_paths and sku_norm not in ['SKUCODE', 'TOTALVALUE']:
                    missing_skus.append((sku_norm, sku_raw, sheet_name))
        
        pdf_skus = set(image_paths.keys())
        
        print(f"\n📈 Statistics:")
        print(f"   Total unique SKUs in Excel: {len(all_excel_skus)}")
        print(f"   SKUs with images from PDF: {len(pdf_skus)}")
        print(f"   SKUs missing images: {len(missing_skus)}")
        
        if missing_skus:
            print(f"\n❌ Missing SKUs (first 20):")
            for sku_norm, sku_raw, sheet in missing_skus[:20]:
                print(f"   {sku_raw:30} (normalized: {sku_norm:25}) in [{sheet}]")
            
            if len(missing_skus) > 20:
                print(f"   ... and {len(missing_skus) - 20} more")
        
        print("=" * 60)
    
    def _select_pdf_file(self) -> str:
        """
        Show file picker to select PDF file

        Returns:
            Path to selected PDF file or empty string
        """
        print("\n Please select a PDF file...")

        try:
            # Create hidden Tkinter window
            root = Tk()
            root.withdraw()  # Hide the root window
            root.attributes('-topmost', True)  # Bring the dialog to the front

            # Show file picker dialog
            file_path = filedialog.askopenfilename(
                title="Select Kohler Pricebook PDF",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            )
            root.destroy()  # Destroy the Tkinter root window

            return file_path if file_path else ""
        
        except Exception as e:
            print(f" Could not show file picker: {str(e)}")
            print(" Please provide PDF path via command-line argument.")
            return ""
        
def main():
    """Main entry point for CLI"""

    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description="Kohler Excel Automation - Auto fill images and generate summaries",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
        epilog="""
Examples:
  # Fill images from PDF (will prompt for PDF file)
  python -m python.main fill-images --excel "Latha.xlsx"
  
  # Fill images with specific PDF
  python -m python.main fill-images --excel "Latha.xlsx" --pdf "Kohler_PriceBook.pdf"
  
  # Create summary sheet
  python -m python.main create-summary --excel "Latha.xlsx"
        """
    )

    parser.add_argument(
        'mode',
        choices=['fill-images', 'create-summary'],
        help='Operation mode'
    )
    parser.add_argument(
        '--excel',
        required=True,
        help='Path to Excel file'
    )
    
    parser.add_argument(
        '--pdf',
        help='Path to PDF file (for fill-images mode, optional - will prompt if not provided)'
    )
    args = parser.parse_args()

     # Create automation instance
    automation = ExcelAutomation()
    
    # Execute based on mode
    if args.mode == 'fill-images':
        success = automation.fill_images_from_pdf(args.excel, args.pdf)
    elif args.mode == 'create-summary':
        success = automation.create_summary(args.excel)
    else:
        print(f"Unknown mode: {args.mode}")
        sys.exit(1)
    
    # Exit with appropriate code
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()