import openpyxl
from openpyxl.drawing.image import Image as XLImage
import sys

def clear_images_from_excel(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    for sheet in wb.worksheets:
        # Remove all images from the sheet
        sheet._images = []
    wb.save(excel_path)
    print(f"Cleared all images from {excel_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python clear_excel_images.py <excel_file>")
        sys.exit(1)
    clear_images_from_excel(sys.argv[1])
