# Kohler Image Automation Tool - User Guide

## Overview
This tool automatically inserts product images from Kohler PDF pricebooks into Excel workbooks, matching images to SKU codes.

---

## Installation

### Option 1: Using Python (Technical Users)
1. Ensure Python 3.8+ is installed
2. Open Terminal
3. Navigate to the automation folder:
   ```bash
   cd /Users/pavankumarv/dev/automation
   ```
4. Activate the virtual environment:
   ```bash
   source venv/bin/activate
   ```
5. Run the application:
   ```bash
   python run_gui.py
   ```

### Option 2: Standalone App (Non-Technical Users)
1. Double-click `Kohler Image Tool.app`
2. If macOS shows a security warning:
   - Go to **System Preferences → Security & Privacy**
   - Click "Open Anyway"

---

## How to Use

### Step 1: Prepare Your Files
- **Excel Workbook**: Your product list with SKU codes in a column
- **PDF Pricebook**: Kohler pricebook with product images (e.g., "Kohler_PriceBook_Nov'25 Edition.pdf")

### Step 2: Launch the Application
- For Python users: Run `python run_gui.py`
- For standalone app: Double-click the app icon

### Step 3: Select Files
1. Click **Browse** next to "Excel Workbook"
   - Select your Excel file (`.xlsx`, `.xlsm`, or `.xlsb`)
2. Click **Browse** next to "PDF Pricebook"
   - Select the Kohler PDF file

### Step 4: Choose Output Option
- **Unchecked** (default): Overwrites original Excel file (faster)
- **Checked**: Creates new file with `_with_images` suffix (safer - preserves original)

> **Tip**: Check the box if customers might change/revert their orders later!

### Step 5: Fill Images
1. Click **Fill Images into Excel**
2. Wait for the process to complete (progress shown in log area)
3. A success message will appear when done

### Step 6: Generate Summary (Optional)
The summary sheet uses an Excel macro:
1. Open your Excel file (with or without images)
2. Press **Alt+F8** (Windows) or **Fn+Option+F8** (Mac)
3. Select **GenerateSummary** from the macro list
4. Click **Run**
5. The SUMMARY sheet will be created automatically

---

## Understanding the Output

### File Naming
- **Overwrite mode**: Original file is updated directly
- **Save As mode**: New file created with suffix
  - Example: `Latha_test.xlsx` → `Latha_test_with_images.xlsx`

### What Gets Copied?
When using "Save As New File", the **entire Excel file** is copied including:
- All sheets (data sheets, summary sheets)
- All formulas and calculations
- All formatting (fonts, colors, borders)
- All existing images
- All VBA macros
- All comments and metadata

Only the **IMAGE column** is modified with new product images.

---

## Troubleshooting

### "File Not Found" Error
- Make sure the Excel and PDF files exist
- Close the file picker and try selecting again

### "Operation Failed" Error
- **Close Excel** if the file is currently open
- Check that the file is not read-only
- Ensure you have write permissions to the folder

### Images Not Appearing
- Verify the Excel has a column named "CODE", "SKU", or "ITEM CODE"
- Verify the Excel has a column named "IMAGE" or "PICTURE"
- Check the log output for missing SKU warnings

### Application Won't Start (macOS)
- Right-click the app → **Open** (first time only)
- Or go to **System Preferences → Security & Privacy → General**
- Click **Open Anyway**

### "An Operation in Progress" Warning
- Wait for the current operation to finish
- Do not close the application while processing

---

## Tips for Best Results

### Excel File Requirements
- Must have a **CODE/SKU column** with product codes
- Must have an **IMAGE column** where images will be inserted
- SKU codes should match those in the PDF (normalized automatically)

### PDF File Requirements
- Must be a Kohler pricebook in PDF format
- Should contain product images and SKU codes
- Images should be near their corresponding SKU codes

### Performance
- Large files (100+ products) may take 2-5 minutes
- Progress is shown in the log area
- Do not close the application while "Processing..."

### File Management
- **For initial import**: Use overwrite mode (faster)
- **For order changes**: Use save-as mode (preserves original)
- **For testing**: Always use save-as mode first

---

## Summary Sheet Generation

The summary sheet is generated using an Excel VBA macro (not this GUI tool).

### Installing the Macro
1. Open Excel file
2. Press **Alt+F11** to open VBA Editor
3. Right-click **VBAProject** → **Import File**
4. Select `excel_macro/GenerateSummary.bas`
5. Close VBA Editor

### Running the Macro
1. Open Excel file
2. Press **Alt+F8** to open Macros dialog
3. Select **GenerateSummary**
4. Click **Run**

### What the Summary Includes
- List of all sheets with their MRP and Offer Price totals
- TOTAL MRP row (sum of all sheets)
- FINAL OFFER VALUE (with GST)
- Terms and Conditions section
- Professional formatting (Bookman Old Style font, pink headers)

---

## Support

### Common Questions

**Q: Can I process multiple Excel files at once?**
A: No, process one file at a time. Restart the app for each file.

**Q: Will this work on Windows?**
A: Yes, if using Python. For standalone app, you need a Windows build.

**Q: Can I customize image sizes?**
A: Images are standardized at 100×100 pixels for consistency.

**Q: What if SKU codes don't match?**
A: The tool normalizes SKUs (removes hyphens, converts to uppercase) automatically.

**Q: Can I undo the operation?**
A: If you used "overwrite mode", no. Always use "save-as mode" for safety!

---

## Version History

### Version 1.0 (January 2026)
- Initial release
- Automated image insertion from PDF to Excel
- Background processing with progress tracking
- Save-as and overwrite modes
- VBA macro for summary generation

---

## Contact
For technical support or feature requests, contact: [Your contact information]
