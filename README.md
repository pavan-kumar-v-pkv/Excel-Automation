# Kohler Excel Automation

Python-powered automation for Kohler quotation Excel workbooks.

## Features

1. **Auto-fill Product Images** - Extracts images from Kohler pricebook PDF based on SKU codes
2. **Auto-generate Summary Sheet** - Scans all data sheets and creates summary with totals

## Project Structure

```
kohler_excel_automation/
├── python/                  # Python automation scripts
│   ├── main.py             # Main orchestrator
│   ├── pdf_parser.py       # PDF parsing logic
│   ├── image_extractor.py  # Image extraction
│   ├── excel_writer.py     # Excel modification
│   └── summary_builder.py  # Summary generation
├── excel/                   # Excel templates
├── tests/                   # Test files
├── data/                    # Sample data
└── temp_images/            # Temporary image storage
```

## Setup

### Prerequisites
- Python 3.8+
- Excel (for testing .xlsm templates)

### Installation

1. Clone the repository
2. Create virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Fill Images from PDF

```bash
python python/main.py fill-images --excel "path/to/workbook.xlsx" --pdf "path/to/pricebook.pdf"
```

### Generate Summary Sheet

```bash
python python/main.py create-summary --excel "path/to/workbook.xlsx"
```

## How It Works

### Image Automation
1. Parses Kohler PDF to locate SKU codes in CODE column
2. Ignores grey reference text and "Must order" dependencies
3. Extracts product images and resizes uniformly
4. Inserts images into Excel at correct SKU rows

### Summary Generation
1. Scans all sheets (except Summary)
2. Locates "TOTAL VALUE" row in each sheet
3. Extracts Total MRP and Total Offer Price
4. Populates Summary sheet

## Development Status

- [x] Project setup
- [ ] PDF parser
- [ ] Image extractor
- [ ] Excel writer
- [ ] Summary builder
- [ ] Main orchestrator
- [ ] Excel VBA integration

## License

Proprietary - Internal use only
