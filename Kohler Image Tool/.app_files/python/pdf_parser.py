"""
PDF Parser Module
Extracts SKU-to-image mappings from Kohler pricebook PDF
"""

import pdfplumber
import re
from typing import Dict, List, Tuple, Optional
from .config import Config

def normalize_sku(sku: str) -> str:
    """Normalize SKU code by stripping whitespace and converting to uppercase"""
    if not sku:
        return ""
    sku = str(sku).strip().replace(' ', '').replace('-', '').replace('_', '').replace('/', '').upper()
    if sku.startswith('K'):
        # Remove leading 'K' if followed by digits
        if len(sku) > 1 and sku[1].isdigit():
            sku = sku[1:]
    return sku

class KohlerPDFParser:
    """Parse Kohler pricebook PDF to extract SKU and image coordinates."""

    def __init__(self, pdf_path: str, config=None):
        """
        Initialize parser with PDF path

        Args:
            pdf_path: Path to Kohler pricebook PDF
            config: Configuration object for customization
        """
        self.pdf_path = pdf_path
        self.config = config or Config()
        self.sku_image_map = {} # SKU -> (page_num, image_bbox)
        self.sku_data_map = {}

    def parse(self) -> Dict[str, dict]:
        """
        Parse PDF and extract SKU mappings
        Returns:
            Dict mapping SKU codes to their data and image coordinates
        """
        print(f"Parsing PDF: {self.pdf_path}")
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                total_pages = len(pdf.pages)
                for page_num, page in enumerate(pdf.pages, start=1):
                    print(f"    Processing page {page_num}/{total_pages}...")
                    self._parse_page(page, page_num)
        except Exception as e:
            print(f"Error parsing PDF: {e}")
            raise

        print("SKUs with images found in PDF (normalized):")
        for sku in self.sku_image_map.keys():
            print(sku)
        print(f" Found {len(self.sku_image_map)} SKUs with images.")
        return self._build_result_map()
    
    def _parse_page(self, page, page_num: int):
        """
        Parse a single PDF page

        Args:
            page: pdfplumber page object
            page_num: Current page number
        """
        # Extract tables from the page
        tables = page.extract_tables()

        if not tables:
            # Try extracting text if not tables
            self._parse_text_mode(page, page_num)
            return
        
        # Extract images from the page
        images = page.images

        # Process each table
        for table in tables:
            self._process_table(table, images, page_num)

    def _parse_text_mode(self, page, page_num: int):
        """
        Fallback: Parse page as text when no tables found

        Args:
            page: pdfplumber page object
            page_num: Current page number
        """
        text = page.extract_text()
        if not text:
            return
        
        # Look for SKU patterns in text
        sku_pattern = r'\b([A-Z0-9]{4,}[-/]?[A-Z0-9]{2,})\b'
        matches = re.findall(sku_pattern, text)

        images = page.images

        for sku in matches:
            if not self._is_valid_sku(sku):
                continue

            self.sku_data_map[sku] = {
                'sku': sku,
                'description': '',
                'mrp': '',
                'page': page_num
            }

            # Associate first available image
            if images and sku not in self.sku_image_map:
                self.sku_image_map[sku] = (page_num, images[0])

    def _process_table(self, table: List[List], images: List[dict], page_num: int):
        """
        Process a table to find SKU codes and match with images

        Args:
            table: Extracted table data
            images: List of image objects on the page
            page_num: Current page number
        """
        if not table or len(table) < 2:
            return
        
        # Find column indices
        header_row = table[0]
        code_col_idx = self._find_column_index(header_row, self.config.CODE_COLUMN_NAMES)
        desc_col_idx = self._find_column_index(header_row, self.config.DESC_COLUMN_NAMES)
        mrp_col_idx = self._find_column_index(header_row, self.config.PRICE_COLUMN_NAMES)

        if code_col_idx is None:
            # No CODE column found, skip this table
            return
        
        # Process data rows
        for row_idx, row in enumerate(table[1:], start=1):
            if not row or len(row) <= code_col_idx:
                continue

            raw_sku_cell = row[code_col_idx]
            if not raw_sku_cell:
                continue
            
            # Extract all SKUs from the cell (handles concatenated SKUs)
            sku_codes = self._extract_all_skus(str(raw_sku_cell))
            
            for sku_code_raw in sku_codes:
                sku_code = normalize_sku(sku_code_raw)

                if not sku_code:
                    continue

                # Store SKU data
                sku_data = {
                    'sku': sku_code_raw,
                    'sku_normalized': sku_code,
                    'description': row[desc_col_idx] if desc_col_idx and len(row) > desc_col_idx else '',
                    'mrp': row[mrp_col_idx] if mrp_col_idx and len(row) > mrp_col_idx else '',
                    'page': page_num,
                    'row_index': row_idx
                }

                # Try to find associated image
                image_bbox = self._find_nearest_image(row_idx, images, len(table))
                
                if image_bbox:
                    self.sku_image_map[sku_code] = (page_num, image_bbox)
                    self.sku_data_map[sku_code] = sku_data
                else:
                    # Store SKU even without image
                    self.sku_data_map[sku_code] = sku_data

    def _find_column_index(self, header_row: List, possible_names: List[str]) -> Optional[int]:
        """
        Find column index by matching header names

        Args:
            header_row: Header row from table
            possible_names: List of possible column names

        Returns:
            Column index or None    
        """
        if not header_row:
            return None
        
        for idx, cell in enumerate(header_row):
            if cell:
                cell_upper = str(cell).upper().strip()
                for name in possible_names:
                    if name.upper() in cell_upper:
                        return idx
        return None
    
    def _clean_sku(self, sku_value) -> Optional[str]:
        """
        Clean and validate SKU code

        Args:
            sku_value: Raw SKU value from table

        Returns:
            Cleaned SKU or None (returns first valid SKU if multiple are concatenated)
        """
        if not sku_value:
            return None
        
        sku = str(sku_value).strip()

        # Remove common prefixes/suffixes
        sku = sku.replace('\n', ' ').strip()
        
        # Try to split concatenated SKUs (e.g., "75890INPCPK75890INMCP")
        # Look for pattern where a "K" followed by digits indicates a new Kohler SKU
        split_pattern = r'(K-?\d+[A-Z0-9-]+)'
        matches = re.findall(split_pattern, sku, re.IGNORECASE)
        if matches and len(matches) >= 1:
            # Return the first valid SKU found
            for match in matches:
                if self._is_valid_sku(match):
                    return match

        # Basic validation
        if not self._is_valid_sku(sku):
            return None
        
        return sku
    
    def _extract_all_skus(self, sku_value: str) -> List[str]:
        """
        Extract all SKUs from a cell value (handles concatenated SKUs)

        Args:
            sku_value: Raw SKU value from table

        Returns:
            List of valid SKU codes found
        """
        if not sku_value:
            return []
        
        sku = str(sku_value).strip().replace('\n', ' ').strip()
        skus = []
        
        # Pattern to match Kohler SKU format: K-XXXXX or just XXXXX where X is alphanumeric
        # Kohler SKUs typically look like: K-75890IN-M-CP or 75890IN-M-CP
        # Also handle patterns like: 75890INPCP, 20746INCP, etc.
        
        # First try to split by common Kohler SKU pattern (K followed by digits)
        # This handles "75890INPCPK75890INMCP" -> ["K75890INMCP"] 
        k_pattern = r'K-?\d+[A-Z0-9-]*(?:IN|T)[A-Z0-9-]*'
        k_matches = re.findall(k_pattern, sku, re.IGNORECASE)
        
        if k_matches:
            for match in k_matches:
                if self._is_valid_sku(match):
                    skus.append(match)
        
        # Also try to find SKUs that don't start with K but have the IN pattern
        # like "75890INPCP", "20746INCP"
        in_pattern = r'\d+IN[A-Z0-9-]*'
        in_matches = re.findall(in_pattern, sku, re.IGNORECASE)
        
        if in_matches:
            for match in in_matches:
                if self._is_valid_sku(match) and match not in skus:
                    skus.append(match)
        
        # If no matches found, try the original value
        if not skus and self._is_valid_sku(sku):
            skus.append(sku)
        
        return skus
    
    def _is_valid_sku(self, sku: str) -> bool:
        """
        Validate SKU format

        Args:
            sku: SKU code to validate

        Returns:
            True if valid SKU format
        """
        if len(sku) < 3:
            return False
        
        # Should have alphanumeric characters
        if not re.search(r'[A-Za-z0-9]', sku):
            return False

        # Should not be just numbers
        if sku.isdigit():
            return False
        
        return True
    
    def _is_grey_reference_row(self, row: List) -> bool:
        """
        Check if row is a grey reference/dependency row

        Args:
            row: Table row

        Returns:
            True if this is a reference row to skip
        """
        row_text = ' '.join([str(cell) for cell in row if cell]).upper()

        # use keywords from config
        for keyword in self.config.SKIP_KEYWORDS:
            if keyword.upper() in row_text:
                return True
            
        return False
    
    def _find_nearest_image(self, row_idx: int, images: List[dict], total_rows: int, table_bbox: dict = None) -> Optional[dict]:
        """
        Find image nearest to the current row based on Y-coordinates

        Args:
            row_idx: Current row index
            images: List of images on the page
            total_rows: Total rows in table
            table_bbox: Optional bounding box of the table for spatial reference

        Returns:
            Image bounding box or None
        """
        if not images:
            return None
        
        # Sort images by their Y position (top coordinate)
        sorted_images = sorted(images, key=lambda img: img.get('top', 0))
        
        # If we have as many or more images as data rows (excluding header),
        # try to map by position in the sorted list
        data_rows = total_rows - 1  # Exclude header row
        
        if len(sorted_images) >= data_rows and row_idx <= len(sorted_images):
            # row_idx is 1-based for data rows, so row_idx-1 gives us 0-based index
            return sorted_images[row_idx - 1]
        
        # If fewer images than rows, try to find the closest image by estimating
        # the Y-position of this row
        if len(sorted_images) > 0 and data_rows > 0:
            # Calculate approximate Y position for this row
            # Assume images span the page somewhat evenly
            first_img_top = sorted_images[0].get('top', 0)
            last_img_top = sorted_images[-1].get('top', first_img_top)
            
            if len(sorted_images) > 1:
                # Estimate where this row's image should be
                row_fraction = (row_idx - 1) / max(data_rows - 1, 1)
                estimated_y = first_img_top + row_fraction * (last_img_top - first_img_top)
                
                # Find closest image to estimated Y position
                closest_img = min(sorted_images, key=lambda img: abs(img.get('top', 0) - estimated_y))
                return closest_img
            else:
                # Only one image, return it
                return sorted_images[0]
        
        return None
    
    def _build_result_map(self) -> Dict[str, dict]:
        """
        Build final result mapping
        Returns: Dict with SKU mappings including image info
        """
        result = {}

        for sku, data in self.sku_data_map.items():
            if sku in self.sku_image_map:
                page_num, image_bbox = self.sku_image_map[sku]
                result[sku] = {
                    **data,
                    'has_image': True,
                    'image_page': page_num,
                    'image_bbox': image_bbox
                }
            else:
                result[sku] = {
                    **data,
                    'has_image': False
                }

        return result
    
def parse_pdf(pdf_path: str, config: Config = None) -> Dict[str, dict]:
    """
    Convenience function to parse PDF
    
    Args:
        pdf_path: Path to Kohler PDF
        config: Optional configuration object
        
    Returns:
        SKU mapping dictionary
    """
    parser = KohlerPDFParser(pdf_path, config)
    return parser.parse()




