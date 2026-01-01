"""
Configuration Module
Centralized configuration for future customization
"""

from typing import List


class Config:
    """Configuration settings for the automation"""
    
    # Excel column names to look for
    CODE_COLUMN_NAMES: List[str] = ['CODE', 'SKU', 'ITEM CODE', 'PRODUCT CODE']
    DESC_COLUMN_NAMES: List[str] = ['DESCRIPTION', 'DESC', 'PRODUCT', 'ITEM']
    PRICE_COLUMN_NAMES: List[str] = ['MRP', 'PRICE', 'LIST PRICE', 'RATE']
    IMAGE_COLUMN_NAMES: List[str] = ['IMAGE', 'PICTURE', 'PHOTO', 'PRODUCT IMAGE']
    
    # Summary sheet identification
    SUMMARY_SHEET_NAMES: List[str] = ['SUMMARY', 'Summary', 'TOTAL', 'Total']
    TOTAL_VALUE_KEYWORDS: List[str] = ['TOTAL VALUE', 'GRAND TOTAL', 'NET TOTAL', 'TOTAL']
    
    # PDF parsing - keywords to skip
    SKIP_KEYWORDS: List[str] = [
        'MUST ORDER',
        'REQUIRED',
        'ACCESSORY',
        'INCLUDED',
        'SEE ALSO',
        'REFERENCE',
        'NOTE:',
        'AVAILABLE IN',
        'SOLD SEPARATELY'
    ]
    
    # Image settings
    IMAGE_TARGET_SIZE: tuple = (100, 100)  # Width, Height in pixels
    IMAGE_FORMAT: str = 'PNG'
    TEMP_IMAGE_DIR: str = 'temp_images'
    
    # Excel settings
    IMAGE_CELL_OFFSET_X: int = 2  # Pixels from left edge
    IMAGE_CELL_OFFSET_Y: int = 2  # Pixels from top edge
    
    def __init__(self, **kwargs):
        """
        Initialize config with optional overrides
        
        Args:
            **kwargs: Key-value pairs to override default settings
        """
        for key, value in kwargs.items():
            if hasattr(self, key):
                setattr(self, key, value)