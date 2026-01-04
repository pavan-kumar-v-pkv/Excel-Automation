"""
Image Extractor Module
Extracts images from PDF pages for further processing
"""

import os
from pathlib import Path
from typing import Dict, Optional
from PIL import Image
import pdfplumber
import io

from .pdf_parser import normalize_sku

class ImageExtractor:
    """Extract and resize images from PDF"""

    def __init__(self, pdf_path: str, config=None):
        """
        Initialize image extractor

        Args:
            pdf_path: Path to PDF file
            config: Configuration object
        """
        from .config import Config

        self.pdf_path = pdf_path
        self.config = config or Config()
        self.output_dir = Path(self.config.TEMP_IMAGE_DIR)
        self.output_dir.mkdir(exist_ok=True)

        # Standard image size from config
        self.target_size = self.config.IMAGE_TARGET_SIZE

    def extract_images(self, sku_map: Dict[str, dict]) -> Dict[str, str]:
        """
        Extract images for all SKUs

        Args:
            sku_map: Dictionary mapping SKU to image info from parser

        Returns:
            Dictionary mapping SKU to save image path
        """
        print(f"\n Extracting images to {self.output_dir}/")

        image_paths = {}
        skus_with_images = [sku for sku, data in sku_map.items() if data.get("has_image")]

        if not skus_with_images:
            print(" No SKUs with images found.")
            return image_paths
        
        print(f" Found {len(skus_with_images)} SKUs with images.")

        with pdfplumber.open(self.pdf_path) as pdf:
            for sku, data in sku_map.items():
                if not data.get("has_image"):
                    continue

                page_num = data.get('image_page', 1)
                image_bbox = data.get('image_bbox')

                if not image_bbox:
                    continue

                # Extract image from PDF page
                try:
                    page = pdf.pages[page_num-1]
                    image_path = self._extract_single_image(page, sku, image_bbox, page_num)

                    if image_path:
                        # Use normalized SKU as key for matching with Excel
                        sku_norm = normalize_sku(sku)
                        image_paths[sku_norm] = image_path
                        print(f"   ✓ {sku} -> {sku_norm}")
                except Exception as e:
                    print(f"   ✗ {sku}: {str(e)}")

        print(f"Successfully extracted {len(image_paths)}/{len(skus_with_images)} images")
        return image_paths
    
    def _extract_single_image(self, page, sku: str, image_info: dict, page_num: int) -> Optional[str]:
        """
        Extract a single image from PDF page

        Args:
            page: pdfplumber page object
            sku: SKU code
            image_info: Image bounding box info
            page_num: Current page number

        Returns:
            Path to saved image file
        """
        try:
            # Method 1: Try to extract image directly from pdfplumber image object
            if 'stream' in image_info:
                result = self._extract_from_stream(image_info, sku)
                if result:
                    return result
                # Fall through to crop method if stream extraction fails

            # Method 2: Crop page region and convert to image
            bbox = self._get_bbox_coordinates(image_info)
            if bbox:
                result = self._extract_from_crop(page, sku, bbox)
                if result:
                    return result
            
            # Method 3: try to get first image on page
            return self._extract_first_image(page, sku, page_num)
        except Exception as e:
            print(f"    Error extracting {sku}: {str(e)}")
            return None
        
    def _extract_from_stream(self, image_info: dict, sku: str) -> Optional[str]:
        """Extract image from PDF image stream"""
        try:
            stream = image_info.get('stream')
            if not stream:
                return None
            
            # Try to get image data
            image_data = stream.get_data()
            pil_image = Image.open(io.BytesIO(image_data))
            
            # Load image fully into memory to avoid "broken data stream" errors
            # This is necessary for JPX/JPEG2000 images that use lazy loading
            pil_image.load()

            # Resize and save
            resized = self._resize_image(pil_image)
            output_path = self._save_image(resized, sku)

            return output_path
        except Exception as e:
            # Silently fail - let other methods try
            return None
        
    def _extract_from_crop(self, page, sku: str, bbox: tuple) -> Optional[str]:
        """Extract image by cropping page region"""
        try:
            # Clamp bbox coordinates to valid page boundaries
            page_width = float(page.width)
            page_height = float(page.height)
            
            x0 = max(0.0, min(float(bbox[0]), page_width))
            y0 = max(0.0, min(float(bbox[1]), page_height))
            x1 = max(0.0, min(float(bbox[2]), page_width))
            y1 = max(0.0, min(float(bbox[3]), page_height))
            
            # Ensure valid dimensions
            if x1 <= x0 or y1 <= y0:
                return None
            
            clamped_bbox = (x0, y0, x1, y1)
            
            # Crop the page to bbox
            cropped = page.crop(clamped_bbox)

            # Convert to image
            img = cropped.to_image(resolution=150)
            pil_image = img.original

            # resize and save
            resized = self._resize_image(pil_image)
            output_path = self._save_image(resized, sku)

            return output_path
        except:
            return None
        
    def _extract_first_image(self, page, sku: str, page_num: int) -> Optional[str]:
        """Fallback: Extract first available image from page"""
        try:
            images = page.images
            if not images:
                return None
            
            # Try first image
            return self._extract_from_stream(images[0], sku)
        except:
            return None
        
    def _get_bbox_coordinates(self, image_info: dict) -> Optional[tuple]:
        """
        Extract bounding box coordinates from image info

        Args:
            image_info: Image information dictionary
        Returns: 
            Tuple of (x0, y0, x1, y1) or None
        """
        # try different coordinate formats:
        if all(k in image_info for k in ['x0', 'top', 'x1', 'bottom']):
            return (
                image_info['x0'],
                image_info['top'],
                image_info['x1'],
                image_info['bottom']
            )
        elif all(k in image_info for k in ['x0', 'y0', 'x1', 'y1']):
            return(
                image_info['x0'],
                image_info['y0'],
                image_info['x1'],
                image_info['y1']
            )
        elif 'bbox' in image_info:
            bbox = image_info['bbox']
            if len(bbox) == 4:
                return tuple(bbox)
        return None
    
    def _resize_image(self, image: Image.Image) -> Image.Image:
        """
        Resize image to standard size while maintaining aspect ratio and transparency
        
        Args:
            image: PIL Image object
            
        Returns:
            Resized image
        """
        # Keep transparency if present
        has_transparency = image.mode in ('RGBA', 'LA', 'P')
        
        if image.mode == 'P':
            image = image.convert('RGBA')
        elif image.mode not in ('RGBA', 'LA', 'RGB'):
            image = image.convert('RGBA' if has_transparency else 'RGB')
        
        # Create thumbnail (maintains aspect ratio)
        image.thumbnail(self.target_size, Image.Resampling.LANCZOS)
        
        # Create transparent or white background based on original
        if has_transparency or image.mode == 'RGBA':
            # Keep transparency - use RGBA with transparent background
            background = Image.new('RGBA', self.target_size, (255, 255, 255, 0))  # Transparent
        else:
            # Use white background
            background = Image.new('RGB', self.target_size, (255, 255, 255))
        
        # Center the image on background
        offset = (
            (self.target_size[0] - image.width) // 2,
            (self.target_size[1] - image.height) // 2
        )
        
        if image.mode == 'RGBA':
            background.paste(image, offset, image)  # Use alpha channel as mask
        else:
            background.paste(image, offset)
        
        return background
    
    def _save_image(self, image: Image.Image, sku: str) -> str:
        """
        Save image to file

        Args:
            image: PIL Image object
            sku: SKU code for filename

        Returns:
            Path to saved image
        """
        safe_filename = self._safe_filename(sku)
        output_path = self.output_dir / f"{safe_filename}.{self.config.IMAGE_FORMAT.lower()}"
        image.save(output_path, self.config.IMAGE_FORMAT)

        return str(output_path)
    
    def _safe_filename(self, sku: str) -> str:
        """
        Create safe filename from SKU

        Args:
            sku: SKU code
        Returns:
            Safe filename
        """
        # Replace unsafe characters
        safe = sku.replace('/', '_').replace('\\', '_').replace(' ', '_')
        safe = ''.join(c for c in safe if c.isalnum() or c in ('_', '-'))
        return safe
    
    def cleanup(self):
        """Remove all temporary images"""
        if self.output_dir.exists():
            count = 0
            for file in self.output_dir.glob('*.*'):
                file.unlink()
                count += 1
            if count > 0:
                print(f"🗑️  Cleaned up {count} temporary images")


def extract_images_from_pdf(pdf_path: str, sku_map: Dict[str, dict], config=None) -> Dict[str, str]:
    """
    Convenience function to extract images
    
    Args:
        pdf_path: Path to PDF
        sku_map: SKU mapping from parser
        config: Optional configuration
        
    Returns:
        Dictionary mapping SKU to image path
    """
    extractor = ImageExtractor(pdf_path, config)
    return extractor.extract_images(sku_map)