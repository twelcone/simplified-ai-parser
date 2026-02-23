"""
Image Extractor Utility

Extracts base64 images from markdown content and replaces them with UUID references.
All images are converted to PNG format for consistency.
"""

import re
import uuid
import base64
import io
from typing import Tuple, Dict
from PIL import Image


def extract_and_replace_images(markdown_content: str) -> Tuple[str, Dict[str, str]]:
    """
    Extract base64 images from markdown, convert to PNG, and replace with UUID references.

    Args:
        markdown_content: Markdown content with inline base64 images

    Returns:
        Tuple of:
        - Modified markdown with UUID image references
        - Dictionary with UUID filenames as keys and base64 data URIs as values (all as PNG)
    """
    images = {}

    # Pattern to match markdown images: ![alt](data:image/format;base64,...)
    img_pattern = re.compile(r'!\[([^\]]*)\]\((data:image/(\w+);base64,([^)]+))\)')

    def replace_image(match):
        alt_text = match.group(1)
        full_data_uri = match.group(2)
        original_format = match.group(3).lower()
        base64_data = match.group(4)

        # Decode base64 to bytes
        try:
            img_bytes = base64.b64decode(base64_data)
        except Exception:
            # If base64 decode fails, skip this image
            img_bytes = None

        if img_bytes:
            # Convert to PNG using PIL
            try:
                with io.BytesIO(img_bytes) as input_buffer:
                    with Image.open(input_buffer) as img:
                        # Convert to RGB if necessary (for transparency handling)
                        if img.mode in ('RGBA', 'LA', 'P'):
                            # Create a white background for transparent images
                            if img.mode == 'RGBA' or (img.mode == 'P' and 'transparency' in img.info):
                                background = Image.new('RGB', img.size, (255, 255, 255))
                                if img.mode == 'P':
                                    img = img.convert('RGBA')
                                background.paste(img, mask=img.split()[3] if img.mode == 'RGBA' else None)
                                img = background
                            else:
                                img = img.convert('RGB')
                        elif img.mode not in ('RGB', 'L'):
                            img = img.convert('RGB')

                        # Save as PNG
                        output_buffer = io.BytesIO()
                        img.save(output_buffer, format='PNG', optimize=True)
                        png_bytes = output_buffer.getvalue()

                        # Encode to base64
                        png_base64 = base64.b64encode(png_bytes).decode('utf-8')
            except Exception:
                # If conversion fails, keep as is but still use PNG extension
                # This handles corrupted or truncated images
                png_base64 = base64_data
        else:
            # If base64 decode failed, keep original
            png_base64 = base64_data

        # Generate UUID for this image
        image_id = uuid.uuid4().hex[:16]
        image_filename = f"{image_id}.png"

        # Store the image mapping (always as PNG)
        images[image_filename] = f"data:image/png;base64,{png_base64}"

        # Replace with UUID reference
        return f"![{alt_text}]({image_filename})"

    # Replace all images in the content
    modified_content = img_pattern.sub(replace_image, markdown_content)

    return modified_content, images