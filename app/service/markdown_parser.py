"""
Markdown Parser

For markdown files, this is essentially a passthrough since the input
is already in the target format.

The original ai-parser converts markdown to HTML (using markdown-it),
then processes images. For a markdown-to-markdown pipeline, we keep
the content as-is but filter out non-base64 images for consistency.
"""

import re
import logging

logger = logging.getLogger(__name__)

SUPPORTED_IMAGE_FORMATS = ["png", "jpg", "jpeg"]


def _filter_images(content: str) -> str:
    """
    Filter markdown images to only keep base64 data URIs with supported formats.
    Matches original behavior that removes non-base64 images from markdown.
    """
    # Pattern to match markdown images: ![alt](src)
    img_pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')

    def process_image(match):
        alt = match.group(1)
        src = match.group(2)

        # Only keep base64 images (matching original behavior)
        if not src.startswith("data:image/"):
            logger.info(f"Removing non-base64 image: {src[:50]}...")
            return ""

        # Check format
        format_match = re.match(r"data:image/(\w+);base64", src)
        if format_match:
            img_format = format_match.group(1).lower()
            if img_format not in SUPPORTED_IMAGE_FORMATS:
                logger.info(f"Removing unsupported image format: {img_format}")
                return ""

        return match.group(0)

    return img_pattern.sub(process_image, content)


def parse_markdown(file_path: str) -> str:
    """
    Read a Markdown file and return its content.

    Processing (matching original behavior):
    - Read the file content
    - Filter out non-base64 images
    - Filter out unsupported image formats (only png/jpg/jpeg allowed)

    Args:
        file_path: Path to the Markdown file

    Returns:
        Markdown content as string
    """
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Filter images to match original behavior
    content = _filter_images(content)

    return content
