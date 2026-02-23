"""
PDF to Markdown Parser

Uses mutool (from mupdf-tools) to convert PDF to HTML,
then markdownify to convert HTML to Markdown.

This matches the original ai-parser PDF conversion flow.
"""

import os
import re
import base64
import tempfile
import logging
from subprocess import run, CalledProcessError
from bs4 import BeautifulSoup
from markdownify import markdownify

logger = logging.getLogger(__name__)

SUPPORTED_IMAGE_FORMATS = ["png", "jpg", "jpeg"]


def _clean_html(html_content: str) -> str:
    """
    Clean HTML content by removing unnecessary elements.
    Matches original document_parser.py _clean_html behavior.
    """
    soup = BeautifulSoup(html_content, "html.parser")

    # Remove head tag
    head = soup.head
    if head:
        head.decompose()

    # Remove style, script, video, audio tags
    for tag_name in ["style", "script", "video", "audio"]:
        for tag in soup.find_all(tag_name):
            tag.decompose()

    # Remove style attributes
    for element in soup.find_all(style=True):
        del element["style"]

    # Remove layout attributes (matching original list)
    attrs_to_remove = [
        "align", "valign", "bgcolor", "sdval", "sdnum",
        "height", "width", "cellspacing", "border", "span",
        "hspace", "vspace", "data-sheets-value",
        "data-sheets-numberformat", "data-sheets-formula"
    ]
    for attr in attrs_to_remove:
        for element in soup.find_all(attrs={attr: True}):
            del element[attr]

    # Unwrap font tags
    for font in soup.find_all("font"):
        font.unwrap()

    # Remove comment indicators
    for indicator in soup.find_all(class_="comment-indicator"):
        indicator.decompose()

    # Remove HTML comments
    html_without_comments = re.sub(r"<!--[\s\S]*?-->", "", str(soup))

    return html_without_comments


def _replace_images_with_base64(html_content: str, base_path: str) -> str:
    """
    Replace image src with base64 data URIs.
    Matches original document_parser.py _replace_images_with_base64.
    """
    soup = BeautifulSoup(html_content, "html.parser")
    all_image_paths = set()

    for img in soup.find_all("img"):
        src = img.get("src")
        if src and not src.startswith("data:image/"):
            image_path = os.path.abspath(os.path.join(base_path, src))
            if os.path.exists(image_path):
                try:
                    with open(image_path, "rb") as f:
                        image_data = f.read()
                    ext = os.path.splitext(image_path)[1][1:].lower()
                    if ext == "jpg":
                        ext = "jpeg"
                    base64_data = base64.b64encode(image_data).decode()
                    img["src"] = f"data:image/{ext};base64,{base64_data}"
                    all_image_paths.add(image_path)
                except Exception as e:
                    logger.warning(f"Failed to process image {image_path}: {e}")

    # Clean up temporary image files (matching original behavior)
    for image_path in all_image_paths:
        try:
            os.remove(image_path)
        except Exception:
            pass

    return str(soup)


def _filter_unsupported_images(html_content: str) -> str:
    """
    Remove images with unsupported formats.
    Matches original _contain_unsupported_image behavior.
    Only keeps png, jpg, jpeg.
    """
    soup = BeautifulSoup(html_content, "html.parser")

    for img in soup.find_all("img"):
        src = img.get("src")
        if not src:
            continue

        extension = ""
        if src.startswith("data:image/"):
            # Extract format from base64 data URI
            matches = re.match(r"^data:image/(\w+);base64", src)
            if matches:
                extension = matches.group(1).lower()
        else:
            # Extract from file extension
            ext_matches = re.search(r"\.([a-zA-Z0-9]+)(\?|$)", src)
            if ext_matches:
                extension = ext_matches.group(1).lower()

        if extension and extension not in SUPPORTED_IMAGE_FORMATS:
            logger.info(f"Removing unsupported image format: {extension}")
            img.decompose()

    return str(soup)


def parse_pdf_to_markdown(file_path: str) -> str:
    """
    Convert a PDF file to Markdown format.

    Processing steps (matching original document_parser.py):
    1. mutool convert PDF to HTML with preserved images
    2. Replace image file references with base64 data URIs
    3. Clean HTML (remove styles, scripts, etc.)
    4. Filter unsupported image formats
    5. Convert HTML to Markdown

    Args:
        file_path: Path to the PDF file

    Returns:
        Markdown content as string
    """
    # Create a temporary directory for the HTML output
    with tempfile.TemporaryDirectory() as temp_dir:
        basename = os.path.splitext(os.path.basename(file_path))[0]
        output_html_path = os.path.join(temp_dir, f"{basename}.html")

        # Step 1: Convert PDF to HTML using mutool (matching original)
        try:
            run(
                [
                    "mutool",
                    "convert",
                    "-F", "html",
                    "-O", "preserve-images",
                    "-o", output_html_path,
                    file_path,
                ],
                check=True,
                capture_output=True
            )
            logger.info(f"PDF {basename} converted to HTML successfully")
        except CalledProcessError as e:
            logger.error(f"mutool conversion failed: {e}")
            raise RuntimeError(
                f"Failed to convert PDF: {e.stderr.decode() if e.stderr else str(e)}"
            )
        except FileNotFoundError:
            raise RuntimeError(
                "mutool not found. Please install mupdf-tools: "
                "apt-get install mupdf-tools"
            )

        # Read the generated HTML
        with open(output_html_path, "r", encoding="utf-8") as f:
            html_content = f.read()

        # Step 2: Replace images with base64 data URIs (matching original)
        html_content = _replace_images_with_base64(html_content, temp_dir)

        # Step 3: Clean the HTML (matching original _clean_html)
        html_content = _clean_html(html_content)

        # Step 4: Filter unsupported images (matching original _contain_unsupported_image)
        html_content = _filter_unsupported_images(html_content)

        # Step 5: Convert HTML to Markdown
        markdown_content = markdownify(
            html_content,
            heading_style="ATX",
            bullets="-",
            strip=["style", "script"]
        )

        # Clean up extra whitespace
        lines = markdown_content.split("\n")
        cleaned_lines = []
        prev_empty = False
        for line in lines:
            is_empty = not line.strip()
            if is_empty and prev_empty:
                continue
            cleaned_lines.append(line)
            prev_empty = is_empty

        return "\n".join(cleaned_lines).strip()
