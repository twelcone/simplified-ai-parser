"""
DOCX to Markdown Parser

Uses mammoth to convert DOCX to HTML, then markdownify to convert to Markdown.
This matches the original ai-parser DOCX conversion flow.
"""

import base64
import re
import mammoth
from bs4 import BeautifulSoup
from markdownify import markdownify


SUPPORTED_IMAGE_FORMATS = ["png", "jpg", "jpeg"]
EMBEDDED_OBJECT_SRC = "embedded_object_src"
EMBEDDED_OBJECT_ICON = "📎"


def _convert_image(image):
    """
    Convert embedded image to base64 data URI.
    Matches original behavior from document_parser.py:1378-1383
    """
    # EMF/WMF are embedded objects, not actual images - mark them for replacement
    if image.content_type in ["image/x-emf", "image/x-wmf"]:
        return {"src": EMBEDDED_OBJECT_SRC}

    with image.open() as image_file:
        image_data = base64.b64encode(image_file.read()).decode("utf-8")
    return {"src": f"data:{image.content_type};base64,{image_data}"}


def _replace_embedded_object_with_icon(html_content: str) -> str:
    """
    Replace embedded object placeholders with icon.
    Matches original document_parser.py:1375-1376
    """
    pattern = re.compile(rf'<img[^>]*{EMBEDDED_OBJECT_SRC}[^>]*>')
    return pattern.sub(EMBEDDED_OBJECT_ICON, html_content)


def _clean_html(html_content: str) -> str:
    """
    Clean HTML content by removing unnecessary elements.
    Matches original document_parser.py _clean_html method.
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


def _filter_unsupported_images(html_content: str) -> str:
    """
    Remove images with unsupported formats from HTML.
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

        if extension and extension not in SUPPORTED_IMAGE_FORMATS:
            img.decompose()

    return str(soup)


def parse_docx_to_markdown(file_path: str) -> str:
    """
    Convert a DOCX file to Markdown format.

    Processing steps (matching original document_parser.py):
    1. mammoth.convert_to_html with image conversion
    2. Replace embedded objects (EMF/WMF) with 📎 icon
    3. Clean HTML (remove styles, scripts, layout attrs)
    4. Filter unsupported image formats
    5. Convert HTML to Markdown

    Args:
        file_path: Path to the DOCX file

    Returns:
        Markdown content as string
    """
    # Step 1: Convert to HTML (matching original line 1094-1097)
    result = mammoth.convert_to_html(
        file_path,
        convert_image=mammoth.images.img_element(_convert_image)
    )
    html_content = result.value

    # Step 2: Replace embedded object placeholders with icon (matching original line 1099)
    html_content = _replace_embedded_object_with_icon(html_content)

    # Step 3: Clean HTML (matching original line 1105)
    html_content = _clean_html(html_content)

    # Step 4: Filter unsupported images (matching original line 1106-1107)
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
