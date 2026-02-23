# Simplified AI Parser

Lightweight document-to-Markdown conversion service.

## Features

- Converts DOCX, XLSX, XLS, XLSM, PDF, PPTX, PPT, and Markdown files to Markdown format
- Single synchronous endpoint
- Base64 embedded images
- No external dependencies (Redis, Azure Storage, etc.)

## Supported File Types

| Extension | Description |
|-----------|-------------|
| `.docx` | Microsoft Word documents |
| `.xlsx` | Microsoft Excel spreadsheets |
| `.xls` | Legacy Excel format (requires LibreOffice) |
| `.xlsm` | Excel with macros |
| `.pptx` | Microsoft PowerPoint presentations |
| `.ppt` | Legacy PowerPoint format (best effort) |
| `.pdf` | PDF documents |
| `.md`, `.markdown` | Markdown files (passthrough) |

## Quick Start

### Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run development server
python server.py
```

### Docker

```bash
# Build
docker build -t simplified-ai-parser .

# Run
docker run -p 7656:7656 simplified-ai-parser
```

## API

### POST /v1/parse-file

Convert a file to Markdown.

**Request:**
- Content-Type: `multipart/form-data`
- Body: `file` (binary upload)

**Response:**
```json
{
  "filename": "document.docx",
  "file_type": "docx",
  "parsed_md_content": "# Heading\n\nContent...",
  "processing_time": 1.234
}
```

### GET /health

Health check endpoint.

## System Requirements

- Python 3.11+
- mupdf-tools (for PDF conversion)
