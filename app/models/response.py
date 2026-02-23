from pydantic import BaseModel
from typing import Dict


class ParseResponse(BaseModel):
    filename: str
    file_type: str
    parsed_md_content: str
    processing_time: float
    images: Dict[str, str]


class ErrorResponse(BaseModel):
    detail: str
