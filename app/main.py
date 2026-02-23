from fastapi import FastAPI
from app.api.parse_route import router as parse_router

app = FastAPI(
    title="Simplified AI Parser",
    description="Lightweight document-to-Markdown conversion service",
    version="0.1.0",
)

app.include_router(parse_router, prefix="/v1")


@app.get("/health")
async def health_check():
    return {"status": "healthy"}
