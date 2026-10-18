import os
import re
from pathlib import Path

from fastapi import HTTPException

from src.constants import GENERATED_IMAGES_DIR


GENERATED_IMAGE_DIR = Path(GENERATED_IMAGES_DIR)
GENERATED_IMAGE_RE = re.compile(
    r"^[a-zA-Z0-9_-]{8,64}\.(png|jpg|jpeg|webp|gif|mp4|mov|webm|mkv|m4v)$"
)
GENERATED_IMAGE_HEADERS = {
    "Cache-Control": "public, max-age=31536000, immutable",
    "X-Content-Type-Options": "nosniff",
}


def resolve_generated_image_path(filename: str) -> Path:
    if not isinstance(filename, str) or not GENERATED_IMAGE_RE.fullmatch(filename):
        raise HTTPException(status_code=400, detail="Invalid filename")
    root = GENERATED_IMAGE_DIR.resolve()
    path = (GENERATED_IMAGE_DIR / filename).resolve()
    try:
        if os.path.commonpath([str(root), str(path)]) != str(root):
            raise ValueError
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid filename")
    if not path.exists():
        raise HTTPException(status_code=404, detail="Image not found")
    return path


def save_generated_image_bytes(image_bytes: bytes, filename: str) -> str:
    """Save raw image bytes into GENERATED_IMAGE_DIR and return the full file path."""
    GENERATED_IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    clean_name = os.path.basename(filename)
    target_path = GENERATED_IMAGE_DIR / clean_name
    target_path.write_bytes(image_bytes)
    return str(target_path)

