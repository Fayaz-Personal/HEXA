# routes/images_routes.py
"""
HEXA Sovereign Industrial AI Workbench — Image Generation Routes

Exposes API endpoints for local image generation, retrieval, and history:
  POST /api/images/generate
  GET  /api/images/{image_id}
  GET  /api/images/history
"""

import os
import time
import uuid
import logging
from pathlib import Path
from typing import Optional, List, Dict, Any

from fastapi import APIRouter, Request, HTTPException, Depends
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

from src.constants import GENERATED_IMAGES_DIR
from src.generated_images import resolve_generated_image_path, GENERATED_IMAGE_HEADERS
from src.settings import load_settings, get_setting
from src.database import SessionLocal, GalleryImage
from src.ai_interaction import do_generate_image

logger = logging.getLogger("hexa.images_routes")

router = APIRouter(prefix="/api/images", tags=["images"])


# ── Pydantic Request / Response Models ───────────────────────────────────────

class ImageGenerateRequest(BaseModel):
    prompt: str = Field(..., min_length=1, description="Natural language image prompt")
    negative_prompt: Optional[str] = Field(None, description="Negative prompt to exclude unwanted concepts")
    width: int = Field(512, ge=64, le=2048, description="Image width in pixels")
    height: int = Field(512, ge=64, le=2048, description="Image height in pixels")
    steps: int = Field(20, ge=1, le=150, description="Number of inference steps")
    quality: str = Field("medium", description="Quality preset: low, medium, high, auto")
    model: Optional[str] = Field(None, description="Model specifier (optional, defaults to settings)")
    session_id: Optional[str] = Field(None, description="Optional linked chat session ID")


class ImageGenerateResponse(BaseModel):
    success: bool
    image_id: str
    image_url: str
    prompt: str
    negative_prompt: Optional[str] = None
    width: int
    height: int
    steps: int
    duration_seconds: float
    model: str
    error: Optional[str] = None


# ── Helpers ──────────────────────────────────────────────────────────────────

def _sanitize_image_identifier(identifier: str) -> str:
    """Sanitize and restrict filename to prevent path traversal attacks."""
    clean = os.path.basename(str(identifier or "").strip())
    if not clean or clean.startswith(".") or ".." in clean:
        raise HTTPException(status_code=400, detail="Invalid image identifier")
    return clean


# ── API Endpoints ────────────────────────────────────────────────────────────

@router.post("/generate", response_model=ImageGenerateResponse)
async def generate_image_endpoint(req: ImageGenerateRequest, request: Request):
    """Generate an image using the local image generation pipeline.
    
    Supports offline air-gapped processing, configurable dimensions, negative prompts,
    and automatic local saving.
    """
    if not get_setting("image_gen_enabled", True):
        raise HTTPException(status_code=403, detail="Image generation is disabled in settings")

    start_time = time.time()
    image_id = f"img_{uuid.uuid4().hex[:12]}"
    filename = f"{image_id}.png"

    # Format instruction payload for generation engine cleanly without raw linebreaks in prompt
    prompt_clean = req.prompt.strip().replace("\n", " ")
    if req.negative_prompt:
        prompt_clean += f" (Negative: {req.negative_prompt.strip().replace('\n', ' ')})"
    
    model_spec = req.model or get_setting("image_model", "") or "local-diffusers"
    size_str = f"{req.width}x{req.height}"
    
    # Payload string for do_generate_image
    gen_content = f"{prompt_clean}\n{model_spec}\n{size_str}\n{req.quality}"

    try:
        res = await do_generate_image(gen_content, session_id=req.session_id)
    except Exception as e:
        logger.error(f"Image generation failed: {e}", exc_info=True)
        duration = round(time.time() - start_time, 2)
        return ImageGenerateResponse(
            success=False,
            image_id=image_id,
            image_url="",
            prompt=req.prompt,
            negative_prompt=req.negative_prompt,
            width=req.width,
            height=req.height,
            steps=req.steps,
            duration_seconds=duration,
            model=model_spec,
            error=str(e),
        )

    if "error" in res and res["error"]:
        duration = round(time.time() - start_time, 2)
        return ImageGenerateResponse(
            success=False,
            image_id=image_id,
            image_url="",
            prompt=req.prompt,
            negative_prompt=req.negative_prompt,
            width=req.width,
            height=req.height,
            steps=req.steps,
            duration_seconds=duration,
            model=model_spec,
            error=str(res["error"]),
        )

    image_url = res.get("image_url", "")
    returned_model = res.get("image_model") or model_spec

    # Extract actual filename from URL or path
    if image_url:
        actual_filename = os.path.basename(image_url.split("?")[0])
        if actual_filename:
            image_id = actual_filename.rsplit(".", 1)[0]
            filename = actual_filename

    # Save to SQLite Gallery database for history tracking
    try:
        db = SessionLocal()
        gallery_entry = GalleryImage(
            id=image_id,
            filename=filename,
            prompt=req.prompt,
            model=returned_model,
            size=size_str,
            quality=req.quality,
        )
        db.add(gallery_entry)
        db.commit()
        db.close()
    except Exception as db_err:
        logger.warning(f"Failed to record generated image in gallery database: {db_err}")

    duration = round(time.time() - start_time, 2)
    api_image_url = f"/api/images/{filename}"

    return ImageGenerateResponse(
        success=True,
        image_id=image_id,
        image_url=api_image_url,
        prompt=req.prompt,
        negative_prompt=req.negative_prompt,
        width=req.width,
        height=req.height,
        steps=req.steps,
        duration_seconds=duration,
        model=returned_model,
        error=None,
    )


@router.get("/history")
async def get_image_history(limit: int = 50, offset: int = 0):
    """Retrieve history of locally generated images."""
    db = SessionLocal()
    try:
        query = db.query(GalleryImage).order_by(GalleryImage.created_at.desc())
        total = query.count()
        items = query.offset(offset).limit(min(limit, 100)).all()

        results = []
        for item in items:
            filename = item.filename or f"{item.id}.png"
            results.append({
                "id": item.id,
                "filename": filename,
                "prompt": item.prompt,
                "model": item.model,
                "size": item.size,
                "quality": item.quality,
                "created_at": item.created_at.isoformat() if hasattr(item.created_at, "isoformat") else str(item.created_at),
                "image_url": f"/api/images/{filename}",
            })

        return {
            "success": True,
            "total": total,
            "items": results,
        }
    except Exception as e:
        logger.error(f"Failed to fetch image history: {e}")
        return {"success": False, "error": str(e), "items": []}
    finally:
        db.close()


@router.get("/{image_id}")
async def get_image_file(image_id: str, request: Request):
    """Serve a generated image file by image_id or filename.
    
    Enforces strict path traversal checks to restrict access exclusively to
    the generated_images directory.
    """
    clean_name = _sanitize_image_identifier(image_id)
    if not clean_name.endswith((".png", ".jpg", ".jpeg", ".webp")):
        clean_name = f"{clean_name}.png"

    file_path = resolve_generated_image_path(clean_name)
    if not file_path or not os.path.isfile(file_path):
        raise HTTPException(status_code=404, detail="Image not found")

    ext = os.path.splitext(file_path)[1].lower()
    media_type = "image/png" if ext == ".png" else "image/jpeg" if ext in (".jpg", ".jpeg") else "image/webp"

    return FileResponse(
        file_path,
        media_type=media_type,
        headers=GENERATED_IMAGE_HEADERS,
    )
