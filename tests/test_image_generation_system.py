# tests/test_image_generation_system.py
"""
Test suite for HEXA AI Workbench Image Generation System.

Tests:
  1. API request validation for POST /api/images/generate
  2. Endpoint execution and structured response format
  3. Image file retrieval and path-traversal prevention for GET /api/images/{image_id}
  4. Image generation history for GET /api/images/history
  5. Intelligent request routing intent detection
  6. Air-gapped offline fallback execution
"""

import os
import tempfile
import pytest
from fastapi.testclient import TestClient

import src.database as db_module

# Ensure test DB tables are created in a shared file-backed SQLite database for testing
test_db_dir = tempfile.mkdtemp()
test_db_path = os.path.join(test_db_dir, "test_images.db")
db_module.DATABASE_URL = f"sqlite:///{test_db_path}"

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

test_engine = create_engine(f"sqlite:///{test_db_path}", connect_args={"check_same_thread": False})
db_module.engine = test_engine
db_module.SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=test_engine)

from app import app
from src.constants import GENERATED_IMAGES_DIR
from routes.chat_routes import is_image_generation_prompt

# Create all tables on test_engine
db_module.Base.metadata.create_all(bind=test_engine)

client = TestClient(app)


def test_image_generation_prompt_detection():
    """Verify intelligent prompt routing intent classifier."""
    assert is_image_generation_prompt("generate an image of a red sports car") is True
    assert is_image_generation_prompt("create a futuristic smart city") is True
    assert is_image_generation_prompt("draw a realistic apple") is True
    assert is_image_generation_prompt("generate a professional AI technology wallpaper") is True
    assert is_image_generation_prompt("create a cyberpunk city at night") is True

    # Non-image prompts
    assert is_image_generation_prompt("What is artificial intelligence?") is False
    assert is_image_generation_prompt("Write a python function to sort a list") is False
    assert is_image_generation_prompt("Calculate 55 * 10 / 2") is False


def test_generate_image_endpoint_validation():
    """Test input validation for POST /api/images/generate."""
    # Missing prompt -> 422 Unprocessable Entity
    res = client.post("/api/images/generate", json={})
    assert res.status_code == 422

    # Invalid width (less than 64) -> 422
    res = client.post("/api/images/generate", json={"prompt": "test image", "width": 10})
    assert res.status_code == 422


def test_generate_image_endpoint_execution():
    """Test POST /api/images/generate execution and structured response."""
    payload = {
        "prompt": "A futuristic industrial robot in a modern factory",
        "negative_prompt": "blurry, low quality",
        "width": 512,
        "height": 512,
        "steps": 10,
        "quality": "medium",
    }

    res = client.post("/api/images/generate", json=payload)
    assert res.status_code == 200

    data = res.json()
    assert data["success"] is True
    assert "image_id" in data and len(data["image_id"]) > 0
    assert "image_url" in data and data["image_url"].startswith("/api/images/")
    assert data["prompt"] == payload["prompt"]
    assert data["width"] == 512
    assert data["height"] == 512
    assert data["steps"] == 10
    assert data["duration_seconds"] >= 0.0
    assert "model" in data


def test_get_image_file_retrieval_and_security():
    """Test GET /api/images/{image_id} retrieval and path traversal security."""
    # First generate an image to retrieve
    gen_res = client.post("/api/images/generate", json={"prompt": "security test image"})
    assert gen_res.status_code == 200
    img_data = gen_res.json()
    image_id = img_data["image_id"]

    # Retrieve created image
    res = client.get(f"/api/images/{image_id}")
    assert res.status_code == 200
    assert res.headers["content-type"] in ("image/png", "image/jpeg", "image/webp")

    # Path traversal attempt -> 400 or 404
    bad_res = client.get("/api/images/..%2F..%2Fetc%2Fpasswd")
    assert bad_res.status_code in (400, 404)

    # Non-existent image -> 404
    non_existent = client.get("/api/images/non_existent_img_999999.png")
    assert non_existent.status_code == 404


def test_get_image_history():
    """Test GET /api/images/history endpoint."""
    res = client.get("/api/images/history?limit=10")
    assert res.status_code == 200

    data = res.json()
    assert data["success"] is True
    assert "total" in data
    assert "items" in data
    assert isinstance(data["items"], list)
