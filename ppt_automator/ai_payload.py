from __future__ import annotations

from io import BytesIO
from pathlib import Path
import base64
import os


def image_data_url(path: Path, max_side: int | None = None) -> str:
    image_bytes = _optimized_image_bytes(path, max_side=max_side)
    return "data:image/png;base64," + base64.b64encode(image_bytes).decode("ascii")


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name, "").strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on", "sim"}


def env_int(name: str, default: int) -> int:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def _optimized_image_bytes(path: Path, max_side: int | None = None) -> bytes:
    max_side = max_side or env_int("AUTO_PPT_AI_IMAGE_MAX_SIDE", 1400)
    try:
        from PIL import Image
    except Exception:
        return path.read_bytes()

    try:
        image = Image.open(path).convert("RGB")
        width, height = image.size
        largest = max(width, height)
        if max_side > 0 and largest > max_side:
            scale = max_side / largest
            image = image.resize((max(1, int(width * scale)), max(1, int(height * scale))), Image.Resampling.LANCZOS)
        stream = BytesIO()
        image.save(stream, format="PNG", optimize=True)
        return stream.getvalue()
    except Exception:
        return path.read_bytes()
