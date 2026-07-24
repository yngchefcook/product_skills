#!/usr/bin/env python3
"""Runtime helpers shared by the portable presentation builder.

This module contains only environment, file, image-download, and QA utilities.
It deliberately avoids slide-specific business logic so ``builder.py`` remains
the single public entry point for deck creation.

All paths are resolved relative to this file. The bundle therefore works after
uploading or extracting it into any Code Interpreter directory.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import re
import urllib.error
import urllib.parse
import urllib.request
from io import BytesIO
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw


BUNDLE_DIR = Path(__file__).resolve().parent
IMAGES_DIR = BUNDLE_DIR / "images"
THEMES_PATH = BUNDLE_DIR / "themes.json"
CATALOG_PATH = BUNDLE_DIR / "image_catalog.json"

REMOTE_IMAGE_TIMEOUT_SECONDS = 15
REMOTE_IMAGE_MAX_BYTES = 25 * 1024 * 1024

__all__ = [
    "BUNDLE_DIR",
    "CATALOG_PATH",
    "IMAGES_DIR",
    "THEMES_PATH",
    "assert_bundle_complete",
    "create_contact_sheet",
    "download_remote_image",
    "environment_report",
    "is_http_url",
    "load_json",
    "sanitize_filename",
]


def load_json(path: str | Path) -> dict:
    """Read a UTF-8 JSON object from *path*."""
    return json.loads(Path(path).read_text(encoding="utf-8"))


def sanitize_filename(value: str, default: str = "presentation") -> str:
    """Return a portable ``.pptx`` basename.

    Directory components, spaces, Cyrillic characters, and unsafe punctuation
    are removed or replaced. The visible presentation title is unaffected.
    """
    name = Path(value or f"{default}.pptx").name
    stem = Path(name).stem
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-") or default
    return f"{stem}.pptx"


def is_http_url(value: object) -> bool:
    """Return whether *value* is an absolute HTTP(S) URL."""
    if not isinstance(value, str):
        return False
    parsed = urllib.parse.urlparse(value.strip())
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def download_remote_image(
    url: str,
    directory: str | Path,
    *,
    timeout: int = REMOTE_IMAGE_TIMEOUT_SECONDS,
    max_bytes: int = REMOTE_IMAGE_MAX_BYTES,
) -> Path:
    """Download, validate, and normalize one remote raster image.

    Args:
        url: Absolute HTTP(S) URL.
        directory: Temporary destination directory.
        timeout: Network timeout in seconds.
        max_bytes: Maximum response size.

    Returns:
        Path to a validated PNG or JPEG.

    Raises:
        RuntimeError: If the response is unavailable, too large, redirected to
            an unsupported scheme, or not a valid raster image.
    """
    if not is_http_url(url):
        raise RuntimeError(f"Remote image URL must use http or https: {url!r}")
    destination = Path(directory)
    destination.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "presentation-code-interpreter/1.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            final_url = response.geturl()
            if not is_http_url(final_url):
                raise RuntimeError(
                    f"Remote image redirected to an unsupported URL: {final_url!r}"
                )
            raw_length = response.headers.get("Content-Length")
            if raw_length and int(raw_length) > max_bytes:
                raise RuntimeError(
                    f"Remote image exceeds {max_bytes // (1024 * 1024)} MiB."
                )
            payload = bytearray()
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                payload.extend(chunk)
                if len(payload) > max_bytes:
                    raise RuntimeError(
                        f"Remote image exceeds {max_bytes // (1024 * 1024)} MiB."
                    )
    except (OSError, ValueError, urllib.error.URLError) as exc:
        raise RuntimeError(f"Could not download remote image {url!r}: {exc}") from exc

    digest = hashlib.sha256(url.encode("utf-8")).hexdigest()[:16]
    try:
        with Image.open(BytesIO(payload)) as image:
            image.load()
            image_format = (image.format or "").upper()
            if image.width < 1 or image.height < 1:
                raise ValueError("image has invalid dimensions")
            if image_format in {"PNG", "JPEG"}:
                suffix = ".png" if image_format == "PNG" else ".jpg"
                target = destination / f"remote-{digest}{suffix}"
                target.write_bytes(payload)
            else:
                target = destination / f"remote-{digest}.png"
                mode = "RGBA" if "A" in image.getbands() else "RGB"
                image.convert(mode).save(target, format="PNG")
    except (OSError, ValueError) as exc:
        raise RuntimeError(f"Remote URL is not a valid raster image: {url!r}") from exc
    return target


def assert_bundle_complete() -> None:
    """Raise ``FileNotFoundError`` when a required uploaded resource is absent."""
    required = [
        BUNDLE_DIR / "builder.py",
        BUNDLE_DIR / "helpers.py",
        BUNDLE_DIR / "image_filenames.py",
        THEMES_PATH,
        CATALOG_PATH,
        IMAGES_DIR,
    ]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise FileNotFoundError(
            "Incomplete presentation bundle:\n- " + "\n- ".join(missing)
        )


def environment_report() -> dict:
    """Return dependency and bundle diagnostics without changing the environment."""
    modules = {
        "Pillow": importlib.util.find_spec("PIL") is not None,
        "python-pptx": importlib.util.find_spec("pptx") is not None,
    }
    return {
        "bundle_dir": str(BUNDLE_DIR),
        "images": len(list(IMAGES_DIR.glob("*"))) if IMAGES_DIR.exists() else 0,
        "themes_file": THEMES_PATH.exists(),
        "catalog_file": CATALOG_PATH.exists(),
        "modules": modules,
        "ready": all(modules.values())
        and THEMES_PATH.exists()
        and CATALOG_PATH.exists(),
    }


def create_contact_sheet(
    image_paths: Iterable[str | Path],
    output_path: str | Path,
    *,
    columns: int = 4,
    thumb_width: int = 480,
    gutter: int = 24,
) -> Path:
    """Create a labeled contact sheet from rendered slide images.

    The contact sheet helps inspect deck-level rhythm. It does not replace
    checking each slide at full resolution.
    """
    paths = [Path(path) for path in image_paths]
    if not paths:
        raise ValueError("At least one image is required for a contact sheet.")
    columns = max(1, int(columns))
    prepared: list[Image.Image] = []
    for path in paths:
        with Image.open(path) as source:
            image = source.convert("RGB")
            height = round(image.height * thumb_width / max(image.width, 1))
            prepared.append(
                image.resize((thumb_width, height), Image.Resampling.LANCZOS)
            )
    cell_height = max(image.height for image in prepared) + 42
    rows = (len(prepared) + columns - 1) // columns
    width = gutter + columns * (thumb_width + gutter)
    height = gutter + rows * (cell_height + gutter)
    canvas = Image.new("RGB", (width, height), "#E8EBEF")
    draw = ImageDraw.Draw(canvas)
    for index, image in enumerate(prepared):
        row, column = divmod(index, columns)
        x = gutter + column * (thumb_width + gutter)
        y = gutter + row * (cell_height + gutter)
        canvas.paste(image, (x, y))
        draw.text((x, y + image.height + 10), f"{index + 1:02d}", fill="#1A2230")
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, format="PNG")
    return output.resolve()
