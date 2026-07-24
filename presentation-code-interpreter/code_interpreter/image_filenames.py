#!/usr/bin/env python3
"""Stable aliases and lookup helpers for bundled presentation imagery.

Use IDs such as ``"ai_neural_orbit"`` in slide specs. The builder resolves
them through this module, so presentation code never depends on upload paths.
"""

from __future__ import annotations

from pathlib import Path

try:
    from .helpers import BUNDLE_DIR, CATALOG_PATH, load_json
except ImportError:
    from helpers import BUNDLE_DIR, CATALOG_PATH, load_json


_CATALOG = load_json(CATALOG_PATH)
IMAGE_META: dict[str, dict] = {
    str(item["id"]): dict(item) for item in _CATALOG.get("items", [])
}
IMAGE_FILES: dict[str, str] = {
    image_id: str((BUNDLE_DIR / item["path"]).resolve())
    for image_id, item in IMAGE_META.items()
}

# Explicit constants improve discoverability in interactive Python sessions.
AI_NEURAL_ORBIT = IMAGE_FILES["ai_neural_orbit"]
AI_HUMAN_COLLABORATION = IMAGE_FILES["ai_human_collaboration"]
CLOUD_COMPUTE_LANDSCAPE = IMAGE_FILES["cloud_compute_landscape"]
CYBERSECURITY_VAULT = IMAGE_FILES["cybersecurity_vault"]
FINANCE_MARKET_FLOW = IMAGE_FILES["finance_market_flow"]
FINANCE_RISK_BALANCE = IMAGE_FILES["finance_risk_balance"]
INVESTMENT_HORIZON = IMAGE_FILES["investment_horizon"]
STRATEGIC_CHOICES = IMAGE_FILES["strategic_choices"]
SUSTAINABLE_FUTURE = IMAGE_FILES["sustainable_future"]
CLOSING_HORIZON = IMAGE_FILES["closing_horizon"]

__all__ = [
    "AI_HUMAN_COLLABORATION",
    "AI_NEURAL_ORBIT",
    "CLOSING_HORIZON",
    "CLOUD_COMPUTE_LANDSCAPE",
    "CYBERSECURITY_VAULT",
    "FINANCE_MARKET_FLOW",
    "FINANCE_RISK_BALANCE",
    "IMAGE_FILES",
    "IMAGE_META",
    "INVESTMENT_HORIZON",
    "STRATEGIC_CHOICES",
    "SUSTAINABLE_FUTURE",
    "catalog_payload",
    "list_images",
    "resolve_image",
]


def catalog_payload() -> dict:
    """Return a copy-safe image catalog payload."""
    return {
        **_CATALOG,
        "items": [dict(item) for item in _CATALOG.get("items", [])],
    }


def resolve_image(value: str | Path | None) -> Path | None:
    """Resolve an image ID or existing filesystem path."""
    if value is None:
        return None
    candidate = Path(str(value)).expanduser()
    if candidate.is_file():
        return candidate.resolve()
    mapped = IMAGE_FILES.get(str(value))
    if mapped and Path(mapped).is_file():
        return Path(mapped)
    return None


def list_images(
    *, keyword: str | None = None, recommended_use: str | None = None
) -> list[dict]:
    """List bundled images, optionally filtered by keyword or slide type."""
    results = []
    for item in _CATALOG.get("items", []):
        keywords = {str(value).lower() for value in item.get("keywords", [])}
        uses = {str(value).lower() for value in item.get("recommended_use", [])}
        if keyword and keyword.lower() not in keywords:
            continue
        if recommended_use and recommended_use.lower() not in uses:
            continue
        results.append(dict(item))
    return results
