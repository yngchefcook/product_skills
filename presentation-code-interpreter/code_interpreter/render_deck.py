#!/usr/bin/env python3
"""Render a PPTX into one PNG per slide using local conversion tools.

The pipeline converts PPTX to a temporary PDF with a headless office converter,
then rasterizes that PDF with a Poppler executable. No network is used.

Public API:
    ``render(deck, output_dir, dpi=160)`` returns the generated PNG paths.

CLI example:
    python3 render_deck.py presentation.pptx --output-dir /tmp/slides
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path


__all__ = ["render"]


def _run(command: list[str]) -> None:
    """Run one external command and raise a readable error on failure."""
    result = subprocess.run(
        command,
        check=False,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"Command failed ({result.returncode}): {' '.join(command)}\n{result.stdout}"
        )


def render(deck: Path, output_dir: Path, dpi: int = 160) -> list[Path]:
    """Render every slide in *deck* to ``slide-*.png`` files.

    Args:
        deck: Path to an existing PPTX.
        output_dir: Directory for PNG output; created when missing.
        dpi: Rasterization resolution, defaulting to 160.

    Returns:
        Sorted paths to the generated slide PNG files.

    Raises:
        RuntimeError: If the office converter or Poppler is unavailable, a
        conversion command fails, or no images are produced.
        OSError: If paths cannot be read or written.

    Notes:
        Existing ``slide-*.png`` files are not deleted automatically. Use a
        clean output directory when the slide count may have changed.
    """
    office = shutil.which("soffice") or shutil.which("libreoffice")
    rasterizer = shutil.which("pdftoppm") or shutil.which("pdftocairo")
    if not office:
        raise RuntimeError("LibreOffice/soffice is not installed.")
    if not rasterizer:
        raise RuntimeError("Poppler (pdftoppm or pdftocairo) is not installed.")
    deck = deck.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    # Keep the intermediate PDF isolated; only final PNG files leave the temp
    # directory.
    with tempfile.TemporaryDirectory(prefix="deck-render-") as tmp:
        tmp_dir = Path(tmp)
        _run(
            [
                office,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp_dir),
                str(deck),
            ]
        )
        pdf = tmp_dir / f"{deck.stem}.pdf"
        if not pdf.exists():
            pdf_candidates = list(tmp_dir.glob("*.pdf"))
            if len(pdf_candidates) != 1:
                raise RuntimeError("Office conversion did not produce a PDF.")
            pdf = pdf_candidates[0]
        prefix = output_dir / "slide"
        if Path(rasterizer).name == "pdftocairo":
            _run(
                [
                    rasterizer,
                    "-png",
                    "-r",
                    str(dpi),
                    str(pdf),
                    str(prefix),
                ]
            )
        else:
            _run(
                [
                    rasterizer,
                    "-png",
                    "-r",
                    str(dpi),
                    str(pdf),
                    str(prefix),
                ]
            )
    images = sorted(output_dir.glob("slide-*.png"))
    if not images:
        raise RuntimeError("No slide images were produced.")
    return images


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Render PPTX slides to PNG using office + Poppler tools.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Example:\n"
            "  python3 render_deck.py deck.pptx --output-dir /tmp/slides --dpi 180\n\n"
            "Requires: soffice/libreoffice and pdftoppm/pdftocairo."
        ),
    )
    parser.add_argument("deck", type=Path, help="Path to the PPTX file to render.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        required=True,
        help="Directory for slide-*.png files.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=160,
        help="Rasterization resolution (default: 160).",
    )
    args = parser.parse_args()
    images = render(args.deck, args.output_dir, args.dpi)
    print(f"Rendered {len(images)} slide(s) into {args.output_dir.resolve()}")


if __name__ == "__main__":
    main()
