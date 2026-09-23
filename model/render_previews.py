"""Render the README images from the board-pack PDF Excel exported.

    python model/render_previews.py        # needs pymupdf and pillow

Images come from Excel's own PDF export rather than screenshots, so what the
README shows is exactly what the workbook prints.
"""
from __future__ import annotations

from pathlib import Path

import pymupdf
from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parent.parent
PDF = ROOT / "docs" / "board_pack.pdf"
OUT = ROOT / "docs" / "img"

PAGES = {  # first line of the page -> image name
    "Kestrel Bay Provisions": "cover",
    "Kestrel Bay Provisions: FY2026 performance and outlook": "dashboard",
    "Profit and loss, FY2026: outlook against budget and prior year": "pnl",
    "Price-volume-mix: why margin moved against budget": "pvm",
    "Rolling forecast by driver": "forecast",
    "Scenarios, sensitivities and the gap to budget": "scenarios",
    "Working capital: cash tied up, and stock close to its shelf life": "working-capital",
    "Channel economics: where the growth came from, and what it earned": "channels",
    "Checks: the model proves its own numbers": "checks",
}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for page in pymupdf.open(PDF):
        name = PAGES.get(page.get_text().strip().split("\n")[0])
        if not name:
            continue
        pix = page.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        box = ImageChops.difference(img, Image.new("RGB", img.size, "white")).getbbox()
        bottom = min(box[3], int(img.height * 0.93))          # leave the page footer out
        img = img.crop((max(box[0] - 20, 0), max(box[1] - 20, 0), min(box[2] + 20, img.width), bottom + 10))
        img.save(OUT / f"{name}.png", optimize=True)
        print(f"docs/img/{name}.png {img.size[0]}x{img.size[1]}")


if __name__ == "__main__":
    main()
