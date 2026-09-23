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

PAGES = {  # first line of the page -> image name (continuation pages are left out)
    "Portland General Electric": "cover",
    "Portland General Electric: 2026 retail revenue, plan and outlook": "dashboard",
    "Profit and loss: Portland General Electric, electric utility (FERC Form 1)": "pnl",
    "Price, volume and mix: where retail revenue growth came from": "pvm",
    "Monthly retail sales: EIA-861M, and how it ties to the FERC filing": "monthly",
    "Weather: how much of the load is the thermostat": "weather",
    "Plan against actual: the 2026 plan, built only from data through December 2025": "plan",
    "Rolling forecast: June 2026 actuals, 6 months forecast": "forecast",
    "Scenarios: what moves the 2026 outlook": "scenarios",
    "Peers: PGE against four Pacific Northwest utilities (FERC Form 1)": "peers",
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
