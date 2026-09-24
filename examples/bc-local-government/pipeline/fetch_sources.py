"""Download and fingerprint the official B.C. local-government workbooks."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import urllib.request
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
MANIFEST = ROOT / "data" / "source_manifest.json"
BASE_URL = (
    "https://www2.gov.bc.ca/assets/gov/british-columbians-our-governments/"
    "local-governments/finance/local-government-statistics"
)
SOURCE_PAGE = (
    "https://www2.gov.bc.ca/gov/content/governments/local-governments/"
    "facts-framework/statistics/statistics"
)
SCHEDULES = ("201", "301", "302", "304", "401", "402", "502", "503", "601_1", "602_1", "706")
YEARS = range(2019, 2025)


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source-dir",
        type=Path,
        help="Hash an existing raw directory instead of downloading (used to reproduce the committed manifest).",
    )
    args = parser.parse_args()
    records: list[dict[str, object]] = []
    RAW.mkdir(parents=True, exist_ok=True)

    for schedule in SCHEDULES:
        for year in YEARS:
            filename = f"schedule{schedule}_{year}.xlsx"
            url = f"{BASE_URL}/{filename}"
            target = RAW / filename
            if args.source_dir:
                source = args.source_dir / filename
                if not source.exists():
                    raise FileNotFoundError(source)
                path = source
            else:
                with urllib.request.urlopen(url, timeout=60) as response, target.open("wb") as handle:
                    shutil.copyfileobj(response, handle)
                path = target
            records.append(
                {
                    "schedule": schedule,
                    "year": year,
                    "filename": filename,
                    "url": url,
                    "bytes": path.stat().st_size,
                    "sha256": digest(path),
                }
            )

    payload = {
        "sourcePage": SOURCE_PAGE,
        "retrievedUtc": datetime.now(timezone.utc).isoformat(),
        "files": records,
    }
    MANIFEST.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {MANIFEST} with {len(records)} source files")


if __name__ == "__main__":
    main()
