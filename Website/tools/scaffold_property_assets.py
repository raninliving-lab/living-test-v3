#!/usr/bin/env python3
"""
Test-only folder scaffolder.

Creates the standard property folder tree under Test/living-test-v3/asset.
This does not copy any images. It only makes folders so Path โฟลเดอร์ has a
real destination on disk.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
ASSET_ROOT = ROOT / "asset"


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def sanitize_segment(value: str) -> str:
    text = clean_text(value)
    text = text.replace("/", "_").replace("\\", "_")
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r'[:*?"<>|]', "", text)
    text = re.sub(r"_+", "_", text)
    return text.strip(" _")


def build_paths(asset_group: str, zone_folder: str, property_id: str) -> list[Path]:
    group_dir = ASSET_ROOT / sanitize_segment(asset_group)
    zone_dir = group_dir / sanitize_segment(zone_folder)
    property_dir = zone_dir / sanitize_segment(property_id)
    return [
        property_dir,
        property_dir / "01_Source",
        property_dir / "02_Working",
        property_dir / "03_Final" / "01_Photos_Final",
        property_dir / "03_Final" / "02_Videos_Final",
        property_dir / "03_Final" / "03_Post_ประกาศ",
        property_dir / "04_Private",
    ]


def main() -> int:
    parser = argparse.ArgumentParser(description="Create Test-only property folders.")
    parser.add_argument("--asset-group", default="PoolvillaForRent")
    parser.add_argument("--zone-folder", required=True)
    parser.add_argument("--property-id", required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    paths = build_paths(args.asset_group, args.zone_folder, args.property_id)
    print("Planned folders:")
    for path in paths:
        print(f" - {path.relative_to(ROOT)}")

    if not args.apply:
        print("Dry-run only. Re-run with --apply to create folders.")
        return 0

    for path in paths:
        path.mkdir(parents=True, exist_ok=True)

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
