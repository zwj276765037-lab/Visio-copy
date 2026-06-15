#!/usr/bin/env python3
"""Create crop pairs for Visio-copy component validation.

Example:
  python crop_compare.py reference.png preview.png --out crops \
    --component left_tensor:32,84,146,118 \
    --component time0_blocks:560,170,170,80

The component bbox is always in reference-image pixels. Preview bboxes are
scaled automatically when the Visio export size differs from the reference.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import json

import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageStat


def parse_component(value: str) -> tuple[str, tuple[int, int, int, int]]:
    try:
        name, raw_box = value.split(":", 1)
        x, y, w, h = [int(v.strip()) for v in raw_box.split(",")]
    except Exception as exc:  # noqa: BLE001
        raise argparse.ArgumentTypeError(
            "component must look like name:x,y,w,h"
        ) from exc
    if not name.strip():
        raise argparse.ArgumentTypeError("component name cannot be empty")
    if w <= 0 or h <= 0:
        raise argparse.ArgumentTypeError("component width and height must be positive")
    return name.strip(), (x, y, x + w, y + h)


def crop_scaled(
    image: Image.Image,
    ref_box: tuple[int, int, int, int],
    ref_size: tuple[int, int],
) -> Image.Image:
    sx = image.width / ref_size[0]
    sy = image.height / ref_size[1]
    x0, y0, x1, y1 = ref_box
    box = (
        round(x0 * sx),
        round(y0 * sy),
        round(x1 * sx),
        round(y1 * sy),
    )
    return image.crop(box)


def save_pair(
    ref: Image.Image,
    preview: Image.Image,
    name: str,
    ref_box: tuple[int, int, int, int],
    out_dir: Path,
    zoom: int,
) -> None:
    ref_crop = ref.crop(ref_box)
    preview_crop = crop_scaled(preview, ref_box, ref.size)
    preview_for_metrics = preview_crop.resize(ref_crop.size, Image.Resampling.BICUBIC)
    diff = ImageChops.difference(ref_crop.convert("RGB"), preview_for_metrics.convert("RGB"))
    stat = ImageStat.Stat(diff)
    arr = np.asarray(diff.convert("L"))
    metrics = {
        "name": name,
        "xywh": [ref_box[0], ref_box[1], ref_box[2] - ref_box[0], ref_box[3] - ref_box[1]],
        "mean_abs_rgb_delta": [round(v, 3) for v in stat.mean],
        "rms_rgb_delta": [round(v, 3) for v in stat.rms],
        "changed_pixel_fraction": round(float((arr > 24).sum()) / float(arr.size), 6) if arr.size else 0.0,
    }

    if zoom != 1:
        ref_crop = ref_crop.resize((ref_crop.width * zoom, ref_crop.height * zoom), Image.Resampling.NEAREST)
        preview_crop = preview_crop.resize((preview_crop.width * zoom, preview_crop.height * zoom), Image.Resampling.NEAREST)

    out_dir.mkdir(parents=True, exist_ok=True)
    ref_path = out_dir / f"{name}.reference.png"
    preview_path = out_dir / f"{name}.preview.png"
    side_path = out_dir / f"{name}.side-by-side.png"
    diff_path = out_dir / f"{name}.diff.png"
    metrics_path = out_dir / f"{name}.metrics.json"
    ref_crop.save(ref_path)
    preview_crop.save(preview_path)
    if zoom != 1:
        diff_out = diff.resize((diff.width * zoom, diff.height * zoom), Image.Resampling.NEAREST)
    else:
        diff_out = diff
    diff_out.save(diff_path)
    metrics_path.write_text(json.dumps(metrics, indent=2), encoding="utf-8")

    pad = 12
    label_h = 18
    width = ref_crop.width + preview_crop.width + diff_out.width + pad * 2
    height = max(ref_crop.height, preview_crop.height, diff_out.height) + label_h
    side = Image.new("RGB", (width, height), "white")
    side.paste(ref_crop.convert("RGB"), (0, label_h))
    side.paste(preview_crop.convert("RGB"), (ref_crop.width + pad, label_h))
    side.paste(diff_out.convert("RGB"), (ref_crop.width + preview_crop.width + pad * 2, label_h))
    draw = ImageDraw.Draw(side)
    draw.text((0, 2), "reference", fill=(0, 0, 0))
    draw.text((ref_crop.width + pad, 2), "preview", fill=(0, 0, 0))
    draw.text((ref_crop.width + preview_crop.width + pad * 2, 2), "diff", fill=(0, 0, 0))
    side.save(side_path)

    print(f"{name}: {ref_path} | {preview_path} | {side_path} | {diff_path} | {metrics_path}")


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("reference", type=Path)
    parser.add_argument("preview", type=Path)
    parser.add_argument("--out", type=Path, default=Path("visio-copy-crops"))
    parser.add_argument("--zoom", type=int, default=3)
    parser.add_argument(
        "--component",
        action="append",
        type=parse_component,
        required=True,
        help="component crop as name:x,y,w,h in reference-image pixels",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)

    ref = Image.open(args.reference).convert("RGB")
    preview = Image.open(args.preview).convert("RGB")
    for name, box in args.component:
        save_pair(ref, preview, name, box, args.out, args.zoom)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
