#!/usr/bin/env python3
"""Analyze color and structure cues in a reference image for Visio redraws.

The output is meant to seed editable Visio primitives, not to auto-trace a final
diagram. It reports dominant colors, connected color regions, crop-local
palettes, edge direction statistics, and simple stack/3D cues.

Examples:
  python analyze_reference_style.py ref.png --top-colors 16 --json
  python analyze_reference_style.py ref.png --component stack_a:40,80,160,120
"""

from __future__ import annotations

import argparse
import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
from PIL import Image


@dataclass(frozen=True)
class Box:
    name: str
    x: int
    y: int
    w: int
    h: int


def rgb_to_hex(rgb: Iterable[int]) -> str:
    r, g, b = [int(v) for v in rgb]
    return f"#{r:02X}{g:02X}{b:02X}"


def parse_component(value: str) -> Box:
    try:
        name, raw_box = value.split(":", 1)
        x, y, w, h = [int(v.strip()) for v in raw_box.split(",")]
    except Exception as exc:  # noqa: BLE001
        raise argparse.ArgumentTypeError("component must look like name:x,y,w,h") from exc
    if not name.strip():
        raise argparse.ArgumentTypeError("component name cannot be empty")
    if w <= 0 or h <= 0:
        raise argparse.ArgumentTypeError("component width and height must be positive")
    return Box(name.strip(), x, y, w, h)


def sample_pixels(arr: np.ndarray, max_pixels: int = 180_000) -> np.ndarray:
    flat = arr.reshape(-1, 3).astype(np.float32)
    if flat.shape[0] <= max_pixels:
        return flat
    stride = max(1, flat.shape[0] // max_pixels)
    return flat[::stride][:max_pixels]


def quantized_palette(arr: np.ndarray, top: int, bin_size: int) -> list[dict[str, object]]:
    quantized = (arr.astype(np.uint16) // bin_size) * bin_size + bin_size // 2
    quantized = np.clip(quantized, 0, 255).astype(np.uint8)
    flat = quantized.reshape(-1, 3)
    colors, counts = np.unique(flat, axis=0, return_counts=True)
    order = np.argsort(counts)[::-1][:top]
    total = int(flat.shape[0])
    palette = []
    for idx in order:
        rgb = colors[idx].tolist()
        palette.append(
            {
                "rgb": [int(v) for v in rgb],
                "hex": rgb_to_hex(rgb),
                "fraction": round(float(counts[idx]) / total, 6),
                "pixels": int(counts[idx]),
            }
        )
    return palette


def kmeans_palette(arr: np.ndarray, k: int, iterations: int = 12) -> list[dict[str, object]]:
    pixels = sample_pixels(arr)
    if pixels.size == 0:
        return []
    unique = np.unique(pixels.astype(np.uint8), axis=0).astype(np.float32)
    k = max(1, min(k, unique.shape[0]))
    if unique.shape[0] <= k:
        centers = unique
    else:
        gray = unique.mean(axis=1)
        seeds = np.linspace(0, unique.shape[0] - 1, k).round().astype(int)
        centers = unique[np.argsort(gray)[seeds]].astype(np.float32)

    labels = np.zeros(pixels.shape[0], dtype=np.int32)
    for _ in range(iterations):
        distances = ((pixels[:, None, :] - centers[None, :, :]) ** 2).sum(axis=2)
        labels = distances.argmin(axis=1)
        new_centers = centers.copy()
        for idx in range(k):
            selected = pixels[labels == idx]
            if selected.size:
                new_centers[idx] = selected.mean(axis=0)
        if np.allclose(new_centers, centers, atol=0.25):
            break
        centers = new_centers

    counts = np.bincount(labels, minlength=k)
    order = np.argsort(counts)[::-1]
    total = int(counts.sum())
    result = []
    for idx in order:
        rgb = np.clip(np.rint(centers[idx]), 0, 255).astype(np.uint8).tolist()
        result.append(
            {
                "rgb": [int(v) for v in rgb],
                "hex": rgb_to_hex(rgb),
                "fraction": round(float(counts[idx]) / total, 6) if total else 0.0,
                "pixels_sampled": int(counts[idx]),
            }
        )
    return result


def edge_metrics(arr: np.ndarray) -> dict[str, object]:
    gray = (
        0.299 * arr[:, :, 0].astype(np.float32)
        + 0.587 * arr[:, :, 1].astype(np.float32)
        + 0.114 * arr[:, :, 2].astype(np.float32)
    )
    if gray.shape[0] < 3 or gray.shape[1] < 3:
        return {"edge_density": 0.0, "direction_fraction": {}}

    gx = np.zeros_like(gray)
    gy = np.zeros_like(gray)
    gx[:, 1:-1] = gray[:, 2:] - gray[:, :-2]
    gy[1:-1, :] = gray[2:, :] - gray[:-2, :]
    mag = np.hypot(gx, gy)
    threshold = max(18.0, float(np.percentile(mag, 88)))
    edge = mag >= threshold
    edge_count = int(edge.sum())
    if edge_count == 0:
        return {"edge_density": 0.0, "direction_fraction": {}}

    angle = np.degrees(np.arctan2(gy[edge], gx[edge]))
    angle = (angle + 180.0) % 180.0
    # Edge normal angle -> visual line orientation.
    line_angle = (angle + 90.0) % 180.0
    buckets = {
        "horizontal": ((line_angle <= 12.0) | (line_angle >= 168.0)).sum(),
        "vertical": ((line_angle >= 78.0) & (line_angle <= 102.0)).sum(),
        "diag_down": ((line_angle >= 28.0) & (line_angle <= 62.0)).sum(),
        "diag_up": ((line_angle >= 118.0) & (line_angle <= 152.0)).sum(),
        "other": edge_count,
    }
    buckets["other"] -= int(
        buckets["horizontal"] + buckets["vertical"] + buckets["diag_down"] + buckets["diag_up"]
    )
    direction_fraction = {
        key: round(float(value) / edge_count, 4) for key, value in buckets.items() if value > 0
    }
    return {
        "edge_density": round(float(edge_count) / float(edge.size), 6),
        "direction_fraction": direction_fraction,
        "edge_pixels": edge_count,
    }


def connected_components(mask: np.ndarray, min_area: int) -> list[dict[str, object]]:
    height, width = mask.shape
    seen = np.zeros(mask.shape, dtype=bool)
    components: list[dict[str, object]] = []

    for y in range(height):
        xs = np.where(mask[y] & ~seen[y])[0]
        for x0 in xs:
            if seen[y, x0] or not mask[y, x0]:
                continue
            stack = [(int(x0), int(y))]
            seen[y, x0] = True
            min_x = max_x = int(x0)
            min_y = max_y = int(y)
            area = 0
            while stack:
                x, yy = stack.pop()
                area += 1
                min_x = min(min_x, x)
                max_x = max(max_x, x)
                min_y = min(min_y, yy)
                max_y = max(max_y, yy)
                for nx, ny in ((x + 1, yy), (x - 1, yy), (x, yy + 1), (x, yy - 1)):
                    if 0 <= nx < width and 0 <= ny < height and not seen[ny, nx] and mask[ny, nx]:
                        seen[ny, nx] = True
                        stack.append((nx, ny))
            if area >= min_area:
                components.append(
                    {
                        "area": area,
                        "bbox": [min_x, min_y, max_x + 1, max_y + 1],
                        "xywh": [min_x, min_y, max_x + 1 - min_x, max_y + 1 - min_y],
                    }
                )
    return sorted(components, key=lambda item: item["area"], reverse=True)


def region_components(arr: np.ndarray, colors: list[dict[str, object]], tolerance: int, min_area: int, top: int) -> list[dict[str, object]]:
    result = []
    rgb_arr = arr.astype(np.int16)
    for color in colors:
        rgb = np.array(color["rgb"], dtype=np.int16)
        distance = np.abs(rgb_arr - rgb).max(axis=2)
        mask = distance <= tolerance
        comps = connected_components(mask, min_area)[:top]
        if comps:
            result.append({"hex": color["hex"], "rgb": color["rgb"], "components": comps})
    return result


def crop_array(arr: np.ndarray, box: Box) -> np.ndarray:
    y0 = max(0, box.y)
    x0 = max(0, box.x)
    y1 = min(arr.shape[0], box.y + box.h)
    x1 = min(arr.shape[1], box.x + box.w)
    return arr[y0:y1, x0:x1, :]


def classify_cues(metrics: dict[str, object]) -> list[str]:
    fractions = metrics.get("direction_fraction", {})
    if not isinstance(fractions, dict):
        return []
    cues = []
    diag = float(fractions.get("diag_down", 0.0)) + float(fractions.get("diag_up", 0.0))
    hv = float(fractions.get("horizontal", 0.0)) + float(fractions.get("vertical", 0.0))
    density = float(metrics.get("edge_density", 0.0))
    if diag >= 0.2 and hv >= 0.25:
        cues.append("likely_2_5d_or_stacked")
    if density >= 0.12 and hv >= 0.45:
        cues.append("dense_grid_or_table")
    if diag >= 0.3:
        cues.append("hatch_or_perspective_edges")
    return cues


def analyze_crop(arr: np.ndarray, box: Box, top_colors: int, bin_size: int) -> dict[str, object]:
    crop = crop_array(arr, box)
    metrics = edge_metrics(crop)
    return {
        "name": box.name,
        "xywh": [box.x, box.y, box.w, box.h],
        "actual_size": [int(crop.shape[1]), int(crop.shape[0])],
        "palette": kmeans_palette(crop, min(top_colors, 10)),
        "quantized_palette": quantized_palette(crop, min(top_colors, 10), bin_size),
        "edge_metrics": metrics,
        "cues": classify_cues(metrics),
    }


def compact_print(result: dict[str, object]) -> None:
    size = result["size"]
    print(f"image={result['image']} size={size[0]}x{size[1]}")
    print("\npalette")
    for item in result["palette"]:
        rgb = ",".join(str(v) for v in item["rgb"])
        print(f"  {item['hex']} rgb=({rgb}) fraction={item['fraction']}")

    print("\nmajor_color_regions")
    for color in result["major_color_regions"]:
        print(f"  {color['hex']}")
        for comp in color["components"]:
            x, y, w, h = comp["xywh"]
            print(f"    area={comp['area']:7d} xywh=({x},{y},{w},{h})")

    if result["components"]:
        print("\ncomponents")
    for comp in result["components"]:
        print(f"  {comp['name']} xywh={tuple(comp['xywh'])} cues={','.join(comp['cues']) or 'none'}")
        for color in comp["palette"][:5]:
            rgb = ",".join(str(v) for v in color["rgb"])
            print(f"    {color['hex']} rgb=({rgb}) fraction={color['fraction']}")
        metrics = comp["edge_metrics"]
        print(f"    edge_density={metrics['edge_density']} directions={metrics['direction_fraction']}")


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("image", type=Path)
    parser.add_argument("--component", action="append", type=parse_component, default=[])
    parser.add_argument("--top-colors", type=int, default=14)
    parser.add_argument("--bin-size", type=int, default=16)
    parser.add_argument("--region-tolerance", type=int, default=12)
    parser.add_argument("--min-area", type=int, default=80)
    parser.add_argument("--top-regions", type=int, default=12)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(list(argv) if argv is not None else None)

    img = Image.open(args.image).convert("RGB")
    arr = np.asarray(img)
    palette = kmeans_palette(arr, args.top_colors)
    result = {
        "image": str(args.image),
        "size": [img.width, img.height],
        "palette": palette,
        "quantized_palette": quantized_palette(arr, args.top_colors, args.bin_size),
        "major_color_regions": region_components(
            arr,
            palette[: min(args.top_colors, 10)],
            args.region_tolerance,
            args.min_area,
            args.top_regions,
        ),
        "global_edge_metrics": edge_metrics(arr),
        "components": [
            analyze_crop(arr, box, args.top_colors, args.bin_size) for box in args.component
        ],
    }
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        compact_print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
