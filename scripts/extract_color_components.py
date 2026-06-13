from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image


DEFAULT_CLASSES = {
    "blue": ((210, 220, 235), (235, 245, 255)),
    "gray": ((205, 205, 205), (225, 225, 225)),
    "pink": ((245, 190, 190), (255, 225, 225)),
    "green": ((215, 230, 205), (240, 250, 235)),
    "whiteish": ((245, 245, 245), (255, 255, 255)),
    "red": ((200, 0, 0), (255, 150, 150)),
    "black": ((0, 0, 0), (80, 80, 80)),
    "navy": ((0, 20, 40), (70, 110, 160)),
}


def connected_components(mask: np.ndarray, min_area: int):
    height, width = mask.shape
    seen = np.zeros(mask.shape, dtype=bool)
    components = []

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


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract major color component bboxes from a reference diagram.")
    parser.add_argument("image", type=Path)
    parser.add_argument("--min-area", type=int, default=100)
    parser.add_argument("--top", type=int, default=20)
    parser.add_argument("--json", action="store_true", help="Print JSON instead of compact text.")
    args = parser.parse_args()

    img = Image.open(args.image).convert("RGB")
    arr = np.asarray(img)
    result = {"image": str(args.image), "size": list(img.size), "classes": {}}

    for name, (lo, hi) in DEFAULT_CLASSES.items():
        lo_arr = np.array(lo, dtype=np.uint8)
        hi_arr = np.array(hi, dtype=np.uint8)
        mask = ((arr >= lo_arr) & (arr <= hi_arr)).all(axis=2)
        result["classes"][name] = connected_components(mask, args.min_area)[: args.top]

    if args.json:
        print(json.dumps(result, indent=2))
        return

    print(f"image={result['image']} size={result['size'][0]}x{result['size'][1]}")
    for name, components in result["classes"].items():
        print(f"\n{name}")
        for comp in components:
            x, y, w, h = comp["xywh"]
            print(f"area={comp['area']:7d} xywh=({x},{y},{w},{h}) bbox={tuple(comp['bbox'])}")


if __name__ == "__main__":
    main()
