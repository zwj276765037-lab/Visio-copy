#!/usr/bin/env python3
"""Disabled trace helper.

visio-copy is native-Visio-only. Do not use raster tracing, SVG generation,
or auto-vectorized pixel rectangles for delivery artifacts.
"""

from __future__ import annotations

raise SystemExit(
    "Disabled: visio-copy must draw with native Visio shapes only; "
    "raster-to-SVG tracing is forbidden."
)
