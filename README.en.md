<p align="right">
  <a href="README.md">简体中文</a> | <strong>English</strong>
</p>

# Visio Copy

`visio-copy` is a Codex skill for recreating raster technical diagrams in Microsoft Visio as editable vector shapes, instead of simply pasting a screenshot.

It is designed for architecture figures, hardware block diagrams, dataflow diagrams, and paper-style module diagrams. The workflow uses Visio COM automation to draw rectangles, arrows, buses, tables, grids, formula labels, and stacked structures, then validates the result through exported previews and component-level crop comparisons.

Usage note: complex linework, dense arrows, stacked structures, 3D/2.5D diagrams, and repeated small modules usually need several repair passes for better results. When using `visio-copy`, try a few rounds patiently; if you can point out specific issues such as misalignment, occlusion, arrow endpoints, layer order, text wrapping, or inaccurate colors, the redraw can usually converge faster to a better editable Visio result.

## What This Skill Does

- Draws editable Visio shapes instead of raster-only copies.
- Keeps the drawing operations in Visio; the final `.vsdx` should consist of editable shapes, lines, text, and layers.
- Uses source-image pixel coordinates as the shared coordinate system.
- Keeps the drawing process in native Visio shapes, connectors, lines, polygons, text, groups, and layers. The reference image is not part of the final Visio page content.
- Provides a PowerShell scaffold for Visio COM drawing.
- Provides Python tools for color-region extraction and side-by-side crop comparison.
- Includes specialized guidance for stacked grids, key matrices, repeated small blocks, and dense tensor-like diagrams.

## 2.0 Effect Showcase: Original vs Copy

The following examples show reference diagrams and `visio-copy` redraw results. The copy-side images are screenshots of editable Visio drawings, not pasted images or non-editable raster tracings.

**Version note:** this is the `2.0` release. Compared with `1.0`, this version improves editable redraw primitives, local layout stability, color/style analysis, and Visio-export audit workflows for complex hardware architecture diagrams. `visio-copy` is still not a one-click pixel-perfect converter; high-quality redraws still require component-level inspection and targeted manual repair, but 2.0 turns many repeated failure modes into reusable drawing rules and primitives.

| Case | Original | Copy |
| --- | --- | --- |
| AQPIM data layout | <img src="assets/showcase/aqpim-data-layout-original.png" width="420" alt="AQPIM data layout original"> | <img src="assets/showcase/aqpim-data-layout-copy.png" width="420" alt="AQPIM data layout Visio copy"> |
| PIM DIMMs routing | <img src="assets/showcase/pim-dimms-original.png" width="420" alt="PIM DIMMs routing original"> | <img src="assets/showcase/pim-dimms-copy.png" width="420" alt="PIM DIMMs routing Visio copy"> |
| Protocol network example | <img src="assets/showcase/protocol-network-original.png" width="420" alt="Protocol network original"> | <img src="assets/showcase/protocol-network-copy.png" width="420" alt="Protocol network Visio copy"> |
| Decoder flow | <img src="assets/showcase/decoder-flow-original.png" width="420" alt="Decoder flow original"> | <img src="assets/showcase/decoder-flow-copy.png" width="420" alt="Decoder flow Visio copy"> |
| DFBM credit management | <img src="assets/showcase/dfbm-credit-original.png" width="420" alt="DFBM credit original"> | <img src="assets/showcase/dfbm-credit-copy.png" width="420" alt="DFBM credit Visio copy"> |

## What Improved Since 1.0

- **Stronger editable Visio primitives:** added and expanded `visio_copy_manual_primitives.ps1` for transparent text, fitted text, rounded modules, tables, message-lane tables, hatched regions, polygons, orthogonal routes, block arrows, isometric router grids, and logic/electrical symbols.
- **More stable text layout:** added rules for text lanes, font caps, transparent text drawn last, rotated-label bboxes, and segmented colored text so short labels such as `PE`, `HBM`, `TSV`, `SRAM`, `NoC`, `P1`, and `D1` do not wrap after PNG export.
- **Less white-mask damage:** normal text should be transparent by default. Broad white masks that hide gray cells, bit bars, arrows, borders, or grids are treated as redraw failures.
- **Better local anchoring:** child modules, bit bars, local arrows, ports, table cells, and labels should be positioned in parent-local coordinates instead of drifting in page-global coordinates.
- **Broader hardware-diagram coverage:** added redraw guidance for chiplet/NoC figures, memory arrays, wafer/die/core/package multi-view diagrams, scheduling token chains, PE grids, dense tables, stacked matrices, DIMM/board/package diagrams, and DRAM/flash/storage-controller figures.
- **Improved color and structure analysis:** added `analyze_reference_style.py` to sample global and crop-local colors, edge directions, and structure cues before drawing.
- **Safer Visio preview export:** added `export_visio_png_safe.ps1` to avoid hangs from direct Visio `Page.Export()` calls, with page-background and off-page-shape cleanup guidance.
- **More systematic crop auditing:** strengthened component-level crop comparison for text, arrow endpoints, layering, repeated-unit counts, cell pitch, line weights, color fidelity, and 2.5D/3D stack ordering.
- **Clearer artifact boundary:** `redraw.vsdx` is reserved for manually reviewed final editable redraws. Automatic scaffolds or first-pass outputs should not overwrite it.

## Repository Layout

```text
.
|-- README.md
|-- README.en.md
|-- SKILL.md
|-- assets/
|   `-- showcase/
|-- agents/
|-- references/
|   |-- redraw-checklist.md
|   `-- stacked-grid-mode.md
|-- scripts/
|   |-- crop_compare.py
|   |-- export_visio_png_safe.ps1
|   |-- extract_color_components.py
|   |-- finalize_visio_copy_page.ps1
|   |-- import_visual_svg_to_visio.ps1
|   |-- raster_to_run_svg.py
|   |-- run_manual_crop_audit.ps1
|   |-- test_visio_copy_primitives_smoke.ps1
|   |-- visio_copy_manual_primitives.ps1
|   `-- visio_manual_redraw_scaffold.ps1
|-- requirements.txt
`-- LICENSE
```

## Requirements

- Windows
- Microsoft Visio desktop application
- PowerShell
- Python 3.10+
- Python packages in `requirements.txt`

Install Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Install As A Codex Skill

Clone this repository into your Codex skills directory:

```powershell
git clone https://github.com/zwj276765037-lab/Visio-copy.git "$env:USERPROFILE\.codex\skills\visio-copy"
```

Then invoke it in Codex with:

```text
$visio-copy
```

## Basic Workflow

1. Prepare the target `.vsdx` file and reference image path.
2. Create a project-specific redraw script from `scripts/visio_manual_redraw_scaffold.ps1`.
3. Set the source image width, height, and pixel-to-Visio scale mapping.
4. Draw in pixel coordinates with helpers such as `Add-RectPx`, `Add-LinePx`, `Add-TextPx`, and `Add-PolygonPx`.
5. Export a preview PNG from Visio.
6. Generate matching component crops from the reference and preview images.
7. Repair geometry, text, arrows, line weights, layer order, and missing elements in batches.
8. After acceptance, clean temporary layers and keep only the final editable vector layer.

## Crop Comparison

Generate side-by-side validation crops:

```powershell
python scripts/crop_compare.py reference.png preview.png --out crops `
  --component left_stack:145,30,300,190 `
  --component right_table:560,40,220,120
```

Each component is defined as:

```text
name:x,y,w,h
```

Coordinates are in reference-image pixels.

If a redraw output directory already contains `manifest.json` with an `audit_components` array, run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/run_manual_crop_audit.ps1 `
  -OutputDir "path\to\redraw-output" `
  -Zoom 3
```

The wrapper reads `source`, `preview`, and `audit_components`, then calls `crop_compare.py` to generate side-by-side crops, diffs, and metrics files.

## Color Component Extraction

Extract rough bounding boxes for common paper-figure color regions:

```powershell
python scripts/extract_color_components.py reference.png --min-area 100 --top 20
```

Use this only to seed coordinates. Final redraw quality still depends on manual component auditing and Visio shape-level drawing.

## Final Cleanup

After the user accepts the redraw:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/finalize_visio_copy_page.ps1 `
  -TargetPath "path\to\diagram.vsdx" `
  -PageName "VisioCopy_Trace" `
  -FinalLayerName "ManualRedraw_HiRes"
```

This backs up the file, deletes shapes outside the final layer, removes temporary tracing layers, and saves the `.vsdx`.

## Disabled Trace Guardrails

The repository keeps two disabled guardrail scripts:

- `scripts/raster_to_run_svg.py`
- `scripts/import_visual_svg_to_visio.ps1`

They are not part of the drawing workflow and intentionally fail when executed. They exist only to make the boundary explicit: raster-to-SVG conversion, SVG/PDF trace import, and auto-vectorized pixel-copy routes are forbidden as normal `visio-copy` delivery methods.

## Notes For Dense Stacked Diagrams

Do not draw separated tensor/key-matrix blocks as one cuboid. Count the visible cells from crops, draw rear hints first, then draw opaque front cells so hidden rear lines do not pass through foreground gaps. Dense stacked diagrams must be validated at 2x or 3x crop scale.

See `references/stacked-grid-mode.md` for detailed rules.

## License

MIT License. See `LICENSE`.
