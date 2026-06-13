# Visio Copy

`visio-copy` is a Codex skill for recreating raster technical diagrams in Microsoft Visio as editable vector shapes.

It is designed for paper-style architecture figures: boxes, arrows, buses, tables, formulas, grids, stacked tensors, and dense labels. The workflow uses a locked reference image as a temporary underlay, redraws the diagram with real Visio shapes through COM automation, exports previews, and validates dense components with side-by-side crops.

## What This Skill Does

- Draws editable Visio shapes instead of pasting screenshots.
- Uses source-image pixel coordinates for repeatable layout.
- Keeps a tracing underlay during iteration, then removes it for final delivery.
- Provides PowerShell helpers for Visio COM drawing.
- Provides Python tools for color-component extraction and crop comparison.
- Includes special guidance for separated stacked-grid and key-matrix diagrams.

## Repository Layout

```text
.
|-- SKILL.md
|-- agents/
|-- references/
|   |-- redraw-checklist.md
|   `-- stacked-grid-mode.md
|-- scripts/
|   |-- crop_compare.py
|   |-- extract_color_components.py
|   |-- finalize_visio_copy_page.ps1
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

Clone this repository into your Codex skills directory, or copy it there:

```powershell
git clone https://github.com/zwj276765037-lab/Visio-copy.git "$env:USERPROFILE\.codex\skills\visio-copy"
```

Then invoke it in Codex with:

```text
$visio-copy
```

## Basic Workflow

1. Put the target `.vsdx` file and reference image in known paths.
2. Use `scripts/visio_manual_redraw_scaffold.ps1` as the base for a project-specific redraw script.
3. Set the source image width and height.
4. Draw in pixel coordinates with helper functions such as `Add-RectPx`, `Add-LinePx`, `Add-TextPx`, and `Add-PolygonPx`.
5. Export a preview PNG from Visio.
6. Compare the preview with the reference by component crops.
7. Repair geometry, text, arrows, line weights, and layer order.
8. After acceptance, run final cleanup to remove tracing underlays and keep only editable vector shapes.

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

## Color Component Extraction

Extract rough bounding boxes for common paper-figure color regions:

```powershell
python scripts/extract_color_components.py reference.png --min-area 100 --top 20
```

Use this only to seed coordinates. Final redraw quality still depends on manual component auditing.

## Visio Final Cleanup

After the user accepts the redraw:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/finalize_visio_copy_page.ps1 `
  -TargetPath "path\to\diagram.vsdx" `
  -PageName "VisioCopy_Trace" `
  -FinalLayerName "ManualRedraw_HiRes"
```

This backs up the file, deletes shapes outside the final layer, removes non-final layers, and saves the `.vsdx`.

## Notes For Dense Stacked Diagrams

Do not draw separated tensor/key-matrix blocks as one cuboid. Count the visible cells from crops, draw rear hints first, then draw opaque front cells so hidden rear lines do not pass through foreground gaps. Validate every dense block at 2x or 3x crop scale.

See `references/stacked-grid-mode.md` for detailed rules.

## License

MIT License. See `LICENSE`.
