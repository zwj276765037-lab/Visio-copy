---
name: visio-copy
description: Recreate a reference raster diagram/image in Microsoft Visio as editable vector shapes using a locked tracing underlay, pixel-coordinate mapping, Visio COM automation, layer ordering, text sizing, and preview exports. Use when the user says visio_copy, visio-copy, asks to draw a PNG/JPG/PDF into a .vsdx, asks to trace or copy a diagram into Visio, or asks for exact layout reproduction with editable boxes, arrows, text, grids, and line weights.
---

# Visio Copy

## Purpose

Use this skill to redraw a reference image into a real Visio diagram. The final result should be editable Visio shapes, not a pasted screenshot, not a pixel-traced SVG, and not a print-style overlay.

This skill is optimized for technical diagrams with boxes, arrows, buses, labels, grids, and clock-domain regions.

## Core Workflow

1. Open the `.vsdx` with Visio COM and keep Visio visible when the user wants to watch.
2. Create a tracing page or use the target page. Set page size from the source image dimensions with one stable mapping, usually `1 px = 0.01 in`.
3. Place the reference image as a locked bottom layer, named like `TraceBase_HiRes`.
4. Split the reference image into small component crops before detailed drawing. Use crops for every recognizable module, not only dense areas: stage headers, outer containers, controller blocks, tables, legends, arrows, text groups, grids, ports, stacks, formulas, and callouts. For each crop, record the bbox, visible cell counts, hatch directions, text boxes, line weights, and layer/occlusion behavior.
5. Build a separate redraw layer, named like `ManualRedraw_HiRes`.
6. Draw real Visio shapes in source-image pixel coordinates:
   - background and large regions
   - major module blocks
   - inner blocks, grids, and ports
   - hollow bus arrows and AXI markers
   - thin orthogonal control lines
   - thick filled data arrows
   - text labels last
7. Export a page preview PNG after every meaningful iteration and inspect it visually.
8. Crop the same component bbox from the exported preview and compare it with the reference crop side by side. Inspect all planned component crops before editing again, so one pass finds multiple issues instead of discovering one issue per redraw.
9. Convert the inspection into a batch repair list. For each issue, record the exact component, bbox, failure type, target coordinate/text/layer change, and script location/helper to edit.
10. Apply all precise repairs in one script-edit pass, then redraw/export once. Do not repeatedly rerender after every single small finding unless a script error or very high-risk layout change needs isolation.
11. Fix by coordinates, layer order, text box sizes, font sizes, and component-specific drawing primitives. Do not hide problems by pasting the source image on top.
12. After the final version is accepted, clean the Visio page for delivery: delete the locked reference image underlay, remove old tracing/preview/experimental layers, and keep only the final editable vector redraw layer on top.

## Component Detail Audit

Before accepting any redraw iteration, run a component-by-component audit. The purpose is to catch failures that a full-page preview hides.

Create a small audit manifest for the image:

```text
component: bui_lut
bbox: x,y,w,h
expected:
  - outer box, gray header, two table columns, five rows
  - all formulas present and centered
  - line weights match neighboring modules
  - no text wrap, no covered text, no unexpected overlap
drawn:
  - script helper/function name or shape group name
status: pass/fix
notes: what is missing or covered
```

For every component in the manifest:

- Compare the reference crop and preview crop side by side at 2x or 3x zoom.
- Check element count: blocks, cells, table rows, table columns, arrows, arrowheads, ports, dots, labels, and repeated stack units.
- Check geometry: x/y position, width/height, gaps, line weight, rounding, fill color, and whether lines that should be horizontal or vertical are exactly horizontal or vertical.
- Check draw order: large white backgrounds and color regions must not cover later text, borders, arrows, or small symbols; if they do, fix the drawing order instead of moving shapes by eye.
- Check text layout: font family, size, bold/italic style, line breaks, alignment, margins, formula-style labels, and whether any word or formula is split, wrapped, clipped, or overlapped.
- Check missing elements explicitly: if the crop contains a visual item that is absent from the redraw, add it or mark it as intentionally skipped with a reason.
- Re-export and recrop after each repair. A component is not done until its crop passes locally.

If a full-page preview looks acceptable but a component crop fails, continue repairing the component. Do not finalize from the full-page view alone.

## Batch Audit Repair Cycle

Use this loop for every substantial redraw:

1. Export one preview from the current Visio page.
2. Generate all component side-by-side crops for the manifest in one command.
3. Inspect every crop and write a batch issue list before making edits.
4. Group issues by repair type:
   - `missing`: add absent arrows, ports, dots, rows, cells, labels, or repeated modules.
   - `geometry`: adjust exact x/y/w/h, gaps, rounding, line weight, or horizontal/vertical alignment.
   - `layering`: change draw order so large fills and rear stack outlines cannot cover foreground text or symbols.
   - `text`: adjust font family, size, bold/italic, box size, margins, formula pieces, and wrap/clipping.
   - `primitive`: replace the wrong helper, such as a cuboid for separated blocks or a rectangle where the source uses a mux/logic gate.
5. Make all low-risk precise script edits together.
6. Rerun the Visio drawing script once and export a fresh preview.
7. Repeat only after the new crop set has been generated and inspected.

Do not use a one-problem-one-redraw loop for normal refinement. It wastes time and tends to miss interactions between module geometry, layering, and text placement.

## Dissatisfaction Recovery Loop

When the user says the redraw is still unsatisfactory, stop doing isolated coordinate nudges. First classify the visible failures across the current preview and component crops:

- `semantic`: the wrong primitive was used, such as a cuboid instead of separated stacked blocks, a connector arrow instead of a hollow bus arrow, or raster/vector trace instead of real Visio shapes.
- `typography`: text boxes, font sizes, bold/italic style, formula helpers, line breaks, margins, or rotated labels are wrong.
- `geometry`: module size, placement, row/column counts, gaps, border weights, and horizontal/vertical alignment do not match.
- `layering`: backgrounds, large fills, table cells, or rear stack hints cover foreground outlines, arrows, or text.
- `missing`: ports, arrows, dots, repeated cells, labels, separator marks, or formula pieces are absent.

Then build one batch repair list covering several components before editing. If several failures share the same cause, repair the reusable helper or primitive first rather than hand-tuning every instance. Examples:

- If formula labels become tiny, clipped, or scattered, fix or replace the formula helper and validate it on one crop before applying it globally.
- If every stacked-grid object reads as one solid cube, replace the stack primitive with separated per-cell blocks and occlusion masks.
- If repeated modules have the wrong cell count, update the component manifest and loop counts, not only the outer bounding box.
- If text is repeatedly covered, repair draw order and text-layer timing instead of moving labels away from the correct coordinates.
- If a more faithful helper makes the crop less readable or more cluttered than the previous iteration, roll it back immediately and keep the stable helper until a replacement is validated in isolation.

Use a staged quality gate:

1. Global layout and major module positions.
2. Component geometry and repeated element counts.
3. Arrow topology, line weights, and layer ordering.
4. Text placement and formula readability.
5. Final cleanup with the underlay removed.

Do not overfit fine text before the component primitive and repeated counts are correct. Do not keep a broken helper just because some instances look close.

## Coordinate Rules

Use source pixels as the source of truth:

```powershell
$Scale = 0.01
$PageWidth = $SourceWidthPx * $Scale
$PageHeight = $SourceHeightPx * $Scale
function To-X { param([double]$Px) return $Px * $Scale }
function To-Y { param([double]$Py) return $PageHeight - ($Py * $Scale) }
```

Record each shape as `(x, y, w, h)` where `(0,0)` is the top-left pixel of the reference image. Convert only inside helper functions.

## Layering Rules

Draw order matters. A common failure is drawing correct shapes but then bringing an entire layer to the front, causing large background blocks to cover text and small elements.

Use this order:

1. page border and large fills
2. region fills
3. module rectangles
4. inner grids and small symbols
5. thin lines and connectors
6. thick filled arrows
7. text labels

Never call `BringToFront()` on every shape in a redraw layer. If a single object needs adjustment, bring only that object forward.

## Final Delivery Cleanup

During iteration, keep the source image underlay and older diagnostic layers only as temporary tracing aids. The final `.vsdx` page should not depend on them.

Before delivery:

- Save a backup first.
- Delete bottom reference-image layers such as `TraceBase_*`, `TraceBase_HiRes`, or `locked_*_trace_base`.
- Delete old trial redraw layers and pixel-trace/preview layers that are not the final accepted version.
- Keep the final vector redraw layer only, with editable Visio shapes.
- Bring the final redraw layer to the visible top only after removing lower underlays; do not use this to hide leftover source images.
- Re-export a final preview after cleanup and inspect it without the underlay present.
- Report that the source underlay has been removed. If the user wants a trace/debug copy, keep it as a separate backup page or backup file, not on the final page.

## Text Rules

Text mismatch is usually caused by text box geometry, not only font size.

- Use separate text-only shapes above the rectangles.
- Set line and fill patterns to `0` for text-only shapes.
- Use `Arial` unless the source clearly uses another font.
- Set all text margins to `0 pt`.
- Reduce font size before letting Visio wrap words.
- Widen text boxes to match the source label extents.
- For compact module labels, prefer smaller font and exact box width over manual line breaks.
- Draw text last so arrows and fills cannot cover labels.

## Arrows And Lines

- Hollow PCIe/AXI arrows should be polygons with white fill, not ordinary connector arrows.
- Thick black and red data paths should be filled polygon arrows or thick orthogonal shafts plus arrowheads.
- Control/config lines should be thin and orthogonal. Avoid diagonal segments unless the source has them.
- Dotted separators are best recreated as small filled rectangles or squares so their spacing matches the reference.
- For vector-array grids, draw the outer white grid rectangle, then internal vertical/horizontal lines, then small arrows and labels.

## Stacked-Grid Mode

Use this mode for 2.5D stacked tensors, memory planes, cube arrays, or repeated grid piles. Do not use a naive loop that draws every layer and every cell: it creates hidden-line clutter and fails to match paper diagrams.

Before drawing a stack, crop it from the reference and count the visible cells. Paper diagrams often use separated blocks, not a continuous cuboid grid. If the source shows separated blocks, every visible cell must have its own rectangle, gap, hatch mark, and front/back ordering.

Represent each stacked object with a component manifest:

```text
bbox = x,y,w,h
front = rows, cols, cell_w, cell_h
depth = layers, dx, dy
cells = matrix of style tokens
visible_edges = front, top_outline, right_outline, rear_hints
labels = text bboxes anchored to the component, not guessed globally
```

Cell style tokens should include at least:

- `cyan`: loaded/front data
- `white`: unloaded data
- `blue_slash`: redundancy slash in a white cell
- `red_slash`: reusable/redundant red hatch in a white cell
- `cross`: iEQK cross cell
- `red`: iQK cell

Draw stacked grids in this order:

1. rear/depth hints as only the visible top/right edges or separated rear cells
2. an occlusion mask for the next/front layer when rear lines would show through gaps
3. front-face cells from the explicit matrix
4. per-cell slashes/crosses bound to individual cells
5. dimension arrows and labels anchored to the component bbox
6. annotations and red callouts

When the preview looks chaotic, inspect whether hidden rear grid lines are being drawn. Most paper figures omit them.

For separated stacked blocks:

- Do not draw a single large cuboid and then subdivide it.
- Do not let rear grid lines pass through the gaps between front cells.
- Draw from back to front, and use white masks between layers when the source hides lower layers.
- Count visible blue/white/red/cross cells from the crop before coding the matrix.
- Verify slash direction per cell: red and blue hatches can use opposite diagonals.
- If the output still reads as one solid cube, the primitive is wrong; replace it with a component-specific block helper instead of tuning coordinates.

For cache/key-matrix stacks in paper figures, inspect whether the rear layers are actually orthogonal stepped caps rather than diagonal perspective faces. If the crop shows U-shaped top steps and mostly vertical rear edges:

- draw the front cell column first as the counted matrix, usually four visible cells for bit planes if the crop shows four divisions.
- draw rear layers as stepped U-shaped top caps plus only the visible right-side vertical hints.
- treat every visible top cap/block as an independent shape with a real white gap. The cap width must be smaller than the column pitch; never let adjacent U-caps share an edge or join into a continuous bridge.
- if a repeated stack appears connected in the preview, first check `cell_w + gap` versus top-cap width, then shorten the cap or add explicit white gutters/masks before changing global coordinates.
- for cache/key-matrix stacks specifically, draw rear/top cap hints before the front cell columns, then draw the front cells with opaque white/colored fills so they occlude lower cap drops. If cap ticks appear inside the front faces or through column gaps, the draw order is wrong.
- do not draw one rear vertical guide through every column gap unless the source visibly shows it. Prefer sparse right-edge depth silhouettes; otherwise separated blocks collapse into one cuboid-like grid.
- run a 2x/3x crop audit that verifies every front column has the source cell count, the top caps have visible white gutters, and no rear hidden line passes through a front block gap.
- avoid diagonal connector lines and rear bottom edges unless the source explicitly shows them.
- keep rear-line weight lighter than front cell borders so the stack does not become a cuboid.

## Use The Bundled Scripts

- `scripts/extract_color_components.py`: run on the reference image to get bounding boxes for major color regions. Use this to seed coordinates instead of guessing.
- `scripts/crop_compare.py`: after exporting a Visio preview, crop matching component bboxes from the reference and preview, then save side-by-side comparisons for detailed local repair.
- `scripts/visio_manual_redraw_scaffold.ps1`: copy or adapt for a target `.vsdx`. It provides Visio COM setup, locked underlay placement, coordinate conversion, layer cleanup, preview export, and drawing helpers.
- `scripts/finalize_visio_copy_page.ps1`: after the final redraw is accepted, back up the `.vsdx`, delete all shapes not on the final layer, remove old underlay/tracing layers, and save a clean final page.

Read `references/redraw-checklist.md` before finishing a redraw task. For 2.5D stack diagrams, also read `references/stacked-grid-mode.md`.

## Validation Checklist

Before finalizing:

- Export a preview PNG from the Visio page.
- Inspect the preview, not just the live Visio canvas.
- During iteration, verify the tracing image is on the bottom layer and locked.
- Verify the manual redraw layer is separate and editable.
- For final delivery, verify the reference image underlay has been removed and only the final vector redraw remains.
- Check that no experimental pixel-trace page or layer remains unless the user requested it as a separate backup/debug copy.
- Confirm text does not wrap unexpectedly or get covered.
- Confirm main arrows are horizontal/vertical where the source is horizontal/vertical.
- For stacked grids, confirm hidden rear grid lines are not drawn unless visible in the source.
- For stacked grids, confirm each slash/cross belongs to a specific cell, not a free-floating diagonal line.
- For every dense component crop, confirm the crop-level preview matches the reference before judging the full page.
- Report the page name, layer names, preview path, and backup path to the user.
