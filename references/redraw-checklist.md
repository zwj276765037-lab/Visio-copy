# Visio Redraw Checklist

Use this checklist for exact diagram recreation tasks.

## Setup

- Keep the original `.vsdx` backed up before each script run.
- Use a dedicated clean redraw page when experimenting.
- Set the page size from the source image dimensions.
- Run `scripts/analyze_reference_style.py` on the full reference before drawing. Save the palette, major color regions, and global edge metrics near the redraw script or in the batch repair notes.
- Keep the source image outside Visio. Use it only for measurement, color sampling, crop audit, and side-by-side review.

## Final Cleanup

- After the user accepts the final redraw, save a backup before cleanup.
- Delete old preview/debug layers and failed experimental redraw layers.
- Keep only the final accepted editable native Visio shape layer on the final page.
- Do not leave any original raster image, imported SVG/PDF trace, pasted screenshot, or auto-vectorized pixel group in the delivered file.
- Re-export the final preview after cleanup and inspect it as a standalone native Visio drawing.

## Shape Construction

- Use real Visio shapes for every logical element.
- Do not use auto-vectorized SVG, raster tracing, or imported image crops as construction shortcuts.
- For dense repeated structures, split the source into component crops before drawing.
- For color-sensitive or 3D components, run `scripts/analyze_reference_style.py --component name:x,y,w,h` and use sampled crop-local colors rather than guessed named colors.
- Do not redraw a whole dense page from memory. Recreate and validate one crop at a time.
- For each large module, define a local origin and place internal child modules, bit bars, labels, and local arrows as offsets from that origin.
- Count repeated cells from the crop before coding loops. Do not infer counts from rough visual width.
- Choose the primitive from the crop: flat rectangle, gradient band, shadowed rectangle, isometric block, depth cap, separated stacked matrix, or component-specific polygons.
- Keep the script idempotent by deleting only known generated layers before redrawing.
- Use exact page and layer names so reruns do not accumulate duplicate shapes.

## Dense Component Repair

- For each crop, save a matching crop from the exported Visio preview.
- Compare reference crop and preview crop side by side.
- Inspect all planned component crops in one pass before editing. Record a batch issue list instead of rerendering after every single issue.
- Apply multiple precise repairs in one script-edit pass, then redraw/export once. Use one-problem-one-redraw only for syntax failures, broken COM automation, or risky changes that must be isolated.
- Verify cell count, cell gaps, hatch direction, front/back occlusion, arrowheads, and text positions in that crop.
- Verify sampled fill colors, line colors, transparency, shadows, gradients, and 3D face colors in that crop.
- Verify internal micro-layout: small child modules should align to local rows/columns, not drift independently inside a correct outer frame.
- Verify text-line collisions: no connector, arrow, grid line, or bar edge should cross a label unless the source explicitly does. Draw text last and reroute or shorten lines first; use masks only when unavoidable.
- Verify overlay bars: if a label/bit bar is drawn over gray repeated cells, it should be segmented: gray cells first, blue outline with no fill, solid `S` tab, white numeric text segment, and transparent tail over the gray cells. A one-piece white-filled bar that erases the gray cells, or a blue outline that starts at the wrong cell offset, is a primitive error.
- Classify every unexpected white background by source: text box, overlay bar, or row/container panel. Do not leave a white row/container fill unless the reference crop clearly contains that panel fill; use transparent fill with only an outline when the row should inherit the parent module background.
- Build a component manifest for all recognizable small modules, not only dense stacks. Include outer containers, headers, tables, formulas, arrows, ports, repeated blocks, labels, and legends.
- For each manifest item, check whether any source element is missing from the redraw. Do not rely on the full-page preview to catch missing small symbols.
- For each manifest item, check whether draw order caused a later-needed object to be covered by a background, region fill, large white block, or table cell.
- For each manifest item, check text layout: font size, bold/italic style, text box width/height, margins, wrap, clipping, formula placement, and overlap with lines or shapes.
- Do not accept a full-page preview if a dense crop is visibly wrong.
- If a stacked object looks like one solid cube but the reference uses separated blocks, replace the drawing primitive instead of tuning coordinates.
- If rear lines show through front-cell gaps but the reference hides them, add layer-footprint masks and redraw from back to front.

## Batch Repair List

Before changing the drawing script, write down all issues found in the current crop set:

| Component | Failure Type | Exact Fix |
| --- | --- | --- |
| qkpu_rows | layering | draw rear stack rectangles first, then front text |
| bui_lut | text | split formula into base/subscript/superscript text pieces |
| ander_tree | missing | add second AND gate and connector for subgroup |
| memory_stack | color-style | replace guessed cyan/red with sampled crop HEX colors and lighten top faces |
| slicing_unit_rows | micro-layout | use row-local anchors so bit bars, labels, and arrows move together |
| dequant_unit0 | collision | shorten connector endpoints to the label edge; use masks only if rerouting fails |

The next script edit should address the whole list where possible, followed by a single redraw/export and a new crop set.

## Unsatisfied Preview Recovery

When a preview is still visibly wrong, diagnose the class of failure before editing:

| Failure Class | Check | Preferred Fix |
| --- | --- | --- |
| semantic | wrong editable primitive or screenshot-like trace | replace the helper/primitive |
| typography | font, formula, line break, margin, or clipping mismatch | fix text helper and textbox geometry |
| geometry | module size, gap, count, or alignment mismatch | correct bbox, loop counts, and row/column spacing |
| micro-layout | outer frame is correct but inner child shapes or labels are offset | introduce parent-local anchors and row/column arrays |
| collision | linework crosses text or text is hidden by later shapes | draw text last, reroute lines, or use masked text |
| color-style | fill, stroke, gradient, shadow, transparency, or 3D face color mismatch | rebuild palette from crop analysis and update shared colors/helpers |
| layering | foreground hidden by fills, tables, or rear stack hints | change draw order, add masks, avoid layer-wide BringToFront |
| missing | absent arrows, ports, dots, labels, or repeated cells | add missing element in the component script |

If the same problem appears in multiple places, repair the shared helper first. Do not keep nudging coordinates around a flawed helper. For example, a formula helper that scatters superscripts should be replaced with a more stable compact formula label or validated in an isolated crop before global use.

If a supposedly more faithful helper makes the component crop less readable, more cluttered, or more overlapped than the previous preview, roll it back. Keep the stable version and record the failed helper in the audit notes instead of compounding the regression.

Apply repairs in this order: global layout, component primitive/counts, arrows and layer ordering, then text polish. Fine text alignment is not meaningful while module geometry or repeated element counts are still wrong.

## Text Repair

- If words split vertically or wrap unexpectedly, widen the text box first.
- If widening is impossible, reduce font size next.
- Do not rely on Visio auto-fit for publication-like diagrams.
- Use separate text shapes placed after arrows and fills.
- Keep labels centered by matching their text box to the corresponding source label bounding box.
- For dense internals, collect text shapes and bring them to front before saving.
- Keep ordinary text transparent. A large white label background covering arrows, bit bars, or frame edges is a redraw failure.
- Use a panel-colored or white masked text background only as a tight knockout behind the text glyphs when rerouting is impossible.

## Preview Diagnosis

- Large blank areas or missing shapes usually mean a large background shape is on top.
- Correct coordinates but wrong visibility usually means layer or draw order is wrong.
- Correct geometry but wrong visual match usually means color/style was guessed. Re-sample the crop palette before moving shapes.
- Text inside blocks breaking into single letters means the text box is too narrow or rotated incorrectly.
- A correct large frame with drifting small labels usually means the child module lacks a local coordinate system.
- Lines crossing labels usually mean the connector endpoint was not routed around the text bbox. Do not solve this by placing a broad white box over the linework.
- A clean preview with wrong editable structure means the method has slipped into screenshot/SVG tracing and should be rejected for Visio-copy tasks.
- A stack with too many visible internal lines usually means hidden rear cells were drawn instead of only visible edges.
- A stack that reads as one cuboid usually means separated cells were not modeled as separate front/back units.
- A shaded 3D block that looks flat usually means front/top/right faces were drawn with one fill color or in the wrong order.
- A final preview that still looks correct only because the original image is underneath is not acceptable; remove the underlay and inspect again.
- A preview that looks acceptable only at full-page scale is not enough; inspect all component crops at enlarged zoom.
