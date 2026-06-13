# Visio Redraw Checklist

Use this checklist for exact diagram recreation tasks.

## Setup

- Keep the original `.vsdx` backed up before each script run.
- Use a dedicated tracing page when experimenting.
- Set the page size from the source image dimensions.
- During tracing iterations, lock the source image underlay:
  - `LockMove = 1`
  - `LockWidth = 1`
  - `LockHeight = 1`
  - `LockDelete = 1`

## Final Cleanup

- After the user accepts the final redraw, save a backup before cleanup.
- Delete the source image underlay and any `TraceBase_*` layers from the final page.
- Delete old bottom tracing layers, preview layers, pixel-trace layers, and failed experimental redraw layers.
- Keep only the final accepted editable vector redraw layer on the final page.
- Do not leave the original raster image hidden behind the vector drawing.
- Re-export the final preview after cleanup, with no underlay present.
- If a debug/trace version is useful, keep it only in a separate backup file or separate backup page, not in the final deliverable page.

## Shape Construction

- Use real Visio shapes for every logical element.
- Avoid auto-vectorized SVG as the final result when the user asks for editable structure.
- Treat auto-vector extraction only as a diagnostic aid.
- For dense repeated structures, split the source into component crops before drawing.
- Do not redraw a whole dense page from memory. Recreate and validate one crop at a time.
- Count repeated cells from the crop before coding loops. Do not infer counts from rough visual width.
- Keep the script idempotent by deleting only known generated layers before redrawing.
- Use exact page and layer names so reruns do not accumulate duplicate shapes.

## Dense Component Repair

- For each crop, save a matching crop from the exported Visio preview.
- Compare reference crop and preview crop side by side.
- Inspect all planned component crops in one pass before editing. Record a batch issue list instead of rerendering after every single issue.
- Apply multiple precise repairs in one script-edit pass, then redraw/export once. Use one-problem-one-redraw only for syntax failures, broken COM automation, or risky changes that must be isolated.
- Verify cell count, cell gaps, hatch direction, front/back occlusion, arrowheads, and text positions in that crop.
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

The next script edit should address the whole list where possible, followed by a single redraw/export and a new crop set.

## Unsatisfied Preview Recovery

When a preview is still visibly wrong, diagnose the class of failure before editing:

| Failure Class | Check | Preferred Fix |
| --- | --- | --- |
| semantic | wrong editable primitive or screenshot-like trace | replace the helper/primitive |
| typography | font, formula, line break, margin, or clipping mismatch | fix text helper and textbox geometry |
| geometry | module size, gap, count, or alignment mismatch | correct bbox, loop counts, and row/column spacing |
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

## Preview Diagnosis

- Large blank areas or missing shapes usually mean a large background shape is on top.
- Correct coordinates but wrong visibility usually means layer or draw order is wrong.
- Text inside blocks breaking into single letters means the text box is too narrow or rotated incorrectly.
- A clean preview with wrong editable structure means the method has slipped into screenshot/SVG tracing and should be rejected for Visio-copy tasks.
- A stack with too many visible internal lines usually means hidden rear cells were drawn instead of only visible edges.
- A stack that reads as one cuboid usually means separated cells were not modeled as separate front/back units.
- A final preview that still looks correct only because the original image is underneath is not acceptable; remove the underlay and inspect again.
- A preview that looks acceptable only at full-page scale is not enough; inspect all component crops at enlarged zoom.
