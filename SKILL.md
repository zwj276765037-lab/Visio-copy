---
name: visio-copy
description: Recreate a reference raster diagram/image in Microsoft Visio as editable native Visio shapes using pixel-coordinate mapping, sampled color palettes, Visio COM automation, layer ordering, text sizing, complex shape primitives, 2.5D/3D stacked diagram reconstruction, crop comparisons, and preview exports. Use when the user says visio_copy, visio-copy, asks to draw a PNG/JPG/PDF into a .vsdx, asks to redraw a diagram in Visio, asks to improve color fidelity, asks to reproduce complex or stacked shapes, or asks for exact layout reproduction with editable boxes, arrows, text, grids, shadows, gradients, and line weights. This skill never uses raster tracing, SVG/PDF trace import, pasted screenshots, or auto-vectorized copies as the final drawing method.
---

# Visio Copy

## Purpose

Use this skill to redraw a reference image into a real Visio diagram. The result must be editable native Visio shapes, not a pasted screenshot, not a source-image underlay, not an SVG/PDF import, not an auto-vectorized copy, and not any other tracing artifact.

This skill is optimized for technical diagrams with boxes, arrows, buses, labels, grids, clock-domain regions, paper-style color fills, shadows, gradients, and 2.5D/3D stacked blocks.

The workflow is Visio-first. Use analysis scripts only to inspect color, geometry, and crop differences; do not use them to create a rasterized final drawing. The reference image may be inspected outside the final Visio page for measurement and auditing, but it must not be inserted as a tracing underlay or remain as visible content.

## Artifact Safety Contract

- `redraw.vsdx` is reserved for a manual or manually reviewed Visio redraw that is acceptable as the current final artifact.
- Automated detection, batch reconstruction, or first-pass shape scaffolds must use names like `scaffold.vsdx`, `auto_first_pass.vsdx`, or `preview.png`; they must not overwrite `redraw.vsdx`.
- `redraw.vsdx` must be made of editable Visio shapes, connectors, text, and layers. Do not deliver a pasted source image, screenshot overlay, pixel-traced raster, or SVG/image conversion as the final redraw.
- Manual redraw scripts should open Visio visibly by default and keep the document open after saving so the drawing process is inspectable. Use `[bool]$Visible = $true` and `[bool]$KeepOpen = $true` in new manual `draw.ps1` scripts; pass `-Visible:$false -KeepOpen:$false` only for deliberate background/batch runs.
- Manual redraw workflows must update `preview.png` in the delivery directory after every accepted draw run. Prefer `scripts/export_visio_png_safe.ps1` after a hidden/closed run, or deliberately sync the latest preview from the active iteration directory. Do not rely on inline `Page.Export()` in visible keep-open scripts unless it has been tested for that figure; it can hang and leave the user thinking drawing failed.
- High-fidelity redraws are not the same as semantic architecture redraws. If the user asks for maximum restoration, treat pixel alignment as a deliverable: use 1:1 export scale, shared grid lines, per-component bboxes, and crop thresholds before claiming the figure is done. The source image may be inspected externally for measurement and audit, but must not be placed on the Visio page as a tracing underlay.

## Fidelity Modes

This skill has one mandatory delivery mode: `strict_visio_drawn`.

`strict_visio_drawn` means every visible element in the delivered `.vsdx` is created by Visio as an editable shape: rectangles, lines, connectors, polygons, tables, logic/electrical symbols, text, arrows, hatch lines, gradients built from Visio shapes, and grouped native shapes. The drawing process should be visible in Visio by default.

Forbidden in all normal use:

- placing the source PNG/JPG/PDF on the Visio page as a tracing underlay;
- importing SVG/PDF/vector traces;
- raster-to-vector conversion;
- pasted screenshots or image crops;
- auto-vectorized pixel rectangles;
- hiding poor redraw quality by leaving a source image underneath;
- switching to a trace workflow because the native Visio redraw is difficult.

The source image may only be used outside the final Visio page as a reference for measurement, color sampling, crop auditing, and visual comparison. If a task cannot be completed at the requested fidelity using native Visio shapes, say that explicitly and continue improving component descriptors and primitives rather than using tracing.

`semantic_editable` and `high_fidelity` are quality targets inside `strict_visio_drawn`, not alternate methods:

- `semantic_editable`: native Visio shapes preserve meaning and approximate layout.
- `high_fidelity`: native Visio shapes preserve source geometry as closely as possible. Every major component must be driven by source-image bboxes and crop audits, not by estimated module placement.

In `high_fidelity` mode:

- Initialize the page with `1 px = 1/96 in` unless there is a verified reason not to. This makes Visio's usual 96-DPI PNG export match the source raster size and avoids half-pixel resampling drift.
- Do not place the source image on the Visio page. Inspect it externally for measurements and crop audits, then draw native Visio shapes on a clean page.
- Draw repeated grids with shared internal lines. Do not draw each cell as a bordered rectangle when adjacent cells share edges; repeated borders create double-thick and uneven lines after export.
- For dense modules, first capture a descriptor: parent bbox, child bboxes, line weights, text lanes, ports, arrow endpoints, and draw order. Do not patch dense diagrams by page-global nudging.
- Treat `changed_pixel_fraction > 0.15` for normal diagram components as a required review, and `> 0.25` as a failed high-fidelity component unless the crop is dominated by unavoidable text/font differences and that limitation is documented.

There is no visual trace fallback in this skill. Scripts such as raster-to-SVG conversion or SVG import are out of scope for `visio-copy` delivery artifacts.

## Practical Redraw Lessons

Apply these practical lessons before writing a manual redraw script:

- Reuse `scripts/visio_copy_manual_primitives.ps1` for manual redraw scripts. Do not copy a fresh set of pixel-coordinate helpers into every figure unless the primitive file is missing a needed operation.
- Export Visio previews with `scripts/export_visio_png_safe.ps1`. Direct inline `Page.Export()` can hang; use the timeout wrapper during large redraw batches.
- For manual redraw outputs, add an `audit_components` array to `manifest.json` and run `scripts/run_manual_crop_audit.ps1 -OutputDir <figure_dir>` after `preview.png` exists. This generates the per-component side-by-side, diff, and metrics files expected by the crop audit workflow.
- Default manual Visio redraw scripts to visible, inspectable runs. Do not require the user to pass `-Visible` just to watch the drawing process; use explicit false boolean arguments when a hidden run is needed.
- In visible `KeepOpen=true` runs, the saved `.vsdx` remains locked by Visio. For another visible iteration, either close that document deliberately or write to a fresh versioned output directory such as `_v5`; do not kill the user's visible Visio window just to overwrite the same file.
- Keep `preview.png` fresh after every manual draw. A stale preview is a workflow failure even if `redraw.vsdx` saved correctly; export with the safe wrapper or sync the newest preview before running crop audits or reporting completion.
- If a preview has a different pixel size than the reference image, fix the scale/export path before trusting crop metrics. A near-constant 0.96 scale means the script probably used `0.01 in/px` with Visio's 96-DPI export; switch to `1/96 in/px` or explicitly rescale before auditing.
- For paper figures with stacked matrices or cache planes, avoid per-cell full borders. Fill cells first, draw one outer frame, then draw internal shared grid lines and rear-plane offsets. This prevents visually jagged line weights.
- For flat 2D plane stacks, draw by depth from rear to front, not by column or cell group. Rear planes should use lighter line weights; front planes should have opaque fills and the strongest shared grid lines. This avoids cuboid-like hidden-line clutter and prevents later columns from incorrectly covering earlier depth layers.
- Use `Add-VisioCopyTextJobPx -Fit` for short labels, rotated labels, and dense table/header text. Many failures were caused by Visio wrapping `Variables`, `Clause`, `Core`, `Network`, or long labels into broken lines.
- Use `Add-VisioCopyTextJobPx -Italic` or `Add-VisioCopyBoxLabelPx -Italic` when source labels are italicized, such as `Offline`/`Online` flow labels. Do not accept upright text as the final style when italic is visually meaningful.
- Use `Add-VisioCopyTextJobPx -Underline` or `Add-VisioCopyBoxLabelPx -Underline` for emphasized source labels, especially red italic underlined callouts in paper figures. Do not draw a separate underline line unless the underline must extend beyond the text run.
- In PowerShell redraw scripts, avoid hard-coded absolute paths that include non-ASCII user names. Build skill/workspace paths from `$env:USERPROFILE`, `$PSScriptRoot`, or caller parameters so script decoding cannot corrupt paths before Visio starts.
- For tiny row/column labels such as `DST`, `SRC`, `PTR`, `IND`, and `Vb`, reserve a wider local text box and reduce the font before accepting any wrap. `-Fit` lowers size but does not by itself guarantee that Visio will keep short labels on one line.
- For short hardware block labels such as `PE`, `Reg`, `ALU`, or `MAC`, still reserve a full internal text lane. A narrowly centered anchor can wrap even two- or three-character labels in exported PNGs.
- For circular operator nodes, define a node-level font cap and expected line breaks before drawing. Labels such as `BConv`, `row-NTT`, and `col-iNTT` can wrap or split inside circles unless the cap is lower than nearby rectangular labels.
- For constrained workflow headers, apply a local font-size cap before widening the label box across the next stage. Short words such as `Encoder` must stay one line without colliding with neighboring step numbers or section titles.
- For code, pseudocode, ISA snippets, and monospace timing labels, pass `-FontFace "Courier New"` or another validated monospaced font to `Add-VisioCopyTextJobPx`. Monospaced fonts need a wider fit heuristic than Arial; otherwise long lines wrap and collide even when `-Fit` is enabled.
- Use `Add-VisioCopyBoxLabelPx -RoundPx` for rounded module boxes in neural-network, pipeline, and hardware block diagrams; do not reimplement a one-off rounded label helper for each figure.
- Use `Add-VisioCopyPageBackgroundPx` at the start of manual scripts when the figure has large white margins. This provides a full-page sentinel that reduces Visio PNG export cropping.
- Before saving/exporting large diagrams, call `Remove-VisioCopyOffPageEmptyShapes` if helper-generated lines or stack outlines may cross the page boundary. Empty off-page shapes can make Visio export a larger canvas, which breaks crop scaling and makes local audits look falsely misaligned.
- Use `Add-VisioCopyOrthogonalRoutePx` for protocol/control/data paths. Treat route color, dash style, arrow direction, and endpoint as explicit topology, not ad hoc line fragments.
- For protocol/timeline message diagrams, treat labels as line-attached components: record line endpoints, label anchor, label angle, and a tight white background only under the label. Do not use broad label masks that hide neighboring protocol arrows, dashed regions, or state boxes.
- For timeline/bar diagrams, define row descriptors before drawing: row label, y-position, segment colors, segment widths, idle/wait gap widths, hatch style, and phase-label anchors. This keeps repeated bars aligned and makes compressed/variable-width rows easy to repair.
- Pipeline timing tables need pill/cell descriptors with per-row labels, column widths, spans, fills, and local font caps. A uniform pill loop is only a first pass when rows have different cell spans or dense text.
- For three-node collective or ring/triangle communication diagrams, define node anchors, arrow edges, center label, bottom value strip, and edge value strip orientation before drawing. If edge value strips are approximated as horizontal cells, record that limitation and consider promoting a rotated grouped-cell primitive after repeated need.
- Use `Add-VisioCopyTablePx` for dense protocol tables such as Cross-VN buffers, credits, lane tables, and memory/controller grids.
- Use `Add-VisioCopyMessageLaneTablePx` for Network/Cache VN request-response inset diagrams.
- Use `Add-VisioCopyHatchedRectPx` for DFBM/floorplan/hatched regions instead of drawing broad white masks or unmanaged hatch lines.
- For hatched memory or GPU-full regions, use `Add-VisioCopyHatchedRectPx` with an explicit hatch color, spacing, and angle. Do not replace hatch semantics with a solid fill or an image crop.
- Use `Add-VisioCopyPolygonPx` for filled trapezoids, custom bus arrow bodies, MUX/DEMUX blocks, and irregular module silhouettes. Do not approximate these with a filled rectangle plus outline when the slanted edge is visually important.
- For flowchart decision diamonds, use explicit four-point polygons when the exact bbox and connector anchors matter. Rotating a rectangle can expand the exported shape and collide with downstream modules or labels.
- Use `Add-VisioCopyBlockDownArrowPx` for hollow section-transition arrows.
- Use `Add-VisioCopyIsoRouterGridPx` for chiplet/router/interposer 2.5D grids. The primitive should return node anchors so vertical/interposer links attach to generated node centers.
- For bit-plane, matrix-plane, cache-plane, or token-plane stacks, first classify whether the source is a set of flat 2D planes or true 3D voxel/cuboid cells. If the source shows staggered rectangular planes with no visible top/side faces per cell, draw complete 2D grids for each plane and order them back-to-front. Do not add cuboid top/side polygons just because planes are offset.
- Use the logic/electrical symbol primitives for circuit-like architecture diagrams: `Add-VisioCopyTriangleSymbolPx`, `Add-VisioCopyMuxSymbolPx`, `Add-VisioCopyAndGateSymbolPx`, and `Add-VisioCopyOrGateSymbolPx`. Draw square/circle legend markers as shapes, not Unicode glyph text, to avoid encoding/export corruption.
- For transistor-level circuit figures, build local glyph helpers for repeated FETs, TG blocks, 6T cells, dot nodes, and BL/BLB/WL labels. Keep exact transistor geometry as a refinement target, but do not fall back to raster tracing when a simplified editable glyph preserves the circuit topology.
- NAND/flash circuit-heavy figures need local helpers for chip hierarchy, NAND strings, bitwise BL columns, latch blocks, and arithmetic grids. For throughput-oriented redraws, line/rect circuit approximations are acceptable when transistor-level geometry is recorded as an explicit limitation.
- Use ASCII formula fallbacks in PowerShell/Visio redraw scripts unless glyph export has already been validated in the target environment. Superscripts, multiplication signs, and math symbols can export as mojibake; prefer readable text such as `K^T(d x L)`, `V(L x d)`, `<=`, and `->`, or draw critical symbols as shapes.
- In narrow matrix/vector cells, avoid raw underscore labels such as `Z_11`, `x_k`, or `Row_0` unless the crop validates cleanly. Visio can wrap at the underscore; use a dedicated subscript primitive or a readable ASCII fallback such as `Z11`, `xk`, and `Row0`.
- Parenthesize coordinate arithmetic inside PowerShell point arrays. Write `@(($X + 12),($Y + 23))`, not `@($X+12,$Y+23)`, because compact expressions can be parsed as array/object addition in reusable icon helpers.
- For thick square-ended buses, hardware icons, and microarchitecture glyphs, draw filled rectangle bars instead of thick Visio lines. Visio line caps can export with rounded ends and drift from the source.
- Build chip/pin icons from editable pin rectangles plus a center square; do not rely on dashed outlines when the source has explicit pin dots or pin bars.
- For bit fields, bit bars, and highlighted binary/amplitude rows, preserve the local token semantics before drawing cells. Split each row into prefix text, highlighted cell groups, and suffix text when the source does; do not convert the whole visible bit string into one evenly colored table.
- For KV-cache, query-token, and memory-residency rows, use an ordered cell descriptor with explicit labels, fills, borders, route anchors, and local font caps. Short cells such as `NewKV` can still wrap unless their text lane is wider than the visible cell label.
- For repeated LLM iteration diagrams, create a local helper that owns the iteration title, KV cache cell, mini grid, token label, and vertical arrows. Repair the helper once rather than fixing each iteration separately.
- For bottom query/token pipelines, use a smaller local token font cap than for the same labels inside large memory panels. Labels such as `LoRA-1` and `New Token` can wrap even when their larger-panel counterparts fit cleanly.
- For segmented labels with mixed colors or styles, such as `polarity:` plus blue bit strings, create separate text jobs with explicit reserved widths for each segment. Do not rely on adjacent fitted text boxes without spacing; font fitting can shift the black label into the colored token area.
- For formula lines with colored fragments, split the formula into explicit text runs. Do not draw the full black formula and then overlay a colored fragment; the export will show duplicated or shadowed text.
- For mixed-color phrases, split each colored phrase segment into its own text anchor. Do not draw a full red phrase and overlay a black suffix such as `in GPU`; the two text layers will collide after export.
- For large italic labels that contain spaces, such as `FPGA DRAM` or code/runtime labels, split the phrase into adjacent transparent text anchors when a single fitted box wraps or breaks words. Prefer stable one-line readability over matching the exact source font size.
- For dense hardware detail panels with child buffers, FIFO lanes, handshake signals, and many 8-bit cells, accept a semantic first pass only if it is explicitly marked as such. Before refinement, build a local descriptor for each buffer/FIFO row: title lane, cell count, cell labels, per-cell fill, arrows, and badge anchors.
- When adding pale subgroup backgrounds inside dense PE-lane or AND-tree panels, draw those fills before divider lines and logic symbols, or redraw the separators immediately after the fill. Otherwise the background repair can silently erase the structural lines it was meant to sit behind.
- Small ports, stacked sockets, score taps, and other micro-symbols must have their own reserved lane. Do not place them under an existing text anchor; even a 5 px native shape can collide with a label after Visio export.
- For CIM arrays, segmented bit bars, and repeated weight grids, use a component-local descriptor that binds row origin, cell pitch, group spans, highlighted cells, local labels, DAC/output lanes, and bar anchors. Uniform row loops are only a first pass; they often make the outer panel correct while gray cells, blue bars, and labels drift apart.
- For mixed photo/rendered-scene plus schematic figures, explicitly mark the photo-like crop in `manifest.json` and redraw it as a clean editable placeholder unless the user specifically asks for raster-like reconstruction. Do not interpret high crop deltas in those regions the same way as table/module failures.
- For vertical VABlock/page-stack or memory-stack diagrams with colored bands and stripe rows, record a band descriptor before drawing: band y-range, fill color, stripe count/positions, dashed boundary y-values, and named route anchors on the stack edge. Uniform stripe loops are only a first pass and often create large crop deltas.
- Do not approximate thick dataflow ribbons by only increasing Visio line weight. Large line weights inflate arrowheads and can cover modules or labels. Prefer a filled-ribbon primitive, or at minimum reduce line weight and move arrow starts outside module boundaries.
- Do not use a single thick double-ended Visio line for bidirectional arrows unless the export crop validates cleanly. The two arrowheads can merge into a star-like blob; draw two one-direction arrows or a filled bidirectional primitive instead.
- Treat legend icons as separate mini-primitives, not scaled-down copies of main dataflow arrows. A thick arrow helper that is acceptable inside a module can over-inflate arrowheads in the legend; use smaller line weights or dedicated filled icon shapes there.
- For compact legends that mix parameter text and color/style swatches, define a local legend layout descriptor with row baselines, icon size, icon/text gaps, and reserved bottom padding. Do not place swatches by eye after a fitted text list; Visio text fitting can wrap one label and make the icon rows collide with the last parameter line.
- For legends in architecture comparison figures, reserve separate lanes for each short hardware token such as `SRAM`, `NoC`, and `Compute units`. Do not reuse the swatch bbox as the text bbox; exported PNGs can split a short token even when the live Visio canvas looks acceptable.
- For thick buses that contain embedded labels, use a narrow label knockout only around the label text when the source visually reserves that space. Do not place transparent dark text directly over a dark bus, and do not use a broad white mask that covers bus arrowheads or neighboring routes.
- Broad filled arrow bands should use `Add-VisioCopyPolygonPx` for the real filled body and arrowhead. If a specific Visio export fails, use a controlled fallback: filled rectangular body, strip-filled arrowhead, and a separate outline polyline. Avoid leaving large arrowheads hollow when the source uses filled bands.
- If `Add-VisioCopyPolygonPx` throws a Visio COM exception in a redraw script, do not abandon the figure or silently skip the shape. Replace the local polygon with an explicit editable fallback, such as offset rectangles for pseudo-3D blocks or thick arrow lines for transition arrows, then record the limitation in the delivery notes.
- For simple horizontal or vertical gradients in architecture blocks, editable strip gradients are an acceptable fallback: draw many thin filled rectangles, then draw the final rounded/outlined border on top. Record this as a gradient approximation in the delivery notes.
- For regular tile arrays with repeated MUX/DEMUX blocks and multiple overlaid bus classes, draw semantic bus layers first, then module boxes/trapezoids, then transparent text. This lets module fills naturally hide bus segments without adding white masks over the diagram.
- For large multi-panel architecture figures with repeated chiplet/PE tiles, define a local descriptor before drawing: x positions, y positions, row fills, cluster frames, local interconnect arrows, inactive tile style, and caption anchor. This is faster and more stable than hand-placing each tile directly in page coordinates.
- For PE-grid architecture figures, define tile/link descriptors for repeated PE boxes, horizontal and vertical bidirectional links, ellipsis markers, crossbars, side buffers, and transpose/reduction units. This avoids drift between repeated rows and columns.
- For symmetric CPU/GPU/CSD or host/device architecture figures, define repeated module descriptors with origin, title, internal child boxes, arrows, and label overrides. Do not hand-place the left and right modules separately when their structure is the same.
- Ultra-wide dense grids need a segment descriptor rather than ad hoc `GridBlock` calls. Record group label, x-range, row/column count, fill pattern, separator marks, outline columns, ellipsis positions, and named column anchors for cross-row arrows.
- Color-coded schedule grids need a grid descriptor plus a separate route layer and legend descriptor. Draw side blocks and grid fills first, route arrows second, and legend/text last so arrows do not disappear under cells or force label masks.
- Scheduling token-chain figures need a descriptor for each row: token label, token width, fill class, side-cap class, row y, and annotation anchor. Narrow chevron tokens need width-aware side caps and a wider text lane; otherwise short labels such as `P1` and `D1` can wrap vertically after export.
- For embedded plots kept as page context in a mixed architecture figure, reserve wide one-line lanes for short tick labels such as `0%`, `20%`, and `90%`. Visio can wrap the percent sign under the number if the tick text box is sized only to the visible glyph width; use a wider box and a local font cap before accepting the crop.
- Dense hardware pin tables and package maps need an explicit table descriptor with column widths, row heights, `rowspan`/`colspan`, per-cell hatch/fill style, and text runs. A uniform grid helper is only a structural first pass and will produce high crop deltas when the source has merged cells.
- Dense numeric matrix/vector tables need a local font cap tied to cell width before rendering. In 30-35 px cells, 20 pt text can wrap two- or three-digit values; reduce to a validated size and audit representative cells before accepting the figure.
- DIMM, package, board, and memory-module diagrams need component-local helpers for board stack layers, connector/notch outlines, pin rows, chip blocks, buffer/data tiles, and repeated bank/PE grids. Keep route layers separate from foreground chips so dashed interconnects do not cover labels.
- For wafer/die/core/package multi-view architecture figures, split the drawing into view-level helpers and reserve explicit one-line lanes for short hardware labels such as `HBM`, `TSV`, `D2D I/F`, and dimension labels. These tokens can wrap in Visio export even when they look safe on the live canvas.
- For die or memory arrays with top and bottom badge rows, reserve the badge lanes before computing inner cell pitch. The last array row must not collide with bottom badges such as `HBM`; repair the component helper rather than nudging only the label.
- Ultra-wide compound figures that mix flowchart, host/device architecture, PE arrays, code pragmas, and matrix grids should be redrawn panel-first. Produce a semantic editable first pass with explicit per-panel limitations instead of spending all time on one dense subpanel; then refine high-delta panels such as nested kernel stacks or matrix grids with local descriptors.
- Multi-column hardware trade-off figures should be redrawn as panel helpers first, then local module/callout helpers. Red italic underlined callouts inside dashed blue frames should use per-line transparent text anchors with a local font cap; a single fitted paragraph often rewraps during Visio export.
- Compiler-to-core microarchitecture figures should use panel-first delivery plus PE-local descriptors: operator tree, icache/control lanes, mask/adapter lanes, buffer stack geometry, bus groups, and I/O interface. A generic buffer-stack helper is a speed pass; high-fidelity PE crops require per-layer offsets and named bus anchors.
- Deployment-plus-pipeline architecture figures need separate descriptors for graph deployment, route classes, core buffer stacks, and timing tables. Dense red address-signal layers should be explicit named route groups instead of ad hoc line loops.
- SSD, flash, and storage-controller diagrams need local helpers for channel buses, die-pair modules, flash-chip queue cell rows, DRAM-chip queue stacks, controller subframes, and numbered badges. Validate short hardware labels beside badges with crops, because labels such as `DRAM` and `SSD` can wrap or be covered even when the full preview looks acceptable.
- DRAM mat and memory-array diagrams need descriptors for repeated cell-array origin, row/column pitch, colored cell pattern, tap rows, bitlines, shared wordlines, selector buses, and IO/sense-amplifier anchors. Treat transistor/capacitor insets and decoder symbols as separate local primitives or explicit editable approximations.
- DRAM organization figures need separate component descriptors for controller/rank/chip stacks, bank cards, subarray dot grids, row buffers, row decoders, bitline/wordline labels, DRAM-cell insets, and sense-amplifier/DCC circuit panels. Keep transistor-level geometry as an explicit refinement target when a semantic editable pass is used.
- QCCD/trap topology diagrams need helpers for trap boxes, channel segments, qubit dots, colored route arrows, dashed dividers, and conflict circles. Treat exact dot ordering and channel hatch as refinements unless they are central to the figure claim.
- QCCD square/ring-step diagrams need a loop descriptor: four trap lanes, corner connector patches, per-lane ion-dot patterns, ellipsis anchors, and a separate route/rotation arrow layer. Large oval approximations are acceptable for a first editable pass, but repeated curved arrows should be promoted into a dedicated curved-arrow primitive when two or more figures need exact arrowhead placement.
- Quantum-circuit labels in QCCD/architecture figures should split the ket bar and the label text. Draw the leading vertical bar as its own thin line, then start the `q0,0>`/`q1,1>` text lane to the right of that bar; do not both draw a bar primitive and include `|` inside the text, because the bar can visually cut through or hide the `q` glyph in Visio exports.
- Memory subarray diagrams with c-groups, B-groups, mask rows, counters, and row buffers need a local descriptor for band y-ranges, row-dot y-positions, column anchors, dotted guide lines, row/column labels, and output arrows. Do not place dot grids by page-global coordinates one row at a time.
- Use flowchart primitives for hardware toolchain and design-flow diagrams: `Add-VisioCopyDocumentPx` for file/document nodes, `Add-VisioCopyCylinderPx` for library/database nodes, and `Add-VisioCopyCutCornerRectPx` for constraint/input nodes.
- For workflow/framework diagrams with icon cards, reserve a local icon lane and a separate text lane inside every card or process box. Even simplified editable icons can collide with labels if they are drawn as decoration without a reserved lane.
- For multi-column flowcharts with a shared bottom merge bar or output lane spanning columns, draw the shared lane after the column-local boxes and audit it with its own full-width crop. Column crops will otherwise clip the shared text and look falsely broken.
- When a connector or vertical arrow is intentionally routed through the gap between words in a flowchart label, split the label into two text anchors around the route slot. Do not draw one wide text box under the connector; the exported crop can make the arrow look like it cuts through a word or removes the space.
- For quadrant comparison diagrams, define panel descriptors for repeated title, operator nodes, PE blocks, fan-in/fan-out guides, time axes, and formula rows. This keeps spatial/temporal or pipeline/share variants consistent.
- For dense graph/network diagrams, use a graph descriptor with node anchors, node class, block membership, and explicit edge lists. A row/column node loop is only a semantic first pass when the source depends on exact dense-vs-sparse edge topology.
- For repeated diagonal DSP/systolic schedule diagrams, define a local descriptor for cell coordinates, active/inactive color class, operator circles, captions, and optional timing lanes. A generic staircase helper is useful for a first pass but should be refined with per-cell coordinates when crop fidelity matters.
- For cropped flowchart side notes or partially visible labels, redraw the visible fragments as editable text and mark the crop context in `manifest.json`. Do not invent the missing full sentence just to make the text box look complete.
- For thick cross-container flow arrows, preserve topology with orthogonal routes first, then decide whether a custom filled arrowhead is needed. Visio default arrowheads can be semantically correct while still producing high local crop deltas.
- For short labels around thick interconnect arrows, such as `FAA`, `NVLink`, and `NCCL`, reserve generous one-line text lanes above and below the link. Do not size these labels to their visible glyph width; Visio can split even three- or four-letter tokens during export.
- Keep component-local origins and helper-level anchors. If a parent frame moves, child boxes, text, route endpoints, and repeated nodes must move together.
- For large framed containers with a title/header plus internal modules, draw the frame as a shape with no centered label, then add the title as a separate transparent text job anchored to the header area. Do not use a full-parent `BoxLabel` when it would center the title over child modules or connector lines.
- For large visual braces around vector/table groups, use explicit line/curve primitives or a dedicated brace helper. Oversized fitted text braces can shrink unpredictably or export as thick blobs that cover nearby cells.
- For protocol-stack, memory, CPU, NIC, and bus containers, reserve explicit local title lanes before placing children. Titles such as `RPC Stack`, `DCOH`, `CPU`, or `NIC Mem` often need a separate wide text anchor; if they are attached to the parent helper, Visio can wrap them or place them over internal icons, rows, and arrows.
- For narrow vertical bus blocks, draw the bus rectangle without a parent label and add only one rotated text anchor. A normal centered parent label plus a rotated label creates duplicate text and apparent collisions.
- For very narrow sidecar badges such as `ATC`, draw the sidecar rectangle unlabeled and place a separate rotated text anchor with a deliberately wider internal lane. A centered `BoxLabel` can wrap the token into broken fragments.
- For badge-like sublabels inside a parent module, such as `Node 1` with an overlaid `NIC 1` badge, draw the parent frame unlabeled, then place the parent title and badge as separate local text/shape components. A centered parent label will collide with the badge.
- For numbered process-step boxes, reserve a badge slot and draw the step text in a separate anchor. Do not draw a centered full label and then place the black number badge on top; the badge will cover the first word or force the text to shrink unpredictably.
- For circular black figure badges such as `A1`, `B10`, or step IDs, make the badge text anchor full-width inside the circle and scale font size by label length. Three-character badges can wrap into two lines even with fitted text if the anchor is too narrow.
- For 90-degree labels, pass the final desired bbox to `Add-VisioCopyTextJobPx` or `Add-VisioCopyBoxLabelPx -Angle`; the helper transposes the text box internally. Do not hand-transpose the bbox unless you deliberately bypass the helper.
- For very narrow vertical labels such as `Crossbar` or `DMA Engine`, validate the export before accepting rotated text. If Visio wraps or shrinks the word into unreadable fragments, use a controlled stacked-text fallback and record that typography limitation in the delivery notes.
- For rotated labels beside dense colored cells, first try a transparent oversized rotation lane that extends beyond the visible narrow strip. This avoids adding a white mask over the cells while giving Visio enough width to keep tokens such as `DevTLB` intact.
- PowerShell variable names are case-insensitive. If a palette is stored in `$C`, never use `$c` as a loop variable; it overwrites the palette and later fills silently degrade to black. Prefer loop names such as `$row`, `$col`, `$idx`, or use `$Palette` for colors.
- In manual PowerShell scripts, call helpers with named parameters after any `[switch]` parameter. Do not use positional `$false` placeholders for optional switches in text helpers; later strings can shift into parameters such as `-Angle` and break the draw run.
- The same rule applies to geometry helpers: never pass `$false` positionally to a `[switch]` before numeric parameters such as `-R`, `-Round`, or `-Weight`. PowerShell can silently bind the boolean to the next numeric parameter and produce zero-size or malformed shapes.

Promote a new primitive into `visio_copy_manual_primitives.ps1` only after it fixes at least two samples or removes clear repeated code. Keep overfit figure-specific helpers inside that figure's `draw.ps1`.

## Manual Primitive Library

`scripts/visio_copy_manual_primitives.ps1` provides reusable building blocks for hand-polished redraws:

```powershell
. "$SkillDir\scripts\visio_copy_manual_primitives.ps1"
Initialize-VisioCopyCanvas -Page $Page -SourceWidthPx 1118 -SourceHeightPx 704
Add-VisioCopyPageBackgroundPx
Add-VisioCopyBoxLabelPx 20 20 120 40 "Controller" "#D9EAF7" 14 -RoundPx 8 -Fit
Add-VisioCopyOrthogonalRoutePx -Points @(@(40,80),@(120,80),@(120,140)) -Color "#000000" -EndArrow 13
Flush-VisioCopyText
```

redraw scripts should prefer this library for:

- common shape/style setup and source-pixel coordinate conversion
- 1:1 96-DPI high-fidelity canvas setup by default
- native Visio shapes only; do not use source-image underlays, imported traces, or pasted image crops
- full-page white background sentinels for stable preview export
- transparent text drawn last
- shared-border grids through `Add-VisioCopySharedGridPx` and `Add-VisioCopyGridCellFillPx`
- fitted labels that reduce font size before wrapping
- rounded module labels through `Add-VisioCopyBoxLabelPx -RoundPx`
- rotated text labels with internal bbox transposition
- orthogonal routes with arrow direction
- tables and message-lane tables
- hatched regions
- hollow block arrows
- isometric router grids with reusable node anchors
- reusable logic/electrical symbols such as inverter triangles, muxes, AND gates, OR gates, and small marker shapes
- flowchart/document primitives such as wavy-bottom documents, cylinders, and cut-corner input boxes

## Core Workflow

1. Open the `.vsdx` with Visio COM visibly by default for manual redraw work, and keep the document open after saving unless the caller explicitly requested a hidden/background run.
2. Create a clean target page with no source image placed on it. For high-fidelity work, set page size from the source image dimensions with `1 px = 1/96 in`, matching Visio's usual 96-DPI PNG export. Use other scales only when the export path has been validated.
3. Keep the source image outside Visio as a measurement and audit reference. Do not place it on the page, even temporarily.
4. Run `scripts/analyze_reference_style.py` on the full reference and on planned component bboxes before detailed drawing. Use its palette, color-region bboxes, edge-direction cues, and crop-local colors to seed the redraw script instead of estimating colors from the full page by eye.
5. Split the reference image into small component crops before detailed drawing. Use crops for every recognizable module, not only dense areas: stage headers, outer containers, controller blocks, tables, legends, arrows, text groups, grids, ports, stacks, formulas, callouts, shadows, gradients, and 3D/2.5D objects. For each crop, record the bbox, sampled colors, visible cell counts, hatch directions, text boxes, line weights, shadow/gradient behavior, and layer/occlusion behavior.
6. Build a separate redraw layer, named like `ManualRedraw_HiRes`.
7. Draw real Visio shapes in source-image pixel coordinates:
   - background and large regions
   - sampled fill/line colors and any shadows or gradients
   - major module blocks
   - inner blocks, grids, and ports
   - 2.5D/3D stack caps, separated blocks, and visible depth hints
   - hollow bus arrows and AXI markers
   - thin orthogonal control lines
   - thick filled data arrows
   - text labels last
8. Run an internal micro-layout pass for every large module. Re-anchor small child modules, labels, bit bars, table cells, arrows, and local text to the parent component bbox instead of nudging them in page-global coordinates.
9. Export a page preview PNG after every meaningful iteration and inspect it visually.
10. Check that the exported preview pixel dimensions match the source image before trusting crop metrics. If the PNG is stretched, padded, or consistently scaled, locate off-page shapes and verify the canvas scale. A preview that is about 0.96x the reference usually means the page used `0.01 in/px` with Visio's 96-DPI export; fix the scale before component auditing.
11. Crop the same component bbox from the exported preview and compare it with the reference crop side by side. For manual redraw directories, store these bboxes in `manifest.json` as `audit_components` and run `scripts/run_manual_crop_audit.ps1 -OutputDir <figure_dir>`. Inspect all planned component crops before editing again, so one pass finds multiple issues instead of discovering one issue per redraw.
12. Convert the inspection into a batch repair list. For each issue, record the exact component, bbox, failure type, target coordinate/text/color/layer change, and script location/helper to edit.
13. Apply all precise repairs in one script-edit pass, then redraw/export once. Do not repeatedly rerender after every single small finding unless a script error or very high-risk layout change needs isolation.
14. Fix by sampled color, coordinates, layer order, text box sizes, font sizes, and component-specific drawing primitives. Do not hide problems by pasting the source image on top.
15. After the final version is accepted, verify the Visio page contains only the final editable native-shape redraw layer and supporting native-shape groups. Remove any experimental layers, imported objects, or accidental image shapes.

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
- Check color/style: sampled fill colors, line colors, gradients, transparency, shadows, and face shading for 3D blocks. If colors differ, sample the crop palette again instead of guessing a nearby named color.
- Check draw order: large white backgrounds and color regions must not cover later text, borders, arrows, or small symbols; if they do, fix the drawing order instead of moving shapes by eye.
- Check text layout: font family, size, bold/italic style, line breaks, alignment, margins, formula-style labels, and whether any word or formula is split, wrapped, clipped, or overlapped.
- Check internal micro-layout: child boxes, bit bars, icons, per-cell labels, and local arrows must be anchored to the large module's crop bbox. If small text is shifted while the big frame is correct, fix the component-local anchor math, not the page-level frame.
- Check text-line collisions explicitly: connector lines and arrows must not cross labels unless the source does. Prefer rerouting the line or shortening the text box; use a masked label background only as a last resort.
- Treat port labels on processor, NoC, memory, or accelerator blocks as anchored micro-components. Reserve a small label lane near the route and attach the connector outside that lane; do not let connector lines run through `Port0`, `Port1`, lane names, queue names, or similar local labels.
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
   - `micro-layout`: fix child-module anchors, local row/column origins, bit-bar widths, label baselines, and local arrow endpoints inside a parent module.
   - `collision`: move/reroute lines or add a masked text label when text and linework overlap.
   - `color-style`: replace guessed fills/lines with sampled HEX/RGB colors, adjust gradient strips, transparency, shadows, or 3D face shading.
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
- `color-style`: fill colors, line colors, shaded faces, transparency, gradients, or shadows do not match the source crop.
- `micro-layout`: the outer module is correct but internal child modules, small labels, bit bars, or local arrows are mispositioned.
- `collision`: text and lines overlap or a connector crosses a label unintentionally.
- `layering`: backgrounds, large fills, table cells, or rear stack hints cover foreground outlines, arrows, or text.
- `missing`: ports, arrows, dots, repeated cells, labels, separator marks, or formula pieces are absent.

Then build one batch repair list covering several components before editing. If several failures share the same cause, repair the reusable helper or primitive first rather than hand-tuning every instance. Examples:

- If formula labels become tiny, clipped, or scattered, fix or replace the formula helper and validate it on one crop before applying it globally.
- If every stacked-grid object reads as one solid cube, replace the stack primitive with separated per-cell blocks and occlusion masks.
- If colors are visibly off across many components, rebuild a named palette from `analyze_reference_style.py` and update shared color constants before touching geometry.
- If a shaded/3D object is flat or visually reversed, redraw it as front face plus separate top/right caps using lighter/darker sampled face colors.
- If repeated modules have the wrong cell count, update the component manifest and loop counts, not only the outer bounding box.
- If a large frame is right but internal text or child boxes are offset, introduce a component-local coordinate origin and helper rather than patching each label in global page coordinates.
- If text is repeatedly covered, repair draw order and text-layer timing instead of moving labels away from the correct coordinates.
- If a connector must pass near a label, draw the connector first and the label last, but keep the text fill transparent. Use `Add-MaskedTextPx` only when the reference itself has a label knockout or when rerouting is impossible.
- If a more faithful helper makes the crop less readable or more cluttered than the previous iteration, roll it back immediately and keep the stable helper until a replacement is validated in isolation.

Use a staged quality gate:

1. Global layout and major module positions.
2. Parent-level component geometry and repeated element counts.
3. Internal micro-layout of child modules, bit bars, labels, and local arrows.
4. Text-line collision removal and text-last layering.
5. Sampled colors, gradients, shadows, and 3D face shading.
6. Arrow topology, line weights, and layer ordering.
7. Text placement and formula readability.
8. Final native-shape purity check: no source image, no imported trace, no pasted screenshot, and no auto-vectorized pixel group.

Do not overfit fine text before the component primitive and repeated counts are correct. Do not keep a broken helper just because some instances look close.

## Coordinate Rules

Use source pixels as the source of truth:

```powershell
$Scale = 1.0 / 96.0
$PageWidth = $SourceWidthPx * $Scale
$PageHeight = $SourceHeightPx * $Scale
function To-X { param([double]$Px) return $Px * $Scale }
function To-Y { param([double]$Py) return $PageHeight - ($Py * $Scale) }
```

Record each shape as `(x, y, w, h)` where `(0,0)` is the top-left pixel of the reference image. Convert only inside helper functions.

## Color And Style Capture

Do not guess colors from the whole-page preview. Sample colors from the component crop where the shape appears.

Run a full-page palette pass first:

```powershell
python scripts/analyze_reference_style.py reference.png --top-colors 18 --json
```

Then rerun it with component bboxes for local colors and structure cues:

```powershell
python scripts/analyze_reference_style.py reference.png `
  --component stack_left:40,80,160,120 `
  --component controller:260,110,180,90
```

Use the output as a named palette in the redraw script:

```powershell
$C = @{
  module_blue = "#DDEAF7"
  line_gray = "#6F7780"
  accent_red = "#E8332A"
}
Add-RectPx 20 20 120 60 $C.line_gray $C.module_blue 1.0
```

Rules:

- Use sampled HEX/RGB colors for fills and lines. `visio_manual_redraw_scaffold.ps1` accepts both `#RRGGBB` and `RGB(r,g,b)`.
- For antialiased boundaries, choose the dominant interior crop color for fill and a darker edge color for stroke.
- For gradients, approximate with `Add-GradientRectPx` using sampled start/end colors and enough strips to look smooth at exported size.
- For shadows, draw a transparent offset shape with `Add-ShadowRectPx` or a custom transparent polygon before the foreground object.
- For 3D faces, use separate sampled or adjusted colors for front, top, and right faces. Do not use one flat fill when the source uses shaded faces.
- After export, use `crop_compare.py` and inspect the diff/metrics files. A high color delta in a crop is a style failure, not a coordinate failure.

## Complex Shapes And 3D Primitives

Choose the primitive from the crop, not from the object's semantic name:

- Flat rounded module: `Add-RectPx` or `Add-ShadowRectPx`.
- Smooth color band: `Add-GradientRectPx`.
- Single 3D block: `Add-IsometricBlockPx` with explicit front/top/right fills.
- 2.5D depth hint behind a front cell: `Add-DepthCapPx` before the front face.
- Separated stacked matrix: `Add-SeparatedStackedMatrixPx` with `GapX`, `GapY`, `DepthDx`, and `DepthDy`.
- Dense table/grid with no perspective: `Add-CellMatrixPx` with exact rows, columns, gaps, and style tokens.

If a complex component cannot be represented by one helper without hidden-line clutter, write a component-specific helper. Prefer explicit polygons and per-cell blocks over a generic cuboid loop.

## Internal Micro-Layout Rules

Most unsatisfactory redraws are not caused by the large frames; they come from small modules, labels, and connectors inside the frames.

For every large module:

- Define a local origin `(module_x, module_y)` and place children as offsets from that origin.
- Keep child rows/columns in local arrays, such as `row_y = @(0,52,104,156)` and `col_x = @(0,120,250)`, instead of scattering absolute page coordinates.
- Anchor bit bars, small cells, and local labels to the same row baseline. If a row moves, all children in that row must move together.
- When a blue label/bit bar sits on top of repeated gray cells, draw it as a segmented primitive: repeated gray cells first, then a blue outline with no fill, a small solid `S` tab, a white text segment only under the numeric label, and a transparent tail where the gray cells remain visible. Use `Add-SegmentedBitBarPx` or an equivalent component helper. Do not use one large white-filled rectangle that hides the cell pattern behind it.
- Use fixed text boxes sized from the crop. Reduce font size before allowing Visio to wrap a short label.
- Treat line endpoints as part of the component manifest. Local arrows should terminate at the child shape edge, not at guessed global points.
- Draw internal fills and lines first, then text labels last. The scaffold records `Add-TextPx` text shapes and brings them to front before saving.
- Keep normal text transparent. A large white text box hiding bit bars, arrows, or frame edges is a failure.
- Use `Add-MaskedTextPx` sparingly, only as a tight mask behind the text glyph area. The mask must not cover adjacent arrows, bit bars, frame borders, or other modules.
- If a line crosses text, first shorten the line, move the endpoint to the text-box edge, or route around the label. Do not hide the collision with a broad white rectangle.
- Fail a component crop if any short label wraps unexpectedly, drifts outside its child shape, or is crossed by a line that is not in the source.

## Layering Rules

Draw order matters. A common failure is drawing correct shapes but then bringing an entire layer to the front, causing large background blocks to cover text and small elements.

Use this order:

1. page border, shadows, and large fills
2. region fills, gradients, and background bands
3. rear/depth caps and visible 3D hints
4. module rectangles and front faces
5. inner grids, per-cell hatches, and small symbols
6. thin lines and connectors
7. thick filled arrows
8. text labels

Never call `BringToFront()` on every shape in a redraw layer. If a single object needs adjustment, bring only that object forward.

## Final Delivery Cleanup

During iteration and delivery, the Visio page must remain a native-shape drawing. Do not keep a source-image underlay or any tracing aid on the page.

Before delivery:

- Save a backup first.
- Delete old trial redraw layers and preview/debug layers that are not the final accepted version.
- Keep only editable native Visio shapes and native-shape groups.
- Check that the document has no source-image shapes, imported SVG/PDF trace objects, pasted screenshot layers, or auto-vectorized pixel groups.
- Re-export a final preview after cleanup and inspect it as a standalone native Visio drawing.
- Report that the delivered page was built only from Visio-created shapes.

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

Before drawing a stack, crop it from the reference and run `analyze_reference_style.py --component name:x,y,w,h` on that crop. Use the edge-direction cues to distinguish flat tables, diagonal hatch marks, 2D plane stacks, isometric depth, true cuboid cell faces, and orthogonal stepped caps. Paper diagrams often use separated flat planes, not a continuous cuboid grid. If the source shows 2D planes, every visible plane should be drawn as a flat grid and layers must be ordered from rear to front. Only draw per-cell top/side faces when those faces are actually visible in the crop.

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
- Prefer `Add-SeparatedStackedMatrixPx` when each visible cell has a real gutter. Tune `GapX`, `GapY`, `DepthDx`, and `DepthDy` from the crop instead of changing the overall bbox first.
- Use `Add-DepthCapPx` for top/right face hints when the front cell matrix must stay opaque and uncluttered.
- If the output still reads as one solid cube, the primitive is wrong; replace it with a component-specific block helper instead of tuning coordinates.

For cache/key-matrix stacks in paper figures, inspect whether the rear layers are actually orthogonal stepped caps rather than diagonal perspective faces. If the crop shows U-shaped top steps and mostly vertical rear edges:

- draw the front cell column first as the counted matrix, usually four visible cells for bit planes if the crop shows four divisions.
- draw rear layers as stepped U-shaped top caps plus only the visible right-side vertical hints.
- treat every visible top cap/block as an independent shape with a real white gap. The cap width must be smaller than the column pitch; never let adjacent U-caps share an edge or join into a continuous bridge.
- For flat 2D plane stacks, use a `FlatPlaneStack` style helper: draw each whole plane as cell rectangles at its layer offset; do not draw cuboid top faces, side faces, or hidden rear grid lines that are not visible. Draw the farthest offset plane first and the nearest/front plane last, so foreground planes cover only the regions they should.
- if a repeated stack appears connected in the preview, first check `cell_w + gap` versus top-cap width, then shorten the cap or add explicit white gutters/masks before changing global coordinates.
- for cache/key-matrix stacks specifically, draw rear/top cap hints before the front cell columns, then draw the front cells with opaque white/colored fills so they occlude lower cap drops. If cap ticks appear inside the front faces or through column gaps, the draw order is wrong.
- do not draw one rear vertical guide through every column gap unless the source visibly shows it. Prefer sparse right-edge depth silhouettes; otherwise separated blocks collapse into one cuboid-like grid.
- run a 2x/3x crop audit that verifies every front column has the source cell count, the top caps have visible white gutters, and no rear hidden line passes through a front block gap.
- avoid diagonal connector lines and rear bottom edges unless the source explicitly shows them.
- keep rear-line weight lighter than front cell borders so the stack does not become a cuboid.

## Use The Bundled Scripts

- `scripts/analyze_reference_style.py`: run before redrawing to extract global and crop-local palettes, major color-region bboxes, edge direction fractions, and cues such as `dense_grid_or_table`, `hatch_or_perspective_edges`, or `likely_2_5d_or_stacked`. Use this as the first tool for color fidelity and complex/3D shape diagnosis.
- `scripts/extract_color_components.py`: run on the reference image to get bounding boxes for major color regions. Use this to seed coordinates instead of guessing.
- `scripts/crop_compare.py`: after exporting a Visio preview, crop matching component bboxes from the reference and preview, then save side-by-side comparisons, diff images, and color-delta metrics for detailed local repair.
- `scripts/run_manual_crop_audit.ps1`: wrapper for manual redraw directories. It reads `manifest.json` `audit_components`, then invokes `crop_compare.py` against `source.png` and `preview.png`.
- `scripts/visio_manual_redraw_scaffold.ps1`: copy or adapt for a target `.vsdx` only after removing any underlay/import behavior. It provides Visio COM setup, coordinate conversion, layer cleanup, preview export, color conversion, transparency, gradient, shadow, polygon, 3D block, separated stacked-matrix, and drawing helpers.
- `scripts/finalize_visio_copy_page.ps1`: after the final redraw is accepted, back up the `.vsdx`, delete all shapes not on the final layer, reject image/import/trace layers, and save a clean final page.

Read `references/redraw-checklist.md` before finishing a redraw task. For 2.5D stack diagrams, also read `references/stacked-grid-mode.md`.

## Validation Checklist

Before finalizing:

- Export a preview PNG from the Visio page.
- Inspect the preview, not just the live Visio canvas.
- Verify the manual redraw layer is separate and editable.
- For final delivery, verify no reference image, imported trace, pasted screenshot, SVG/PDF trace, or auto-vectorized pixel group exists on the page.
- Check that no experimental trace page or trace layer remains in the delivered file.
- Confirm text does not wrap unexpectedly or get covered.
- Confirm sampled fills, lines, gradients, shadows, and 3D face colors match component crops closely enough.
- Confirm main arrows are horizontal/vertical where the source is horizontal/vertical.
- For stacked grids, confirm hidden rear grid lines are not drawn unless visible in the source.
- For stacked/3D objects, confirm separated cells stay separated, top/right caps do not merge into a continuous cuboid, and face shading matches the source direction.
- For stacked grids, confirm each slash/cross belongs to a specific cell, not a free-floating diagonal line.
- For every dense component crop, confirm the crop-level preview matches the reference before judging the full page.
- Report the page name, layer names, preview path, and backup path to the user.
