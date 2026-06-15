# Stacked-Grid Mode

Use this reference when the source image contains 2.5D tensors, bit planes, stacked memory arrays, cube piles, or repeated cell grids.

## Diagnosis

If a redraw looks chaotic, check for these failures first:

- A generic loop drew every hidden layer and every cell.
- Rear layers were drawn as filled rectangles instead of transparent outlines.
- Rear-layer lines pass through the gaps between front cells.
- A separated block stack was approximated as one continuous cuboid.
- The visible blue/white/red/cross cell count was guessed instead of counted from a crop.
- The fill colors were guessed globally instead of sampled from the stack crop.
- The hatch direction was copied globally, even though red and blue hatch marks may use opposite diagonals.
- Red or blue hatch lines were drawn across a region instead of inside individual cells.
- Front, top, and right faces were drawn with one flat color, so the stack lost its 3D direction.
- Text and dimension arrows were positioned globally rather than anchored to the component.
- A previous source image size was reused after the screenshot changed.

## Crop-First Workflow

For dense stacked diagrams, do not start by coding the whole page. First split the reference into component crops:

1. Crop each stack, row of blocks, legend, arrow group, text group, and callout into separate images.
2. For each crop, write down the bbox in source pixels.
3. Run `scripts/analyze_reference_style.py reference.png --component name:x,y,w,h` and record the crop palette plus edge-direction cues.
4. Count visible front cells and visible rear layers.
5. Mark each cell style in a matrix: `cyan`, `white`, `blue_slash`, `red_slash`, `cross`, or `red`.
6. Record hatch direction for each style, not just the color.
7. Record sampled front/top/right face colors or the lighten/darken rule used for face shading.
8. Record which layer hides which lower layer.
9. Draw only that component and export a preview.
10. Crop the same bbox from the preview and compare it side by side with the reference crop.
11. Iterate on the component until the local crop matches before moving to the next component.

This prevents a clean-looking full-page preview from hiding local count, hatch, or occlusion errors.

## Component Manifest

Write a small manifest before drawing a complex stack:

```text
name: left_query_tensor
bbox: 55,119,91,68
front: rows=4 cols=6 cell=12x15 gap=5x2
depth: layers=5 dx=5 dy=4
cell_styles:
  row0: blue_slash blue_slash blue_slash blue_slash red_slash red_slash
  row1: blue_slash blue_slash blue_slash blue_slash red_slash red_slash
  row2: blue_slash cyan blue_slash cyan red_slash red_slash
  row3: cyan cyan cyan cyan red_slash red_slash
visible_edges: front_cells, top_step_edges, right_rear_edges
colors: front=#5CD3EF top=#8BE6F5 right=#44A9C4 line=#000000
occlusion: draw layers back-to-front and mask the next/front layer footprint before drawing it
```

The exact matrix should be adjusted per source image, but the redraw should always be matrix-driven.

## Drawing Rules

- Treat separated block stacks as separated blocks. Do not draw a large cuboid and subdivide it.
- Draw rear/depth hints first, but only the visible top/right edges or visible rear cells.
- Draw layers back to front. Before drawing a front layer, use a white mask over the layer footprint if rear lines would otherwise show through gaps.
- Draw the front matrix next, using the counted cell matrix from the crop.
- Use sampled crop colors in the matrix palette. For shaded blocks, use a lighter top face and darker right face unless the crop shows different sampled colors.
- For orthogonal U-shaped cache/key-matrix caps, every cap is a separate block. Use `cap_width < column_pitch` and preserve a visible white gutter between caps; do not let adjacent caps share a border or become one continuous bridge.
- For cache/key-matrix stacks, draw top caps and right-edge depth hints before the front cells. The front cell fills must hide lower cap drops; visible tick marks inside front faces mean the layer order is wrong.
- Do not place rear vertical hints in every front-column gap by default. Keep hidden-line hints sparse and usually confined to the right silhouette unless the reference crop clearly shows them.
- Fail the crop if front columns have the wrong cell count, if any U-cap touches the next cap, or if a rear hidden line turns the separated blocks into one continuous cuboid.
- Add hatch and cross marks inside the cells after each cell rectangle.
- Use per-style slash direction. For example, red hatch may be `\` while blue hatch is `/` in the same component.
- Avoid long red diagonal guide lines unless they are annotation arrows in the source.
- Keep rear layer count visually approximate only after front visible cell counts are exact.
- Export a crop of the component for comparison before fixing the whole figure.
- If the stack still reads as one solid cube, discard the generic helper for that component and write a component-specific helper.

## Visio Helpers

Prefer helper functions like:

```powershell
Add-OutlineRectPx $x $y $w $h
Add-CellStyledPx $x $y $cellW $cellH "#FFFFFF" "#FF0000" $false "\"
Add-VisibleStackOutlinePx $x $y $w $h $layers $dx $dy
Add-DepthCapPx $x $y $cellW $cellH $dx $dy "#000000" "#5CD3EF" "#8BE6F5" "#44A9C4"
Add-SeparatedStackedMatrixPx $x $y $matrix 13 17 3 3 6 4 1 $palette
```

Do not use a full `for layer -> for row -> for col` cell loop for rear layers unless the source explicitly shows every rear cell and the occlusion has been handled. In most paper-style diagrams, rear cells are partially hidden by later layers.
