# Grouping and clean figure defaults

## Native groups

Use slide-owned text boxes, freeforms and pictures, with unique object IDs and no `noGrp` locks. Use an outer figure group and named panel subgroups; a dense mesh can remain a batched freeform inside its panel. Do not put selectable figure components on the master or layout. The actual slide background can stay a background.

If Group is unavailable, inspect the selected object types and their slide ownership. Placeholders, tables, chart/OLE objects and some imported graphic objects need separate handling. Rebuild the required figure content as ordinary text/paths/pictures when faithful reconstruction is possible. Preserve editable data-driven charts separately if required; explain any incompatible grouping requirement instead of flattening the whole figure into an image.

The portable writer now defaults to a figure group. `group: false` at scene root produces flat objects. A `kind: group` element contains an `elements` array, whose positions remain in slide coordinates. It writes matched `off/chOff` and `ext/chExt` bounds, so introducing a group does not move its children. Bounds account for rotated text. Panel grouping must preserve stacking order, clipping and relative alignment.

Test both the hierarchy and regrouping of the fully ungrouped leaf objects. Then move and resize the group, edit a child label and line, save and reopen. A single green XML check is insufficient. Use the real target PowerPoint renderer when available. [Microsoft Group API](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.group) and [Ungroup API](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.ungroup).

## Titles and figure text

For a plot spanning `left` to `right`, title box `x=left`, `width=right-left`, alignment=center. Use the plot region rather than the figure canvas including axis labels or a separate colorbar. Leave enough height for the font. Check the visible text as well as the box alignment, especially with math and font substitutions. A figure-wide title, if requested, centers on the overall panel area. Do not add a title when the intended figure has none.

Keep only text that helps interpret the science. Necessary units, legends, sample sizes, significance labels and scale bars remain. Explanations of tools, generation, editability, rendering settings and analysis provenance belong in the response or figure caption unless specifically requested inside the image.

The optional scene roles are explicit metadata, not automatic text classification. Only text deliberately tagged `role: helper` is hidden by default; `showHelperText: true` retains it. Untagged text is preserved. Never delete labels based on being gray, small or unfamiliar.

## Contrast and resolution

At the final placed figure size, start with opaque black axes/borders/ticks at 1.2 pt, data curves at 1.5 pt and mesh at 0.7 pt. Use black text; retain meaningful series colors. Adjust dense mesh widths to prevent dark merging. Explicit user colors and widths override these defaults. A sparse figure does not need a decorative background grid, subtitle or footer.

Native paths/text stay vector in PPTX. Raster preview/export size is `round(width_pt / 72 * dpi)` by `round(height_pt / 72 * dpi)`. Default mixed text/line manuscript previews to 600 dpi when no target is specified. For continuous fields, check source pixel dimensions divided by placed inches, normally at least 300 ppi. Exporting a 72 ppi image at 600 dpi does not recover missing detail. Do not sharpen or alter scientific data to make a figure appear crisper.

Inspect at final publication size and zoom in on edges. Increasing zoom is not a substitute for checking readability at the intended size. Export TIFF from the same rendered version, with matching dimensions, explicit resolution metadata and lossless compression.
