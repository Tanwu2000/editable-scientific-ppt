---
name: editable-scientific-ppt
description: Create or repair scientific figures as native editable PowerPoint objects, with editable text, curves and mesh lines, correct mathematical labels and preclipped subplot geometry. Use for manuscript figures that must remain editable in PPT, especially when SVG conversion causes overlaps or escaping lines.
metadata:
  version: "0.1.2"
---

# Editable scientific PowerPoint figures

Produce a PPTX whose text can be edited as text and whose curves can be edited as shapes. Preserve the user's layout, units and scientific evidence. This skill concerns figure construction, not general slide storytelling or scientific parameter fitting.

## Choose the source

Prefer plotting code and source data, then a vector source with live text. A screenshot alone cannot recover exact data or obscured labels: reconstruct only supported content, disclose approximations, and request the missing source when it affects accuracy. Do not trace a raster chart and call its coordinates measured data.

Inspect the reference before rebuilding. Identify panel bounds, figure dimensions, fonts, labels, legend positions and each layer's required editability. Honor explicit styling; when unspecified, Arial with black text is a useful starting point, not a compulsory replacement for the user's style.

## Essential construction decisions

- Deliver the complete figure as a native PowerPoint group, with named panel subgroups where useful. Keep every label, path and raster as an editable child. Verify grouping, ungrouping and regrouping of all leaf objects, including mixed text/vector/image selections; move and resize the group, then save and reopen. Read [grouping-and-style.md](references/grouping-and-style.md) for nongroupable placeholders, tables, chart/OLE objects and group coordinate transforms. Do not treat individual editability as proof of groupability.
- Center each panel title on its plotting area using an explicit text box spanning that area, not the slide center or a guessed text width. Keep panel letters separate. Use a clean manuscript layout: no unsolicited subtitles, gray microtext, generation notes, editing instructions or repeated explanatory captions inside the figure. Keep necessary units, legends, scale bars, statistical labels and scientific annotations; put workflow explanations outside the artwork.
- Default axes, ticks, frames, borders and text to opaque black; use Arial unless a requested style says otherwise. At final figure size start with axes/borders 1.2 pt, curves 1.5 pt and mesh 0.7 pt. Keep series colors meaningful and avoid faded strokes, translucent edges or hairlines. These are adjustable starting values, not forced overrides of explicit user styling. Check dense meshes for over-darkening. Read [grouping-and-style.md](references/grouping-and-style.md) for raster resolution and final-size checks.
- Use native text boxes for titles, axes, ticks, legends and annotations. Keep Greek letters and units as Unicode, and simple super/subscripts as rich-text runs with baseline offsets. For complex equations prefer native equation objects if supported; explain any fallback to a picture. Converting lettering to outlines does not preserve text editing.
- Use native lines or freeform paths for curves, mesh, boundaries, markers and arrows. Name objects by panel and layer. Batch dense meshes into manageable shapes, while keeping text independent. Distinguish editable paths from data-linked charts; editable paths do not provide an Excel data table.
- Apply all transforms and geometrically clip paths to subplot bounds **before** writing native shapes. Do not rely on SVG `clipPath` surviving PowerPoint's conversion. White masking rectangles conceal geometry but do not repair it.
- Filled polygons and open polylines need different clipping. Keep stroke/fill behavior, subpaths and layer order. Read [geometry-and-text.md](references/geometry-and-text.md) when adapting SVG or mathematical labels.
- Continuous fields, microscopy images or very dense color rasters may remain images. Keep surrounding text and axes independently editable, and disclose raster layers. Do not flatten the entire figure merely to pass a visual check.
- Export SVG with live text when it is an intermediate source. Outlined SVG is useful for static appearance but cannot provide editable font sizes. Do not use PowerPoint's “Convert to Shape” as the required editing step.

## Implementation

Use the native presentation authoring library available in the environment. The portable helper [scripts/native_pptx.py](scripts/native_pptx.py) is a bounded fallback: it writes a one-slide PPTX from explicit scene coordinates with Python's standard library. It is **not** a general SVG/PDF importer. The scene schema and commands are in [scene-format.md](references/scene-format.md). The included synthetic example exercises clipping, rotation, Unicode and subscript text.

Keep the original source and intermediate geometry separate from final output. If adapting a converter, reject unsupported transforms, path commands, fonts or masks instead of silently dropping them. Do not copy machine-specific runtime paths into reusable scripts.

For reproducible examples and actual PowerPoint comparison renders, see [examples/gallery/README.md](examples/gallery/README.md). Its SVG adapter is limited to those Matplotlib fixtures, not a general importer.

## Verification and delivery

1. Check the PPTX structurally with `scripts/inspect_pptx.py`: confirm native text and vector objects, readable XML, slide dimensions and absence of external links. Structural success alone does not establish visual correctness.
2. Render and inspect the final deck, preferably in the user's target PowerPoint version. Check math, rotated axes, panel clipping, title centers, absence of unnecessary microtext, black border contrast, raster orientation and legend overlap at final size. Read [powerpoint-qa.md](references/powerpoint-qa.md) for grouping and edit/save/reopen checks. Report honestly when PowerPoint is unavailable.
3. Test changes on a disposable copy, keeping the delivered layout unchanged. Never close, save over or quit the user's unrelated open presentations.
4. If requested, export matching PNG and TIFF from the final figure version; use 600 dpi by default for manuscript mixed line/text figures unless a specification says otherwise. Match physical size and pixel dimensions; a DPI tag or upscaling alone does not add detail. Keep vectors native and check embedded field/image effective resolution at placed size (normally at least 300 ppi for continuous-tone layers). State which layers remain raster images and which objects are editable.

Using this skill does not authorize uploading manuscript data, installing dependencies, publishing a repository or scheduling updates. Follow the current user's scope and the environment's permissions.
