# Verify in PowerPoint

Use a test copy of the generated PPTX, opening it without altering any user presentation.

On Windows with desktop PowerPoint installed, the optional `tools/check_powerpoint.ps1` opens only the supplied file, renders a PNG, edits a disposable copy, and verifies persistence. It does not quit PowerPoint. Use a new output folder:

```powershell
./tools/check_powerpoint.ps1 -InputPptx ./build/example.pptx -OutputDirectory ./build/powerpoint-qa
```

This uses the PowerPoint COM interface and may require interactive desktop permissions; follow the host's permission flow. Inspect the produced PNG separately. The script cannot judge its visual quality.

1. Compare its appearance with the approved figure: title centering, panel positions, Greek letters, subscripts, rotated y labels, color field orientation, curves at plot boundaries, legend location and stroke contrast.
2. Select a native title text box, increase its font size and change its text. Select a mesh/curve freeform and alter line width/color. A picture's resize handles alone do not establish editability.
3. Save the test copy, close only that copy, reopen it, and confirm both edits persisted. Verify the original deliverable was not modified.
4. If requested, export PNG at `round(width_points / 72 * dpi)` by `round(height_points / 72 * dpi)`. TIFF should come from that same rendered version with the requested resolution metadata and suitable compression. Inspect the raster output.

If PowerPoint is unavailable, use another renderer for a visual check and report that it is not a PowerPoint compatibility test. The XML audit cannot check glyph placement or renderer-specific problems. A single successful test does not guarantee every Office version or operating system.
