# Geometry and text fidelity

## Why SVG import breaks

SVG import and conversion to shapes are separate operations. A picture may display correctly while native conversion changes text positioning or loses clip paths. Math text often consists of individually positioned `tspan` glyphs rather than one simple string. Font substitution, anchor differences and discarded transforms may stack them. These are diagnoses to verify in the actual source, not claims that every PowerPoint version behaves identically.

## Recovering a figure

Use source plot coordinates where possible. Resolve transforms and unit conversions once; PowerPoint uses 914400 EMU/inch, or 12700 EMU/point. Match the original figure size instead of forcing a manuscript panel into a 16:9 slide.

Clip open lines with a segment/rectangle algorithm such as Liang–Barsky. Assemble consecutive retained segments into connected polylines. Clip a filled polygon separately, for example with Sutherland–Hodgman. Unsupported holes, compound fills, nonrectangular masks and Bézier curves require an appropriate geometry library or source-level reconstruction. Never silently interpret all paths as lines. For approximated curves, specify a tolerance in output units and verify it visually.

Reconstruct mathematical labels into ordered native rich-text runs. Derive ordering from logical content or verified source positions, not blindly from XML traversal order. Use the actual font's metrics when positioning text. Font baseline offsets in DrawingML are percentages multiplied by 1000; positive values raise text. A subscript also usually uses a smaller run size. Zero text-box margins, no wrapping, and no automatic shrink-to-fit make positioning predictable, but boxes still need enough spare width for PowerPoint rendering.

Native shapes are generally editable in batches. Thousands of mesh segments can share one freeform object with multiple subpaths; a user can then change that layer's stroke without selecting thousands of objects. Explain this granularity. Do not merge labels with paths.

Do not redistribute fonts. Check font availability and clearly report substitution. Support Unicode µ, δ, γ and minus signs without replacing them with images.

Legacy symbol fonts may put a mathematical glyph in an ordinary character slot. In the gallery fixture, Matplotlib's cmsy10 stores proportional-to at `/`; copying that slash into Arial changes the formula. Decode the verified glyph to Unicode and use a font that contains it (Cambria Math in this fixture). Also check rendered super/subscript size: the tested PowerPoint build shrinks baseline-shifted runs to 65%, so supplying an already reduced SVG font size applies the shrink twice. The fixture adapter compensates for that behavior; verify it again for other renderers.
