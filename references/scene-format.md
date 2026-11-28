# Portable scene exporter

Requires Python 3.10+; standard library only. No network calls, COM, shell execution or third-party runtime is used by the exporter. Inputs are explicit coordinates, not arbitrary SVG.

```sh
python scripts/native_pptx.py examples/minimal_scene.json build/example.pptx
python scripts/inspect_pptx.py build/example.pptx
python -m unittest discover -s tests -v
```

All coordinates, slide dimensions and font sizes are in points (72 points = 1 inch). Origin: top left; y increases downward. Elements are drawn in array order. The root has `width`, `height`, and `elements`.

Since v0.1.2, root `group` defaults to `true`, wrapping the figure in a native group; set `false` for flat output. Explicit groups use `{"kind":"group","name":"Panel A","elements":[...]}`. Child coordinates remain absolute slide coordinates, with an identity group transform. A single existing group is not wrapped again. Image paths remain relative to the scene file.

Optional `role` values: `title`, `axis`, `border`, `curve`, `mesh`, `helper`, `label`. A text element with `role: title` requires `centerOn: [left,right]`; this sets its x/width and center alignment. Other text position/size fields are still required. Axis/border stroke defaults to black at 1.2 pt, curves to 1.5 pt and mesh to 0.7 pt; explicit colors/widths are preserved. Only elements deliberately marked `helper` are omitted, unless root `showHelperText: true` is set. Untagged scientific labels are preserved. See [grouping-and-style.md](grouping-and-style.md).

Supported element kinds:

- `text`: `name`, `x`, `y`, `width`, `height`, optional `rotation` in degrees, `alignment` (`left`, `center`, `right`), `font`, `color`, `size`, and either `text` or `runs`. Runs contain `text` and `style` with `size`, `font`, `color`, `bold`, `italic`, `baseline` (DrawingML integer units). Set explicit boxes with spare space; the exporter does not measure or fit text.
- `shape`: `name`, `paths` (each has `points: [[x,y],...]` and `closed: true/false`), `fill` and `stroke` (`#RRGGBB` or `none`), `lineWidth`, optional `dash: true` and rectangular `clip: [xmin,ymin,xmax,ymax]`. Straight polylines and convex closed polygons are supported. Filled shapes with holes and concave clipped polygons are outside this helper's scope; use an appropriate authoring backend.
- `image`: `name`, `file` (PNG/JPEG relative to the scene JSON's directory), `x`, `y`, `width`, `height`. Images must already be cropped and oriented correctly; shape clipping does not apply to them. No remote images or external relationships are created. Preserve aspect ratio in the supplied dimensions.

Unknown fields are rejected, so an unsupported property cannot silently disappear. For example, opacity/gradient fills, arrows and Béziers are not understood: model supported geometry explicitly or use a more capable backend. Colors must be opaque 6-digit hex values.

The exporter refuses to overwrite an existing PPTX. Use a fresh build path. PNG/TIFF conversion is not bundled because a correct raster export requires a rendering engine (PowerPoint or another verified renderer), not merely a file extension change.
