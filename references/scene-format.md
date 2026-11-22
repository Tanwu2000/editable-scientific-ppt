# Portable scene exporter

Requires Python 3.10+; standard library only. No network calls, COM, shell execution or third-party runtime is used by the exporter. Inputs are explicit coordinates, not arbitrary SVG.

```sh
python scripts/native_pptx.py examples/minimal_scene.json build/example.pptx
python scripts/inspect_pptx.py build/example.pptx
python -m unittest discover -s tests -v
```

All coordinates, slide dimensions and font sizes are in points (72 points = 1 inch). Origin: top left; y increases downward. Elements are drawn in array order. The root has `width`, `height`, and `elements`.

Supported element kinds:

- `text`: `name`, `x`, `y`, `width`, `height`, optional `rotation` in degrees, `alignment` (`left`, `center`, `right`), `font`, `color`, `size`, and either `text` or `runs`. Runs contain `text` and `style` with `size`, `font`, `color`, `bold`, `italic`, `baseline` (DrawingML integer units). Set explicit boxes with spare space; the exporter does not measure or fit text.
- `shape`: `name`, `paths` (each has `points: [[x,y],...]` and `closed: true/false`), `fill` and `stroke` (`#RRGGBB` or `none`), `lineWidth`, optional `dash: true` and rectangular `clip: [xmin,ymin,xmax,ymax]`. Straight polylines and convex closed polygons are supported. Filled shapes with holes and concave clipped polygons are outside this helper's scope; use an appropriate authoring backend.
- `image`: `name`, `file` (PNG/JPEG relative to the scene JSON's directory), `x`, `y`, `width`, `height`. Images must already be cropped and oriented correctly; shape clipping does not apply to them. No remote images or external relationships are created. Preserve aspect ratio in the supplied dimensions.

Unknown fields are rejected, so an unsupported property cannot silently disappear. For example, opacity/gradient fills, arrows and Béziers are not understood: model supported geometry explicitly or use a more capable backend. Colors must be opaque 6-digit hex values.

The exporter refuses to overwrite an existing PPTX. Use a fresh build path. PNG/TIFF conversion is not bundled because a correct raster export requires a rendering engine (PowerPoint or another verified renderer), not merely a file extension change.
