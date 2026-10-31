"""Compare actual browser output to independent LibreOffice reference geometry.

This checks colored object extents, not cross-renderer text antialiasing. The
separate Playwright snapshots cover same-browser pixels with fixed fonts.
No baseline is generated or rewritten by this check.
"""
from pathlib import Path
import argparse
import hashlib
import json
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent


def bounds(image, box, predicate):
    points = [(x, y) for y in range(box[1], min(box[3], image.height))
              for x in range(box[0], min(box[2], image.width)) if predicate(image.getpixel((x, y))) ]
    if not points:
        raise ValueError("An expected visible object has no pixels")
    return [min(x for x, y in points), min(y for x, y in points), max(x for x, y in points), max(y for x, y in points)]


def main(browser):
    result = {"browser": browser, "reference": "Independent LibreOffice source PDF raster, examples/pptx/reference.json", "checks": [], "inputs": {}}
    def pair(name, page):
        native = ROOT / f"examples/pptx/reference/{name}-{page}.png"
        actual = ROOT / f".cache/presentation-preview/{browser}/source-{name}.pptx-{page}.png"
        for path in [native, actual]:
            result["inputs"][str(path.relative_to(ROOT))] = hashlib.sha256(path.read_bytes()).hexdigest()
        # Native PDF raster has a one-pixel quantization rim (1281 x 721).
        # Compare the shared 1280 x 720 viewport without resampling text or colors.
        return Image.open(native).convert("RGB"), Image.open(actual).convert("RGB")
    def check(label, images, box, predicate, tolerance):
        expected, actual = [bounds(image, box, predicate) for image in images]
        difference = max(abs(a - b) for a, b in zip(expected, actual))
        result["checks"].append({"label": label, "referenceBounds": expected, "browserBounds": actual, "maxDifferencePixels": difference, "tolerance": tolerance})
        if difference > tolerance:
            raise ValueError(f"{label}: geometry differs by {difference}px: {expected} vs {actual}")
    shapes = pair("04-shapes", 1)
    nonwhite = lambda c: min(c) < 235
    for index in range(16):
        x, y = round((40 + index % 4 * 225) * 4 / 3), round((30 + index // 4 * 125) * 4 / 3)
        check(f"preset-{index + 1}", shapes, (x - 2, y - 2, x + 264, y + 132), nonwhite, 4)
    blue = lambda c: max(abs(c[i] - [68, 114, 196][i]) for i in range(3)) < 8
    column = pair("09-charts", 1)
    for index in range(3):
        check(f"column-{index + 1}", column, (160 + index * 318, 250, 360 + index * 318, 550), blue, 3)
    check("line-plot", pair("09-charts", 2), (150, 250, 1000, 540), blue, 3)
    check("pie-plot", pair("09-charts", 3), (300, 240, 850, 580), blue, 3)
    target = ROOT / f".cache/presentation-preview/{browser}-reference-result.json"
    target.write_text(json.dumps(result, indent=2) + "\n")
    print(f"Verified {len(result['checks'])} independent native/browser object extents ({browser}); result: {target.relative_to(ROOT)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--browser", default="chromium", choices=["chromium", "firefox", "webkit"])
    main(parser.parse_args().browser)
