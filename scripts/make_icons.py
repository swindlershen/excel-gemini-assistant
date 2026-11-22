"""Generates add-in icons (gradient rounded square with a 4-point sparkle)."""
from PIL import Image, ImageDraw
import os

OUT = os.path.join(os.path.dirname(__file__), "..", "public", "assets")
os.makedirs(OUT, exist_ok=True)


def make(size: int) -> Image.Image:
    s = size * 8  # supersample
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    grad = Image.new("RGBA", (s, s))
    px = grad.load()
    c1, c2 = (26, 115, 232), (142, 68, 239)
    for y in range(s):
        for x in range(s):
            t = (x + y) / (2 * s)
            px[x, y] = tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3)) + (255,)
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, s - 1, s - 1], radius=int(s * 0.22), fill=255)
    img.paste(grad, (0, 0), mask)
    d = ImageDraw.Draw(img)
    c, r, w = s / 2, s * 0.34, s * 0.075
    pts = [(c, c - r), (c + w, c - w), (c + r, c), (c + w, c + w), (c, c + r), (c - w, c + w), (c - r, c), (c - w, c - w)]
    d.polygon(pts, fill=(255, 255, 255, 255))
    return img.resize((size, size), Image.LANCZOS)


for size in (16, 32, 64, 80, 128):
    make(size).save(os.path.join(OUT, f"icon-{size}.png"))
print("icons written to", os.path.abspath(OUT))
