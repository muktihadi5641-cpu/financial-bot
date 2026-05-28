"""One-shot: generate PWA icons for Financial Dashboard."""
from PIL import Image, ImageDraw
import os

OUT = os.path.dirname(os.path.abspath(__file__))

def draw_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d   = ImageDraw.Draw(img)

    r = size // 5
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=(15, 38, 102))

    # Three rising bars (bar chart)
    margin  = int(size * 0.18)
    gap     = int(size * 0.05)
    bw      = (size - 2 * margin - 2 * gap) // 3
    bottom  = int(size * 0.76)
    heights = [int(size * 0.28), int(size * 0.42), int(size * 0.56)]
    bar_r   = max(3, bw // 6)

    for i, h in enumerate(heights):
        x0 = margin + i * (bw + gap)
        x1 = x0 + bw
        y0 = bottom - h
        y1 = bottom
        d.rounded_rectangle([x0, y0, x1, y1], radius=bar_r, fill=(255, 255, 255))

    # Small trend line
    pts = []
    for i, h in enumerate(heights):
        cx = margin + i * (bw + gap) + bw // 2
        cy = bottom - h - int(size * 0.055)
        pts.append((cx, cy))
    lw = max(2, size // 64)
    for i in range(len(pts) - 1):
        d.line([pts[i], pts[i + 1]], fill=(96, 165, 250), width=lw)
    for pt in pts:
        dot = lw + 1
        d.ellipse([pt[0] - dot, pt[1] - dot, pt[0] + dot, pt[1] + dot], fill=(96, 165, 250))

    return img.convert("RGB")

for sz, name in [(512, "icon-512.png"), (192, "icon-192.png"), (180, "apple-touch-icon.png")]:
    img = draw_icon(sz)
    img.save(os.path.join(OUT, name))
    print(f"  {name}  ({sz}x{sz})")

print("Icons created.")
