"""Draw the pads-eco logo (one press → component swap) and export PNG + ICO.

python tool/make_icon.py   → tool/pads-eco.png (1024) and tool/pads-eco.ico (16..256)
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

PAPER = "#EEF1F5"
INK = "#141C26"
COPPER = "#B5652B"
PASS_C = "#237A4C"
HERE = Path(__file__).resolve().parent


def draw(size: int, rounded_bg: bool = True, wire=INK) -> Image.Image:
    """The check mark *is* the schematic net: a wire bent into a tick, pin dots at
    both ends, and the swapped component (copper capacitor) sitting on the long stroke."""
    import math

    S = size
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    u = S / 100  # design units

    if rounded_bg:
        d.rounded_rectangle((0, 0, S - 1, S - 1), radius=int(22 * u), fill=PAPER)

    w = max(2, round(6 * u))  # wire stroke

    # tick geometry
    p1 = (18 * u, 56 * u)
    p2 = (42 * u, 80 * u)
    p3 = (84 * u, 26 * u)

    def dot(p, r, color):
        d.ellipse((p[0] - r, p[1] - r, p[0] + r, p[1] + r), fill=color)

    # short stroke
    d.line((*p1, *p2), fill=wire, width=w)
    dot(p2, w / 2, wire)  # round joint

    # long stroke, broken for the capacitor
    dx, dy = p3[0] - p2[0], p3[1] - p2[1]
    L = math.hypot(dx, dy)
    ux, uy = dx / L, dy / L          # along the stroke
    nx, ny = -uy, ux                 # perpendicular
    t = 0.56                         # component position along the stroke
    mx, my = p2[0] + dx * t, p2[1] + dy * t
    gap = 4.5 * u                    # half distance between plates
    ph = 12 * u                      # plate half-length
    a = (mx - ux * gap, my - uy * gap)
    b = (mx + ux * gap, my + uy * gap)
    d.line((*p2, a[0] - ux * 1.5 * u, a[1] - uy * 1.5 * u), fill=wire, width=w)
    d.line((b[0] + ux * 1.5 * u, b[1] + uy * 1.5 * u, *p3), fill=wire, width=w)

    # capacitor plates (copper), perpendicular to the stroke
    pw = max(2, round(5.5 * u))
    for c in (a, b):
        d.line((c[0] - nx * ph, c[1] - ny * ph, c[0] + nx * ph, c[1] + ny * ph), fill=COPPER, width=pw)
        dot((c[0] - nx * ph, c[1] - ny * ph), pw / 2, COPPER)
        dot((c[0] + nx * ph, c[1] + ny * ph), pw / 2, COPPER)

    # pin ends: schematic-style terminal dots
    dot(p1, w * 0.9, wire)
    dot(p3, w * 0.9, wire)

    return img


def main() -> None:
    big = draw(1024)
    big.save(HERE / "pads-eco.png")
    sizes = [256, 128, 64, 48, 32, 16]
    frames = [draw(s) for s in sizes]
    frames[0].save(HERE / "pads-eco.ico", format="ICO", sizes=[(s, s) for s in sizes], append_images=frames[1:])
    print("wrote", HERE / "pads-eco.png", HERE / "pads-eco.ico")


if __name__ == "__main__":
    main()
