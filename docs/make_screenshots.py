"""Regenerates the README images from examples/out/*.pdf (run examples/showcase.py first)."""
import io
import sys
from pathlib import Path

import fitz
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from PIL import Image  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from tairdoc import mpl_style  # noqa: E402

PDF = ROOT / "examples" / "out"
SHOTS = ROOT / "docs" / "screenshots"
SHOTS.mkdir(parents=True, exist_ok=True)
PRESETS = ("article", "journal", "essay")


def page(preset, n, dpi):
    doc = fitz.open(str(PDF / f"specimen_{preset}.pdf"))
    return Image.open(io.BytesIO(doc[n].get_pixmap(dpi=dpi).tobytes("png"))).convert("RGB")


def crop(preset, n, needle, above_cm, below_cm, dpi=220, name=None):
    """Crop a horizontal band of a page around the first match of `needle`."""
    doc = fitz.open(str(PDF / f"specimen_{preset}.pdf"))
    pg = doc[n]
    hit = pg.search_for(needle)[0]
    cm = 72 / 2.54
    clip = fitz.Rect(0, max(0, hit.y0 - above_cm * cm), pg.rect.width, min(pg.rect.height, hit.y0 + below_cm * cm))
    pix = pg.get_pixmap(dpi=dpi, clip=clip)
    pix.save(str(SHOTS / name))


# full pages
for p in PRESETS:
    doc = fitz.open(str(PDF / f"specimen_{p}.pdf"))
    for i in range(len(doc)):
        page(p, i, 110).save(SHOTS / f"{p}_p{i + 1}.png", optimize=True)

# hero: first pages side by side
mpl_style("article")
fig, axes = plt.subplots(1, 3, figsize=(15, 7.6))
for ax, p in zip(axes, PRESETS):
    ax.imshow(page(p, 0, 110))
    ax.set_xticks([]); ax.set_yticks([])
    for s in ax.spines.values():
        s.set_visible(True); s.set_color("#c8c8c8"); s.set_linewidth(0.8)
    ax.set_xlabel(p, fontsize=15, style="italic", labelpad=10)
fig.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.07, wspace=0.06)
fig.savefig(ROOT / "docs" / "hero.png", dpi=130, facecolor="white")
plt.close(fig)

# details
crop("article", 1, "What the presets do about it", 0.35, 11.0, name="detail_table_equation.png")
crop("essay", 0, "Why the line length matters", 0.6, 7.5, name="detail_sidenotes.png")
crop("journal", 0, "On the Length", 1.6, 9.0, name="detail_journal_title.png")
print("written:", sorted(x.name for x in SHOTS.iterdir()))
