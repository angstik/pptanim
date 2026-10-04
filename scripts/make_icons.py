"""Régénère les icônes de src/icons (à lancer seulement si le dessin change).
Dépendance : Pillow. Police : Poppins Bold, chemin à adapter selon la machine."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parent.parent / "src" / "icons"
FONT = "/usr/share/fonts/truetype/google-fonts/Poppins-Bold.ttf"
ACCENT, WHITE, PALE = (181, 71, 31, 255), (255, 255, 255, 255), (255, 225, 205, 255)


def draw(size, maskable=False):
    s = size * 4  # suréchantillonnage
    im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    if maskable:  # fond plein, motif dans la zone sûre (80 % au centre)
        d.rectangle([0, 0, s, s], fill=ACCENT)
        k = 0.36
    else:
        d.rounded_rectangle([0, 0, s - 1, s - 1], radius=s // 5, fill=ACCENT)
        k = 0.5
    f = ImageFont.truetype(FONT, int(s * k))
    d.text((s * (0.5 - 0.4 * k), s * 0.52), "A", font=f, fill=WHITE, anchor="mm")
    d.text((s * (0.5 + 0.36 * k), s * (0.52 - 0.16 * k)), "a", font=f, fill=PALE, anchor="mm")
    return im.resize((size, size), Image.LANCZOS)


OUT.mkdir(parents=True, exist_ok=True)
for n in (16, 32, 64, 80, 192, 512):
    draw(n).save(OUT / f"icon-{n}.png")
draw(32).save(OUT / "favicon-32.png")
draw(512, maskable=True).save(OUT / "icon-maskable-512.png")
# iOS n'applique pas de transparence : fond plein, sans coins arrondis.
draw(180, maskable=True).save(OUT / "apple-touch-icon-180.png")
print(sorted(p.name for p in OUT.iterdir()))
