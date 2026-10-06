"""Fabrique les fichiers d'essai : deux présentations (sans et avec animation existante)
et l'image PNG transparente de la zone de texte, rendue par LibreOffice."""
import subprocess, sys, copy, zipfile
from pathlib import Path
from lxml import etree
from PIL import Image, ImageDraw, ImageFilter
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

OUT = Path(__file__).parent / "fixtures"
OUT.mkdir(exist_ok=True)
DPI = 288  # 4 px par point
TEXT = ['Bonjour à tous,', 'voici le "scrambler" d\'été !']
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def background():
    w, h = 1600, 900
    img = Image.new("RGB", (w, h))
    px = img.load()
    for y in range(h):
        for x in range(w):
            px[x, y] = (20 + 60 * x // w, 40 + 70 * y // h, 90 + 80 * (x + y) // (w + h))
    d = ImageDraw.Draw(img)
    for i in range(14):
        d.ellipse([i * 130 - 80, 500 - (i % 4) * 110, i * 130 + 240, 820 - (i % 4) * 110], fill=(30 + i * 9, 120 + i * 5, 150))
    img = img.filter(ImageFilter.GaussianBlur(18))
    path = OUT / "fond.jpg"
    img.save(path, quality=88)
    return path


def text_box(slide, color):
    tb = slide.shapes.add_textbox(Inches(1.2), Inches(2.3), Inches(11), Inches(3.4))
    tb.name = "Titre"
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(TEXT):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = line
        r.font.size = Pt(54)
        r.font.bold = True
        r.font.name = "DejaVu Sans"
        r.font.color.rgb = color
    return tb


def deck(with_bg=True, color=RGBColor(255, 255, 255)):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    if with_bg:
        s.shapes.add_picture(str(background()), 0, 0, prs.slide_width, prs.slide_height)
    tb = text_box(s, color)
    return prs, s, tb


EXISTING = """<p:timing xmlns:p="{P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{ID}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="500"/><p:tgtEl><p:spTgt spid="{ID}"/></p:tgtEl></p:cBhvr></p:animEffect>
</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>
</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="{ID}" grpId="0" animBg="1"/></p:bldLst></p:timing>"""


def render(pptx, dpi=DPI, page=1):
    """Rend une diapo (la première par défaut) en PNG via LibreOffice puis pdftoppm."""
    pptx = Path(pptx)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(pptx.parent), str(pptx)],
                   check=True, capture_output=True, timeout=180)
    pdf = pptx.with_suffix(".pdf")
    base = pptx.with_suffix("")
    base = base.with_name(f"{base.name}-p{page}") if page != 1 else base
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", "-f", str(page), "-l", str(page), "-singlefile", str(pdf), str(base)], check=True)
    return base.with_suffix(".png")


def main():
    # A : fond image + zone de texte, aucune animation
    prs, s, tb = deck()
    prs.save(OUT / "a.pptx")

    # B : la même avec un bloc déjà animé (apparition en fondu au clic)
    prs, s, tb = deck()
    rect = s.shapes.add_shape(1, Inches(1.2), Inches(5.4), Inches(3), Inches(1))
    rect.name = "Bloc"
    xml = EXISTING.replace("{P}", P).replace("{ID}", str(rect.shape_id))
    s._element.append(etree.fromstring(xml))
    prs.save(OUT / "b.pptx")

    # C : plusieurs diapos. La zone à animer est sur la 4e diapo affichée mais dans le fichier slide3.xml
    # (l'ordre d'affichage diffère de la numérotation) ; une autre diapo porte une zone de même nom.
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    def plain(lines, name="Texte"):
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        for i, line in enumerate(lines):
            tb = sl.shapes.add_textbox(Inches(1), Inches(1 + 1.6 * i), Inches(9), Inches(1.2))
            tb.name = name if i == 0 else f"{name} {i + 1}"
            r = tb.text_frame.paragraphs[0].add_run()
            r.text = line
            r.font.size = Pt(40)
            r.font.name = "DejaVu Sans"
        return sl
    plain(["Sommaire", "Trois parties, une conclusion"])
    plain(["Bonjour à tous, voici un autre texte", "Deuxième zone de la diapo"], name="Titre")
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(str(background()), 0, 0, prs.slide_width, prs.slide_height)
    plain_note = s.shapes.add_textbox(Inches(1.2), Inches(0.6), Inches(8), Inches(0.8))
    plain_note.name = "Surtitre"
    r = plain_note.text_frame.paragraphs[0].add_run(); r.text = "Chapitre 2"; r.font.size = Pt(24); r.font.name = "DejaVu Sans"
    text_box(s, RGBColor(255, 255, 255))
    rect = s.shapes.add_shape(1, Inches(1.2), Inches(5.4), Inches(3), Inches(1))
    rect.name = "Bloc"
    s._element.append(etree.fromstring(EXISTING.replace("{P}", P).replace("{ID}", str(rect.shape_id))))
    plain(["Merci !"])
    prs.save(OUT / "c.pptx")
    # python-pptx renumérote les fichiers à l'enregistrement : on déplace donc la diapo dans le XML lui-même.
    # Ordre affiché : slide1, slide2, slide4 (« Merci ! »), slide3 (la cible).
    src = zipfile.ZipFile(OUT / "c.pptx")
    parts = {i.filename: src.read(i.filename) for i in src.infolist()}
    src.close()
    pres = etree.fromstring(parts["ppt/presentation.xml"])
    lst = pres.find(f"{{{P}}}sldIdLst")
    third = lst[2]
    lst.remove(third)
    lst.append(third)
    parts["ppt/presentation.xml"] = etree.tostring(pres, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(OUT / "c.pptx", "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in parts.items():
            out.writestr(name, data)

    # Image de la forme : le texte seul, noir sur blanc, rendu par LibreOffice, puis blanc + alpha
    prs, s, tb = deck(with_bg=False, color=RGBColor(0, 0, 0))
    prs.save(OUT / "_texte.pptx")
    page = Image.open(render(OUT / "_texte.pptx")).convert("L")
    k = DPI / 914400
    box = [round(v * k) for v in (tb.left, tb.top, tb.left + tb.width, tb.top + tb.height)]
    gray = page.crop(box)
    alpha = gray.point(lambda v: 255 - v)
    shape = Image.new("RGBA", gray.size, (255, 255, 255, 0))
    shape.putalpha(alpha)
    shape.save(OUT / "titre.png")

    # Une image de texte qui ne correspond à aucune zone de la présentation
    from PIL import ImageFont
    other = Image.new("RGBA", (1400, 300), (0, 0, 0, 0))
    ImageDraw.Draw(other).text((30, 40), "Rien à voir ici", font=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 150), fill=(255, 255, 255, 255))
    other.save(OUT / "autre.png")
    print("fixtures:", sorted(p.name for p in OUT.iterdir()), "image", shape.size)


if __name__ == "__main__":
    main()
