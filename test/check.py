"""Contrôle les pptx produits par e2e.mjs : structure du paquet, chronologie, rendu avant/après."""
import os, sys, zipfile
from pathlib import Path
from lxml import etree
from PIL import Image, ImageChops, ImageStat
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).parent))
from make_fixtures import render

HERE = Path(__file__).parent
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ns = {"p": P, "a": "http://schemas.openxmlformats.org/drawingml/2006/main", "r": R}
fails = []
SCHEMA = etree.XMLSchema(etree.parse(os.environ["PML_XSD"])) if os.environ.get("PML_XSD") else None


def check(cond, msg):
    if not cond:
        fails.append(msg)
    return cond


# Durée totale attendue par cas (secondes), telle que réglée dans e2e.mjs
TOTALS = {"b-emoji-image": 6.0, "b-emoji-separe": 6.0, "b-emoji-unique": 6.0, "a-emoji-aleatoire": 4.5}
EXPECTED_LINES = {"b-emoji-unique": 1}


def emoji_checks(name, sld, new, mine, rel, ns):
    """Contrôles propres à l'effet emoji : images rangées hors diapo, fichiers partagés, durée totale exacte."""
    stem = name.replace(".pptx", "")
    parked = [p for p in new if int(p.find("p:spPr/a:xfrm/a:off", ns).get("x")) < 0]
    letters = [p for p in new if p not in parked]
    check(len(letters) == 37, f"{name}: {len(letters)} images de lettres, 37 attendues")
    for p in parked:
        off, ext = p.find("p:spPr/a:xfrm/a:off", ns), p.find("p:spPr/a:xfrm/a:ext", ns)
        check(int(off.get("x")) + int(ext.get("cx")) < 0, f"{name}: emoji rangé mais encore visible sur la diapo")
        check(abs(int(ext.get("cx")) - int(ext.get("cy"))) < 0.01 * int(ext.get("cx")), f"{name}: emoji non carré")
    files = {rel[p.find("p:blipFill/a:blip", ns).get(f"{{{R}}}embed")] for p in parked}
    check(len(files) <= 9 and len(files) < len(parked), f"{name}: {len(files)} fichiers pour {len(parked)} emoji (partage attendu)")
    # Même taille d'emoji sur une même ligne
    by_line = {}
    for p in parked:
        by_line.setdefault(p.find("p:spPr/a:xfrm/a:off", ns).get("y"), set()).add(p.find("p:spPr/a:xfrm/a:ext", ns).get("cx"))
    check(all(len(v) == 1 for v in by_line.values()), f"{name}: tailles d'emoji différentes sur une même ligne")

    # Fin de la dernière animation = durée totale demandée
    def delay(ctn):
        c = ctn.find("p:stCondLst/p:cond", ns)
        return int(c.get("delay")) if c is not None and c.get("delay", "").isdigit() else 0
    end = 0
    for e in mine:
        for b in e.xpath("./p:childTnLst//p:cBhvr/p:cTn", namespaces=ns):
            end = max(end, delay(e) + delay(b) + int(b.get("dur")))
    total = TOTALS[stem] * 1000
    check(abs(end - total) <= 2, f"{name}: fin à {end} ms, durée totale demandée {total:.0f} ms")
    if stem in EXPECTED_LINES:
        check(len(mine) == EXPECTED_LINES[stem], f"{name}: {len(mine)} animations, {EXPECTED_LINES[stem]} attendue(s)")
    # Chaque lettre apparaît une fois ; chaque emoji apparaît puis disparaît
    vis = {}
    for st in sld.xpath(".//p:set[p:cBhvr/p:attrNameLst/p:attrName='style.visibility']", namespaces=ns):
        spid = st.find("p:cBhvr/p:tgtEl/p:spTgt", ns).get("spid")
        vis.setdefault(spid, []).append(st.find("p:to/p:strVal", ns).get("val"))
    ids = lambda pics: [p.find("p:nvPicPr/p:cNvPr", ns).get("id") for p in pics]
    check(all(vis.get(i) == ["visible"] for i in ids(letters)), f"{name}: une lettre n'apparaît pas exactement une fois")
    check(all(sorted(vis.get(i, [])) == ["hidden", "visible"] for i in ids(parked)), f"{name}: un emoji n'apparaît pas puis ne disparaît pas")
    return f"{len(parked)} emoji dans {len(files)} fichiers, fin à {end} ms"


def one(path, source, before_png):
    name = path.name
    z = zipfile.ZipFile(path)
    check(z.testzip() is None, f"{name}: zip corrompu")
    sld = etree.fromstring(z.read("ppt/slides/slide1.xml"))
    rels = etree.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels"))
    src = etree.fromstring(zipfile.ZipFile(source).read("ppt/slides/slide1.xml"))

    # Ordre des enfants de la diapo : la chronologie vient avant extLst
    tags = [etree.QName(c).localname for c in sld]
    check(tags.index("timing") > tags.index("cSld") and (("extLst" not in tags) or tags.index("timing") < tags.index("extLst")),
          f"{name}: place de p:timing {tags}")
    check(tags.count("timing") == 1, f"{name}: {tags.count('timing')} chronologies")

    # Identifiants
    ids = [e.get("id") for e in sld.iter(f"{{{P}}}cNvPr")]
    check(len(ids) == len(set(ids)), f"{name}: identifiants de formes en double")
    ctn = [e.get("id") for e in sld.iter(f"{{{P}}}cTn")]
    check(len(ctn) == len(set(ctn)), f"{name}: identifiants cTn en double")
    targets = {e.get("spid") for e in sld.iter(f"{{{P}}}spTgt")}
    check(targets <= set(ids), f"{name}: animations visant des formes absentes {targets - set(ids)}")

    # Images et relations
    pics = sld.findall(".//p:pic", ns)
    new = [p for p in pics if p.find("p:nvPicPr/p:cNvPr", ns).get("name").startswith("PPTAnim ")]
    rel = {r.get("Id"): r.get("Target") for r in rels}
    check(len(rel) == len(list(rels)), f"{name}: relations en double")
    for p in new:
        rid = p.find("p:blipFill/a:blip", ns).get(f"{{{R}}}embed")
        t = rel.get(rid, "")
        check(("ppt/" + t.replace("../", "")) in z.namelist(), f"{name}: image manquante pour {rid}")
    ct = z.read("[Content_Types].xml").decode()
    check('Extension="png"' in ct, f"{name}: type png non déclaré")

    # Chaque image est animée, un seul déclencheur au clic, la zone de texte d'origine est masquée
    new_ids = {p.find("p:nvPicPr/p:cNvPr", ns).get("id") for p in new}
    effects = sld.xpath(".//p:cTn[@presetClass]", namespaces=ns)
    mine = [e for e in effects if e.xpath(".//p:spTgt/@spid", namespaces=ns)[0] in new_ids]
    animated = {t for e in mine for t in e.xpath(".//p:spTgt/@spid", namespaces=ns)}
    check(animated == new_ids and len(new) > 0, f"{name}: {len(animated)} images animées sur {len(new)}")
    check(sum(e.get("nodeType") == "clickEffect" for e in mine) == 1, f"{name}: il faut un seul déclencheur au clic")
    emoji = "emoji" in name
    if not emoji:
        check(len(mine) == len(new), f"{name}: {len(mine)} animations pour {len(new)} images")
    else:
        extra = emoji_checks(name, sld, new, mine, rel, ns)
    titre = sld.xpath(".//p:sp/p:nvSpPr/p:cNvPr[@name='Titre']", namespaces=ns)[0]
    check(titre.get("hidden") == "1", f"{name}: forme d'origine non masquée")

    # Les animations déjà présentes sont intactes, dans le même ordre, avant les nôtres
    canon = lambda e: etree.tostring(e, method="c14n")
    old_src = src.xpath(".//p:cTn[@presetClass]", namespaces=ns)
    old_out = effects[: len(old_src)]
    check([canon(e) for e in old_src] == [canon(e) for e in old_out], f"{name}: animations existantes modifiées")
    check([canon(e) for e in src.xpath(".//p:bldLst", namespaces=ns)] == [canon(e) for e in sld.xpath(".//p:bldLst", namespaces=ns)],
          f"{name}: bldLst modifié")

    # Schéma PresentationML (facultatif : PML_XSD=chemin/vers/pml.xsd)
    if SCHEMA is not None:
        check(SCHEMA.validate(sld), f"{name}: schéma : {[str(e.message)[:140] for e in SCHEMA.error_log][:3]}")

    # python-pptx relit le fichier
    prs = Presentation(str(path))
    check(len(prs.slides) == 1, f"{name}: nombre de diapos")

    # Rendu LibreOffice : image finale identique à l'originale (lettres à leur place, texte d'origine masqué)
    after = Image.open(render(path, dpi=96)).convert("RGB")
    before = Image.open(before_png).convert("RGB")
    diff = ImageChops.difference(before, after)
    mean = sum(ImageStat.Stat(diff).mean) / 3
    worst = max(ImageStat.Stat(diff).extrema, key=lambda e: e[1])[1]
    diff.point(lambda v: min(255, v * 4)).save(path.with_name(path.stem + "-diff.png"))
    check(mean < 1.0, f"{name}: rendu différent de l'original (écart moyen {mean:.2f})")
    print(f"{name}: {len(new)} images, {len(mine)} animations + {len(old_src)} existantes, écart de rendu moyen {mean:.3f} (max {worst})"
          + (f" | {extra}" if emoji else ""))


def main():
    fx, out = HERE / "fixtures", HERE / "out"
    before = {d: render(fx / f"{d}.pptx", dpi=96) for d in "ab"}
    for path in sorted(out.glob("*.pptx")):
        d = path.name[0]
        one(path, fx / f"{d}.pptx", before[d])
    print("\nÉCHECS :\n" + "\n".join(fails) if fails else "\nTous les contrôles passent.")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
