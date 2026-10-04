# PPT Anim

Application web installable (PWA) qui remplace une zone de texte PowerPoint par des images de lettres
animées avec de **vraies animations PowerPoint**. Elles s'ajoutent à la chronologie de la diapo, à côté
des animations déjà présentes, et se règlent ensuite dans le volet Animations.

Application publiée : <https://angstik.github.io/pptanim/>

Effet disponible : le scrambler (ordonner ou mélanger), par lettre, par mot ou par ligne, dans l'ordre ou
aléatoire. Tout le traitement se fait dans le navigateur : aucun fichier n'est envoyé sur Internet.

## Utilisation

1. Dans PowerPoint : clic droit sur la zone de texte, « Enregistrer en tant qu'image », format PNG.
2. Dans l'application : choisir le .pptx et le PNG (ou les déposer sur la page, ou coller l'image avec
   Ctrl+V), puis la diapo et la zone de texte, « Analyser ».
3. Régler l'effet, vérifier l'aperçu, « Télécharger le .pptx animé ».
4. Ouvrir le fichier obtenu dans PowerPoint et lancer le diaporama.

Le bouton « Aide » de l'application reprend ces étapes avec un schéma pour chacune.

L'application s'installe depuis le navigateur (bouton « Installer » ou menu du navigateur) et fonctionne
ensuite hors ligne.

## Limites connues

- Texte sur fond transparent uniquement (pas de remplissage, d'ombre ni de halo sur la zone de texte).
- Zone de texte hors groupe, non pivotée, avec une position propre (pas un espace réservé jamais déplacé).
- Lettres liées ou qui se touchent (cursives, certaines paires en italique) : le mot reste entier.
- Écriture de gauche à droite.
- Démarrage : « au clic » pour la première lettre, « avec la précédente » pour les autres ; le reste se
  règle dans le volet Animations (sélectionner toutes les lignes « PPTAnim » pour les déplacer).
- Le texte devient des images : la zone de texte d'origine est conservée, masquée, pour pouvoir recommencer.
- Le collage direct d'une zone de texte copiée depuis PowerPoint dépend de ce que PowerPoint place dans le
  presse-papiers (image PNG à fond transparent attendue) ; ce point n'a pas été vérifié avec PowerPoint.

## Développement

Aucune dépendance pour construire :

```
node scripts/build.mjs      # src/ -> dist/ (version, service worker, contrôles de chemins)
node scripts/serve.mjs      # http://localhost:8080/pptanim/ (même chemin que GitHub Pages)
```

Le build échoue si un fichier suppose une publication à la racine du domaine, si le manifeste ne
correspond pas à `/pptanim/` ou si la liste de cache du service worker est incohérente.

| Dossier | Contenu |
|---|---|
| `src/` | l'application : `index.html`, `app.js`, `pwa.js`, `core/`, `icons/`, `manifest.webmanifest`, JSZip |
| `scripts/` | `build.mjs`, `serve.mjs`, modèle du service worker, générateur d'icônes |
| `office/` | manifeste du complément PowerPoint (facultatif) |
| `test/` | essais de bout en bout |
| `.github/workflows/pages.yml` | publication sur GitHub Pages à chaque push sur `main` |

### Publication

Le workflow construit `dist/` et le publie avec les actions officielles GitHub Pages. Dans les réglages
du dépôt, la source de Pages doit être « GitHub Actions » (Settings > Pages > Build and deployment).

Le service worker met en cache tous les fichiers de l'application sous un nom lié à leur contenu : toute
modification publiée crée un nouveau cache, et l'application propose alors de recharger.

### Essais

```
npm install && npx playwright install chromium
python3 test/make_fixtures.py   # python-pptx, Pillow, lxml, LibreOffice, pdftoppm
node scripts/build.mjs
node test/pwa.mjs               # manifeste, service worker, installabilité, hors ligne, chemins
node test/e2e.mjs               # génération de pptx dans Chromium
python3 test/check.py           # structure et rendu des pptx (PML_XSD=chemin/pml.xsd pour le schéma)
```

## Complément PowerPoint (facultatif)

La même application sert de volet dans PowerPoint via `taskpane.html` : elle part de la zone de texte
sélectionnée et insère la diapo animée après l'originale. Il faut Microsoft 365 version 2601 ou plus.
Charger `office/manifest.xml` dans PowerPoint, par exemple avec
`npx office-addin-debugging start office/manifest.xml desktop --app powerpoint`.
Ce mode n'a tourné qu'avec une API Office simulée ; il reste à valider dans PowerPoint.

L'historique des versions est dans [CHANGE.log](CHANGE.log).
