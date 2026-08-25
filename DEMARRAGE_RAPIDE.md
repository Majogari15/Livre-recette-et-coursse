# Démarrage rapide

Ce dossier contient tout ce qu'il faut pour utiliser Mon Livre de Recettes.

## Option A — Utiliser directement avec Python (le plus simple pour tester)

Double-cliquez sur **`main.pyw`**. Ça fonctionne tout de suite, à condition
d'avoir Python installé sur votre PC (gratuit sur
https://www.python.org/downloads/, cocher "Add Python to PATH" pendant
l'installation).

Certaines fonctionnalités (photos, export PDF/Excel, QR code, import photo)
nécessitent en plus quelques modules Python. Ouvrez une invite de commandes
dans ce dossier et tapez :
```
pip install pillow reportlab openpyxl qrcode pytesseract
```

## Option B — Créer un fichier .exe autonome (pour un usage sans Python)

Si vous voulez un `.exe` qui fonctionne **sans avoir Python installé**
(pratique pour l'utiliser sur un autre PC, le partager, ou juste avoir une
icône à double-cliquer) :

1. Assurez-vous d'avoir Python installé sur CE PC, juste le temps de la
   construction (voir Option A) — une fois le `.exe` généré, il n'aura
   plus besoin de Python du tout, y compris sur d'autres PC.
2. Double-cliquez sur **`Construire_le_exe.bat`**.
3. Laissez faire — le script installe tout ce qu'il faut et construit
   l'exécutable automatiquement (1 à 2 minutes).
4. Une fois terminé, votre application se trouve dans le dossier `dist`,
   sous le nom **`MesRecettes.exe`**, accompagnée automatiquement de tous
   les fichiers nécessaires à son fonctionnement (`ingredients_par_defaut.json`,
   `valeurs_nutritionnelles.json`, `ingredient_allergenes.json` et
   `LISEZ-MOI.txt`).
5. Vous pouvez déplacer le dossier `dist` entier où vous voulez (clé USB,
   Bureau, autre PC...) — gardez tous les fichiers de ce dossier ensemble.
   Le fichier **`LISEZ-MOI.txt`** à l'intérieur explique tout ce qu'il faut
   savoir pour utiliser cette version `.exe` (sans aucune référence à
   Python, puisque vous n'en aurez plus besoin).

> Ce script doit être exécuté sur Windows (pas depuis ce chat) : téléchargez
> le dossier, puis lancez `Construire_le_exe.bat` sur votre propre PC.

## Contenu de ce dossier
- `main.pyw` : le code source de l'application (nécessite Python — voir
  Option A), se lance sans fenêtre noire de console
- `ingredients_par_defaut.json`, `valeurs_nutritionnelles.json`,
  `ingredient_allergenes.json` : les bases de données fournies (ingrédients
  courants, valeurs nutritionnelles, allergènes)
- `Construire_le_exe.bat` : script pour générer le `.exe` automatiquement
  (Option B)
- `LISEZ-MOI.md` : le guide complet d'utilisation de toutes les
  fonctionnalités, pour ceux qui utilisent `main.pyw` avec Python
- `LISEZ-MOI.txt` : le même type de guide mais **spécifique à la version
  `.exe`** (sans rien sur Python/pip) — ce fichier n'a d'utilité qu'une
  fois copié aux côtés de `MesRecettes.exe` après construction (le script
  de l'Option B s'en charge automatiquement)

Au premier lancement (quelle que soit l'option choisie), l'application
créera automatiquement à côté d'elle : `recipes.json`, `ingredients.json`,
`weekly_plan.json`, `menus.json`, `trash.json`, `settings.json`, un dossier
`images/` et un dossier `backups/`. Gardez tous ces fichiers ensemble.
