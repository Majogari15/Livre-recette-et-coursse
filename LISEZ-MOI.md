# Mon Livre de Recettes — Application de bureau

Application de bureau (Windows, macOS, Linux) pour gérer vos recettes de
cuisine, créée avec Python + Tkinter, avec une identité visuelle chaleureuse
(palette terracotta/crème) plutôt que le gris par défaut de Windows, et un
thème sombre disponible en un clic.

## Contenu du dossier
- `main.py` / `main.pyw` : le code de l'application (identiques, `.pyw` évite
  la fenêtre noire de la console)
- `ingredients_par_defaut.json` : liste d'environ 1000 ingrédients de cuisine
  courants fournie avec l'application (voir ci-dessous)
- `valeurs_nutritionnelles.json` : base de valeurs nutritionnelles estimées
  (kcal, protéines, glucides, lipides) fournie avec l'application, pour les
  ~1000 ingrédients courants — voir la section "Coût et valeurs
  nutritionnelles" plus bas
- `ingredient_allergenes.json` : base des allergènes présents dans les
  ~1000 ingrédients courants, fournie avec l'application (voir la section
  "Détection automatique des allergènes" plus bas)
- `recipes.json` : créé automatiquement dès que vous enregistrez votre
  première recette (votre base de données de recettes)
- `ingredients.json` : créé automatiquement, contient la liste des
  ingrédients réutilisables dans les menus déroulants
- `ingredient_custom_data.json` : créé automatiquement si vous modifiez les
  allergènes/valeurs nutritionnelles d'un ingrédient ou en créez un nouveau
  (voir "Gérer les ingrédients")
- `ingredient_prices.json` : créé automatiquement si vous renseignez des prix
  d'ingrédients (pour l'estimation du coût des recettes)
- `ingredient_dismissed_pairs.json` : créé automatiquement si vous indiquez
  qu'une paire d'ingrédients détectée comme doublon probable n'en est pas un
  (voir "🔤 Vérifier les doublons")
- `weekly_plan.json` : créé automatiquement, contient votre planning de la
  semaine
- `menus.json` : créé automatiquement, contient vos menus enregistrés
- `saved_shopping_lists.json` : créé automatiquement si vous enregistrez une
  liste de courses pour plus tard
- `trash.json` : créé automatiquement, contient les recettes envoyées à la
  corbeille
- `recent_views.json` : créé automatiquement, mémorise vos dernières
  recettes consultées pour la page d'accueil
- `backups/` : créé automatiquement, contient les sauvegardes automatiques
  périodiques (voir plus bas)
- `settings.json` : créé automatiquement, mémorise votre choix de thème
  (clair/sombre), l'acceptation de la clause de responsabilité, et le
  dossier de sauvegarde cloud si vous en configurez un
- `images/` : créé automatiquement, contient les photos de vos recettes et
  de votre journal de cuisine

> ⚠️ Important : gardez `ingredients_par_defaut.json`, `valeurs_nutritionnelles.json`
> **et** `ingredient_allergenes.json` dans le même dossier que `main.py` (et
> à côté du `.exe` si vous en générez un). Le premier permet de pré-remplir la
> liste des ~1000 ingrédients courants, le deuxième sert à l'estimation des
> valeurs nutritionnelles, le troisième à la détection automatique des
> allergènes.

## 1. Installer les dépendances

Il vous faut Python installé (gratuit sur https://www.python.org/downloads/,
cocher "Add Python to PATH" pendant l'installation).

Tkinter est déjà inclus avec Python. Il faut en plus installer quelques
modules pour les photos, l'export PDF/Excel, les QR codes et l'import depuis
une photo. Ouvrez une invite de commandes dans le dossier du projet et
tapez :

```
pip install pillow reportlab openpyxl qrcode pytesseract
```

- **pillow** : permet d'afficher les photos des recettes et les QR codes
- **reportlab** : permet d'exporter la liste de courses / une recette en PDF
- **openpyxl** : permet d'exporter la liste de courses en Excel (.xlsx)
- **qrcode** : permet d'exporter une recette sous forme de QR code
- **pytesseract** : permet d'importer une recette depuis une photo (OCR) —
  voir l'encadré ci-dessous, une étape supplémentaire est nécessaire

Si vous ne les installez pas, l'application fonctionne quand même, mais sans
les fonctionnalités correspondantes (un message vous le rappelle au
démarrage). L'export en texte (.txt) et l'import de recette depuis un lien
internet, eux, fonctionnent toujours sans dépendance supplémentaire.

> ⚠️ **Cas particulier de l'import depuis une photo** : contrairement aux
> autres modules ci-dessus, `pytesseract` ne fait que *piloter* un
> programme externe appelé **Tesseract OCR**, qui doit être installé
> séparément sur votre PC (ce n'est pas un simple `pip install`) :
> - **Windows** : téléchargez l'installateur depuis
>   https://github.com/UB-Mannheim/tesseract/wiki, installez-le (cochez la
>   langue française si proposée), puis redémarrez l'application.
> - **macOS** : `brew install tesseract tesseract-lang` (via Homebrew).
> - **Linux** : `sudo apt install tesseract-ocr tesseract-ocr-fra` (Debian/
>   Ubuntu) ou l'équivalent pour votre distribution.
>
> Sans Tesseract OCR installé, le bouton "📷 Importer une recette depuis une
> photo" reste accessible mais affiche un message clair vous renvoyant à ces
> instructions plutôt que de planter.

## 2. Lancer l'application

Double-cliquez sur `main.pyw` (recommandé, aucune fenêtre noire ne s'ouvre),
ou sur `main.py`. Vous pouvez aussi utiliser une invite de commandes :

```
python main.py
```

> 💡 Une fenêtre noire (l'invite de commandes) s'ouvre en arrière-plan quand
> vous lancez `main.py` avec Python ? C'est normal, mais évitable :
> - **Le plus simple** : utilisez `main.pyw` au lieu de `main.py` — même
>   contenu, mais Windows le lance sans console.
> - En ligne de commande, utilisez `pythonw main.py` au lieu de `python main.py`
>   (notez le "w" à la fin de "python").
> - Si vous générez un `.exe` avec PyInstaller (voir plus bas), l'option
>   `--windowed` déjà indiquée dans les instructions supprime aussi cette
>   fenêtre noire automatiquement.

> 💡 Plusieurs fenêtres de l'application (page d'accueil, "Ajouter une
> recette", "Toutes les recettes", "Voir une recette précise", "Planning de
> la semaine", "Nouveau menu") s'ouvrent à la hauteur de votre écran pour
> afficher un maximum de contenu sans avoir à agrandir la fenêtre à la
> main. Sous Windows, cette hauteur tient automatiquement compte de la
> barre des tâches (l'application ne se retrouve jamais partiellement
> masquée derrière). Un petit espace vide est aussi laissé en bas de
> chaque liste déroulante, pour pouvoir descendre l'ascenseur un peu plus
> bas que le dernier élément et le voir entièrement.

## 3. Comment ça marche

**⚠ Clause de responsabilité (au tout premier lancement)**
Avant de pouvoir utiliser l'application pour la première fois, une fenêtre
affiche une clause de responsabilité (notamment sur la gestion des
allergènes, qui reste une aide informative et ne remplace jamais une
vérification personnelle des étiquettes des produits). Cochez "J'ai lu et
j'accepte les conditions ci-dessus" pour activer le bouton "Continuer" et
accéder à l'application. Ce texte n'apparaît **qu'une seule fois** : votre
acceptation est mémorisée dans `settings.json`, les lancements suivants
vont directement à la page d'accueil.

**🏠 Page d'accueil**
En haut à gauche, le bouton **"☕ Faire un don"** ouvre dans votre navigateur
la page https://buymeacoffee.com/majogari, si vous souhaitez soutenir le
développement de l'application.

En haut à droite, le bouton **"🌙 Thème sombre" / "☀️ Thème clair"** bascule
l'apparence de toute l'application. Votre choix est mémorisé automatiquement
d'une utilisation à l'autre. La page d'accueil et toute nouvelle fenêtre que
vous ouvrez ensuite adoptent immédiatement le thème choisi ; si une fenêtre
secondaire était déjà ouverte au moment de la bascule, refermez-la et
rouvrez-la pour qu'elle reflète pleinement le nouveau thème.

Trois **filtres rapides** ("⭐ Favoris", "⏱️ Rapide (≤ 30 min)",
"🥗 Végétarien") ouvrent directement "Modifier / Supprimer une recette" avec
la liste déjà filtrée, sans avoir à passer par la recherche manuelle. Le
filtre "Végétarien" se base sur l'étiquette "végétarien"/"végétarienne" que
vous avez éventuellement ajoutée à vos recettes. Un bouton "✕ Retirer le
filtre" apparaît dans la fenêtre pour revenir à la liste complète.

En plus des boutons donnant accès à toutes les fonctionnalités, la page
d'accueil affiche deux sections pratiques :
- **"📅 Aujourd'hui"** : si vous avez rempli le planning de la semaine, les
  repas prévus pour aujourd'hui (petit-déjeuner, déjeuner, dîner) s'affichent
  directement ici, avec un bouton "👁" pour ouvrir chaque recette en un clic ;
- **"🕘 Récemment consultées"** : les 8 dernières recettes que vous avez
  affichées dans "Voir une recette précise" (la plus récente en haut, sans
  doublon). Sélectionnez-en une puis "👁 Ouvrir" (ou double-cliquez dessus)
  pour la rouvrir directement, sans avoir à la rechercher.

**🥕 Gérer les ingrédients**
Un cinquième bouton permet de gérer la liste des ingrédients réutilisables :
- une **barre de recherche** en haut filtre instantanément la liste affichée
  au fur et à mesure que vous tapez (recherche insensible aux accents et à la
  casse, sur n'importe quelle partie du nom) ;
- une **barre de défilement** à droite de la liste permet de parcourir
  facilement tous les ingrédients ;
- tapez un nom puis "➕ Ajouter", ou sélectionnez un ingrédient existant puis
  "✏️ Modifier" : dans les deux cas, la même fenêtre s'ouvre, où vous pouvez
  renseigner :
  - le **nom** (un renommage met à jour automatiquement toutes les recettes
    qui utilisent cet ingrédient) ;
  - les **allergènes présents** (cases à cocher) ;
  - les **valeurs nutritionnelles** pour 100 g/100 ml (calories, protéines,
    glucides, lipides) — laissez vide si vous ne les connaissez pas ;
  - le **prix** (voir "💰 Gérer les prix" plus bas).
  Ce que vous renseignez ici est prioritaire sur les bases fournies avec
  l'application (~1000 ingrédients) : par exemple, si vous corrigez la valeur
  calorique d'un ingrédient courant, c'est votre valeur qui sera utilisée
  partout. Le bouton "🗑️ Supprimer cet ingrédient" est aussi disponible
  directement dans cette fenêtre lors d'une modification ;
  > En créant un nouvel ingrédient, l'application vérifie automatiquement
  > qu'il ne s'agit pas simplement du singulier ou du pluriel d'un
  > ingrédient déjà existant (ex. tenter de créer "Tomates" alors que
  > "Tomate" existe déjà) et vous invite à utiliser directement l'ingrédient
  > existant plutôt que d'en créer un doublon — pour que chaque ingrédient
  > n'apparaisse toujours qu'une seule fois dans la liste ;
- "📚 Charger les ~1000 ingrédients courants" ajoute d'un coup tous les
  ingrédients de la liste fournie avec l'application qui ne sont pas déjà
  présents (aucun doublon, aucune suppression) ;
- "🔤 Vérifier les doublons / fautes de frappe" analyse toute votre liste
  d'ingrédients et détecte les paires qui se ressemblent à **90 % ou plus** —
  un pluriel non fusionné ("Tomate" / "Tomates"), une faute de frappe
  ("Echalotte" / "Échalote")... Chaque paire trouvée s'affiche avec son
  pourcentage de similarité. Vous pouvez sélectionner **plusieurs paires à la
  fois** (Ctrl+clic ou Maj+clic) :
  - pour **une seule paire**, "🔗 Fusionner la sélection" vous demande
    laquelle des deux graphies garder ;
  - pour **plusieurs paires**, la fusion se fait automatiquement pour
    chacune : l'ingrédient le moins utilisé dans vos recettes est fusionné
    vers celui utilisé dans le plus grand nombre de recettes (pratique pour
    nettoyer une longue liste sans traiter les paires une par une).
  Dans tous les cas, la fusion renomme l'ingrédient partout où il est
  utilisé dans vos recettes, comme un renommage classique.
  Si une paire proposée n'est **pas réellement un doublon** (deux
  ingrédients différents qui se ressemblent juste beaucoup), sélectionnez-la
  et cliquez sur **"✕ Ce n'est pas un doublon"** : elle disparaît de cette
  analyse, aujourd'hui et lors de toutes les analyses suivantes.

> Au tout premier lancement de l'application (avant toute création de
> recette), la liste des ~1000 ingrédients les plus utilisés en cuisine est
> automatiquement chargée, pour que les menus déroulants soient tout de suite
> bien fournis. Si vous avez déjà utilisé l'application avant cette mise à
> jour, utilisez simplement le bouton "📚 Charger les ~1000 ingrédients
> courants" pour les ajouter à votre liste existante.

**🔎 Recherche par ingrédient**
Cette fenêtre fait l'inverse d'une recherche classique : au lieu de chercher
une recette par son nom, elle trouve **toutes les recettes qui utilisent un
ingrédient donné**. Recherchez et sélectionnez l'ingrédient dans la liste
(ou double-cliquez dessus), puis "🔍 Voir les recettes qui l'utilisent" :
la liste des recettes concernées s'affiche, avec la quantité nécessaire pour
1 personne dans chacune — pratique pour savoir quoi cuisiner avec un
ingrédient qui traîne, ou pour repérer toutes les recettes à ajuster si un
ingrédient devient difficile à trouver. Sélectionnez une recette dans les
résultats puis "📖 Consulter la recette sélectionnée" (ou double-cliquez
dessus) pour l'ouvrir directement dans "Voir une recette précise".

**💰 Gérer les prix (dans "Gérer les ingrédients")**
Cette fenêtre permet de renseigner un prix pour les ingrédients qui vous
intéressent — inutile de tous les faire, seuls ceux avec un prix connu
compteront dans l'estimation. Recherchez un ingrédient, indiquez son prix et
son unité de référence (**kg**, **L**, **pièce**, **cuillère à soupe** ou
**cuillère à café**), puis "💾 Enregistrer le prix". "🗑 Effacer le prix"
retire le prix enregistré. Un prix au kg s'applique automatiquement aux
recettes utilisant des grammes (Gr), un prix au litre aux recettes utilisant
des centilitres (cl) ; les prix en pièce/cuillère s'appliquent tels quels.

> Il n'existe aucune source de prix fiable en ligne pour ce type
> d'application locale (les prix varient trop selon le magasin, la région,
> les promotions...) : c'est donc vous qui renseignez vos propres prix, ce
> qui reste la seule méthode qui donne une estimation réellement pertinente
> pour votre budget.

**🥗 Coût et valeurs nutritionnelles estimés**
Une fois des prix renseignés, le **coût estimé** d'une recette s'affiche
automatiquement dans "Voir une recette précise", dans les exports PDF
(recette individuelle et livre de recettes) et dans "Comparer deux
recettes". De la même façon, les **valeurs nutritionnelles estimées**
(calories, protéines, glucides, lipides) s'affichent partout, calculées à
partir de la base `valeurs_nutritionnelles.json` fournie avec l'application
(environ 1000 ingrédients) — sans rien à configurer de votre côté pour la
nutrition, contrairement au coût.

Dans les deux cas, si certains ingrédients de la recette n'ont pas de prix
renseigné (pour le coût) ou ne sont pas reconnus dans la base nutritionnelle
(par exemple un ingrédient que vous avez créé vous-même), une mention
"estimation partielle, X/Y ingrédients pris en compte" vous le signale
clairement, plutôt que d'afficher un chiffre trompeur. Les ingrédients
exprimés en **"pièce"** ou en unité **"autre"** personnalisée ne sont jamais
comptés dans les valeurs nutritionnelles (le poids d'"une pièce" varie trop
selon l'ingrédient pour être généralisé) — vous pouvez toujours renseigner un
prix pour un ingrédient à la pièce, cela fonctionne normalement pour le coût.

> Les valeurs nutritionnelles sont des **estimations générales par famille
> d'aliments** (ex. "poisson maigre", "fromage à pâte dure", "légume-feuille"),
> pas des mesures de laboratoire spécifiques à une marque ou une variété
> précise — à prendre comme un ordre de grandeur utile, pas une valeur
> médicale exacte.

**⚖️ Comparer deux recettes**
Choisissez une recette A et une recette B dans les deux menus déroulants,
puis "⚖️ Comparer" : un tableau côte à côte affiche leur catégorie, favori,
note, difficulté, temps de préparation/cuisson/total, nombre de fois
cuisinée, **coût estimé** et **calories estimées** (pour le nombre de
personnes par défaut de chacune), suivi de la liste des **ingrédients
communs aux deux recettes** et de ceux qui ne se trouvent que dans l'une ou
l'autre — pratique pour choisir entre deux variantes d'un même plat.

**🌐 Importer une recette depuis un lien**
Collez l'adresse (URL) d'une page de recette trouvée sur internet, puis
"🌐 Récupérer la recette" : l'application télécharge la page et tente d'en
extraire automatiquement le nom, les ingrédients (avec quantité et unité),
les étapes de préparation, les temps de préparation/cuisson, **et la photo
de la recette** (déjà téléchargée et prête dans la galerie), puis ouvre
directement le formulaire "Ajouter une recette" pré-rempli — il ne vous
reste qu'à vérifier et enregistrer.

Cela fonctionne avec les sites qui utilisent le format de données standard
« Schema.org Recipe », utilisé par la grande majorité des sites de cuisine
(mais pas tous). Une **connexion internet est nécessaire**. Le repérage des
quantités et unités dans le texte n'est pas toujours parfait selon la façon
dont le site rédige ses ingrédients ; relisez et corrigez si besoin avant
d'enregistrer. La photo n'est récupérée que si le site en indique une dans
ses données structurées ; si aucune photo n'est trouvée (ou si son
téléchargement échoue), la recette s'importe quand même, simplement sans
photo — vous pourrez toujours en ajouter une manuellement. Si aucune donnée
de recette n'est trouvée sur la page, un message vous l'indique et vous
pouvez créer la recette manuellement. Si un ingrédient importé n'est qu'une
variante singulier/pluriel d'un ingrédient déjà dans votre liste (ex. le
site utilise "Tomates" alors que vous avez déjà "Tomate"), l'application
réutilise automatiquement votre ingrédient existant au lieu d'en créer un
doublon, aussi bien dans la liste que dans la recette importée elle-même —
ce qui permet aussi de conserver la détection des allergènes et des valeurs
nutritionnelles pour cet ingrédient.

**📷 Importer une recette depuis une photo**
Prenez en photo (ou scannez) une recette manuscrite, une carte de recette ou
une page de livre de cuisine, puis choisissez cette image via
"📁 Choisir une photo". Cliquez sur "🔍 Extraire le texte" : le texte visible
sur la photo est reconnu automatiquement (reconnaissance optique de
caractères, OCR) et s'affiche dans une zone modifiable, à relire et corriger
avant de créer la recette avec "➡️ Créer la recette avec ce texte" — cela
ouvre le formulaire "Ajouter une recette" avec ce texte dans la description
et la photo déjà attachée.

> Contrairement à l'import depuis un lien, une photo n'a pas de structure
> ingrédients/étapes qu'on puisse deviner automatiquement : le texte extrait
> est un bloc brut, à vous de le relire et de répartir vous-même le nom, les
> ingrédients et les étapes dans le formulaire. C'est aussi plus rapide que
> de tout retaper à la main depuis une recette papier.

Cette fonctionnalité nécessite le module `pytesseract` **et** le programme
externe **Tesseract OCR** installé séparément sur votre PC (voir la section
"Installer les dépendances" en tout début de ce document pour les
instructions par système). Sans cela, le bouton "🔍 Extraire le texte" vous
indique clairement ce qu'il manque plutôt que de planter.

**➕ Ajouter une recette**
Le formulaire est organisé en deux colonnes pour limiter le défilement :
à gauche le nom, la catégorie, les temps/difficulté et les étiquettes, à
droite les allergènes ; plus bas, les ingrédients à gauche et la
description/notes personnelles à droite. La fenêtre s'ouvre à la hauteur de
votre écran pour afficher un maximum de contenu d'un coup.

Donnez un nom, cochez éventuellement "⭐ Marquer comme recette favorite" et
notez la recette avec **1 à 5 étoiles cliquables** ("Ma note" — cliquer sur
une étoile déjà sélectionnée remet la note à zéro). Choisissez une catégorie
(**Entrée**, **Plat**, **Dessert**, **Apéro**, **Boisson**, **Sauce** ou
**Autre**), et indiquez si vous le
souhaitez un **temps de préparation**, un **temps de cuisson** (en minutes),
une **difficulté** (Facile / Moyen / Difficile) et un **nombre de personnes
par défaut** (utilisé pour préremplir "Voir une recette précise" et le
planning de la semaine).

Le champ **Étiquettes** (séparées par des virgules, ex. "végétarien, sans
gluten, rapide, économique") permet d'ajouter vos propres mots-clés libres,
en plus de la catégorie fixe. Elles sont reconnues par toutes les barres de
recherche de l'application (taper une étiquette dans "🔍 Rechercher" retrouve
aussi les recettes qui la portent).

Les cases à cocher **Allergènes présents** (Gluten, Lactose, Œufs,
Arachides, Fruits à coque, Soja, Poisson, Crustacés, Sésame, Céleri,
Moutarde, Sulfites, Lupin, Mollusques — les 14 allergènes à déclaration
obligatoire en Europe) se cochent **et se décochent automatiquement** au fil
de la saisie des ingrédients, dans la section plus bas du formulaire : choisir
"Farine de blé" coche aussitôt "Gluten", et si vous remplacez ensuite cet
ingrédient par autre chose qui n'en contient pas, "Gluten" se décoche tout
seul. Cette synchronisation automatique **ne touche jamais** à une case que
vous auriez cochée vous-même sans qu'un ingrédient de la recette ne la
justifie (par exemple pour signaler un risque de contamination croisée) —
seuls les allergènes que la détection a elle-même cochés peuvent être
décochés par la suite. Les recettes importées depuis un lien internet (voir
plus haut) ont elles aussi leurs allergènes détectés automatiquement dès
l'ouverture du formulaire, sans action de votre part.

Le bouton **"🔍 Détecter automatiquement"**, à côté du titre "Allergènes
présents", relance cette même synchronisation sur l'ensemble des ingrédients
déjà saisis — pratique après avoir modifié plusieurs lignes d'un coup, ou
après avoir changé les allergènes d'un ingrédient dans "Gérer les
ingrédients". Juste en dessous, un texte en **rouge vif** rappelle que
cette détection n'est qu'indicative : vérifiez toujours les allergènes sur
les étiquettes des produits physiques avant de cuisiner pour quelqu'un
ayant une allergie.

Pour les photos, cliquez sur "📷 Ajouter une photo" autant de fois que
nécessaire : vous pouvez attacher **plusieurs photos** à une même recette
(sélection multiple possible dans la fenêtre de choix de fichier). Chaque
photo ajoutée apparaît dans une galerie avec un bouton "🗑 Retirer" pour
l'enlever avant d'enregistrer.

Une zone **Description** (jusqu'à 2056 caractères, avec un compteur affiché en
dessous) permet de noter les étapes de préparation, des astuces, ou toute
autre information utile sur la recette. Une seconde zone **Notes
personnelles** (jusqu'à 500 caractères) est prévue pour vos propres
remarques après l'avoir cuisinée — par exemple "trop salé, réduire le sel la
prochaine fois" ou "j'ai adoré, à refaire pour Noël".

Pour chaque ligne d'ingrédient : tapez les premières lettres dans la case
"Ingrédient" pour voir apparaître une liste de suggestions juste en dessous,
qui **reste affichée et se met à jour pendant que vous tapez** (recherche
insensible aux accents et à la casse — par exemple taper "e" affiche aussi
bien "Eau" que "Échalote" ou "Épices"). Cliquez sur la suggestion voulue pour
la sélectionner (ou utilisez la flèche bas puis Entrée). Cliquer dans une case
vide affiche aussi toutes les suggestions disponibles. Si vous validez un nom
qui ne correspond à aucun ingrédient existant, un message vous demandera
d'utiliser d'abord le bouton "🥕 Nouvel ingrédient". Indiquez ensuite la
quantité **pour 1 personne**, puis l'unité :
- **Gr** (gramme)
- **cl** (centilitre)
- **pièce**
- **cuillère à soupe**
- **cuillère à café**
- **autre** : un champ de texte apparaît à côté, pour taper l'unité de votre
  choix (kg, L, botte, gousse...)

Si l'ingrédient recherché n'existe pas encore, cliquez sur "🥕 Nouvel
ingrédient" en haut de la section : il sera ajouté à la liste et disponible
immédiatement dans les menus déroulants. Le bouton "+ Ajouter un ingrédient"
et le bouton "Enregistrer" restent toujours juste en dessous du dernier
ingrédient ajouté, où que vous en soyez dans le formulaire.

> Remarque : les recettes créées avant cette mise à jour avec les anciennes
> unités (kg, L...) sont automatiquement reconverties à l'ouverture : "g"
> devient "Gr", "cl" reste "cl", et tout le reste (kg, L, etc.) passe en
> "autre" avec le texte d'origine conservé.

> Remarque sur le classement alphabétique : les ingrédients contenant "œ"
> (comme "Œuf" ou "Bœuf") sont automatiquement renommés avec "oe" (donc
> "Oeuf", "Boeuf") pour apparaître avec les mots en "o". Les mots commençant
> par un "e" accentué (Échalote, Épices, Édulcorant...) sont désormais
> classés avec les mots en "e" plutôt qu'à la fin de la liste.

**🧾 Voir toutes les recettes (liste de courses)**
Une **barre de recherche** et un menu **"Trier par :"** permettent de
retrouver et d'organiser rapidement la liste des recettes affichées (utile si
vous en avez beaucoup). Un encadré **"Filtrer par ingrédient"** permet
d'affiner davantage : cliquez dans un champ et **tapez les premières
lettres** de l'ingrédient recherché pour faire apparaître une liste de
suggestions qui se réduit au fur et à mesure — exactement comme pour choisir
un ingrédient dans le formulaire de recette.
- **"Je veux :"** — jusqu'à 2 ingrédients ; seules les recettes qui
  contiennent **tous** les ingrédients choisis restent affichées ;
- **"Je ne veux pas :"** — jusqu'à 2 ingrédients ; les recettes contenant
  **l'un ou l'autre** de ces ingrédients sont masquées ;
- "Réinitialiser" vide les 4 champs et réaffiche toutes les recettes.

Ces filtres se combinent avec la barre de recherche (une recette doit
satisfaire toutes les conditions actives pour rester affichée) — pratique
pour trouver rapidement "une recette avec du poulet mais sans crème
fraîche", par exemple.

Les recettes favorites sont signalées par une ⭐. Si
vous avez utilisé "🛒 Ajouter à la liste de courses" depuis "Voir une recette
précise", ces recettes sont **déjà ajoutées à la liste de courses** avec le
bon nombre de personnes dès l'ouverture de cette fenêtre. À droite du
nombre de personnes de chaque recette :
- **"🛒 Ajouter aux courses"** ajoute immédiatement les ingrédients de cette
  recette (à la quantité de personnes indiquée) à la liste de courses
  affichée en bas de la fenêtre. Un ingrédient déjà présent (venant d'une
  autre recette ou ajouté manuellement) voit sa quantité **s'additionner**
  plutôt que de créer une ligne en double. Ajouter une même recette une
  seconde fois avec un nombre de personnes différent **remplace** la
  quantité précédente pour cette recette (plutôt que de la compter deux
  fois) ;
- **"✏️ Modifier"** ouvre directement le formulaire d'édition de cette
  recette, sans avoir à passer par "Modifier / Supprimer une recette".

Ajoutez ainsi autant de recettes que vous voulez à la liste, une par une,
en ajustant le nombre de personnes de chacune avant de cliquer sur
"🛒 Ajouter aux courses". La liste de courses en bas de la fenêtre se met à
jour à chaque ajout, **regroupée par rayon de magasin** (Fruits & Légumes,
Viandes & Poissons, Crèmerie, Boulangerie & Pâtisserie, Épicerie, Herbes &
Épices, Boissons, Autre) plutôt qu'en une simple liste alphabétique —
pratique pour suivre l'ordre des rayons pendant les courses. Une fois votre
liste complète :
- "📤 Exporter" ouvre une petite fenêtre pour choisir le format d'export :
  - "📝 Exporter en texte" enregistre la liste dans un fichier `.txt` ;
  - "📊 Exporter en Excel" enregistre la liste dans un fichier `.xlsx` (une
    feuille "Recettes" et une feuille "Ingrédients" avec une colonne "Rayon") ;
  - "📄 Exporter en PDF" enregistre la liste dans un fichier `.pdf` mis en
    forme, avec les mêmes regroupements par rayon.
  Ce même bouton (et cette même fenêtre de choix de format) est aussi
  disponible depuis "Planning de la semaine" et "Mes menus" ;
- "🖨️ Imprimer" envoie directement la liste à votre imprimante par défaut
  (elle est d'abord générée en PDF en coulisses, puis envoyée à l'impression) ;
- "☑️ Mode courses (cocher au fur et à mesure)" ouvre la liste calculée sous
  forme de cases à cocher, organisées par rayon, pour pointer chaque article
  au fur et à mesure que vous le mettez dans le caddie (un compteur affiche
  votre progression) — ce même bouton est aussi disponible depuis le
  "Planning de la semaine" et "Mes menus" ;
- "🗑 Vider la liste de courses" efface la mémoire de présélection issue de
  "🛒 Ajouter à la liste de courses", les ingrédients ajoutés manuellement,
  et **remet la liste de courses affichée à zéro**.

Si la liste de courses affichée contient des ingrédients au moment où vous
fermez cette fenêtre, un message vous prévient qu'elle sera perdue et vous
laisse annuler la fermeture — pensez à "💾 Enregistrer cette liste pour plus
tard" avant de fermer si vous voulez la conserver.

**✏️ Modifier ou supprimer un ingrédient de la liste calculée**
Dès qu'un ingrédient apparaît dans la liste de courses affichée (que ce soit
via "🛒 Ajouter aux courses" sur "Toutes les recettes", ou via "Calculer la
liste de courses" sur "Planning de la semaine"/"Nouveau menu"), chaque ligne
est modifiable : changez la **quantité** directement dans son champ
(validez avec Entrée ou en cliquant ailleurs), ou cliquez sur "🗑" à droite
d'une ligne pour retirer complètement cet ingrédient de la liste — pratique
si vous avez déjà ce produit chez vous. Ces modifications sont prises en
compte par tous les exports, l'impression et le "Mode courses", tant que
vous n'ajoutez pas une nouvelle recette ou ne relancez pas le calcul (ce qui
recalcule sans perdre vos modifications précédentes sur "Toutes les
recettes", mais repart de zéro sur "Planning"/"Nouveau menu").

**💾 Enregistrer / 📂 Charger une liste de courses**
Disponible sur "Toutes les recettes", "Planning de la semaine" et "Nouveau
menu" : "💾 Enregistrer cette liste pour plus tard" sauvegarde la liste
actuellement affichée (avec vos éventuelles modifications de quantité et
suppressions) sous un nom de votre choix, pour la retrouver un autre jour —
pratique pour une liste de courses récurrente, ou pour continuer vos
courses en plusieurs fois. "📂 Charger une liste enregistrée" ouvre la
liste de toutes vos listes sauvegardées (avec leur nombre d'ingrédients et
leur date), pour en **charger** une (elle remplace alors la liste
actuellement affichée, quelle que soit la fenêtre depuis laquelle vous
l'avez enregistrée) ou en **supprimer** une définitivement.

**➕ Ajouter un ingrédient à la liste de courses**
Ce bouton (disponible sur "Toutes les recettes", "Planning de la semaine" et
"Nouveau menu") permet d'ajouter à la liste de courses un ou plusieurs
articles qui ne viennent d'aucune recette — du papier essuie-tout, des sacs
poubelle, ou tout autre ingrédient que vous voulez juste ne pas oublier.
Choisissez l'ingrédient (ou créez-en un nouveau directement depuis cette
fenêtre via "🥕 Nouvel ingrédient" s'il n'existe pas encore dans votre
liste), indiquez une quantité et une unité (libre : pièce, boîte, paquet,
rouleau...), puis "➕ Ajouter à la liste" : l'ingrédient rejoint une **liste
d'attente**, et les champs se vident pour enchaîner rapidement la saisie du
suivant, sans avoir à rouvrir la fenêtre à chaque fois. Une fois tous vos
articles ajoutés à la liste d'attente (retirez-en un avec "🗑 Retirer de la
liste d'attente" si besoin), cliquez sur "✅ Valider tous ces ingrédients"
pour les envoyer d'un coup dans la liste de courses, regroupés par rayon
comme les autres ingrédients — la liste affichée s'actualise immédiatement.

Chaque bouton d'export/impression recalcule automatiquement la liste à partir
de la sélection actuelle.

> Les quantités totales sont automatiquement converties vers une unité plus
> parlante quand elles deviennent grandes : les **grammes passent en
> kilogrammes au-delà de 1000 g** (ex. 1500 g → 1,5 kg), et les
> **centilitres passent en litres au-delà de 100 cl** (ex. 250 cl → 2,5 L).
> Cette conversion s'applique partout où un total est calculé : liste de
> courses, planning de la semaine et menus.

> Le classement par rayon repose sur la reconnaissance du nom de
> l'ingrédient. Il fonctionne très bien avec les ~1000 ingrédients fournis et
> la plupart des noms courants, mais un ingrédient au nom très inhabituel
> pourra atterrir dans la catégorie "Autre" plutôt que dans le bon rayon.

> Partout où l'on choisit une recette dans une liste (liste de courses,
> "Voir une recette précise", "Modifier/Supprimer", livre de recettes,
> planning, menus...), chaque recette affiche directement à côté de son nom
> son **temps total** (préparation + cuisson), sa **difficulté** et ses
> **allergènes** éventuels (précédés de ⚠) entre parenthèses — pratique pour
> comparer d'un coup d'œil ou repérer une allergie sans avoir à ouvrir
> chaque recette.

**🍽️ Voir une recette précise**
La fenêtre s'ouvre à la hauteur de votre écran, et l'affichage d'une
recette est divisé en deux panneaux côte à côte : à gauche les ingrédients
et les informations (allergènes, coût, nutrition), à droite la description
et les notes personnelles. Les boutons d'action sont alignés en rangées de
4 pour un accès rapide.

Une **barre de recherche** et un menu **"Trier :"** (Nom, Temps de
préparation, Difficulté, Note, Ajoutées récemment) permettent de retrouver
rapidement une recette (là aussi, les favorites sont marquées d'une ⭐, et la
note éventuelle s'affiche en étoiles à côté du nom ; la recherche reconnaît
aussi les étiquettes). Sélectionnez une recette dans la liste — le nombre de
personnes se préremplit automatiquement avec la valeur par défaut de la
recette — puis "Afficher la recette" pour voir sa galerie de photos, sa note,
son temps de préparation/cuisson, sa difficulté, ses ingrédients recalculés,
sa description et vos notes personnelles.

Une fois une recette affichée, ajustez rapidement les portions avec les
boutons **−1 / +1 / ÷2 / ×2** à côté du nombre de personnes : l'affichage se
met à jour immédiatement, sans avoir à retaper le nombre ni recliquer sur
"Afficher la recette".

- "📄 Exporter en PDF" génère un document avec le nom, la catégorie, la note,
  le temps de préparation/cuisson, la difficulté, la première photo, les
  ingrédients (recalculés pour le nombre de personnes affiché), la
  description et les notes personnelles.
- "🖨️ Imprimer" envoie directement cette même mise en page à votre
  imprimante par défaut, sans passer par un fichier à ouvrir manuellement.
- "🛒 Ajouter à la liste de courses" mémorise la recette affichée (avec son
  nombre de personnes) pour qu'elle soit automatiquement ajoutée à la liste
  de courses la prochaine fois que vous ouvrirez "Voir toutes les recettes" —
  pratique pour construire sa liste de courses recette par recette plutôt
  que de tout sélectionner d'un coup.
- "🍳 J'ai cuisiné ça !" incrémente un compteur "nombre de fois cuisinée" pour
  cette recette (visible dans les Statistiques) et enregistre la date du
  jour. Une petite fenêtre s'ouvre ensuite pour ajouter, si vous le
  souhaitez, une **note** (ex. "un peu trop salé, réduire le sel") et/ou une
  **photo** du résultat — ou cliquez simplement sur "Passer" pour ignorer
  cette étape.
- "📔 Journal de cuisine" affiche l'historique de toutes vos notes et photos
  pour cette recette (la plus récente en premier), avec le nombre total de
  fois cuisinée — pratique pour se souvenir de ce qui a marché (ou pas) la
  dernière fois.
- "🖥️ Mode cuisine (plein écran)" ouvre la recette dans une fenêtre maximisée,
  en très gros caractères, sans menus ni distractions — pratique à lire posé
  à côté des fourneaux. Les boutons "−" et "+" en haut à gauche ajustent le
  nombre de personnes sans quitter ce mode. Appuyez sur **Échap** ou cliquez
  sur "✕ Fermer" pour revenir à l'écran normal, ou sur **F11** pour basculer
  en plein écran natif (masque complètement la barre des tâches) si vous le
  souhaitez.
- "📱 QR Code" génère un QR code contenant le nom et les ingrédients de la
  recette (adaptés au nombre de personnes affiché), à scanner avec
  l'appareil photo d'un téléphone pour l'emporter sans imprimer ni
  transférer de fichier. Le bouton "💾 Enregistrer en image (PNG)" permet de
  le sauvegarder. Pour les recettes très longues, le contenu encodé est
  automatiquement résumé (nom + ingrédients uniquement) afin de rester
  facilement scannable.
- "⏲️ Minuteurs" ouvre une fenêtre où vous pouvez régler et démarrer
  **plusieurs minuteurs indépendants**, empilés les uns sous les autres —
  pratique pour chronométrer plusieurs étapes en même temps (ex. un pour les
  pâtes, un pour la sauce). Le premier minuteur est pré-rempli avec le temps
  de cuisson (ou de préparation) de la recette affichée ; "➕ Ajouter un
  minuteur" en ajoute d'autres, chacun réglable séparément (minutes,
  secondes) avec ses propres boutons ▶️ Démarrer / ⏸️ Pause / 🔄
  Réinitialiser, et "🗑" pour le retirer. Une seule fenêtre de minuteurs est
  utilisée pour toute l'application : la rouvrir depuis une autre recette y
  ajoute simplement un minuteur de plus, sans en ouvrir une deuxième. Cette
  fenêtre **reste toujours visible au premier plan**, même par-dessus le
  mode cuisine en plein écran, pour ne jamais perdre de vue vos minuteurs en
  cours. **Quand un minuteur arrive à zéro, sa ligne clignote en rouge et un
  signal sonore retentit** jusqu'à ce que vous cliquiez dessus (ou que vous
  le redémarriez/réinitialisiez) pour arrêter l'alarme.

**✏️ Modifier / Supprimer une recette**
Une **barre de recherche** et un menu **"Trier par :"** filtrent et
organisent la liste (avec favoris ⭐ et note en étoiles affichés). Sélectionnez
une recette puis :
- "✏️ Modifier" ouvre le même formulaire que pour l'ajout, pré-rempli, pour
  changer le nom, les photos, le temps, la difficulté, la note, les notes
  personnelles ou les ingrédients ;
- "📋 Dupliquer" crée une copie complète de la recette (nommée
  "(copie)"), avec ses propres fichiers photo indépendants — pratique pour
  créer une variante sans tout retaper ;
- "🗑️ Supprimer" envoie la recette à la **corbeille** (après confirmation) —
  elle n'est pas effacée définitivement, voir ci-dessous.
Vous pouvez aussi supprimer une recette directement depuis l'écran de
modification.

**🗑️ Corbeille**
Une recette supprimée (depuis "Modifier/Supprimer" ou depuis l'écran de
modification) n'est jamais effacée immédiatement : elle est déplacée dans la
corbeille, photos comprises. Cette fenêtre liste les recettes supprimées avec
leur date de suppression, et permet :
- "♻️ Restaurer" : remet la recette sélectionnée dans votre livre de recettes
  (si une recette du même nom existe déjà, la version restaurée est renommée
  "(restaurée)" pour éviter tout conflit) ;
- "🗑️ Supprimer définitivement" : efface pour de bon la recette sélectionnée
  et ses photos (irréversible) ;
- "🧹 Vider la corbeille" : efface définitivement tout son contenu d'un coup
  (irréversible, avec confirmation).

**💾 Importer / Exporter les données**
- "📤 Exporter toutes mes données" enregistre un fichier `.zip` contenant vos
  recettes, vos ingrédients et vos photos — pratique pour faire une
  sauvegarde ou transférer vos données vers un autre ordinateur.
- "📥 Importer des données" lit un fichier `.zip` exporté précédemment. On
  vous demande alors :
  - **Fusionner** : ajoute les recettes et ingrédients importés à ceux déjà
    présents (une recette en double est renommée "(importé)" pour éviter
    d'écraser la vôtre) ;
  - **Remplacer** : efface les données actuelles et les remplace entièrement
    par celles du fichier importé.

**🗄️ Sauvegardes automatiques** (dans le même écran, plus bas)
L'application crée automatiquement une sauvegarde complète (recettes,
ingrédients, photos) au démarrage, au maximum une fois toutes les 24 heures,
sans rien vous demander. Les 10 sauvegardes les plus récentes sont
conservées (les plus anciennes sont supprimées automatiquement) dans le
dossier `backups/`, sous le nom `sauvegarde_auto_mesrecettes_AAAA-MM-JJ_HHMMSS.zip`
— ce préfixe distinctif évite toute confusion avec les fichiers de
sauvegarde d'un autre logiciel qui utiliserait le même dossier (notamment
utile pour le dossier cloud, voir ci-dessous). Dans cette fenêtre :
- la liste affiche chaque sauvegarde automatique avec sa date et sa taille ;
- "💾 Sauvegarder maintenant" force la création d'une sauvegarde
  immédiatement, sans attendre le prochain démarrage ;
- "♻️ Restaurer la sélection" restaure la sauvegarde choisie (avec le même
  choix Fusionner/Remplacer que pour un import classique).

**☁️ Sauvegarde automatique dans le cloud** (dans le même écran, tout en bas)
Pour que vos sauvegardes soient aussi envoyées en ligne automatiquement,
cliquez sur "📁 Choisir un dossier cloud" et sélectionnez un dossier
synchronisé par un client déjà installé sur votre PC — par exemple votre
dossier **Google Drive**, **OneDrive** ou **Dropbox** local. À partir de là,
chaque sauvegarde automatique (créée au démarrage ou via "💾 Sauvegarder
maintenant") est aussi copiée dans ce dossier ; c'est ensuite le client cloud
lui-même qui se charge de l'envoyer en ligne, exactement comme n'importe quel
autre fichier que vous y déposeriez. "🚫 Désactiver" arrête cette copie
automatique (les sauvegardes déjà envoyées restent en ligne).

> L'application ne se connecte à aucun service en ligne elle-même — elle
> copie simplement le fichier dans un dossier local que votre client cloud
> (déjà installé et connecté à votre compte) surveille et synchronise tout
> seul. Vous n'avez donc rien à configurer côté Google/Microsoft/Dropbox
> dans l'application.

**🧊 Que puis-je cuisiner ?**
Cette fenêtre vous aide à trouver une recette à partir de ce que vous avez
déjà chez vous. À l'ouverture, quelques **ingrédients de base courants sont
déjà cochés** dans "Ce que j'ai" (Sel, Poivre, Huile de tournesol, Huile
d'olive, Beurre, Farine, Sucre, Vinaigre, Moutarde, Riz, Pâtes, Lait) — le
genre de choses qu'on a presque toujours dans un placard ou un frigo.
Retirez-en si certains ne s'appliquent pas chez vous (double-clic ou
"🗑 Retirer"). À gauche, recherchez et ajoutez (double-clic ou bouton "➕
Ajouter →") les autres ingrédients que vous possédez ; ils passent dans la
colonne de droite "Ce que j'ai". Une fois votre sélection faite, cliquez sur
"🔍 Voir les recettes réalisables" :
- les recettes **✅ réalisables** (tous les ingrédients sont cochés) ;
- les recettes **🟡 presque réalisables** (il ne manque qu'1 à 3
  ingrédients, listés) apparaissent ensuite.
Sélectionnez une recette dans les résultats puis "📖 Consulter la recette
sélectionnée" (ou double-cliquez dessus) pour l'ouvrir directement dans
"Voir une recette précise".

**📅 Planning de la semaine**
La fenêtre se présente comme une **vraie grille calendrier** : les 7 jours de
la semaine (Lundi à Dimanche) en colonnes, et les **7 créneaux de repas** en
lignes :
- Petit-déjeuner
- Déjeuner — Entrée, Déjeuner — Plat, Déjeuner — Dessert
- Dîner — Entrée, Dîner — Plat, Dîner — Dessert

Dans chaque case, choisissez une recette et le nombre de personnes (👤). Les
noms des jours de la semaine **restent toujours visibles en haut**, même en
faisant défiler la grille vers le bas pour voir les créneaux de repas
suivants. La grille défile horizontalement et verticalement si la fenêtre
est trop petite pour tout afficher.

"💾 Enregistrer le planning" sauvegarde vos choix pour la prochaine fois que
vous ouvrez cette fenêtre. "🗑 Tout effacer" vide le planning. "📆 Exporter
vers un calendrier (.ics)" enregistre un fichier `.ics` que vous pouvez
importer dans Google Agenda, Outlook ou Calendrier (Apple) : chaque repas
prévu (petit-déjeuner, déjeuner, dîner) devient un évènement qui **se répète
automatiquement chaque semaine** au même jour et à la même heure — le
déjeuner et le dîner regroupent en un seul évènement l'entrée, le plat et le
dessert prévus. Une fois vos repas de la semaine choisis :
- "Calculer la liste de courses de la semaine" additionne les ingrédients
  nécessaires pour tous les créneaux planifiés (tous jours et repas
  confondus), regroupés par rayon ;
- le bouton "📤 Exporter" et "🖨️ Imprimer" fonctionnent comme pour la liste
  de courses classique (voir "Toutes les recettes" plus haut), y compris la
  possibilité de modifier une quantité, retirer un ingrédient, ou enregistrer
  la liste pour plus tard.

**📋 Mes menus**
Créez et sauvegardez des combinaisons de plusieurs recettes (par exemple
entrée + plat + dessert) pour une occasion particulière. "➕ Nouveau menu"
ouvre un formulaire où vous donnez un nom au menu, puis ajoutez des recettes
une à une (avec leur nombre de personnes) via "+ Ajouter". La liste des
recettes du menu s'affiche, avec un bouton "🗑 Retirer du menu". "💾
Enregistrer le menu" sauvegarde la combinaison pour la retrouver plus tard
dans la liste "Mes menus" (bouton "👁 Ouvrir" pour la rouvrir, "🗑 Supprimer"
pour l'effacer). Comme pour le planning, vous pouvez calculer la liste de
courses du menu, l'exporter (bouton "📤 Exporter" : texte/Excel/PDF) et
l'imprimer, modifier une quantité ou retirer un ingrédient directement dans
la liste affichée, ajouter des ingrédients qui ne viennent d'aucune recette,
et enregistrer/charger cette liste pour plus tard — les recettes y sont
regroupées par catégorie (Apéro, Entrée, Plat, Sauce, Dessert, Boisson,
Autre) dans l'ordre logique d'un repas.

**📊 Statistiques**
Un aperçu synthétique de votre livre de recettes : nombre total de recettes,
répartition par catégorie et par difficulté, nombre de favoris, note moyenne
et recettes notées 5 étoiles, recettes les plus cuisinées (grâce au bouton
"🍳 J'ai cuisiné ça !" dans "Voir une recette précise"), étiquettes les plus
utilisées, deux sections "recettes oubliées" pour redécouvrir votre livre
(les recettes **jamais cuisinées**, et celles **pas cuisinées depuis plus de
90 jours**), ainsi que :
- **💰 Coût moyen par personne**, calculé sur les recettes ayant au moins un
  prix d'ingrédient renseigné (voir "💰 Gérer les prix") ;
- **🥗 Calories moyennes par personne**, calculées sur les recettes dont les
  ingrédients sont reconnus dans la base nutritionnelle ;
- **📈 un graphique** du nombre de recettes cuisinées par mois, sur les 12
  derniers mois, pour visualiser vos habitudes de cuisine dans le temps.

**📖 Exporter le livre de recettes**
Réunit plusieurs recettes en **un seul PDF**, façon vrai livre de cuisine.
Filtrez par catégorie si besoin, cochez/décochez les recettes à inclure
("Tout cocher" / "Tout décocher" pour aller vite), puis "📄 Générer le PDF du
livre" : le document commence par une page de sommaire, suivie d'une recette
par page (photo, ingrédients pour le nombre de personnes par défaut de
chaque recette, description, allergènes...) — pratique pour imprimer ou
relier votre livre de recettes personnel, ou en offrir une version papier.
**Toutes les pages sont numérotées** ("Page X / Y" en bas de chaque page),
et le **sommaire indique le numéro de page de chaque recette** en face de
son nom, pour retrouver rapidement une recette dans un livre imprimé —
y compris pour les recettes dont le contenu déborde sur plusieurs pages
(beaucoup d'ingrédients, longue description...), dont la pagination est
prise en compte automatiquement.

Toutes les recettes sont sauvegardées dans `recipes.json`, et les photos dans
le dossier `images/`. Gardez toujours `main.py`, `recipes.json`,
`ingredients.json` et `images/` ensemble si vous déplacez ou sauvegardez le
projet — ou utilisez simplement l'export en .zip qui regroupe tout.

> À propos des boutons "🖨️ Imprimer" : l'application tente d'abord d'envoyer
> le document directement à votre imprimante par défaut. Si votre système ne
> le permet pas (cas fréquent selon le lecteur PDF installé), le PDF
> s'ouvre alors automatiquement dans votre lecteur habituel : il ne reste
> qu'à faire **Ctrl+P** (ou cliquer sur son bouton Imprimer) pour lancer
> l'impression depuis là. Ce n'est qu'en cas d'échec total que l'application
> vous indique simplement le chemin du PDF généré pour l'ouvrir vous-même.

