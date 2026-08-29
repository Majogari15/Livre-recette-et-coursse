import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import copy
import difflib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import io
import random
import threading
import webbrowser
import tempfile
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from datetime import datetime, timedelta

# Pillow est nécessaire pour afficher les photos des recettes.
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# reportlab est nécessaire pour l'export PDF de la liste de courses.
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as pdf_canvas
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# openpyxl est nécessaire pour l'export Excel de la liste de courses.
try:
    from openpyxl import Workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# qrcode est nécessaire pour exporter une recette sous forme de QR code.
try:
    import qrcode
    QRCODE_AVAILABLE = True
except ImportError:
    QRCODE_AVAILABLE = False

# pytesseract est nécessaire pour importer une recette depuis une photo (OCR).
# Il ne suffit pas de l'installer via pip : il nécessite aussi le programme
# Tesseract OCR installé séparément sur le système (voir le LISEZ-MOI).
try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except ImportError:
    PYTESSERACT_AVAILABLE = False

# pyttsx3 est nécessaire pour la lecture à voix haute en mode cuisine.
# Fonctionne hors ligne (s'appuie sur la synthèse vocale déjà installée sur
# le système : SAPI5 sous Windows).
try:
    import pyttsx3
    PYTTSX3_AVAILABLE = True
except ImportError:
    PYTTSX3_AVAILABLE = False


if getattr(sys, "frozen", False):
    # Application compilée avec PyInstaller (.exe) : __file__ pointe vers un
    # dossier temporaire d'extraction interne à Windows (pas vers le dossier
    # où se trouve le .exe), donc on utilise plutôt l'emplacement réel de
    # l'exécutable pour que les fichiers de données (recettes, ingrédients,
    # photos...) soient bien lus/écrits à côté du .exe, là où l'utilisateur
    # les voit et les place.
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "recipes.json")
INGREDIENTS_FILE = os.path.join(BASE_DIR, "ingredients.json")
DEFAULT_INGREDIENTS_FILE = os.path.join(BASE_DIR, "ingredients_par_defaut.json")
NUTRITION_DATA_FILE = os.path.join(BASE_DIR, "valeurs_nutritionnelles.json")
INGREDIENT_ALLERGENS_FILE = os.path.join(BASE_DIR, "ingredient_allergenes.json")
INGREDIENT_SUBSTITUTIONS_FILE = os.path.join(BASE_DIR, "ingredient_substitutions.json")
# Fichiers de traduction, par langue autre que le français (langue de
# référence des données, jamais dans ce dictionnaire). Pour ajouter une
# langue supplémentaire à l'avenir, il suffit d'ajouter une entrée ici et
# de fournir les deux fichiers JSON correspondants.
INGREDIENT_SUBSTITUTIONS_TRANSLATION_FILES = {
    "en": os.path.join(BASE_DIR, "ingredient_substitutions_en.json"),
    "es": os.path.join(BASE_DIR, "ingredient_substitutions_es.json"),
    "de": os.path.join(BASE_DIR, "ingredient_substitutions_de.json"),
}
INGREDIENT_TRANSLATIONS_FILES = {
    "en": os.path.join(BASE_DIR, "ingredient_translations_en.json"),
    "es": os.path.join(BASE_DIR, "ingredient_translations_es.json"),
    "de": os.path.join(BASE_DIR, "ingredient_translations_de.json"),
}
FLAG_FILES = {
    "fr": os.path.join(BASE_DIR, "flag_fr.png"),
    "en": os.path.join(BASE_DIR, "flag_uk.png"),
    "es": os.path.join(BASE_DIR, "flag_es.png"),
    "de": os.path.join(BASE_DIR, "flag_de.png"),
}
# Codes de langue attendus par Tesseract OCR (différents des codes ISO à 2
# lettres utilisés partout ailleurs dans l'application) pour l'import de
# recette depuis une photo : suit la langue actuellement sélectionnée dans
# l'interface, sur l'hypothèse que le texte manuscrit ou imprimé photographié
# est dans cette même langue. Nécessite que le paquet linguistique
# correspondant soit installé pour Tesseract (voir LISEZ-MOI).
TESSERACT_LANG_CODES = {
    "fr": "fra",
    "en": "eng",
    "es": "spa",
    "de": "deu",
}
INGREDIENT_OVERRIDES_FILE = os.path.join(BASE_DIR, "ingredient_custom_data.json")
INGREDIENT_PRICES_FILE = os.path.join(BASE_DIR, "ingredient_prices.json")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
WEEKLY_PLAN_FILE = os.path.join(BASE_DIR, "weekly_plan.json")
WEEKLY_PLAN_HISTORY_FILE = os.path.join(BASE_DIR, "weekly_plan_history.json")
WEEKLY_PLAN_TEMPLATES_FILE = os.path.join(BASE_DIR, "weekly_plan_templates.json")
MENUS_FILE = os.path.join(BASE_DIR, "menus.json")
TRASH_FILE = os.path.join(BASE_DIR, "trash.json")
BACKUPS_DIR = os.path.join(BASE_DIR, "backups")
RECENT_VIEWS_FILE = os.path.join(BASE_DIR, "recent_views.json")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

os.makedirs(IMAGES_DIR, exist_ok=True)

# Fichiers de données personnelles inclus dans une sauvegarde complète (tout
# ce qui n'est pas fourni avec l'application elle-même : les bases
# d'ingrédients/allergènes/nutrition/substitutions ne changent pas d'un
# utilisateur à l'autre et ne sont donc pas incluses).
USER_DATA_FILES = [
    "recipes.json", "ingredients.json", "ingredient_custom_data.json",
    "ingredient_prices.json", "ingredient_dismissed_pairs.json",
    "weekly_plan.json", "weekly_plan_history.json", "weekly_plan_templates.json",
    "menus.json", "saved_shopping_lists.json", "pantry.json",
    "trash.json", "recent_views.json", "settings.json",
]


def export_full_backup(zip_path):
    """Crée une sauvegarde complète (.zip) de toutes les données
    utilisateur : recettes, ingrédients personnalisés, garde-manger,
    planning (+ historique et modèles), menus, listes de courses
    enregistrées, corbeille, réglages, et le dossier des photos."""
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename in USER_DATA_FILES:
            filepath = os.path.join(BASE_DIR, filename)
            if os.path.exists(filepath):
                zf.write(filepath, filename)
        if os.path.isdir(IMAGES_DIR):
            for root, dirs, files in os.walk(IMAGES_DIR):
                for f in files:
                    full = os.path.join(root, f)
                    rel = os.path.relpath(full, BASE_DIR)
                    zf.write(full, rel)


def import_full_backup(zip_path):
    """Restaure une sauvegarde complète : extrait tous les fichiers dans le
    dossier de l'application, en écrasant les fichiers existants de même
    nom (les recettes/photos/réglages actuels seront remplacés)."""
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(BASE_DIR)


def load_default_ingredients():
    """Charge la liste des ~1000 ingrédients de cuisine les plus courants,
    fournie avec l'application (fichier ingredients_par_defaut.json)."""
    if os.path.exists(DEFAULT_INGREDIENTS_FILE):
        try:
            with open(DEFAULT_INGREDIENTS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            return []
    return []


def normalize_oe(text):
    """Remplace les ligatures œ/Œ par 'oe'/'Oe' afin qu'un mot comme « Œuf »
    s'affiche et se classe avec les autres mots commençant par « o »."""
    return text.replace("œ", "oe").replace("Œ", "Oe")


def ingredient_sort_key(text):
    """Clé de tri qui ignore les accents (é, è, ê, à, ç...) afin qu'un mot
    comme « Échalote » se classe avec les autres mots en « e », et qui
    convertit d'abord œ/Œ en oe/Oe pour un classement cohérent avec « o »."""
    normalized = normalize_oe(text)
    decomposed = unicodedata.normalize("NFD", normalized)
    stripped = "".join(c for c in decomposed if unicodedata.category(c) != "Mn")
    return stripped.lower()


def print_file(path):
    """Tente d'envoyer un fichier directement à l'imprimante par défaut du
    système. Si ce n'est pas possible (aucune application n'ayant enregistré
    de « impression silencieuse » pour ce type de fichier — un cas fréquent
    sur Windows selon le lecteur PDF installé), ouvre le fichier dans
    l'application par défaut à la place, pour que l'utilisateur puisse
    l'imprimer d'un clic depuis celle-ci.

    Retourne "printed" (envoyé directement à l'imprimante), "opened" (ouvert
    dans l'application par défaut, à imprimer manuellement), ou None (échec
    complet)."""
    if os.name == "nt":
        try:
            os.startfile(path, "print")
            return "printed"
        except OSError:
            try:
                os.startfile(path)
                return "opened"
            except OSError:
                return None
    else:
        try:
            subprocess.run(["lp", path], check=True)
            return "printed"
        except Exception:
            try:
                subprocess.run(["xdg-open", path], check=True)
                return "opened"
            except Exception:
                return None


def report_print_result(result, temp_path, subject):
    """Affiche le message adapté selon le résultat de print_file()."""
    if result == "printed":
        messagebox.showinfo("Impression", f"Envoyé à l'imprimante : {subject}.")
    elif result == "opened":
        messagebox.showinfo(
            "Ouvert pour impression",
            f"Impossible d'envoyer directement à l'imprimante depuis l'application.\n\n"
            f"Le PDF ({subject}) a été ouvert dans votre lecteur habituel : "
            "utilisez Ctrl+P (ou le bouton Imprimer) pour l'imprimer depuis là."
        )
    else:
        messagebox.showerror(
            "Impression impossible",
            f"Impossible d'ouvrir ou d'imprimer automatiquement.\n"
            f"Le PDF a tout de même été généré ici, vous pouvez l'ouvrir et l'imprimer manuellement :\n{temp_path}"
        )


def get_temp_pdf_path(prefix):
    """Crée un chemin unique dans le dossier temporaire du système, utilisé
    pour générer un PDF juste avant de l'envoyer à l'imprimante."""
    return os.path.join(tempfile.gettempdir(), f"{prefix}_{uuid.uuid4().hex}.pdf")


# Classement des ingrédients par rayon de magasin, utilisé pour regrouper la
# liste de courses. Chaque rayon est associé à une liste de mots-clés ; le
# premier rayon dont un mot-clé apparaît dans le nom de l'ingrédient est
# retenu. L'ordre de la liste correspond à un parcours de magasin classique.
RAYON_KEYWORDS = [
    ("Fruits & Légumes", [
        "ail", "oignon", "echalote", "poireau", "carotte", "celeri", "panais",
        "navet", "betterave", "radis", "pomme de terre", "patate", "topinambour",
        "tomate", "concombre", "courgette", "aubergine", "poivron", "piment",
        "chou", "brocoli", "epinard", "blette", "oseille", "roquette", "mache",
        "laitue", "batavia", "endive", "chicoree", "cresson", "pissenlit",
        "fenouil", "artichaut", "asperge", "petit pois", "haricot vert",
        "haricot beurre", "fève", "mais", "courge", "potiron", "citrouille",
        "champignon", "girolle", "cepe", "truffe", "igname", "manioc", "gombo",
        "salsifis", "cardon", "crosne", "chayotte", "cornichon", "avocat",
        "pomme", "poire", "banane", "orange", "clementine", "mandarine",
        "pamplemousse", "pomelo", "citron", "kiwi", "fraise", "framboise",
        "myrtille", "mure", "groseille", "cassis", "cerise", "griotte",
        "abricot", "peche", "nectarine", "prune", "mirabelle", "reine-claude",
        "raisin", "melon", "pasteque", "ananas", "mangue", "papaye",
        "fruit de la passion", "litchi", "grenade", "figue", "datte", "coing",
        "kaki", "rhubarbe", "noix de coco", "kumquat", "goyave", "carambole",
        "persil", "basilic", "thym", "romarin", "origan",
        "marjolaine", "sauge", "laurier", "menthe", "ciboulette", "cerfeuil",
        "estragon", "aneth", "gingembre frais", "curcuma frais", "citronnelle",
        "germe de soja", "pousse", "algue",
    ]),
    ("Viandes & Poissons", [
        "boeuf", "bœuf", "veau", "porc", "lard", "bacon", "jambon", "saucisse",
        "chorizo", "andouille", "boudin", "saucisson", "pancetta", "agneau",
        "mouton", "poulet", "coq", "dinde", "canard", "magret", "confit de",
        "foie gras", "oie", "pintade", "caille", "pigeon", "lapin", "gibier",
        "chevreuil", "sanglier", "cheval", "steak", "viande", "kefta",
        "saumon", "truite", "cabillaud", "morue", "merlu", "colin", "lieu",
        "bar", "loup de mer", "dorade", "daurade", "thon", "espadon",
        "maquereau", "sardine", "anchois", "hareng", "sole", "turbot",
        "flétan", "raie", "rouget", "saint-pierre", "lotte", "baudroie",
        "congre", "anguille", "carpe", "brochet", "perche", "sandre",
        "tilapia", "panga", "poisson", "surimi", "caviar", "tarama",
        "crevette", "gambas", "langoustine", "homard", "langouste", "crabe",
        "tourteau", "etrille", "moule", "huitre", "huître", "palourde",
        "coque", "praire", "bulot", "bigorneau", "couteau", "petoncle",
        "coquille saint-jacques", "calamar", "encornet", "seiche", "poulpe",
        "oursin", "ormeau", "escargot", "grenouille",
    ]),
    ("Crèmerie", [
        "lait", "creme", "crème", "beurre", "margarine", "yaourt",
        "fromage", "faisselle", "petit-suisse", "mascarpone", "ricotta",
        "cottage", "emmental", "gruyere", "comte", "beaufort", "cantal",
        "reblochon", "morbier", "tomme", "camembert", "brie", "coulommiers",
        "munster", "epoisses", "maroilles", "livarot", "pont-l'eveque",
        "roquefort", "bleu", "fourme", "gorgonzola", "chevre", "chèvre",
        "crottin", "feta", "halloumi", "mozzarella", "burrata", "parmesan",
        "pecorino", "grana padano", "provolone", "gouda", "edam", "cheddar",
        "raclette", "fondue", "babeurre", "kefir", "skyr", "oeuf", "œuf",
    ]),
    ("Boulangerie & Pâtisserie", [
        "farine", "levure", "bicarbonate", "chocolat", "cacao", "praline",
        "praliné", "nougat", "caramel", "gelatine", "agar-agar", "pectine",
        "pate feuilletee", "pate brisee", "pate sablee", "pate a choux",
        "pate a pizza", "pate filo", "pate a crepes", "pate a gaufres",
        "genoise", "biscuit", "boudoir", "speculoos", "meringue",
        "poudre d'amande", "poudre de noisette", "amande effilee",
        "fruits confits", "nappage", "glacage", "fondant", "marron glace",
        "chataigne", "châtaigne", "pain", "baguette", "chapelure", "croutons",
        "biscotte", "cracker", "sucre", "cassonade", "vergeoise", "miel",
        "sirop", "melasse", "vanille",
    ]),
    ("Épicerie", [
        "riz", "semoule", "couscous", "boulgour", "quinoa", "epeautre",
        "orge", "sarrasin", "avoine", "ble", "blé", "fecule", "maizena",
        "tapioca", "pate", "spaghetti", "penne", "fusilli", "tagliatelle",
        "macaroni", "lasagne", "nouille", "vermicelle", "lentille",
        "pois chiche", "soja", "haricot rouge", "haricot blanc",
        "haricot noir", "huile", "graisse", "saindoux", "ghee",
        "moutarde", "ketchup", "mayonnaise", "vinaigre", "sauce",
        "concentre de tomate", "coulis", "pesto", "tapenade", "houmous",
        "tahini", "confiture", "marmelade", "gelee", "chutney", "pickles",
        "cornichons au vinaigre", "câpres", "capres", "olive", "raifort",
        "wasabi", "bouillon", "fond de veau", "fumet", "tofu", "seitan",
        "tempeh", "conserve", "boite", "boîte", "chips", "nachos",
        "pop-corn", "biscuit apero", "cacahuete grillee",
    ]),
    ("Herbes & Épices", [
        "poivre", "sel", "paprika", "cumin", "coriandre en", "cannelle",
        "muscade", "girofle", "cardamome", "anis", "curry", "garam masala",
        "ras el hanout", "za'atar", "sumac", "safran", "vanille en poudre",
        "reglisse", "genievre", "moutarde en poudre", "herbes de provence",
        "bouquet garni", "quatre epices", "piment de cayenne",
        "piment d'espelette", "chili en poudre", "epices", "épices",
        "curcuma en poudre", "gingembre en poudre", "sesame", "sésame",
        "graines de", "colorant",
    ]),
    ("Boissons", [
        "vin", "champagne", "porto", "madere", "marsala", "vermouth",
        "cognac", "armagnac", "calvados", "rhum", "whisky", "bourbon", "gin",
        "vodka", "grand marnier", "cointreau", "amaretto", "kirsch",
        "eau de vie", "biere", "bière", "cidre", "cafe", "café", "the", "thé",
        "eau gazeuse", "jus de", "sirop de grenadine", "sirop de menthe",
    ]),
]


def get_ingredient_rayon(name):
    """Retourne le rayon de magasin associé à un ingrédient, en se basant sur
    des mots-clés (recherche insensible aux accents et à la casse). Renvoie
    'Autre' si aucun mot-clé ne correspond."""
    key = ingredient_sort_key(name)
    for rayon, keywords in RAYON_KEYWORDS:
        for kw in keywords:
            if ingredient_sort_key(kw) in key:
                return rayon
    return "Autre"


RAYON_ORDER = [
    "Fruits & Légumes", "Viandes & Poissons", "Crèmerie",
    "Boulangerie & Pâtisserie", "Épicerie", "Herbes & Épices", "Boissons", "Autre",
]

# Table de correspondance des rayons par langue, sur le même principe que
# le dictionnaire d'ingrédients : la donnée réelle (regroupement, tri,
# stockage dans les listes de courses enregistrées) reste toujours le nom
# français ci-dessus, cette table ne sert qu'à l'affichage.
RAYON_TRANSLATIONS = {
    "en": {
        "fruits & légumes": "Fruits & Vegetables",
        "viandes & poissons": "Meat & Fish",
        "crèmerie": "Dairy",
        "boulangerie & pâtisserie": "Bakery & Pastry",
        "épicerie": "Grocery",
        "herbes & épices": "Herbs & Spices",
        "boissons": "Beverages",
        "autre": "Other",
    },
    "es": {
        "fruits & légumes": "Frutas y verduras",
        "viandes & poissons": "Carnes y pescados",
        "crèmerie": "Lácteos",
        "boulangerie & pâtisserie": "Panadería y repostería",
        "épicerie": "Almacén",
        "herbes & épices": "Hierbas y especias",
        "boissons": "Bebidas",
        "autre": "Otro",
    },
    "de": {
        "fruits & légumes": "Obst & Gemüse",
        "viandes & poissons": "Fleisch & Fisch",
        "crèmerie": "Milchprodukte",
        "boulangerie & pâtisserie": "Bäckerei & Konditorei",
        "épicerie": "Lebensmittel",
        "herbes & épices": "Kräuter & Gewürze",
        "boissons": "Getränke",
        "autre": "Sonstiges",
    },
}


def translate_rayon_name(rayon):
    """Retourne le nom d'affichage d'un rayon de magasin dans la langue
    actuellement sélectionnée. Comme translate_ingredient_name(), ne
    change jamais la donnée réelle utilisée pour le regroupement, le tri
    ou le stockage — seulement ce qui est montré à l'écran ou écrit dans
    un export."""
    if not rayon:
        return rayon
    if CURRENT_LANGUAGE == "fr":
        return rayon
    return RAYON_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(rayon.strip().lower(), rayon)


# Tables de correspondance pour les catégories et niveaux de difficulté,
# par langue, sur le même principe que les rayons et les ingrédients : la
# donnée réelle stockée dans chaque recette reste toujours en français
# (ces valeurs sont des clés de comparaison/filtre à de nombreux
# endroits), ces tables ne servent qu'à l'affichage et à la résolution
# des menus déroulants (qui sont toujours en liste fermée "readonly",
# donc sans ambiguïté possible contrairement aux noms d'ingrédients en
# texte libre).
CATEGORY_TRANSLATIONS = {
    "en": {
        "petit-déjeuner": "Breakfast",
        "entrée": "Starter",
        "plat": "Main course",
        "dessert": "Dessert",
        "apéro": "Appetizer",
        "boisson": "Drink",
        "sauce": "Sauce",
        "autre": "Other",
    },
    "es": {
        "petit-déjeuner": "Desayuno",
        "entrée": "Entrante",
        "plat": "Plato principal",
        "dessert": "Postre",
        "apéro": "Aperitivo",
        "boisson": "Bebida",
        "sauce": "Salsa",
        "autre": "Otro",
    },
    "de": {
        "petit-déjeuner": "Frühstück",
        "entrée": "Vorspeise",
        "plat": "Hauptgericht",
        "dessert": "Dessert",
        "apéro": "Aperitif",
        "boisson": "Getränk",
        "sauce": "Sauce",
        "autre": "Sonstiges",
    },
}

DIFFICULTY_TRANSLATIONS = {
    "en": {
        "facile": "Easy",
        "moyen": "Medium",
        "difficile": "Hard",
    },
    "es": {
        "facile": "Fácil",
        "moyen": "Medio",
        "difficile": "Difícil",
    },
    "de": {
        "facile": "Einfach",
        "moyen": "Mittel",
        "difficile": "Schwer",
    },
}


def translate_category_name(category):
    """Retourne le nom d'affichage d'une catégorie de recette dans la
    langue actuellement sélectionnée. Ne change jamais la donnée réelle
    stockée dans la recette."""
    if not category:
        return category
    if CURRENT_LANGUAGE == "fr":
        return category
    return CATEGORY_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(category.strip().lower(), category)


def translate_difficulty_name(difficulty):
    """Retourne le nom d'affichage d'un niveau de difficulté dans la
    langue actuellement sélectionnée. Ne change jamais la donnée réelle
    stockée dans la recette."""
    if not difficulty:
        return difficulty
    if CURRENT_LANGUAGE == "fr":
        return difficulty
    return DIFFICULTY_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(difficulty.strip().lower(), difficulty)


def resolve_category_input(displayed_value, category_options):
    """Résout une catégorie sélectionnée dans un menu déroulant (donc
    éventuellement affichée traduite) vers son nom canonique français
    exact. Les menus concernés sont toujours en liste fermée
    (state='readonly'), donc chaque valeur affichée correspond
    exactement à une seule catégorie française : pas d'ambiguïté possible
    ici, contrairement aux noms d'ingrédients en texte libre."""
    if not displayed_value:
        return displayed_value
    for cat in category_options:
        if cat.strip().lower() == displayed_value.strip().lower():
            return cat
    for cat in category_options:
        if translate_category_name(cat).strip().lower() == displayed_value.strip().lower():
            return cat
    return displayed_value


def resolve_difficulty_input(displayed_value, difficulty_options):
    """Équivalent de resolve_category_input() pour le niveau de
    difficulté."""
    if not displayed_value:
        return displayed_value
    for diff in difficulty_options:
        if diff.strip().lower() == displayed_value.strip().lower():
            return diff
    for diff in difficulty_options:
        if translate_difficulty_name(diff).strip().lower() == displayed_value.strip().lower():
            return diff
    return displayed_value


# Table de correspondance pour les options du menu déroulant de tri des
# recettes (RECIPE_SORT_OPTIONS), même principe que les catégories et
# difficultés : liste fermée en lecture seule, la donnée réelle comparée
# dans recipe_sort_key() reste toujours la valeur française d'origine.
SORT_OPTION_TRANSLATIONS = {
    "en": {
        "nom (a-z)": "Name (A-Z)",
        "temps de préparation": "Prep time",
        "difficulté": "Difficulty",
        "note": "Rating",
        "ajoutées récemment": "Recently added",
    },
    "es": {
        "nom (a-z)": "Nombre (A-Z)",
        "temps de préparation": "Tiempo de preparación",
        "difficulté": "Dificultad",
        "note": "Valoración",
        "ajoutées récemment": "Añadidas recientemente",
    },
    "de": {
        "nom (a-z)": "Name (A-Z)",
        "temps de préparation": "Zubereitungszeit",
        "difficulté": "Schwierigkeit",
        "note": "Bewertung",
        "ajoutées récemment": "Kürzlich hinzugefügt",
    },
}


def translate_sort_option(option):
    """Retourne le nom d'affichage d'une option de tri de recette dans la
    langue actuellement sélectionnée. Ne change jamais la valeur comparée
    dans recipe_sort_key()."""
    if not option:
        return option
    if CURRENT_LANGUAGE == "fr":
        return option
    return SORT_OPTION_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(option.strip().lower(), option)


def resolve_sort_option_input(displayed_value, sort_options):
    """Résout une option de tri sélectionnée dans le menu déroulant
    (éventuellement affichée traduite) vers sa valeur canonique française
    exacte. Même principe que resolve_category_input() : liste fermée,
    donc sans ambiguïté."""
    if not displayed_value:
        return displayed_value
    for opt in sort_options:
        if opt.strip().lower() == displayed_value.strip().lower():
            return opt
    for opt in sort_options:
        if translate_sort_option(opt).strip().lower() == displayed_value.strip().lower():
            return opt
    return displayed_value


# Table de correspondance pour les unités de mesure des ingrédients
# (UNIT_OPTIONS, PRICE_UNIT_OPTIONS et les quelques unités supplémentaires
# de contenant utilisées dans le garde-manger et l'ajout manuel à la
# liste de courses). Les unités du système métrique (Gr, Kilo, cl, Litre)
# utilisent le même symbole dans les 4 langues (g/kg/cl/L partout dans le
# monde) : elles ne figurent donc pas ici, translate_unit_name() les
# laisse simplement inchangées comme n'importe quel mot sans traduction
# connue. Seuls les mots qui diffèrent réellement d'une langue à l'autre
# ont une entrée.
UNIT_TRANSLATIONS = {
    "en": {
        "gr": "g",
        "kilo": "kg",
        "litre": "L",
        "pièce": "piece",
        "cuillère à soupe": "tbsp",
        "cuillère à café": "tsp",
        "autre": "other",
        "boîte": "can",
        "paquet": "pack",
        "rouleau": "roll",
        "bouteille": "bottle",
    },
    "es": {
        "gr": "g",
        "kilo": "kg",
        "litre": "L",
        "pièce": "unidad",
        "cuillère à soupe": "cucharada",
        "cuillère à café": "cucharadita",
        "autre": "otro",
        "boîte": "lata",
        "paquet": "paquete",
        "rouleau": "rollo",
        "bouteille": "botella",
    },
    "de": {
        "gr": "g",
        "kilo": "kg",
        "litre": "L",
        "pièce": "Stück",
        "cuillère à soupe": "EL",
        "cuillère à café": "TL",
        "autre": "andere",
        "boîte": "Dose",
        "paquet": "Packung",
        "rouleau": "Rolle",
        "bouteille": "Flasche",
    },
}


def translate_unit_name(unit):
    """Retourne le nom d'affichage d'une unité de mesure dans la langue
    actuellement sélectionnée. Les unités du système métrique (Gr, Kilo,
    cl, Litre, kg, L) n'ont pas d'entrée dans la table et restent donc
    inchangées : ce sont les mêmes symboles dans toutes les langues
    disponibles. Ne change jamais la donnée réelle stockée dans la
    recette ou le prix d'un ingrédient."""
    if not unit:
        return unit
    if CURRENT_LANGUAGE == "fr":
        return unit
    return UNIT_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(unit.strip().lower(), unit)


def resolve_unit_input(displayed_value, unit_options):
    """Résout une unité sélectionnée dans un menu déroulant fermé
    (état 'readonly', donc sans ambiguïté) vers sa valeur canonique
    française exacte. Pour les champs d'unité en texte libre (garde-
    manger, ajout manuel), voir resolve_unit_input_best_effort()
    ci-dessous à la place."""
    if not displayed_value:
        return displayed_value
    for u in unit_options:
        if u.strip().lower() == displayed_value.strip().lower():
            return u
    for u in unit_options:
        if translate_unit_name(u).strip().lower() == displayed_value.strip().lower():
            return u
    return displayed_value


def resolve_unit_input_best_effort(typed_value, unit_options):
    """Équivalent de resolve_unit_input(), mais pour les champs d'unité en
    texte libre : si la valeur tapée ou sélectionnée correspond à une
    unité connue (en français ou traduite), la résout vers son nom
    canonique français ; sinon, la retourne telle quelle, pour ne pas
    entraver la saisie d'une unité personnalisée absente de la liste
    (ex. « sachet »)."""
    if not typed_value:
        return typed_value
    return resolve_unit_input(typed_value, unit_options)


# Tables de correspondance pour les jours de la semaine et les créneaux de
# repas, par langue. Contrairement aux catégories/difficultés, ces
# valeurs servent aussi de CLÉS de dictionnaire dans le planning de la
# semaine et son historique (weekly_plan.json, weekly_plan_history.json,
# weekly_plan_templates.json) : ces tables ne doivent JAMAIS être
# utilisées pour changer une clé de stockage, uniquement pour ce qui est
# affiché à l'écran ou écrit dans un export (planning ICS, PDF...).
WEEKDAY_TRANSLATIONS = {
    "en": {
        "lundi": "Monday",
        "mardi": "Tuesday",
        "mercredi": "Wednesday",
        "jeudi": "Thursday",
        "vendredi": "Friday",
        "samedi": "Saturday",
        "dimanche": "Sunday",
    },
    "es": {
        "lundi": "Lunes",
        "mardi": "Martes",
        "mercredi": "Miércoles",
        "jeudi": "Jueves",
        "vendredi": "Viernes",
        "samedi": "Sábado",
        "dimanche": "Domingo",
    },
    "de": {
        "lundi": "Montag",
        "mardi": "Dienstag",
        "mercredi": "Mittwoch",
        "jeudi": "Donnerstag",
        "vendredi": "Freitag",
        "samedi": "Samstag",
        "dimanche": "Sonntag",
    },
}

MEALSLOT_TRANSLATIONS = {
    "en": {
        "petit-déjeuner": "Breakfast",
        "déjeuner — entrée": "Lunch — Starter",
        "déjeuner — plat": "Lunch — Main",
        "déjeuner — dessert": "Lunch — Dessert",
        "dîner — entrée": "Dinner — Starter",
        "dîner — plat": "Dinner — Main",
        "dîner — dessert": "Dinner — Dessert",
    },
    "es": {
        "petit-déjeuner": "Desayuno",
        "déjeuner — entrée": "Almuerzo — Entrante",
        "déjeuner — plat": "Almuerzo — Plato principal",
        "déjeuner — dessert": "Almuerzo — Postre",
        "dîner — entrée": "Cena — Entrante",
        "dîner — plat": "Cena — Plato principal",
        "dîner — dessert": "Cena — Postre",
    },
    "de": {
        "petit-déjeuner": "Frühstück",
        "déjeuner — entrée": "Mittagessen — Vorspeise",
        "déjeuner — plat": "Mittagessen — Hauptgericht",
        "déjeuner — dessert": "Mittagessen — Dessert",
        "dîner — entrée": "Abendessen — Vorspeise",
        "dîner — plat": "Abendessen — Hauptgericht",
        "dîner — dessert": "Abendessen — Dessert",
    },
}


def translate_weekday_name(day):
    """Retourne le nom d'affichage d'un jour de la semaine dans la langue
    actuellement sélectionnée. Ne doit jamais servir à changer une clé de
    stockage : le planning de la semaine et son historique restent
    toujours indexés par le nom français."""
    if not day:
        return day
    if CURRENT_LANGUAGE == "fr":
        return day
    return WEEKDAY_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(day.strip().lower(), day)


def translate_mealslot_name(slot):
    """Retourne le nom d'affichage d'un créneau de repas dans la langue
    actuellement sélectionnée. Même principe que translate_weekday_name()
    : jamais utilisée pour changer une clé de stockage."""
    if not slot:
        return slot
    if CURRENT_LANGUAGE == "fr":
        return slot
    return MEALSLOT_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(slot.strip().lower(), slot)


def rating_stars(n):
    """Retourne une représentation textuelle d'une note sur 5 étoiles,
    ex: rating_stars(3) -> '★★★☆☆'."""
    try:
        n = int(n or 0)
    except (TypeError, ValueError):
        n = 0
    n = max(0, min(5, n))
    return "★" * n + "☆" * (5 - n)


def recipe_matches_search(recipe, search_key):
    """Vérifie si une recette correspond à une recherche (déjà normalisée via
    ingredient_sort_key), en comparant le nom ET les étiquettes."""
    if not search_key:
        return True
    if search_key in ingredient_sort_key(recipe["name"]):
        return True
    for tag in recipe.get("tags", []):
        if search_key in ingredient_sort_key(tag):
            return True
    return False


RECIPE_SORT_OPTIONS = ["Nom (A-Z)", "Temps de préparation", "Difficulté", "Note", "Ajoutées récemment"]
_DIFFICULTY_ORDER = {"Facile": 1, "Moyen": 2, "Difficile": 3}


def find_similar_recipes(recipe, all_recipes, limit=5):
    """Suggère des recettes proches d'une recette donnée, en se basant sur
    la même catégorie, les étiquettes en commun et les ingrédients en
    commun (chacun apportant des points, la catégorie comptant le plus).
    Ne retourne que des recettes avec au moins un point de similarité, et
    jamais la recette elle-même."""
    target_name = recipe.get("name")
    target_category = recipe.get("category", "Autre")
    target_tags = {ingredient_sort_key(t) for t in recipe.get("tags", [])}
    target_ingredients = {ingredient_sort_key(ing["name"]) for ing in recipe.get("ingredients", [])}

    scored = []
    for other in all_recipes:
        if other.get("name") == target_name:
            continue
        score = 0
        if other.get("category", "Autre") == target_category:
            score += 2
        other_tags = {ingredient_sort_key(t) for t in other.get("tags", [])}
        score += len(target_tags & other_tags)
        other_ingredients = {ingredient_sort_key(ing["name"]) for ing in other.get("ingredients", [])}
        score += min(len(target_ingredients & other_ingredients), 5)
        if score > 0:
            scored.append((score, other))

    scored.sort(key=lambda pair: (-pair[0], ingredient_sort_key(pair[1]["name"])))
    return [r for score, r in scored[:limit]]


def recipe_sort_key(recipe, option):
    """Retourne une clé de tri pour une recette selon l'option choisie parmi
    RECIPE_SORT_OPTIONS."""
    if option == "Temps de préparation":
        try:
            return float(recipe.get("prep_time") or 0)
        except (TypeError, ValueError):
            return 0.0
    if option == "Difficulté":
        return _DIFFICULTY_ORDER.get(recipe.get("difficulty"), 0)
    if option == "Note":
        return -int(recipe.get("rating", 0) or 0)  # négatif : meilleure note en premier
    if option == "Ajoutées récemment":
        return recipe.get("created_at") or ""
    return ingredient_sort_key(recipe["name"])


def format_recipe_list_label(recipe):
    """Construit le libellé affiché pour une recette dans les listes de
    sélection : favori, catégorie, nom, note, et — pour qu'on les voie tout
    de suite lors d'un tri — le temps total, la difficulté et les
    allergènes éventuels."""
    cat = translate_category_name(recipe.get("category", "Autre"))
    star = "⭐ " if recipe.get("favorite") else ""
    wish = "💭 " if recipe.get("wishlist") else ""
    rating = recipe.get("rating", 0)
    rating_suffix = f" {rating_stars(rating)}" if rating else ""
    label = f"{star}{wish}[{cat}] {recipe['name']}{rating_suffix}"

    info_bits = []
    prep = recipe.get("prep_time")
    cook = recipe.get("cook_time")
    if prep or cook:
        try:
            total = float(prep or 0) + float(cook or 0)
            total_display = int(total) if total == int(total) else total
        except (TypeError, ValueError):
            total_display = None
        if total_display is not None:
            info_bits.append(f"{total_display} min")
    difficulty = recipe.get("difficulty")
    if difficulty:
        info_bits.append(translate_difficulty_name(difficulty))
    allergens = recipe.get("allergens") or []
    if allergens:
        info_bits.append(f"⚠ {', '.join(translate_allergen_name(a) for a in allergens)}")
    if info_bits:
        label += f"  ({' · '.join(info_bits)})"
    return label


def load_recipes():
    """Charge les recettes depuis le fichier JSON local."""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def save_recipes(recipes):
    """Sauvegarde la liste des recettes dans le fichier JSON local."""
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(recipes, f, ensure_ascii=False, indent=2)


def load_ingredients():
    """Charge la liste des ingrédients connus (triée, sans doublons)."""
    if os.path.exists(INGREDIENTS_FILE):
        try:
            with open(INGREDIENTS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return sorted(dict.fromkeys(normalize_oe(i) for i in data), key=ingredient_sort_key)
        except Exception:
            return []
    return []


def save_ingredients(ingredients):
    """Sauvegarde la liste des ingrédients connus (triée, sans doublons)."""
    cleaned_map = {}
    for i in ingredients:
        name = normalize_oe(i.strip())
        if name:
            cleaned_map.setdefault(name.lower(), name)
    cleaned = sorted(cleaned_map.values(), key=ingredient_sort_key)
    with open(INGREDIENTS_FILE, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, ensure_ascii=False, indent=2)
    return cleaned


def sync_ingredients_from_recipes():
    """S'assure que tout ingrédient déjà utilisé dans une recette existante
    figure bien dans la liste des ingrédients connus (utile lors de la
    première utilisation de cette fonctionnalité, ou après import de
    données). Lors du tout premier lancement (aucun ingredients.json), la
    liste des ~1000 ingrédients courants fournis avec l'application est
    utilisée comme point de départ."""
    first_run = not os.path.exists(INGREDIENTS_FILE)
    known = load_default_ingredients() if first_run else load_ingredients()
    known_lower = {i.lower() for i in known}
    changed = first_run
    for recipe in load_recipes():
        for ing in recipe.get("ingredients", []):
            name = ing.get("name", "").strip()
            if name and name.lower() not in known_lower:
                known.append(name)
                known_lower.add(name.lower())
                changed = True
    if changed:
        return save_ingredients(known)
    return known


def merge_default_ingredients():
    """Ajoute à la liste actuelle les ingrédients courants fournis avec
    l'application qui ne seraient pas déjà présents. Retourne le nombre
    d'ingrédients réellement ajoutés."""
    current = load_ingredients()
    current_lower = {i.lower() for i in current}
    defaults = load_default_ingredients()
    added = 0
    for name in defaults:
        if name.lower() not in current_lower:
            current.append(name)
            current_lower.add(name.lower())
            added += 1
    save_ingredients(current)
    return added


def rename_ingredient_everywhere(old_name, new_name):
    """Renomme un ingrédient dans toutes les recettes qui l'utilisent."""
    recipes = load_recipes()
    changed = False
    for recipe in recipes:
        for ing in recipe.get("ingredients", []):
            if ing.get("name", "").strip().lower() == old_name.strip().lower():
                ing["name"] = new_name
                changed = True
    if changed:
        save_recipes(recipes)


def count_ingredient_usage(name):
    """Compte le nombre de recettes utilisant cet ingrédient."""
    count = 0
    for recipe in load_recipes():
        for ing in recipe.get("ingredients", []):
            if ing.get("name", "").strip().lower() == name.strip().lower():
                count += 1
                break
    return count


# ---------------------------------------------------------------------------
# Estimation du coût d'une recette : les prix sont renseignés par
# l'utilisateur (aucune source de prix en ligne n'est disponible/fiable pour
# une application locale), un ingrédient par un, dans "Gérer les prix".
# ---------------------------------------------------------------------------

PRICE_UNIT_OPTIONS = ["kg", "L", "pièce", "cuillère à soupe", "cuillère à café"]


def load_ingredient_prices():
    """Retourne {nom_ingredient_en_minuscules: {"name":, "price":, "unit":}}."""
    if os.path.exists(INGREDIENT_PRICES_FILE):
        try:
            with open(INGREDIENT_PRICES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            return {}
    return {}


def save_ingredient_prices(prices):
    with open(INGREDIENT_PRICES_FILE, "w", encoding="utf-8") as f:
        json.dump(prices, f, ensure_ascii=False, indent=2)


def get_ingredient_price(name):
    return load_ingredient_prices().get(name.strip().lower())


def set_ingredient_price(name, price, unit):
    """price=None efface le prix enregistré pour cet ingrédient."""
    prices = load_ingredient_prices()
    key = name.strip().lower()
    if price is None:
        prices.pop(key, None)
    else:
        prices[key] = {"name": name.strip(), "price": price, "unit": unit}
    save_ingredient_prices(prices)


def compute_recipe_cost(recipe, persons):
    """Retourne (coût_total_estimé, nb_ingrédients_avec_prix_connu,
    nb_ingrédients_total). Un ingrédient sans prix renseigné, ou dont
    l'unité ne peut pas être convertie vers celle du prix, est ignoré
    (le coût rendu est donc une estimation partielle si des prix manquent)."""
    prices = load_ingredient_prices()
    total = 0.0
    known = 0
    total_count = len(recipe["ingredients"])
    for ing in recipe["ingredients"]:
        price_info = prices.get(ing["name"].strip().lower())
        if not price_info:
            continue
        qty = ing["quantity"] * persons
        unit_lower = ing["unit"].strip().lower()
        price = price_info["price"]
        price_unit = price_info["unit"]
        contribution = None
        if price_unit == "kg" and unit_lower in ("gr", "g", "gramme", "grammes"):
            contribution = (qty / 1000.0) * price
        elif price_unit == "L" and unit_lower == "cl":
            contribution = (qty / 100.0) * price
        elif price_unit.lower() == unit_lower:
            contribution = qty * price
        if contribution is not None:
            total += contribution
            known += 1
    return total, known, total_count


# ---------------------------------------------------------------------------
# Estimation des valeurs nutritionnelles d'une recette, à partir d'une base
# de valeurs typiques par ingrédient (kcal / protéines / glucides / lipides
# pour 100 g ou 100 ml), fournie avec l'application.
# ---------------------------------------------------------------------------

_nutrition_cache = None

# Équivalence approximative de chaque unité de recette vers des grammes (ou
# millilitres, assimilés à des grammes pour les liquides — approximation
# standard). Les unités "pièce" et "autre" ne sont pas converties : le poids
# d'une "pièce" dépend trop de l'ingrédient pour être généralisé.
UNIT_TO_GRAMS = {
    "gr": 1.0, "g": 1.0, "gramme": 1.0, "grammes": 1.0,
    "kilo": 1000.0, "kg": 1000.0,
    "cl": 10.0,
    "litre": 1000.0, "l": 1000.0,
    "cuillère à soupe": 15.0,
    "cuillère à café": 5.0,
}


PANTRY_FILE = os.path.join(BASE_DIR, "pantry.json")


def load_pantry():
    """Charge le garde-manger : dict {clé normalisée: {"name","quantity","unit"}}."""
    if os.path.exists(PANTRY_FILE):
        try:
            with open(PANTRY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}


def save_pantry(pantry):
    with open(PANTRY_FILE, "w", encoding="utf-8") as f:
        json.dump(pantry, f, ensure_ascii=False, indent=2)


def set_pantry_item(name, quantity, unit, threshold=None):
    pantry = load_pantry()
    pantry[ingredient_sort_key(name)] = {
        "name": name, "quantity": quantity, "unit": unit, "threshold": threshold
    }
    save_pantry(pantry)
    return pantry


def get_low_stock_pantry_items():
    """Retourne les articles du garde-manger dont la quantité est passée
    sous le seuil d'alerte défini par l'utilisateur (uniquement ceux où un
    seuil a été renseigné)."""
    pantry = load_pantry()
    return [
        entry for entry in pantry.values()
        if entry.get("threshold") is not None and entry.get("quantity", 0) < entry["threshold"]
    ]


def remove_pantry_item(name):
    pantry = load_pantry()
    pantry.pop(ingredient_sort_key(name), None)
    save_pantry(pantry)
    return pantry


def convert_to_grams_equivalent(quantity, unit):
    """Convertit une quantité vers un équivalent en grammes si l'unité est
    connue (voir UNIT_TO_GRAMS), sinon retourne None (comparaison impossible,
    par exemple pour « pièce » ou une unité personnalisée)."""
    factor = UNIT_TO_GRAMS.get((unit or "").strip().lower())
    if factor is None:
        return None
    try:
        return float(quantity) * factor
    except (TypeError, ValueError):
        return None


def pantry_stock_status(ingredient_name, needed_qty, needed_unit, pantry):
    """Compare la quantité nécessaire d'un ingrédient à celle disponible dans
    le garde-manger. Retourne :
    - "absent" : l'ingrédient n'est pas du tout dans le garde-manger
    - "suffisant" : la quantité en stock couvre le besoin
    - "insuffisant" : l'ingrédient est en stock mais pas en quantité suffisante
    - "inconnu" : l'ingrédient est en stock mais les unités ne sont pas
      comparables (ex. « pièce » contre « Gr »), impossible de conclure
    """
    entry = pantry.get(ingredient_sort_key(ingredient_name))
    if entry is None:
        return "absent"
    have_qty, have_unit = entry.get("quantity", 0), entry.get("unit", "")
    if (have_unit or "").strip().lower() == (needed_unit or "").strip().lower():
        return "suffisant" if have_qty >= needed_qty else "insuffisant"
    have_grams = convert_to_grams_equivalent(have_qty, have_unit)
    needed_grams = convert_to_grams_equivalent(needed_qty, needed_unit)
    if have_grams is None or needed_grams is None:
        return "inconnu"
    return "suffisant" if have_grams >= needed_grams else "insuffisant"


def decrement_pantry_for_recipe(recipe, persons):
    """Décompte du garde-manger les ingrédients d'une recette qui viennent
    d'être cuisinée, dans la limite du raisonnable : ne fait rien pour un
    ingrédient absent du garde-manger ou dont l'unité n'est pas comparable
    (ne devine jamais), et ne descend jamais sous zéro. Retourne le nombre
    d'ingrédients réellement décomptés."""
    pantry = load_pantry()
    default_persons = recipe.get("default_persons", 1) or 1
    try:
        ratio = float(persons) / float(default_persons)
    except (TypeError, ValueError, ZeroDivisionError):
        ratio = 1
    decremented = 0
    for ing in recipe.get("ingredients", []):
        key = ingredient_sort_key(ing["name"])
        entry = pantry.get(key)
        if entry is None:
            continue
        needed_qty = ing["quantity"] * ratio
        have_unit = (entry.get("unit") or "").strip().lower()
        needed_unit = (ing.get("unit") or "").strip().lower()
        if have_unit == needed_unit:
            entry["quantity"] = max(0, entry.get("quantity", 0) - needed_qty)
            decremented += 1
            continue
        have_grams = convert_to_grams_equivalent(entry.get("quantity", 0), entry.get("unit", ""))
        needed_grams = convert_to_grams_equivalent(needed_qty, ing.get("unit", ""))
        if have_grams is not None and needed_grams is not None:
            new_grams = max(0, have_grams - needed_grams)
            # Reconvertit dans l'unité d'origine du garde-manger (toujours
            # en grammes-équivalent ici, donc conversion directe possible).
            factor = UNIT_TO_GRAMS.get(have_unit, 1.0)
            entry["quantity"] = new_grams / factor if factor else new_grams
            decremented += 1
    if decremented:
        save_pantry(pantry)
    return decremented


def load_nutrition_data():
    """Charge (une seule fois, en cache) la base de valeurs nutritionnelles
    fournie avec l'application."""
    global _nutrition_cache
    if _nutrition_cache is not None:
        return _nutrition_cache
    data = {}
    if os.path.exists(NUTRITION_DATA_FILE):
        try:
            with open(NUTRITION_DATA_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
                if isinstance(raw, dict):
                    data = {k.strip().lower(): v for k, v in raw.items()}
        except Exception:
            data = {}
    _nutrition_cache = data
    return _nutrition_cache


def get_ingredient_nutrition(name):
    """Retourne le dict nutrition {kcal, protein_g, carbs_g, fat_g} pour un
    ingrédient (pour 100 g/100 ml), ou None si inconnu. Une surcharge
    personnelle est prioritaire sur la base fournie."""
    override = get_ingredient_override(name)
    if override and "nutrition" in override and override["nutrition"]:
        return override["nutrition"]
    return load_nutrition_data().get(name.strip().lower())


def compute_recipe_nutrition(recipe, persons):
    """Retourne (totaux {kcal, protein_g, carbs_g, fat_g}, nb_ingrédients pris
    en compte, nb_ingrédients_total). Les ingrédients inconnus de la base, ou
    exprimés en "pièce"/unité personnalisée, sont exclus du total (estimation
    partielle dans ce cas)."""
    totals = {"kcal": 0.0, "protein_g": 0.0, "carbs_g": 0.0, "fat_g": 0.0}
    converted = 0
    total_count = len(recipe["ingredients"])
    for ing in recipe["ingredients"]:
        nutri = get_ingredient_nutrition(ing["name"])
        if not nutri:
            continue
        grams_per_unit = UNIT_TO_GRAMS.get(ing["unit"].strip().lower())
        if grams_per_unit is None:
            continue
        grams = ing["quantity"] * persons * grams_per_unit
        factor = grams / 100.0
        totals["kcal"] += nutri.get("kcal", 0) * factor
        totals["protein_g"] += nutri.get("protein_g", 0) * factor
        totals["carbs_g"] += nutri.get("carbs_g", 0) * factor
        totals["fat_g"] += nutri.get("fat_g", 0) * factor
        converted += 1
    return totals, converted, total_count


# ---------------------------------------------------------------------------
# Allergènes présents dans chaque ingrédient, à partir d'une base fournie
# avec l'application (les ~1000 ingrédients courants). Sert à détecter
# automatiquement les allergènes d'une recette à partir de ses ingrédients.
# ---------------------------------------------------------------------------

_ingredient_allergens_cache = None


def load_ingredient_allergens():
    """Charge (une seule fois, en cache) la base des allergènes par
    ingrédient fournie avec l'application."""
    global _ingredient_allergens_cache
    if _ingredient_allergens_cache is not None:
        return _ingredient_allergens_cache
    data = {}
    if os.path.exists(INGREDIENT_ALLERGENS_FILE):
        try:
            with open(INGREDIENT_ALLERGENS_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
                if isinstance(raw, dict):
                    data = {k.strip().lower(): v for k, v in raw.items()}
        except Exception:
            data = {}
    _ingredient_allergens_cache = data
    return _ingredient_allergens_cache


# ---------------------------------------------------------------------------
# Surcharges personnelles par ingrédient (allergènes et/ou valeurs
# nutritionnelles modifiés par l'utilisateur, ou définis pour un ingrédient
# créé par lui). Prioritaires sur les bases fournies avec l'application.
# ---------------------------------------------------------------------------

def load_ingredient_overrides():
    if os.path.exists(INGREDIENT_OVERRIDES_FILE):
        try:
            with open(INGREDIENT_OVERRIDES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            return {}
    return {}


def save_ingredient_overrides(overrides):
    with open(INGREDIENT_OVERRIDES_FILE, "w", encoding="utf-8") as f:
        json.dump(overrides, f, ensure_ascii=False, indent=2)


def get_ingredient_override(name):
    return load_ingredient_overrides().get(name.strip().lower())


def set_ingredient_override(name, allergens=None, nutrition=None, substitutions=None):
    """allergens : liste (peut être vide) ou None pour ne pas y toucher.
    nutrition : dict {kcal, protein_g, carbs_g, fat_g} ou None pour ne pas
    y toucher (passer un dict vide {} pour effacer la surcharge nutrition).
    substitutions : liste de {"nom","note"} ou None pour ne pas y toucher
    (passer une liste vide [] pour effacer la surcharge et revenir à la
    base de substitutions fournie avec l'application, s'il y en a une)."""
    overrides = load_ingredient_overrides()
    key = name.strip().lower()
    entry = overrides.get(key, {})
    if allergens is not None:
        entry["allergens"] = allergens
    if nutrition is not None:
        entry["nutrition"] = nutrition
    if substitutions is not None:
        entry["substitutions"] = substitutions
    if entry:
        overrides[key] = entry
    else:
        overrides.pop(key, None)
    save_ingredient_overrides(overrides)


def rename_ingredient_override(old_name, new_name):
    """Transfère une éventuelle surcharge de l'ancien nom vers le nouveau,
    lors d'un renommage d'ingrédient."""
    overrides = load_ingredient_overrides()
    old_key = old_name.strip().lower()
    new_key = new_name.strip().lower()
    if old_key in overrides and old_key != new_key:
        overrides[new_key] = overrides.pop(old_key)
        save_ingredient_overrides(overrides)


_default_substitutions_cache = None


def load_default_substitutions():
    """Charge (une seule fois, en cache) la base de substitutions courantes
    fournie avec l'application."""
    global _default_substitutions_cache
    if _default_substitutions_cache is not None:
        return _default_substitutions_cache
    data = {}
    if os.path.exists(INGREDIENT_SUBSTITUTIONS_FILE):
        try:
            with open(INGREDIENT_SUBSTITUTIONS_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
                if isinstance(raw, dict):
                    data = {k.strip().lower(): v for k, v in raw.items()}
        except Exception:
            data = {}
    _default_substitutions_cache = data
    return _default_substitutions_cache


_ingredient_translations_cache = {}


def load_ingredient_translations(lang):
    """Charge (une seule fois par langue, en cache) le dictionnaire de
    correspondance des ~1000 ingrédients courants vers la langue donnée,
    fourni avec l'application. Un ingrédient absent de ce dictionnaire
    (par exemple un ingrédient personnalisé ajouté par l'utilisateur)
    n'a simplement pas de traduction : voir translate_ingredient_name()."""
    if lang in _ingredient_translations_cache:
        return _ingredient_translations_cache[lang]
    data = {}
    file_path = INGREDIENT_TRANSLATIONS_FILES.get(lang)
    if file_path and os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
                if isinstance(raw, dict):
                    data = {k.strip().lower(): v for k, v in raw.items()}
        except Exception:
            data = {}
    _ingredient_translations_cache[lang] = data
    return data


def translate_ingredient_name(name):
    """Retourne le nom d'affichage d'un ingrédient dans la langue
    actuellement sélectionnée. La donnée réelle (recherche, tri,
    comparaison, clé de stockage des allergènes/prix/substituts...) reste
    TOUJOURS le nom français d'origine, quelle que soit la langue de
    l'interface : cette fonction ne change que ce qui est affiché à
    l'écran, jamais ce qui est enregistré ou comparé en interne. Un
    ingrédient personnalisé sans traduction connue s'affiche simplement
    dans son nom français d'origine, comme le reste de l'interface pas
    encore traduite."""
    if not name:
        return name
    if CURRENT_LANGUAGE == "fr":
        return name
    translations = load_ingredient_translations(CURRENT_LANGUAGE)
    translated = translations.get(name.strip().lower())
    return translated if translated else name


_ingredient_reverse_translations_cache = {}


def load_ingredient_reverse_translations(lang):
    """Construit (une seule fois par langue, en cache) la correspondance
    inverse vers le français, à partir du dictionnaire de traduction de
    la langue donnée. Certains noms traduits correspondent à plusieurs
    ingrédients français distincts (ex. « peanut » pour « Arachide » et
    « Cacahuète ») : dans ce cas, le nom français le plus court est
    choisi par convention (à égalité, ordre alphabétique), comme candidat
    le plus généraliste — une approximation raisonnable plutôt qu'une
    ambiguïté bloquante."""
    if lang in _ingredient_reverse_translations_cache:
        return _ingredient_reverse_translations_cache[lang]
    forward = load_ingredient_translations(lang)
    grouped = {}
    for fr_lower, translated_name in forward.items():
        grouped.setdefault(translated_name.strip().lower(), []).append(fr_lower)
    reverse = {}
    for translated_lower, fr_list in grouped.items():
        fr_list.sort(key=lambda s: (len(s), s))
        reverse[translated_lower] = fr_list[0]
    _ingredient_reverse_translations_cache[lang] = reverse
    return reverse


def resolve_ingredient_input(typed_name, ingredient_names):
    """Résout un nom d'ingrédient tapé par l'utilisateur vers son nom
    canonique français exact (tel qu'il apparaît dans ingredient_names),
    quelle que soit la langue dans laquelle il a été tapé. Essaie d'abord
    une correspondance directe (le français reste toujours la langue de
    référence des données), puis, si l'interface n'est pas en français,
    une correspondance via le dictionnaire de traduction inverse de la
    langue courante. Retourne None si rien ne correspond, exactement
    comme l'ancienne logique de correspondance stricte au nom français."""
    if not typed_name:
        return None
    typed_key = typed_name.strip().lower()
    if not typed_key:
        return None
    for n in ingredient_names:
        if n.strip().lower() == typed_key:
            return n
    if CURRENT_LANGUAGE != "fr":
        reverse = load_ingredient_reverse_translations(CURRENT_LANGUAGE)
        fr_key = reverse.get(typed_key)
        if fr_key:
            for n in ingredient_names:
                if n.strip().lower() == fr_key:
                    return n
    return None


def get_display_ingredient_values(ingredient_names):
    """Retourne la liste des noms d'ingrédients à proposer dans un champ
    de saisie, une liste déroulante ou un filtre, dans la langue
    actuellement sélectionnée (traduits quand une traduction est connue,
    sinon inchangés). Utilisé pour l'autocomplétion : le nom RÉELLEMENT
    sélectionné ou tapé doit ensuite toujours être résolu via
    resolve_ingredient_input() avant d'être enregistré."""
    if CURRENT_LANGUAGE == "fr":
        return list(ingredient_names)
    return [translate_ingredient_name(n) for n in ingredient_names]


def get_ingredient_substitutions(name):
    """Retourne la liste de substituts connus pour un ingrédient (une
    surcharge personnelle, si vous en avez défini une, est toujours
    prioritaire sur la base fournie avec l'application), ou une liste vide
    si aucun substitut n'est connu. Toujours en français : c'est la
    donnée de référence, utilisée pour toute comparaison ou correspondance
    interne (voir get_display_ingredient_substitutions() pour un
    affichage traduit)."""
    key = name.strip().lower()
    override = load_ingredient_overrides().get(key)
    if override is not None and "substitutions" in override:
        return override["substitutions"]
    return load_default_substitutions().get(key, [])


_default_substitutions_translations_cache = {}


def load_default_substitutions_translated(lang):
    """Charge (une seule fois par langue, en cache) la traduction de la
    base de substitutions courante fournie avec l'application, dans la
    langue donnée."""
    if lang in _default_substitutions_translations_cache:
        return _default_substitutions_translations_cache[lang]
    data = {}
    file_path = INGREDIENT_SUBSTITUTIONS_TRANSLATION_FILES.get(lang)
    if file_path and os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
                if isinstance(raw, dict):
                    data = {k.strip().lower(): v for k, v in raw.items()}
        except Exception:
            data = {}
    _default_substitutions_translations_cache[lang] = data
    return data


def get_display_ingredient_substitutions(name):
    """Retourne la liste de substituts connus pour un ingrédient, dans la
    langue actuellement sélectionnée, pour AFFICHAGE uniquement — ne
    jamais utiliser cette fonction pour une comparaison ou une
    correspondance interne (utiliser get_ingredient_substitutions() dans
    ce cas, qui reste toujours en français). Une surcharge personnelle
    (substituts que vous avez vous-même ajoutés ou modifiés) est toujours
    utilisée telle quelle, jamais traduite automatiquement, puisqu'elle a
    été tapée par vous dans la langue de votre choix."""
    key = name.strip().lower()
    override = load_ingredient_overrides().get(key)
    if override is not None and "substitutions" in override:
        return override["substitutions"]
    if CURRENT_LANGUAGE != "fr":
        translated = load_default_substitutions_translated(CURRENT_LANGUAGE).get(key)
        if translated is not None:
            return translated
    return load_default_substitutions().get(key, [])


def has_known_substitutions():
    """Ensemble des clés normalisées de tous les ingrédients ayant au moins
    un substitut connu (base fournie ou surcharge personnelle)."""
    keys = set(load_default_substitutions().keys())
    overrides = load_ingredient_overrides()
    for key, entry in overrides.items():
        if entry.get("substitutions"):
            keys.add(key)
        elif "substitutions" in entry and not entry["substitutions"]:
            keys.discard(key)  # surcharge vide = substituts explicitement effacés
    return keys


def revert_ingredient_substitutions_to_default(name):
    """Retire complètement la surcharge de substituts d'un ingrédient (sans
    toucher à ses éventuelles surcharges d'allergènes/nutrition), pour
    revenir à ce que propose la base fournie avec l'application — à ne pas
    confondre avec le fait d'enregistrer une liste vide, qui bloquerait au
    contraire tout substitut pour cet ingrédient."""
    overrides = load_ingredient_overrides()
    key = name.strip().lower()
    if key in overrides and "substitutions" in overrides[key]:
        del overrides[key]["substitutions"]
        if not overrides[key]:
            overrides.pop(key)
        save_ingredient_overrides(overrides)


def get_ingredient_allergens(name):
    """Retourne la liste des allergènes connus pour un ingrédient donné
    (liste vide si l'ingrédient n'est pas reconnu ou n'en contient aucun).
    Une surcharge personnelle est prioritaire sur la base fournie."""
    override = get_ingredient_override(name)
    if override and "allergens" in override:
        return override["allergens"]
    return load_ingredient_allergens().get(name.strip().lower(), [])


def compute_recipe_allergens(ingredients):
    """Retourne l'ensemble des allergènes détectés à partir d'une liste
    d'ingrédients de recette, dans l'ordre de la liste ALLERGENS."""
    detected = set()
    for ing in ingredients:
        detected.update(get_ingredient_allergens(ing["name"]))
    return [a for a in ALLERGENS if a in detected]


def find_similar_ingredient_pairs(names, threshold=0.82):
    """Retourne une liste de paires (nom_a, nom_b, score) d'ingrédients dont
    les noms se ressemblent fortement (accents/casse ignorés), pouvant
    indiquer un doublon ou une faute de frappe (ex. "Tomate" / "Tomates",
    "Echalotte" / "Échalote"). Les ingrédients sont d'abord regroupés par
    leurs deux premières lettres pour limiter le nombre de comparaisons sur
    de grandes listes."""
    normalized = [(n, ingredient_sort_key(n)) for n in names]
    buckets = {}
    for name, key in normalized:
        prefix = key[:2] if len(key) >= 2 else key
        buckets.setdefault(prefix, []).append((name, key))

    pairs = []
    for bucket in buckets.values():
        for i in range(len(bucket)):
            name_a, key_a = bucket[i]
            for j in range(i + 1, len(bucket)):
                name_b, key_b = bucket[j]
                if key_a == key_b:
                    continue
                is_plural_variant = (
                    key_a in (key_b + "s", key_b + "x") or key_b in (key_a + "s", key_a + "x")
                )
                ratio = difflib.SequenceMatcher(None, key_a, key_b).ratio()
                if is_plural_variant or ratio >= threshold:
                    pairs.append((name_a, name_b, ratio if not is_plural_variant else max(ratio, 0.9)))

    pairs.sort(key=lambda t: -t[2])
    return pairs


DISMISSED_DUPLICATE_PAIRS_FILE = os.path.join(BASE_DIR, "ingredient_dismissed_pairs.json")


def load_dismissed_pairs():
    """Retourne l'ensemble des paires d'ingrédients que l'utilisateur a
    explicitement indiquées comme n'étant PAS des doublons, pour ne plus les
    proposer lors des prochaines analyses. Chaque paire est représentée par
    un tuple trié de deux clés normalisées, indépendant de l'ordre."""
    if os.path.exists(DISMISSED_DUPLICATE_PAIRS_FILE):
        try:
            with open(DISMISSED_DUPLICATE_PAIRS_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
                if isinstance(raw, list):
                    return {
                        tuple(sorted(pair)) for pair in raw
                        if isinstance(pair, list) and len(pair) == 2
                    }
        except Exception:
            pass
    return set()


def save_dismissed_pairs(pairs_set):
    with open(DISMISSED_DUPLICATE_PAIRS_FILE, "w", encoding="utf-8") as f:
        json.dump([list(pair) for pair in sorted(pairs_set)], f, ensure_ascii=False, indent=2)


def add_dismissed_pair(name_a, name_b):
    pairs = load_dismissed_pairs()
    pairs.add(tuple(sorted([ingredient_sort_key(name_a), ingredient_sort_key(name_b)])))
    save_dismissed_pairs(pairs)


def is_pair_dismissed(name_a, name_b, dismissed_pairs):
    return tuple(sorted([ingredient_sort_key(name_a), ingredient_sort_key(name_b)])) in dismissed_pairs


def find_plural_duplicate(name, existing_names):
    """Retourne le nom déjà présent dans existing_names qui n'est qu'une
    variante singulier/pluriel de `name` (même règle que le vérificateur de
    doublons : suffixe "s" ou "x"), ou None si aucune correspondance. Sert à
    empêcher qu'un même ingrédient se retrouve deux fois dans la liste sous
    deux graphies différentes (ex. "Tomate" et "Tomates")."""
    key = ingredient_sort_key(name)
    for existing in existing_names:
        existing_key = ingredient_sort_key(existing)
        if existing_key == key:
            continue  # correspondance exacte : gérée séparément
        if key in (existing_key + "s", existing_key + "x") or existing_key in (key + "s", key + "x"):
            return existing
    return None


def copy_image_to_store(source_path):
    """Copie une image choisie par l'utilisateur dans le dossier images/
    et retourne le nom de fichier généré (à stocker dans la recette)."""
    ext = os.path.splitext(source_path)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"):
        ext = ".png"
    new_filename = f"{uuid.uuid4().hex}{ext}"
    dest_path = os.path.join(IMAGES_DIR, new_filename)
    shutil.copy2(source_path, dest_path)
    return new_filename


def delete_image_file(image_filename):
    """Supprime un fichier image du dossier images/ (si présent)."""
    if not image_filename:
        return
    path = os.path.join(IMAGES_DIR, image_filename)
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass


def load_thumbnail(image_filename, size=(240, 180)):
    """Retourne un objet ImageTk.PhotoImage pour affichage, ou None si
    l'image n'existe pas ou si Pillow n'est pas installé."""
    if not image_filename or not PIL_AVAILABLE:
        return None
    path = os.path.join(IMAGES_DIR, image_filename)
    if not os.path.exists(path):
        return None
    try:
        img = Image.open(path)
        img.thumbnail(size)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None


def get_recipe_images(recipe):
    """Retourne la liste des noms de fichiers image d'une recette, en gérant
    la compatibilité avec l'ancien format à une seule photo (clé 'image')."""
    images = recipe.get("images")
    if images:
        return list(images)
    legacy = recipe.get("image")
    return [legacy] if legacy else []


def delete_recipe_images(recipe):
    """Supprime tous les fichiers photo associés à une recette."""
    for fname in get_recipe_images(recipe):
        delete_image_file(fname)


def duplicate_recipe_images(recipe):
    """Copie physiquement tous les fichiers photo d'une recette sous de
    nouveaux noms, pour qu'une recette dupliquée ait ses propres fichiers
    indépendants (supprimer l'une ne doit pas affecter l'autre)."""
    new_names = []
    for fname in get_recipe_images(recipe):
        src = os.path.join(IMAGES_DIR, fname)
        if os.path.exists(src):
            new_names.append(copy_image_to_store(src))
    return new_names


# ---------------------------------------------------------------------------
# Fonctions partagées de calcul / export de liste de courses, réutilisées par
# "Voir toutes les recettes", le planificateur de repas et les menus.
# ---------------------------------------------------------------------------

def format_quantity_with_unit(qty, unit):
    """Convertit automatiquement une grande quantité vers une unité plus
    parlante pour un total de liste de courses : les grammes passent en
    kilogrammes au-delà de 1000 g (1 kg = 1000 g), et les centilitres
    passent en litres au-delà de 100 cl (1 L = 100 cl)."""
    unit_lower = (unit or "").strip().lower()
    if unit_lower in ("gr", "g", "gramme", "grammes") and qty >= 1000:
        qty = qty / 1000
        unit = "kg"
    elif unit_lower == "cl" and qty >= 100:
        qty = qty / 100
        unit = "L"
    qty = round(qty, 2)
    if qty == int(qty):
        qty = int(qty)
    return qty, unit


def compute_grouped_totals(recipe_persons_pairs):
    """À partir d'une liste de (recette, nombre_de_personnes), calcule le
    total de chaque ingrédient nécessaire, regroupé par rayon de magasin
    (dans l'ordre RAYON_ORDER). Retourne grouped_totals :
    [(rayon, [(nom, quantité, unité), ...]), ...]. Les grandes quantités
    (≥ 1000 g, ≥ 100 cl) sont automatiquement affichées en kg / L."""
    totals = {}
    for recipe, persons in recipe_persons_pairs:
        for ing in recipe["ingredients"]:
            key = (ing["name"].strip().lower(), ing["unit"].strip().lower())
            totals[key] = totals.get(key, 0) + ing["quantity"] * persons

    by_rayon = {}
    for (name, unit), qty in totals.items():
        display_qty, display_unit = format_quantity_with_unit(qty, unit)
        rayon = get_ingredient_rayon(name)
        by_rayon.setdefault(rayon, []).append((name.capitalize(), display_qty, display_unit))

    grouped_totals = []
    for rayon in RAYON_ORDER:
        if rayon in by_rayon:
            items = sorted(by_rayon[rayon], key=lambda x: ingredient_sort_key(x[0]))
            grouped_totals.append((rayon, items))
    return grouped_totals


def grouped_totals_from_flat_items(items):
    """Reconvertit une liste plate d'ingrédients modifiable
    [{'name','quantity','unit','rayon'}, ...] (utilisée pour l'édition
    ligne par ligne d'une liste de courses déjà calculée) vers le même
    format que compute_grouped_totals : [(rayon, [(nom, qté, unité), ...]), ...]."""
    by_rayon = {}
    for item in items:
        by_rayon.setdefault(item["rayon"], []).append((item["name"], item["quantity"], item["unit"]))
    grouped_totals = []
    for rayon in RAYON_ORDER:
        if rayon in by_rayon:
            entries = sorted(by_rayon[rayon], key=lambda x: ingredient_sort_key(x[0]))
            grouped_totals.append((rayon, entries))
    return grouped_totals


SAVED_SHOPPING_LISTS_FILE = os.path.join(BASE_DIR, "saved_shopping_lists.json")


def load_saved_shopping_lists():
    if os.path.exists(SAVED_SHOPPING_LISTS_FILE):
        try:
            with open(SAVED_SHOPPING_LISTS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            pass
    return []


def save_saved_shopping_lists(lists):
    with open(SAVED_SHOPPING_LISTS_FILE, "w", encoding="utf-8") as f:
        json.dump(lists, f, ensure_ascii=False, indent=2)


def write_shopping_list_txt(path, title, chosen_recipes, grouped_totals):
    """Écrit une liste de courses au format texte. chosen_recipes est une
    liste de (libellé_affiché, personnes)."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"=== {title} ===\n\n")
        f.write(t("shoppingexport_generated_on", date=datetime.now().strftime("%d/%m/%Y %H:%M")) + "\n\n")
        f.write(t("shoppingexport_selected_recipes") + "\n")
        for label, persons in chosen_recipes:
            f.write(f"- {label} ({persons} pers.)\n")
        for rayon, items in grouped_totals:
            f.write(f"\n{translate_rayon_name(rayon)} :\n")
            for name, qty, unit in items:
                unit_display = f" {translate_unit_name(unit)}" if unit else ""
                f.write(f"- {translate_ingredient_name(name)} : {qty}{unit_display}\n")


def build_shopping_list_workbook(chosen_recipes, grouped_totals):
    """Construit un classeur Excel (openpyxl) pour une liste de courses.
    Nécessite que OPENPYXL_AVAILABLE soit vrai."""
    wb = Workbook()
    ws_recipes = wb.active
    ws_recipes.title = t("shoppingexport_excel_sheet_recipes")
    ws_recipes.append([t("shoppingexport_excel_col_recipe"), t("shoppingexport_excel_col_persons")])
    for label, persons in chosen_recipes:
        ws_recipes.append([label, persons])

    ws_ing = wb.create_sheet(t("shoppingexport_excel_sheet_ingredients"))
    ws_ing.append([
        t("shoppingexport_excel_col_rayon"), t("shoppingexport_excel_col_ingredient"),
        t("shoppingexport_excel_col_total_qty"), t("shoppingexport_excel_col_unit")
    ])
    for rayon, items in grouped_totals:
        for name, qty, unit in items:
            ws_ing.append([translate_rayon_name(rayon), translate_ingredient_name(name), qty, unit])
    return wb


def build_shopping_list_pdf(path, title, chosen_recipes, grouped_totals):
    """Construit un PDF pour une liste de courses. Nécessite que
    REPORTLAB_AVAILABLE soit vrai."""
    c = pdf_canvas.Canvas(path, pagesize=A4)
    width, height = A4
    y = height - 2 * cm

    c.setFont("Helvetica-Bold", 18)
    c.drawString(2 * cm, y, title)
    y -= 1 * cm

    c.setFont("Helvetica", 10)
    c.drawString(2 * cm, y, t("shoppingexport_generated_on", date=datetime.now().strftime("%d/%m/%Y %H:%M")))
    y -= 1 * cm

    c.setFont("Helvetica-Bold", 12)
    c.drawString(2 * cm, y, t("shoppingexport_selected_recipes"))
    y -= 0.6 * cm
    c.setFont("Helvetica", 10)
    for label, persons in chosen_recipes:
        c.drawString(2.3 * cm, y, f"- {label} ({persons} pers.)")
        y -= 0.5 * cm
        if y < 3 * cm:
            c.showPage()
            y = height - 2 * cm

    for rayon, items in grouped_totals:
        y -= 0.4 * cm
        if y < 3.5 * cm:
            c.showPage()
            y = height - 2 * cm
        c.setFont("Helvetica-Bold", 12)
        c.drawString(2 * cm, y, f"{translate_rayon_name(rayon)} :")
        y -= 0.6 * cm
        c.setFont("Helvetica", 10)
        for name, qty, unit in items:
            unit_display = f" {translate_unit_name(unit)}" if unit else ""
            c.drawString(2.3 * cm, y, f"- {translate_ingredient_name(name)} : {qty}{unit_display}")
            y -= 0.5 * cm
            if y < 2 * cm:
                c.showPage()
                y = height - 2 * cm

    c.save()


ALLERGENS = ["Gluten", "Lactose", "Œufs", "Arachides", "Fruits à coque",
             "Soja", "Poisson", "Crustacés", "Sésame", "Céleri", "Moutarde",
             "Sulfites", "Lupin", "Mollusques"]

# Table de correspondance des allergènes par langue, sur le même principe
# que les catégories et difficultés : la donnée réelle stockée dans
# chaque recette (recipe["allergens"]) reste toujours en français — ces
# valeurs sont comparées lors de la détection automatique et des filtres
# —, cette table ne sert qu'à l'affichage. Les cases à cocher des
# allergènes sont elles aussi toujours en liste fermée, donc sans
# ambiguïté possible.
ALLERGEN_TRANSLATIONS = {
    "en": {
        "gluten": "Gluten",
        "lactose": "Lactose",
        "œufs": "Eggs",
        "arachides": "Peanuts",
        "fruits à coque": "Tree nuts",
        "soja": "Soy",
        "poisson": "Fish",
        "crustacés": "Shellfish",
        "sésame": "Sesame",
        "céleri": "Celery",
        "moutarde": "Mustard",
        "sulfites": "Sulphites",
        "lupin": "Lupin",
        "mollusques": "Molluscs",
    },
    "es": {
        "gluten": "Gluten",
        "lactose": "Lactosa",
        "œufs": "Huevos",
        "arachides": "Cacahuetes",
        "fruits à coque": "Frutos de cáscara",
        "soja": "Soja",
        "poisson": "Pescado",
        "crustacés": "Crustáceos",
        "sésame": "Sésamo",
        "céleri": "Apio",
        "moutarde": "Mostaza",
        "sulfites": "Sulfitos",
        "lupin": "Altramuces",
        "mollusques": "Moluscos",
    },
    "de": {
        "gluten": "Gluten",
        "lactose": "Laktose",
        "œufs": "Eier",
        "arachides": "Erdnüsse",
        "fruits à coque": "Schalenfrüchte",
        "soja": "Soja",
        "poisson": "Fisch",
        "crustacés": "Krebstiere",
        "sésame": "Sesam",
        "céleri": "Sellerie",
        "moutarde": "Senf",
        "sulfites": "Sulfite",
        "lupin": "Lupinen",
        "mollusques": "Weichtiere",
    },
}


def translate_allergen_name(allergen):
    """Retourne le nom d'affichage d'un allergène dans la langue
    actuellement sélectionnée. Ne change jamais la donnée réelle stockée
    dans la recette (recipe["allergens"]), utilisée pour la détection
    automatique et les comparaisons."""
    if not allergen:
        return allergen
    if CURRENT_LANGUAGE == "fr":
        return allergen
    return ALLERGEN_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(allergen.strip().lower(), allergen)

# Ingrédients de base qu'on a presque toujours sous la main, pré-cochés par
# défaut dans "Que puis-je cuisiner ?" (l'utilisateur reste libre de les
# retirer au cas par cas).
PANTRY_STAPLES = [
    "Sel", "Poivre", "Huile de tournesol", "Huile d'olive", "Beurre",
    "Farine", "Sucre", "Vinaigre", "Moutarde", "Riz", "Pâtes", "Lait",
]


def draw_recipe_content(c, recipe, persons, width, height):
    """Dessine le contenu complet d'une recette (titre, infos, photo,
    ingrédients, description, notes) sur un canevas reportlab déjà créé, à
    partir du haut d'une page. Retourne la position verticale (y) atteinte à
    la fin, pour pouvoir enchaîner d'autres contenus sur la même page si
    besoin (sinon l'appelant peut faire c.showPage() lui-même)."""
    y = height - 2 * cm

    star = "⭐ " if recipe.get("favorite") else ""
    c.setFont("Helvetica-Bold", 18)
    c.drawString(2 * cm, y, f"{star}{recipe['name']}")
    y -= 0.9 * cm

    cat = translate_category_name(recipe.get("category", "Autre"))
    c.setFont("Helvetica", 10)
    c.drawString(2 * cm, y, t("recipepdf_category_persons", cat=cat, persons=persons))
    y -= 0.6 * cm

    rating = recipe.get("rating", 0)
    if rating:
        c.drawString(2 * cm, y, t("recipepdf_rating", stars=rating_stars(rating)))
        y -= 0.6 * cm

    info_bits = []
    if recipe.get("prep_time"):
        info_bits.append(t("recipepdf_prep", time=recipe['prep_time']))
    if recipe.get("cook_time"):
        info_bits.append(t("recipepdf_cook", time=recipe['cook_time']))
    if recipe.get("difficulty"):
        info_bits.append(t("recipepdf_difficulty", value=translate_difficulty_name(recipe['difficulty'])))
    if info_bits:
        c.drawString(2 * cm, y, "   |   ".join(info_bits))
        y -= 0.6 * cm

    allergens = recipe.get("allergens") or []
    if allergens:
        c.setFillColorRGB(0.7, 0.2, 0.2)
        c.drawString(2 * cm, y, t("recipepdf_allergens", list=", ".join(translate_allergen_name(a) for a in allergens)))
        c.setFillColorRGB(0, 0, 0)
        y -= 0.6 * cm

    y -= 0.2 * cm

    # Photo (la première disponible)
    images = get_recipe_images(recipe)
    if images:
        img_path = os.path.join(IMAGES_DIR, images[0])
        if os.path.exists(img_path):
            try:
                img_reader = ImageReader(img_path)
                iw, ih = img_reader.getSize()
                max_w = 8 * cm
                scale = max_w / iw
                draw_w = max_w
                draw_h = ih * scale
                if y - draw_h < 3 * cm:
                    c.showPage()
                    y = height - 2 * cm
                c.drawImage(img_reader, 2 * cm, y - draw_h, width=draw_w, height=draw_h,
                            preserveAspectRatio=True, mask="auto")
                y -= draw_h + 0.6 * cm
            except Exception:
                pass

    if y < 4 * cm:
        c.showPage()
        y = height - 2 * cm

    c.setFont("Helvetica-Bold", 12)
    c.drawString(2 * cm, y, t("recipepdf_ingredients_heading"))
    y -= 0.6 * cm
    c.setFont("Helvetica", 10)
    for ing in recipe["ingredients"]:
        qty = round(ing["quantity"] * persons, 2)
        unit = f" {translate_unit_name(ing['unit'])}" if ing["unit"] else ""
        c.drawString(2.3 * cm, y, f"- {translate_ingredient_name(ing['name']).capitalize()} : {qty}{unit}")
        y -= 0.5 * cm
        if y < 3 * cm:
            c.showPage()
            y = height - 2 * cm

    cost, cost_known, cost_total = compute_recipe_cost(recipe, persons)
    nutrition, nutri_known, nutri_total = compute_recipe_nutrition(recipe, persons)
    if cost_known or nutri_known:
        y -= 0.3 * cm
        if y < 3 * cm:
            c.showPage()
            y = height - 2 * cm
        c.setFont("Helvetica-Oblique", 9)
        if cost_known:
            partial = "" if cost_known == cost_total else t("recipepdf_partial_suffix", known=cost_known, total=cost_total)
            c.drawString(2 * cm, y, t("recipepdf_cost", cost=f"{cost:.2f}", partial=partial))
            y -= 0.45 * cm
        if nutri_known:
            partial = "" if nutri_known == nutri_total else t("recipepdf_partial_suffix", known=nutri_known, total=nutri_total)
            c.drawString(
                2 * cm, y,
                t(
                    "recipepdf_nutrition", partial=partial,
                    kcal=f"{nutrition['kcal']:.0f}", protein=f"{nutrition['protein_g']:.0f}",
                    carbs=f"{nutrition['carbs_g']:.0f}", fat=f"{nutrition['fat_g']:.0f}"
                )
            )
            y -= 0.45 * cm
        c.setFont("Helvetica", 10)

    def draw_wrapped_section(title, text, y):
        y -= 0.4 * cm
        if y < 4 * cm:
            c.showPage()
            y = height - 2 * cm
        c.setFont("Helvetica-Bold", 12)
        c.drawString(2 * cm, y, title)
        y -= 0.6 * cm
        c.setFont("Helvetica", 10)
        for paragraph in text.split("\n"):
            words = paragraph.split(" ")
            line = ""
            for word in words:
                test_line = f"{line} {word}".strip()
                if c.stringWidth(test_line, "Helvetica", 10) > (width - 4 * cm):
                    c.drawString(2 * cm, y, line)
                    y -= 0.5 * cm
                    if y < 2 * cm:
                        c.showPage()
                        y = height - 2 * cm
                    line = word
                else:
                    line = test_line
            if line:
                c.drawString(2 * cm, y, line)
                y -= 0.5 * cm
                if y < 2 * cm:
                    c.showPage()
                    y = height - 2 * cm
        return y

    description = recipe.get("description", "").strip()
    if description:
        y = draw_wrapped_section(t("recipepdf_description_heading"), description, y)

    personal_notes = recipe.get("personal_notes", "").strip()
    if personal_notes:
        y = draw_wrapped_section(t("recipepdf_notes_heading"), personal_notes, y)

    return y


def build_cookbook_pdf(path, recipes_with_persons):
    """Construit un PDF regroupant plusieurs recettes à la suite (une page de
    titre listant le sommaire, puis une recette par page), avec toutes les
    pages numérotées et le numéro de page de chaque recette indiqué en face
    de son nom dans le sommaire."""
    c = pdf_canvas.Canvas(path, pagesize=A4)
    width, height = A4

    # ---- Pré-calcul (rendu à blanc, jeté ensuite) du nombre de pages :
    # une recette peut elle-même s'étaler sur plusieurs pages selon son
    # contenu (photos, longue description...), donc il faut d'abord simuler
    # tout le document pour connaître la page de démarrage de chaque
    # recette et le nombre total de pages, avant de dessiner le sommaire. ----
    def _count_summary_pages():
        y = height - 3 * cm - 1 * cm - 1.2 * cm - 0.7 * cm
        pages = 1
        for _ in recipes_with_persons:
            y -= 0.5 * cm
            if y < 2 * cm:
                pages += 1
                y = height - 2 * cm
        return pages

    def _count_recipe_pages(recipe, persons):
        dummy = pdf_canvas.Canvas(io.BytesIO(), pagesize=(width, height))
        count = [1]
        real_show_page = dummy.showPage

        def counting_show_page():
            count[0] += 1
            real_show_page()

        dummy.showPage = counting_show_page
        draw_recipe_content(dummy, recipe, persons, width, height)
        return count[0]

    summary_page_count = _count_summary_pages()
    recipe_start_pages = []
    running_page = summary_page_count + 1
    for recipe, persons in recipes_with_persons:
        recipe_start_pages.append(running_page)
        running_page += _count_recipe_pages(recipe, persons)
    total_pages = running_page - 1

    # ---- Numérotation automatique : on intercepte chaque saut de page
    # (y compris ceux internes à draw_recipe_content) pour dessiner le pied
    # de page juste avant de passer à la suivante. ----
    current_page = [1]
    real_show_page = c.showPage

    def numbered_show_page():
        c.setFont("Helvetica", 8)
        c.drawCentredString(width / 2, 1 * cm, t("cookbookpdf_page_number", current=current_page[0], total=total_pages))
        real_show_page()
        current_page[0] += 1

    c.showPage = numbered_show_page

    # Page(s) de titre / sommaire, avec le numéro de page de chaque recette
    y = height - 3 * cm
    c.setFont("Helvetica-Bold", 24)
    c.drawString(2 * cm, y, t("home_window_title"))
    y -= 1 * cm
    c.setFont("Helvetica", 10)
    c.drawString(2 * cm, y, t("cookbookpdf_generated_on", date=datetime.now().strftime("%d/%m/%Y")))
    y -= 1.2 * cm
    c.setFont("Helvetica-Bold", 13)
    c.drawString(2 * cm, y, t("cookbookpdf_summary_heading"))
    y -= 0.7 * cm
    c.setFont("Helvetica", 10)
    for i, (recipe, persons) in enumerate(recipes_with_persons):
        cat = translate_category_name(recipe.get("category", "Autre"))
        c.drawString(2.3 * cm, y, t("cookbookpdf_summary_line", cat=cat, name=recipe['name']))
        c.drawRightString(width - 2 * cm, y, str(recipe_start_pages[i]))
        y -= 0.5 * cm
        if y < 2 * cm:
            c.showPage()
            y = height - 2 * cm

    for recipe, persons in recipes_with_persons:
        c.showPage()
        draw_recipe_content(c, recipe, persons, width, height)

    # La toute dernière page ne passe jamais par un showPage() suivant : son
    # pied de page doit être dessiné explicitement ici, juste avant de
    # sauvegarder. On restaure d'abord showPage() à son comportement
    # d'origine, car Canvas.save() l'appelle lui-même en interne pour
    # finaliser la dernière page — sans cela, le pied de page serait dessiné
    # une seconde fois par erreur.
    c.showPage = real_show_page
    c.setFont("Helvetica", 8)
    c.drawCentredString(width / 2, 1 * cm, t("cookbookpdf_page_number", current=current_page[0], total=total_pages))
    c.save()


WEEKDAYS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]


def load_weekly_plan():
    """Charge le planning de la semaine : {jour: {'recipe_name':.., 'persons':..}}."""
    if os.path.exists(WEEKLY_PLAN_FILE):
        try:
            with open(WEEKLY_PLAN_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            return {}
    return {}


def save_weekly_plan(plan):
    with open(WEEKLY_PLAN_FILE, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False, indent=2)


def get_current_week_key():
    """Clé identifiant la semaine calendaire actuelle (année-numéro de
    semaine ISO), utilisée pour qu'un même planning enregistré plusieurs
    fois dans la même semaine ne crée qu'une seule entrée d'historique."""
    year, week, _ = datetime.now().isocalendar()
    return f"{year}-S{week:02d}"


def load_weekly_plan_history():
    """Liste des plannings archivés : [{'week_start','plan','saved_at'}, ...]."""
    if os.path.exists(WEEKLY_PLAN_HISTORY_FILE):
        try:
            with open(WEEKLY_PLAN_HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            pass
    return []


def save_weekly_plan_history(history):
    with open(WEEKLY_PLAN_HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


WEEKLY_PLAN_HISTORY_RETENTION = 26  # environ 6 mois d'historique


def archive_current_week(plan):
    """Enregistre un instantané du planning actuel dans l'historique, sous
    la clé de la semaine calendaire en cours : ré-enregistrer plusieurs fois
    dans la même semaine met simplement à jour son entrée, sans créer de
    doublon. Ne fait rien si le planning est entièrement vide."""
    if not any(plan.values()):
        return
    week_key = get_current_week_key()
    history = load_weekly_plan_history()
    history = [h for h in history if h.get("week_start") != week_key]
    history.append({
        "week_start": week_key,
        "plan": plan,
        "saved_at": datetime.now().strftime("%Y-%m-%d"),
    })
    history.sort(key=lambda h: h.get("week_start", ""))
    history = history[-WEEKLY_PLAN_HISTORY_RETENTION:]
    save_weekly_plan_history(history)


def load_weekly_plan_templates():
    """Modèles de semaine réutilisables : {nom_du_modele: plan}."""
    if os.path.exists(WEEKLY_PLAN_TEMPLATES_FILE):
        try:
            with open(WEEKLY_PLAN_TEMPLATES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}


def save_weekly_plan_templates(templates):
    with open(WEEKLY_PLAN_TEMPLATES_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Export du planning de la semaine vers un fichier .ics (iCalendar), lisible
# par Google Agenda, Outlook, Apple Calendrier, etc.
# ---------------------------------------------------------------------------

_ICS_MEAL_TIMES = {
    "Petit-déjeuner": ("0800", "0830"),
    "Déjeuner": ("1230", "1330"),
    "Dîner": ("1930", "2030"),
}
_ICS_MEAL_GROUPS = {
    "Petit-déjeuner": ["Petit-déjeuner"],
    "Déjeuner": ["Déjeuner — Entrée", "Déjeuner — Plat", "Déjeuner — Dessert"],
    "Dîner": ["Dîner — Entrée", "Dîner — Plat", "Dîner — Dessert"],
}
# Traductions pour l'affichage dans le calendrier exporté (.ics)
# uniquement, par langue : ces libellés ne sont ni des clés de
# _ICS_MEAL_GROUPS/_ICS_MEAL_TIMES, ni des créneaux complets de
# MEAL_SLOTS (juste la période du repas, ou juste le type de plat), donc
# traités séparément des autres tables.
_ICS_MEAL_PERIOD_TRANSLATIONS = {
    "en": {"petit-déjeuner": "Breakfast", "déjeuner": "Lunch", "dîner": "Dinner"},
    "es": {"petit-déjeuner": "Desayuno", "déjeuner": "Almuerzo", "dîner": "Cena"},
    "de": {"petit-déjeuner": "Frühstück", "déjeuner": "Mittagessen", "dîner": "Abendessen"},
}
_ICS_COURSE_LABEL_TRANSLATIONS = {
    "en": {"entrée": "Starter", "plat": "Main", "dessert": "Dessert"},
    "es": {"entrée": "Entrante", "plat": "Plato principal", "dessert": "Postre"},
    "de": {"entrée": "Vorspeise", "plat": "Hauptgericht", "dessert": "Dessert"},
}


def _translate_ics_meal_period(period):
    if CURRENT_LANGUAGE == "fr":
        return period
    return _ICS_MEAL_PERIOD_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(period.strip().lower(), period)


def _translate_ics_course_label(label):
    if not label or CURRENT_LANGUAGE == "fr":
        return label
    return _ICS_COURSE_LABEL_TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(label.strip().lower(), label)


def _escape_ics_text(text):
    return (text.replace("\\", "\\\\")
                .replace(";", "\\;")
                .replace(",", "\\,")
                .replace("\n", "\\n"))


def _fold_ics_line(line):
    """Replie une ligne selon la limite de 75 octets recommandée par la
    norme iCalendar (RFC 5545), pour une compatibilité maximale."""
    if len(line.encode("utf-8")) <= 75:
        return line
    parts = []
    current = ""
    for ch in line:
        if len((current + ch).encode("utf-8")) > 74:
            parts.append(current)
            current = " " + ch
        else:
            current += ch
    parts.append(current)
    return "\r\n".join(parts)


def build_weekly_plan_ics(plan):
    """Construit le contenu d'un fichier .ics à partir du planning de la
    semaine : un évènement hebdomadaire récurrent par repas (petit-déjeuner,
    déjeuner, dîner) et par jour où quelque chose est prévu."""
    weekday_index = {d: i for i, d in enumerate(WEEKDAYS)}  # Lundi=0 ... Dimanche=6
    today = datetime.now().date()
    now_stamp = datetime.now().strftime("%Y%m%dT%H%M%S")

    lines = ["BEGIN:VCALENDAR", "VERSION:2.0",
             "PRODID:-//Mon Livre de Recettes//FR", "CALSCALE:GREGORIAN"]
    event_count = 0

    for day, slots_data in plan.items():
        if day not in weekday_index or not slots_data:
            continue
        delta_days = (weekday_index[day] - today.weekday()) % 7
        event_date = today + timedelta(days=delta_days)
        date_str = event_date.strftime("%Y%m%d")

        for meal_period, slot_names in _ICS_MEAL_GROUPS.items():
            components = []
            for slot in slot_names:
                info = slots_data.get(slot)
                if info and info.get("recipe_name"):
                    label = _translate_ics_course_label(slot.split(" — ", 1)[1]) if " — " in slot else None
                    components.append(f"{label} : {info['recipe_name']}" if label else info["recipe_name"])
            if not components:
                continue

            start_time, end_time = _ICS_MEAL_TIMES[meal_period]
            summary_names = [c.split(" : ", 1)[1] if " : " in c else c for c in components]
            summary = f"{_translate_ics_meal_period(meal_period)} : " + ", ".join(summary_names)
            description = "\\n".join(components)

            event_count += 1
            uid = f"recette-{event_count}-{uuid.uuid4().hex}@monlivrederecettes"

            lines.append("BEGIN:VEVENT")
            lines.append(_fold_ics_line(f"UID:{uid}"))
            lines.append(f"DTSTAMP:{now_stamp}")
            lines.append(f"DTSTART:{date_str}T{start_time}00")
            lines.append(f"DTEND:{date_str}T{end_time}00")
            lines.append("RRULE:FREQ=WEEKLY")
            lines.append(_fold_ics_line(f"SUMMARY:{_escape_ics_text(summary)}"))
            lines.append(_fold_ics_line(f"DESCRIPTION:{_escape_ics_text(description)}"))
            lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def load_menus():
    """Charge la liste des menus enregistrés :
    [{'name':.., 'items': [{'recipe_name':.., 'persons':..}, ...]}, ...]."""
    if os.path.exists(MENUS_FILE):
        try:
            with open(MENUS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            return []
    return []


def save_menus(menus):
    with open(MENUS_FILE, "w", encoding="utf-8") as f:
        json.dump(menus, f, ensure_ascii=False, indent=2)


def find_recipe_by_name(recipes, name):
    return next((r for r in recipes if r.get("name") == name), None)


# ---------------------------------------------------------------------------
# Import d'une recette depuis un lien internet (données structurées
# Schema.org "Recipe", utilisées par la plupart des sites de cuisine).
# ---------------------------------------------------------------------------

_FRACTION_MAP = {"½": 0.5, "¼": 0.25, "¾": 0.75, "⅓": 1 / 3, "⅔": 2 / 3}

_UNIT_ALIASES = {
    "g": "Gr", "gr": "Gr", "gramme": "Gr", "grammes": "Gr",
    "cl": "cl",
    "cas": "cuillère à soupe", "c.a.s": "cuillère à soupe",
    "cuillere": "cuillère à soupe", "cuillères": "cuillère à soupe", "cuillère": "cuillère à soupe",
    "cuillere a soupe": "cuillère à soupe", "cuillère à soupe": "cuillère à soupe",
    "cuillères à soupe": "cuillère à soupe", "cuilleres a soupe": "cuillère à soupe",
    "cac": "cuillère à café",
    "cuillere a cafe": "cuillère à café", "cuillère à café": "cuillère à café",
    "cuillères à café": "cuillère à café", "cuilleres a cafe": "cuillère à café",
    "piece": "pièce", "pièce": "pièce", "pièces": "pièce", "pieces": "pièce",
}


def parse_quantity_token(token):
    token = token.strip()
    if token in _FRACTION_MAP:
        return _FRACTION_MAP[token]
    m = re.match(r"^(\d+)\s*/\s*(\d+)$", token)
    if m:
        return int(m.group(1)) / int(m.group(2))
    token = token.replace(",", ".")
    try:
        return float(token)
    except ValueError:
        return None


def _match_unit_tokens(tokens, start):
    """Essaie de faire correspondre 3, puis 2, puis 1 token(s) à partir de
    `start` à une unité connue (ex. « cuillères à soupe »), en testant la
    séquence la plus longue en premier. Insensible aux accents/à la casse."""
    for length in (3, 2, 1):
        end = start + length
        if end <= len(tokens):
            candidate = " ".join(t.strip(".,") for t in tokens[start:end])
            candidate_key = ingredient_sort_key(candidate)
            for alias, unit in _UNIT_ALIASES.items():
                if ingredient_sort_key(alias) == candidate_key:
                    return unit, end
    return None, start


def parse_ingredient_line(line):
    """Tente d'extraire (nom, quantité, unité) d'une ligne d'ingrédient en
    texte libre, telle que fournie par un site de recettes. Retourne un
    résultat raisonnable même quand le texte ne suit pas un format standard
    (l'utilisateur pourra toujours corriger à la main après import)."""
    original = line.strip()
    if not original:
        return None
    tokens = original.split()
    if not tokens:
        return None

    qty = parse_quantity_token(tokens[0])
    unit = None
    rest_start = 0
    if qty is not None:
        rest_start = 1
        if len(tokens) > 1:
            unit, rest_start = _match_unit_tokens(tokens, 1)

    name = " ".join(tokens[rest_start:]).strip()
    name = re.sub(r"^(de |d['’]|of )", "", name, flags=re.IGNORECASE).strip()

    if not name:
        name = original
        qty = qty or 1
        unit = unit or "pièce"

    return {"name": name.capitalize(), "quantity": qty or 1, "unit": unit or "pièce"}


def parse_iso8601_duration_minutes(duration):
    """Convertit une durée ISO 8601 (ex. 'PT1H30M') en nombre de minutes."""
    if not duration:
        return None
    match = re.search(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", str(duration))
    if not match:
        return None
    hours = int(match.group(1) or 0)
    minutes = int(match.group(2) or 0)
    seconds = int(match.group(3) or 0)
    total = hours * 60 + minutes + (1 if seconds >= 30 else 0)
    return total if total > 0 else None


def _find_recipe_jsonld(data):
    """Cherche récursivement un objet Schema.org de type Recipe dans une
    structure JSON-LD (qui peut être un objet, une liste, ou contenir
    '@graph')."""
    if isinstance(data, dict):
        type_value = data.get("@type")
        types = type_value if isinstance(type_value, list) else [type_value]
        if any(t and "recipe" in str(t).lower() for t in types):
            return data
        if "@graph" in data:
            found = _find_recipe_jsonld(data["@graph"])
            if found:
                return found
        for value in data.values():
            found = _find_recipe_jsonld(value)
            if found:
                return found
    elif isinstance(data, list):
        for item in data:
            found = _find_recipe_jsonld(item)
            if found:
                return found
    return None


def fetch_recipe_from_url(url):
    """Télécharge une page de recette et tente d'en extraire le contenu à
    partir des données structurées Schema.org (JSON-LD), un format utilisé
    par la grande majorité des sites de recettes. Lève une exception avec un
    message clair en cas d'échec."""
    request = urllib.request.Request(
        url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    )
    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            raw = response.read()
            charset = response.headers.get_content_charset() or "utf-8"
            page_html = raw.decode(charset, errors="replace")
    except urllib.error.URLError as e:
        raise RuntimeError(f"Impossible d'accéder à cette adresse : {e}")
    except Exception as e:
        raise RuntimeError(f"Erreur lors du téléchargement de la page : {e}")

    scripts = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        page_html, flags=re.DOTALL | re.IGNORECASE
    )

    recipe_data = None
    for script in scripts:
        try:
            parsed = json.loads(script.strip())
        except (json.JSONDecodeError, ValueError):
            continue
        found = _find_recipe_jsonld(parsed)
        if found:
            recipe_data = found
            break

    if recipe_data is None:
        raise RuntimeError(
            "Aucune recette structurée n'a été trouvée sur cette page.\n\n"
            "Cet import fonctionne avec les sites qui utilisent le format "
            "standard « Schema.org Recipe » (la plupart des grands sites de "
            "cuisine). Vous pouvez toujours créer la recette manuellement."
        )

    def clean_text(value):
        if isinstance(value, list):
            value = " ".join(str(v) for v in value)
        return html.unescape(re.sub(r"<[^>]+>", " ", str(value or ""))).strip()

    name = clean_text(recipe_data.get("name", "")) or "Recette importée"

    raw_ingredients = recipe_data.get("recipeIngredient") or recipe_data.get("ingredients") or []
    if isinstance(raw_ingredients, str):
        raw_ingredients = [raw_ingredients]
    ingredients = []
    for line in raw_ingredients:
        parsed_ing = parse_ingredient_line(clean_text(line))
        if parsed_ing:
            ingredients.append(parsed_ing)

    instructions = recipe_data.get("recipeInstructions") or []
    steps = []
    if isinstance(instructions, str):
        steps = [instructions]
    elif isinstance(instructions, list):
        for item in instructions:
            if isinstance(item, dict):
                text = item.get("text") or item.get("name") or ""
                steps.append(clean_text(text))
            else:
                steps.append(clean_text(item))
    steps = [s for s in steps if s]
    description = "\n".join(f"{i + 1}. {s}" for i, s in enumerate(steps))

    prep_time = parse_iso8601_duration_minutes(recipe_data.get("prepTime"))
    cook_time = parse_iso8601_duration_minutes(recipe_data.get("cookTime"))

    yield_value = recipe_data.get("recipeYield")
    default_persons = None
    if yield_value:
        if isinstance(yield_value, list):
            yield_value = yield_value[0] if yield_value else None
        match = re.search(r"\d+", str(yield_value or ""))
        if match:
            default_persons = int(match.group())

    if not ingredients:
        raise RuntimeError(
            "Une recette a été trouvée sur cette page, mais aucun ingrédient "
            "n'a pu en être extrait. Vous pouvez créer la recette "
            "manuellement à la place."
        )

    # Récupération de la photo de la recette, si le site en indique une.
    images = []
    image_url = _extract_recipe_image_url(recipe_data.get("image"))
    if image_url:
        image_url = urllib.parse.urljoin(url, image_url)  # gère les URLs relatives
        downloaded = download_image_to_store(image_url)
        if downloaded:
            images.append(downloaded)

    return {
        "name": name[:100],
        "description": description[:2056],
        "ingredients": ingredients,
        "prep_time": str(prep_time) if prep_time else "",
        "cook_time": str(cook_time) if cook_time else "",
        "default_persons": default_persons or 4,
        "images": images,
        "source_url": url,
    }


def _extract_recipe_image_url(image):
    """Extrait une URL d'image utilisable depuis le champ 'image' d'une
    donnée structurée Schema.org, qui peut prendre plusieurs formes :
    chaîne, liste de chaînes, objet ImageObject, ou liste d'objets."""
    if not image:
        return None
    if isinstance(image, list):
        image = image[0] if image else None
    if isinstance(image, dict):
        return image.get("url") or image.get("@id")
    if isinstance(image, str):
        return image.strip() or None
    return None


def download_image_to_store(image_url, timeout=15):
    """Télécharge une image depuis une URL et l'enregistre dans le dossier
    images/. Retourne le nom de fichier généré, ou None en cas d'échec (page
    introuvable, ce n'est pas une image, pas de connexion...)."""
    try:
        request = urllib.request.Request(
            image_url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        )
        with urllib.request.urlopen(request, timeout=timeout) as response:
            content_type = (response.headers.get_content_type() or "").lower()
            if content_type and not content_type.startswith("image/"):
                return None
            data = response.read()
    except Exception:
        return None

    if not data:
        return None

    ext_map = {
        "image/jpeg": ".jpg", "image/jpg": ".jpg", "image/png": ".png",
        "image/webp": ".webp", "image/gif": ".gif",
    }
    ext = ext_map.get(content_type)
    if not ext:
        guessed = os.path.splitext(image_url.split("?")[0])[1].lower()
        ext = guessed if guessed in (".jpg", ".jpeg", ".png", ".webp", ".gif") else ".jpg"

    new_filename = f"{uuid.uuid4().hex}{ext}"
    dest_path = os.path.join(IMAGES_DIR, new_filename)
    try:
        with open(dest_path, "wb") as f:
            f.write(data)
    except OSError:
        return None
    return new_filename


# ---------------------------------------------------------------------------
# Corbeille : les recettes supprimées y sont déplacées (avec leurs photos
# conservées) au lieu d'être effacées immédiatement, pour pouvoir les
# récupérer en cas d'erreur.
# ---------------------------------------------------------------------------

def load_trash():
    if os.path.exists(TRASH_FILE):
        try:
            with open(TRASH_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            return []
    return []


def save_trash(trash):
    with open(TRASH_FILE, "w", encoding="utf-8") as f:
        json.dump(trash, f, ensure_ascii=False, indent=2)


def move_recipe_to_trash(recipe):
    """Déplace une recette vers la corbeille (ses photos ne sont PAS
    supprimées, pour pouvoir tout restaurer intact)."""
    trash = load_trash()
    trash.insert(0, {"recipe": recipe, "deleted_at": datetime.now().isoformat()})
    save_trash(trash)


# ---------------------------------------------------------------------------
# Historique des recettes récemment consultées (affiché sur la page d'accueil)
# ---------------------------------------------------------------------------

RECENT_VIEWS_MAX = 8


def load_recent_view_names():
    if os.path.exists(RECENT_VIEWS_FILE):
        try:
            with open(RECENT_VIEWS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            return []
    return []


def record_recipe_view(name):
    names = load_recent_view_names()
    names = [n for n in names if n != name]
    names.insert(0, name)
    names = names[:RECENT_VIEWS_MAX]
    with open(RECENT_VIEWS_FILE, "w", encoding="utf-8") as f:
        json.dump(names, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Palette et style visuel de l'application — une identité chaleureuse et
# gourmande plutôt que le gris par défaut de Windows, appliquée d'un coup à
# toute l'application (fenêtre principale et toutes les fenêtres ouvertes
# ensuite), sans avoir à retoucher chaque écran un par un. Deux palettes
# (clair/sombre) sont disponibles, basculables depuis la page d'accueil.
# ---------------------------------------------------------------------------

LIGHT_PALETTE = {
    "BG": "#FBF6EF",           # fond général, blanc cassé chaleureux
    "CARD": "#FFFFFF",         # fond des zones "carte"
    "BORDER": "#E8DCC8",       # bordures discrètes
    "TEXT": "#332B22",         # texte principal, brun très foncé
    "TEXT_MUTED": "#8A7D68",   # texte secondaire
    "ACCENT": "#D97B3F",       # orange terracotta — actions principales
    "ACCENT_DARK": "#B8622C",  # survol / pression
    "ACCENT_LIGHT": "#F3D9C4",  # fonds légers accentués
    "GREEN": "#5C8A57",        # vert sauge — validations, favoris, succès
    "ERROR": "#B4483A",        # rouge terracotta foncé — erreurs/avertissements
}

DARK_PALETTE = {
    "BG": "#211F1C",           # fond général, brun-noir doux
    "CARD": "#2C2A25",         # fond des zones "carte"
    "BORDER": "#43403A",       # bordures discrètes
    "TEXT": "#EFE8DB",         # texte principal, crème clair
    "TEXT_MUTED": "#AB9F8C",   # texte secondaire
    "ACCENT": "#E08A4F",       # orange terracotta, plus lumineux sur fond sombre
    "ACCENT_DARK": "#F0A868",  # survol / pression (plus clair que ACCENT en sombre)
    "ACCENT_LIGHT": "#4A3B2C",  # fonds légers accentués
    "GREEN": "#7CB273",        # vert sauge, plus lumineux
    "ERROR": "#E0897A",        # rouge corail, plus lumineux
}

COLOR_BG = LIGHT_PALETTE["BG"]
COLOR_CARD = LIGHT_PALETTE["CARD"]
COLOR_BORDER = LIGHT_PALETTE["BORDER"]
COLOR_TEXT = LIGHT_PALETTE["TEXT"]
COLOR_TEXT_MUTED = LIGHT_PALETTE["TEXT_MUTED"]
COLOR_ACCENT = LIGHT_PALETTE["ACCENT"]
COLOR_ACCENT_DARK = LIGHT_PALETTE["ACCENT_DARK"]
COLOR_ACCENT_LIGHT = LIGHT_PALETTE["ACCENT_LIGHT"]
COLOR_GREEN = LIGHT_PALETTE["GREEN"]
COLOR_ERROR = LIGHT_PALETTE["ERROR"]


def apply_palette(dark):
    """Met à jour les constantes de couleur globales selon le thème choisi.
    Doit être suivi d'un appel à configure_app_style() pour que les styles
    ttk et la base d'options Tk reflètent les nouvelles couleurs."""
    global COLOR_BG, COLOR_CARD, COLOR_BORDER, COLOR_TEXT, COLOR_TEXT_MUTED
    global COLOR_ACCENT, COLOR_ACCENT_DARK, COLOR_ACCENT_LIGHT, COLOR_GREEN, COLOR_ERROR
    palette = DARK_PALETTE if dark else LIGHT_PALETTE
    COLOR_BG = palette["BG"]
    COLOR_CARD = palette["CARD"]
    COLOR_BORDER = palette["BORDER"]
    COLOR_TEXT = palette["TEXT"]
    COLOR_TEXT_MUTED = palette["TEXT_MUTED"]
    COLOR_ACCENT = palette["ACCENT"]
    COLOR_ACCENT_DARK = palette["ACCENT_DARK"]
    COLOR_ACCENT_LIGHT = palette["ACCENT_LIGHT"]
    COLOR_GREEN = palette["GREEN"]
    COLOR_ERROR = palette["ERROR"]


def get_dark_mode_preference():
    return bool(load_settings().get("dark_mode", False))


def set_dark_mode_preference(value):
    settings = load_settings()
    settings["dark_mode"] = bool(value)
    save_settings(settings)


FONT_SCALE = 1.0


def apply_font_scale(large_text):
    """Met à jour l'échelle globale des polices selon le mode « Texte
    agrandi » (accessibilité). Doit être appelé avant la construction de
    toute fenêtre pour que sf() reflète le bon facteur."""
    global FONT_SCALE
    FONT_SCALE = 1.3 if large_text else 1.0


def sf(size):
    """Renvoie une taille de police mise à l'échelle selon le mode « Texte
    agrandi » (accessibilité) — à utiliser à la place d'un nombre littéral
    dans tous les tuples de police, ex. ("Segoe UI", sf(10), "bold")."""
    return round(size * FONT_SCALE)


def gs(size):
    """Renvoie une dimension de fenêtre (largeur, hauteur, ou plafond de
    hauteur) mise à l'échelle selon le mode « Texte agrandi », pour que les
    fenêtres restent assez grandes pour leur contenu affiché en plus gros
    caractères et qu'aucun bouton ne se retrouve coupé ou masqué."""
    return round(size * FONT_SCALE)


def get_large_text_preference():
    return bool(load_settings().get("large_text", False))


def set_large_text_preference(value):
    settings = load_settings()
    settings["large_text"] = bool(value)
    save_settings(settings)


CURRENT_LANGUAGE = "fr"

# Traductions disponibles, langue par langue. Le français n'a pas besoin
# d'être dupliqué ici : il sert de texte de secours (voir t() ci-dessous)
# pour toute clé pas encore traduite dans une autre langue — l'application
# reste donc entièrement utilisable pendant qu'on ajoute les traductions
# progressivement, une langue et une fenêtre à la fois.
FRENCH_STRINGS = {
    "home_window_title": "Mon Livre de Recettes",
    "home_banner_title": "👨‍🍳 Mon Livre de Recettes",
    "home_banner_subtitle": "Toutes vos recettes, à portée de main",
    "home_donate_button": "☕ Faire un don",
    "home_dark_theme": "🌙 Thème sombre",
    "home_light_theme": "☀️ Thème clair",
    "home_large_text_on": "🔎 Texte agrandi",
    "home_large_text_off": "🔎 Texte normal",
    "home_daily_recipe_title": "🎲 Recette du jour",
    "home_open_button": "👁 Ouvrir",
    "home_quick_filter_favorites": "⭐ Favoris",
    "home_quick_filter_quick": "⏱️ Rapide (≤ 30 min)",
    "home_quick_filter_vegetarian": "🥗 Végétarien",
    "home_quick_filter_wishlist": "💭 Envies",
    "home_wishlist_reminder": (
        "💭 {count} recette(s) en liste d'envies depuis plus de {days} jours — "
        "et si vous les essayiez ? (cliquez pour les voir)"
    ),
    "home_low_stock_reminder": (
        "📦 {count} ingrédient(s) presque épuisé(s) dans votre garde-manger : "
        "{names} — cliquez pour les ajouter à la liste de courses"
    ),
    "home_btn_add_recipe": "➕  Ajouter une recette",
    "home_btn_import_url": "🌐  Importer une recette depuis un lien",
    "home_btn_import_photo": "📷  Importer une recette depuis une photo",
    "home_btn_view_all_recipes": "🧾  Voir toutes les recettes (liste de courses)",
    "home_btn_view_one_recipe": "🍽️  Voir une recette précise",
    "home_btn_manage_recipes": "✏️  Modifier / Supprimer une recette",
    "home_btn_compare_recipes": "⚖️  Comparer deux recettes",
    "home_btn_manage_ingredients": "🥕  Gérer les ingrédients",
    "home_btn_ingredient_search": "🔎  Recherche par ingrédient",
    "home_btn_what_can_i_cook": "🧊  Que puis-je cuisiner ?",
    "home_btn_pantry": "📦  Mon garde-manger",
    "home_btn_unit_converter": "🔄  Convertisseur d'unités",
    "home_btn_weekly_plan": "📅  Planning de la semaine",
    "home_btn_menus": "📋  Mes menus",
    "home_btn_statistics": "📊  Statistiques",
    "home_btn_export_cookbook": "📖  Exporter le livre de recettes",
    "home_btn_import_export": "💾  Importer / Exporter les données",
    "home_btn_trash": "🗑️  Corbeille",
    "home_today_title": "📅 Aujourd'hui",
    "home_recent_title": "🕘 Récemment consultées",
    "home_wishlist_title": "💭 Recettes à essayer",
    "home_new_draw_button": "🎲 Nouveau tirage",
    "home_footer_recipe_count": "{count} recette(s) enregistrée(s)",
    "home_nothing_planned": (
        "Rien de planifié pour {day}. Remplissez le « 📅 Planning de la semaine » pour le voir ici."
    ),
    "home_no_recent_recipe": "Aucune recette consultée pour le moment.",
    "home_no_wishlist_recipe": "Aucune recette dans votre liste d'envies pour le moment.",
    "warning_pillow": "Pillow non installé : les photos ne s'afficheront pas (pip install pillow)",
    "warning_reportlab": "reportlab non installé : export PDF indisponible (pip install reportlab)",
    "warning_openpyxl": "openpyxl non installé : export Excel indisponible (pip install openpyxl)",
    "warning_qrcode": "qrcode non installé : export QR code indisponible (pip install qrcode)",
    "warning_pytesseract": (
        "pytesseract non installé : import depuis une photo indisponible "
        "(pip install pytesseract, + Tesseract OCR)"
    ),

    # ---- Titres de dialogue génériques, réutilisés dans toute l'application ----
    "common_error": "Erreur",
    "common_info": "Info",
    "common_confirm": "Confirmer",
    "common_success": "Succès",
    "common_module_missing": "Module manquant",
    "common_all_categories": "Toutes",
    "common_export_failed": "L'export a échoué :\n{error}",
    "common_export_success_title": "Export réussi",
    "common_print_failed": "La préparation de l'impression a échoué :\n{error}",
    "common_reset_button": "Réinitialiser",
    "common_want_label": "Je veux :",
    "common_exclude_label": "Je ne veux pas :",
    "common_tags_filter_label": "Étiquettes (toutes requises) :",
    "common_filter_hint": "Tapez les premières lettres pour filtrer la liste.",
    "common_search_label": "🔍 Rechercher :",
    "common_sort_by_label": "Trier par :",
    "common_category_label": "Catégorie :",
    "common_edit_button": "✏️ Modifier",
    "common_unknown_ingredient_title": "Ingrédient inconnu",
    "common_unknown_ingredient_simple_message": (
        "« {name} » ne correspond à aucun ingrédient enregistré.\nChoisissez-en un dans la liste déroulante."
    ),
    "common_ingredient_label": "Ingrédient :",
    "common_quantity_label": "Quantité :",
    "common_unit_label": "Unité :",
    "common_new_ingredient_button": "🥕 Nouvel ingrédient",
    "common_save_button": "💾 Enregistrer",

    # ---- PantryWindow (Mon garde-manger) ----
    "pantry_title": "Mon garde-manger",
    "pantry_heading": "📦 Mon garde-manger",
    "pantry_intro": (
        "Indiquez ce que vous avez chez vous et en quelle quantité.\n"
        "« Que puis-je cuisiner ? » pourra alors vérifier si vous en avez assez,\n"
        "et proposer de décompter automatiquement le stock après avoir cuisiné."
    ),
    "pantry_threshold_label": "Seuil d'alerte (optionnel) :",
    "pantry_help_text": (
        "Pour AJOUTER un article : indiquez l'ingrédient (créez-le d'abord avec\n"
        "« 🥕 Nouvel ingrédient » s'il n'est pas encore dans votre liste), la\n"
        "quantité et l'unité, puis cliquez sur « 💾 Enregistrer ».\n"
        "Pour MODIFIER un article déjà présent : cliquez une fois dessus dans la\n"
        "liste ci-dessous — cela charge ses valeurs dans les champs ci-dessus,\n"
        "sans rien enregistrer : changez les valeurs souhaitées PUIS cliquez sur\n"
        "« 💾 Enregistrer » pour que le changement soit pris en compte.\n"
        "Le seuil d'alerte déclenche un rappel sur la page d'accueil dès que la\n"
        "quantité passe en dessous (laissez vide pour ne jamais être alerté)."
    ),
    "pantry_remove_button": "🗑 Retirer du garde-manger",
    "pantry_empty": "Votre garde-manger est vide pour le moment.",
    "pantry_threshold_suffix": " (seuil : {threshold})",
    "pantry_error_ingredient_required": "Merci d'indiquer un ingrédient.",
    "pantry_error_invalid_quantity": "Quantité invalide.",
    "pantry_error_invalid_threshold": "Seuil d'alerte invalide (laissez vide si vous n'en voulez pas).",
    "pantry_select_ingredient_first": "Sélectionnez un ingrédient dans la liste.",
    "pantry_remove_confirm_message": "Retirer « {name} » du garde-manger ?",

    # ---- WhatCanICookWindow (Que puis-je cuisiner ?) ----
    "cook_title": "Que puis-je cuisiner ?",
    "cook_instructions_label": "Indiquez les ingrédients que vous avez chez vous :",
    "cook_staples_hint": (
        "Quelques ingrédients de base courants sont déjà cochés ci-contre\n"
        "(sel, huile, farine...) — retirez ceux que vous n'avez pas."
    ),
    "cook_all_ingredients_label": "Tous les ingrédients :",
    "cook_add_button": "➕ Ajouter →",
    "cook_have_label": "Ce que j'ai :",
    "cook_remove_button": "🗑 Retirer",
    "cook_load_from_pantry_button": "📦 Charger depuis mon garde-manger",
    "cook_compute_button": "🔍 Voir les recettes réalisables",
    "cook_open_selected_button": "📖 Consulter la recette sélectionnée",
    "cook_pantry_empty_title": "Info",
    "cook_pantry_empty_message": (
        "Votre garde-manger est vide pour le moment. Ouvrez « 📦 Mon "
        "garde-manger » depuis la page d'accueil pour y ajouter des ingrédients."
    ),
    "cook_loaded_title": "Chargé",
    "cook_loaded_message": "{count} ingrédient(s) ajouté(s) depuis votre garde-manger.",
    "cook_add_ingredient_first": "Ajoutez au moins un ingrédient que vous avez.",
    "cook_feasible_header": "✅ Réalisables avec ce que vous avez :",
    "cook_insufficient_quantity": "  ⚠️ quantité insuffisante : {list}",
    "cook_none_feasible": "Aucune recette n'est réalisable à 100 % avec ces ingrédients.",
    "cook_substitutable_header": "🔄 Réalisables en utilisant un substitut :",
    "cook_almost_header": "🟡 Presque (il manque 1 à 3 ingrédients) :",
    "cook_missing_label": "   {name} (manque : {list})",
    "cook_no_results": "Essayez d'ajouter d'autres ingrédients à votre sélection.",
    "cook_select_recipe_from_results": "Sélectionnez une recette dans la liste des résultats.",
    "cook_select_recipe_row": "Sélectionnez une ligne correspondant à une recette.",

    # ---- WeeklyPlanHistoryWindow (Historique des semaines passées) ----
    "weekhistory_title": "Historique des semaines passées",
    "weekhistory_heading": "🕘 Historique des semaines passées",
    "weekhistory_intro": (
        "Chaque semaine où vous enregistrez le planning est archivée ici\n"
        "automatiquement (jusqu'à 26 semaines, environ 6 mois), pour éviter\n"
        "de refaire deux fois la même chose de trop près."
    ),
    "weekhistory_reload_button": "♻️ Recharger dans le planning actuel",
    "weekhistory_delete_button": "🗑 Supprimer cette semaine",
    "weekhistory_no_archived_weeks": "Aucune semaine archivée.",
    "weekhistory_week_label": "Semaine {week}",
    "weekhistory_saved_on": "Enregistré le {date}\n\n",
    "weekhistory_day_heading": "{day} :\n",
    "weekhistory_slot_line": "   {slot} : {recipe} ({persons} pers.)\n",
    "weekhistory_empty_week": "(Planning vide pour cette semaine.)",
    "weekhistory_select_week_first": "Sélectionnez une semaine dans la liste.",
    "weekhistory_reload_confirm_message": (
        "Recharger le planning de la semaine {week} dans le "
        "planning actuel ?\n\nCela remplacera les recettes actuellement affichées "
        "(pensez à enregistrer le planning en cours avant, si vous voulez le garder)."
    ),
    "weekhistory_delete_confirm_message": "Supprimer définitivement l'archive de la semaine {week} ?",

    # ---- WeeklyPlanTemplatesWindow (Modèles de semaine) ----
    "weektemplates_title": "Modèles de semaine",
    "weektemplates_heading": "📋 Modèles de semaine",
    "weektemplates_intro": (
        "Enregistrez le planning actuellement affiché comme modèle\n"
        "réutilisable, pour l'appliquer d'un clic à une autre semaine\n"
        "plutôt que de tout resaisir."
    ),
    "weektemplates_name_label": "Nom du nouveau modèle :",
    "weektemplates_save_button": "💾 Enregistrer le planning actuel comme modèle",
    "weektemplates_apply_button": "📋 Appliquer ce modèle",
    "weektemplates_delete_button": "🗑 Supprimer ce modèle",
    "weektemplates_none_saved": "Aucun modèle enregistré pour le moment.",
    "weektemplates_error_name_required": "Merci d'indiquer un nom pour ce modèle.",
    "weektemplates_empty_plan": "Le planning actuellement affiché est vide : rien à enregistrer comme modèle.",
    "weektemplates_saved_message": "Modèle « {name} » enregistré.",
    "weektemplates_select_template_first": "Sélectionnez un modèle dans la liste.",
    "weektemplates_apply_confirm_message": (
        "Appliquer le modèle « {name} » au planning actuel ?\n\n"
        "Cela remplacera les recettes actuellement affichées (pensez à "
        "enregistrer le planning en cours avant, si vous voulez le garder)."
    ),
    "weektemplates_delete_confirm_message": "Supprimer définitivement le modèle « {name} » ?",
    "common_none_option": "-- Aucune --",

    # ---- WeeklyPlanWindow (Planning de la semaine) ----
    "weekplan_title": "Planning de la semaine",
    "weekplan_subtitle": "Vue calendrier : jours en colonnes, repas en lignes.",
    "weekplan_save_button": "💾 Enregistrer le planning",
    "weekplan_clear_button": "🗑 Tout effacer",
    "weekplan_export_ics_button": "📆 Exporter vers un calendrier (.ics)",
    "weekplan_compute_button": "Calculer la liste de courses de la semaine",
    "weekplan_checklist_button": "☑️ Mode courses",
    "weekplan_empty_list_message": (
        "Aucune liste calculée pour le moment.\n"
        "Cliquez sur « Calculer la liste de courses de la semaine » ci-dessus,\n"
        "ou chargez une liste enregistrée."
    ),
    "weekplan_total_list_heading": "=== Liste de courses de la semaine ===",
    "weekplan_calculate_list_for_export": (
        "Calculez d'abord une liste de courses (bouton « Calculer la liste de courses de la semaine »)."
    ),
    "weekplan_invalid_persons_for_slot": "Nombre de personnes invalide pour {day} — {slot}.",
    "weekplan_saved_message": "Le planning de la semaine a été enregistré.",
    "weekplan_clear_confirm_message": "Effacer tout le planning de la semaine ?",
    "weekplan_assign_recipe_first": "Assignez au moins une recette à un créneau de la semaine.",
    "weekplan_export_ics_title": "Exporter le planning vers un calendrier",
    "weekplan_ics_export_success_message": (
        "Planning exporté :\n{path}\n\n"
        "Importez ce fichier dans Google Agenda, Outlook ou Calendrier "
        "pour voir vos repas s'y répéter chaque semaine."
    ),
    "weekplan_assign_or_manual": (
        "Assignez au moins une recette à un créneau de la semaine, "
        "ou ajoutez un ingrédient manuellement."
    ),
    "weekplan_export_shopping_list_title": "Enregistrer la liste de courses",
    "weekplan_shopping_list_title": "Liste de courses de la semaine",
    "weekplan_list_saved_message": "Liste enregistrée :\n{path}",
    "weekplan_excel_module_missing": "L'export Excel nécessite : pip install openpyxl",
    "weekplan_pdf_module_missing": "L'export PDF nécessite : pip install reportlab",
    "weekplan_print_module_missing": "L'impression nécessite : pip install reportlab",
    "weekplan_print_label": "la liste de courses de la semaine",

    # ---- ManageIngredientsWindow (Gérer les ingrédients) ----
    "manageing_title": "Gérer les ingrédients",
    "manageing_list_label": "Liste des ingrédients enregistrés :",
    "manageing_add_button": "➕ Ajouter",
    "manageing_edit_button": "✏️ Modifier",
    "manageing_delete_button": "🗑️ Supprimer",
    "manageing_load_defaults_button": "📚 Charger les ~1000 ingrédients courants",
    "manageing_spell_check_button": "🔤 Vérifier les doublons / fautes de frappe",
    "manageing_prices_button": "💰 Gérer les prix (pour le coût des recettes)",
    "manageing_substitutions_button": "🔄 Gérer les substitutions",
    "manageing_edit_hint": (
        "\"Modifier\" permet de changer le nom (mis à jour\n"
        "partout où l'ingrédient est utilisé), ses allergènes,\n"
        "ses valeurs nutritionnelles et son prix."
    ),
    "manageing_select_ingredient_first": "Sélectionnez un ingrédient dans la liste.",
    "manageing_delete_confirm_message": "Supprimer « {name} » de la liste des ingrédients ?",
    "manageing_delete_usage_warning": (
        "\n\nAttention : il est utilisé dans {count} recette(s). "
        "Ces recettes conserveront cet ingrédient, mais il ne sera "
        "plus proposé dans le menu déroulant, sauf si vous le rajoutez."
    ),
    "manageing_missing_file_title": "Fichier manquant",
    "manageing_missing_file_message": (
        "Le fichier ingredients_par_defaut.json est introuvable.\n"
        "Assurez-vous qu'il se trouve dans le même dossier que main.py."
    ),
    "manageing_done_title": "Terminé",
    "manageing_defaults_added_message": "{count} nouvel(aux) ingrédient(s) ajouté(s) à partir de la liste courante.",
    "manageing_defaults_none_added": "Tous les ingrédients courants étaient déjà présents.",

    # ---- SubstitutionEditWindow (Substituts pour un ingrédient précis) ----
    "subedit_title": "Substituts pour « {name} »",
    "subedit_heading": "🔄 Substituts pour « {name} »",
    "subedit_disclaimer": (
        "Une substitution est un conseil culinaire, pas une équivalence\n"
        "garantie : le résultat peut varier selon la recette."
    ),
    "subedit_remove_button": "🗑 Retirer le substitut sélectionné",
    "subedit_add_frame_title": "Ajouter un substitut",
    "subedit_name_label": "Nom :",
    "subedit_note_label": "Note (optionnelle) :",
    "subedit_add_to_list_button": "➕ Ajouter à la liste",
    "subedit_revert_button": "🔄 Revenir à la base fournie",
    "subedit_cancel_button": "Annuler",
    "subedit_no_substitute_yet": "Aucun substitut pour le moment.",
    "subedit_error_name_required": "Merci d'indiquer un nom de substitut.",
    "subedit_select_to_remove": "Sélectionnez un substitut à retirer.",
    "subedit_revert_confirm_message": (
        "Retirer votre liste personnalisée et revenir aux substituts fournis\n"
        "avec l'application pour « {name} » ?"
    ),

    # ---- ManageSubstitutionsWindow (Gérer les substitutions) ----
    "managesub_title": "Gérer les substitutions",
    "managesub_heading": "🔄 Substitutions d'ingrédients",
    "managesub_intro": (
        "Consultez ou modifiez les substituts suggérés pour un ingrédient.\n"
        "Une substitution est un conseil culinaire, pas une équivalence garantie."
    ),
    "managesub_manage_button": "✏️ Gérer ses substituts",
    "managesub_hint": (
        "Double-cliquez sur un ingrédient de la liste pour voir ou modifier ses\n"
        "substituts, ou tapez un nom ci-dessus (y compris un ingrédient qui n'a\n"
        "pas encore de substitut connu) puis « ✏️ Gérer ses substituts »."
    ),
    "managesub_none_with_substitute": "Aucun ingrédient avec substitut pour le moment.",
    "managesub_substitute_count": "{name} ({count} substitut{plural})",
    "managesub_error_ingredient_required": "Merci d'indiquer un ingrédient.",
    "managesub_unknown_ingredient_message": (
        "« {name} » ne correspond à aucun ingrédient enregistré.\n"
        "Choisissez-en un dans la liste déroulante, ou créez-le d'abord "
        "depuis « 🥕 Gérer les ingrédients »."
    ),

    # ---- IngredientPricesWindow (Gérer les prix des ingrédients) ----
    "ingprices_title": "Gérer les prix des ingrédients",
    "ingprices_heading": "💰 Prix des ingrédients",
    "ingprices_intro": (
        "Renseignez un prix pour les ingrédients qui vous\n"
        "intéressent — inutile de tous les faire. Le coût\n"
        "d'une recette est estimé à partir de ces prix."
    ),
    "ingprices_price_label": "Prix (€) :",
    "ingprices_for_one_label": "pour 1",
    "ingprices_save_button": "💾 Enregistrer le prix",
    "ingprices_clear_button": "🗑 Effacer le prix",
    "ingprices_units_note": (
        "kg ↔ recettes en Gr   ·   L ↔ recettes en cl   ·   les prix\n"
        "en pièce/cuillère s'appliquent tels quels."
    ),
    "ingprices_no_price_set": "  —  (prix non renseigné)",
    "ingprices_price_suffix": "  —  {price} € / {unit}",
    "ingprices_error_invalid_price": "Entrez un prix valide (nombre positif).",
    "ingprices_saved_message": "Prix enregistré pour « {name} ».",

    # ---- IngredientEditWindow (Nouvel ingrédient / Modifier un ingrédient) ----
    "ingedit_title_edit": "Modifier un ingrédient",
    "ingedit_title_new": "Nouvel ingrédient",
    "ingedit_heading_edit": "✏️ Modifier l'ingrédient",
    "ingedit_heading_new": "➕ Nouvel ingrédient",
    "ingedit_name_label": "Nom :",
    "ingedit_allergens_label": "Allergènes présents :",
    "ingedit_nutrition_label": "Valeurs nutritionnelles (pour 100 g / 100 ml) :",
    "ingedit_nutri_kcal": "Calories (kcal)",
    "ingedit_nutri_protein": "Protéines (g)",
    "ingedit_nutri_carbs": "Glucides (g)",
    "ingedit_nutri_fat": "Lipides (g)",
    "ingedit_nutrition_hint": "Laissez vide si vous ne connaissez pas ces valeurs.",
    "ingedit_price_label": "Prix :",
    "ingedit_save_button": "💾 Enregistrer",
    "ingedit_delete_button": "🗑️ Supprimer cet ingrédient",
    "ingedit_error_invalid_field": "« {field} » doit être un nombre positif (ou vide).",
    "ingedit_error_name_required": "Merci d'indiquer un nom d'ingrédient.",
    "ingedit_error_already_exists": "L'ingrédient « {name} » existe déjà.",
    "ingedit_error_plural_duplicate": (
        "« {name} » n'est qu'une variante singulier/pluriel de "
        "l'ingrédient déjà existant « {existing} ». Pour éviter les "
        "doublons dans la liste, utilisez directement « {existing} »."
    ),
    "ingedit_nutri_field_kcal": "Calories",
    "ingedit_nutri_field_protein": "Protéines",
    "ingedit_nutri_field_carbs": "Glucides",
    "ingedit_nutri_field_fat": "Lipides",
    "ingedit_error_invalid_price": "Le prix doit être un nombre positif (ou vide).",
    "ingedit_saved_message": "« {name} » a été enregistré.",

    # ---- IngredientSpellCheckWindow (Vérification orthographique) ----
    "spellcheck_title": "Vérification orthographique des ingrédients",
    "spellcheck_heading": (
        "Paires d'ingrédients qui se ressemblent à 90 % ou plus\n"
        "(doublons probables ou fautes de frappe) :"
    ),
    "spellcheck_multi_select_hint": (
        "Sélection multiple possible (Ctrl+clic ou Maj+clic) pour\n"
        "fusionner plusieurs paires d'un coup."
    ),
    "spellcheck_merge_button": "🔗 Fusionner la sélection",
    "spellcheck_not_duplicate_button": "✕ Ce n'est pas un doublon",
    "spellcheck_rerun_button": "🔄 Relancer l'analyse",
    "spellcheck_footer_hint": (
        "Pour une seule paire, on vous demande laquelle des deux\n"
        "graphies garder. Pour plusieurs paires à la fois, l'ingrédient\n"
        "le moins utilisé dans vos recettes est automatiquement fusionné\n"
        "vers celui utilisé dans le plus de recettes.\n"
        "« Ce n'est pas un doublon » retire définitivement la ou les\n"
        "paires sélectionnées de cette analyse, aujourd'hui et à l'avenir."
    ),
    "spellcheck_none_found": "Aucun doublon probable détecté. 🎉",
    "spellcheck_pair_line": "{a}   ↔   {b}     ({percent} % similaires)",
    "spellcheck_select_pair_first": "Sélectionnez au moins une paire dans la liste.",
    "spellcheck_dismissed_message": (
        "{count} paire(s) marquée(s) comme n'étant pas des "
        "doublons. Elles ne seront plus proposées lors des prochaines analyses."
    ),
    "spellcheck_merge_dialog_title": "Fusionner",
    "spellcheck_merge_dialog_message": (
        "Fusionner « {a} » et « {b} » ?\n\n"
        "Oui = tout renommer en « {a} »\n"
        "Non = tout renommer en « {b} »\n"
        "Annuler = ne rien faire"
    ),
    "spellcheck_merged_title": "Fusionné",
    "spellcheck_merged_one_message": "« {removed} » a été fusionné avec « {kept} ».",
    "spellcheck_merge_multi_confirm": (
        "Fusionner automatiquement ces {count} paires ?\n\n"
        "Pour chaque paire, l'ingrédient le moins utilisé dans vos "
        "recettes sera fusionné vers celui utilisé dans le plus de "
        "recettes (le premier par ordre alphabétique en cas d'égalité)."
    ),
    "spellcheck_merged_multi_message": "{count} paire(s) fusionnée(s).",

    # ---- CompareRecipesWindow (Comparer deux recettes) ----
    "compare_title": "Comparer deux recettes",
    "compare_recipe_a_label": "Recette A :",
    "compare_recipe_b_label": "Recette B :",
    "compare_button": "⚖️ Comparer",
    "compare_choose_each_list": "Choisissez une recette dans chaque liste.",
    "compare_field_category": "Catégorie :",
    "compare_field_favorite": "Favori :",
    "compare_yes": "⭐ Oui",
    "compare_no": "Non",
    "compare_field_rating": "Note :",
    "compare_field_difficulty": "Difficulté :",
    "compare_field_prep": "Préparation :",
    "compare_field_cook": "Cuisson :",
    "compare_field_total_time": "Temps total :",
    "compare_field_cooked": "Cuisinée :",
    "compare_times_suffix": "{count} fois",
    "compare_field_cost": "Coût estimé :",
    "compare_field_nutrition": "Nutrition (kcal) :",
    "compare_field_ingredient_count": "Nb. ingrédients :",
    "compare_common_ingredients": "🟰 Communs ({count})",
    "compare_only_a": "🅰️ Uniquement « {name} » ({count})",
    "compare_only_b": "🅱️ Uniquement « {name} » ({count})",
    "compare_none": "Aucun",

    # ---- StatisticsWindow (Statistiques) ----
    "stats_title": "Statistiques",
    "stats_heading": "=== Statistiques ===\n\n",
    "stats_total_recipes": "Nombre total de recettes : {count}\n\n",
    "stats_by_category": "Répartition par catégorie :\n",
    "stats_category_line": "  - {category} : {count}\n",
    "stats_by_difficulty": "Répartition par difficulté :\n",
    "stats_difficulty_line": "  - {difficulty} : {count}\n",
    "stats_difficulty_unspecified": "Non renseignée",
    "stats_favorites_count": "Recettes favorites : {count}\n\n",
    "stats_avg_rating": "Note moyenne (recettes notées) : {avg} / 5 ({count} recette(s) notée(s))\n\n",
    "stats_no_rated_recipe": "Note moyenne : aucune recette notée pour le moment.\n\n",
    "stats_five_star_heading": "Recette(s) notée(s) 5 étoiles :\n",
    "stats_recipe_line": "  - {name}\n",
    "stats_most_cooked_heading": "Recettes les plus cuisinées :\n",
    "stats_cooked_line": "  - {name} : {count} fois\n",
    "stats_none_cooked_yet": (
        "  Aucune recette marquée comme cuisinée pour le moment.\n"
        "  (bouton « 🍳 J'ai cuisiné ça ! » dans « Voir une recette précise »)\n"
    ),
    "stats_most_used_tags_heading": "Étiquettes les plus utilisées :\n",
    "stats_tag_line": "  - {tag} : {count}\n",
    "stats_never_cooked_heading": "🕸️ Recettes jamais cuisinées :\n",
    "stats_and_others": "  ... et {count} autre(s)\n",
    "stats_all_cooked": "  Toutes vos recettes ont déjà été cuisinées au moins une fois. 👏\n",
    "stats_stale_heading": "🕰️ Pas cuisinées depuis plus de {days} jours :\n",
    "stats_stale_line": "  - {name} (il y a {days} jours)\n",
    "stats_no_stale_recipe": "  Aucune recette dans ce cas pour le moment.\n",
    "stats_avg_cost_heading": "💰 Coût moyen par personne :\n",
    "stats_avg_cost_line": (
        "  {avg} € en moyenne, sur {count} recette(s) avec au "
        "moins un prix connu ({without_price} sans prix renseigné)\n"
    ),
    "stats_no_priced_recipe": (
        "  Aucune recette avec un prix renseigné pour le moment.\n"
        "  (voir « 💰 Gérer les prix » dans « Gérer les ingrédients »)\n"
    ),
    "stats_avg_kcal_heading": "🥗 Calories moyennes par personne :\n",
    "stats_avg_kcal_line": (
        "  {avg} kcal en moyenne, sur {count} recette(s) avec des "
        "ingrédients reconnus dans la base nutritionnelle\n"
    ),
    "stats_no_recognized_recipe": "  Aucune recette avec des ingrédients reconnus pour le moment.\n",
    "stats_monthly_chart_title": "📈 Recettes cuisinées par mois (12 derniers mois)",
    "stats_heatmap_title": "🗓️ Calendrier des jours cuisinés (12 derniers mois)",
    "stats_heatmap_legend": "Moins ⬜ 🟨 🟧 🟥 Plus",
    "stats_day_labels": "L,M,M,J,V,S,D",
    "stats_month_labels_short": "Jan,Fév,Mar,Avr,Mai,Jun,Jul,Aoû,Sep,Oct,Nov,Déc",
    "stats_month_labels_lower": "jan,fév,mar,avr,mai,jun,jul,aoû,sep,oct,nov,déc",

    # ---- draw_recipe_content (contenu PDF d'une recette, export seule ou dans le livre de recettes) ----
    "recipepdf_category_persons": "Catégorie : {cat}    Pour {persons} personne(s)",
    "recipepdf_rating": "Note : {stars}",
    "recipepdf_prep": "Préparation : {time} min",
    "recipepdf_cook": "Cuisson : {time} min",
    "recipepdf_difficulty": "Difficulté : {value}",
    "recipepdf_allergens": "⚠ Allergènes : {list}",
    "recipepdf_ingredients_heading": "Ingrédients :",
    "recipepdf_cost": "Coût estimé : {cost} €{partial}",
    "recipepdf_partial_suffix": " (partiel, {known}/{total})",
    "recipepdf_nutrition": "Nutrition estimée{partial} : {kcal} kcal · {protein}g prot. · {carbs}g gluc. · {fat}g lip.",
    "recipepdf_description_heading": "Description :",
    "recipepdf_notes_heading": "Notes personnelles :",

    # ---- build_cookbook_pdf (livre de recettes PDF) ----
    "cookbookpdf_page_number": "Page {current} / {total}",
    "cookbookpdf_generated_on": "Généré le {date}",
    "cookbookpdf_summary_heading": "Sommaire",
    "cookbookpdf_summary_line": "- [{cat}] {name}",

    # ---- CookbookExportWindow (Exporter le livre de recettes) ----
    "cookbookexport_title": "Exporter le livre de recettes",
    "cookbookexport_heading": "📖 Exporter le livre de recettes",
    "cookbookexport_intro": (
        "Sélectionnez les recettes à inclure dans un seul PDF,\n"
        "façon livre de cuisine."
    ),
    "cookbookexport_filter_label": "Filtrer par catégorie :",
    "cookbookexport_check_all_button": "Tout cocher",
    "cookbookexport_uncheck_all_button": "Tout décocher",
    "cookbookexport_generate_button": "📄 Générer le PDF du livre",
    "cookbookexport_error_select_recipe": "Sélectionnez au moins une recette.",
    "cookbookexport_save_dialog_title": "Enregistrer le livre de recettes",
    "cookbookexport_saved_message": "Livre de recettes enregistré :\n{path}",

    # ---- ImportExportWindow (Importer / Exporter les données) ----
    "importexport_title": "Importer / Exporter les données",
    "importexport_heading": "Sauvegarder ou transférer vos données",
    "importexport_export_intro": (
        "L'export crée un fichier .zip contenant absolument toutes\n"
        "vos données : recettes, photos, ingrédients personnalisés,\n"
        "prix, substituts, garde-manger, planning et son historique,\n"
        "menus, listes de courses enregistrées, corbeille et\n"
        "réglages — pour tout sauvegarder ou tout transférer sur un\n"
        "autre ordinateur en un seul fichier."
    ),
    "importexport_export_button": "📤 Exporter toutes mes données (.zip)",
    "importexport_import_intro": (
        "L'import lit un fichier .zip précédemment exporté.\n"
        "\"Fusionner\" ajoute les recettes/photos en double sous un\n"
        "nouveau nom plutôt que de les perdre, et complète le reste\n"
        "(garde-manger, menus, listes...) sans rien supprimer.\n"
        "\"Remplacer\" écrase tout, y compris les réglages et le\n"
        "planning en cours."
    ),
    "importexport_import_button": "📥 Importer des données (.zip)",
    "importexport_auto_backups_heading": "🗄️ Sauvegardes automatiques",
    "importexport_auto_backups_intro": (
        "Une sauvegarde est créée automatiquement au démarrage de\n"
        "l'application (au maximum une par {hours}h), et les "
        "{retention} plus\nrécentes sont conservées ici."
    ),
    "importexport_backup_now_button": "💾 Sauvegarder maintenant",
    "importexport_restore_selected_button": "♻️ Restaurer la sélection",
    "importexport_cloud_heading": "☁️ Sauvegarde automatique dans le cloud",
    "importexport_cloud_intro": (
        "Choisissez un dossier synchronisé par un client déjà\n"
        "installé sur ce PC (Google Drive, OneDrive, Dropbox...).\n"
        "Chaque sauvegarde automatique y sera aussi copiée, et ce\n"
        "client se chargera de l'envoyer dans le cloud tout seul."
    ),
    "importexport_choose_cloud_button": "📁 Choisir un dossier cloud",
    "importexport_disable_button": "🚫 Désactiver",
    "importexport_cloud_enabled": "✅ Activé : {folder}",
    "importexport_cloud_not_configured": "Non configuré pour le moment.",
    "importexport_choose_folder_title": "Choisir un dossier synchronisé (Google Drive, OneDrive, Dropbox...)",
    "importexport_cloud_configured_title": "Dossier configuré",
    "importexport_cloud_configured_message": (
        "Dossier cloud configuré :\n{folder}\n\n"
        "Voulez-vous y copier une sauvegarde dès maintenant ?"
    ),
    "importexport_disabled_title": "Désactivé",
    "importexport_disabled_message": "La sauvegarde automatique dans le cloud est désactivée.",
    "importexport_backup_date_line": "{date}   ({size} Ko)",
    "importexport_no_backups": "Aucune sauvegarde automatique pour le moment.",
    "importexport_backup_failed": "La sauvegarde a échoué :\n{error}",
    "importexport_backup_created_title": "Sauvegarde créée",
    "importexport_backup_created_message": "Une nouvelle sauvegarde automatique a été créée.",
    "importexport_select_backup_first": "Sélectionnez une sauvegarde dans la liste.",
    "importexport_restore_mode_title": "Mode de restauration",
    "importexport_restore_mode_message": (
        "Comment restaurer cette sauvegarde ?\n\n"
        "Oui = Fusionner (ajouter aux données actuelles, sans rien supprimer)\n"
        "Non = Remplacer entièrement les données actuelles\n"
        "Annuler = ne rien faire"
    ),
    "importexport_restore_failed": "La restauration a échoué :\n{error}",
    "importexport_restore_done_title": "Restauration terminée",
    "importexport_restore_done_message": "Les données ont été restaurées avec succès.",
    "importexport_export_data_title": "Exporter mes données",
    "importexport_export_data_success": "Vos données ont été exportées vers :\n{path}",
    "importexport_choose_archive_title": "Choisir une archive à importer",
    "importexport_import_mode_title": "Mode d'import",
    "importexport_import_mode_message": (
        "Comment importer ces données ?\n\n"
        "Oui = Fusionner (ajouter aux données actuelles, sans rien supprimer)\n"
        "Non = Remplacer entièrement les données actuelles\n"
        "Annuler = ne rien faire"
    ),
    "importexport_import_failed": "L'import a échoué :\n{error}",
    "importexport_import_done_title": "Import terminé",
    "importexport_import_done_message": "Les données ont été importées avec succès.",

    # ---- ShoppingChecklistWindow (Mode courses) ----
    "checklist_instruction": "Cochez chaque article au fur et à mesure de vos courses.",
    "checklist_check_all_button": "☑️ Tout cocher",
    "checklist_uncheck_all_button": "⬜ Tout décocher",
    "checklist_progress_label": "{done} / {total} article(s) coché(s)",

    # ---- ExportFormatDialog (Choisir un format d'export) ----
    "exportformat_title": "Choisir un format d'export",
    "exportformat_heading": "📤 Exporter la liste de courses",
    "exportformat_choose_label": "Choisissez le format d'export souhaité :",
    "exportformat_txt_button": "📝 Exporter en texte (.txt)",
    "exportformat_excel_button": "📊 Exporter en Excel (.xlsx)",
    "exportformat_pdf_button": "📄 Exporter en PDF (.pdf)",
    "exportformat_cancel_button": "Annuler",

    # ---- MenuManagerWindow (Mes menus) ----
    "menumanager_title": "Mes menus",
    "menumanager_list_label": "Mes menus enregistrés :",
    "menumanager_new_button": "➕ Nouveau menu",
    "menumanager_recipe_count": "{name} ({count} recette(s))",
    "menumanager_select_menu_first": "Sélectionnez un menu dans la liste.",
    "menumanager_delete_confirm": "Supprimer le menu « {name} » ?",

    # ---- MenuFormWindow (Nouveau menu / Modifier le menu) ----
    "menuform_title_edit": "Modifier le menu",
    "menuform_title_new": "Nouveau menu",
    "menuform_name_label": "Nom du menu :",
    "menuform_add_recipe_label": "Ajouter une recette au menu :",
    "menuform_persons_short_label": "pers. :",
    "menuform_add_button": "+ Ajouter",
    "menuform_recipes_label": "Recettes du menu :",
    "menuform_remove_button": "🗑 Retirer du menu",
    "menuform_save_button": "💾 Enregistrer le menu",
    "menuform_compute_button": "Calculer la liste de courses du menu",
    "menuform_empty_list_message": (
        "Aucune liste calculée pour le moment.\n"
        "Cliquez sur « Calculer la liste de courses du menu » ci-dessus,\n"
        "ou chargez une liste enregistrée."
    ),
    "menuform_total_list_heading": "=== Liste de courses du menu ===",
    "menuform_item_row_label": "[{cat}] {name} ({persons} pers.)",
    "menuform_select_recipe_to_remove": "Sélectionnez une recette du menu à retirer.",
    "menuform_error_name_required": "Merci d'indiquer un nom de menu.",
    "menuform_error_no_recipe": "Ajoutez au moins une recette au menu.",
    "menuform_saved_message": "Le menu « {name} » a été enregistré.",
    "menuform_calculate_list_for_export": (
        "Calculez d'abord une liste de courses (bouton « Calculer la liste de courses du menu »)."
    ),
    "menuform_add_recipe_or_manual": "Ajoutez au moins une recette au menu, ou ajoutez un ingrédient manuellement.",
    "menuform_shopping_list_title": "Menu : {name}",
    "menuform_print_label": "le menu « {name} »",

    # ---- ImportFromUrlWindow (Importer une recette depuis un lien) ----
    "importurl_title": "Importer une recette depuis un lien",
    "importurl_heading": "🌐 Importer une recette depuis un lien",
    "importurl_intro": (
        "Collez l'adresse (URL) d'une page de recette. Cela fonctionne\n"
        "avec la plupart des grands sites de cuisine (qui utilisent un\n"
        "format de données standard). Une connexion internet est requise."
    ),
    "importurl_fetch_button": "🌐 Récupérer la recette",
    "importurl_after_import_note": (
        "Après import, vérifiez et complétez la recette si besoin\n"
        "(le repérage des quantités et unités n'est pas toujours parfait)."
    ),
    "importurl_paste_url_first": "Collez d'abord une adresse internet (URL).",
    "importurl_fetching": "Récupération en cours...",
    "importurl_failed_title": "Échec de l'import",

    # ---- ImportFromPhotoWindow (Importer une recette depuis une photo) ----
    "importphoto_title": "Importer une recette depuis une photo",
    "importphoto_heading": "📷 Importer une recette depuis une photo",
    "importphoto_intro": (
        "Prenez en photo (ou scannez) une recette manuscrite ou une\n"
        "page de livre de cuisine, puis choisissez l'image ici. Le texte\n"
        "en est extrait automatiquement, mais reste à relire et organiser\n"
        "vous-même (contrairement à l'import depuis un lien, une photo n'a\n"
        "pas de structure ingrédients/étapes que l'on puisse deviner)."
    ),
    "importphoto_module_warning": (
        "⚠ Cette fonctionnalité nécessite le module 'pytesseract'\n"
        "ET le programme Tesseract OCR installé séparément sur ce PC.\n"
        "Voir le LISEZ-MOI pour les instructions d'installation."
    ),
    "importphoto_no_photo_chosen": "Aucune photo choisie",
    "importphoto_choose_button": "📁 Choisir une photo",
    "importphoto_extract_button": "🔍 Extraire le texte",
    "importphoto_extracted_text_label": "Texte extrait (modifiable) :",
    "importphoto_create_button": "➡️ Créer la recette avec ce texte",
    "importphoto_choose_photo_title": "Choisir une photo de recette",
    "importphoto_choose_first": "Choisissez d'abord une photo.",
    "importphoto_ocr_module_missing": (
        "Cette fonctionnalité nécessite le module 'pytesseract'\n"
        "(pip install pytesseract) ET le programme Tesseract OCR\n"
        "installé séparément sur ce PC. Voir le LISEZ-MOI."
    ),
    "importphoto_extraction_failed_title": "Échec de l'extraction",
    "importphoto_extraction_failed_message": (
        "La reconnaissance de texte a échoué. Vérifiez que Tesseract OCR "
        "est bien installé sur ce PC et accessible.\n\nDétail : {error}"
    ),
    "importphoto_no_text_extracted": (
        "Aucun texte n'a pu être extrait de cette photo. Essayez une image "
        "plus nette, mieux cadrée ou mieux éclairée."
    ),
    "importphoto_no_text_title": "Aucun texte",
    "importphoto_no_text_confirm": (
        "Aucun texte n'a été extrait ou saisi. Créer quand même une "
        "recette vide (avec juste la photo) ?"
    ),

    # ---- TrashWindow (Corbeille) ----
    "trash_title": "Corbeille",
    "trash_heading": "🗑️ Recettes supprimées",
    "trash_intro": (
        "Les photos des recettes de la corbeille sont conservées\n"
        "jusqu'à leur suppression définitive."
    ),
    "trash_restore_button": "♻️ Restaurer",
    "trash_delete_forever_button": "🗑️ Supprimer définitivement",
    "trash_empty_button": "🧹 Vider la corbeille",
    "trash_unnamed_recipe": "(sans nom)",
    "trash_unknown_date": "date inconnue",
    "trash_entry_line": "{name}  —  supprimée le {date}",
    "trash_is_empty": "La corbeille est vide.",
    "trash_select_recipe_first": "Sélectionnez une recette dans la corbeille.",
    "trash_restored_suffix": "{name} (restaurée)",
    "trash_restored_title": "Restaurée",
    "trash_restored_message": "« {name} » a été restaurée.",
    "trash_delete_forever_confirm": "Supprimer définitivement « {name} » ?\n\nCette action est irréversible.",
    "trash_deleted_title": "Supprimée",
    "trash_deleted_message": "La recette a été définitivement supprimée.",
    "trash_already_empty": "La corbeille est déjà vide.",
    "trash_empty_confirm": (
        "Supprimer définitivement les {count} recette(s) de la corbeille ?\n\n"
        "Cette action est irréversible."
    ),
    "trash_emptied_title": "Corbeille vidée",
    "trash_emptied_message": "La corbeille a été vidée.",

    # ---- CookingModeWindow (Mode cuisine plein écran) ----
    "cookingmode_title": "Mode cuisine — {name}",
    "cookingmode_close_button": "✕ Fermer (Échap)",
    "cookingmode_cooked_button": "🍳 J'ai cuisiné ça !",
    "cookingmode_fullscreen_hint": "F11 : plein écran",
    "cookingmode_persons_suffix": "{persons} pers.",
    "cookingmode_speech_button": "🔊 Lire à voix haute",
    "cookingmode_speech_stop_button": "⏹ Arrêter la lecture",
    "cookingmode_volume_percent": "{percent} %",
    "cookingmode_tts_module_missing": "La lecture à voix haute nécessite le module 'pyttsx3'.\nInstallez-le avec : pip install pyttsx3",
    "cookingmode_no_description_to_read": "Cette recette n'a pas de description à lire (le champ description est vide).",
    "cookingmode_ingredients_heading": "Ingrédients",
    "cookingmode_prep_label": "Préparation : {time} min",
    "cookingmode_cook_label": "Cuisson : {time} min",
    "cookingmode_difficulty_label": "Difficulté : {value}",
    "cookingmode_preparation_heading": "Préparation",
    "cookingmode_personal_notes_heading": "Notes personnelles",

    # ---- IngredientSearchWindow (Recherche par ingrédient) ----
    "ingsearch_title": "Recherche par ingrédient",
    "ingsearch_question_label": "Quel ingrédient recherchez-vous ?",
    "ingsearch_view_recipes_button": "🔍 Voir les recettes qui l'utilisent",
    "ingsearch_view_selected_button": "📖 Consulter la recette sélectionnée",
    "ingsearch_no_recipe_uses": "Aucune recette n'utilise « {name} » pour le moment.",
    "ingsearch_recipes_using": "Recettes utilisant « {name} » ({count}) :",
    "ingsearch_result_line": "{star}[{cat}] {name} ({qty}{unit} pour 1 personne)",
    "ingsearch_select_result_first": "Sélectionnez une recette dans la liste des résultats.",

    # ---- TimerRow (une ligne de minuteur dans TimersWindow) ----
    "timerrow_minutes_label": "Min :",
    "timerrow_seconds_label": "Sec :",
    "timerrow_error_invalid_duration": "Durée invalide.",
    "timerrow_set_duration_first": "Réglez une durée avant de démarrer.",

    # ---- CookLogEntryDialog (Ajouter au journal de cuisine) ----
    "cooklogentry_title": "📔 Ajouter au journal de cuisine",
    "cooklogentry_heading": "🍳 « {name} »",
    "cooklogentry_intro": "Comment était-ce ? Une note et/ou une photo\n(facultatif, vous pouvez aussi passer directement).",
    "cooklogentry_no_photo_chosen": "Aucune photo choisie",
    "cooklogentry_choose_photo_button": "📷 Choisir une photo",
    "cooklogentry_skip_button": "Passer",
    "cooklogentry_choose_photo_title": "Choisir une photo",

    # ---- CookLogWindow (Journal de cuisine) ----
    "cooklog_title": "📔 Journal de cuisine — {name}",
    "cooklog_heading": "📔 {name}",
    "cooklog_times_cooked": "Cuisinée {count} fois au total",
    "cooklog_no_entry": "Aucune note enregistrée pour le moment.\nUtilisez « 🍳 J'ai cuisiné ça ! » pour en ajouter une.",
    "cooklog_no_note": "(pas de note)",

    # ---- TimersWindow (Minuteurs) ----
    "timers_title": "⏲️ Minuteurs",
    "timers_intro": (
        "Réglez chaque minuteur puis ▶️ pour le démarrer.\n"
        "À la fin, la ligne clignote en rouge avec un signal sonore."
    ),
    "timers_add_button": "➕ Ajouter un minuteur",

    # ---- QRCodeWindow (QR Code d'une recette) ----
    "qrcode_title": "QR Code — {name}",
    "qrcode_intro": (
        "Scannez avec l'appareil photo ou une application de\n"
        "lecture de QR code pour voir le nom et les ingrédients."
    ),
    "qrcode_save_button": "💾 Enregistrer en image (PNG)",
    "qrcode_truncated_warning": (
        "⚠️ La recette est longue : le QR code contient un\n"
        "résumé tronqué (nom + ingrédients uniquement)."
    ),
    "qrcode_encoded_ingredients_heading": "Ingrédients ({persons} pers.) :",
    "qrcode_save_dialog_title": "Enregistrer le QR code",
    "qrcode_save_failed": "L'enregistrement a échoué :\n{error}",
    "qrcode_saved_message": "QR code enregistré :\n{path}",

    # ---- UnitConverterWindow (Convertisseur d'unités) ----
    "unitconv_title": "Convertisseur d'unités",
    "unitconv_heading": "🔄 Convertisseur d'unités",
    "unitconv_intro": (
        "Conversion approximative basée sur la densité de l'eau pour\n"
        "les unités de volume (ml, cl, L, tasse, cuillères) : fiable pour\n"
        "les liquides, approximative pour des solides comme la farine\n"
        "ou le sucre, dont la densité réelle diffère un peu."
    ),
    "unitconv_quantity_label": "Quantité :",
    "unitconv_from_label": "De :",
    "unitconv_to_label": "Vers :",
    "unitconv_convert_button": "Convertir",
    "unitconv_error_invalid_quantity": "Quantité invalide.",
    "unitconv_result": "{quantity} {from_unit} ≈ {result} {to_unit}",
    "unitconv_gram": "Gramme (g)",
    "unitconv_kilogram": "Kilogramme (kg)",
    "unitconv_ounce": "Once (oz)",
    "unitconv_pound": "Livre (lb)",
    "unitconv_milliliter": "Millilitre (ml)",
    "unitconv_centiliter": "Centilitre (cl)",
    "unitconv_liter": "Litre (L)",
    "unitconv_teaspoon": "Cuillère à café (5 ml)",
    "unitconv_tablespoon": "Cuillère à soupe (15 ml)",
    "unitconv_cup": "Tasse US (240 ml)",

    # ---- DisclaimerWindow (Clause de responsabilité) ----
    "disclaimer_title": "Clause de responsabilité",
    "disclaimer_heading": "⚠ Clause de responsabilité",
    "disclaimer_intro": "Merci de lire ce texte avant d'utiliser l'application.",
    "disclaimer_checkbox": "J'ai lu et j'accepte les conditions ci-dessus",
    "disclaimer_continue_button": "Continuer",
    "disclaimer_quit_button": "Quitter l'application",
    "disclaimer_text": (
        "ARTICLE 1 – EXCLUSION ET LIMITATION DE RESPONSABILITÉ\n\n"
        "1.1. Alertes médicales et gestion des allergènes\n\n"
        "L'Application propose une fonctionnalité permettant à l'Utilisateur de renseigner, "
        "modifier et configurer ses propres critères d'allergies et d'allergènes. "
        "L'Utilisateur reconnaît expressément que :\n\n"
        "• L'exactitude et la mise à jour de ces informations relèvent de sa seule et unique responsabilité.\n"
        "• L'Application est un outil informatique d'aide à la consultation de recettes et ne remplace en "
        "aucun cas un avis médical, un diagnostic ou le contrôle humain des ingrédients.\n"
        "• L'Éditeur ne saurait être tenu pour responsable en cas de mauvaise saisie, d'omission, de "
        "configuration erronée par l'Utilisateur, ou de réaction allergique (intolérance, choc anaphylactique, "
        "etc.) survenue après la consommation d'un plat. Il incombe à l'Utilisateur de vérifier "
        "systématiquement les étiquettes et la composition réelle de chaque ingrédient physique avant toute "
        "préparation ou ingestion.\n\n"
        "1.2. Fourniture « en l'état » et gratuité\n\n"
        "L'Application est mise à disposition de l'Utilisateur à titre entièrement gratuit. Elle est fournie "
        "« en l'état » et « selon sa disponibilité », sans aucune garantie d'absence d'erreurs, de bugs "
        "informatiques ou d'interruptions. L'Éditeur ne garantit pas que les fonctionnalités de l'Application "
        "répondront aux besoins spécifiques de l'Utilisateur.\n\n"
        "1.3. Dommages matériels et immatériels\n\n"
        "L'Éditeur décline toute responsabilité pour les dommages directs ou indirects causés à l'Utilisateur "
        "ou à des tiers. Plus particulièrement, l'Éditeur ne pourra être poursuivi pour :\n\n"
        "• Une panne, une surchauffe, un dysfonctionnement ou une détérioration du matériel informatique ou "
        "du smartphone de l'Utilisateur lors de l'utilisation de l'Application.\n"
        "• Une perte de données informatiques, une altération de fichiers ou un piratage du système de "
        "l'Utilisateur.\n\n"
        "En raison de la gratuité du service, si la responsabilité de l'Éditeur devait être engagée par un "
        "tribunal, le montant des dommages et intérêts serait expressément plafonné à la somme de zéro "
        "euro (0 €)."
    ),

    # ---- AddManualIngredientDialog (Ajouter des ingrédients à la liste de courses) ----
    "addmanual_title": "Ajouter des ingrédients à la liste de courses",
    "addmanual_heading": "➕ Ajouter des ingrédients à la liste de courses",
    "addmanual_intro": (
        "Ajoutez autant d'ingrédients que vous voulez à la liste\n"
        "d'attente ci-dessous, puis validez-les tous d'un coup."
    ),
    "addmanual_new_ingredient_button": "🥕 Nouvel ingrédient",
    "addmanual_add_to_list_button": "➕ Ajouter à la liste",
    "addmanual_staged_label": "Ingrédients en attente de validation :",
    "addmanual_remove_staged_button": "🗑 Retirer de la liste d'attente",
    "addmanual_confirm_all_button": "✅ Valider tous ces ingrédients",
    "addmanual_close_button": "Fermer",
    "addmanual_select_staged_first": "Sélectionnez un ingrédient dans la liste d'attente.",
    "addmanual_add_staged_first": "Ajoutez au moins un ingrédient à la liste d'attente avant de valider.",
    "addmanual_confirmed_message": "{count} ingrédient(s) ajouté(s) à la liste de courses.",

    # ---- Fonctions d'export de liste de courses (texte/Excel/PDF) ----
    "shoppingexport_generated_on": "Générée le {date}",
    "shoppingexport_selected_recipes": "Recettes sélectionnées :",
    "shoppingexport_excel_sheet_recipes": "Recettes",
    "shoppingexport_excel_col_recipe": "Recette",
    "shoppingexport_excel_col_persons": "Nombre de personnes",
    "shoppingexport_excel_sheet_ingredients": "Ingrédients",
    "shoppingexport_excel_col_rayon": "Rayon",
    "shoppingexport_excel_col_ingredient": "Ingrédient",
    "shoppingexport_excel_col_total_qty": "Quantité totale",
    "shoppingexport_excel_col_unit": "Unité",

    # ---- SavedShoppingListsWindow (Listes de courses enregistrées) ----
    "savedlists_title": "Listes de courses enregistrées",
    "savedlists_heading": "📂 Listes de courses enregistrées",
    "savedlists_load_button": "📂 Charger",
    "savedlists_delete_button": "🗑 Supprimer",
    "savedlists_none_saved": "Aucune liste enregistrée pour le moment.",
    "savedlists_entry_line": "{name} — {count} ingrédient(s) — {date}",
    "savedlists_select_list_first": "Sélectionnez une liste dans la liste.",
    "savedlists_delete_confirm": "Supprimer définitivement la liste « {name} » ?",

    # ---- QuickSearchWindow (Recherche rapide, Ctrl+K) ----
    "quicksearch_title": "Recherche rapide",
    "quicksearch_heading": "🔍 Recherche rapide de recette",
    "quicksearch_no_results": "Aucune recette trouvée.",
    "quicksearch_footer_hint": "Entrée pour ouvrir, Échap pour fermer.",

    # ---- AllRecipesWindow (Voir toutes les recettes / liste de courses) ----
    "allrecipes_title": "Toutes les recettes - Liste de courses",
    "allrecipes_select_label": "Sélectionnez les recettes et le nombre de personnes :",
    "allrecipes_ingredient_filter_title": "Filtrer par ingrédient",
    "allrecipes_persons_count_label": "Nb. personnes :",
    "allrecipes_add_to_cart_button": "🛒 Ajouter aux courses",
    "allrecipes_checklist_mode_button": "☑️ Mode courses (cocher au fur et à mesure)",
    "allrecipes_clear_list_button": "🗑 Vider la liste de courses",
    "allrecipes_export_button": "📤 Exporter",
    "allrecipes_print_button": "🖨️ Imprimer",
    "allrecipes_add_manual_ingredient_button": "➕ Ajouter un ingrédient à la liste de courses",
    "allrecipes_save_list_button": "💾 Enregistrer cette liste pour plus tard",
    "allrecipes_load_list_button": "📂 Charger une liste enregistrée",
    "allrecipes_invalid_persons": "Nombre de personnes invalide pour « {name} ».",
    "allrecipes_empty_list_message": (
        "Votre liste de courses est vide pour le moment.\n"
        "Cliquez sur « 🛒 Ajouter aux courses » en face d'une recette,\n"
        "ajoutez un ingrédient manuellement, ou chargez une liste enregistrée."
    ),
    "allrecipes_total_list_heading": "=== Liste de courses totale ===",
    "allrecipes_manual_items_note": "({count} ingrédient(s) ajouté(s) manuellement inclus)",
    "allrecipes_invalid_quantity": "Quantité invalide.",
    "allrecipes_calculate_list_first": "Calculez d'abord une liste de courses avant de l'enregistrer.",
    "allrecipes_save_list_dialog_title": "Enregistrer la liste",
    "allrecipes_save_list_dialog_prompt": "Nom pour cette liste :",
    "allrecipes_list_saved_title": "Enregistré",
    "allrecipes_list_saved_message": "Liste « {name} » enregistrée pour plus tard.",
    "allrecipes_empty_list_for_export": (
        "La liste de courses est vide. Ajoutez au moins une recette "
        "(bouton « 🛒 Ajouter aux courses ») ou un ingrédient manuel."
    ),
    "allrecipes_export_txt_title": "Enregistrer la liste de courses en texte",
    "allrecipes_export_excel_title": "Enregistrer la liste de courses en Excel",
    "allrecipes_export_pdf_title": "Enregistrer la liste de courses en PDF",
    "allrecipes_export_saved_message": "Liste de courses enregistrée :\n{path}",
    "allrecipes_excel_module_missing": "L'export Excel nécessite le module 'openpyxl'.\nInstallez-le avec : pip install openpyxl",
    "allrecipes_pdf_module_missing": "L'export PDF nécessite le module 'reportlab'.\nInstallez-le avec : pip install reportlab",
    "allrecipes_print_module_missing": (
        "L'impression nécessite le module 'reportlab' pour générer la mise en page.\n"
        "Installez-le avec : pip install reportlab"
    ),
    "allrecipes_print_label": "la liste de courses",
    "allrecipes_shopping_list_title": "Liste de courses",
    "allrecipes_close_confirm_title": "Fermer la liste de courses ?",
    "allrecipes_close_confirm_message": (
        "La liste de courses affichée n'est pas enregistrée : elle sera "
        "définitivement perdue si vous fermez cette fenêtre maintenant.\n\n"
        "Astuce : utilisez « 💾 Enregistrer cette liste pour plus tard » "
        "avant de fermer si vous voulez la conserver.\n\n"
        "Fermer quand même ?"
    ),

    # ---- ManageRecipesWindow (Modifier / Supprimer une recette) ----
    "managerecipes_title": "Modifier / Supprimer une recette",
    "managerecipes_select_label": "Sélectionnez une recette :",
    "managerecipes_filter_favorites": "⭐ Favoris uniquement",
    "managerecipes_filter_quick": "⏱️ Recettes rapides (≤ 30 min) uniquement",
    "managerecipes_filter_vegetarian": "🥗 Recettes végétariennes uniquement",
    "managerecipes_filter_wishlist": "💭 Liste d'envies uniquement",
    "managerecipes_remove_filter_button": "✕ Retirer le filtre",
    "managerecipes_search_label": "🔍 Rechercher :",
    "managerecipes_sort_label": "Trier par :",
    "managerecipes_category_label": "Catégorie :",
    "managerecipes_edit_button": "✏️ Modifier",
    "managerecipes_duplicate_button": "📋 Dupliquer",
    "managerecipes_delete_button": "🗑️ Supprimer",
    "managerecipes_select_recipe_first": "Sélectionnez une recette dans la liste.",
    "managerecipes_duplicate_suffix": "(copie)",
    "managerecipes_duplicated_title": "Dupliquée",
    "managerecipes_duplicated_message": "« {original} » a été dupliquée sous le nom « {new} ».",
    "managerecipes_delete_confirm_message": (
        "Envoyer la recette « {name} » à la corbeille ?\n\n"
        "Vous pourrez la restaurer plus tard depuis le bouton « 🗑️ Corbeille »."
    ),
    "managerecipes_deleted_title": "Envoyée à la corbeille",
    "managerecipes_deleted_message": "La recette a été déplacée vers la corbeille.",

    # ---- OneRecipeWindow (Voir une recette précise) ----
    "onerecipe_window_title": "Voir une recette",
    "onerecipe_choose_recipe_label": "Choisissez une recette :",
    "onerecipe_search_label": "🔍 Rechercher :",
    "onerecipe_sort_label": "Trier :",
    "onerecipe_category_label": "Catégorie :",
    "onerecipe_persons_label": "Nombre de personnes :",
    "onerecipe_btn_show": "Afficher la recette",
    "onerecipe_btn_export_pdf": "📄 Exporter en PDF",
    "onerecipe_btn_print": "🖨️ Imprimer",
    "onerecipe_btn_add_to_shopping": "🛒 Ajouter à la liste de courses",
    "onerecipe_btn_cooked": "🍳 J'ai cuisiné ça !",
    "onerecipe_btn_cooking_mode": "🖥️ Mode cuisine (plein écran)",
    "onerecipe_btn_qr": "📱 QR Code",
    "onerecipe_btn_timers": "⏲️ Minuteurs",
    "onerecipe_btn_cook_log": "📔 Journal de cuisine",
    "onerecipe_btn_substitutions": "🔄 Substituts possibles",
    "onerecipe_edit_button": "✏️ Modifier",
    "onerecipe_ingredients_info_label": "Ingrédients et informations :",
    "onerecipe_description_notes_label": "Description et notes :",
    "onerecipe_similar_label": "Recettes similaires :",
    "onerecipe_no_photo": "(aucune photo)",
    "onerecipe_preview_unavailable": "(aperçu indisponible)",
    "onerecipe_select_recipe_first": "Sélectionnez une recette dans la liste.",
    "onerecipe_display_first": "Affichez d'abord une recette avec « Afficher la recette ».",
    "onerecipe_invalid_persons": "Nombre de personnes invalide.",
    "onerecipe_added_to_shopping_title": "Ajouté",
    "onerecipe_added_to_shopping_message": (
        "« {name} » ({persons} pers.) sera automatiquement ajoutée à la liste de courses "
        "la prochaine fois que vous ouvrirez « Voir toutes les recettes »."
    ),
    "onerecipe_pantry_decrement_title": "Garde-manger",
    "onerecipe_pantry_decrement_prompt": "Décompter les ingrédients de « {name} » ({persons} pers.) de votre garde-manger ?",
    "onerecipe_pantry_updated_title": "Garde-manger mis à jour",
    "onerecipe_pantry_updated_message": "{count} ingrédient(s) décompté(s) de votre garde-manger.",
    "onerecipe_pantry_none_decremented": (
        "Aucun ingrédient de cette recette n'a pu être décompté "
        "(absent du garde-manger, ou unité non comparable)."
    ),
    "onerecipe_marked_title": "Marqué",
    "onerecipe_marked_message": "« {name} » a été marquée comme cuisinée aujourd'hui !",
    "onerecipe_no_substitutes_title": "Aucun substitut connu",
    "onerecipe_no_substitutes_message": (
        "Aucun ingrédient de cette recette n'a de substitut connu pour le moment.\n\n"
        "Vous pouvez en ajouter vous-même depuis « 🥕 Gérer les ingrédients » > "
        "« 🔄 Gérer les substitutions »."
    ),
    "onerecipe_substitutes_title": "Substituts possibles — {name}",
    "onerecipe_substitutes_heading": "🔄 Substituts possibles pour « {name} »",
    "onerecipe_substitutes_disclaimer": (
        "Suggestions culinaires, pas des équivalences garanties :\nle résultat peut varier selon la recette."
    ),
    "onerecipe_close_button": "Fermer",
    "onerecipe_rating_label": "Note : {stars}",
    "onerecipe_prep_label": "Préparation : {time} min",
    "onerecipe_cook_label": "Cuisson : {time} min",
    "onerecipe_difficulty_label": "Difficulté : {value}",
    "onerecipe_allergens_label": "⚠ Allergènes : {list}",
    "onerecipe_cost_label": "💰 Coût estimé : {cost} €{partial}",
    "onerecipe_cost_partial": " (estimation partielle, {known}/{total} ingrédients avec prix connu)",
    "onerecipe_nutrition_partial": " (estimation partielle, {known}/{total} ingrédients reconnus)",
    "onerecipe_nutrition_label": (
        "🥗 Valeurs nutritionnelles estimées{partial} :\n"
        "   {kcal} kcal · {protein} g protéines · {carbs} g glucides · {fat} g lipides\n"
    ),
    "onerecipe_description_heading": "--- Description ---\n{text}\n",
    "onerecipe_notes_heading": "\n--- Notes personnelles ---\n{text}\n",
    "onerecipe_no_description_notes": "(Aucune description ni note personnelle pour cette recette.)",
    "onerecipe_export_pdf_title": "Exporter la recette en PDF",
    "onerecipe_export_success_title": "Export réussi",
    "onerecipe_export_success_message": "Recette exportée :\n{path}",
    "onerecipe_export_failed": "L'export a échoué :\n{error}",
    "onerecipe_print_failed": "La préparation de l'impression a échoué :\n{error}",
    "onerecipe_pdf_module_missing": "L'export PDF nécessite le module 'reportlab'.\nInstallez-le avec : pip install reportlab",
    "onerecipe_print_module_missing": (
        "L'impression nécessite le module 'reportlab' pour générer la mise en page.\n"
        "Installez-le avec : pip install reportlab"
    ),
    "onerecipe_qr_module_missing": "L'export en QR code nécessite le module 'qrcode'.\nInstallez-le avec : pip install qrcode",
    "onerecipe_qr_pillow_missing": (
        "L'export en QR code nécessite aussi le module 'Pillow'.\nInstallez-le avec : pip install pillow"
    ),
    "onerecipe_default_timer_label": "Minuteur",

    # ---- RecipeFormWindow (Ajouter / Modifier une recette) ----
    "recipeform_title_edit": "Modifier la recette",
    "recipeform_title_add": "Ajouter une recette",
    "recipeform_name_label": "Nom de la recette :",
    "recipeform_favorite_checkbox": "⭐ Marquer comme recette favorite",
    "recipeform_wishlist_checkbox": "💭 Ajouter à ma liste d'envies (à essayer)",
    "recipeform_rating_label": "Ma note :",
    "recipeform_category_label": "Catégorie :",
    "recipeform_prep_time_label": "Préparation (min) :",
    "recipeform_cook_time_label": "Cuisson (min) :",
    "recipeform_difficulty_label": "Difficulté :",
    "recipeform_default_persons_label": "   Personnes par défaut :",
    "recipeform_tags_label": "Étiquettes (séparées par des virgules) :",
    "recipeform_tags_example": "ex. végétarien, sans gluten, rapide, économique",
    "recipeform_allergens_label": "Allergènes présents :",
    "recipeform_detect_allergens_button": "🔍 Détecter automatiquement",
    "recipeform_allergens_disclaimer": (
        "Ceci n'est qu'à titre informatif, vérifiez toujours les\n"
        "allergènes sur les étiquettes des produits physiques."
    ),
    "recipeform_allergens_auto_note": (
        "La détection automatique se base sur les ingrédients de la\n"
        "recette déjà saisis ci-dessous : elle coche et décoche les\n"
        "cases en fonction, sans jamais toucher à celles que vous\n"
        "auriez cochées vous-même sans lien avec un ingrédient détecté."
    ),
    "recipeform_photos_label": "Photos :",
    "recipeform_add_photo_button": "📷 Ajouter une photo",
    "recipeform_description_label": "Description (informations, étapes, astuces...) :",
    "recipeform_notes_label": "Notes personnelles (avis, ajustements pour la prochaine fois...) :",
    "recipeform_ingredients_label": "Ingrédients (quantité pour 1 personne) :",
    "recipeform_new_ingredient_button": "🥕 Nouvel ingrédient",
    "recipeform_no_ingredients_registered": (
        "Aucun ingrédient enregistré. Cliquez sur « 🥕 Nouvel ingrédient »\npour en créer un premier."
    ),
    "recipeform_header_ingredient": "Ingrédient",
    "recipeform_header_quantity": "Quantité",
    "recipeform_header_unit": "Unité",
    "recipeform_header_other": "(si autre)",
    "recipeform_add_ingredient_button": "+ Ajouter un ingrédient",
    "recipeform_save_button": "Enregistrer",
    "recipeform_delete_button": "Supprimer cette recette",
    "recipeform_char_counter": "{count} / {max} caractères",
    "recipeform_add_ingredients_first": "Ajoutez d'abord des ingrédients à la recette.",
    "recipeform_allergens_updated_title": "Allergènes mis à jour",
    "recipeform_allergens_updated_added": "ajouté(s) : {list}",
    "recipeform_allergens_updated_removed": "retiré(s) : {list}",
    "recipeform_allergens_updated_message": "Allergène(s) {parts}.",
    "recipeform_allergens_no_change": "Aucun changement : les allergènes cochés correspondent déjà aux ingrédients.",
    "recipeform_choose_photos_title": "Choisir une ou plusieurs photos",
    "recipeform_no_photo": "(aucune photo)",
    "recipeform_preview_unavailable": "(aperçu\nindisponible)",
    "recipeform_remove_photo_button": "🗑 Retirer",
    "recipeform_new_ingredient_dialog_title": "Nouvel ingrédient",
    "recipeform_new_ingredient_dialog_prompt": "Nom du nouvel ingrédient :",
    "recipeform_ingredient_already_exists": "L'ingrédient « {name} » existe déjà.",
    "recipeform_ingredient_added_title": "Ajouté",
    "recipeform_ingredient_added_message": (
        "L'ingrédient « {name} » a été ajouté.\nSélectionnez-le dans une des listes déroulantes."
    ),
    "recipeform_error_name_required": "Merci d'indiquer un nom de recette.",
    "recipeform_error_prep_time": "Le temps de préparation doit être un nombre positif (ou vide).",
    "recipeform_error_cook_time": "Le temps de cuisson doit être un nombre positif (ou vide).",
    "recipeform_unknown_ingredient_title": "Ingrédient inconnu",
    "recipeform_unknown_ingredient_message": (
        "« {name} » ne correspond à aucun ingrédient enregistré.\n"
        "Choisissez-en un dans la liste déroulante, ou cliquez sur "
        "« 🥕 Nouvel ingrédient » pour l'ajouter d'abord."
    ),
    "recipeform_error_invalid_quantity": "Quantité invalide pour '{name}'.",
    "recipeform_error_custom_unit_required": "Précisez l'unité personnalisée pour '{name}'.",
    "recipeform_error_no_valid_ingredient": "Ajoutez au moins un ingrédient valide.",
    "recipeform_duplicate_ingredient_title": "Ingrédient en double",
    "recipeform_duplicate_ingredient_message": "« {list} » apparaît plusieurs fois dans cette recette.\n\nEnregistrer quand même ?",
    "recipeform_saved_message": "La recette « {name} » a été enregistrée.",
    "recipeform_delete_confirm_message": (
        "Envoyer la recette « {name} » à la corbeille ?\n\n"
        "Vous pourrez la restaurer plus tard depuis le bouton « 🗑️ Corbeille »."
    ),
    "recipeform_deleted_title": "Envoyée à la corbeille",
    "recipeform_deleted_message": "La recette a été déplacée vers la corbeille.",
}

TRANSLATIONS = {
    "en": {
        "home_window_title": "My Recipe Book",
        "home_banner_title": "👨‍🍳 My Recipe Book",
        "home_banner_subtitle": "All your recipes, right at hand",
        "home_donate_button": "☕ Donate",
        "home_dark_theme": "🌙 Dark theme",
        "home_light_theme": "☀️ Light theme",
        "home_large_text_on": "🔎 Larger text",
        "home_large_text_off": "🔎 Normal text",
        "home_daily_recipe_title": "🎲 Recipe of the day",
        "home_open_button": "👁 Open",
        "home_quick_filter_favorites": "⭐ Favorites",
        "home_quick_filter_quick": "⏱️ Quick (≤ 30 min)",
        "home_quick_filter_vegetarian": "🥗 Vegetarian",
        "home_quick_filter_wishlist": "💭 Wish list",
        "home_wishlist_reminder": (
            "💭 {count} recipe(s) on your wish list for over {days} days — "
            "how about trying them? (click to see them)"
        ),
        "home_low_stock_reminder": (
            "📦 {count} ingredient(s) running low in your pantry: "
            "{names} — click to add them to the shopping list"
        ),
        "home_btn_add_recipe": "➕  Add a recipe",
        "home_btn_import_url": "🌐  Import a recipe from a link",
        "home_btn_import_photo": "📷  Import a recipe from a photo",
        "home_btn_view_all_recipes": "🧾  View all recipes (shopping list)",
        "home_btn_view_one_recipe": "🍽️  View a specific recipe",
        "home_btn_manage_recipes": "✏️  Edit / Delete a recipe",
        "home_btn_compare_recipes": "⚖️  Compare two recipes",
        "home_btn_manage_ingredients": "🥕  Manage ingredients",
        "home_btn_ingredient_search": "🔎  Search by ingredient",
        "home_btn_what_can_i_cook": "🧊  What can I cook?",
        "home_btn_pantry": "📦  My pantry",
        "home_btn_unit_converter": "🔄  Unit converter",
        "home_btn_weekly_plan": "📅  Weekly meal plan",
        "home_btn_menus": "📋  My menus",
        "home_btn_statistics": "📊  Statistics",
        "home_btn_export_cookbook": "📖  Export the cookbook",
        "home_btn_import_export": "💾  Import / Export data",
        "home_btn_trash": "🗑️  Trash",
        "home_today_title": "📅 Today",
        "home_recent_title": "🕘 Recently viewed",
        "home_wishlist_title": "💭 Recipes to try",
        "home_new_draw_button": "🎲 New picks",
        "home_footer_recipe_count": "{count} recipe(s) saved",
        "home_nothing_planned": (
            "Nothing planned for {day}. Fill in the « 📅 Weekly meal plan » to see it here."
        ),
        "home_no_recent_recipe": "No recipe viewed yet.",
        "home_no_wishlist_recipe": "No recipe on your wish list yet.",
        "warning_pillow": "Pillow not installed: photos won't display (pip install pillow)",
        "warning_reportlab": "reportlab not installed: PDF export unavailable (pip install reportlab)",
        "warning_openpyxl": "openpyxl not installed: Excel export unavailable (pip install openpyxl)",
        "warning_qrcode": "qrcode not installed: QR code export unavailable (pip install qrcode)",
        "warning_pytesseract": (
            "pytesseract not installed: import from photo unavailable "
            "(pip install pytesseract, + Tesseract OCR)"
        ),

        # ---- Common dialog titles, reused throughout the application ----
        "common_error": "Error",
        "common_info": "Info",
        "common_confirm": "Confirm",
        "common_success": "Success",
        "common_module_missing": "Missing module",
        "common_all_categories": "All",
        "common_export_failed": "Export failed:\n{error}",
        "common_export_success_title": "Export successful",
        "common_print_failed": "Print preparation failed:\n{error}",
        "common_reset_button": "Reset",
        "common_want_label": "I want:",
        "common_exclude_label": "I don't want:",
        "common_tags_filter_label": "Tags (all required):",
        "common_filter_hint": "Type the first letters to filter the list.",
        "common_search_label": "🔍 Search:",
        "common_sort_by_label": "Sort by:",
        "common_category_label": "Category:",
        "common_edit_button": "✏️ Edit",
        "common_unknown_ingredient_title": "Unknown ingredient",
        "common_unknown_ingredient_simple_message": (
            "« {name} » doesn't match any registered ingredient.\nChoose one from the dropdown list."
        ),
        "common_ingredient_label": "Ingredient:",
        "common_quantity_label": "Quantity:",
        "common_unit_label": "Unit:",
        "common_new_ingredient_button": "🥕 New ingredient",
        "common_save_button": "💾 Save",

        # ---- PantryWindow (My pantry) ----
        "pantry_title": "My pantry",
        "pantry_heading": "📦 My pantry",
        "pantry_intro": (
            "Enter what you have at home and how much of it.\n"
            "« What can I cook? » can then check if you have enough,\n"
            "and offer to automatically deduct stock after you cook."
        ),
        "pantry_threshold_label": "Alert threshold (optional):",
        "pantry_help_text": (
            "To ADD an item: enter the ingredient (create it first with\n"
            "« 🥕 New ingredient » if it's not in your list yet), the\n"
            "quantity and unit, then click « 💾 Save ».\n"
            "To EDIT an existing item: click on it once in the\n"
            "list below — this loads its values into the fields above,\n"
            "without saving anything: change the values you want, THEN click\n"
            "« 💾 Save » for the change to take effect.\n"
            "The alert threshold triggers a reminder on the home page as soon as the\n"
            "quantity drops below it (leave empty to never be alerted)."
        ),
        "pantry_remove_button": "🗑 Remove from pantry",
        "pantry_empty": "Your pantry is empty for now.",
        "pantry_threshold_suffix": " (threshold: {threshold})",
        "pantry_error_ingredient_required": "Please enter an ingredient.",
        "pantry_error_invalid_quantity": "Invalid quantity.",
        "pantry_error_invalid_threshold": "Invalid alert threshold (leave empty if you don't want one).",
        "pantry_select_ingredient_first": "Select an ingredient from the list.",
        "pantry_remove_confirm_message": "Remove « {name} » from the pantry?",

        # ---- WhatCanICookWindow (What can I cook?) ----
        "cook_title": "What can I cook?",
        "cook_instructions_label": "Enter the ingredients you have at home:",
        "cook_staples_hint": (
            "A few common staple ingredients are already checked here\n"
            "(salt, oil, flour...) — remove any you don't have."
        ),
        "cook_all_ingredients_label": "All ingredients:",
        "cook_add_button": "➕ Add →",
        "cook_have_label": "What I have:",
        "cook_remove_button": "🗑 Remove",
        "cook_load_from_pantry_button": "📦 Load from my pantry",
        "cook_compute_button": "🔍 See feasible recipes",
        "cook_open_selected_button": "📖 View selected recipe",
        "cook_pantry_empty_title": "Info",
        "cook_pantry_empty_message": (
            "Your pantry is empty for now. Open « 📦 My "
            "pantry » from the home page to add ingredients."
        ),
        "cook_loaded_title": "Loaded",
        "cook_loaded_message": "{count} ingredient(s) added from your pantry.",
        "cook_add_ingredient_first": "Add at least one ingredient you have.",
        "cook_feasible_header": "✅ Feasible with what you have:",
        "cook_insufficient_quantity": "  ⚠️ insufficient quantity: {list}",
        "cook_none_feasible": "No recipe is 100% feasible with these ingredients.",
        "cook_substitutable_header": "🔄 Feasible using a substitute:",
        "cook_almost_header": "🟡 Almost (missing 1 to 3 ingredients):",
        "cook_missing_label": "   {name} (missing: {list})",
        "cook_no_results": "Try adding more ingredients to your selection.",
        "cook_select_recipe_from_results": "Select a recipe from the results list.",
        "cook_select_recipe_row": "Select a row corresponding to a recipe.",

        # ---- WeeklyPlanHistoryWindow (Past weeks history) ----
        "weekhistory_title": "Past weeks history",
        "weekhistory_heading": "🕘 Past weeks history",
        "weekhistory_intro": (
            "Every week you save the meal plan, it is archived here\n"
            "automatically (up to 26 weeks, about 6 months), to avoid\n"
            "repeating the same thing too soon."
        ),
        "weekhistory_reload_button": "♻️ Reload into current plan",
        "weekhistory_delete_button": "🗑 Delete this week",
        "weekhistory_no_archived_weeks": "No archived weeks.",
        "weekhistory_week_label": "Week {week}",
        "weekhistory_saved_on": "Saved on {date}\n\n",
        "weekhistory_day_heading": "{day}:\n",
        "weekhistory_slot_line": "   {slot}: {recipe} ({persons} servings)\n",
        "weekhistory_empty_week": "(Empty plan for this week.)",
        "weekhistory_select_week_first": "Select a week from the list.",
        "weekhistory_reload_confirm_message": (
            "Reload the plan for week {week} into the "
            "current plan?\n\nThis will replace the recipes currently shown "
            "(consider saving the current plan first if you want to keep it)."
        ),
        "weekhistory_delete_confirm_message": "Permanently delete the archive for week {week}?",

        # ---- WeeklyPlanTemplatesWindow (Weekly plan templates) ----
        "weektemplates_title": "Weekly plan templates",
        "weektemplates_heading": "📋 Weekly plan templates",
        "weektemplates_intro": (
            "Save the currently displayed plan as a reusable\n"
            "template, to apply it to another week in one click\n"
            "instead of re-entering everything."
        ),
        "weektemplates_name_label": "Name for the new template:",
        "weektemplates_save_button": "💾 Save current plan as template",
        "weektemplates_apply_button": "📋 Apply this template",
        "weektemplates_delete_button": "🗑 Delete this template",
        "weektemplates_none_saved": "No template saved yet.",
        "weektemplates_error_name_required": "Please enter a name for this template.",
        "weektemplates_empty_plan": "The currently displayed plan is empty: nothing to save as a template.",
        "weektemplates_saved_message": "Template « {name} » saved.",
        "weektemplates_select_template_first": "Select a template from the list.",
        "weektemplates_apply_confirm_message": (
            "Apply the template « {name} » to the current plan?\n\n"
            "This will replace the recipes currently shown (consider "
            "saving the current plan first if you want to keep it)."
        ),
        "weektemplates_delete_confirm_message": "Permanently delete the template « {name} »?",
        "common_none_option": "-- None --",

        # ---- WeeklyPlanWindow (Weekly meal plan) ----
        "weekplan_title": "Weekly meal plan",
        "weekplan_subtitle": "Calendar view: days in columns, meals in rows.",
        "weekplan_save_button": "💾 Save the plan",
        "weekplan_clear_button": "🗑 Clear all",
        "weekplan_export_ics_button": "📆 Export to a calendar (.ics)",
        "weekplan_compute_button": "Calculate the week's shopping list",
        "weekplan_checklist_button": "☑️ Shopping mode",
        "weekplan_empty_list_message": (
            "No list calculated yet.\n"
            "Click « Calculate the week's shopping list » above,\n"
            "or load a saved list."
        ),
        "weekplan_total_list_heading": "=== Weekly shopping list ===",
        "weekplan_calculate_list_for_export": (
            "First calculate a shopping list (« Calculate the week's shopping list » button)."
        ),
        "weekplan_invalid_persons_for_slot": "Invalid number of servings for {day} — {slot}.",
        "weekplan_saved_message": "The weekly meal plan has been saved.",
        "weekplan_clear_confirm_message": "Clear the entire weekly meal plan?",
        "weekplan_assign_recipe_first": "Assign at least one recipe to a slot in the week.",
        "weekplan_export_ics_title": "Export the plan to a calendar",
        "weekplan_ics_export_success_message": (
            "Plan exported:\n{path}\n\n"
            "Import this file into Google Calendar, Outlook, or Calendar "
            "to see your meals repeat there every week."
        ),
        "weekplan_assign_or_manual": (
            "Assign at least one recipe to a slot in the week, "
            "or add an ingredient manually."
        ),
        "weekplan_export_shopping_list_title": "Save the shopping list",
        "weekplan_shopping_list_title": "Weekly shopping list",
        "weekplan_list_saved_message": "List saved:\n{path}",
        "weekplan_excel_module_missing": "Excel export requires: pip install openpyxl",
        "weekplan_pdf_module_missing": "PDF export requires: pip install reportlab",
        "weekplan_print_module_missing": "Printing requires: pip install reportlab",
        "weekplan_print_label": "the weekly shopping list",

        # ---- ManageIngredientsWindow (Manage ingredients) ----
        "manageing_title": "Manage ingredients",
        "manageing_list_label": "List of registered ingredients:",
        "manageing_add_button": "➕ Add",
        "manageing_edit_button": "✏️ Edit",
        "manageing_delete_button": "🗑️ Delete",
        "manageing_load_defaults_button": "📚 Load the ~1000 common ingredients",
        "manageing_spell_check_button": "🔤 Check for duplicates / typos",
        "manageing_prices_button": "💰 Manage prices (for recipe cost)",
        "manageing_substitutions_button": "🔄 Manage substitutions",
        "manageing_edit_hint": (
            "\"Edit\" lets you change the name (updated\n"
            "everywhere the ingredient is used), its allergens,\n"
            "its nutritional values, and its price."
        ),
        "manageing_select_ingredient_first": "Select an ingredient from the list.",
        "manageing_delete_confirm_message": "Delete « {name} » from the ingredient list?",
        "manageing_delete_usage_warning": (
            "\n\nWarning: it is used in {count} recipe(s). "
            "These recipes will keep this ingredient, but it will no longer "
            "be suggested in the dropdown menu, unless you add it back."
        ),
        "manageing_missing_file_title": "Missing file",
        "manageing_missing_file_message": (
            "The file ingredients_par_defaut.json could not be found.\n"
            "Make sure it is in the same folder as main.py."
        ),
        "manageing_done_title": "Done",
        "manageing_defaults_added_message": "{count} new ingredient(s) added from the common list.",
        "manageing_defaults_none_added": "All the common ingredients were already present.",

        # ---- SubstitutionEditWindow (Substitutes for a specific ingredient) ----
        "subedit_title": "Substitutes for « {name} »",
        "subedit_heading": "🔄 Substitutes for « {name} »",
        "subedit_disclaimer": (
            "A substitution is a cooking suggestion, not a guaranteed\n"
            "equivalence: the result may vary depending on the recipe."
        ),
        "subedit_remove_button": "🗑 Remove selected substitute",
        "subedit_add_frame_title": "Add a substitute",
        "subedit_name_label": "Name:",
        "subedit_note_label": "Note (optional):",
        "subedit_add_to_list_button": "➕ Add to list",
        "subedit_revert_button": "🔄 Revert to the built-in list",
        "subedit_cancel_button": "Cancel",
        "subedit_no_substitute_yet": "No substitute yet.",
        "subedit_error_name_required": "Please enter a substitute name.",
        "subedit_select_to_remove": "Select a substitute to remove.",
        "subedit_revert_confirm_message": (
            "Remove your custom list and revert to the substitutes provided\n"
            "with the application for « {name} »?"
        ),

        # ---- ManageSubstitutionsWindow (Manage substitutions) ----
        "managesub_title": "Manage substitutions",
        "managesub_heading": "🔄 Ingredient substitutions",
        "managesub_intro": (
            "View or edit the suggested substitutes for an ingredient.\n"
            "A substitution is a cooking suggestion, not a guaranteed equivalence."
        ),
        "managesub_manage_button": "✏️ Manage its substitutes",
        "managesub_hint": (
            "Double-click an ingredient in the list to view or edit its\n"
            "substitutes, or type a name above (including an ingredient with\n"
            "no known substitute yet) then « ✏️ Manage its substitutes »."
        ),
        "managesub_none_with_substitute": "No ingredient with a substitute yet.",
        "managesub_substitute_count": "{name} ({count} substitute{plural})",
        "managesub_error_ingredient_required": "Please enter an ingredient.",
        "managesub_unknown_ingredient_message": (
            "« {name} » doesn't match any registered ingredient.\n"
            "Choose one from the dropdown list, or create it first "
            "from « 🥕 Manage ingredients »."
        ),

        # ---- IngredientPricesWindow (Manage ingredient prices) ----
        "ingprices_title": "Manage ingredient prices",
        "ingprices_heading": "💰 Ingredient prices",
        "ingprices_intro": (
            "Enter a price for the ingredients you care\n"
            "about — no need to do them all. A recipe's\n"
            "cost is estimated from these prices."
        ),
        "ingprices_price_label": "Price (€):",
        "ingprices_for_one_label": "per 1",
        "ingprices_save_button": "💾 Save the price",
        "ingprices_clear_button": "🗑 Clear the price",
        "ingprices_units_note": (
            "kg ↔ recipes in g   ·   L ↔ recipes in cl   ·   prices\n"
            "per piece/spoon apply as-is."
        ),
        "ingprices_no_price_set": "  —  (no price set)",
        "ingprices_price_suffix": "  —  {price} € / {unit}",
        "ingprices_error_invalid_price": "Enter a valid price (positive number).",
        "ingprices_saved_message": "Price saved for « {name} ».",

        # ---- IngredientEditWindow (New ingredient / Edit an ingredient) ----
        "ingedit_title_edit": "Edit an ingredient",
        "ingedit_title_new": "New ingredient",
        "ingedit_heading_edit": "✏️ Edit the ingredient",
        "ingedit_heading_new": "➕ New ingredient",
        "ingedit_name_label": "Name:",
        "ingedit_allergens_label": "Allergens present:",
        "ingedit_nutrition_label": "Nutritional values (per 100 g / 100 ml):",
        "ingedit_nutri_kcal": "Calories (kcal)",
        "ingedit_nutri_protein": "Protein (g)",
        "ingedit_nutri_carbs": "Carbs (g)",
        "ingedit_nutri_fat": "Fat (g)",
        "ingedit_nutrition_hint": "Leave blank if you don't know these values.",
        "ingedit_price_label": "Price:",
        "ingedit_save_button": "💾 Save",
        "ingedit_delete_button": "🗑️ Delete this ingredient",
        "ingedit_error_invalid_field": "« {field} » must be a positive number (or empty).",
        "ingedit_error_name_required": "Please enter an ingredient name.",
        "ingedit_error_already_exists": "The ingredient « {name} » already exists.",
        "ingedit_error_plural_duplicate": (
            "« {name} » is just a singular/plural variant of the "
            "already existing ingredient « {existing} ». To avoid "
            "duplicates in the list, use « {existing} » directly."
        ),
        "ingedit_nutri_field_kcal": "Calories",
        "ingedit_nutri_field_protein": "Protein",
        "ingedit_nutri_field_carbs": "Carbs",
        "ingedit_nutri_field_fat": "Fat",
        "ingedit_error_invalid_price": "The price must be a positive number (or empty).",
        "ingedit_saved_message": "« {name} » has been saved.",

        # ---- IngredientSpellCheckWindow (Spelling check) ----
        "spellcheck_title": "Ingredient spelling check",
        "spellcheck_heading": (
            "Ingredient pairs that are 90% similar or more\n"
            "(probable duplicates or typos):"
        ),
        "spellcheck_multi_select_hint": (
            "Multiple selection possible (Ctrl+click or Shift+click) to\n"
            "merge several pairs at once."
        ),
        "spellcheck_merge_button": "🔗 Merge selection",
        "spellcheck_not_duplicate_button": "✕ Not a duplicate",
        "spellcheck_rerun_button": "🔄 Re-run the scan",
        "spellcheck_footer_hint": (
            "For a single pair, you'll be asked which of the two\n"
            "spellings to keep. For several pairs at once, the\n"
            "less-used ingredient in your recipes is automatically merged\n"
            "into the one used in the most recipes.\n"
            "« Not a duplicate » permanently removes the selected\n"
            "pair(s) from this scan, now and in the future."
        ),
        "spellcheck_none_found": "No probable duplicate detected. 🎉",
        "spellcheck_pair_line": "{a}   ↔   {b}     ({percent}% similar)",
        "spellcheck_select_pair_first": "Select at least one pair from the list.",
        "spellcheck_dismissed_message": (
            "{count} pair(s) marked as not being "
            "duplicates. They won't be suggested again in future scans."
        ),
        "spellcheck_merge_dialog_title": "Merge",
        "spellcheck_merge_dialog_message": (
            "Merge « {a} » and « {b} »?\n\n"
            "Yes = rename everything to « {a} »\n"
            "No = rename everything to « {b} »\n"
            "Cancel = do nothing"
        ),
        "spellcheck_merged_title": "Merged",
        "spellcheck_merged_one_message": "« {removed} » has been merged with « {kept} ».",
        "spellcheck_merge_multi_confirm": (
            "Automatically merge these {count} pairs?\n\n"
            "For each pair, the less-used ingredient in your "
            "recipes will be merged into the one used in the most "
            "recipes (the first alphabetically in case of a tie)."
        ),
        "spellcheck_merged_multi_message": "{count} pair(s) merged.",

        # ---- CompareRecipesWindow (Compare two recipes) ----
        "compare_title": "Compare two recipes",
        "compare_recipe_a_label": "Recipe A:",
        "compare_recipe_b_label": "Recipe B:",
        "compare_button": "⚖️ Compare",
        "compare_choose_each_list": "Choose a recipe from each list.",
        "compare_field_category": "Category:",
        "compare_field_favorite": "Favorite:",
        "compare_yes": "⭐ Yes",
        "compare_no": "No",
        "compare_field_rating": "Rating:",
        "compare_field_difficulty": "Difficulty:",
        "compare_field_prep": "Prep:",
        "compare_field_cook": "Cook:",
        "compare_field_total_time": "Total time:",
        "compare_field_cooked": "Cooked:",
        "compare_times_suffix": "{count} times",
        "compare_field_cost": "Estimated cost:",
        "compare_field_nutrition": "Nutrition (kcal):",
        "compare_field_ingredient_count": "Ingredients:",
        "compare_common_ingredients": "🟰 Common ({count})",
        "compare_only_a": "🅰️ Only in « {name} » ({count})",
        "compare_only_b": "🅱️ Only in « {name} » ({count})",
        "compare_none": "None",

        # ---- StatisticsWindow (Statistics) ----
        "stats_title": "Statistics",
        "stats_heading": "=== Statistics ===\n\n",
        "stats_total_recipes": "Total number of recipes: {count}\n\n",
        "stats_by_category": "Breakdown by category:\n",
        "stats_category_line": "  - {category}: {count}\n",
        "stats_by_difficulty": "Breakdown by difficulty:\n",
        "stats_difficulty_line": "  - {difficulty}: {count}\n",
        "stats_difficulty_unspecified": "Not specified",
        "stats_favorites_count": "Favorite recipes: {count}\n\n",
        "stats_avg_rating": "Average rating (rated recipes): {avg} / 5 ({count} recipe(s) rated)\n\n",
        "stats_no_rated_recipe": "Average rating: no recipe rated yet.\n\n",
        "stats_five_star_heading": "5-star recipe(s):\n",
        "stats_recipe_line": "  - {name}\n",
        "stats_most_cooked_heading": "Most cooked recipes:\n",
        "stats_cooked_line": "  - {name}: {count} times\n",
        "stats_none_cooked_yet": (
            "  No recipe marked as cooked yet.\n"
            "  (« 🍳 I cooked this! » button in « View a specific recipe »)\n"
        ),
        "stats_most_used_tags_heading": "Most used tags:\n",
        "stats_tag_line": "  - {tag}: {count}\n",
        "stats_never_cooked_heading": "🕸️ Never cooked recipes:\n",
        "stats_and_others": "  ... and {count} more\n",
        "stats_all_cooked": "  All your recipes have already been cooked at least once. 👏\n",
        "stats_stale_heading": "🕰️ Not cooked in over {days} days:\n",
        "stats_stale_line": "  - {name} ({days} days ago)\n",
        "stats_no_stale_recipe": "  No recipe in this case for now.\n",
        "stats_avg_cost_heading": "💰 Average cost per person:\n",
        "stats_avg_cost_line": (
            "  {avg} € on average, across {count} recipe(s) with at "
            "least one known price ({without_price} without a price set)\n"
        ),
        "stats_no_priced_recipe": (
            "  No recipe with a price set yet.\n"
            "  (see « 💰 Manage prices » in « Manage ingredients »)\n"
        ),
        "stats_avg_kcal_heading": "🥗 Average calories per person:\n",
        "stats_avg_kcal_line": (
            "  {avg} kcal on average, across {count} recipe(s) with "
            "ingredients recognized in the nutrition database\n"
        ),
        "stats_no_recognized_recipe": "  No recipe with recognized ingredients yet.\n",
        "stats_monthly_chart_title": "📈 Recipes cooked per month (last 12 months)",
        "stats_heatmap_title": "🗓️ Calendar of days cooked (last 12 months)",
        "stats_heatmap_legend": "Less ⬜ 🟨 🟧 🟥 More",
        "stats_day_labels": "M,T,W,T,F,S,S",
        "stats_month_labels_short": "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",
        "stats_month_labels_lower": "jan,feb,mar,apr,may,jun,jul,aug,sep,oct,nov,dec",

        # ---- draw_recipe_content (recipe PDF content, standalone or in cookbook) ----
        "recipepdf_category_persons": "Category: {cat}    For {persons} serving(s)",
        "recipepdf_rating": "Rating: {stars}",
        "recipepdf_prep": "Prep: {time} min",
        "recipepdf_cook": "Cook: {time} min",
        "recipepdf_difficulty": "Difficulty: {value}",
        "recipepdf_allergens": "⚠ Allergens: {list}",
        "recipepdf_ingredients_heading": "Ingredients:",
        "recipepdf_cost": "Estimated cost: {cost} €{partial}",
        "recipepdf_partial_suffix": " (partial, {known}/{total})",
        "recipepdf_nutrition": "Estimated nutrition{partial}: {kcal} kcal · {protein}g protein · {carbs}g carbs · {fat}g fat",
        "recipepdf_description_heading": "Description:",
        "recipepdf_notes_heading": "Personal notes:",

        # ---- build_cookbook_pdf (PDF cookbook) ----
        "cookbookpdf_page_number": "Page {current} / {total}",
        "cookbookpdf_generated_on": "Generated on {date}",
        "cookbookpdf_summary_heading": "Table of Contents",
        "cookbookpdf_summary_line": "- [{cat}] {name}",

        # ---- CookbookExportWindow (Export the cookbook) ----
        "cookbookexport_title": "Export the cookbook",
        "cookbookexport_heading": "📖 Export the cookbook",
        "cookbookexport_intro": (
            "Select the recipes to include in a single PDF,\n"
            "cookbook-style."
        ),
        "cookbookexport_filter_label": "Filter by category:",
        "cookbookexport_check_all_button": "Check all",
        "cookbookexport_uncheck_all_button": "Uncheck all",
        "cookbookexport_generate_button": "📄 Generate the book PDF",
        "cookbookexport_error_select_recipe": "Select at least one recipe.",
        "cookbookexport_save_dialog_title": "Save the cookbook",
        "cookbookexport_saved_message": "Cookbook saved:\n{path}",

        # ---- ImportExportWindow (Import / Export data) ----
        "importexport_title": "Import / Export data",
        "importexport_heading": "Back up or transfer your data",
        "importexport_export_intro": (
            "The export creates a .zip file containing absolutely all\n"
            "your data: recipes, photos, custom ingredients,\n"
            "prices, substitutes, pantry, plan and its history,\n"
            "menus, saved shopping lists, trash and\n"
            "settings — to back up or transfer everything to\n"
            "another computer in a single file."
        ),
        "importexport_export_button": "📤 Export all my data (.zip)",
        "importexport_import_intro": (
            "The import reads a previously exported .zip file.\n"
            "\"Merge\" adds duplicate recipes/photos under a\n"
            "new name rather than losing them, and completes the rest\n"
            "(pantry, menus, lists...) without deleting anything.\n"
            "\"Replace\" overwrites everything, including settings and the\n"
            "current plan."
        ),
        "importexport_import_button": "📥 Import data (.zip)",
        "importexport_auto_backups_heading": "🗄️ Automatic backups",
        "importexport_auto_backups_intro": (
            "A backup is automatically created when the\n"
            "application starts (at most one every {hours}h), and the "
            "{retention} most\nrecent are kept here."
        ),
        "importexport_backup_now_button": "💾 Back up now",
        "importexport_restore_selected_button": "♻️ Restore selection",
        "importexport_cloud_heading": "☁️ Automatic cloud backup",
        "importexport_cloud_intro": (
            "Choose a folder synced by a client already\n"
            "installed on this PC (Google Drive, OneDrive, Dropbox...).\n"
            "Each automatic backup will also be copied there, and this\n"
            "client will take care of sending it to the cloud on its own."
        ),
        "importexport_choose_cloud_button": "📁 Choose a cloud folder",
        "importexport_disable_button": "🚫 Disable",
        "importexport_cloud_enabled": "✅ Enabled: {folder}",
        "importexport_cloud_not_configured": "Not configured yet.",
        "importexport_choose_folder_title": "Choose a synced folder (Google Drive, OneDrive, Dropbox...)",
        "importexport_cloud_configured_title": "Folder configured",
        "importexport_cloud_configured_message": (
            "Cloud folder configured:\n{folder}\n\n"
            "Do you want to copy a backup there right now?"
        ),
        "importexport_disabled_title": "Disabled",
        "importexport_disabled_message": "Automatic cloud backup is disabled.",
        "importexport_backup_date_line": "{date}   ({size} KB)",
        "importexport_no_backups": "No automatic backup yet.",
        "importexport_backup_failed": "Backup failed:\n{error}",
        "importexport_backup_created_title": "Backup created",
        "importexport_backup_created_message": "A new automatic backup has been created.",
        "importexport_select_backup_first": "Select a backup from the list.",
        "importexport_restore_mode_title": "Restore mode",
        "importexport_restore_mode_message": (
            "How do you want to restore this backup?\n\n"
            "Yes = Merge (add to current data, without deleting anything)\n"
            "No = Fully replace the current data\n"
            "Cancel = do nothing"
        ),
        "importexport_restore_failed": "Restore failed:\n{error}",
        "importexport_restore_done_title": "Restoration complete",
        "importexport_restore_done_message": "The data has been restored successfully.",
        "importexport_export_data_title": "Export my data",
        "importexport_export_data_success": "Your data has been exported to:\n{path}",
        "importexport_choose_archive_title": "Choose an archive to import",
        "importexport_import_mode_title": "Import mode",
        "importexport_import_mode_message": (
            "How do you want to import this data?\n\n"
            "Yes = Merge (add to current data, without deleting anything)\n"
            "No = Fully replace the current data\n"
            "Cancel = do nothing"
        ),
        "importexport_import_failed": "Import failed:\n{error}",
        "importexport_import_done_title": "Import complete",
        "importexport_import_done_message": "The data has been imported successfully.",

        # ---- ShoppingChecklistWindow (Shopping mode) ----
        "checklist_instruction": "Check off each item as you go through your shopping.",
        "checklist_check_all_button": "☑️ Check all",
        "checklist_uncheck_all_button": "⬜ Uncheck all",
        "checklist_progress_label": "{done} / {total} item(s) checked",

        # ---- ExportFormatDialog (Choose an export format) ----
        "exportformat_title": "Choose an export format",
        "exportformat_heading": "📤 Export the shopping list",
        "exportformat_choose_label": "Choose the desired export format:",
        "exportformat_txt_button": "📝 Export as text (.txt)",
        "exportformat_excel_button": "📊 Export as Excel (.xlsx)",
        "exportformat_pdf_button": "📄 Export as PDF (.pdf)",
        "exportformat_cancel_button": "Cancel",

        # ---- MenuManagerWindow (My menus) ----
        "menumanager_title": "My menus",
        "menumanager_list_label": "My saved menus:",
        "menumanager_new_button": "➕ New menu",
        "menumanager_recipe_count": "{name} ({count} recipe(s))",
        "menumanager_select_menu_first": "Select a menu from the list.",
        "menumanager_delete_confirm": "Delete the menu « {name} »?",

        # ---- MenuFormWindow (New menu / Edit the menu) ----
        "menuform_title_edit": "Edit the menu",
        "menuform_title_new": "New menu",
        "menuform_name_label": "Menu name:",
        "menuform_add_recipe_label": "Add a recipe to the menu:",
        "menuform_persons_short_label": "servings:",
        "menuform_add_button": "+ Add",
        "menuform_recipes_label": "Menu recipes:",
        "menuform_remove_button": "🗑 Remove from menu",
        "menuform_save_button": "💾 Save the menu",
        "menuform_compute_button": "Calculate the menu's shopping list",
        "menuform_empty_list_message": (
            "No list calculated yet.\n"
            "Click « Calculate the menu's shopping list » above,\n"
            "or load a saved list."
        ),
        "menuform_total_list_heading": "=== Menu shopping list ===",
        "menuform_item_row_label": "[{cat}] {name} ({persons} servings)",
        "menuform_select_recipe_to_remove": "Select a menu recipe to remove.",
        "menuform_error_name_required": "Please enter a menu name.",
        "menuform_error_no_recipe": "Add at least one recipe to the menu.",
        "menuform_saved_message": "The menu « {name} » has been saved.",
        "menuform_calculate_list_for_export": (
            "First calculate a shopping list (« Calculate the menu's shopping list » button)."
        ),
        "menuform_add_recipe_or_manual": "Add at least one recipe to the menu, or add an ingredient manually.",
        "menuform_shopping_list_title": "Menu: {name}",
        "menuform_print_label": "the menu « {name} »",

        # ---- ImportFromUrlWindow (Import a recipe from a link) ----
        "importurl_title": "Import a recipe from a link",
        "importurl_heading": "🌐 Import a recipe from a link",
        "importurl_intro": (
            "Paste the address (URL) of a recipe page. This works\n"
            "with most major cooking sites (which use a\n"
            "standard data format). An internet connection is required."
        ),
        "importurl_fetch_button": "🌐 Fetch the recipe",
        "importurl_after_import_note": (
            "After import, check and complete the recipe if needed\n"
            "(detecting quantities and units isn't always perfect)."
        ),
        "importurl_paste_url_first": "First paste a web address (URL).",
        "importurl_fetching": "Fetching...",
        "importurl_failed_title": "Import failed",

        # ---- ImportFromPhotoWindow (Import a recipe from a photo) ----
        "importphoto_title": "Import a recipe from a photo",
        "importphoto_heading": "📷 Import a recipe from a photo",
        "importphoto_intro": (
            "Take a photo of (or scan) a handwritten recipe or a\n"
            "cookbook page, then choose the image here. The text\n"
            "is extracted automatically, but still needs to be reviewed and organized\n"
            "yourself (unlike importing from a link, a photo has no\n"
            "ingredients/steps structure that can be guessed)."
        ),
        "importphoto_module_warning": (
            "⚠ This feature requires the 'pytesseract' module\n"
            "AND the Tesseract OCR program installed separately on this PC.\n"
            "See the README for installation instructions."
        ),
        "importphoto_no_photo_chosen": "No photo chosen",
        "importphoto_choose_button": "📁 Choose a photo",
        "importphoto_extract_button": "🔍 Extract text",
        "importphoto_extracted_text_label": "Extracted text (editable):",
        "importphoto_create_button": "➡️ Create the recipe with this text",
        "importphoto_choose_photo_title": "Choose a recipe photo",
        "importphoto_choose_first": "First choose a photo.",
        "importphoto_ocr_module_missing": (
            "This feature requires the 'pytesseract' module\n"
            "(pip install pytesseract) AND the Tesseract OCR program\n"
            "installed separately on this PC. See the README."
        ),
        "importphoto_extraction_failed_title": "Extraction failed",
        "importphoto_extraction_failed_message": (
            "Text recognition failed. Check that Tesseract OCR "
            "is properly installed on this PC and accessible.\n\nDetail: {error}"
        ),
        "importphoto_no_text_extracted": (
            "No text could be extracted from this photo. Try a sharper, "
            "better-framed, or better-lit image."
        ),
        "importphoto_no_text_title": "No text",
        "importphoto_no_text_confirm": (
            "No text was extracted or entered. Create an empty "
            "recipe anyway (with just the photo)?"
        ),

        # ---- TrashWindow (Trash) ----
        "trash_title": "Trash",
        "trash_heading": "🗑️ Deleted recipes",
        "trash_intro": (
            "Photos of recipes in the trash are kept\n"
            "until they are permanently deleted."
        ),
        "trash_restore_button": "♻️ Restore",
        "trash_delete_forever_button": "🗑️ Delete permanently",
        "trash_empty_button": "🧹 Empty trash",
        "trash_unnamed_recipe": "(unnamed)",
        "trash_unknown_date": "unknown date",
        "trash_entry_line": "{name}  —  deleted on {date}",
        "trash_is_empty": "The trash is empty.",
        "trash_select_recipe_first": "Select a recipe from the trash.",
        "trash_restored_suffix": "{name} (restored)",
        "trash_restored_title": "Restored",
        "trash_restored_message": "« {name} » has been restored.",
        "trash_delete_forever_confirm": "Permanently delete « {name} »?\n\nThis action cannot be undone.",
        "trash_deleted_title": "Deleted",
        "trash_deleted_message": "The recipe has been permanently deleted.",
        "trash_already_empty": "The trash is already empty.",
        "trash_empty_confirm": (
            "Permanently delete the {count} recipe(s) in the trash?\n\n"
            "This action cannot be undone."
        ),
        "trash_emptied_title": "Trash emptied",
        "trash_emptied_message": "The trash has been emptied.",

        # ---- CookingModeWindow (Fullscreen cooking mode) ----
        "cookingmode_title": "Cooking mode — {name}",
        "cookingmode_close_button": "✕ Close (Esc)",
        "cookingmode_cooked_button": "🍳 I cooked this!",
        "cookingmode_fullscreen_hint": "F11: fullscreen",
        "cookingmode_persons_suffix": "{persons} servings",
        "cookingmode_speech_button": "🔊 Read aloud",
        "cookingmode_speech_stop_button": "⏹ Stop reading",
        "cookingmode_volume_percent": "{percent}%",
        "cookingmode_tts_module_missing": "Reading aloud requires the 'pyttsx3' module.\nInstall it with: pip install pyttsx3",
        "cookingmode_no_description_to_read": "This recipe has no description to read (the description field is empty).",
        "cookingmode_ingredients_heading": "Ingredients",
        "cookingmode_prep_label": "Prep: {time} min",
        "cookingmode_cook_label": "Cook: {time} min",
        "cookingmode_difficulty_label": "Difficulty: {value}",
        "cookingmode_preparation_heading": "Preparation",
        "cookingmode_personal_notes_heading": "Personal notes",

        # ---- IngredientSearchWindow (Search by ingredient) ----
        "ingsearch_title": "Search by ingredient",
        "ingsearch_question_label": "Which ingredient are you looking for?",
        "ingsearch_view_recipes_button": "🔍 See recipes using it",
        "ingsearch_view_selected_button": "📖 View selected recipe",
        "ingsearch_no_recipe_uses": "No recipe uses « {name} » yet.",
        "ingsearch_recipes_using": "Recipes using « {name} » ({count}):",
        "ingsearch_result_line": "{star}[{cat}] {name} ({qty}{unit} per person)",
        "ingsearch_select_result_first": "Select a recipe from the results list.",

        # ---- TimerRow (a timer row within TimersWindow) ----
        "timerrow_minutes_label": "Min:",
        "timerrow_seconds_label": "Sec:",
        "timerrow_error_invalid_duration": "Invalid duration.",
        "timerrow_set_duration_first": "Set a duration before starting.",

        # ---- CookLogEntryDialog (Add to cooking log) ----
        "cooklogentry_title": "📔 Add to cooking log",
        "cooklogentry_heading": "🍳 « {name} »",
        "cooklogentry_intro": "How was it? A note and/or a photo\n(optional, you can also just skip this).",
        "cooklogentry_no_photo_chosen": "No photo chosen",
        "cooklogentry_choose_photo_button": "📷 Choose a photo",
        "cooklogentry_skip_button": "Skip",
        "cooklogentry_choose_photo_title": "Choose a photo",

        # ---- CookLogWindow (Cooking log) ----
        "cooklog_title": "📔 Cooking log — {name}",
        "cooklog_heading": "📔 {name}",
        "cooklog_times_cooked": "Cooked {count} times in total",
        "cooklog_no_entry": "No note saved yet.\nUse « 🍳 I cooked this! » to add one.",
        "cooklog_no_note": "(no note)",

        # ---- TimersWindow (Timers) ----
        "timers_title": "⏲️ Timers",
        "timers_intro": (
            "Set each timer then ▶️ to start it.\n"
            "When done, the row flashes red with a sound alert."
        ),
        "timers_add_button": "➕ Add a timer",

        # ---- QRCodeWindow (Recipe QR Code) ----
        "qrcode_title": "QR Code — {name}",
        "qrcode_intro": (
            "Scan with the camera app or a QR code\n"
            "reader app to see the name and ingredients."
        ),
        "qrcode_save_button": "💾 Save as image (PNG)",
        "qrcode_truncated_warning": (
            "⚠️ The recipe is long: the QR code contains a\n"
            "truncated summary (name + ingredients only)."
        ),
        "qrcode_encoded_ingredients_heading": "Ingredients ({persons} servings):",
        "qrcode_save_dialog_title": "Save the QR code",
        "qrcode_save_failed": "Save failed:\n{error}",
        "qrcode_saved_message": "QR code saved:\n{path}",

        # ---- UnitConverterWindow (Unit converter) ----
        "unitconv_title": "Unit converter",
        "unitconv_heading": "🔄 Unit converter",
        "unitconv_intro": (
            "Approximate conversion based on the density of water for\n"
            "volume units (ml, cl, L, cup, spoons): reliable for\n"
            "liquids, approximate for solids like flour\n"
            "or sugar, whose actual density differs a bit."
        ),
        "unitconv_quantity_label": "Quantity:",
        "unitconv_from_label": "From:",
        "unitconv_to_label": "To:",
        "unitconv_convert_button": "Convert",
        "unitconv_error_invalid_quantity": "Invalid quantity.",
        "unitconv_result": "{quantity} {from_unit} ≈ {result} {to_unit}",
        "unitconv_gram": "Gram (g)",
        "unitconv_kilogram": "Kilogram (kg)",
        "unitconv_ounce": "Ounce (oz)",
        "unitconv_pound": "Pound (lb)",
        "unitconv_milliliter": "Milliliter (ml)",
        "unitconv_centiliter": "Centiliter (cl)",
        "unitconv_liter": "Liter (L)",
        "unitconv_teaspoon": "Teaspoon (5 ml)",
        "unitconv_tablespoon": "Tablespoon (15 ml)",
        "unitconv_cup": "US Cup (240 ml)",

        # ---- DisclaimerWindow (Disclaimer) ----
        "disclaimer_title": "Disclaimer",
        "disclaimer_heading": "⚠ Disclaimer",
        "disclaimer_intro": "Please read this text before using the application.",
        "disclaimer_checkbox": "I have read and accept the terms above",
        "disclaimer_continue_button": "Continue",
        "disclaimer_quit_button": "Quit the application",
        "disclaimer_text": (
            "ARTICLE 1 – EXCLUSION AND LIMITATION OF LIABILITY\n\n"
            "1.1. Medical alerts and allergen management\n\n"
            "The Application offers a feature allowing the User to enter, modify, and configure their own "
            "allergy and allergen criteria. The User expressly acknowledges that:\n\n"
            "• The accuracy and upkeep of this information is the User's sole and exclusive responsibility.\n"
            "• The Application is a software tool to assist with browsing recipes and does not, under any "
            "circumstances, replace medical advice, a diagnosis, or human verification of ingredients.\n"
            "• The Publisher cannot be held liable for incorrect entries, omissions, misconfiguration by the "
            "User, or an allergic reaction (intolerance, anaphylactic shock, etc.) occurring after eating a "
            "dish. It is the User's responsibility to systematically check the labels and actual composition "
            "of each physical ingredient before any preparation or consumption.\n\n"
            "1.2. Provided \"as is\" and free of charge\n\n"
            "The Application is made available to the User entirely free of charge. It is provided \"as is\" "
            "and \"as available\", without any guarantee of being free of errors, software bugs, or "
            "interruptions. The Publisher does not guarantee that the Application's features will meet the "
            "User's specific needs.\n\n"
            "1.3. Material and immaterial damages\n\n"
            "The Publisher disclaims all liability for direct or indirect damages caused to the User or to "
            "third parties. In particular, the Publisher cannot be held liable for:\n\n"
            "• A failure, overheating, malfunction, or deterioration of the User's computer hardware or "
            "smartphone while using the Application.\n"
            "• A loss of computer data, alteration of files, or hacking of the User's system.\n\n"
            "Because the service is provided free of charge, should the Publisher's liability be established "
            "by a court, the amount of damages would be expressly capped at zero euros (€0)."
        ),

        # ---- AddManualIngredientDialog (Add ingredients to shopping list) ----
        "addmanual_title": "Add ingredients to the shopping list",
        "addmanual_heading": "➕ Add ingredients to the shopping list",
        "addmanual_intro": (
            "Add as many ingredients as you want to the pending\n"
            "list below, then confirm them all at once."
        ),
        "addmanual_new_ingredient_button": "🥕 New ingredient",
        "addmanual_add_to_list_button": "➕ Add to list",
        "addmanual_staged_label": "Ingredients pending confirmation:",
        "addmanual_remove_staged_button": "🗑 Remove from pending list",
        "addmanual_confirm_all_button": "✅ Confirm all these ingredients",
        "addmanual_close_button": "Close",
        "addmanual_select_staged_first": "Select an ingredient from the pending list.",
        "addmanual_add_staged_first": "Add at least one ingredient to the pending list before confirming.",
        "addmanual_confirmed_message": "{count} ingredient(s) added to the shopping list.",

        # ---- Shopping list export functions (text/Excel/PDF) ----
        "shoppingexport_generated_on": "Generated on {date}",
        "shoppingexport_selected_recipes": "Selected recipes:",
        "shoppingexport_excel_sheet_recipes": "Recipes",
        "shoppingexport_excel_col_recipe": "Recipe",
        "shoppingexport_excel_col_persons": "Number of servings",
        "shoppingexport_excel_sheet_ingredients": "Ingredients",
        "shoppingexport_excel_col_rayon": "Aisle",
        "shoppingexport_excel_col_ingredient": "Ingredient",
        "shoppingexport_excel_col_total_qty": "Total quantity",
        "shoppingexport_excel_col_unit": "Unit",

        # ---- SavedShoppingListsWindow (Saved shopping lists) ----
        "savedlists_title": "Saved shopping lists",
        "savedlists_heading": "📂 Saved shopping lists",
        "savedlists_load_button": "📂 Load",
        "savedlists_delete_button": "🗑 Delete",
        "savedlists_none_saved": "No list saved yet.",
        "savedlists_entry_line": "{name} — {count} ingredient(s) — {date}",
        "savedlists_select_list_first": "Select a list from the list.",
        "savedlists_delete_confirm": "Permanently delete the list « {name} »?",

        # ---- QuickSearchWindow (Quick search, Ctrl+K) ----
        "quicksearch_title": "Quick search",
        "quicksearch_heading": "🔍 Quick recipe search",
        "quicksearch_no_results": "No recipe found.",
        "quicksearch_footer_hint": "Enter to open, Esc to close.",

        # ---- AllRecipesWindow (View all recipes / shopping list) ----
        "allrecipes_title": "All recipes - Shopping list",
        "allrecipes_select_label": "Select recipes and the number of servings:",
        "allrecipes_ingredient_filter_title": "Filter by ingredient",
        "allrecipes_persons_count_label": "Servings:",
        "allrecipes_add_to_cart_button": "🛒 Add to shopping list",
        "allrecipes_checklist_mode_button": "☑️ Shopping mode (check off as you go)",
        "allrecipes_clear_list_button": "🗑 Clear shopping list",
        "allrecipes_export_button": "📤 Export",
        "allrecipes_print_button": "🖨️ Print",
        "allrecipes_add_manual_ingredient_button": "➕ Add an ingredient to the shopping list",
        "allrecipes_save_list_button": "💾 Save this list for later",
        "allrecipes_load_list_button": "📂 Load a saved list",
        "allrecipes_invalid_persons": "Invalid number of servings for « {name} ».",
        "allrecipes_empty_list_message": (
            "Your shopping list is empty for now.\n"
            "Click « 🛒 Add to shopping list » next to a recipe,\n"
            "add an ingredient manually, or load a saved list."
        ),
        "allrecipes_total_list_heading": "=== Total shopping list ===",
        "allrecipes_manual_items_note": "({count} manually added ingredient(s) included)",
        "allrecipes_invalid_quantity": "Invalid quantity.",
        "allrecipes_calculate_list_first": "First calculate a shopping list before saving it.",
        "allrecipes_save_list_dialog_title": "Save the list",
        "allrecipes_save_list_dialog_prompt": "Name for this list:",
        "allrecipes_list_saved_title": "Saved",
        "allrecipes_list_saved_message": "List « {name} » saved for later.",
        "allrecipes_empty_list_for_export": (
            "The shopping list is empty. Add at least one recipe "
            "(« 🛒 Add to shopping list » button) or a manual ingredient."
        ),
        "allrecipes_export_txt_title": "Save the shopping list as text",
        "allrecipes_export_excel_title": "Save the shopping list as Excel",
        "allrecipes_export_pdf_title": "Save the shopping list as PDF",
        "allrecipes_export_saved_message": "Shopping list saved:\n{path}",
        "allrecipes_excel_module_missing": "Excel export requires the 'openpyxl' module.\nInstall it with: pip install openpyxl",
        "allrecipes_pdf_module_missing": "PDF export requires the 'reportlab' module.\nInstall it with: pip install reportlab",
        "allrecipes_print_module_missing": (
            "Printing requires the 'reportlab' module to generate the layout.\n"
            "Install it with: pip install reportlab"
        ),
        "allrecipes_print_label": "the shopping list",
        "allrecipes_shopping_list_title": "Shopping list",
        "allrecipes_close_confirm_title": "Close the shopping list?",
        "allrecipes_close_confirm_message": (
            "The displayed shopping list is not saved: it will be "
            "permanently lost if you close this window now.\n\n"
            "Tip: use « 💾 Save this list for later » "
            "before closing if you want to keep it.\n\n"
            "Close anyway?"
        ),

        # ---- ManageRecipesWindow (Edit / Delete a recipe) ----
        "managerecipes_title": "Edit / Delete a recipe",
        "managerecipes_select_label": "Select a recipe:",
        "managerecipes_filter_favorites": "⭐ Favorites only",
        "managerecipes_filter_quick": "⏱️ Quick recipes (≤ 30 min) only",
        "managerecipes_filter_vegetarian": "🥗 Vegetarian recipes only",
        "managerecipes_filter_wishlist": "💭 Wish list only",
        "managerecipes_remove_filter_button": "✕ Remove filter",
        "managerecipes_search_label": "🔍 Search:",
        "managerecipes_sort_label": "Sort by:",
        "managerecipes_category_label": "Category:",
        "managerecipes_edit_button": "✏️ Edit",
        "managerecipes_duplicate_button": "📋 Duplicate",
        "managerecipes_delete_button": "🗑️ Delete",
        "managerecipes_select_recipe_first": "Select a recipe from the list.",
        "managerecipes_duplicate_suffix": "(copy)",
        "managerecipes_duplicated_title": "Duplicated",
        "managerecipes_duplicated_message": "« {original} » has been duplicated as « {new} ».",
        "managerecipes_delete_confirm_message": (
            "Send the recipe « {name} » to the trash?\n\n"
            "You can restore it later from the « 🗑️ Trash » button."
        ),
        "managerecipes_deleted_title": "Sent to trash",
        "managerecipes_deleted_message": "The recipe has been moved to the trash.",

        # ---- OneRecipeWindow (View a specific recipe) ----
        "onerecipe_window_title": "View a recipe",
        "onerecipe_choose_recipe_label": "Choose a recipe:",
        "onerecipe_search_label": "🔍 Search:",
        "onerecipe_sort_label": "Sort:",
        "onerecipe_category_label": "Category:",
        "onerecipe_persons_label": "Number of servings:",
        "onerecipe_btn_show": "Show the recipe",
        "onerecipe_btn_export_pdf": "📄 Export as PDF",
        "onerecipe_btn_print": "🖨️ Print",
        "onerecipe_btn_add_to_shopping": "🛒 Add to shopping list",
        "onerecipe_btn_cooked": "🍳 I cooked this!",
        "onerecipe_btn_cooking_mode": "🖥️ Cooking mode (fullscreen)",
        "onerecipe_btn_qr": "📱 QR Code",
        "onerecipe_btn_timers": "⏲️ Timers",
        "onerecipe_btn_cook_log": "📔 Cooking log",
        "onerecipe_btn_substitutions": "🔄 Possible substitutes",
        "onerecipe_edit_button": "✏️ Edit",
        "onerecipe_ingredients_info_label": "Ingredients and information:",
        "onerecipe_description_notes_label": "Description and notes:",
        "onerecipe_similar_label": "Similar recipes:",
        "onerecipe_no_photo": "(no photo)",
        "onerecipe_preview_unavailable": "(preview unavailable)",
        "onerecipe_select_recipe_first": "Select a recipe from the list.",
        "onerecipe_display_first": "First display a recipe with « Show the recipe ».",
        "onerecipe_invalid_persons": "Invalid number of servings.",
        "onerecipe_added_to_shopping_title": "Added",
        "onerecipe_added_to_shopping_message": (
            "« {name} » ({persons} servings) will be automatically added to the shopping list "
            "next time you open « View all recipes »."
        ),
        "onerecipe_pantry_decrement_title": "Pantry",
        "onerecipe_pantry_decrement_prompt": "Deduct the ingredients of « {name} » ({persons} servings) from your pantry?",
        "onerecipe_pantry_updated_title": "Pantry updated",
        "onerecipe_pantry_updated_message": "{count} ingredient(s) deducted from your pantry.",
        "onerecipe_pantry_none_decremented": (
            "None of this recipe's ingredients could be deducted "
            "(missing from the pantry, or unit not comparable)."
        ),
        "onerecipe_marked_title": "Marked",
        "onerecipe_marked_message": "« {name} » has been marked as cooked today!",
        "onerecipe_no_substitutes_title": "No known substitute",
        "onerecipe_no_substitutes_message": (
            "None of this recipe's ingredients has a known substitute yet.\n\n"
            "You can add one yourself from « 🥕 Manage ingredients » > "
            "« 🔄 Manage substitutions »."
        ),
        "onerecipe_substitutes_title": "Possible substitutes — {name}",
        "onerecipe_substitutes_heading": "🔄 Possible substitutes for « {name} »",
        "onerecipe_substitutes_disclaimer": (
            "Culinary suggestions, not guaranteed equivalences:\nthe result may vary depending on the recipe."
        ),
        "onerecipe_close_button": "Close",
        "onerecipe_rating_label": "Rating: {stars}",
        "onerecipe_prep_label": "Prep: {time} min",
        "onerecipe_cook_label": "Cook: {time} min",
        "onerecipe_difficulty_label": "Difficulty: {value}",
        "onerecipe_allergens_label": "⚠ Allergens: {list}",
        "onerecipe_cost_label": "💰 Estimated cost: {cost} €{partial}",
        "onerecipe_cost_partial": " (partial estimate, {known}/{total} ingredients with known price)",
        "onerecipe_nutrition_partial": " (partial estimate, {known}/{total} ingredients recognized)",
        "onerecipe_nutrition_label": (
            "🥗 Estimated nutritional values{partial}:\n"
            "   {kcal} kcal · {protein} g protein · {carbs} g carbs · {fat} g fat\n"
        ),
        "onerecipe_description_heading": "--- Description ---\n{text}\n",
        "onerecipe_notes_heading": "\n--- Personal notes ---\n{text}\n",
        "onerecipe_no_description_notes": "(No description or personal notes for this recipe.)",
        "onerecipe_export_pdf_title": "Export the recipe as PDF",
        "onerecipe_export_success_title": "Export successful",
        "onerecipe_export_success_message": "Recipe exported:\n{path}",
        "onerecipe_export_failed": "Export failed:\n{error}",
        "onerecipe_print_failed": "Print preparation failed:\n{error}",
        "onerecipe_pdf_module_missing": "PDF export requires the 'reportlab' module.\nInstall it with: pip install reportlab",
        "onerecipe_print_module_missing": (
            "Printing requires the 'reportlab' module to generate the layout.\n"
            "Install it with: pip install reportlab"
        ),
        "onerecipe_qr_module_missing": "QR code export requires the 'qrcode' module.\nInstall it with: pip install qrcode",
        "onerecipe_qr_pillow_missing": (
            "QR code export also requires the 'Pillow' module.\nInstall it with: pip install pillow"
        ),
        "onerecipe_default_timer_label": "Timer",

        # ---- RecipeFormWindow (Add / Edit a recipe) ----
        "recipeform_title_edit": "Edit recipe",
        "recipeform_title_add": "Add a recipe",
        "recipeform_name_label": "Recipe name:",
        "recipeform_favorite_checkbox": "⭐ Mark as favorite recipe",
        "recipeform_wishlist_checkbox": "💭 Add to my wish list (to try)",
        "recipeform_rating_label": "My rating:",
        "recipeform_category_label": "Category:",
        "recipeform_prep_time_label": "Prep time (min):",
        "recipeform_cook_time_label": "Cook time (min):",
        "recipeform_difficulty_label": "Difficulty:",
        "recipeform_default_persons_label": "   Default servings:",
        "recipeform_tags_label": "Tags (comma-separated):",
        "recipeform_tags_example": "e.g. vegetarian, gluten-free, quick, budget-friendly",
        "recipeform_allergens_label": "Allergens present:",
        "recipeform_detect_allergens_button": "🔍 Detect automatically",
        "recipeform_allergens_disclaimer": (
            "This is for informational purposes only, always check\n"
            "allergens on the actual product labels."
        ),
        "recipeform_allergens_auto_note": (
            "Automatic detection is based on the recipe's ingredients\n"
            "already entered below: it checks and unchecks boxes\n"
            "accordingly, without ever touching boxes you would have\n"
            "checked yourself unrelated to a detected ingredient."
        ),
        "recipeform_photos_label": "Photos:",
        "recipeform_add_photo_button": "📷 Add a photo",
        "recipeform_description_label": "Description (information, steps, tips...):",
        "recipeform_notes_label": "Personal notes (review, adjustments for next time...):",
        "recipeform_ingredients_label": "Ingredients (quantity for 1 person):",
        "recipeform_new_ingredient_button": "🥕 New ingredient",
        "recipeform_no_ingredients_registered": (
            "No ingredient registered yet. Click « 🥕 New ingredient »\nto create your first one."
        ),
        "recipeform_header_ingredient": "Ingredient",
        "recipeform_header_quantity": "Quantity",
        "recipeform_header_unit": "Unit",
        "recipeform_header_other": "(if other)",
        "recipeform_add_ingredient_button": "+ Add an ingredient",
        "recipeform_save_button": "Save",
        "recipeform_delete_button": "Delete this recipe",
        "recipeform_char_counter": "{count} / {max} characters",
        "recipeform_add_ingredients_first": "First add ingredients to the recipe.",
        "recipeform_allergens_updated_title": "Allergens updated",
        "recipeform_allergens_updated_added": "added: {list}",
        "recipeform_allergens_updated_removed": "removed: {list}",
        "recipeform_allergens_updated_message": "Allergen(s) {parts}.",
        "recipeform_allergens_no_change": "No change: the checked allergens already match the ingredients.",
        "recipeform_choose_photos_title": "Choose one or more photos",
        "recipeform_no_photo": "(no photo)",
        "recipeform_preview_unavailable": "(preview\nunavailable)",
        "recipeform_remove_photo_button": "🗑 Remove",
        "recipeform_new_ingredient_dialog_title": "New ingredient",
        "recipeform_new_ingredient_dialog_prompt": "Name of the new ingredient:",
        "recipeform_ingredient_already_exists": "The ingredient « {name} » already exists.",
        "recipeform_ingredient_added_title": "Added",
        "recipeform_ingredient_added_message": (
            "The ingredient « {name} » has been added.\nSelect it from one of the dropdown lists."
        ),
        "recipeform_error_name_required": "Please enter a recipe name.",
        "recipeform_error_prep_time": "Prep time must be a positive number (or empty).",
        "recipeform_error_cook_time": "Cook time must be a positive number (or empty).",
        "recipeform_unknown_ingredient_title": "Unknown ingredient",
        "recipeform_unknown_ingredient_message": (
            "« {name} » doesn't match any registered ingredient.\n"
            "Choose one from the dropdown list, or click "
            "« 🥕 New ingredient » to add it first."
        ),
        "recipeform_error_invalid_quantity": "Invalid quantity for '{name}'.",
        "recipeform_error_custom_unit_required": "Please specify the custom unit for '{name}'.",
        "recipeform_error_no_valid_ingredient": "Add at least one valid ingredient.",
        "recipeform_duplicate_ingredient_title": "Duplicate ingredient",
        "recipeform_duplicate_ingredient_message": "« {list} » appears multiple times in this recipe.\n\nSave anyway?",
        "recipeform_saved_message": "The recipe « {name} » has been saved.",
        "recipeform_delete_confirm_message": (
            "Send the recipe « {name} » to the trash?\n\n"
            "You can restore it later from the « 🗑️ Trash » button."
        ),
        "recipeform_deleted_title": "Sent to trash",
        "recipeform_deleted_message": "The recipe has been moved to the trash.",
    },
    "es": {
        'home_window_title': 'Mi Libro de Recetas',
        'home_banner_title': '👨\u200d🍳 Mi Libro de Recetas',
        'home_banner_subtitle': 'Todas tus recetas, al alcance de la mano',
        'home_donate_button': '☕ Hacer una donación',
        'home_dark_theme': '🌙 Tema oscuro',
        'home_light_theme': '☀️ Tema claro',
        'home_large_text_on': '🔎 Texto ampliado',
        'home_large_text_off': '🔎 Texto normal',
        'home_daily_recipe_title': '🎲 Receta del día',
        'home_open_button': '👁 Abrir',
        'home_quick_filter_favorites': '⭐ Favoritos',
        'home_quick_filter_quick': '⏱️ Rápido (≤ 30 min)',
        'home_quick_filter_vegetarian': '🥗 Vegetariano',
        'home_quick_filter_wishlist': '💭 Deseos',
        'home_wishlist_reminder': '💭 {count} receta(s) en tu lista de deseos desde hace más de {days} días — ¿por qué no las pruebas? (haz clic para verlas)',
        'home_low_stock_reminder': '📦 {count} ingrediente(s) casi agotado(s) en tu despensa: {names} — haz clic para añadirlos a la lista de compras',
        'home_btn_add_recipe': '➕  Añadir una receta',
        'home_btn_import_url': '🌐  Importar una receta desde un enlace',
        'home_btn_import_photo': '📷  Importar una receta desde una foto',
        'home_btn_view_all_recipes': '🧾  Ver todas las recetas (lista de compras)',
        'home_btn_view_one_recipe': '🍽️  Ver una receta concreta',
        'home_btn_manage_recipes': '✏️  Modificar / Eliminar una receta',
        'home_btn_compare_recipes': '⚖️  Comparar dos recetas',
        'home_btn_manage_ingredients': '🥕  Gestionar los ingredientes',
        'home_btn_ingredient_search': '🔎  Búsqueda por ingrediente',
        'home_btn_what_can_i_cook': '🧊  ¿Qué puedo cocinar?',
        'home_btn_pantry': '📦  Mi despensa',
        'home_btn_unit_converter': '🔄  Conversor de unidades',
        'home_btn_weekly_plan': '📅  Planificación semanal',
        'home_btn_menus': '📋  Mis menús',
        'home_btn_statistics': '📊  Estadísticas',
        'home_btn_export_cookbook': '📖  Exportar el libro de recetas',
        'home_btn_import_export': '💾  Importar / Exportar datos',
        'home_btn_trash': '🗑️  Papelera',
        'home_today_title': '📅 Hoy',
        'home_recent_title': '🕘 Vistas recientemente',
        'home_wishlist_title': '💭 Recetas para probar',
        'home_new_draw_button': '🎲 Nuevo sorteo',
        'home_footer_recipe_count': '{count} receta(s) guardada(s)',
        'home_nothing_planned': 'Nada planificado para {day}. Completa la « 📅 Planificación semanal » para verlo aquí.',
        'home_no_recent_recipe': 'Ninguna receta consultada por el momento.',
        'home_no_wishlist_recipe': 'Ninguna receta en tu lista de deseos por el momento.',
        'warning_pillow': 'Pillow no está instalado: las fotos no se mostrarán (pip install pillow)',
        'warning_reportlab': 'reportlab no está instalado: exportación a PDF no disponible (pip install reportlab)',
        'warning_openpyxl': 'openpyxl no está instalado: exportación a Excel no disponible (pip install openpyxl)',
        'warning_qrcode': 'qrcode no está instalado: exportación de código QR no disponible (pip install qrcode)',
        'warning_pytesseract': 'pytesseract no está instalado: importación desde foto no disponible (pip install pytesseract, + Tesseract OCR)',
        'common_error': 'Error',
        'common_info': 'Información',
        'common_confirm': 'Confirmar',
        'common_success': 'Éxito',
        'common_module_missing': 'Módulo faltante',
        'common_all_categories': 'Todas',
        'common_export_failed': 'La exportación falló:\n{error}',
        'common_export_success_title': 'Exportación exitosa',
        'common_print_failed': 'La preparación de la impresión falló:\n{error}',
        'common_reset_button': 'Restablecer',
        'common_want_label': 'Quiero:',
        'common_exclude_label': 'No quiero:',
        'common_tags_filter_label': 'Etiquetas (todas requeridas):',
        'common_filter_hint': 'Escribe las primeras letras para filtrar la lista.',
        'common_search_label': '🔍 Buscar:',
        'common_sort_by_label': 'Ordenar por:',
        'common_category_label': 'Categoría:',
        'common_edit_button': '✏️ Modificar',
        'common_unknown_ingredient_title': 'Ingrediente desconocido',
        'common_unknown_ingredient_simple_message': '« {name} » no corresponde a ningún ingrediente registrado.\nElige uno de la lista desplegable.',
        'common_ingredient_label': 'Ingrediente:',
        'common_quantity_label': 'Cantidad:',
        'common_unit_label': 'Unidad:',
        'common_new_ingredient_button': '🥕 Nuevo ingrediente',
        'common_save_button': '💾 Guardar',
        'pantry_title': 'Mi despensa',
        'pantry_heading': '📦 Mi despensa',
        'pantry_intro': 'Indica lo que tienes en casa y en qué cantidad.\n«¿Qué puedo cocinar?» podrá entonces comprobar si tienes suficiente,\ny proponer descontar automáticamente el stock después de cocinar.',
        'pantry_threshold_label': 'Umbral de alerta (opcional):',
        'pantry_help_text': 'Para AÑADIR un artículo: indica el ingrediente (créalo primero con\n« 🥕 Nuevo ingrediente » si aún no está en tu lista), la\ncantidad y la unidad, luego haz clic en « 💾 Guardar ».\nPara MODIFICAR un artículo ya existente: haz clic una vez sobre él en la\nlista de abajo — esto carga sus valores en los campos de arriba,\nsin guardar nada: cambia los valores deseados Y LUEGO haz clic en\n« 💾 Guardar » para que el cambio se aplique.\nEl umbral de alerta activa un recordatorio en la página de inicio en cuanto\nla cantidad baja de ese nivel (déjalo vacío para no ser alertado nunca).',
        'pantry_remove_button': '🗑 Quitar de la despensa',
        'pantry_empty': 'Tu despensa está vacía por el momento.',
        'pantry_threshold_suffix': ' (umbral: {threshold})',
        'pantry_error_ingredient_required': 'Por favor, indica un ingrediente.',
        'pantry_error_invalid_quantity': 'Cantidad no válida.',
        'pantry_error_invalid_threshold': 'Umbral de alerta no válido (déjalo vacío si no quieres uno).',
        'pantry_select_ingredient_first': 'Selecciona un ingrediente de la lista.',
        'pantry_remove_confirm_message': '¿Quitar « {name} » de la despensa?',
        'cook_title': '¿Qué puedo cocinar?',
        'cook_instructions_label': 'Indica los ingredientes que tienes en casa:',
        'cook_staples_hint': 'Algunos ingredientes básicos comunes ya están marcados al lado\n(sal, aceite, harina...) — quita los que no tengas.',
        'cook_all_ingredients_label': 'Todos los ingredientes:',
        'cook_add_button': '➕ Añadir →',
        'cook_have_label': 'Lo que tengo:',
        'cook_remove_button': '🗑 Quitar',
        'cook_load_from_pantry_button': '📦 Cargar desde mi despensa',
        'cook_compute_button': '🔍 Ver las recetas viables',
        'cook_open_selected_button': '📖 Consultar la receta seleccionada',
        'cook_pantry_empty_title': 'Información',
        'cook_pantry_empty_message': 'Tu despensa está vacía por el momento. Abre « 📦 Mi despensa » desde la página de inicio para añadir ingredientes.',
        'cook_loaded_title': 'Cargado',
        'cook_loaded_message': '{count} ingrediente(s) añadido(s) desde tu despensa.',
        'cook_add_ingredient_first': 'Añade al menos un ingrediente que tengas.',
        'cook_feasible_header': '✅ Viables con lo que tienes:',
        'cook_insufficient_quantity': '  ⚠️ cantidad insuficiente: {list}',
        'cook_none_feasible': 'Ninguna receta es 100% viable con estos ingredientes.',
        'cook_substitutable_header': '🔄 Viables usando un sustituto:',
        'cook_almost_header': '🟡 Casi (faltan de 1 a 3 ingredientes):',
        'cook_missing_label': '   {name} (falta: {list})',
        'cook_no_results': 'Intenta añadir más ingredientes a tu selección.',
        'cook_select_recipe_from_results': 'Selecciona una receta de la lista de resultados.',
        'cook_select_recipe_row': 'Selecciona una fila correspondiente a una receta.',
        'weekhistory_title': 'Historial de semanas pasadas',
        'weekhistory_heading': '🕘 Historial de semanas pasadas',
        'weekhistory_intro': 'Cada semana en la que guardas la planificación se archiva aquí\nautomáticamente (hasta 26 semanas, unos 6 meses), para evitar\nrepetir dos veces lo mismo con demasiada frecuencia.',
        'weekhistory_reload_button': '♻️ Recargar en la planificación actual',
        'weekhistory_delete_button': '🗑 Eliminar esta semana',
        'weekhistory_no_archived_weeks': 'Ninguna semana archivada.',
        'weekhistory_week_label': 'Semana {week}',
        'weekhistory_saved_on': 'Guardado el {date}\n\n',
        'weekhistory_day_heading': '{day}:\n',
        'weekhistory_slot_line': '   {slot}: {recipe} ({persons} pers.)\n',
        'weekhistory_empty_week': '(Planificación vacía para esta semana.)',
        'weekhistory_select_week_first': 'Selecciona una semana de la lista.',
        'weekhistory_reload_confirm_message': '¿Recargar la planificación de la semana {week} en la planificación actual?\n\nEsto reemplazará las recetas actualmente mostradas (recuerda guardar la planificación en curso antes, si quieres conservarla).',
        'weekhistory_delete_confirm_message': '¿Eliminar definitivamente el archivo de la semana {week}?',
        'weektemplates_title': 'Plantillas de semana',
        'weektemplates_heading': '📋 Plantillas de semana',
        'weektemplates_intro': 'Guarda la planificación actualmente mostrada como plantilla\nreutilizable, para aplicarla con un clic a otra semana\nen lugar de volver a introducirlo todo.',
        'weektemplates_name_label': 'Nombre de la nueva plantilla:',
        'weektemplates_save_button': '💾 Guardar la planificación actual como plantilla',
        'weektemplates_apply_button': '📋 Aplicar esta plantilla',
        'weektemplates_delete_button': '🗑 Eliminar esta plantilla',
        'weektemplates_none_saved': 'Ninguna plantilla guardada por el momento.',
        'weektemplates_error_name_required': 'Por favor, indica un nombre para esta plantilla.',
        'weektemplates_empty_plan': 'La planificación actualmente mostrada está vacía: nada que guardar como plantilla.',
        'weektemplates_saved_message': 'Plantilla « {name} » guardada.',
        'weektemplates_select_template_first': 'Selecciona una plantilla de la lista.',
        'weektemplates_apply_confirm_message': '¿Aplicar la plantilla « {name} » a la planificación actual?\n\nEsto reemplazará las recetas actualmente mostradas (recuerda guardar la planificación en curso antes, si quieres conservarla).',
        'weektemplates_delete_confirm_message': '¿Eliminar definitivamente la plantilla « {name} »?',
        'common_none_option': '-- Ninguna --',
        'weekplan_title': 'Planificación semanal',
        'weekplan_subtitle': 'Vista de calendario: días en columnas, comidas en filas.',
        'weekplan_save_button': '💾 Guardar la planificación',
        'weekplan_clear_button': '🗑 Borrar todo',
        'weekplan_export_ics_button': '📆 Exportar a un calendario (.ics)',
        'weekplan_compute_button': 'Calcular la lista de compras de la semana',
        'weekplan_checklist_button': '☑️ Modo compras',
        'weekplan_empty_list_message': 'Ninguna lista calculada por el momento.\nHaz clic en « Calcular la lista de compras de la semana » arriba,\no carga una lista guardada.',
        'weekplan_total_list_heading': '=== Lista de compras de la semana ===',
        'weekplan_calculate_list_for_export': 'Primero calcula una lista de compras (botón « Calcular la lista de compras de la semana »).',
        'weekplan_invalid_persons_for_slot': 'Número de personas no válido para {day} — {slot}.',
        'weekplan_saved_message': 'La planificación de la semana ha sido guardada.',
        'weekplan_clear_confirm_message': '¿Borrar toda la planificación de la semana?',
        'weekplan_assign_recipe_first': 'Asigna al menos una receta a un espacio de la semana.',
        'weekplan_export_ics_title': 'Exportar la planificación a un calendario',
        'weekplan_ics_export_success_message': 'Planificación exportada:\n{path}\n\nImporta este archivo en Google Calendar, Outlook o Calendario para ver tus comidas repetirse cada semana.',
        'weekplan_assign_or_manual': 'Asigna al menos una receta a un espacio de la semana, o añade un ingrediente manualmente.',
        'weekplan_export_shopping_list_title': 'Guardar la lista de compras',
        'weekplan_shopping_list_title': 'Lista de compras de la semana',
        'weekplan_list_saved_message': 'Lista guardada:\n{path}',
        'weekplan_excel_module_missing': 'La exportación a Excel requiere: pip install openpyxl',
        'weekplan_pdf_module_missing': 'La exportación a PDF requiere: pip install reportlab',
        'weekplan_print_module_missing': 'La impresión requiere: pip install reportlab',
        'weekplan_print_label': 'la lista de compras de la semana',
        'manageing_title': 'Gestionar los ingredientes',
        'manageing_list_label': 'Lista de ingredientes guardados:',
        'manageing_add_button': '➕ Añadir',
        'manageing_edit_button': '✏️ Modificar',
        'manageing_delete_button': '🗑️ Eliminar',
        'manageing_load_defaults_button': '📚 Cargar los ~1000 ingredientes comunes',
        'manageing_spell_check_button': '🔤 Comprobar duplicados / errores tipográficos',
        'manageing_prices_button': '💰 Gestionar los precios (para el coste de las recetas)',
        'manageing_substitutions_button': '🔄 Gestionar las sustituciones',
        'manageing_edit_hint': '"Modificar" permite cambiar el nombre (actualizado\nen todos los lugares donde se usa el ingrediente), sus alérgenos,\nsus valores nutricionales y su precio.',
        'manageing_select_ingredient_first': 'Selecciona un ingrediente de la lista.',
        'manageing_delete_confirm_message': '¿Eliminar « {name} » de la lista de ingredientes?',
        'manageing_delete_usage_warning': '\n\nAtención: se usa en {count} receta(s). Estas recetas conservarán este ingrediente, pero ya no se propondrá en el menú desplegable, a menos que lo vuelvas a añadir.',
        'manageing_missing_file_title': 'Archivo faltante',
        'manageing_missing_file_message': 'No se encuentra el archivo ingredients_par_defaut.json.\nAsegúrate de que esté en la misma carpeta que main.py.',
        'manageing_done_title': 'Completado',
        'manageing_defaults_added_message': '{count} nuevo(s) ingrediente(s) añadido(s) de la lista común.',
        'manageing_defaults_none_added': 'Todos los ingredientes comunes ya estaban presentes.',
        'subedit_title': 'Sustitutos para « {name} »',
        'subedit_heading': '🔄 Sustitutos para « {name} »',
        'subedit_disclaimer': 'Una sustitución es un consejo culinario, no una equivalencia\ngarantizada: el resultado puede variar según la receta.',
        'subedit_remove_button': '🗑 Quitar el sustituto seleccionado',
        'subedit_add_frame_title': 'Añadir un sustituto',
        'subedit_name_label': 'Nombre:',
        'subedit_note_label': 'Nota (opcional):',
        'subedit_add_to_list_button': '➕ Añadir a la lista',
        'subedit_revert_button': '🔄 Volver a la base proporcionada',
        'subedit_cancel_button': 'Cancelar',
        'subedit_no_substitute_yet': 'Ningún sustituto por el momento.',
        'subedit_error_name_required': 'Por favor, indica un nombre de sustituto.',
        'subedit_select_to_remove': 'Selecciona un sustituto para quitar.',
        'subedit_revert_confirm_message': '¿Quitar tu lista personalizada y volver a los sustitutos proporcionados con la aplicación para « {name} »?',
        'managesub_title': 'Gestionar las sustituciones',
        'managesub_heading': '🔄 Sustituciones de ingredientes',
        'managesub_intro': 'Consulta o modifica los sustitutos sugeridos para un ingrediente.\nUna sustitución es un consejo culinario, no una equivalencia garantizada.',
        'managesub_manage_button': '✏️ Gestionar sus sustitutos',
        'managesub_hint': 'Haz doble clic en un ingrediente de la lista para ver o modificar sus\nsustitutos, o escribe un nombre arriba (incluso un ingrediente que aún\nno tenga sustituto conocido) y luego « ✏️ Gestionar sus sustitutos ».',
        'managesub_none_with_substitute': 'Ningún ingrediente con sustituto por el momento.',
        'managesub_substitute_count': '{name} ({count} sustituto{plural})',
        'managesub_error_ingredient_required': 'Por favor, indica un ingrediente.',
        'managesub_unknown_ingredient_message': '« {name} » no corresponde a ningún ingrediente registrado.\nElige uno de la lista desplegable, o créalo primero desde « 🥕 Gestionar los ingredientes ».',
        'ingprices_title': 'Gestionar los precios de los ingredientes',
        'ingprices_heading': '💰 Precios de los ingredientes',
        'ingprices_intro': 'Indica un precio para los ingredientes que te\ninteresen — no es necesario hacerlo con todos. El coste\nde una receta se estima a partir de estos precios.',
        'ingprices_price_label': 'Precio (€):',
        'ingprices_for_one_label': 'por 1',
        'ingprices_save_button': '💾 Guardar el precio',
        'ingprices_clear_button': '🗑 Borrar el precio',
        'ingprices_units_note': 'kg ↔ recetas en Gr   ·   L ↔ recetas en cl   ·   los precios\npor unidad/cucharada se aplican tal cual.',
        'ingprices_no_price_set': '  —  (precio no indicado)',
        'ingprices_price_suffix': '  —  {price} € / {unit}',
        'ingprices_error_invalid_price': 'Introduce un precio válido (número positivo).',
        'ingprices_saved_message': 'Precio guardado para « {name} ».',
        'ingedit_title_edit': 'Modificar un ingrediente',
        'ingedit_title_new': 'Nuevo ingrediente',
        'ingedit_heading_edit': '✏️ Modificar el ingrediente',
        'ingedit_heading_new': '➕ Nuevo ingrediente',
        'ingedit_name_label': 'Nombre:',
        'ingedit_allergens_label': 'Alérgenos presentes:',
        'ingedit_nutrition_label': 'Valores nutricionales (por 100 g / 100 ml):',
        'ingedit_nutri_kcal': 'Calorías (kcal)',
        'ingedit_nutri_protein': 'Proteínas (g)',
        'ingedit_nutri_carbs': 'Carbohidratos (g)',
        'ingedit_nutri_fat': 'Grasas (g)',
        'ingedit_nutrition_hint': 'Déjalo vacío si no conoces estos valores.',
        'ingedit_price_label': 'Precio:',
        'ingedit_save_button': '💾 Guardar',
        'ingedit_delete_button': '🗑️ Eliminar este ingrediente',
        'ingedit_error_invalid_field': '« {field} » debe ser un número positivo (o estar vacío).',
        'ingedit_error_name_required': 'Por favor, indica un nombre de ingrediente.',
        'ingedit_error_already_exists': 'El ingrediente « {name} » ya existe.',
        'ingedit_error_plural_duplicate': '« {name} » es solo una variante singular/plural del ingrediente ya existente « {existing} ». Para evitar duplicados en la lista, usa directamente « {existing} ».',
        'ingedit_nutri_field_kcal': 'Calorías',
        'ingedit_nutri_field_protein': 'Proteínas',
        'ingedit_nutri_field_carbs': 'Carbohidratos',
        'ingedit_nutri_field_fat': 'Grasas',
        'ingedit_error_invalid_price': 'El precio debe ser un número positivo (o estar vacío).',
        'ingedit_saved_message': '« {name} » ha sido guardado.',
        'spellcheck_title': 'Comprobación ortográfica de ingredientes',
        'spellcheck_heading': 'Pares de ingredientes que se parecen en un 90% o más\n(probables duplicados o errores tipográficos):',
        'spellcheck_multi_select_hint': 'Selección múltiple posible (Ctrl+clic o Mayús+clic) para\nfusionar varios pares a la vez.',
        'spellcheck_merge_button': '🔗 Fusionar la selección',
        'spellcheck_not_duplicate_button': '✕ No es un duplicado',
        'spellcheck_rerun_button': '🔄 Repetir el análisis',
        'spellcheck_footer_hint': 'Para un solo par, se te preguntará cuál de las dos\ngrafías conservar. Para varios pares a la vez, el ingrediente\nmenos usado en tus recetas se fusiona automáticamente\ncon el usado en más recetas.\n« No es un duplicado » retira definitivamente el o los\npares seleccionados de este análisis, ahora y en el futuro.',
        'spellcheck_none_found': 'No se detectó ningún duplicado probable. 🎉',
        'spellcheck_pair_line': '{a}   ↔   {b}     ({percent}% similares)',
        'spellcheck_select_pair_first': 'Selecciona al menos un par de la lista.',
        'spellcheck_dismissed_message': '{count} par(es) marcado(s) como no duplicados. Ya no se propondrán en los próximos análisis.',
        'spellcheck_merge_dialog_title': 'Fusionar',
        'spellcheck_merge_dialog_message': '¿Fusionar « {a} » y « {b} »?\n\nSí = renombrar todo como « {a} »\nNo = renombrar todo como « {b} »\nCancelar = no hacer nada',
        'spellcheck_merged_title': 'Fusionado',
        'spellcheck_merged_one_message': '« {removed} » se ha fusionado con « {kept} ».',
        'spellcheck_merge_multi_confirm': '¿Fusionar automáticamente estos {count} pares?\n\nPara cada par, el ingrediente menos usado en tus recetas se fusionará con el usado en más recetas (el primero por orden alfabético en caso de empate).',
        'spellcheck_merged_multi_message': '{count} par(es) fusionado(s).',
        'compare_title': 'Comparar dos recetas',
        'compare_recipe_a_label': 'Receta A:',
        'compare_recipe_b_label': 'Receta B:',
        'compare_button': '⚖️ Comparar',
        'compare_choose_each_list': 'Elige una receta en cada lista.',
        'compare_field_category': 'Categoría:',
        'compare_field_favorite': 'Favorito:',
        'compare_yes': '⭐ Sí',
        'compare_no': 'No',
        'compare_field_rating': 'Valoración:',
        'compare_field_difficulty': 'Dificultad:',
        'compare_field_prep': 'Preparación:',
        'compare_field_cook': 'Cocción:',
        'compare_field_total_time': 'Tiempo total:',
        'compare_field_cooked': 'Cocinada:',
        'compare_times_suffix': '{count} veces',
        'compare_field_cost': 'Coste estimado:',
        'compare_field_nutrition': 'Nutrición (kcal):',
        'compare_field_ingredient_count': 'N.º ingredientes:',
        'compare_common_ingredients': '🟰 Comunes ({count})',
        'compare_only_a': '🅰️ Solo en « {name} » ({count})',
        'compare_only_b': '🅱️ Solo en « {name} » ({count})',
        'compare_none': 'Ninguno',
        'stats_title': 'Estadísticas',
        'stats_heading': '=== Estadísticas ===\n\n',
        'stats_total_recipes': 'Número total de recetas: {count}\n\n',
        'stats_by_category': 'Distribución por categoría:\n',
        'stats_category_line': '  - {category}: {count}\n',
        'stats_by_difficulty': 'Distribución por dificultad:\n',
        'stats_difficulty_line': '  - {difficulty}: {count}\n',
        'stats_difficulty_unspecified': 'No especificada',
        'stats_favorites_count': 'Recetas favoritas: {count}\n\n',
        'stats_avg_rating': 'Valoración media (recetas valoradas): {avg} / 5 ({count} receta(s) valorada(s))\n\n',
        'stats_no_rated_recipe': 'Valoración media: ninguna receta valorada por el momento.\n\n',
        'stats_five_star_heading': 'Receta(s) valorada(s) con 5 estrellas:\n',
        'stats_recipe_line': '  - {name}\n',
        'stats_most_cooked_heading': 'Recetas más cocinadas:\n',
        'stats_cooked_line': '  - {name}: {count} veces\n',
        'stats_none_cooked_yet': '  Ninguna receta marcada como cocinada por el momento.\n  (botón « 🍳 ¡Cociné esto! » en « Ver una receta concreta »)\n',
        'stats_most_used_tags_heading': 'Etiquetas más usadas:\n',
        'stats_tag_line': '  - {tag}: {count}\n',
        'stats_never_cooked_heading': '🕸️ Recetas nunca cocinadas:\n',
        'stats_and_others': '  ... y {count} más\n',
        'stats_all_cooked': '  Todas tus recetas ya se han cocinado al menos una vez. 👏\n',
        'stats_stale_heading': '🕰️ No cocinadas desde hace más de {days} días:\n',
        'stats_stale_line': '  - {name} (hace {days} días)\n',
        'stats_no_stale_recipe': '  Ninguna receta en este caso por el momento.\n',
        'stats_avg_cost_heading': '💰 Coste medio por persona:\n',
        'stats_avg_cost_line': '  {avg} € de media, sobre {count} receta(s) con al menos un precio conocido ({without_price} sin precio indicado)\n',
        'stats_no_priced_recipe': '  Ninguna receta con precio indicado por el momento.\n  (ver « 💰 Gestionar los precios » en « Gestionar los ingredientes »)\n',
        'stats_avg_kcal_heading': '🥗 Calorías medias por persona:\n',
        'stats_avg_kcal_line': '  {avg} kcal de media, sobre {count} receta(s) con ingredientes reconocidos en la base nutricional\n',
        'stats_no_recognized_recipe': '  Ninguna receta con ingredientes reconocidos por el momento.\n',
        'stats_monthly_chart_title': '📈 Recetas cocinadas por mes (últimos 12 meses)',
        'stats_heatmap_title': '🗓️ Calendario de días cocinados (últimos 12 meses)',
        'stats_heatmap_legend': 'Menos ⬜ 🟨 🟧 🟥 Más',
        'stats_day_labels': 'L,M,X,J,V,S,D',
        'stats_month_labels_short': 'Ene,Feb,Mar,Abr,May,Jun,Jul,Ago,Sep,Oct,Nov,Dic',
        'stats_month_labels_lower': 'ene,feb,mar,abr,may,jun,jul,ago,sep,oct,nov,dic',
        'recipepdf_category_persons': 'Categoría: {cat}    Para {persons} persona(s)',
        'recipepdf_rating': 'Valoración: {stars}',
        'recipepdf_prep': 'Preparación: {time} min',
        'recipepdf_cook': 'Cocción: {time} min',
        'recipepdf_difficulty': 'Dificultad: {value}',
        'recipepdf_allergens': '⚠ Alérgenos: {list}',
        'recipepdf_ingredients_heading': 'Ingredientes:',
        'recipepdf_cost': 'Coste estimado: {cost} €{partial}',
        'recipepdf_partial_suffix': ' (parcial, {known}/{total})',
        'recipepdf_nutrition': 'Nutrición estimada{partial}: {kcal} kcal · {protein}g prot. · {carbs}g carb. · {fat}g grasa',
        'recipepdf_description_heading': 'Descripción:',
        'recipepdf_notes_heading': 'Notas personales:',
        'cookbookpdf_page_number': 'Página {current} / {total}',
        'cookbookpdf_generated_on': 'Generado el {date}',
        'cookbookpdf_summary_heading': 'Índice',
        'cookbookpdf_summary_line': '- [{cat}] {name}',
        'cookbookexport_title': 'Exportar el libro de recetas',
        'cookbookexport_heading': '📖 Exportar el libro de recetas',
        'cookbookexport_intro': 'Selecciona las recetas a incluir en un solo PDF,\nestilo libro de cocina.',
        'cookbookexport_filter_label': 'Filtrar por categoría:',
        'cookbookexport_check_all_button': 'Marcar todo',
        'cookbookexport_uncheck_all_button': 'Desmarcar todo',
        'cookbookexport_generate_button': '📄 Generar el PDF del libro',
        'cookbookexport_error_select_recipe': 'Selecciona al menos una receta.',
        'cookbookexport_save_dialog_title': 'Guardar el libro de recetas',
        'cookbookexport_saved_message': 'Libro de recetas guardado:\n{path}',
        'importexport_title': 'Importar / Exportar datos',
        'importexport_heading': 'Respaldar o transferir tus datos',
        'importexport_export_intro': 'La exportación crea un archivo .zip que contiene absolutamente\ntodos tus datos: recetas, fotos, ingredientes personalizados,\nprecios, sustitutos, despensa, planificación y su historial,\nmenús, listas de compras guardadas, papelera y\nconfiguración — para respaldar o transferir todo a\notro ordenador en un solo archivo.',
        'importexport_export_button': '📤 Exportar todos mis datos (.zip)',
        'importexport_import_intro': 'La importación lee un archivo .zip exportado previamente.\n"Fusionar" añade las recetas/fotos duplicadas con un\nnuevo nombre en lugar de perderlas, y completa el resto\n(despensa, menús, listas...) sin eliminar nada.\n"Reemplazar" sobrescribe todo, incluida la configuración y la\nplanificación en curso.',
        'importexport_import_button': '📥 Importar datos (.zip)',
        'importexport_auto_backups_heading': '🗄️ Copias de seguridad automáticas',
        'importexport_auto_backups_intro': 'Se crea una copia de seguridad automáticamente al iniciar la\naplicación (como máximo una cada {hours}h), y las {retention} más\nrecientes se conservan aquí.',
        'importexport_backup_now_button': '💾 Hacer copia de seguridad ahora',
        'importexport_restore_selected_button': '♻️ Restaurar la selección',
        'importexport_cloud_heading': '☁️ Copia de seguridad automática en la nube',
        'importexport_cloud_intro': 'Elige una carpeta sincronizada por un cliente ya\ninstalado en este PC (Google Drive, OneDrive, Dropbox...).\nCada copia de seguridad automática también se copiará allí, y este\ncliente se encargará de enviarla a la nube por sí solo.',
        'importexport_choose_cloud_button': '📁 Elegir una carpeta en la nube',
        'importexport_disable_button': '🚫 Desactivar',
        'importexport_cloud_enabled': '✅ Activado: {folder}',
        'importexport_cloud_not_configured': 'No configurado por el momento.',
        'importexport_choose_folder_title': 'Elegir una carpeta sincronizada (Google Drive, OneDrive, Dropbox...)',
        'importexport_cloud_configured_title': 'Carpeta configurada',
        'importexport_cloud_configured_message': 'Carpeta en la nube configurada:\n{folder}\n\n¿Quieres copiar una copia de seguridad ahora mismo?',
        'importexport_disabled_title': 'Desactivado',
        'importexport_disabled_message': 'La copia de seguridad automática en la nube está desactivada.',
        'importexport_backup_date_line': '{date}   ({size} KB)',
        'importexport_no_backups': 'Ninguna copia de seguridad automática por el momento.',
        'importexport_backup_failed': 'La copia de seguridad falló:\n{error}',
        'importexport_backup_created_title': 'Copia de seguridad creada',
        'importexport_backup_created_message': 'Se ha creado una nueva copia de seguridad automática.',
        'importexport_select_backup_first': 'Selecciona una copia de seguridad de la lista.',
        'importexport_restore_mode_title': 'Modo de restauración',
        'importexport_restore_mode_message': '¿Cómo restaurar esta copia de seguridad?\n\nSí = Fusionar (añadir a los datos actuales, sin eliminar nada)\nNo = Reemplazar completamente los datos actuales\nCancelar = no hacer nada',
        'importexport_restore_failed': 'La restauración falló:\n{error}',
        'importexport_restore_done_title': 'Restauración completada',
        'importexport_restore_done_message': 'Los datos se han restaurado correctamente.',
        'importexport_export_data_title': 'Exportar mis datos',
        'importexport_export_data_success': 'Tus datos han sido exportados a:\n{path}',
        'importexport_choose_archive_title': 'Elegir un archivo para importar',
        'importexport_import_mode_title': 'Modo de importación',
        'importexport_import_mode_message': '¿Cómo importar estos datos?\n\nSí = Fusionar (añadir a los datos actuales, sin eliminar nada)\nNo = Reemplazar completamente los datos actuales\nCancelar = no hacer nada',
        'importexport_import_failed': 'La importación falló:\n{error}',
        'importexport_import_done_title': 'Importación completada',
        'importexport_import_done_message': 'Los datos se han importado correctamente.',
        'checklist_instruction': 'Marca cada artículo a medida que hagas tus compras.',
        'checklist_check_all_button': '☑️ Marcar todo',
        'checklist_uncheck_all_button': '⬜ Desmarcar todo',
        'checklist_progress_label': '{done} / {total} artículo(s) marcado(s)',
        'exportformat_title': 'Elegir un formato de exportación',
        'exportformat_heading': '📤 Exportar la lista de compras',
        'exportformat_choose_label': 'Elige el formato de exportación deseado:',
        'exportformat_txt_button': '📝 Exportar como texto (.txt)',
        'exportformat_excel_button': '📊 Exportar como Excel (.xlsx)',
        'exportformat_pdf_button': '📄 Exportar como PDF (.pdf)',
        'exportformat_cancel_button': 'Cancelar',
        'menumanager_title': 'Mis menús',
        'menumanager_list_label': 'Mis menús guardados:',
        'menumanager_new_button': '➕ Nuevo menú',
        'menumanager_recipe_count': '{name} ({count} receta(s))',
        'menumanager_select_menu_first': 'Selecciona un menú de la lista.',
        'menumanager_delete_confirm': '¿Eliminar el menú « {name} »?',
        'menuform_title_edit': 'Modificar el menú',
        'menuform_title_new': 'Nuevo menú',
        'menuform_name_label': 'Nombre del menú:',
        'menuform_add_recipe_label': 'Añadir una receta al menú:',
        'menuform_persons_short_label': 'pers.:',
        'menuform_add_button': '+ Añadir',
        'menuform_recipes_label': 'Recetas del menú:',
        'menuform_remove_button': '🗑 Quitar del menú',
        'menuform_save_button': '💾 Guardar el menú',
        'menuform_compute_button': 'Calcular la lista de compras del menú',
        'menuform_empty_list_message': 'Ninguna lista calculada por el momento.\nHaz clic en « Calcular la lista de compras del menú » arriba,\no carga una lista guardada.',
        'menuform_total_list_heading': '=== Lista de compras del menú ===',
        'menuform_item_row_label': '[{cat}] {name} ({persons} pers.)',
        'menuform_select_recipe_to_remove': 'Selecciona una receta del menú para quitar.',
        'menuform_error_name_required': 'Por favor, indica un nombre de menú.',
        'menuform_error_no_recipe': 'Añade al menos una receta al menú.',
        'menuform_saved_message': 'El menú « {name} » ha sido guardado.',
        'menuform_calculate_list_for_export': 'Primero calcula una lista de compras (botón « Calcular la lista de compras del menú »).',
        'menuform_add_recipe_or_manual': 'Añade al menos una receta al menú, o añade un ingrediente manualmente.',
        'menuform_shopping_list_title': 'Menú: {name}',
        'menuform_print_label': 'el menú « {name} »',
        'importurl_title': 'Importar una receta desde un enlace',
        'importurl_heading': '🌐 Importar una receta desde un enlace',
        'importurl_intro': 'Pega la dirección (URL) de una página de receta. Esto funciona\ncon la mayoría de los grandes sitios de cocina (que usan un\nformato de datos estándar). Se requiere conexión a internet.',
        'importurl_fetch_button': '🌐 Obtener la receta',
        'importurl_after_import_note': 'Después de importar, revisa y completa la receta si es necesario\n(la detección de cantidades y unidades no siempre es perfecta).',
        'importurl_paste_url_first': 'Primero pega una dirección de internet (URL).',
        'importurl_fetching': 'Obteniendo...',
        'importurl_failed_title': 'Error al importar',
        'importphoto_title': 'Importar una receta desde una foto',
        'importphoto_heading': '📷 Importar una receta desde una foto',
        'importphoto_intro': 'Toma una foto (o escanea) una receta manuscrita o una\npágina de un libro de cocina, luego elige la imagen aquí. El texto\nse extrae automáticamente, pero aún debes revisarlo y organizarlo\ntú mismo (a diferencia de la importación desde un enlace, una foto\nno tiene una estructura de ingredientes/pasos que se pueda adivinar).',
        'importphoto_module_warning': "⚠ Esta función requiere el módulo 'pytesseract'\nY el programa Tesseract OCR instalado por separado en este PC.\nConsulta el LÉEME para las instrucciones de instalación.",
        'importphoto_no_photo_chosen': 'Ninguna foto elegida',
        'importphoto_choose_button': '📁 Elegir una foto',
        'importphoto_extract_button': '🔍 Extraer el texto',
        'importphoto_extracted_text_label': 'Texto extraído (editable):',
        'importphoto_create_button': '➡️ Crear la receta con este texto',
        'importphoto_choose_photo_title': 'Elegir una foto de receta',
        'importphoto_choose_first': 'Primero elige una foto.',
        'importphoto_ocr_module_missing': "Esta función requiere el módulo 'pytesseract'\n(pip install pytesseract) Y el programa Tesseract OCR\ninstalado por separado en este PC. Consulta el LÉEME.",
        'importphoto_extraction_failed_title': 'Error de extracción',
        'importphoto_extraction_failed_message': 'El reconocimiento de texto falló. Comprueba que Tesseract OCR esté correctamente instalado en este PC y sea accesible.\n\nDetalle: {error}',
        'importphoto_no_text_extracted': 'No se pudo extraer texto de esta foto. Prueba con una imagen más nítida, mejor encuadrada o mejor iluminada.',
        'importphoto_no_text_title': 'Sin texto',
        'importphoto_no_text_confirm': 'No se ha extraído ni introducido ningún texto. ¿Crear de todos modos una receta vacía (solo con la foto)?',
        'trash_title': 'Papelera',
        'trash_heading': '🗑️ Recetas eliminadas',
        'trash_intro': 'Las fotos de las recetas de la papelera se conservan\nhasta su eliminación definitiva.',
        'trash_restore_button': '♻️ Restaurar',
        'trash_delete_forever_button': '🗑️ Eliminar definitivamente',
        'trash_empty_button': '🧹 Vaciar la papelera',
        'trash_unnamed_recipe': '(sin nombre)',
        'trash_unknown_date': 'fecha desconocida',
        'trash_entry_line': '{name}  —  eliminada el {date}',
        'trash_is_empty': 'La papelera está vacía.',
        'trash_select_recipe_first': 'Selecciona una receta en la papelera.',
        'trash_restored_suffix': '{name} (restaurada)',
        'trash_restored_title': 'Restaurada',
        'trash_restored_message': '« {name} » ha sido restaurada.',
        'trash_delete_forever_confirm': '¿Eliminar definitivamente « {name} »?\n\nEsta acción es irreversible.',
        'trash_deleted_title': 'Eliminada',
        'trash_deleted_message': 'La receta ha sido eliminada definitivamente.',
        'trash_already_empty': 'La papelera ya está vacía.',
        'trash_empty_confirm': '¿Eliminar definitivamente las {count} receta(s) de la papelera?\n\nEsta acción es irreversible.',
        'trash_emptied_title': 'Papelera vaciada',
        'trash_emptied_message': 'La papelera ha sido vaciada.',
        'cookingmode_title': 'Modo cocina — {name}',
        'cookingmode_close_button': '✕ Cerrar (Esc)',
        'cookingmode_cooked_button': '🍳 ¡Cociné esto!',
        'cookingmode_fullscreen_hint': 'F11: pantalla completa',
        'cookingmode_persons_suffix': '{persons} pers.',
        'cookingmode_speech_button': '🔊 Leer en voz alta',
        'cookingmode_speech_stop_button': '⏹ Detener la lectura',
        'cookingmode_volume_percent': '{percent}%',
        'cookingmode_tts_module_missing': "La lectura en voz alta requiere el módulo 'pyttsx3'.\nInstálalo con: pip install pyttsx3",
        'cookingmode_no_description_to_read': 'Esta receta no tiene descripción para leer (el campo de descripción está vacío).',
        'cookingmode_ingredients_heading': 'Ingredientes',
        'cookingmode_prep_label': 'Preparación: {time} min',
        'cookingmode_cook_label': 'Cocción: {time} min',
        'cookingmode_difficulty_label': 'Dificultad: {value}',
        'cookingmode_preparation_heading': 'Preparación',
        'cookingmode_personal_notes_heading': 'Notas personales',
        'ingsearch_title': 'Búsqueda por ingrediente',
        'ingsearch_question_label': '¿Qué ingrediente estás buscando?',
        'ingsearch_view_recipes_button': '🔍 Ver las recetas que lo usan',
        'ingsearch_view_selected_button': '📖 Consultar la receta seleccionada',
        'ingsearch_no_recipe_uses': 'Ninguna receta usa « {name} » por el momento.',
        'ingsearch_recipes_using': 'Recetas que usan « {name} » ({count}):',
        'ingsearch_result_line': '{star}[{cat}] {name} ({qty}{unit} por 1 persona)',
        'ingsearch_select_result_first': 'Selecciona una receta de la lista de resultados.',
        'timerrow_minutes_label': 'Min:',
        'timerrow_seconds_label': 'Seg:',
        'timerrow_error_invalid_duration': 'Duración no válida.',
        'timerrow_set_duration_first': 'Establece una duración antes de iniciar.',
        'cooklogentry_title': '📔 Añadir al diario de cocina',
        'cooklogentry_heading': '🍳 « {name} »',
        'cooklogentry_intro': '¿Qué tal estuvo? Una nota y/o una foto\n(opcional, también puedes omitir esto).',
        'cooklogentry_no_photo_chosen': 'Ninguna foto elegida',
        'cooklogentry_choose_photo_button': '📷 Elegir una foto',
        'cooklogentry_skip_button': 'Omitir',
        'cooklogentry_choose_photo_title': 'Elegir una foto',
        'cooklog_title': '📔 Diario de cocina — {name}',
        'cooklog_heading': '📔 {name}',
        'cooklog_times_cooked': 'Cocinada {count} veces en total',
        'cooklog_no_entry': 'Ninguna nota guardada por el momento.\nUsa « 🍳 ¡Cociné esto! » para añadir una.',
        'cooklog_no_note': '(sin nota)',
        'timers_title': '⏲️ Temporizadores',
        'timers_intro': 'Ajusta cada temporizador y luego pulsa ▶️ para iniciarlo.\nAl terminar, la fila parpadea en rojo con una alerta sonora.',
        'timers_add_button': '➕ Añadir un temporizador',
        'qrcode_title': 'Código QR — {name}',
        'qrcode_intro': 'Escanea con la cámara o una aplicación de lectura de\ncódigos QR para ver el nombre y los ingredientes.',
        'qrcode_save_button': '💾 Guardar como imagen (PNG)',
        'qrcode_truncated_warning': '⚠️ La receta es larga: el código QR contiene un\nresumen truncado (solo nombre + ingredientes).',
        'qrcode_encoded_ingredients_heading': 'Ingredientes ({persons} pers.):',
        'qrcode_save_dialog_title': 'Guardar el código QR',
        'qrcode_save_failed': 'El guardado falló:\n{error}',
        'qrcode_saved_message': 'Código QR guardado:\n{path}',
        'unitconv_title': 'Conversor de unidades',
        'unitconv_heading': '🔄 Conversor de unidades',
        'unitconv_intro': 'Conversión aproximada basada en la densidad del agua para\nlas unidades de volumen (ml, cl, L, taza, cucharas): fiable para\nlíquidos, aproximada para sólidos como la harina\no el azúcar, cuya densidad real difiere un poco.',
        'unitconv_quantity_label': 'Cantidad:',
        'unitconv_from_label': 'De:',
        'unitconv_to_label': 'A:',
        'unitconv_convert_button': 'Convertir',
        'unitconv_error_invalid_quantity': 'Cantidad no válida.',
        'unitconv_result': '{quantity} {from_unit} ≈ {result} {to_unit}',
        'unitconv_gram': 'Gramo (g)',
        'unitconv_kilogram': 'Kilogramo (kg)',
        'unitconv_ounce': 'Onza (oz)',
        'unitconv_pound': 'Libra (lb)',
        'unitconv_milliliter': 'Mililitro (ml)',
        'unitconv_centiliter': 'Centilitro (cl)',
        'unitconv_liter': 'Litro (L)',
        'unitconv_teaspoon': 'Cucharadita (5 ml)',
        'unitconv_tablespoon': 'Cucharada (15 ml)',
        'unitconv_cup': 'Taza EE. UU. (240 ml)',
        'disclaimer_title': 'Cláusula de responsabilidad',
        'disclaimer_heading': '⚠ Cláusula de responsabilidad',
        'disclaimer_intro': 'Por favor, lee este texto antes de usar la aplicación.',
        'disclaimer_checkbox': 'He leído y acepto las condiciones anteriores',
        'disclaimer_continue_button': 'Continuar',
        'disclaimer_quit_button': 'Salir de la aplicación',
        'disclaimer_text': 'ARTÍCULO 1 – EXCLUSIÓN Y LIMITACIÓN DE RESPONSABILIDAD\n\n1.1. Alertas médicas y gestión de alérgenos\n\nLa Aplicación ofrece una función que permite al Usuario indicar, modificar y configurar sus propios criterios de alergias y alérgenos. El Usuario reconoce expresamente que:\n\n• La exactitud y la actualización de esta información es responsabilidad exclusiva del Usuario.\n• La Aplicación es una herramienta informática de ayuda para consultar recetas y no sustituye en ningún caso un consejo médico, un diagnóstico o el control humano de los ingredientes.\n• El Editor no podrá ser considerado responsable en caso de entrada incorrecta, omisión, configuración errónea por parte del Usuario, o reacción alérgica (intolerancia, choque anafiláctico, etc.) ocurrida después de consumir un plato. Corresponde al Usuario verificar sistemáticamente las etiquetas y la composición real de cada ingrediente físico antes de cualquier preparación o ingestión.\n\n1.2. Suministro «tal cual» y gratuidad\n\nLa Aplicación se pone a disposición del Usuario de forma completamente gratuita. Se proporciona «tal cual» y «según su disponibilidad», sin garantía alguna de ausencia de errores, fallos informáticos o interrupciones. El Editor no garantiza que las funciones de la Aplicación satisfagan las necesidades específicas del Usuario.\n\n1.3. Daños materiales e inmateriales\n\nEl Editor rechaza toda responsabilidad por los daños directos o indirectos causados al Usuario o a terceros. En particular, el Editor no podrá ser demandado por:\n\n• Una avería, sobrecalentamiento, mal funcionamiento o deterioro del equipo informático o del smartphone del Usuario al usar la Aplicación.\n• Una pérdida de datos informáticos, alteración de archivos o pirateo del sistema del Usuario.\n\nDebido a la gratuidad del servicio, si la responsabilidad del Editor fuera declarada por un tribunal, el importe de los daños y perjuicios quedaría expresamente limitado a la suma de cero euros (0 €).',
        'addmanual_title': 'Añadir ingredientes a la lista de compras',
        'addmanual_heading': '➕ Añadir ingredientes a la lista de compras',
        'addmanual_intro': 'Añade todos los ingredientes que quieras a la lista\nde espera de abajo, y luego confírmalos todos a la vez.',
        'addmanual_new_ingredient_button': '🥕 Nuevo ingrediente',
        'addmanual_add_to_list_button': '➕ Añadir a la lista',
        'addmanual_staged_label': 'Ingredientes pendientes de confirmación:',
        'addmanual_remove_staged_button': '🗑 Quitar de la lista de espera',
        'addmanual_confirm_all_button': '✅ Confirmar todos estos ingredientes',
        'addmanual_close_button': 'Cerrar',
        'addmanual_select_staged_first': 'Selecciona un ingrediente de la lista de espera.',
        'addmanual_add_staged_first': 'Añade al menos un ingrediente a la lista de espera antes de confirmar.',
        'addmanual_confirmed_message': '{count} ingrediente(s) añadido(s) a la lista de compras.',
        'shoppingexport_generated_on': 'Generada el {date}',
        'shoppingexport_selected_recipes': 'Recetas seleccionadas:',
        'shoppingexport_excel_sheet_recipes': 'Recetas',
        'shoppingexport_excel_col_recipe': 'Receta',
        'shoppingexport_excel_col_persons': 'Número de personas',
        'shoppingexport_excel_sheet_ingredients': 'Ingredientes',
        'shoppingexport_excel_col_rayon': 'Sección',
        'shoppingexport_excel_col_ingredient': 'Ingrediente',
        'shoppingexport_excel_col_total_qty': 'Cantidad total',
        'shoppingexport_excel_col_unit': 'Unidad',
        'savedlists_title': 'Listas de compras guardadas',
        'savedlists_heading': '📂 Listas de compras guardadas',
        'savedlists_load_button': '📂 Cargar',
        'savedlists_delete_button': '🗑 Eliminar',
        'savedlists_none_saved': 'Ninguna lista guardada por el momento.',
        'savedlists_entry_line': '{name} — {count} ingrediente(s) — {date}',
        'savedlists_select_list_first': 'Selecciona una lista de la lista.',
        'savedlists_delete_confirm': '¿Eliminar definitivamente la lista « {name} »?',
        'quicksearch_title': 'Búsqueda rápida',
        'quicksearch_heading': '🔍 Búsqueda rápida de recetas',
        'quicksearch_no_results': 'No se encontró ninguna receta.',
        'quicksearch_footer_hint': 'Intro para abrir, Esc para cerrar.',
        'allrecipes_title': 'Todas las recetas - Lista de compras',
        'allrecipes_select_label': 'Selecciona las recetas y el número de personas:',
        'allrecipes_ingredient_filter_title': 'Filtrar por ingrediente',
        'allrecipes_persons_count_label': 'N.º personas:',
        'allrecipes_add_to_cart_button': '🛒 Añadir a la compra',
        'allrecipes_checklist_mode_button': '☑️ Modo compras (marcar a medida)',
        'allrecipes_clear_list_button': '🗑 Vaciar la lista de compras',
        'allrecipes_export_button': '📤 Exportar',
        'allrecipes_print_button': '🖨️ Imprimir',
        'allrecipes_add_manual_ingredient_button': '➕ Añadir un ingrediente a la lista de compras',
        'allrecipes_save_list_button': '💾 Guardar esta lista para más tarde',
        'allrecipes_load_list_button': '📂 Cargar una lista guardada',
        'allrecipes_invalid_persons': 'Número de personas no válido para « {name} ».',
        'allrecipes_empty_list_message': 'Tu lista de compras está vacía por el momento.\nHaz clic en « 🛒 Añadir a la compra » junto a una receta,\nañade un ingrediente manualmente, o carga una lista guardada.',
        'allrecipes_total_list_heading': '=== Lista de compras total ===',
        'allrecipes_manual_items_note': '({count} ingrediente(s) añadido(s) manualmente incluido(s))',
        'allrecipes_invalid_quantity': 'Cantidad no válida.',
        'allrecipes_calculate_list_first': 'Primero calcula una lista de compras antes de guardarla.',
        'allrecipes_save_list_dialog_title': 'Guardar la lista',
        'allrecipes_save_list_dialog_prompt': 'Nombre para esta lista:',
        'allrecipes_list_saved_title': 'Guardado',
        'allrecipes_list_saved_message': 'Lista « {name} » guardada para más tarde.',
        'allrecipes_empty_list_for_export': 'La lista de compras está vacía. Añade al menos una receta (botón « 🛒 Añadir a la compra ») o un ingrediente manual.',
        'allrecipes_export_txt_title': 'Guardar la lista de compras como texto',
        'allrecipes_export_excel_title': 'Guardar la lista de compras como Excel',
        'allrecipes_export_pdf_title': 'Guardar la lista de compras como PDF',
        'allrecipes_export_saved_message': 'Lista de compras guardada:\n{path}',
        'allrecipes_excel_module_missing': "La exportación a Excel requiere el módulo 'openpyxl'.\nInstálalo con: pip install openpyxl",
        'allrecipes_pdf_module_missing': "La exportación a PDF requiere el módulo 'reportlab'.\nInstálalo con: pip install reportlab",
        'allrecipes_print_module_missing': "La impresión requiere el módulo 'reportlab' para generar el diseño.\nInstálalo con: pip install reportlab",
        'allrecipes_print_label': 'la lista de compras',
        'allrecipes_shopping_list_title': 'Lista de compras',
        'allrecipes_close_confirm_title': '¿Cerrar la lista de compras?',
        'allrecipes_close_confirm_message': 'La lista de compras mostrada no está guardada: se perderá definitivamente si cierras esta ventana ahora.\n\nConsejo: usa « 💾 Guardar esta lista para más tarde » antes de cerrar si quieres conservarla.\n\n¿Cerrar de todos modos?',
        'managerecipes_title': 'Modificar / Eliminar una receta',
        'managerecipes_select_label': 'Selecciona una receta:',
        'managerecipes_filter_favorites': '⭐ Solo favoritos',
        'managerecipes_filter_quick': '⏱️ Solo recetas rápidas (≤ 30 min)',
        'managerecipes_filter_vegetarian': '🥗 Solo recetas vegetarianas',
        'managerecipes_filter_wishlist': '💭 Solo lista de deseos',
        'managerecipes_remove_filter_button': '✕ Quitar el filtro',
        'managerecipes_search_label': '🔍 Buscar:',
        'managerecipes_sort_label': 'Ordenar por:',
        'managerecipes_category_label': 'Categoría:',
        'managerecipes_edit_button': '✏️ Modificar',
        'managerecipes_duplicate_button': '📋 Duplicar',
        'managerecipes_delete_button': '🗑️ Eliminar',
        'managerecipes_select_recipe_first': 'Selecciona una receta de la lista.',
        'managerecipes_duplicate_suffix': '(copia)',
        'managerecipes_duplicated_title': 'Duplicada',
        'managerecipes_duplicated_message': '« {original} » se ha duplicado con el nombre « {new} ».',
        'managerecipes_delete_confirm_message': '¿Enviar la receta « {name} » a la papelera?\n\nPodrás restaurarla más tarde desde el botón « 🗑️ Papelera ».',
        'managerecipes_deleted_title': 'Enviada a la papelera',
        'managerecipes_deleted_message': 'La receta se ha movido a la papelera.',
        'onerecipe_window_title': 'Ver una receta',
        'onerecipe_choose_recipe_label': 'Elige una receta:',
        'onerecipe_search_label': '🔍 Buscar:',
        'onerecipe_sort_label': 'Ordenar:',
        'onerecipe_category_label': 'Categoría:',
        'onerecipe_persons_label': 'Número de personas:',
        'onerecipe_btn_show': 'Mostrar la receta',
        'onerecipe_btn_export_pdf': '📄 Exportar a PDF',
        'onerecipe_btn_print': '🖨️ Imprimir',
        'onerecipe_btn_add_to_shopping': '🛒 Añadir a la lista de compras',
        'onerecipe_btn_cooked': '🍳 ¡Cociné esto!',
        'onerecipe_btn_cooking_mode': '🖥️ Modo cocina (pantalla completa)',
        'onerecipe_btn_qr': '📱 Código QR',
        'onerecipe_btn_timers': '⏲️ Temporizadores',
        'onerecipe_btn_cook_log': '📔 Diario de cocina',
        'onerecipe_btn_substitutions': '🔄 Sustitutos posibles',
        'onerecipe_edit_button': '✏️ Modificar',
        'onerecipe_ingredients_info_label': 'Ingredientes e información:',
        'onerecipe_description_notes_label': 'Descripción y notas:',
        'onerecipe_similar_label': 'Recetas similares:',
        'onerecipe_no_photo': '(sin foto)',
        'onerecipe_preview_unavailable': '(vista previa no disponible)',
        'onerecipe_select_recipe_first': 'Selecciona una receta de la lista.',
        'onerecipe_display_first': 'Primero muestra una receta con « Mostrar la receta ».',
        'onerecipe_invalid_persons': 'Número de personas no válido.',
        'onerecipe_added_to_shopping_title': 'Añadido',
        'onerecipe_added_to_shopping_message': '« {name} » ({persons} pers.) se añadirá automáticamente a la lista de compras la próxima vez que abras « Ver todas las recetas ».',
        'onerecipe_pantry_decrement_title': 'Despensa',
        'onerecipe_pantry_decrement_prompt': '¿Descontar los ingredientes de « {name} » ({persons} pers.) de tu despensa?',
        'onerecipe_pantry_updated_title': 'Despensa actualizada',
        'onerecipe_pantry_updated_message': '{count} ingrediente(s) descontado(s) de tu despensa.',
        'onerecipe_pantry_none_decremented': 'No se pudo descontar ningún ingrediente de esta receta (ausente de la despensa, o unidad no comparable).',
        'onerecipe_marked_title': 'Marcada',
        'onerecipe_marked_message': '¡« {name} » ha sido marcada como cocinada hoy!',
        'onerecipe_no_substitutes_title': 'Ningún sustituto conocido',
        'onerecipe_no_substitutes_message': 'Ningún ingrediente de esta receta tiene un sustituto conocido por el momento.\n\nPuedes añadir uno tú mismo desde « 🥕 Gestionar los ingredientes » > « 🔄 Gestionar las sustituciones ».',
        'onerecipe_substitutes_title': 'Sustitutos posibles — {name}',
        'onerecipe_substitutes_heading': '🔄 Sustitutos posibles para « {name} »',
        'onerecipe_substitutes_disclaimer': 'Sugerencias culinarias, no equivalencias garantizadas:\nel resultado puede variar según la receta.',
        'onerecipe_close_button': 'Cerrar',
        'onerecipe_rating_label': 'Valoración: {stars}',
        'onerecipe_prep_label': 'Preparación: {time} min',
        'onerecipe_cook_label': 'Cocción: {time} min',
        'onerecipe_difficulty_label': 'Dificultad: {value}',
        'onerecipe_allergens_label': '⚠ Alérgenos: {list}',
        'onerecipe_cost_label': '💰 Coste estimado: {cost} €{partial}',
        'onerecipe_cost_partial': ' (estimación parcial, {known}/{total} ingredientes con precio conocido)',
        'onerecipe_nutrition_partial': ' (estimación parcial, {known}/{total} ingredientes reconocidos)',
        'onerecipe_nutrition_label': '🥗 Valores nutricionales estimados{partial}:\n   {kcal} kcal · {protein} g proteínas · {carbs} g carbohidratos · {fat} g grasas\n',
        'onerecipe_description_heading': '--- Descripción ---\n{text}\n',
        'onerecipe_notes_heading': '\n--- Notas personales ---\n{text}\n',
        'onerecipe_no_description_notes': '(Ninguna descripción ni nota personal para esta receta.)',
        'onerecipe_export_pdf_title': 'Exportar la receta a PDF',
        'onerecipe_export_success_title': 'Exportación exitosa',
        'onerecipe_export_success_message': 'Receta exportada:\n{path}',
        'onerecipe_export_failed': 'La exportación falló:\n{error}',
        'onerecipe_print_failed': 'La preparación de la impresión falló:\n{error}',
        'onerecipe_pdf_module_missing': "La exportación a PDF requiere el módulo 'reportlab'.\nInstálalo con: pip install reportlab",
        'onerecipe_print_module_missing': "La impresión requiere el módulo 'reportlab' para generar el diseño.\nInstálalo con: pip install reportlab",
        'onerecipe_qr_module_missing': "La exportación a código QR requiere el módulo 'qrcode'.\nInstálalo con: pip install qrcode",
        'onerecipe_qr_pillow_missing': "La exportación a código QR también requiere el módulo 'Pillow'.\nInstálalo con: pip install pillow",
        'onerecipe_default_timer_label': 'Temporizador',
        'recipeform_title_edit': 'Modificar la receta',
        'recipeform_title_add': 'Añadir una receta',
        'recipeform_name_label': 'Nombre de la receta:',
        'recipeform_favorite_checkbox': '⭐ Marcar como receta favorita',
        'recipeform_wishlist_checkbox': '💭 Añadir a mi lista de deseos (para probar)',
        'recipeform_rating_label': 'Mi valoración:',
        'recipeform_category_label': 'Categoría:',
        'recipeform_prep_time_label': 'Preparación (min):',
        'recipeform_cook_time_label': 'Cocción (min):',
        'recipeform_difficulty_label': 'Dificultad:',
        'recipeform_default_persons_label': '   Personas por defecto:',
        'recipeform_tags_label': 'Etiquetas (separadas por comas):',
        'recipeform_tags_example': 'ej. vegetariano, sin gluten, rápido, económico',
        'recipeform_allergens_label': 'Alérgenos presentes:',
        'recipeform_detect_allergens_button': '🔍 Detectar automáticamente',
        'recipeform_allergens_disclaimer': 'Esto es solo informativo, verifica siempre los\nalérgenos en las etiquetas de los productos físicos.',
        'recipeform_allergens_auto_note': 'La detección automática se basa en los ingredientes de la\nreceta ya introducidos abajo: marca y desmarca las\ncasillas en consecuencia, sin tocar nunca las que\nhubieras marcado tú mismo sin relación con un ingrediente detectado.',
        'recipeform_photos_label': 'Fotos:',
        'recipeform_add_photo_button': '📷 Añadir una foto',
        'recipeform_description_label': 'Descripción (información, pasos, consejos...):',
        'recipeform_notes_label': 'Notas personales (opinión, ajustes para la próxima vez...):',
        'recipeform_ingredients_label': 'Ingredientes (cantidad para 1 persona):',
        'recipeform_new_ingredient_button': '🥕 Nuevo ingrediente',
        'recipeform_no_ingredients_registered': 'Ningún ingrediente registrado. Haz clic en « 🥕 Nuevo ingrediente »\npara crear el primero.',
        'recipeform_header_ingredient': 'Ingrediente',
        'recipeform_header_quantity': 'Cantidad',
        'recipeform_header_unit': 'Unidad',
        'recipeform_header_other': '(si es otro)',
        'recipeform_add_ingredient_button': '+ Añadir un ingrediente',
        'recipeform_save_button': 'Guardar',
        'recipeform_delete_button': 'Eliminar esta receta',
        'recipeform_char_counter': '{count} / {max} caracteres',
        'recipeform_add_ingredients_first': 'Primero añade ingredientes a la receta.',
        'recipeform_allergens_updated_title': 'Alérgenos actualizados',
        'recipeform_allergens_updated_added': 'añadido(s): {list}',
        'recipeform_allergens_updated_removed': 'quitado(s): {list}',
        'recipeform_allergens_updated_message': 'Alérgeno(s) {parts}.',
        'recipeform_allergens_no_change': 'Sin cambios: los alérgenos marcados ya corresponden a los ingredientes.',
        'recipeform_choose_photos_title': 'Elegir una o varias fotos',
        'recipeform_no_photo': '(sin foto)',
        'recipeform_preview_unavailable': '(vista previa\nno disponible)',
        'recipeform_remove_photo_button': '🗑 Quitar',
        'recipeform_new_ingredient_dialog_title': 'Nuevo ingrediente',
        'recipeform_new_ingredient_dialog_prompt': 'Nombre del nuevo ingrediente:',
        'recipeform_ingredient_already_exists': 'El ingrediente « {name} » ya existe.',
        'recipeform_ingredient_added_title': 'Añadido',
        'recipeform_ingredient_added_message': 'El ingrediente « {name} » ha sido añadido.\nSelecciónalo en una de las listas desplegables.',
        'recipeform_error_name_required': 'Por favor, indica un nombre de receta.',
        'recipeform_error_prep_time': 'El tiempo de preparación debe ser un número positivo (o estar vacío).',
        'recipeform_error_cook_time': 'El tiempo de cocción debe ser un número positivo (o estar vacío).',
        'recipeform_unknown_ingredient_title': 'Ingrediente desconocido',
        'recipeform_unknown_ingredient_message': '« {name} » no corresponde a ningún ingrediente registrado.\nElige uno de la lista desplegable, o haz clic en « 🥕 Nuevo ingrediente » para añadirlo primero.',
        'recipeform_error_invalid_quantity': "Cantidad no válida para '{name}'.",
        'recipeform_error_custom_unit_required': "Especifica la unidad personalizada para '{name}'.",
        'recipeform_error_no_valid_ingredient': 'Añade al menos un ingrediente válido.',
        'recipeform_duplicate_ingredient_title': 'Ingrediente duplicado',
        'recipeform_duplicate_ingredient_message': '« {list} » aparece varias veces en esta receta.\n\n¿Guardar de todos modos?',
        'recipeform_saved_message': 'La receta « {name} » ha sido guardada.',
        'recipeform_delete_confirm_message': '¿Enviar la receta « {name} » a la papelera?\n\nPodrás restaurarla más tarde desde el botón « 🗑️ Papelera ».',
        'recipeform_deleted_title': 'Enviada a la papelera',
        'recipeform_deleted_message': 'La receta se ha movido a la papelera.',
    },
    "de": {
        'home_window_title': 'Mein Rezeptbuch',
        'home_banner_title': '👨\u200d🍳 Mein Rezeptbuch',
        'home_banner_subtitle': 'Alle Ihre Rezepte griffbereit',
        'home_donate_button': '☕ Spenden',
        'home_dark_theme': '🌙 Dunkles Design',
        'home_light_theme': '☀️ Helles Design',
        'home_large_text_on': '🔎 Vergrößerte Schrift',
        'home_large_text_off': '🔎 Normale Schrift',
        'home_daily_recipe_title': '🎲 Rezept des Tages',
        'home_open_button': '👁 Öffnen',
        'home_quick_filter_favorites': '⭐ Favoriten',
        'home_quick_filter_quick': '⏱️ Schnell (≤ 30 Min.)',
        'home_quick_filter_vegetarian': '🥗 Vegetarisch',
        'home_quick_filter_wishlist': '💭 Wunschliste',
        'home_wishlist_reminder': '💭 {count} Rezept(e) seit mehr als {days} Tagen auf Ihrer Wunschliste — wie wäre es, sie auszuprobieren? (klicken, um sie anzuzeigen)',
        'home_low_stock_reminder': '📦 {count} Zutat(en) in Ihrer Vorratskammer fast aufgebraucht: {names} — klicken, um sie zur Einkaufsliste hinzuzufügen',
        'home_btn_add_recipe': '➕  Rezept hinzufügen',
        'home_btn_import_url': '🌐  Rezept von einem Link importieren',
        'home_btn_import_photo': '📷  Rezept von einem Foto importieren',
        'home_btn_view_all_recipes': '🧾  Alle Rezepte anzeigen (Einkaufsliste)',
        'home_btn_view_one_recipe': '🍽️  Ein bestimmtes Rezept anzeigen',
        'home_btn_manage_recipes': '✏️  Rezept ändern / löschen',
        'home_btn_compare_recipes': '⚖️  Zwei Rezepte vergleichen',
        'home_btn_manage_ingredients': '🥕  Zutaten verwalten',
        'home_btn_ingredient_search': '🔎  Suche nach Zutat',
        'home_btn_what_can_i_cook': '🧊  Was kann ich kochen?',
        'home_btn_pantry': '📦  Meine Vorratskammer',
        'home_btn_unit_converter': '🔄  Einheitenumrechner',
        'home_btn_weekly_plan': '📅  Wochenplan',
        'home_btn_menus': '📋  Meine Menüs',
        'home_btn_statistics': '📊  Statistiken',
        'home_btn_export_cookbook': '📖  Kochbuch exportieren',
        'home_btn_import_export': '💾  Daten importieren / exportieren',
        'home_btn_trash': '🗑️  Papierkorb',
        'home_today_title': '📅 Heute',
        'home_recent_title': '🕘 Kürzlich angesehen',
        'home_wishlist_title': '💭 Rezepte zum Ausprobieren',
        'home_new_draw_button': '🎲 Neue Auswahl',
        'home_footer_recipe_count': '{count} gespeicherte(s) Rezept(e)',
        'home_nothing_planned': 'Nichts geplant für {day}. Füllen Sie den « 📅 Wochenplan » aus, um es hier zu sehen.',
        'home_no_recent_recipe': 'Noch kein Rezept angesehen.',
        'home_no_wishlist_recipe': 'Noch kein Rezept auf Ihrer Wunschliste.',
        'warning_pillow': 'Pillow nicht installiert: Fotos werden nicht angezeigt (pip install pillow)',
        'warning_reportlab': 'reportlab nicht installiert: PDF-Export nicht verfügbar (pip install reportlab)',
        'warning_openpyxl': 'openpyxl nicht installiert: Excel-Export nicht verfügbar (pip install openpyxl)',
        'warning_qrcode': 'qrcode nicht installiert: QR-Code-Export nicht verfügbar (pip install qrcode)',
        'warning_pytesseract': 'pytesseract nicht installiert: Import von Foto nicht verfügbar (pip install pytesseract, + Tesseract OCR)',
        'common_error': 'Fehler',
        'common_info': 'Info',
        'common_confirm': 'Bestätigen',
        'common_success': 'Erfolg',
        'common_module_missing': 'Modul fehlt',
        'common_all_categories': 'Alle',
        'common_export_failed': 'Export fehlgeschlagen:\n{error}',
        'common_export_success_title': 'Export erfolgreich',
        'common_print_failed': 'Vorbereitung des Drucks fehlgeschlagen:\n{error}',
        'common_reset_button': 'Zurücksetzen',
        'common_want_label': 'Ich möchte:',
        'common_exclude_label': 'Ich möchte nicht:',
        'common_tags_filter_label': 'Tags (alle erforderlich):',
        'common_filter_hint': 'Geben Sie die ersten Buchstaben ein, um die Liste zu filtern.',
        'common_search_label': '🔍 Suchen:',
        'common_sort_by_label': 'Sortieren nach:',
        'common_category_label': 'Kategorie:',
        'common_edit_button': '✏️ Ändern',
        'common_unknown_ingredient_title': 'Unbekannte Zutat',
        'common_unknown_ingredient_simple_message': '« {name} » entspricht keiner gespeicherten Zutat.\nWählen Sie eine aus der Dropdown-Liste aus.',
        'common_ingredient_label': 'Zutat:',
        'common_quantity_label': 'Menge:',
        'common_unit_label': 'Einheit:',
        'common_new_ingredient_button': '🥕 Neue Zutat',
        'common_save_button': '💾 Speichern',
        'pantry_title': 'Meine Vorratskammer',
        'pantry_heading': '📦 Meine Vorratskammer',
        'pantry_intro': 'Geben Sie an, was Sie zu Hause haben und in welcher Menge.\n„Was kann ich kochen?“ kann dann prüfen, ob Sie genug davon haben,\nund vorschlagen, den Bestand nach dem Kochen automatisch abzuziehen.',
        'pantry_threshold_label': 'Warnschwelle (optional):',
        'pantry_help_text': 'Um einen Artikel HINZUZUFÜGEN: Geben Sie die Zutat an (legen Sie sie zuerst mit\n« 🥕 Neue Zutat » an, falls sie noch nicht in Ihrer Liste ist), die\nMenge und die Einheit, und klicken Sie dann auf « 💾 Speichern ».\nUm einen bereits vorhandenen Artikel zu ÄNDERN: Klicken Sie einmal darauf in der\nListe unten — dies lädt seine Werte in die Felder oben,\nohne etwas zu speichern: Ändern Sie die gewünschten Werte UND klicken Sie dann auf\n« 💾 Speichern », damit die Änderung übernommen wird.\nDie Warnschwelle löst eine Erinnerung auf der Startseite aus, sobald die\nMenge darunter fällt (leer lassen, wenn Sie nie benachrichtigt werden möchten).',
        'pantry_remove_button': '🗑 Aus der Vorratskammer entfernen',
        'pantry_empty': 'Ihre Vorratskammer ist derzeit leer.',
        'pantry_threshold_suffix': ' (Schwelle: {threshold})',
        'pantry_error_ingredient_required': 'Bitte geben Sie eine Zutat an.',
        'pantry_error_invalid_quantity': 'Ungültige Menge.',
        'pantry_error_invalid_threshold': 'Ungültige Warnschwelle (leer lassen, wenn Sie keine möchten).',
        'pantry_select_ingredient_first': 'Wählen Sie eine Zutat aus der Liste aus.',
        'pantry_remove_confirm_message': '« {name} » aus der Vorratskammer entfernen?',
        'cook_title': 'Was kann ich kochen?',
        'cook_instructions_label': 'Geben Sie die Zutaten an, die Sie zu Hause haben:',
        'cook_staples_hint': 'Einige gängige Grundzutaten sind bereits nebenan angekreuzt\n(Salz, Öl, Mehl...) — entfernen Sie diejenigen, die Sie nicht haben.',
        'cook_all_ingredients_label': 'Alle Zutaten:',
        'cook_add_button': '➕ Hinzufügen →',
        'cook_have_label': 'Was ich habe:',
        'cook_remove_button': '🗑 Entfernen',
        'cook_load_from_pantry_button': '📦 Aus meiner Vorratskammer laden',
        'cook_compute_button': '🔍 Machbare Rezepte anzeigen',
        'cook_open_selected_button': '📖 Ausgewähltes Rezept anzeigen',
        'cook_pantry_empty_title': 'Info',
        'cook_pantry_empty_message': 'Ihre Vorratskammer ist derzeit leer. Öffnen Sie « 📦 Meine Vorratskammer » von der Startseite aus, um Zutaten hinzuzufügen.',
        'cook_loaded_title': 'Geladen',
        'cook_loaded_message': '{count} Zutat(en) aus Ihrer Vorratskammer hinzugefügt.',
        'cook_add_ingredient_first': 'Fügen Sie mindestens eine Zutat hinzu, die Sie haben.',
        'cook_feasible_header': '✅ Machbar mit dem, was Sie haben:',
        'cook_insufficient_quantity': '  ⚠️ unzureichende Menge: {list}',
        'cook_none_feasible': 'Mit diesen Zutaten ist kein Rezept zu 100% machbar.',
        'cook_substitutable_header': '🔄 Machbar mit einem Ersatz:',
        'cook_almost_header': '🟡 Fast (es fehlen 1 bis 3 Zutaten):',
        'cook_missing_label': '   {name} (fehlt: {list})',
        'cook_no_results': 'Versuchen Sie, weitere Zutaten zu Ihrer Auswahl hinzuzufügen.',
        'cook_select_recipe_from_results': 'Wählen Sie ein Rezept aus der Ergebnisliste aus.',
        'cook_select_recipe_row': 'Wählen Sie eine Zeile aus, die einem Rezept entspricht.',
        'weekhistory_title': 'Verlauf vergangener Wochen',
        'weekhistory_heading': '🕘 Verlauf vergangener Wochen',
        'weekhistory_intro': 'Jede Woche, in der Sie den Plan speichern, wird hier automatisch\narchiviert (bis zu 26 Wochen, etwa 6 Monate), um zu vermeiden,\ndasselbe zu oft zu wiederholen.',
        'weekhistory_reload_button': '♻️ In den aktuellen Plan neu laden',
        'weekhistory_delete_button': '🗑 Diese Woche löschen',
        'weekhistory_no_archived_weeks': 'Keine archivierte Woche.',
        'weekhistory_week_label': 'Woche {week}',
        'weekhistory_saved_on': 'Gespeichert am {date}\n\n',
        'weekhistory_day_heading': '{day}:\n',
        'weekhistory_slot_line': '   {slot}: {recipe} ({persons} Pers.)\n',
        'weekhistory_empty_week': '(Plan für diese Woche leer.)',
        'weekhistory_select_week_first': 'Wählen Sie eine Woche aus der Liste aus.',
        'weekhistory_reload_confirm_message': 'Den Plan der Woche {week} in den aktuellen Plan neu laden?\n\nDies ersetzt die aktuell angezeigten Rezepte (denken Sie daran, den laufenden Plan vorher zu speichern, wenn Sie ihn behalten möchten).',
        'weekhistory_delete_confirm_message': 'Das Archiv der Woche {week} endgültig löschen?',
        'weektemplates_title': 'Wochenvorlagen',
        'weektemplates_heading': '📋 Wochenvorlagen',
        'weektemplates_intro': 'Speichern Sie den aktuell angezeigten Plan als wiederverwendbare\nVorlage, um ihn mit einem Klick auf eine andere Woche anzuwenden,\nanstatt alles neu einzugeben.',
        'weektemplates_name_label': 'Name der neuen Vorlage:',
        'weektemplates_save_button': '💾 Aktuellen Plan als Vorlage speichern',
        'weektemplates_apply_button': '📋 Diese Vorlage anwenden',
        'weektemplates_delete_button': '🗑 Diese Vorlage löschen',
        'weektemplates_none_saved': 'Noch keine Vorlage gespeichert.',
        'weektemplates_error_name_required': 'Bitte geben Sie einen Namen für diese Vorlage an.',
        'weektemplates_empty_plan': 'Der aktuell angezeigte Plan ist leer: nichts als Vorlage zu speichern.',
        'weektemplates_saved_message': 'Vorlage « {name} » gespeichert.',
        'weektemplates_select_template_first': 'Wählen Sie eine Vorlage aus der Liste aus.',
        'weektemplates_apply_confirm_message': 'Die Vorlage « {name} » auf den aktuellen Plan anwenden?\n\nDies ersetzt die aktuell angezeigten Rezepte (denken Sie daran, den laufenden Plan vorher zu speichern, wenn Sie ihn behalten möchten).',
        'weektemplates_delete_confirm_message': 'Die Vorlage « {name} » endgültig löschen?',
        'common_none_option': '-- Keine --',
        'weekplan_title': 'Wochenplan',
        'weekplan_subtitle': 'Kalenderansicht: Tage in Spalten, Mahlzeiten in Zeilen.',
        'weekplan_save_button': '💾 Plan speichern',
        'weekplan_clear_button': '🗑 Alles löschen',
        'weekplan_export_ics_button': '📆 In einen Kalender exportieren (.ics)',
        'weekplan_compute_button': 'Einkaufsliste der Woche berechnen',
        'weekplan_checklist_button': '☑️ Einkaufsmodus',
        'weekplan_empty_list_message': 'Noch keine Liste berechnet.\nKlicken Sie oben auf « Einkaufsliste der Woche berechnen »,\noder laden Sie eine gespeicherte Liste.',
        'weekplan_total_list_heading': '=== Einkaufsliste der Woche ===',
        'weekplan_calculate_list_for_export': 'Berechnen Sie zuerst eine Einkaufsliste (Schaltfläche « Einkaufsliste der Woche berechnen »).',
        'weekplan_invalid_persons_for_slot': 'Ungültige Personenanzahl für {day} — {slot}.',
        'weekplan_saved_message': 'Der Wochenplan wurde gespeichert.',
        'weekplan_clear_confirm_message': 'Den gesamten Wochenplan löschen?',
        'weekplan_assign_recipe_first': 'Weisen Sie mindestens ein Rezept einem Termin der Woche zu.',
        'weekplan_export_ics_title': 'Plan in einen Kalender exportieren',
        'weekplan_ics_export_success_message': 'Plan exportiert:\n{path}\n\nImportieren Sie diese Datei in Google Kalender, Outlook oder Kalender, um Ihre Mahlzeiten dort jede Woche wiederholt zu sehen.',
        'weekplan_assign_or_manual': 'Weisen Sie mindestens ein Rezept einem Termin der Woche zu, oder fügen Sie eine Zutat manuell hinzu.',
        'weekplan_export_shopping_list_title': 'Einkaufsliste speichern',
        'weekplan_shopping_list_title': 'Einkaufsliste der Woche',
        'weekplan_list_saved_message': 'Liste gespeichert:\n{path}',
        'weekplan_excel_module_missing': 'Der Excel-Export erfordert: pip install openpyxl',
        'weekplan_pdf_module_missing': 'Der PDF-Export erfordert: pip install reportlab',
        'weekplan_print_module_missing': 'Das Drucken erfordert: pip install reportlab',
        'weekplan_print_label': 'die Einkaufsliste der Woche',
        'manageing_title': 'Zutaten verwalten',
        'manageing_list_label': 'Liste der gespeicherten Zutaten:',
        'manageing_add_button': '➕ Hinzufügen',
        'manageing_edit_button': '✏️ Ändern',
        'manageing_delete_button': '🗑️ Löschen',
        'manageing_load_defaults_button': '📚 Die ~1000 gängigen Zutaten laden',
        'manageing_spell_check_button': '🔤 Duplikate / Tippfehler prüfen',
        'manageing_prices_button': '💰 Preise verwalten (für Rezeptkosten)',
        'manageing_substitutions_button': '🔄 Ersatzzutaten verwalten',
        'manageing_edit_hint': '„Ändern“ ermöglicht es, den Namen zu ändern (überall aktualisiert,\nwo die Zutat verwendet wird), ihre Allergene,\nihre Nährwerte und ihren Preis.',
        'manageing_select_ingredient_first': 'Wählen Sie eine Zutat aus der Liste aus.',
        'manageing_delete_confirm_message': '« {name} » aus der Zutatenliste löschen?',
        'manageing_delete_usage_warning': '\n\nAchtung: sie wird in {count} Rezept(en) verwendet. Diese Rezepte behalten diese Zutat, aber sie wird nicht mehr im Dropdown-Menü vorgeschlagen, es sei denn, Sie fügen sie erneut hinzu.',
        'manageing_missing_file_title': 'Datei fehlt',
        'manageing_missing_file_message': 'Die Datei ingredients_par_defaut.json wurde nicht gefunden.\nStellen Sie sicher, dass sie sich im selben Ordner wie main.py befindet.',
        'manageing_done_title': 'Fertig',
        'manageing_defaults_added_message': '{count} neue Zutat(en) aus der gängigen Liste hinzugefügt.',
        'manageing_defaults_none_added': 'Alle gängigen Zutaten waren bereits vorhanden.',
        'subedit_title': 'Ersatzzutaten für « {name} »',
        'subedit_heading': '🔄 Ersatzzutaten für « {name} »',
        'subedit_disclaimer': 'Ein Ersatz ist ein kulinarischer Ratschlag, keine garantierte\nGleichwertigkeit: Das Ergebnis kann je nach Rezept variieren.',
        'subedit_remove_button': '🗑 Ausgewählten Ersatz entfernen',
        'subedit_add_frame_title': 'Ersatz hinzufügen',
        'subedit_name_label': 'Name:',
        'subedit_note_label': 'Notiz (optional):',
        'subedit_add_to_list_button': '➕ Zur Liste hinzufügen',
        'subedit_revert_button': '🔄 Zur mitgelieferten Basis zurückkehren',
        'subedit_cancel_button': 'Abbrechen',
        'subedit_no_substitute_yet': 'Noch kein Ersatz.',
        'subedit_error_name_required': 'Bitte geben Sie einen Namen für den Ersatz an.',
        'subedit_select_to_remove': 'Wählen Sie einen zu entfernenden Ersatz aus.',
        'subedit_revert_confirm_message': 'Ihre benutzerdefinierte Liste entfernen und zu den mit der Anwendung gelieferten Ersatzzutaten für « {name} » zurückkehren?',
        'managesub_title': 'Ersatzzutaten verwalten',
        'managesub_heading': '🔄 Zutatenersatz',
        'managesub_intro': 'Sehen oder ändern Sie die vorgeschlagenen Ersatzzutaten für eine Zutat.\nEin Ersatz ist ein kulinarischer Ratschlag, keine garantierte Gleichwertigkeit.',
        'managesub_manage_button': '✏️ Ihre Ersatzzutaten verwalten',
        'managesub_hint': 'Doppelklicken Sie auf eine Zutat in der Liste, um ihre\nErsatzzutaten anzuzeigen oder zu ändern, oder geben Sie oben einen Namen ein (auch eine Zutat, die noch\nkeinen bekannten Ersatz hat) und dann « ✏️ Ihre Ersatzzutaten verwalten ».',
        'managesub_none_with_substitute': 'Noch keine Zutat mit Ersatz.',
        'managesub_substitute_count': '{name} ({count} Ersatzzutat(en){plural})',
        'managesub_error_ingredient_required': 'Bitte geben Sie eine Zutat an.',
        'managesub_unknown_ingredient_message': '« {name} » entspricht keiner gespeicherten Zutat.\nWählen Sie eine aus der Dropdown-Liste aus, oder legen Sie sie zuerst über « 🥕 Zutaten verwalten » an.',
        'ingprices_title': 'Zutatenpreise verwalten',
        'ingprices_heading': '💰 Zutatenpreise',
        'ingprices_intro': 'Geben Sie einen Preis für die Zutaten an, die Sie\ninteressieren — es ist nicht nötig, alle einzutragen. Die Kosten\neines Rezepts werden anhand dieser Preise geschätzt.',
        'ingprices_price_label': 'Preis (€):',
        'ingprices_for_one_label': 'für 1',
        'ingprices_save_button': '💾 Preis speichern',
        'ingprices_clear_button': '🗑 Preis löschen',
        'ingprices_units_note': 'kg ↔ Rezepte in Gr   ·   L ↔ Rezepte in cl   ·   Preise\npro Stück/Löffel gelten unverändert.',
        'ingprices_no_price_set': '  —  (kein Preis angegeben)',
        'ingprices_price_suffix': '  —  {price} € / {unit}',
        'ingprices_error_invalid_price': 'Geben Sie einen gültigen Preis ein (positive Zahl).',
        'ingprices_saved_message': 'Preis für « {name} » gespeichert.',
        'ingedit_title_edit': 'Zutat ändern',
        'ingedit_title_new': 'Neue Zutat',
        'ingedit_heading_edit': '✏️ Zutat ändern',
        'ingedit_heading_new': '➕ Neue Zutat',
        'ingedit_name_label': 'Name:',
        'ingedit_allergens_label': 'Enthaltene Allergene:',
        'ingedit_nutrition_label': 'Nährwerte (pro 100 g / 100 ml):',
        'ingedit_nutri_kcal': 'Kalorien (kcal)',
        'ingedit_nutri_protein': 'Eiweiß (g)',
        'ingedit_nutri_carbs': 'Kohlenhydrate (g)',
        'ingedit_nutri_fat': 'Fett (g)',
        'ingedit_nutrition_hint': 'Leer lassen, wenn Sie diese Werte nicht kennen.',
        'ingedit_price_label': 'Preis:',
        'ingedit_save_button': '💾 Speichern',
        'ingedit_delete_button': '🗑️ Diese Zutat löschen',
        'ingedit_error_invalid_field': '« {field} » muss eine positive Zahl sein (oder leer).',
        'ingedit_error_name_required': 'Bitte geben Sie einen Zutatennamen an.',
        'ingedit_error_already_exists': 'Die Zutat « {name} » existiert bereits.',
        'ingedit_error_plural_duplicate': '« {name} » ist nur eine Singular-/Plural-Variante der bereits vorhandenen Zutat « {existing} ». Um Duplikate in der Liste zu vermeiden, verwenden Sie direkt « {existing} ».',
        'ingedit_nutri_field_kcal': 'Kalorien',
        'ingedit_nutri_field_protein': 'Eiweiß',
        'ingedit_nutri_field_carbs': 'Kohlenhydrate',
        'ingedit_nutri_field_fat': 'Fett',
        'ingedit_error_invalid_price': 'Der Preis muss eine positive Zahl sein (oder leer).',
        'ingedit_saved_message': '« {name} » wurde gespeichert.',
        'spellcheck_title': 'Rechtschreibprüfung der Zutaten',
        'spellcheck_heading': 'Zutatenpaare, die sich zu 90% oder mehr ähneln\n(vermutliche Duplikate oder Tippfehler):',
        'spellcheck_multi_select_hint': 'Mehrfachauswahl möglich (Strg+Klick oder Umschalt+Klick), um\nmehrere Paare auf einmal zusammenzuführen.',
        'spellcheck_merge_button': '🔗 Auswahl zusammenführen',
        'spellcheck_not_duplicate_button': '✕ Kein Duplikat',
        'spellcheck_rerun_button': '🔄 Analyse erneut starten',
        'spellcheck_footer_hint': 'Bei einem einzelnen Paar werden Sie gefragt, welche der beiden\nSchreibweisen beibehalten werden soll. Bei mehreren Paaren gleichzeitig wird die\nin Ihren Rezepten am wenigsten verwendete Zutat automatisch\nmit der in den meisten Rezepten verwendeten zusammengeführt.\n„Kein Duplikat“ entfernt das oder die ausgewählten\nPaare endgültig aus dieser Analyse, heute und in Zukunft.',
        'spellcheck_none_found': 'Kein wahrscheinliches Duplikat gefunden. 🎉',
        'spellcheck_pair_line': '{a}   ↔   {b}     ({percent}% ähnlich)',
        'spellcheck_select_pair_first': 'Wählen Sie mindestens ein Paar aus der Liste aus.',
        'spellcheck_dismissed_message': '{count} Paar(e) als kein Duplikat markiert. Sie werden bei zukünftigen Analysen nicht mehr vorgeschlagen.',
        'spellcheck_merge_dialog_title': 'Zusammenführen',
        'spellcheck_merge_dialog_message': '« {a} » und « {b} » zusammenführen?\n\nJa = alles in « {a} » umbenennen\nNein = alles in « {b} » umbenennen\nAbbrechen = nichts tun',
        'spellcheck_merged_title': 'Zusammengeführt',
        'spellcheck_merged_one_message': '« {removed} » wurde mit « {kept} » zusammengeführt.',
        'spellcheck_merge_multi_confirm': 'Diese {count} Paare automatisch zusammenführen?\n\nFür jedes Paar wird die in Ihren Rezepten am wenigsten verwendete Zutat mit der in den meisten Rezepten verwendeten zusammengeführt (bei Gleichstand die erste in alphabetischer Reihenfolge).',
        'spellcheck_merged_multi_message': '{count} Paar(e) zusammengeführt.',
        'compare_title': 'Zwei Rezepte vergleichen',
        'compare_recipe_a_label': 'Rezept A:',
        'compare_recipe_b_label': 'Rezept B:',
        'compare_button': '⚖️ Vergleichen',
        'compare_choose_each_list': 'Wählen Sie ein Rezept in jeder Liste aus.',
        'compare_field_category': 'Kategorie:',
        'compare_field_favorite': 'Favorit:',
        'compare_yes': '⭐ Ja',
        'compare_no': 'Nein',
        'compare_field_rating': 'Bewertung:',
        'compare_field_difficulty': 'Schwierigkeit:',
        'compare_field_prep': 'Zubereitung:',
        'compare_field_cook': 'Kochzeit:',
        'compare_field_total_time': 'Gesamtzeit:',
        'compare_field_cooked': 'Gekocht:',
        'compare_times_suffix': '{count} Mal',
        'compare_field_cost': 'Geschätzte Kosten:',
        'compare_field_nutrition': 'Nährwerte (kcal):',
        'compare_field_ingredient_count': 'Anz. Zutaten:',
        'compare_common_ingredients': '🟰 Gemeinsam ({count})',
        'compare_only_a': '🅰️ Nur in « {name} » ({count})',
        'compare_only_b': '🅱️ Nur in « {name} » ({count})',
        'compare_none': 'Keine',
        'stats_title': 'Statistiken',
        'stats_heading': '=== Statistiken ===\n\n',
        'stats_total_recipes': 'Gesamtzahl der Rezepte: {count}\n\n',
        'stats_by_category': 'Verteilung nach Kategorie:\n',
        'stats_category_line': '  - {category}: {count}\n',
        'stats_by_difficulty': 'Verteilung nach Schwierigkeit:\n',
        'stats_difficulty_line': '  - {difficulty}: {count}\n',
        'stats_difficulty_unspecified': 'Nicht angegeben',
        'stats_favorites_count': 'Lieblingsrezepte: {count}\n\n',
        'stats_avg_rating': 'Durchschnittliche Bewertung (bewertete Rezepte): {avg} / 5 ({count} bewertete(s) Rezept(e))\n\n',
        'stats_no_rated_recipe': 'Durchschnittliche Bewertung: noch kein Rezept bewertet.\n\n',
        'stats_five_star_heading': 'Mit 5 Sternen bewertete(s) Rezept(e):\n',
        'stats_recipe_line': '  - {name}\n',
        'stats_most_cooked_heading': 'Am häufigsten gekochte Rezepte:\n',
        'stats_cooked_line': '  - {name}: {count} Mal\n',
        'stats_none_cooked_yet': '  Noch kein Rezept als gekocht markiert.\n  (Schaltfläche « 🍳 Habe ich gekocht! » in « Ein bestimmtes Rezept anzeigen »)\n',
        'stats_most_used_tags_heading': 'Am häufigsten verwendete Tags:\n',
        'stats_tag_line': '  - {tag}: {count}\n',
        'stats_never_cooked_heading': '🕸️ Nie gekochte Rezepte:\n',
        'stats_and_others': '  ... und {count} weitere\n',
        'stats_all_cooked': '  Alle Ihre Rezepte wurden bereits mindestens einmal gekocht. 👏\n',
        'stats_stale_heading': '🕰️ Seit mehr als {days} Tagen nicht gekocht:\n',
        'stats_stale_line': '  - {name} (vor {days} Tagen)\n',
        'stats_no_stale_recipe': '  Derzeit kein Rezept in diesem Fall.\n',
        'stats_avg_cost_heading': '💰 Durchschnittliche Kosten pro Person:\n',
        'stats_avg_cost_line': '  {avg} € im Durchschnitt, über {count} Rezept(e) mit mindestens einem bekannten Preis ({without_price} ohne angegebenen Preis)\n',
        'stats_no_priced_recipe': '  Noch kein Rezept mit angegebenem Preis.\n  (siehe « 💰 Preise verwalten » in « Zutaten verwalten »)\n',
        'stats_avg_kcal_heading': '🥗 Durchschnittliche Kalorien pro Person:\n',
        'stats_avg_kcal_line': '  {avg} kcal im Durchschnitt, über {count} Rezept(e) mit in der Nährwertdatenbank erkannten Zutaten\n',
        'stats_no_recognized_recipe': '  Noch kein Rezept mit erkannten Zutaten.\n',
        'stats_monthly_chart_title': '📈 Gekochte Rezepte pro Monat (letzte 12 Monate)',
        'stats_heatmap_title': '🗓️ Kalender der Kochtage (letzte 12 Monate)',
        'stats_heatmap_legend': 'Weniger ⬜ 🟨 🟧 🟥 Mehr',
        'stats_day_labels': 'M,D,M,D,F,S,S',
        'stats_month_labels_short': 'Jan,Feb,Mär,Apr,Mai,Jun,Jul,Aug,Sep,Okt,Nov,Dez',
        'stats_month_labels_lower': 'jan,feb,mär,apr,mai,jun,jul,aug,sep,okt,nov,dez',
        'recipepdf_category_persons': 'Kategorie: {cat}    Für {persons} Person(en)',
        'recipepdf_rating': 'Bewertung: {stars}',
        'recipepdf_prep': 'Zubereitung: {time} Min.',
        'recipepdf_cook': 'Kochzeit: {time} Min.',
        'recipepdf_difficulty': 'Schwierigkeit: {value}',
        'recipepdf_allergens': '⚠ Allergene: {list}',
        'recipepdf_ingredients_heading': 'Zutaten:',
        'recipepdf_cost': 'Geschätzte Kosten: {cost} €{partial}',
        'recipepdf_partial_suffix': ' (teilweise, {known}/{total})',
        'recipepdf_nutrition': 'Geschätzte Nährwerte{partial}: {kcal} kcal · {protein}g Eiweiß · {carbs}g Kohlenhydrate · {fat}g Fett',
        'recipepdf_description_heading': 'Beschreibung:',
        'recipepdf_notes_heading': 'Persönliche Notizen:',
        'cookbookpdf_page_number': 'Seite {current} / {total}',
        'cookbookpdf_generated_on': 'Erstellt am {date}',
        'cookbookpdf_summary_heading': 'Inhaltsverzeichnis',
        'cookbookpdf_summary_line': '- [{cat}] {name}',
        'cookbookexport_title': 'Kochbuch exportieren',
        'cookbookexport_heading': '📖 Kochbuch exportieren',
        'cookbookexport_intro': 'Wählen Sie die Rezepte aus, die in ein einzelnes PDF im\nKochbuchstil aufgenommen werden sollen.',
        'cookbookexport_filter_label': 'Nach Kategorie filtern:',
        'cookbookexport_check_all_button': 'Alle ankreuzen',
        'cookbookexport_uncheck_all_button': 'Alle abwählen',
        'cookbookexport_generate_button': '📄 PDF des Buches erstellen',
        'cookbookexport_error_select_recipe': 'Wählen Sie mindestens ein Rezept aus.',
        'cookbookexport_save_dialog_title': 'Kochbuch speichern',
        'cookbookexport_saved_message': 'Kochbuch gespeichert:\n{path}',
        'importexport_title': 'Daten importieren / exportieren',
        'importexport_heading': 'Ihre Daten sichern oder übertragen',
        'importexport_export_intro': 'Der Export erstellt eine .zip-Datei mit absolut allen\nIhren Daten: Rezepte, Fotos, benutzerdefinierte Zutaten,\nPreise, Ersatzzutaten, Vorratskammer, Plan und dessen Verlauf,\nMenüs, gespeicherte Einkaufslisten, Papierkorb und\nEinstellungen — um alles in einer einzigen Datei zu sichern oder auf einen\nanderen Computer zu übertragen.',
        'importexport_export_button': '📤 Alle meine Daten exportieren (.zip)',
        'importexport_import_intro': 'Der Import liest eine zuvor exportierte .zip-Datei.\n„Zusammenführen“ fügt doppelte Rezepte/Fotos unter einem\nneuen Namen hinzu, anstatt sie zu verlieren, und ergänzt den Rest\n(Vorratskammer, Menüs, Listen...), ohne etwas zu löschen.\n„Ersetzen“ überschreibt alles, einschließlich der Einstellungen und des\naktuellen Plans.',
        'importexport_import_button': '📥 Daten importieren (.zip)',
        'importexport_auto_backups_heading': '🗄️ Automatische Sicherungen',
        'importexport_auto_backups_intro': 'Beim Start der Anwendung wird automatisch eine Sicherung erstellt\n(höchstens eine alle {hours} Std.), und die {retention} neuesten\nwerden hier aufbewahrt.',
        'importexport_backup_now_button': '💾 Jetzt sichern',
        'importexport_restore_selected_button': '♻️ Auswahl wiederherstellen',
        'importexport_cloud_heading': '☁️ Automatische Cloud-Sicherung',
        'importexport_cloud_intro': 'Wählen Sie einen Ordner, der von einem bereits auf diesem PC\ninstallierten Client synchronisiert wird (Google Drive, OneDrive, Dropbox...).\nJede automatische Sicherung wird dorthin kopiert, und dieser\nClient kümmert sich selbst um das Senden in die Cloud.',
        'importexport_choose_cloud_button': '📁 Cloud-Ordner auswählen',
        'importexport_disable_button': '🚫 Deaktivieren',
        'importexport_cloud_enabled': '✅ Aktiviert: {folder}',
        'importexport_cloud_not_configured': 'Derzeit nicht konfiguriert.',
        'importexport_choose_folder_title': 'Synchronisierten Ordner auswählen (Google Drive, OneDrive, Dropbox...)',
        'importexport_cloud_configured_title': 'Ordner konfiguriert',
        'importexport_cloud_configured_message': 'Cloud-Ordner konfiguriert:\n{folder}\n\nMöchten Sie jetzt eine Sicherung dorthin kopieren?',
        'importexport_disabled_title': 'Deaktiviert',
        'importexport_disabled_message': 'Die automatische Cloud-Sicherung ist deaktiviert.',
        'importexport_backup_date_line': '{date}   ({size} KB)',
        'importexport_no_backups': 'Noch keine automatische Sicherung.',
        'importexport_backup_failed': 'Sicherung fehlgeschlagen:\n{error}',
        'importexport_backup_created_title': 'Sicherung erstellt',
        'importexport_backup_created_message': 'Eine neue automatische Sicherung wurde erstellt.',
        'importexport_select_backup_first': 'Wählen Sie eine Sicherung aus der Liste aus.',
        'importexport_restore_mode_title': 'Wiederherstellungsmodus',
        'importexport_restore_mode_message': 'Wie soll diese Sicherung wiederhergestellt werden?\n\nJa = Zusammenführen (zu aktuellen Daten hinzufügen, ohne etwas zu löschen)\nNein = Aktuelle Daten vollständig ersetzen\nAbbrechen = nichts tun',
        'importexport_restore_failed': 'Wiederherstellung fehlgeschlagen:\n{error}',
        'importexport_restore_done_title': 'Wiederherstellung abgeschlossen',
        'importexport_restore_done_message': 'Die Daten wurden erfolgreich wiederhergestellt.',
        'importexport_export_data_title': 'Meine Daten exportieren',
        'importexport_export_data_success': 'Ihre Daten wurden exportiert nach:\n{path}',
        'importexport_choose_archive_title': 'Zu importierendes Archiv auswählen',
        'importexport_import_mode_title': 'Importmodus',
        'importexport_import_mode_message': 'Wie sollen diese Daten importiert werden?\n\nJa = Zusammenführen (zu aktuellen Daten hinzufügen, ohne etwas zu löschen)\nNein = Aktuelle Daten vollständig ersetzen\nAbbrechen = nichts tun',
        'importexport_import_failed': 'Import fehlgeschlagen:\n{error}',
        'importexport_import_done_title': 'Import abgeschlossen',
        'importexport_import_done_message': 'Die Daten wurden erfolgreich importiert.',
        'checklist_instruction': 'Kreuzen Sie jeden Artikel während Ihres Einkaufs an.',
        'checklist_check_all_button': '☑️ Alle ankreuzen',
        'checklist_uncheck_all_button': '⬜ Alle abwählen',
        'checklist_progress_label': '{done} / {total} Artikel angekreuzt',
        'exportformat_title': 'Exportformat auswählen',
        'exportformat_heading': '📤 Einkaufsliste exportieren',
        'exportformat_choose_label': 'Wählen Sie das gewünschte Exportformat:',
        'exportformat_txt_button': '📝 Als Text exportieren (.txt)',
        'exportformat_excel_button': '📊 Als Excel exportieren (.xlsx)',
        'exportformat_pdf_button': '📄 Als PDF exportieren (.pdf)',
        'exportformat_cancel_button': 'Abbrechen',
        'menumanager_title': 'Meine Menüs',
        'menumanager_list_label': 'Meine gespeicherten Menüs:',
        'menumanager_new_button': '➕ Neues Menü',
        'menumanager_recipe_count': '{name} ({count} Rezept(e))',
        'menumanager_select_menu_first': 'Wählen Sie ein Menü aus der Liste aus.',
        'menumanager_delete_confirm': 'Das Menü « {name} » löschen?',
        'menuform_title_edit': 'Menü ändern',
        'menuform_title_new': 'Neues Menü',
        'menuform_name_label': 'Name des Menüs:',
        'menuform_add_recipe_label': 'Rezept zum Menü hinzufügen:',
        'menuform_persons_short_label': 'Pers.:',
        'menuform_add_button': '+ Hinzufügen',
        'menuform_recipes_label': 'Rezepte des Menüs:',
        'menuform_remove_button': '🗑 Aus dem Menü entfernen',
        'menuform_save_button': '💾 Menü speichern',
        'menuform_compute_button': 'Einkaufsliste des Menüs berechnen',
        'menuform_empty_list_message': 'Noch keine Liste berechnet.\nKlicken Sie oben auf « Einkaufsliste des Menüs berechnen »,\noder laden Sie eine gespeicherte Liste.',
        'menuform_total_list_heading': '=== Einkaufsliste des Menüs ===',
        'menuform_item_row_label': '[{cat}] {name} ({persons} Pers.)',
        'menuform_select_recipe_to_remove': 'Wählen Sie ein Rezept des Menüs zum Entfernen aus.',
        'menuform_error_name_required': 'Bitte geben Sie einen Namen für das Menü an.',
        'menuform_error_no_recipe': 'Fügen Sie mindestens ein Rezept zum Menü hinzu.',
        'menuform_saved_message': 'Das Menü « {name} » wurde gespeichert.',
        'menuform_calculate_list_for_export': 'Berechnen Sie zuerst eine Einkaufsliste (Schaltfläche « Einkaufsliste des Menüs berechnen »).',
        'menuform_add_recipe_or_manual': 'Fügen Sie mindestens ein Rezept zum Menü hinzu, oder fügen Sie eine Zutat manuell hinzu.',
        'menuform_shopping_list_title': 'Menü: {name}',
        'menuform_print_label': 'das Menü « {name} »',
        'importurl_title': 'Rezept von einem Link importieren',
        'importurl_heading': '🌐 Rezept von einem Link importieren',
        'importurl_intro': 'Fügen Sie die Adresse (URL) einer Rezeptseite ein. Dies funktioniert\nmit den meisten großen Kochseiten (die ein Standarddatenformat verwenden). Eine Internetverbindung ist erforderlich.',
        'importurl_fetch_button': '🌐 Rezept abrufen',
        'importurl_after_import_note': 'Überprüfen und vervollständigen Sie das Rezept nach dem Import bei Bedarf\n(die Erkennung von Mengen und Einheiten ist nicht immer perfekt).',
        'importurl_paste_url_first': 'Fügen Sie zuerst eine Internetadresse (URL) ein.',
        'importurl_fetching': 'Wird abgerufen...',
        'importurl_failed_title': 'Import fehlgeschlagen',
        'importphoto_title': 'Rezept von einem Foto importieren',
        'importphoto_heading': '📷 Rezept von einem Foto importieren',
        'importphoto_intro': 'Fotografieren (oder scannen) Sie ein handgeschriebenes Rezept oder eine\nKochbuchseite, und wählen Sie dann das Bild hier aus. Der Text\nwird automatisch extrahiert, muss aber noch selbst überprüft und organisiert\nwerden (im Gegensatz zum Import von einem Link hat ein Foto\nkeine Zutaten-/Schritte-Struktur, die erraten werden kann).',
        'importphoto_module_warning': "⚠ Diese Funktion erfordert das Modul 'pytesseract'\nUND das separat auf diesem PC installierte Programm Tesseract OCR.\nSiehe LIESMICH für die Installationsanweisungen.",
        'importphoto_no_photo_chosen': 'Kein Foto ausgewählt',
        'importphoto_choose_button': '📁 Foto auswählen',
        'importphoto_extract_button': '🔍 Text extrahieren',
        'importphoto_extracted_text_label': 'Extrahierter Text (bearbeitbar):',
        'importphoto_create_button': '➡️ Rezept mit diesem Text erstellen',
        'importphoto_choose_photo_title': 'Rezeptfoto auswählen',
        'importphoto_choose_first': 'Wählen Sie zuerst ein Foto aus.',
        'importphoto_ocr_module_missing': "Diese Funktion erfordert das Modul 'pytesseract'\n(pip install pytesseract) UND das separat auf diesem PC installierte Programm Tesseract OCR. Siehe LIESMICH.",
        'importphoto_extraction_failed_title': 'Extraktion fehlgeschlagen',
        'importphoto_extraction_failed_message': 'Die Texterkennung ist fehlgeschlagen. Überprüfen Sie, ob Tesseract OCR auf diesem PC korrekt installiert und zugänglich ist.\n\nDetails: {error}',
        'importphoto_no_text_extracted': 'Aus diesem Foto konnte kein Text extrahiert werden. Versuchen Sie ein schärferes, besser gerahmtes oder besser beleuchtetes Bild.',
        'importphoto_no_text_title': 'Kein Text',
        'importphoto_no_text_confirm': 'Es wurde kein Text extrahiert oder eingegeben. Trotzdem ein leeres Rezept erstellen (nur mit dem Foto)?',
        'trash_title': 'Papierkorb',
        'trash_heading': '🗑️ Gelöschte Rezepte',
        'trash_intro': 'Fotos von Rezepten im Papierkorb werden bis zu ihrer\nendgültigen Löschung aufbewahrt.',
        'trash_restore_button': '♻️ Wiederherstellen',
        'trash_delete_forever_button': '🗑️ Endgültig löschen',
        'trash_empty_button': '🧹 Papierkorb leeren',
        'trash_unnamed_recipe': '(unbenannt)',
        'trash_unknown_date': 'unbekanntes Datum',
        'trash_entry_line': '{name}  —  gelöscht am {date}',
        'trash_is_empty': 'Der Papierkorb ist leer.',
        'trash_select_recipe_first': 'Wählen Sie ein Rezept im Papierkorb aus.',
        'trash_restored_suffix': '{name} (wiederhergestellt)',
        'trash_restored_title': 'Wiederhergestellt',
        'trash_restored_message': '« {name} » wurde wiederhergestellt.',
        'trash_delete_forever_confirm': '« {name} » endgültig löschen?\n\nDiese Aktion kann nicht rückgängig gemacht werden.',
        'trash_deleted_title': 'Gelöscht',
        'trash_deleted_message': 'Das Rezept wurde endgültig gelöscht.',
        'trash_already_empty': 'Der Papierkorb ist bereits leer.',
        'trash_empty_confirm': 'Die {count} Rezept(e) im Papierkorb endgültig löschen?\n\nDiese Aktion kann nicht rückgängig gemacht werden.',
        'trash_emptied_title': 'Papierkorb geleert',
        'trash_emptied_message': 'Der Papierkorb wurde geleert.',
        'cookingmode_title': 'Kochmodus — {name}',
        'cookingmode_close_button': '✕ Schließen (Esc)',
        'cookingmode_cooked_button': '🍳 Habe ich gekocht!',
        'cookingmode_fullscreen_hint': 'F11: Vollbild',
        'cookingmode_persons_suffix': '{persons} Pers.',
        'cookingmode_speech_button': '🔊 Vorlesen',
        'cookingmode_speech_stop_button': '⏹ Vorlesen stoppen',
        'cookingmode_volume_percent': '{percent}%',
        'cookingmode_tts_module_missing': "Das Vorlesen erfordert das Modul 'pyttsx3'.\nInstallieren Sie es mit: pip install pyttsx3",
        'cookingmode_no_description_to_read': 'Dieses Rezept hat keine Beschreibung zum Vorlesen (das Beschreibungsfeld ist leer).',
        'cookingmode_ingredients_heading': 'Zutaten',
        'cookingmode_prep_label': 'Zubereitung: {time} Min.',
        'cookingmode_cook_label': 'Kochzeit: {time} Min.',
        'cookingmode_difficulty_label': 'Schwierigkeit: {value}',
        'cookingmode_preparation_heading': 'Zubereitung',
        'cookingmode_personal_notes_heading': 'Persönliche Notizen',
        'ingsearch_title': 'Suche nach Zutat',
        'ingsearch_question_label': 'Welche Zutat suchen Sie?',
        'ingsearch_view_recipes_button': '🔍 Rezepte anzeigen, die sie verwenden',
        'ingsearch_view_selected_button': '📖 Ausgewähltes Rezept anzeigen',
        'ingsearch_no_recipe_uses': 'Derzeit verwendet kein Rezept « {name} ».',
        'ingsearch_recipes_using': 'Rezepte, die « {name} » verwenden ({count}):',
        'ingsearch_result_line': '{star}[{cat}] {name} ({qty}{unit} für 1 Person)',
        'ingsearch_select_result_first': 'Wählen Sie ein Rezept aus der Ergebnisliste aus.',
        'timerrow_minutes_label': 'Min.:',
        'timerrow_seconds_label': 'Sek.:',
        'timerrow_error_invalid_duration': 'Ungültige Dauer.',
        'timerrow_set_duration_first': 'Stellen Sie eine Dauer ein, bevor Sie starten.',
        'cooklogentry_title': '📔 Zum Kochtagebuch hinzufügen',
        'cooklogentry_heading': '🍳 « {name} »',
        'cooklogentry_intro': 'Wie war es? Eine Notiz und/oder ein Foto\n(optional, Sie können dies auch überspringen).',
        'cooklogentry_no_photo_chosen': 'Kein Foto ausgewählt',
        'cooklogentry_choose_photo_button': '📷 Foto auswählen',
        'cooklogentry_skip_button': 'Überspringen',
        'cooklogentry_choose_photo_title': 'Foto auswählen',
        'cooklog_title': '📔 Kochtagebuch — {name}',
        'cooklog_heading': '📔 {name}',
        'cooklog_times_cooked': 'Insgesamt {count} Mal gekocht',
        'cooklog_no_entry': 'Noch keine Notiz gespeichert.\nVerwenden Sie « 🍳 Habe ich gekocht! », um eine hinzuzufügen.',
        'cooklog_no_note': '(keine Notiz)',
        'timers_title': '⏲️ Timer',
        'timers_intro': 'Stellen Sie jeden Timer ein und drücken Sie dann ▶️, um ihn zu starten.\nAm Ende blinkt die Zeile rot mit einem akustischen Signal.',
        'timers_add_button': '➕ Timer hinzufügen',
        'qrcode_title': 'QR-Code — {name}',
        'qrcode_intro': 'Scannen Sie mit der Kamera oder einer App zum Lesen\nvon QR-Codes, um Name und Zutaten anzuzeigen.',
        'qrcode_save_button': '💾 Als Bild speichern (PNG)',
        'qrcode_truncated_warning': '⚠️ Das Rezept ist lang: Der QR-Code enthält eine\ngekürzte Zusammenfassung (nur Name + Zutaten).',
        'qrcode_encoded_ingredients_heading': 'Zutaten ({persons} Pers.):',
        'qrcode_save_dialog_title': 'QR-Code speichern',
        'qrcode_save_failed': 'Speichern fehlgeschlagen:\n{error}',
        'qrcode_saved_message': 'QR-Code gespeichert:\n{path}',
        'unitconv_title': 'Einheitenumrechner',
        'unitconv_heading': '🔄 Einheitenumrechner',
        'unitconv_intro': 'Ungefähre Umrechnung basierend auf der Dichte von Wasser für\nVolumeneinheiten (ml, cl, L, Tasse, Löffel): zuverlässig für\nFlüssigkeiten, ungefähr für Feststoffe wie Mehl\noder Zucker, deren tatsächliche Dichte etwas abweicht.',
        'unitconv_quantity_label': 'Menge:',
        'unitconv_from_label': 'Von:',
        'unitconv_to_label': 'Zu:',
        'unitconv_convert_button': 'Umrechnen',
        'unitconv_error_invalid_quantity': 'Ungültige Menge.',
        'unitconv_result': '{quantity} {from_unit} ≈ {result} {to_unit}',
        'unitconv_gram': 'Gramm (g)',
        'unitconv_kilogram': 'Kilogramm (kg)',
        'unitconv_ounce': 'Unze (oz)',
        'unitconv_pound': 'Pfund (lb)',
        'unitconv_milliliter': 'Milliliter (ml)',
        'unitconv_centiliter': 'Zentiliter (cl)',
        'unitconv_liter': 'Liter (L)',
        'unitconv_teaspoon': 'Teelöffel (5 ml)',
        'unitconv_tablespoon': 'Esslöffel (15 ml)',
        'unitconv_cup': 'US-Tasse (240 ml)',
        'disclaimer_title': 'Haftungsausschluss',
        'disclaimer_heading': '⚠ Haftungsausschluss',
        'disclaimer_intro': 'Bitte lesen Sie diesen Text, bevor Sie die Anwendung verwenden.',
        'disclaimer_checkbox': 'Ich habe die obigen Bedingungen gelesen und akzeptiere sie',
        'disclaimer_continue_button': 'Weiter',
        'disclaimer_quit_button': 'Anwendung beenden',
        'disclaimer_text': 'ARTIKEL 1 – HAFTUNGSAUSSCHLUSS UND -BESCHRÄNKUNG\n\n1.1. Medizinische Hinweise und Allergenverwaltung\n\nDie Anwendung bietet eine Funktion, mit der der Nutzer seine eigenen Allergie- und Allergenkriterien angeben, ändern und konfigurieren kann. Der Nutzer erkennt ausdrücklich an, dass:\n\n• Die Richtigkeit und Aktualität dieser Informationen liegt allein in seiner Verantwortung.\n• Die Anwendung ist ein Software-Hilfsmittel zum Durchsuchen von Rezepten und ersetzt in keinem Fall einen ärztlichen Rat, eine Diagnose oder die menschliche Kontrolle der Zutaten.\n• Der Herausgeber kann nicht haftbar gemacht werden für fehlerhafte Eingaben, Auslassungen, Fehlkonfigurationen durch den Nutzer oder eine allergische Reaktion (Unverträglichkeit, anaphylaktischer Schock usw.), die nach dem Verzehr eines Gerichts auftritt. Es liegt in der Verantwortung des Nutzers, systematisch die Etiketten und die tatsächliche Zusammensetzung jeder physischen Zutat vor jeder Zubereitung oder Einnahme zu überprüfen.\n\n1.2. Bereitstellung „wie besehen“ und Kostenlosigkeit\n\nDie Anwendung wird dem Nutzer vollständig kostenlos zur Verfügung gestellt. Sie wird „wie besehen“ und „je nach Verfügbarkeit“ bereitgestellt, ohne jegliche Garantie für die Fehlerfreiheit, Software-Fehler oder Unterbrechungen. Der Herausgeber garantiert nicht, dass die Funktionen der Anwendung den spezifischen Bedürfnissen des Nutzers entsprechen.\n\n1.3. Materielle und immaterielle Schäden\n\nDer Herausgeber lehnt jegliche Haftung für direkte oder indirekte Schäden ab, die dem Nutzer oder Dritten entstehen. Insbesondere kann der Herausgeber nicht belangt werden für:\n\n• Einen Ausfall, eine Überhitzung, eine Fehlfunktion oder eine Beschädigung der Computerhardware oder des Smartphones des Nutzers bei der Verwendung der Anwendung.\n• Einen Verlust von Computerdaten, eine Veränderung von Dateien oder ein Eindringen in das System des Nutzers.\n\nAufgrund der Kostenlosigkeit des Dienstes wäre, sollte die Haftung des Herausgebers durch ein Gericht festgestellt werden, der Schadensersatzbetrag ausdrücklich auf null Euro (0 €) begrenzt.',
        'addmanual_title': 'Zutaten zur Einkaufsliste hinzufügen',
        'addmanual_heading': '➕ Zutaten zur Einkaufsliste hinzufügen',
        'addmanual_intro': 'Fügen Sie so viele Zutaten wie gewünscht zur\nWarteliste unten hinzu, und bestätigen Sie sie dann alle auf einmal.',
        'addmanual_new_ingredient_button': '🥕 Neue Zutat',
        'addmanual_add_to_list_button': '➕ Zur Liste hinzufügen',
        'addmanual_staged_label': 'Zutaten, die auf Bestätigung warten:',
        'addmanual_remove_staged_button': '🗑 Von der Warteliste entfernen',
        'addmanual_confirm_all_button': '✅ Alle diese Zutaten bestätigen',
        'addmanual_close_button': 'Schließen',
        'addmanual_select_staged_first': 'Wählen Sie eine Zutat aus der Warteliste aus.',
        'addmanual_add_staged_first': 'Fügen Sie mindestens eine Zutat zur Warteliste hinzu, bevor Sie bestätigen.',
        'addmanual_confirmed_message': '{count} Zutat(en) zur Einkaufsliste hinzugefügt.',
        'shoppingexport_generated_on': 'Erstellt am {date}',
        'shoppingexport_selected_recipes': 'Ausgewählte Rezepte:',
        'shoppingexport_excel_sheet_recipes': 'Rezepte',
        'shoppingexport_excel_col_recipe': 'Rezept',
        'shoppingexport_excel_col_persons': 'Anzahl der Personen',
        'shoppingexport_excel_sheet_ingredients': 'Zutaten',
        'shoppingexport_excel_col_rayon': 'Abteilung',
        'shoppingexport_excel_col_ingredient': 'Zutat',
        'shoppingexport_excel_col_total_qty': 'Gesamtmenge',
        'shoppingexport_excel_col_unit': 'Einheit',
        'savedlists_title': 'Gespeicherte Einkaufslisten',
        'savedlists_heading': '📂 Gespeicherte Einkaufslisten',
        'savedlists_load_button': '📂 Laden',
        'savedlists_delete_button': '🗑 Löschen',
        'savedlists_none_saved': 'Noch keine Liste gespeichert.',
        'savedlists_entry_line': '{name} — {count} Zutat(en) — {date}',
        'savedlists_select_list_first': 'Wählen Sie eine Liste aus der Liste aus.',
        'savedlists_delete_confirm': 'Die Liste « {name} » endgültig löschen?',
        'quicksearch_title': 'Schnellsuche',
        'quicksearch_heading': '🔍 Schnellsuche nach Rezept',
        'quicksearch_no_results': 'Kein Rezept gefunden.',
        'quicksearch_footer_hint': 'Eingabetaste zum Öffnen, Esc zum Schließen.',
        'allrecipes_title': 'Alle Rezepte - Einkaufsliste',
        'allrecipes_select_label': 'Wählen Sie die Rezepte und die Personenanzahl aus:',
        'allrecipes_ingredient_filter_title': 'Nach Zutat filtern',
        'allrecipes_persons_count_label': 'Anz. Personen:',
        'allrecipes_add_to_cart_button': '🛒 Zum Einkauf hinzufügen',
        'allrecipes_checklist_mode_button': '☑️ Einkaufsmodus (schrittweise ankreuzen)',
        'allrecipes_clear_list_button': '🗑 Einkaufsliste leeren',
        'allrecipes_export_button': '📤 Exportieren',
        'allrecipes_print_button': '🖨️ Drucken',
        'allrecipes_add_manual_ingredient_button': '➕ Zutat zur Einkaufsliste hinzufügen',
        'allrecipes_save_list_button': '💾 Diese Liste für später speichern',
        'allrecipes_load_list_button': '📂 Gespeicherte Liste laden',
        'allrecipes_invalid_persons': 'Ungültige Personenanzahl für « {name} ».',
        'allrecipes_empty_list_message': 'Ihre Einkaufsliste ist derzeit leer.\nKlicken Sie auf « 🛒 Zum Einkauf hinzufügen » neben einem Rezept,\nfügen Sie eine Zutat manuell hinzu, oder laden Sie eine gespeicherte Liste.',
        'allrecipes_total_list_heading': '=== Gesamte Einkaufsliste ===',
        'allrecipes_manual_items_note': '({count} manuell hinzugefügte Zutat(en) enthalten)',
        'allrecipes_invalid_quantity': 'Ungültige Menge.',
        'allrecipes_calculate_list_first': 'Berechnen Sie zuerst eine Einkaufsliste, bevor Sie sie speichern.',
        'allrecipes_save_list_dialog_title': 'Liste speichern',
        'allrecipes_save_list_dialog_prompt': 'Name für diese Liste:',
        'allrecipes_list_saved_title': 'Gespeichert',
        'allrecipes_list_saved_message': 'Liste « {name} » für später gespeichert.',
        'allrecipes_empty_list_for_export': 'Die Einkaufsliste ist leer. Fügen Sie mindestens ein Rezept (Schaltfläche « 🛒 Zum Einkauf hinzufügen ») oder eine manuelle Zutat hinzu.',
        'allrecipes_export_txt_title': 'Einkaufsliste als Text speichern',
        'allrecipes_export_excel_title': 'Einkaufsliste als Excel speichern',
        'allrecipes_export_pdf_title': 'Einkaufsliste als PDF speichern',
        'allrecipes_export_saved_message': 'Einkaufsliste gespeichert:\n{path}',
        'allrecipes_excel_module_missing': "Der Excel-Export erfordert das Modul 'openpyxl'.\nInstallieren Sie es mit: pip install openpyxl",
        'allrecipes_pdf_module_missing': "Der PDF-Export erfordert das Modul 'reportlab'.\nInstallieren Sie es mit: pip install reportlab",
        'allrecipes_print_module_missing': "Das Drucken erfordert das Modul 'reportlab', um das Layout zu erstellen.\nInstallieren Sie es mit: pip install reportlab",
        'allrecipes_print_label': 'die Einkaufsliste',
        'allrecipes_shopping_list_title': 'Einkaufsliste',
        'allrecipes_close_confirm_title': 'Einkaufsliste schließen?',
        'allrecipes_close_confirm_message': 'Die angezeigte Einkaufsliste ist nicht gespeichert: Sie geht endgültig verloren, wenn Sie dieses Fenster jetzt schließen.\n\nTipp: Verwenden Sie « 💾 Diese Liste für später speichern » vor dem Schließen, wenn Sie sie behalten möchten.\n\nTrotzdem schließen?',
        'managerecipes_title': 'Rezept ändern / löschen',
        'managerecipes_select_label': 'Wählen Sie ein Rezept aus:',
        'managerecipes_filter_favorites': '⭐ Nur Favoriten',
        'managerecipes_filter_quick': '⏱️ Nur schnelle Rezepte (≤ 30 Min.)',
        'managerecipes_filter_vegetarian': '🥗 Nur vegetarische Rezepte',
        'managerecipes_filter_wishlist': '💭 Nur Wunschliste',
        'managerecipes_remove_filter_button': '✕ Filter entfernen',
        'managerecipes_search_label': '🔍 Suchen:',
        'managerecipes_sort_label': 'Sortieren nach:',
        'managerecipes_category_label': 'Kategorie:',
        'managerecipes_edit_button': '✏️ Ändern',
        'managerecipes_duplicate_button': '📋 Duplizieren',
        'managerecipes_delete_button': '🗑️ Löschen',
        'managerecipes_select_recipe_first': 'Wählen Sie ein Rezept aus der Liste aus.',
        'managerecipes_duplicate_suffix': '(Kopie)',
        'managerecipes_duplicated_title': 'Dupliziert',
        'managerecipes_duplicated_message': '« {original} » wurde unter dem Namen « {new} » dupliziert.',
        'managerecipes_delete_confirm_message': 'Das Rezept « {name} » in den Papierkorb verschieben?\n\nSie können es später über die Schaltfläche « 🗑️ Papierkorb » wiederherstellen.',
        'managerecipes_deleted_title': 'In den Papierkorb verschoben',
        'managerecipes_deleted_message': 'Das Rezept wurde in den Papierkorb verschoben.',
        'onerecipe_window_title': 'Ein Rezept anzeigen',
        'onerecipe_choose_recipe_label': 'Wählen Sie ein Rezept aus:',
        'onerecipe_search_label': '🔍 Suchen:',
        'onerecipe_sort_label': 'Sortieren:',
        'onerecipe_category_label': 'Kategorie:',
        'onerecipe_persons_label': 'Anzahl der Personen:',
        'onerecipe_btn_show': 'Rezept anzeigen',
        'onerecipe_btn_export_pdf': '📄 Als PDF exportieren',
        'onerecipe_btn_print': '🖨️ Drucken',
        'onerecipe_btn_add_to_shopping': '🛒 Zur Einkaufsliste hinzufügen',
        'onerecipe_btn_cooked': '🍳 Habe ich gekocht!',
        'onerecipe_btn_cooking_mode': '🖥️ Kochmodus (Vollbild)',
        'onerecipe_btn_qr': '📱 QR-Code',
        'onerecipe_btn_timers': '⏲️ Timer',
        'onerecipe_btn_cook_log': '📔 Kochtagebuch',
        'onerecipe_btn_substitutions': '🔄 Mögliche Ersatzzutaten',
        'onerecipe_edit_button': '✏️ Ändern',
        'onerecipe_ingredients_info_label': 'Zutaten und Informationen:',
        'onerecipe_description_notes_label': 'Beschreibung und Notizen:',
        'onerecipe_similar_label': 'Ähnliche Rezepte:',
        'onerecipe_no_photo': '(kein Foto)',
        'onerecipe_preview_unavailable': '(Vorschau nicht verfügbar)',
        'onerecipe_select_recipe_first': 'Wählen Sie ein Rezept aus der Liste aus.',
        'onerecipe_display_first': 'Zeigen Sie zuerst ein Rezept mit « Rezept anzeigen » an.',
        'onerecipe_invalid_persons': 'Ungültige Personenanzahl.',
        'onerecipe_added_to_shopping_title': 'Hinzugefügt',
        'onerecipe_added_to_shopping_message': '« {name} » ({persons} Pers.) wird beim nächsten Öffnen von « Alle Rezepte anzeigen » automatisch zur Einkaufsliste hinzugefügt.',
        'onerecipe_pantry_decrement_title': 'Vorratskammer',
        'onerecipe_pantry_decrement_prompt': 'Die Zutaten von « {name} » ({persons} Pers.) von Ihrer Vorratskammer abziehen?',
        'onerecipe_pantry_updated_title': 'Vorratskammer aktualisiert',
        'onerecipe_pantry_updated_message': '{count} Zutat(en) von Ihrer Vorratskammer abgezogen.',
        'onerecipe_pantry_none_decremented': 'Keine Zutat dieses Rezepts konnte abgezogen werden (nicht in der Vorratskammer, oder Einheit nicht vergleichbar).',
        'onerecipe_marked_title': 'Markiert',
        'onerecipe_marked_message': '« {name} » wurde heute als gekocht markiert!',
        'onerecipe_no_substitutes_title': 'Kein bekannter Ersatz',
        'onerecipe_no_substitutes_message': 'Derzeit hat keine Zutat dieses Rezepts einen bekannten Ersatz.\n\nSie können selbst einen hinzufügen über « 🥕 Zutaten verwalten » > « 🔄 Ersatzzutaten verwalten ».',
        'onerecipe_substitutes_title': 'Mögliche Ersatzzutaten — {name}',
        'onerecipe_substitutes_heading': '🔄 Mögliche Ersatzzutaten für « {name} »',
        'onerecipe_substitutes_disclaimer': 'Kulinarische Vorschläge, keine garantierten Gleichwertigkeiten:\nDas Ergebnis kann je nach Rezept variieren.',
        'onerecipe_close_button': 'Schließen',
        'onerecipe_rating_label': 'Bewertung: {stars}',
        'onerecipe_prep_label': 'Zubereitung: {time} Min.',
        'onerecipe_cook_label': 'Kochzeit: {time} Min.',
        'onerecipe_difficulty_label': 'Schwierigkeit: {value}',
        'onerecipe_allergens_label': '⚠ Allergene: {list}',
        'onerecipe_cost_label': '💰 Geschätzte Kosten: {cost} €{partial}',
        'onerecipe_cost_partial': ' (teilweise Schätzung, {known}/{total} Zutaten mit bekanntem Preis)',
        'onerecipe_nutrition_partial': ' (teilweise Schätzung, {known}/{total} erkannte Zutaten)',
        'onerecipe_nutrition_label': '🥗 Geschätzte Nährwerte{partial}:\n   {kcal} kcal · {protein} g Eiweiß · {carbs} g Kohlenhydrate · {fat} g Fett\n',
        'onerecipe_description_heading': '--- Beschreibung ---\n{text}\n',
        'onerecipe_notes_heading': '\n--- Persönliche Notizen ---\n{text}\n',
        'onerecipe_no_description_notes': '(Keine Beschreibung oder persönliche Notiz für dieses Rezept.)',
        'onerecipe_export_pdf_title': 'Rezept als PDF exportieren',
        'onerecipe_export_success_title': 'Export erfolgreich',
        'onerecipe_export_success_message': 'Rezept exportiert:\n{path}',
        'onerecipe_export_failed': 'Export fehlgeschlagen:\n{error}',
        'onerecipe_print_failed': 'Vorbereitung des Drucks fehlgeschlagen:\n{error}',
        'onerecipe_pdf_module_missing': "Der PDF-Export erfordert das Modul 'reportlab'.\nInstallieren Sie es mit: pip install reportlab",
        'onerecipe_print_module_missing': "Das Drucken erfordert das Modul 'reportlab', um das Layout zu erstellen.\nInstallieren Sie es mit: pip install reportlab",
        'onerecipe_qr_module_missing': "Der QR-Code-Export erfordert das Modul 'qrcode'.\nInstallieren Sie es mit: pip install qrcode",
        'onerecipe_qr_pillow_missing': "Der QR-Code-Export erfordert außerdem das Modul 'Pillow'.\nInstallieren Sie es mit: pip install pillow",
        'onerecipe_default_timer_label': 'Timer',
        'recipeform_title_edit': 'Rezept ändern',
        'recipeform_title_add': 'Rezept hinzufügen',
        'recipeform_name_label': 'Name des Rezepts:',
        'recipeform_favorite_checkbox': '⭐ Als Lieblingsrezept markieren',
        'recipeform_wishlist_checkbox': '💭 Zu meiner Wunschliste hinzufügen (auszuprobieren)',
        'recipeform_rating_label': 'Meine Bewertung:',
        'recipeform_category_label': 'Kategorie:',
        'recipeform_prep_time_label': 'Zubereitung (Min.):',
        'recipeform_cook_time_label': 'Kochzeit (Min.):',
        'recipeform_difficulty_label': 'Schwierigkeit:',
        'recipeform_default_persons_label': '   Standardpersonenanzahl:',
        'recipeform_tags_label': 'Tags (durch Kommas getrennt):',
        'recipeform_tags_example': 'z. B. vegetarisch, glutenfrei, schnell, günstig',
        'recipeform_allergens_label': 'Enthaltene Allergene:',
        'recipeform_detect_allergens_button': '🔍 Automatisch erkennen',
        'recipeform_allergens_disclaimer': 'Dies dient nur zur Information, überprüfen Sie immer die\nAllergene auf den Etiketten der physischen Produkte.',
        'recipeform_allergens_auto_note': 'Die automatische Erkennung basiert auf den unten bereits\neingegebenen Zutaten des Rezepts: Sie kreuzt entsprechend Felder\nan und ab, ohne jemals die Felder zu berühren, die Sie selbst\nohne Bezug zu einer erkannten Zutat angekreuzt hätten.',
        'recipeform_photos_label': 'Fotos:',
        'recipeform_add_photo_button': '📷 Foto hinzufügen',
        'recipeform_description_label': 'Beschreibung (Informationen, Schritte, Tipps...):',
        'recipeform_notes_label': 'Persönliche Notizen (Meinung, Anpassungen für nächstes Mal...):',
        'recipeform_ingredients_label': 'Zutaten (Menge für 1 Person):',
        'recipeform_new_ingredient_button': '🥕 Neue Zutat',
        'recipeform_no_ingredients_registered': 'Keine Zutat gespeichert. Klicken Sie auf « 🥕 Neue Zutat »,\num die erste anzulegen.',
        'recipeform_header_ingredient': 'Zutat',
        'recipeform_header_quantity': 'Menge',
        'recipeform_header_unit': 'Einheit',
        'recipeform_header_other': '(falls andere)',
        'recipeform_add_ingredient_button': '+ Zutat hinzufügen',
        'recipeform_save_button': 'Speichern',
        'recipeform_delete_button': 'Dieses Rezept löschen',
        'recipeform_char_counter': '{count} / {max} Zeichen',
        'recipeform_add_ingredients_first': 'Fügen Sie zuerst Zutaten zum Rezept hinzu.',
        'recipeform_allergens_updated_title': 'Allergene aktualisiert',
        'recipeform_allergens_updated_added': 'hinzugefügt: {list}',
        'recipeform_allergens_updated_removed': 'entfernt: {list}',
        'recipeform_allergens_updated_message': 'Allergen(e) {parts}.',
        'recipeform_allergens_no_change': 'Keine Änderung: Die angekreuzten Allergene entsprechen bereits den Zutaten.',
        'recipeform_choose_photos_title': 'Ein oder mehrere Fotos auswählen',
        'recipeform_no_photo': '(kein Foto)',
        'recipeform_preview_unavailable': '(Vorschau\nnicht verfügbar)',
        'recipeform_remove_photo_button': '🗑 Entfernen',
        'recipeform_new_ingredient_dialog_title': 'Neue Zutat',
        'recipeform_new_ingredient_dialog_prompt': 'Name der neuen Zutat:',
        'recipeform_ingredient_already_exists': 'Die Zutat « {name} » existiert bereits.',
        'recipeform_ingredient_added_title': 'Hinzugefügt',
        'recipeform_ingredient_added_message': 'Die Zutat « {name} » wurde hinzugefügt.\nWählen Sie sie in einer der Dropdown-Listen aus.',
        'recipeform_error_name_required': 'Bitte geben Sie einen Namen für das Rezept an.',
        'recipeform_error_prep_time': 'Die Zubereitungszeit muss eine positive Zahl sein (oder leer).',
        'recipeform_error_cook_time': 'Die Kochzeit muss eine positive Zahl sein (oder leer).',
        'recipeform_unknown_ingredient_title': 'Unbekannte Zutat',
        'recipeform_unknown_ingredient_message': '« {name} » entspricht keiner gespeicherten Zutat.\nWählen Sie eine aus der Dropdown-Liste aus, oder klicken Sie auf « 🥕 Neue Zutat », um sie zuerst hinzuzufügen.',
        'recipeform_error_invalid_quantity': "Ungültige Menge für '{name}'.",
        'recipeform_error_custom_unit_required': "Geben Sie die benutzerdefinierte Einheit für '{name}' an.",
        'recipeform_error_no_valid_ingredient': 'Fügen Sie mindestens eine gültige Zutat hinzu.',
        'recipeform_duplicate_ingredient_title': 'Doppelte Zutat',
        'recipeform_duplicate_ingredient_message': '„{list}“ kommt in diesem Rezept mehrmals vor.\n\nTrotzdem speichern?',
        'recipeform_saved_message': 'Das Rezept « {name} » wurde gespeichert.',
        'recipeform_delete_confirm_message': 'Das Rezept « {name} » in den Papierkorb verschieben?\n\nSie können es später über die Schaltfläche « 🗑️ Papierkorb » wiederherstellen.',
        'recipeform_deleted_title': 'In den Papierkorb verschoben',
        'recipeform_deleted_message': 'Das Rezept wurde in den Papierkorb verschoben.',
    },
}


def detect_system_language():
    """Détecte la langue du système d'exploitation pour proposer une
    langue de démarrage sensée au tout premier lancement (avant qu'aucune
    préférence n'ait jamais été enregistrée). Ne reconnaît que le
    français, l'anglais, l'espagnol et l'allemand pour l'instant (les
    seules langues disponibles) : toute autre langue système retombe sur
    le français, la langue de référence de l'application. Repose sur
    locale.getlocale()/getdefaultlocale(), qui peuvent échouer ou
    renvoyer None selon la configuration du système — dans ce cas, on
    retombe aussi sur le français plutôt que de risquer une erreur au
    démarrage."""
    try:
        import locale
        lang_code = None
        try:
            lang_code, _ = locale.getlocale()
        except (ValueError, TypeError):
            pass
        if not lang_code:
            try:
                lang_code, _ = locale.getdefaultlocale()
            except (ValueError, TypeError):
                pass
        if lang_code:
            lang_code_lower = lang_code.lower()
            if lang_code_lower.startswith("en"):
                return "en"
            if lang_code_lower.startswith("es"):
                return "es"
            if lang_code_lower.startswith("de"):
                return "de"
    except Exception:
        pass
    return "fr"


def get_language_preference():
    settings = load_settings()
    if "language" not in settings:
        # Premier lancement : aucune préférence enregistrée pour l'instant,
        # on propose la langue du système plutôt que de toujours démarrer
        # en français. Ce choix devient ensuite la préférence enregistrée,
        # comme si l'utilisateur l'avait choisie lui-même (bascule normale
        # possible à tout moment via le bouton 🌐).
        detected = detect_system_language()
        settings["language"] = detected
        save_settings(settings)
        return detected
    return settings.get("language", "fr")


def set_language_preference(value):
    settings = load_settings()
    settings["language"] = value
    save_settings(settings)


def apply_language(lang):
    global CURRENT_LANGUAGE
    CURRENT_LANGUAGE = lang


def t(key, **kwargs):
    """Retourne le texte de l'interface pour la clé donnée, dans la langue
    actuellement sélectionnée. Si cette clé n'a pas encore été traduite
    dans la langue courante, retourne le texte français de référence — la
    partie pas encore traduite reste donc lisible plutôt que d'afficher un
    identifiant technique."""
    text = None
    if CURRENT_LANGUAGE != "fr":
        text = TRANSLATIONS.get(CURRENT_LANGUAGE, {}).get(key)
    if text is None:
        text = FRENCH_STRINGS.get(key, key)
    if kwargs:
        try:
            return text.format(**kwargs)
        except (KeyError, IndexError, ValueError):
            return text
    return text


def get_disclaimer_accepted():
    return bool(load_settings().get("disclaimer_accepted", False))


def set_disclaimer_accepted(value):
    settings = load_settings()
    settings["disclaimer_accepted"] = bool(value)
    save_settings(settings)


class DisclaimerWindow(tk.Toplevel):
    """Clause de responsabilité affichée obligatoirement au tout premier
    lancement de l'application. Tant qu'elle n'est pas acceptée (case cochée
    puis bouton « Continuer »), l'application ne peut pas être utilisée."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.accepted = False
        # La hauteur est plafonnée à l'espace écran réellement disponible
        # (moins une petite marge), pour que le bouton « Continuer » reste
        # toujours visible même sur un écran de petite hauteur ou avec le
        # mode « Texte agrandi » déjà activé lors d'une session précédente.
        window_height = min(gs(660), get_usable_screen_height(self) - 40)
        self.geometry(f"{gs(640)}x{window_height}")
        self.minsize(gs(480), min(gs(460), window_height))
        self.resizable(True, True)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._quit_app)
        self._build_ui()

    def _build_ui(self):
        # Reconstruit tout le contenu de la fenêtre dans la langue
        # actuellement sélectionnée. Appelée à l'ouverture, puis à chaque
        # fois que la langue est changée via le sélecteur ci-dessous —
        # avant même d'avoir accepté les conditions, une personne
        # anglophone ou hispanophone peut ainsi lire ce texte dans sa
        # propre langue dès le premier lancement.
        previously_checked = self.accept_var.get() if hasattr(self, "accept_var") else False
        for child in self.winfo_children():
            child.destroy()
        self.title(t("disclaimer_title"))

        # ---- Sélecteur de langue, tout en haut : mêmes principe et
        # habillage que le menu déroulant de la page d'accueil. ----
        lang_bar = ttk.Frame(self)
        lang_bar.pack(fill="x", padx=15, pady=(10, 0))
        language_names = {"fr": "Français", "en": "English", "es": "Español", "de": "Deutsch"}
        current_flag = self.app.flag_photos.get(self.app.language)
        menubutton_kwargs = {"text": language_names.get(self.app.language, "Français")}
        if current_flag is not None:
            menubutton_kwargs["image"] = current_flag
            menubutton_kwargs["compound"] = "left"
        else:
            menubutton_kwargs["text"] = "🌐 " + menubutton_kwargs["text"]
        language_menubutton = ttk.Menubutton(lang_bar, style="Secondary.TMenubutton", **menubutton_kwargs)
        language_menu = tk.Menu(language_menubutton, tearoff=False)
        for lang_code in ("fr", "en", "es", "de"):
            item_kwargs = {
                "label": language_names[lang_code],
                "command": lambda lc=lang_code: self._set_language(lc),
            }
            lang_flag = self.app.flag_photos.get(lang_code)
            if lang_flag is not None:
                item_kwargs["image"] = lang_flag
                item_kwargs["compound"] = "left"
            language_menu.add_command(**item_kwargs)
        language_menubutton["menu"] = language_menu
        language_menubutton.pack(side="right")

        ttk.Label(self, text=t("disclaimer_heading"), font=("Segoe UI", sf(14), "bold"),
                  foreground=COLOR_ERROR).pack(pady=(15, 5))
        ttk.Label(self, text=t("disclaimer_intro"),
                  font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED).pack(pady=(0, 10))

        text_frame = ttk.Frame(self)
        text_frame.pack(fill="both", expand=True, padx=15)
        text_widget = tk.Text(text_frame, wrap="word", padx=10, pady=10, font=("Segoe UI", sf(10)))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        text_widget.insert("1.0", t("disclaimer_text"))
        text_widget.config(state="disabled")

        self.accept_var = tk.BooleanVar(value=previously_checked)
        check = ttk.Checkbutton(
            self, text=t("disclaimer_checkbox"),
            variable=self.accept_var, command=self._on_toggle
        )
        check.pack(pady=(12, 5))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=(0, 15))
        self.continue_button = ttk.Button(
            btn_frame, text=t("disclaimer_continue_button"),
            state="normal" if previously_checked else "disabled", command=self._on_continue
        )
        self.continue_button.grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("disclaimer_quit_button"),
                   style="Secondary.TButton", command=self._quit_app).grid(row=0, column=1, padx=5)

    def _set_language(self, lang):
        if lang == self.app.language:
            return
        self.app.language = lang
        set_language_preference(lang)
        apply_language(lang)
        self._build_ui()

    def _on_toggle(self):
        self.continue_button.config(state="normal" if self.accept_var.get() else "disabled")

    def _on_continue(self):
        if not self.accept_var.get():
            return
        self.accepted = True
        set_disclaimer_accepted(True)
        self.destroy()

    def _quit_app(self):
        self.accepted = False
        self.destroy()
        self.app.destroy()
        sys.exit(0)


# Espace vide laissé en bas des listes défilantes (environ 2 cm à 96 DPI),
# pour pouvoir descendre l'ascenseur un peu plus bas que le dernier élément
# et le voir entièrement, même si la fenêtre se termine juste au-dessus de
# la barre des tâches.
SCROLL_BOTTOM_PADDING = 76


def get_usable_screen_height(widget):
    """Retourne la hauteur d'écran réellement utilisable (écran total moins
    la barre des tâches Windows), pour qu'une fenêtre réglée à la hauteur de
    l'écran ne se retrouve jamais partiellement masquée derrière elle. Sur
    les systèmes où cette information n'est pas disponible (macOS, Linux, ou
    en cas d'erreur), retombe simplement sur la hauteur d'écran totale."""
    try:
        import ctypes

        class RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

        rect = RECT()
        SPI_GETWORKAREA = 0x0030
        if ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0):
            height = rect.bottom - rect.top
            if height > 0:
                return height
    except Exception:
        pass
    return widget.winfo_screenheight()


def configure_app_style(root):
    """Configure l'apparence de toute l'application : couleurs ttk (boutons,
    champs, onglets...) ainsi que les widgets Tkinter bruts (Listbox, Text,
    Canvas...) via la base d'options de Tk, qui s'applique automatiquement à
    tous les widgets créés ensuite dans n'importe quelle fenêtre."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    base_font = ("Segoe UI", sf(10))

    style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=base_font)
    style.configure("TFrame", background=COLOR_BG)
    style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT)
    style.configure("TCheckbutton", background=COLOR_BG, foreground=COLOR_TEXT)
    style.map("TCheckbutton", background=[("active", COLOR_BG)])

    style.configure("TButton", background=COLOR_ACCENT, foreground="white",
                     font=base_font, padding=(12, 7), borderwidth=0, relief="flat")
    style.map("TButton",
              background=[("active", COLOR_ACCENT_DARK), ("disabled", "#D8CBB8")],
              foreground=[("disabled", "#F4EEE3")])

    style.configure("TEntry", fieldbackground=COLOR_CARD, foreground=COLOR_TEXT,
                     bordercolor=COLOR_BORDER, lightcolor=COLOR_BORDER, darkcolor=COLOR_BORDER,
                     padding=5)
    style.configure("TCombobox", fieldbackground=COLOR_CARD, foreground=COLOR_TEXT,
                     bordercolor=COLOR_BORDER, arrowcolor=COLOR_ACCENT_DARK, padding=5)
    style.map("TCombobox", fieldbackground=[("readonly", COLOR_CARD)])

    style.configure("TLabelframe", background=COLOR_BG, bordercolor=COLOR_BORDER)
    style.configure("TLabelframe.Label", background=COLOR_BG, foreground=COLOR_ACCENT_DARK,
                     font=("Segoe UI", sf(10), "bold"))

    style.configure("TNotebook", background=COLOR_BG, bordercolor=COLOR_BORDER)
    style.configure("TNotebook.Tab", background=COLOR_ACCENT_LIGHT, foreground=COLOR_TEXT,
                     padding=(12, 6), font=base_font)
    style.map("TNotebook.Tab",
              background=[("selected", COLOR_ACCENT)],
              foreground=[("selected", "white")])

    style.configure("TScrollbar", background=COLOR_ACCENT, troughcolor=COLOR_BG,
                     bordercolor=COLOR_BG, arrowcolor="white")
    style.map("TScrollbar", background=[("active", COLOR_ACCENT_DARK)])

    style.configure("TSeparator", background=COLOR_BORDER)

    style.configure("TPanedwindow", background=COLOR_BG)

    # Boutons secondaires plus discrets (utilisés pour des actions annexes) :
    # à activer au cas par cas avec style="Secondary.TButton" si besoin plus tard.
    style.configure("Secondary.TButton", background=COLOR_CARD, foreground=COLOR_ACCENT_DARK,
                     padding=(12, 7))
    style.map("Secondary.TButton", background=[("active", COLOR_ACCENT_LIGHT)])

    # Menu déroulant de langue : même habillage visuel que les boutons
    # secondaires ci-dessus, pour rester cohérent dans la barre du haut.
    style.configure("Secondary.TMenubutton", background=COLOR_CARD, foreground=COLOR_ACCENT_DARK,
                     padding=(12, 7))
    style.map("Secondary.TMenubutton", background=[("active", COLOR_ACCENT_LIGHT)])

    # Variantes "carte" (fond blanc) pour les encadrés mis en valeur sur la
    # page d'accueil, afin que les widgets ttk placés dedans (Label, Button)
    # aient le même fond que la carte plutôt que le fond général de la page.
    style.configure("Card.TFrame", background=COLOR_CARD)
    style.configure("Card.TLabel", background=COLOR_CARD, foreground=COLOR_TEXT)

    # ---- Widgets Tkinter bruts (non gérés par ttk) : Listbox, Text, Canvas,
    # Button (celles en tk.Button, ex. mode cuisine), via la base d'options.
    # S'applique à toute fenêtre créée dans l'application, sans exception. ----
    root.option_add("*Font", "{Segoe UI} 10")
    root.option_add("*Background", COLOR_BG)
    root.option_add("*Foreground", COLOR_TEXT)

    root.option_add("*Listbox.Background", COLOR_CARD)
    root.option_add("*Listbox.Foreground", COLOR_TEXT)
    root.option_add("*Listbox.selectBackground", COLOR_ACCENT)
    root.option_add("*Listbox.selectForeground", "white")
    root.option_add("*Listbox.borderWidth", 1)
    root.option_add("*Listbox.relief", "solid")
    root.option_add("*Listbox.highlightThickness", 0)

    root.option_add("*Text.Background", COLOR_CARD)
    root.option_add("*Text.Foreground", COLOR_TEXT)
    root.option_add("*Text.borderWidth", 1)
    root.option_add("*Text.relief", "solid")
    root.option_add("*Text.highlightThickness", 0)

    root.option_add("*Canvas.Background", COLOR_BG)
    root.option_add("*Canvas.highlightThickness", 0)

    root.option_add("*Button.Background", COLOR_ACCENT)
    root.option_add("*Button.Foreground", "white")
    root.option_add("*Button.activeBackground", COLOR_ACCENT_DARK)
    root.option_add("*Button.activeForeground", "white")
    root.option_add("*Button.relief", "flat")
    root.option_add("*Button.borderWidth", 0)
    root.option_add("*Button.padX", 10)
    root.option_add("*Button.padY", 5)

    root.option_add("*Entry.Background", COLOR_CARD)
    root.option_add("*Entry.Foreground", COLOR_TEXT)
    root.option_add("*Entry.relief", "solid")
    root.option_add("*Entry.borderWidth", 1)

    return style


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(t("home_window_title"))
        # Le mode « Texte agrandi » doit être chargé et appliqué AVANT tout
        # calcul de géométrie ci-dessous (via gs()), sans quoi la fenêtre
        # s'ouvrirait à sa taille normale au premier lancement même si ce
        # mode était déjà activé lors d'une session précédente.
        self.large_text = get_large_text_preference()
        apply_font_scale(self.large_text)
        self.language = get_language_preference()
        apply_language(self.language)
        # Icônes de drapeaux pour le bouton de langue, chargées une seule
        # fois ici (et non à chaque reconstruction de la page d'accueil)
        # pour éviter de relire le fichier à chaque bascule. Une référence
        # doit être conservée sur self, sans quoi Tkinter "oublierait"
        # l'image (garbage collection) et le bouton se retrouverait sans
        # icône. Si un fichier est absent (ex. copie incomplète de
        # l'application), le bouton reste utilisable, juste sans icône
        # pour cette langue-là.
        self.flag_photos = {}
        for lang_code, flag_path in FLAG_FILES.items():
            try:
                if os.path.exists(flag_path):
                    self.flag_photos[lang_code] = tk.PhotoImage(file=flag_path)
            except tk.TclError:
                pass
        # La hauteur de la fenêtre correspond à la hauteur de l'écran sur
        # lequel l'application est lancée, pour profiter de tout l'espace
        # vertical disponible dès le démarrage. La largeur est fixée pour
        # accueillir les boutons et les cartes "Aujourd'hui"/"Récemment
        # consultées" côte à côte.
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(1000)}x{screen_height}+40+0")
        self.minsize(gs(720), gs(500))
        self.resizable(True, True)

        self.dark_mode = get_dark_mode_preference()
        apply_palette(self.dark_mode)
        configure_app_style(self)
        self.configure(background=COLOR_BG)

        # ---- Clause de responsabilité : obligatoire au tout premier
        # lancement, l'application reste inutilisable tant qu'elle n'est
        # pas acceptée. ----
        if not get_disclaimer_accepted():
            self.withdraw()
            disclaimer = DisclaimerWindow(self)
            self.wait_window(disclaimer)
            if not disclaimer.accepted:
                return  # _quit_app() a déjà fermé l'application (sys.exit)
            self.deiconify()

        self.recipes = load_recipes()
        self.ingredient_names = sync_ingredients_from_recipes()
        self.shopping_selection = {}  # sélection en cours pour la liste de courses (nom -> personnes)
        self.timers_window = None  # fenêtre unique des minuteurs, créée à la demande

        # Recherche rapide de recette (Ctrl+K), accessible depuis n'importe
        # quelle fenêtre de l'application, y compris par-dessus une fenêtre
        # modale (bind_all s'applique quelle que soit la fenêtre au premier
        # plan).
        self.bind_all("<Control-k>", lambda e: self.open_quick_search())

        # Sauvegarde automatique périodique (silencieuse, ne bloque jamais le démarrage)
        try:
            maybe_create_auto_backup()
        except Exception:
            pass

        self._build_home_ui()

    def open_quick_search(self):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        QuickSearchWindow(self)

    def open_donate_page(self):
        webbrowser.open("https://buymeacoffee.com/majogari")

    def toggle_dark_mode(self):
        """Bascule entre thème clair et sombre : met à jour la palette, les
        styles ttk (effet immédiat sur toutes les fenêtres déjà ouvertes),
        puis reconstruit entièrement la page d'accueil pour que ses widgets
        Tkinter bruts (bannière, cartes...) reflètent aussi les nouvelles
        couleurs. Les fenêtres secondaires déjà ouvertes devront être
        refermées puis rouvertes pour refléter pleinement le nouveau thème."""
        self.dark_mode = not self.dark_mode
        set_dark_mode_preference(self.dark_mode)
        apply_palette(self.dark_mode)
        configure_app_style(self)
        self.configure(background=COLOR_BG)
        try:
            self.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        # Ne détruit que les widgets propres à la page d'accueil : les
        # fenêtres secondaires (Toplevel) déjà ouvertes — comme les
        # minuteurs en cours — ne doivent surtout pas être affectées.
        for child in self.winfo_children():
            if not isinstance(child, tk.Toplevel):
                child.destroy()
        self._build_home_ui()

    def toggle_large_text(self):
        """Bascule le mode « Texte agrandi » (accessibilité, malvoyants) :
        met à jour l'échelle globale des polices et reconstruit la page
        d'accueil pour l'appliquer immédiatement. Les fenêtres secondaires
        déjà ouvertes gardent leur taille de police d'origine — fermez-les
        et rouvrez-les pour qu'elles s'affichent avec le nouveau réglage,
        y compris leur propre taille de fenêtre, recalculée en
        conséquence pour ne rien couper ni masquer."""
        self.large_text = not self.large_text
        set_large_text_preference(self.large_text)
        apply_font_scale(self.large_text)
        configure_app_style(self)
        try:
            self.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        for child in self.winfo_children():
            if not isinstance(child, tk.Toplevel):
                child.destroy()
        self._build_home_ui()

    def set_language(self, lang):
        """Change la langue de l'interface vers celle choisie directement
        dans le menu déroulant (français, anglais ou espagnol — d'autres
        langues pourront être ajoutées de la même façon). Les parties de
        l'interface pas encore traduites dans la langue choisie restent
        affichées en français (voir t())."""
        if lang == self.language:
            return
        self.language = lang
        set_language_preference(self.language)
        apply_language(self.language)
        try:
            self.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        for child in self.winfo_children():
            if not isinstance(child, tk.Toplevel):
                child.destroy()
        self._build_home_ui()

    def _build_home_ui(self):
        # Efface l'éventuel contenu déjà construit, pour que cette méthode
        # puisse être rappelée en toute sécurité (ex. après avoir modifié le
        # garde-manger) sans dupliquer boutons et bannières.
        for child in self.winfo_children():
            child.destroy()

        # ---- Barre supérieure fixe (hors zone de défilement) : bouton de
        # don en haut à gauche, bascule de thème en haut à droite — tous
        # deux toujours visibles quel que soit le défilement. ----
        top_bar = tk.Frame(self, background=COLOR_BG)
        top_bar.pack(fill="x")
        ttk.Button(top_bar, text=t("home_donate_button"), style="Secondary.TButton",
                   command=self.open_donate_page).pack(side="left", padx=10, pady=8)
        toggle_text = t("home_light_theme") if self.dark_mode else t("home_dark_theme")
        ttk.Button(top_bar, text=toggle_text, style="Secondary.TButton",
                   command=self.toggle_dark_mode).pack(side="right", padx=10, pady=8)
        large_text_label = t("home_large_text_off") if self.large_text else t("home_large_text_on")
        ttk.Button(top_bar, text=large_text_label, style="Secondary.TButton",
                   command=self.toggle_large_text).pack(side="right", padx=(10, 0), pady=8)
        # Menu déroulant de langue : affiche la langue actuellement
        # sélectionnée (avec son propre drapeau), et propose un choix
        # direct des 3 langues disponibles plutôt qu'un simple cycle —
        # plus explicite, et on peut choisir n'importe laquelle en un
        # seul clic sans repasser par les autres.
        language_names = {"fr": "Français", "en": "English", "es": "Español", "de": "Deutsch"}
        current_flag = self.flag_photos.get(self.language)
        menubutton_kwargs = {"text": language_names.get(self.language, "Français")}
        if current_flag is not None:
            menubutton_kwargs["image"] = current_flag
            menubutton_kwargs["compound"] = "left"
        else:
            # Repli textuel si le fichier de drapeau est absent, plutôt
            # que d'afficher un bouton sans aucune indication de langue.
            menubutton_kwargs["text"] = "🌐 " + menubutton_kwargs["text"]
        language_menubutton = ttk.Menubutton(top_bar, style="Secondary.TMenubutton", **menubutton_kwargs)
        language_menu = tk.Menu(language_menubutton, tearoff=False)
        for lang_code in ("fr", "en", "es", "de"):
            item_kwargs = {
                "label": language_names[lang_code],
                "command": lambda lc=lang_code: self.set_language(lc),
            }
            lang_flag = self.flag_photos.get(lang_code)
            if lang_flag is not None:
                item_kwargs["image"] = lang_flag
                item_kwargs["compound"] = "left"
            language_menu.add_command(**item_kwargs)
        language_menubutton["menu"] = language_menu
        language_menubutton.pack(side="right", padx=(0, 0), pady=8)

        # ---- Conteneur scrollable pour toute la page d'accueil : ainsi, quel
        # que soit le nombre de boutons ou la taille de la fenêtre, rien ne
        # peut jamais être coupé ou invisible. Le pied de page reste fixé en
        # bas, en dehors de la zone qui défile. ----
        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        content = ttk.Frame(canvas)
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content, anchor="n")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)

        banner = tk.Frame(content, background=COLOR_ACCENT)
        banner.pack(fill="x")
        tk.Label(banner, text=t("home_banner_title"), font=("Segoe UI", sf(21), "bold"),
                 background=COLOR_ACCENT, foreground="white").pack(pady=(18, 2))
        tk.Label(banner, text=t("home_banner_subtitle"),
                 font=("Segoe UI", sf(10)), background=COLOR_ACCENT,
                 foreground=COLOR_ACCENT_LIGHT).pack(pady=(0, 16))

        # ---- Recette du jour : tirée au sort une fois par jour (mémorisée
        # dans settings.json), pas à chaque ouverture de l'application. ----
        daily_recipe = get_daily_recipe(self.recipes)
        if daily_recipe is not None:
            daily_card = tk.Frame(content, background=COLOR_CARD, highlightbackground=COLOR_BORDER,
                                   highlightthickness=1, cursor="hand2")
            daily_card.pack(padx=20, pady=(15, 0), fill="x")
            daily_inner = ttk.Frame(daily_card, style="Card.TFrame")
            daily_inner.pack(fill="x", padx=15, pady=10)
            star = "⭐ " if daily_recipe.get("favorite") else ""
            cat = translate_category_name(daily_recipe.get("category", "Autre"))
            ttk.Label(daily_inner, text=t("home_daily_recipe_title"), font=("Segoe UI", sf(11), "bold"),
                      style="Card.TLabel", foreground=COLOR_ACCENT_DARK).pack(side="left")
            ttk.Label(daily_inner, text=f"{star}[{cat}] {daily_recipe['name']}",
                      style="Card.TLabel", font=("Segoe UI", sf(11))).pack(side="left", padx=(15, 0))
            ttk.Button(daily_inner, text=t("home_open_button"),
                       command=lambda r=daily_recipe: self._open_daily_recipe(r)).pack(side="right")

        # ---- Filtres rapides : accès direct à une liste déjà filtrée ----
        quick_filters_frame = ttk.Frame(content)
        quick_filters_frame.pack(padx=20, pady=(15, 0), fill="x")
        ttk.Button(quick_filters_frame, text=t("home_quick_filter_favorites"), style="Secondary.TButton",
                   command=lambda: self.open_manage_recipes(quick_filter="favoris")).pack(
            side="left", expand=True, fill="x", padx=(0, 4))
        ttk.Button(quick_filters_frame, text=t("home_quick_filter_quick"), style="Secondary.TButton",
                   command=lambda: self.open_manage_recipes(quick_filter="rapide")).pack(
            side="left", expand=True, fill="x", padx=4)
        ttk.Button(quick_filters_frame, text=t("home_quick_filter_vegetarian"), style="Secondary.TButton",
                   command=lambda: self.open_manage_recipes(quick_filter="vegetarien")).pack(
            side="left", expand=True, fill="x", padx=4)
        ttk.Button(quick_filters_frame, text=t("home_quick_filter_wishlist"), style="Secondary.TButton",
                   command=lambda: self.open_manage_recipes(quick_filter="envie")).pack(
            side="left", expand=True, fill="x", padx=(4, 0))

        # ---- Rappel : recettes en liste d'envies depuis longtemps ----
        WISHLIST_REMINDER_DAYS = 90
        stale_wishlist = []
        for r in self.recipes:
            if r.get("wishlist") and r.get("wishlist_since"):
                try:
                    since = datetime.fromisoformat(r["wishlist_since"])
                    if (datetime.now() - since).days >= WISHLIST_REMINDER_DAYS:
                        stale_wishlist.append(r)
                except (ValueError, TypeError):
                    pass
        if stale_wishlist:
            reminder_frame = tk.Frame(content, background=COLOR_ACCENT_LIGHT,
                                       highlightbackground=COLOR_BORDER, highlightthickness=1, cursor="hand2")
            reminder_frame.pack(padx=20, pady=(10, 0), fill="x")
            reminder_label = tk.Label(
                reminder_frame,
                text=t("home_wishlist_reminder", count=len(stale_wishlist), days=WISHLIST_REMINDER_DAYS),
                background=COLOR_ACCENT_LIGHT, foreground=COLOR_ACCENT_DARK, font=("Segoe UI", sf(9), "bold"),
                wraplength=560, justify="center", cursor="hand2"
            )
            reminder_label.pack(padx=10, pady=8)
            reminder_frame.bind("<Button-1>", lambda e: self.open_manage_recipes(quick_filter="envie"))
            reminder_label.bind("<Button-1>", lambda e: self.open_manage_recipes(quick_filter="envie"))

        # ---- Rappel : articles du garde-manger sous leur seuil d'alerte ----
        low_stock = get_low_stock_pantry_items()
        if low_stock:
            stock_frame = tk.Frame(content, background=COLOR_ACCENT_LIGHT,
                                    highlightbackground=COLOR_BORDER, highlightthickness=1, cursor="hand2")
            stock_frame.pack(padx=20, pady=(10, 0), fill="x")
            stock_names = ", ".join(sorted(e["name"] for e in low_stock))
            stock_label = tk.Label(
                stock_frame,
                text=t("home_low_stock_reminder", count=len(low_stock), names=stock_names),
                background=COLOR_ACCENT_LIGHT, foreground=COLOR_ACCENT_DARK, font=("Segoe UI", sf(9), "bold"),
                wraplength=560, justify="center", cursor="hand2"
            )
            stock_label.pack(padx=10, pady=8)
            stock_frame.bind("<Button-1>", lambda e=None, items=low_stock: self._open_low_stock_to_cart(items))
            stock_label.bind("<Button-1>", lambda e=None, items=low_stock: self._open_low_stock_to_cart(items))

        # ---- Deux colonnes côte à côte : les boutons à gauche, les cartes
        # "Aujourd'hui" et "Récemment consultées" à droite. ----
        columns_frame = ttk.Frame(content)
        columns_frame.pack(padx=15, pady=(15, 0), fill="x")

        left_col = ttk.Frame(columns_frame)
        left_col.pack(side="left", fill="y", padx=(0, 15), anchor="n")

        right_col = ttk.Frame(columns_frame)
        right_col.pack(side="left", fill="both", expand=True, anchor="n")

        grid_frame = ttk.Frame(left_col)
        grid_frame.pack(fill="x")
        grid_frame.columnconfigure(0, weight=1)

        buttons = [
            (t("home_btn_add_recipe"), self.open_add_recipe),
            (t("home_btn_import_url"), self.open_import_from_url),
            (t("home_btn_import_photo"), self.open_import_from_photo),
            (t("home_btn_view_all_recipes"), self.open_all_recipes),
            (t("home_btn_view_one_recipe"), self.open_one_recipe),
            (t("home_btn_manage_recipes"), self.open_manage_recipes),
            (t("home_btn_compare_recipes"), self.open_compare_recipes),
            (t("home_btn_manage_ingredients"), self.open_manage_ingredients),
            (t("home_btn_ingredient_search"), self.open_ingredient_search),
            (t("home_btn_what_can_i_cook"), self.open_what_can_i_cook),
            (t("home_btn_pantry"), self.open_pantry),
            (t("home_btn_unit_converter"), self.open_unit_converter),
            (t("home_btn_weekly_plan"), self.open_weekly_plan),
            (t("home_btn_menus"), self.open_menus),
            (t("home_btn_statistics"), self.open_statistics),
            (t("home_btn_export_cookbook"), self.open_cookbook_export),
            (t("home_btn_import_export"), self.open_import_export),
            (t("home_btn_trash"), self.open_trash),
        ]
        for i, (text, command) in enumerate(buttons):
            ttk.Button(grid_frame, text=text, command=command).grid(
                row=i, column=0, padx=4, pady=4, sticky="ew"
            )

        # ---- Repas du jour ----
        today_card = tk.Frame(right_col, background=COLOR_CARD, highlightbackground=COLOR_BORDER,
                               highlightthickness=1)
        today_card.pack(fill="x")
        ttk.Label(today_card, text=t("home_today_title"), font=("Segoe UI", sf(12), "bold"),
                  style="Card.TLabel", foreground=COLOR_ACCENT_DARK).pack(anchor="w", padx=12, pady=(10, 4))
        self.today_frame = ttk.Frame(today_card, style="Card.TFrame")
        self.today_frame.pack(padx=12, pady=(0, 12), fill="x")
        self._refresh_today_meals()

        # ---- Récemment consultées ----
        recent_card = tk.Frame(right_col, background=COLOR_CARD, highlightbackground=COLOR_BORDER,
                                highlightthickness=1)
        recent_card.pack(fill="x", pady=(15, 0))
        ttk.Label(recent_card, text=t("home_recent_title"), font=("Segoe UI", sf(12), "bold"),
                  style="Card.TLabel", foreground=COLOR_ACCENT_DARK).pack(anchor="w", padx=12, pady=(10, 4))
        recent_frame = ttk.Frame(recent_card, style="Card.TFrame")
        recent_frame.pack(padx=12, pady=(0, 12), fill="x")
        self.recent_listbox = tk.Listbox(recent_frame, height=5, font=("Segoe UI", sf(9)))
        self.recent_listbox.pack(side="left", fill="x", expand=True)
        self.recent_listbox.bind("<Double-Button-1>", lambda e: self.open_recent_selected())
        ttk.Button(recent_frame, text=t("home_open_button"), command=self.open_recent_selected).pack(
            side="left", padx=(8, 0))
        self._refresh_recent_views()

        # ---- Recettes à essayer (liste d'envies), tirage au sort ----
        wishlist_card = tk.Frame(right_col, background=COLOR_CARD, highlightbackground=COLOR_BORDER,
                                  highlightthickness=1)
        wishlist_card.pack(fill="x", pady=(15, 0))
        ttk.Label(wishlist_card, text=t("home_wishlist_title"), font=("Segoe UI", sf(12), "bold"),
                  style="Card.TLabel", foreground=COLOR_ACCENT_DARK).pack(anchor="w", padx=12, pady=(10, 4))
        wishlist_frame = ttk.Frame(wishlist_card, style="Card.TFrame")
        wishlist_frame.pack(padx=12, pady=(0, 12), fill="x")
        self.wishlist_listbox = tk.Listbox(wishlist_frame, height=8, font=("Segoe UI", sf(9)))
        self.wishlist_listbox.pack(side="left", fill="x", expand=True)
        self.wishlist_listbox.bind("<Double-Button-1>", lambda e: self.open_wishlist_selected())
        wishlist_btn_col = ttk.Frame(wishlist_frame, style="Card.TFrame")
        wishlist_btn_col.pack(side="left", padx=(8, 0))
        ttk.Button(wishlist_btn_col, text=t("home_open_button"), command=self.open_wishlist_selected).pack(fill="x")
        ttk.Button(wishlist_btn_col, text=t("home_new_draw_button"),
                   command=self._refresh_wishlist_sample).pack(fill="x", pady=(4, 0))
        self.wishlist_sample = []
        self._refresh_wishlist_sample()

        warnings = []
        if not PIL_AVAILABLE:
            warnings.append(t("warning_pillow"))
        if not REPORTLAB_AVAILABLE:
            warnings.append(t("warning_reportlab"))
        if not OPENPYXL_AVAILABLE:
            warnings.append(t("warning_openpyxl"))
        if not QRCODE_AVAILABLE:
            warnings.append(t("warning_qrcode"))
        if not PYTESSERACT_AVAILABLE:
            warnings.append(t("warning_pytesseract"))
        if warnings:
            ttk.Label(content, text="\n".join(warnings), foreground=COLOR_ERROR, font=("Segoe UI", sf(8)),
                      justify="center").pack(pady=(10, 10))
        else:
            ttk.Label(content, text="", font=("Segoe UI", sf(4))).pack(pady=(5, 5))

        tk.Frame(content, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

        self.footer = ttk.Label(self, text=t("home_footer_recipe_count", count=len(self.recipes)),
                                 font=("Segoe UI", sf(9)))
        self.footer.pack(side="bottom", pady=15)

    def refresh_recipes(self):
        self.recipes = load_recipes()
        self.footer.config(text=t("home_footer_recipe_count", count=len(self.recipes)))

    def _refresh_today_meals(self):
        for child in self.today_frame.winfo_children():
            child.destroy()

        today_name = WEEKDAYS[datetime.now().weekday()]  # 0=Lundi ... 6=Dimanche
        plan = load_weekly_plan()
        day_data = plan.get(today_name) or {}

        entries = []
        for slot in WeeklyPlanWindow.MEAL_SLOTS:
            slot_data = day_data.get(slot)
            if slot_data and slot_data.get("recipe_name"):
                entries.append((slot, slot_data["recipe_name"], slot_data.get("persons", 1)))

        if not entries:
            ttk.Label(
                self.today_frame,
                text=t("home_nothing_planned", day=(
                    translate_weekday_name(today_name).lower() if CURRENT_LANGUAGE == "fr"
                    else translate_weekday_name(today_name)
                )),
                font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED, wraplength=560, justify="left",
                style="Card.TLabel"
            ).pack(anchor="w", pady=3)
            return

        ttk.Label(self.today_frame, text=translate_weekday_name(today_name), font=("Segoe UI", sf(9), "bold"),
                  foreground=COLOR_TEXT_MUTED, style="Card.TLabel").pack(anchor="w")
        for slot, recipe_name, persons in entries:
            row = ttk.Frame(self.today_frame, style="Card.TFrame")
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=f"{translate_mealslot_name(slot)} :", width=18, anchor="w", style="Card.TLabel").pack(side="left")
            ttk.Label(row, text=f"{recipe_name} ({persons} pers.)", anchor="w",
                      style="Card.TLabel").pack(side="left")
            ttk.Button(row, text="👁", width=3,
                       command=lambda n=recipe_name: self._open_today_recipe(n)).pack(side="left", padx=5)

    def _open_today_recipe(self, recipe_name):
        win = OneRecipeWindow(self, initial_recipe_name=recipe_name)
        self.wait_window(win)
        self._refresh_recent_views()

    def _refresh_recent_views(self):
        self.recent_listbox.delete(0, tk.END)
        names = load_recent_view_names()
        self._recent_recipes = []
        for name in names:
            recipe = find_recipe_by_name(self.recipes, name)
            if recipe is None:
                continue  # la recette a été supprimée depuis
            self._recent_recipes.append(recipe)
            self.recent_listbox.insert(tk.END, format_recipe_list_label(recipe))
        if not self._recent_recipes:
            self.recent_listbox.insert(tk.END, t("home_no_recent_recipe"))

    def _open_daily_recipe(self, recipe):
        OneRecipeWindow(self, initial_recipe_name=recipe["name"])

    def _open_low_stock_to_cart(self, items):
        win = AllRecipesWindow(self)
        # Le seuil d'alerte sert de quantité suggérée à racheter (une
        # estimation raisonnable de « combien en garder en stock »).
        to_add = [{"name": e["name"], "quantity": e["threshold"], "unit": e["unit"]} for e in items]
        win.add_manual_items(to_add)

    def open_recent_selected(self):
        sel = self.recent_listbox.curselection()
        if not sel or not getattr(self, "_recent_recipes", None):
            return
        recipe = self._recent_recipes[sel[0]]
        win = OneRecipeWindow(self, initial_recipe_name=recipe["name"])
        self.wait_window(win)
        self._refresh_recent_views()

    def _refresh_wishlist_sample(self):
        """Tire au sort jusqu'à 10 recettes parmi celles de la liste d'envies
        (« à essayer »), pour donner une nouvelle idée à chaque tirage plutôt
        que de toujours montrer les mêmes en premier."""
        self.wishlist_listbox.delete(0, tk.END)
        candidates = [r for r in self.recipes if r.get("wishlist")]
        if not candidates:
            self.wishlist_listbox.insert(tk.END, t("home_no_wishlist_recipe"))
            self.wishlist_sample = []
            return
        sample_size = min(10, len(candidates))
        self.wishlist_sample = random.sample(candidates, sample_size)
        for r in self.wishlist_sample:
            cat = translate_category_name(r.get("category", "Autre"))
            self.wishlist_listbox.insert(tk.END, f"[{cat}] {r['name']}")

    def open_wishlist_selected(self):
        sel = self.wishlist_listbox.curselection()
        if not sel or not self.wishlist_sample:
            return
        recipe = self.wishlist_sample[sel[0]]
        win = OneRecipeWindow(self, initial_recipe_name=recipe["name"])
        self.wait_window(win)
        self._refresh_wishlist_sample()

    def refresh_ingredients(self):
        self.ingredient_names = load_ingredients()

    # ---------- Ajouter une recette ----------
    def open_add_recipe(self):
        RecipeFormWindow(self, recipe_index=None)

    # ---------- Voir toutes les recettes ----------
    def open_all_recipes(self):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        AllRecipesWindow(self)

    # ---------- Voir une recette précise ----------
    def open_one_recipe(self):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        win = OneRecipeWindow(self)
        self.wait_window(win)
        self._refresh_recent_views()

    # ---------- Modifier / Supprimer une recette ----------
    def open_manage_recipes(self, quick_filter=None):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        win = ManageRecipesWindow(self, quick_filter=quick_filter)
        self.wait_window(win)
        self._refresh_recent_views()

    # ---------- Gérer les ingrédients ----------
    def open_manage_ingredients(self):
        ManageIngredientsWindow(self)

    # ---------- Corbeille ----------
    def open_trash(self):
        win = TrashWindow(self)
        self.wait_window(win)
        self._refresh_recent_views()

    # ---------- Recherche par ingrédient ----------
    def open_ingredient_search(self):
        IngredientSearchWindow(self)

    # ---------- Comparer deux recettes ----------
    def open_compare_recipes(self):
        if len(self.recipes) < 2:
            messagebox.showinfo("Info", "Il faut au moins 2 recettes enregistrées pour pouvoir comparer.")
            return
        CompareRecipesWindow(self)

    # ---------- Importer une recette depuis un lien ----------
    def open_import_from_url(self):
        ImportFromUrlWindow(self)

    # ---------- Importer une recette depuis une photo ----------
    def open_import_from_photo(self):
        ImportFromPhotoWindow(self)

    # ---------- Importer / Exporter les données ----------
    def open_import_export(self):
        ImportExportWindow(self)

    # ---------- Que puis-je cuisiner ? ----------
    def open_what_can_i_cook(self):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        WhatCanICookWindow(self)

    # ---------- Mon garde-manger ----------
    def open_pantry(self):
        win = PantryWindow(self)
        self.wait_window(win)
        self._build_home_ui()

    def open_unit_converter(self):
        UnitConverterWindow(self)

    # ---------- Planning de la semaine ----------
    def open_weekly_plan(self):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        win = WeeklyPlanWindow(self)
        self.wait_window(win)
        self._refresh_today_meals()

    # ---------- Mes menus ----------
    def open_menus(self):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        MenuManagerWindow(self)

    # ---------- Statistiques ----------
    def open_statistics(self):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        StatisticsWindow(self)

    # ---------- Exporter le livre de recettes ----------
    def open_cookbook_export(self):
        if not self.recipes:
            messagebox.showinfo("Info", "Aucune recette enregistrée pour le moment.")
            return
        CookbookExportWindow(self)


class RecipeFormWindow(tk.Toplevel):
    """Fenêtre d'ajout OU de modification d'une recette (nom, catégorie, photo,
    description, ingrédients pour 1 personne). Si recipe_index est fourni, la
    fenêtre s'ouvre en mode modification, pré-remplie avec la recette
    existante."""

    CATEGORY_OPTIONS = ["Petit-déjeuner", "Entrée", "Plat", "Dessert", "Apéro", "Boisson", "Sauce", "Autre"]
    DIFFICULTY_OPTIONS = ["Facile", "Moyen", "Difficile"]
    MAX_DESC_LEN = 2056
    MAX_NOTES_LEN = 500

    def __init__(self, app, recipe_index=None, prefill=None):
        super().__init__(app)
        self.app = app
        self.recipe_index = recipe_index
        self.editing = recipe_index is not None
        self.existing_recipe = app.recipes[recipe_index] if self.editing else None
        self.prefill = prefill if not self.editing else None
        self.ingredient_names = load_ingredients()

        # Galerie de photos : chaque élément est ("existing", nom_de_fichier)
        # pour une photo déjà enregistrée, ou ("new", chemin_source) pour une
        # photo qui vient d'être choisie et sera copiée lors de l'enregistrement.
        self.gallery_items = []
        if self.editing:
            self.gallery_items = [("existing", fname) for fname in get_recipe_images(self.existing_recipe)]
        elif self.prefill:
            # La photo a déjà été téléchargée et enregistrée dans images/ par
            # fetch_recipe_from_url ; on la référence donc comme "existing".
            self.gallery_items = [("existing", fname) for fname in self.prefill.get("images", [])]
        self._gallery_thumb_refs = []  # garder une référence pour éviter le garbage collector

        self.title(t("recipeform_title_edit") if self.editing else t("recipeform_title_add"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(1220)}x{screen_height}+40+0")
        self.minsize(gs(760), gs(500))
        self.grab_set()

        # ---- Conteneur scrollable pour tout le formulaire ----
        # Tout le contenu (nom, catégorie, photo, description, ingrédients,
        # boutons) est placé dans ce conteneur qui défile d'un bloc : ainsi,
        # les boutons "+ Ajouter un ingrédient" / "Enregistrer" restent
        # toujours juste après le dernier ingrédient, où qu'il soit.
        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.content_frame = ttk.Frame(self.canvas)
        self.content_frame.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ---- Rangée 1 : infos générales à gauche, allergènes à droite ----
        row1 = ttk.Frame(self.content_frame)
        row1.pack(fill="both", expand=True)
        row1_left = ttk.Frame(row1)
        row1_left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        row1_right = ttk.Frame(row1)
        row1_right.pack(side="left", fill="both", expand=True, padx=(10, 0), anchor="n")

        ttk.Label(row1_left, text=t("recipeform_name_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(15, 5))
        self.name_entry = ttk.Entry(row1_left, width=42)
        self.name_entry.pack()
        if self.editing:
            self.name_entry.insert(0, self.existing_recipe["name"])
        elif self.prefill:
            self.name_entry.insert(0, self.prefill.get("name", ""))

        self.favorite_var = tk.BooleanVar(
            value=self.existing_recipe.get("favorite", False) if self.editing else False
        )
        ttk.Checkbutton(row1_left, text=t("recipeform_favorite_checkbox"),
                         variable=self.favorite_var).pack(pady=(8, 0))

        self.wishlist_var = tk.BooleanVar(
            value=self.existing_recipe.get("wishlist", False) if self.editing else False
        )
        ttk.Checkbutton(row1_left, text=t("recipeform_wishlist_checkbox"),
                         variable=self.wishlist_var).pack(pady=(4, 0))

        # ---- Note personnelle (1 à 5 étoiles, cliquables) ----
        self.rating_value = self.existing_recipe.get("rating", 0) if self.editing else 0
        rating_frame = ttk.Frame(row1_left)
        rating_frame.pack(pady=(8, 0))
        ttk.Label(rating_frame, text=t("recipeform_rating_label")).pack(side="left", padx=(0, 5))
        self.rating_star_labels = []
        for i in range(1, 6):
            lbl = ttk.Label(rating_frame, text="☆", font=("Segoe UI", sf(14)), cursor="hand2")
            lbl.pack(side="left")
            lbl.bind("<Button-1>", lambda e, i=i: self._set_rating(i))
            self.rating_star_labels.append(lbl)
        self._refresh_rating_stars()

        ttk.Label(row1_left, text=t("recipeform_category_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(15, 5))
        self.category_combo = ttk.Combobox(row1_left, values=[translate_category_name(c) for c in self.CATEGORY_OPTIONS],
                                            state="readonly", width=20)
        self.category_combo.set(
            translate_category_name(self.existing_recipe.get("category", "Plat") if self.editing else "Plat")
        )
        self.category_combo.pack()

        # ---- Temps de préparation / cuisson / difficulté ----
        times_frame = ttk.Frame(row1_left)
        times_frame.pack(pady=(15, 5))
        ttk.Label(times_frame, text=t("recipeform_prep_time_label")).grid(row=0, column=0, padx=3, sticky="e")
        self.prep_time_entry = ttk.Entry(times_frame, width=6)
        self.prep_time_entry.grid(row=0, column=1, padx=3)
        ttk.Label(times_frame, text=t("recipeform_cook_time_label")).grid(row=0, column=2, padx=3, sticky="e")
        self.cook_time_entry = ttk.Entry(times_frame, width=6)
        self.cook_time_entry.grid(row=0, column=3, padx=3)
        if self.editing:
            self.prep_time_entry.insert(0, str(self.existing_recipe.get("prep_time", "") or ""))
            self.cook_time_entry.insert(0, str(self.existing_recipe.get("cook_time", "") or ""))
        elif self.prefill:
            self.prep_time_entry.insert(0, str(self.prefill.get("prep_time", "") or ""))
            self.cook_time_entry.insert(0, str(self.prefill.get("cook_time", "") or ""))

        difficulty_frame = ttk.Frame(row1_left)
        difficulty_frame.pack(pady=(5, 5))
        ttk.Label(difficulty_frame, text=t("recipeform_difficulty_label")).pack(side="left", padx=3)
        self.difficulty_combo = ttk.Combobox(difficulty_frame, values=[translate_difficulty_name(d) for d in self.DIFFICULTY_OPTIONS],
                                              state="readonly", width=15)
        self.difficulty_combo.set(
            translate_difficulty_name(self.existing_recipe.get("difficulty", "Facile") if self.editing else "Facile")
        )
        self.difficulty_combo.pack(side="left", padx=3)
        ttk.Label(difficulty_frame, text=t("recipeform_default_persons_label")).pack(side="left", padx=(10, 3))
        self.default_persons_entry = ttk.Entry(difficulty_frame, width=5)
        if self.editing:
            default_persons_value = str(self.existing_recipe.get("default_persons", 4))
        elif self.prefill:
            default_persons_value = str(self.prefill.get("default_persons", 4))
        else:
            default_persons_value = "4"
        self.default_persons_entry.insert(0, default_persons_value)
        self.default_persons_entry.pack(side="left")

        # ---- Étiquettes libres ----
        ttk.Label(row1_left, text=t("recipeform_tags_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(15, 5))
        self.tags_entry = ttk.Entry(row1_left, width=42)
        self.tags_entry.pack()
        if self.editing and self.existing_recipe.get("tags"):
            self.tags_entry.insert(0, ", ".join(self.existing_recipe["tags"]))
        ttk.Label(row1_left, text=t("recipeform_tags_example"),
                  font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED).pack()

        # ---- Allergènes ----
        allergens_header = ttk.Frame(row1_right)
        allergens_header.pack(fill="x", padx=10, pady=(15, 5))
        ttk.Label(allergens_header, text=t("recipeform_allergens_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(side="left")
        ttk.Button(allergens_header, text=t("recipeform_detect_allergens_button"),
                   command=self.detect_allergens_from_ingredients).pack(side="right")
        ttk.Label(
            row1_right,
            text=t("recipeform_allergens_disclaimer"),
            font=("Segoe UI", sf(8), "bold"), foreground="#FF0000", justify="center"
        ).pack(pady=(0, 8))
        allergens_frame = ttk.Frame(row1_right)
        allergens_frame.pack()
        if self.editing:
            existing_allergens = set(self.existing_recipe.get("allergens", []))
        elif self.prefill and self.prefill.get("ingredients"):
            # Recette importée depuis un lien : détection automatique dès l'ouverture
            existing_allergens = set(compute_recipe_allergens(self.prefill["ingredients"]))
        else:
            existing_allergens = set()
        self.allergen_vars = {}
        for i, allergen in enumerate(ALLERGENS):
            var = tk.BooleanVar(value=allergen in existing_allergens)
            self.allergen_vars[allergen] = var
            ttk.Checkbutton(allergens_frame, text=translate_allergen_name(allergen), variable=var).grid(
                row=i // 2, column=i % 2, sticky="w", padx=8, pady=2
            )
        # Sert à ne jamais décocher un allergène que l'utilisateur aurait
        # coché lui-même sans lien avec un ingrédient détecté : on ne
        # décoche automatiquement que ce que la détection a elle-même coché.
        if self.editing:
            self._auto_detected_allergens = set(
                compute_recipe_allergens(self.existing_recipe.get("ingredients", []))
            )
        elif self.prefill and self.prefill.get("ingredients"):
            self._auto_detected_allergens = set(compute_recipe_allergens(self.prefill["ingredients"]))
        else:
            self._auto_detected_allergens = set()
        ttk.Label(
            row1_right,
            text=t("recipeform_allergens_auto_note"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 5))

        # ---- Photos (galerie) ----
        ttk.Label(self.content_frame, text=t("recipeform_photos_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(15, 5))
        gallery_outer = ttk.Frame(self.content_frame)
        gallery_outer.pack(fill="x", padx=10)
        self.gallery_canvas = tk.Canvas(gallery_outer, height=130, highlightthickness=0)
        gallery_scrollbar = ttk.Scrollbar(gallery_outer, orient="horizontal",
                                           command=self.gallery_canvas.xview)
        self.gallery_frame = ttk.Frame(self.gallery_canvas)
        self.gallery_frame.bind(
            "<Configure>", lambda e: self.gallery_canvas.configure(scrollregion=self.gallery_canvas.bbox("all"))
        )
        self.gallery_canvas.create_window((0, 0), window=self.gallery_frame, anchor="nw")
        self.gallery_canvas.configure(xscrollcommand=gallery_scrollbar.set)
        self.gallery_canvas.pack(fill="x")
        gallery_scrollbar.pack(fill="x")
        ttk.Button(self.content_frame, text=t("recipeform_add_photo_button"),
                   command=self.choose_images).pack(pady=5)
        self._refresh_gallery()

        # ---- Rangée 2 : ingrédients à gauche, description/notes à droite ----
        row2 = ttk.Frame(self.content_frame)
        row2.pack(fill="both", expand=True)
        row2_left = ttk.Frame(row2)
        row2_left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        row2_right = ttk.Frame(row2)
        row2_right.pack(side="left", fill="both", expand=True, padx=(10, 0), anchor="n")

        # ---- Description ----
        ttk.Label(row2_right, text=t("recipeform_description_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(15, 5))
        desc_frame = ttk.Frame(row2_right)
        desc_frame.pack(padx=10)
        self.description_text = tk.Text(desc_frame, height=10, width=58, wrap="word", font=("Segoe UI", sf(10)))
        self.description_text.pack()
        if self.editing and self.existing_recipe.get("description"):
            self.description_text.insert("1.0", self.existing_recipe["description"])
        elif self.prefill and self.prefill.get("description"):
            self.description_text.insert("1.0", self.prefill["description"])
        self.desc_counter_label = ttk.Label(row2_right, text="", font=("Segoe UI", sf(8)),
                                             foreground=COLOR_TEXT_MUTED)
        self.desc_counter_label.pack(pady=(2, 0))
        self.description_text.bind("<<Modified>>", self._on_description_modified)
        self._on_description_modified()

        # ---- Notes personnelles ----
        ttk.Label(row2_right, text=t("recipeform_notes_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(15, 5))
        notes_frame = ttk.Frame(row2_right)
        notes_frame.pack(padx=10)
        self.notes_text = tk.Text(notes_frame, height=5, width=58, wrap="word", font=("Segoe UI", sf(10)))
        self.notes_text.pack()
        if self.editing and self.existing_recipe.get("personal_notes"):
            self.notes_text.insert("1.0", self.existing_recipe["personal_notes"])
        self.notes_counter_label = ttk.Label(row2_right, text="", font=("Segoe UI", sf(8)),
                                              foreground=COLOR_TEXT_MUTED)
        self.notes_counter_label.pack(pady=(2, 0))
        self.notes_text.bind("<<Modified>>", self._on_notes_modified)
        self._on_notes_modified()

        # ---- Ingrédients ----
        ing_header_frame = ttk.Frame(row2_left)
        ing_header_frame.pack(fill="x", padx=10, pady=(15, 5))
        ttk.Label(ing_header_frame, text=t("recipeform_ingredients_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(side="left")
        ttk.Button(ing_header_frame, text=t("recipeform_new_ingredient_button"),
                   command=self.add_new_ingredient_global).pack(side="right")

        if not self.ingredient_names:
            ttk.Label(row2_left,
                      text=t("recipeform_no_ingredients_registered"),
                      font=("Segoe UI", sf(8)), foreground=COLOR_ERROR, justify="center").pack()

        self.rows_frame = ttk.Frame(row2_left)
        self.rows_frame.pack(fill="x", padx=10)

        header = ttk.Frame(self.rows_frame)
        header.pack(fill="x", pady=2)
        ttk.Label(header, text=t("recipeform_header_ingredient"), width=17,
                  font=("Segoe UI", sf(9), "bold")).grid(row=0, column=0)
        ttk.Label(header, text=t("recipeform_header_quantity"), width=9,
                  font=("Segoe UI", sf(9), "bold")).grid(row=0, column=1)
        ttk.Label(header, text=t("recipeform_header_unit"), width=15,
                  font=("Segoe UI", sf(9), "bold")).grid(row=0, column=2)
        ttk.Label(header, text=t("recipeform_header_other"), width=10,
                  font=("Segoe UI", sf(9), "bold")).grid(row=0, column=3)

        self.ingredient_rows = []
        if self.editing and self.existing_recipe["ingredients"]:
            for ing in self.existing_recipe["ingredients"]:
                self.add_ingredient_row(ing["name"], ing["quantity"], ing["unit"])
        elif self.prefill and self.prefill.get("ingredients"):
            for ing in self.prefill["ingredients"]:
                self.add_ingredient_row(ing["name"], ing["quantity"], ing["unit"])
        else:
            self.add_ingredient_row()

        # Ces deux éléments sont recréés/déplacés à chaque ajout de ligne afin
        # de toujours rester juste en dessous du dernier ingrédient.
        self.add_ingredient_button = ttk.Button(
            self.rows_frame, text=t("recipeform_add_ingredient_button"), command=lambda: self.add_ingredient_row()
        )
        self.add_ingredient_button.pack(pady=10)

        self.bottom_actions_frame = ttk.Frame(self.rows_frame)
        self.bottom_actions_frame.pack(pady=(0, 20))
        ttk.Button(self.bottom_actions_frame, text=t("recipeform_save_button"),
                   command=self.save_recipe).grid(row=0, column=0, padx=5)
        if self.editing:
            ttk.Button(self.bottom_actions_frame, text=t("recipeform_delete_button"),
                       command=self.delete_recipe).grid(row=0, column=1, padx=5)

        tk.Frame(self.content_frame, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _on_description_modified(self, event=None):
        self.description_text.edit_modified(False)
        content = self.description_text.get("1.0", "end-1c")
        if len(content) > self.MAX_DESC_LEN:
            content = content[: self.MAX_DESC_LEN]
            self.description_text.delete("1.0", "end")
            self.description_text.insert("1.0", content)
            self.description_text.edit_modified(False)
        self.desc_counter_label.config(text=t("recipeform_char_counter", count=len(content), max=self.MAX_DESC_LEN))

    def _on_notes_modified(self, event=None):
        self.notes_text.edit_modified(False)
        content = self.notes_text.get("1.0", "end-1c")
        if len(content) > self.MAX_NOTES_LEN:
            content = content[: self.MAX_NOTES_LEN]
            self.notes_text.delete("1.0", "end")
            self.notes_text.insert("1.0", content)
            self.notes_text.edit_modified(False)
        self.notes_counter_label.config(text=t("recipeform_char_counter", count=len(content), max=self.MAX_NOTES_LEN))

    def _set_rating(self, value):
        self.rating_value = 0 if self.rating_value == value else value
        self._refresh_rating_stars()

    def _refresh_rating_stars(self):
        for i, lbl in enumerate(self.rating_star_labels, start=1):
            lbl.config(text="★" if i <= self.rating_value else "☆")

    def detect_allergens_from_ingredients(self):
        current_ingredients = []
        for name_e, qty_e, unit_e, custom_e in self.ingredient_rows:
            ing_name = name_e.get().strip()
            if ing_name:
                current_ingredients.append({"name": ing_name})
        if not current_ingredients:
            messagebox.showinfo(t("common_info"), t("recipeform_add_ingredients_first"))
            return

        before = {a for a, var in self.allergen_vars.items() if var.get()}
        self._sync_allergens_from_ingredients()
        after = {a for a, var in self.allergen_vars.items() if var.get()}

        added = sorted(after - before)
        removed = sorted(before - after)
        if added or removed:
            parts = []
            if added:
                parts.append(t("recipeform_allergens_updated_added", list=", ".join(translate_allergen_name(a) for a in added)))
            if removed:
                parts.append(t("recipeform_allergens_updated_removed", list=", ".join(translate_allergen_name(a) for a in removed)))
            messagebox.showinfo(
                t("recipeform_allergens_updated_title"),
                t("recipeform_allergens_updated_message", parts=" ; ".join(parts))
            )
        else:
            messagebox.showinfo(t("common_info"), t("recipeform_allergens_no_change"))

    def choose_images(self):
        paths = filedialog.askopenfilenames(
            title=t("recipeform_choose_photos_title"),
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.webp")]
        )
        for path in paths:
            self.gallery_items.append(("new", path))
        if paths:
            self._refresh_gallery()

    def _remove_gallery_item(self, index):
        if 0 <= index < len(self.gallery_items):
            self.gallery_items.pop(index)
            self._refresh_gallery()

    def _refresh_gallery(self):
        for child in self.gallery_frame.winfo_children():
            child.destroy()
        self._gallery_thumb_refs = []

        if not self.gallery_items:
            ttk.Label(self.gallery_frame, text=t("recipeform_no_photo")).pack(side="left", padx=10, pady=10)
            return

        for idx, (kind, ref) in enumerate(self.gallery_items):
            cell = ttk.Frame(self.gallery_frame)
            cell.pack(side="left", padx=5, pady=5)

            thumb = None
            if PIL_AVAILABLE:
                try:
                    if kind == "existing":
                        thumb = load_thumbnail(ref, size=(110, 90))
                    else:
                        img = Image.open(ref)
                        img.thumbnail((110, 90))
                        thumb = ImageTk.PhotoImage(img)
                except Exception:
                    thumb = None

            if thumb is not None:
                self._gallery_thumb_refs.append(thumb)
                ttk.Label(cell, image=thumb).pack()
            else:
                ttk.Label(cell, text=t("recipeform_preview_unavailable"), justify="center").pack()

            ttk.Button(cell, text=t("recipeform_remove_photo_button"), width=10,
                       command=lambda i=idx: self._remove_gallery_item(i)).pack(pady=(3, 0))

    UNIT_OPTIONS = ["Gr", "Kilo", "cl", "Litre", "pièce", "cuillère à soupe", "cuillère à café", "autre"]

    @staticmethod
    def _map_unit_for_edit(unit):
        """Convertit une unité stockée (éventuellement héritée d'une ancienne
        version) vers (valeur du menu déroulant, texte personnalisé)."""
        u = (unit or "").strip()
        u_lower = u.lower()
        if u_lower in ("gr", "g", "gramme", "grammes"):
            return "Gr", ""
        if u_lower == "cl":
            return "cl", ""
        if u_lower in ("pièce", "piece", "pieces", "pièces"):
            return "pièce", ""
        if u_lower in ("cuillère à soupe", "cuillere a soupe", "c. à soupe", "cas"):
            return "cuillère à soupe", ""
        if u_lower in ("cuillère à café", "cuillere a café", "cuillere a cafe", "c. à café", "cac"):
            return "cuillère à café", ""
        if u == "":
            return "Gr", ""
        if u == "autre":
            return "autre", ""
        return "autre", u

    @staticmethod
    def _filter_ingredients(full_values, typed):
        if not typed:
            return full_values
        typed_key = ingredient_sort_key(typed)
        filtered = [v for v in full_values if ingredient_sort_key(v).startswith(typed_key)]
        if not filtered:
            filtered = [v for v in full_values if typed_key in ingredient_sort_key(v)]
        return filtered

    def _hide_suggestions(self, entry):
        popup = getattr(entry, "_suggestion_popup", None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
            entry._suggestion_popup = None
            entry._suggestion_listbox = None

    def _show_suggestions(self, entry, filtered):
        self._hide_suggestions(entry)
        if not filtered:
            return
        popup = tk.Toplevel(entry)
        popup.wm_overrideredirect(True)
        try:
            popup.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        width = max(entry.winfo_width(), 160)
        height = min(6, len(filtered)) * 20
        popup.wm_geometry(f"{width}x{height}+{x}+{y}")

        listbox = tk.Listbox(popup, height=min(6, len(filtered)), exportselection=False, font=("Segoe UI", sf(9)))
        listbox.pack(fill="both", expand=True)
        for v in filtered:
            listbox.insert(tk.END, v)

        def choose(event=None):
            sel = listbox.curselection()
            if sel:
                value = listbox.get(sel[0])
                entry.delete(0, tk.END)
                entry.insert(0, value)
                self._sync_allergens_from_ingredients()
            self._hide_suggestions(entry)
            entry.focus_set()

        listbox.bind("<ButtonRelease-1>", choose)
        listbox.bind("<Return>", choose)
        entry._suggestion_popup = popup
        entry._suggestion_listbox = listbox

    def _sync_allergens_from_ingredients(self):
        """Recalcule les allergènes détectés à partir de TOUS les ingrédients
        actuellement saisis, et coche/décoche les cases en conséquence. Ne
        décoche jamais une case pour un allergène que l'utilisateur aurait
        cochée lui-même sans qu'un ingrédient détecté ne soit à l'origine
        (seuls les allergènes que la détection a elle-même cochés peuvent
        être décochés automatiquement par la suite)."""
        if not hasattr(self, "allergen_vars"):
            return
        current_ingredients = []
        for name_e, qty_e, unit_e, custom_e in self.ingredient_rows:
            typed_name = name_e.get().strip()
            if typed_name:
                # Le champ peut afficher le nom traduit dans la langue
                # actuellement sélectionnée (ex. « Ground beef » en
                # anglais) : la base d'allergènes est indexée en français,
                # donc on résout d'abord vers le nom canonique avant toute
                # recherche, sans quoi la détection échouerait
                # silencieusement dans une langue autre que le français.
                resolved_name = resolve_ingredient_input(typed_name, self.app.ingredient_names) or typed_name
                current_ingredients.append({"name": resolved_name})
        detected = set(compute_recipe_allergens(current_ingredients))

        for allergen in self._auto_detected_allergens - detected:
            if allergen in self.allergen_vars:
                self.allergen_vars[allergen].set(False)
        for allergen in detected:
            if allergen in self.allergen_vars:
                self.allergen_vars[allergen].set(True)

        self._auto_detected_allergens = detected

    def _on_ingredient_keyrelease(self, event, entry):
        if event.keysym == "Down":
            listbox = getattr(entry, "_suggestion_listbox", None)
            if listbox is not None:
                listbox.focus_set()
                listbox.selection_set(0)
            return
        if event.keysym in ("Escape",):
            self._hide_suggestions(entry)
            return
        if event.keysym in ("Return", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R",
                              "Caps_Lock", "Alt_L", "Alt_R", "Left", "Right"):
            return
        full_values = getattr(entry, "full_values", [])
        typed = entry.get()
        filtered = self._filter_ingredients(full_values, typed)
        if filtered:
            self._show_suggestions(entry, filtered)
        else:
            self._hide_suggestions(entry)

    def _on_ingredient_focus_in(self, event, entry):
        full_values = getattr(entry, "full_values", [])
        typed = entry.get()
        filtered = self._filter_ingredients(full_values, typed)
        if filtered:
            self._show_suggestions(entry, filtered)

    def _on_ingredient_focus_out(self, event, entry):
        # Recalcule les allergènes à partir de l'ensemble des ingrédients
        # actuellement saisis (coche ou décoche selon le résultat).
        self._sync_allergens_from_ingredients()
        # Petit délai pour laisser le temps à un clic dans la liste de
        # suggestions de s'exécuter avant qu'elle ne soit masquée.
        entry.after(200, lambda: self._hide_suggestions(entry))

    def add_ingredient_row(self, name="", qty="", unit=""):
        # On détache temporairement les boutons de fin de liste pour que la
        # nouvelle ligne s'insère bien avant eux, puis on les replace après.
        has_controls = hasattr(self, "add_ingredient_button")
        if has_controls:
            self.add_ingredient_button.pack_forget()
            self.bottom_actions_frame.pack_forget()

        row = ttk.Frame(self.rows_frame)
        row.pack(fill="x", pady=2)
        values = list(self.ingredient_names)
        if name and name not in values:
            values = sorted(values + [name], key=ingredient_sort_key)
        name_e = ttk.Entry(row, width=17)  # champ libre avec suggestions déroulantes personnalisées
        name_e.full_values = get_display_ingredient_values(values)
        name_e._suggestion_popup = None
        name_e._suggestion_listbox = None
        if name:
            name_e.insert(0, translate_ingredient_name(name))
        name_e.grid(row=0, column=0, padx=2)
        name_e.bind("<KeyRelease>", lambda e, ent=name_e: self._on_ingredient_keyrelease(e, ent))
        name_e.bind("<FocusIn>", lambda e, ent=name_e: self._on_ingredient_focus_in(e, ent))
        name_e.bind("<FocusOut>", lambda e, ent=name_e: self._on_ingredient_focus_out(e, ent))
        qty_e = ttk.Entry(row, width=9)
        qty_e.insert(0, "" if qty == "" else str(qty))
        qty_e.grid(row=0, column=1, padx=2)

        combo_value, custom_text = self._map_unit_for_edit(unit)
        unit_e = ttk.Combobox(row, width=15, state="readonly",
                               values=[translate_unit_name(u) for u in self.UNIT_OPTIONS])
        unit_e.set(translate_unit_name(combo_value))
        unit_e.grid(row=0, column=2, padx=2)

        custom_e = ttk.Entry(row, width=10)
        custom_e.insert(0, custom_text)
        custom_e.grid(row=0, column=3, padx=2)
        if combo_value != "autre":
            custom_e.grid_remove()

        def on_unit_change(event, u=unit_e, c=custom_e):
            if u.get() == "autre":
                c.grid()
            else:
                c.delete(0, tk.END)
                c.grid_remove()

        unit_e.bind("<<ComboboxSelected>>", on_unit_change)
        self.ingredient_rows.append((name_e, qty_e, unit_e, custom_e))

        if has_controls:
            self.add_ingredient_button.pack(pady=10)
            self.bottom_actions_frame.pack(pady=(0, 20))
            # Fait défiler la fenêtre pour amener la nouvelle ligne en vue
            self.update_idletasks()
            self.canvas.yview_moveto(1.0)

    def add_new_ingredient_global(self):
        new_name = simpledialog.askstring(
            t("recipeform_new_ingredient_dialog_title"), t("recipeform_new_ingredient_dialog_prompt"), parent=self
        )
        if not new_name:
            return
        new_name = normalize_oe(new_name.strip())
        if not new_name:
            return
        existing_match = resolve_ingredient_input(new_name, self.ingredient_names)
        if existing_match is not None:
            messagebox.showinfo(
                t("common_info"), t("recipeform_ingredient_already_exists", name=existing_match)
            )
            return

        ingredients = load_ingredients()
        ingredients.append(new_name)
        self.ingredient_names = save_ingredients(ingredients)
        self.app.refresh_ingredients()

        # Met à jour les suggestions déjà affichées dans cette fenêtre
        for name_e, qty_e, unit_e, custom_e in self.ingredient_rows:
            name_e.full_values = get_display_ingredient_values(self.ingredient_names)

        messagebox.showinfo(
            t("recipeform_ingredient_added_title"),
            t("recipeform_ingredient_added_message", name=new_name)
        )

    @staticmethod
    def _validate_time_field(raw_value, label):
        """Valide un champ de temps optionnel (minutes). Retourne
        (valeur_normalisée, ok). Une valeur vide est acceptée (temps non
        renseigné)."""
        raw_value = raw_value.strip()
        if not raw_value:
            return "", True
        try:
            value = float(raw_value.replace(",", "."))
            if value < 0:
                raise ValueError
        except ValueError:
            return None, False
        if value == int(value):
            return str(int(value)), True
        return str(value), True

    def save_recipe(self):
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror(t("common_error"), t("recipeform_error_name_required"))
            return

        prep_time, ok_prep = self._validate_time_field(self.prep_time_entry.get(), "préparation")
        if not ok_prep:
            messagebox.showerror(t("common_error"), t("recipeform_error_prep_time"))
            return
        cook_time, ok_cook = self._validate_time_field(self.cook_time_entry.get(), "cuisson")
        if not ok_cook:
            messagebox.showerror(t("common_error"), t("recipeform_error_cook_time"))
            return

        ingredients = []
        for name_e, qty_e, unit_e, custom_e in self.ingredient_rows:
            ing_name_raw = name_e.get().strip()
            qty_str = qty_e.get().strip().replace(",", ".")
            unit_choice = resolve_unit_input(unit_e.get().strip(), self.UNIT_OPTIONS)
            if not ing_name_raw:
                continue
            canonical = resolve_ingredient_input(ing_name_raw, self.ingredient_names)
            if canonical is None:
                messagebox.showerror(
                    t("recipeform_unknown_ingredient_title"),
                    t("recipeform_unknown_ingredient_message", name=ing_name_raw)
                )
                return
            ing_name = canonical
            try:
                qty = float(qty_str) if qty_str else 0
            except ValueError:
                messagebox.showerror(t("common_error"), t("recipeform_error_invalid_quantity", name=ing_name))
                return
            if unit_choice == "autre":
                unit = custom_e.get().strip()
                if not unit:
                    messagebox.showerror(
                        t("common_error"), t("recipeform_error_custom_unit_required", name=ing_name)
                    )
                    return
            else:
                unit = unit_choice
            ingredients.append({"name": ing_name, "quantity": qty, "unit": unit})

        if not ingredients:
            messagebox.showerror(t("common_error"), t("recipeform_error_no_valid_ingredient"))
            return

        # Détection de doublons : le même ingrédient saisi dans plusieurs
        # lignes est presque toujours une erreur de saisie (copier-coller,
        # ligne ajoutée deux fois par mégarde). On prévient avant
        # d'enregistrer, mais sans bloquer complètement — au cas où ce
        # serait volontaire (deux quantités séparées pour un même
        # ingrédient utilisé à deux endroits différents de la recette).
        seen_keys = set()
        duplicate_names = []
        for ing in ingredients:
            key = ingredient_sort_key(ing["name"])
            if key in seen_keys:
                if ing["name"] not in duplicate_names:
                    duplicate_names.append(ing["name"])
            else:
                seen_keys.add(key)
        if duplicate_names:
            duplicate_display = ", ".join(translate_ingredient_name(n) for n in duplicate_names)
            if not messagebox.askyesno(
                t("recipeform_duplicate_ingredient_title"),
                t("recipeform_duplicate_ingredient_message", list=duplicate_display)
            ):
                return

        # Construit la liste finale des photos : celles déjà enregistrées qui
        # n'ont pas été retirées, plus celles nouvellement choisies (copiées
        # sur le disque à ce moment-là).
        final_images = []
        for kind, ref in self.gallery_items:
            if kind == "existing":
                final_images.append(ref)
            else:
                final_images.append(copy_image_to_store(ref))

        if self.editing:
            old_images = get_recipe_images(self.existing_recipe)
            for fname in old_images:
                if fname not in final_images:
                    delete_image_file(fname)

        try:
            default_persons = float(self.default_persons_entry.get().strip().replace(",", "."))
            if default_persons <= 0:
                raise ValueError
        except ValueError:
            default_persons = 4
        if default_persons == int(default_persons):
            default_persons = int(default_persons)

        raw_tags = [t.strip() for t in self.tags_entry.get().split(",") if t.strip()]
        # Déduplique les étiquettes identiques à la casse près (ex. « rapide »
        # et « Rapide » tapées par erreur toutes les deux), en conservant la
        # première graphie rencontrée et l'ordre de saisie.
        tags = []
        seen_tag_keys = set()
        for tag in raw_tags:
            tag_key = ingredient_sort_key(tag)
            if tag_key not in seen_tag_keys:
                seen_tag_keys.add(tag_key)
                tags.append(tag)

        # La date de mise en liste d'envies n'est (re)fixée que lors du
        # passage de "pas en liste" à "en liste", pour pouvoir signaler plus
        # tard les recettes qui y attendent depuis longtemps.
        was_wishlist = self.existing_recipe.get("wishlist", False) if self.editing else False
        now_wishlist = self.wishlist_var.get()
        if now_wishlist and was_wishlist:
            wishlist_since = self.existing_recipe.get("wishlist_since") or datetime.now().isoformat()
        elif now_wishlist:
            wishlist_since = datetime.now().isoformat()
        else:
            wishlist_since = None

        recipes = load_recipes()
        recipe_data = {
            "name": name,
            "category": resolve_category_input(self.category_combo.get(), self.CATEGORY_OPTIONS),
            "favorite": self.favorite_var.get(),
            "wishlist": now_wishlist,
            "wishlist_since": wishlist_since,
            "rating": self.rating_value,
            "prep_time": prep_time,
            "cook_time": cook_time,
            "difficulty": resolve_difficulty_input(self.difficulty_combo.get(), self.DIFFICULTY_OPTIONS),
            "default_persons": default_persons,
            "tags": tags,
            "allergens": [a for a, var in self.allergen_vars.items() if var.get()],
            "description": self.description_text.get("1.0", "end-1c").strip()[: self.MAX_DESC_LEN],
            "personal_notes": self.notes_text.get("1.0", "end-1c").strip()[: self.MAX_NOTES_LEN],
            "ingredients": ingredients,
            "images": final_images,
            "created_at": self.existing_recipe.get("created_at") if self.editing else datetime.now().isoformat(),
            "times_cooked": self.existing_recipe.get("times_cooked", 0) if self.editing else 0,
            "cooked_dates": self.existing_recipe.get("cooked_dates", []) if self.editing else [],
        }
        if not recipe_data["created_at"]:
            recipe_data["created_at"] = datetime.now().isoformat()

        if self.editing:
            recipes[self.recipe_index] = recipe_data
        else:
            recipes.append(recipe_data)

        save_recipes(recipes)
        self.app.refresh_recipes()
        messagebox.showinfo(t("common_success"), t("recipeform_saved_message", name=name))
        self.destroy()

    def delete_recipe(self):
        if not messagebox.askyesno(
            t("common_confirm"),
            t("recipeform_delete_confirm_message", name=self.existing_recipe['name'])
        ):
            return
        recipes = load_recipes()
        removed = recipes.pop(self.recipe_index)
        save_recipes(recipes)
        move_recipe_to_trash(removed)
        self.app.refresh_recipes()
        messagebox.showinfo(t("recipeform_deleted_title"), t("recipeform_deleted_message"))
        self.destroy()



class ManageRecipesWindow(tk.Toplevel):
    """Fenêtre listant les recettes pour choisir laquelle modifier, dupliquer
    ou supprimer."""

    def __init__(self, app, quick_filter=None):
        super().__init__(app)
        self.app = app
        self.title(t("managerecipes_title"))
        self.geometry(f"{gs(560)}x{gs(580)}")
        self.grab_set()
        self.filtered_indices = []  # correspondance ligne affichée -> index réel dans app.recipes
        self.quick_filter = quick_filter

        ttk.Label(self, text=t("managerecipes_select_label"), font=("Segoe UI", sf(11), "bold")).pack(pady=(10, 5))

        quick_filter_label_keys = {
            "favoris": "managerecipes_filter_favorites",
            "rapide": "managerecipes_filter_quick",
            "vegetarien": "managerecipes_filter_vegetarian",
            "envie": "managerecipes_filter_wishlist",
        }
        if quick_filter in quick_filter_label_keys:
            filter_bar = ttk.Frame(self)
            filter_bar.pack(fill="x", padx=15, pady=(0, 5))
            ttk.Label(filter_bar, text=t(quick_filter_label_keys[quick_filter]),
                      foreground=COLOR_ACCENT_DARK, font=("Segoe UI", sf(9), "bold")).pack(side="left")
            ttk.Button(filter_bar, text=t("managerecipes_remove_filter_button"),
                       command=self._clear_quick_filter).pack(side="right")

        top_frame = ttk.Frame(self)
        top_frame.pack(pady=(0, 5), fill="x", padx=15)
        ttk.Label(top_frame, text=t("managerecipes_search_label")).pack(side="left")
        self.search_entry = ttk.Entry(top_frame, width=18)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate())

        sort_frame = ttk.Frame(self)
        sort_frame.pack(pady=(0, 5), fill="x", padx=15)
        ttk.Label(sort_frame, text=t("managerecipes_sort_label")).pack(side="left")
        self.sort_combo = ttk.Combobox(sort_frame, values=[translate_sort_option(o) for o in RECIPE_SORT_OPTIONS], state="readonly", width=22)
        self.sort_combo.set(translate_sort_option(RECIPE_SORT_OPTIONS[0]))
        self.sort_combo.pack(side="left", padx=5)
        self.sort_combo.bind("<<ComboboxSelected>>", lambda e: self._populate())
        ttk.Label(sort_frame, text=t("managerecipes_category_label")).pack(side="left", padx=(10, 0))
        self.category_filter_combo = ttk.Combobox(
            sort_frame, values=[t("common_all_categories")] + [translate_category_name(c) for c in RecipeFormWindow.CATEGORY_OPTIONS],
            state="readonly", width=16
        )
        self.category_filter_combo.set(t("common_all_categories"))
        self.category_filter_combo.pack(side="left", padx=5)
        self.category_filter_combo.bind("<<ComboboxSelected>>", lambda e: self._populate())

        list_frame = ttk.Frame(self)
        list_frame.pack(pady=5, padx=15, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, width=56, height=15, font=("Segoe UI", sf(9)))
        list_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=list_scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        list_scrollbar.pack(side="right", fill="y")
        self._populate()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text=t("managerecipes_edit_button"), command=self.edit_selected).grid(
            row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("managerecipes_duplicate_button"), command=self.duplicate_selected).grid(
            row=0, column=1, padx=5)
        ttk.Button(btn_frame, text=t("managerecipes_delete_button"), command=self.delete_selected).grid(
            row=0, column=2, padx=5)

    def _clear_quick_filter(self):
        self.quick_filter = None
        self.destroy()
        ManageRecipesWindow(self.app)

    def _matches_quick_filter(self, recipe):
        if self.quick_filter == "favoris":
            return bool(recipe.get("favorite"))
        if self.quick_filter == "rapide":
            try:
                total = float(recipe.get("prep_time") or 0) + float(recipe.get("cook_time") or 0)
            except (TypeError, ValueError):
                return False
            return 0 < total <= 30
        if self.quick_filter == "vegetarien":
            # Les étiquettes sont du texte libre saisi par l'utilisateur,
            # jamais traduites automatiquement (contrairement aux
            # catégories, allergènes...). Mais ce filtre rapide est une
            # fonctionnalité intégrée à l'application, pas une étiquette
            # quelconque : il reconnaît donc le mot "végétarien" dans les
            # 4 langues disponibles, pour fonctionner même si vous avez
            # tagué vos recettes dans une langue autre que le français.
            tag_keys = {ingredient_sort_key(t) for t in recipe.get("tags", [])}
            vegetarian_keywords = {
                "vegetarien", "vegetarienne",  # français
                "vegetarian",  # anglais
                "vegetariano", "vegetariana",  # espagnol
                "vegetarisch",  # allemand
            }
            return bool(tag_keys & vegetarian_keywords)
        if self.quick_filter == "envie":
            return bool(recipe.get("wishlist"))
        return True

    def _populate(self):
        self.listbox.delete(0, tk.END)
        self.filtered_indices = []
        search = self.search_entry.get().strip() if hasattr(self, "search_entry") else ""
        search_key = ingredient_sort_key(search) if search else ""
        option = resolve_sort_option_input(self.sort_combo.get(), RECIPE_SORT_OPTIONS)
        all_categories_label = t("common_all_categories")
        category_filter = (
            self.category_filter_combo.get() if hasattr(self, "category_filter_combo") else all_categories_label
        )
        category_filter = resolve_category_input(category_filter, RecipeFormWindow.CATEGORY_OPTIONS)
        indexed = list(enumerate(self.app.recipes))
        indexed = [pair for pair in indexed if recipe_matches_search(pair[1], search_key)]
        indexed = [pair for pair in indexed if self._matches_quick_filter(pair[1])]
        if category_filter and category_filter != all_categories_label:
            indexed = [pair for pair in indexed if pair[1].get("category", "Autre") == category_filter]
        reverse = option in ("Ajoutées récemment",)
        indexed.sort(key=lambda pair: recipe_sort_key(pair[1], option), reverse=reverse)
        for idx, recipe in indexed:
            self.listbox.insert(tk.END, format_recipe_list_label(recipe))
            self.filtered_indices.append(idx)

    def _selected_index(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showinfo(t("common_info"), t("managerecipes_select_recipe_first"))
            return None
        return self.filtered_indices[sel[0]]

    def edit_selected(self):
        idx = self._selected_index()
        if idx is None:
            return
        self.destroy()
        RecipeFormWindow(self.app, recipe_index=idx)

    def duplicate_selected(self):
        idx = self._selected_index()
        if idx is None:
            return
        recipes = load_recipes()
        original = recipes[idx]
        new_recipe = copy.deepcopy(original)
        new_recipe["name"] = f"{original['name']} {t('managerecipes_duplicate_suffix')}"
        new_recipe["images"] = duplicate_recipe_images(original)
        new_recipe.pop("image", None)
        recipes.append(new_recipe)
        save_recipes(recipes)
        self.app.refresh_recipes()
        self._populate()
        messagebox.showinfo(
            t("managerecipes_duplicated_title"),
            t("managerecipes_duplicated_message", original=original['name'], new=new_recipe['name'])
        )

    def delete_selected(self):
        idx = self._selected_index()
        if idx is None:
            return
        recipe = self.app.recipes[idx]
        if not messagebox.askyesno(
            t("common_confirm"),
            t("managerecipes_delete_confirm_message", name=recipe['name'])
        ):
            return
        recipes = load_recipes()
        removed = recipes.pop(idx)
        save_recipes(recipes)
        move_recipe_to_trash(removed)
        self.app.refresh_recipes()
        self._populate()
        messagebox.showinfo(t("managerecipes_deleted_title"), t("managerecipes_deleted_message"))


class TrashWindow(tk.Toplevel):
    """Corbeille : liste les recettes supprimées récemment, avec la
    possibilité de les restaurer ou de les effacer définitivement."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("trash_title"))
        self.geometry(f"{gs(560)}x{gs(520)}")
        self.grab_set()

        ttk.Label(self, text=t("trash_heading"), font=("Segoe UI", sf(13), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self,
            text=t("trash_intro"),
            justify="center", font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED
        ).pack(pady=(0, 10))

        list_frame = ttk.Frame(self)
        list_frame.pack(padx=15, pady=5, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, width=56, height=14, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._populate()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text=t("trash_restore_button"), command=self.restore_selected).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("trash_delete_forever_button"),
                   command=self.delete_selected_forever).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text=t("trash_empty_button"),
                   command=self.empty_trash).grid(row=0, column=2, padx=5)

    def _populate(self):
        self.listbox.delete(0, tk.END)
        self.trash = load_trash()
        for entry in self.trash:
            recipe = entry.get("recipe", {})
            try:
                deleted_at = datetime.fromisoformat(entry.get("deleted_at", ""))
                date_display = deleted_at.strftime("%d/%m/%Y à %H:%M")
            except ValueError:
                date_display = t("trash_unknown_date")
            self.listbox.insert(
                tk.END,
                t("trash_entry_line", name=recipe.get('name', t('trash_unnamed_recipe')), date=date_display)
            )
        if not self.trash:
            self.listbox.insert(tk.END, t("trash_is_empty"))

    def _selected_index(self):
        sel = self.listbox.curselection()
        if not sel or not self.trash:
            messagebox.showinfo(t("common_info"), t("trash_select_recipe_first"))
            return None
        return sel[0]

    def restore_selected(self):
        idx = self._selected_index()
        if idx is None:
            return
        entry = self.trash[idx]
        recipe = entry["recipe"]

        recipes = load_recipes()
        existing_names_lower = {r["name"].strip().lower() for r in recipes}
        if recipe.get("name", "").strip().lower() in existing_names_lower:
            recipe["name"] = t("trash_restored_suffix", name=recipe['name'])
        recipes.append(recipe)
        save_recipes(recipes)

        trash = load_trash()
        trash.pop(idx)
        save_trash(trash)

        self.app.refresh_recipes()
        self._populate()
        messagebox.showinfo(t("trash_restored_title"), t("trash_restored_message", name=recipe['name']))

    def delete_selected_forever(self):
        idx = self._selected_index()
        if idx is None:
            return
        entry = self.trash[idx]
        recipe = entry["recipe"]
        if not messagebox.askyesno(
            t("common_confirm"),
            t("trash_delete_forever_confirm", name=recipe['name'])
        ):
            return
        delete_recipe_images(recipe)
        trash = load_trash()
        trash.pop(idx)
        save_trash(trash)
        self._populate()
        messagebox.showinfo(t("trash_deleted_title"), t("trash_deleted_message"))

    def empty_trash(self):
        trash = load_trash()
        if not trash:
            messagebox.showinfo(t("common_info"), t("trash_already_empty"))
            return
        if not messagebox.askyesno(
            t("common_confirm"),
            t("trash_empty_confirm", count=len(trash))
        ):
            return
        for entry in trash:
            delete_recipe_images(entry["recipe"])
        save_trash([])
        self._populate()
        messagebox.showinfo(t("trash_emptied_title"), t("trash_emptied_message"))


class ManageIngredientsWindow(tk.Toplevel):
    """Fenêtre pour ajouter, renommer ou supprimer un ingrédient de la liste
    réutilisable dans les recettes."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("manageing_title"))
        self.geometry(f"{gs(420)}x{gs(680)}")
        self.grab_set()

        ttk.Label(self, text=t("manageing_list_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(10, 5))

        search_frame = ttk.Frame(self)
        search_frame.pack(pady=(0, 5), fill="x", padx=15)
        ttk.Label(search_frame, text=t("common_search_label")).pack(side="left")
        self.search_entry = ttk.Entry(search_frame, width=28)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate())

        list_frame = ttk.Frame(self)
        list_frame.pack(pady=5, padx=15, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, width=36, height=14, font=("Segoe UI", sf(9)))
        list_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=list_scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        list_scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<Double-Button-1>", lambda e: self.edit_selected())
        self._populate()

        add_frame = ttk.Frame(self)
        add_frame.pack(pady=10)
        self.new_entry = ttk.Entry(add_frame, width=25)
        self.new_entry.grid(row=0, column=0, padx=5)
        ttk.Button(add_frame, text=t("manageing_add_button"), command=self.add_ingredient).grid(row=0, column=1)
        self.new_entry.bind("<Return>", lambda e: self.add_ingredient())

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("manageing_edit_button"), command=self.edit_selected).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("manageing_delete_button"), command=self.delete_selected).grid(row=0, column=1, padx=5)

        ttk.Button(self, text=t("manageing_load_defaults_button"),
                   command=self.load_defaults).pack(pady=(5, 5))
        ttk.Button(self, text=t("manageing_spell_check_button"),
                   command=self.open_spell_check).pack(pady=(0, 5))
        ttk.Button(self, text=t("manageing_prices_button"),
                   command=self.open_prices).pack(pady=(0, 5))
        ttk.Button(self, text=t("manageing_substitutions_button"),
                   command=self.open_substitutions).pack(pady=(0, 10))

        ttk.Label(self, text=t("manageing_edit_hint"),
                  font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center").pack(pady=(0, 10))

    def _populate(self):
        self.listbox.delete(0, tk.END)
        search = self.search_entry.get().strip() if hasattr(self, "search_entry") else ""
        if search:
            search_key = ingredient_sort_key(search)
            names = [
                n for n in self.app.ingredient_names
                if search_key in ingredient_sort_key(n) or search_key in ingredient_sort_key(translate_ingredient_name(n))
            ]
        else:
            names = self.app.ingredient_names
        self._names = names
        for name in names:
            self.listbox.insert(tk.END, translate_ingredient_name(name))

    def _selected_name(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showinfo(t("common_info"), t("manageing_select_ingredient_first"))
            return None
        return self._names[sel[0]]

    def add_ingredient(self):
        prefill_name = normalize_oe(self.new_entry.get().strip())
        self.new_entry.delete(0, tk.END)
        IngredientEditWindow(self.app, manage_window=self, existing_name=None,
                              prefill_name=prefill_name)

    def edit_selected(self):
        name = self._selected_name()
        if name is None:
            return
        IngredientEditWindow(self.app, manage_window=self, existing_name=name)

    def delete_selected(self):
        name = self._selected_name()
        if name is None:
            return
        usage = count_ingredient_usage(name)
        message = t("manageing_delete_confirm_message", name=translate_ingredient_name(name))
        if usage:
            message += t("manageing_delete_usage_warning", count=usage)
        if not messagebox.askyesno(t("common_confirm"), message):
            return
        ingredients = [n for n in load_ingredients() if n.lower() != name.lower()]
        self.app.ingredient_names = save_ingredients(ingredients)
        self._populate()

    def load_defaults(self):
        if not os.path.exists(DEFAULT_INGREDIENTS_FILE):
            messagebox.showerror(
                t("manageing_missing_file_title"),
                t("manageing_missing_file_message")
            )
            return
        added = merge_default_ingredients()
        self.app.ingredient_names = load_ingredients()
        self._populate()
        if added:
            messagebox.showinfo(
                t("manageing_done_title"),
                t("manageing_defaults_added_message", count=added)
            )
        else:
            messagebox.showinfo(t("manageing_done_title"), t("manageing_defaults_none_added"))

    def open_spell_check(self):
        IngredientSpellCheckWindow(self.app, manage_window=self)

    def open_prices(self):
        IngredientPricesWindow(self.app)

    def open_substitutions(self):
        ManageSubstitutionsWindow(self.app)


class SubstitutionEditWindow(tk.Toplevel):
    """Fenêtre pour consulter/modifier la liste des substituts suggérés pour
    un ingrédient précis."""

    def __init__(self, app, ingredient_name, parent_window=None):
        super().__init__(parent_window or app)
        self.app = app
        self.ingredient_name = ingredient_name
        self.title(t("subedit_title", name=ingredient_name))
        self.geometry(f"{gs(480)}x{gs(650)}")
        self.minsize(gs(420), gs(550))
        self.resizable(True, True)
        self.grab_set()

        ttk.Label(self, text=t("subedit_heading", name=ingredient_name),
                  font=("Segoe UI", sf(12), "bold"), wraplength=440, justify="center").pack(pady=(15, 5))
        ttk.Label(
            self, text=t("subedit_disclaimer"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

        self.current_list = [dict(s) for s in get_ingredient_substitutions(ingredient_name)]

        list_frame = ttk.Frame(self)
        list_frame.pack(padx=15, pady=5, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, height=6, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._refresh_listbox()

        ttk.Button(self, text=t("subedit_remove_button"),
                   command=self.remove_selected).pack(pady=(0, 10))

        add_frame = ttk.LabelFrame(self, text=t("subedit_add_frame_title"))
        add_frame.pack(padx=15, pady=(0, 10), fill="x")
        ttk.Label(add_frame, text=t("subedit_name_label")).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.new_name_entry = ttk.Entry(add_frame, width=28)
        self.new_name_entry.full_values = get_display_ingredient_values(sorted(self.app.ingredient_names, key=ingredient_sort_key))
        self.new_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.new_name_entry.bind("<KeyRelease>", lambda e: self._on_name_entry_keyrelease(e))
        self.new_name_entry.bind("<FocusIn>", lambda e: self._on_name_entry_focus_in(e))
        self.new_name_entry.bind("<FocusOut>", lambda e: self._on_name_entry_focus_out(e))
        ttk.Label(add_frame, text=t("subedit_note_label")).grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.new_note_entry = ttk.Entry(add_frame, width=28)
        self.new_note_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(add_frame, text=t("subedit_add_to_list_button"), command=self.add_substitute).grid(
            row=2, column=0, columnspan=2, pady=(5, 5))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("common_save_button"), command=self.save_and_close).grid(row=0, column=0, padx=5)
        if load_default_substitutions().get(ingredient_name.strip().lower()):
            ttk.Button(btn_frame, text=t("subedit_revert_button"),
                       command=self.revert_to_default).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text=t("subedit_cancel_button"), style="Secondary.TButton",
                   command=self.destroy).grid(row=0, column=2, padx=5)

        tk.Frame(self, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        if not self.current_list:
            self.listbox.insert(tk.END, t("subedit_no_substitute_yet"))
            return
        for sub in self.current_list:
            note = f" — {sub['note']}" if sub.get("note") else ""
            self.listbox.insert(tk.END, f"{sub['nom']}{note}")

    def add_substitute(self):
        nom = self.new_name_entry.get().strip()
        if not nom:
            messagebox.showerror(t("common_error"), t("subedit_error_name_required"))
            return
        note = self.new_note_entry.get().strip()
        self.current_list.append({"nom": nom, "note": note})
        self._refresh_listbox()
        self.new_name_entry.delete(0, tk.END)
        self.new_note_entry.delete(0, tk.END)
        self.new_name_entry.focus_set()

    def remove_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.current_list:
            messagebox.showinfo(t("common_info"), t("subedit_select_to_remove"))
            return
        del self.current_list[sel[0]]
        self._refresh_listbox()

    def save_and_close(self):
        set_ingredient_override(self.ingredient_name, substitutions=self.current_list)
        self.destroy()

    def revert_to_default(self):
        if messagebox.askyesno(
            t("common_confirm"),
            t("subedit_revert_confirm_message", name=self.ingredient_name)
        ):
            revert_ingredient_substitutions_to_default(self.ingredient_name)
            self.destroy()

    # ---- Autocomplétion du champ Nom (même principe que les autres listes
    # déroulantes d'ingrédients de l'application) ----

    def _hide_name_suggestions(self):
        popup = getattr(self.new_name_entry, "_suggestion_popup", None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
            self.new_name_entry._suggestion_popup = None
            self.new_name_entry._suggestion_listbox = None

    def _show_name_suggestions(self, filtered):
        self._hide_name_suggestions()
        if not filtered:
            return
        entry = self.new_name_entry
        popup = tk.Toplevel(entry)
        popup.wm_overrideredirect(True)
        try:
            popup.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        width = max(entry.winfo_width(), 160)
        height = min(6, len(filtered)) * 20
        popup.wm_geometry(f"{width}x{height}+{x}+{y}")

        listbox = tk.Listbox(popup, height=min(6, len(filtered)), exportselection=False, font=("Segoe UI", sf(9)))
        listbox.pack(fill="both", expand=True)
        for v in filtered:
            listbox.insert(tk.END, v)

        def choose(event=None):
            sel = listbox.curselection()
            if sel:
                value = listbox.get(sel[0])
                entry.delete(0, tk.END)
                entry.insert(0, value)
            self._hide_name_suggestions()
            entry.focus_set()

        listbox.bind("<ButtonRelease-1>", choose)
        listbox.bind("<Return>", choose)
        entry._suggestion_popup = popup
        entry._suggestion_listbox = listbox

    def _on_name_entry_keyrelease(self, event):
        if event.keysym == "Down":
            listbox = getattr(self.new_name_entry, "_suggestion_listbox", None)
            if listbox is not None:
                listbox.focus_set()
                listbox.selection_set(0)
            return
        if event.keysym == "Escape":
            self._hide_name_suggestions()
            return
        if event.keysym == "Return":
            self._hide_name_suggestions()
            self.add_substitute()
            return
        if event.keysym in ("Tab", "Shift_L", "Shift_R", "Control_L", "Control_R",
                            "Caps_Lock", "Alt_L", "Alt_R", "Left", "Right"):
            return
        filtered = self._filter_ingredient_values(self.new_name_entry.full_values, self.new_name_entry.get())
        if filtered:
            self._show_name_suggestions(filtered)
        else:
            self._hide_name_suggestions()

    def _on_name_entry_focus_in(self, event):
        filtered = self._filter_ingredient_values(self.new_name_entry.full_values, self.new_name_entry.get())
        if filtered:
            self._show_name_suggestions(filtered)

    def _on_name_entry_focus_out(self, event):
        self.new_name_entry.after(200, self._hide_name_suggestions)

    @staticmethod
    def _filter_ingredient_values(full_values, typed):
        if not typed:
            return full_values
        typed_key = ingredient_sort_key(typed)
        filtered = [v for v in full_values if ingredient_sort_key(v).startswith(typed_key)]
        if not filtered:
            filtered = [v for v in full_values if typed_key in ingredient_sort_key(v)]
        return filtered


class ManageSubstitutionsWindow(tk.Toplevel):
    """Fenêtre pour consulter et gérer les substituts d'ingrédients suggérés
    par l'application, ou en ajouter/modifier pour vos propres besoins."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("managesub_title"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(480)}x{min(screen_height, gs(700))}+40+20")
        self.minsize(gs(420), gs(460))
        self.resizable(True, True)
        self.grab_set()

        ttk.Label(self, text=t("managesub_heading"), font=("Segoe UI", sf(14), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self, text=t("managesub_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

        search_frame = ttk.Frame(self)
        search_frame.pack(pady=(0, 5), fill="x", padx=15)
        ttk.Label(search_frame, text=t("common_search_label")).pack(side="left")
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate())

        list_frame = ttk.Frame(self)
        list_frame.pack(pady=5, padx=15, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<Double-Button-1>", lambda e: self.edit_selected())

        self.displayed_names = []
        self._populate()

        add_frame = ttk.Frame(self)
        add_frame.pack(pady=(5, 5), padx=15, fill="x")
        ttk.Label(add_frame, text=t("common_ingredient_label")).pack(side="left")
        self.name_entry = ttk.Entry(add_frame, width=20)
        self.name_entry.full_values = get_display_ingredient_values(sorted(self.app.ingredient_names, key=ingredient_sort_key))
        self.name_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.name_entry.bind("<KeyRelease>", lambda e: self._on_name_entry_keyrelease(e))
        self.name_entry.bind("<FocusIn>", lambda e: self._on_name_entry_focus_in(e))
        self.name_entry.bind("<FocusOut>", lambda e: self._on_name_entry_focus_out(e))
        ttk.Button(add_frame, text=t("managesub_manage_button"), command=self.edit_typed).pack(side="left", padx=(5, 0))

        ttk.Label(
            self, text=t("managesub_hint"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(10, 5))

        tk.Frame(self, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _populate(self):
        self.listbox.delete(0, tk.END)
        search = self.search_entry.get().strip()
        search_key = ingredient_sort_key(search) if search else ""
        keys_with_subs = has_known_substitutions()
        name_by_key = {ingredient_sort_key(n): n for n in self.app.ingredient_names}
        self.displayed_names = []
        for key in sorted(keys_with_subs):
            name = name_by_key.get(key, key.capitalize())
            if search_key and search_key not in ingredient_sort_key(name) \
                    and search_key not in ingredient_sort_key(translate_ingredient_name(name)):
                continue
            self.displayed_names.append(name)
        if not self.displayed_names:
            self.listbox.insert(tk.END, t("managesub_none_with_substitute"))
            return
        for name in self.displayed_names:
            count = len(get_ingredient_substitutions(name))
            plural = "s" if count > 1 else ""
            self.listbox.insert(
                tk.END, t("managesub_substitute_count", name=translate_ingredient_name(name), count=count, plural=plural)
            )

    def edit_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.displayed_names:
            messagebox.showinfo(t("common_info"), t("manageing_select_ingredient_first"))
            return
        name = self.displayed_names[sel[0]]
        self._open_editor(name)

    def edit_typed(self):
        typed = normalize_oe(self.name_entry.get().strip())
        if not typed:
            messagebox.showerror(t("common_error"), t("managesub_error_ingredient_required"))
            return
        canonical = resolve_ingredient_input(typed, self.app.ingredient_names)
        if canonical is None:
            messagebox.showerror(
                t("common_unknown_ingredient_title"),
                t("managesub_unknown_ingredient_message", name=typed)
            )
            return
        self._open_editor(canonical)

    def _open_editor(self, name):
        win = SubstitutionEditWindow(self.app, name, parent_window=self)
        self.wait_window(win)
        self._populate()

    # ---- Autocomplétion du champ ingrédient (même principe que les autres
    # listes déroulantes d'ingrédients de l'application) ----

    def _hide_name_suggestions(self):
        popup = getattr(self.name_entry, "_suggestion_popup", None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
            self.name_entry._suggestion_popup = None
            self.name_entry._suggestion_listbox = None

    def _show_name_suggestions(self, filtered):
        self._hide_name_suggestions()
        if not filtered:
            return
        entry = self.name_entry
        popup = tk.Toplevel(entry)
        popup.wm_overrideredirect(True)
        try:
            popup.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        width = max(entry.winfo_width(), 160)
        height = min(6, len(filtered)) * 20
        popup.wm_geometry(f"{width}x{height}+{x}+{y}")

        listbox = tk.Listbox(popup, height=min(6, len(filtered)), exportselection=False, font=("Segoe UI", sf(9)))
        listbox.pack(fill="both", expand=True)
        for v in filtered:
            listbox.insert(tk.END, v)

        def choose(event=None):
            sel = listbox.curselection()
            if sel:
                value = listbox.get(sel[0])
                entry.delete(0, tk.END)
                entry.insert(0, value)
            self._hide_name_suggestions()
            entry.focus_set()

        listbox.bind("<ButtonRelease-1>", choose)
        listbox.bind("<Return>", choose)
        entry._suggestion_popup = popup
        entry._suggestion_listbox = listbox

    def _on_name_entry_keyrelease(self, event):
        if event.keysym == "Down":
            listbox = getattr(self.name_entry, "_suggestion_listbox", None)
            if listbox is not None:
                listbox.focus_set()
                listbox.selection_set(0)
            return
        if event.keysym == "Escape":
            self._hide_name_suggestions()
            return
        if event.keysym == "Return":
            self._hide_name_suggestions()
            self.edit_typed()
            return
        if event.keysym in ("Tab", "Shift_L", "Shift_R", "Control_L", "Control_R",
                            "Caps_Lock", "Alt_L", "Alt_R", "Left", "Right"):
            return
        filtered = self._filter_ingredient_values(self.name_entry.full_values, self.name_entry.get())
        if filtered:
            self._show_name_suggestions(filtered)
        else:
            self._hide_name_suggestions()

    def _on_name_entry_focus_in(self, event):
        filtered = self._filter_ingredient_values(self.name_entry.full_values, self.name_entry.get())
        if filtered:
            self._show_name_suggestions(filtered)

    def _on_name_entry_focus_out(self, event):
        self.name_entry.after(200, self._hide_name_suggestions)

    @staticmethod
    def _filter_ingredient_values(full_values, typed):
        if not typed:
            return full_values
        typed_key = ingredient_sort_key(typed)
        filtered = [v for v in full_values if ingredient_sort_key(v).startswith(typed_key)]
        if not filtered:
            filtered = [v for v in full_values if typed_key in ingredient_sort_key(v)]
        return filtered


class IngredientPricesWindow(tk.Toplevel):
    """Fenêtre pour renseigner le prix de vos ingrédients, utilisé ensuite
    pour estimer le coût de vos recettes. Les prix sont saisis par vous —
    aucune source de prix en ligne n'est utilisée."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("ingprices_title"))
        self.geometry(f"{gs(480)}x{gs(620)}")
        self.grab_set()

        ttk.Label(self, text=t("ingprices_heading"), font=("Segoe UI", sf(13), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self,
            text=t("ingprices_intro"),
            justify="center", font=("Segoe UI", sf(9))
        ).pack(pady=(0, 10))

        search_frame = ttk.Frame(self)
        search_frame.pack(pady=(0, 5), fill="x", padx=15)
        ttk.Label(search_frame, text=t("common_search_label")).pack(side="left")
        self.search_entry = ttk.Entry(search_frame, width=28)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate())

        list_frame = ttk.Frame(self)
        list_frame.pack(pady=5, padx=15, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, width=48, height=13, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._load_selected_price())
        self._populate()

        edit_frame = ttk.Frame(self)
        edit_frame.pack(pady=10, padx=15, fill="x")
        ttk.Label(edit_frame, text=t("ingprices_price_label")).grid(row=0, column=0, padx=3)
        self.price_entry = ttk.Entry(edit_frame, width=8)
        self.price_entry.grid(row=0, column=1, padx=3)
        ttk.Label(edit_frame, text=t("ingprices_for_one_label")).grid(row=0, column=2, padx=3)
        self.unit_combo = ttk.Combobox(edit_frame, values=[translate_unit_name(u) for u in PRICE_UNIT_OPTIONS],
                                        state="readonly", width=15)
        self.unit_combo.set(translate_unit_name(PRICE_UNIT_OPTIONS[0]))
        self.unit_combo.grid(row=0, column=3, padx=3)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("ingprices_save_button"),
                   command=self.save_price).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("ingprices_clear_button"),
                   command=self.clear_price).grid(row=0, column=1, padx=5)

        ttk.Label(
            self,
            text=t("ingprices_units_note"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

    def _populate(self):
        self.listbox.delete(0, tk.END)
        search = self.search_entry.get().strip()
        search_key = ingredient_sort_key(search) if search else ""
        self._names = []
        for name in self.app.ingredient_names:
            if search_key and search_key not in ingredient_sort_key(name) \
                    and search_key not in ingredient_sort_key(translate_ingredient_name(name)):
                continue
            self._names.append(name)
            price_info = get_ingredient_price(name)
            if price_info:
                suffix = t("ingprices_price_suffix", price=f"{price_info['price']:.2f}", unit=price_info['unit'])
            else:
                suffix = t("ingprices_no_price_set")
            self.listbox.insert(tk.END, f"{translate_ingredient_name(name)}{suffix}")

    def _selected_name(self):
        sel = self.listbox.curselection()
        if not sel or not self._names:
            return None
        return self._names[sel[0]]

    def _load_selected_price(self):
        name = self._selected_name()
        if name is None:
            return
        price_info = get_ingredient_price(name)
        self.price_entry.delete(0, tk.END)
        if price_info:
            self.price_entry.insert(0, str(price_info["price"]))
            self.unit_combo.set(translate_unit_name(price_info["unit"]))
        else:
            self.unit_combo.set(translate_unit_name(PRICE_UNIT_OPTIONS[0]))

    def save_price(self):
        name = self._selected_name()
        if name is None:
            messagebox.showinfo(t("common_info"), t("manageing_select_ingredient_first"))
            return
        raw_price = self.price_entry.get().strip().replace(",", ".")
        try:
            price = float(raw_price)
            if price < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(t("common_error"), t("ingprices_error_invalid_price"))
            return
        set_ingredient_price(name, price, resolve_unit_input(self.unit_combo.get(), PRICE_UNIT_OPTIONS))
        self._populate()
        messagebox.showinfo(t("allrecipes_list_saved_title"), t("ingprices_saved_message", name=name))

    def clear_price(self):
        name = self._selected_name()
        if name is None:
            messagebox.showinfo(t("common_info"), t("manageing_select_ingredient_first"))
            return
        set_ingredient_price(name, None, None)
        self.price_entry.delete(0, tk.END)
        self._populate()


class IngredientEditWindow(tk.Toplevel):
    """Fenêtre unifiée pour ajouter un nouvel ingrédient ou modifier un
    ingrédient existant : nom, allergènes, valeurs nutritionnelles et prix."""

    def __init__(self, app, manage_window=None, existing_name=None, prefill_name="", parent_window=None):
        # `parent_window` permet de préciser la vraie fenêtre parente Tkinter
        # (ex. la fenêtre qui a ouvert cet éditeur), différente de `app` qui
        # sert uniquement de référence aux données. Sans cela, fermer cette
        # fenêtre pourrait faire remonter la page d'accueil au premier plan
        # au lieu de la fenêtre depuis laquelle elle a été ouverte.
        super().__init__(parent_window or app)
        self.app = app
        self.manage_window = manage_window
        self.existing_name = existing_name
        self.editing = existing_name is not None
        self.title(t("ingedit_title_edit") if self.editing else t("ingedit_title_new"))
        self.geometry(f"{gs(480)}x{gs(680)}")
        self.grab_set()

        ttk.Label(self, text=t("ingedit_heading_edit") if self.editing else t("ingedit_heading_new"),
                  font=("Segoe UI", sf(13), "bold")).pack(pady=(15, 10))

        ttk.Label(self, text=t("ingedit_name_label"), font=("Segoe UI", sf(10), "bold")).pack()
        self.name_entry = ttk.Entry(self, width=40)
        self.name_entry.pack(pady=(2, 10))
        self.name_entry.insert(0, existing_name if self.editing else prefill_name)

        ttk.Label(self, text=t("ingedit_allergens_label"), font=("Segoe UI", sf(10), "bold")).pack(pady=(5, 5))
        allergens_frame = ttk.Frame(self)
        allergens_frame.pack()
        existing_allergens = set(get_ingredient_allergens(existing_name)) if self.editing else set()
        self.allergen_vars = {}
        for i, allergen in enumerate(ALLERGENS):
            var = tk.BooleanVar(value=allergen in existing_allergens)
            self.allergen_vars[allergen] = var
            ttk.Checkbutton(allergens_frame, text=translate_allergen_name(allergen), variable=var).grid(
                row=i // 3, column=i % 3, sticky="w", padx=8, pady=2
            )

        ttk.Label(self, text=t("ingedit_nutrition_label"),
                  font=("Segoe UI", sf(10), "bold")).pack(pady=(15, 5))
        nutri_frame = ttk.Frame(self)
        nutri_frame.pack()
        existing_nutri = (get_ingredient_nutrition(existing_name) or {}) if self.editing else {}
        nutri_labels = [("kcal", t("ingedit_nutri_kcal")), ("protein_g", t("ingedit_nutri_protein")),
                         ("carbs_g", t("ingedit_nutri_carbs")), ("fat_g", t("ingedit_nutri_fat"))]
        self.nutri_entries = {}
        for i, (key, label) in enumerate(nutri_labels):
            ttk.Label(nutri_frame, text=f"{label} :").grid(row=i, column=0, sticky="e", padx=5, pady=3)
            entry = ttk.Entry(nutri_frame, width=10)
            if key in existing_nutri:
                entry.insert(0, str(existing_nutri[key]))
            entry.grid(row=i, column=1, sticky="w", padx=5, pady=3)
            self.nutri_entries[key] = entry
        ttk.Label(self, text=t("ingedit_nutrition_hint"),
                  font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED).pack()

        ttk.Label(self, text=t("ingedit_price_label"), font=("Segoe UI", sf(10), "bold")).pack(pady=(15, 5))
        price_frame = ttk.Frame(self)
        price_frame.pack()
        existing_price = get_ingredient_price(existing_name) if self.editing else None
        ttk.Label(price_frame, text=t("ingprices_price_label")).grid(row=0, column=0, padx=3)
        self.price_entry = ttk.Entry(price_frame, width=8)
        if existing_price:
            self.price_entry.insert(0, str(existing_price["price"]))
        self.price_entry.grid(row=0, column=1, padx=3)
        ttk.Label(price_frame, text=t("ingprices_for_one_label")).grid(row=0, column=2, padx=3)
        self.unit_combo = ttk.Combobox(price_frame, values=[translate_unit_name(u) for u in PRICE_UNIT_OPTIONS], state="readonly", width=15)
        self.unit_combo.set(translate_unit_name(existing_price["unit"]) if existing_price else translate_unit_name(PRICE_UNIT_OPTIONS[0]))
        self.unit_combo.grid(row=0, column=3, padx=3)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text=t("ingedit_save_button"), command=self.save).grid(row=0, column=0, padx=5)
        if self.editing:
            ttk.Button(btn_frame, text=t("ingedit_delete_button"),
                       command=self.delete_ingredient).grid(row=0, column=1, padx=5)

    def _parse_float_or_none(self, entry, field_label):
        raw = entry.get().strip().replace(",", ".")
        if not raw:
            return None, True
        try:
            value = float(raw)
            if value < 0:
                raise ValueError
            return value, True
        except ValueError:
            messagebox.showerror(t("common_error"), t("ingedit_error_invalid_field", field=field_label))
            return None, False

    def save(self):
        new_name = normalize_oe(self.name_entry.get().strip())
        if not new_name:
            messagebox.showerror(t("common_error"), t("ingedit_error_name_required"))
            return

        other_names = [
            n for n in self.app.ingredient_names
            if not (self.editing and n.lower() == self.existing_name.lower())
        ]
        if new_name.lower() in [n.lower() for n in other_names]:
            messagebox.showerror(t("common_error"), t("ingedit_error_already_exists", name=new_name))
            return

        plural_match = find_plural_duplicate(new_name, other_names)
        if plural_match:
            messagebox.showerror(
                t("common_error"),
                t("ingedit_error_plural_duplicate", name=new_name, existing=plural_match)
            )
            return

        nutrition = {}
        nutri_field_labels = {"kcal": t("ingedit_nutri_field_kcal"), "protein_g": t("ingedit_nutri_field_protein"),
                               "carbs_g": t("ingedit_nutri_field_carbs"), "fat_g": t("ingedit_nutri_field_fat")}
        for key, entry in self.nutri_entries.items():
            value, ok = self._parse_float_or_none(entry, nutri_field_labels[key])
            if not ok:
                return
            if value is not None:
                nutrition[key] = value

        raw_price = self.price_entry.get().strip().replace(",", ".")
        price = None
        if raw_price:
            try:
                price = float(raw_price)
                if price < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(t("common_error"), t("ingedit_error_invalid_price"))
                return

        allergens = [a for a, var in self.allergen_vars.items() if var.get()]

        if self.editing:
            old_name = self.existing_name
            if new_name.lower() != old_name.lower():
                ingredients = [new_name if n.lower() == old_name.lower() else n
                               for n in load_ingredients()]
                self.app.ingredient_names = save_ingredients(ingredients)
                rename_ingredient_everywhere(old_name, new_name)
                rename_ingredient_override(old_name, new_name)
                old_price = get_ingredient_price(old_name)
                if old_price:
                    set_ingredient_price(old_name, None, None)
                    set_ingredient_price(new_name, old_price["price"], old_price["unit"])
                self.app.refresh_recipes()
        else:
            ingredients = load_ingredients()
            ingredients.append(new_name)
            self.app.ingredient_names = save_ingredients(ingredients)

        set_ingredient_override(new_name, allergens=allergens, nutrition=nutrition)
        if raw_price:
            set_ingredient_price(new_name, price, resolve_unit_input(self.unit_combo.get(), PRICE_UNIT_OPTIONS))
        elif self.editing:
            set_ingredient_price(new_name, None, None)

        if self.manage_window is not None:
            self.manage_window.app.ingredient_names = self.app.ingredient_names
            self.manage_window._populate()

        messagebox.showinfo(t("allrecipes_list_saved_title"), t("ingedit_saved_message", name=new_name))
        self.destroy()

    def delete_ingredient(self):
        name = self.existing_name
        usage = count_ingredient_usage(name)
        message = t("manageing_delete_confirm_message", name=translate_ingredient_name(name))
        if usage:
            message += t("manageing_delete_usage_warning", count=usage)
        if not messagebox.askyesno(t("common_confirm"), message):
            return
        ingredients = [n for n in load_ingredients() if n.lower() != name.lower()]
        self.app.ingredient_names = save_ingredients(ingredients)
        set_ingredient_override(name, allergens=[], nutrition={})
        set_ingredient_price(name, None, None)
        if self.manage_window is not None:
            self.manage_window.app.ingredient_names = self.app.ingredient_names
            self.manage_window._populate()
        self.destroy()


class IngredientSpellCheckWindow(tk.Toplevel):
    """Détecte les paires d'ingrédients qui se ressemblent fortement (90 % de
    similarité ou plus — pluriels non fusionnés, fautes de frappe) et propose
    de les fusionner, une par une ou plusieurs à la fois."""

    SIMILARITY_THRESHOLD = 0.90

    def __init__(self, app, manage_window=None):
        super().__init__(app)
        self.app = app
        self.manage_window = manage_window
        self.title(t("spellcheck_title"))
        self.geometry(f"{gs(580)}x{gs(560)}")
        self.grab_set()

        ttk.Label(
            self,
            text=t("spellcheck_heading"),
            font=("Segoe UI", sf(11), "bold"), justify="center"
        ).pack(pady=10)
        ttk.Label(
            self,
            text=t("spellcheck_multi_select_hint"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 5))

        list_frame = ttk.Frame(self)
        list_frame.pack(padx=15, pady=5, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, width=64, height=16, selectmode="extended", font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.pairs = []
        self._scan()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("spellcheck_merge_button"),
                   command=self.merge_selected).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("spellcheck_not_duplicate_button"),
                   command=self.dismiss_selected).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text=t("spellcheck_rerun_button"),
                   command=self._scan).grid(row=0, column=2, padx=5)

        ttk.Label(
            self,
            text=t("spellcheck_footer_hint"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

    def _scan(self):
        self.listbox.delete(0, tk.END)
        all_pairs = find_similar_ingredient_pairs(
            self.app.ingredient_names, threshold=self.SIMILARITY_THRESHOLD
        )
        dismissed = load_dismissed_pairs()
        self.pairs = [
            (name_a, name_b, ratio) for (name_a, name_b, ratio) in all_pairs
            if not is_pair_dismissed(name_a, name_b, dismissed)
        ]
        if not self.pairs:
            self.listbox.insert(tk.END, t("spellcheck_none_found"))
            return
        for name_a, name_b, ratio in self.pairs:
            self.listbox.insert(tk.END, t("spellcheck_pair_line", a=name_a, b=name_b, percent=int(ratio * 100)))

    def dismiss_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.pairs:
            messagebox.showinfo(t("common_info"), t("spellcheck_select_pair_first"))
            return
        selected_pairs = [self.pairs[i] for i in sel]
        for name_a, name_b, ratio in selected_pairs:
            add_dismissed_pair(name_a, name_b)
        self._scan()
        messagebox.showinfo(
            t("common_info"),
            t("spellcheck_dismissed_message", count=len(selected_pairs))
        )

    def _merge_pair(self, keep, remove):
        ingredients = [n for n in load_ingredients() if n.lower() != remove.lower()]
        self.app.ingredient_names = save_ingredients(ingredients)
        rename_ingredient_everywhere(remove, keep)

    def merge_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.pairs:
            messagebox.showinfo(t("common_info"), t("spellcheck_select_pair_first"))
            return
        selected_pairs = [self.pairs[i] for i in sel]

        if len(selected_pairs) == 1:
            name_a, name_b, ratio = selected_pairs[0]
            choice = messagebox.askyesnocancel(
                t("spellcheck_merge_dialog_title"),
                t("spellcheck_merge_dialog_message", a=name_a, b=name_b)
            )
            if choice is None:
                return
            keep, remove = (name_a, name_b) if choice else (name_b, name_a)
            self._merge_pair(keep, remove)
            self.app.refresh_recipes()
            if self.manage_window is not None:
                self.manage_window.app.ingredient_names = self.app.ingredient_names
                self.manage_window._populate()
            messagebox.showinfo(t("spellcheck_merged_title"), t("spellcheck_merged_one_message", removed=remove, kept=keep))
        else:
            if not messagebox.askyesno(
                t("common_confirm"),
                t("spellcheck_merge_multi_confirm", count=len(selected_pairs))
            ):
                return
            for name_a, name_b, ratio in selected_pairs:
                usage_a = count_ingredient_usage(name_a)
                usage_b = count_ingredient_usage(name_b)
                if usage_a > usage_b:
                    keep, remove = name_a, name_b
                elif usage_b > usage_a:
                    keep, remove = name_b, name_a
                else:
                    keep, remove = sorted([name_a, name_b], key=ingredient_sort_key)
                self._merge_pair(keep, remove)
            self.app.refresh_recipes()
            if self.manage_window is not None:
                self.manage_window.app.ingredient_names = self.app.ingredient_names
                self.manage_window._populate()
            messagebox.showinfo(t("spellcheck_merged_title"), t("spellcheck_merged_multi_message", count=len(selected_pairs)))

        self._scan()


def build_full_backup_zip(path):
    """Construit une archive ZIP contenant l'intégralité des données
    utilisateur (recettes, ingrédients personnalisés, garde-manger,
    planning et historique, menus, listes de courses enregistrées,
    corbeille, réglages...) et toutes les photos. Utilisé aussi bien par
    l'export manuel que par les sauvegardes automatiques."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename in USER_DATA_FILES:
            filepath = os.path.join(BASE_DIR, filename)
            if os.path.exists(filepath):
                zf.write(filepath, arcname=filename)
        if os.path.isdir(IMAGES_DIR):
            for fname in os.listdir(IMAGES_DIR):
                full = os.path.join(IMAGES_DIR, fname)
                if os.path.isfile(full):
                    zf.write(full, arcname=f"images/{fname}")


def restore_from_zip(path, merge):
    """Restaure des données à partir d'une archive ZIP (export manuel ou
    sauvegarde automatique). merge=True fusionne avec les données actuelles
    (recettes/ingrédients/photos en double sont renommés plutôt qu'écrasés ;
    les autres données personnelles — garde-manger, prix, substituts,
    listes de courses enregistrées, menus, historique de planning,
    corbeille... — sont ajoutées à celles déjà présentes, sans rien
    perdre) ; merge=False remplace tout, y compris les réglages et le
    planning actif."""
    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()
        imported_recipes = []
        imported_ingredients = []
        if "recipes.json" in names:
            imported_recipes = json.loads(zf.read("recipes.json").decode("utf-8"))
        if "ingredients.json" in names:
            imported_ingredients = json.loads(zf.read("ingredients.json").decode("utf-8"))
        image_entries = [n for n in names if n.startswith("images/") and not n.endswith("/")]

        if not merge:
            if os.path.isdir(IMAGES_DIR):
                for fname in os.listdir(IMAGES_DIR):
                    full = os.path.join(IMAGES_DIR, fname)
                    if os.path.isfile(full):
                        os.remove(full)
            for entry in image_entries:
                fname = os.path.basename(entry)
                with open(os.path.join(IMAGES_DIR, fname), "wb") as out:
                    out.write(zf.read(entry))
            save_recipes(imported_recipes)
            save_ingredients(imported_ingredients)
        else:
            rename_map = {}
            for entry in image_entries:
                old_fname = os.path.basename(entry)
                ext = os.path.splitext(old_fname)[1]
                new_fname = f"{uuid.uuid4().hex}{ext}"
                with open(os.path.join(IMAGES_DIR, new_fname), "wb") as out:
                    out.write(zf.read(entry))
                rename_map[old_fname] = new_fname

            existing_recipes = load_recipes()
            existing_names_lower = {r["name"].strip().lower() for r in existing_recipes}
            for r in imported_recipes:
                img = r.get("image")
                if img and img in rename_map:
                    r["image"] = rename_map[img]
                imgs = r.get("images")
                if imgs:
                    r["images"] = [rename_map.get(f, f) for f in imgs]
                if r.get("name", "").strip().lower() in existing_names_lower:
                    r["name"] = f"{r['name']} (importé)"
                existing_recipes.append(r)
                existing_names_lower.add(r["name"].strip().lower())
            save_recipes(existing_recipes)

            existing_ingredients = load_ingredients()
            save_ingredients(existing_ingredients + imported_ingredients)

        # ---- Le reste des données personnelles : garde-manger, prix et
        # substituts personnalisés, listes de courses enregistrées, menus,
        # historique/modèles de planning, corbeille, historique récent,
        # doublons d'ingrédients ignorés. Chaque type de fichier a sa
        # propre logique de fusion adaptée à sa structure. ----
        dict_merge_files = [
            "ingredient_custom_data.json", "ingredient_prices.json", "pantry.json",
            "weekly_plan_templates.json",
        ]
        named_list_merge_files = {
            "saved_shopping_lists.json": "name",
            "menus.json": "name",
            "weekly_plan_history.json": "week_start",
        }
        simple_list_merge_files = ["trash.json", "recent_views.json"]
        replace_only_files = ["settings.json", "weekly_plan.json"]

        def _read_json_entry(filename, expected_type):
            try:
                data = json.loads(zf.read(filename).decode("utf-8"))
                return data if isinstance(data, expected_type) else expected_type()
            except Exception:
                return expected_type()

        def _read_existing_json(filepath, expected_type):
            if not os.path.exists(filepath):
                return expected_type()
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    return data if isinstance(data, expected_type) else expected_type()
            except Exception:
                return expected_type()

        for filename in dict_merge_files:
            if filename not in names:
                continue
            imported = _read_json_entry(filename, dict)
            filepath = os.path.join(BASE_DIR, filename)
            if merge:
                existing = _read_existing_json(filepath, dict)
                existing.update(imported)
                imported = existing
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(imported, f, ensure_ascii=False, indent=2)

        for filename, key_field in named_list_merge_files.items():
            if filename not in names:
                continue
            imported = _read_json_entry(filename, list)
            filepath = os.path.join(BASE_DIR, filename)
            if merge:
                existing = _read_existing_json(filepath, list)
                by_key = {item.get(key_field): item for item in existing if isinstance(item, dict)}
                for item in imported:
                    if isinstance(item, dict):
                        by_key[item.get(key_field)] = item
                imported = list(by_key.values())
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(imported, f, ensure_ascii=False, indent=2)

        for filename in simple_list_merge_files:
            if filename not in names:
                continue
            imported = _read_json_entry(filename, list)
            filepath = os.path.join(BASE_DIR, filename)
            if merge:
                imported = _read_existing_json(filepath, list) + imported
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(imported, f, ensure_ascii=False, indent=2)

        if "ingredient_dismissed_pairs.json" in names:
            imported_pairs = _read_json_entry("ingredient_dismissed_pairs.json", list)
            pairs = {tuple(sorted(p)) for p in imported_pairs if isinstance(p, list) and len(p) == 2}
            if merge:
                pairs |= load_dismissed_pairs()
            save_dismissed_pairs(pairs)

        if not merge:
            for filename in replace_only_files:
                if filename in names:
                    imported = json.loads(zf.read(filename).decode("utf-8"))
                    filepath = os.path.join(BASE_DIR, filename)
                    with open(filepath, "w", encoding="utf-8") as f:
                        json.dump(imported, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Paramètres de l'application (dossier de sauvegarde cloud, etc.)
# ---------------------------------------------------------------------------

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            return {}
    return {}


def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def get_daily_recipe(recipes):
    """Retourne la recette du jour : la même toute la journée (mémorisée
    dans settings.json), renouvelée aléatoirement chaque nouveau jour, parmi
    toutes les recettes (pas seulement la liste d'envies)."""
    if not recipes:
        return None
    settings = load_settings()
    today = datetime.now().strftime("%Y-%m-%d")
    if settings.get("daily_recipe_date") == today:
        match = next((r for r in recipes if r["name"] == settings.get("daily_recipe_name")), None)
        if match is not None:
            return match
    chosen = random.choice(recipes)
    settings["daily_recipe_date"] = today
    settings["daily_recipe_name"] = chosen["name"]
    save_settings(settings)
    return chosen


def get_cloud_backup_folder():
    folder = load_settings().get("cloud_backup_folder") or ""
    return folder if folder and os.path.isdir(folder) else ""


def set_cloud_backup_folder(path):
    settings = load_settings()
    if path:
        settings["cloud_backup_folder"] = path
    else:
        settings.pop("cloud_backup_folder", None)
    save_settings(settings)


AUTO_BACKUP_RETENTION = 10  # nombre de sauvegardes automatiques conservées
AUTO_BACKUP_MIN_INTERVAL_HOURS = 24  # fréquence minimale entre deux sauvegardes auto
# Préfixe distinctif des fichiers de sauvegarde, pour ne jamais les confondre
# avec ceux d'un autre programme dans un même dossier (ex. cloud partagé).
AUTO_BACKUP_PREFIX = "sauvegarde_auto_mesrecettes_"


def list_auto_backups():
    """Retourne la liste des sauvegardes automatiques existantes
    (chemin complet), les plus récentes en premier."""
    if not os.path.isdir(BACKUPS_DIR):
        return []
    files = [
        os.path.join(BACKUPS_DIR, f) for f in os.listdir(BACKUPS_DIR)
        if f.startswith(AUTO_BACKUP_PREFIX) and f.endswith(".zip")
    ]
    files.sort(key=os.path.getmtime, reverse=True)
    return files


def create_auto_backup():
    """Crée une nouvelle sauvegarde automatique horodatée, puis supprime les
    plus anciennes au-delà de AUTO_BACKUP_RETENTION. Si un dossier cloud est
    configuré (Google Drive, OneDrive, Dropbox...), une copie y est aussi
    déposée : le client cloud déjà installé sur le PC se charge ensuite de
    l'envoyer en ligne automatiquement."""
    os.makedirs(BACKUPS_DIR, exist_ok=True)
    filename = f"{AUTO_BACKUP_PREFIX}{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.zip"
    path = os.path.join(BACKUPS_DIR, filename)
    build_full_backup_zip(path)

    existing = list_auto_backups()
    for old_path in existing[AUTO_BACKUP_RETENTION:]:
        try:
            os.remove(old_path)
        except OSError:
            pass

    cloud_folder = get_cloud_backup_folder()
    if cloud_folder:
        try:
            cloud_path = os.path.join(cloud_folder, filename)
            shutil.copy2(path, cloud_path)
            cloud_backups = sorted(
                [os.path.join(cloud_folder, f) for f in os.listdir(cloud_folder)
                 if f.startswith(AUTO_BACKUP_PREFIX) and f.endswith(".zip")],
                key=os.path.getmtime, reverse=True
            )
            for old_cloud_path in cloud_backups[AUTO_BACKUP_RETENTION:]:
                try:
                    os.remove(old_cloud_path)
                except OSError:
                    pass
        except Exception:
            pass  # un souci côté dossier cloud ne doit jamais faire échouer la sauvegarde locale

    return path


def migrate_old_backup_filenames():
    """Renomme les sauvegardes créées avant l'ajout du préfixe distinctif
    'mesrecettes' (ex. 'sauvegarde_auto_2026-01-01_120000.zip') vers le
    nouveau format, pour qu'elles restent reconnues par l'application au
    lieu de devenir invisibles."""
    if not os.path.isdir(BACKUPS_DIR):
        return
    old_pattern = re.compile(r"^sauvegarde_auto_(\d{4}-\d{2}-\d{2}_\d{6})\.zip$")
    for fname in os.listdir(BACKUPS_DIR):
        match = old_pattern.match(fname)
        if match:
            new_path = os.path.join(BACKUPS_DIR, f"{AUTO_BACKUP_PREFIX}{match.group(1)}.zip")
            if not os.path.exists(new_path):
                try:
                    os.rename(os.path.join(BACKUPS_DIR, fname), new_path)
                except OSError:
                    pass


def maybe_create_auto_backup():
    """Crée automatiquement une sauvegarde si aucune n'existe encore ou si la
    dernière date de plus de AUTO_BACKUP_MIN_INTERVAL_HOURS. Ne fait jamais
    planter l'application en cas de problème (disque plein, permissions...)."""
    try:
        migrate_old_backup_filenames()
        if not os.path.exists(DATA_FILE):
            return  # rien à sauvegarder pour un tout premier lancement
        existing = list_auto_backups()
        if existing:
            last_mtime = os.path.getmtime(existing[0])
            age_hours = (datetime.now().timestamp() - last_mtime) / 3600
            if age_hours < AUTO_BACKUP_MIN_INTERVAL_HOURS:
                return
        create_auto_backup()
    except Exception:
        pass


class ImportExportWindow(tk.Toplevel):
    """Fenêtre pour exporter ou importer l'ensemble des données (recettes,
    ingrédients et photos) sous forme d'une seule archive ZIP, et pour
    consulter/restaurer les sauvegardes automatiques."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("importexport_title"))
        self.geometry(f"{gs(480)}x{gs(940)}")
        self.minsize(gs(440), gs(500))
        self.resizable(True, True)
        self.grab_set()

        ttk.Label(self, text=t("importexport_heading"),
                  font=("Segoe UI", sf(13), "bold")).pack(pady=15)

        ttk.Label(
            self,
            text=t("importexport_export_intro"),
            justify="center", font=("Segoe UI", sf(9))
        ).pack(pady=(0, 10))

        ttk.Button(self, text=t("importexport_export_button"),
                   width=42, command=self.export_data).pack(pady=6)

        ttk.Label(
            self,
            text=t("importexport_import_intro"),
            justify="center", font=("Segoe UI", sf(9))
        ).pack(pady=(10, 10))

        ttk.Button(self, text=t("importexport_import_button"),
                   width=42, command=self.import_data).pack(pady=6)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=20, pady=15)

        ttk.Label(self, text=t("importexport_auto_backups_heading"), font=("Segoe UI", sf(12), "bold")).pack()
        ttk.Label(
            self,
            text=t("importexport_auto_backups_intro", hours=AUTO_BACKUP_MIN_INTERVAL_HOURS, retention=AUTO_BACKUP_RETENTION),
            justify="center", font=("Segoe UI", sf(9))
        ).pack(pady=(5, 10))

        list_frame = ttk.Frame(self)
        list_frame.pack(padx=15, fill="both", expand=True)
        self.backup_listbox = tk.Listbox(list_frame, height=8, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.backup_listbox.yview)
        self.backup_listbox.configure(yscrollcommand=scrollbar.set)
        self.backup_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._populate_backups()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("importexport_backup_now_button"),
                   command=self.backup_now).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("importexport_restore_selected_button"),
                   command=self.restore_selected).grid(row=0, column=1, padx=5)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=20, pady=15)

        ttk.Label(self, text=t("importexport_cloud_heading"),
                  font=("Segoe UI", sf(12), "bold")).pack()
        ttk.Label(
            self,
            text=t("importexport_cloud_intro"),
            justify="center", font=("Segoe UI", sf(9))
        ).pack(pady=(5, 8))

        self.cloud_folder_label = ttk.Label(self, text="", font=("Segoe UI", sf(9), "bold"),
                                             foreground="#266", wraplength=400, justify="center")
        self.cloud_folder_label.pack(pady=(0, 8))
        self._refresh_cloud_label()

        cloud_btn_frame = ttk.Frame(self)
        cloud_btn_frame.pack(pady=(0, 15))
        ttk.Button(cloud_btn_frame, text=t("importexport_choose_cloud_button"),
                   command=self.choose_cloud_folder).grid(row=0, column=0, padx=5)
        ttk.Button(cloud_btn_frame, text=t("importexport_disable_button"),
                   command=self.disable_cloud_backup).grid(row=0, column=1, padx=5)

        tk.Frame(self, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _refresh_cloud_label(self):
        folder = get_cloud_backup_folder()
        if folder:
            self.cloud_folder_label.config(text=t("importexport_cloud_enabled", folder=folder))
        else:
            self.cloud_folder_label.config(text=t("importexport_cloud_not_configured"), foreground=COLOR_TEXT_MUTED)

    def choose_cloud_folder(self):
        folder = filedialog.askdirectory(
            title=t("importexport_choose_folder_title")
        )
        if not folder:
            return
        set_cloud_backup_folder(folder)
        self._refresh_cloud_label()
        if messagebox.askyesno(
            t("importexport_cloud_configured_title"),
            t("importexport_cloud_configured_message", folder=folder)
        ):
            self.backup_now()

    def disable_cloud_backup(self):
        if not get_cloud_backup_folder():
            return
        set_cloud_backup_folder(None)
        self._refresh_cloud_label()
        messagebox.showinfo(t("importexport_disabled_title"), t("importexport_disabled_message"))

    def _populate_backups(self):
        self.backup_listbox.delete(0, tk.END)
        self.backups = list_auto_backups()
        for path in self.backups:
            mtime = datetime.fromtimestamp(os.path.getmtime(path))
            size_kb = os.path.getsize(path) / 1024
            self.backup_listbox.insert(
                tk.END, t("importexport_backup_date_line", date=mtime.strftime('%d/%m/%Y à %H:%M'), size=f"{size_kb:.0f}")
            )
        if not self.backups:
            self.backup_listbox.insert(tk.END, t("importexport_no_backups"))

    def backup_now(self):
        try:
            create_auto_backup()
        except Exception as e:
            messagebox.showerror(t("common_error"), t("importexport_backup_failed", error=e))
            return
        self._populate_backups()
        messagebox.showinfo(t("importexport_backup_created_title"), t("importexport_backup_created_message"))

    def restore_selected(self):
        sel = self.backup_listbox.curselection()
        if not sel or not self.backups:
            messagebox.showinfo(t("common_info"), t("importexport_select_backup_first"))
            return
        path = self.backups[sel[0]]

        mode = messagebox.askyesnocancel(
            t("importexport_restore_mode_title"),
            t("importexport_restore_mode_message")
        )
        if mode is None:
            return
        try:
            restore_from_zip(path, merge=bool(mode))
        except Exception as e:
            messagebox.showerror(t("common_error"), t("importexport_restore_failed", error=e))
            return
        self.app.refresh_recipes()
        self.app.refresh_ingredients()
        messagebox.showinfo(t("importexport_restore_done_title"), t("importexport_restore_done_message"))
        self.destroy()

    def export_data(self):
        path = filedialog.asksaveasfilename(
            title=t("importexport_export_data_title"),
            defaultextension=".zip",
            filetypes=[("Archive ZIP", "*.zip")],
            initialfile="mes_recettes_export.zip"
        )
        if not path:
            return
        try:
            build_full_backup_zip(path)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("importexport_export_data_success", path=path))

    def import_data(self):
        path = filedialog.askopenfilename(
            title=t("importexport_choose_archive_title"),
            filetypes=[("Archive ZIP", "*.zip")]
        )
        if not path:
            return

        mode = messagebox.askyesnocancel(
            t("importexport_import_mode_title"),
            t("importexport_import_mode_message")
        )
        if mode is None:
            return

        try:
            restore_from_zip(path, merge=bool(mode))
        except Exception as e:
            messagebox.showerror(t("common_error"), t("importexport_import_failed", error=e))
            return

        self.app.refresh_recipes()
        self.app.refresh_ingredients()
        messagebox.showinfo(t("importexport_import_done_title"), t("importexport_import_done_message"))
        self.destroy()


class ShoppingChecklistWindow(tk.Toplevel):
    """Affiche une liste de courses déjà calculée sous forme de cases à
    cocher, pour pointer les articles au fur et à mesure des courses."""

    def __init__(self, app, grouped_totals, title=None):
        super().__init__(app)
        self.app = app
        if title is None:
            title = t("allrecipes_shopping_list_title")
        self.title(f"☑️ {title}")
        self.geometry(f"{gs(480)}x{gs(600)}")
        self.grab_set()

        ttk.Label(self, text=f"☑️ {title}", font=("Segoe UI", sf(14), "bold")).pack(pady=(15, 5))
        ttk.Label(self, text=t("checklist_instruction"),
                  font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED).pack(pady=(0, 10))

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=15)
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        rows_frame = ttk.Frame(canvas)
        rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.checks = []
        for rayon, items in grouped_totals:
            ttk.Label(rows_frame, text=translate_rayon_name(rayon), font=("Segoe UI", sf(11), "bold")).pack(
                anchor="w", pady=(12, 3))
            for name, qty, unit in items:
                unit_display = f" {translate_unit_name(unit)}" if unit else ""
                var = tk.BooleanVar()
                lbl_text = f"{translate_ingredient_name(name)} : {qty}{unit_display}"
                chk = ttk.Checkbutton(rows_frame, text=lbl_text, variable=var,
                                       command=lambda: None)
                chk.pack(anchor="w", padx=10, pady=1)
                self.checks.append((var, chk, lbl_text))
                var.trace_add("write", lambda *args, v=var, c=chk: self._update_style(v, c))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("checklist_check_all_button"), command=self.check_all).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("checklist_uncheck_all_button"), command=self.uncheck_all).grid(row=0, column=1, padx=5)

        self.progress_label = ttk.Label(self, text="", font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED)
        self.progress_label.pack(pady=(0, 10))
        self._update_progress()

    def _update_style(self, var, chk):
        style_name = "Checked.TCheckbutton" if var.get() else "TCheckbutton"
        try:
            style = ttk.Style(self)
            style.configure("Checked.TCheckbutton", foreground="#999")
            chk.configure(style=style_name)
        except tk.TclError:
            pass
        self._update_progress()

    def _update_progress(self):
        total = len(self.checks)
        done = sum(1 for var, chk, text in self.checks if var.get())
        self.progress_label.config(text=t("checklist_progress_label", done=done, total=total))

    def check_all(self):
        for var, chk, text in self.checks:
            var.set(True)

    def uncheck_all(self):
        for var, chk, text in self.checks:
            var.set(False)


class ExportFormatDialog(tk.Toplevel):
    """Petite fenêtre pour choisir le format d'export (texte, Excel ou PDF)
    d'une liste de courses, réutilisable depuis n'importe quelle fenêtre qui
    propose ces 3 exports (« Toutes les recettes », « Planning de la
    semaine », « Nouveau menu »)."""

    def __init__(self, parent_window, export_txt_callback, export_excel_callback, export_pdf_callback):
        super().__init__(parent_window)
        self.title(t("exportformat_title"))
        self.geometry(f"{gs(380)}x{gs(280)}")
        self.resizable(False, False)
        self.grab_set()

        ttk.Label(self, text=t("exportformat_heading"),
                  font=("Segoe UI", sf(12), "bold"), wraplength=340, justify="center").pack(pady=(20, 5))
        ttk.Label(self, text=t("exportformat_choose_label"),
                  font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED).pack(pady=(0, 15))

        ttk.Button(self, text=t("exportformat_txt_button"),
                   command=lambda: self._run(export_txt_callback)).pack(pady=6, padx=40, fill="x")
        ttk.Button(self, text=t("exportformat_excel_button"),
                   command=lambda: self._run(export_excel_callback)).pack(pady=6, padx=40, fill="x")
        ttk.Button(self, text=t("exportformat_pdf_button"),
                   command=lambda: self._run(export_pdf_callback)).pack(pady=6, padx=40, fill="x")
        ttk.Button(self, text=t("exportformat_cancel_button"), style="Secondary.TButton",
                   command=self.destroy).pack(pady=(15, 10))

    def _run(self, callback):
        self.destroy()
        callback()


class AddManualIngredientDialog(tk.Toplevel):
    """Permet d'ajouter un ou plusieurs ingrédients (avec quantité) directement
    à une liste de courses, indépendamment des recettes sélectionnées — par
    exemple pour du papier essuie-tout ou tout autre article à ne pas
    oublier. Réutilisable depuis "Toutes les recettes", "Planning de la
    semaine" et "Nouveau menu" : `target_window` doit juste exposer une
    méthode `add_manual_items(items)`."""

    def __init__(self, app, target_window):
        super().__init__(target_window)
        self.app = app
        self.target_window = target_window
        self.staged_items = []  # ingrédients ajoutés à la liste d'attente, pas encore validés
        self.title(t("addmanual_title"))
        self.geometry(f"{gs(600)}x{gs(560)}")
        self.minsize(gs(420), gs(480))
        self.resizable(True, True)
        self.grab_set()

        ttk.Label(self, text=t("addmanual_heading"),
                  font=("Segoe UI", sf(12), "bold"), wraplength=420, justify="center").pack(pady=(15, 5))
        ttk.Label(
            self, text=t("addmanual_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

        name_frame = ttk.Frame(self)
        name_frame.pack(pady=5)
        ttk.Label(name_frame, text=t("common_ingredient_label")).grid(row=0, column=0, padx=5, sticky="e")
        self.name_combo = ttk.Combobox(name_frame, values=get_display_ingredient_values(sorted(self.app.ingredient_names, key=ingredient_sort_key)),
                                        width=26)
        self.name_combo.grid(row=0, column=1, padx=5)
        ttk.Button(name_frame, text=t("addmanual_new_ingredient_button"),
                   command=self.create_new_ingredient).grid(row=0, column=2, padx=5)

        qty_frame = ttk.Frame(self)
        qty_frame.pack(pady=10)
        ttk.Label(qty_frame, text=t("common_quantity_label")).grid(row=0, column=0, padx=5)
        self.qty_entry = ttk.Entry(qty_frame, width=8)
        self.qty_entry.insert(0, "1")
        self.qty_entry.grid(row=0, column=1, padx=5)
        ttk.Label(qty_frame, text=t("common_unit_label")).grid(row=0, column=2, padx=5)
        self.unit_options = RecipeFormWindow.UNIT_OPTIONS[:-1] + ["boîte", "paquet", "rouleau", "bouteille"]
        self.unit_combo = ttk.Combobox(qty_frame, values=[translate_unit_name(u) for u in self.unit_options], width=14)  # texte libre autorisé
        self.unit_combo.set(translate_unit_name("pièce"))
        self.unit_combo.grid(row=0, column=3, padx=5)
        ttk.Button(qty_frame, text=t("addmanual_add_to_list_button"),
                   command=self.stage_item).grid(row=0, column=4, padx=(10, 0))
        self.name_combo.bind("<Return>", lambda e: self.stage_item())
        self.qty_entry.bind("<Return>", lambda e: self.stage_item())

        ttk.Label(self, text=t("addmanual_staged_label"),
                  font=("Segoe UI", sf(10), "bold")).pack(pady=(10, 3))
        staged_frame = ttk.Frame(self)
        staged_frame.pack(padx=15, fill="both", expand=True)
        self.staged_listbox = tk.Listbox(staged_frame, height=10, font=("Segoe UI", sf(9)))
        staged_scrollbar = ttk.Scrollbar(staged_frame, orient="vertical", command=self.staged_listbox.yview)
        self.staged_listbox.configure(yscrollcommand=staged_scrollbar.set)
        self.staged_listbox.pack(side="left", fill="both", expand=True)
        staged_scrollbar.pack(side="right", fill="y")
        ttk.Button(self, text=t("addmanual_remove_staged_button"),
                   command=self.remove_staged).pack(pady=(5, 0))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text=t("addmanual_confirm_all_button"),
                   command=self.confirm_all).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("addmanual_close_button"), style="Secondary.TButton",
                   command=self.close).grid(row=0, column=1, padx=5)
        self.protocol("WM_DELETE_WINDOW", self.close)

    def close(self):
        self.destroy()
        # Sans ceci, la fenêtre d'origine (ex. "Toutes les recettes") peut
        # se retrouver derrière la page d'accueil une fois cette fenêtre
        # fermée, plutôt que de rester au premier plan.
        self.target_window.lift()
        self.target_window.focus_force()

    def create_new_ingredient(self):
        typed = normalize_oe(self.name_combo.get().strip())
        win = IngredientEditWindow(self.app, manage_window=None, existing_name=None,
                                    prefill_name=typed, parent_window=self)
        self.wait_window(win)
        self.name_combo["values"] = get_display_ingredient_values(sorted(self.app.ingredient_names, key=ingredient_sort_key))

    def stage_item(self):
        name = normalize_oe(self.name_combo.get().strip())
        if not name:
            messagebox.showerror(t("common_error"), t("pantry_error_ingredient_required"))
            return
        # Cette liste accepte aussi des articles hors base de données (ex.
        # essuie-tout) : si le nom tapé correspond à un ingrédient connu
        # (en français ou en anglais), on le normalise vers son nom
        # canonique français ; sinon, on garde le texte tel quel.
        canonical = resolve_ingredient_input(name, self.app.ingredient_names)
        if canonical is not None:
            name = canonical
        try:
            quantity = float(self.qty_entry.get().strip().replace(",", "."))
            if quantity <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(t("common_error"), t("allrecipes_invalid_quantity"))
            return
        unit = resolve_unit_input_best_effort(self.unit_combo.get().strip(), self.unit_options)
        self.staged_items.append({"name": name, "quantity": quantity, "unit": unit})
        unit_display = f" {translate_unit_name(unit)}" if unit else ""
        self.staged_listbox.insert(tk.END, f"{translate_ingredient_name(name)} : {quantity}{unit_display}")

        # Prêt pour la saisie suivante : on vide juste le nom et la
        # quantité repasse à 1, pour enchaîner rapidement plusieurs ajouts.
        self.name_combo.set("")
        self.qty_entry.delete(0, tk.END)
        self.qty_entry.insert(0, "1")
        self.name_combo.focus_set()

    def remove_staged(self):
        sel = self.staged_listbox.curselection()
        if not sel:
            messagebox.showinfo(t("common_info"), t("addmanual_select_staged_first"))
            return
        for i in reversed(sel):
            self.staged_listbox.delete(i)
            del self.staged_items[i]

    def confirm_all(self):
        if not self.staged_items:
            messagebox.showinfo(t("common_info"), t("addmanual_add_staged_first"))
            return
        self.target_window.add_manual_items(self.staged_items)
        messagebox.showinfo(t("onerecipe_added_to_shopping_title"), t("addmanual_confirmed_message", count=len(self.staged_items)))
        self.close()


class SavedShoppingListsWindow(tk.Toplevel):
    """Fenêtre pour recharger ou supprimer une liste de courses enregistrée
    précédemment via « 💾 Enregistrer cette liste pour plus tard ». Réutilisable
    depuis n'importe quelle fenêtre de liste de courses éditable :
    `target_window` doit exposer une méthode `load_saved_list(items)`."""

    def __init__(self, app, target_window):
        super().__init__(target_window)
        self.app = app
        self.target_window = target_window
        self.title(t("savedlists_title"))
        self.geometry(f"{gs(500)}x{gs(440)}")
        self.minsize(gs(420), gs(360))
        self.resizable(True, True)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.close)

        ttk.Label(self, text=t("savedlists_heading"),
                  font=("Segoe UI", sf(12), "bold")).pack(pady=(15, 10))

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=15)
        self.listbox = tk.Listbox(list_frame, height=12, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<Double-Button-1>", lambda e: self.load_selected())

        self.saved_lists = []
        self._populate()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text=t("savedlists_load_button"), command=self.load_selected).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("savedlists_delete_button"), command=self.delete_selected).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text=t("addmanual_close_button"), style="Secondary.TButton",
                   command=self.close).grid(row=0, column=2, padx=5)

    def _populate(self):
        self.listbox.delete(0, tk.END)
        self.saved_lists = load_saved_shopping_lists()
        self.saved_lists.sort(key=lambda l: l.get("created_at", ""), reverse=True)
        if not self.saved_lists:
            self.listbox.insert(tk.END, t("savedlists_none_saved"))
            return
        for saved in self.saved_lists:
            n_items = len(saved.get("items", []))
            self.listbox.insert(
                tk.END, t("savedlists_entry_line", name=saved['name'], count=n_items, date=saved.get('created_at', '?'))
            )

    def load_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.saved_lists:
            messagebox.showinfo(t("common_info"), t("savedlists_select_list_first"))
            return
        saved = self.saved_lists[sel[0]]
        self.target_window.load_saved_list(saved["items"])
        self.close()

    def delete_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.saved_lists:
            messagebox.showinfo(t("common_info"), t("savedlists_select_list_first"))
            return
        saved = self.saved_lists[sel[0]]
        if not messagebox.askyesno(t("common_confirm"), t("savedlists_delete_confirm", name=saved['name'])):
            return
        all_lists = load_saved_shopping_lists()
        all_lists = [l for l in all_lists if l["name"] != saved["name"]]
        save_saved_shopping_lists(all_lists)
        self._populate()

    def close(self):
        self.destroy()
        self.target_window.lift()
        self.target_window.focus_force()


class AllRecipesWindow(tk.Toplevel):
    """Fenêtre listant toutes les recettes avec sélection + nombre de personnes,
    pour calculer et exporter en PDF la quantité totale d'ingrédients nécessaire."""

    SORT_OPTIONS = RECIPE_SORT_OPTIONS

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("allrecipes_title"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(1220)}x{screen_height}+40+0")
        self.minsize(gs(720), gs(500))
        self.resizable(True, True)
        self.grab_set()
        self.manual_items = []  # ingrédients ajoutés manuellement (hors recettes) : [{"name","quantity","unit"}]
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        ttk.Label(self, text=t("allrecipes_select_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(10, 5))

        top_frame = ttk.Frame(self)
        top_frame.pack(pady=(0, 5), fill="x", padx=15)
        ttk.Label(top_frame, text=t("common_search_label")).pack(side="left")
        self.search_entry = ttk.Entry(top_frame, width=20)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", lambda e: self._filter_rows())
        ttk.Label(top_frame, text=t("common_sort_by_label")).pack(side="left", padx=(10, 2))
        self.sort_combo = ttk.Combobox(top_frame, values=[translate_sort_option(o) for o in self.SORT_OPTIONS], state="readonly", width=18)
        self.sort_combo.set(translate_sort_option(self.SORT_OPTIONS[0]))
        self.sort_combo.pack(side="left")
        self.sort_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_sort())
        ttk.Label(top_frame, text=t("common_category_label")).pack(side="left", padx=(10, 2))
        self.category_filter_combo = ttk.Combobox(
            top_frame, values=[t("common_all_categories")] + [translate_category_name(c) for c in RecipeFormWindow.CATEGORY_OPTIONS],
            state="readonly", width=16
        )
        self.category_filter_combo.set(t("common_all_categories"))
        self.category_filter_combo.pack(side="left")
        self.category_filter_combo.bind("<<ComboboxSelected>>", lambda e: self._filter_rows())

        ingredient_filter_frame = ttk.LabelFrame(self, text=t("allrecipes_ingredient_filter_title"))
        ingredient_filter_frame.pack(pady=(0, 8), padx=15, fill="x")
        ingredient_values = get_display_ingredient_values(sorted(self.app.ingredient_names, key=ingredient_sort_key))

        ttk.Label(ingredient_filter_frame, text=t("common_want_label")).grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.want_entries = []
        for i in range(2):
            entry = self._make_ingredient_filter_entry(ingredient_filter_frame, ingredient_values)
            entry.grid(row=0, column=1 + i, padx=5, pady=3)
            self.want_entries.append(entry)

        ttk.Label(ingredient_filter_frame, text=t("common_exclude_label")).grid(row=1, column=0, sticky="w", padx=5, pady=3)
        self.exclude_entries = []
        for i in range(2):
            entry = self._make_ingredient_filter_entry(ingredient_filter_frame, ingredient_values)
            entry.grid(row=1, column=1 + i, padx=5, pady=3)
            self.exclude_entries.append(entry)

        ttk.Label(ingredient_filter_frame, text=t("common_tags_filter_label")).grid(
            row=2, column=0, sticky="w", padx=5, pady=3)
        self.all_tags = sorted({tag for r in self.app.recipes for tag in r.get("tags", [])}, key=ingredient_sort_key)
        self.tag_filter_entries = []
        for i in range(2):
            entry = self._make_ingredient_filter_entry(ingredient_filter_frame, self.all_tags)
            entry.grid(row=2, column=1 + i, padx=5, pady=3)
            self.tag_filter_entries.append(entry)

        ttk.Button(ingredient_filter_frame, text=t("common_reset_button"),
                   command=self._reset_ingredient_filters).grid(row=0, column=3, rowspan=3, padx=8)
        ttk.Label(ingredient_filter_frame, text=t("common_filter_hint"),
                  font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED).grid(
            row=3, column=0, columnspan=4, sticky="w", padx=5, pady=(0, 3))

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10)

        canvas = tk.Canvas(container, height=240, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        rows_frame = ttk.Frame(canvas)
        rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.checks = []
        self.current_items = []       # liste plate éditable [{'name','quantity','unit','rayon'}, ...]
        self.last_chosen_recipes = []  # recettes ajoutées au panier (pour les en-têtes d'export)
        for row_index, recipe in enumerate(self.app.recipes):
            row = ttk.Frame(rows_frame)
            row.grid(row=row_index, column=0, sticky="ew", pady=4)
            ttk.Label(row, text=format_recipe_list_label(recipe), width=110, anchor="w").grid(
                row=0, column=0, sticky="w")
            ttk.Label(row, text=t("allrecipes_persons_count_label")).grid(row=0, column=1)
            pers_entry = ttk.Entry(row, width=5)

            # Préremplit à partir d'une sélection faite depuis "Voir une
            # recette précise" (bouton "Ajouter à la liste de courses"),
            # sinon utilise le nombre de personnes par défaut de la recette.
            preselected = self.app.shopping_selection.get(recipe["name"])
            if preselected is not None:
                pers_entry.insert(0, str(preselected))
            else:
                pers_entry.insert(0, str(recipe.get("default_persons") or 1))
            pers_entry.grid(row=0, column=2, padx=5)
            ttk.Button(row, text=t("allrecipes_add_to_cart_button"), width=18,
                       command=lambda r=recipe, e=pers_entry: self._add_recipe_to_cart(r, e)).grid(
                row=0, column=3, padx=5)
            ttk.Button(row, text=t("common_edit_button"), width=10,
                       command=lambda idx=row_index: self._edit_recipe(idx)).grid(row=0, column=4, padx=5)
            self.checks.append((recipe, pers_entry, row))

        tk.Frame(rows_frame, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).grid(
            row=len(self.app.recipes), column=0, sticky="ew")

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        for col in range(4):
            btn_frame.columnconfigure(col, weight=1)
        ttk.Button(btn_frame, text=t("allrecipes_checklist_mode_button"),
                   command=self.open_checklist).grid(row=0, column=0, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_clear_list_button"),
                   command=self.clear_selection).grid(row=0, column=2, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_export_button"),
                   command=self.open_export_dialog).grid(row=1, column=0, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_print_button"),
                   command=self.print_shopping_list).grid(row=1, column=2, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_add_manual_ingredient_button"),
                   command=self.open_add_manual_ingredient).grid(
            row=2, column=0, columnspan=4, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_save_list_button"),
                   command=self.save_list_for_later).grid(row=3, column=0, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_load_list_button"),
                   command=self.open_saved_lists).grid(row=3, column=2, columnspan=2, padx=5, pady=3, sticky="ew")

        # ---- Zone de résultat éditable : chaque ingrédient peut voir sa
        # quantité modifiée ou être retiré, sans devoir tout recalculer. ----
        result_container = ttk.Frame(self)
        result_container.pack(pady=10, padx=15, fill="both", expand=True)
        result_canvas = tk.Canvas(result_container, highlightthickness=0)
        result_scrollbar = ttk.Scrollbar(result_container, orient="vertical", command=result_canvas.yview)
        self.result_frame = ttk.Frame(result_canvas)
        self.result_frame.bind(
            "<Configure>", lambda e: result_canvas.configure(scrollregion=result_canvas.bbox("all"))
        )
        result_canvas.create_window((0, 0), window=self.result_frame, anchor="nw")
        result_canvas.configure(yscrollcommand=result_scrollbar.set)
        result_canvas.pack(side="left", fill="both", expand=True)
        result_scrollbar.pack(side="right", fill="y")

        def _on_result_mousewheel(event):
            result_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        result_canvas.bind("<Enter>", lambda e: result_canvas.bind_all("<MouseWheel>", _on_result_mousewheel))
        result_canvas.bind("<Leave>", lambda e: result_canvas.unbind_all("<MouseWheel>"))

        # Ajoute automatiquement au panier les recettes présélectionnées
        # depuis "Voir une recette précise" (bouton "🛒 Ajouter à la liste de
        # courses"), pour ne pas perdre cette présélection maintenant qu'il
        # n'y a plus de case à cocher à valider soi-même.
        for recipe, pers_entry, row in self.checks:
            if recipe["name"] in self.app.shopping_selection:
                self._add_recipe_to_cart(recipe, pers_entry)

        self._render_shopping_list()

    def open_add_manual_ingredient(self):
        AddManualIngredientDialog(self.app, self)

    def open_export_dialog(self):
        ExportFormatDialog(self, self.export_txt, self.export_excel, self.export_pdf)

    def add_manual_items(self, items):
        self.manual_items.extend(items)
        for item in items:
            self._merge_item_into_current(
                item["name"], item["quantity"], item["unit"], get_ingredient_rayon(item["name"])
            )
        self._render_shopping_list()

    def _edit_recipe(self, index):
        win = RecipeFormWindow(self.app, recipe_index=index)
        self.wait_window(win)
        self.app.refresh_recipes()
        # La modification peut avoir changé le nom, les temps, les
        # allergènes... : on reconstruit la fenêtre pour que tout
        # s'affiche à jour (libellés, tri, filtres).
        self.destroy()
        AllRecipesWindow(self.app)

    def _apply_sort(self):
        option = resolve_sort_option_input(self.sort_combo.get(), self.SORT_OPTIONS)
        reverse = option in ("Ajoutées récemment",)
        ordered = sorted(self.checks, key=lambda t: recipe_sort_key(t[0], option), reverse=reverse)
        for new_index, (recipe, pers_entry, row) in enumerate(ordered):
            row.grid(row=new_index, column=0, sticky="ew", pady=4)
        self.checks = ordered
        self._filter_rows()

    def _on_close(self):
        if getattr(self, "current_items", None):
            if not messagebox.askyesno(
                t("allrecipes_close_confirm_title"),
                t("allrecipes_close_confirm_message"),
                icon="warning"
            ):
                return  # l'utilisateur annule la fermeture
        self.destroy()

    def clear_selection(self):
        self.app.shopping_selection.clear()
        self.manual_items = []
        self.current_items = []
        self.last_chosen_recipes = []
        self._render_shopping_list()

    def _filter_rows(self):
        search = self.search_entry.get().strip()
        search_key = ingredient_sort_key(search) if search else ""
        category_filter = resolve_category_input(self.category_filter_combo.get(), RecipeFormWindow.CATEGORY_OPTIONS)

        known_keys = {ingredient_sort_key(n) for n in self.app.ingredient_names}
        known_tag_keys = {ingredient_sort_key(t) for t in self.all_tags}

        def valid_typed_names(entries, known, resolve_as_ingredient=False):
            names = []
            for e in entries:
                txt = e.get().strip()
                if not txt:
                    continue
                if resolve_as_ingredient:
                    resolved = resolve_ingredient_input(txt, self.app.ingredient_names)
                    if resolved is not None:
                        names.append(resolved)
                elif ingredient_sort_key(txt) in known:
                    names.append(txt)
            return names

        want_names = valid_typed_names(self.want_entries, known_keys, resolve_as_ingredient=True)
        exclude_names = valid_typed_names(self.exclude_entries, known_keys, resolve_as_ingredient=True)
        want_keys = {ingredient_sort_key(n) for n in want_names}
        exclude_keys = {ingredient_sort_key(n) for n in exclude_names}
        tag_names = valid_typed_names(self.tag_filter_entries, known_tag_keys)
        tag_keys = {ingredient_sort_key(n) for n in tag_names}

        for recipe, pers_entry, row in self.checks:
            if not recipe_matches_search(recipe, search_key):
                row.grid_remove()
                continue
            if category_filter and category_filter != t("common_all_categories") and recipe.get("category", "Autre") != category_filter:
                row.grid_remove()
                continue
            if tag_keys:
                recipe_tag_keys = {ingredient_sort_key(t) for t in recipe.get("tags", [])}
                if not tag_keys.issubset(recipe_tag_keys):
                    row.grid_remove()
                    continue

            recipe_ing_keys = {ingredient_sort_key(ing["name"]) for ing in recipe["ingredients"]}
            if want_keys and not want_keys.issubset(recipe_ing_keys):
                row.grid_remove()
                continue
            if exclude_keys and (exclude_keys & recipe_ing_keys):
                row.grid_remove()
                continue

            row.grid()

    def _reset_ingredient_filters(self):
        for entry in self.want_entries + self.exclude_entries + self.tag_filter_entries:
            entry.delete(0, tk.END)
        self._filter_rows()

    # ---- Autocomplétion des champs de filtre par ingrédient (même principe
    # que le choix d'ingrédient dans le formulaire de recette) ----

    def _make_ingredient_filter_entry(self, parent, values):
        entry = ttk.Entry(parent, width=22)
        entry.full_values = values
        entry.bind("<KeyRelease>", lambda e: self._on_filter_entry_keyrelease(e, entry))
        entry.bind("<FocusIn>", lambda e: self._on_filter_entry_focus_in(e, entry))
        entry.bind("<FocusOut>", lambda e: self._on_filter_entry_focus_out(e, entry))
        return entry

    def _hide_filter_suggestions(self, entry):
        popup = getattr(entry, "_suggestion_popup", None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
            entry._suggestion_popup = None
            entry._suggestion_listbox = None

    def _show_filter_suggestions(self, entry, filtered):
        self._hide_filter_suggestions(entry)
        if not filtered:
            return
        popup = tk.Toplevel(entry)
        popup.wm_overrideredirect(True)
        try:
            popup.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        width = max(entry.winfo_width(), 160)
        height = min(6, len(filtered)) * 20
        popup.wm_geometry(f"{width}x{height}+{x}+{y}")

        listbox = tk.Listbox(popup, height=min(6, len(filtered)), exportselection=False, font=("Segoe UI", sf(9)))
        listbox.pack(fill="both", expand=True)
        for v in filtered:
            listbox.insert(tk.END, v)

        def choose(event=None):
            sel = listbox.curselection()
            if sel:
                value = listbox.get(sel[0])
                entry.delete(0, tk.END)
                entry.insert(0, value)
            self._hide_filter_suggestions(entry)
            entry.focus_set()
            self._filter_rows()

        listbox.bind("<ButtonRelease-1>", choose)
        listbox.bind("<Return>", choose)
        entry._suggestion_popup = popup
        entry._suggestion_listbox = listbox

    def _on_filter_entry_keyrelease(self, event, entry):
        if event.keysym == "Down":
            listbox = getattr(entry, "_suggestion_listbox", None)
            if listbox is not None:
                listbox.focus_set()
                listbox.selection_set(0)
            return
        if event.keysym in ("Escape",):
            self._hide_filter_suggestions(entry)
            return
        if event.keysym in ("Return", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R",
                              "Caps_Lock", "Alt_L", "Alt_R", "Left", "Right"):
            return
        full_values = getattr(entry, "full_values", [])
        typed = entry.get()
        filtered = self._filter_ingredients(full_values, typed)
        if filtered:
            self._show_filter_suggestions(entry, filtered)
        else:
            self._hide_filter_suggestions(entry)

    def _on_filter_entry_focus_in(self, event, entry):
        full_values = getattr(entry, "full_values", [])
        typed = entry.get()
        filtered = self._filter_ingredients(full_values, typed)
        if filtered:
            self._show_filter_suggestions(entry, filtered)

    def _on_filter_entry_focus_out(self, event, entry):
        entry.after(200, lambda: self._hide_filter_suggestions(entry))
        self._filter_rows()

    @staticmethod
    def _filter_ingredients(full_values, typed):
        if not typed:
            return full_values
        typed_key = ingredient_sort_key(typed)
        filtered = [v for v in full_values if ingredient_sort_key(v).startswith(typed_key)]
        if not filtered:
            filtered = [v for v in full_values if typed_key in ingredient_sort_key(v)]
        return filtered

    def _merge_item_into_current(self, name, qty, unit, rayon):
        """Ajoute un ingrédient à la liste déjà affichée, en cumulant sa
        quantité avec une ligne existante si le même ingrédient (même nom,
        même unité) y figure déjà, plutôt que de créer une ligne en double."""
        key = ingredient_sort_key(name)
        for item in self.current_items:
            if ingredient_sort_key(item["name"]) == key and item["unit"] == unit:
                item["quantity"] += qty
                return
        self.current_items.append({"name": name, "quantity": qty, "unit": unit, "rayon": rayon})

    def _add_recipe_to_cart(self, recipe, pers_entry):
        try:
            persons = float(pers_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("allrecipes_invalid_persons", name=recipe['name']))
            return

        grouped_totals = compute_grouped_totals([(recipe, persons)])
        for rayon, items in grouped_totals:
            for name, qty, unit in items:
                self._merge_item_into_current(name, qty, unit, rayon)

        # Met à jour la liste des recettes utilisées (pour les en-têtes des
        # exports) : si cette recette avait déjà été ajoutée, on remplace
        # son nombre de personnes plutôt que d'avoir une entrée en double.
        self.last_chosen_recipes = [
            (n, p) for (n, p) in self.last_chosen_recipes if n != recipe["name"]
        ]
        self.last_chosen_recipes.append((recipe["name"], persons))

        self._render_shopping_list()

    def _grouped_current_items(self):
        """Regroupe self.current_items par rayon, en conservant l'ordre des
        rayons et le tri alphabétique au sein de chaque rayon. Retourne une
        liste de (rayon, [indices dans self.current_items, triés])."""
        by_rayon = {}
        for i, item in enumerate(self.current_items):
            by_rayon.setdefault(item["rayon"], []).append(i)
        grouped = []
        for rayon in RAYON_ORDER:
            if rayon in by_rayon:
                idxs = sorted(by_rayon[rayon], key=lambda i: ingredient_sort_key(self.current_items[i]["name"]))
                grouped.append((rayon, idxs))
        return grouped

    def _render_shopping_list(self):
        for child in self.result_frame.winfo_children():
            child.destroy()

        if not self.current_items:
            ttk.Label(
                self.result_frame,
                text=t("allrecipes_empty_list_message"),
                foreground=COLOR_TEXT_MUTED, justify="center"
            ).pack(pady=20)
            return

        ttk.Label(self.result_frame, text=t("allrecipes_total_list_heading"),
                  font=("Segoe UI", sf(11), "bold")).pack(anchor="w", pady=(5, 2))
        if self.manual_items:
            ttk.Label(
                self.result_frame,
                text=t("allrecipes_manual_items_note", count=len(self.manual_items)),
                font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED
            ).pack(anchor="w")

        for rayon, idxs in self._grouped_current_items():
            ttk.Label(self.result_frame, text=translate_rayon_name(rayon), font=("Segoe UI", sf(10), "bold"),
                      foreground=COLOR_ACCENT_DARK).pack(anchor="w", pady=(12, 4))
            for idx in idxs:
                item = self.current_items[idx]
                row = ttk.Frame(self.result_frame)
                row.pack(fill="x", pady=1)
                ttk.Label(row, text=f"- {translate_ingredient_name(item['name'])}", width=30, anchor="w").pack(side="left")
                qty_entry = ttk.Entry(row, width=8)
                qty_entry.insert(0, str(item["quantity"]))
                qty_entry.pack(side="left", padx=3)
                qty_entry.bind("<FocusOut>", lambda e, i=idx, ent=qty_entry: self._update_item_quantity(i, ent))
                qty_entry.bind("<Return>", lambda e, i=idx, ent=qty_entry: self._update_item_quantity(i, ent))
                ttk.Label(row, text=translate_unit_name(item["unit"]), width=18, anchor="w").pack(side="left", padx=3)
                ttk.Button(row, text="🗑", width=3,
                           command=lambda i=idx: self._delete_item(i)).pack(side="left", padx=3)

        tk.Frame(self.result_frame, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _update_item_quantity(self, index, entry):
        if index >= len(self.current_items):
            return
        try:
            new_qty = float(entry.get().strip().replace(",", "."))
            if new_qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(t("common_error"), t("allrecipes_invalid_quantity"))
            entry.delete(0, tk.END)
            entry.insert(0, str(self.current_items[index]["quantity"]))
            return
        self.current_items[index]["quantity"] = new_qty

    def _delete_item(self, index):
        del self.current_items[index]
        self._render_shopping_list()

    def save_list_for_later(self):
        if not self.current_items:
            messagebox.showinfo(t("common_info"), t("allrecipes_calculate_list_first"),
                                 parent=self)
            return
        name = simpledialog.askstring(
            t("allrecipes_save_list_dialog_title"), t("allrecipes_save_list_dialog_prompt"), parent=self
        )
        if not name:
            self.lift()
            self.focus_force()
            return
        name = name.strip()
        if not name:
            self.lift()
            self.focus_force()
            return
        lists = load_saved_shopping_lists()
        lists = [l for l in lists if l["name"].lower() != name.lower()]  # remplace une liste de même nom
        lists.append({
            "name": name,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "items": [dict(item) for item in self.current_items],
        })
        save_saved_shopping_lists(lists)
        messagebox.showinfo(t("allrecipes_list_saved_title"), t("allrecipes_list_saved_message", name=name), parent=self)
        # Une messagebox sans fenêtre parente peut parfois faire remonter la
        # page d'accueil au premier plan une fois fermée : on force cette
        # fenêtre à revenir au premier plan par sécurité.
        self.lift()
        self.focus_force()

    def open_saved_lists(self):
        SavedShoppingListsWindow(self.app, self)

    def load_saved_list(self, items):
        self.current_items = [dict(item) for item in items]
        self.last_chosen_recipes = []  # une liste chargée n'est pas liée à une sélection de recettes
        self._render_shopping_list()

    def _current_export_data(self):
        """Renvoie (chosen_recipes, grouped_totals) à partir de la liste
        actuellement affichée (self.current_items), en tenant compte des
        modifications de quantité et des suppressions faites à la main."""
        if not self.current_items:
            messagebox.showinfo(t("common_info"), t("allrecipes_empty_list_for_export"))
            return None
        grouped_totals = grouped_totals_from_flat_items(self.current_items)
        return self.last_chosen_recipes, grouped_totals

    def export_txt(self):
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result

        path = filedialog.asksaveasfilename(
            title=t("allrecipes_export_txt_title"),
            defaultextension=".txt",
            filetypes=[("Fichier texte", "*.txt")],
            initialfile="liste_de_courses.txt"
        )
        if not path:
            return
        try:
            write_shopping_list_txt(path, t("allrecipes_shopping_list_title"), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("allrecipes_export_saved_message", path=path))

    def export_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("allrecipes_excel_module_missing"))
            return

        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result

        path = filedialog.asksaveasfilename(
            title=t("allrecipes_export_excel_title"),
            defaultextension=".xlsx",
            filetypes=[("Fichier Excel", "*.xlsx")],
            initialfile="liste_de_courses.xlsx"
        )
        if not path:
            return

        try:
            wb = build_shopping_list_workbook(chosen_recipes, grouped_totals)
            wb.save(path)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("allrecipes_export_saved_message", path=path))

    def export_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("allrecipes_pdf_module_missing"))
            return

        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result

        path = filedialog.asksaveasfilename(
            title=t("allrecipes_export_pdf_title"),
            defaultextension=".pdf",
            filetypes=[("Fichier PDF", "*.pdf")],
            initialfile="liste_de_courses.pdf"
        )
        if not path:
            return

        try:
            build_shopping_list_pdf(path, t("allrecipes_shopping_list_title"), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("allrecipes_export_saved_message", path=path))

    def print_shopping_list(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("allrecipes_print_module_missing"))
            return

        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result

        temp_path = get_temp_pdf_path("liste_de_courses")
        try:
            build_shopping_list_pdf(temp_path, t("allrecipes_shopping_list_title"), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_print_failed", error=e))
            return

        result_status = print_file(temp_path)
        report_print_result(result_status, temp_path, t("allrecipes_print_label"))

    def open_checklist(self):
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        ShoppingChecklistWindow(self.app, grouped_totals, title=t("allrecipes_shopping_list_title"))


class QuickSearchWindow(tk.Toplevel):
    """Recherche rapide de recette, accessible depuis n'importe quelle
    fenêtre de l'application via le raccourci Ctrl+K."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("quicksearch_title"))
        self.geometry(f"{gs(480)}x{gs(420)}")
        self.minsize(gs(400), gs(360))
        self.resizable(True, True)
        self.grab_set()
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            pass

        ttk.Label(self, text=t("quicksearch_heading"),
                  font=("Segoe UI", sf(12), "bold")).pack(pady=(15, 5))
        self.search_entry = ttk.Entry(self, font=("Segoe UI", sf(11)))
        self.search_entry.pack(padx=15, pady=(0, 10), fill="x")
        self.search_entry.focus_set()
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate())
        self.search_entry.bind("<Return>", lambda e: self._open_first_or_selected())
        self.search_entry.bind("<Down>", lambda e: self.listbox.focus_set())
        self.bind("<Escape>", lambda e: self.destroy())

        list_frame = ttk.Frame(self)
        list_frame.pack(padx=15, pady=(0, 10), fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<Return>", lambda e: self._open_selected())
        self.listbox.bind("<Double-Button-1>", lambda e: self._open_selected())

        self.matched_names = []
        self._populate()

        ttk.Label(self, text=t("quicksearch_footer_hint"),
                  font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED).pack(pady=(0, 10))

    def _populate(self):
        self.listbox.delete(0, tk.END)
        search = self.search_entry.get().strip()
        search_key = ingredient_sort_key(search) if search else ""
        self.matched_names = []
        for recipe in self.app.recipes:
            if search_key and not recipe_matches_search(recipe, search_key):
                continue
            self.listbox.insert(tk.END, format_recipe_list_label(recipe))
            self.matched_names.append(recipe["name"])
        if not self.matched_names:
            self.listbox.insert(tk.END, t("quicksearch_no_results"))

    def _open_selected(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.matched_names):
            return
        name = self.matched_names[sel[0]]
        self.destroy()
        OneRecipeWindow(self.app, initial_recipe_name=name)

    def _open_first_or_selected(self):
        sel = self.listbox.curselection()
        if sel:
            self._open_selected()
        elif self.matched_names:
            name = self.matched_names[0]
            self.destroy()
            OneRecipeWindow(self.app, initial_recipe_name=name)


class OneRecipeWindow(tk.Toplevel):
    """Fenêtre pour afficher une recette précise (avec ses photos), avec
    quantités recalculées selon le nombre de personnes choisi."""

    def __init__(self, app, initial_recipe_name=None):
        super().__init__(app)
        self.app = app
        self.title(t("onerecipe_window_title"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(1100)}x{screen_height}+40+0")
        self.minsize(gs(760), gs(500))
        self.resizable(True, True)
        self.grab_set()
        self._gallery_thumb_refs = []
        self.current_recipe = None
        self.filtered_indices = []

        ttk.Label(self, text=t("onerecipe_choose_recipe_label"), font=("Segoe UI", sf(11), "bold")).pack(pady=(10, 5))

        top_frame = ttk.Frame(self)
        top_frame.pack(pady=(0, 5), fill="x", padx=15)
        ttk.Label(top_frame, text=t("onerecipe_search_label")).pack(side="left")
        self.search_entry = ttk.Entry(top_frame, width=16)
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate())
        ttk.Label(top_frame, text=t("onerecipe_sort_label")).pack(side="left", padx=(5, 2))
        self.sort_combo = ttk.Combobox(top_frame, values=[translate_sort_option(o) for o in RECIPE_SORT_OPTIONS], state="readonly", width=16)
        self.sort_combo.set(translate_sort_option(RECIPE_SORT_OPTIONS[0]))
        self.sort_combo.pack(side="left")
        self.sort_combo.bind("<<ComboboxSelected>>", lambda e: self._populate())
        ttk.Label(top_frame, text=t("onerecipe_category_label")).pack(side="left", padx=(5, 2))
        self.category_filter_combo = ttk.Combobox(
            top_frame, values=[t("common_all_categories")] + [translate_category_name(c) for c in RecipeFormWindow.CATEGORY_OPTIONS],
            state="readonly", width=14
        )
        self.category_filter_combo.set(t("common_all_categories"))
        self.category_filter_combo.pack(side="left")
        self.category_filter_combo.bind("<<ComboboxSelected>>", lambda e: self._populate())

        list_frame = ttk.Frame(self)
        list_frame.pack(pady=5, padx=15, fill="both")
        list_canvas = tk.Canvas(list_frame, height=170, highlightthickness=0)
        list_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=list_canvas.yview)
        self.rows_frame = ttk.Frame(list_canvas)
        self.rows_frame.bind("<Configure>", lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all")))
        list_canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        list_canvas.configure(yscrollcommand=list_scrollbar.set)
        list_canvas.pack(side="left", fill="both", expand=True)
        list_scrollbar.pack(side="right", fill="y")
        self.selected_actual_index = None
        self._selected_row_widgets = None
        self._row_widgets = []  # [(actual_index, row_frame, label_widget), ...] pour la ligne affichée
        self._populate()

        # ---- Galerie de photos ----
        gallery_outer = ttk.Frame(self)
        gallery_outer.pack(fill="x", padx=15, pady=(10, 0))
        self.gallery_canvas = tk.Canvas(gallery_outer, height=140, highlightthickness=0)
        gallery_scrollbar = ttk.Scrollbar(gallery_outer, orient="horizontal",
                                           command=self.gallery_canvas.xview)
        self.gallery_frame = ttk.Frame(self.gallery_canvas)
        self.gallery_frame.bind(
            "<Configure>", lambda e: self.gallery_canvas.configure(scrollregion=self.gallery_canvas.bbox("all"))
        )
        self.gallery_canvas.create_window((0, 0), window=self.gallery_frame, anchor="nw")
        self.gallery_canvas.configure(xscrollcommand=gallery_scrollbar.set)
        self.gallery_canvas.pack(fill="x")
        gallery_scrollbar.pack(fill="x")

        # ---- Nombre de personnes + ajustement rapide ----
        persons_frame = ttk.Frame(self)
        persons_frame.pack(pady=(10, 5))
        ttk.Label(persons_frame, text=t("onerecipe_persons_label")).grid(row=0, column=0, columnspan=4, pady=(0, 5))
        self.pers_entry = ttk.Entry(persons_frame, width=8)
        self.pers_entry.insert(0, "1")
        self.pers_entry.grid(row=1, column=0, padx=3)
        ttk.Button(persons_frame, text="−1", width=4,
                   command=lambda: self._adjust_persons(delta=-1)).grid(row=1, column=1, padx=3)
        ttk.Button(persons_frame, text="+1", width=4,
                   command=lambda: self._adjust_persons(delta=1)).grid(row=1, column=2, padx=3)
        ttk.Button(persons_frame, text="÷2", width=4,
                   command=lambda: self._adjust_persons(factor=0.5)).grid(row=1, column=3, padx=3)
        ttk.Button(persons_frame, text="×2", width=4,
                   command=lambda: self._adjust_persons(factor=2)).grid(row=1, column=4, padx=3)

        # ---- Boutons d'action, alignés en rangées de 4 ----
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10, padx=15, fill="x")
        action_buttons = [
            (t("onerecipe_btn_show"), self.show_recipe),
            (t("onerecipe_btn_export_pdf"), self.export_recipe_pdf),
            (t("onerecipe_btn_print"), self.print_recipe),
            (t("onerecipe_btn_add_to_shopping"), self.add_to_shopping_list),
            (t("onerecipe_btn_cooked"), self.mark_as_cooked),
            (t("onerecipe_btn_cooking_mode"), self.open_cooking_mode),
            (t("onerecipe_btn_qr"), self.show_qr_code),
            (t("onerecipe_btn_timers"), self.open_timers),
            (t("onerecipe_btn_cook_log"), self.open_cook_log),
            (t("onerecipe_btn_substitutions"), self.show_substitutions),
        ]
        for col in range(4):
            btn_frame.columnconfigure(col, weight=1)
        for i, (text, command) in enumerate(action_buttons):
            row, col = divmod(i, 4)
            ttk.Button(btn_frame, text=text, command=command).grid(
                row=row, column=col, padx=4, pady=4, sticky="ew"
            )

        # ---- Deux panneaux côte à côte : ingrédients/infos à gauche,
        # description/notes à droite. ----
        results_frame = ttk.Frame(self)
        results_frame.pack(pady=5, padx=15, fill="both", expand=True)

        left_results = ttk.Frame(results_frame)
        left_results.pack(side="left", fill="both", expand=True, padx=(0, 8))
        ttk.Label(left_results, text=t("onerecipe_ingredients_info_label"),
                  font=("Segoe UI", sf(9), "bold")).pack(anchor="w")
        self.result_text = tk.Text(left_results, width=48, height=14, wrap="word", font=("Segoe UI", sf(10)))
        self.result_text.pack(fill="both", expand=True)

        right_results = ttk.Frame(results_frame)
        right_results.pack(side="left", fill="both", expand=True, padx=(8, 0))
        ttk.Label(right_results, text=t("onerecipe_description_notes_label"),
                  font=("Segoe UI", sf(9), "bold")).pack(anchor="w")
        self.description_result_text = tk.Text(right_results, width=48, height=14, wrap="word", font=("Segoe UI", sf(10)))
        self.description_result_text.pack(fill="both", expand=True)

        # ---- Recettes similaires : suggestions basées sur la catégorie, les
        # étiquettes et les ingrédients en commun. ----
        self.similar_frame = ttk.Frame(self)
        self.similar_frame.pack(pady=(0, 10), padx=15, fill="x")

        # Espace vide en bas de la fenêtre, pour que "Recettes similaires"
        # ne se retrouve jamais collé au bord inférieur (ou caché derrière
        # la barre des tâches sur certains systèmes).
        tk.Frame(self, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

        if initial_recipe_name:
            for idx, row, label in self._row_widgets:
                if self.app.recipes[idx]["name"] == initial_recipe_name:
                    self._select_row(idx, row, show=True)
                    break

    def _populate(self):
        for child in self.rows_frame.winfo_children():
            child.destroy()
        self._row_widgets = []
        self.filtered_indices = []
        search = self.search_entry.get().strip()
        search_key = ingredient_sort_key(search) if search else ""
        option = resolve_sort_option_input(self.sort_combo.get(), RECIPE_SORT_OPTIONS)
        category_filter = resolve_category_input(self.category_filter_combo.get(), RecipeFormWindow.CATEGORY_OPTIONS)
        indexed = list(enumerate(self.app.recipes))
        indexed = [pair for pair in indexed if recipe_matches_search(pair[1], search_key)]
        if category_filter and category_filter != t("common_all_categories"):
            indexed = [pair for pair in indexed if pair[1].get("category", "Autre") == category_filter]
        reverse = option in ("Ajoutées récemment",)
        indexed.sort(key=lambda pair: recipe_sort_key(pair[1], option), reverse=reverse)
        for row_index, (idx, recipe) in enumerate(indexed):
            row = tk.Frame(self.rows_frame, background=COLOR_BG)
            row.grid(row=row_index, column=0, sticky="ew", pady=1)
            label = tk.Label(row, text=format_recipe_list_label(recipe), background=COLOR_BG,
                              anchor="w", cursor="hand2", padx=4, font=("Segoe UI", sf(10)))
            label.pack(side="left", fill="x", expand=True)
            label.bind("<Button-1>", lambda e, i=idx, r=row: self._select_row(i, r))
            label.bind("<Double-Button-1>", lambda e, i=idx, r=row: self._select_row(i, r, show=True))
            ttk.Button(row, text=t("onerecipe_edit_button"), width=10,
                       command=lambda i=idx: self._edit_recipe(i)).pack(side="right", padx=4)
            self._row_widgets.append((idx, row, label))
            self.filtered_indices.append(idx)
        if self.selected_actual_index is not None:
            for idx, row, label in self._row_widgets:
                if idx == self.selected_actual_index:
                    row.configure(background=COLOR_ACCENT_LIGHT)
                    label.configure(background=COLOR_ACCENT_LIGHT)
                    self._selected_row_widgets = (row, label)
                    break

    def _select_row(self, actual_index, row, show=False):
        if self._selected_row_widgets is not None:
            prev_row, prev_label = self._selected_row_widgets
            try:
                prev_row.configure(background=COLOR_BG)
                prev_label.configure(background=COLOR_BG)
            except tk.TclError:
                pass  # la ligne précédente a pu être détruite par un _populate entre-temps
        label = None
        for idx, r, l in self._row_widgets:
            if r is row:
                label = l
                break
        row.configure(background=COLOR_ACCENT_LIGHT)
        if label is not None:
            label.configure(background=COLOR_ACCENT_LIGHT)
        self._selected_row_widgets = (row, label)
        self.selected_actual_index = actual_index
        if show:
            self.show_recipe()

    def _edit_recipe(self, index):
        win = RecipeFormWindow(self.app, recipe_index=index)
        self.wait_window(win)
        self.app.refresh_recipes()
        was_displaying_this = (
            self.current_recipe is not None
            and self.selected_actual_index == index
        )
        self._populate()
        if was_displaying_this and index < len(self.app.recipes):
            self.current_recipe = self.app.recipes[index]
            self._display_recipe(self.current_recipe)

    def _refresh_gallery(self, recipe):
        for child in self.gallery_frame.winfo_children():
            child.destroy()
        self._gallery_thumb_refs = []

        images = get_recipe_images(recipe)
        if not images:
            ttk.Label(self.gallery_frame, text=t("onerecipe_no_photo")).pack(side="left", padx=10, pady=10)
            return

        for fname in images:
            thumb = load_thumbnail(fname, size=(160, 120))
            cell = ttk.Frame(self.gallery_frame)
            cell.pack(side="left", padx=5, pady=5)
            if thumb is not None:
                self._gallery_thumb_refs.append(thumb)
                ttk.Label(cell, image=thumb).pack()
            else:
                ttk.Label(cell, text=t("onerecipe_preview_unavailable")).pack()

    def _adjust_persons(self, factor=None, delta=None):
        try:
            current = float(self.pers_entry.get().strip().replace(",", "."))
        except ValueError:
            current = 1.0
        if delta is not None:
            current = current + delta
        if factor is not None:
            current = current * factor
        current = max(0.5, current)
        if current == int(current):
            current = int(current)
        self.pers_entry.delete(0, tk.END)
        self.pers_entry.insert(0, str(current))
        if self.current_recipe is not None:
            self._display_recipe(self.current_recipe)

    def show_recipe(self):
        if self.selected_actual_index is None or self.selected_actual_index >= len(self.app.recipes):
            messagebox.showinfo(t("common_info"), t("onerecipe_select_recipe_first"))
            return
        idx = self.selected_actual_index
        recipe = self.app.recipes[idx]
        is_new_selection = self.current_recipe is not recipe
        self.current_recipe = recipe
        if is_new_selection:
            self.pers_entry.delete(0, tk.END)
            self.pers_entry.insert(0, str(recipe.get("default_persons", 1) or 1))
            record_recipe_view(recipe["name"])
        self._display_recipe(recipe)

    def add_to_shopping_list(self):
        if self.current_recipe is None:
            messagebox.showinfo(t("common_info"), t("onerecipe_display_first"))
            return
        try:
            persons = float(self.pers_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("onerecipe_invalid_persons"))
            return
        self.app.shopping_selection[self.current_recipe["name"]] = persons
        messagebox.showinfo(
            t("onerecipe_added_to_shopping_title"),
            t("onerecipe_added_to_shopping_message", name=self.current_recipe['name'], persons=persons)
        )

    def mark_as_cooked(self):
        if self.current_recipe is None:
            messagebox.showinfo(t("common_info"), t("onerecipe_display_first"))
            return
        recipes = load_recipes()
        target_name = self.current_recipe.get("name")
        for r in recipes:
            if r is self.current_recipe or r.get("name") == target_name:
                r["times_cooked"] = r.get("times_cooked", 0) + 1
                cooked_dates = r.get("cooked_dates", [])
                cooked_dates.append(datetime.now().strftime("%Y-%m-%d"))
                r["cooked_dates"] = cooked_dates
                self.current_recipe = r
                break
        save_recipes(recipes)
        self.app.refresh_recipes()

        # Propose de décompter du garde-manger les ingrédients utilisés, si
        # l'utilisateur en tient un (jamais aucune supposition sur les
        # ingrédients absents ou aux unités non comparables, voir
        # decrement_pantry_for_recipe).
        pantry = load_pantry()
        if pantry:
            try:
                persons = float(self.pers_entry.get().strip().replace(",", "."))
            except ValueError:
                persons = self.current_recipe.get("default_persons", 1) or 1
            if messagebox.askyesno(
                t("onerecipe_pantry_decrement_title"),
                t("onerecipe_pantry_decrement_prompt", name=target_name, persons=persons)
            ):
                count = decrement_pantry_for_recipe(self.current_recipe, persons)
                if count:
                    messagebox.showinfo(
                        t("onerecipe_pantry_updated_title"),
                        t("onerecipe_pantry_updated_message", count=count)
                    )
                else:
                    messagebox.showinfo(t("common_info"), t("onerecipe_pantry_none_decremented"))

        def _on_log_done(note, photo_filename):
            recipes2 = load_recipes()
            for r in recipes2:
                if r.get("name") == target_name:
                    cook_log = r.get("cook_log", [])
                    cook_log.append({
                        "date": datetime.now().strftime("%Y-%m-%d"),
                        "note": note,
                        "photo": photo_filename,
                    })
                    r["cook_log"] = cook_log
                    self.current_recipe = r
                    break
            save_recipes(recipes2)
            self.app.refresh_recipes()
            messagebox.showinfo(t("onerecipe_marked_title"), t("onerecipe_marked_message", name=target_name))

        CookLogEntryDialog(self.app, target_name, _on_log_done)

    def open_cook_log(self):
        if self.current_recipe is None:
            messagebox.showinfo(t("common_info"), t("onerecipe_display_first"))
            return
        CookLogWindow(self.app, self.current_recipe)

    def show_substitutions(self):
        if self.current_recipe is None:
            messagebox.showinfo(t("common_info"), t("onerecipe_display_first"))
            return
        recipe = self.current_recipe
        entries = []
        seen_keys = set()
        for ing in recipe["ingredients"]:
            key = ingredient_sort_key(ing["name"])
            if key in seen_keys:
                continue
            subs = get_display_ingredient_substitutions(ing["name"])
            if subs:
                seen_keys.add(key)
                entries.append((ing["name"], subs))
        if not entries:
            messagebox.showinfo(
                t("onerecipe_no_substitutes_title"),
                t("onerecipe_no_substitutes_message")
            )
            return

        win = tk.Toplevel(self)
        win.title(t("onerecipe_substitutes_title", name=recipe['name']))
        win.geometry("520x520")
        win.minsize(440, 400)
        win.resizable(True, True)
        win.grab_set()

        ttk.Label(win, text=t("onerecipe_substitutes_heading", name=recipe['name']),
                  font=("Segoe UI", sf(12), "bold"), wraplength=480, justify="center").pack(pady=(15, 5))
        ttk.Label(
            win, text=t("onerecipe_substitutes_disclaimer"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

        container = ttk.Frame(win)
        container.pack(fill="both", expand=True, padx=15, pady=5)
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for ing_name, subs in entries:
            ttk.Label(inner, text=translate_ingredient_name(ing_name).capitalize(), font=("Segoe UI", sf(10), "bold"),
                      foreground=COLOR_ACCENT_DARK).pack(anchor="w", pady=(10, 2))
            for sub in subs:
                note = f" — {sub['note']}" if sub.get("note") else ""
                ttk.Label(inner, text=f"  • {sub['nom']}{note}", wraplength=440,
                          justify="left").pack(anchor="w")

        tk.Frame(win, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")
        ttk.Button(win, text=t("onerecipe_close_button"), command=win.destroy).pack(pady=10)

    def _display_recipe(self, recipe):
        try:
            persons = float(self.pers_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("onerecipe_invalid_persons"))
            return

        self._refresh_gallery(recipe)

        self.result_text.delete("1.0", tk.END)
        self.description_result_text.delete("1.0", tk.END)
        cat = translate_category_name(recipe.get("category", "Autre"))
        star = "⭐ " if recipe.get("favorite") else ""
        self.result_text.insert(tk.END, f"=== {star}[{cat}] {recipe['name']} ({persons} pers.) ===\n\n")

        rating = recipe.get("rating", 0)
        if rating:
            self.result_text.insert(tk.END, t("onerecipe_rating_label", stars=rating_stars(rating)) + "\n\n")

        info_bits = []
        if recipe.get("prep_time"):
            info_bits.append(t("onerecipe_prep_label", time=recipe['prep_time']))
        if recipe.get("cook_time"):
            info_bits.append(t("onerecipe_cook_label", time=recipe['cook_time']))
        if recipe.get("difficulty"):
            info_bits.append(t("onerecipe_difficulty_label", value=translate_difficulty_name(recipe['difficulty'])))
        if info_bits:
            self.result_text.insert(tk.END, " | ".join(info_bits) + "\n\n")

        allergens = recipe.get("allergens") or []
        if allergens:
            self.result_text.insert(tk.END, t("onerecipe_allergens_label", list=", ".join(translate_allergen_name(a) for a in allergens)) + "\n\n")

        for ing in recipe["ingredients"]:
            qty = round(ing["quantity"] * persons, 2)
            unit = f" {translate_unit_name(ing['unit'])}" if ing["unit"] else ""
            self.result_text.insert(tk.END, f"- {translate_ingredient_name(ing['name']).capitalize()} : {qty}{unit}\n")

        cost, cost_known, cost_total = compute_recipe_cost(recipe, persons)
        if cost_known:
            partial = "" if cost_known == cost_total else t("onerecipe_cost_partial", known=cost_known, total=cost_total)
            self.result_text.insert(tk.END, "\n" + t("onerecipe_cost_label", cost=f"{cost:.2f}", partial=partial) + "\n")

        nutrition, nutri_known, nutri_total = compute_recipe_nutrition(recipe, persons)
        if nutri_known:
            partial = "" if nutri_known == nutri_total else t(
                "onerecipe_nutrition_partial", known=nutri_known, total=nutri_total
            )
            self.result_text.insert(
                tk.END,
                t(
                    "onerecipe_nutrition_label", partial=partial,
                    kcal=f"{nutrition['kcal']:.0f}", protein=f"{nutrition['protein_g']:.0f}",
                    carbs=f"{nutrition['carbs_g']:.0f}", fat=f"{nutrition['fat_g']:.0f}"
                )
            )

        description = recipe.get("description", "").strip()
        if description:
            self.description_result_text.insert(tk.END, t("onerecipe_description_heading", text=description))
        personal_notes = recipe.get("personal_notes", "").strip()
        if personal_notes:
            self.description_result_text.insert(tk.END, t("onerecipe_notes_heading", text=personal_notes))
        if not description and not personal_notes:
            self.description_result_text.insert(tk.END, t("onerecipe_no_description_notes"))

        self._render_similar_recipes(recipe)

    def _render_similar_recipes(self, recipe):
        for child in self.similar_frame.winfo_children():
            child.destroy()
        similar = find_similar_recipes(recipe, self.app.recipes, limit=5)
        if not similar:
            return
        ttk.Label(self.similar_frame, text=t("onerecipe_similar_label"),
                  font=("Segoe UI", sf(9), "bold"), foreground=COLOR_ACCENT_DARK).pack(anchor="w")
        links_frame = ttk.Frame(self.similar_frame)
        links_frame.pack(anchor="w", pady=(2, 0))
        for other in similar:
            cat = translate_category_name(other.get("category", "Autre"))
            btn = tk.Label(
                links_frame, text=f"[{cat}] {other['name']}", foreground=COLOR_ACCENT_DARK,
                cursor="hand2", font=("Segoe UI", sf(9), "underline")
            )
            btn.pack(side="left", padx=(0, 12))
            btn.bind("<Button-1>", lambda e, name=other["name"]: self._open_similar_recipe(name))

    def _open_similar_recipe(self, recipe_name):
        for idx, row, label in self._row_widgets:
            if self.app.recipes[idx]["name"] == recipe_name:
                self._select_row(idx, row, show=True)
                return
        # La recette n'est pas dans la liste actuellement filtrée (recherche/
        # catégorie en cours) : on réinitialise les filtres pour la retrouver.
        self.search_entry.delete(0, tk.END)
        self.category_filter_combo.set(t("common_all_categories"))
        self._populate()
        for idx, row, label in self._row_widgets:
            if self.app.recipes[idx]["name"] == recipe_name:
                self._select_row(idx, row, show=True)
                return

    @staticmethod
    def _build_recipe_pdf(path, recipe, persons):
        c = pdf_canvas.Canvas(path, pagesize=A4)
        width, height = A4
        draw_recipe_content(c, recipe, persons, width, height)
        c.save()

    def export_recipe_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("onerecipe_pdf_module_missing"))
            return
        if self.current_recipe is None:
            messagebox.showinfo(t("common_info"), t("onerecipe_display_first"))
            return
        try:
            persons = float(self.pers_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("onerecipe_invalid_persons"))
            return

        recipe = self.current_recipe
        path = filedialog.asksaveasfilename(
            title=t("onerecipe_export_pdf_title"),
            defaultextension=".pdf",
            filetypes=[("Fichier PDF", "*.pdf")],
            initialfile=f"{recipe['name']}.pdf"
        )
        if not path:
            return

        try:
            self._build_recipe_pdf(path, recipe, persons)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("onerecipe_export_failed", error=e))
            return
        messagebox.showinfo(t("onerecipe_export_success_title"), t("onerecipe_export_success_message", path=path))

    def print_recipe(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("onerecipe_print_module_missing"))
            return
        if self.current_recipe is None:
            messagebox.showinfo(t("common_info"), t("onerecipe_display_first"))
            return
        try:
            persons = float(self.pers_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("onerecipe_invalid_persons"))
            return

        recipe = self.current_recipe
        temp_path = get_temp_pdf_path("recette")
        try:
            self._build_recipe_pdf(temp_path, recipe, persons)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("onerecipe_print_failed", error=e))
            return

        result_status = print_file(temp_path)
        report_print_result(result_status, temp_path, f"« {recipe['name']} »")

    def show_qr_code(self):
        if not QRCODE_AVAILABLE:
            messagebox.showerror(
                t("common_module_missing"),
                t("onerecipe_qr_module_missing")
            )
            return
        if not PIL_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("onerecipe_qr_pillow_missing"))
            return
        if self.current_recipe is None:
            messagebox.showinfo(t("common_info"), t("onerecipe_display_first"))
            return
        try:
            persons = float(self.pers_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("onerecipe_invalid_persons"))
            return
        QRCodeWindow(self.app, self.current_recipe, persons)

    def open_timers(self):
        recipe = self.current_recipe
        label_text = recipe["name"] if recipe else t("onerecipe_default_timer_label")
        default_minutes = 10
        if recipe is not None:
            try:
                if recipe.get("cook_time"):
                    default_minutes = max(1, round(float(recipe["cook_time"])))
                elif recipe.get("prep_time"):
                    default_minutes = max(1, round(float(recipe["prep_time"])))
            except (TypeError, ValueError):
                pass

        # Cette fenêtre est modale (grab_set), ce qui bloquerait toute autre
        # fenêtre de l'application — y compris la fenêtre des minuteurs, qui
        # doit rester utilisable en même temps qu'on consulte la recette.
        # On relâche donc le grab ici ; les deux fenêtres restent ensuite
        # utilisables librement en parallèle.
        try:
            self.grab_release()
        except tk.TclError:
            pass

        # Une seule fenêtre de minuteurs pour toute l'application : si elle
        # est déjà ouverte, on y ajoute simplement un minuteur de plus au
        # lieu d'en ouvrir une deuxième.
        existing = getattr(self.app, "timers_window", None)
        if existing is not None and existing.winfo_exists():
            existing.add_timer(label_text, default_minutes)
            existing.deiconify()
            existing.lift()
            existing.focus_force()
        else:
            self.app.timers_window = TimersWindow(self.app, label_text, default_minutes)

    def open_cooking_mode(self):
        if self.current_recipe is None:
            messagebox.showinfo(t("common_info"), t("onerecipe_display_first"))
            return
        try:
            persons = float(self.pers_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("onerecipe_invalid_persons"))
            return
        # CookingModeWindow n'a pas son propre grab_set() : sans relâcher
        # celui de cette fenêtre, le mode cuisine serait inutilisable (même
        # pas fermable) tant que "Voir une recette précise" reste ouverte.
        try:
            self.grab_release()
        except tk.TclError:
            pass
        CookingModeWindow(self.app, self.current_recipe, persons)


class CookingModeWindow(tk.Toplevel):
    """Affichage plein écran, en gros caractères et sans menus, d'une
    recette — pratique à consulter en cuisinant, posé à côté des
    fourneaux."""

    def __init__(self, app, recipe, persons):
        super().__init__(app)
        self.app = app
        self.recipe = recipe
        self.persons = persons
        self.title(t("cookingmode_title", name=recipe['name']))
        self.configure(bg="white")
        self._is_fullscreen = False
        # Une fenêtre maximisée (plutôt qu'un vrai plein écran) s'ouvre
        # instantanément : le vrai plein écran provoque, sur certains
        # systèmes Windows, un blocage de plusieurs secondes le temps que la
        # transition d'affichage se fasse. Le plein écran natif reste
        # disponible via F11 pour qui le souhaite.
        try:
            self.state("zoomed")
        except tk.TclError:
            self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")
        self.tts_engine = None
        self.tts_thread = None
        self.speech_volume = 1.0       # 0.0 (muet) à 1.0 (plein volume)
        self.bind("<Escape>", lambda e: self._on_close())
        self.bind("<F11>", lambda e: self._toggle_fullscreen())
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        top_bar = tk.Frame(self, bg="white")
        top_bar.pack(fill="x", pady=10)
        tk.Button(top_bar, text=t("cookingmode_close_button"), font=("Segoe UI", sf(13)),
                  command=self._on_close).pack(side="right", padx=30)
        tk.Button(top_bar, text=t("cookingmode_cooked_button"), font=("Segoe UI", sf(13)),
                  command=self.mark_as_cooked).pack(side="right", padx=(10, 0))
        tk.Label(top_bar, text=t("cookingmode_fullscreen_hint"), font=("Segoe UI", sf(9)),
                 bg="white", fg="#999").pack(side="right", padx=10)

        volume_frame = tk.Frame(top_bar, bg="white")
        volume_frame.pack(side="right", padx=(10, 0))
        tk.Button(volume_frame, text="🔊+", font=("Segoe UI", sf(11)), width=4,
                  command=lambda: self._adjust_volume(0.1)).pack(side="right")
        self.volume_label = tk.Label(volume_frame, text=t("cookingmode_volume_percent", percent=100), font=("Segoe UI", sf(10)),
                                      bg="white", fg="#666", width=5)
        self.volume_label.pack(side="right", padx=3)
        tk.Button(volume_frame, text="🔉−", font=("Segoe UI", sf(11)), width=4,
                  command=lambda: self._adjust_volume(-0.1)).pack(side="right")

        self.speech_button = tk.Button(top_bar, text=t("cookingmode_speech_button"), font=("Segoe UI", sf(13)),
                                        command=self.toggle_speech)
        self.speech_button.pack(side="right", padx=(10, 0))

        pers_frame = tk.Frame(top_bar, bg="white")
        pers_frame.pack(side="left", padx=30)
        tk.Button(pers_frame, text="−", font=("Segoe UI", sf(14), "bold"), width=3,
                  command=lambda: self._adjust(-1)).pack(side="left")
        self.pers_label = tk.Label(pers_frame, text=t("cookingmode_persons_suffix", persons=self._fmt(persons)),
                                    font=("Segoe UI", sf(14)), bg="white")
        self.pers_label.pack(side="left", padx=10)
        tk.Button(pers_frame, text="+", font=("Segoe UI", sf(14), "bold"), width=3,
                  command=lambda: self._adjust(1)).pack(side="left")

        outer = tk.Frame(self, bg="white")
        outer.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(outer, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.content = tk.Frame(self.canvas, bg="white")
        self.content.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.content, anchor="n")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True, padx=60)
        scrollbar.pack(side="right", fill="y")

        self._render()

    def mark_as_cooked(self):
        recipes = load_recipes()
        target_name = self.recipe.get("name")
        for r in recipes:
            if r is self.recipe or r.get("name") == target_name:
                r["times_cooked"] = r.get("times_cooked", 0) + 1
                cooked_dates = r.get("cooked_dates", [])
                cooked_dates.append(datetime.now().strftime("%Y-%m-%d"))
                r["cooked_dates"] = cooked_dates
                self.recipe = r
                break
        save_recipes(recipes)
        self.app.refresh_recipes()

        # Propose de décompter du garde-manger les ingrédients utilisés, si
        # l'utilisateur en tient un (jamais aucune supposition sur les
        # ingrédients absents ou aux unités non comparables, voir
        # decrement_pantry_for_recipe).
        pantry = load_pantry()
        if pantry:
            if messagebox.askyesno(
                t("onerecipe_pantry_decrement_title"),
                t("onerecipe_pantry_decrement_prompt", name=target_name, persons=self._fmt(self.persons))
            ):
                count = decrement_pantry_for_recipe(self.recipe, self.persons)
                if count:
                    messagebox.showinfo(
                        t("onerecipe_pantry_updated_title"),
                        t("onerecipe_pantry_updated_message", count=count)
                    )
                else:
                    messagebox.showinfo(
                        t("common_info"),
                        t("onerecipe_pantry_none_decremented")
                    )

        def _on_log_done(note, photo_filename):
            recipes2 = load_recipes()
            for r in recipes2:
                if r.get("name") == target_name:
                    cook_log = r.get("cook_log", [])
                    cook_log.append({
                        "date": datetime.now().strftime("%Y-%m-%d"),
                        "note": note,
                        "photo": photo_filename,
                    })
                    r["cook_log"] = cook_log
                    self.recipe = r
                    break
            save_recipes(recipes2)
            self.app.refresh_recipes()
            messagebox.showinfo(t("onerecipe_marked_title"), t("onerecipe_marked_message", name=target_name))

        CookLogEntryDialog(self.app, target_name, _on_log_done)

    def toggle_speech(self):
        if self.tts_thread is not None and self.tts_thread.is_alive():
            self.stop_speech()
        else:
            self.start_speech()

    def start_speech(self):
        if not PYTTSX3_AVAILABLE:
            messagebox.showerror(
                t("common_module_missing"),
                t("cookingmode_tts_module_missing")
            )
            return
        text = (self.recipe.get("description") or "").strip()
        if not text:
            messagebox.showinfo(
                t("common_info"), t("cookingmode_no_description_to_read")
            )
            return

        self.speech_button.config(text=t("cookingmode_speech_stop_button"))

        def _run():
            try:
                engine = pyttsx3.init()
                engine.setProperty("volume", self.speech_volume)
                self.tts_engine = engine
                engine.say(text)
                engine.runAndWait()
            except Exception:
                pass
            finally:
                self.tts_engine = None
                try:
                    self.after(0, lambda: self.speech_button.config(text=t("cookingmode_speech_button")))
                except tk.TclError:
                    pass  # la fenêtre a pu être fermée pendant la lecture

        self.tts_thread = threading.Thread(target=_run, daemon=True)
        self.tts_thread.start()

    def stop_speech(self):
        if self.tts_engine is not None:
            try:
                self.tts_engine.stop()
            except Exception:
                pass
        try:
            self.speech_button.config(text=t("cookingmode_speech_button"))
        except tk.TclError:
            pass

    def _adjust_volume(self, delta):
        self.speech_volume = round(max(0.0, min(1.0, self.speech_volume + delta)), 2)
        self.volume_label.config(text=t("cookingmode_volume_percent", percent=round(self.speech_volume * 100)))
        if self.tts_engine is not None:
            try:
                self.tts_engine.setProperty("volume", self.speech_volume)
            except Exception:
                pass

    def _on_close(self):
        self.stop_speech()
        self.destroy()

    def _toggle_fullscreen(self):
        self._is_fullscreen = not self._is_fullscreen
        try:
            self.attributes("-fullscreen", self._is_fullscreen)
        except tk.TclError:
            self._is_fullscreen = False

    @staticmethod
    def _fmt(value):
        return int(value) if value == int(value) else value

    def _adjust(self, delta):
        self.persons = max(0.5, self.persons + delta)
        self.pers_label.config(text=t("cookingmode_persons_suffix", persons=self._fmt(self.persons)))
        self._render()

    def _render(self):
        for child in self.content.winfo_children():
            child.destroy()

        recipe = self.recipe
        star = "⭐ " if recipe.get("favorite") else ""
        tk.Label(self.content, text=f"{star}{recipe['name']}", font=("Segoe UI", sf(34), "bold"),
                 bg="white", wraplength=1000, justify="center").pack(pady=(10, 5))

        info_bits = []
        if recipe.get("prep_time"):
            info_bits.append(t("cookingmode_prep_label", time=recipe['prep_time']))
        if recipe.get("cook_time"):
            info_bits.append(t("cookingmode_cook_label", time=recipe['cook_time']))
        if recipe.get("difficulty"):
            info_bits.append(t("cookingmode_difficulty_label", value=translate_difficulty_name(recipe['difficulty'])))
        if info_bits:
            tk.Label(self.content, text="   |   ".join(info_bits), font=("Segoe UI", sf(16)),
                     bg="white", fg="#555").pack(pady=(0, 20))

        tk.Label(self.content, text=t("cookingmode_ingredients_heading"), font=("Segoe UI", sf(22), "bold"),
                 bg="white").pack(pady=(10, 8), anchor="w", fill="x")
        for ing in recipe["ingredients"]:
            qty = round(ing["quantity"] * self.persons, 2)
            if qty == int(qty):
                qty = int(qty)
            unit = f" {translate_unit_name(ing['unit'])}" if ing["unit"] else ""
            tk.Label(self.content, text=f"•  {translate_ingredient_name(ing['name']).capitalize()} : {qty}{unit}",
                     font=("Segoe UI", sf(18)), bg="white", anchor="w", justify="left",
                     wraplength=1000).pack(fill="x", pady=3, anchor="w")

        description = recipe.get("description", "").strip()
        if description:
            tk.Label(self.content, text=t("cookingmode_preparation_heading"), font=("Segoe UI", sf(22), "bold"),
                     bg="white").pack(pady=(25, 8), anchor="w", fill="x")
            tk.Label(self.content, text=description, font=("Segoe UI", sf(16)), bg="white",
                     justify="left", anchor="w", wraplength=1000).pack(fill="x", anchor="w")

        personal_notes = recipe.get("personal_notes", "").strip()
        if personal_notes:
            tk.Label(self.content, text=t("cookingmode_personal_notes_heading"), font=("Segoe UI", sf(20), "bold"),
                     bg="white", fg="#555").pack(pady=(20, 8), anchor="w", fill="x")
            tk.Label(self.content, text=personal_notes, font=("Segoe UI", sf(14)), bg="white",
                     fg="#555", justify="left", anchor="w", wraplength=1000).pack(fill="x", anchor="w")

        tk.Label(self.content, text="", bg="white").pack(pady=30)  # marge basse


class IngredientSearchWindow(tk.Toplevel):
    """Recherche inversée : à partir d'un ingrédient, retrouver toutes les
    recettes qui l'utilisent."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("ingsearch_title"))
        self.geometry(f"{gs(480)}x{gs(600)}")
        self.grab_set()

        ttk.Label(self, text=t("ingsearch_question_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(10, 5))

        search_frame = ttk.Frame(self)
        search_frame.pack(fill="x", padx=15, pady=(0, 5))
        ttk.Label(search_frame, text="🔍").pack(side="left")
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate_ingredients())

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", padx=15, pady=5)
        self.ing_listbox = tk.Listbox(list_frame, height=10, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.ing_listbox.yview)
        self.ing_listbox.configure(yscrollcommand=scrollbar.set)
        self.ing_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.ing_listbox.bind("<Double-Button-1>", lambda e: self.search_recipes())
        self._populate_ingredients()

        ttk.Button(self, text=t("ingsearch_view_recipes_button"),
                   command=self.search_recipes).pack(pady=8)

        self.result_label = ttk.Label(self, text="", font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED)
        self.result_label.pack(padx=15, anchor="w")

        result_frame = ttk.Frame(self)
        result_frame.pack(pady=5, padx=15, fill="both", expand=True)
        self.result_listbox = tk.Listbox(result_frame, height=12, font=("Segoe UI", sf(9)))
        result_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_listbox.yview)
        self.result_listbox.configure(yscrollcommand=result_scrollbar.set)
        self.result_listbox.pack(side="left", fill="both", expand=True)
        result_scrollbar.pack(side="right", fill="y")
        self.result_listbox.bind("<Double-Button-1>", lambda e: self.open_selected_recipe())
        self.matched_recipe_names = []

        ttk.Button(self, text=t("ingsearch_view_selected_button"),
                   command=self.open_selected_recipe).pack(pady=(5, 10))

    def _populate_ingredients(self):
        search = self.search_entry.get().strip()
        search_key = ingredient_sort_key(search) if search else ""
        self.ing_listbox.delete(0, tk.END)
        self.displayed_ing_names = []
        for name in self.app.ingredient_names:
            if search_key and search_key not in ingredient_sort_key(name) \
                    and search_key not in ingredient_sort_key(translate_ingredient_name(name)):
                continue
            self.ing_listbox.insert(tk.END, translate_ingredient_name(name))
            self.displayed_ing_names.append(name)

    def search_recipes(self):
        sel = self.ing_listbox.curselection()
        if not sel:
            messagebox.showinfo(t("common_info"), t("manageing_select_ingredient_first"))
            return
        target_name = self.displayed_ing_names[sel[0]]
        target_key = ingredient_sort_key(target_name)

        matches = []
        for recipe in self.app.recipes:
            for ing in recipe["ingredients"]:
                if ingredient_sort_key(ing["name"]) == target_key:
                    matches.append((recipe, ing))
                    break

        self.result_listbox.delete(0, tk.END)
        self.matched_recipe_names = []
        if not matches:
            self.result_label.config(text=t("ingsearch_no_recipe_uses", name=translate_ingredient_name(target_name)))
            return

        self.result_label.config(
            text=t("ingsearch_recipes_using", name=translate_ingredient_name(target_name), count=len(matches))
        )
        for recipe, ing in matches:
            star = "⭐ " if recipe.get("favorite") else ""
            cat = translate_category_name(recipe.get("category", "Autre"))
            qty = ing["quantity"]
            if qty == int(qty):
                qty = int(qty)
            unit = f" {translate_unit_name(ing['unit'])}" if ing["unit"] else ""
            self.result_listbox.insert(
                tk.END, t("ingsearch_result_line", star=star, cat=cat, name=recipe['name'], qty=qty, unit=unit)
            )
            self.matched_recipe_names.append(recipe["name"])

    def open_selected_recipe(self):
        sel = self.result_listbox.curselection()
        if not sel:
            messagebox.showinfo(t("common_info"), t("ingsearch_select_result_first"))
            return
        recipe_name = self.matched_recipe_names[sel[0]]
        OneRecipeWindow(self.app, initial_recipe_name=recipe_name)


class TimerRow(tk.Frame):
    """Un minuteur réglable et indépendant, affiché comme une ligne à
    l'intérieur de TimersWindow. Quand il arrive à zéro, la ligne clignote
    en rouge et un signal sonore retentit jusqu'à ce que l'utilisateur
    interagisse avec elle (démarrer, réinitialiser, ou simplement cliquer
    dessus)."""

    def __init__(self, parent, timers_window, label=None, minutes=10):
        super().__init__(parent, background=COLOR_CARD, highlightbackground=COLOR_BORDER,
                          highlightthickness=1)
        if label is None:
            label = t("onerecipe_default_timer_label")
        self.timers_window = timers_window
        self.remaining_seconds = max(0, int(minutes) * 60)
        self.running = False
        self.finished = False
        self._after_id = None
        self._flash_after_id = None
        self._flash_on = False

        top_row = tk.Frame(self, background=COLOR_CARD)
        top_row.pack(fill="x", padx=8, pady=(8, 2))
        self.name_entry = ttk.Entry(top_row, width=18)
        self.name_entry.insert(0, label)
        self.name_entry.pack(side="left")
        self.display_label = tk.Label(top_row, text=self._format_time(), font=("Segoe UI", sf(18), "bold"),
                                       background=COLOR_CARD, foreground=COLOR_ACCENT_DARK, width=7)
        self.display_label.pack(side="left", padx=10)
        ttk.Button(top_row, text="🗑", width=3,
                   command=lambda: self.timers_window.remove_timer(self)).pack(side="right")

        bottom_row = tk.Frame(self, background=COLOR_CARD)
        bottom_row.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(bottom_row, text=t("timerrow_minutes_label"), style="Card.TLabel").pack(side="left")
        self.minutes_entry = ttk.Entry(bottom_row, width=4)
        self.minutes_entry.insert(0, str(minutes))
        self.minutes_entry.pack(side="left", padx=(2, 8))
        ttk.Label(bottom_row, text=t("timerrow_seconds_label"), style="Card.TLabel").pack(side="left")
        self.seconds_entry = ttk.Entry(bottom_row, width=4)
        self.seconds_entry.insert(0, "0")
        self.seconds_entry.pack(side="left", padx=(2, 8))

        self.start_button = ttk.Button(bottom_row, text="▶️", width=3, command=self.start)
        self.start_button.pack(side="left", padx=2)
        self.pause_button = ttk.Button(bottom_row, text="⏸️", width=3, command=self.pause, state="disabled")
        self.pause_button.pack(side="left", padx=2)
        ttk.Button(bottom_row, text="🔄", width=3, command=self.reset).pack(side="left", padx=2)

        # Cliquer n'importe où sur la ligne (ou sur le gros affichage du
        # temps) fait taire l'alarme si le minuteur est terminé.
        self.bind("<Button-1>", lambda e: self._stop_flash())
        self.display_label.bind("<Button-1>", lambda e: self._stop_flash())

        self._refresh_display()

    def _format_time(self):
        mins, secs = divmod(max(0, self.remaining_seconds), 60)
        return f"{mins:02d}:{secs:02d}"

    def _refresh_display(self):
        self.display_label.config(text=self._format_time())

    def start(self):
        if self.finished:
            self._stop_flash()
        if self.running:
            return
        if self.remaining_seconds <= 0:
            try:
                minutes = int(self.minutes_entry.get().strip() or 0)
                seconds = int(self.seconds_entry.get().strip() or 0)
            except ValueError:
                messagebox.showerror(t("common_error"), t("timerrow_error_invalid_duration"))
                return
            self.remaining_seconds = max(0, minutes * 60 + seconds)
            if self.remaining_seconds <= 0:
                messagebox.showinfo(t("common_info"), t("timerrow_set_duration_first"))
                return
        self.running = True
        self.minutes_entry.config(state="disabled")
        self.seconds_entry.config(state="disabled")
        self.start_button.config(state="disabled")
        self.pause_button.config(state="normal")
        self._tick()

    def _tick(self):
        if not self.running:
            return
        if self.remaining_seconds <= 0:
            self._on_finished()
            return
        self.remaining_seconds -= 1
        self._refresh_display()
        self._after_id = self.after(1000, self._tick)

    def pause(self):
        self.running = False
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        self.start_button.config(state="normal")
        self.pause_button.config(state="disabled")

    def reset(self):
        self.running = False
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        self._stop_flash()
        self.minutes_entry.config(state="normal")
        self.seconds_entry.config(state="normal")
        self.start_button.config(state="normal")
        self.pause_button.config(state="disabled")
        try:
            minutes = int(self.minutes_entry.get().strip() or 0)
            seconds = int(self.seconds_entry.get().strip() or 0)
        except ValueError:
            minutes, seconds = 0, 0
        self.remaining_seconds = max(0, minutes * 60 + seconds)
        self._refresh_display()

    def _on_finished(self):
        self.running = False
        self.finished = True
        self._after_id = None
        self.pause_button.config(state="disabled")
        self.start_button.config(state="normal")
        self.minutes_entry.config(state="normal")
        self.seconds_entry.config(state="normal")
        self._refresh_display()
        self._start_flash()

    def _set_children_bg(self, widget, color):
        for child in widget.winfo_children():
            if isinstance(child, tk.Frame):
                child.configure(background=color)
                self._set_children_bg(child, color)
            elif isinstance(child, tk.Label):
                child.configure(background=color)

    def _start_flash(self):
        self._flash_on = False
        self._flash_step()

    def _flash_step(self):
        if not self.finished:
            return
        self._flash_on = not self._flash_on
        color = COLOR_ERROR if self._flash_on else COLOR_CARD
        text_color = "white" if self._flash_on else COLOR_ACCENT_DARK
        self.configure(background=color)
        self._set_children_bg(self, color)
        self.display_label.config(background=color, foreground=text_color)
        try:
            self.bell()
        except tk.TclError:
            pass
        self._flash_after_id = self.after(500, self._flash_step)

    def _stop_flash(self):
        if not self.finished and self._flash_after_id is None:
            return
        self.finished = False
        if self._flash_after_id is not None:
            try:
                self.after_cancel(self._flash_after_id)
            except Exception:
                pass
            self._flash_after_id = None
        self.configure(background=COLOR_CARD)
        self._set_children_bg(self, COLOR_CARD)
        self.display_label.config(background=COLOR_CARD, foreground=COLOR_ACCENT_DARK)

    def cancel(self):
        """Arrête tout minuterie/clignotement en cours (appelé à la
        fermeture de la fenêtre ou à la suppression de cette ligne)."""
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
        if self._flash_after_id is not None:
            try:
                self.after_cancel(self._flash_after_id)
            except Exception:
                pass


class CookLogEntryDialog(tk.Toplevel):
    """Petite fenêtre pour ajouter, juste après avoir marqué une recette
    comme cuisinée, une note et/ou une photo optionnelles au journal de
    cuisine de cette recette."""

    def __init__(self, app, recipe_name, on_done):
        super().__init__(app)
        self.app = app
        self.on_done = on_done
        self.photo_path = None
        self.title(t("cooklogentry_title"))
        self.geometry(f"{gs(420)}x{gs(400)}")
        self.resizable(False, False)
        self.grab_set()

        ttk.Label(self, text=t("cooklogentry_heading", name=recipe_name), font=("Segoe UI", sf(12), "bold"),
                  wraplength=380, justify="center").pack(pady=(15, 2))
        ttk.Label(self, text=t("cooklogentry_intro"),
                  font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center").pack(pady=(0, 10))

        self.note_text = tk.Text(self, height=7, width=42, wrap="word", font=("Segoe UI", sf(10)))
        self.note_text.pack(padx=15, pady=(0, 10))

        photo_frame = ttk.Frame(self)
        photo_frame.pack(pady=(0, 10))
        self.photo_label = ttk.Label(photo_frame, text=t("cooklogentry_no_photo_chosen"), foreground=COLOR_TEXT_MUTED)
        self.photo_label.pack(side="left", padx=(0, 8))
        ttk.Button(photo_frame, text=t("cooklogentry_choose_photo_button"), command=self.choose_photo).pack(side="left")

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text=t("common_save_button"), command=self.save).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("cooklogentry_skip_button"), style="Secondary.TButton", command=self.skip).grid(row=0, column=1, padx=5)

        self.protocol("WM_DELETE_WINDOW", self.skip)

    def choose_photo(self):
        path = filedialog.askopenfilename(
            title=t("cooklogentry_choose_photo_title"),
            filetypes=[("Images", "*.jpg *.jpeg *.png *.webp *.gif *.bmp")]
        )
        if path:
            self.photo_path = path
            self.photo_label.config(text=os.path.basename(path), foreground=COLOR_TEXT)

    def save(self):
        note = self.note_text.get("1.0", "end-1c").strip()
        photo_filename = copy_image_to_store(self.photo_path) if self.photo_path else None
        self.on_done(note, photo_filename)
        self.destroy()

    def skip(self):
        self.on_done("", None)
        self.destroy()


class CookLogWindow(tk.Toplevel):
    """Affiche l'historique des fois où une recette a été cuisinée, avec les
    notes et photos éventuellement ajoutées à chaque fois (la plus récente
    en premier)."""

    def __init__(self, app, recipe):
        super().__init__(app)
        self.app = app
        self.recipe = recipe
        self.title(t("cooklog_title", name=recipe['name']))
        self.geometry(f"{gs(480)}x{gs(600)}")
        self.minsize(gs(400), gs(400))
        self.resizable(True, True)
        self.grab_set()

        ttk.Label(self, text=t("cooklog_heading", name=recipe['name']), font=("Segoe UI", sf(13), "bold"),
                  wraplength=440, justify="center").pack(pady=(15, 2))
        times_cooked = recipe.get("times_cooked", 0)
        ttk.Label(self, text=t("cooklog_times_cooked", count=times_cooked),
                  font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED).pack(pady=(0, 10))

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        rows_frame = ttk.Frame(canvas)
        rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=rows_frame, anchor="n")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        self._thumb_refs = []
        cook_log = list(recipe.get("cook_log", []))
        cook_log.sort(key=lambda e: e.get("date", ""), reverse=True)

        if not cook_log:
            ttk.Label(
                rows_frame,
                text=t("cooklog_no_entry"),
                foreground=COLOR_TEXT_MUTED, justify="center"
            ).pack(pady=30)

        for entry in cook_log:
            entry_card = tk.Frame(rows_frame, background=COLOR_CARD, highlightbackground=COLOR_BORDER,
                                   highlightthickness=1)
            entry_card.pack(fill="x", pady=6, padx=2)
            try:
                date_display = datetime.fromisoformat(entry["date"]).strftime("%d/%m/%Y")
            except (KeyError, ValueError, TypeError):
                date_display = entry.get("date", "?")
            ttk.Label(entry_card, text=date_display, font=("Segoe UI", sf(10), "bold"),
                      style="Card.TLabel", foreground=COLOR_ACCENT_DARK).pack(anchor="w", padx=10, pady=(8, 2))

            photo_filename = entry.get("photo")
            if photo_filename:
                thumb = load_thumbnail(photo_filename, size=(220, 160))
                if thumb is not None:
                    self._thumb_refs.append(thumb)
                    tk.Label(entry_card, image=thumb, background=COLOR_CARD).pack(padx=10, pady=4)

            note = (entry.get("note") or "").strip()
            if note:
                ttk.Label(entry_card, text=note, style="Card.TLabel", wraplength=420,
                          justify="left").pack(anchor="w", padx=10, pady=(0, 8))
            else:
                ttk.Label(entry_card, text=t("cooklog_no_note"), style="Card.TLabel",
                          foreground=COLOR_TEXT_MUTED, font=("Segoe UI", sf(8))).pack(anchor="w", padx=10, pady=(0, 8))


class TimersWindow(tk.Toplevel):
    """Fenêtre unique regroupant plusieurs minuteurs indépendants et
    réglables, pour chronométrer différentes étapes d'une recette en même
    temps (ex. un pour les pâtes, un pour la sauce...). Le bouton
    "➕ Ajouter un minuteur" empile un nouveau minuteur sous les précédents."""

    def __init__(self, app, initial_label=None, initial_minutes=10):
        super().__init__(app)
        self.app = app
        if initial_label is None:
            initial_label = t("onerecipe_default_timer_label")
        self.title(t("timers_title"))
        self.geometry(f"{gs(380)}x{gs(560)}")
        self.minsize(gs(340), gs(300))
        self.resizable(True, True)
        # Reste visible au premier plan même par-dessus une autre fenêtre
        # maximisée (ex. le mode cuisine) : on veut toujours voir les
        # minuteurs en cours, quoi qu'on affiche par ailleurs.
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            pass

        ttk.Label(self, text=t("timers_title"), font=("Segoe UI", sf(13), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self, text=t("timers_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 8))

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10)
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.rows_frame = ttk.Frame(canvas)
        self.rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.rows_frame, anchor="n")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.timer_rows = []
        self.add_timer(initial_label, initial_minutes)

        ttk.Button(self, text=t("timers_add_button"),
                   command=lambda: self.add_timer()).pack(pady=10)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def add_timer(self, label=None, minutes=10):
        if label is None:
            label = t("onerecipe_default_timer_label")
        row = TimerRow(self.rows_frame, self, label, minutes)
        row.pack(fill="x", pady=6, padx=4)
        self.timer_rows.append(row)
        self.update_idletasks()

    def remove_timer(self, row):
        row.cancel()
        row.destroy()
        if row in self.timer_rows:
            self.timer_rows.remove(row)

    def _on_close(self):
        for row in list(self.timer_rows):
            row.cancel()
        if getattr(self.app, "timers_window", None) is self:
            self.app.timers_window = None
        self.destroy()


class QRCodeWindow(tk.Toplevel):
    """Affiche une recette (nom + ingrédients) sous forme de QR code, à
    scanner avec un téléphone, et permet de l'enregistrer en image PNG."""

    MAX_CHARS = 800  # limite la taille du texte encodé pour rester scannable

    def __init__(self, app, recipe, persons):
        super().__init__(app)
        self.app = app
        self.recipe = recipe
        self.title(t("qrcode_title", name=recipe['name']))
        self.geometry(f"{gs(420)}x{gs(540)}")
        self.resizable(False, False)
        self.grab_set()

        text = self._build_text(recipe, persons)

        qr = qrcode.QRCode(border=2)
        qr.add_data(text)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        qr_img = qr_img.resize((340, 340))
        self._qr_img = qr_img
        self._photo = ImageTk.PhotoImage(qr_img)

        ttk.Label(self, text=t("qrcode_title", name=recipe['name']),
                  font=("Segoe UI", sf(12), "bold"), wraplength=380, justify="center").pack(pady=10)
        ttk.Label(self, image=self._photo).pack(pady=5)
        ttk.Label(
            self,
            text=t("qrcode_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=5)

        ttk.Button(self, text=t("qrcode_save_button"),
                   command=self.save_image).pack(pady=10)

        if len(text) >= self.MAX_CHARS:
            ttk.Label(
                self,
                text=t("qrcode_truncated_warning"),
                font=("Segoe UI", sf(8)), foreground=COLOR_ERROR, justify="center"
            ).pack(pady=(0, 10))

    @classmethod
    def _build_text(cls, recipe, persons):
        lines = [recipe["name"], "", t("qrcode_encoded_ingredients_heading", persons=persons)]
        for ing in recipe["ingredients"]:
            qty = round(ing["quantity"] * persons, 2)
            if qty == int(qty):
                qty = int(qty)
            unit = f" {translate_unit_name(ing['unit'])}" if ing["unit"] else ""
            lines.append(f"- {translate_ingredient_name(ing['name']).capitalize()} : {qty}{unit}")
        text = "\n".join(lines)
        if len(text) > cls.MAX_CHARS:
            text = text[: cls.MAX_CHARS - 3] + "..."
        return text

    def save_image(self):
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", self.recipe["name"])
        path = filedialog.asksaveasfilename(
            title=t("qrcode_save_dialog_title"),
            defaultextension=".png",
            filetypes=[("Image PNG", "*.png")],
            initialfile=f"qrcode_{safe_name}.png"
        )
        if not path:
            return
        try:
            self._qr_img.save(path)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("qrcode_save_failed", error=e))
            return
        messagebox.showinfo(t("allrecipes_list_saved_title"), t("qrcode_saved_message", path=path))


CONVERTER_UNIT_KEYS = [
    ("unitconv_gram", 1.0),
    ("unitconv_kilogram", 1000.0),
    ("unitconv_ounce", 28.35),
    ("unitconv_pound", 453.6),
    ("unitconv_milliliter", 1.0),
    ("unitconv_centiliter", 10.0),
    ("unitconv_liter", 1000.0),
    ("unitconv_teaspoon", 5.0),
    ("unitconv_tablespoon", 15.0),
    ("unitconv_cup", 240.0),
]


class UnitConverterWindow(tk.Toplevel):
    """Petit outil de conversion d'unités indépendant de toute recette,
    pratique pour une recette trouvée ailleurs (ex. en tasses/onces)."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("unitconv_title"))
        self.geometry(f"{gs(440)}x{gs(460)}")
        self.minsize(gs(400), gs(400))
        self.resizable(True, True)
        self.grab_set()

        # Construit les unités traduites à l'exécution (dans la langue
        # actuellement sélectionnée), plutôt qu'un dictionnaire figé en
        # français au chargement du module.
        self.converter_units = {t(key): factor for key, factor in CONVERTER_UNIT_KEYS}

        ttk.Label(self, text=t("unitconv_heading"), font=("Segoe UI", sf(14), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self,
            text=t("unitconv_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 15))

        form = ttk.Frame(self)
        form.pack(pady=5)
        ttk.Label(form, text=t("unitconv_quantity_label")).grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.qty_entry = ttk.Entry(form, width=10)
        self.qty_entry.insert(0, "1")
        self.qty_entry.grid(row=0, column=1, padx=5, pady=5)

        unit_names = list(self.converter_units.keys())
        ttk.Label(form, text=t("unitconv_from_label")).grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.from_combo = ttk.Combobox(form, values=unit_names, state="readonly", width=22)
        self.from_combo.set(unit_names[0])
        self.from_combo.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(form, text=t("unitconv_to_label")).grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.to_combo = ttk.Combobox(form, values=unit_names, state="readonly", width=22)
        self.to_combo.set(unit_names[1])
        self.to_combo.grid(row=2, column=1, padx=5, pady=5)

        ttk.Button(self, text=t("unitconv_convert_button"), command=self.convert).pack(pady=10)
        self.result_label = ttk.Label(self, text="", font=("Segoe UI", sf(12), "bold"),
                                       foreground=COLOR_ACCENT_DARK)
        self.result_label.pack(pady=5)

        self.qty_entry.bind("<Return>", lambda e: self.convert())

    def convert(self):
        try:
            quantity = float(self.qty_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("unitconv_error_invalid_quantity"))
            return
        from_unit = self.from_combo.get()
        to_unit = self.to_combo.get()
        grams_equivalent = quantity * self.converter_units[from_unit]
        result = grams_equivalent / self.converter_units[to_unit]
        result_display = round(result, 3)
        if result_display == int(result_display):
            result_display = int(result_display)
        self.result_label.config(
            text=t("unitconv_result", quantity=quantity, from_unit=from_unit, result=result_display, to_unit=to_unit)
        )


class PantryWindow(tk.Toplevel):
    """Fenêtre de suivi du garde-manger : indiquez ce que vous avez chez vous
    et en quelle quantité, pour que « Que puis-je cuisiner ? » puisse vérifier
    non seulement la présence d'un ingrédient mais aussi si vous en avez
    assez, et pour pouvoir décompter automatiquement le stock après avoir
    cuisiné une recette."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("pantry_title"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(680)}x{min(screen_height, gs(860))}+40+20")
        self.minsize(gs(560), gs(500))
        self.resizable(True, True)
        self.grab_set()

        ttk.Label(self, text=t("pantry_heading"), font=("Segoe UI", sf(14), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self, text=t("pantry_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

        add_frame = ttk.Frame(self)
        add_frame.pack(pady=5, padx=15, fill="x")
        ttk.Label(add_frame, text=t("common_ingredient_label")).grid(row=0, column=0, padx=5, sticky="e")
        self.name_entry = ttk.Entry(add_frame, width=24)
        self.name_entry.full_values = get_display_ingredient_values(sorted(self.app.ingredient_names, key=ingredient_sort_key))
        self.name_entry.grid(row=0, column=1, padx=5)
        self.name_entry.bind("<KeyRelease>", lambda e: self._on_name_entry_keyrelease(e))
        self.name_entry.bind("<FocusIn>", lambda e: self._on_name_entry_focus_in(e))
        self.name_entry.bind("<FocusOut>", lambda e: self._on_name_entry_focus_out(e))
        ttk.Label(add_frame, text=t("common_quantity_label")).grid(row=0, column=2, padx=5)
        self.qty_entry = ttk.Entry(add_frame, width=7)
        self.qty_entry.insert(0, "1")
        self.qty_entry.grid(row=0, column=3, padx=5)
        self.unit_options = RecipeFormWindow.UNIT_OPTIONS[:-1] + ["boîte", "paquet", "rouleau", "bouteille"]
        self.unit_combo = ttk.Combobox(add_frame, values=[translate_unit_name(u) for u in self.unit_options], width=13)
        self.unit_combo.set(translate_unit_name("pièce"))
        self.unit_combo.grid(row=0, column=4, padx=5)
        ttk.Label(add_frame, text=t("pantry_threshold_label")).grid(row=1, column=0, padx=5, pady=(6, 0), sticky="e")
        self.threshold_entry = ttk.Entry(add_frame, width=7)
        self.threshold_entry.grid(row=1, column=1, padx=5, pady=(6, 0), sticky="w")
        ttk.Button(add_frame, text=t("common_save_button"), command=self.save_item).grid(
            row=1, column=2, padx=5, pady=(6, 0))
        ttk.Button(add_frame, text=t("common_new_ingredient_button"),
                   command=self.create_new_ingredient).grid(row=1, column=3, columnspan=2, padx=5, pady=(6, 0))

        ttk.Label(
            self, text=t("pantry_help_text"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 5))

        list_frame = ttk.Frame(self)
        list_frame.pack(pady=10, padx=15, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._load_selected_for_edit())

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=(0, 15))
        ttk.Button(btn_frame, text=t("pantry_remove_button"),
                   command=self.remove_selected).grid(row=0, column=0, padx=5)

        tk.Frame(self, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

        self.pantry_entries_ordered = []
        self._populate()

    def _populate(self):
        self.listbox.delete(0, tk.END)
        pantry = load_pantry()
        self.pantry_entries_ordered = sorted(pantry.values(), key=lambda e: ingredient_sort_key(e["name"]))
        if not self.pantry_entries_ordered:
            self.listbox.insert(tk.END, t("pantry_empty"))
            return
        for entry in self.pantry_entries_ordered:
            unit_display = f" {entry['unit']}" if entry["unit"] else ""
            qty = entry["quantity"]
            qty_display = int(qty) if qty == int(qty) else round(qty, 2)
            threshold = entry.get("threshold")
            low_stock = threshold is not None and qty < threshold
            prefix = "⚠️ " if low_stock else ""
            suffix = t("pantry_threshold_suffix", threshold=threshold) if threshold is not None else ""
            self.listbox.insert(tk.END, f"{prefix}{translate_ingredient_name(entry['name'])} : {qty_display}{unit_display}{suffix}")

    def create_new_ingredient(self):
        typed = normalize_oe(self.name_entry.get().strip())
        win = IngredientEditWindow(self.app, manage_window=None, existing_name=None,
                                    prefill_name=typed, parent_window=self)
        self.wait_window(win)
        self.name_entry.full_values = get_display_ingredient_values(sorted(self.app.ingredient_names, key=ingredient_sort_key))

    def _prefill_existing(self):
        typed = self.name_entry.get().strip()
        canonical = resolve_ingredient_input(typed, self.app.ingredient_names)
        if canonical is None:
            return
        pantry = load_pantry()
        entry = pantry.get(ingredient_sort_key(canonical))
        if entry:
            self.qty_entry.delete(0, tk.END)
            self.qty_entry.insert(0, str(entry["quantity"]))
            self.unit_combo.set(translate_unit_name(entry["unit"]))
            self.threshold_entry.delete(0, tk.END)
            if entry.get("threshold") is not None:
                self.threshold_entry.insert(0, str(entry["threshold"]))

    def _load_selected_for_edit(self):
        sel = self.listbox.curselection()
        if not sel or not self.pantry_entries_ordered:
            return
        entry = self.pantry_entries_ordered[sel[0]]
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, translate_ingredient_name(entry["name"]))
        self.qty_entry.delete(0, tk.END)
        self.qty_entry.insert(0, str(entry["quantity"]))
        self.unit_combo.set(translate_unit_name(entry["unit"]))
        self.threshold_entry.delete(0, tk.END)
        if entry.get("threshold") is not None:
            self.threshold_entry.insert(0, str(entry["threshold"]))

    # ---- Autocomplétion du champ ingrédient (même principe que les autres
    # listes déroulantes d'ingrédients de l'application) ----

    def _hide_name_suggestions(self):
        popup = getattr(self.name_entry, "_suggestion_popup", None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
            self.name_entry._suggestion_popup = None
            self.name_entry._suggestion_listbox = None

    def _show_name_suggestions(self, filtered):
        self._hide_name_suggestions()
        if not filtered:
            return
        entry = self.name_entry
        popup = tk.Toplevel(entry)
        popup.wm_overrideredirect(True)
        try:
            popup.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        width = max(entry.winfo_width(), 160)
        height = min(6, len(filtered)) * 20
        popup.wm_geometry(f"{width}x{height}+{x}+{y}")

        listbox = tk.Listbox(popup, height=min(6, len(filtered)), exportselection=False, font=("Segoe UI", sf(9)))
        listbox.pack(fill="both", expand=True)
        for v in filtered:
            listbox.insert(tk.END, v)

        def choose(event=None):
            sel = listbox.curselection()
            if sel:
                value = listbox.get(sel[0])
                entry.delete(0, tk.END)
                entry.insert(0, value)
                self._prefill_existing()
            self._hide_name_suggestions()
            entry.focus_set()

        listbox.bind("<ButtonRelease-1>", choose)
        listbox.bind("<Return>", choose)
        entry._suggestion_popup = popup
        entry._suggestion_listbox = listbox

    def _on_name_entry_keyrelease(self, event):
        if event.keysym == "Down":
            listbox = getattr(self.name_entry, "_suggestion_listbox", None)
            if listbox is not None:
                listbox.focus_set()
                listbox.selection_set(0)
            return
        if event.keysym == "Escape":
            self._hide_name_suggestions()
            return
        if event.keysym == "Return":
            self._hide_name_suggestions()
            self.save_item()
            return
        if event.keysym in ("Tab", "Shift_L", "Shift_R", "Control_L", "Control_R",
                            "Caps_Lock", "Alt_L", "Alt_R", "Left", "Right"):
            return
        filtered = self._filter_ingredient_values(self.name_entry.full_values, self.name_entry.get())
        if filtered:
            self._show_name_suggestions(filtered)
        else:
            self._hide_name_suggestions()

    def _on_name_entry_focus_in(self, event):
        filtered = self._filter_ingredient_values(self.name_entry.full_values, self.name_entry.get())
        if filtered:
            self._show_name_suggestions(filtered)

    def _on_name_entry_focus_out(self, event):
        self.name_entry.after(200, self._hide_name_suggestions)

    @staticmethod
    def _filter_ingredient_values(full_values, typed):
        if not typed:
            return full_values
        typed_key = ingredient_sort_key(typed)
        filtered = [v for v in full_values if ingredient_sort_key(v).startswith(typed_key)]
        if not filtered:
            filtered = [v for v in full_values if typed_key in ingredient_sort_key(v)]
        return filtered

    def save_item(self):
        name = normalize_oe(self.name_entry.get().strip())
        if not name:
            messagebox.showerror(t("common_error"), t("pantry_error_ingredient_required"))
            return
        canonical = resolve_ingredient_input(name, self.app.ingredient_names)
        if canonical is None:
            messagebox.showerror(
                t("common_unknown_ingredient_title"),
                t("common_unknown_ingredient_simple_message", name=name)
            )
            return
        try:
            quantity = float(self.qty_entry.get().strip().replace(",", "."))
            if quantity < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(t("common_error"), t("pantry_error_invalid_quantity"))
            return
        threshold_str = self.threshold_entry.get().strip()
        threshold = None
        if threshold_str:
            try:
                threshold = float(threshold_str.replace(",", "."))
                if threshold < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(t("common_error"), t("pantry_error_invalid_threshold"))
                return
        unit = resolve_unit_input_best_effort(self.unit_combo.get().strip(), self.unit_options)
        set_pantry_item(canonical, quantity, unit, threshold)
        self._populate()
        self.name_entry.delete(0, tk.END)
        self.qty_entry.delete(0, tk.END)
        self.qty_entry.insert(0, "1")
        self.threshold_entry.delete(0, tk.END)

    def remove_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.pantry_entries_ordered:
            messagebox.showinfo(t("common_info"), t("pantry_select_ingredient_first"))
            return
        entry = self.pantry_entries_ordered[sel[0]]
        if not messagebox.askyesno(t("common_confirm"), t("pantry_remove_confirm_message", name=entry['name'])):
            return
        remove_pantry_item(entry["name"])
        self._populate()


class WhatCanICookWindow(tk.Toplevel):
    """Fenêtre pour indiquer les ingrédients qu'on a sous la main, et voir
    quelles recettes sont réalisables (ou presque)."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("cook_title"))
        self.geometry(f"{gs(640)}x{gs(660)}")
        self.grab_set()
        # Pré-coche les ingrédients de base qu'on a presque toujours sous la
        # main (correspondance exacte, insensible à la casse, avec la liste
        # d'ingrédients de l'utilisateur — un ingrédient de base absent de sa
        # liste est simplement ignoré ici).
        available_lower = {n.lower(): n for n in self.app.ingredient_names}
        self.have_names = [
            available_lower[staple.lower()] for staple in PANTRY_STAPLES
            if staple.lower() in available_lower
        ]

        ttk.Label(self, text=t("cook_instructions_label"),
                  font=("Segoe UI", sf(11), "bold")).pack(pady=(10, 2))
        ttk.Label(
            self,
            text=t("cook_staples_hint"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 5))

        columns = ttk.Frame(self)
        columns.pack(fill="both", expand=False, padx=15, pady=5)

        left = ttk.Frame(columns)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        ttk.Label(left, text=t("cook_all_ingredients_label")).pack()
        search_frame = ttk.Frame(left)
        search_frame.pack(fill="x", pady=3)
        ttk.Label(search_frame, text="🔍").pack(side="left")
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=3)
        self.search_entry.bind("<KeyRelease>", lambda e: self._populate_all())
        self.all_listbox = tk.Listbox(left, height=14, font=("Segoe UI", sf(9)))
        self.all_listbox.pack(fill="both", expand=True)
        self.all_listbox.bind("<Double-Button-1>", lambda e: self._add_selected())
        ttk.Button(left, text=t("cook_add_button"), command=self._add_selected).pack(pady=5)

        right = ttk.Frame(columns)
        right.pack(side="left", fill="both", expand=True)
        ttk.Label(right, text=t("cook_have_label")).pack()
        self.have_listbox = tk.Listbox(right, height=16, font=("Segoe UI", sf(9)))
        self.have_listbox.pack(fill="both", expand=True, pady=(3, 0))
        self.have_listbox.bind("<Double-Button-1>", lambda e: self._remove_selected())
        ttk.Button(right, text=t("cook_remove_button"), command=self._remove_selected).pack(pady=5)
        ttk.Button(right, text=t("cook_load_from_pantry_button"),
                   command=self._load_from_pantry).pack(pady=(0, 5))

        self._populate_all()
        self._populate_have()

        ttk.Button(self, text=t("cook_compute_button"),
                   command=self.compute_feasible).pack(pady=8)

        result_frame = ttk.Frame(self)
        result_frame.pack(pady=5, padx=15, fill="both", expand=True)
        self.result_listbox = tk.Listbox(result_frame, height=14, font=("Segoe UI", sf(9)))
        result_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_listbox.yview)
        self.result_listbox.configure(yscrollcommand=result_scrollbar.set)
        self.result_listbox.pack(side="left", fill="both", expand=True)
        result_scrollbar.pack(side="right", fill="y")
        self.result_listbox.bind("<Double-Button-1>", lambda e: self.open_selected_recipe())
        self.feasible_recipe_names = []  # correspondance ligne -> nom de recette (None pour les en-têtes)

        ttk.Button(self, text=t("cook_open_selected_button"),
                   command=self.open_selected_recipe).pack(pady=(5, 10))

    def _populate_all(self):
        search = self.search_entry.get().strip()
        search_key = ingredient_sort_key(search) if search else ""
        self.all_listbox.delete(0, tk.END)
        self.displayed_all_names = []
        for name in self.app.ingredient_names:
            if name in self.have_names:
                continue
            if search_key and search_key not in ingredient_sort_key(name) \
                    and search_key not in ingredient_sort_key(translate_ingredient_name(name)):
                continue
            self.all_listbox.insert(tk.END, translate_ingredient_name(name))
            self.displayed_all_names.append(name)

    def _populate_have(self):
        self.have_listbox.delete(0, tk.END)
        for name in self.have_names:
            self.have_listbox.insert(tk.END, translate_ingredient_name(name))

    def _add_selected(self):
        sel = self.all_listbox.curselection()
        for i in sel:
            name = self.displayed_all_names[i]
            if name not in self.have_names:
                self.have_names.append(name)
        self._populate_have()
        self._populate_all()

    def _remove_selected(self):
        sel = self.have_listbox.curselection()
        names_to_remove = {self.have_names[i] for i in sel}
        self.have_names = [n for n in self.have_names if n not in names_to_remove]
        self._populate_have()
        self._populate_all()

    def _load_from_pantry(self):
        pantry = load_pantry()
        if not pantry:
            messagebox.showinfo(t("cook_pantry_empty_title"), t("cook_pantry_empty_message"))
            return
        added = 0
        for entry in pantry.values():
            if entry["name"] not in self.have_names:
                self.have_names.append(entry["name"])
                added += 1
        self._populate_have()
        self._populate_all()
        messagebox.showinfo(t("cook_loaded_title"), t("cook_loaded_message", count=added))

    def compute_feasible(self):
        have_keys = {ingredient_sort_key(n) for n in self.have_names}
        if not have_keys:
            messagebox.showinfo(t("common_info"), t("cook_add_ingredient_first"))
            return

        results = []
        for recipe in self.app.recipes:
            seen = set()
            missing = []
            for ing in recipe["ingredients"]:
                key = ingredient_sort_key(ing["name"])
                if key not in have_keys and key not in seen:
                    seen.add(key)
                    missing.append(ing["name"])
            results.append((recipe, missing))

        feasible = [r for r in results if not r[1]]
        remaining = [r for r in results if r[1] and len(r[1]) <= 3]

        # Pour les recettes avec 1 à 3 ingrédients manquants, vérifie si un
        # substitut connu pour l'ingrédient manquant est déjà dans "Ce que
        # j'ai" : si TOUS les ingrédients manquants ont un substitut
        # disponible, la recette devient réalisable avec substitution.
        substitutable = []
        almost = []
        for recipe, missing in remaining:
            subs_used = {}
            all_covered = True
            for ing_name in missing:
                sub_name = None
                for sub in get_ingredient_substitutions(ing_name):
                    if ingredient_sort_key(sub["nom"]) in have_keys:
                        sub_name = sub["nom"]
                        break
                if sub_name:
                    subs_used[ing_name] = sub_name
                else:
                    all_covered = False
            if all_covered:
                substitutable.append((recipe, missing, subs_used))
            else:
                almost.append((recipe, missing))

        self.result_listbox.delete(0, tk.END)
        self.feasible_recipe_names = []

        def add_header(text):
            self.result_listbox.insert(tk.END, text)
            self.feasible_recipe_names.append(None)

        if feasible:
            add_header(t("cook_feasible_header"))
            pantry = load_pantry()
            for recipe, missing in sorted(feasible, key=lambda pair: ingredient_sort_key(pair[0]["name"])):
                star = "⭐ " if recipe.get("favorite") else ""
                warning = ""
                if pantry:
                    insufficient = [
                        translate_ingredient_name(ing["name"]).capitalize() for ing in recipe["ingredients"]
                        if pantry_stock_status(ing["name"], ing["quantity"], ing["unit"], pantry) == "insuffisant"
                    ]
                    if insufficient:
                        warning = t("cook_insufficient_quantity", list=", ".join(insufficient))
                self.result_listbox.insert(tk.END, f"   {star}{recipe['name']}{warning}")
                self.feasible_recipe_names.append(recipe["name"])
        else:
            add_header(t("cook_none_feasible"))

        if substitutable:
            add_header("")
            add_header(t("cook_substitutable_header"))
            for recipe, missing, subs_used in sorted(substitutable, key=lambda pair: ingredient_sort_key(pair[0]["name"])):
                details = ", ".join(
                    f"{translate_ingredient_name(m).capitalize()} → {translate_ingredient_name(subs_used[m])}"
                    for m in missing
                )
                self.result_listbox.insert(tk.END, f"   {recipe['name']} ({details})")
                self.feasible_recipe_names.append(recipe["name"])

        if almost:
            add_header("")
            add_header(t("cook_almost_header"))
            for recipe, missing in sorted(almost, key=lambda pair: (len(pair[1]), ingredient_sort_key(pair[0]["name"]))):
                missing_display = ", ".join(translate_ingredient_name(m).capitalize() for m in missing)
                self.result_listbox.insert(tk.END, t("cook_missing_label", name=recipe['name'], list=missing_display))
                self.feasible_recipe_names.append(recipe["name"])

        if not feasible and not substitutable and not almost:
            add_header(t("cook_no_results"))

    def open_selected_recipe(self):
        sel = self.result_listbox.curselection()
        if not sel:
            messagebox.showinfo(t("common_info"), t("cook_select_recipe_from_results"))
            return
        recipe_name = self.feasible_recipe_names[sel[0]]
        if recipe_name is None:
            messagebox.showinfo(t("common_info"), t("cook_select_recipe_row"))
            return
        OneRecipeWindow(self.app, initial_recipe_name=recipe_name)


class WeeklyPlanHistoryWindow(tk.Toplevel):
    """Fenêtre pour consulter les plannings de semaines passées (jusqu'à 26
    semaines d'historique), et éventuellement en recharger un dans le
    planning actuel — pratique pour éviter de refaire deux fois la même
    chose de trop près."""

    def __init__(self, app, parent_window):
        super().__init__(parent_window)
        self.app = app
        self.parent_window = parent_window
        self.title(t("weekhistory_title"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(620)}x{min(screen_height, gs(760))}+40+20")
        self.minsize(gs(500), gs(500))
        self.resizable(True, True)
        self.grab_set()

        ttk.Label(self, text=t("weekhistory_heading"),
                  font=("Segoe UI", sf(14), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self, text=t("weekhistory_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

        list_frame = ttk.Frame(self)
        list_frame.pack(padx=15, pady=5, fill="both", expand=True)
        self.week_listbox = tk.Listbox(list_frame, width=16, font=("Segoe UI", sf(9)))
        week_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.week_listbox.yview)
        self.week_listbox.configure(yscrollcommand=week_scrollbar.set)
        self.week_listbox.pack(side="left", fill="y")
        week_scrollbar.pack(side="left", fill="y")
        self.week_listbox.bind("<<ListboxSelect>>", lambda e: self._show_week_detail())

        self.detail_text = tk.Text(list_frame, wrap="word", width=40, height=18, font=("Segoe UI", sf(9)))
        self.detail_text.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self.history = []
        self._populate()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("weekhistory_reload_button"),
                   command=self.reload_selected).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("weekhistory_delete_button"),
                   command=self.delete_selected).grid(row=0, column=1, padx=5)

        tk.Frame(self, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _populate(self):
        self.week_listbox.delete(0, tk.END)
        self.detail_text.delete("1.0", tk.END)
        self.history = sorted(load_weekly_plan_history(), key=lambda h: h.get("week_start", ""), reverse=True)
        if not self.history:
            self.week_listbox.insert(tk.END, t("weekhistory_no_archived_weeks"))
            return
        for entry in self.history:
            self.week_listbox.insert(tk.END, t("weekhistory_week_label", week=entry.get('week_start', '?')))

    def _show_week_detail(self):
        self.detail_text.delete("1.0", tk.END)
        sel = self.week_listbox.curselection()
        if not sel or not self.history or sel[0] >= len(self.history):
            return
        entry = self.history[sel[0]]
        plan = entry.get("plan", {})
        self.detail_text.insert(tk.END, t("weekhistory_saved_on", date=entry.get('saved_at', '?')))
        has_content = False
        for day in WEEKDAYS:
            day_data = plan.get(day) or {}
            filled_slots = {slot: info for slot, info in day_data.items() if info}
            if not filled_slots:
                continue
            has_content = True
            self.detail_text.insert(tk.END, t("weekhistory_day_heading", day=translate_weekday_name(day)))
            for slot, info in filled_slots.items():
                self.detail_text.insert(
                    tk.END,
                    t("weekhistory_slot_line", slot=translate_mealslot_name(slot), recipe=info.get('recipe_name', '?'),
                      persons=info.get('persons', '?'))
                )
            self.detail_text.insert(tk.END, "\n")
        if not has_content:
            self.detail_text.insert(tk.END, t("weekhistory_empty_week"))

    def reload_selected(self):
        sel = self.week_listbox.curselection()
        if not sel or not self.history or sel[0] >= len(self.history):
            messagebox.showinfo(t("common_info"), t("weekhistory_select_week_first"))
            return
        entry = self.history[sel[0]]
        if not messagebox.askyesno(
            t("common_confirm"),
            t("weekhistory_reload_confirm_message", week=entry.get('week_start', '?'))
        ):
            return
        self.parent_window.apply_plan(entry.get("plan", {}))
        self.destroy()

    def delete_selected(self):
        sel = self.week_listbox.curselection()
        if not sel or not self.history or sel[0] >= len(self.history):
            messagebox.showinfo(t("common_info"), t("weekhistory_select_week_first"))
            return
        entry = self.history[sel[0]]
        week_key = entry.get("week_start")
        if not messagebox.askyesno(
            t("common_confirm"), t("weekhistory_delete_confirm_message", week=week_key)
        ):
            return
        all_history = load_weekly_plan_history()
        all_history = [h for h in all_history if h.get("week_start") != week_key]
        save_weekly_plan_history(all_history)
        self._populate()


class WeeklyPlanTemplatesWindow(tk.Toplevel):
    """Fenêtre pour enregistrer le planning actuel comme modèle réutilisable
    (ex. « Semaine légère », « Semaine végétarienne »), et pour appliquer
    d'un clic un modèle déjà enregistré à un planning vide."""

    def __init__(self, app, parent_window):
        super().__init__(parent_window)
        self.app = app
        self.parent_window = parent_window
        self.title(t("weektemplates_title"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(560)}x{min(screen_height, gs(700))}+40+20")
        self.minsize(gs(460), gs(460))
        self.resizable(True, True)
        self.grab_set()

        ttk.Label(self, text=t("weektemplates_heading"), font=("Segoe UI", sf(14), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self, text=t("weektemplates_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center"
        ).pack(pady=(0, 10))

        save_frame = ttk.Frame(self)
        save_frame.pack(pady=(0, 10), padx=15, fill="x")
        ttk.Label(save_frame, text=t("weektemplates_name_label")).pack(side="left")
        self.new_name_entry = ttk.Entry(save_frame)
        self.new_name_entry.pack(side="left", padx=5, fill="x", expand=True)
        ttk.Button(save_frame, text=t("weektemplates_save_button"),
                   command=self.save_as_template).pack(side="left", padx=(5, 0))

        list_frame = ttk.Frame(self)
        list_frame.pack(padx=15, pady=5, fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, font=("Segoe UI", sf(9)))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<Double-Button-1>", lambda e: self.apply_selected())

        self.template_names = []
        self._populate()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("weektemplates_apply_button"),
                   command=self.apply_selected).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("weektemplates_delete_button"),
                   command=self.delete_selected).grid(row=0, column=1, padx=5)

        tk.Frame(self, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _populate(self):
        self.listbox.delete(0, tk.END)
        templates = load_weekly_plan_templates()
        self.template_names = sorted(templates.keys(), key=ingredient_sort_key)
        if not self.template_names:
            self.listbox.insert(tk.END, t("weektemplates_none_saved"))
            return
        for name in self.template_names:
            self.listbox.insert(tk.END, name)

    def save_as_template(self):
        name = self.new_name_entry.get().strip()
        if not name:
            messagebox.showerror(t("common_error"), t("weektemplates_error_name_required"))
            return
        plan, _ = self.parent_window._collect_selection()
        if plan is None:
            return
        if not any(plan.values()):
            messagebox.showinfo(t("common_info"), t("weektemplates_empty_plan"))
            return
        templates = load_weekly_plan_templates()
        templates[name] = plan
        save_weekly_plan_templates(templates)
        self.new_name_entry.delete(0, tk.END)
        self._populate()
        messagebox.showinfo(t("allrecipes_list_saved_title"), t("weektemplates_saved_message", name=name))

    def apply_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.template_names or sel[0] >= len(self.template_names):
            messagebox.showinfo(t("common_info"), t("weektemplates_select_template_first"))
            return
        name = self.template_names[sel[0]]
        if not messagebox.askyesno(
            t("common_confirm"),
            t("weektemplates_apply_confirm_message", name=name)
        ):
            return
        templates = load_weekly_plan_templates()
        self.parent_window.apply_plan(templates.get(name, {}))
        self.destroy()

    def delete_selected(self):
        sel = self.listbox.curselection()
        if not sel or not self.template_names or sel[0] >= len(self.template_names):
            messagebox.showinfo(t("common_info"), t("weektemplates_select_template_first"))
            return
        name = self.template_names[sel[0]]
        if not messagebox.askyesno(t("common_confirm"), t("weektemplates_delete_confirm_message", name=name)):
            return
        templates = load_weekly_plan_templates()
        templates.pop(name, None)
        save_weekly_plan_templates(templates)
        self._populate()


class WeeklyPlanWindow(tk.Toplevel):
    """Planning des repas de la semaine (petit-déjeuner, déjeuner en 3 temps,
    dîner en 3 temps), avec génération automatique de la liste de courses
    pour l'ensemble des recettes planifiées."""

    MEAL_SLOTS = [
        "Petit-déjeuner",
        "Déjeuner — Entrée", "Déjeuner — Plat", "Déjeuner — Dessert",
        "Dîner — Entrée", "Dîner — Plat", "Dîner — Dessert",
    ]

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("weekplan_title"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(1080)}x{screen_height}+40+0")
        self.minsize(gs(600), gs(400))
        self.resizable(True, True)
        self.grab_set()

        self.manual_items = []  # ingrédients ajoutés manuellement (hors planning) : [{"name","quantity","unit"}]
        self.plan = load_weekly_plan()  # {jour: {créneau: {'recipe_name':.., 'persons':..}}}
        recipe_names = [r["name"] for r in self.app.recipes]

        ttk.Label(self, text=t("weekplan_title"), font=("Segoe UI", sf(14), "bold")).pack(pady=10)
        ttk.Label(self, text=t("weekplan_subtitle"),
                  font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED).pack()

        # Largeurs de colonnes forcées identiquement dans l'en-tête fixe et
        # dans la grille défilante en dessous, pour qu'elles restent alignées
        # verticalement quel que soit le contenu de chaque cellule.
        COL0_WIDTH = 140
        DAY_COL_WIDTH = 130
        # Largeur approximative de l'ascenseur vertical de la grille, pour
        # compenser côté en-tête et garder les colonnes de jours bien alignées
        # avec celles de la grille (qui dispose d'un peu moins de largeur
        # utile à cause de cet ascenseur).
        SCROLLBAR_WIDTH_ESTIMATE = 18

        # ---- En-tête des jours de la semaine, fixe : reste toujours visible
        # à l'écran, même en faisant défiler la grille vers le bas. ----
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=(10, 10 + SCROLLBAR_WIDTH_ESTIMATE))
        header_frame.grid_columnconfigure(0, minsize=COL0_WIDTH)
        ttk.Label(header_frame, text="", width=17).grid(row=0, column=0, padx=2, pady=2)
        for col, day in enumerate(WEEKDAYS, start=1):
            header_frame.grid_columnconfigure(col, minsize=DAY_COL_WIDTH)
            ttk.Label(header_frame, text=translate_weekday_name(day), font=("Segoe UI", sf(9), "bold"),
                      foreground=COLOR_ACCENT_DARK, anchor="center").grid(
                row=0, column=col, padx=3, pady=(2, 6), sticky="ew")
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10)

        grid_container = ttk.Frame(self)
        grid_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        h_scrollbar = ttk.Scrollbar(grid_container, orient="horizontal")
        v_scrollbar = ttk.Scrollbar(grid_container, orient="vertical")
        canvas = tk.Canvas(grid_container, highlightthickness=0,
                            xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        h_scrollbar.config(command=canvas.xview)
        v_scrollbar.config(command=canvas.yview)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        calendar_frame = ttk.Frame(canvas)
        calendar_frame.grid_columnconfigure(0, minsize=COL0_WIDTH)
        for col in range(1, len(WEEKDAYS) + 1):
            calendar_frame.grid_columnconfigure(col, minsize=DAY_COL_WIDTH)
        calendar_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=calendar_frame, anchor="nw")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Les créneaux de repas (lignes), sans ligne d'en-tête ici puisqu'elle
        # est maintenant affichée séparément, fixe, au-dessus de la grille.
        self.widgets = {}  # (jour, créneau) -> (combo, pers_entry)
        for row_index, slot in enumerate(self.MEAL_SLOTS):
            ttk.Label(calendar_frame, text=translate_mealslot_name(slot), font=("Segoe UI", sf(9)), anchor="w",
                      width=17, wraplength=120, justify="left").grid(
                row=row_index, column=0, padx=(2, 6), pady=4, sticky="w")
            for col, day in enumerate(WEEKDAYS, start=1):
                cell = tk.Frame(calendar_frame, background=COLOR_CARD, highlightbackground=COLOR_BORDER,
                                 highlightthickness=1)
                cell.grid(row=row_index, column=col, padx=2, pady=2, sticky="nsew")
                day_data = self.plan.get(day) or {}
                slot_data = day_data.get(slot) or {}
                combo = ttk.Combobox(cell, values=[t("common_none_option")] + recipe_names,
                                      state="readonly", width=13)
                combo.set(slot_data.get("recipe_name") or t("common_none_option"))
                combo.pack(padx=3, pady=(3, 1))
                pers_frame = ttk.Frame(cell, style="Card.TFrame")
                pers_frame.pack(padx=3, pady=(0, 3))
                ttk.Label(pers_frame, text="👤", style="Card.TLabel").pack(side="left")
                pers_entry = ttk.Entry(pers_frame, width=3)
                pers_entry.insert(0, str(slot_data.get("persons", 4)))
                pers_entry.pack(side="left")
                self.widgets[(day, slot)] = (combo, pers_entry)

        tk.Frame(calendar_frame, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).grid(
            row=len(self.MEAL_SLOTS), column=0, columnspan=len(WEEKDAYS) + 1, sticky="ew")

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=8)
        for col in range(4):
            btn_frame.columnconfigure(col, weight=1)
        ttk.Button(btn_frame, text=t("weekplan_save_button"),
                   command=self.save_plan).grid(row=0, column=0, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("weekplan_clear_button"),
                   command=self.clear_plan).grid(row=0, column=1, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("weekplan_export_ics_button"),
                   command=self.export_ics).grid(row=0, column=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("weekplan_compute_button"),
                   command=self.compute).grid(row=0, column=3, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_export_button"), command=self.open_export_dialog).grid(
            row=1, column=0, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_print_button"), command=self.print_list).grid(
            row=1, column=2, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("weekplan_checklist_button"),
                   command=self.open_checklist).grid(row=2, column=0, columnspan=4, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_add_manual_ingredient_button"),
                   command=self.open_add_manual_ingredient).grid(
            row=3, column=0, columnspan=4, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_save_list_button"),
                   command=self.save_list_for_later).grid(row=4, column=0, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("allrecipes_load_list_button"),
                   command=self.open_saved_lists).grid(row=4, column=2, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("weekhistory_heading"),
                   command=self.open_history).grid(row=5, column=0, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(btn_frame, text=t("weektemplates_heading"),
                   command=self.open_templates).grid(row=5, column=2, columnspan=2, padx=5, pady=3, sticky="ew")

        # ---- Zone de résultat éditable : chaque ingrédient peut voir sa
        # quantité modifiée ou être retiré, sans devoir tout recalculer. ----
        self.current_items = []       # liste plate éditable [{'name','quantity','unit','rayon'}, ...]
        self.last_chosen_recipes = []  # recettes utilisées lors du dernier calcul (pour les exports)

        result_container = ttk.Frame(self)
        result_container.pack(pady=10, padx=15, fill="both", expand=True)
        result_canvas = tk.Canvas(result_container, highlightthickness=0)
        result_scrollbar = ttk.Scrollbar(result_container, orient="vertical", command=result_canvas.yview)
        self.result_frame = ttk.Frame(result_canvas)
        self.result_frame.bind(
            "<Configure>", lambda e: result_canvas.configure(scrollregion=result_canvas.bbox("all"))
        )
        result_canvas.create_window((0, 0), window=self.result_frame, anchor="nw")
        result_canvas.configure(yscrollcommand=result_scrollbar.set)
        result_canvas.pack(side="left", fill="both", expand=True)
        result_scrollbar.pack(side="right", fill="y")

        def _on_result_mousewheel(event):
            result_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        result_canvas.bind("<Enter>", lambda e: result_canvas.bind_all("<MouseWheel>", _on_result_mousewheel))
        result_canvas.bind("<Leave>", lambda e: result_canvas.unbind_all("<MouseWheel>"))

        self._render_shopping_list()

    def open_add_manual_ingredient(self):
        AddManualIngredientDialog(self.app, self)

    def open_export_dialog(self):
        ExportFormatDialog(self, self.export_txt, self.export_excel, self.export_pdf)

    def add_manual_items(self, items):
        self.manual_items.extend(items)
        self.compute()  # actualise immédiatement la liste de courses affichée

    def _grouped_current_items(self):
        by_rayon = {}
        for i, item in enumerate(self.current_items):
            by_rayon.setdefault(item["rayon"], []).append(i)
        grouped = []
        for rayon in RAYON_ORDER:
            if rayon in by_rayon:
                idxs = sorted(by_rayon[rayon], key=lambda i: ingredient_sort_key(self.current_items[i]["name"]))
                grouped.append((rayon, idxs))
        return grouped

    def _render_shopping_list(self):
        for child in self.result_frame.winfo_children():
            child.destroy()

        if not self.current_items:
            ttk.Label(
                self.result_frame,
                text=t("weekplan_empty_list_message"),
                foreground=COLOR_TEXT_MUTED, justify="center"
            ).pack(pady=20)
            return

        ttk.Label(self.result_frame, text=t("weekplan_total_list_heading"),
                  font=("Segoe UI", sf(11), "bold")).pack(anchor="w", pady=(5, 2))
        if self.manual_items:
            ttk.Label(
                self.result_frame,
                text=t("allrecipes_manual_items_note", count=len(self.manual_items)),
                font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED
            ).pack(anchor="w")

        for rayon, idxs in self._grouped_current_items():
            ttk.Label(self.result_frame, text=translate_rayon_name(rayon), font=("Segoe UI", sf(10), "bold"),
                      foreground=COLOR_ACCENT_DARK).pack(anchor="w", pady=(12, 4))
            for idx in idxs:
                item = self.current_items[idx]
                row = ttk.Frame(self.result_frame)
                row.pack(fill="x", pady=1)
                ttk.Label(row, text=f"- {translate_ingredient_name(item['name'])}", width=30, anchor="w").pack(side="left")
                qty_entry = ttk.Entry(row, width=8)
                qty_entry.insert(0, str(item["quantity"]))
                qty_entry.pack(side="left", padx=3)
                qty_entry.bind("<FocusOut>", lambda e, i=idx, ent=qty_entry: self._update_item_quantity(i, ent))
                qty_entry.bind("<Return>", lambda e, i=idx, ent=qty_entry: self._update_item_quantity(i, ent))
                ttk.Label(row, text=translate_unit_name(item["unit"]), width=18, anchor="w").pack(side="left", padx=3)
                ttk.Button(row, text="🗑", width=3,
                           command=lambda i=idx: self._delete_item(i)).pack(side="left", padx=3)

        tk.Frame(self.result_frame, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _update_item_quantity(self, index, entry):
        if index >= len(self.current_items):
            return
        try:
            new_qty = float(entry.get().strip().replace(",", "."))
            if new_qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(t("common_error"), t("allrecipes_invalid_quantity"))
            entry.delete(0, tk.END)
            entry.insert(0, str(self.current_items[index]["quantity"]))
            return
        self.current_items[index]["quantity"] = new_qty

    def _delete_item(self, index):
        del self.current_items[index]
        self._render_shopping_list()

    def save_list_for_later(self):
        if not self.current_items:
            messagebox.showinfo(t("common_info"), t("allrecipes_calculate_list_first"),
                                 parent=self)
            return
        name = simpledialog.askstring(
            t("allrecipes_save_list_dialog_title"), t("allrecipes_save_list_dialog_prompt"), parent=self
        )
        if not name:
            self.lift()
            self.focus_force()
            return
        name = name.strip()
        if not name:
            self.lift()
            self.focus_force()
            return
        lists = load_saved_shopping_lists()
        lists = [l for l in lists if l["name"].lower() != name.lower()]
        lists.append({
            "name": name,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "items": [dict(item) for item in self.current_items],
        })
        save_saved_shopping_lists(lists)
        messagebox.showinfo(t("allrecipes_list_saved_title"), t("allrecipes_list_saved_message", name=name), parent=self)
        self.lift()
        self.focus_force()

    def open_saved_lists(self):
        SavedShoppingListsWindow(self.app, self)

    def load_saved_list(self, items):
        self.current_items = [dict(item) for item in items]
        self.last_chosen_recipes = []
        self._render_shopping_list()

    def _current_export_data(self):
        if not self.current_items:
            messagebox.showinfo(t("common_info"), t("weekplan_calculate_list_for_export"))
            return None
        grouped_totals = grouped_totals_from_flat_items(self.current_items)
        return self.last_chosen_recipes, grouped_totals

    def _collect_selection(self):
        new_plan = {}
        pairs = []
        for day in WEEKDAYS:
            for slot in self.MEAL_SLOTS:
                combo, pers_entry = self.widgets[(day, slot)]
                name = combo.get()
                if not name or name == t("common_none_option"):
                    continue
                try:
                    persons = float(pers_entry.get().strip().replace(",", "."))
                except ValueError:
                    messagebox.showerror(
                        t("common_error"),
                        t("weekplan_invalid_persons_for_slot", day=translate_weekday_name(day), slot=translate_mealslot_name(slot))
                    )
                    return None, None
                recipe = find_recipe_by_name(self.app.recipes, name)
                if recipe is None:
                    continue
                new_plan.setdefault(day, {})[slot] = {"recipe_name": name, "persons": persons}
                pairs.append((recipe, persons))
        return new_plan, pairs

    def save_plan(self):
        new_plan, pairs = self._collect_selection()
        if new_plan is None:
            return
        save_weekly_plan(new_plan)
        archive_current_week(new_plan)
        self.plan = new_plan
        messagebox.showinfo(t("allrecipes_list_saved_title"), t("weekplan_saved_message"))

    def clear_plan(self):
        if not messagebox.askyesno(t("common_confirm"), t("weekplan_clear_confirm_message")):
            return
        for (day, slot), (combo, pers_entry) in self.widgets.items():
            combo.set(t("common_none_option"))
        save_weekly_plan({})
        self.plan = {}

    def apply_plan(self, plan):
        """Recharge un planning donné (depuis l'historique ou un modèle)
        dans les cases actuellement affichées, sans reconstruire la
        fenêtre. Ne l'enregistre pas automatiquement : il faut toujours
        cliquer sur « 💾 Enregistrer le planning » pour le conserver."""
        for (day, slot), (combo, pers_entry) in self.widgets.items():
            day_data = plan.get(day) or {}
            slot_data = day_data.get(slot) or {}
            combo.set(slot_data.get("recipe_name") or t("common_none_option"))
            pers_entry.delete(0, tk.END)
            pers_entry.insert(0, str(slot_data.get("persons", 4)))

    def open_history(self):
        WeeklyPlanHistoryWindow(self.app, self)

    def open_templates(self):
        WeeklyPlanTemplatesWindow(self.app, self)

    def export_ics(self):
        new_plan, pairs = self._collect_selection()
        if new_plan is None:
            return
        if not new_plan:
            messagebox.showinfo(t("common_info"), t("weekplan_assign_recipe_first"))
            return
        path = filedialog.asksaveasfilename(
            title=t("weekplan_export_ics_title"),
            defaultextension=".ics",
            filetypes=[("Fichier calendrier (.ics)", "*.ics")],
            initialfile="planning_repas.ics"
        )
        if not path:
            return
        try:
            content = build_weekly_plan_ics(new_plan)
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(
            t("common_export_success_title"),
            t("weekplan_ics_export_success_message", path=path)
        )

    def compute(self):
        new_plan, pairs = self._collect_selection()
        if new_plan is None:
            return None
        if not pairs and not self.manual_items:
            messagebox.showinfo(t("common_info"), t("weekplan_assign_or_manual"))
            return None

        if self.manual_items:
            pairs = list(pairs) + [({"ingredients": list(self.manual_items)}, 1)]

        grouped_totals = compute_grouped_totals(pairs)
        day_order = {d: i for i, d in enumerate(WEEKDAYS)}
        slot_order = {s: i for i, s in enumerate(self.MEAL_SLOTS)}
        chosen_recipes = []
        for day, slots in new_plan.items():
            for slot, info in slots.items():
                chosen_recipes.append((f"{day} — {slot} : {info['recipe_name']}", info["persons"]))
        chosen_recipes.sort(
            key=lambda t: (day_order.get(t[0].split(" — ")[0], 99),
                            slot_order.get(t[0].split(" — ")[1].split(" : ")[0], 99))
        )

        self.current_items = []
        for rayon, items in grouped_totals:
            for name, qty, unit in items:
                self.current_items.append({"name": name, "quantity": qty, "unit": unit, "rayon": rayon})
        self.last_chosen_recipes = chosen_recipes
        self._render_shopping_list()

        return chosen_recipes, grouped_totals

    def export_txt(self):
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        path = filedialog.asksaveasfilename(
            title=t("weekplan_export_shopping_list_title"), defaultextension=".txt",
            filetypes=[("Fichier texte", "*.txt")], initialfile="liste_de_courses_semaine.txt"
        )
        if not path:
            return
        try:
            write_shopping_list_txt(path, t("weekplan_shopping_list_title"), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("weekplan_list_saved_message", path=path))

    def export_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("weekplan_excel_module_missing"))
            return
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        path = filedialog.asksaveasfilename(
            title=t("weekplan_export_shopping_list_title"), defaultextension=".xlsx",
            filetypes=[("Fichier Excel", "*.xlsx")], initialfile="liste_de_courses_semaine.xlsx"
        )
        if not path:
            return
        try:
            wb = build_shopping_list_workbook(chosen_recipes, grouped_totals)
            wb.save(path)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("weekplan_list_saved_message", path=path))

    def export_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("weekplan_pdf_module_missing"))
            return
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        path = filedialog.asksaveasfilename(
            title=t("weekplan_export_shopping_list_title"), defaultextension=".pdf",
            filetypes=[("Fichier PDF", "*.pdf")], initialfile="liste_de_courses_semaine.pdf"
        )
        if not path:
            return
        try:
            build_shopping_list_pdf(path, t("weekplan_shopping_list_title"), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("weekplan_list_saved_message", path=path))

    def print_list(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("weekplan_print_module_missing"))
            return
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        temp_path = get_temp_pdf_path("planning_semaine")
        try:
            build_shopping_list_pdf(temp_path, t("weekplan_shopping_list_title"), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_print_failed", error=e))
            return
        result_status = print_file(temp_path)
        report_print_result(result_status, temp_path, t("weekplan_print_label"))

    def open_checklist(self):
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        ShoppingChecklistWindow(self.app, grouped_totals, title=t("weekplan_shopping_list_title"))


class MenuManagerWindow(tk.Toplevel):
    """Gestion des menus (combinaisons de plusieurs recettes) : création,
    édition, suppression."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("menumanager_title"))
        self.geometry(f"{gs(420)}x{gs(480)}")
        self.grab_set()

        ttk.Label(self, text=t("menumanager_list_label"), font=("Segoe UI", sf(11), "bold")).pack(pady=(10, 5))

        self.listbox = tk.Listbox(self, width=40, height=14, font=("Segoe UI", sf(9)))
        self.listbox.pack(pady=5, padx=15, fill="both", expand=True)
        self.menus = []
        self._populate()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text=t("menumanager_new_button"), command=self.new_menu).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text=t("home_open_button"), command=self.open_menu).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text=t("managerecipes_delete_button"), command=self.delete_menu).grid(row=0, column=2, padx=5)

    def _populate(self):
        self.listbox.delete(0, tk.END)
        self.menus = load_menus()
        for menu in self.menus:
            self.listbox.insert(tk.END, t("menumanager_recipe_count", name=menu['name'], count=len(menu.get('items', []))))

    def _selected_index(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showinfo(t("common_info"), t("menumanager_select_menu_first"))
            return None
        return sel[0]

    def new_menu(self):
        MenuFormWindow(self.app, self, menu_index=None)

    def open_menu(self):
        idx = self._selected_index()
        if idx is None:
            return
        MenuFormWindow(self.app, self, menu_index=idx)

    def delete_menu(self):
        idx = self._selected_index()
        if idx is None:
            return
        menus = load_menus()
        name = menus[idx]["name"]
        if not messagebox.askyesno(t("common_confirm"), t("menumanager_delete_confirm", name=name)):
            return
        menus.pop(idx)
        save_menus(menus)
        self._populate()


class MenuFormWindow(tk.Toplevel):
    """Fenêtre de création/édition d'un menu : nom + liste de recettes avec
    nombre de personnes, export/impression de sa liste de courses."""

    CATEGORY_ORDER = {"Apéro": 0, "Entrée": 1, "Plat": 2, "Sauce": 3,
                       "Dessert": 4, "Boisson": 5, "Autre": 6}

    def __init__(self, app, manager, menu_index=None):
        super().__init__(app)
        self.app = app
        self.manager = manager
        self.menu_index = menu_index
        self.editing = menu_index is not None
        menus = load_menus()
        self.existing_menu = menus[menu_index] if self.editing else None

        self.title(t("menuform_title_edit") if self.editing else t("menuform_title_new"))
        screen_height = get_usable_screen_height(self)
        self.geometry(f"{gs(700)}x{screen_height}+40+0")
        self.minsize(gs(560), gs(400))
        self.resizable(True, True)
        self.grab_set()

        self.manual_items = []  # ingrédients ajoutés manuellement (hors menu) : [{"name","quantity","unit"}]

        ttk.Label(self, text=t("menuform_name_label"), font=("Segoe UI", sf(11), "bold")).pack(pady=(10, 5))
        self.name_entry = ttk.Entry(self, width=40)
        self.name_entry.pack()
        if self.editing:
            self.name_entry.insert(0, self.existing_menu["name"])

        self.items = [dict(it) for it in self.existing_menu.get("items", [])] if self.editing else []

        ttk.Label(self, text=t("menuform_add_recipe_label"), font=("Segoe UI", sf(10), "bold")).pack(pady=(15, 5))
        add_frame = ttk.Frame(self)
        add_frame.pack(padx=15, fill="x")
        recipe_names = [r["name"] for r in self.app.recipes]
        self.recipe_combo = ttk.Combobox(add_frame, values=recipe_names, state="readonly", width=26)
        self.recipe_combo.pack(side="left", padx=(0, 5))
        if recipe_names:
            self.recipe_combo.current(0)
        ttk.Label(add_frame, text=t("menuform_persons_short_label")).pack(side="left")
        self.add_persons_entry = ttk.Entry(add_frame, width=5)
        self.add_persons_entry.insert(0, "4")
        self.add_persons_entry.pack(side="left", padx=5)
        ttk.Button(add_frame, text=t("menuform_add_button"), command=self.add_item).pack(side="left", padx=5)

        ttk.Label(self, text=t("menuform_recipes_label"), font=("Segoe UI", sf(10), "bold")).pack(pady=(15, 5))
        self.items_listbox = tk.Listbox(self, width=55, height=7, font=("Segoe UI", sf(9)))
        self.items_listbox.pack(padx=15, fill="x")
        self._refresh_items_listbox()
        ttk.Button(self, text=t("menuform_remove_button"), command=self.remove_item).pack(pady=5)

        ttk.Button(self, text=t("menuform_save_button"), command=self.save_menu).pack(pady=8)

        export_frame = ttk.Frame(self)
        export_frame.pack(pady=5)
        for col in range(4):
            export_frame.columnconfigure(col, weight=1)
        ttk.Button(export_frame, text=t("menuform_compute_button"),
                   command=self.compute).grid(row=0, column=0, columnspan=4, padx=5, pady=3, sticky="ew")
        ttk.Button(export_frame, text=t("allrecipes_export_button"), command=self.open_export_dialog).grid(
            row=1, column=0, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(export_frame, text=t("allrecipes_print_button"), command=self.print_list).grid(
            row=1, column=2, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(export_frame, text=t("weekplan_checklist_button"),
                   command=self.open_checklist).grid(row=2, column=0, columnspan=4, padx=5, pady=3, sticky="ew")
        ttk.Button(export_frame, text=t("allrecipes_add_manual_ingredient_button"),
                   command=self.open_add_manual_ingredient).grid(
            row=3, column=0, columnspan=4, padx=5, pady=3, sticky="ew")
        ttk.Button(export_frame, text=t("allrecipes_save_list_button"),
                   command=self.save_list_for_later).grid(row=4, column=0, columnspan=2, padx=5, pady=3, sticky="ew")
        ttk.Button(export_frame, text=t("allrecipes_load_list_button"),
                   command=self.open_saved_lists).grid(row=4, column=2, columnspan=2, padx=5, pady=3, sticky="ew")

        # ---- Zone de résultat éditable : chaque ingrédient peut voir sa
        # quantité modifiée ou être retiré, sans devoir tout recalculer. ----
        self.current_items = []       # liste plate éditable [{'name','quantity','unit','rayon'}, ...]
        self.last_chosen_recipes = []  # recettes utilisées lors du dernier calcul (pour les exports)

        result_container = ttk.Frame(self)
        result_container.pack(pady=10, padx=15, fill="both", expand=True)
        result_canvas = tk.Canvas(result_container, highlightthickness=0)
        result_scrollbar = ttk.Scrollbar(result_container, orient="vertical", command=result_canvas.yview)
        self.result_frame = ttk.Frame(result_canvas)
        self.result_frame.bind(
            "<Configure>", lambda e: result_canvas.configure(scrollregion=result_canvas.bbox("all"))
        )
        result_canvas.create_window((0, 0), window=self.result_frame, anchor="nw")
        result_canvas.configure(yscrollcommand=result_scrollbar.set)
        result_canvas.pack(side="left", fill="both", expand=True)
        result_scrollbar.pack(side="right", fill="y")

        def _on_result_mousewheel(event):
            result_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        result_canvas.bind("<Enter>", lambda e: result_canvas.bind_all("<MouseWheel>", _on_result_mousewheel))
        result_canvas.bind("<Leave>", lambda e: result_canvas.unbind_all("<MouseWheel>"))

        self._render_shopping_list()

    def open_add_manual_ingredient(self):
        AddManualIngredientDialog(self.app, self)

    def open_export_dialog(self):
        ExportFormatDialog(self, self.export_txt, self.export_excel, self.export_pdf)

    def add_manual_items(self, items):
        self.manual_items.extend(items)
        self.compute()  # actualise immédiatement la liste de courses affichée

    def _grouped_current_items(self):
        by_rayon = {}
        for i, item in enumerate(self.current_items):
            by_rayon.setdefault(item["rayon"], []).append(i)
        grouped = []
        for rayon in RAYON_ORDER:
            if rayon in by_rayon:
                idxs = sorted(by_rayon[rayon], key=lambda i: ingredient_sort_key(self.current_items[i]["name"]))
                grouped.append((rayon, idxs))
        return grouped

    def _render_shopping_list(self):
        for child in self.result_frame.winfo_children():
            child.destroy()

        if not self.current_items:
            ttk.Label(
                self.result_frame,
                text=t("menuform_empty_list_message"),
                foreground=COLOR_TEXT_MUTED, justify="center"
            ).pack(pady=20)
            return

        ttk.Label(self.result_frame, text=t("menuform_total_list_heading"),
                  font=("Segoe UI", sf(11), "bold")).pack(anchor="w", pady=(5, 2))
        if self.manual_items:
            ttk.Label(
                self.result_frame,
                text=t("allrecipes_manual_items_note", count=len(self.manual_items)),
                font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED
            ).pack(anchor="w")

        for rayon, idxs in self._grouped_current_items():
            ttk.Label(self.result_frame, text=translate_rayon_name(rayon), font=("Segoe UI", sf(10), "bold"),
                      foreground=COLOR_ACCENT_DARK).pack(anchor="w", pady=(12, 4))
            for idx in idxs:
                item = self.current_items[idx]
                row = ttk.Frame(self.result_frame)
                row.pack(fill="x", pady=1)
                ttk.Label(row, text=f"- {translate_ingredient_name(item['name'])}", width=30, anchor="w").pack(side="left")
                qty_entry = ttk.Entry(row, width=8)
                qty_entry.insert(0, str(item["quantity"]))
                qty_entry.pack(side="left", padx=3)
                qty_entry.bind("<FocusOut>", lambda e, i=idx, ent=qty_entry: self._update_item_quantity(i, ent))
                qty_entry.bind("<Return>", lambda e, i=idx, ent=qty_entry: self._update_item_quantity(i, ent))
                ttk.Label(row, text=translate_unit_name(item["unit"]), width=18, anchor="w").pack(side="left", padx=3)
                ttk.Button(row, text="🗑", width=3,
                           command=lambda i=idx: self._delete_item(i)).pack(side="left", padx=3)

        tk.Frame(self.result_frame, height=SCROLL_BOTTOM_PADDING, background=COLOR_BG).pack(fill="x")

    def _update_item_quantity(self, index, entry):
        if index >= len(self.current_items):
            return
        try:
            new_qty = float(entry.get().strip().replace(",", "."))
            if new_qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(t("common_error"), t("allrecipes_invalid_quantity"))
            entry.delete(0, tk.END)
            entry.insert(0, str(self.current_items[index]["quantity"]))
            return
        self.current_items[index]["quantity"] = new_qty

    def _delete_item(self, index):
        del self.current_items[index]
        self._render_shopping_list()

    def save_list_for_later(self):
        if not self.current_items:
            messagebox.showinfo(t("common_info"), t("allrecipes_calculate_list_first"),
                                 parent=self)
            return
        name = simpledialog.askstring(
            t("allrecipes_save_list_dialog_title"), t("allrecipes_save_list_dialog_prompt"), parent=self
        )
        if not name:
            self.lift()
            self.focus_force()
            return
        name = name.strip()
        if not name:
            self.lift()
            self.focus_force()
            return
        lists = load_saved_shopping_lists()
        lists = [l for l in lists if l["name"].lower() != name.lower()]
        lists.append({
            "name": name,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "items": [dict(item) for item in self.current_items],
        })
        save_saved_shopping_lists(lists)
        messagebox.showinfo(t("allrecipes_list_saved_title"), t("allrecipes_list_saved_message", name=name), parent=self)
        self.lift()
        self.focus_force()

    def open_saved_lists(self):
        SavedShoppingListsWindow(self.app, self)

    def load_saved_list(self, items):
        self.current_items = [dict(item) for item in items]
        self.last_chosen_recipes = []
        self._render_shopping_list()

    def _current_export_data(self):
        if not self.current_items:
            messagebox.showinfo(t("common_info"), t("menuform_calculate_list_for_export"))
            return None
        grouped_totals = grouped_totals_from_flat_items(self.current_items)
        return self.last_chosen_recipes, grouped_totals

    def _refresh_items_listbox(self):
        self.items_listbox.delete(0, tk.END)
        for item in self.items:
            recipe = find_recipe_by_name(self.app.recipes, item["recipe_name"])
            cat = translate_category_name(recipe.get("category", "Autre")) if recipe else "?"
            self.items_listbox.insert(
                tk.END, t("menuform_item_row_label", cat=cat, name=item['recipe_name'], persons=item['persons'])
            )

    def add_item(self):
        name = self.recipe_combo.get()
        if not name:
            return
        try:
            persons = float(self.add_persons_entry.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror(t("common_error"), t("onerecipe_invalid_persons"))
            return
        self.items.append({"recipe_name": name, "persons": persons})
        self._refresh_items_listbox()

    def remove_item(self):
        sel = self.items_listbox.curselection()
        if not sel:
            messagebox.showinfo(t("common_info"), t("menuform_select_recipe_to_remove"))
            return
        self.items.pop(sel[0])
        self._refresh_items_listbox()

    def save_menu(self):
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror(t("common_error"), t("menuform_error_name_required"))
            return
        if not self.items:
            messagebox.showerror(t("common_error"), t("menuform_error_no_recipe"))
            return
        menus = load_menus()
        menu_data = {"name": name, "items": self.items}
        if self.editing:
            menus[self.menu_index] = menu_data
        else:
            menus.append(menu_data)
        save_menus(menus)
        self.manager._populate()
        messagebox.showinfo(t("allrecipes_list_saved_title"), t("menuform_saved_message", name=name))

    def _collect_pairs(self):
        pairs = []
        chosen_recipes = []
        sorted_items = sorted(
            self.items,
            key=lambda it: self.CATEGORY_ORDER.get(
                (find_recipe_by_name(self.app.recipes, it["recipe_name"]) or {}).get("category", "Autre"), 4
            )
        )
        for item in sorted_items:
            recipe = find_recipe_by_name(self.app.recipes, item["recipe_name"])
            if recipe is None:
                continue
            persons = item["persons"]
            pairs.append((recipe, persons))
            cat = translate_category_name(recipe.get("category", "Autre"))
            chosen_recipes.append((f"{cat} — {recipe['name']}", persons))
        return pairs, chosen_recipes

    def compute(self):
        pairs, chosen_recipes = self._collect_pairs()
        if not pairs and not self.manual_items:
            messagebox.showinfo(
                t("common_info"), t("menuform_add_recipe_or_manual")
            )
            return None
        if self.manual_items:
            pairs = list(pairs) + [({"ingredients": list(self.manual_items)}, 1)]
        grouped_totals = compute_grouped_totals(pairs)

        self.current_items = []
        for rayon, items in grouped_totals:
            for name, qty, unit in items:
                self.current_items.append({"name": name, "quantity": qty, "unit": unit, "rayon": rayon})
        self.last_chosen_recipes = chosen_recipes
        self._render_shopping_list()

        return chosen_recipes, grouped_totals

    def export_txt(self):
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        menu_name = self.name_entry.get().strip() or "menu"
        path = filedialog.asksaveasfilename(
            title=t("weekplan_export_shopping_list_title"), defaultextension=".txt",
            filetypes=[("Fichier texte", "*.txt")], initialfile=f"{menu_name}.txt"
        )
        if not path:
            return
        try:
            write_shopping_list_txt(path, t("menuform_shopping_list_title", name=menu_name), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("weekplan_list_saved_message", path=path))

    def export_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("weekplan_excel_module_missing"))
            return
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        menu_name = self.name_entry.get().strip() or "menu"
        path = filedialog.asksaveasfilename(
            title=t("weekplan_export_shopping_list_title"), defaultextension=".xlsx",
            filetypes=[("Fichier Excel", "*.xlsx")], initialfile=f"{menu_name}.xlsx"
        )
        if not path:
            return
        try:
            wb = build_shopping_list_workbook(chosen_recipes, grouped_totals)
            wb.save(path)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("weekplan_list_saved_message", path=path))

    def export_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("weekplan_pdf_module_missing"))
            return
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        menu_name = self.name_entry.get().strip() or "menu"
        path = filedialog.asksaveasfilename(
            title=t("weekplan_export_shopping_list_title"), defaultextension=".pdf",
            filetypes=[("Fichier PDF", "*.pdf")], initialfile=f"{menu_name}.pdf"
        )
        if not path:
            return
        try:
            build_shopping_list_pdf(path, t("menuform_shopping_list_title", name=menu_name), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("weekplan_list_saved_message", path=path))

    def print_list(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(t("common_module_missing"), t("weekplan_print_module_missing"))
            return
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        menu_name = self.name_entry.get().strip() or "menu"
        temp_path = get_temp_pdf_path("menu")
        try:
            build_shopping_list_pdf(temp_path, t("menuform_shopping_list_title", name=menu_name), chosen_recipes, grouped_totals)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_print_failed", error=e))
            return
        result_status = print_file(temp_path)
        report_print_result(result_status, temp_path, t("menuform_print_label", name=menu_name))

    def open_checklist(self):
        result = self._current_export_data()
        if result is None:
            return
        chosen_recipes, grouped_totals = result
        menu_name = self.name_entry.get().strip() or "menu"
        ShoppingChecklistWindow(self.app, grouped_totals, title=t("menuform_shopping_list_title", name=menu_name))


class ImportFromUrlWindow(tk.Toplevel):
    """Importe une recette à partir d'un lien internet (fonctionne avec les
    sites utilisant le format de données standard Schema.org Recipe)."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("importurl_title"))
        self.geometry(f"{gs(520)}x{gs(280)}")
        self.resizable(False, False)
        self.grab_set()

        ttk.Label(self, text=t("importurl_heading"),
                  font=("Segoe UI", sf(13), "bold")).pack(pady=15)
        ttk.Label(
            self,
            text=t("importurl_intro"),
            justify="center", font=("Segoe UI", sf(9))
        ).pack(pady=(0, 15))

        self.url_entry = ttk.Entry(self, width=55)
        self.url_entry.pack(pady=5, padx=20, fill="x")
        self.url_entry.bind("<Return>", lambda e: self.fetch())

        self.status_label = ttk.Label(self, text="", font=("Segoe UI", sf(9)), foreground=COLOR_TEXT_MUTED)
        self.status_label.pack(pady=(5, 5))

        self.fetch_button = ttk.Button(self, text=t("importurl_fetch_button"), command=self.fetch)
        self.fetch_button.pack(pady=10)

        ttk.Label(
            self,
            text=t("importurl_after_import_note"),
            justify="center", font=("Segoe UI", sf(8)), foreground="#999"
        ).pack(pady=(5, 0))

    def fetch(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showinfo(t("common_info"), t("importurl_paste_url_first"))
            return
        if not url.lower().startswith(("http://", "https://")):
            url = "https://" + url

        self.fetch_button.config(state="disabled")
        self.status_label.config(text=t("importurl_fetching"))
        self.update()

        try:
            recipe_data = fetch_recipe_from_url(url)
        except Exception as e:
            self.status_label.config(text="")
            self.fetch_button.config(state="normal")
            messagebox.showerror(t("importurl_failed_title"), str(e))
            return

        # Enregistre automatiquement tout ingrédient qui n'existe pas encore.
        # Si l'ingrédient importé n'est qu'une variante singulier/pluriel
        # d'un ingrédient déjà connu (ex. la recette utilise "Tomates" alors
        # que la liste a déjà "Tomate"), on réutilise directement la forme
        # existante plutôt que de créer un doublon — à la fois dans la liste
        # ET dans le nom de l'ingrédient de la recette elle-même, pour que
        # la détection des allergènes/valeurs nutritionnelles continue de
        # fonctionner sur cet ingrédient.
        known_lower = {n.lower() for n in self.app.ingredient_names}
        ingredients_list = load_ingredients()
        changed = False
        for ing in recipe_data["ingredients"]:
            if ing["name"].lower() in known_lower:
                continue
            plural_match = find_plural_duplicate(ing["name"], ingredients_list)
            if plural_match:
                ing["name"] = plural_match
                continue
            ingredients_list.append(ing["name"])
            known_lower.add(ing["name"].lower())
            changed = True
        if changed:
            self.app.ingredient_names = save_ingredients(ingredients_list)

        self.status_label.config(text="")
        self.fetch_button.config(state="normal")
        self.destroy()
        RecipeFormWindow(self.app, recipe_index=None, prefill=recipe_data)


class ImportFromPhotoWindow(tk.Toplevel):
    """Importe une recette à partir d'une photo (recette manuscrite ou page
    d'un livre de cuisine), en extrayant le texte par reconnaissance optique
    de caractères (OCR). Contrairement à l'import depuis un lien, le texte
    extrait n'est pas automatiquement organisé en ingrédients/étapes — il
    est présenté à relire et corriger avant de créer la recette."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("importphoto_title"))
        self.geometry(f"{gs(540)}x{gs(640)}")
        self.minsize(gs(460), gs(500))
        self.resizable(True, True)
        self.grab_set()
        self.photo_path = None
        self._preview_ref = None

        ttk.Label(self, text=t("importphoto_heading"),
                  font=("Segoe UI", sf(13), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self,
            text=t("importphoto_intro"),
            font=("Segoe UI", sf(8)), foreground=COLOR_TEXT_MUTED, justify="center", wraplength=480
        ).pack(pady=(0, 10))

        if not PYTESSERACT_AVAILABLE:
            ttk.Label(
                self,
                text=t("importphoto_module_warning"),
                foreground=COLOR_ERROR, font=("Segoe UI", sf(9), "bold"), justify="center"
            ).pack(pady=10)

        self.preview_label = ttk.Label(self, text=t("importphoto_no_photo_chosen"), foreground=COLOR_TEXT_MUTED)
        self.preview_label.pack(pady=5)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text=t("importphoto_choose_button"),
                   command=self.choose_photo).grid(row=0, column=0, padx=5)
        self.extract_button = ttk.Button(btn_frame, text=t("importphoto_extract_button"),
                                          command=self.extract_text, state="disabled")
        self.extract_button.grid(row=0, column=1, padx=5)

        ttk.Label(self, text=t("importphoto_extracted_text_label"), font=("Segoe UI", sf(9), "bold")).pack(
            pady=(10, 3))
        self.text_box = tk.Text(self, height=14, wrap="word", font=("Segoe UI", sf(10)))
        self.text_box.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        ttk.Button(self, text=t("importphoto_create_button"),
                   command=self.create_recipe).pack(pady=(0, 15))

    def choose_photo(self):
        path = filedialog.askopenfilename(
            title=t("importphoto_choose_photo_title"),
            filetypes=[("Images", "*.jpg *.jpeg *.png *.webp *.bmp *.tiff")]
        )
        if not path:
            return
        self.photo_path = path
        self.preview_label.config(text=os.path.basename(path), foreground=COLOR_TEXT)

        if PIL_AVAILABLE:
            try:
                img = Image.open(path)
                img.thumbnail((300, 220))
                self._preview_ref = ImageTk.PhotoImage(img)
                self.preview_label.config(image=self._preview_ref, text="", compound="top")
            except Exception:
                pass

        if PYTESSERACT_AVAILABLE:
            self.extract_button.config(state="normal")

    def extract_text(self):
        if not self.photo_path:
            messagebox.showinfo(t("common_info"), t("importphoto_choose_first"))
            return
        if not PYTESSERACT_AVAILABLE:
            messagebox.showerror(
                t("common_module_missing"),
                t("importphoto_ocr_module_missing")
            )
            return
        try:
            image = Image.open(self.photo_path) if PIL_AVAILABLE else self.photo_path
            tesseract_lang = TESSERACT_LANG_CODES.get(CURRENT_LANGUAGE, "fra")
            extracted = pytesseract.image_to_string(image, lang=tesseract_lang)
        except Exception as e:
            messagebox.showerror(
                t("importphoto_extraction_failed_title"),
                t("importphoto_extraction_failed_message", error=e)
            )
            return
        extracted = extracted.strip()
        self.text_box.delete("1.0", tk.END)
        if extracted:
            self.text_box.insert("1.0", extracted)
        else:
            messagebox.showinfo(
                t("common_info"),
                t("importphoto_no_text_extracted")
            )

    def create_recipe(self):
        raw_text = self.text_box.get("1.0", "end-1c").strip()
        if not raw_text:
            if not messagebox.askyesno(
                t("importphoto_no_text_title"),
                t("importphoto_no_text_confirm")
            ):
                return

        images = []
        if self.photo_path:
            copied = copy_image_to_store(self.photo_path)
            if copied:
                images.append(copied)

        prefill = {
            "name": "",
            "description": raw_text[:2056],
            "ingredients": [],
            "prep_time": "",
            "cook_time": "",
            "default_persons": 4,
            "images": images,
        }
        self.destroy()
        RecipeFormWindow(self.app, recipe_index=None, prefill=prefill)


class CookbookExportWindow(tk.Toplevel):
    """Exporte plusieurs recettes réunies en un seul PDF façon livre de
    cuisine (page de sommaire puis une recette par page)."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("cookbookexport_title"))
        self.geometry(f"{gs(560)}x{gs(600)}")
        self.grab_set()

        ttk.Label(self, text=t("cookbookexport_heading"),
                  font=("Segoe UI", sf(13), "bold")).pack(pady=(15, 5))
        ttk.Label(
            self,
            text=t("cookbookexport_intro"),
            justify="center", font=("Segoe UI", sf(9))
        ).pack(pady=(0, 10))

        filter_frame = ttk.Frame(self)
        filter_frame.pack(pady=(0, 5), fill="x", padx=15)
        ttk.Label(filter_frame, text=t("cookbookexport_filter_label")).pack(side="left")
        self.category_filter = ttk.Combobox(
            filter_frame, values=[t("common_all_categories")] + [translate_category_name(c) for c in RecipeFormWindow.CATEGORY_OPTIONS],
            state="readonly", width=15
        )
        self.category_filter.set(t("common_all_categories"))
        self.category_filter.pack(side="left", padx=5)
        self.category_filter.bind("<<ComboboxSelected>>", lambda e: self._populate())
        ttk.Button(filter_frame, text=t("cookbookexport_check_all_button"), command=self.check_all).pack(side="left", padx=5)
        ttk.Button(filter_frame, text=t("cookbookexport_uncheck_all_button"), command=self.uncheck_all).pack(side="left")

        list_frame = ttk.Frame(self)
        list_frame.pack(padx=15, pady=5, fill="both", expand=True)
        canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        self.rows_frame = ttk.Frame(canvas)
        self.rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.checks = []  # (var, recipe)
        self._populate()

        ttk.Button(self, text=t("cookbookexport_generate_button"),
                   command=self.export_pdf).pack(pady=15)

    def _populate(self):
        for child in self.rows_frame.winfo_children():
            child.destroy()
        self.checks = []
        category = resolve_category_input(self.category_filter.get(), RecipeFormWindow.CATEGORY_OPTIONS)
        for recipe in self.app.recipes:
            if category != t("common_all_categories") and recipe.get("category", "Autre") != category:
                continue
            var = tk.BooleanVar(value=True)
            ttk.Checkbutton(self.rows_frame, text=format_recipe_list_label(recipe),
                             variable=var).pack(anchor="w", pady=2)
            self.checks.append((var, recipe))

    def check_all(self):
        for var, recipe in self.checks:
            var.set(True)

    def uncheck_all(self):
        for var, recipe in self.checks:
            var.set(False)

    def export_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(
                t("common_module_missing"),
                t("onerecipe_pdf_module_missing")
            )
            return
        selected = [(recipe, recipe.get("default_persons", 4) or 4)
                    for var, recipe in self.checks if var.get()]
        if not selected:
            messagebox.showinfo(t("common_info"), t("cookbookexport_error_select_recipe"))
            return

        path = filedialog.asksaveasfilename(
            title=t("cookbookexport_save_dialog_title"),
            defaultextension=".pdf",
            filetypes=[("Fichier PDF", "*.pdf")],
            initialfile="mon_livre_de_recettes.pdf"
        )
        if not path:
            return
        try:
            build_cookbook_pdf(path, selected)
        except Exception as e:
            messagebox.showerror(t("common_error"), t("common_export_failed", error=e))
            return
        messagebox.showinfo(t("common_export_success_title"), t("cookbookexport_saved_message", path=path))


class CompareRecipesWindow(tk.Toplevel):
    """Compare deux recettes côte à côte : temps, difficulté, note, et
    ingrédients communs / différents."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("compare_title"))
        self.geometry(f"{gs(900)}x{gs(700)}")
        self.grab_set()

        recipe_names = [r["name"] for r in self.app.recipes]

        picker_frame = ttk.Frame(self)
        picker_frame.pack(pady=15, padx=15, fill="x")

        ttk.Label(picker_frame, text=t("compare_recipe_a_label"), font=("Segoe UI", sf(10), "bold")).grid(
            row=0, column=0, sticky="w", padx=(0, 5))
        self.combo_a = ttk.Combobox(picker_frame, values=recipe_names, state="readonly", width=32)
        self.combo_a.grid(row=0, column=1, padx=5)
        if recipe_names:
            self.combo_a.current(0)

        ttk.Label(picker_frame, text=t("compare_recipe_b_label"), font=("Segoe UI", sf(10), "bold")).grid(
            row=1, column=0, sticky="w", padx=(0, 5), pady=(8, 0))
        self.combo_b = ttk.Combobox(picker_frame, values=recipe_names, state="readonly", width=32)
        self.combo_b.grid(row=1, column=1, padx=5, pady=(8, 0))
        if len(recipe_names) > 1:
            self.combo_b.current(1)
        elif recipe_names:
            self.combo_b.current(0)

        ttk.Button(picker_frame, text=t("compare_button"), command=self.compare).grid(
            row=0, column=2, rowspan=2, padx=15)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.result_frame = ttk.Frame(canvas)
        self.result_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.result_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

    def compare(self):
        name_a = self.combo_a.get()
        name_b = self.combo_b.get()
        if not name_a or not name_b:
            messagebox.showinfo(t("common_info"), t("compare_choose_each_list"))
            return
        recipe_a = find_recipe_by_name(self.app.recipes, name_a)
        recipe_b = find_recipe_by_name(self.app.recipes, name_b)
        if recipe_a is None or recipe_b is None:
            return

        for child in self.result_frame.winfo_children():
            child.destroy()

        # ---- Tableau comparatif (vraies colonnes de grille, toujours bien
        # alignées, contrairement à un padding par espaces dans du texte) ----
        table = ttk.Frame(self.result_frame)
        table.pack(fill="x", pady=(0, 15))
        table.columnconfigure(0, weight=0)
        table.columnconfigure(1, weight=1)
        table.columnconfigure(2, weight=1)

        ttk.Label(table, text="", width=16).grid(row=0, column=0)
        ttk.Label(table, text=name_a, font=("Segoe UI", sf(10), "bold"),
                  foreground=COLOR_ACCENT_DARK, wraplength=260).grid(row=0, column=1, padx=8, pady=(0, 6), sticky="w")
        ttk.Label(table, text=name_b, font=("Segoe UI", sf(10), "bold"),
                  foreground=COLOR_ACCENT_DARK, wraplength=260).grid(row=0, column=2, padx=8, pady=(0, 6), sticky="w")
        ttk.Separator(table, orient="horizontal").grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 6))

        row_counter = [2]

        def field_line(label, value_a, value_b):
            r = row_counter[0]
            ttk.Label(table, text=label, font=("Segoe UI", sf(9), "bold")).grid(
                row=r, column=0, sticky="w", pady=3, padx=(0, 8))
            ttk.Label(table, text=str(value_a), wraplength=260, justify="left").grid(
                row=r, column=1, sticky="w", padx=8, pady=3)
            ttk.Label(table, text=str(value_b), wraplength=260, justify="left").grid(
                row=r, column=2, sticky="w", padx=8, pady=3)
            row_counter[0] += 1

        cat_a = translate_category_name(recipe_a.get("category", "Autre"))
        cat_b = translate_category_name(recipe_b.get("category", "Autre"))
        field_line(t("compare_field_category"), cat_a, cat_b)

        fav_a = t("compare_yes") if recipe_a.get("favorite") else t("compare_no")
        fav_b = t("compare_yes") if recipe_b.get("favorite") else t("compare_no")
        field_line(t("compare_field_favorite"), fav_a, fav_b)

        field_line(t("compare_field_rating"), rating_stars(recipe_a.get("rating", 0)), rating_stars(recipe_b.get("rating", 0)))
        field_line(
            t("compare_field_difficulty"),
            translate_difficulty_name(recipe_a.get("difficulty")) or "—",
            translate_difficulty_name(recipe_b.get("difficulty")) or "—"
        )

        prep_a = recipe_a.get("prep_time") or "—"
        prep_b = recipe_b.get("prep_time") or "—"
        field_line(t("compare_field_prep"), f"{prep_a} min" if prep_a != "—" else "—", f"{prep_b} min" if prep_b != "—" else "—")

        cook_a = recipe_a.get("cook_time") or "—"
        cook_b = recipe_b.get("cook_time") or "—"
        field_line(t("compare_field_cook"), f"{cook_a} min" if cook_a != "—" else "—", f"{cook_b} min" if cook_b != "—" else "—")

        def total_minutes(r):
            try:
                return float(r.get("prep_time") or 0) + float(r.get("cook_time") or 0)
            except (TypeError, ValueError):
                return 0
        total_a, total_b = total_minutes(recipe_a), total_minutes(recipe_b)
        field_line(t("compare_field_total_time"), f"{total_a:.0f} min" if total_a else "—", f"{total_b:.0f} min" if total_b else "—")

        field_line(
            t("compare_field_cooked"),
            t("compare_times_suffix", count=recipe_a.get('times_cooked', 0)),
            t("compare_times_suffix", count=recipe_b.get('times_cooked', 0))
        )

        persons_a = recipe_a.get("default_persons", 1) or 1
        persons_b = recipe_b.get("default_persons", 1) or 1
        cost_a, cost_known_a, cost_total_a = compute_recipe_cost(recipe_a, persons_a)
        cost_b, cost_known_b, cost_total_b = compute_recipe_cost(recipe_b, persons_b)
        cost_display_a = f"{cost_a:.2f} € ({persons_a} p.)" if cost_known_a else "—"
        cost_display_b = f"{cost_b:.2f} € ({persons_b} p.)" if cost_known_b else "—"
        field_line(t("compare_field_cost"), cost_display_a, cost_display_b)

        nutri_a, nutri_known_a, nutri_total_a = compute_recipe_nutrition(recipe_a, persons_a)
        nutri_b, nutri_known_b, nutri_total_b = compute_recipe_nutrition(recipe_b, persons_b)
        kcal_display_a = f"{nutri_a['kcal']:.0f} kcal ({persons_a} p.)" if nutri_known_a else "—"
        kcal_display_b = f"{nutri_b['kcal']:.0f} kcal ({persons_b} p.)" if nutri_known_b else "—"
        field_line(t("compare_field_nutrition"), kcal_display_a, kcal_display_b)

        ing_names_a = {ing["name"].strip().lower(): ing["name"] for ing in recipe_a["ingredients"]}
        ing_names_b = {ing["name"].strip().lower(): ing["name"] for ing in recipe_b["ingredients"]}
        field_line(t("compare_field_ingredient_count"), len(ing_names_a), len(ing_names_b))

        common_keys = set(ing_names_a) & set(ing_names_b)
        only_a_keys = set(ing_names_a) - set(ing_names_b)
        only_b_keys = set(ing_names_b) - set(ing_names_a)

        # ---- Ingrédients : trois colonnes côte à côte, elles aussi bien
        # alignées (communs / uniquement A / uniquement B) ----
        ing_frame = ttk.Frame(self.result_frame)
        ing_frame.pack(fill="x")
        ing_frame.columnconfigure(0, weight=1)
        ing_frame.columnconfigure(1, weight=1)
        ing_frame.columnconfigure(2, weight=1)

        def ingredient_column(parent, col, title, keys, names_map):
            col_frame = ttk.Frame(parent)
            col_frame.grid(row=0, column=col, sticky="nw", padx=8)
            ttk.Label(col_frame, text=title, font=("Segoe UI", sf(9), "bold"),
                      foreground=COLOR_ACCENT_DARK, wraplength=220, justify="left").pack(anchor="w", pady=(0, 4))
            if keys:
                for key in sorted(keys, key=ingredient_sort_key):
                    ttk.Label(col_frame, text=f"• {translate_ingredient_name(names_map[key]).capitalize()}",
                              wraplength=220, justify="left").pack(anchor="w", pady=1)
            else:
                ttk.Label(col_frame, text=t("compare_none"), foreground=COLOR_TEXT_MUTED).pack(anchor="w")

        ingredient_column(ing_frame, 0, t("compare_common_ingredients", count=len(common_keys)), common_keys, ing_names_a)
        ingredient_column(ing_frame, 1, t("compare_only_a", name=name_a, count=len(only_a_keys)), only_a_keys, ing_names_a)
        ingredient_column(ing_frame, 2, t("compare_only_b", name=name_b, count=len(only_b_keys)), only_b_keys, ing_names_b)


class StatisticsWindow(tk.Toplevel):
    """Fenêtre affichant des statistiques simples sur les recettes."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title(t("stats_title"))
        self.geometry(f"{gs(560)}x{gs(820)}")
        self.minsize(gs(480), gs(500))
        self.resizable(True, True)
        self.grab_set()

        text_frame = ttk.Frame(self)
        text_frame.pack(fill="both", expand=True, padx=15, pady=(15, 5))
        text = tk.Text(text_frame, wrap="word", height=22, font=("Segoe UI", sf(9)))
        text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=text_scrollbar.set)
        text.pack(side="left", fill="both", expand=True)
        text_scrollbar.pack(side="right", fill="y")

        recipes = self.app.recipes
        total = len(recipes)
        text.insert(tk.END, t("stats_heading"))
        text.insert(tk.END, t("stats_total_recipes", count=total))

        cat_counts = {}
        for r in recipes:
            cat = r.get("category", "Autre")
            cat_counts[cat] = cat_counts.get(cat, 0) + 1
        text.insert(tk.END, t("stats_by_category"))
        for cat in RecipeFormWindow.CATEGORY_OPTIONS:
            if cat in cat_counts:
                text.insert(tk.END, t("stats_category_line", category=translate_category_name(cat), count=cat_counts[cat]))
        text.insert(tk.END, "\n")

        diff_counts = {}
        for r in recipes:
            diff = r.get("difficulty") or t("stats_difficulty_unspecified")
            diff_counts[diff] = diff_counts.get(diff, 0) + 1
        text.insert(tk.END, t("stats_by_difficulty"))
        for diff in ["Facile", "Moyen", "Difficile", t("stats_difficulty_unspecified")]:
            if diff in diff_counts:
                text.insert(tk.END, t("stats_difficulty_line", difficulty=translate_difficulty_name(diff), count=diff_counts[diff]))
        text.insert(tk.END, "\n")

        fav_count = sum(1 for r in recipes if r.get("favorite"))
        text.insert(tk.END, t("stats_favorites_count", count=fav_count))

        rated = [r.get("rating", 0) for r in recipes if r.get("rating")]
        if rated:
            avg_rating = sum(rated) / len(rated)
            text.insert(tk.END, t("stats_avg_rating", avg=f"{avg_rating:.1f}", count=len(rated)))
        else:
            text.insert(tk.END, t("stats_no_rated_recipe"))

        best_rated = [r for r in recipes if r.get("rating", 0) == 5]
        if best_rated:
            text.insert(tk.END, t("stats_five_star_heading"))
            for r in best_rated[:10]:
                text.insert(tk.END, t("stats_recipe_line", name=r['name']))
            text.insert(tk.END, "\n")

        cooked = sorted(
            (r for r in recipes if r.get("times_cooked", 0) > 0),
            key=lambda r: r.get("times_cooked", 0), reverse=True
        )
        text.insert(tk.END, t("stats_most_cooked_heading"))
        if cooked:
            for r in cooked[:10]:
                text.insert(tk.END, t("stats_cooked_line", name=r['name'], count=r['times_cooked']))
        else:
            text.insert(tk.END, t("stats_none_cooked_yet"))
        text.insert(tk.END, "\n")

        tag_counts = {}
        for r in recipes:
            for tag in r.get("tags", []):
                tag_counts[tag] = tag_counts.get(tag, 0) + 1
        if tag_counts:
            text.insert(tk.END, t("stats_most_used_tags_heading"))
            for tag, count in sorted(tag_counts.items(), key=lambda pair: -pair[1])[:10]:
                text.insert(tk.END, t("stats_tag_line", tag=tag, count=count))
            text.insert(tk.END, "\n")

        # ---- Recettes oubliées ----
        never_cooked = [r for r in recipes if r.get("times_cooked", 0) == 0]
        text.insert(tk.END, t("stats_never_cooked_heading"))
        if never_cooked:
            shown = never_cooked[:15]
            for r in shown:
                text.insert(tk.END, t("stats_recipe_line", name=r['name']))
            remaining = len(never_cooked) - len(shown)
            if remaining > 0:
                text.insert(tk.END, t("stats_and_others", count=remaining))
        else:
            text.insert(tk.END, t("stats_all_cooked"))
        text.insert(tk.END, "\n")

        stale_cutoff_days = 90
        now = datetime.now()
        stale = []
        for r in recipes:
            cooked_dates = r.get("cooked_dates") or []
            if not cooked_dates:
                continue
            try:
                last_date = max(datetime.fromisoformat(d) for d in cooked_dates)
            except ValueError:
                continue
            if (now - last_date).days >= stale_cutoff_days:
                stale.append((r, (now - last_date).days))
        stale.sort(key=lambda pair: -pair[1])
        text.insert(tk.END, t("stats_stale_heading", days=stale_cutoff_days))
        if stale:
            for r, days in stale[:15]:
                text.insert(tk.END, t("stats_stale_line", name=r['name'], days=days))
        else:
            text.insert(tk.END, t("stats_no_stale_recipe"))
        text.insert(tk.END, "\n")

        # ---- Coût moyen ----
        costs_per_person = []
        for r in recipes:
            persons = r.get("default_persons", 1) or 1
            cost, known, _ = compute_recipe_cost(r, persons)
            if known > 0:
                costs_per_person.append(cost / persons)
        text.insert(tk.END, t("stats_avg_cost_heading"))
        if costs_per_person:
            avg_cost = sum(costs_per_person) / len(costs_per_person)
            without_price = total - len(costs_per_person)
            text.insert(
                tk.END,
                t("stats_avg_cost_line", avg=f"{avg_cost:.2f}", count=len(costs_per_person), without_price=without_price)
            )
        else:
            text.insert(tk.END, t("stats_no_priced_recipe"))
        text.insert(tk.END, "\n")

        # ---- Calories moyennes ----
        kcal_per_person = []
        for r in recipes:
            persons = r.get("default_persons", 1) or 1
            nutrition, known, _ = compute_recipe_nutrition(r, persons)
            if known > 0:
                kcal_per_person.append(nutrition["kcal"] / persons)
        text.insert(tk.END, t("stats_avg_kcal_heading"))
        if kcal_per_person:
            avg_kcal = sum(kcal_per_person) / len(kcal_per_person)
            text.insert(
                tk.END,
                t("stats_avg_kcal_line", avg=f"{avg_kcal:.0f}", count=len(kcal_per_person))
            )
        else:
            text.insert(tk.END, t("stats_no_recognized_recipe"))

        text.config(state="disabled")

        # ---- Graphique d'évolution mensuelle des recettes cuisinées ----
        ttk.Label(self, text=t("stats_monthly_chart_title"),
                  font=("Segoe UI", sf(10), "bold")).pack(pady=(8, 2))
        chart_canvas = tk.Canvas(self, height=180, background=COLOR_CARD, highlightthickness=1,
                                  highlightbackground=COLOR_BORDER)
        chart_canvas.pack(fill="x", padx=15, pady=(0, 15))
        self.after(50, lambda: self._draw_monthly_chart(chart_canvas, recipes))

        # ---- Calendrier visuel (façon contributions GitHub) ----
        ttk.Label(self, text=t("stats_heatmap_title"),
                  font=("Segoe UI", sf(10), "bold")).pack(pady=(0, 2))
        heatmap_container = ttk.Frame(self)
        heatmap_container.pack(fill="x", padx=15, pady=(0, 5))
        heatmap_canvas = tk.Canvas(heatmap_container, height=120, background=COLOR_CARD,
                                    highlightthickness=1, highlightbackground=COLOR_BORDER)
        heatmap_hscroll = ttk.Scrollbar(heatmap_container, orient="horizontal", command=heatmap_canvas.xview)
        heatmap_canvas.configure(xscrollcommand=heatmap_hscroll.set)
        heatmap_canvas.pack(fill="x")
        heatmap_hscroll.pack(fill="x")
        ttk.Label(self, text=t("stats_heatmap_legend"), font=("Segoe UI", sf(8)),
                  foreground=COLOR_TEXT_MUTED).pack(pady=(0, 15))
        self.after(50, lambda: self._draw_cooking_heatmap(heatmap_canvas, recipes))

    def _draw_cooking_heatmap(self, canvas, recipes):
        """Dessine un calendrier visuel façon « contributions GitHub » des
        jours où au moins une recette a été cuisinée, sur les 12 derniers
        mois (toutes recettes confondues)."""
        counts = {}
        for r in recipes:
            for date_str in r.get("cooked_dates") or []:
                counts[date_str] = counts.get(date_str, 0) + 1

        today = datetime.now().date()
        start = today - timedelta(days=today.weekday())  # lundi de la semaine actuelle
        start = start - timedelta(weeks=52)

        cell_size, cell_gap = 12, 3
        margin_left, margin_top = 22, 16
        max_count = max(counts.values(), default=0) or 1

        def color_for_count(count):
            if count == 0:
                return COLOR_BORDER
            ratio = count / max_count
            if ratio < 0.34:
                return COLOR_ACCENT_LIGHT
            elif ratio < 0.67:
                return COLOR_ACCENT
            return COLOR_ACCENT_DARK

        canvas.delete("all")
        day_labels = t("stats_day_labels").split(",")
        for row in range(7):
            canvas.create_text(
                margin_left - 12, margin_top + row * (cell_size + cell_gap) + cell_size / 2,
                text=day_labels[row], font=("Segoe UI", sf(7)), fill=COLOR_TEXT_MUTED
            )

        month_labels_fr = t("stats_month_labels_short").split(",")
        current = start
        week_index = 0
        last_month = None
        while current <= today:
            if current.day <= 7 and current.month != last_month:
                canvas.create_text(
                    margin_left + week_index * (cell_size + cell_gap), margin_top - 9,
                    text=month_labels_fr[current.month - 1], font=("Segoe UI", sf(7)),
                    fill=COLOR_TEXT_MUTED, anchor="w"
                )
                last_month = current.month

            row = current.weekday()
            date_str = current.strftime("%Y-%m-%d")
            count = counts.get(date_str, 0)
            x0 = margin_left + week_index * (cell_size + cell_gap)
            y0 = margin_top + row * (cell_size + cell_gap)
            canvas.create_rectangle(
                x0, y0, x0 + cell_size, y0 + cell_size,
                fill=color_for_count(count), outline=COLOR_BG
            )
            if row == 6:
                week_index += 1
            current += timedelta(days=1)

        canvas.configure(scrollregion=canvas.bbox("all"))
        # Fait défiler la vue tout à droite (les semaines les plus récentes),
        # pour que l'utilisateur voie d'emblée les jours les plus récents.
        canvas.xview_moveto(1.0)

    def _draw_monthly_chart(self, canvas, recipes):
        """Dessine un histogramme simple (sans dépendance externe) du nombre
        de recettes cuisinées par mois, sur les 12 derniers mois."""
        now = datetime.now()
        months = []
        for i in range(11, -1, -1):
            year, month = now.year, now.month - i
            while month <= 0:
                month += 12
                year -= 1
            months.append((year, month))

        counts = {ym: 0 for ym in months}
        for r in recipes:
            for date_str in r.get("cooked_dates") or []:
                try:
                    d = datetime.fromisoformat(date_str)
                except (ValueError, TypeError):
                    continue
                ym = (d.year, d.month)
                if ym in counts:
                    counts[ym] += 1

        canvas.update_idletasks()
        width = max(canvas.winfo_width(), 480)
        height = 180
        margin_bottom, margin_top = 28, 14
        chart_height = height - margin_bottom - margin_top
        max_count = max(max(counts.values(), default=0), 1)
        bar_width = (width - 20) / len(months)

        month_labels_fr = t("stats_month_labels_lower").split(",")

        canvas.delete("all")
        for i, ym in enumerate(months):
            count = counts[ym]
            bar_height = (count / max_count) * chart_height
            x0 = 10 + i * bar_width + 3
            x1 = 10 + (i + 1) * bar_width - 3
            y1 = height - margin_bottom
            y0 = y1 - bar_height
            canvas.create_rectangle(x0, y0, x1, y1, fill=COLOR_ACCENT, outline="")
            if count > 0:
                canvas.create_text((x0 + x1) / 2, y0 - 8, text=str(count),
                                    font=("Segoe UI", sf(8)), fill=COLOR_TEXT)
            canvas.create_text((x0 + x1) / 2, y1 + 12, text=month_labels_fr[ym[1] - 1],
                                font=("Segoe UI", sf(7)), fill=COLOR_TEXT_MUTED)


if __name__ == "__main__":
    app = App()
    app.mainloop()
