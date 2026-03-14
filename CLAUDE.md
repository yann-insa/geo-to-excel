# CLAUDE.md — GEO → Excel Converter

## Résumé du projet

Application desktop Python/tkinter qui convertit des fichiers géographiques (KMZ, KML, DXF) en fichier Excel (.xlsx) au format **« exportations ouvrages » réseau fibre** utilisé en interne pour la gestion d'infrastructure fibre optique sur le réseau de la Ville de Lyon (VDL).

Le contexte métier est la gestion de réseau fibre optique pour la Métropole de Lyon. Les fichiers DXF viennent d'AutoCAD/QGIS et contiennent des tracés de câbles FO. Les fichiers KMZ/KML viennent de Google Earth.

---

## Architecture

Le code principal est dans `src/geo_to_excel/app.py` (~1260 lignes). Fichier unique pour la logique métier + IHM.

### Arborescence du projet :

```
├── src/
│   └── geo_to_excel/
│       ├── __init__.py         # Version du package
│       └── app.py              # Script principal (tout-en-un)
├── scripts/
│   ├── build_exe.bat           # Build exe (CMD)
│   └── build_exe.ps1           # Build exe (PowerShell)
├── tests/
│   └── __init__.py
├── .gitignore
├── CLAUDE.md                   # Ce fichier (contexte pour Claude Code)
├── LICENSE
├── pyproject.toml              # Configuration projet Python
└── README.md
```

### Sections du code `app.py` (dans l'ordre) :

```
Ligne   Section
──────  ─────────────────────────────────────────────────
1-15    Docstring, metadata
17-24   Imports (tkinter, xml, zipfile, math, os, sys, threading)
26-77   Constantes: COLORS (thème UI dark), FONT_FAMILY, PROJECTIONS (dict label→EPSG)
79-215  Détection automatique de projection (detect_projection, detect_projection_display)
218-280 Utilitaires géométriques (haversine, calc_length_km, WKT, COLOR_NAME_MAP, ACI_COLORS)
282-310 Parseur KML: _kml_color_to_hex, _parse_kml_coordinates
311-418 parse_kml_file() — parsing KML/KMZ complet
420-644 parse_dxf_file() — parsing DXF avec reprojection pyproj
646-700 Reverse geocoding Lyon (_detect_lyon_arrondissement, _extract_fo_from_layer, GESTIONNAIRES)
701-810 build_xlsx() — génération Excel avec openpyxl
812-845 convert() — pipeline unifié d'entrée
847-855 _CancelledError — exception pour annulation utilisateur
856-1248 class App(tk.Tk) — interface graphique complète
1250-1260 main() et __main__
```

### Fonctions clés :

| Fonction | Rôle |
|----------|------|
| `detect_projection(dxf_path)` | Analyse les plages X/Y des entités DXF pour identifier la projection (Lambert-93, CC42-50, Lambert I/II/III, UTM, WGS84). Retourne `(epsg, label, confiance)`. |
| `parse_kml_file(path, progress_cb)` | Parse KML/KMZ. Gère les KMZ corrompus (BadZipFile), KML mal formés (ParseError). Extrait LineString et Point. |
| `parse_dxf_file(path, epsg, progress_cb)` | Parse DXF via `ezdxf`. Reprojette toutes les coordonnées vers WGS84 via `pyproj.Transformer`. Supporte LINE, LWPOLYLINE, POLYLINE, SPLINE, POINT, CIRCLE, ARC, ELLIPSE, INSERT. |
| `build_xlsx(rows, path, gestionnaire, progress_cb)` | Génère le fichier Excel au format exact « exportations ouvrages ». |
| `convert(input, output, epsg, gestionnaire, progress_cb)` | Pipeline unifié : détecte le format → parse → build xlsx. |
| `_detect_lyon_arrondissement(lon, lat)` | Reverse geocoding simplifié : détermine l'arrondissement de Lyon par distance euclidienne aux centres des 9 arrondissements. |
| `_extract_fo_from_layer(layer_name)` | Extrait la capacité FO du nom de layer DXF. Ex: `_012_FO_TRACE` → `12FO`, `_144_FO_TRACE` → `144FO`. |

---

## Format de sortie Excel (critique)

Le fichier Excel doit EXACTEMENT respecter ce format. C'est le format d'import de l'outil de gestion réseau fibre.

### Structure :

- **Feuille** : `Réseau fibre`
- **Ligne 1** : 2 cellules fusionnées avec fond gris `#B3AEAD`, police Calibri gras noir centré
  - `A1:T1` = `OUVRAGES`
  - `U1:AC1` = `GC`
- **Ligne 2** : 29 colonnes d'en-tête (même style)
- **Ligne 3+** : données

### Colonnes (A=1 à AC=29) :

| Col | Lettre | Nom | Remplissage |
|-----|--------|-----|-------------|
| 1 | A | ID | **VIDE** |
| 2 | B | SECTEUR | Toujours `LYON VP` |
| 3 | C | VILLE | Auto-détecté : `LYON - 8E ARRONDISSEMENT` (par reverse geocoding des coordonnées) |
| 4 | D | ZONE | `Non renseigné` |
| 5 | E | NUMERO | Auto-incrémenté : `RF0001`, `RF0002`... |
| 6 | F | PARENT | vide |
| 7 | G | PARENT CLASSE D'OUVRAGES | vide |
| 8 | H | MULTIVILLES | vide |
| 9 | I | LATITUDE | Rempli uniquement pour les POINTs |
| 10 | J | LONGITUDE | Rempli uniquement pour les POINTs |
| 11 | K | COORDS WKT | `LINESTRING Z (lon lat alt,...)` ou `POINT Z (lon lat alt)` |
| 12 | L | COULEUR | Nom français si connu (ROUGE, BLEU...) sinon vide |
| 13 | M | LONGUEUR | En km, calculé par haversine |
| 14 | N | PATTERN | **VIDE** |
| 15 | O | REMARQUES | **VIDE** |
| 16 | P | FICHIERS | vide |
| 17 | Q | IMAGES | vide |
| 18 | R | ICONE | vide |
| 19 | S | CONDITION ICONE | vide |
| 20 | T | SPECIFIQUE | vide |
| 21 | U | ELEMENT ID | **VIDE** |
| 22 | V | NOM ELEMENT | Toujours `Fibre` |
| 23 | W | GESTIONNAIRE | Choisi par l'utilisateur dans l'IHM (liste déroulante) |
| 24 | X | COMMENTAIRE | vide |
| 25 | Y | MODE DE POSE | vide |
| 26 | Z | COULEUR PROPRIETAIRE | **VIDE** |
| 27 | AA | LAYER | Capacité FO extraite du layer DXF : `12FO`, `24FO`, `144FO`... |
| 28 | AB | COUCHE | Idem que LAYER |
| 29 | AC | COLOR | Code hex couleur `#FF0000` |

### Colonnes volontairement VIDES (ne pas remplir) :
- ID (col A), PATTERN (col N), REMARQUES (col O), ELEMENT ID (col U), COULEUR PROPRIETAIRE (col Z)

---

## Détection automatique de projection

Fonctionne uniquement pour la **France métropolitaine**. Basée sur les plages de coordonnées X/Y médianes des entités du DXF :

| Condition (X_med, Y_med) | Projection |
|--------------------------|------------|
| X 50k-1.3M, Y 5.9M-7.3M | Lambert-93 (EPSG:2154) |
| X 50k-1.3M, Y 1.5M-2.8M | Lambert II étendu (EPSG:27572) |
| X 50k-1.3M, Y 100k-1.5M | Lambert I (EPSG:27571) |
| X 50k-1.3M, Y 2.8M-4.1M | Lambert III (EPSG:27573) |
| X -700k-700k, Y 4M-5.4M | Lambert IV Corse (EPSG:27574) |
| X 900k-2.4M, Y variable | CC42-CC50 (zone = (Y_med - 1.2M) / 1M + 42) |
| X -10 à 15, Y 40 à 52 | WGS84 déjà en lat/lon (EPSG:4326) |
| Y 4.4M-5.9M, X 200k-800k | UTM 30N/31N/32N (validé par reprojection) |

Pour les **Coniques Conformes (CC)**, le numéro de zone est déduit du Y médian :
- CC42 (Y centre ~1.2M) → Perpignan/Corse
- CC43 (Y centre ~2.2M) → Montpellier/Marseille
- CC44 (Y centre ~3.2M) → Toulouse/Nice
- CC45 (Y centre ~4.2M) → Bordeaux/Grenoble
- **CC46 (Y centre ~5.2M) → Lyon/Saint-Étienne** ← le plus fréquent pour ce projet
- CC47 (Y centre ~6.2M) → Paris/Nantes/Tours
- CC48 (Y centre ~7.2M) → Paris/Lille/Rennes
- CC49 (Y centre ~8.2M) → Lille/Calais
- CC50 (Y centre ~9.2M) → Dunkerque

---

## Reverse geocoding arrondissements de Lyon

Basé sur la distance euclidienne aux centres approximatifs (WGS84) des 9 arrondissements :

```python
_LYON_ARR_CENTERS = {
    1: (4.8320, 45.7690),  # Terreaux/Pentes
    2: (4.8280, 45.7560),  # Presqu'île/Bellecour
    3: (4.8570, 45.7580),  # Part-Dieu/Villette
    4: (4.8270, 45.7740),  # Croix-Rousse
    5: (4.8130, 45.7590),  # Vieux-Lyon/Point du Jour
    6: (4.8510, 45.7710),  # Brotteaux/Tête d'Or
    7: (4.8410, 45.7400),  # Gerland/Jean Macé
    8: (4.8720, 45.7330),  # Monplaisir/Mermoz  ← zone principale du projet
    9: (4.8020, 45.7790),  # Vaise/Duchère
}
```

Ne fonctionne que si les coordonnées WGS84 tombent dans la zone Lyon (lon 4.75-4.95, lat 45.70-45.80). Sinon retourne "Non renseigné".

**Limitation connue** : le reverse geocoding par distance au centre est approximatif. Pour les entités aux frontières entre arrondissements, le résultat peut être incorrect. Une amélioration serait d'utiliser les vraies limites administratives (polygones shapefile/geojson de l'IGN).

---

## Liste des gestionnaires

Les gestionnaires disponibles dans la liste déroulante IHM :

```python
GESTIONNAIRES = [
    "CRITER",           # Réseau CRITER (gestion de trafic)
    "VDL-Aérien",       # Ville de Lyon - câbles aériens
    "VDL-Souterrain",   # Ville de Lyon - câbles souterrains
    "VDL-Divers",       # Ville de Lyon - divers
    "VDL-Rocade",       # Ville de Lyon - rocade
    "Eclairage public", # Réseau d'éclairage public
]
```

---

## Extraction de la capacité FO depuis les layers DXF

Les fichiers DXF utilisent des noms de layer qui encodent la capacité en fibres optiques :

| Layer DXF | Extraction | Résultat |
|-----------|-----------|----------|
| `_012_FO_TRACE` | regex `_?(\d+)_?FO` | `12FO` |
| `_024_FO_TRACE` | | `24FO` |
| `_144_FO_TRACE` | | `144FO` |
| `_096_FO_BACKBONE` | | `96FO` |

Si le pattern ne matche pas, le nom de layer brut est utilisé tel quel.

---

## Interface graphique (tkinter)

### Thème
- Dark theme avec fond `#0F1117`, accent violet `#6C63FF`
- Police : Segoe UI (Windows), SF Pro Display (macOS), Ubuntu (Linux)
- Style des combobox personnalisé via ttk.Style "Dark.TCombobox"

### Éléments de l'IHM (de haut en bas)
1. **Titre** : "◆ GEO → Excel"
2. **Zone de sélection fichier** : bouton "Parcourir…", affiche nom/taille/chemin
3. **Bloc projection** (visible uniquement pour DXF) : label de détection auto + combobox projection + champ EPSG custom
4. **Bloc gestionnaire** (toujours visible après sélection fichier) : combobox avec la liste GESTIONNAIRES
5. **Barre de progression** : canvas 8px avec couleur accent/success
6. **Boutons** : "▶ Convertir en Excel" + "■ Annuler" côte à côte
7. **Label status** : messages succès/erreur/annulation
8. **Footer** : version

### Mécanisme d'annulation
- `threading.Event` (`_cancel_event`) vérifié dans le progress callback
- Lève `_CancelledError` qui est catchée dans `_do_convert`
- Le fichier Excel partiel est supprimé si l'annulation intervient

### Bug Python 3.13 corrigé
Les lambdas dans `self.after(0, lambda: ...)` ne doivent PAS capturer directement des variables de bloc `except` (Python 3.13 libère `exc` à la sortie du bloc). Solution : pré-évaluer dans une variable locale avant le lambda, ou utiliser des paramètres par défaut `lambda m=msg: ...`.

---

## Dépendances

| Package | Version testée | Rôle |
|---------|---------------|------|
| `openpyxl` | 3.1+ | Lecture/écriture Excel .xlsx |
| `ezdxf` | 1.4+ | Parsing fichiers DXF (AutoCAD) |
| `pyproj` | 3.7+ | Reprojection CRS (Lambert→WGS84 etc.) |
| `tkinter` | stdlib | Interface graphique (inclus avec Python) |
| `pyinstaller` | 6+ | Build executable Windows (optionnel, pour le build uniquement) |

### Installation

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate

pip install openpyxl ezdxf pyproj
```

---

## Build exécutable Windows

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name GEO_to_Excel \
    --hidden-import=openpyxl \
    --hidden-import=ezdxf \
    --hidden-import=pyproj \
    --hidden-import=pyproj.database \
    --hidden-import=pyproj._datadir \
    --hidden-import=pyproj._crs \
    --hidden-import=pyproj._transformer \
    --hidden-import=pyproj._geod \
    --hidden-import=certifi \
    --collect-data=pyproj \
    --collect-data=certifi \
    --noconfirm --clean \
    geo_to_excel.py
```

L'exe résultant est dans `dist/GEO_to_Excel.exe` (~80-120 Mo). Autonome, ne nécessite pas Python.

Scripts de build fournis : `build_exe.bat` (CMD) et `build_exe.ps1` (PowerShell).

---

## Tests à effectuer après modification

### Tests fonctionnels critiques :

1. **haversine** : même point = 0, Paris→Lyon ≈ 392 km
2. **calc_length_km** : liste vide = 0, 1 point = 0
3. **_kml_color_to_hex** : `ff0000ff` → `#FF0000`, chaînes invalides → `""`
4. **_parse_kml_coordinates** : vide, données corrompues, point-virgule final, sans altitude
5. **aci_to_hex** : index 0/-1/256 → `""`, index 1 → `#FF0000`
6. **parse_kml_file** : fichier KMZ réel, KML corrompu (ParseError), KMZ corrompu (BadZipFile)
7. **parse_dxf_file** : Lambert-93→WGS84 (vérifier lon 4.8-4.9, lat 45.7-45.8 pour Lyon), WGS84 passthrough, DXF corrompu
8. **detect_projection** : Lambert-93 (X~845k,Y~6520k), CC46 (X~1843k,Y~5175k), WGS84 (lon/lat directement), DXF vide → None
9. **build_xlsx** : vérifier en-têtes fusionnées, SECTEUR=LYON VP, VILLE=arrondissement, LAYER=12FO/24FO, GESTIONNAIRE rempli, ID/PATTERN/REMARQUES/ELEMENT ID/COULEUR PROPRIETAIRE tous vides
10. **convert pipeline** : KMZ→XLSX, DXF→XLSX, DXF sans EPSG → erreur, format .shp → erreur
11. **_detect_lyon_arrondissement** : coordonnées Mermoz → "LYON - 8E ARRONDISSEMENT"
12. **_extract_fo_from_layer** : `_012_FO_TRACE` → `12FO`

### Fichiers de test disponibles :
- `mermoz_nord_fibre.dxf` — DXF réel en CC46, 20 LWPOLYLINE, layers _012_FO/_024_FO/_144_FO
- `Backbone_96FO_Mermoz_Pavillon.kmz` — KMZ réel, 1 LineString, 61 points, ~5.6 km

---

## Améliorations possibles (backlog)

1. **Reverse geocoding plus précis** : utiliser les polygones administratifs IGN des arrondissements de Lyon au lieu de la distance au centre
2. **Support d'autres villes** : étendre le reverse geocoding au-delà de Lyon (Villeurbanne, Vénissieux, etc.)
3. **Import shapefile (.shp)** : ajouter le support via `pyshp` ou `fiona`
4. **Import GeoJSON** : format de plus en plus courant
5. **Export multi-feuilles** : une feuille par layer/capacité FO
6. **Drag & drop** : supporter le glisser-déposer de fichiers sur la fenêtre (tkinterdnd2)
7. **Sauvegarde des préférences** : mémoriser le dernier gestionnaire et la dernière projection choisie
8. **Mode batch** : convertir plusieurs fichiers d'un coup
9. **Logging** : remplacer les messages console par un vrai module logging
10. **Tests unitaires** : créer un fichier `test_geo_to_excel.py` avec pytest

---

## Conventions de code

- **Python 3.10+** requis (f-strings, type hints optionnels)
- **Pas de classes pour la logique métier** : tout est en fonctions pures sauf l'IHM (class App)
- **Progress callback** : signature `progress_cb(message: str, pourcentage: int)`
- **Segments** : chaque entité géographique est un dict avec les clés : `name`, `type`, `layer`, `coords` (list de tuples (lon,lat,alt)), `wkt`, `length`, `color`, et optionnellement `latitude`/`longitude` pour les points
- **Annulation** : vérifiée dans le progress callback, lève `_CancelledError`
- **Lambdas tkinter** : ne JAMAIS capturer directement une variable `except as exc` dans un lambda — Python 3.13 crash. Toujours pré-évaluer.

---

## Fichiers du projet

```
src/geo_to_excel/app.py    # Script principal (tout-en-un)
src/geo_to_excel/__init__.py
scripts/build_exe.bat      # Script batch Windows pour builder l'exe
scripts/build_exe.ps1      # Script PowerShell pour builder l'exe
tests/                     # Tests (à implémenter)
pyproject.toml             # Configuration projet Python
README.md                  # Documentation utilisateur
CLAUDE.md                  # Ce fichier (contexte pour Claude Code)
LICENSE                    # Licence MIT
.gitignore
```
