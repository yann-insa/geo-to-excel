# GEO → Excel Converter

Convertisseur KMZ / KML / DXF vers le format Excel « exportations ouvrages » réseau fibre.

## Fonctionnalités

- **Formats supportés** : KMZ, KML, DXF
- **Détection automatique de projection** (France métropolitaine) : Lambert-93, CC42–CC50, Lambert I/II/III/IV, UTM, WGS84
- **Reprojection automatique** vers WGS84 (lat/lon)
- **Bouton Annuler** pour stopper les conversions longues
- **Fichier de sortie** au même emplacement que le fichier d'entrée

## Structure du projet

```
├── src/
│   └── geo_to_excel/
│       ├── __init__.py
│       └── app.py              # Script principal
├── scripts/
│   ├── build_exe.bat           # Build exe (CMD)
│   └── build_exe.ps1           # Build exe (PowerShell)
├── tests/
│   └── __init__.py
├── .gitignore
├── CLAUDE.md
├── LICENSE
├── pyproject.toml
└── README.md
```

## Installation

### Option 1 : Exécutable Windows (sans Python)

1. Clonez le dépôt
2. Lancez `scripts\build_exe.bat` (CMD) ou `scripts\build_exe.ps1` (PowerShell)
3. L'exécutable sera créé dans `dist\GEO_to_Excel.exe`
4. Copiez `GEO_to_Excel.exe` où vous voulez — aucune dépendance requise

**Prérequis pour le build uniquement** : Python 3.10+ installé avec pip.

### Option 2 : Script Python direct

```bash
# Créer un environnement virtuel
python -m venv .venv

# Activer (Windows)
.venv\Scripts\activate

# Installer les dépendances
pip install -e .

# Lancer
geo-to-excel
# ou
python src/geo_to_excel/app.py
```

## Projections détectées automatiquement

| Système | EPSG | Zone typique |
|---------|------|-------------|
| Lambert-93 | 2154 | France entière |
| CC42 | 3942 | Perpignan / Corse |
| CC43 | 3943 | Montpellier / Marseille |
| CC44 | 3944 | Toulouse / Nice |
| CC45 | 3945 | Bordeaux / Grenoble |
| CC46 | 3946 | Lyon / Saint-Étienne |
| CC47 | 3947 | Paris / Nantes / Tours |
| CC48 | 3948 | Paris / Lille / Rennes |
| CC49 | 3949 | Lille / Calais |
| CC50 | 3950 | Dunkerque |
| Lambert II étendu | 27572 | France (ancien) |
| UTM 31N | 32631 | France centre |

## Licence

MIT
