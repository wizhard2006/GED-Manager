# GED-Manager

Gestion Électronique de Documents — Classification automatique par OCR et apprentissage.

## Fonctionnalités

- **Acquisition** : importer un fichier PDF existant OU scanner via CANON DR-C125
- **OCR** : extraction automatique du texte (Tesseract, supporte PDF scannés)
- **Classification** : détection de l'origine/société par mots-clés
- **Apprentissage** : mémorise vos choix → classement automatique les fois suivantes
- **Quarantaine** : documents non identifiés mis de côté pour traitement manuel
- **Recherche** : retrouver un document dans l'historique
- **Logs** : trace de chaque action (fichier, date, destination)

## Installation

### 1. Prérequis

- **Python 3.10+** (testé sur 3.13)  
- **Tesseract OCR** : https://github.com/UB-Mannheim/tesseract/wiki  
  → Installer dans `C:\Program Files\Tesseract-OCR\`

### 2. Cloner le projet

```cmd
git clone https://github.com/wizhard2006/GED-Manager.git
cd GED-Manager
```

### 3. Installer les dépendances

```cmd
pip install -r requirements.txt
```

### 4. Lancer l'application

```cmd
python main.py
```

## Configuration

Au premier lancement, allez dans **Paramètres** pour définir :

| Paramètre | Exemple | Description |
|-----------|---------|-------------|
| Dossier racine GED | `D:\` | Racine où seront créés les sous-dossiers |
| Dossier quarantaine | `D:\_A_Classer` | Documents non identifiés |
| Chemin Tesseract | `C:\Program Files\Tesseract-OCR\tesseract.exe` | Moteur OCR |
| Dossier sortie scanner | `C:\Users\...\Scan` | Où votre scanner dépose les fichiers |

## Structure des dossiers générée

```
D:\
├── Automobile\
│   ├── Citroen\
│   └── ControleTechnique\
├── Banque\
│   └── CreditAgricole\
├── Energie\
│   └── EDF\
├── Assurance\
│   └── MAIF\
└── _A_Classer\     ← Quarantaine
```

## Structure du projet

```
GED-Manager/
├── main.py
├── requirements.txt
├── ged_config.ini          ← Créé automatiquement
├── ged_manager.db          ← Base de données SQLite
├── ged_manager.log         ← Historique des actions
├── gui/
│   └── main_window.py      ← Interface graphique PySimpleGUI
├── core/
│   ├── ocr_engine.py       ← Extraction texte (Tesseract + PyMuPDF)
│   ├── classifier.py       ← Détection origine/société
│   ├── file_manager.py     ← Déplacement/copie/quarantaine
│   └── mapper.py           ← Système d'apprentissage
├── scanner/
│   └── scanner_interface.py ← CANON DR-C125 (WIA + surveillance dossier)
├── db/
│   └── database.py         ← SQLite (mappages, historique, quarantaine)
└── utils/
    ├── config.py            ← Configuration (ged_config.ini)
    └── logger.py            ← Logs horodatés
```

## Flux d'utilisation

```
Lancer main.py
      ↓
[Fichier existant] ou [Scanner]
      ↓
OCR + analyse du texte
      ↓
Mots-clés détectés → mappage connu ?
  ├── OUI → Proposition automatique → Confirmer → Classé
  └── NON → Saisie manuelle du dossier → [Mémoriser ?] → Classé
      ↓ (si incertain)
    Quarantaine → traitement ultérieur
```
