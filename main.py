"""
GED-Manager MVP - Point d'entrée
Gestion Électronique de Documents avec classification automatique
"""

import sys
import os
from gui.main_window import GEDManagerApp

def main():
    """Lancer l'application principale"""
    app = GEDManagerApp()
    app.run()

if __name__ == "__main__":
    main()