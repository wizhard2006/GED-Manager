"""
GED-Manager — Point d'entrée principal
Gestion Électronique de Documents avec classification automatique
"""

import sys
import os

# Assurer que les modules locaux sont trouvables
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.main_window import GEDManagerApp


def main():
    app = GEDManagerApp()
    app.run()


if __name__ == "__main__":
    main()
