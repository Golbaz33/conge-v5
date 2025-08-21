# Fichier : main.py
# Point d'entrée principal de l'application. Il initialise la configuration,
# la base de données, la logique métier, puis lance l'interface graphique.

import tkinter as tk
from tkinter import messagebox
import sys
import os
import logging

# Imports des modules de l'application, centralisés en haut du fichier.
from utils.config_loader import load_config, CONFIG
from db.database import DatabaseManager
from core.conges.manager import CongeManager
from ui.main_window import MainWindow


# --- SECTION 1 : Configuration des chemins d'accès ---
# Détermine le répertoire de base de l'application pour un accès fiable aux fichiers.
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()

CONFIG_PATH = os.path.join(BASE_DIR, "config.yaml")


# --- SECTION 2 : Chargement de la configuration ---
# Charge les paramètres depuis config.yaml. C'est une étape critique,
# l'application s'arrête si elle échoue.
try:
    load_config(CONFIG_PATH)
except Exception as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Erreur Critique de Configuration", f"Impossible de charger la configuration:\n{e}")
    sys.exit(1)


# --- SECTION 3 : Démarrage de l'application ---
if __name__ == "__main__":

    # Initialisation de l'environnement (chemins, logs).
    CERTIFICATS_DIR_ABS = os.path.join(BASE_DIR, CONFIG['db']['certificates_dir'])
    DB_PATH_ABS = os.path.join(BASE_DIR, CONFIG['db']['filename'])
    LOG_FILE_PATH = os.path.join(BASE_DIR, "conges.log")
    logging.basicConfig(filename=LOG_FILE_PATH, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')

    # Boucle permettant un redémarrage propre de l'application.
    restart_app = True
    while restart_app:
        restart_app = False

        # Connexion et mise à jour de la base de données.
        db_manager = DatabaseManager(DB_PATH_ABS)
        if not db_manager.connect():
            sys.exit(1)

        try:
            db_manager.run_migrations()
        except Exception as e:
            logging.critical(f"Échec de la migration de la BDD. Arrêt. Erreur : {e}")
            messagebox.showerror("Erreur de Migration DB", f"La mise à jour de la base de données a échoué.\nErreur: {e}")
            db_manager.close()
            sys.exit(1)

        # Initialisation du gestionnaire métier et lancement de l'interface.
        conge_manager = CongeManager(db_manager, CERTIFICATS_DIR_ABS)
        
        print(f"--- Lancement de {CONFIG['app']['title']} v{CONFIG['app']['version']} ---")
        app = MainWindow(conge_manager, BASE_DIR)
        app.mainloop()

        # Nettoyage à la fermeture.
        if hasattr(app, 'restart_on_close') and app.restart_on_close:
            restart_app = True
        
        db_manager.close()
    
    print("--- Application fermée, connexion à la base de données terminée. ---")