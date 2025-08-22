# Fichier : utils/date_utils.py
# Version finale corrigée avec validation de date stricte et gestion d'erreur.

from datetime import datetime, timedelta, date
import sqlite3
import logging
from utils.config_loader import CONFIG

# --- Gestion optionnelle de la bibliothèque holidays ---
try:
    import holidays
    HOLIDAYS_AVAILABLE = True
except ImportError:
    HOLIDAYS_AVAILABLE = False
    logging.warning("Bibliothèque 'holidays' non trouvée. Seuls les jours fériés personnalisés seront chargés.")

# --- Fonctions de formatage (ajustées pour la nouvelle validation) ---

def format_date_for_display(date_str_sql):
    """Convertit une date du format SQL (YYYY-MM-DD) en format affichable (DD/MM/YYYY)."""
    if not date_str_sql:
        return ""
    try:
        if hasattr(date_str_sql, 'strftime'):
            return date_str_sql.strftime("%d/%m/%Y")
        validated_date = validate_date(date_str_sql)
        return validated_date.strftime("%d/%m/%Y") if validated_date else str(date_str_sql)
    except (ValueError, TypeError, AttributeError):
        return str(date_str_sql)

def format_date_for_display_short(date_obj):
    """Convertit un objet date en format affichable court (JJ/MM/AA)."""
    if not date_obj:
        return ""
    try:
        if hasattr(date_obj, 'strftime'):
            return date_obj.strftime("%d/%m/%y")
        validated_date = validate_date(str(date_obj))
        return validated_date.strftime("%d/%m/%y") if validated_date else str(date_obj)
    except (ValueError, TypeError, AttributeError):
        return str(date_obj)

# --- Fonction de validation (corrigée) ---

def validate_date(date_str):
    """
    Valide et convertit une chaîne de caractères en objet datetime de manière stricte.
    Accepte les formats JJ/MM/AAAA, JJ-MM-AAAA, et AAAA-MM-JJ.
    Retourne None si le format est invalide.
    """
    if not date_str:
        return None
    
    if isinstance(date_str, datetime):
        return date_str
    if isinstance(date_str, date):
        return datetime.combine(date_str, datetime.min.time())
    if not isinstance(date_str, str):
        return None

    # Nettoie la chaîne pour ne garder que la partie date
    date_part = str(date_str).strip().split(" ")[0]
    accepted_formats = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"]
    
    for fmt in accepted_formats:
        try:
            return datetime.strptime(date_part, fmt)
        except (ValueError, TypeError):
            continue
            
    return None

# --- Fonctions de calcul (ajustées pour la nouvelle validation) ---

def get_holidays_set_for_period(db_manager, start_year, end_year):
    """Charge les jours fériés (officiels et personnalisés) pour une période donnée."""
    country_code = CONFIG['conges']['holidays_country']
    all_h = {}
    
    for year in range(start_year, end_year + 2):
        # Charge les jours fériés officiels si la bibliothèque est disponible
        if HOLIDAYS_AVAILABLE:
            try:
                all_h.update(holidays.country_holidays(country_code, years=year))
            except Exception as e:
                logging.error(f"Erreur lors de la récupération des jours fériés officiels pour {year}: {e}")

        # Charge les jours fériés personnalisés depuis la base de données
        try:
            if db_manager and db_manager.conn:
                db_h = db_manager.get_holidays_for_year(str(year))
                for date_str, name, type in db_h:
                    validated_date = validate_date(date_str)
                    if validated_date:
                        all_h[validated_date.date()] = name
        except sqlite3.Error as e:
            logging.error(f"Erreur lors du chargement des jours fériés personnalisés pour {year}: {e}")
            
    return set(all_h.keys())

def jours_ouvres(date_debut, date_fin, holidays_set):
    """Calcule le nombre de jours ouvrés entre deux dates, en excluant les jours fériés."""
    if not date_debut or not date_fin or date_fin < date_debut:
        return 0
    jours = 0
    current_day = date_debut.date() if isinstance(date_debut, datetime) else date_debut
    end_day = date_fin.date() if isinstance(date_fin, datetime) else date_fin
    while current_day <= end_day:
        if current_day.weekday() < 5 and current_day not in holidays_set:
            jours += 1
        current_day += timedelta(days=1)
    return jours

def calculate_reprise_date(end_date, holidays_set):
    """Calcule la date de reprise de service."""
    if not end_date:
        return None
    reprise_date = end_date.date() if isinstance(end_date, datetime) else end_date
    reprise_date += timedelta(days=1)
    while reprise_date.weekday() >= 5 or reprise_date in holidays_set: 
        reprise_date += timedelta(days=1)
    return reprise_date