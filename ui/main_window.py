# Fichier : ui/main_window.py
# Ce fichier utilise les nouvelles fonctions de formatage de date sans nécessiter de modification.

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from collections import defaultdict
import logging
import os
import sqlite3
import threading
from datetime import datetime, date
import subprocess
import sys

from core.conges.manager import CongeManager
from core.constants import SoldeStatus
from ui.forms.agent_form import AgentForm
from ui.forms.conge_form import CongeForm
from ui.widgets.secondary_windows import AdminWindow, JustificatifsWindow
from utils.file_utils import export_agents_to_excel, export_all_conges_to_excel, import_agents_from_excel, generate_decision_from_template
from utils.date_utils import format_date_for_display, format_date_for_display_short, calculate_reprise_date
from utils.config_loader import CONFIG


def treeview_sort_column(tv, col, reverse):
    """Fonction utilitaire pour trier une colonne de Treeview."""
    items_list = [(tv.set(k, col), k) for k in tv.get_children('')]
    
    numeric_cols = ['Solde Total', 'Jours', 'PPR']
    if 'Solde ' in col:
        numeric_cols.append(col)
        
    try:
        if col in numeric_cols:
            items_list.sort(key=lambda t: float(str(t[0]).replace('j', '').replace(',', '.').strip()), reverse=reverse)
        else:
            items_list.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
    except (ValueError, IndexError):
        items_list.sort(key=lambda t: str(t[0]), reverse=reverse)
        
    for index, (val, k) in enumerate(items_list):
        tv.move(k, '', index)
        
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))


class MainWindow(tk.Tk):
    def __init__(self, manager: CongeManager, base_dir: str):
        super().__init__()
        self.manager = manager
        self.base_dir = base_dir
        self.title(f"{CONFIG['app']['title']} - v{CONFIG['app']['version']}")
        self.minsize(1400, 700)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self.annee_exercice = self.manager.get_annee_exercice()
        
        self.current_page = 1
        self.items_per_page = 50
        self.total_pages = 1
        
        self.restart_on_close = False

        self.create_widgets()
        self.refresh_all()

    def on_close(self):
        if messagebox.askokcancel("Quitter", "Voulez-vous vraiment quitter ?"):
            self.destroy()

    def trigger_restart(self):
        """Active le drapeau de redémarrage et ferme la fenêtre."""
        self.restart_on_close = True
        self.destroy()

    def set_status(self, message):
        self.status_var.set(message)
        self.update_idletasks()

    def create_widgets(self):
        """Crée et organise tous les widgets de l'interface principale."""
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("Treeview", rowheight=25, font=('Helvetica', 10))
        style.configure("Treeview.Heading", font=('Helvetica', 10, 'bold'), relief="raised")
        style.configure("TLabel", font=('Helvetica', 11))
        style.configure("TButton", font=('Helvetica', 10))
        style.configure("TLabelframe.Label", font=('Helvetica', 12, 'bold'))

        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_pane = ttk.Frame(main_pane, padding=5)
        main_pane.add(left_pane, weight=3)

        agents_frame = ttk.LabelFrame(left_pane, text="Agents")
        agents_frame.pack(fill=tk.BOTH, expand=True)

        search_frame = ttk.Frame(agents_frame)
        search_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(search_frame, text="Rechercher:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.search_agents())
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(fill=tk.X, expand=True, side=tk.LEFT)
        
        an_n, an_n1, an_n2 = self.annee_exercice, self.annee_exercice - 1, self.annee_exercice - 2
        self.cols_agents = ["ID", "Nom", "Prénom", "PPR", "Grade", f"Solde {an_n2}", f"Solde {an_n1}", f"Solde {an_n}", "Solde Total"]
        self.list_agents = ttk.Treeview(agents_frame, columns=self.cols_agents, show="headings", selectmode="browse")
        
        for col in self.cols_agents:
            self.list_agents.heading(col, text=col, command=lambda c=col: treeview_sort_column(self.list_agents, c, False))

        self.list_agents.column("ID", width=0, stretch=False)
        self.list_agents.column("Nom", width=120)
        self.list_agents.column("Prénom", width=120)
        self.list_agents.column("PPR", width=80, anchor="center")
        self.list_agents.column("Grade", width=100)
        self.list_agents.column(f"Solde {an_n2}", width=80, anchor="center")
        self.list_agents.column(f"Solde {an_n1}", width=80, anchor="center")
        self.list_agents.column(f"Solde {an_n}", width=80, anchor="center")
        self.list_agents.column("Solde Total", width=90, anchor="center")
        
        self.list_agents.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.list_agents.bind("<<TreeviewSelect>>", self.on_agent_select)
        self.list_agents.bind("<Double-1>", lambda e: self.modify_selected_agent())

        pagination_frame = ttk.Frame(agents_frame)
        pagination_frame.pack(fill=tk.X, padx=5, pady=5)
        self.prev_button = ttk.Button(pagination_frame, text="<< Précédent", command=self.prev_page)
        self.prev_button.pack(side=tk.LEFT)
        self.page_label = ttk.Label(pagination_frame, text="Page 1 / 1")
        self.page_label.pack(side=tk.LEFT, expand=True)
        self.next_button = ttk.Button(pagination_frame, text="Suivant >>", command=self.next_page)
        self.next_button.pack(side=tk.RIGHT)
        
        self.btn_frame_agents = ttk.Frame(agents_frame)
        self.btn_frame_agents.pack(fill=tk.X, padx=5, pady=(0, 5))
        ttk.Button(self.btn_frame_agents, text="Ajouter", command=self.add_agent_ui).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(self.btn_frame_agents, text="Modifier", command=self.modify_selected_agent).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(self.btn_frame_agents, text="Supprimer", command=self.delete_selected_agent).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        
        self.io_frame_agents = ttk.Frame(agents_frame)
        self.io_frame_agents.pack(fill=tk.X, padx=5, pady=(5, 5))
        ttk.Button(self.io_frame_agents, text="Importer Agents (Excel)", command=self.import_agents).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(self.io_frame_agents, text="Exporter Agents (Excel)", command=self.export_agents).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        right_pane = ttk.PanedWindow(main_pane, orient=tk.VERTICAL)
        main_pane.add(right_pane, weight=2)
        
        conges_frame = ttk.LabelFrame(right_pane, text="Congés de l'agent sélectionné")
        right_pane.add(conges_frame, weight=3)
        
        filter_frame = ttk.Frame(conges_frame)
        filter_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(filter_frame, text="Filtrer par type:").pack(side=tk.LEFT, padx=(0, 5))
        self.conge_filter_var = tk.StringVar(value="Tous")
        conge_filter_combo = ttk.Combobox(filter_frame, textvariable=self.conge_filter_var, values=["Tous"] + CONFIG['ui']['types_conge'], state="readonly")
        conge_filter_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        conge_filter_combo.bind("<<ComboboxSelected>>", self.on_agent_select)
        
        cols_conges = ("CongeID", "Certificat", "Type", "Début", "Fin", "Date Reprise", "Jours", "Justification", "Intérimaire")
        self.list_conges = ttk.Treeview(conges_frame, columns=cols_conges, show="headings", selectmode="browse")
        
        for col in cols_conges:
            self.list_conges.heading(col, text=col, command=lambda c=col: treeview_sort_column(self.list_conges, c, False))
        
        self.list_conges.column("CongeID", width=0, stretch=False)
        self.list_conges.column("Certificat", width=80, anchor="center")
        self.list_conges.column("Type", width=120)
        self.list_conges.column("Début", width=90, anchor="center")
        self.list_conges.column("Fin", width=90, anchor="center")
        self.list_conges.column("Date Reprise", width=90, anchor="center")
        self.list_conges.column("Jours", width=50, anchor="center")
        self.list_conges.column("Intérimaire", width=150)
        self.list_conges.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.list_conges.tag_configure("summary", background="#e6f2ff", font=("Helvetica", 10, "bold"))
        self.list_conges.tag_configure("annule", foreground="grey", font=('Helvetica', 10, 'overstrike'))
        self.list_conges.bind("<Double-1>", lambda e: self.on_conge_double_click())
        self.list_conges.bind("<<TreeviewSelect>>", self._update_conge_action_buttons_state)
        
        self.btn_frame_conges = ttk.Frame(conges_frame)
        self.btn_frame_conges.pack(fill=tk.X, padx=5, pady=(0, 5))
        ttk.Button(self.btn_frame_conges, text="Ajouter", command=self.add_conge_ui).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(self.btn_frame_conges, text="Modifier", command=self.modify_selected_conge).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(self.btn_frame_conges, text="Supprimer", command=self.delete_selected_conge).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.btn_generate_decision = ttk.Button(self.btn_frame_conges, text="Générer Décision", command=self.on_generate_decision_click, state="disabled")
        self.btn_generate_decision.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        
        stats_frame = ttk.LabelFrame(right_pane, text="Tableau de Bord")
        right_pane.add(stats_frame, weight=1)
        
        on_leave_frame = ttk.LabelFrame(stats_frame, text="Agents Actuellement en Congé")
        on_leave_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        cols_on_leave = ("Agent", "PPR", "Type Congé", "Date de Reprise")
        self.list_on_leave = ttk.Treeview(on_leave_frame, columns=cols_on_leave, show="headings", height=8)
        for col in cols_on_leave:
            self.list_on_leave.heading(col, text=col)
        self.list_on_leave.column("Agent", width=200)
        self.list_on_leave.column("PPR", width=100, anchor="center")
        self.list_on_leave.column("Type Congé", width=150)
        self.list_on_leave.column("Date de Reprise", width=120, anchor="center")
        self.list_on_leave.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.global_actions_frame = ttk.Frame(stats_frame)
        self.global_actions_frame.pack(fill=tk.X, padx=5, pady=(5, 5))
        ttk.Button(self.global_actions_frame, text="Actualiser", command=self.refresh_stats).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(self.global_actions_frame, text="Suivi Justificatifs", command=self.open_justificatifs_suivi).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(self.global_actions_frame, text="Administration", command=self.open_admin_window).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(self.global_actions_frame, text="Exporter Tous les Congés", command=self.export_conges).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        
        self.status_var = tk.StringVar(value="Prêt.")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def open_admin_window(self):
        AdminWindow(self, self.manager)
        
    def get_selected_agent_id(self):
        selection = self.list_agents.selection()
        return int(self.list_agents.item(selection[0])["values"][0]) if selection else None
        
    def get_selected_conge_id(self):
        selection = self.list_conges.selection()
        if not selection:
            return None
        item = self.list_conges.item(selection[0])
        if "summary" in item["tags"]:
            return None
        return int(item["values"][0]) if item["values"] else None
        
    def add_agent_ui(self):
        AgentForm(self, self.manager)
        
    def modify_selected_agent(self):
        agent_id = self.get_selected_agent_id()
        if agent_id:
            AgentForm(self, self.manager, agent_id_to_modify=agent_id)
        else:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un agent à modifier.")
            
    def delete_selected_agent(self):
        agent_id = self.get_selected_agent_id()
        if not agent_id:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un agent à supprimer.")
            return
        agent = self.manager.get_agent_by_id(agent_id)
        if not agent:
            messagebox.showerror("Erreur", "Agent introuvable.")
            return
        agent_nom = f"{agent.nom} {agent.prenom}"
        if messagebox.askyesno("Confirmation", f"Supprimer l'agent '{agent_nom}' et tous ses congés ?\nCette action est irréversible."):
            try:
                if self.manager.delete_agent(agent.id):
                    self.set_status(f"Agent '{agent_nom}' supprimé.")
                    self.refresh_all()
            except Exception as e:
                messagebox.showerror("Erreur de suppression", f"Une erreur est survenue : {e}")
                
    def add_conge_ui(self):
        agent_id = self.get_selected_agent_id()
        if agent_id:
            CongeForm(self, self.manager, agent_id)
        else:
            messagebox.showwarning("Aucun agent", "Veuillez sélectionner un agent.")
            
    def modify_selected_conge(self):
        agent_id = self.get_selected_agent_id()
        conge_id = self.get_selected_conge_id()
        if agent_id and conge_id:
            CongeForm(self, self.manager, agent_id, conge_id=conge_id)
        else:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un congé à modifier.")
    
    def delete_selected_conge(self):
        conge_id = self.get_selected_conge_id()
        agent_id = self.get_selected_agent_id()
        if not conge_id:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un congé à supprimer.")
            return
        try:
            if messagebox.askyesno("Confirmation", "Êtes-vous sûr de vouloir supprimer ce congé ?"):
                if self.manager.delete_conge(conge_id):
                    self.set_status("Congé supprimé.")
                    self.refresh_all(agent_id)
        except (ValueError, sqlite3.Error) as e:
            messagebox.showerror("Erreur de suppression", str(e))
        except Exception as e:
            logging.error(f"Erreur inattendue suppression congé: {e}", exc_info=True)
            messagebox.showerror("Erreur Inattendue", f"Une erreur est survenue: {e}")

    def refresh_all(self, agent_to_select_id=None):
        self.annee_exercice = self.manager.get_annee_exercice()
        current_selection = agent_to_select_id or self.get_selected_agent_id()
        self.refresh_agents_list(current_selection)
        self.refresh_stats()
        
    def refresh_agents_list(self, agent_to_select_id=None):
        for row in self.list_agents.get_children():
            self.list_agents.delete(row)
        term = self.search_var.get().strip().lower() or None
        total_items = self.manager.get_agents_count(term)
        self.total_pages = max(1, (total_items + self.items_per_page - 1) // self.items_per_page)
        self.current_page = min(self.current_page, self.total_pages)
        offset = (self.current_page - 1) * self.items_per_page
        agents = self.manager.get_all_agents(term=term, limit=self.items_per_page, offset=offset)
        selected_item_id = None
        an_n, an_n1, an_n2 = self.annee_exercice, self.annee_exercice - 1, self.annee_exercice - 2
        for agent in agents:
            soldes_par_annee = {s.annee: s.solde for s in agent.soldes_annuels if s.statut == SoldeStatus.ACTIF}
            solde_n2 = soldes_par_annee.get(an_n2, 0.0)
            solde_n1 = soldes_par_annee.get(an_n1, 0.0)
            solde_n = soldes_par_annee.get(an_n, 0.0)
            solde_total = agent.get_solde_total_actif()
            agent_values = (agent.id, agent.nom, agent.prenom, agent.ppr, agent.grade, f"{solde_n2:.1f} j", f"{solde_n1:.1f} j", f"{solde_n:.1f} j", f"{solde_total:.1f} j")
            item_id = self.list_agents.insert("", "end", values=agent_values)
            if agent.id == agent_to_select_id:
                selected_item_id = item_id
        if selected_item_id:
            self.list_agents.selection_set(selected_item_id)
            self.list_agents.focus(selected_item_id)
        self.on_agent_select()
        self.page_label.config(text=f"Page {self.current_page} / {self.total_pages}")
        self.prev_button.config(state="normal" if self.current_page > 1 else "disabled")
        self.next_button.config(state="normal" if self.current_page < self.total_pages else "disabled")
        self.set_status(f"{len(agents)} agents affichés sur {total_items} au total.")
        
    def refresh_conges_list(self, agent_id):
        self.list_conges.delete(*self.list_conges.get_children())
        filtre = self.conge_filter_var.get()
        conges_data = self.manager.get_conges_for_agent(agent_id)
        conges_par_annee = defaultdict(list)
        for c in conges_data:
            if filtre != "Tous" and c.type_conge != filtre:
                continue
            try:
                conges_par_annee[c.date_debut.year].append(c)
            except AttributeError:
                logging.warning(f"Date invalide ou nulle pour congé ID {c.id}")
        for annee in sorted(conges_par_annee.keys(), reverse=True):
            total_jours = sum(c.jours_pris for c in conges_par_annee[annee] if c.type_conge == 'Congé annuel' and c.statut == 'Actif')
            summary_id = self.list_conges.insert("", "end", values=("", "", f"📅 ANNÉE {annee}", "", "", "", total_jours, f"{total_jours} jours pris", ""), tags=("summary",), open=True)
            holidays_set = self.manager.get_holidays_set_for_period(annee, annee + 1)
            for conge in sorted(conges_par_annee[annee], key=lambda c: c.date_debut):
                cert_status = "✅ Fourni" if self.manager.get_certificat_for_conge(conge.id) else "❌ Manquant" if conge.type_conge == 'Congé de maladie' else ""
                interim_info = ""
                if conge.interim_id:
                    interim = self.manager.get_agent_by_id(conge.interim_id)
                    interim_info = f"{interim.nom} {interim.prenom}" if interim else "Agent Supprimé"
                tags = ('annule',) if conge.statut == 'Annulé' else ()
                reprise_date = calculate_reprise_date(conge.date_fin, holidays_set)
                reprise_date_str = format_date_for_display_short(reprise_date) if reprise_date else ""
                self.list_conges.insert(summary_id, "end", values=(conge.id, cert_status, conge.type_conge, format_date_for_display_short(conge.date_debut), format_date_for_display_short(conge.date_fin), reprise_date_str, conge.jours_pris, conge.justif or "", interim_info), tags=tags)

    def refresh_stats(self):
        for row in self.list_on_leave.get_children():
            self.list_on_leave.delete(row)
        try:
            holidays_set = self.manager.get_holidays_set_for_period(self.annee_exercice, self.annee_exercice + 1)
            agents_on_leave_data = self.manager.get_agents_on_leave_today()
            for nom, prenom, ppr, type_conge, date_fin_str in agents_on_leave_data:
                # La date de la DB est déjà un objet date/datetime grâce à `detect_types`
                reprise_date = calculate_reprise_date(date_fin_str, holidays_set)
                reprise_date_display = format_date_for_display(reprise_date)
                self.list_on_leave.insert("", "end", values=(f"{nom} {prenom}", ppr, type_conge, reprise_date_display))
        except (sqlite3.Error, AttributeError) as e:
            self.list_on_leave.insert("", "end", values=(f"Erreur: {e}", "", "", ""))

    def search_agents(self):
        self.current_page = 1
        self.refresh_agents_list()
    
    def on_agent_select(self, event=None):
        if self.get_selected_agent_id():
            self.refresh_conges_list(self.get_selected_agent_id())
        else:
            self.list_conges.delete(*self.list_conges.get_children())
        self._update_conge_action_buttons_state()

    def _update_conge_action_buttons_state(self, event=None):
        conge_id = self.get_selected_conge_id()
        state = "normal" if conge_id else "disabled"
        self.btn_generate_decision.config(state=state)

    def on_generate_decision_click(self):
        conge_id = self.get_selected_conge_id()
        agent_id = self.get_selected_agent_id()
        if not conge_id or not agent_id:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un congé.")
            return

        conge = self.manager.get_conge_by_id(conge_id)
        agent = self.manager.get_agent_by_id(agent_id)

        if not conge or not agent:
            messagebox.showerror("Erreur", "Impossible de récupérer les informations du congé ou de l'agent.")
            return

        templates_dir_name = CONFIG.get('paths', {}).get('templates_dir', 'templates')
        grade_str = agent.grade.lower().replace(" ", "_")
        template_name = f"{grade_str}.docx"
        template_path = os.path.join(self.base_dir, templates_dir_name, template_name)

        if not os.path.exists(template_path):
            messagebox.showerror("Modèle manquant", f"Le modèle pour le grade '{agent.grade}' est introuvable.\nIl devrait être ici : {template_path}")
            return
            
        details_solde_str = ""
        if conge.type_conge == "Congé annuel":
            details = self.manager.get_deduction_details(agent.id, conge.jours_pris)
            parts = []
            for year, days in sorted(details.items()):
                days_int = int(round(days))
                jour_text = "jour" if days_int == 1 else "jours"
                parts.append(f"{days_int} {jour_text} au titre de l'année {year}")
            details_solde_str = " et ".join(parts)

        holidays_set = self.manager.get_holidays_set_for_period(conge.date_fin.year, conge.date_fin.year + 1)
        date_reprise = calculate_reprise_date(conge.date_fin, holidays_set)

        context = {
            "{{nom_complet}}": f"{agent.nom} {agent.prenom}", "{{grade}}": agent.grade, "{{ppr}}": agent.ppr,
            "{{date_debut}}": format_date_for_display(conge.date_debut), "{{date_fin}}": format_date_for_display(conge.date_fin),
            "{{date_reprise}}": format_date_for_display(date_reprise) if date_reprise else "N/A",
            "{{jours_pris}}": str(conge.jours_pris), "{{details_solde}}": details_solde_str,
            "{{date_aujourdhui}}": date.today().strftime("%d/%m/%Y")
        }

        initial_filename = f"Decision_Conge_{agent.nom}_{conge.date_debut.strftime('%Y-%m-%d')}.docx"
        save_path = filedialog.asksaveasfilename(
            title="Enregistrer la décision", initialfile=initial_filename, defaultextension=".docx",
            filetypes=[("Documents Word", "*.docx"), ("Tous les fichiers", "*.*")]
        )

        if not save_path:
            return

        try:
            generate_decision_from_template(template_path, save_path, context)
            if messagebox.askyesno("Succès", "La décision a été générée.\nVoulez-vous ouvrir le fichier ?", parent=self):
                self._open_file(save_path)
        except Exception as e:
            messagebox.showerror("Erreur de Génération", f"Une erreur est survenue:\n{e}", parent=self)

    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.refresh_agents_list(self.get_selected_agent_id())
            
    def next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.refresh_agents_list(self.get_selected_agent_id())
            
    def on_conge_double_click(self):
        conge_id = self.get_selected_conge_id()
        if not conge_id:
            return
        
        cert = self.manager.get_certificat_for_conge(conge_id)
        if cert and cert[2] and os.path.exists(cert[2]):
            try:
                self._open_file(cert[2])
            except Exception as e:
                messagebox.showerror("Erreur d'ouverture", f"Impossible d'ouvrir le fichier:\n{e}", parent=self)
        else:
            self.modify_selected_conge()

    def open_justificatifs_suivi(self):
        JustificatifsWindow(self, self.manager)
    
    def export_agents(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Fichiers Excel", "*.xlsx")], title="Exporter la liste des agents", initialfile=f"Export_Agents_{datetime.now().strftime('%Y-%m-%d')}.xlsx")
        if not save_path:
            return
        db_path = self.manager.db.db_file
        cert_path = self.manager.certificats_dir
        self._run_long_task(lambda: export_agents_to_excel(db_path, cert_path, save_path), self._on_task_complete, "Exportation des agents en cours...")

    def export_conges(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Fichiers Excel", "*.xlsx")], title="Exporter tous les congés", initialfile=f"Export_Conges_Total_{datetime.now().strftime('%Y-%m-%d')}.xlsx")
        if not save_path:
            return
        db_path = self.manager.db.db_file
        cert_path = self.manager.certificats_dir
        self._run_long_task(lambda: export_all_conges_to_excel(db_path, cert_path, save_path), self._on_task_complete, "Exportation de tous les congés en cours...")

    def import_agents(self):
        source_path = filedialog.askopenfilename(title="Sélectionner un fichier Excel à importer", filetypes=[("Fichiers Excel", "*.xlsx")])
        if not source_path:
            return
        db_path = self.manager.db.db_file
        cert_path = self.manager.certificats_dir
        self._run_long_task(lambda: import_agents_from_excel(db_path, cert_path, source_path), self._on_import_complete, "Importation des agents depuis Excel en cours...")

    def _open_file(self, filepath):
        filepath = os.path.realpath(filepath)
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin":
                subprocess.run(["open", filepath], check=True)
            else:
                subprocess.run(["xdg-open", filepath], check=True)
        except Exception as e:
            messagebox.showerror("Erreur d'Ouverture", f"Impossible d'ouvrir le fichier:\n{e}", parent=self)
            
    def _run_long_task(self, task_lambda, on_complete, status_message):
        self.set_status(status_message)
        self.config(cursor="watch")
        self._toggle_buttons_state("disabled")
        result_container = []
        
        def task_wrapper():
            try:
                result_container.append(task_lambda())
            except Exception as e:
                result_container.append(e)
                
        worker_thread = threading.Thread(target=task_wrapper)
        worker_thread.start()
        self._check_thread_completion(worker_thread, result_container, on_complete)
    
    def _check_thread_completion(self, thread, result_container, on_complete):
        if thread.is_alive():
            self.after(100, lambda: self._check_thread_completion(thread, result_container, on_complete))
        else:
            result = result_container[0] if result_container else None
            on_complete(result)
            self.config(cursor="")
            self._toggle_buttons_state("normal")
            self.set_status("Prêt.")
    
    def _on_task_complete(self, result):
        if isinstance(result, Exception):
            messagebox.showerror("Erreur", f"L'opération a échoué:\n{result}")
        elif result:
            messagebox.showinfo("Succès", result)
    
    def _on_import_complete(self, result):
        self._on_task_complete(result)
        if not isinstance(result, Exception):
            self.refresh_all()

    def _toggle_buttons_state(self, state):
        for frame in [self.btn_frame_agents, self.io_frame_agents, self.btn_frame_conges, self.global_actions_frame]:
            for child in frame.winfo_children():
                if isinstance(child, (ttk.Button, ttk.Combobox)):
                    child.config(state=state)