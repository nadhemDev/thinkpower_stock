# -*- coding: utf-8 -*-
"""Stock Think Power - Interface Graphique Moderne & Intelligente (CustomTkinter).
Gestionnaire de stock universel pour tout type d'activité commerciale, industrielle ou logistique.
Base de données autonome Excel (Stock_Think_Power_Base.xlsx) située à côté de l'application.
"""
import os
import sys
import traceback
from datetime import datetime
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog, Menu

from stock_think_power import (
    DB, APP_NAME, EXPORT_DIR, DB_PATH, BASE_DIR, calc, etat, fmt_date, user_date,
    to_float, now_iso, export_pdf, export_excel, export_bon_pdf, export_reorder_pdf,
    get_resource_path, get_base_dir
)

# ----------------------------------------------------------------- Palette UI
BG = "#0b0f19"          # Fond principal ultra sombre moderne
SIDE = "#111827"        # Barre latérale sombre élégante
PANEL = "#1e293b"       # Panneaux & cartes de contenu
CARD = "#172033"        # Fond des cartes KPI
FIELD = "#0f172a"       # Champs de saisie
BORDER = "#334155"      # Bordures discrètes
CONSOLE = "#090d16"     # Fond de console / aperçu

BLUE = "#3b82f6"        # Bleu primaire
BLUE_HOVER = "#2563eb"
CYAN = "#06b6d4"        # Cyan accent
CYAN_HOVER = "#0891b2"
GREEN = "#10b981"       # Vert succès / en stock
GREEN_HOVER = "#059669"
RED = "#ef4444"         # Rouge alerte / rupture / suppression
RED_HOVER = "#dc2626"
ORANGE = "#f59e0b"      # Orange stock faible / avertissement
ORANGE_HOVER = "#d97706"
PURPLE = "#8b5cf6"      # Violet inventaire / IA

TEXT_WHITE = "#f8fafc"
TEXT_MUTED = "#94a3b8"

HOVER_MAP = {
    BLUE: BLUE_HOVER,
    CYAN: CYAN_HOVER,
    GREEN: GREEN_HOVER,
    RED: RED_HOVER,
    ORANGE: ORANGE_HOVER,
    PURPLE: "#7c3aed",
    PANEL: "#27354a",
    SIDE: "#161f33"
}

ETAT_LABEL = {"ok": "● En Stock", "faible": "▲ Faible", "rupture": "✖ Rupture"}
ETAT_TAG = {"ok": "ok", "faible": "critique", "rupture": "rupture"}
FILTERS = {"Tous": "all", "En Stock": "ok", "Stock Faible": "alert", "En Rupture": "rupture"}


# ----------------------------------------------------------------- Widgets Helper
def panel(parent, **kw):
    return ctk.CTkFrame(parent, fg_color=PANEL, corner_radius=12, border_width=1, border_color=BORDER, **kw)


def label(parent, text, color=TEXT_WHITE, size=13, bold=False, **kw):
    return ctk.CTkLabel(
        parent,
        text=text,
        text_color=color,
        font=("Segoe UI", size, "bold" if bold else "normal"),
        **kw
    )


def entry(parent, width=260, **kw):
    return ctk.CTkEntry(
        parent,
        width=width,
        height=38,
        fg_color=FIELD,
        border_color=BORDER,
        corner_radius=8,
        text_color=TEXT_WHITE,
        placeholder_text_color=TEXT_MUTED,
        **kw
    )


def combo(parent, width=260, command=None, values=None):
    return ctk.CTkComboBox(
        parent,
        width=width,
        height=38,
        fg_color=FIELD,
        border_color=BORDER,
        button_color=BLUE,
        button_hover_color=BLUE_HOVER,
        corner_radius=8,
        text_color=TEXT_WHITE,
        dropdown_fg_color=PANEL,
        dropdown_text_color=TEXT_WHITE,
        dropdown_hover_color=BLUE,
        values=values or [""],
        command=command
    )


def button(parent, text, color, command, width=None, height=38, icon=None):
    kw = {"width": width} if width else {}
    return ctk.CTkButton(
        parent,
        text=text,
        height=height,
        fg_color=color,
        hover_color=HOVER_MAP.get(color, color),
        corner_radius=8,
        font=("Segoe UI", 12, "bold"),
        text_color=TEXT_WHITE,
        command=command,
        **kw
    )


def card(parent, title, icon="📊", color=BLUE):
    f = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=12, border_width=1, border_color=BORDER)
    top = ctk.CTkFrame(f, fg_color="transparent")
    top.pack(fill="x", padx=14, pady=(10, 0))
    label(top, icon, color, 14, True).pack(side="left")
    label(top, f"  {title}", TEXT_MUTED, 11, False).pack(side="left")
    val_lbl = label(f, "-", color, 18, True)
    val_lbl.pack(anchor="w", padx=14, pady=(2, 10))
    return f, val_lbl


def make_tree(parent, columns, widths):
    box = ctk.CTkFrame(parent, fg_color=PANEL, corner_radius=10)
    tree = ttk.Treeview(box, columns=columns, show="headings", selectmode="browse")
    left_align = {"Désignation", "Article", "Note", "Client/Fourn.", "Société / Client", "Agent", "Catégorie", "Emplacement", "Remarque"}
    center_align = {"État", "Type", "ID", "Réf", "N° Bon", "Bon", "TVA %", "TVA", "Option 10%", "Unité", "Statut"}

    for c, w in zip(columns, widths):
        tree.heading(c, text=c)
        if c in left_align:
            anchor = "w"
        elif c in center_align:
            anchor = "center"
        else:
            anchor = "e"
        tree.column(c, width=w, anchor=anchor, minwidth=35)

    sb_y = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
    sb_x = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

    tree.grid(row=0, column=0, sticky="nsew")
    sb_y.grid(row=0, column=1, sticky="ns")
    sb_x.grid(row=1, column=0, sticky="ew")
    box.grid_rowconfigure(0, weight=1)
    box.grid_columnconfigure(0, weight=1)

    tree.tag_configure("rupture", foreground="#f87171")
    tree.tag_configure("critique", foreground="#fbbf24")
    tree.tag_configure("ok", foreground="#34d399")
    tree.tag_configure("odd", background="#131c2e")
    tree.tag_configure("even", background="#0f172a")

    return box, tree


def fail(ex, parent=None):
    messagebox.showerror(APP_NAME, str(ex), parent=parent)


# ----------------------------------------------------------- Modales / Dialogues
class ArticleDialog(ctk.CTkToplevel):
    """Ajout ou modification d'une fiche article universelle."""

    def __init__(self, app, article=None):
        super().__init__(app, fg_color=BG)
        self.app, self.article = app, article
        a = article
        is_edit = bool(a)

        self.title("Modifier l'article" if is_edit else "Nouvel article en stock")
        self.geometry("580x680")
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()

        # Centrage
        self.update_idletasks()
        x = app.winfo_x() + (app.winfo_width() - 580) // 2
        y = app.winfo_y() + (app.winfo_height() - 680) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

        title_txt = f"✏️  Modifier l'article #{a['id']}" if is_edit else "➕  Création d'un article en stock"
        label(self, title_txt, BLUE, 18, True).pack(anchor="w", padx=24, pady=(16, 8))

        form = panel(self)
        form.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        form.grid_columnconfigure((0, 1), weight=1)
        self.e = {}

        def field(r, c, key, text, width, val, span=1):
            label(form, text, TEXT_MUTED, 11).grid(row=r * 2, column=c, columnspan=span, sticky="w", padx=14, pady=(6, 2))
            w = entry(form, width)
            w.insert(0, val)
            w.grid(row=r * 2 + 1, column=c, columnspan=span, sticky="ew", padx=14)
            self.e[key] = w

        # Ligne 0 : Réf / Code barre & Agent
        field(0, 0, "ref", "Référence / Code Barre (SKU)", 230, a.get("ref", "") if a else "")
        label(form, "Agent responsable", TEXT_MUTED, 11).grid(row=0, column=1, sticky="w", padx=14, pady=(6, 2))
        self.agent = combo(form, 230)
        self.agent.configure(values=app.db.agents() or ["Admin"])
        self.agent.set(app.db.last_agent or "Admin")
        self.agent.grid(row=1, column=1, sticky="ew", padx=14)

        # Ligne 1 : Désignation
        field(1, 0, "designation", "Matière / Désignation *", 500, a["designation"] if a else "", span=2)

        # Ligne 2 : Catégorie & Unité
        label(form, "Catégorie / Rayon", TEXT_MUTED, 11).grid(row=4, column=0, sticky="w", padx=14, pady=(6, 2))
        self.cat_cb = combo(form, 230)
        self.cat_cb.configure(values=app.db.categories())
        self.cat_cb.set(a.get("categorie", "Général") if a else "Général")
        self.cat_cb.grid(row=5, column=0, sticky="ew", padx=14)

        label(form, "Unité de mesure", TEXT_MUTED, 11).grid(row=4, column=1, sticky="w", padx=14, pady=(6, 2))
        self.unit_cb = combo(form, 230)
        self.unit_cb.configure(values=app.db.units())
        self.unit_cb.set(a.get("unite", "pcs") if a else "pcs")
        self.unit_cb.grid(row=5, column=1, sticky="ew", padx=14)

        # Ligne 3 : Emplacement & Quantité existante
        field(3, 0, "emplacement", "Emplacement magasin (Rayon/Casier)", 230, a.get("emplacement", "") if a else "")
        field(3, 1, "qty", "Quantité existante en stock", 230, f"{a['nombre_existant']:g}" if a else "0")

        # Ligne 4 : Stock min (alerte) & PU HT
        cur = app.db.currency
        field(4, 0, "seuil", "Stock Minimum (seuil alerte)", 230, f"{a['seuil_alerte']:g}" if a else "5")
        field(4, 1, "pu", f"Prix Unitaire HT ({cur})", 230, f"{a['pu_ht']:.3f}" if a else "0.000")

        # Ligne 5 : TVA % & Note
        field(5, 0, "tva", "Taux TVA (%)", 230, f"{a['taux_tva']:g}" if a else "19")
        field(5, 1, "note", "Remarque / Note", 230, a.get("note", "") if a else "")

        # Option 10%
        self.opt10 = ctk.CTkCheckBox(
            form,
            text=f"Option 10% ({app.db.mode10})",
            text_color=TEXT_WHITE,
            fg_color=BLUE,
            hover_color=BLUE_HOVER
        )
        if a and a.get("appliquer_10_pct"):
            self.opt10.select()
        self.opt10.grid(row=12, column=0, columnspan=2, sticky="w", padx=14, pady=(8, 6))

        # Preview dynamique
        self.preview_lbl = label(form, "", CYAN, 11, bold=True)
        self.preview_lbl.grid(row=13, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 6))
        self.update_preview()

        for k in ("qty", "pu", "tva"):
            self.e[k].bind("<KeyRelease>", lambda e: self.update_preview())

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=20, pady=(0, 16))
        button(btns, "Annuler", PANEL, self.destroy, width=120, height=42).pack(side="left")
        button(btns, "💾  Enregistrer dans Excel", BLUE, self.save, height=42).pack(side="right", fill="x", expand=True, padx=(10, 0))

        self.bind("<Return>", lambda e: self.save())
        self.bind("<Escape>", lambda e: self.destroy())

    def update_preview(self):
        try:
            qty = to_float(self.e["qty"].get(), "Qté")
            pu = to_float(self.e["pu"].get(), "PU")
            tva = to_float(self.e["tva"].get(), "TVA")
            c = calc(qty, pu, tva, bool(self.opt10.get()), self.app.db.mode10)
            cur = self.app.db.currency
            self.preview_lbl.configure(text=f"Aperçu : PU TTC = {pu * (1 + tva/100):.3f} {cur}   |   Total TTC = {c['ttc']:.3f} {cur}")
        except Exception:
            self.preview_lbl.configure(text="")

    def save(self):
        try:
            designation = self.e["designation"].get().strip()
            if not designation:
                raise ValueError("La désignation de l'article est obligatoire.")

            data = {
                "ref": self.e["ref"].get().strip(),
                "designation": designation,
                "categorie": self.cat_cb.get().strip(),
                "unite": self.unit_cb.get().strip(),
                "emplacement": self.e["emplacement"].get().strip(),
                "qty": to_float(self.e["qty"].get(), "Nombre existant"),
                "seuil": to_float(self.e["seuil"].get(), "Stock minimum"),
                "pu": to_float(self.e["pu"].get(), "Prix unitaire HT"),
                "tva": to_float(self.e["tva"].get(), "TVA"),
                "opt10": int(self.opt10.get()),
                "note": self.e["note"].get().strip()
            }
            agent = self.agent.get().strip() or "Admin"
            aid = self.article["id"] if self.article else None
            self.app.db.save_article(data, aid, agent)
        except Exception as ex:
            fail(ex, self)
            return

        self.app.refresh_all()
        self.app.msg(f"Article « {data['designation']} » enregistré dans Excel.")
        self.destroy()


class ReceptionDialog(ctk.CTkToplevel):
    """Réception Fournisseur & Entrée de Stock complète."""

    def __init__(self, app, article=None):
        super().__init__(app, fg_color=BG)
        self.app, self.article = app, article
        self.title("Réception Fournisseur / Entrée de Stock")
        self.geometry("520x560")
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()

        self.update_idletasks()
        x = app.winfo_x() + (app.winfo_width() - 520) // 2
        y = app.winfo_y() + (app.winfo_height() - 560) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

        label(self, "📥  Entrée de Stock / Arrivage Fournisseur", GREEN, 17, True).pack(anchor="w", padx=22, pady=(16, 6))

        box = panel(self)
        box.pack(fill="both", expand=True, padx=22, pady=(0, 14))

        # Article sélectionné
        label(box, "Article concerné *", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(10, 2))
        self.art_cb = combo(box, 460)
        art_names = [f"{a.get('ref', '')} - {a['designation']}" for a in app.db.articles()]
        self.art_cb.configure(values=art_names or [""])
        if article:
            self.art_cb.set(f"{article.get('ref', '')} - {article['designation']}")
        elif art_names:
            self.art_cb.set(art_names[0])
        self.art_cb.pack(anchor="w", padx=16)

        # Quantité reçue
        label(box, "Quantité reçue à ajouter (+)", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(8, 2))
        self.qty_e = entry(box, 460)
        self.qty_e.insert(0, "10")
        self.qty_e.pack(anchor="w", padx=16)

        # Fournisseur
        label(box, "Fournisseur / Société expéditrice", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(8, 2))
        self.fourn_e = entry(box, 460, placeholder_text="Ex: Usine Plastique, Fournisseur SARL...")
        self.fourn_e.pack(anchor="w", padx=16)

        # N° Bon de livraison
        label(box, "N° Bon de Livraison (BL) Fournisseur", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(8, 2))
        self.bl_e = entry(box, 460, placeholder_text="Ex: BL-2026-904...")
        self.bl_e.pack(anchor="w", padx=16)

        # Agent
        label(box, "Agent réceptionnaire", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(8, 2))
        self.agent_cb = combo(box, 460)
        self.agent_cb.configure(values=app.db.agents() or ["Admin"])
        self.agent_cb.set(app.db.last_agent or "Admin")
        self.agent_cb.pack(anchor="w", padx=16)

        # Note
        label(box, "Remarque / Lot", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(8, 2))
        self.note_e = entry(box, 460, placeholder_text="Optionnel (ex: Lot #45, contrôle qualité OK)...")
        self.note_e.pack(anchor="w", padx=16, pady=(0, 10))

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=22, pady=(0, 16))
        button(btns, "Annuler", PANEL, self.destroy, width=120, height=42).pack(side="left")
        button(btns, "✔  Valider la Réception", GREEN, self.validate, height=42).pack(side="right", fill="x", expand=True, padx=(10, 0))

        self.bind("<Return>", lambda e: self.validate())
        self.bind("<Escape>", lambda e: self.destroy())

    def validate(self):
        try:
            raw_art = self.art_cb.get().strip()
            if not raw_art:
                raise ValueError("Veuillez sélectionner un article.")
            # Extraire la référence ou désignation
            parts = raw_art.split(" - ", 1)
            target = None
            if len(parts) == 2:
                target = self.app.db.article_by_ref_or_barcode(parts[0])
            if not target:
                target = self.app.db.article_by_name(parts[-1])
            if not target:
                raise ValueError("Article sélectionné introuvable dans la base.")

            qty = to_float(self.qty_e.get(), "Quantité")
            if qty <= 0:
                raise ValueError("La quantité reçue doit être strictement positive.")

            fourn = self.fourn_e.get().strip()
            bl = self.bl_e.get().strip()
            agent = self.agent_cb.get().strip() or "Admin"
            note = self.note_e.get().strip()

            self.app.db.add_stock(
                target["id"], qty, agent=agent, note=note,
                fournisseur=fourn, bon_fournisseur=bl
            )
        except Exception as ex:
            fail(ex, self)
            return

        self.app.refresh_all()
        self.app.msg(f"Entrée de +{qty:g} {target.get('unite', 'pcs')} enregistrée pour « {target['designation']} » dans Excel.")
        self.destroy()


class SettingsDialog(ctk.CTkToplevel):
    """Configuration de l'entreprise et des paramètres de l'application."""

    def __init__(self, app):
        super().__init__(app, fg_color=BG)
        self.app = app
        self.title("Paramètres de l'entreprise & Devise")
        self.geometry("480x420")
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()

        self.update_idletasks()
        x = app.winfo_x() + (app.winfo_width() - 480) // 2
        y = app.winfo_y() + (app.winfo_height() - 420) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

        label(self, "⚙️  Paramètres de l'application", BLUE, 18, True).pack(anchor="w", padx=24, pady=(18, 10))

        box = panel(self)
        box.pack(fill="both", expand=True, padx=22, pady=(0, 16))

        label(box, "Nom de l'entreprise / Société", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(12, 2))
        self.soc_e = entry(box, 420)
        self.soc_e.insert(0, app.db.company_name)
        self.soc_e.pack(anchor="w", padx=16)

        label(box, "Symbole Devise (ex: DT, €, $, MAD, DZD, CFA)", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(10, 2))
        self.cur_e = entry(box, 420)
        self.cur_e.insert(0, app.db.currency)
        self.cur_e.pack(anchor="w", padx=16)

        label(box, "Mode de calcul de l'option 10%", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(10, 2))
        self.mode_cb = combo(box, 420, values=["remise", "frais"])
        self.mode_cb.set(app.db.mode10)
        self.mode_cb.pack(anchor="w", padx=16)

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=22, pady=(0, 16))
        button(btns, "Annuler", PANEL, self.destroy, width=120, height=42).pack(side="left")
        button(btns, "💾  Sauvegarder", BLUE, self.save, height=42).pack(side="right", fill="x", expand=True, padx=(10, 0))

    def save(self):
        soc = self.soc_e.get().strip() or "THINK POWER"
        cur = self.cur_e.get().strip() or "DT"
        mode = self.mode_cb.get().strip() or "remise"

        self.app.db.set_setting("societe_nom", soc)
        self.app.db.set_setting("devise", cur)
        self.app.db.set_setting("mode10", mode)

        self.app.refresh_all()
        self.app.msg(f"Paramètres sauvegardés : Devise = {cur}, Entreprise = {soc}.")
        self.destroy()


# ----------------------------------------------------------------- Application Principale
class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=BG)
        ctk.set_appearance_mode("dark")

        # Configuration Treeview ttk
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background=FIELD,
            fieldbackground=FIELD,
            foreground=TEXT_WHITE,
            rowheight=30,
            borderwidth=0,
            font=("Segoe UI", 10)
        )
        style.configure(
            "Treeview.Heading",
            background=PANEL,
            foreground=TEXT_WHITE,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padding=6
        )
        style.map("Treeview", background=[("selected", BLUE)], foreground=[("selected", TEXT_WHITE)])
        style.map("Treeview.Heading", background=[("active", BLUE)])

        # Initialisation Base Excel
        self.db = DB()

        self.title(f"{self.db.company_name} — Gestion de Stock Intelligente")
        self.geometry("1280x760")
        self.minsize(1050, 650)

        # Icône de l'application
        ico = get_resource_path("app_icon.ico")
        if os.path.exists(ico):
            try:
                self.iconbitmap(ico)
            except Exception:
                pass

        # Centrage écran
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        start_x = max(0, (screen_w - 1280) // 2)
        start_y = max(0, (screen_h - 760) // 2)
        self.geometry(f"1280x760+{start_x}+{start_y}")

        self.sel_art = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.build_sidebar()

        main_area = ctk.CTkFrame(self, fg_color="transparent")
        main_area.grid(row=0, column=1, sticky="nsew")
        main_area.grid_columnconfigure(0, weight=1)
        main_area.grid_rowconfigure(1, weight=1)

        # En-tête supérieur
        top_bar = ctk.CTkFrame(main_area, fg_color="transparent", height=50)
        top_bar.pack(fill="x", padx=20, pady=(12, 0))

        self.page_title = label(top_bar, "", TEXT_WHITE, 20, True)
        self.page_title.pack(side="left")

        # Badge Base Excel & Devise
        db_badge = ctk.CTkFrame(top_bar, fg_color=CARD, corner_radius=8, border_width=1, border_color=BORDER)
        db_badge.pack(side="right", padx=(10, 0))
        label(db_badge, "🟢 Base Excel :", GREEN, 11, True).pack(side="left", padx=(10, 4), pady=5)
        label(db_badge, os.path.basename(DB_PATH), TEXT_WHITE, 11).pack(side="left", padx=(0, 6), pady=5)
        button(db_badge, "🔄", PANEL, self.sync_reload, width=32, height=26).pack(side="left", padx=4, pady=3)

        self.clock = label(top_bar, "", TEXT_MUTED, 11)
        self.clock.pack(side="right", padx=10)

        # Corps dynamique des pages
        self.body = ctk.CTkFrame(main_area, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=20, pady=(6, 4))

        # Barre d'état inférieure
        self.status = label(main_area, "", TEXT_MUTED, 11, anchor="w")
        self.status.pack(side="bottom", fill="x", padx=20, pady=(2, 6))

        # Pages
        self.pages = {
            "stock": self.build_stock(),
            "sortie": self.build_sortie(),
            "smart": self.build_smart(),
            "histo": self.build_histo()
        }

        self.go("stock")
        self.tick()

    # ------------------------------------------------------------- Navigation
    def build_sidebar(self):
        side = ctk.CTkFrame(self, fg_color=SIDE, corner_radius=0, width=235)
        side.grid(row=0, column=0, sticky="ns")
        side.grid_propagate(False)

        brand_box = ctk.CTkFrame(side, fg_color="transparent")
        brand_box.pack(fill="x", padx=18, pady=(20, 16))
        self.brand_lbl = label(brand_box, f"⚡ {self.db.company_name.upper()}", BLUE, 16, True)
        self.brand_lbl.pack(anchor="w")
        label(brand_box, "Smart Stock Manager", TEXT_MUTED, 11).pack(anchor="w")

        self.nav = {}
        items = (
            ("stock", "📦   Stock & Articles"),
            ("sortie", "📤   Sortie de stock"),
            ("smart", "🧠   Smart Analytics & Réappro"),
            ("histo", "📋   Historique & Journal")
        )

        for key, text in items:
            b = ctk.CTkButton(
                side,
                text=text,
                anchor="w",
                height=42,
                corner_radius=8,
                fg_color="transparent",
                hover_color="#1f293d",
                font=("Segoe UI", 12, "bold"),
                text_color=TEXT_WHITE,
                command=lambda k=key: self.go(k)
            )
            b.pack(fill="x", padx=10, pady=2)
            self.nav[key] = b

        sep = ctk.CTkFrame(side, height=1, fg_color=BORDER)
        sep.pack(side="bottom", fill="x", padx=12, pady=(0, 10))

        # Raccourcis bas de page
        button(side, "⚙️  Paramètres & Devise", PANEL, lambda: SettingsDialog(self), height=36).pack(side="bottom", fill="x", padx=12, pady=(0, 6))
        button(side, "📂  Ouvrir Base Excel", GREEN, self.open_db, height=36).pack(side="bottom", fill="x", padx=12, pady=(0, 6))
        button(side, "🗂  Dossier Application", PANEL, lambda: self.open_path(BASE_DIR), height=36).pack(side="bottom", fill="x", padx=12, pady=(0, 6))

    TITLES = {
        "stock": "📦  Gestion Universelle du Stock & Articles",
        "sortie": "📤  Enregistrement des Sorties & Scanner",
        "smart": "🧠  Tableau de Bord Intelligent & Conseiller de Réapprovisionnement",
        "histo": "📋  Historique des Mouvements & Rapports"
    }

    def go(self, key):
        for f in self.pages.values():
            f.pack_forget()
        self.pages[key].pack(fill="both", expand=True)
        for k, b in self.nav.items():
            b.configure(fg_color=BLUE if k == key else "transparent")
        self.page_title.configure(text=self.TITLES[key])
        self.refresh_all()

    def tick(self):
        self.clock.configure(text=datetime.now().strftime("%d/%m/%Y   %H:%M:%S"))
        self.after(1000, self.tick)

    def msg(self, text):
        cur = self.db.currency
        self.status.configure(text=f"●  {text}   |   Devise active : {cur}   |   Fichier : {DB_PATH}")

    def open_path(self, path):
        try:
            if not os.path.exists(path):
                fail(f"Le fichier ou dossier spécifié n'existe pas :\n{path}")
                return
            if os.name == "nt":
                os.startfile(path)
            else:
                messagebox.showinfo(APP_NAME, path)
        except OSError as ex:
            if getattr(ex, "winerror", None) == 1155 or "1155" in str(ex):
                messagebox.showwarning(
                    APP_NAME,
                    f"Le fichier a bien été généré mais aucun logiciel n'est configuré pour l'ouvrir directement sur ce poste.\n\n"
                    f"Fichier : {path}\n\n"
                    "Vous pouvez l'ouvrir avec Excel, LibreOffice ou un lecteur PDF."
                )
            else:
                fail(f"Impossible d'ouvrir le fichier :\n{ex}")
        except Exception as ex:
            fail(f"Impossible d'ouvrir le chemin :\n{ex}")

    def open_db(self):
        self.open_path(DB_PATH)

    def sync_reload(self):
        try:
            self.db.reload()
            self.brand_lbl.configure(text=f"⚡ {self.db.company_name.upper()}")
            self.refresh_all()
            self.msg("Données rechargées avec succès depuis le fichier Excel.")
        except Exception as ex:
            fail(f"Erreur de rechargement :\n{ex}")

    def refresh_all(self):
        self.refresh_stock()
        self.refresh_pick()
        self.refresh_smart()
        if hasattr(self, "s_agent"):
            self.s_agent.configure(values=self.db.agents() or ["Admin"])
        self.refresh_histo()
        if not self.status.cget("text"):
            self.msg("Prêt — Enregistrement en temps réel activé.")

    # ------------------------------------------------------------- Page 1 : Stock
    def build_stock(self):
        f = panel(self.body)

        # Cartes KPI
        cards_box = ctk.CTkFrame(f, fg_color="transparent")
        cards_box.pack(fill="x", padx=14, pady=(12, 4))

        self.kpi = {}
        for key, text, icon, color in (
            ("refs", "Total Références", "📦", BLUE),
            ("units", "Unités Physiques", "🔢", PURPLE),
            ("ttc", "Valeur Totale TTC", "💰", CYAN),
            ("alert", "Alertes (Rupture / Faible)", "⚠️", ORANGE),
            ("sorties", "Sorties du mois", "📉", GREEN)
        ):
            c, v = card(cards_box, text, icon, color)
            c.pack(side="left", expand=True, fill="x", padx=3)
            self.kpi[key] = v

        # Barre de recherche & filtres
        bar = ctk.CTkFrame(f, fg_color="transparent")
        bar.pack(fill="x", padx=14, pady=(4, 6))

        self.search = entry(bar, 260, placeholder_text="🔎  Réf, désignation, rayon, note...")
        self.search.pack(side="left")
        self.search.bind("<KeyRelease>", lambda e: self.refresh_stock())

        self.cat_filter = ctk.CTkComboBox(
            bar,
            values=["Toutes les catégories"],
            width=190,
            height=36,
            fg_color=FIELD,
            border_color=BORDER,
            dropdown_fg_color=PANEL,
            command=lambda v: self.refresh_stock()
        )
        self.cat_filter.pack(side="left", padx=8)

        self.flt = ctk.CTkSegmentedButton(
            bar,
            values=list(FILTERS),
            command=lambda v: self.refresh_stock(),
            selected_color=BLUE,
            selected_hover_color=BLUE_HOVER,
            height=36,
            corner_radius=8
        )
        self.flt.set("Tous")
        self.flt.pack(side="left", padx=4)

        # Boutons d'action principaux
        button(bar, "➕  Nouvel Article", BLUE, lambda: ArticleDialog(self), width=130, height=36).pack(side="right", padx=3)
        button(bar, "📥  Réception Fournisseur", GREEN, lambda: ReceptionDialog(self), width=165, height=36).pack(side="right", padx=3)

        # Tableau des articles
        cols = ("Réf", "Désignation", "Catégorie", "Unité", "En Stock", "Seuil", "Emplacement", "PU HT", "PU TTC", "Total TTC", "État")
        box, self.tree = make_tree(f, cols, (95, 230, 130, 60, 75, 65, 110, 85, 85, 95, 95))
        box.pack(fill="both", expand=True, padx=14, pady=4)
        self.tree.bind("<Double-1>", lambda e: self.edit_article())

        # Barre d'actions secondaires
        action_bar = ctk.CTkFrame(f, fg_color="transparent")
        action_bar.pack(fill="x", padx=14, pady=(4, 10))

        button(action_bar, "📤  Sortir cet article", ORANGE, self.sortie_from_stock, width=150).pack(side="left", padx=(0, 6))
        button(action_bar, "✏️  Modifier", PANEL, self.edit_article, width=100).pack(side="left", padx=3)
        button(action_bar, "🧮  Inventaire Physique", CYAN, self.open_inventory, width=160).pack(side="left", padx=3)
        button(action_bar, "📗  Importer Excel", PANEL, self.import_excel, width=130).pack(side="left", padx=3)
        button(action_bar, "🗑  Supprimer", PANEL, self.del_article, width=110).pack(side="right")

        return f

    def refresh_stock(self):
        self.tree.delete(*self.tree.get_children())
        mode = self.db.mode10
        cur = self.db.currency

        # Actualiser les catégories disponibles dans le filtre
        cats = ["Toutes les catégories"] + self.db.categories()
        self.cat_filter.configure(values=cats)

        cat_sel = self.cat_filter.get()
        filter_cat = "" if cat_sel == "Toutes les catégories" else cat_sel

        arts = self.db.articles(self.search.get(), FILTERS[self.flt.get()], categorie=filter_cat)

        for i, a in enumerate(arts):
            c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], mode)
            e = etat(a)
            pu_ttc = a["pu_ht"] * (1 + a["taux_tva"] / 100)
            self.tree.insert(
                "", "end",
                iid=str(a["id"]),
                tags=(ETAT_TAG[e], "odd" if i % 2 else "even"),
                values=(
                    a.get("ref", f"REF-{a['id']:04d}"),
                    a["designation"],
                    a.get("categorie", "Général"),
                    a.get("unite", "pcs"),
                    f"{a['nombre_existant']:g}",
                    f"{a['seuil_alerte']:g}",
                    a.get("emplacement", "-"),
                    f"{a['pu_ht']:.3f}",
                    f"{pu_ttc:.3f}",
                    f"{c['ttc']:.3f}",
                    ETAT_LABEL[e]
                )
            )

        k = self.db.kpis()
        self.kpi["refs"].configure(text=str(k["refs"]))
        self.kpi["units"].configure(text=f"{k['total_units']:g}")
        self.kpi["ttc"].configure(text=f"{k['ttc']:.3f} {cur}")
        self.kpi["alert"].configure(
            text=f"{k['rupt']} ruptures  |  {k['crit']} faibles",
            text_color=RED if k["rupt"] else (ORANGE if k["crit"] else GREEN)
        )
        self.kpi["sorties"].configure(text=f"{k['sorties']:g}")

    def selected_article(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Veuillez d'abord sélectionner un article dans le tableau.")
            return None
        return self.db.article(int(sel[0]))

    def edit_article(self):
        a = self.selected_article()
        if a:
            ArticleDialog(self, a)

    def sortie_from_stock(self):
        a = self.selected_article()
        if a:
            self.sel_art = a["id"]
            self.go("sortie")
            self.set_article(a["id"])

    def del_article(self):
        a = self.selected_article()
        if not a:
            return

        conf = messagebox.askyesno(
            APP_NAME,
            f"Voulez-vous vraiment supprimer définitivement « {a['designation']} » ?\n\n"
            "Cette ligne sera supprimée immédiatement du fichier Excel.",
            icon="warning"
        )
        if not conf:
            return

        try:
            self.db.delete_article(a["id"])
        except Exception as ex:
            fail(ex)
            return

        self.refresh_all()
        self.msg(f"Article « {a['designation']} » supprimé du fichier Excel.")

    def import_excel(self):
        path = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel ou CSV de stock",
            filetypes=[("Fichiers Excel", "*.xlsx *.xlsm")]
        )
        if not path:
            return

        try:
            n = self.db.import_excel(path)
        except Exception as ex:
            fail(f"Importation impossible :\n{ex}")
            return

        self.refresh_all()
        self.msg(f"{n} article(s) importé(s) / mis à jour depuis {os.path.basename(path)}")
        messagebox.showinfo(APP_NAME, f"{n} article(s) importé(s) avec succès dans Excel !")

    def open_inventory(self):
        d = ctk.CTkToplevel(self, fg_color=BG)
        d.title("Inventaire physique & régularisation")
        d.geometry("480x480")
        d.resizable(False, False)
        d.transient(self)
        d.grab_set()

        d.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 480) // 2
        y = self.winfo_y() + (self.winfo_height() - 480) // 2
        d.geometry(f"+{max(0, x)}+{max(0, y)}")

        label(d, "🧮  Régularisation d'inventaire", CYAN, 18, True).pack(anchor="w", padx=24, pady=(18, 8))
        box = panel(d)
        box.pack(fill="both", expand=True, padx=22, pady=(0, 16))

        a0 = self.db.article(int(self.tree.selection()[0])) if self.tree.selection() else None

        label(box, "Article à vérifier", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(14, 2))
        art_cb = combo(box, 400)
        art_cb.configure(values=[f"{a.get('ref','')} - {a['designation']}" for a in self.db.articles()] or [""])
        if a0:
            art_cb.set(f"{a0.get('ref','')} - {a0['designation']}")
        art_cb.pack(anchor="w", padx=16)

        label(box, "Quantité physique réelle comptée", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(10, 2))
        qty_e = entry(box, 400)
        qty_e.pack(anchor="w", padx=16)

        label(box, "Agent responsable", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(10, 2))
        agent_cb = combo(box, 400)
        agent_cb.configure(values=self.db.agents() or ["Admin"])
        agent_cb.set(self.db.last_agent or "Admin")
        agent_cb.pack(anchor="w", padx=16, pady=(0, 16))

        def validate():
            try:
                raw = art_cb.get().strip()
                name = raw.split(" - ", 1)[-1] if " - " in raw else raw
                a = self.db.article_by_name(name)
                if not a:
                    raise ValueError("Veuillez sélectionner un article valide.")
                phys = to_float(qty_e.get(), "Quantité physique")
                agent = agent_cb.get().strip() or "Admin"
                ecart = phys - a["nombre_existant"]
                c = calc(abs(ecart), a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], self.db.mode10)
                cur = self.db.currency

                conf = messagebox.askyesno(
                    APP_NAME,
                    f"Article : {a['designation']}\n"
                    f"Stock théorique : {a['nombre_existant']:g} {a.get('unite','pcs')}\n"
                    f"Stock physique : {phys:g} {a.get('unite','pcs')}\n"
                    f"Écart : {ecart:+g}  (Valeur : {c['ttc']:.3f} {cur} TTC)\n\n"
                    "Enregistrer la régularisation directement dans Excel ?",
                    parent=d
                )
                if not conf:
                    return

                self.db.add_movement(a["id"], "INVENTAIRE", phys, agent, "", f"Régularisation inventaire (écart {ecart:+g})")
            except Exception as ex:
                fail(ex, d)
                return

            self.refresh_all()
            self.msg(f"Inventaire enregistré par {agent} dans Excel.")
            d.destroy()

        btns = ctk.CTkFrame(d, fg_color="transparent")
        btns.pack(fill="x", padx=22, pady=(0, 16))
        button(btns, "Annuler", PANEL, d.destroy, width=110, height=42).pack(side="left")
        button(btns, "✔  Régulariser dans Excel", CYAN, validate, height=42).pack(side="right", fill="x", expand=True, padx=(10, 0))

    # ------------------------------------------------------------- Page 2 : Sortie
    def build_sortie(self):
        f = ctk.CTkFrame(self.body, fg_color="transparent")
        f.grid_columnconfigure(0, weight=4)
        f.grid_columnconfigure(1, weight=5)
        f.grid_rowconfigure(0, weight=1)

        left = panel(f)
        right = panel(f)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        # ---- Formulaire Sortie (Gauche)
        label(left, "📤  Nouvelle Sortie de Stock", ORANGE, 17, True).pack(anchor="w", padx=18, pady=(14, 6))

        # Carte article sélectionné
        self.art_box = ctk.CTkFrame(left, fg_color=CARD, corner_radius=10, border_width=1, border_color=BORDER)
        self.art_box.pack(fill="x", padx=18, pady=(4, 8))
        label(self.art_box, "Matière / Article sélectionné :", TEXT_MUTED, 11).pack(anchor="w", padx=12, pady=(6, 1))
        self.s_art = label(self.art_box, "— Aucun article sélectionné —", CYAN, 13, True, wraplength=340, justify="left")
        self.s_art.pack(anchor="w", padx=12)
        self.s_exist = label(self.art_box, "Scanner un code barre ou choisir à droite →", TEXT_MUTED, 11)
        self.s_exist.pack(anchor="w", padx=12, pady=(2, 6))

        # Scanner ou recherche directe par code-barres / référence
        scan_bar = ctk.CTkFrame(left, fg_color="transparent")
        scan_bar.pack(fill="x", padx=18, pady=(2, 6))
        label(scan_bar, "⚡ Code-barres ou Réf :", TEXT_MUTED, 11).pack(side="left")
        self.scan_e = entry(scan_bar, 180, placeholder_text="Scanner ou taper...")
        self.scan_e.pack(side="left", padx=6)
        self.scan_e.bind("<KeyRelease>", self.on_barcode_scan)

        # Quantité
        self.qty_label = label(left, "Quantité demandée *", TEXT_MUTED, 11)
        self.qty_label.pack(anchor="w", padx=18, pady=(4, 2))
        q_box = ctk.CTkFrame(left, fg_color="transparent")
        q_box.pack(anchor="w", padx=18)
        button(q_box, "−", BLUE, lambda: self.step(-1), width=44, height=36).pack(side="left")
        self.s_qty = entry(q_box, 110, justify="center")
        self.s_qty.pack(side="left", padx=6)
        self.s_qty.bind("<KeyRelease>", lambda e: self.check_qty())
        button(q_box, "+", BLUE, lambda: self.step(1), width=44, height=36).pack(side="left")

        self.s_live = label(left, "", TEXT_MUTED, 11, True)
        self.s_live.pack(anchor="w", padx=18, pady=(2, 4))

        # Bénéficiaire & Agent
        label(left, "Société / Client bénéficiaire *", TEXT_MUTED, 11).pack(anchor="w", padx=18, pady=(4, 2))
        self.s_soc = entry(left, 360, placeholder_text="Ex: Entreprise ABC, Service Chantier...")
        self.s_soc.pack(fill="x", padx=18)

        label(left, "Agent émetteur *", TEXT_MUTED, 11).pack(anchor="w", padx=18, pady=(6, 2))
        self.s_agent = combo(left, 360)
        self.s_agent.configure(values=self.db.agents() or ["Admin"])
        self.s_agent.set(self.db.last_agent or "Admin")
        self.s_agent.pack(fill="x", padx=18)

        label(left, "Date & Heure", TEXT_MUTED, 11).pack(anchor="w", padx=18, pady=(6, 2))
        self.s_date = entry(left, 360)
        self.s_date.insert(0, datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.s_date.pack(fill="x", padx=18)

        label(left, "Note / Justification", TEXT_MUTED, 11).pack(anchor="w", padx=18, pady=(6, 2))
        self.s_note = entry(left, 360, placeholder_text="Optionnel (ex: Projet Z, Bon commande #12)...")
        self.s_note.pack(fill="x", padx=18)

        button(left, "✔  VALIDER LA SORTIE (EXCEL)", ORANGE, self.do_sortie, height=42).pack(fill="x", padx=18, pady=(12, 14))

        # ---- Sélecteur Articles & Aperçu (Droite)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(2, weight=3)
        right.grid_rowconfigure(4, weight=2)

        label(right, "1 · Sélectionner l'article dans le stock", TEXT_WHITE, 14, True).grid(row=0, column=0, sticky="w", padx=16, pady=(14, 4))
        self.p_search = entry(right, 320, placeholder_text="🔎  Réf ou désignation...")
        self.p_search.grid(row=1, column=0, sticky="w", padx=16, pady=(0, 6))
        self.p_search.bind("<KeyRelease>", lambda e: self.refresh_pick())

        box, self.pick = make_tree(right, ("Réf", "Désignation", "En Stock", "État"), (90, 220, 80, 90))
        box.grid(row=2, column=0, sticky="nsew", padx=16)
        self.pick.bind("<<TreeviewSelect>>", self.on_pick)

        label(right, "2 · Détails & Historique récent de l'article", TEXT_WHITE, 14, True).grid(row=3, column=0, sticky="w", padx=16, pady=(8, 4))
        self.console = ctk.CTkTextbox(right, fg_color=CONSOLE, font=("Consolas", 11), corner_radius=10, state="disabled")
        self.console.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 14))
        self.say("Sélectionnez ou scannez un article pour afficher sa fiche complète.")

        return f

    def say(self, text):
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.insert("1.0", text)
        self.console.configure(state="disabled")

    def on_barcode_scan(self, event=None):
        code = self.scan_e.get().strip()
        if not code:
            return
        a = self.db.article_by_ref_or_barcode(code)
        if a:
            self.set_article(a["id"])

    def refresh_pick(self):
        self.pick.unbind("<<TreeviewSelect>>")
        self.pick.delete(*self.pick.get_children())
        arts = self.db.articles(self.p_search.get())
        for i, a in enumerate(arts):
            e = etat(a)
            self.pick.insert(
                "", "end",
                iid=str(a["id"]),
                tags=(ETAT_TAG[e], "odd" if i % 2 else "even"),
                values=(
                    a.get("ref", ""),
                    a["designation"],
                    f"{a['nombre_existant']:g} {a.get('unite', '')}",
                    ETAT_LABEL[e]
                )
            )
        if self.sel_art and self.pick.exists(str(self.sel_art)):
            self.pick.selection_set(str(self.sel_art))
            self.pick.see(str(self.sel_art))
        self.pick.after(100, lambda: self.pick.bind("<<TreeviewSelect>>", self.on_pick))

    def on_pick(self, _=None):
        sel = self.pick.selection()
        if sel:
            self.set_article(int(sel[0]))

    def set_article(self, aid):
        a = self.db.article(aid)
        if not a:
            return
        self.sel_art = aid
        cur = self.db.currency
        u = a.get("unite", "pcs")
        emp = a.get("emplacement", "Non défini")
        self.s_art.configure(text=f"[{a.get('ref', '')}] {a['designation']}")
        self.s_exist.configure(
            text=f"Stock : {a['nombre_existant']:g} {u}   (Seuil : {a['seuil_alerte']:g} {u})   |   Rayon : {emp}"
        )
        self.qty_label.configure(text=f"Quantité demandée ({u}) *")
        if not self.s_qty.get().strip():
            self.s_qty.insert(0, "1")

        c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], self.db.mode10)
        pu_ttc = a["pu_ht"] * (1 + a["taux_tva"] / 100)

        txt = (
            f"Article         : {a['designation']}\n"
            f"Référence / SKU : {a.get('ref', '-')}\n"
            f"Catégorie       : {a.get('categorie', 'Général')}\n"
            f"Emplacement     : {emp}\n"
            f"Stock disponible: {a['nombre_existant']:g} {u}\n"
            f"Prix unitaire HT: {a['pu_ht']:.3f} {cur}\n"
            f"Prix unitaire TTC: {pu_ttc:.3f} {cur}\n"
            f"Valeur en stock : {c['ttc']:.3f} {cur} TTC\n"
        )
        if etat(a) != "ok":
            txt += "\n⚠️  ATTENTION : NIVEAU CRITIQUE OU EN RUPTURE !\n"

        txt += "\nDerniers mouvements enregistrés :\n"
        for m in self.db.article_movements(aid, limit=5):
            q = -m["qte"] if m["type"] == "SORTIE" else m["qte"]
            txt += f"  • {fmt_date(m['date'])[:10]}  {m['type']:<11} {q:+g} {u}  ({m.get('agent', '-')})\n"

        self.say(txt)
        self.check_qty()

    def step(self, d):
        try:
            v = to_float(self.s_qty.get(), "")
        except ValueError:
            v = 0
        self.s_qty.delete(0, "end")
        self.s_qty.insert(0, f"{max(0, v + d):g}")
        self.check_qty()

    def check_qty(self):
        a = self.db.article(self.sel_art) if self.sel_art else None
        if not a:
            self.s_live.configure(text="")
            return
        try:
            q = to_float(self.s_qty.get(), "")
        except ValueError:
            self.s_live.configure(text="Quantité invalide", text_color=ORANGE)
            return

        u = a.get("unite", "pcs")
        if q > a["nombre_existant"]:
            self.s_live.configure(text=f"⛔ Stock insuffisant ({a['nombre_existant']:g} {u} dispo)", text_color=RED)
        else:
            reste = a["nombre_existant"] - q
            self.s_live.configure(text=f"✔ Reste après sortie : {reste:g} {u}", text_color=GREEN)

    def do_sortie(self):
        try:
            agent = self.s_agent.get().strip() or "Admin"
            soc = self.s_soc.get().strip()
            if not self.sel_art:
                raise ValueError("Veuillez d'abord choisir un article dans la liste à droite ou le scanner.")
            if not soc:
                raise ValueError("Le nom de la société ou client bénéficiaire est obligatoire.")
            qty = to_float(self.s_qty.get(), "Quantité")

            date_str = self.s_date.get().strip()
            date_iso = user_date(date_str) if date_str else now_iso()

            m = self.db.add_movement(
                self.sel_art, "SORTIE", qty, agent, soc,
                self.s_note.get().strip(), date_iso
            )
        except Exception as ex:
            fail(ex)
            return

        a = self.db.article(self.sel_art)
        cur = self.db.currency
        u = a.get("unite", "pcs")
        txt = (
            f"✅  SORTIE ENREGISTRÉE DANS EXCEL — N° Bon : {m['bon']}\n\n"
            f"Article  : {a['designation']} ({a.get('ref', '')})\n"
            f"Bénéf.   : {soc}\n"
            f"Agent    : {agent}\n"
            f"Quantité : {m['qte']:g} {u}\n"
            f"Reste    : {m['apres']:g} {u}\n"
            f"Total TTC: {m['total_ttc']:.3f} {cur}\n"
        )
        if etat(a) != "ok":
            txt += "\n⚠️  STOCK FAIBLE / RUPTURE : Penser à commander !\n"

        self.s_qty.delete(0, "end")
        self.s_note.delete(0, "end")
        self.s_date.delete(0, "end")
        self.s_date.insert(0, datetime.now().strftime("%d/%m/%Y %H:%M"))

        self.refresh_all()
        self.s_exist.configure(text=f"Stock actuel : {a['nombre_existant']:g} {u}   (Seuil : {a['seuil_alerte']:g} {u})")
        self.s_live.configure(text="")
        self.say(txt)
        self.msg(f"Sortie {m['bon']} enregistrée par {agent} dans Excel.")

        # Proposition de Bon PDF
        if messagebox.askyesno(APP_NAME, f"Sortie {m['bon']} validée !\n\nVoulez-vous ouvrir le Bon de Sortie officiel en PDF ?"):
            try:
                os.makedirs(EXPORT_DIR, exist_ok=True)
                pdf_path = os.path.join(EXPORT_DIR, f"Bon_{m['bon']}.pdf")
                export_bon_pdf(self.db, m, pdf_path)
                self.open_path(pdf_path)
            except Exception as ex:
                fail(f"Erreur PDF :\n{ex}")

    # ------------------------------------------------------------- Page 3 : Smart Advisor
    def build_smart(self):
        f = panel(self.body)

        top_bar = ctk.CTkFrame(f, fg_color="transparent")
        top_bar.pack(fill="x", padx=16, pady=(14, 6))

        label(top_bar, "🧠  Conseiller Intelligent de Réapprovisionnement", CYAN, 17, True).pack(side="left")
        self.reorder_badge = label(top_bar, "", ORANGE, 13, True)
        self.reorder_badge.pack(side="left", padx=16)

        button(top_bar, "📄  Télécharger Bon de Commande Suggéré (PDF)", BLUE, self.do_reorder_pdf, width=280).pack(side="right")

        # Zone centrale partagée (Gauche: Tableau Réappro, Droite: Insights & Rotation)
        center = ctk.CTkFrame(f, fg_color="transparent")
        center.pack(fill="both", expand=True, padx=14, pady=6)
        center.grid_columnconfigure(0, weight=6)
        center.grid_columnconfigure(1, weight=4)
        center.grid_rowconfigure(0, weight=1)

        # Tableau des articles à commander
        left_box = ctk.CTkFrame(center, fg_color="transparent")
        left_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        label(left_box, "Articles sous le seuil de sécurité nécessitant une commande :", TEXT_MUTED, 11).pack(anchor="w", pady=(0, 4))

        cols = ("Réf", "Désignation", "Rayon", "En Stock", "Seuil", "Qté Suggérée", "Budget Estimé TTC", "Statut")
        t_box, self.rtree = make_tree(left_box, cols, (80, 180, 95, 75, 65, 80, 110, 85))
        t_box.pack(fill="both", expand=True)

        # Panneau de droite : Insights & Rotation
        right_box = panel(center)
        right_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        label(right_box, "🔥  Top 5 - Articles Forte Rotation (Sorties 30j)", GREEN, 13, True).pack(anchor="w", padx=14, pady=(12, 4))
        self.top_box = ctk.CTkTextbox(right_box, fg_color=FIELD, font=("Consolas", 11), height=140, corner_radius=8, state="disabled")
        self.top_box.pack(fill="x", padx=14, pady=(0, 10))

        label(right_box, "💤  Articles Dormants (Sur-stockage - Aucun mouvement 30j)", ORANGE, 13, True).pack(anchor="w", padx=14, pady=(4, 4))
        self.dead_box = ctk.CTkTextbox(right_box, fg_color=FIELD, font=("Consolas", 11), height=150, corner_radius=8, state="disabled")
        self.dead_box.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        return f

    def refresh_smart(self):
        self.rtree.delete(*self.rtree.get_children())
        analysis = self.db.smart_analysis()
        cur = self.db.currency
        reorder = analysis["reorder_list"]

        self.reorder_badge.configure(
            text=f"{len(reorder)} article(s) à réapprovisionner  |  Budget estimé : {analysis['reorder_total_cost']:.3f} {cur}"
        )

        for i, r in enumerate(reorder):
            tag = "rupture" if r["statut"] == "Rupture" else "critique"
            self.rtree.insert(
                "", "end",
                iid=str(r["id"]),
                tags=(tag, "odd" if i % 2 else "even"),
                values=(
                    r["ref"],
                    r["designation"],
                    r["emplacement"] or r["categorie"],
                    f"{r['actuel']:g} {r['unite']}",
                    f"{r['seuil']:g} {r['unite']}",
                    f"{r['suggested']:g} {r['unite']}",
                    f"{r['cost_ttc']:.3f} {cur}",
                    r["statut"]
                )
            )

        # Top sorties
        self.top_box.configure(state="normal")
        self.top_box.delete("1.0", "end")
        if analysis["top_sorties"]:
            for item in analysis["top_sorties"]:
                a = item["article"]
                self.top_box.insert("end", f" • {a['designation']:<26} : {item['qte']:g} {a.get('unite','pcs')} sorties\n")
        else:
            self.top_box.insert("end", "Aucune sortie enregistrée ce mois-ci.\n")
        self.top_box.configure(state="disabled")

        # Articles dormants
        self.dead_box.configure(state="normal")
        self.dead_box.delete("1.0", "end")
        if analysis["dead_stock"]:
            for a in analysis["dead_stock"]:
                self.dead_box.insert("end", f" • [{a.get('ref','-')}] {a['designation']:<24} ({a['nombre_existant']:g} {a.get('unite','pcs')})\n")
        else:
            self.dead_box.insert("end", "Aucun article dormant identifié.\n")
        self.dead_box.configure(state="disabled")

    def do_reorder_pdf(self):
        os.makedirs(EXPORT_DIR, exist_ok=True)
        path = os.path.join(EXPORT_DIR, f"Bon_Commande_Suggere_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        try:
            export_reorder_pdf(self.db, path)
            self.open_path(path)
        except Exception as ex:
            fail(f"Impossible de générer le bon de commande :\n{ex}")

    # ------------------------------------------------------------- Page 4 : Historique
    def build_histo(self):
        f = panel(self.body)

        # Filtres
        bar = ctk.CTkFrame(f, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=(12, 6))

        self.h_search = entry(bar, 180, placeholder_text="Filtrer texte/bon/réf...")
        self.h_agent = entry(bar, 130, placeholder_text="Filtrer Agent...")
        self.h_soc = entry(bar, 150, placeholder_text="Filtrer Client/Fourn...")
        self.h_d1 = entry(bar, 120, placeholder_text="Du JJ/MM/AAAA")
        self.h_d2 = entry(bar, 120, placeholder_text="Au JJ/MM/AAAA")
        self.h_typ = ctk.CTkOptionMenu(
            bar,
            values=["Tous", "SORTIE", "ENTREE", "INVENTAIRE", "NOUVEAU", "MODIF", "SUPPRESSION", "IMPORT"],
            width=130,
            height=38,
            fg_color=PANEL,
            button_color=BLUE,
            button_hover_color=BLUE_HOVER
        )

        for w in (self.h_search, self.h_agent, self.h_soc, self.h_d1, self.h_d2, self.h_typ):
            w.pack(side="left", padx=(0, 6))

        button(bar, "🔎  Filtrer", BLUE, self.refresh_histo, width=85).pack(side="left", padx=3)
        button(bar, "↺  Effacer", PANEL, self.reset_histo_filters, width=85).pack(side="left", padx=2)

        # Tableau historique
        cols = ("N° Bon", "Date", "Type", "Réf", "Article", "Agent", "Client/Fourn.", "Avant", "Mvt", "Reste", "Total TTC", "Note")
        box, self.htree = make_tree(f, cols, (95, 115, 90, 85, 180, 100, 140, 50, 50, 50, 80, 160))
        box.pack(fill="both", expand=True, padx=16, pady=4)

        # Actions d'export & suppression directe
        ex_bar = ctk.CTkFrame(f, fg_color="transparent")
        ex_bar.pack(fill="x", padx=16, pady=(4, 12))

        button(ex_bar, "📗  Exporter Rapport Excel (.xlsx)", GREEN, lambda: self.do_export("xlsx"), width=230, height=38).pack(side="left", padx=(0, 10))
        button(ex_bar, "📕  Exporter Rapport PDF", RED, lambda: self.do_export("pdf"), width=210, height=38).pack(side="left")
        button(ex_bar, "🗑  Supprimer la ligne sélectionnée", PANEL, self.del_movement, width=240, height=38).pack(side="right")

        return f

    def reset_histo_filters(self):
        self.h_search.delete(0, "end")
        self.h_agent.delete(0, "end")
        self.h_soc.delete(0, "end")
        self.h_d1.delete(0, "end")
        self.h_d2.delete(0, "end")
        self.h_typ.set("Tous")
        self.refresh_histo()

    def histo_filters(self):
        t = self.h_typ.get()
        return dict(
            search=self.h_search.get().strip(),
            agent=self.h_agent.get().strip(),
            societe=self.h_soc.get().strip(),
            d1=user_date(self.h_d1.get()),
            d2=user_date(self.h_d2.get(), end=True),
            typ="" if t == "Tous" else t
        )

    def refresh_histo(self):
        try:
            flt = self.histo_filters()
        except ValueError as ex:
            fail(ex)
            return

        self.htree.delete(*self.htree.get_children())
        tags = {
            "SORTIE": "rupture",
            "SUPPRESSION": "rupture",
            "ENTREE": "ok",
            "NOUVEAU": "ok",
            "INVENTAIRE": "critique",
            "MODIF": "critique"
        }
        movs = self.db.movements(**flt)
        for i, m in enumerate(movs):
            q = f"{m.get('qte', 0):g}" if m.get("type") == "SORTIE" else f"{m.get('qte', 0):+g}"
            self.htree.insert(
                "", "end",
                iid=str(m["id"]),
                tags=(tags.get(m.get("type", ""), ""), "odd" if i % 2 else "even"),
                values=(
                    m.get("bon", ""),
                    fmt_date(m.get("date")),
                    m.get("type", ""),
                    m.get("ref", ""),
                    m.get("designation", ""),
                    m.get("agent", ""),
                    m.get("societe", ""),
                    f"{m.get('avant', 0):g}",
                    q,
                    f"{m.get('apres', 0):g}",
                    f"{m.get('total_ttc', 0):.3f}",
                    m.get("note", "")
                )
            )

    def del_movement(self):
        sel = self.htree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Veuillez sélectionner une ligne d'historique à supprimer.")
            return

        mid = int(sel[0])
        conf = messagebox.askyesno(
            APP_NAME,
            "Voulez-vous supprimer cette ligne de l'historique ?\n\n"
            "Elle sera effacée directement de la feuille Journal du fichier Excel.",
            icon="warning"
        )
        if not conf:
            return

        try:
            self.db.delete_movement(mid)
        except Exception as ex:
            fail(ex)
            return

        self.refresh_histo()
        self.msg("Ligne supprimée directement du fichier Excel.")

    def do_export(self, kind):
        try:
            flt = self.histo_filters()
        except ValueError as ex:
            fail(ex)
            return
        movs = self.db.movements(**flt)
        os.makedirs(EXPORT_DIR, exist_ok=True)
        path = os.path.join(EXPORT_DIR, f"Rapport_Stock_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{kind}")
        period = f"{self.h_d1.get() or 'début'} au {self.h_d2.get() or datetime.now().strftime('%d/%m/%Y')}"
        try:
            if kind == "pdf":
                export_pdf(self.db, movs, path, period)
            else:
                export_excel(self.db, movs, path)
        except Exception as ex:
            fail(f"Export impossible :\n{ex}")
            return
        self.open_path(path)


# ----------------------------------------------------------------- Point d'entrée
def excepthook(exc_type, exc_value, exc_traceback):
    """Gestionnaire global des erreurs non capturées pour les postes sans console."""
    err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    try:
        log_path = os.path.join(get_base_dir(), "erreur_crash.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"\n--- {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n{err_msg}\n")
    except Exception:
        pass
    try:
        messagebox.showerror(APP_NAME, f"Une erreur inattendue est survenue :\n\n{err_msg}")
    except Exception:
        pass
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


sys.excepthook = excepthook

if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except Exception as ex:
        messagebox.showerror(APP_NAME, f"Erreur critique au démarrage :\n{ex}")
