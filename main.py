# -*- coding: utf-8 -*-
"""Stock Think Power - Interface Graphique Moderne (CustomTkinter).
Base de données Excel autonome (Stock_Think_Power_Base.xlsx) située à côté du .exe.
"""
import os
import sys
import traceback
from datetime import datetime
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog, Menu

from stock_think_power import (
    DB, APP_NAME, EXPORT_DIR, DB_PATH, BASE_DIR, calc, etat, fmt_date, user_date,
    to_float, now_iso, export_pdf, export_excel, export_bon_pdf, get_resource_path, get_base_dir
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
PURPLE = "#8b5cf6"      # Violet inventaire

TEXT_WHITE = "#f8fafc"
TEXT_MUTED = "#94a3b8"

HOVER_MAP = {
    BLUE: BLUE_HOVER,
    CYAN: CYAN_HOVER,
    GREEN: GREEN_HOVER,
    RED: RED_HOVER,
    ORANGE: ORANGE_HOVER,
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
    val_lbl = label(f, "-", color, 20, True)
    val_lbl.pack(anchor="w", padx=14, pady=(2, 10))
    return f, val_lbl


def make_tree(parent, columns, widths):
    box = ctk.CTkFrame(parent, fg_color=PANEL, corner_radius=10)
    tree = ttk.Treeview(box, columns=columns, show="headings", selectmode="browse")
    left_align = {"Désignation", "Article", "Note", "Société / Client", "Société / Bénéficiaire", "Agent", "Remarque"}
    center_align = {"État", "Type", "ID", "N° Bon", "Bon", "TVA %", "TVA", "Option 10%"}

    for c, w in zip(columns, widths):
        tree.heading(c, text=c)
        if c in left_align:
            anchor = "w"
        elif c in center_align:
            anchor = "center"
        else:
            anchor = "e"
        tree.column(c, width=w, anchor=anchor, minwidth=40)

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
    """Ajout ou modification d'une fiche article."""

    def __init__(self, app, article=None):
        super().__init__(app, fg_color=BG)
        self.app, self.article = app, article
        a = article
        is_edit = bool(a)

        self.title("Modifier l'article" if is_edit else "Nouvel article")
        self.geometry("540x630")
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()

        # Centrage
        self.update_idletasks()
        x = app.winfo_x() + (app.winfo_width() - 540) // 2
        y = app.winfo_y() + (app.winfo_height() - 630) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

        title_txt = f"✏️  Modifier l'article #{a['id']}" if is_edit else "➕  Nouvel article en stock"
        label(self, title_txt, BLUE, 18, True).pack(anchor="w", padx=24, pady=(18, 10))

        form = panel(self)
        form.pack(fill="both", expand=True, padx=22, pady=(0, 16))
        form.grid_columnconfigure((0, 1), weight=1)
        self.e = {}

        def field(r, c, key, text, width, val, span=1):
            label(form, text, TEXT_MUTED, 11).grid(row=r * 2, column=c, columnspan=span, sticky="w", padx=16, pady=(8, 2))
            w = entry(form, width)
            w.insert(0, val)
            w.grid(row=r * 2 + 1, column=c, columnspan=span, sticky="ew", padx=16)
            self.e[key] = w

        label(form, "Agent responsable", TEXT_MUTED, 11).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(10, 2))
        self.agent = combo(form, 460)
        self.agent.configure(values=app.db.agents() or ["Admin"])
        self.agent.set(app.db.last_agent or "Admin")
        self.agent.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16)

        field(1, 0, "designation", "Matière / Désignation *", 460, a["designation"] if a else "", span=2)
        field(2, 0, "qty", "Quantité existante", 215, f"{a['nombre_existant']:g}" if a else "0")
        field(2, 1, "seuil", "Stock minimum (alerte)", 215, f"{a['seuil_alerte']:g}" if a else "5")
        field(3, 0, "pu", "Prix Unitaire HT (DT)", 215, f"{a['pu_ht']:.3f}" if a else "0.000")
        field(3, 1, "tva", "Taux TVA (%)", 215, f"{a['taux_tva']:g}" if a else "19")
        field(4, 0, "note", "Remarque / Note", 460, a.get("note", "") if a else "", span=2)

        self.opt10 = ctk.CTkCheckBox(
            form,
            text="Option 10% (selon paramètre : remise ou frais)",
            text_color=TEXT_WHITE,
            fg_color=BLUE,
            hover_color=BLUE_HOVER
        )
        if a and a.get("appliquer_10_pct"):
            self.opt10.select()
        self.opt10.grid(row=10, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 14))

        # Preview dynamique
        self.preview_lbl = label(form, "", CYAN, 11, bold=True)
        self.preview_lbl.grid(row=11, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 8))
        self.update_preview()

        for k in ("qty", "pu", "tva"):
            self.e[k].bind("<KeyRelease>", lambda e: self.update_preview())

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=22, pady=(0, 16))
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
            self.preview_lbl.configure(text=f"Aperçu : PU TTC = {pu * (1 + tva/100):.3f} DT   |   Total TTC = {c['ttc']:.3f} DT")
        except Exception:
            self.preview_lbl.configure(text="")

    def save(self):
        try:
            designation = self.e["designation"].get().strip()
            if not designation:
                raise ValueError("La désignation de l'article est obligatoire.")

            data = {
                "designation": designation,
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
        self.app.msg(f"Article « {data['designation']} » enregistré directement dans Excel.")
        self.destroy()


class QuickAddStockDialog(ctk.CTkToplevel):
    """Ajout direct de stock (réapprovisionnement / entrée rapide) depuis la page Stock."""

    def __init__(self, app, article):
        super().__init__(app, fg_color=BG)
        self.app, self.article = app, article
        self.title("Ajouter du Stock")
        self.geometry("460x420")
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()

        # Centrage
        self.update_idletasks()
        x = app.winfo_x() + (app.winfo_width() - 460) // 2
        y = app.winfo_y() + (app.winfo_height() - 420) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

        label(self, "➕  Entrée de Stock / Réappro", GREEN, 17, True).pack(anchor="w", padx=22, pady=(18, 6))

        box = panel(self)
        box.pack(fill="both", expand=True, padx=22, pady=(0, 14))

        label(box, f"Article : {article['designation']}", TEXT_WHITE, 14, True).pack(anchor="w", padx=16, pady=(14, 2))
        self.curr_lbl = label(box, f"Stock actuel disponible : {article['nombre_existant']:g}", CYAN, 12)
        self.curr_lbl.pack(anchor="w", padx=16, pady=(0, 10))

        label(box, "Quantité reçue à ajouter (+)", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(6, 2))
        self.qty_e = entry(box, 400)
        self.qty_e.insert(0, "10")
        self.qty_e.pack(anchor="w", padx=16)

        self.res_lbl = label(box, "", GREEN, 12, True)
        self.res_lbl.pack(anchor="w", padx=16, pady=(4, 6))
        self.qty_e.bind("<KeyRelease>", lambda e: self.update_res())

        label(box, "Agent responsable", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(4, 2))
        self.agent_cb = combo(box, 400)
        self.agent_cb.configure(values=app.db.agents() or ["Admin"])
        self.agent_cb.set(app.db.last_agent or "Admin")
        self.agent_cb.pack(anchor="w", padx=16)

        label(box, "Remarque / Réf bon de livraison", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(6, 2))
        self.note_e = entry(box, 400, placeholder_text="Ex: Arrivage fournisseur X, Bon #123...")
        self.note_e.pack(anchor="w", padx=16, pady=(0, 14))

        self.update_res()

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=22, pady=(0, 16))
        button(btns, "Annuler", PANEL, self.destroy, width=110, height=40).pack(side="left")
        button(btns, "✔  Valider l'entrée", GREEN, self.validate, height=40).pack(side="right", fill="x", expand=True, padx=(10, 0))

        self.bind("<Return>", lambda e: self.validate())
        self.bind("<Escape>", lambda e: self.destroy())

    def update_res(self):
        try:
            q = to_float(self.qty_e.get(), "Quantité")
            tot = self.article["nombre_existant"] + q
            self.res_lbl.configure(text=f"Nouveau stock après entrée : {tot:g}", text_color=GREEN)
        except Exception:
            self.res_lbl.configure(text="Quantité invalide", text_color=ORANGE)

    def validate(self):
        try:
            qty = to_float(self.qty_e.get(), "Quantité")
            if qty <= 0:
                raise ValueError("La quantité doit être strictement positive.")
            agent = self.agent_cb.get().strip() or "Admin"
            note = self.note_e.get().strip()
            self.app.db.add_stock(self.article["id"], qty, agent=agent, note=note)
        except Exception as ex:
            fail(ex, self)
            return

        self.app.refresh_all()
        self.app.msg(f"Stock augmenté de +{qty:g} pour « {self.article['designation']} » enregistré dans Excel.")
        self.destroy()


# ----------------------------------------------------------------- Application Principale
class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=BG)
        ctk.set_appearance_mode("dark")

        # Configuration de style moderne pour Treeview ttk
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background=FIELD,
            fieldbackground=FIELD,
            foreground=TEXT_WHITE,
            rowheight=32,
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

        self.title(f"{APP_NAME} — Gestion de Stock & Logistique")
        self.geometry("1220x740")
        self.minsize(1000, 620)

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
        start_x = max(0, (screen_w - 1220) // 2)
        start_y = max(0, (screen_h - 740) // 2)
        self.geometry(f"1220x740+{start_x}+{start_y}")

        self.sel_art = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.build_sidebar()

        main_area = ctk.CTkFrame(self, fg_color="transparent")
        main_area.grid(row=0, column=1, sticky="nsew")
        main_area.grid_columnconfigure(0, weight=1)
        main_area.grid_rowconfigure(1, weight=1)

        # En-tête supérieur
        top_bar = ctk.CTkFrame(main_area, fg_color="transparent", height=54)
        top_bar.pack(fill="x", padx=24, pady=(14, 0))

        self.page_title = label(top_bar, "", TEXT_WHITE, 22, True)
        self.page_title.pack(side="left")

        # Indicateur DB Excel connecté & Actualiser
        db_badge = ctk.CTkFrame(top_bar, fg_color=CARD, corner_radius=8, border_width=1, border_color=BORDER)
        db_badge.pack(side="right", padx=(12, 0))
        label(db_badge, "🟢 Base Excel :", GREEN, 11, True).pack(side="left", padx=(10, 4), pady=6)
        label(db_badge, os.path.basename(DB_PATH), TEXT_WHITE, 11).pack(side="left", padx=(0, 6), pady=6)
        button(db_badge, "🔄 Actualiser", PANEL, self.sync_reload, width=80, height=28).pack(side="left", padx=4, pady=3)

        self.clock = label(top_bar, "", TEXT_MUTED, 12)
        self.clock.pack(side="right")

        # Corps dynamique des pages
        self.body = ctk.CTkFrame(main_area, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=24, pady=(8, 4))

        # Barre d'état inférieure
        self.status = label(main_area, "", TEXT_MUTED, 11, anchor="w")
        self.status.pack(side="bottom", fill="x", padx=24, pady=(2, 6))

        # Pages
        self.pages = {
            "stock": self.build_stock(),
            "sortie": self.build_sortie(),
            "histo": self.build_histo()
        }

        self.go("stock")
        self.tick()

    # ------------------------------------------------------------- Navigation & Structure
    def build_sidebar(self):
        side = ctk.CTkFrame(self, fg_color=SIDE, corner_radius=0, width=240)
        side.grid(row=0, column=0, sticky="ns")
        side.grid_propagate(False)

        brand_box = ctk.CTkFrame(side, fg_color="transparent")
        brand_box.pack(fill="x", padx=20, pady=(24, 20))
        label(brand_box, "⚡ THINK POWER", BLUE, 18, True).pack(anchor="w")
        label(brand_box, "Stock & Logistique", TEXT_MUTED, 11).pack(anchor="w")

        self.nav = {}
        items = (
            ("stock", "📦   Stock & Articles"),
            ("sortie", "📤   Sortie de stock"),
            ("histo", "📋   Historique & Journal")
        )

        for key, text in items:
            b = ctk.CTkButton(
                side,
                text=text,
                anchor="w",
                height=44,
                corner_radius=8,
                fg_color="transparent",
                hover_color="#1f293d",
                font=("Segoe UI", 13, "bold"),
                text_color=TEXT_WHITE,
                command=lambda k=key: self.go(k)
            )
            b.pack(fill="x", padx=12, pady=3)
            self.nav[key] = b

        # Séparateur / Actions rapides
        sep = ctk.CTkFrame(side, height=1, fg_color=BORDER)
        sep.pack(side="bottom", fill="x", padx=14, pady=(0, 12))

        button(side, "📂  Ouvrir Base Excel", GREEN, self.open_db, height=38).pack(side="bottom", fill="x", padx=14, pady=(0, 8))
        button(side, "🗂  Dossier Application", PANEL, lambda: self.open_path(BASE_DIR), height=38).pack(side="bottom", fill="x", padx=14, pady=(0, 8))

    TITLES = {
        "stock": "📦  Gestion du Stock & Articles",
        "sortie": "📤  Enregistrement des Sorties",
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
        self.status.configure(text=f"●  {text}   |   Base Excel active : {DB_PATH}")

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
                    f"Le fichier a bien été créé mais aucun logiciel n'est configuré pour l'ouvrir directement sur ce poste.\n\n"
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
            self.refresh_all()
            self.msg("Données rechargées avec succès depuis le fichier Excel.")
        except Exception as ex:
            fail(f"Erreur de rechargement :\n{ex}")

    def refresh_all(self):
        self.refresh_stock()
        self.refresh_pick()
        if hasattr(self, "s_agent"):
            self.s_agent.configure(values=self.db.agents() or ["Admin"])
        self.refresh_histo()
        if not self.status.cget("text"):
            self.msg("Prêt — Toutes les données sont sauvegardées en temps réel dans le fichier Excel.")

    # ------------------------------------------------------------- Page 1 : Stock
    def build_stock(self):
        f = panel(self.body)

        # Cartes KPI
        cards_box = ctk.CTkFrame(f, fg_color="transparent")
        cards_box.pack(fill="x", padx=16, pady=(14, 6))

        self.kpi = {}
        for key, text, icon, color in (
            ("refs", "Total Références", "📦", BLUE),
            ("ttc", "Valeur Stock TTC", "💰", CYAN),
            ("alert", "Alertes (Rupture / Faible)", "⚠️", ORANGE),
            ("sorties", "Sorties du mois", "📉", GREEN)
        ):
            c, v = card(cards_box, text, icon, color)
            c.pack(side="left", expand=True, fill="x", padx=4)
            self.kpi[key] = v

        # Barre de recherche & filtres
        bar = ctk.CTkFrame(f, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=(6, 6))

        self.search = entry(bar, 300, placeholder_text="🔎  Rechercher désignation ou note...")
        self.search.pack(side="left")
        self.search.bind("<KeyRelease>", lambda e: self.refresh_stock())

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
        self.flt.pack(side="left", padx=12)

        # Mode option 10%
        self.mode_menu = ctk.CTkOptionMenu(
            bar,
            values=["remise", "frais"],
            width=95,
            height=36,
            fg_color=PANEL,
            button_color=BLUE,
            button_hover_color=BLUE_HOVER,
            command=self.set_mode
        )
        self.mode_menu.set(self.db.mode10)
        self.mode_menu.pack(side="right")
        label(bar, "Option 10% :", TEXT_MUTED, 12).pack(side="right", padx=6)

        # Barre d'actions
        acts = ctk.CTkFrame(f, fg_color="transparent")
        acts.pack(fill="x", padx=16, pady=(4, 8))

        button(acts, "➕  Nouvel article", BLUE, lambda: ArticleDialog(self), width=140).pack(side="left", padx=(0, 6))
        button(acts, "➕  Ajouter Stock", GREEN, self.open_quick_add, width=135).pack(side="left", padx=(0, 6))
        button(acts, "✏️  Modifier", PANEL, self.edit_article, width=110).pack(side="left", padx=(0, 6))
        button(acts, "📤  Faire Sortie", ORANGE, self.sortie_from_stock, width=125).pack(side="left", padx=(0, 6))
        button(acts, "🧮  Inventaire", CYAN, self.open_inventory, width=120).pack(side="left", padx=(0, 6))
        button(acts, "📥  Importer Excel", PANEL, self.import_excel, width=135).pack(side="left", padx=(0, 6))
        button(acts, "🗑  Supprimer", RED, self.del_article, width=115).pack(side="right")

        # Tableau des articles
        cols = ("ID", "Désignation", "État", "En Stock", "Seuil", "PU HT", "PU TTC", "Valeur TTC", "Date Saisie", "Note")
        box, self.tree = make_tree(f, cols, (45, 260, 95, 75, 65, 85, 85, 95, 120, 200))
        box.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        self.tree.bind("<Double-1>", lambda e: self.edit_article())
        self.tree.bind("<Button-3>", self.show_stock_context_menu)

        # Menu contextuel clic-droit
        self.stock_menu = Menu(self, tearoff=0, background=PANEL, foreground=TEXT_WHITE, activebackground=BLUE)
        self.stock_menu.add_command(label="✏️  Modifier l'article", command=self.edit_article)
        self.stock_menu.add_command(label="➕  Ajouter du Stock (Entrée)", command=self.open_quick_add)
        self.stock_menu.add_command(label="📤  Faire une sortie", command=self.sortie_from_stock)
        self.stock_menu.add_separator()
        self.stock_menu.add_command(label="🗑  Supprimer définitivement de l'Excel", command=self.del_article)

        return f

    def show_stock_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.stock_menu.post(event.x_root, event.y_root)

    def set_mode(self, v):
        try:
            self.db.set_mode10(v)
        except RuntimeError as ex:
            fail(ex)
        self.refresh_stock()

    def refresh_stock(self):
        self.tree.delete(*self.tree.get_children())
        mode = self.db.mode10
        arts = self.db.articles(self.search.get(), FILTERS[self.flt.get()])

        for i, a in enumerate(arts):
            c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], mode)
            e = etat(a)
            pu_ttc = a["pu_ht"] * (1 + a["taux_tva"] / 100)
            self.tree.insert(
                "", "end",
                iid=str(a["id"]),
                tags=(ETAT_TAG[e], "odd" if i % 2 else "even"),
                values=(
                    a["id"],
                    a["designation"],
                    ETAT_LABEL[e],
                    f"{a['nombre_existant']:g}",
                    f"{a['seuil_alerte']:g}",
                    f"{a['pu_ht']:.3f}",
                    f"{pu_ttc:.3f}",
                    f"{c['ttc']:.3f}",
                    fmt_date(a["date_saisie"]),
                    a.get("note", "")
                )
            )

        k = self.db.kpis()
        self.kpi["refs"].configure(text=str(k["refs"]))
        self.kpi["ttc"].configure(text=f"{k['ttc']:.3f} DT")
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

    def open_quick_add(self):
        a = self.selected_article()
        if a:
            QuickAddStockDialog(self, a)

    def sortie_from_stock(self):
        a = self.selected_article()
        if a:
            self.sel_art = a["id"]
            self.go("sortie")
            self.set_article(a["id"])

    def del_article(self):
        """Supprime directement et immédiatement la ligne du fichier Excel."""
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
        self.msg(f"Article « {a['designation']} » supprimé directement du fichier Excel.")

    def import_excel(self):
        path = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel",
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
        self.msg(f"{n} article(s) importé(s) depuis {os.path.basename(path)}")
        messagebox.showinfo(APP_NAME, f"{n} article(s) importé(s) / mis à jour avec succès dans Excel.")

    def open_inventory(self):
        d = ctk.CTkToplevel(self, fg_color=BG)
        d.title("Inventaire physique régularisation")
        d.geometry("480x480")
        d.resizable(False, False)
        d.transient(self)
        d.grab_set()

        # Centrage
        d.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 480) // 2
        y = self.winfo_y() + (self.winfo_height() - 480) // 2
        d.geometry(f"+{max(0, x)}+{max(0, y)}")

        label(d, "🧮  Inventaire physique", CYAN, 18, True).pack(anchor="w", padx=24, pady=(18, 8))
        box = panel(d)
        box.pack(fill="both", expand=True, padx=22, pady=(0, 16))

        a0 = self.db.article(int(self.tree.selection()[0])) if self.tree.selection() else None

        label(box, "Article / Matière à vérifier", TEXT_MUTED, 11).pack(anchor="w", padx=16, pady=(14, 2))
        art_cb = combo(box, 400)
        art_cb.configure(values=[a["designation"] for a in self.db.articles()] or [""])
        art_cb.set(a0["designation"] if a0 else "")
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
                a = self.db.article_by_name(art_cb.get())
                if not a:
                    raise ValueError("Veuillez sélectionner un article valide.")
                phys = to_float(qty_e.get(), "Quantité physique")
                agent = agent_cb.get().strip() or "Admin"
                ecart = phys - a["nombre_existant"]
                c = calc(abs(ecart), a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], self.db.mode10)

                conf = messagebox.askyesno(
                    APP_NAME,
                    f"Article : {a['designation']}\n"
                    f"Stock théorique : {a['nombre_existant']:g}\n"
                    f"Stock physique : {phys:g}\n"
                    f"Écart : {ecart:+g}  (Valeur : {c['ttc']:.3f} DT TTC)\n\n"
                    "Enregistrer la régularisation directement dans Excel ?",
                    parent=d
                )
                if not conf:
                    return

                self.db.add_movement(a["id"], "INVENTAIRE", phys, agent, "", "Régularisation inventaire")
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
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        # ---- Formulaire Sortie (Gauche)
        label(left, "📤  Nouvelle Sortie de Stock", ORANGE, 17, True).pack(anchor="w", padx=20, pady=(16, 6))

        # Carte article sélectionné
        self.art_box = ctk.CTkFrame(left, fg_color=CARD, corner_radius=10, border_width=1, border_color=BORDER)
        self.art_box.pack(fill="x", padx=20, pady=(4, 10))
        label(self.art_box, "Matière sélectionnée :", TEXT_MUTED, 11).pack(anchor="w", padx=14, pady=(8, 2))
        self.s_art = label(self.art_box, "— Aucune matière sélectionnée —", CYAN, 14, True, wraplength=340, justify="left")
        self.s_art.pack(anchor="w", padx=14)
        self.s_exist = label(self.art_box, "Veuillez choisir un article dans la liste à droite →", TEXT_MUTED, 11)
        self.s_exist.pack(anchor="w", padx=14, pady=(2, 8))

        # Quantité
        label(left, "Quantité demandée *", TEXT_MUTED, 11).pack(anchor="w", padx=20, pady=(4, 2))
        q_box = ctk.CTkFrame(left, fg_color="transparent")
        q_box.pack(anchor="w", padx=20)
        button(q_box, "−", BLUE, lambda: self.step(-1), width=44, height=36).pack(side="left")
        self.s_qty = entry(q_box, 110, justify="center")
        self.s_qty.pack(side="left", padx=6)
        self.s_qty.bind("<KeyRelease>", lambda e: self.check_qty())
        button(q_box, "+", BLUE, lambda: self.step(1), width=44, height=36).pack(side="left")

        self.s_live = label(left, "", TEXT_MUTED, 11, True)
        self.s_live.pack(anchor="w", padx=20, pady=(2, 6))

        # Bénéficiaire & Agent
        label(left, "Société / Client bénéficiaire *", TEXT_MUTED, 11).pack(anchor="w", padx=20, pady=(4, 2))
        self.s_soc = entry(left, 360, placeholder_text="Ex: Entreprise ABC, Projet Y...")
        self.s_soc.pack(fill="x", padx=20)

        label(left, "Agent émetteur *", TEXT_MUTED, 11).pack(anchor="w", padx=20, pady=(8, 2))
        self.s_agent = combo(left, 360)
        self.s_agent.configure(values=self.db.agents() or ["Admin"])
        self.s_agent.set(self.db.last_agent or "Admin")
        self.s_agent.pack(fill="x", padx=20)

        label(left, "Date & Heure", TEXT_MUTED, 11).pack(anchor="w", padx=20, pady=(8, 2))
        self.s_date = entry(left, 360)
        self.s_date.insert(0, datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.s_date.pack(fill="x", padx=20)

        label(left, "Note / Justification", TEXT_MUTED, 11).pack(anchor="w", padx=20, pady=(8, 2))
        self.s_note = entry(left, 360, placeholder_text="Optionnel...")
        self.s_note.pack(fill="x", padx=20)

        button(left, "✔  VALIDER LA SORTIE (EXCEL)", ORANGE, self.do_sortie, height=44).pack(fill="x", padx=20, pady=(16, 16))

        # ---- Sélecteur Articles & Aperçu (Droite)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(2, weight=3)
        right.grid_rowconfigure(4, weight=2)

        label(right, "1 · Choisir la matière dans le stock", TEXT_WHITE, 15, True).grid(row=0, column=0, sticky="w", padx=18, pady=(16, 4))
        self.p_search = entry(right, 320, placeholder_text="🔎  Filtrer les matières...")
        self.p_search.grid(row=1, column=0, sticky="w", padx=18, pady=(0, 6))
        self.p_search.bind("<KeyRelease>", lambda e: self.refresh_pick())

        box, self.pick = make_tree(right, ("Désignation", "En Stock", "État"), (280, 80, 100))
        box.grid(row=2, column=0, sticky="nsew", padx=18)
        self.pick.bind("<<TreeviewSelect>>", self.on_pick)

        label(right, "2 · Détails & Récapitulatif", TEXT_WHITE, 14, True).grid(row=3, column=0, sticky="w", padx=18, pady=(10, 4))
        self.console = ctk.CTkTextbox(right, fg_color=CONSOLE, font=("Consolas", 11), corner_radius=10, state="disabled")
        self.console.grid(row=4, column=0, sticky="nsew", padx=18, pady=(0, 16))
        self.say("Sélectionnez une matière dans le tableau ci-dessus pour afficher son détail.")

        return f

    def say(self, text):
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.insert("1.0", text)
        self.console.configure(state="disabled")

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
                values=(a["designation"], f"{a['nombre_existant']:g}", ETAT_LABEL[e])
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
        self.s_art.configure(text=a["designation"])
        self.s_exist.configure(text=f"Stock actuel : {a['nombre_existant']:g}   (Seuil d'alerte : {a['seuil_alerte']:g})")
        if not self.s_qty.get().strip():
            self.s_qty.insert(0, "1")

        c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], self.db.mode10)
        pu_ttc = a["pu_ht"] * (1 + a["taux_tva"] / 100)

        txt = (
            f"Article         : {a['designation']}\n"
            f"Stock disponible: {a['nombre_existant']:g}\n"
            f"Prix unitaire HT: {a['pu_ht']:.3f} DT\n"
            f"Prix unitaire TTC: {pu_ttc:.3f} DT\n"
            f"Valeur en stock : {c['ttc']:.3f} DT TTC\n"
        )
        if etat(a) != "ok":
            txt += "\n⚠️  ATTENTION : STOCK FAIBLE OU EN RUPTURE !\n"

        txt += "\nDerniers mouvements enregistrés :\n"
        for m in self.db.article_movements(aid, limit=4):
            q = -m["qte"] if m["type"] == "SORTIE" else m["qte"]
            txt += f"  • {fmt_date(m['date'])[:10]}  {m['type']:<11} {q:+g}  ({m.get('agent', '-')})\n"

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

        if q > a["nombre_existant"]:
            self.s_live.configure(text=f"⛔ Stock insuffisant ({a['nombre_existant']:g} disponible)", text_color=RED)
        else:
            reste = a["nombre_existant"] - q
            self.s_live.configure(text=f"✔ Reste en stock après sortie : {reste:g}", text_color=GREEN)

    def do_sortie(self):
        try:
            agent = self.s_agent.get().strip() or "Admin"
            soc = self.s_soc.get().strip()
            if not self.sel_art:
                raise ValueError("Veuillez d'abord choisir une matière dans la liste à droite.")
            if not soc:
                raise ValueError("La société ou client bénéficiaire est obligatoire.")
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
        txt = (
            f"✅  SORTIE ENREGISTRÉE DANS EXCEL — N° Bon : {m['bon']}\n\n"
            f"Article  : {a['designation']}\n"
            f"Bénéf.   : {soc}\n"
            f"Agent    : {agent}\n"
            f"Quantité : {m['qte']:g}\n"
            f"Reste    : {m['apres']:g}\n"
            f"Total TTC: {m['total_ttc']:.3f} DT\n"
        )
        if etat(a) != "ok":
            txt += "\n⚠️  STOCK FAIBLE / RUPTURE : Penser à commander !\n"

        self.s_qty.delete(0, "end")
        self.s_note.delete(0, "end")
        self.s_date.delete(0, "end")
        self.s_date.insert(0, datetime.now().strftime("%d/%m/%Y %H:%M"))

        self.refresh_all()
        self.s_exist.configure(text=f"Stock actuel : {a['nombre_existant']:g}   (Seuil : {a['seuil_alerte']:g})")
        self.s_live.configure(text="")
        self.say(txt)
        self.msg(f"Sortie {m['bon']} enregistrée par {agent} dans Excel.")

        # Proposition de Bon PDF
        if messagebox.askyesno(APP_NAME, f"Sortie {m['bon']} validée !\n\nVoulez-vous générer et ouvrir le Bon de Sortie en PDF ?"):
            try:
                os.makedirs(EXPORT_DIR, exist_ok=True)
                pdf_path = os.path.join(EXPORT_DIR, f"Bon_{m['bon']}.pdf")
                export_bon_pdf(self.db, m, pdf_path)
                self.open_path(pdf_path)
            except Exception as ex:
                fail(f"Erreur PDF :\n{ex}")

    # ------------------------------------------------------------- Page 3 : Historique
    def build_histo(self):
        f = panel(self.body)

        # Filtres
        bar = ctk.CTkFrame(f, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=(14, 6))

        self.h_agent = entry(bar, 140, placeholder_text="Filtrer Agent...")
        self.h_soc = entry(bar, 160, placeholder_text="Filtrer Bénéficiaire...")
        self.h_d1 = entry(bar, 125, placeholder_text="Du JJ/MM/AAAA")
        self.h_d2 = entry(bar, 125, placeholder_text="Au JJ/MM/AAAA")
        self.h_typ = ctk.CTkOptionMenu(
            bar,
            values=["Tous", "SORTIE", "ENTREE", "INVENTAIRE", "NOUVEAU", "MODIF", "SUPPRESSION", "IMPORT"],
            width=135,
            height=38,
            fg_color=PANEL,
            button_color=BLUE,
            button_hover_color=BLUE_HOVER
        )

        for w in (self.h_agent, self.h_soc, self.h_d1, self.h_d2, self.h_typ):
            w.pack(side="left", padx=(0, 6))

        button(bar, "🔎  Filtrer", BLUE, self.refresh_histo, width=95).pack(side="left", padx=4)
        button(bar, "↺  Réinitialiser", PANEL, self.reset_histo_filters, width=105).pack(side="left", padx=2)

        # Tableau historique
        cols = ("N° Bon", "Date", "Type", "Article", "Agent", "Société / Bénéficiaire", "Avant", "Mvt", "Reste", "Total TTC", "Note")
        box, self.htree = make_tree(f, cols, (95, 120, 95, 200, 105, 150, 55, 55, 55, 85, 180))
        box.pack(fill="both", expand=True, padx=16, pady=6)

        # Actions d'export & suppression directe
        ex_bar = ctk.CTkFrame(f, fg_color="transparent")
        ex_bar.pack(fill="x", padx=16, pady=(4, 14))

        button(ex_bar, "📗  Exporter Rapport Excel (.xlsx)", GREEN, lambda: self.do_export("xlsx"), width=230, height=40).pack(side="left", padx=(0, 10))
        button(ex_bar, "📕  Exporter Rapport PDF", RED, lambda: self.do_export("pdf"), width=210, height=40).pack(side="left")

        # Bouton supprimer ligne sélectionnée
        button(ex_bar, "🗑  Supprimer la ligne de l'historique", PANEL, self.del_movement, width=240, height=40).pack(side="right")

        return f

    def reset_histo_filters(self):
        self.h_agent.delete(0, "end")
        self.h_soc.delete(0, "end")
        self.h_d1.delete(0, "end")
        self.h_d2.delete(0, "end")
        self.h_typ.set("Tous")
        self.refresh_histo()

    def histo_filters(self):
        t = self.h_typ.get()
        return dict(
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
        """Supprime directement une ligne du journal dans Excel."""
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
