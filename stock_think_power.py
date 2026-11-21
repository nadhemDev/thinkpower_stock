# -*- coding: utf-8 -*-
"""Stock Think Power - moteur de données intelligent & universel.
Compatible avec tout type de stock (matériel, quincaillerie, pièces détachées, commerce, alimentaire, etc.).
La base de données est un fichier Excel (Stock_Think_Power_Base.xlsx) placé à côté du .exe.
Toutes les opérations (entrées, sorties, inventaires, régularisations, achats) sont enregistrées en temps réel.
"""
import os
import re
import sys
import copy
import shutil
from contextlib import contextmanager
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from fpdf import FPDF

APP_NAME = "Stock Think Power"


def get_base_dir():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    # Vérification des droits d'écriture dans le dossier de base
    try:
        test_file = os.path.join(base, f".wr_test_{os.getpid()}.tmp")
        with open(test_file, "w") as f:
            f.write("1")
        os.remove(test_file)
        return base
    except Exception:
        fallback = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "StockThinkPower")
        os.makedirs(fallback, exist_ok=True)
        return fallback


def get_resource_path(relative_path):
    """Retourne le chemin absolu vers une ressource empaquetée avec PyInstaller (sys._MEIPASS) ou locale."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


BASE_DIR = get_base_dir()
DB_PATH = os.path.join(BASE_DIR, "Stock_Think_Power_Base.xlsx")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
EXPORT_DIR = os.path.join(BASE_DIR, "exports")
DT = "#,##0.000"
DISP = "%d/%m/%Y %H:%M"
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F538D")

STOCK_HEAD = [
    "ID", "Référence", "Matière / Désignation", "Catégorie", "Unité", "Emplacement",
    "Nombre Existant", "Stock Minimum", "PU HT", "TVA %",
    "Option 10%", "PU TTC", "Total HTVA", "Total TTC", "Date de saisie", "État", "Note"
]
JOURNAL_HEAD = [
    "ID", "Bon", "Date", "Type", "Agent", "Société / Client / Fournisseur", "Article", "Réf",
    "ID Article", "Avant", "Qté", "Reste", "PU HT", "HTVA net", "TVA", "TTC", "Note"
]

DEFAULT_CATEGORIES = [
    "Général", "Matières Premières", "Électricité & Câblage", "Outillage & Équipement",
    "Quincaillerie & Visserie", "Pièces Détachées", "Consommables", "Sécurité & EPI", "Informatique"
]
DEFAULT_UNITS = ["pcs", "kg", "m", "L", "carton", "paquet", "sac", "boîte", "rouleau", "lot"]


# ----------------------------------------------------------------- Utilitaires
def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def to_iso(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    s = str(v or "").strip()
    for f in (DISP, "%d/%m/%Y", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d %H:%M")
        except ValueError:
            pass
    return now_iso()


def fmt_date(iso):
    try:
        return datetime.strptime(str(iso)[:16], "%Y-%m-%d %H:%M").strftime(DISP)
    except Exception:
        return str(iso or "")


def user_date(text, end=False):
    """JJ/MM/AAAA [HH:MM] -> ISO. Vide -> None."""
    text = (text or "").strip()
    if not text:
        return None
    for f in ("%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            d = datetime.strptime(text, f)
            if f == "%d/%m/%Y" and end:
                d = d.replace(hour=23, minute=59)
            return d.strftime("%Y-%m-%d %H:%M")
        except ValueError:
            pass
    raise ValueError("Date invalide (format attendu : JJ/MM/AAAA ou JJ/MM/AAAA HH:MM).")


def parse_num(v, default=0.0):
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().upper().replace(",", ".")
    if not s or s.startswith("#"):
        return default
    m = re.search(r"\d+(?:\.\d+)?", s)
    if not m:
        return 0.5 if "DEMI" in s else default
    n = float(m.group())
    return n + 0.5 if "DEMI" in s else n


def to_float(txt, label=""):
    try:
        s = str(txt or "").strip().replace(" ", "").replace(",", ".").rstrip("%").strip()
        for cur in ("DT", "dt", "Dt", "TND", "tnd", "€", "$", "MAD", "DZD", "CFA"):
            if s.endswith(cur):
                s = s[:-len(cur)].strip()
        return float(s)
    except ValueError:
        lbl = label if label else "Valeur"
        raise ValueError(f"{lbl} : nombre invalide (« {txt} »).")


def parse_date_cell(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M"), ""
    s = str(v or "").strip()
    if not s or s.startswith("#"):
        return now_iso(), ""
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        try:
            d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            return d.strftime("%Y-%m-%d 00:00"), (s[:m.start()] + s[m.end():]).strip(" ()-")
        except ValueError:
            pass
    return now_iso(), s


def calc(qty, pu, tva, opt10, mode):
    qty = float(qty)
    pu = float(pu)
    tva = float(tva)
    htva = qty * pu
    adj = htva * 0.10 if opt10 else 0.0
    if mode == "remise":
        adj = -adj
    net = max(0.0, htva + adj)
    tva_m = net * tva / 100.0
    return {"htva": htva, "adj": adj, "net": net, "tva": tva_m, "ttc": net + tva_m}


def etat(a):
    if a["nombre_existant"] <= 0:
        return "rupture"
    return "faible" if a["nombre_existant"] <= a["seuil_alerte"] else "ok"


def sheet_head(ws, headers, widths):
    ws.append(headers)
    for c in ws[1]:
        c.font, c.fill, c.alignment = HEAD_FONT, HEAD_FILL, Alignment(horizontal="center", vertical="center")
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


# ------------------------------------------------------------ Base de données Excel
class DB:
    def __init__(self, path=None):
        self.path = path or os.path.join(get_base_dir(), "Stock_Think_Power_Base.xlsx")
        self.arts = []
        self.movs = []
        self.settings = {}
        self.reload()

    def reload(self):
        """Recharge toutes les données depuis le fichier Excel physique."""
        self.arts, self.movs, self.settings = [], [], {}
        if os.path.exists(self.path):
            self.backup()
            self.load()
        else:
            self.save()

    def backup(self):
        if not os.path.exists(self.path):
            return
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            dst = os.path.join(BACKUP_DIR, f"Base_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            shutil.copy2(self.path, dst)
            files = sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith("Base_") and f.endswith(".xlsx")])
            for old in files[:-20]:
                try:
                    os.remove(os.path.join(BACKUP_DIR, old))
                except Exception:
                    pass
        except Exception:
            pass

    def load(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)

        def rows(name):
            if name not in wb.sheetnames:
                return []
            data = list(wb[name].iter_rows(values_only=True))
            if not data:
                return []
            head = [str(h or "").strip() for h in data[0]]
            return [dict(zip(head, r)) for r in data[1:] if any(c not in (None, "") for c in r)]

        for r in rows("Stock"):
            name = str(r.get("Matière / Désignation") or r.get("Désignation") or "").strip()
            if not name or name.upper().startswith("TOTAL"):
                continue
            aid = int(parse_num(r.get("ID")))
            ref = str(r.get("Référence") or r.get("Réf") or r.get("Code Barre") or "").strip()
            if not ref:
                ref = f"REF-{aid:04d}"
            cat = str(r.get("Catégorie") or r.get("Famille") or "Général").strip()
            unite = str(r.get("Unité") or "pcs").strip()
            emp = str(r.get("Emplacement") or r.get("Rayon") or "").strip()

            self.arts.append(dict(
                id=aid,
                ref=ref,
                designation=name,
                categorie=cat,
                unite=unite,
                emplacement=emp,
                nombre_existant=parse_num(r.get("Nombre Existant") or r.get("Stock")),
                seuil_alerte=parse_num(r.get("Stock Minimum") or r.get("Seuil"), 5),
                pu_ht=parse_num(r.get("PU HT") or r.get("Prix HT")),
                taux_tva=parse_num(r.get("TVA %") or r.get("TVA"), 19),
                appliquer_10_pct=str(r.get("Option 10%") or "").strip().lower() in ("oui", "1", "true", "x"),
                date_saisie=to_iso(r.get("Date de saisie")),
                note=str(r.get("Note") or "").strip()
            ))

        used = set()
        for a in self.arts:
            if a["id"] <= 0 or a["id"] in used:
                a["id"] = max([x["id"] for x in self.arts] + [0]) + 1
            used.add(a["id"])

        for r in rows("Journal"):
            self.movs.append(dict(
                id=int(parse_num(r.get("ID"))),
                bon=str(r.get("Bon") or ""),
                date=to_iso(r.get("Date")),
                type=str(r.get("Type") or ""),
                agent=str(r.get("Agent") or ""),
                societe=str(r.get("Société / Client / Fournisseur") or r.get("Société / Client") or ""),
                designation=str(r.get("Article") or ""),
                ref=str(r.get("Réf") or ""),
                article_id=int(parse_num(r.get("ID Article"))),
                avant=parse_num(r.get("Avant")),
                qte=parse_num(r.get("Qté")),
                apres=parse_num(r.get("Reste")),
                pu_ht=parse_num(r.get("PU HT")),
                total_htva=parse_num(r.get("HTVA net")),
                tva=parse_num(r.get("TVA")),
                total_ttc=parse_num(r.get("TTC")),
                note=str(r.get("Note") or "")
            ))

        if "Paramètres" in wb.sheetnames:
            for k, v in list(wb["Paramètres"].iter_rows(values_only=True))[1:]:
                if k:
                    self.settings[str(k)] = str(v if v is not None else "")

    def save(self):
        """Enregistre l'ensemble des données dans Stock_Think_Power_Base.xlsx."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Stock"
        # 17 colonnes
        sheet_head(ws, STOCK_HEAD, (6, 14, 36, 18, 8, 14, 15, 12, 12, 8, 11, 12, 14, 14, 17, 11, 28))
        mode = self.mode10
        colors = {"ok": "1E8449", "faible": "D68910", "rupture": "C0392B"}

        for i, a in enumerate(sorted(self.arts, key=lambda x: str(x["designation"]).lower()), start=2):
            c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], mode)
            e = etat(a)
            ws.append([
                a["id"],
                a.get("ref", f"REF-{a['id']:04d}"),
                a["designation"],
                a.get("categorie", "Général"),
                a.get("unite", "pcs"),
                a.get("emplacement", ""),
                a["nombre_existant"],
                a["seuil_alerte"],
                a["pu_ht"],
                a["taux_tva"],
                "Oui" if a["appliquer_10_pct"] else "Non",
                a["pu_ht"] * (1 + a["taux_tva"] / 100),
                c["net"],
                c["ttc"],
                fmt_date(a["date_saisie"]),
                {"ok": "OK", "faible": "Faible", "rupture": "Rupture"}[e],
                a["note"]
            ])
            for col in ["I", "L", "M", "N"]:
                ws[col + str(i)].number_format = DT
            ws["P" + str(i)].font = Font(bold=True, color=colors[e])

        n = len(self.arts) + 1
        if self.arts:
            ws.append(["", "", "TOTAL", "", "", "", "", "", "", "", "", "", f"=SUM(M2:M{n})", f"=SUM(N2:N{n})"])
            for col in ["M", "N"]:
                ws[col + str(n + 1)].number_format = DT
            for c in ws[n + 1]:
                c.font = Font(bold=True)
            ws.auto_filter.ref = f"A1:Q{n}"

        j = wb.create_sheet("Journal")
        sheet_head(j, JOURNAL_HEAD, (6, 13, 17, 13, 16, 26, 34, 14, 10, 8, 8, 8, 11, 12, 11, 12, 30))
        for i, m in enumerate(sorted(self.movs, key=lambda x: (x["date"], x["id"]), reverse=True), start=2):
            j.append([
                m["id"],
                m["bon"],
                fmt_date(m["date"]),
                m["type"],
                m["agent"],
                m["societe"],
                m["designation"],
                m.get("ref", ""),
                m["article_id"],
                m["avant"],
                m["qte"],
                m["apres"],
                m["pu_ht"],
                m["total_htva"],
                m["tva"],
                m["total_ttc"],
                m["note"]
            ])
            for col in ["M", "N", "O", "P"]:
                j[col + str(i)].number_format = DT
        if self.movs:
            j.auto_filter.ref = f"A1:Q{len(self.movs) + 1}"

        p = wb.create_sheet("Paramètres")
        sheet_head(p, ["Clé", "Valeur"], (24, 30))
        for k, v in self.settings.items():
            p.append([k, v])

        tmp = self.path + ".tmp"
        try:
            wb.save(tmp)
            if os.path.exists(self.path):
                try:
                    os.replace(tmp, self.path)
                except Exception:
                    os.remove(self.path)
                    os.rename(tmp, self.path)
            else:
                os.rename(tmp, self.path)
        except PermissionError:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except Exception:
                    pass
            raise RuntimeError(
                f"Le fichier Excel « {os.path.basename(self.path)} » est actuellement ouvert dans Excel.\n"
                "Veuillez le fermer puis réessayer."
            )

    @contextmanager
    def tx(self):
        snap = copy.deepcopy((self.arts, self.movs, self.settings))
        try:
            yield
            self.save()
        except Exception:
            self.arts, self.movs, self.settings = snap
            raise

    # -- Propriétés configurables
    @property
    def mode10(self):
        return self.settings.get("mode10", "remise")

    def set_mode10(self, v):
        with self.tx():
            self.settings["mode10"] = v

    @property
    def currency(self):
        return self.settings.get("devise", "DT")

    @property
    def company_name(self):
        return self.settings.get("societe_nom", "THINK POWER")

    def set_setting(self, k, v):
        with self.tx():
            self.settings[k] = str(v)

    @property
    def last_agent(self):
        return self.settings.get("dernier_agent", "")

    # -- Catégories & Unités
    def categories(self):
        found = {str(a.get("categorie") or "").strip() for a in self.arts if a.get("categorie")}
        all_cats = set(DEFAULT_CATEGORIES) | found
        return sorted([c for c in all_cats if c], key=str.lower)

    def units(self):
        found = {str(a.get("unite") or "").strip() for a in self.arts if a.get("unite")}
        all_u = set(DEFAULT_UNITS) | found
        return sorted([u for u in all_u if u], key=str.lower)

    # -- Lecture & Recherche intelligente
    def articles(self, search="", mode="all", categorie="Tous"):
        s = search.lower().strip()
        rows = self.arts

        if s:
            rows = [
                a for a in rows if (
                    s in a["designation"].lower()
                    or s in a.get("ref", "").lower()
                    or s in a.get("categorie", "").lower()
                    or s in a.get("emplacement", "").lower()
                    or s in a.get("note", "").lower()
                )
            ]

        if categorie and categorie != "Tous":
            rows = [a for a in rows if a.get("categorie", "Général").lower() == categorie.lower()]

        if mode == "alert":
            rows = [a for a in rows if a["nombre_existant"] <= a["seuil_alerte"]]
        elif mode == "rupture":
            rows = [a for a in rows if a["nombre_existant"] <= 0]
        elif mode == "ok":
            rows = [a for a in rows if a["nombre_existant"] > a["seuil_alerte"]]

        return sorted(rows, key=lambda a: str(a["designation"]).lower())

    def article(self, aid):
        return next((a for a in self.arts if a["id"] == aid), None)

    def article_by_ref_or_barcode(self, code):
        """Recherche exacte par référence SKU ou code barre scanné."""
        c = (code or "").strip().lower()
        if not c:
            return None
        return next((a for a in self.arts if a.get("ref", "").lower() == c), None)

    def article_by_name(self, name):
        n = (name or "").strip().lower()
        return next((a for a in self.arts if a["designation"].lower() == n), None)

    def agents(self):
        names = {m.get("agent", "").strip() for m in self.movs if m.get("agent", "").strip()}
        if self.last_agent:
            names.add(self.last_agent.strip())
        return sorted([n for n in names if n], key=str.lower)

    def movements(self, agent="", societe="", d1=None, d2=None, typ="", search=""):
        out = []
        s = (search or "").strip().lower()
        for m in self.movs:
            if agent and agent.lower() not in m.get("agent", "").lower():
                continue
            if societe and societe.lower() not in m.get("societe", "").lower():
                continue
            if d1 and m.get("date", "") < d1:
                continue
            if d2 and m.get("date", "") > d2:
                continue
            if typ and m.get("type", "") != typ:
                continue
            if s and not (
                s in m.get("designation", "").lower()
                or s in m.get("bon", "").lower()
                or s in m.get("ref", "").lower()
                or s in m.get("note", "").lower()
            ):
                continue
            out.append(m)
        return sorted(out, key=lambda m: (m.get("date", ""), m.get("id", 0)), reverse=True)

    def article_movements(self, aid, limit=6):
        return sorted(
            [m for m in self.movs if m.get("article_id") == aid],
            key=lambda m: (m.get("date", ""), m.get("id", 0)),
            reverse=True
        )[:limit]

    def kpis(self):
        mode, htva, ttc, total_units = self.mode10, 0.0, 0.0, 0.0
        for a in self.arts:
            c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], mode)
            htva += c["net"]
            ttc += c["ttc"]
            total_units += a["nombre_existant"]
        month = datetime.now().strftime("%Y-%m")
        return {
            "refs": len(self.arts),
            "total_units": total_units,
            "htva": htva,
            "ttc": ttc,
            "rupt": sum(1 for a in self.arts if etat(a) == "rupture"),
            "crit": sum(1 for a in self.arts if etat(a) == "faible"),
            "sorties": sum(m["qte"] for m in self.movs if m.get("type") == "SORTIE" and str(m.get("date", "")).startswith(month))
        }

    # -- Analyse Intelligente & Conseiller de Réapprovisionnement
    def smart_analysis(self):
        """Fournit les métriques avancées d'aide à la décision pour le gestionnaire de stock."""
        month = datetime.now().strftime("%Y-%m")
        # Sorties par article sur 30 jours
        sorties_map = {}
        for m in self.movs:
            if m.get("type") == "SORTIE":
                aid = m.get("article_id")
                sorties_map[aid] = sorties_map.get(aid, 0.0) + m.get("qte", 0.0)

        # Articles à commander (Réapprovisionnement suggéré)
        reorder = []
        for a in self.arts:
            if a["nombre_existant"] <= a["seuil_alerte"]:
                # Quantité suggérée pour retrouver un niveau de confort (seuil * 2)
                suggested = max(1.0, (a["seuil_alerte"] * 2.0) - a["nombre_existant"])
                cost = suggested * a["pu_ht"] * (1 + a["taux_tva"] / 100)
                reorder.append({
                    "id": a["id"],
                    "ref": a.get("ref", ""),
                    "designation": a["designation"],
                    "categorie": a.get("categorie", "Général"),
                    "unite": a.get("unite", "pcs"),
                    "emplacement": a.get("emplacement", ""),
                    "actuel": a["nombre_existant"],
                    "seuil": a["seuil_alerte"],
                    "suggested": suggested,
                    "pu_ht": a["pu_ht"],
                    "cost_ttc": cost,
                    "statut": "Rupture" if a["nombre_existant"] <= 0 else "Stock Faible"
                })

        # Top articles en rotation
        sorted_sorties = sorted(
            [{"article": self.article(aid), "qte": q} for aid, q in sorties_map.items() if self.article(aid)],
            key=lambda x: x["qte"],
            reverse=True
        )[:5]

        # Articles dormants (aucun mouvement et stock disponible)
        dead_stock = [
            a for a in self.arts
            if a["id"] not in sorties_map and a["nombre_existant"] > 0
        ][:8]

        return {
            "reorder_list": reorder,
            "reorder_total_cost": sum(x["cost_ttc"] for x in reorder),
            "top_sorties": sorted_sorties,
            "dead_stock": dead_stock,
            "total_articles": len(self.arts),
        }

    # -- Écriture
    def _next(self, rows):
        return max([r["id"] for r in rows] + [0]) + 1

    def next_bon(self, prefix="BS"):
        y = datetime.now().year
        n = sum(1 for m in self.movs if str(m.get("bon", "")).startswith(f"{prefix}-{y}-"))
        return f"{prefix}-{y}-{n + 1:03d}"

    def _log(self, a, typ, agent, note, avant, apres, qte=None, societe="", date=None, bon="", pu=0.0, cost=None):
        cost = cost or {"net": 0.0, "tva": 0.0, "ttc": 0.0}
        m = dict(
            id=self._next(self.movs),
            bon=bon,
            date=date or now_iso(),
            type=typ,
            agent=agent or self.last_agent or "Admin",
            societe=societe,
            designation=a["designation"] if a else "",
            ref=a.get("ref", "") if a else "",
            article_id=a["id"] if a else 0,
            avant=avant,
            qte=(apres - avant) if qte is None else qte,
            apres=apres,
            pu_ht=pu,
            total_htva=cost["net"],
            tva=cost["tva"],
            total_ttc=cost["ttc"],
            note=note
        )
        self.movs.append(m)
        if agent:
            self.settings["dernier_agent"] = agent
        return m

    def save_article(self, data, aid=None, agent=""):
        agent = agent.strip() or self.last_agent or "Admin"
        with self.tx():
            dup = self.article_by_name(data["designation"])
            if dup and dup["id"] != aid:
                raise ValueError(f"Un article avec la désignation « {data['designation']} » existe déjà.")

            ref_code = data.get("ref", "").strip()
            if ref_code:
                dup_ref = self.article_by_ref_or_barcode(ref_code)
                if dup_ref and dup_ref["id"] != aid:
                    raise ValueError(f"La référence / code-barres « {ref_code} » est déjà utilisée par « {dup_ref['designation']} ».")

            if aid:
                a = self.article(aid)
                old, typ, note = a["nombre_existant"], "MODIF", "Fiche article modifiée"
            else:
                a = {"id": self._next(self.arts)}
                self.arts.append(a)
                old, typ, note = 0, "NOUVEAU", "Article créé"

            ref_final = ref_code if ref_code else f"REF-{a['id']:04d}"
            a.update(
                ref=ref_final,
                designation=data["designation"].strip(),
                categorie=data.get("categorie", "Général").strip() or "Général",
                unite=data.get("unite", "pcs").strip() or "pcs",
                emplacement=data.get("emplacement", "").strip(),
                nombre_existant=float(data["qty"]),
                seuil_alerte=float(data["seuil"]),
                pu_ht=float(data["pu"]),
                taux_tva=float(data["tva"]),
                appliquer_10_pct=bool(data["opt10"]),
                note=data.get("note", "").strip(),
                date_saisie=now_iso()
            )
            self._log(a, typ, agent, note, old, a["nombre_existant"])
        return a

    def delete_article(self, aid, agent=""):
        self.backup()
        with self.tx():
            a = self.article(aid)
            if a and agent:
                self._log(a, "SUPPRESSION", agent, "Article supprimé", a["nombre_existant"], 0)
            self.arts = [x for x in self.arts if x["id"] != aid]

    def delete_movement(self, mid):
        self.backup()
        with self.tx():
            self.movs = [m for m in self.movs if m["id"] != mid]

    def add_stock(self, aid, qty, agent="", note="", date_iso=None, fournisseur="", bon_fournisseur=""):
        """Entrée de stock / réception fournisseur."""
        qty = float(qty)
        if qty <= 0:
            raise ValueError("La quantité ajoutée doit être supérieure à zéro.")
        a = self.article(aid)
        if not a:
            raise ValueError("Article introuvable.")
        agent = (agent or self.last_agent or "Admin").strip()
        avant = a["nombre_existant"]
        apres = avant + qty
        cost = calc(qty, a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], self.mode10)
        bon_ref = bon_fournisseur.strip() or self.next_bon(prefix="BL")
        with self.tx():
            a = self.article(aid)
            a["nombre_existant"] = apres
            a["date_saisie"] = now_iso()
            m = self._log(
                a, "ENTREE", agent,
                note or f"Réception fournisseur : {fournisseur or 'Non précisé'}",
                avant, apres, qty, fournisseur, date_iso, bon_ref, a["pu_ht"], cost
            )
        return m

    def add_movement(self, aid, typ, qty, agent, societe="", note="", date_iso=None):
        agent = (agent or "").strip()
        if not agent:
            raise ValueError("Le nom de l'agent est obligatoire.")
        a = self.article(aid)
        if not a:
            raise ValueError("Article introuvable.")
        avant, bon = a["nombre_existant"], ""
        if typ == "SORTIE":
            if qty <= 0:
                raise ValueError("La quantité demandée doit être > 0.")
            if qty > avant:
                raise ValueError(f"Stock insuffisant : {avant:g} disponible(s), {qty:g} demandé(s).")
            apres, bon = avant - qty, self.next_bon(prefix="BS")
        elif typ == "INVENTAIRE":
            if qty < 0:
                raise ValueError("La quantité physique ne peut pas être négative.")
            apres, qty = qty, qty - avant
        else:
            raise ValueError(f"Type de mouvement inconnu : {typ}")

        cost = calc(abs(qty), a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], self.mode10)
        with self.tx():
            a = self.article(aid)
            a["nombre_existant"], a["date_saisie"] = apres, now_iso()
            m = self._log(a, typ, agent, note, avant, apres, qty, societe, date_iso, bon, a["pu_ht"], cost)
        return m

    def import_excel(self, path, agent=""):
        agent = (agent or self.last_agent or "Admin").strip()
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["Stock"] if "Stock" in wb.sheetnames else wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        hdr = next((i for i, r in enumerate(rows[:30]) if any(
            c and ("mati" in str(c).lower() or "design" in str(c).lower() or "nom" in str(c).lower() or "art" in str(c).lower()) for c in r)), None)
        if hdr is None:
            raise ValueError("En-tête « Matière / Désignation / Article » introuvable dans la feuille sélectionnée.")
        cols = {}
        for j, c in enumerate(rows[hdr]):
            t = str(c or "").lower().replace(" ", "").replace(".", "")
            if any(k in t for k in ("mati", "design", "nom", "article")) and "name" not in cols:
                cols["name"] = j
            elif any(k in t for k in ("ref", "code", "sku", "bare")):
                cols.setdefault("ref", j)
            elif any(k in t for k in ("cat", "famille", "rayon")):
                cols.setdefault("cat", j)
            elif any(k in t for k in ("unit", "mesure")):
                cols.setdefault("unite", j)
            elif any(k in t for k in ("emp", "casier", "place", "loc")):
                cols.setdefault("emp", j)
            elif "ttc" in t:
                continue
            elif "puht" in t or ("pu" in t and "ht" in t) or "prix" in t:
                cols.setdefault("pu", j)
            elif any(k in t for k in ("exist", "quant", "qt", "nombre", "stock")):
                cols.setdefault("qty", j)
            elif any(k in t for k in ("seuil", "min", "alerte")):
                cols.setdefault("seuil", j)
            elif "remarque" in t or "note" in t:
                cols.setdefault("note", j)
            elif "date" in t:
                cols.setdefault("date", j)

        def cell(r, k):
            return r[cols[k]] if k in cols and cols[k] < len(r) else None

        self.backup()
        n = 0
        with self.tx():
            for r in rows[hdr + 1:]:
                name = str(cell(r, "name") or "").strip()
                if not name or name.startswith("#") or name.lower().startswith("total"):
                    continue
                date_iso, extra = parse_date_cell(cell(r, "date"))
                note = " | ".join(x for x in (str(cell(r, "note") or "").strip().lstrip("#"), extra) if x)
                ref = str(cell(r, "ref") or "").strip()
                cat = str(cell(r, "cat") or "Général").strip() or "Général"
                unite = str(cell(r, "unite") or "pcs").strip() or "pcs"
                emp = str(cell(r, "emp") or "").strip()

                a = self.article_by_name(name)
                if not a:
                    aid = self._next(self.arts)
                    a = {
                        "id": aid,
                        "ref": ref or f"REF-{aid:04d}",
                        "categorie": cat,
                        "unite": unite,
                        "emplacement": emp,
                        "seuil_alerte": parse_num(cell(r, "seuil"), 5.0),
                        "taux_tva": 19.0,
                        "appliquer_10_pct": False
                    }
                    self.arts.append(a)
                else:
                    if ref:
                        a["ref"] = ref
                    if cat != "Général":
                        a["categorie"] = cat
                    if unite != "pcs":
                        a["unite"] = unite
                    if emp:
                        a["emplacement"] = emp
                    if cell(r, "seuil") is not None:
                        a["seuil_alerte"] = parse_num(cell(r, "seuil"), a["seuil_alerte"])

                a.update(
                    designation=name,
                    nombre_existant=parse_num(cell(r, "qty")),
                    pu_ht=parse_num(cell(r, "pu")),
                    note=note,
                    date_saisie=date_iso
                )
                n += 1
            self._log(None, "IMPORT", agent, f"{n} article(s) importé(s) depuis {os.path.basename(path)}", 0, 0, 0)
        return n


# -------------------------------------------------------------------- Exports PDF & Excel
def pdf_safe(s):
    return str(s if s is not None else "").encode("latin-1", "replace").decode("latin-1")


def money(x):
    return f"{float(x):.3f}"


class ReportPDF(FPDF):
    title_txt = ""
    company_name = "THINK POWER"

    def header(self):
        self.set_font("Helvetica", "B", 15)
        self.set_text_color(31, 83, 141)
        self.cell(0, 8, f"{self.company_name.upper()} — GESTION DE STOCK INTELLIGENTE", new_x="LMARGIN", new_y="NEXT")
        self.set_font("Helvetica", "", 10)
        self.set_text_color(80, 80, 80)
        self.cell(0, 6, pdf_safe(self.title_txt), new_x="LMARGIN", new_y="NEXT")
        self.ln(2)
        self.set_text_color(0, 0, 0)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 6, f"Page {self.page_no()} - Document officiel généré le {datetime.now().strftime(DISP)}", align="C")


def pdf_table(pdf, headers, widths, rows, aligns):
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(31, 83, 141)
    pdf.set_text_color(255, 255, 255)
    for h, w in zip(headers, widths):
        pdf.cell(w, 7, pdf_safe(h), border=1, fill=True, align="C")
    pdf.ln()
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(0, 0, 0)
    for r in rows:
        for v, w, al in zip(r, widths, aligns):
            pdf.cell(w, 6, pdf_safe(v)[: int(w * 0.62)], border=1, align=al)
        pdf.ln()


def export_pdf(db, movs, path, period_txt):
    mode = db.mode10
    cur = db.currency
    pdf = ReportPDF(orientation="L", format="A4")
    pdf.company_name = db.company_name
    pdf.title_txt = f"RAPPORT GÉNÉRAL DU STOCK ET MOUVEMENTS — {period_txt}"
    pdf.add_page()
    tot = {"htva": 0.0, "adj": 0.0, "net": 0.0, "tva": 0.0, "ttc": 0.0}
    for a in db.arts:
        c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], mode)
        for k in tot:
            tot[k] += c[k]

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, f"Synthèse financière du stock ({cur})", new_x="LMARGIN", new_y="NEXT")
    pdf_table(
        pdf,
        ["Total HTVA", f"Ajustement 10% ({mode})", "HTVA net", "TVA", "Total TTC"],
        [50, 55, 50, 45, 50],
        [[money(tot["htva"]), money(tot["adj"]), money(tot["net"]), money(tot["tva"]), money(tot["ttc"])]],
        ["R"] * 5
    )
    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Mouvements du journal", new_x="LMARGIN", new_y="NEXT")
    rows = [
        [
            fmt_date(m.get("date")),
            m.get("type", ""),
            m.get("ref", ""),
            m.get("designation", ""),
            m.get("agent", ""),
            m.get("societe", ""),
            f"{m.get('avant', 0):g}",
            f"{m.get('qte', 0):g}",
            f"{m.get('apres', 0):g}",
            money(m.get("total_ttc", 0))
        ] for m in movs
    ]
    pdf_table(
        pdf,
        ["Date", "Type", "Réf", "Article", "Agent", "Client/Fourn.", "Avant", "Qté", "Reste", f"TTC ({cur})"],
        [28, 20, 24, 60, 26, 36, 15, 15, 15, 25],
        rows,
        ["L", "L", "L", "L", "L", "L", "R", "R", "R", "R"]
    )
    pdf.ln(12)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(130, 6, "Visa Responsable Stock")
    pdf.cell(0, 6, "Visa Direction", new_x="LMARGIN", new_y="NEXT")
    try:
        pdf.output(path)
    except PermissionError:
        raise RuntimeError(
            f"Impossible d'enregistrer « {os.path.basename(path)} » car ce fichier est actuellement ouvert dans un lecteur PDF.\n"
            "Veuillez fermer votre lecteur PDF puis réessayer."
        )


def export_bon_pdf(db, m, path):
    cur = db.currency
    pdf = ReportPDF(format="A4")
    pdf.company_name = db.company_name
    pdf.title_txt = f"BON OFFICIEL {m.get('bon', '')}"
    pdf.add_page()
    items = [
        ("N° Bon", m.get("bon", "")),
        ("Date & Heure", fmt_date(m.get("date"))),
        ("Nature de l'opération", m.get("type", "SORTIE")),
        ("Agent responsable", m.get("agent", "")),
        ("Bénéficiaire / Société", m.get("societe", "")),
        ("Référence / Code barre", m.get("ref", "")),
        ("Désignation de l'article", m.get("designation", "")),
        ("Stock avant opération", f"{m.get('avant', 0):g}"),
        ("Quantité livrée / mouvementée", f"{m.get('qte', 0):g}"),
        ("Stock restant en magasin", f"{m.get('apres', 0):g}"),
        ("Prix Unitaire HT", f"{money(m.get('pu_ht', 0))} {cur}"),
        ("Valeur Totale TTC", f"{money(m.get('total_ttc', 0))} {cur}"),
        ("Note / Remarque", m.get("note", ""))
    ]
    for label, val in items:
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(60, 7, pdf_safe(label + " :"))
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 7, pdf_safe(val), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(14)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(95, 6, "Signature & Cachet Émetteur")
    pdf.cell(0, 6, "Signature & Visa Réceptionnaire", new_x="LMARGIN", new_y="NEXT")
    try:
        pdf.output(path)
    except PermissionError:
        raise RuntimeError(
            f"Impossible d'enregistrer « {os.path.basename(path)} » car ce fichier est actuellement ouvert dans un lecteur PDF.\n"
            "Veuillez fermer votre lecteur PDF puis réessayer."
        )


def export_reorder_pdf(db, path):
    """Génère la feuille d'aide à la commande de réapprovisionnement intelligente."""
    analysis = db.smart_analysis()
    reorder = analysis["reorder_list"]
    cur = db.currency

    pdf = ReportPDF(orientation="L", format="A4")
    pdf.company_name = db.company_name
    pdf.title_txt = f"FEUILLE DE RÉAPPROVISIONNEMENT & COMMANDE SUGGÉRÉE ({len(reorder)} articles)"
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, f"Estimation du budget de réapprovisionnement : {money(analysis['reorder_total_cost'])} {cur} TTC", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    rows = [
        [
            r["ref"],
            r["designation"],
            r["categorie"],
            r["emplacement"],
            f"{r['actuel']:g} {r['unite']}",
            f"{r['seuil']:g} {r['unite']}",
            f"{r['suggested']:g} {r['unite']}",
            money(r["pu_ht"]),
            money(r["cost_ttc"]),
            r["statut"]
        ] for r in reorder
    ]
    pdf_table(
        pdf,
        ["Réf", "Désignation", "Catégorie", "Emplacement", "En Stock", "Seuil", "Qté Suggérée", f"PU HT ({cur})", f"Total TTC ({cur})", "Statut"],
        [24, 60, 32, 24, 20, 18, 22, 22, 26, 26],
        rows or [["-", "Aucun article nécessitant un réapprovisionnement", "-", "-", "-", "-", "-", "-", "-", "Stock OK"]],
        ["L", "L", "L", "L", "R", "R", "R", "R", "R", "C"]
    )
    pdf.ln(12)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(130, 6, "Visa Responsable des Achats")
    pdf.cell(0, 6, "Approbation Direction", new_x="LMARGIN", new_y="NEXT")
    try:
        pdf.output(path)
    except PermissionError:
        raise RuntimeError(
            f"Impossible d'enregistrer « {os.path.basename(path)} » car ce fichier est actuellement ouvert dans un lecteur PDF.\n"
            "Veuillez fermer votre lecteur PDF puis réessayer."
        )


def export_excel(db, movs, path):
    cur = db.currency
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "État du Stock"
    sheet_head(
        ws,
        ["Référence", "Matière / Désignation", "Catégorie", "Unité", "Emplacement",
         "Nombre Existant", "Seuil", "PU HT", "TVA %", "Option 10%", "PU TTC",
         "Total HTVA", "Total TTC", "Date de saisie", "Note"],
        (14, 36, 18, 8, 14, 15, 8, 12, 8, 11, 12, 14, 14, 17, 30)
    )
    arts = db.articles()
    for i, a in enumerate(arts, start=2):
        c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], db.mode10)
        ws.append([
            a.get("ref", ""), a["designation"], a.get("categorie", "Général"), a.get("unite", "pcs"),
            a.get("emplacement", ""), a["nombre_existant"], a["seuil_alerte"], a["pu_ht"], a["taux_tva"],
            "Oui" if a["appliquer_10_pct"] else "Non", a["pu_ht"] * (1 + a["taux_tva"] / 100), c["net"], c["ttc"],
            fmt_date(a["date_saisie"]), a["note"]
        ])
        for col in ["H", "K", "L", "M"]:
            ws[col + str(i)].number_format = DT

    n = max(len(arts) + 1, 2)
    h = wb.create_sheet("Historique")
    sheet_head(
        h,
        ["Bon", "Date", "Type", "Réf", "Article", "Agent", "Client / Fournisseur", "Avant", "Qté", "Reste", "PU HT",
         "HTVA net", "TVA", "TTC", "Note"],
        (13, 17, 13, 14, 34, 16, 24, 8, 8, 8, 11, 12, 11, 12, 30)
    )
    for i, m in enumerate(movs, start=2):
        h.append([
            m.get("bon", ""), fmt_date(m.get("date")), m.get("type", ""), m.get("ref", ""),
            m.get("designation", ""), m.get("agent", ""), m.get("societe", ""),
            m.get("avant", 0), m.get("qte", 0), m.get("apres", 0), m.get("pu_ht", 0),
            m.get("total_htva", 0), m.get("tva", 0), m.get("total_ttc", 0), m.get("note", "")
        ])
        for col in ["K", "L", "M", "N"]:
            h[col + str(i)].number_format = DT

    v = wb.create_sheet("Valorisation")
    sheet_head(v, ["Indicateur", f"Montant ({cur})"], (38, 18))
    v.append(["Total HTVA net du stock", f"=SUM('État du Stock'!L2:L{n})"])
    v.append(["Total TTC du stock", f"=SUM('État du Stock'!M2:M{n})"])
    v.append(["Total TVA du stock", "=B3-B2"])
    v.append(["Total TTC des mouvements listés", f"=SUM(Historique!N2:N{max(len(movs) + 1, 2)})"])
    for r in range(2, 6):
        v.cell(r, 2).number_format = DT

    try:
        wb.save(path)
    except PermissionError:
        raise RuntimeError(
            f"Impossible d'enregistrer « {os.path.basename(path)} » car ce fichier est actuellement ouvert dans Excel.\n"
            "Veuillez fermer Excel puis réessayer."
        )
