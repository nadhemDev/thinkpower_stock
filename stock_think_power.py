# -*- coding: utf-8 -*-
"""Stock Think Power - moteur de données.
La base de données EST un fichier Excel (Stock_Think_Power_Base.xlsx) placé à côté du .exe.
Toutes les opérations (ajout, modification, suppression, sortie, inventaire)
sont enregistrées directement et immédiatement dans ce fichier Excel.
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
        # Repli propre sur AppData local si l'exécutable est placé dans un dossier protégé (ex: Program Files)
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
    "ID", "Matière / Désignation", "Nombre Existant", "Stock Minimum", "PU HT", "TVA %",
    "Option 10%", "PU TTC", "Total HTVA", "Total TTC", "Date de saisie", "État", "Note"
]
JOURNAL_HEAD = [
    "ID", "Bon", "Date", "Type", "Agent", "Société / Client", "Article", "ID Article",
    "Avant", "Qté", "Reste", "PU HT", "HTVA net", "TVA", "TTC", "Note"
]


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
        for cur in ("DT", "dt", "Dt", "TND", "tnd", "€", "$"):
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
            # Nettoyage automatique : garder les 20 plus récents
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
            name = str(r.get("Matière / Désignation") or "").strip()
            if not name or name.upper().startswith("TOTAL"):
                continue
            self.arts.append(dict(
                id=int(parse_num(r.get("ID"))),
                designation=name,
                nombre_existant=parse_num(r.get("Nombre Existant")),
                seuil_alerte=parse_num(r.get("Stock Minimum"), 5),
                pu_ht=parse_num(r.get("PU HT")),
                taux_tva=parse_num(r.get("TVA %"), 19),
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
                societe=str(r.get("Société / Client") or ""),
                designation=str(r.get("Article") or ""),
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
        """Enregistre l'ensemble des données dans Stock_Think_Power_Base.xlsx.
        Si une ligne est supprimée de self.arts ou self.movs, elle est directement supprimée du fichier Excel.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Stock"
        sheet_head(ws, STOCK_HEAD, (6, 40, 15, 14, 12, 8, 11, 12, 13, 13, 17, 11, 34))
        mode = self.mode10
        colors = {"ok": "1E8449", "faible": "D68910", "rupture": "C0392B"}

        for i, a in enumerate(sorted(self.arts, key=lambda x: str(x["designation"]).lower()), start=2):
            c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], mode)
            e = etat(a)
            ws.append([
                a["id"],
                a["designation"],
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
            for col in "EHIJ":
                ws[col + str(i)].number_format = DT
            ws["L" + str(i)].font = Font(bold=True, color=colors[e])

        n = len(self.arts) + 1
        if self.arts:
            ws.append(["", "TOTAL", "", "", "", "", "", "", f"=SUM(I2:I{n})", f"=SUM(J2:J{n})"])
            for col in "IJ":
                ws[col + str(n + 1)].number_format = DT
            for c in ws[n + 1]:
                c.font = Font(bold=True)
            ws.auto_filter.ref = f"A1:M{n}"

        j = wb.create_sheet("Journal")
        sheet_head(j, JOURNAL_HEAD, (6, 13, 17, 13, 16, 24, 36, 10, 8, 8, 8, 11, 12, 11, 12, 36))
        for i, m in enumerate(sorted(self.movs, key=lambda x: (x["date"], x["id"]), reverse=True), start=2):
            j.append([
                m["id"],
                m["bon"],
                fmt_date(m["date"]),
                m["type"],
                m["agent"],
                m["societe"],
                m["designation"],
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
            for col in "LMNO":
                j[col + str(i)].number_format = DT
        if self.movs:
            j.auto_filter.ref = f"A1:P{len(self.movs) + 1}"

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

    # -- Paramètres
    @property
    def mode10(self):
        return self.settings.get("mode10", "remise")

    def set_mode10(self, v):
        with self.tx():
            self.settings["mode10"] = v

    @property
    def last_agent(self):
        return self.settings.get("dernier_agent", "")

    # -- Lecture
    def articles(self, search="", mode="all"):
        s = search.lower().strip()
        rows = [a for a in self.arts if s in a["designation"].lower() or s in a.get("note", "").lower()]
        if mode == "alert":
            rows = [a for a in rows if a["nombre_existant"] <= a["seuil_alerte"]]
        elif mode == "rupture":
            rows = [a for a in rows if a["nombre_existant"] <= 0]
        elif mode == "ok":
            rows = [a for a in rows if a["nombre_existant"] > a["seuil_alerte"]]
        return sorted(rows, key=lambda a: str(a["designation"]).lower())

    def article(self, aid):
        return next((a for a in self.arts if a["id"] == aid), None)

    def article_by_name(self, name):
        n = (name or "").strip().lower()
        return next((a for a in self.arts if a["designation"].lower() == n), None)

    def agents(self):
        names = {m.get("agent", "").strip() for m in self.movs if m.get("agent", "").strip()}
        if self.last_agent:
            names.add(self.last_agent.strip())
        return sorted([n for n in names if n], key=str.lower)

    def movements(self, agent="", societe="", d1=None, d2=None, typ=""):
        out = []
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
            out.append(m)
        return sorted(out, key=lambda m: (m.get("date", ""), m.get("id", 0)), reverse=True)

    def article_movements(self, aid, limit=6):
        return sorted(
            [m for m in self.movs if m.get("article_id") == aid],
            key=lambda m: (m.get("date", ""), m.get("id", 0)),
            reverse=True
        )[:limit]

    def kpis(self):
        mode, htva, ttc = self.mode10, 0.0, 0.0
        for a in self.arts:
            c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], mode)
            htva += c["net"]
            ttc += c["ttc"]
        month = datetime.now().strftime("%Y-%m")
        return {
            "refs": len(self.arts),
            "htva": htva,
            "ttc": ttc,
            "rupt": sum(1 for a in self.arts if etat(a) == "rupture"),
            "crit": sum(1 for a in self.arts if etat(a) == "faible"),
            "sorties": sum(m["qte"] for m in self.movs if m.get("type") == "SORTIE" and str(m.get("date", "")).startswith(month))
        }

    # -- Écriture
    def _next(self, rows):
        return max([r["id"] for r in rows] + [0]) + 1

    def next_bon(self):
        y = datetime.now().year
        n = sum(1 for m in self.movs if m.get("type") == "SORTIE" and str(m.get("bon", "")).startswith(f"BS-{y}-"))
        return f"BS-{y}-{n + 1:03d}"

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
            if aid:
                a = self.article(aid)
                old, typ, note = a["nombre_existant"], "MODIF", "Fiche article modifiée"
            else:
                a = {"id": self._next(self.arts)}
                self.arts.append(a)
                old, typ, note = 0, "NOUVEAU", "Article créé"
            a.update(
                designation=data["designation"].strip(),
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
        """Supprime définitivement la ligne de l'article de la base Excel."""
        self.backup()
        with self.tx():
            a = self.article(aid)
            if a and agent:
                self._log(a, "SUPPRESSION", agent, "Article supprimé", a["nombre_existant"], 0)
            self.arts = [x for x in self.arts if x["id"] != aid]

    def delete_movement(self, mid):
        """Supprime directement une ligne du journal dans la base Excel."""
        self.backup()
        with self.tx():
            self.movs = [m for m in self.movs if m["id"] != mid]

    def add_stock(self, aid, qty, agent="", note="", date_iso=None):
        """Ajout direct de stock (réapprovisionnement / entrée) depuis l'onglet Stock."""
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
        with self.tx():
            a = self.article(aid)
            a["nombre_existant"] = apres
            a["date_saisie"] = now_iso()
            m = self._log(
                a, "ENTREE", agent,
                note or "Réapprovisionnement / Entrée de stock",
                avant, apres, qty, "", date_iso, "", a["pu_ht"], cost
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
            apres, bon = avant - qty, self.next_bon()
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
            c and ("mati" in str(c).lower() or "design" in str(c).lower()) for c in r)), None)
        if hdr is None:
            raise ValueError("En-tête « Matière / Désignation » introuvable dans la feuille sélectionnée.")
        cols = {}
        for j, c in enumerate(rows[hdr]):
            t = str(c or "").lower().replace(" ", "").replace(".", "")
            if ("mati" in t or "design" in t) and "name" not in cols:
                cols["name"] = j
            elif "ttc" in t:
                continue
            elif "puht" in t or ("pu" in t and "ht" in t):
                cols.setdefault("pu", j)
            elif any(k in t for k in ("exist", "quant", "qt", "nombre")):
                cols.setdefault("qty", j)
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
                a = self.article_by_name(name)
                if not a:
                    a = {"id": self._next(self.arts), "seuil_alerte": 5.0, "taux_tva": 19.0, "appliquer_10_pct": False}
                    self.arts.append(a)
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

    def header(self):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(31, 83, 141)
        self.cell(0, 8, "THINK POWER - GESTION DE STOCK", new_x="LMARGIN", new_y="NEXT")
        self.set_font("Helvetica", "", 10)
        self.set_text_color(80, 80, 80)
        self.cell(0, 6, pdf_safe(self.title_txt), new_x="LMARGIN", new_y="NEXT")
        self.ln(2)
        self.set_text_color(0, 0, 0)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 6, f"Page {self.page_no()} - Généré le {datetime.now().strftime(DISP)}", align="C")


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
    pdf = ReportPDF(orientation="L", format="A4")
    pdf.title_txt = f"ÉTAT D'INVENTAIRE ET MOUVEMENTS DE STOCK - {period_txt}"
    pdf.add_page()
    tot = {"htva": 0.0, "adj": 0.0, "net": 0.0, "tva": 0.0, "ttc": 0.0}
    for a in db.arts:
        c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], mode)
        for k in tot:
            tot[k] += c[k]
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Synthèse financière du stock (DT)", new_x="LMARGIN", new_y="NEXT")
    pdf_table(
        pdf,
        ["Total HTVA", f"Ajustement 10% ({mode})", "HTVA net", "TVA", "Total TTC"],
        [50, 55, 50, 45, 50],
        [[money(tot["htva"]), money(tot["adj"]), money(tot["net"]), money(tot["tva"]), money(tot["ttc"])]],
        ["R"] * 5
    )
    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Mouvements (par agent / société bénéficiaire)", new_x="LMARGIN", new_y="NEXT")
    rows = [
        [
            fmt_date(m.get("date")),
            m.get("type", ""),
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
        ["Date", "Type", "Article", "Agent", "Société / Client", "Avant", "Qté", "Reste", "TTC (DT)"],
        [32, 24, 66, 30, 40, 17, 17, 17, 28],
        rows,
        ["L", "L", "L", "L", "L", "R", "R", "R", "R"]
    )
    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Alertes de réapprovisionnement", new_x="LMARGIN", new_y="NEXT")
    alerts = [
        [a["designation"], f"{a['nombre_existant']:g}", f"{a['seuil_alerte']:g}", money(a["pu_ht"])]
        for a in db.articles(mode="alert")
    ]
    pdf_table(
        pdf,
        ["Article", "Existant", "Seuil", "PU HT (DT)"],
        [120, 30, 30, 40],
        alerts or [["Aucune alerte", "", "", ""]],
        ["L", "R", "R", "R"]
    )
    pdf.ln(12)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(130, 6, "Visa Agent Inventaire")
    pdf.cell(0, 6, "Visa Direction", new_x="LMARGIN", new_y="NEXT")
    try:
        pdf.output(path)
    except PermissionError:
        raise RuntimeError(
            f"Impossible d'enregistrer « {os.path.basename(path)} » car ce fichier est actuellement ouvert dans un lecteur PDF.\n"
            "Veuillez fermer votre lecteur PDF puis réessayer."
        )


def export_bon_pdf(db, m, path):
    pdf = ReportPDF(format="A4")
    pdf.title_txt = f"BON DE SORTIE {m.get('bon', '')}"
    pdf.add_page()
    items = [
        ("N° Bon", m.get("bon", "")),
        ("Date", fmt_date(m.get("date"))),
        ("Agent émetteur", m.get("agent", "")),
        ("Société / Bénéficiaire", m.get("societe", "")),
        ("Matière / Article", m.get("designation", "")),
        ("Stock avant sortie", f"{m.get('avant', 0):g}"),
        ("Quantité livrée", f"{m.get('qte', 0):g}"),
        ("Stock restant", f"{m.get('apres', 0):g}"),
        ("Prix Unitaire HT", f"{money(m.get('pu_ht', 0))} DT"),
        ("Valeur Totale TTC", f"{money(m.get('total_ttc', 0))} DT"),
        ("Remarque / Note", m.get("note", ""))
    ]
    for label, val in items:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(55, 8, pdf_safe(label + " :"))
        pdf.set_font("Helvetica", "", 11)
        pdf.multi_cell(0, 8, pdf_safe(val), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(16)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(95, 6, "Signature & Visa Agent")
    pdf.cell(0, 6, "Signature Réceptionnaire", new_x="LMARGIN", new_y="NEXT")
    try:
        pdf.output(path)
    except PermissionError:
        raise RuntimeError(
            f"Impossible d'enregistrer « {os.path.basename(path)} » car ce fichier est actuellement ouvert dans un lecteur PDF.\n"
            "Veuillez fermer votre lecteur PDF puis réessayer."
        )


def export_excel(db, movs, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "État du Stock"
    sheet_head(
        ws,
        ["Matière / Désignation", "Nombre Existant", "Seuil", "PU HT", "TVA %", "Option 10%", "PU TTC",
         "Total HTVA", "Total TTC", "Date de saisie", "Note"],
        (38, 15, 8, 12, 8, 11, 12, 14, 14, 17, 30)
    )
    arts = db.articles()
    for i, a in enumerate(arts, start=2):
        c = calc(a["nombre_existant"], a["pu_ht"], a["taux_tva"], a["appliquer_10_pct"], db.mode10)
        ws.append([
            a["designation"], a["nombre_existant"], a["seuil_alerte"], a["pu_ht"], a["taux_tva"],
            "Oui" if a["appliquer_10_pct"] else "Non", a["pu_ht"] * (1 + a["taux_tva"] / 100), c["net"], c["ttc"],
            fmt_date(a["date_saisie"]), a["note"]
        ])
        for col in "DGHI":
            ws[col + str(i)].number_format = DT

    n = max(len(arts) + 1, 2)
    h = wb.create_sheet("Historique")
    sheet_head(
        h,
        ["Bon", "Date", "Type", "Article", "Agent", "Société / Client", "Avant", "Qté", "Reste", "PU HT",
         "HTVA net", "TVA", "TTC", "Note"],
        (13, 17, 13, 36, 16, 24, 8, 8, 8, 11, 12, 11, 12, 30)
    )
    for i, m in enumerate(movs, start=2):
        h.append([
            m.get("bon", ""), fmt_date(m.get("date")), m.get("type", ""), m.get("designation", ""),
            m.get("agent", ""), m.get("societe", ""), m.get("avant", 0), m.get("qte", 0),
            m.get("apres", 0), m.get("pu_ht", 0), m.get("total_htva", 0), m.get("tva", 0),
            m.get("total_ttc", 0), m.get("note", "")
        ])
        for col in "JKLM":
            h[col + str(i)].number_format = DT

    v = wb.create_sheet("Valorisation")
    sheet_head(v, ["Libellé", "Montant (DT)"], (38, 18))
    v.append(["Total HTVA net du stock", f"=SUM('État du Stock'!H2:H{n})"])
    v.append(["Total TTC du stock", f"=SUM('État du Stock'!I2:I{n})"])
    v.append(["Total TVA", "=B3-B2"])
    v.append(["Total TTC des mouvements listés", f"=SUM(Historique!M2:M{max(len(movs) + 1, 2)})"])
    for r in range(2, 6):
        v.cell(r, 2).number_format = DT
    try:
        wb.save(path)
    except PermissionError:
        raise RuntimeError(
            f"Impossible d'enregistrer « {os.path.basename(path)} » car ce fichier est actuellement ouvert dans Excel.\n"
            "Veuillez fermer Excel puis réessayer."
        )
