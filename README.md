# ⚡ THINK POWER — Gestion de Stock Intelligente & Universelle

Application moderne, autonome et intelligente de gestion de stock, de logistique et de réapprovisionnement, conçue pour tout type d'activité (matières premières, quincaillerie, pièces détachées, commerce, alimentaire, consommables, outillage, etc.).

---

## 🌟 Points Forts & Fonctionnalités Intelligentes

### 1. 📦 Gestion Universelle des Stocks
- **Multi-catégories & Familles d'articles** : organisation par rayon/famille avec filtre instantané.
- **Unités de mesure universelles** : pièces (`pcs`), kilogrammes (`kg`), mètres (`m`), litres (`L`), cartons, rouleaux, sacs, etc.
- **Emplacement précis en magasin** : suivi du rayon, de l'étagère ou du casier de stockage.
- **Scanner & Recherche Rapide par Code-Barres / Réf (SKU)** : sélection instantanée en scannant le code-barres ou en tapant la référence.
- **Multi-devise & Multi-société** : personnalisez le nom de votre entreprise et votre devise (`DT`, `€`, `$`, `MAD`, `DZD`, `CFA`...).

### 2. 🧠 Smart Advisor (Conseiller de Réapprovisionnement Intelligent)
- **Détection proactive des pénuries** : alerte automatique dès qu'un article passe sous son seuil de sécurité.
- **Calcul automatique du réapprovisionnement** : calcul de la quantité optimale à commander pour chaque référence en alerte.
- **Estimation budgétaire en temps réel** : valorisation financière globale du besoin de réapprovisionnement.
- **Bon de Commande Suggéré (PDF 1-clic)** : génération instantanée d'une feuille officielle de commande fournisseur prête pour visa et signature.
- **Analyse de rotation** :
  - 🔥 **Top 5 Forte Rotation** : articles les plus demandés sur les 30 derniers jours.
  - 💤 **Articles Dormants (Sur-stockage)** : détection des stocks sans mouvement pour éviter l'immobilisation de trésorerie.

### 3. 📥 Entrées Fournisseur & 📤 Sorties Sécurisées
- **Réception Fournisseur complète** : enregistrement des livraisons avec nom du fournisseur, N° de Bon de Livraison (BL), quantité et agent.
- **Sortie de stock avec contrôle temps réel** : vérification immédiate de la disponibilité pour bloquer les ruptures.
- **Édition automatique de Bons Officiels (PDF)** : bons de sortie et de réception avec visa émetteur et signature client/bénéficiaire.

### 4. 🗄️ Base de Données Autonome Excel
- **Zéro base de données complexe à installer** : toutes les données sont enregistrées en temps réel dans `Stock_Think_Power_Base.xlsx`.
- **Sauvegardes automatiques (Backups)** : chaque modification génère un backup horodaté dans le dossier `backups/`.
- **Import / Export universel** : importez n'importe quel fichier Excel ou exportez des rapports financiers complets en `.xlsx` et `.pdf`.

### 5. 💻 100% Portable (Fonctionne sur n'importe quel PC Windows)
- Aucun besoin d'installer Python sur les postes utilisateurs.
- Intègre directement les dépendances graphiques, les DLLs Visual C++ Runtime et les thèmes sombres.

---

## 🚀 Lancement & Compilation

### Lancer depuis les sources
```bash
py -m pip install -r requirements.txt
py main.py
```

### Compiler en `.exe` autonome tout-en-un
Double-cliquez simplement sur **`build_exe.bat`** ou lancez :
```bash
py -m PyInstaller --clean --noconfirm Stock_Think_Power.spec
```
L'exécutable portable est généré dans `dist\Stock_Think_Power.exe`.

---

## 📁 Structure du Projet

```
StockApp/
├── main.py                    # Interface graphique moderne CustomTkinter
├── stock_think_power.py       # Moteur de données Excel & logique métier intelligente
├── Stock_Think_Power.spec     # Configuration PyInstaller portable
├── build_exe.bat              # Script automatisé de compilation Windows
├── requirements.txt           # Dépendances Python requises
├── app_icon.ico               # Icône applicative vectorielle haute résolution
├── Stock_Think_Power_Base.xlsx# Base de données Excel temps réel
└── exports/                   # Dossier de génération des PDF et rapports Excel
```

---

## 📄 Licence & Droits
Développé par **nadhemDev** pour Think Power Stock. Tous droits réservés.
