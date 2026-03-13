# 📋 Analyseur de Contrats d'Assurance-Crédit

Extraction automatique des données contractuelles depuis des polices PDF/DOCX/TXT.  
Optimisé pour le format **Atradius Modula** — compatible Coface, Euler Hermes, et autres.

---

## Fichiers du projet

| Fichier | Rôle |
|---|---|
| `app.py` | Interface Streamlit (4 onglets) |
| `extractor.py` | Extraction texte PDF / DOCX / TXT |
| `regex_rules.py` | Patterns regex Atradius Modula |
| `nlp_parser.py` | Analyse sémantique via API Claude |
| `excel_exporter.py` | Export Excel ponctuel (session) |
| `db_manager.py` | Base persistante Excel (upsert, suppression, filtre) |
| `utils.py` | Utilitaires (log, formatage) |
| `requirements.txt` | Dépendances Python |

---

## Installation

```bash
git clone https://github.com/[votre-org]/credit-insurance-analyzer
cd credit-insurance-analyzer
pip install -r requirements.txt
```

---

## Lancement

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
export CIA_DB_PATH="C:\Users\...\SharePoint\contrats_base.xlsx"   # Windows
# ou
export CIA_DB_PATH="/Users/.../OneDrive-Entreprise/contrats_base.xlsx"  # Mac

streamlit run app.py
```

Le fichier Excel de base (`CIA_DB_PATH`) est **créé automatiquement** au premier enregistrement.  
Laissez vide pour utiliser `contrats_base.xlsx` dans le dossier courant.

---

## Champs extraits

| Champ | Exemple |
|---|---|
| N° Police | 1453314 USD |
| Assuré | BNP PARIBAS FACTOR (AMARIS CORP) |
| Assureur | Atradius Crédito y Caución S.A. |
| Courtier | WILLIS TOWERS WATSON FRANCE |
| Taux de prime | 0,052% |
| Prime provisionnelle | 4.510 |
| Pourcentage assuré | 95% |
| Délai d'indemnisation | 6 mois |
| Délai max crédit | 180 jours |
| Date de prise d'effet | 1er Décembre 2025 |
| Date d'échéance | 31 Décembre 2026 |
| Durée | 13 mois |
| Devise | Dollar US |
| Limite de décaissement | CAD 2.000.000 |
| Zone discrétionnaire | 23.000 |
| Groupe de polices | Liste complète (37 polices) |

---

## Base de données partagée (SharePoint / Teams)

1. Synchroniser le dossier SharePoint via le client OneDrive
2. Renseigner le chemin local synchronisé dans `CIA_DB_PATH`
3. Tous les membres du groupe voient les mises à jour en temps réel via la sync OneDrive

---

## Variables d'environnement

| Variable | Description | Défaut |
|---|---|---|
| `ANTHROPIC_API_KEY` | Clé API Claude (analyse IA) | — |
| `CIA_DB_PATH` | Chemin vers le fichier Excel base | `contrats_base.xlsx` |
