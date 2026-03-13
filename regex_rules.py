"""
Règles regex adaptées aux polices Atradius (format Modula)
Testées sur police 1453314 USD — BNP PARIBAS FACTOR (AMARIS CORP)

Champs extraits :
  - numero_police, assure, assureur, courtier
  - taux_prime, prime_provisionnelle, quotites_garanties
  - delai_indemnisation, delai_max_credit
  - date_prise_effet, date_echeance, duree_police
  - devise, limite_decaissement, zone_discretionnaire
  - groupe_polices  (liste complète)
"""

import re
from typing import Optional, Tuple, List, Dict


# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────

def _first_match(patterns: list, text: str) -> Tuple[Optional[str], float]:
    for pattern, confidence in patterns:
        m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if m:
            val = m.group(1).strip()
            val = re.sub(r'\s+', ' ', val)
            return val, confidence
    return None, 0.0


def _clean(val: Optional[str]) -> Optional[str]:
    if not val:
        return None
    return val.strip().strip('.,;:')


# ─────────────────────────────────────────────────────────────
# Extracteurs individuels
# ─────────────────────────────────────────────────────────────

def extract_numero_police(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'Police\s*:\s*(\d{6,8}(?:\s+[A-Z]{2,4})?)', 0.95),
        (r'N[o°]?\s*de\s*police\s*[:\-]?\s*(\d+(?:\s+[A-Z]{2,4})?)', 0.85),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_assure(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'Assur[eé]\s+Org\s+ID:\d+\s+([A-Z][^\n]+?)(?:\n\d|\nOrg|\nCourtier|\nDate)', 0.92),
        (r"d[eé]nomm[eé] l.Assur[eé],\s*\n([A-Z][^\n]+)", 0.90),
        (r'^Assur[eé]\s*\n(.+?)(?:\n|$)', 0.80),
        (r'Assur[eé]\s*[:\-]\s*(.+?)(?:\n|Courtier|$)', 0.75),
    ]
    val, conf = _first_match(patterns, text)
    if val:
        lines = [l.strip() for l in val.split('\n') if l.strip()]
        val = lines[0] if lines else val
        val = re.sub(r'\s+(Org\s+ID|Courtier|Date|Dur[eé]e).*$', '', val, flags=re.IGNORECASE)
    return _clean(val), conf


def extract_assureur(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'(Atradius\s+Cr[eé]dito\s+y\s+Cauc[ií][oó]n\s+S\.A\.(?:\s+de\s+Seguros\s+y\s+Reaseguros)?)', 0.95),
        (r'(Atradius[^,\n]{0,60})', 0.85),
        (r"[Ll]'[Aa]ssureur[,\s]+(.+?)(?:\n|,|$)", 0.70),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_courtier(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        # Format Atradius CP : "Courtier NOM\nOrg ID:XXXX ..."  (nom sur même ligne que label)
        (r'Courtier\s+([A-Z][^\n]+?)(?:\nOrg\s+ID|\n\d)', 0.92),
        (r'Courtier\s+Org\s+ID:\d+\s+([A-Z][^\n]+?)(?:\n\d|\nOrg|\nDate)', 0.88),
        (r'^Courtier\s*\n(.+?)(?:\n|$)', 0.85),
        (r'Courtier\s*[:\-]\s*(.+?)(?:\n|$)', 0.80),
    ]
    val, conf = _first_match(patterns, text)
    if val:
        lines = [l.strip() for l in val.split('\n') if l.strip()]
        val = lines[0] if lines else val
        val = re.sub(r'\s+(Org\s+ID|Date|Dur[eé]e).*$', '', val, flags=re.IGNORECASE)
    return _clean(val), conf


def extract_taux_prime(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'Taux\s+de\s+prime\s+Risque\s+Cr[eé]dit\s*\n.{0,80}?([\d]+[,\.][\d]+\s*%)', 0.95),
        (r'(?:taux|rate)\s+(?:de\s+)?prime\s*[:\-]?\s*([\d,\.]+\s*%)', 0.85),
        (r'(\d+[,\.]\d+\s*%)\s*(?:\n|$)', 0.70),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_prime_provisionnelle(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        # "4.510 (total)" — ligne totale Atradius
        (r'\n([\d\s\.]+\d)\s*\(total\)', 0.95),
        (r'[Pp]rime\s+[Pp]rovisionnelle\s*[:\-]?\s*([\d\s\.,]+(?:\s*[A-Z]{3})?)', 0.75),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_quotites_garanties(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'[Pp]ourcentage\s+assur[eé]\s+([\d]+\s*%)', 0.95),
        (r'[Qq]uotit[eé]s?\s+(?:garanties?|assur[eé]es?)\s*[:\-]?\s*([\d]+\s*%)', 0.90),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_delai_indemnisation(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r"[Dd][eé]lai\s+d.indemnisation\s+([\d]+\s+mois)", 0.95),
        (r"[Dd][eé]lai\s+d.indemnisation\s*[:\-]?\s*([\d]+\s+(?:mois|jours?))", 0.90),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_delai_max_credit(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'[Dd][eé]lai\s+maximum\s+de\s+cr[eé]dit\s+consenti\s+([\d]+\s+jours?)', 0.95),
        (r'[Dd][eé]lai\s+(?:max(?:imum)?\s+de\s+)?cr[eé]dit\s*[:\-]?\s*([\d]+\s+(?:jours?|mois))', 0.85),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_date_prise_effet(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r"Date\s+de\s+prise\s+d.effet\s+(\d{1,2}(?:er)?\s+\w+\s+\d{4})", 0.95),
        (r"Date\s+de\s+prise\s+d.effet\s+(\d{2}[/-]\d{2}[/-]\d{4})", 0.90),
        (r"prise\s+d.effet\s*[:\-]?\s*(\d{1,2}[/-]\d{2}[/-]\d{4})", 0.85),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_date_echeance(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r"jusqu.au\s+(\d{1,2}(?:er)?\s+\w+\s+\d{4})\s+inclus", 0.95),
        (r"Ann[eé]e\s+d.assurance\s+.+?jusqu.au\s+(\d{1,2}(?:er)?\s+\w+\s+\d{4})", 0.90),
        (r"[eé]ch[eé]ance\s*[:\-]?\s*(\d{1,2}[/-]\d{2}[/-]\d{4})", 0.85),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_duree_police(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'Dur[eé]e\s+de\s+la\s+Police\s+([\d]+\s+mois)', 0.95),
        (r'dur[eé]e\s*[:\-]?\s*([\d]+\s+(?:mois|ans?))', 0.85),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_devise(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'Devise\s+de\s+la\s+police\s+([^\n]+?)(?:\n|$)', 0.95),
        (r'[Dd]evise\s*[:\-]?\s*([A-Z]{3}|Dollar[\w\s]*|Euro[\w\s]*)', 0.85),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_limite_decaissement(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        # "Maximum d'indemnité pour un CAD 2.000.000 ou\ngroupe de polices 73 fois..."
        (r"Maximum\s+d.indemni[ts][eé]\s+pour\s+un\s+([A-Z]{2,3}\s+[\d\s\.]+(?:\s+ou\s+[\d]+\s+fois[^\n]+)?)", 0.95),
        (r"[Ll]imite\s+(?:de\s+)?d[eé]caissement\s*[:\-]?\s*([^\n]+)", 0.85),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


def extract_zone_discretionnaire(text: str) -> Tuple[Optional[str], float]:
    patterns = [
        (r'Montant\s+du\s+Credit\s+Check\s+([\d\s\.,]+)', 0.95),
        (r'[Cc]redit\s+[Cc]heck\s*[:\-]?\s*([\d\s\.,]+)', 0.90),
        (r'zone\s+discr[eé]tionnaire\s*[:\-]?\s*([\d\s\.,]+)', 0.80),
    ]
    val, conf = _first_match(patterns, text)
    return _clean(val), conf


# ─────────────────────────────────────────────────────────────
# Extraction du groupe de polices (liste complète)
# ─────────────────────────────────────────────────────────────

def extract_groupe_polices(text: str) -> Tuple[Optional[List[Dict]], float]:
    """
    Extrait la liste complète des polices du groupe.
    Retourne une liste de dicts :
      [{"numero": "1090335", "devise": None, "assure": "AMARIS FRANCE SAS"}, ...]
    """
    # Pattern Atradius : "Police: NNNNNN [DEVISE] NOM"  ou  "Police: NNNNNN - DEVISE NOM"
    pattern = re.compile(
        r'Police\s*:\s*(\d{6,8})'               # numéro
        r'(?:\s*[-–]?\s*([A-Z]{2,4}))?'         # devise optionnelle
        r'\s+([A-Z][^\n]{2,80})',                # nom assuré
        re.MULTILINE
    )

    seen: set = set()
    polices: List[Dict] = []

    for m in pattern.finditer(text):
        numero = m.group(1).strip()
        devise = m.group(2).strip() if m.group(2) else None
        nom_raw = m.group(3).strip()

        # Nettoyer la date de modification éventuelle en fin de ligne (ex: 01-01-2026)
        nom = re.sub(r'\s+\d{2}-\d{2}-\d{4}.*$', '', nom_raw).strip()
        nom = re.sub(r'\s{2,}', ' ', nom)
        # Supprimer artefacts PDF de début (ex: "MCL_Addr_Marker")
        nom = re.sub(r'^[A-Za-z_]+Marker\s*', '', nom).strip()
        # Supprimer préfixe pays isolé (ex: "Autriche ", "Suisse - ")
        nom = re.sub(r'^(?:Autriche|Suisse|Austria|Belgium|France|Germany)\s*[-–]?\s*', '', nom, flags=re.IGNORECASE).strip()

        if numero not in seen:
            seen.add(numero)
            polices.append({
                "numero": numero,
                "devise": devise,
                "assure": nom,
            })

    if polices:
        return polices, 0.92
    return None, 0.0


# ─────────────────────────────────────────────────────────────
# Fonction principale — point d'entrée
# ─────────────────────────────────────────────────────────────

def extract_all_fields(text: str) -> dict:
    """
    Extrait tous les champs et retourne un dict :
      { champ: {"value": ..., "confidence": float, "source": "regex"} }
    """
    extractors = {
        "numero_police":        extract_numero_police,
        "assure":               extract_assure,
        "assureur":             extract_assureur,
        "courtier":             extract_courtier,
        "taux_prime":           extract_taux_prime,
        "prime_provisionnelle": extract_prime_provisionnelle,
        "quotites_garanties":   extract_quotites_garanties,
        "delai_indemnisation":  extract_delai_indemnisation,
        "delai_max_credit":     extract_delai_max_credit,
        "date_prise_effet":     extract_date_prise_effet,
        "date_echeance":        extract_date_echeance,
        "duree_police":         extract_duree_police,
        "devise":               extract_devise,
        "limite_decaissement":  extract_limite_decaissement,
        "zone_discretionnaire": extract_zone_discretionnaire,
    }

    results = {}
    for field, func in extractors.items():
        val, conf = func(text)
        results[field] = {
            "value":      val,
            "confidence": conf,
            "source":     "regex" if val else None,
        }

    # Groupe de polices — structure liste, traitement séparé
    polices, conf = extract_groupe_polices(text)
    results["groupe_polices"] = {
        "value":      polices,
        "confidence": conf,
        "source":     "regex" if polices else None,
    }

    return results
