"""
regex_rules.py
Règles regex pour l'extraction des champs contractuels
"""

import re
from typing import Optional, Tuple

# ─────────────────────────────────────────────
# PATTERNS PAR CHAMP
# ─────────────────────────────────────────────

PATTERNS = {

    "assure": [
        r"(?:assuré|assure|policyholder|insured)\s*[:\-–]\s*([A-Z][^\n,;]{3,80})",
        r"(?:le présent contrat est souscrit par|souscripteur)\s*[:\-–]?\s*([A-Z][^\n,;]{3,80})",
        r"between\s+([A-Z][^\n,;]{3,60})\s+(?:hereinafter|as the insured)",
    ],

    "assureur": [
        r"(?:assureur|compagnie d'assurance|insurer|underwriter)\s*[:\-–]\s*([A-Z][^\n,;]{3,80})",
        r"(?:la société|l'entreprise)\s+([A-Z][^\n,;]{3,60})\s+(?:ci-après|en tant qu'assureur)",
        r"(?:issued by|underwritten by)\s+([A-Z][^\n,;]{3,80})",
    ],

    "taux_prime": [
        r"(?:taux de prime|taux prime|prime rate|insurance rate|taux d'assurance)\s*[:\-–]?\s*([\d,\.]+\s*%)",
        r"(?:premium rate|rate of premium)\s*[:\-–]?\s*([\d,\.]+\s*%)",
        r"([\d,\.]+\s*%)\s*(?:du chiffre d'affaires|of turnover|of sales)",
        r"(?:taux)\s*[:\-–]?\s*([\d,\.]+\s*‰)",  # pour mille
    ],

    "minimum_prime": [
        r"(?:prime minimum|minimum de prime|minimum premium|prime minimale)\s*[:\-–]?\s*([\d\s,\.]+\s*(?:EUR|€|\$|USD|GBP)?)",
        r"(?:minimum annual premium|MAP)\s*[:\-–]?\s*([\d\s,\.]+\s*(?:EUR|€|\$|USD)?)",
        r"(?:la prime ne pourra être inférieure à)\s*([\d\s,\.]+\s*(?:EUR|€)?)",
    ],

    "limite_decaissement": [
        r"(?:limite de décaissement|plafond d'indemnisation|maximum d'indemnisation|indemnity limit)\s*[:\-–]?\s*([\d\s,\.]+\s*(?:EUR|€|\$|USD|M€|K€)?)",
        r"(?:maximum aggregate|aggregate limit|limit of indemnity)\s*[:\-–]?\s*([\d\s,\.]+\s*(?:EUR|€|\$|USD)?)",
        r"(?:aggregate indemnity limit|AIL)\s*[:\-–]?\s*([\d\s,\.]+\s*(?:EUR|€|\$|USD)?)",
    ],

    "zone_discretionnaire": [
        r"(?:zone discrétionnaire|non.dénommé|non dénommé|discretionary limit|non.named buyers?)\s*[:\-–]?\s*([\d\s,\.]+\s*(?:EUR|€|\$|USD|%)?)",
        r"(?:limite discrétionnaire|acheteurs non dénommés)\s*[:\-–]?\s*([\d\s,\.]+\s*(?:EUR|€|\$|USD)?)",
        r"(?:buyer's? discretionary limit|BDL)\s*[:\-–]?\s*([\d\s,\.]+\s*(?:EUR|€|\$|USD)?)",
    ],

    "quotites_garanties": [
        r"(?:quotité garantie|quotités garanties|percentage of cover|coverage ratio)\s*[:\-–]?\s*([\d,\.]+\s*%(?:\s*[àa/-]\s*[\d,\.]+\s*%)?)",
        r"(?:taux de couverture|taux d'indemnisation|indemnification rate)\s*[:\-–]?\s*([\d,\.]+\s*%)",
        r"(?:covered at|coverage of)\s*([\d,\.]+\s*%)",
        r"(?:l'assureur couvre|guarantee up to)\s*([\d,\.]+\s*%)",
    ],

    "date_echeance": [
        r"(?:date d'échéance|échéance|expiry date|renewal date|date de renouvellement)\s*[:\-–]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        r"(?:valid until|expires on|valable jusqu'au)\s*[:\-–]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
        r"(?:au|until|to)\s+(\d{1,2}\s+(?:janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre|january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4})",
        r"(?:du\s+\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\s+au\s+)(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
    ],
}


def extract_with_regex(field: str, text: str) -> Tuple[Optional[str], float]:
    """
    Tente d'extraire un champ via regex.
    Retourne (valeur, score_confiance).
    """
    patterns = PATTERNS.get(field, [])
    text_lower = text  # garder la casse pour les noms propres

    for i, pattern in enumerate(patterns):
        try:
            match = re.search(pattern, text_lower, re.IGNORECASE | re.MULTILINE)
            if match:
                value = match.group(1).strip()
                # Score décroissant selon l'ordre des patterns (le 1er est le plus fiable)
                confidence = max(0.5, 0.95 - i * 0.1)
                return value, confidence
        except re.error:
            continue

    return None, 0.0


def extract_all_fields_regex(text: str) -> dict:
    """Extrait tous les champs avec regex."""
    results = {}
    for field in PATTERNS:
        value, confidence = extract_with_regex(field, text)
        results[field] = {
            "value": value or "Non trouvé",
            "confidence": confidence,
            "method": "regex" if value else "none"
        }
    return results
