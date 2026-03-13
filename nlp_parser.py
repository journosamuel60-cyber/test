"""
nlp_parser.py
Analyse NLP des contrats via l'API Claude.
Utilisé comme couche sémantique au-dessus des regex.
"""

import json
import logging
import os
import re
from typing import Optional

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# CHAMPS EXTRAITS
# ─────────────────────────────────────────────

SCALAR_FIELDS = [
    "numero_police", "assure", "assureur", "courtier",
    "taux_prime", "prime_provisionnelle", "quotites_garanties",
    "delai_indemnisation", "delai_max_credit",
    "date_prise_effet", "date_echeance", "duree_police",
    "devise", "limite_decaissement", "zone_discretionnaire",
]

STRUCTURED_FIELDS = ["groupe_polices"]
ALL_FIELDS = SCALAR_FIELDS + STRUCTURED_FIELDS


# ─────────────────────────────────────────────
# PROMPT SYSTÈME
# ─────────────────────────────────────────────

SYSTEM_PROMPT = """Tu es un expert en analyse de contrats d'assurance-crédit (credit insurance).
Tu parles couramment le français et l'anglais.

Ta mission : extraire des informations contractuelles précises depuis le texte fourni.
Le document peut être une police Atradius format Modula, Coface, Euler Hermes, ou autre assureur-crédit.

Réponds UNIQUEMENT avec un objet JSON valide, sans texte avant ou après.
Si une information est absente ou ambiguë, utilise la valeur "Non trouvé".

Format de réponse obligatoire :
{
  "numero_police": "Numéro de la police (ex: 1453314 USD)",
  "assure": "Nom de l'assuré / souscripteur",
  "assureur": "Nom de l'assureur / compagnie",
  "courtier": "Nom du courtier / broker (si présent, sinon Non trouvé)",
  "taux_prime": "Taux de prime (ex: 0,052% ou 1.2‰)",
  "prime_provisionnelle": "Prime provisionnelle totale (ex: 4.510 USD ou 12 000 EUR)",
  "quotites_garanties": "Pourcentage assuré / couverture (ex: 95%)",
  "delai_indemnisation": "Délai d'indemnisation (ex: 6 mois)",
  "delai_max_credit": "Délai maximum de crédit consenti (ex: 180 jours)",
  "date_prise_effet": "Date de prise d'effet / début de la police",
  "date_echeance": "Date d'échéance / fin de la police",
  "duree_police": "Durée de la police (ex: 13 mois)",
  "devise": "Devise de la police (ex: Dollar US, EUR, CAD)",
  "limite_decaissement": "Plafond total d'indemnisation (ex: CAD 2.000.000)",
  "zone_discretionnaire": "Limite discrétionnaire / Credit Check (ex: 23.000)"
}

Termes équivalents :
- Numéro police = policy number, police N°
- Assuré = policyholder, insured, souscripteur
- Assureur = insurer, underwriter, compagnie d'assurance
- Courtier = broker, intermédiaire
- Taux de prime = prime rate, taux d'assurance
- Prime provisionnelle = provisional premium, prime totale, acompte
- Pourcentage assuré = coverage ratio, quotités garanties, taux d'indemnisation
- Délai d'indemnisation = indemnification period, waiting period
- Délai maximum de crédit = maximum credit term
- Date de prise d'effet = effective date, inception date
- Date d'échéance = expiry date, renewal date, fin de police
- Durée de la police = policy term
- Limite de décaissement = aggregate limit, AIL, maximum d'indemnité
- Zone discrétionnaire = non-dénommé, discretionary limit, BDL, Credit Check
"""


# ─────────────────────────────────────────────
# EXTRACTION VIA LLM
# ─────────────────────────────────────────────

def extract_with_llm(text: str, api_key: Optional[str] = None) -> dict:
    try:
        import anthropic

        max_chars = 15000
        if len(text) > max_chars:
            text_truncated = text[:8000] + "\n\n[...]\n\n" + text[-7000:]
            logger.info(f"Texte tronqué: {len(text)} → ~{max_chars} caractères")
        else:
            text_truncated = text

        client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1200,
            system=SYSTEM_PROMPT,
            messages=[{
                "role": "user",
                "content": f"Analyse ce contrat d'assurance-crédit et extrait les informations demandées :\n\n{text_truncated}"
            }]
        )

        response_text = message.content[0].text.strip()
        logger.debug(f"Réponse LLM: {response_text[:500]}")
        return _parse_llm_response(response_text)

    except ImportError:
        logger.warning("anthropic non installé, LLM désactivé")
        return {}
    except Exception as e:
        logger.error(f"Erreur LLM: {e}")
        return {}


def _parse_llm_response(response_text: str) -> dict:
    cleaned = re.sub(r"```(?:json)?", "", response_text).strip().rstrip("```").strip()

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        logger.warning("JSON invalide, tentative d'extraction partielle")
        data = {}
        for field in SCALAR_FIELDS:
            match = re.search(rf'"{field}"\s*:\s*"([^"]*)"', cleaned)
            if match:
                data[field] = match.group(1)

    result = {}
    for field, value in data.items():
        is_found = value not in ["Non trouvé", "", None, "N/A", "non trouvé"]
        result[field] = {
            "value": value if is_found else "Non trouvé",
            "confidence": 0.85 if is_found else 0.0,
            "method": "llm" if is_found else "none"
        }
    return result


# ─────────────────────────────────────────────
# FUSION REGEX + LLM
# ─────────────────────────────────────────────

def merge_results(regex_results: dict, llm_results: dict) -> dict:
    """
    Fusionne regex et LLM.
    - Champs scalaires : meilleure confiance gagne.
    - groupe_polices : toujours depuis regex (structure liste).
    """
    merged = {}

    for field in SCALAR_FIELDS:
        regex = regex_results.get(field, {"value": None, "confidence": 0.0})
        llm   = llm_results.get(field,   {"value": None, "confidence": 0.0})

        regex_found = regex.get("value") not in [None, "Non trouvé", ""]
        llm_found   = llm.get("value")   not in [None, "Non trouvé", ""]

        if regex_found and llm_found:
            if regex.get("confidence", 0) >= llm.get("confidence", 0):
                merged[field] = {**regex, "method": "regex+llm"}
            else:
                merged[field] = {**llm, "method": "llm+regex"}
        elif regex_found:
            merged[field] = {**regex, "method": "regex"}
        elif llm_found:
            merged[field] = {**llm, "method": "llm"}
        else:
            merged[field] = {"value": "Non trouvé", "confidence": 0.0, "method": "none"}

    # groupe_polices — liste structurée, regex uniquement
    gp = regex_results.get("groupe_polices", {"value": None, "confidence": 0.0})
    merged["groupe_polices"] = {**gp, "method": "regex" if gp.get("value") else "none"}

    return merged
