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
# PROMPT SYSTÈME
# ─────────────────────────────────────────────

SYSTEM_PROMPT = """Tu es un expert en analyse de contrats d'assurance-crédit (credit insurance).
Tu parles couramment le français et l'anglais.

Ta mission : extraire des informations contractuelles précises depuis le texte fourni.

Réponds UNIQUEMENT avec un objet JSON valide, sans texte avant ou après.
Si une information est absente ou ambiguë, utilise la valeur "Non trouvé".

Format de réponse obligatoire :
{
  "assure": "Nom de l'assuré / souscripteur",
  "assureur": "Nom de l'assureur / compagnie",
  "taux_prime": "Taux de prime (ex: 0.12% ou 1.2‰)",
  "minimum_prime": "Prime minimum annuelle (ex: 5 000 EUR)",
  "limite_decaissement": "Plafond total d'indemnisation (ex: 500 000 EUR)",
  "zone_discretionnaire": "Limite discrétionnaire / non-dénommés (ex: 30 000 EUR)",
  "quotites_garanties": "Pourcentage de couverture (ex: 85% ou 80%-90%)",
  "date_echeance": "Date d'échéance du contrat (ex: 31/12/2025)"
}

Termes équivalents à reconnaître :
- Assuré = policyholder, insured, souscripteur
- Assureur = insurer, underwriter, compagnie d'assurance
- Taux de prime = prime rate, taux d'assurance, insurance rate
- Minimum de prime = minimum premium, prime plancher, MAP
- Limite de décaissement = aggregate limit, AIL, plafond d'indemnisation
- Zone discrétionnaire = non-dénommé, discretionary limit, non-named buyers, BDL
- Quotités garanties = coverage ratio, percentage of cover, taux d'indemnisation
- Date d'échéance = expiry date, renewal date, date de renouvellement
"""


# ─────────────────────────────────────────────
# EXTRACTION VIA LLM
# ─────────────────────────────────────────────

def extract_with_llm(text: str, api_key: Optional[str] = None) -> dict:
    """
    Envoie le texte à Claude pour extraction sémantique.
    Retourne un dict avec les champs extraits et scores de confiance.
    """
    try:
        import anthropic

        # Tronquer le texte si trop long (max ~6000 tokens de contexte utile)
        max_chars = 15000
        if len(text) > max_chars:
            # Prendre début + fin (les infos importantes sont souvent aux deux extrémités)
            text_truncated = text[:8000] + "\n\n[...]\n\n" + text[-7000:]
            logger.info(f"Texte tronqué: {len(text)} → ~{max_chars} caractères")
        else:
            text_truncated = text

        client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1000,
            system=SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": f"Analyse ce contrat d'assurance-crédit et extrait les informations demandées :\n\n{text_truncated}"
                }
            ]
        )

        response_text = message.content[0].text.strip()
        logger.debug(f"Réponse LLM: {response_text[:500]}")

        # Parser le JSON retourné
        parsed = _parse_llm_response(response_text)
        return parsed

    except ImportError:
        logger.warning("anthropic non installé, LLM désactivé")
        return {}
    except Exception as e:
        logger.error(f"Erreur LLM: {e}")
        return {}


def _parse_llm_response(response_text: str) -> dict:
    """Parse la réponse JSON du LLM avec gestion des erreurs."""
    # Nettoyer les éventuels backticks markdown
    cleaned = re.sub(r"```(?:json)?", "", response_text).strip()
    cleaned = cleaned.rstrip("```").strip()

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        # Tentative d'extraction partielle
        logger.warning("JSON invalide, tentative d'extraction partielle")
        data = {}
        for field in ["assure", "assureur", "taux_prime", "minimum_prime",
                      "limite_decaissement", "zone_discretionnaire",
                      "quotites_garanties", "date_echeance"]:
            pattern = rf'"{field}"\s*:\s*"([^"]*)"'
            match = re.search(pattern, cleaned)
            if match:
                data[field] = match.group(1)

    # Ajouter scores de confiance
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
    Fusionne les résultats regex et LLM.
    Stratégie : prendre le résultat avec la meilleure confiance.
    """
    fields = [
        "assure", "assureur", "taux_prime", "minimum_prime",
        "limite_decaissement", "zone_discretionnaire",
        "quotites_garanties", "date_echeance"
    ]

    merged = {}
    for field in fields:
        regex = regex_results.get(field, {"value": "Non trouvé", "confidence": 0.0, "method": "none"})
        llm = llm_results.get(field, {"value": "Non trouvé", "confidence": 0.0, "method": "none"})

        regex_found = regex["value"] != "Non trouvé"
        llm_found = llm["value"] != "Non trouvé"

        if regex_found and llm_found:
            # Les deux trouvent quelque chose : prendre le plus confiant
            if regex["confidence"] >= llm["confidence"]:
                merged[field] = {**regex, "method": "regex+llm"}
            else:
                merged[field] = {**llm, "method": "llm+regex"}
        elif regex_found:
            merged[field] = {**regex, "method": "regex"}
        elif llm_found:
            merged[field] = {**llm, "method": "llm"}
        else:
            merged[field] = {"value": "Non trouvé", "confidence": 0.0, "method": "none"}

    return merged
