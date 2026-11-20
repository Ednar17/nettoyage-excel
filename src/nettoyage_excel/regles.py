"""Règles de lecture d'une cellule : chaque fonction renvoie la valeur propre, ou None si elle est inutilisable.

Aucune règle ne devine : une valeur ambiguë ou illisible est refusée, pour être montrée à un humain.
"""

from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime

MOIS = {
    "janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
    "juillet": 7, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12,
}
STATUTS = {
    "paye": "Payé", "payee": "Payé",
    "attente": "En attente", "en attente": "En attente",
    "annule": "Annulé", "annulee": "Annulé",
}
FORMATS_DATE = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y", "%d-%m-%Y", "%d/%m/%y")
EMAIL = re.compile(r"^[a-z0-9._%+-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}$")


def sans_accents(valeur: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", valeur) if unicodedata.category(c) != "Mn")


def texte(valeur) -> str:
    """Convertit une cellule en texte, sans espaces en trop (y compris les espaces insécables)."""
    return re.sub(r"\s+", " ", str(valeur)).strip() if valeur is not None else ""


def affichage(valeur) -> str:
    """Comme `texte`, mais une date Excel est montrée au format français."""
    return valeur.strftime("%d/%m/%Y") if isinstance(valeur, (date, datetime)) else texte(valeur)


def nom_propre(valeur) -> str:
    """« jean-PIERRE   dupont » -> « Jean-Pierre Dupont »."""
    return re.sub(r"[^\s-]+", lambda m: m.group().capitalize(), texte(valeur))


def lire_date(valeur) -> date | None:
    if isinstance(valeur, datetime):
        return valeur.date()
    if isinstance(valeur, date):
        return valeur
    brut = texte(valeur)
    for forme in FORMATS_DATE:
        try:
            return datetime.strptime(brut, forme).date()
        except ValueError:
            pass
    en_lettres = re.fullmatch(r"(\d{1,2}) (\w+) (\d{4})", sans_accents(brut.lower()))
    if en_lettres and en_lettres.group(2) in MOIS:
        try:
            return date(int(en_lettres.group(3)), MOIS[en_lettres.group(2)], int(en_lettres.group(1)))
        except ValueError:
            return None
    return None


def lire_montant(valeur) -> float | None:
    """Accepte « 1 234,50 € », « 1.234,50 », « 1,234.50 », 1234.5... et renvoie 1234.5."""
    if isinstance(valeur, bool):
        return None
    if isinstance(valeur, (int, float)):
        return round(float(valeur), 2)

    # La devise n'est acceptée qu'au début ou à la fin : « 12 € 50 » est ambigu, donc refusé.
    brut = re.sub(r"\s", "", texte(valeur))
    brut = re.sub(r"(?i)^(€|eur)|(€|eur)$", "", brut)
    if not re.fullmatch(r"-?[\d.,]+", brut):
        return None

    if "," in brut and "." in brut:
        # Le dernier des deux signes est la virgule décimale, l'autre sépare les milliers.
        decimal, milliers = (",", ".") if brut.rfind(",") > brut.rfind(".") else (".", ",")
        brut = brut.replace(milliers, "").replace(decimal, ".")
    elif brut.count(",") == 1:
        brut = brut.replace(",", ".")
    elif brut.count(",") > 1 or brut.count(".") > 1:
        if not re.fullmatch(r"-?\d{1,3}([.,]\d{3})+", brut):
            return None
        brut = brut.replace(",", "").replace(".", "")

    try:
        return round(float(brut), 2)
    except ValueError:
        return None


def lire_telephone(valeur) -> str | None:
    """Renvoie un numéro français au format « 06 12 34 56 78 », ou None."""
    if isinstance(valeur, bool):
        return None
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)  # Excel a stocké le numéro comme un nombre
    brut = texte(valeur).replace("(0)", "")
    chiffres = re.sub(r"\D", "", brut)

    if chiffres.startswith("0033"):
        chiffres = "0" + chiffres[4:]
    elif brut.startswith("+33") or (chiffres.startswith("33") and len(chiffres) == 11):
        chiffres = "0" + chiffres[2:]
    elif len(chiffres) == 9 and chiffres[0] != "0":
        chiffres = "0" + chiffres  # zéro initial perdu par Excel

    if not re.fullmatch(r"0[1-9]\d{8}", chiffres):
        return None
    return " ".join(chiffres[i : i + 2] for i in range(0, 10, 2))


def lire_email(valeur) -> str | None:
    adresse = texte(valeur).lower()
    return adresse if EMAIL.match(adresse) else None


def lire_statut(valeur) -> str | None:
    return STATUTS.get(sans_accents(texte(valeur).lower()))
