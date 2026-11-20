"""Tri des lignes : propres, doublons, vides, à vérifier."""

from __future__ import annotations

from collections import Counter

from .regles import (
    affichage, lire_date, lire_email, lire_montant, lire_statut, lire_telephone, nom_propre, texte,
)

COLONNES = ["Date commande", "Client", "Email", "Téléphone", "Ville", "Montant", "Statut"]

LECTEURS = {
    "Date commande": lire_date,
    "Client": nom_propre,
    "Email": lire_email,
    "Téléphone": lire_telephone,
    "Ville": nom_propre,
    "Montant": lire_montant,
    "Statut": lire_statut,
}


def nettoyer(lignes: list[tuple[int, dict]]) -> tuple[list[dict], list[dict], Counter]:
    """Prend des couples (numéro de ligne, ligne brute).

    Renvoie les lignes propres, les lignes à vérifier (valeurs d'origine conservées,
    avec la raison) et le décompte de ce qui a été fait. Aucune ligne n'est perdue :
    propres + à vérifier + doublons + vides = lignes reçues.
    """
    propres, a_verifier, bilan = [], [], Counter()
    deja_vues = set()

    for numero, ligne in lignes:
        if not any(texte(ligne.get(c)) for c in COLONNES):
            bilan["lignes vides supprimées"] += 1
            continue

        nettoyee = {colonne: lecteur(ligne.get(colonne)) for colonne, lecteur in LECTEURS.items()}

        problemes = [
            f"{colonne} : « {affichage(ligne.get(colonne))} »" if texte(ligne.get(colonne)) else f"{colonne} manquant"
            for colonne, valeur in nettoyee.items()
            if valeur is None or valeur == ""
        ]
        if problemes:
            a_verifier.append({
                **{c: affichage(ligne.get(c)) for c in COLONNES},
                "Ligne d'origine": numero,
                "Problème": " ; ".join(problemes),
            })
            bilan["lignes à vérifier"] += 1
            continue

        cle = (nettoyee["Date commande"], nettoyee["Email"], nettoyee["Montant"])
        if cle in deja_vues:
            bilan["doublons supprimés"] += 1
            continue
        deja_vues.add(cle)
        propres.append(nettoyee)

    propres.sort(key=lambda l: (l["Date commande"], l["Client"]))
    bilan["lignes propres"] = len(propres)
    return propres, a_verifier, bilan
