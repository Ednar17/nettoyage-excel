"""Fabrique un fichier Excel volontairement en désordre, avec des données fictives.

Usage : python creer_exemple.py
Produit : exemple_avant.xlsx
"""

import random
from datetime import date, timedelta

from openpyxl import Workbook

random.seed(7)

PRENOMS = ["Camille", "Lucas", "Inès", "Hugo", "Léa", "Nathan", "Chloé", "Yanis", "Manon", "Théo", "Sarah", "Jules"]
NOMS = ["Martin", "Bernard", "Petit", "Durand", "Leroy", "Moreau", "Simon", "Laurent", "Lefebvre", "Roux", "Fournier", "Girard"]
VILLES = ["Nanterre", "Paris", "Versailles", "Saint-Denis", "Créteil", "Boulogne-Billancourt", "Montreuil", "Argenteuil"]
STATUTS = {
    "Payé": ["Payé", "payé", "PAYE", "Payee", "payée "],
    "En attente": ["En attente", "en attente", "attente", "EN ATTENTE"],
    "Annulé": ["Annulé", "annule", "ANNULÉ", "Annulée"],
}
MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]


def date_en_desordre(jour: date):
    """Renvoie la même date sous un format différent à chaque fois."""
    forme = random.choice(["excel", "fr", "iso", "tirets", "texte"])
    if forme == "excel":
        return jour
    if forme == "fr":
        return jour.strftime("%d/%m/%Y")
    if forme == "iso":
        return jour.strftime("%Y-%m-%d")
    if forme == "tirets":
        return jour.strftime("%d-%m-%y")
    return f"{jour.day} {MOIS[jour.month - 1]} {jour.year}"


def telephone_en_desordre(chiffres: str) -> str:
    forme = random.choice(["colle", "espaces", "points", "plus33", "tirets"])
    if forme == "colle":
        return chiffres
    paires = [chiffres[i : i + 2] for i in range(0, 10, 2)]
    if forme == "espaces":
        return " ".join(paires)
    if forme == "points":
        return ".".join(paires)
    if forme == "tirets":
        return "-".join(paires)
    return "+33" + chiffres[1:]


def montant_en_desordre(valeur: float):
    forme = random.choice(["nombre", "virgule", "euro", "espace"])
    if forme == "nombre":
        return valeur
    texte = f"{valeur:.2f}".replace(".", ",")
    if forme == "virgule":
        return texte
    if forme == "euro":
        return texte + " €"
    entier, decimales = texte.split(",")
    if len(entier) > 3:
        entier = entier[:-3] + " " + entier[-3:]
    return f"{entier},{decimales} EUR"


def nom_en_desordre(prenom: str, nom: str) -> str:
    forme = random.choice(["propre", "majuscules", "minuscules", "espaces"])
    complet = f"{prenom} {nom}"
    if forme == "majuscules":
        return complet.upper()
    if forme == "minuscules":
        return complet.lower()
    if forme == "espaces":
        return f"  {prenom}   {nom} "
    return complet


def main() -> None:
    lignes = []
    debut = date(2025, 1, 6)
    for i in range(60):
        prenom, nom = random.choice(PRENOMS), random.choice(NOMS)
        email = f"{prenom}.{nom}@exemple.fr".lower().replace("é", "e").replace("è", "e").replace("ï", "i")
        telephone = "0" + random.choice("67") + "".join(random.choice("0123456789") for _ in range(8))
        statut = random.choice(list(STATUTS))
        lignes.append(
            [
                date_en_desordre(debut + timedelta(days=random.randint(0, 300))),
                nom_en_desordre(prenom, nom),
                random.choice([email, email.upper(), " " + email]),
                telephone_en_desordre(telephone),
                random.choice([v := random.choice(VILLES), v.upper(), v.lower()]),
                montant_en_desordre(round(random.uniform(35, 2400), 2)),
                random.choice(STATUTS[statut]),
            ]
        )

    # Problèmes ajoutés à la main : ce sont eux que le script doit repérer.
    lignes[4][2] = "camille.martin@exemple"  # e-mail sans extension
    lignes[11][2] = "lucas.bernard.exemple.fr"  # e-mail sans @
    lignes[19][2] = ""  # e-mail manquant
    lignes[23][3] = "06 12 34"  # téléphone incomplet
    lignes[31][3] = "inconnu"  # téléphone non renseigné
    lignes[37][0] = "32/13/2025"  # date impossible
    lignes[44][0] = "bientôt"  # date illisible
    lignes[52][5] = "à définir"  # montant illisible

    # Doublons : même commande saisie deux fois, avec une présentation différente.
    for source in (2, 9, 15, 28, 40, 55):
        copie = list(lignes[source])
        copie[1] = str(copie[1]).upper()
        copie[2] = "  " + str(copie[2]).upper()
        lignes.append(copie)

    random.shuffle(lignes)
    for position in (10, 25, 48):  # lignes vides au milieu du fichier
        lignes.insert(position, [None] * 7)

    classeur = Workbook()
    feuille = classeur.active
    feuille.title = "Commandes"
    feuille.append(["Date commande", "Client", "Email", "Téléphone", "Ville", "Montant", "Statut"])
    for ligne in lignes:
        feuille.append(ligne)
    classeur.save("exemple_avant.xlsx")
    print(f"exemple_avant.xlsx créé : {len(lignes)} lignes")


if __name__ == "__main__":
    main()
