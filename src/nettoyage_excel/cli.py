"""Ligne de commande : nettoyer-excel <entree> <sortie.xlsx>"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .classeur import LIBELLES_BILAN, ErreurFichier, ecrire, lire
from .traitement import nettoyer


def main(argv: list[str] | None = None) -> int:
    """Renvoie 0 si le fichier a été traité, 2 si le fichier ou les arguments posent problème."""
    parser = argparse.ArgumentParser(
        prog="nettoyer-excel",
        description="Nettoie un fichier de commandes (Excel ou CSV) et produit un classeur Excel prêt à l'emploi.",
    )
    parser.add_argument("entree", type=Path, help="fichier à nettoyer (.xlsx, .xlsm ou .csv)")
    parser.add_argument("sortie", type=Path, help="classeur à créer (.xlsx)")
    args = parser.parse_args(argv)

    try:
        if args.sortie.suffix.lower() != ".xlsx":
            raise ErreurFichier("Le fichier de sortie doit se terminer par .xlsx")
        if args.sortie.resolve() == args.entree.resolve():
            raise ErreurFichier("Le fichier de sortie doit être différent du fichier d'origine, qui n'est jamais modifié.")

        lignes = lire(args.entree)
        propres, a_verifier, bilan = nettoyer(lignes)
        ecrire(args.sortie, propres, a_verifier, bilan, len(lignes))
    except ErreurFichier as erreur:
        print(f"Erreur : {erreur}", file=sys.stderr)
        return 2

    print(f"{len(lignes)} lignes lues dans {args.entree.name}")
    for libelle in LIBELLES_BILAN:
        print(f"  {libelle:<24} {bilan[libelle]:>5}")
    print(f"Résultat : {args.sortie}")
    return 0
