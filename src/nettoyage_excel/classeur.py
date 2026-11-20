"""Lecture du fichier du client (Excel ou CSV) et écriture du classeur de résultat."""

from __future__ import annotations

import csv
import io
import zipfile
from collections import Counter
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from .regles import texte
from .traitement import COLONNES

FORMAT_DATE = "DD/MM/YYYY"
FORMAT_EUROS = '#,##0.00 "€"'
LIBELLES_BILAN = ("lignes propres", "doublons supprimés", "lignes vides supprimées", "lignes à vérifier")


class ErreurFichier(Exception):
    """Le fichier du client ne peut pas être traité ; le message est fait pour lui être montré."""


def _rangees_excel(chemin: Path) -> list[tuple]:
    try:
        classeur = load_workbook(chemin, data_only=True, read_only=True)
    except (InvalidFileException, zipfile.BadZipFile, KeyError, OSError, ValueError) as erreur:
        raise ErreurFichier(f"Impossible d'ouvrir {chemin.name} : ce n'est pas un fichier Excel lisible.") from erreur
    try:
        return list(classeur.active.iter_rows(values_only=True))
    finally:
        classeur.close()


def _rangees_csv(chemin: Path) -> list[tuple]:
    brut = chemin.read_bytes()
    for encodage in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            contenu = brut.decode(encodage)
            break
        except UnicodeDecodeError:
            continue
    premiere = contenu.splitlines()[0] if contenu.strip() else ""
    separateur = ";" if premiere.count(";") > premiere.count(",") else ","
    return [tuple(r) for r in csv.reader(io.StringIO(contenu), delimiter=separateur)]


def lire(entree: str | Path) -> list[tuple[int, dict]]:
    """Renvoie les lignes du fichier sous la forme (numéro de ligne dans le fichier, {colonne: valeur})."""
    chemin = Path(entree)
    if not chemin.is_file():
        raise ErreurFichier(f"Fichier introuvable : {chemin}")

    suffixe = chemin.suffix.lower()
    if suffixe == ".csv":
        rangees = _rangees_csv(chemin)
    elif suffixe in (".xlsx", ".xlsm"):
        rangees = _rangees_excel(chemin)
    else:
        raise ErreurFichier(f"Format non pris en charge : {suffixe or chemin.name}. Formats acceptés : .xlsx, .xlsm, .csv")

    entetes = [texte(c) for c in rangees[0]] if rangees else []
    manquantes = [c for c in COLONNES if c not in entetes]
    if manquantes:
        raise ErreurFichier(f"Colonnes absentes du fichier : {', '.join(manquantes)}")

    lignes = []
    for numero, rangee in enumerate(rangees[1:], start=2):
        ligne = {entete: (rangee[i] if i < len(rangee) else None) for i, entete in enumerate(entetes) if entete}
        lignes.append((numero, ligne))
    return lignes


def _mettre_en_forme(feuille, largeurs: list[int], filtre: bool = True) -> None:
    for cellule in feuille[1]:
        cellule.font = Font(bold=True, color="FFFFFF")
        cellule.fill = PatternFill("solid", fgColor="1F4E78")
        cellule.alignment = Alignment(vertical="center")
    for index, largeur in enumerate(largeurs, start=1):
        feuille.column_dimensions[get_column_letter(index)].width = largeur
    feuille.freeze_panes = "A2"
    if filtre and feuille.max_row > 1:
        feuille.auto_filter.ref = feuille.dimensions


def ecrire(sortie: str | Path, propres: list[dict], a_verifier: list[dict], bilan: Counter, lues: int) -> None:
    """Écrit le classeur de résultat : « Données propres », « À vérifier », « Synthèse »."""
    classeur = Workbook()

    feuille = classeur.active
    feuille.title = "Données propres"
    feuille.append(COLONNES)
    for ligne in propres:
        feuille.append([ligne[c] for c in COLONNES])
    for rang in range(2, feuille.max_row + 1):
        feuille.cell(rang, 1).number_format = FORMAT_DATE
        feuille.cell(rang, 6).number_format = FORMAT_EUROS
    _mettre_en_forme(feuille, [15, 24, 32, 17, 24, 14, 13])

    feuille = classeur.create_sheet("À vérifier")
    entetes = ["Ligne d'origine", "Problème"] + COLONNES
    feuille.append(entetes)
    for ligne in a_verifier:
        feuille.append([ligne[c] for c in entetes])
    _mettre_en_forme(feuille, [15, 42, 15, 24, 32, 17, 24, 14, 13])

    feuille = classeur.create_sheet("Synthèse")
    feuille.append(["Indicateur", "Valeur"])
    feuille.append(["Lignes lues", lues])
    for libelle in LIBELLES_BILAN:
        feuille.append([libelle.capitalize(), bilan[libelle]])
    feuille.append([])
    feuille.append(["Chiffre d'affaires par statut", ""])
    feuille.cell(feuille.max_row, 1).font = Font(bold=True)
    totaux = Counter()
    for ligne in propres:
        totaux[ligne["Statut"]] += ligne["Montant"]
    for statut, total in sorted(totaux.items()):
        feuille.append([statut, round(total, 2)])
        feuille.cell(feuille.max_row, 2).number_format = FORMAT_EUROS
    _mettre_en_forme(feuille, [32, 16], filtre=False)

    try:
        classeur.save(sortie)
    except PermissionError as erreur:
        raise ErreurFichier(
            f"Impossible d'écrire {Path(sortie).name} : le fichier est sans doute ouvert dans Excel. Fermez-le et relancez."
        ) from erreur
