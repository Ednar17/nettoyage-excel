"""Outils partagés par les tests."""

import csv
import os
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

from nettoyage_excel import COLONNES

LIGNE_VALIDE = {
    "Date commande": "15/03/2025", "Client": "camille martin", "Email": "camille@exemple.fr",
    "Téléphone": "0612345678", "Ville": "paris", "Montant": "100,00 €", "Statut": "payé",
}


def ligne(**remplacements) -> dict:
    """Une ligne valide, dont on peut remplacer des colonnes : ligne(Email="x")."""
    return {**LIGNE_VALIDE, **remplacements}


def numerotees(*lignes) -> list[tuple[int, dict]]:
    """Numérote les lignes comme dans un fichier : la première ligne de données est la ligne 2."""
    return list(enumerate(lignes, start=2))


def creer_xlsx(chemin, lignes: list[list], entetes: list[str] = COLONNES):
    classeur = Workbook()
    feuille = classeur.active
    feuille.append(entetes)
    for valeurs in lignes:
        feuille.append(valeurs)
    classeur.save(chemin)
    return chemin


def creer_csv(chemin, lignes: list[list], entetes: list[str] = COLONNES, separateur=";", encodage="utf-8"):
    with open(chemin, "w", encoding=encodage, newline="") as fichier:
        ecrivain = csv.writer(fichier, delimiter=separateur)
        ecrivain.writerow(entetes)
        ecrivain.writerows(lignes)
    return chemin


@pytest.fixture
def valeurs_valides() -> list:
    return [LIGNE_VALIDE[c] for c in COLONNES]


RACINE = Path(__file__).resolve().parent.parent
ENVIRONNEMENT = {**os.environ, "PYTHONIOENCODING": "utf-8"}


def lancer(*arguments, cwd=None) -> subprocess.CompletedProcess:
    """Lance la vraie ligne de commande, dans un processus séparé, comme le ferait un client."""
    return subprocess.run(
        [sys.executable, "-m", "nettoyage_excel", *map(str, arguments)],
        capture_output=True, text=True, encoding="utf-8", cwd=cwd, env=ENVIRONNEMENT, timeout=120,
    )


@pytest.fixture(scope="session")
def fichier_demo(tmp_path_factory) -> Path:
    """Le fichier en désordre de la démo (69 lignes, 8 erreurs glissées, 6 doublons, 3 lignes vides)."""
    dossier = tmp_path_factory.mktemp("demo")
    subprocess.run(
        [sys.executable, str(RACINE / "exemples" / "creer_exemple.py")],
        cwd=dossier, check=True, capture_output=True, env=ENVIRONNEMENT, timeout=120,
    )
    return dossier / "exemple_avant.xlsx"
