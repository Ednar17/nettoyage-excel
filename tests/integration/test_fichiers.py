"""Tests d'intégration : lecture, tri et écriture assemblés, sur de vrais fichiers Excel et CSV."""

from datetime import datetime

import pytest
from openpyxl import Workbook, load_workbook

from conftest import creer_csv, creer_xlsx
from nettoyage_excel import COLONNES, ErreurFichier, ecrire, lire, nettoyer
from nettoyage_excel.classeur import FORMAT_DATE, FORMAT_EUROS


def traiter(entree, sortie):
    lignes = lire(entree)
    propres, a_verifier, bilan = nettoyer(lignes)
    ecrire(sortie, propres, a_verifier, bilan, len(lignes))
    return load_workbook(sortie)


def test_aller_retour_excel(tmp_path):
    entree = creer_xlsx(tmp_path / "avant.xlsx", [
        [datetime(2025, 3, 15), "camille MARTIN", "Camille@Exemple.fr", 612345678, "PARIS", "1 234,50 €", "payee"],
        [None] * 7,
        ["bientôt", "Lucas Petit", "lucas@exemple.fr", "0712345678", "Nanterre", 80, "attente"],
    ])

    classeur = traiter(entree, tmp_path / "apres.xlsx")

    assert classeur.sheetnames == ["Données propres", "À vérifier", "Synthèse"]
    donnees = list(classeur["Données propres"].iter_rows(values_only=True))
    assert donnees == [
        tuple(COLONNES),
        (datetime(2025, 3, 15), "Camille Martin", "camille@exemple.fr", "06 12 34 56 78", "Paris", 1234.5, "Payé"),
    ]

    verif = list(classeur["À vérifier"].iter_rows(values_only=True))
    assert verif[1][0] == 4  # numéro de la ligne dans le fichier d'origine
    assert verif[1][1] == "Date commande : « bientôt »"
    assert verif[1][3] == "Lucas Petit"  # valeurs d'origine conservées

    synthese = dict(r[:2] for r in classeur["Synthèse"].iter_rows(values_only=True) if r[0])
    assert synthese["Lignes lues"] == 3
    assert synthese["Lignes propres"] == 1
    assert synthese["Lignes vides supprimées"] == 1
    assert synthese["Lignes à vérifier"] == 1
    assert synthese["Doublons supprimés"] == 0
    assert synthese["Payé"] == 1234.5


def test_le_classeur_produit_est_mis_en_forme_pour_excel(tmp_path, valeurs_valides):
    entree = creer_xlsx(tmp_path / "avant.xlsx", [valeurs_valides])

    feuille = traiter(entree, tmp_path / "apres.xlsx")["Données propres"]

    assert feuille["A2"].number_format == FORMAT_DATE
    assert feuille["F2"].number_format == FORMAT_EUROS
    assert feuille["A2"].is_date and isinstance(feuille["F2"].value, (int, float))  # vraies dates, vrais nombres
    assert feuille.freeze_panes == "A2"
    assert feuille.auto_filter.ref == "A1:G2"
    assert feuille["A1"].font.bold


def test_colonnes_dans_un_autre_ordre_avec_une_colonne_en_plus(tmp_path):
    entree = creer_xlsx(
        tmp_path / "avant.xlsx",
        [["payé", 50, "Paris", "0612345678", "a@exemple.fr", "Ana Roux", "01/02/2025", "rien"]],
        entetes=["Statut", " Montant ", "Ville", "Téléphone", "Email", "Client", "Date commande", "Commentaire"],
    )

    propres, a_verifier, _ = nettoyer(lire(entree))

    assert a_verifier == []
    assert propres[0]["Client"] == "Ana Roux" and propres[0]["Montant"] == 50.0


def test_seule_la_premiere_feuille_est_lue(tmp_path, valeurs_valides):
    classeur = Workbook()
    classeur.active.append(COLONNES)
    classeur.active.append(valeurs_valides)
    classeur.create_sheet("Archives").append(["autre chose"])
    classeur.save(tmp_path / "avant.xlsx")

    assert len(lire(tmp_path / "avant.xlsx")) == 1


@pytest.mark.parametrize("separateur, encodage", [(";", "utf-8"), (",", "utf-8"), (";", "utf-8-sig"), (";", "cp1252")])
def test_lecture_csv_selon_separateur_et_encodage(tmp_path, separateur, encodage):
    montant = "1234.50" if separateur == "," else "1 234,50 €"
    entree = creer_csv(
        tmp_path / "avant.csv",
        [["15/03/2025", "Élodie Étienne", "elodie@exemple.fr", "06 12 34 56 78", "Créteil", montant, "payé"]],
        separateur=separateur, encodage=encodage,
    )

    propres, a_verifier, _ = nettoyer(lire(entree))

    assert a_verifier == []
    assert propres[0]["Client"] == "Élodie Étienne"
    assert propres[0]["Ville"] == "Créteil"
    assert propres[0]["Montant"] == 1234.5


def test_csv_avec_une_ligne_trop_courte(tmp_path):
    chemin = tmp_path / "avant.csv"
    chemin.write_text(";".join(COLONNES) + "\n15/03/2025;Ana Roux\n", encoding="utf-8")

    _, a_verifier, _ = nettoyer(lire(chemin))

    assert a_verifier[0]["Ligne d'origine"] == 2
    assert "Email manquant" in a_verifier[0]["Problème"]


def test_fichier_avec_seulement_les_entetes(tmp_path):
    entree = creer_xlsx(tmp_path / "avant.xlsx", [])

    classeur = traiter(entree, tmp_path / "apres.xlsx")

    assert classeur["Données propres"].max_row == 1
    assert classeur["À vérifier"].max_row == 1


@pytest.mark.parametrize("preparer, message", [
    (lambda d: d / "absent.xlsx", "Fichier introuvable"),
    (lambda d: creer_xlsx(d / "court.xlsx", [["01/02/2025", "Ana"]], entetes=["Date commande", "Client"]), "Colonnes absentes"),
    (lambda d: (Workbook().save(d / "vide.xlsx"), d / "vide.xlsx")[1], "Colonnes absentes"),
    (lambda d: (d.joinpath("faux.xlsx").write_text("pas un excel"), d / "faux.xlsx")[1], "pas un fichier Excel lisible"),
    (lambda d: (d.joinpath("notes.txt").write_text("x"), d / "notes.txt")[1], "Format non pris en charge"),
    (lambda d: (d.joinpath("vide.csv").write_text(""), d / "vide.csv")[1], "Colonnes absentes"),
])
def test_fichier_inutilisable_donne_un_message_clair(tmp_path, preparer, message):
    with pytest.raises(ErreurFichier, match=message):
        lire(preparer(tmp_path))


def test_message_de_colonnes_absentes_les_nomme(tmp_path):
    entree = creer_xlsx(tmp_path / "court.xlsx", [], entetes=["Date commande", "Client", "Email", "Ville", "Statut"])

    with pytest.raises(ErreurFichier, match="Téléphone, Montant"):
        lire(entree)
