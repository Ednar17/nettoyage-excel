"""Tests système : le programme complet, lancé dans un vrai processus, de la commande au fichier produit."""

import hashlib
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook

from conftest import creer_csv, creer_xlsx, lancer
from nettoyage_excel import cli


def empreinte(chemin: Path) -> str:
    return hashlib.sha256(chemin.read_bytes()).hexdigest()


def test_la_demo_complete_donne_les_chiffres_annonces(fichier_demo, tmp_path):
    sortie = tmp_path / "apres.xlsx"

    resultat = lancer(fichier_demo, sortie)

    assert resultat.returncode == 0, resultat.stderr
    assert resultat.stderr == ""
    assert "69 lignes lues dans exemple_avant.xlsx" in resultat.stdout
    chiffres = {l.split()[-1] for l in resultat.stdout.splitlines()[1:5]}
    assert [l.split()[-1] for l in resultat.stdout.splitlines()[1:5]] == ["52", "6", "3", "8"], chiffres
    assert load_workbook(sortie)["Données propres"].max_row == 53  # 52 lignes + l'en-tête


def test_fichier_csv_en_entree(tmp_path, valeurs_valides):
    entree = creer_csv(tmp_path / "avant.csv", [valeurs_valides])

    resultat = lancer(entree, tmp_path / "apres.xlsx")

    assert resultat.returncode == 0, resultat.stderr
    assert load_workbook(tmp_path / "apres.xlsx")["Données propres"]["B2"].value == "Camille Martin"


def test_chemins_avec_espaces_et_accents(tmp_path, valeurs_valides):
    dossier = tmp_path / "Mes fichiers clients" / "été 2025"
    dossier.mkdir(parents=True)
    entree = creer_xlsx(dossier / "commandes à nettoyer.xlsx", [valeurs_valides])
    sortie = dossier / "commandes nettoyées.xlsx"

    resultat = lancer(entree, sortie)

    assert resultat.returncode == 0, resultat.stderr
    assert sortie.is_file()


def test_chemins_relatifs_depuis_le_dossier_courant(tmp_path, valeurs_valides):
    creer_xlsx(tmp_path / "avant.xlsx", [valeurs_valides])

    resultat = lancer("avant.xlsx", "apres.xlsx", cwd=tmp_path)

    assert resultat.returncode == 0, resultat.stderr
    assert (tmp_path / "apres.xlsx").is_file()


def test_fichier_introuvable(tmp_path):
    resultat = lancer(tmp_path / "absent.xlsx", tmp_path / "apres.xlsx")

    assert resultat.returncode == 2
    assert "Erreur : Fichier introuvable" in resultat.stderr
    assert "Traceback" not in resultat.stderr
    assert not (tmp_path / "apres.xlsx").exists()


def test_colonnes_absentes(tmp_path):
    entree = creer_xlsx(tmp_path / "avant.xlsx", [["x", "y"]], entetes=["Nom", "Prix"])

    resultat = lancer(entree, tmp_path / "apres.xlsx")

    assert resultat.returncode == 2
    assert "Colonnes absentes du fichier : Date commande, Client" in resultat.stderr
    assert "Traceback" not in resultat.stderr
    assert not (tmp_path / "apres.xlsx").exists()


def test_refuse_d_ecraser_le_fichier_d_origine(tmp_path, valeurs_valides):
    entree = creer_xlsx(tmp_path / "avant.xlsx", [valeurs_valides])
    avant = empreinte(entree)

    resultat = lancer(entree, entree)

    assert resultat.returncode == 2
    assert "différent du fichier d'origine" in resultat.stderr
    assert empreinte(entree) == avant


def test_sortie_qui_n_est_pas_un_xlsx(tmp_path, valeurs_valides):
    entree = creer_xlsx(tmp_path / "avant.xlsx", [valeurs_valides])

    resultat = lancer(entree, tmp_path / "apres.csv")

    assert resultat.returncode == 2
    assert "doit se terminer par .xlsx" in resultat.stderr


def test_fichier_qui_n_est_pas_un_excel(tmp_path):
    faux = tmp_path / "faux.xlsx"
    faux.write_text("ceci n'est pas un classeur")

    resultat = lancer(faux, tmp_path / "apres.xlsx")

    assert resultat.returncode == 2
    assert "pas un fichier Excel lisible" in resultat.stderr
    assert "Traceback" not in resultat.stderr


def test_arguments_manquants():
    resultat = lancer()

    assert resultat.returncode == 2
    assert "usage: nettoyer-excel" in resultat.stderr


def test_aide():
    resultat = lancer("--help")

    assert resultat.returncode == 0
    assert "Nettoie un fichier de commandes" in resultat.stdout


def test_la_commande_installee_existe():
    dossier = Path(sys.executable).parent
    assert any((dossier / nom).exists() for nom in ("nettoyer-excel.exe", "nettoyer-excel"))


def test_fichier_de_sortie_ouvert_dans_excel(tmp_path, valeurs_valides, monkeypatch, capsys):
    """Sous Windows, Excel verrouille le fichier ouvert : l'écriture échoue avec PermissionError."""
    entree = creer_xlsx(tmp_path / "avant.xlsx", [valeurs_valides])

    def refuser(self, chemin):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(Workbook, "save", refuser)
    code = cli.main([str(entree), str(tmp_path / "apres.xlsx")])

    assert code == 2
    assert "sans doute ouvert dans Excel" in capsys.readouterr().err
