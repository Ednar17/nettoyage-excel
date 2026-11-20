"""Tests de recette : chaque promesse faite au client dans l'annonce est vérifiée sur le fichier de démo.

Chaque test porte le nom d'une phrase de l'annonce. Si l'un d'eux échoue, l'annonce ment.
"""

import hashlib
import random
import re
import socket
import time
from datetime import date, datetime, timedelta

import pytest
from openpyxl import load_workbook

from conftest import creer_csv, creer_xlsx, lancer
from nettoyage_excel import COLONNES, cli

STATUTS = {"Payé", "En attente", "Annulé"}


@pytest.fixture(scope="module")
def resultat(fichier_demo, tmp_path_factory):
    """Le classeur rendu au client pour le fichier de démo."""
    sortie = tmp_path_factory.mktemp("recette") / "apres.xlsx"
    execution = lancer(fichier_demo, sortie)
    assert execution.returncode == 0, execution.stderr
    return load_workbook(sortie), sortie


def lignes(feuille) -> list[dict]:
    rangees = list(feuille.iter_rows(values_only=True))
    return [dict(zip(rangees[0], r)) for r in rangees[1:]]


def synthese(classeur) -> dict:
    return dict(r[:2] for r in classeur["Synthèse"].iter_rows(values_only=True) if r[0])


def test_dates_montants_telephones_emails_et_noms_au_meme_format_partout(resultat):
    propres = lignes(resultat[0]["Données propres"])

    assert len(propres) == 52
    for l in propres:
        assert isinstance(l["Date commande"], datetime), l
        assert isinstance(l["Montant"], (int, float)) and l["Montant"] > 0, l
        assert re.fullmatch(r"0[1-9]( \d\d){4}", l["Téléphone"]), l
        assert l["Email"] == l["Email"].strip().lower() and "@" in l["Email"], l
        assert l["Client"] == l["Client"].strip() and "  " not in l["Client"], l
        assert l["Client"][0].isupper() and not l["Client"].isupper(), l
        assert l["Ville"][0].isupper() and not l["Ville"].isupper(), l
        assert l["Statut"] in STATUTS, l


def test_les_doublons_et_les_lignes_vides_sont_supprimes(resultat):
    propres = lignes(resultat[0]["Données propres"])
    cles = [(l["Date commande"], l["Email"], l["Montant"]) for l in propres]

    assert len(cles) == len(set(cles))
    assert all(any(v not in (None, "") for v in l.values()) for l in propres)
    assert synthese(resultat[0])["Doublons supprimés"] == 6
    assert synthese(resultat[0])["Lignes vides supprimées"] == 3


def test_les_lignes_douteuses_sont_listees_avec_la_raison(resultat):
    a_verifier = lignes(resultat[0]["À vérifier"])
    problemes = " | ".join(l["Problème"] for l in a_verifier)

    assert len(a_verifier) == 8
    for erreur_glissee in (
        "camille.martin@exemple", "lucas.bernard.exemple.fr", "Email manquant", "06 12 34",
        "inconnu", "32/13/2025", "bientôt", "à définir",
    ):
        assert erreur_glissee in problemes, erreur_glissee
    assert all(isinstance(l["Ligne d'origine"], int) and l["Ligne d'origine"] >= 2 for l in a_verifier)


def test_je_ne_supprime_et_n_invente_jamais_une_donnee(resultat, fichier_demo):
    classeur = resultat[0]
    bilan = synthese(classeur)

    # Aucune ligne ne disparaît sans être comptée.
    assert bilan["Lignes lues"] == 69
    assert bilan["Lignes lues"] == (bilan["Lignes propres"] + bilan["Doublons supprimés"]
                                    + bilan["Lignes vides supprimées"] + bilan["Lignes à vérifier"])

    # Une ligne douteuse garde ses valeurs d'origine, à son numéro de ligne d'origine.
    origine = list(load_workbook(fichier_demo).active.iter_rows(values_only=True))
    for l in lignes(classeur["À vérifier"]):
        brute = dict(zip(origine[0], origine[l["Ligne d'origine"] - 1]))
        assert l["Client"] == " ".join(str(brute["Client"]).split())
        if isinstance(brute["Email"], str):
            assert l["Email"] == brute["Email"].strip()

    # Aucun e-mail propre ne vient d'ailleurs que du fichier d'origine.
    emails_origine = {str(r[2]).strip().lower() for r in origine[1:] if r[2]}
    assert {l["Email"] for l in lignes(classeur["Données propres"])} <= emails_origine


def test_la_synthese_compte_juste(resultat):
    classeur = resultat[0]
    propres = lignes(classeur["Données propres"])
    bilan = synthese(classeur)

    assert bilan["Lignes propres"] == len(propres)
    assert bilan["Lignes à vérifier"] == len(lignes(classeur["À vérifier"]))
    for statut in STATUTS:
        attendu = round(sum(l["Montant"] for l in propres if l["Statut"] == statut), 2)
        assert bilan[statut] == pytest.approx(attendu, abs=0.01)


def test_votre_fichier_d_origine_n_est_jamais_modifie(fichier_demo, tmp_path):
    avant = hashlib.sha256(fichier_demo.read_bytes()).hexdigest()

    assert lancer(fichier_demo, tmp_path / "apres.xlsx").returncode == 0

    assert hashlib.sha256(fichier_demo.read_bytes()).hexdigest() == avant


def test_vous_pourrez_relancer_le_script_sur_vos_prochains_fichiers(resultat, tmp_path):
    """Relancé sur un fichier déjà propre, le script ne change rien et ne signale rien."""
    classeur, sortie = resultat
    seconde = tmp_path / "apres_2.xlsx"

    assert lancer(sortie, seconde).returncode == 0

    second = load_workbook(seconde)
    assert lignes(second["Données propres"]) == lignes(classeur["Données propres"])
    assert synthese(second)["Lignes à vérifier"] == 0
    assert synthese(second)["Doublons supprimés"] == 0


def test_ca_marche_avec_un_csv(fichier_demo, tmp_path):
    origine = list(load_workbook(fichier_demo).active.iter_rows(values_only=True))
    en_csv = [[v.strftime("%d/%m/%Y") if isinstance(v, (date, datetime)) else ("" if v is None else v) for v in r]
              for r in origine[1:]]
    creer_csv(tmp_path / "avant.csv", en_csv)

    assert lancer(fichier_demo, tmp_path / "xlsx.xlsx").returncode == 0
    assert lancer(tmp_path / "avant.csv", tmp_path / "csv.xlsx").returncode == 0

    depuis_xlsx = lignes(load_workbook(tmp_path / "xlsx.xlsx")["Données propres"])
    depuis_csv = lignes(load_workbook(tmp_path / "csv.xlsx")["Données propres"])
    assert depuis_csv == depuis_xlsx


def test_un_fichier_jusqu_a_5000_lignes(tmp_path):
    hasard = random.Random(1)
    debut = date(2025, 1, 1)
    rangees = [[
        (debut + timedelta(days=hasard.randrange(300))).strftime("%d/%m/%Y"),
        f"client {i}", f"client{i}@exemple.fr", f"06{hasard.randrange(10**8):08d}",
        "paris", f"{hasard.uniform(10, 5000):.2f}".replace(".", ",") + " €", "payé",
    ] for i in range(5000)]
    entree = creer_xlsx(tmp_path / "gros.xlsx", rangees)

    depart = time.perf_counter()
    execution = lancer(entree, tmp_path / "apres.xlsx")
    duree = time.perf_counter() - depart

    assert execution.returncode == 0, execution.stderr
    assert load_workbook(tmp_path / "apres.xlsx")["Données propres"].max_row == 5001
    assert duree < 30, f"{duree:.1f} s pour 5 000 lignes"


def test_vos_donnees_ne_sont_envoyees_a_aucun_service_en_ligne(fichier_demo, tmp_path, monkeypatch):
    """Le traitement complet tourne avec le réseau coupé."""
    def reseau_interdit(*args, **kwargs):
        raise AssertionError("tentative de connexion réseau pendant le nettoyage")

    monkeypatch.setattr(socket, "socket", reseau_interdit)
    monkeypatch.setattr(socket, "create_connection", reseau_interdit)

    assert cli.main([str(fichier_demo), str(tmp_path / "apres.xlsx")]) == 0


def test_le_classeur_rendu_contient_les_trois_feuilles_annoncees(resultat):
    classeur = resultat[0]

    assert classeur.sheetnames == ["Données propres", "À vérifier", "Synthèse"]
    assert [c.value for c in classeur["Données propres"][1]] == COLONNES
