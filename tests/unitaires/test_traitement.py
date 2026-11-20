"""Tests unitaires : le tri des lignes (propres, doublons, vides, à vérifier), sans aucun fichier."""

from datetime import date

from conftest import ligne, numerotees
from nettoyage_excel import COLONNES, nettoyer


def test_ligne_valide_est_normalisee():
    propres, a_verifier, bilan = nettoyer(numerotees(ligne()))

    assert a_verifier == []
    assert propres == [{
        "Date commande": date(2025, 3, 15), "Client": "Camille Martin", "Email": "camille@exemple.fr",
        "Téléphone": "06 12 34 56 78", "Ville": "Paris", "Montant": 100.0, "Statut": "Payé",
    }]
    assert bilan["lignes propres"] == 1


def test_doublon_detecte_malgre_une_presentation_differente():
    copie = ligne(Client="CAMILLE MARTIN", Email=" CAMILLE@EXEMPLE.FR", Montant=100, **{"Date commande": "2025-03-15"})
    propres, _, bilan = nettoyer(numerotees(ligne(), copie))

    assert len(propres) == 1
    assert bilan["doublons supprimés"] == 1


def test_la_premiere_occurrence_d_un_doublon_est_celle_gardee():
    propres, _, _ = nettoyer(numerotees(ligne(Ville="Paris"), ligne(Ville="Lyon")))
    assert propres[0]["Ville"] == "Paris"


def test_deux_commandes_du_meme_client_ne_sont_pas_des_doublons():
    propres, _, bilan = nettoyer(numerotees(ligne(), ligne(Montant="250,00 €")))

    assert len(propres) == 2
    assert bilan["doublons supprimés"] == 0


def test_lignes_vides_ignorees():
    vide = {c: None for c in COLONNES}
    espaces = {c: "   " for c in COLONNES}
    propres, a_verifier, bilan = nettoyer(numerotees(vide, ligne(), espaces))

    assert len(propres) == 1 and a_verifier == []
    assert bilan["lignes vides supprimées"] == 2


def test_ligne_douteuse_mise_de_cote_avec_sa_raison_et_son_numero():
    propres, a_verifier, bilan = nettoyer(numerotees(ligne(), ligne(Email="pas-un-email", Montant="à définir")))

    assert len(propres) == 1
    assert bilan["lignes à vérifier"] == 1
    assert a_verifier[0]["Ligne d'origine"] == 3
    assert "Email : « pas-un-email »" in a_verifier[0]["Problème"]
    assert "Montant : « à définir »" in a_verifier[0]["Problème"]


def test_valeur_manquante_signalee():
    _, a_verifier, _ = nettoyer(numerotees(ligne(Email=None)))
    assert a_verifier[0]["Problème"] == "Email manquant"


def test_client_ou_ville_manquant_signale():
    _, a_verifier, _ = nettoyer(numerotees(ligne(Client="  "), ligne(Ville=None)))
    assert [l["Problème"] for l in a_verifier] == ["Client manquant", "Ville manquant"]


def test_montant_a_zero_est_accepte():
    propres, a_verifier, _ = nettoyer(numerotees(ligne(Montant=0)))
    assert a_verifier == [] and propres[0]["Montant"] == 0.0


def test_un_doublon_d_une_ligne_douteuse_reste_a_verifier():
    _, a_verifier, bilan = nettoyer(numerotees(ligne(Email="x"), ligne(Email="x")))
    assert len(a_verifier) == 2 and bilan["doublons supprimés"] == 0


def test_resultat_trie_par_date_puis_par_client():
    propres, _, _ = nettoyer(numerotees(
        ligne(**{"Date commande": "20/03/2025"}),
        ligne(**{"Date commande": "01/03/2025"}, Client="zoé", Email="z@exemple.fr"),
        ligne(**{"Date commande": "01/03/2025"}, Client="adam", Email="a@exemple.fr"),
    ))
    assert [(p["Date commande"].day, p["Client"]) for p in propres] == [(1, "Adam"), (1, "Zoé"), (20, "Camille Martin")]


def test_aucune_entree_donne_un_bilan_a_zero():
    propres, a_verifier, bilan = nettoyer([])
    assert propres == [] and a_verifier == [] and bilan["lignes propres"] == 0


def test_colonne_absente_d_une_ligne_traitee_comme_manquante():
    incomplete = {"Client": "Ana Roux"}
    _, a_verifier, _ = nettoyer(numerotees(incomplete))
    assert "Email manquant" in a_verifier[0]["Problème"]
