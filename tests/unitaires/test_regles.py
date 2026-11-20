"""Tests unitaires : chaque règle de lecture d'une cellule, isolément."""

from datetime import date, datetime

import pytest

from nettoyage_excel.regles import (
    affichage, lire_date, lire_email, lire_montant, lire_statut, lire_telephone, nom_propre, texte,
)


@pytest.mark.parametrize("brut, attendu", [
    ("15/03/2025", date(2025, 3, 15)),
    ("2025-03-15", date(2025, 3, 15)),
    ("15-03-25", date(2025, 3, 15)),
    ("15-03-2025", date(2025, 3, 15)),
    ("15/03/25", date(2025, 3, 15)),
    ("15 mars 2025", date(2025, 3, 15)),
    ("1 août 2025", date(2025, 8, 1)),
    ("3 Février 2025", date(2025, 2, 3)),
    ("  15/03/2025  ", date(2025, 3, 15)),
    ("29/02/2024", date(2024, 2, 29)),  # année bissextile
    (datetime(2025, 3, 15, 10, 30), date(2025, 3, 15)),
    (date(2025, 3, 15), date(2025, 3, 15)),
])
def test_dates_reconnues(brut, attendu):
    assert lire_date(brut) == attendu


@pytest.mark.parametrize("brut", [
    "32/13/2025", "30/02/2025", "29/02/2025", "31 février 2025", "bientôt", "", None, "15 truc 2025", 45000,
])
def test_dates_refusees(brut):
    assert lire_date(brut) is None


@pytest.mark.parametrize("brut, attendu", [
    (1234.5, 1234.5),
    (80, 80.0),
    (0, 0.0),
    (19.999, 20.0),
    ("1234,50", 1234.5),
    ("1234.50", 1234.5),
    ("1 234,50 €", 1234.5),
    ("1 234,50 EUR", 1234.5),
    ("1 234,50 €", 1234.5),  # espaces insécables, fréquents dans Excel
    ("1 234,50", 1234.5),  # espace fine insécable
    ("1.234,50", 1234.5),  # milliers à l'européenne
    ("1,234.50", 1234.5),  # milliers à l'anglaise
    ("1.234.567", 1234567.0),
    ("12 345 678,90 €", 12345678.9),
    ("-45,00", -45.0),
    ("€ 99", 99.0),
    ("99 eur", 99.0),
])
def test_montants_reconnus(brut, attendu):
    assert lire_montant(brut) == attendu


@pytest.mark.parametrize("brut", ["à définir", "", None, "douze", "12,5,3", "1.2.3", True, "12 € 50", "--5"])
def test_montants_refuses(brut):
    assert lire_montant(brut) is None


@pytest.mark.parametrize("brut", [
    "0612345678",
    "06 12 34 56 78",
    "06.12.34.56.78",
    "06-12-34-56-78",
    "+33612345678",
    "+33 6 12 34 56 78",
    "0033612345678",
    "+33 (0)6 12 34 56 78",
    "33612345678",
    612345678,  # Excel a transformé le numéro en nombre : le zéro a disparu
    612345678.0,
    "612345678",
])
def test_telephones_reconnus(brut):
    assert lire_telephone(brut) == "06 12 34 56 78"


@pytest.mark.parametrize("brut", [
    "06 12 34", "inconnu", "", None, "06123456789", "1234567890", "+44 20 7946 0958", "0012345678", True, 6.5,
])
def test_telephones_refuses(brut):
    assert lire_telephone(brut) is None


@pytest.mark.parametrize("brut, attendu", [
    ("Camille.Martin@Exemple.FR", "camille.martin@exemple.fr"),
    ("  camille@exemple.fr ", "camille@exemple.fr"),
    ("c.martin+pro@mail.exemple.co.uk", "c.martin+pro@mail.exemple.co.uk"),
])
def test_emails_reconnus(brut, attendu):
    assert lire_email(brut) == attendu


@pytest.mark.parametrize("brut", [
    "camille@exemple", "camille.exemple.fr", "", None, "a@b@c.fr", "camille @exemple.fr", "@exemple.fr", "camille@.fr",
])
def test_emails_refuses(brut):
    assert lire_email(brut) is None


@pytest.mark.parametrize("brut, attendu", [
    ("PAYE", "Payé"), ("payée ", "Payé"), ("Payee", "Payé"), ("Payé", "Payé"),
    ("attente", "En attente"), ("EN ATTENTE", "En attente"), ("En  attente", "En attente"),
    ("annule", "Annulé"), ("ANNULÉE", "Annulé"),
])
def test_statuts_reconnus(brut, attendu):
    assert lire_statut(brut) == attendu


@pytest.mark.parametrize("brut", ["", None, "remboursé", "?", "payé ?"])
def test_statuts_refuses(brut):
    assert lire_statut(brut) is None


@pytest.mark.parametrize("brut, attendu", [
    ("  yanis   LEFEBVRE ", "Yanis Lefebvre"),
    ("jean-pierre dupont", "Jean-Pierre Dupont"),
    ("ÉLODIE ÉTIENNE", "Élodie Étienne"),
    ("boulogne-billancourt", "Boulogne-Billancourt"),
    (None, ""),
    ("", ""),
])
def test_noms_propres(brut, attendu):
    assert nom_propre(brut) == attendu


def test_texte_retire_tous_les_espaces_en_trop():
    assert texte("  a   b\t\nc ") == "a b c"
    assert texte(None) == ""
    assert texte(12) == "12"


def test_affichage_montre_les_dates_excel_en_francais():
    assert affichage(datetime(2025, 7, 2)) == "02/07/2025"
    assert affichage(" x ") == "x"
