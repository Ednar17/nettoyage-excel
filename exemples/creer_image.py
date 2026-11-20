"""Fabrique l'image avant/après de l'annonce à partir des deux fichiers de démo.

Usage : python creer_image.py
Produit : avant_apres.png (1280 x 720)
"""

from PIL import Image, ImageDraw, ImageFont

from nettoyage_excel import lire, nettoyer
from nettoyage_excel.regles import affichage

LARGEUR, HAUTEUR = 1280, 720
MARGE = 48
COLONNES = ["Date commande", "Client", "Téléphone", "Montant", "Statut"]
LARGEURS = [220, 300, 250, 210, 204]
NB_LIGNES = 5

ENCRE, ENCRE_DOUCE, FOND, TRAIT = "#0b0b0b", "#52514e", "#ffffff", "#e1e0d9"
ROUGE, ROUGE_FOND = "#b3261e", "#fdecea"
VERT, VERT_FOND = "#006300", "#e6f4e6"
ENTETE = "#1f4e78"


def police(taille: int, gras: bool = False):
    fichier = "segoeuib.ttf" if gras else "segoeui.ttf"
    return ImageFont.truetype(f"C:/Windows/Fonts/{fichier}", taille)


def tableau(dessin, haut: int, etiquette: str, couleur: str, fond: str, lignes: list[list[str]]) -> int:
    dessin.rounded_rectangle([MARGE, haut, MARGE + 118, haut + 36], radius=8, fill=fond)
    dessin.text((MARGE + 59, haut + 18), etiquette, font=police(20, True), fill=couleur, anchor="mm")
    haut += 48

    x = MARGE
    dessin.rectangle([MARGE, haut, LARGEUR - MARGE, haut + 34], fill=ENTETE)
    for titre, largeur in zip(COLONNES, LARGEURS):
        dessin.text((x + 14, haut + 17), titre, font=police(18, True), fill="#ffffff", anchor="lm")
        x += largeur
    haut += 34

    for ligne in lignes:
        x = MARGE
        for valeur, largeur in zip(ligne, LARGEURS):
            dessin.text((x + 14, haut + 17), valeur, font=police(19), fill=ENCRE, anchor="lm")
            x += largeur
        haut += 34
        dessin.line([MARGE, haut, LARGEUR - MARGE, haut], fill=TRAIT, width=1)
    return haut


def main() -> None:
    brutes = lire("exemple_avant.xlsx")
    propres, _, bilan = nettoyer(brutes)

    # On montre les mêmes commandes avant et après : les premières lignes valides du fichier.
    choisies = []
    for numero, ligne in brutes:
        resultat, a_verifier, _ = nettoyer([(numero, ligne)])
        if resultat and len(choisies) < NB_LIGNES:
            choisies.append((ligne, resultat[0]))

    avant = [[affichage(brute[c]).replace(".", ",") if c == "Montant" and isinstance(brute[c], float)
              else str(brute[c]) for c in COLONNES] for brute, _ in choisies]
    apres = [[propre["Date commande"].strftime("%d/%m/%Y"), propre["Client"], propre["Téléphone"],
              f"{propre['Montant']:,.2f} €".replace(",", " ").replace(".", ","), propre["Statut"]]
             for _, propre in choisies]

    image = Image.new("RGB", (LARGEUR, HAUTEUR), FOND)
    dessin = ImageDraw.Draw(image)
    dessin.text((MARGE, 34), "Votre fichier Excel nettoyé automatiquement", font=police(34, True), fill=ENCRE)

    bas = tableau(dessin, 108, "AVANT", ROUGE, ROUGE_FOND, avant)
    bas = tableau(dessin, bas + 26, "APRÈS", VERT, VERT_FOND, apres)

    resume = (f"Sur {len(brutes)} lignes : {bilan['doublons supprimés']} doublons supprimés, "
              f"{bilan['lignes vides supprimées']} lignes vides retirées, "
              f"{bilan['lignes à vérifier']} lignes douteuses listées à part.")
    dessin.text((MARGE, bas + 28), resume, font=police(21), fill=ENCRE_DOUCE)

    image.save("avant_apres.png")
    print("avant_apres.png créé")


if __name__ == "__main__":
    main()
