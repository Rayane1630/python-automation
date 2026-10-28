"""
EXEMPLE : CONSOLIDER LES FICHIERS REÇUS PAR MAIL ET LES RENVOYER
================================================================

Scénario typique du matin :
    1. récupérer tous les mails du jour dont l'objet contient "Valorisation" ;
    2. enregistrer leurs pièces jointes Excel dans un dossier daté ;
    3. empiler tous les fichiers dans un seul classeur, nommé avec la date ;
    4. préparer un mail avec le classeur en pièce jointe et un tableau récapitulatif.

À adapter : les XXXX (dossier), le mot-clé de l'objet et le destinataire.
"""

from excel_outils import consolider_dossier, creer_excel, dossier_du_jour, nom_fichier
from outlook_outils import chercher_mails, envoyer_mail, sauver_pieces_jointes, tableau_html

DOSSIER_RACINE = r"XXXXXXXXXXXXXXXXXXXXXXXX"
MOT_CLE_OBJET = "Valorisation"
DESTINATAIRE = "XXXXXXXX@XXXXXXXX.com"


def main():
    # 1. Mails du jour
    mails = chercher_mails(sujet_contient=MOT_CLE_OBJET, depuis_jours=0, avec_pj=True)
    print(f"{len(mails)} mail(s) trouvé(s)")
    if not mails:
        return

    # 2. Pièces jointes dans RACINE/AAAA/MM/AAAAMMJJ
    dossier = dossier_du_jour(DOSSIER_RACINE)
    for mail in mails:
        for chemin in sauver_pieces_jointes(mail, dossier, extensions=[".xlsx", ".xls", ".csv"]):
            print(f"  enregistré : {chemin.name}")

    # 3. Consolidation dans un seul classeur daté
    donnees = consolider_dossier(dossier, motif="*.xls*")
    recap = donnees.groupby("Fichier_source").size().rename("Nb_lignes").reset_index()
    fichier = creer_excel(dossier / nom_fichier("Consolidation_" + MOT_CLE_OBJET),
                          {"Donnees": donnees, "Recap": recap})
    print(f"Classeur créé : {fichier}")

    # 4. Mail prêt à relire (envoyer=False : rien ne part sans validation)
    corps = ("<p>Bonjour,</p>"
             f"<p>Ci-joint la consolidation des fichiers {MOT_CLE_OBJET} du jour.</p>"
             + tableau_html(recap) + "<p>Bonne journée,</p>")
    envoyer_mail(DESTINATAIRE, f"Consolidation {MOT_CLE_OBJET}", corps,
                 pieces_jointes=[fichier], envoyer=False)


if __name__ == "__main__":
    main()
