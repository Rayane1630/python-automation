"""
OUTILS OUTLOOK (WINDOWS)
========================

Piloter Outlook depuis Python : lire les mails, récupérer les pièces
jointes, préparer et envoyer des mails. Utilise pywin32, donc fonctionne
uniquement sous Windows avec Outlook installé et ouvert.

    pip install pywin32

Exemple rapide
--------------
    from outlook_outils import chercher_mails, sauver_pieces_jointes, envoyer_mail

    mails = chercher_mails(sujet_contient="NAV", depuis_jours=1)
    for m in mails:
        sauver_pieces_jointes(m, r"XXXXXXXX", extensions=[".xlsx"])

    envoyer_mail("collegue@exemple.com", "Contrôle NAV", "<p>Bonjour,</p>",
                 pieces_jointes=[r"XXXXXXXX\\Controle_NAV_20261028.xlsx"])
"""

from datetime import datetime, timedelta
from pathlib import Path

try:
    import win32com.client
except ImportError:                    # permet d'importer le module hors Windows
    win32com = None

BOITE_DE_RECEPTION = 6                 # constante Outlook olFolderInbox
ELEMENT_MAIL = 0                       # constante Outlook olMailItem


# ---------------------------------------------------------------------------
# CONNEXION
# ---------------------------------------------------------------------------

def outlook():
    """Renvoie l'application Outlook ouverte sur le poste."""
    if win32com is None:
        raise RuntimeError("pywin32 n'est pas installé (pip install pywin32), "
                           "ou on n'est pas sous Windows.")
    return win32com.client.Dispatch("Outlook.Application")


def dossier_outlook(chemin=None, boite=None):
    """
    Renvoie un dossier Outlook.
        dossier_outlook()                              -> Boîte de réception
        dossier_outlook("Contreparties/BNP")           -> sous-dossier de la Boîte de réception
        dossier_outlook("Boîte de réception", boite="equipe@exemple.com")
                                                       -> boîte partagée
    """
    ns = outlook().GetNamespace("MAPI")
    if boite:
        racine = ns.Folders[boite]
        parties = (chemin or "Boîte de réception").split("/")
        dossier = racine.Folders[parties[0]]
        parties = parties[1:]
    else:
        dossier = ns.GetDefaultFolder(BOITE_DE_RECEPTION)
        parties = chemin.split("/") if chemin else []
    for nom in parties:
        dossier = dossier.Folders[nom]
    return dossier


# ---------------------------------------------------------------------------
# LECTURE DES MAILS
# ---------------------------------------------------------------------------

def chercher_mails(sujet_contient=None, expediteur_contient=None, depuis_jours=1,
                   dossier=None, avec_pj=False):
    """
    Renvoie la liste des mails correspondant aux critères, du plus récent au plus ancien.
        sujet_contient      : texte cherché dans l'objet (insensible à la casse)
        expediteur_contient : texte cherché dans l'adresse ou le nom de l'expéditeur
        depuis_jours        : 1 = depuis hier même heure, 0 = depuis minuit aujourd'hui
        dossier             : objet renvoyé par dossier_outlook() (Boîte de réception par défaut)
        avec_pj             : True pour ne garder que les mails avec pièce jointe
    """
    dossier = dossier or dossier_outlook()
    elements = dossier.Items
    elements.Sort("[ReceivedTime]", True)          # du plus récent au plus ancien

    if depuis_jours == 0:
        limite = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        limite = datetime.now() - timedelta(days=depuis_jours)

    resultats = []
    for mail in elements:
        if getattr(mail, "Class", None) != 43:     # 43 = mail (ignore invitations, etc.)
            continue
        recu = mail.ReceivedTime.replace(tzinfo=None)
        if recu < limite:
            break                                  # liste triée : on peut s'arrêter
        if sujet_contient and sujet_contient.lower() not in (mail.Subject or "").lower():
            continue
        expediteur = f"{mail.SenderName} {mail.SenderEmailAddress}".lower()
        if expediteur_contient and expediteur_contient.lower() not in expediteur:
            continue
        if avec_pj and mail.Attachments.Count == 0:
            continue
        resultats.append(mail)
    return resultats


def resume_mail(mail):
    """Dictionnaire lisible d'un mail (pratique pour un DataFrame de suivi)."""
    return {
        "Recu": mail.ReceivedTime.replace(tzinfo=None),
        "Expediteur": mail.SenderName,
        "Objet": mail.Subject,
        "Nb_PJ": mail.Attachments.Count,
    }


def sauver_pieces_jointes(mail, dossier, extensions=None, prefixe_date=True):
    """
    Enregistre les pièces jointes d'un mail dans un dossier.
        extensions   : ex. [".xlsx", ".pdf"] pour filtrer (None = tout)
        prefixe_date : ajoute la date de réception devant le nom (évite d'écraser)
    Renvoie la liste des chemins enregistrés.
    """
    dossier = Path(dossier)
    dossier.mkdir(parents=True, exist_ok=True)
    chemins = []
    for pj in mail.Attachments:
        nom = pj.FileName
        if extensions and Path(nom).suffix.lower() not in [e.lower() for e in extensions]:
            continue
        if prefixe_date:
            nom = f"{mail.ReceivedTime:%Y%m%d}_{nom}"
        chemin = dossier / nom
        pj.SaveAsFile(str(chemin))
        chemins.append(chemin)
    return chemins


# ---------------------------------------------------------------------------
# ENVOI
# ---------------------------------------------------------------------------

def envoyer_mail(a, objet, corps_html, cc=None, pieces_jointes=None,
                 envoyer=False, garder_signature=True):
    """
    Prépare un mail Outlook.
        a, cc          : une adresse ou une liste d'adresses
        corps_html     : texte du mail en HTML (ex. "<p>Bonjour,</p>")
        pieces_jointes : liste de chemins de fichiers
        envoyer        : False = le mail s'ouvre pour relecture (recommandé)
                         True  = envoi direct
        garder_signature : ajoute la signature Outlook par défaut sous le corps
    """
    def joindre(adresses):
        return "; ".join(adresses) if isinstance(adresses, (list, tuple)) else (adresses or "")

    mail = outlook().CreateItem(ELEMENT_MAIL)
    mail.To = joindre(a)
    mail.CC = joindre(cc)
    mail.Subject = objet

    if garder_signature:
        mail.Display()                       # l'affichage insère la signature
        mail.HTMLBody = corps_html + mail.HTMLBody
    else:
        mail.HTMLBody = corps_html

    for chemin in pieces_jointes or []:
        mail.Attachments.Add(str(Path(chemin).resolve()))

    if envoyer:
        mail.Send()
    else:
        mail.Display()
    return mail


def tableau_html(df, format_nombres="{:,.2f}"):
    """Transforme un DataFrame en tableau HTML propre à coller dans un mail."""
    return df.to_html(index=False, border=0, float_format=format_nombres.format,
                      justify="center").replace(
        "<table", '<table style="border-collapse:collapse;font-family:Arial;font-size:10pt"'
    ).replace("<th>", '<th style="background:#1F3864;color:white;padding:4px 8px">'
    ).replace("<td>", '<td style="border-bottom:1px solid #ddd;padding:4px 8px">')
