# Python Automation: Outlook & Excel

Small, reusable functions to automate daily back-office and fund management tasks: fetch emails, save attachments, build dated Excel files, consolidate folders and send reports.

| File | Functions |
|---|---|
| `outlook_outils.py` | `chercher_mails`, `sauver_pieces_jointes`, `envoyer_mail`, `tableau_html`, `dossier_outlook` (shared mailboxes, subfolders) |
| `excel_outils.py` | `nom_fichier`, `dossier_du_jour`, `dernier_fichier`, `renommer_avec_date`, `lire_excel`, `consolider_dossier`, `creer_excel`, `ajouter_onglet`, `mettre_en_forme` |
| `exemple_consolidation_mails.py` | End-to-end example: today's emails → attachments → one consolidated, formatted workbook → email ready to review |

```python
from excel_outils import nom_fichier, creer_excel
creer_excel(nom_fichier("Controle_NAV"), {"NAV": df}, {"NAV": "#,##0.00"})
# -> Controle_NAV_20261028.xlsx, formatted headers, frozen panes, filters, fitted columns
```

The Outlook functions need Windows with Outlook running (`pywin32`). Emails open for review by default and are sent only with `envoyer=True`.

```bash
pip install -r requirements.txt
```

Code comments are in French.
