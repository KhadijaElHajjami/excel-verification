import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import warnings
from openpyxl import load_workbook
import urllib.parse
import webbrowser


def charger_fichier_1():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_fichier_1.delete(0, tk.END)
    entry_fichier_1.insert(0, filepath)

def charger_fichier_2():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_fichier_2.delete(0, tk.END)
    entry_fichier_2.insert(0, filepath)

def executer_script():
    fichier_1 = entry_fichier_1.get()
    fichier_2 = entry_fichier_2.get()

    # Votre code de traitement ici...
    # Désactiver les avertissements spécifiques de validation de données
    warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

    # Données de typologie
    data = {
        "Final_Plant_Name": [
            "WATTRELOS", "VILLIERS LE BEL", "VIENNE", "TRAPPES", "TOULOUSE LT", "TOULON", "STRASBOURG",
            "ST THIBAULT ICS", "ST THIBAULT HORE", "ST QUENTIN", "ST GEOURS", "ST GENIS LAVAL", "SENS",
            "ROUEN", "ROANNE", "QUIMPER", "PLOUDANIEL", "PERSAN", "PAU", "PARIS OUEST", "PANTIN", "NIMES",
            "NICE2", "NICE CARROS", "NANTERRE", "NANCY", "MONTLOUIS", "MIOS", "MEAUX", "MARSEILLE",
            "LOURDES", "LOUDUN", "LOUDEAC", "LIMOGES", "LILLE", "LE MANS", "GUILERS", "GUERANDE",
            "GRENOBLE", "GENNEVILLIERS", "FOUGERES", "CLISSON", "CLERMONT FERRAND", "CHATEAUROUX",
            "CAULNES", "CARCASSONNE", "CALAIS", "CAEN", "BRETIGNY", "BORDEAUX", "BLOIS", "BEZONS",
            "ANGERS", "AIX LES MILLES", "AIX LES BAINS", "AIX"
        ],
        "Typology": [
            "FL/WW", "FL", "WW", "WW", "FL/WW", "FL", "FL/WW", "WW", "WW", "FL/WW", "FL/WW", "FL/WW",
            "FL/WW", "FL/WW", "FL/WW", "FL/WW", "FL", "FL", "FL", "FL", "FL/WW", "WW", "FL", "FL/WW",
            "FL", "FL/WW", "WW", "FL", "FL", "FL/WW", "FL", "FL", "WW", "WW", "WW", "WW", "FL/WW", "FL",
            "FL/WW", "FL/WW", "FL", "FL/WW", "FL/WW", "FL/WW", "FL/WW", "FL", "WW", "FL/WW", "FL/WW",
            "FL/WW", "FL/WW", "FL/WW", "FL/WW", "FL", "FL/WW", "FL/WW"
        ]
    }

    # Creating DataFrame
    df_typology = pd.DataFrame(data)

    try:
        wb = load_workbook(fichier_1)
        ws = wb.active
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du chargement du fichier Excel: {e}")
        return

    try:
        df_1 = pd.read_excel(fichier_1, sheet_name='Synthèse en détail')
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du chargement du fichier Excel: {e}")
        return

    df_june = df_1[['Unnamed: 1', 'Unnamed: 7', 'Unnamed: 13', 'Unnamed: 21']]
    df_june.columns = ['Centre', 'Réforme LP', 'Total CST LP', 'Total CST VT']

    df_june['Centre'] = df_june['Centre'].str.strip().str.upper()
    df_typology['Final_Plant_Name'] = df_typology['Final_Plant_Name'].str.strip().str.upper()

    df_combined = pd.merge(df_june, df_typology, how='left', left_on='Centre', right_on='Final_Plant_Name')

    centres_concernes = df_combined[
        (((df_combined['Réforme LP'] == 0) | (df_combined['Total CST LP'] == 0)) & 
        (df_combined['Typology'].isin(["FL", "FL/WW"]))) |
        ((df_combined['Total CST VT'] == 0) & df_combined['Typology'].isin(["WW", "FL/WW"]))
    ]['Centre']

    centres_concernes = centres_concernes.dropna().tolist()

    try:
        df_2 = pd.read_excel(fichier_2)
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du chargement du deuxième fichier Excel: {e}")
        return

    emails_concernes = df_2[df_2['Plant name'].isin(centres_concernes)]['email'].tolist()

    subject = 'Synthèse CST'
    body = (
    "Bonjour à tous,\n\n"
    "Suite à la consolidation de vos fichiers « Suivi CST » par le contrôle de gestion, nous avons constaté qu’il n’y avait pas d’engagés au global Linge Plat / VT ou bien pas d’engagés pour la réforme Linge Plat sur le mois de Mois_Année.\n\n"
    "Pouvez-vous svp mettre à jour vos fichiers Suivi CST jusqu’à mai 2024, directement sur vos temporaires « H:\\NomCentre\\Communs\\Temporaires\\Suivi CST ».\n"
    "La dernière consolidation du contrôle de gestion est prévue xxxx à 14h.\n\n"
    "Merci d’avance,\n"
    "Svp,"
)


    def prepare_mailto_url(emails, subject, body):
        mailto_base = "mailto:"
        recipients = ';'.join(emails) # Utiliser ';' comme séparateur
        # Encoding subject and body separately
        subject_encoded = urllib.parse.quote(subject)
        body_encoded = urllib.parse.quote(body, safe='')  # Encode everything including spaces and newlines
        mailto_url = f"{mailto_base}{recipients}?subject={subject_encoded}&body={body_encoded}"
        return mailto_url

    mailto_url = prepare_mailto_url(emails_concernes, subject, body)
    webbrowser.open(mailto_url)

# Interface utilisateur avec tkinter
root = tk.Tk()
root.title("Automatisation de la vérification des fichiers Excel et envoi d'emails")

frame = tk.Frame(root)
frame.pack(pady=20, padx=20)

label_fichier_1 = tk.Label(frame, text="Chemin du fichier Excel 1:")
label_fichier_1.grid(row=0, column=0, padx=10, pady=5)

entry_fichier_1 = tk.Entry(frame, width=50)
entry_fichier_1.grid(row=0, column=1, padx=10, pady=5)

button_fichier_1 = tk.Button(frame, text="Parcourir...", command=charger_fichier_1)
button_fichier_1.grid(row=0, column=2, padx=10, pady=5)

label_fichier_2 = tk.Label(frame, text="Chemin du fichier Excel 2:")
label_fichier_2.grid(row=1, column=0, padx=10, pady=5)

entry_fichier_2 = tk.Entry(frame, width=50)
entry_fichier_2.grid(row=1, column=1, padx=10, pady=5)

button_fichier_2 = tk.Button(frame, text="Parcourir...", command=charger_fichier_2)
button_fichier_2.grid(row=1, column=2, padx=10, pady=5)

button_exec = tk.Button(frame, text="Exécuter", command=executer_script)
button_exec.grid(row=2, columnspan=3, pady=20)

root.mainloop()
