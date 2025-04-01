import os
import numpy as np
import re
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = 'input/input.xlsx'  # Pas aan naar je bestand
sheet_name = 'Overzicht' # Naam van het werkblad
output_dir = 'output/plots' # Naam van de output map

try:
    data = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    print("Sheet succesvol geladen")
except Exception as e:
    print(f"Fout tijdens laden bestand: {e}")
    exit()

def sanitize_filename(filename):
    """Aanpassen bestandsnamen, vervangt ongeldige karakters."""
    filename = re.sub(r'[\/:*?"<>|()]', '', filename)  # Remove () and other forbidden characters
    filename = filename.replace(" ", "_")  # Replace spaces with underscores
    filename = re.sub(r'[^A-Za-z0-9_.-]', '_', filename)  # Replace other invalid characters
    return filename

def prepare_toetsnaam(data):
    """Concatenate de 'OP code' en 'Toets Code' kolommen tot een nieuwe kolom 'ExamenCode' en voegt deze toe aan de DataFrame."""
    if 'OP code' not in data.columns or 'Toets Code' not in data.columns:
        print("Fout: Vereiste kolommen 'OP code' en/of 'Toets Code' niet gevonden.")
        return data  # Return de ongewijzigde data in plaats van exit()

    data['ExamenCode'] = data['OP code'].astype(str) + "_" + data['Toets Code'].astype(str)
    return data  # Zorg dat de gewijzigde DataFrame wordt teruggegeven


def process_eindresultaat_column(data):
    """
    Verwerkt de 'Eindresultaat'-kolom:
    - Vervangt 'O', 'V', 'G' door 4, 6 en 8. (Onvoldoende, Voldoende, Goed)
    - Vervangt 'VO' door 10 en 'NVO' door 0. (Voldaan en Niet Voldaan)
    - Vervangt 'J' door 0. (Blokkade Cum Laude)
    - Converteert waarden met een komma naar numeriek.
    
    Returns:
        pd.DataFrame: De aangepaste DataFrame.
    """
    if 'Eindresultaat' not in data.columns:
        print("Fout: kolom 'Eindresultaat' niet gevonden.")
        exit()

    replace_map = {'O': 4, 'V': 6, 'G': 8, 'VO': 10, 'NVO': 0, 'J': 0}
    data.loc[:, 'Eindresultaat'] = data['Eindresultaat'].replace(replace_map)
    data.loc[:, 'Eindresultaat'] = (
        data['Eindresultaat'].astype(str).str.replace(',', '.').astype(float)
    )

    return data

# Verwerk de kolom
data = process_eindresultaat_column(data)

def plot_results_per_team_exam(data, output_dir=output_dir):
    """
    Maakt grafieken van de verdeling van Eindresultaten per opleiding en examen.
    
    Args:
        data (pd.DataFrame): De invoer-DataFrame.
        output_dir (str): Directory om grafieken op te slaan.
    """
    if 'Opleiding Naam' not in data.columns or 'ExamenCode' not in data.columns:
        print("Fout: Vereiste kolommen 'Opleiding Naam' en/of 'ExamenCode' niet gevonden.")
        return

    # Unieke combinaties van Opleiding en Examen
    for (team, exam), group in data.groupby(['Opleiding Naam', 'ExamenCode']):
        plt.figure(figsize=(8, 5))
        ax = sns.histplot(group['Eindresultaat'], bins=np.arange(0.5, 11.5, 1), kde=False, color='steelblue')

        # Instellingen voor de grafiek
        plt.xticks(range(1, 11))  # X-as van 1 t/m 10
        plt.xlim(0.5, 10.5)  # Limieten om bins mooi weer te geven
        plt.xlabel('Resultaten') # Label voor X-as
        plt.ylabel('Aantal studenten') # Label voor Y-as
        plt.title(f"{team} | {exam}") # Titel van de grafiek

        # Voeg nummers boven de balken toe
        max_y = 0  # Houd bij wat de hoogste waarde is
        for p in ax.patches:
            height = p.get_height()
            if height > 0:  # Alleen labels tonen als er een waarde is
                ax.annotate(f'{int(height)}', 
                            (p.get_x() + p.get_width() / 2., height), 
                            ha='center', va='bottom', fontsize=10, color='black')
                max_y = max(max_y, height)  # Update hoogste waarde

        # Zorg ervoor dat de Y-as alleen hele getallen toont
        ax.set_yticks(np.arange(0, max_y + 1, 1))

        # Opslaan
        team_dir = os.path.join(output_dir, sanitize_filename(team))
        os.makedirs(team_dir, exist_ok=True)
        filename = os.path.join(team_dir, sanitize_filename(f"{exam}.png"))
        
        plt.savefig(filename, bbox_inches='tight')
        plt.close()

        print(f"Grafiek opgeslagen: {filename}")

# Maak de grafieken
data = prepare_toetsnaam(data)
plot_results_per_team_exam(data)