import os
import numpy as np
import re
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import argparse

def main():
    parser = argparse.ArgumentParser(description='Process exam results from different report types.')
    parser.add_argument('--report-type', choices=['dc030', 'all_results'], default='dc030',
                        help='Type of report to process: dc030 (default) or all_results')
    parser.add_argument('--input-file', default='input/input.xlsx',
                        help='Path to the input Excel file (default: input/input.xlsx)')
    parser.add_argument('--sheet-name', default='Overzicht',
                        help='Name of the worksheet to process (default: Overzicht)')
    parser.add_argument('--output-dir', default='output/plots',
                        help='Directory for output plots (default: output/plots)')
    args = parser.parse_args()

    # Load the Excel file
    try:
        data = pd.read_excel(args.input_file, sheet_name=args.sheet_name, engine='openpyxl')
        print(f"Sheet '{args.sheet_name}' succesvol geladen uit {args.input_file}")
    except Exception as e:
        print(f"Fout tijdens laden bestand: {e}")
        exit()

    # Process data based on report type
    if args.report_type == 'dc030':
        process_dc030_report(data, args.output_dir)
    elif args.report_type == 'all_results':
        process_all_results_report(data, args.output_dir)

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
        return data

    replace_map = {'O': 4, 'V': 6, 'G': 8, 'VO': 10, 'NVO': 0, 'J': 0}
    data.loc[:, 'Eindresultaat'] = data['Eindresultaat'].replace(replace_map)
    data.loc[:, 'Eindresultaat'] = (
        data['Eindresultaat'].astype(str).str.replace(',', '.').astype(float)
    )

    return data

def plot_results_per_team_exam(data, output_dir):
    """
    Maakt grafieken van de verdeling van Eindresultaten per opleiding en examen.
    
    Args:
        data (pd.DataFrame): De invoer-DataFrame.
        output_dir (str): Directory om grafieken op te slaan.
    """
    if 'Opleiding Naam' not in data.columns or 'ExamenCode' not in data.columns:
        print("Fout: Vereiste kolommen 'Opleiding Naam' en/of 'ExamenCode' niet gevonden.")
        return

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

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

def process_dc030_report(data, output_dir):
    """Process the DC030 report format."""
    data = prepare_toetsnaam(data)
    data = process_eindresultaat_column(data)
    plot_results_per_team_exam(data, output_dir)

def process_all_results_report(data, output_dir):
    """Process the all results report format."""
    # Assume the new report might have different column names or structure
    # Map the column names to match our expected format if needed
    
    # Check for required columns in the all_results format
    required_columns = ['Student ID', 'Opleiding Naam', 'Toets Naam', 'Resultaat', 'Datum']
    alternate_columns = {
        'Student ID': ['StudentID', 'Studentnummer', 'Student Nummer'],
        'Opleiding Naam': ['OpleidingNaam', 'Opleiding', 'Opleidingnaam'],
        'Toets Naam': ['ToetsNaam', 'Examen Naam', 'ExamenNaam', 'Toetsnaam'],
        'Resultaat': ['Cijfer', 'Eindresultaat', 'Score'],
        'Datum': ['DatumAfname', 'Afnamedatum', 'Datum Afname']
    }
    
    # Try to map column names if they don't match exactly
    column_mapping = {}
    for req_col, alt_cols in alternate_columns.items():
        if req_col in data.columns:
            column_mapping[req_col] = req_col
        else:
            for alt_col in alt_cols:
                if alt_col in data.columns:
                    column_mapping[alt_col] = req_col
                    break
    
    # Rename columns if necessary
    if column_mapping:
        data = data.rename(columns=column_mapping)
    
    # Process result column similar to Eindresultaat
    if 'Resultaat' in data.columns:
        replace_map = {'O': 4, 'V': 6, 'G': 8, 'VO': 10, 'NVO': 0, 'J': 0}
        data.loc[:, 'Resultaat'] = data['Resultaat'].replace(replace_map)
        data.loc[:, 'Resultaat'] = (
            data['Resultaat'].astype(str).str.replace(',', '.').astype(float)
        )
    else:
        print("Fout: kolom 'Resultaat' niet gevonden in all_results rapport.")
        return
    
    # Generate plots per Opleiding and Toets
    if 'Opleiding Naam' in data.columns and 'Toets Naam' in data.columns:
        # Create output directories
        os.makedirs(output_dir, exist_ok=True)
        
        # Create basic histogram plots
        plot_all_results_histograms(data, output_dir)
        
        # Create time series plots to show progression
        plot_results_over_time(data, output_dir)
        
        # Create comparison plots (average scores per test)
        plot_average_scores(data, output_dir)
    else:
        print("Fout: Vereiste kolommen niet gevonden in all_results rapport.")

def plot_all_results_histograms(data, output_dir):
    """Create histograms for all results, grouped by Opleiding and Toets."""
    for (team, exam), group in data.groupby(['Opleiding Naam', 'Toets Naam']):
        plt.figure(figsize=(8, 5))
        ax = sns.histplot(group['Resultaat'], bins=np.arange(0.5, 11.5, 1), kde=False, color='steelblue')

        # Instellingen voor de grafiek
        plt.xticks(range(1, 11))
        plt.xlim(0.5, 10.5)
        plt.xlabel('Resultaten')
        plt.ylabel('Aantal pogingen')
        plt.title(f"{team} | {exam} - Alle resultaten")

        # Voeg nummers boven de balken toe
        max_y = 0
        for p in ax.patches:
            height = p.get_height()
            if height > 0:
                ax.annotate(f'{int(height)}', 
                            (p.get_x() + p.get_width() / 2., height), 
                            ha='center', va='bottom', fontsize=10, color='black')
                max_y = max(max_y, height)

        ax.set_yticks(np.arange(0, max_y + 1, 1))

        # Opslaan
        team_dir = os.path.join(output_dir, 'histograms', sanitize_filename(team))
        os.makedirs(team_dir, exist_ok=True)
        filename = os.path.join(team_dir, sanitize_filename(f"{exam}_all_results.png"))
        
        plt.savefig(filename, bbox_inches='tight')
        plt.close()

        print(f"Histogram opgeslagen: {filename}")

def plot_results_over_time(data, output_dir):
    """Create time series plots to show progression of results over time."""
    # Convert date column to datetime if it's not already
    if 'Datum' in data.columns:
        data['Datum'] = pd.to_datetime(data['Datum'], errors='coerce')
        
        # Group by Opleiding, Toets, and sort by date
        for (team, exam), group in data.groupby(['Opleiding Naam', 'Toets Naam']):
            # Sort by date
            group = group.sort_values('Datum')
            
            plt.figure(figsize=(10, 6))
            
            # Calculate moving average to show trends
            if len(group) > 5:  # Only if we have enough data points
                group['rolling_avg'] = group['Resultaat'].rolling(window=5, min_periods=1).mean()
                
                # Plot individual results as scatter and rolling average as line
                plt.scatter(group['Datum'], group['Resultaat'], alpha=0.5, label='Individual results')
                plt.plot(group['Datum'], group['rolling_avg'], color='red', linewidth=2, label='5-point moving average')
                plt.legend()
            else:
                # Just scatter plot if not enough data for moving average
                plt.scatter(group['Datum'], group['Resultaat'])
            
            plt.xlabel('Datum')
            plt.ylabel('Resultaat')
            plt.title(f"{team} | {exam} - Resultaten over tijd")
            plt.grid(True, linestyle='--', alpha=0.7)
            
            # Make sure y-axis shows relevant range for grades
            plt.ylim(0, 10.5)
            
            # Draw line at passing grade (typically 5.5)
            plt.axhline(y=5.5, color='green', linestyle='-', alpha=0.5)
            
            # Format the date axis nicely
            plt.gcf().autofmt_xdate()
            
            # Opslaan
            team_dir = os.path.join(output_dir, 'time_series', sanitize_filename(team))
            os.makedirs(team_dir, exist_ok=True)
            filename = os.path.join(team_dir, sanitize_filename(f"{exam}_time_series.png"))
            
            plt.savefig(filename, bbox_inches='tight')
            plt.close()
            
            print(f"Tijdreeks grafiek opgeslagen: {filename}")

def plot_average_scores(data, output_dir):
    """Create bar charts showing average scores per test across opleidingen."""
    # Group by Toets Naam and calculate average score
    avg_scores = data.groupby('Toets Naam')['Resultaat'].agg(['mean', 'count']).reset_index()
    avg_scores = avg_scores.sort_values('mean', ascending=False)
    
    # Only plot exams with significant number of attempts (e.g., > 5)
    avg_scores = avg_scores[avg_scores['count'] > 5]
    
    plt.figure(figsize=(12, 8))
    bars = plt.barh(avg_scores['Toets Naam'], avg_scores['mean'], color='steelblue')
    
    # Add count annotations to bars
    for i, bar in enumerate(bars):
        plt.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2, 
                 f"n={int(avg_scores.iloc[i]['count'])}", 
                 va='center')
    
    plt.xlabel('Gemiddeld resultaat')
    plt.ylabel('Toets')
    plt.title('Gemiddelde resultaten per toets (met meer dan 5 pogingen)')
    plt.axvline(x=5.5, color='red', linestyle='--', alpha=0.7)  # Pass/fail line
    plt.grid(axis='x', linestyle='--', alpha=0.7)
    
    # Ensure x-axis shows appropriate grade range
    plt.xlim(0, 10)
    
    # Save the overall comparison
    os.makedirs(os.path.join(output_dir, 'comparison'), exist_ok=True)
    filename = os.path.join(output_dir, 'comparison', 'average_scores_by_exam.png')
    plt.savefig(filename, bbox_inches='tight')
    plt.close()
    
    print(f"Vergelijkingsgrafiek opgeslagen: {filename}")
    
    # Now do the same but per opleiding
    for team, group in data.groupby('Opleiding Naam'):
        if len(group) < 10:  # Skip opleidingen with very few results
            continue
            
        team_avg = group.groupby('Toets Naam')['Resultaat'].agg(['mean', 'count']).reset_index()
        team_avg = team_avg[team_avg['count'] > 3]  # Only exams with more than 3 attempts
        team_avg = team_avg.sort_values('mean', ascending=False)
        
        if len(team_avg) < 3:  # Skip if not enough exams
            continue
            
        plt.figure(figsize=(12, 8))
        bars = plt.barh(team_avg['Toets Naam'], team_avg['mean'], color='steelblue')
        
        for i, bar in enumerate(bars):
            plt.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2, 
                     f"n={int(team_avg.iloc[i]['count'])}", 
                     va='center')
        
        plt.xlabel('Gemiddeld resultaat')
        plt.ylabel('Toets')
        plt.title(f'{team} - Gemiddelde resultaten per toets')
        plt.axvline(x=5.5, color='red', linestyle='--', alpha=0.7)
        plt.grid(axis='x', linestyle='--', alpha=0.7)
        plt.xlim(0, 10)
        
        team_dir = os.path.join(output_dir, 'comparison', sanitize_filename(team))
        os.makedirs(team_dir, exist_ok=True)
        filename = os.path.join(team_dir, sanitize_filename('average_scores_by_exam.png'))
        plt.savefig(filename, bbox_inches='tight')
        plt.close()
        
        print(f"Vergelijkingsgrafiek voor {team} opgeslagen: {filename}")

if __name__ == "__main__":
    main()