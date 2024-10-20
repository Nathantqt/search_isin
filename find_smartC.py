import pandas as pd
import os
import re

# Chemin vers le fichier Excel
file_path = r"C:\Users\natha\Desktop\Perso\VBA\ISIN File New.xlsx"
# Répertoire à vérifier si le code n'est pas trouvé
directory_path = r"C:\Users\natha\Desktop\Perso\Issuances 2024"

# Les noms des feuilles à parcourir
sheets = ["SwissTrade", "NonSwissTrade", "NordicTrade"]

# Fonction pour rechercher un code dans toutes les feuilles
def search_code_in_excel(file_path, code):
    try:
        # Parcourir chaque feuille du fichier Excel
        for sheet in sheets:
            # Lire la feuille dans un DataFrame en mode lecture seule
            df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
            
            # Convertir toutes les colonnes en chaînes de caractères pour faciliter la recherche
            df = df.astype(str)
            
            # Vérifier si le code existe dans une des colonnes
            found_row = df[
                (df['Valoren'] == code) |
                (df['ISIN'] == code) |
                (df['Series Code'] == code) |
                (df['SmartD'] == code) |
                (df['Common Code'] == code)
            ]
            
            # Si une ligne est trouvée, l'afficher
            if not found_row.empty:
                for idx, row in found_row.iterrows():
                    print("\nDétails de la ligne trouvée :")
                    for col in df.columns:
                        if col not in ['Description', 'Notional']:
                            print(f"{col}: {row[col]}")
                            if col == 'SmartD' : 
                                smartD = row[col]
                return smartD  # Quitter après avoir trouvé le code
        
        # Si le code n'a pas été trouvé dans aucune des feuilles, vérifier le répertoire
        print(f"Le code '{code}' n'a été trouvé dans aucune feuille.")
        smartD = check_directory_for_code(directory_path, code)
        return smartD
        
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier Excel : {str(e)}")

# Fonction pour vérifier si le code est dans le nom d'un dossier
def check_directory_for_code(directory_path, code):
    try:
        smartD = 0
        # Lister tous les dossiers dans le répertoire spécifié
        for folder_name in os.listdir(directory_path):
            if code in folder_name:
                print(f"Dossier trouvé : {folder_name}")
                match = re.search(r'\b(?:CE|EI)\w+', folder_name)
                if match:
                    smartD = match.group()
                    return smartD
                return smartD
                # Quitter après avoir trouvé un dossier
        
        # Si aucun dossier ne contient le code
        print(f"Aucun dossier ne contient le code '{code}'.")
        
    except Exception as e:
        print(f"Erreur lors de l'accès au répertoire : {str(e)}")

# Fonction principale pour exécuter le script
def main():
    print("Bienvenue dans le système de recherche ISIN.")
    smartD = 0
    while True:
        # Entrée utilisateur
        code_input = input("Entrez un code (Valoren, ISIN, Series Code, ou SmartD) ou 'q' pour quitter : ")
        
        if code_input.lower() == 'q':
            break  # Quitter la boucle si l'utilisateur tape 'q'
        elif code_input.lower() == 'y':
            print(smartD) # Recherche du SmartC Code
        
        else :
            smartD = search_code_in_excel(file_path, code_input)
            print(smartD)

# Exécuter la fonction principale
if __name__ == "__main__":
    main()