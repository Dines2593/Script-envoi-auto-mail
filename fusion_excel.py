import pandas as pd
import os

# Chemin du répertoire contenant les fichiers Excel
dossier = r'H:\Desktop\\alternance\\fichier_a_fusionner'

# Initialiser un DataFrame vide
dataframe_final = pd.DataFrame()

# Parcourir tous les fichiers dans le répertoire
for fichier in os.listdir(dossier):
    if fichier.endswith(".xls"):  # Assurez-vous que le fichier est un fichier Excel
        chemin_fichier = os.path.join(dossier, fichier)
        try:
            df_temp = pd.read_excel(chemin_fichier)
            dataframe_final = pd.concat([dataframe_final, df_temp], ignore_index=True)
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier {fichier}: {str(e)}")

# Écrire le DataFrame final dans un nouveau fichier Excel avec le moteur openpyxl
try:
    fichier_final = os.path.join(dossier, 'fichier_final.xlsx')
    dataframe_final.to_excel(fichier_final, index=False, engine='openpyxl')
    print(f"Fusion réussie. Le fichier final a été créé avec succès à l'emplacement : {fichier_final}")
except Exception as e:
    print(f"Erreur lors de l'écriture du fichier final : {str(e)}")
