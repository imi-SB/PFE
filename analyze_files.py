import pandas as pd
import os
import sys

files = [
    "PMP 2022 (1).xls",
    "PMP Condit PLF 2021 (1).xlsx",
    "CL_EXPORT_ETAT_CONSOMATION3060118986151578542 (1).xls",
    "LISTE DE CONTRÔLE DISPOSITIF DE SECURITE (1).xlsx",
    "tree_data.xlsx"
]

base_path = r"c:\Users\Administrateur\Desktop\PFE"
output_file = os.path.join(base_path, "analysis_output.txt")

with open(output_file, 'w', encoding='utf-8') as out:
    for f in files:
        full_path = os.path.join(base_path, f)
        out.write(f"\n{'='*60}\n")
        out.write(f"FICHIER: {f}\n")
        out.write('='*60 + "\n")
        try:
            # Essayer de lire toutes les feuilles
            xl = pd.ExcelFile(full_path)
            out.write(f"Feuilles: {xl.sheet_names}\n")
            
            for sheet in xl.sheet_names[:2]:  # Max 2 feuilles
                df = pd.read_excel(full_path, sheet_name=sheet)
                out.write(f"\n--- Feuille: {sheet} ---\n")
                out.write(f"Shape: {df.shape}\n")
                out.write(f"Colonnes: {list(df.columns)}\n")
                out.write(f"\nApercu (3 premieres lignes):\n")
                out.write(df.head(3).to_string() + "\n")
        except Exception as e:
            out.write(f"ERREUR: {e}\n")

    out.write("\n\nFIN DE L'ANALYSE\n")

print(f"Analyse terminee. Resultats dans: {output_file}")
