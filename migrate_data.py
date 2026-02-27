"""
Script de migration des données tree_data.xlsx vers tree_data_internal.xlsx
Ce script reconstruit la hiérarchie parent/enfant à partir du niveau de chaque ligne.
"""
import pandas as pd
import uuid
import os

# Fichiers
SOURCE_FILE = "tree_data.xlsx"
TARGET_FILE = "tree_data_internal.xlsx"

def migrate():
    if not os.path.exists(SOURCE_FILE):
        print(f"Erreur: Le fichier {SOURCE_FILE} n'existe pas.")
        return
    
    print(f"Lecture de {SOURCE_FILE}...")
    df = pd.read_excel(SOURCE_FILE)
    
    # Ignorer les lignes d'en-tête si le fichier a une zone de titre
    # Chercher la ligne qui contient "Position" ou "PartNumber"
    header_row = None
    for i, row in df.iterrows():
        row_values = [str(v).strip() for v in row.values if pd.notna(v)]
        if "Position" in row_values or "PartNumber" in row_values:
            header_row = i
            break
    
    if header_row is not None and header_row > 0:
        print(f"En-tête trouvé à la ligne {header_row}, réajustement...")
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
    
    print(f"Colonnes trouvées: {list(df.columns)}")
    print(f"Nombre de lignes: {len(df)}")
    
    # Vérifier les colonnes requises
    required_cols = ["Position", "PartNumber", "Description"]
    has_niveau = "Niveau" in df.columns
    
    for col in required_cols:
        if col not in df.columns:
            print(f"Erreur: Colonne '{col}' manquante.")
            return
    
    # Construire les données avec ID et ParentID
    result = []
    parent_stack = []  # Liste de tuples (niveau, id)
    
    for i, row in df.iterrows():
        position = str(row["Position"]) if pd.notna(row["Position"]) else ""
        part_number = str(row["PartNumber"]) if pd.notna(row["PartNumber"]) else ""
        description = str(row["Description"]) if pd.notna(row["Description"]) else ""
        
        # Déterminer le niveau
        if has_niveau and pd.notna(row["Niveau"]):
            niveau = int(row["Niveau"])
        else:
            # Estimer le niveau par l'indentation de la position (ex: 1, 1.1, 1.1.1)
            niveau = position.count('.') if position else 0
        
        # Ignorer les lignes vides
        if not part_number.strip():
            continue
        
        # Générer un nouvel ID
        new_id = str(uuid.uuid4())
        
        # Trouver le parent
        parent_id = ""
        if niveau > 0:
            # Chercher le dernier élément avec un niveau inférieur
            while parent_stack and parent_stack[-1][0] >= niveau:
                parent_stack.pop()
            if parent_stack:
                parent_id = parent_stack[-1][1]
        else:
            # Niveau 0 = racine, vider la pile
            parent_stack = []
        
        # Ajouter à la pile des parents potentiels
        parent_stack.append((niveau, new_id))
        
        result.append({
            "ID": new_id,
            "ParentID": parent_id,
            "Position": position,
            "PartNumber": part_number,
            "Description": description,
            "Niveau": niveau
        })
    
    # Créer le DataFrame résultat
    df_result = pd.DataFrame(result, columns=["ID", "ParentID", "Position", "PartNumber", "Description", "Niveau"])
    
    # Sauvegarder
    df_result.to_excel(TARGET_FILE, index=False)
    print(f"\n✅ Migration réussie!")
    print(f"   • {len(result)} éléments migrés")
    print(f"   • Fichier créé: {TARGET_FILE}")
    print(f"\nVous pouvez maintenant lancer l'application avec run_app.bat")

if __name__ == "__main__":
    migrate()
    input("\nAppuyez sur Entrée pour fermer...")
