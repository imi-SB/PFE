import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import uuid
import os

# Nom du fichier Excel pour la sauvegarde
EXCEL_FILE = "tree_data.xlsx"

class TreeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestionnaire d'Arbre - Part Number & Description")
        self.root.geometry("900x600")

        # Cadre pour les boutons
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        # Boutons d'action
        tk.Button(button_frame, text="Ajouter Racine", command=self.add_root).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Ajouter Frère", command=self.add_sibling).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Ajouter Enfant", command=self.add_child).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Importer Masse", command=self.import_bulk_children).pack(side=tk.LEFT, padx=5) # Nouveau bouton
        tk.Button(button_frame, text="Modifier", command=self.edit_node).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Supprimer", command=self.delete_node).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Sauvegarder", command=self.save_data).pack(side=tk.RIGHT, padx=5)
        
        # Boutons Copier/Coller
        tk.Button(button_frame, text="Copier Branche", command=self.copy_node).pack(side=tk.LEFT, padx=5)
        self.btn_paste = tk.Button(button_frame, text="Coller Branche", command=self.paste_node, state=tk.DISABLED)
        self.btn_paste.pack(side=tk.LEFT, padx=5)

        # Arbre (Treeview)
        # Colonnes: Part Number (sera dans la colonne #0 pour l'arborescence), Position, Description
        self.tree = ttk.Treeview(self.root, columns=("Position", "Description"))
        
        # Configuration de la colonne #0 (l'arbre lui-même) -> Part Number
        self.tree.heading("#0", text="Part Number (Arbre)", anchor=tk.W)
        self.tree.heading("Position", text="Position", anchor=tk.W)
        self.tree.heading("Description", text="Description", anchor=tk.W)
        
        # Configuration des colonnes
        self.tree.column("#0", stretch=tk.YES, width=300) 
        self.tree.column("Position", stretch=tk.NO, width=80) # Position est souvent court
        self.tree.column("Description", stretch=tk.YES, width=420)

        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Ajout d'une barre de défilement verticale
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)  
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.place(relx=1.0, rely=0.0, relheight=1.0, anchor="ne") 

        # Dictionnaire pour stocker les données en mémoire (id -> {parent_id, position, part_number, description})
        self.data_store = {}
        
        # Presse-papiers pour copier/coller des branches
        self.clipboard = None

        # Chargement initial des données
        self.load_data()

    def load_data(self):
        """Charge les données depuis le fichier Excel s'il existe."""
        if os.path.exists(EXCEL_FILE):
            try:
                df = pd.read_excel(EXCEL_FILE)
                # Assurons-nous que le fichier n'est pas vide et a les bonnes colonnes (ou compatibles)
                if not df.empty:
                    # Nettoyer l'arbre actuel
                    for item in self.tree.get_children():
                        self.tree.delete(item)
                    self.data_store = {}

                    # Convertir le DataFrame en dictionnaire pour un accès rapide
                    nodes = {}
                    for index, row in df.iterrows():
                        node_id = str(row["ID"])
                        parent_id = str(row["ParentID"]) if "ParentID" in df.columns and pd.notna(row["ParentID"]) else ""
                        part_number = str(row["PartNumber"]) if "PartNumber" in df.columns else ""
                        description = str(row["Description"]) if "Description" in df.columns and pd.notna(row["Description"]) else ""
                        position = str(row["Position"]) if "Position" in df.columns and pd.notna(row["Position"]) else ""

                        nodes[node_id] = {"parent_id": parent_id, "part_number": part_number, "description": description, "position": position}

                    # Reconstruire l'arbre
                    to_add = list(nodes.keys())
                    added = set()

                    # On ajoute d'abord les racines (parent_id vide ou non trouvé dans nodes)
                    for node_id in to_add[:]:
                        parent_id = nodes[node_id]["parent_id"]
                        if not parent_id or parent_id not in nodes:
                            self.insert_node_in_tree("", node_id, nodes[node_id]["position"], nodes[node_id]["part_number"], nodes[node_id]["description"])
                            self.data_store[node_id] = nodes[node_id]
                            added.add(node_id)
                            to_add.remove(node_id)
                    
                    # Ensuite on ajoute les enfants itérativement
                    last_count = len(to_add) + 1
                    while to_add:
                        current_count = len(to_add)
                        if current_count == last_count:
                            print("Attention: Des orphelins ont été détectés et ignorés.")
                            break 
                        last_count = current_count

                        for node_id in to_add[:]:
                            parent_id = nodes[node_id]["parent_id"]
                            if parent_id in added:
                                self.insert_node_in_tree(parent_id, node_id, nodes[node_id]["position"], nodes[node_id]["part_number"], nodes[node_id]["description"])
                                self.data_store[node_id] = nodes[node_id]
                                added.add(node_id)
                                to_add.remove(node_id)

            except Exception as e:
                messagebox.showerror("Erreur de chargement", f"Impossible de charger le fichier Excel:\n{e}")

    def save_data(self):
        """Sauvegarde les données dans Excel avec groupement (Outlining)."""
        export_data = []
        
        # Parcours récursif pour obtenir l'ordre visuel et les niveaux
        def traverse(parent_id, level):
            children = self.tree.get_children(parent_id)
            for child_id in children:
                node_data = self.data_store[child_id]
                export_data.append({
                    "ID": child_id,
                    "ParentID": node_data["parent_id"],
                    "Position": node_data.get("position", ""),
                    "PartNumber": node_data["part_number"],
                    "Description": node_data["description"],
                    "_Level": level 
                })
                traverse(child_id, level + 1)

        traverse("", 0)

        if not export_data:
            # Si vide, on crée juste les headers
            df = pd.DataFrame(columns=["ID", "ParentID", "Position", "PartNumber", "Description"])
            df.to_excel(EXCEL_FILE, index=False)
            return

        df = pd.DataFrame(export_data, columns=["ID", "ParentID", "Position", "PartNumber", "Description", "_Level"])
        
        try:
            # On sauvegarde sans la colonne interne _Level
            df_to_save = df.drop(columns=["_Level"])
            df_to_save.to_excel(EXCEL_FILE, index=False)
            
            # Post-traitement avec OpenPyXL pour le groupement
            try:
                import openpyxl
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb.active
                
                # IMPORTANT: Pour que les '+' soient sur la ligne du parent (en haut du groupe)
                ws.sheet_properties.outlinePr.summaryBelow = False
                
                # Appliquer les niveaux de plan (grouping)
                # Les données commencent à la ligne 2 (1 = Header)
                for i, row_data in enumerate(export_data):
                    level = row_data["_Level"]
                    if level > 0:
                        # row index = i + 2 (car 1-based + header)
                        ws.row_dimensions[i + 2].outlineLevel = level
                
                wb.save(EXCEL_FILE)
                print("Sauvegarde avec groupement Excel effectuée.")
                
            except ImportError:
                print("OpenPyXL non trouvé pour le groupement, sauvegarde simple effectuée.")
                
        except Exception as e:
            messagebox.showerror("Erreur de sauvegarde", f"Impossible de sauvegarder le fichier Excel:\n{e}\nVérifiez que le fichier n'est pas ouvert ailleurs.")

    def insert_node_in_tree(self, parent_id, node_id, position, part_number, description):
        """Helper pour insérer dans l'arbre visuel."""
        # text=part_number pour l'afficher dans la colonne #0 (l'arbre)
        # values=(position, description)
        self.tree.insert(parent_id, 'end', iid=node_id, text=part_number, values=(position, description))
        
        # On ouvre le parent pour montrer le nouvel enfant
        if parent_id:
            self.tree.item(parent_id, open=True)

    def add_root(self):
        self.prompt_and_add_node("")

    def add_sibling(self):
        """Ajoute un noeud au même niveau que la sélection actuelle."""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un noeud pour lui ajouter un frère.")
            return
        
        node_id = selected_item[0]
        # Trouver le parent du noeud sélectionné
        parent_id = self.tree.parent(node_id)
        
        self.prompt_and_add_node(parent_id)

    def add_child(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un noeud parent.")
            return
        parent_id = selected_item[0]
        self.prompt_and_add_node(parent_id)

    def import_bulk_children(self):
        """Ouvre une fenêtre pour importer plusieurs enfants d'un coup via copier-coller."""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un noeud parent pour les items importés.")
            return
        parent_id = selected_item[0]
        
        # Fenêtre de dialogue
        dialog = tk.Toplevel(self.root)
        dialog.title("Importer en masse via texte")
        dialog.geometry("700x500")

        tk.Label(dialog, text="Collez vos données ci-dessous :").pack(anchor=tk.W, padx=10, pady=(10, 0))
        text_area = tk.Text(dialog, height=15)
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Cadre pour les paramètres de découpage
        frame = tk.LabelFrame(dialog, text="Délimitation des colonnes (index de caractère, 0 = début de ligne)")
        frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Toggle pour basculer entre deux modes
        toggle_frame = tk.Frame(frame)
        toggle_frame.pack(padx=5, pady=5)
        
        mode_var = tk.BooleanVar(value=True)  # True = Mode normal, False = Mode alternatif
        
        grid_frame = tk.Frame(frame)
        grid_frame.pack(padx=5, pady=5)

        # Ligne 1: Position
        tk.Label(grid_frame, text="Index Début Position :").grid(row=0, column=0, sticky=tk.E, padx=5)
        entry_pos_start = tk.Entry(grid_frame, width=5)
        entry_pos_start.insert(0, "0") 
        entry_pos_start.grid(row=0, column=1, padx=5)

        tk.Label(grid_frame, text="Index Fin Position :").grid(row=0, column=2, sticky=tk.E, padx=5)
        entry_pos_end = tk.Entry(grid_frame, width=5)
        entry_pos_end.insert(0, "3") 
        entry_pos_end.grid(row=0, column=3, padx=5)

        # Ligne 2: Part Number
        tk.Label(grid_frame, text="Index Début Part Number :").grid(row=1, column=0, sticky=tk.E, padx=5)
        entry_pn_start = tk.Entry(grid_frame, width=5)
        entry_pn_start.insert(0, "4") 
        entry_pn_start.grid(row=1, column=1, padx=5)

        tk.Label(grid_frame, text="Index Fin Part Number :").grid(row=1, column=2, sticky=tk.E, padx=5)
        entry_pn_end = tk.Entry(grid_frame, width=5)
        entry_pn_end.insert(0, "16") 
        entry_pn_end.grid(row=1, column=3, padx=5)

        # Ligne 3: Description
        tk.Label(grid_frame, text="Index Début Description :").grid(row=2, column=0, sticky=tk.E, padx=5)
        entry_desc_start = tk.Entry(grid_frame, width=5)
        entry_desc_start.insert(0, "19") 
        entry_desc_start.grid(row=2, column=1, padx=5)

        tk.Label(grid_frame, text="Caractère de fin (Optionnel) :").grid(row=2, column=2, sticky=tk.E, padx=5)
        entry_desc_end_char = tk.Entry(grid_frame, width=5)
        entry_desc_end_char.insert(0, ".") 
        entry_desc_end_char.grid(row=2, column=3, padx=5)

        # Fonction pour basculer entre les modes
        def toggle_mode():
            if mode_var.get():  # Mode normal
                entry_pos_end.delete(0, tk.END)
                entry_pos_end.insert(0, "3")
                entry_pn_start.delete(0, tk.END)
                entry_pn_start.insert(0, "4")
                entry_pn_end.delete(0, tk.END)
                entry_pn_end.insert(0, "16")
                entry_desc_start.delete(0, tk.END)
                entry_desc_start.insert(0, "19")
            else:  # Mode alternatif
                entry_pos_end.delete(0, tk.END)
                entry_pos_end.insert(0, "6")
                entry_pn_start.delete(0, tk.END)
                entry_pn_start.insert(0, "7")
                entry_pn_end.delete(0, tk.END)
                entry_pn_end.insert(0, "19")
                entry_desc_start.delete(0, tk.END)
                entry_desc_start.insert(0, "21")
            update_highlights()
        
        # Checkbox pour le toggle avec commande
        tk.Checkbutton(toggle_frame, text="Mode Normal (décocher pour Mode Alternatif)", 
                       variable=mode_var, command=toggle_mode).pack(side=tk.LEFT)

        # Configuration des tags pour les couleurs
        text_area.tag_configure("pos", background="#ffeb3b") # Jaune
        text_area.tag_configure("pn", background="#80deea")  # Cyan
        text_area.tag_configure("desc", background="#c8e6c9") # Vert clair

        def update_highlights(event=None):
            text_area.tag_remove("pos", "1.0", tk.END)
            text_area.tag_remove("pn", "1.0", tk.END)
            text_area.tag_remove("desc", "1.0", tk.END)
            
            try:
                pos_start = int(entry_pos_start.get())
                pos_end = int(entry_pos_end.get())
                pn_start = int(entry_pn_start.get())
                pn_end = int(entry_pn_end.get())
                desc_start = int(entry_desc_start.get())
                stop_char = entry_desc_end_char.get()
            except ValueError:
                return 

            content = text_area.get("1.0", tk.END)
            lines = content.split('\n')
            
            for i, line in enumerate(lines):
                line_idx = i + 1
                if not line.strip(): continue
                
                # Position tagging
                if len(line) > pos_start:
                    p_end = min(len(line), pos_end)
                    if p_end > pos_start:
                        text_area.tag_add("pos", f"{line_idx}.{pos_start}", f"{line_idx}.{p_end}")
                
                # PN tagging
                if len(line) > pn_start:
                    p_end = min(len(line), pn_end)
                    if p_end > pn_start:
                        text_area.tag_add("pn", f"{line_idx}.{pn_start}", f"{line_idx}.{p_end}")
                
                # Desc tagging
                if len(line) > desc_start:
                    raw_desc = line[desc_start:]
                    d_end = len(line)
                    
                    if stop_char:
                        found_idx = raw_desc.find(stop_char)
                        if found_idx != -1:
                            d_end = desc_start + found_idx
                    
                    if d_end > desc_start:
                        text_area.tag_add("desc", f"{line_idx}.{desc_start}", f"{line_idx}.{d_end}")

        # Bindings pour mise à jour temps réel
        text_area.bind('<KeyRelease>', update_highlights)
        entry_pos_start.bind('<KeyRelease>', update_highlights)
        entry_pos_end.bind('<KeyRelease>', update_highlights)
        entry_pn_start.bind('<KeyRelease>', update_highlights)
        entry_pn_end.bind('<KeyRelease>', update_highlights)
        entry_desc_start.bind('<KeyRelease>', update_highlights)
        entry_desc_end_char.bind('<KeyRelease>', update_highlights)

        # Appel initial pour colorer si des valeurs par défaut sont présentes
        dialog.after(100, update_highlights) # Petite pause pour s'assurer que la fenêtre est rendue

        def do_import():
            try:
                raw_text = text_area.get("1.0", tk.END).strip()
                if not raw_text: return
                
                pos_start = int(entry_pos_start.get())
                pos_end = int(entry_pos_end.get())
                pn_start = int(entry_pn_start.get())
                pn_end = int(entry_pn_end.get())
                desc_start = int(entry_desc_start.get())
                stop_char = entry_desc_end_char.get()
                
                lines = raw_text.split('\n')
                count = 0
                for line in lines:
                    line = line.rstrip() 
                    if not line: continue
                    try:
                        # 1. Extraction Position
                        position = ""
                        if len(line) > pos_start:
                            current_pos_end = min(len(line), pos_end)
                            position = line[pos_start:current_pos_end].strip()

                        # 2. Extraction Part Number - Condition critique
                        if len(line) <= pn_start: continue 
                        
                        current_pn_end = min(len(line), pn_end)
                        part_number = line[pn_start:current_pn_end].strip()
                        
                        # 3. Extraction Description
                        description = ""
                        if len(line) > desc_start:
                            raw_desc = line[desc_start:]
                            if stop_char and stop_char in raw_desc:
                                description = raw_desc.split(stop_char, 1)[0].strip()
                            else:
                                description = raw_desc.strip()
                        
                        if part_number: 
                            new_id = str(uuid.uuid4())
                            # Update parameters including position
                            self.insert_node_in_tree(parent_id, new_id, position, part_number, description)
                            self.data_store[new_id] = {
                                "parent_id": parent_id,
                                "position": position,
                                "part_number": part_number,
                                "description": description
                            }
                            count += 1
                    except IndexError:
                        continue 
                
                self.save_data() 
                messagebox.showinfo("Succès", f"{count} éléments importés.")
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez entrer des nombres entiers valides pour les index.")

        tk.Button(dialog, text="Lancer l'Importation", command=do_import, bg="#dddddd", height=2).pack(pady=10)

    def prompt_and_add_node(self, parent_id):
        # Dialogue pour Position
        position = simpledialog.askstring("Nouveau Noeud", "Entrez la Position (Optionnel):", initialvalue="")
        if position is None: position = ""

        # Dialogue pour Part Number
        part_number = simpledialog.askstring("Nouveau Noeud", "Entrez le Part Number:")
        if part_number is None: return # Annulé

        # Dialogue pour Description
        description = simpledialog.askstring("Nouveau Noeud", "Entrez la Description:")
        if description is None: description = ""

        # Création
        new_id = str(uuid.uuid4())
        self.insert_node_in_tree(parent_id, new_id, position, part_number, description)
        
        # Mise à jour des données
        self.data_store[new_id] = {
            "parent_id": parent_id,
            "position": position,
            "part_number": part_number,
            "description": description
        }

        # Sauvegarde temps réel
        self.save_data()

    def edit_node(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un noeud à modifier.")
            return
        node_id = selected_item[0]
        
        # Récupérer les données
        current_data = self.data_store.get(node_id)
        if not current_data: return 

        new_pos = simpledialog.askstring("Modifier", "Modifier Position:", initialvalue=current_data.get("position", ""))
        if new_pos is None: return

        new_pn = simpledialog.askstring("Modifier", "Modifier Part Number:", initialvalue=current_data["part_number"])
        if new_pn is None: return

        new_desc = simpledialog.askstring("Modifier", "Modifier Description:", initialvalue=current_data["description"])
        if new_desc is None: return

        # Mise à jour arbre
        self.tree.item(node_id, text=new_pn, values=(new_pos, new_desc))

        # Mise à jour données
        self.data_store[node_id]["position"] = new_pos
        self.data_store[node_id]["part_number"] = new_pn
        self.data_store[node_id]["description"] = new_desc

        # Sauvegarde
        self.save_data()

    def delete_node(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un noeud à supprimer.")
            return
        node_id = selected_item[0]
        
        if messagebox.askyesno("Confirmation", "Voulez-vous vraiment supprimer ce noeud et TOUS ses enfants?"):
            self.delete_node_recursive(node_id)
            self.save_data()

    def delete_node_recursive(self, node_id):
        # Supprimer visuellement (supprime aussi les enfants dans Treeview)
        # Mais on doit aussi nettoyer self.data_store récursivement
        
        # Trouver les enfants dans data_store
        children = [k for k, v in self.data_store.items() if v["parent_id"] == node_id]
        for child in children:
            self.delete_node_recursive(child)
            
        # Supprimer soi-même
        if node_id in self.data_store:
            del self.data_store[node_id]
        
        # Si le noeud existe encore dans l'arbre (c'est la racine de la suppression), on le supprime
        if self.tree.exists(node_id):
            self.tree.delete(node_id)

    def copy_node(self):
        """Copie le noeud sélectionné et toute sa descendance dans le presse-papiers."""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un noeud à copier.")
            return
        
        node_id = selected_item[0]
        
        # Fonction récursive pour construire la structure de données de la branche
        def build_subtree(current_id):
            node_data = self.data_store.get(current_id).copy()
            # On ne garde pas l'ID ni le parent_id car ils changeront au collage
            # Mais on garde les données pour les recréer
            
            children = [k for k, v in self.data_store.items() if v["parent_id"] == current_id]
            children_data = []
            for child_id in children:
                children_data.append(build_subtree(child_id))
            
            return {
                "data": node_data,
                "children": children_data
            }
            
        self.clipboard = build_subtree(node_id)
        self.btn_paste.config(state=tk.NORMAL)
        messagebox.showinfo("Copié", "Branche copiée dans le presse-papiers.")

    def paste_node(self):
        """Colle le contenu du presse-papiers sous le noeud sélectionné."""
        if not self.clipboard:
            messagebox.showwarning("Presse-papiers vide", "Rien à coller.")
            return

        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un parent pour coller.")
            return
        
        parent_id = selected_item[0]
        
        # Fonction récursive pour recréer les noeuds
        def recreate_subtree(subtree_structure, current_parent_id):
            node_data = subtree_structure["data"]
            
            new_id = str(uuid.uuid4())
            position = node_data.get("position", "")
            part_number = node_data.get("part_number", "")
            description = node_data.get("description", "")
            
            # Insérer dans l'arbre et le data_store
            self.insert_node_in_tree(current_parent_id, new_id, position, part_number, description)
            self.data_store[new_id] = {
                "parent_id": current_parent_id,
                "position": position,
                "part_number": part_number,
                "description": description
            }
            
            # Gérer les enfants
            for child_struct in subtree_structure["children"]:
                recreate_subtree(child_struct, new_id)
                
        try:
            recreate_subtree(self.clipboard, parent_id)
            self.save_data()
            messagebox.showinfo("Succès", "Branche collée avec succès. Les enfants ont été clonés.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du collage : {e}")

if __name__ == "__main__":
    # Vérification des dépendances au lancement
    try:
        import pandas
        import openpyxl
    except ImportError:
        messagebox.showerror("Erreur de dépendances", "Les bibliothèques 'pandas' et 'openpyxl' sont requises.\nVeuillez les installer avec: pip install pandas openpyxl")
        exit()

    root = tk.Tk()
    app = TreeApp(root)
    root.mainloop()
