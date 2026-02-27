# Guide d'utilisation - Gestionnaire d'Arbre
Ce programme vous permet de créer une arborescence de "Part Number" et "Description" et de sauvegarder le tout dans un fichier Excel en temps réel.
Le fichier Excel généré contient des **groupes (plan)** que vous pouvez plier/déplier comme dans l'application.

## Installation

1. Assurez-vous d'avoir Python installé sur votre machine.
2. Le dossier contient un fichier `requirements.txt` avec les bibliothèques nécessaires (`pandas`, `openpyxl`).

## Lancement Rapide

Double-cliquez simplement sur le fichier **`run_app.bat`**. 
Il installera automatiquement les librairies manquantes et lancera le programme.

## Utilisation

- **Ajouter Racine**: Crée un nouveau noeud au premier niveau.
- **Ajouter Frère**: Crée un noeud au même niveau que le noeud sélectionné.
- **Ajouter Enfant**: Crée un sous-noeud pour l'élément sélectionné.
- **Importer Masse**: Permet de coller du texte et d'importer plusieurs enfants d'un coup en définissant les colonnes (index de début/fin).
- **Modifier**: Change le texte du noeud sélectionné.
- **Supprimer**: Efface le noeud sélectionné et TOUS ses enfants.
- **Sauvegarde**: La sauvegarde est automatique à chaque modification (création, édition, suppression).

## Données
Le fichier `tree_data.xlsx` est créé automatiquement dans le même dossier. Il sert de base de données. Si vous fermez et relancez le programme, vous retrouverez votre arbre tel quel.
