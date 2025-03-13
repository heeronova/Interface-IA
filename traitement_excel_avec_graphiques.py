"""
Script démontrant comment utiliser la classe ChartGenerator pour:
1. Traiter un fichier Excel existant (supprimer des lignes, filtrer, etc.)
2. Générer des graphiques à partir des données modifiées
3. Intégrer les graphiques dans un nouveau fichier Excel
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from ChartGenerator import ChartGenerator
import os
import argparse

def creer_fichier_exemple(output_path):
    """
    Crée un fichier Excel d'exemple avec des données aléatoires.
    
    Args:
        output_path: Chemin du fichier Excel à créer
    
    Returns:
        DataFrame pandas contenant les données générées
    """
    print(f"Création d'un fichier Excel d'exemple: {output_path}")
    
    # Créer des données aléatoires
    np.random.seed(42)
    dates = pd.date_range('20250101', periods=20)
    
    # Créer un DataFrame avec plusieurs colonnes
    df = pd.DataFrame({
        'Date': dates,
        'Produit': np.random.choice(['A', 'B', 'C', 'D'], size=20),
        'Région': np.random.choice(['Nord', 'Sud', 'Est', 'Ouest'], size=20),
        'Ventes': np.random.randint(100, 1000, size=20),
        'Coûts': np.random.randint(50, 800, size=20),
        'Profit': np.random.randint(10, 500, size=20),
        'Satisfaction': np.random.randint(1, 6, size=20)  # Note de 1 à 5
    })
    
    # Calculer le profit réel (Ventes - Coûts)
    df['Profit Réel'] = df['Ventes'] - df['Coûts']
    
    # Ajouter quelques valeurs manquantes
    df.loc[2, 'Ventes'] = np.nan
    df.loc[5, 'Coûts'] = np.nan
    df.loc[8, 'Profit'] = np.nan
    
    # Ajouter quelques valeurs aberrantes
    df.loc[10, 'Ventes'] = 9999
    df.loc[15, 'Coûts'] = 9999
    
    # Enregistrer le DataFrame dans un fichier Excel
    df.to_excel(output_path, index=False)
    
    print(f"Fichier Excel créé avec succès: {output_path}")
    return df

def traiter_fichier_excel(input_path, output_path):
    """
    Traite un fichier Excel existant et génère des graphiques.
    
    Args:
        input_path: Chemin du fichier Excel d'entrée
        output_path: Chemin du fichier Excel de sortie
    
    Returns:
        Chemin du fichier Excel de sortie
    """
    print(f"Traitement du fichier Excel: {input_path}")
    
    # Définir les opérations à effectuer sur les données
    operations = [
        # Supprimer les lignes avec des valeurs manquantes
        {"type": "drop_rows", "rows": [2, 5, 8]},
        
        # Supprimer les valeurs aberrantes (lignes 10 et 15)
        {"type": "drop_rows", "rows": [10, 15]},
        
        # Filtrer pour ne garder que les produits A et B
        {"type": "filter", "column": "Produit", "condition": "equals", "value": "A"},
        
        # Trier par profit réel décroissant
        {"type": "sort", "column": "Profit Réel", "ascending": False}
    ]
    
    # Traiter les données
    df_processed = ChartGenerator.process_excel_data(input_path, operations)
    
    print("Données traitées:")
    print(df_processed.head())
    
    # Définir les configurations des graphiques
    chart_configs = [
        {
            "type": "bar",
            "x_column": "Date",
            "y_column": "Ventes",
            "title": "Ventes du produit A par date"
        },
        {
            "type": "line",
            "x_column": "Date",
            "y_column": "Profit Réel",
            "title": "Évolution du profit réel"
        },
        {
            "type": "pie",
            "labels_column": "Région",
            "values_column": "Ventes",
            "title": "Répartition des ventes par région"
        },
        {
            "type": "scatter",
            "x_column": "Ventes",
            "y_column": "Profit Réel",
            "title": "Relation entre ventes et profit réel"
        }
    ]
    
    # Créer les graphiques et les exporter dans un fichier Excel
    excel_path = ChartGenerator.create_charts_from_dataframe(df_processed, chart_configs, output_path)
    
    print(f"Données traitées et graphiques enregistrés dans: {excel_path}")
    return excel_path

def main():
    """
    Fonction principale qui exécute le traitement d'un fichier Excel.
    """
    parser = argparse.ArgumentParser(description="Traitement de fichier Excel avec génération de graphiques")
    parser.add_argument("--input", help="Chemin du fichier Excel d'entrée")
    parser.add_argument("--output", default="excel_avec_graphiques.xlsx", help="Chemin du fichier Excel de sortie")
    args = parser.parse_args()
    
    # Si aucun fichier d'entrée n'est spécifié, créer un fichier d'exemple
    if args.input is None:
        input_path = "donnees_exemple_complet.xlsx"
        creer_fichier_exemple(input_path)
    else:
        input_path = args.input
    
    # Traiter le fichier Excel et générer des graphiques
    traiter_fichier_excel(input_path, args.output)
    
    print("\nTraitement terminé avec succès!")
    print(f"Vous pouvez maintenant ouvrir le fichier {args.output} pour voir les graphiques.")

if __name__ == "__main__":
    main()