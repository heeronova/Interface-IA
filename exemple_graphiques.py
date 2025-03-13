"""
Script d'exemple pour démontrer l'utilisation de la classe ChartGenerator
pour créer des graphiques et les intégrer dans Excel.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from ChartGenerator import ChartGenerator
import os

def exemple_graphiques_simples():
    """
    Exemple de création de graphiques simples et leur intégration dans Excel.
    """
    print("Création de graphiques simples...")
    
    # Générer des données d'exemple
    sample_data = ChartGenerator.generate_sample_data()
    
    # Créer différents types de graphiques
    line_chart = ChartGenerator.create_line_chart(
        sample_data['line']['simple'], 
        title="Exemple de graphique en courbes",
        xlabel="Axe X",
        ylabel="Valeurs"
    )
    
    bar_chart = ChartGenerator.create_bar_chart(
        sample_data['bar'],
        title="Exemple de graphique à barres",
        xlabel="Catégories",
        ylabel="Valeurs"
    )
    
    pie_chart = ChartGenerator.create_pie_chart(
        sample_data['pie'],
        title="Exemple de graphique en camembert"
    )
    
    scatter_chart = ChartGenerator.create_scatter_plot(
        sample_data['scatter'],
        title="Exemple de nuage de points",
        xlabel="X",
        ylabel="Y"
    )
    
    histogram_chart = ChartGenerator.create_histogram(
        sample_data['histogram'],
        title="Exemple d'histogramme",
        xlabel="Valeurs",
        ylabel="Fréquence"
    )
    
    # Créer un fichier Excel avec tous les graphiques
    charts_data = [
        (line_chart, "Courbes"),
        (bar_chart, "Barres"),
        (pie_chart, "Camembert"),
        (scatter_chart, "Nuage de points"),
        (histogram_chart, "Histogramme")
    ]
    
    output_path = "exemples_graphiques.xlsx"
    ChartGenerator.create_excel_with_charts(charts_data, output_path)
    
    print(f"Graphiques créés et enregistrés dans {output_path}")
    return output_path

def exemple_graphiques_a_partir_de_donnees():
    """
    Exemple de création de graphiques à partir de données dans un DataFrame.
    """
    print("Création de graphiques à partir de données...")
    
    # Créer un DataFrame d'exemple
    np.random.seed(42)
    dates = pd.date_range('20250101', periods=12)
    df = pd.DataFrame({
        'Date': dates,
        'Ventes': np.random.randint(100, 1000, size=12),
        'Dépenses': np.random.randint(50, 800, size=12),
        'Profit': np.random.randint(10, 500, size=12),
        'Région': np.random.choice(['Nord', 'Sud', 'Est', 'Ouest'], size=12)
    })
    
    # Enregistrer les données brutes dans un fichier Excel
    df.to_excel("donnees_exemple.xlsx", index=False)
    print("Données d'exemple enregistrées dans donnees_exemple.xlsx")
    
    # Définir les configurations des graphiques
    chart_configs = [
        {
            "type": "line",
            "x_column": "Date",
            "y_column": "Ventes",
            "title": "Évolution des ventes"
        },
        {
            "type": "bar",
            "x_column": "Date",
            "y_column": "Dépenses",
            "title": "Dépenses mensuelles"
        },
        {
            "type": "pie",
            "labels_column": "Région",
            "values_column": "Profit",
            "title": "Répartition des profits par région"
        },
        {
            "type": "scatter",
            "x_column": "Ventes",
            "y_column": "Profit",
            "title": "Relation entre ventes et profits"
        }
    ]
    
    # Créer les graphiques et les exporter dans un fichier Excel
    output_path = "graphiques_donnees.xlsx"
    ChartGenerator.create_charts_from_dataframe(df, chart_configs, output_path)
    
    print(f"Graphiques créés à partir des données et enregistrés dans {output_path}")
    return output_path

def exemple_manipulation_donnees_excel():
    """
    Exemple de manipulation de données Excel et création de graphiques.
    """
    print("Manipulation de données Excel...")
    
    # Vérifier si le fichier de données existe déjà
    input_path = "donnees_exemple.xlsx"
    if not os.path.exists(input_path):
        exemple_graphiques_a_partir_de_donnees()
    
    # Définir les opérations à effectuer sur les données
    operations = [
        {
            "type": "filter",
            "column": "Ventes",
            "condition": "greater_than",
            "value": 500
        },
        {
            "type": "sort",
            "column": "Profit",
            "ascending": False
        }
    ]
    
    # Traiter les données
    df_processed = ChartGenerator.process_excel_data(input_path, operations)
    
    # Définir les configurations des graphiques
    chart_configs = [
        {
            "type": "bar",
            "x_column": "Date",
            "y_column": "Ventes",
            "title": "Ventes supérieures à 500"
        },
        {
            "type": "line",
            "x_column": "Date",
            "y_column": "Profit",
            "title": "Profits (triés par ordre décroissant)"
        }
    ]
    
    # Créer les graphiques et les exporter dans un fichier Excel
    output_path = "graphiques_donnees_traitees.xlsx"
    ChartGenerator.create_charts_from_dataframe(df_processed, chart_configs, output_path)
    
    print(f"Données traitées et graphiques enregistrés dans {output_path}")
    return output_path

def main():
    """
    Fonction principale qui exécute tous les exemples.
    """
    print("=== Démonstration des fonctionnalités de génération de graphiques ===")
    
    # Exemple 1: Graphiques simples
    exemple_graphiques_simples()
    
    # Exemple 2: Graphiques à partir de données
    exemple_graphiques_a_partir_de_donnees()
    
    # Exemple 3: Manipulation de données Excel
    exemple_manipulation_donnees_excel()
    
    print("\nTous les exemples ont été exécutés avec succès!")
    print("Vous pouvez maintenant ouvrir les fichiers Excel générés pour voir les graphiques.")

if __name__ == "__main__":
    main()