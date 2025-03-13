# Fonctionnalités de Génération de Graphiques

Ce module permet de générer des graphiques à partir de données et de les intégrer dans des fichiers Excel. Il offre une interface simple pour créer différents types de graphiques (courbes, histogrammes, camemberts, nuages de points, etc.) et les exporter dans des feuilles Excel séparées.

## Fonctionnalités

- Génération de différents types de graphiques:
  - Courbes (simples et multiples)
  - Histogrammes
  - Camemberts
  - Nuages de points
  - Barres
- Intégration automatique des graphiques dans des feuilles Excel séparées
- Graphiques en haute résolution avec titres, légendes et étiquettes
- Manipulation des données Excel (suppression de lignes, filtrage, tri, etc.)
- Création de graphiques à partir de DataFrames pandas

## Installation des dépendances

Pour utiliser ce module, vous devez installer les dépendances suivantes:

```bash
pip install matplotlib pandas openpyxl pillow numpy
```

## Utilisation

### Création de graphiques simples

```python
from ChartGenerator import ChartGenerator

# Générer des données d'exemple
sample_data = ChartGenerator.generate_sample_data()

# Créer un graphique en courbes
line_chart = ChartGenerator.create_line_chart(
    sample_data['line']['simple'], 
    title="Exemple de graphique en courbes",
    xlabel="Axe X",
    ylabel="Valeurs"
)

# Créer un graphique à barres
bar_chart = ChartGenerator.create_bar_chart(
    sample_data['bar'],
    title="Exemple de graphique à barres",
    xlabel="Catégories",
    ylabel="Valeurs"
)

# Créer un fichier Excel avec les graphiques
charts_data = [
    (line_chart, "Courbes"),
    (bar_chart, "Barres")
]

output_path = "exemples_graphiques.xlsx"
ChartGenerator.create_excel_with_charts(charts_data, output_path)
```

### Création de graphiques à partir d'un DataFrame

```python
import pandas as pd
from ChartGenerator import ChartGenerator

# Créer ou charger un DataFrame
df = pd.read_excel("donnees.xlsx")

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
        "x_column": "Produit",
        "y_column": "Quantité",
        "title": "Ventes par produit"
    },
    {
        "type": "pie",
        "labels_column": "Région",
        "values_column": "Ventes",
        "title": "Répartition des ventes par région"
    }
]

# Créer les graphiques et les exporter dans un fichier Excel
output_path = "graphiques_donnees.xlsx"
ChartGenerator.create_charts_from_dataframe(df, chart_configs, output_path)
```

### Traitement de données Excel et génération de graphiques

```python
from ChartGenerator import ChartGenerator

# Définir les opérations à effectuer sur les données
operations = [
    # Supprimer des lignes
    {"type": "drop_rows", "rows": [2, 5, 8]},
    
    # Filtrer les données
    {"type": "filter", "column": "Produit", "condition": "equals", "value": "A"},
    
    # Trier les données
    {"type": "sort", "column": "Ventes", "ascending": False}
]

# Traiter les données
input_path = "donnees.xlsx"
df_processed = ChartGenerator.process_excel_data(input_path, operations)

# Définir les configurations des graphiques
chart_configs = [
    {
        "type": "bar",
        "x_column": "Produit",
        "y_column": "Ventes",
        "title": "Ventes par produit (après traitement)"
    }
]

# Créer les graphiques et les exporter dans un fichier Excel
output_path = "graphiques_donnees_traitees.xlsx"
ChartGenerator.create_charts_from_dataframe(df_processed, chart_configs, output_path)
```

## Exemples

Deux scripts d'exemple sont fournis pour montrer comment utiliser les fonctionnalités de génération de graphiques:

1. `exemple_graphiques.py` - Montre comment créer différents types de graphiques et les intégrer dans Excel.
2. `traitement_excel_avec_graphiques.py` - Montre comment traiter un fichier Excel existant, y appliquer des modifications, puis générer des graphiques à partir des données modifiées.

Pour exécuter les exemples:

```bash
python exemple_graphiques.py
python traitement_excel_avec_graphiques.py
```

## Intégration dans l'application Chat

La classe `ChartGenerator` est utilisée par les méthodes dans `chart_functions.py` pour générer des graphiques et les intégrer dans Excel. Ces méthodes sont appelées par l'interface utilisateur de l'application Chat.

Pour utiliser les fonctionnalités de génération de graphiques dans l'application Chat:

1. Cliquez sur le menu "Fichier" > "Créateur de graphiques"
2. Sélectionnez le type de graphique souhaité
3. Entrez les données pour le graphique
4. Cliquez sur "Aperçu" pour voir le graphique
5. Cliquez sur "Exporter vers Excel" pour enregistrer le graphique dans un fichier Excel

Vous pouvez également générer des exemples de graphiques en cliquant sur le menu "Fichier" > "Générer des exemples de graphiques".