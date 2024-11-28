import os
import json
import pandas as pd

def load_config(filepath):
    """Charge les données de configuration depuis un fichier JSON."""
    if os.path.exists(filepath):
        with open(filepath, 'r') as f:
            return json.load(f)
    return {}

def save_config(filepath, data):
    """Enregistre les données de configuration dans un fichier JSON."""
    if filepath and os.path.dirname(filepath):
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
    with open(filepath, 'w') as f:
        json.dump(data, f, indent=4)

def calculate_forecast(start_date, end_date, channels, products, config_file):
    config = load_config(config_file)
    audience_file = config.get('audience')
    if not audience_file:
        raise ValueError("Le fichier d'audience n'est pas configuré.")
    
    # Charger les données d'audience
    data = pd.read_excel(audience_file)
    
    # Filtrer les données
    filtered_data = data[
        (data['date'] >= pd.to_datetime(start_date)) &
        (data['date'] <= pd.to_datetime(end_date)) &
        (data['channel'].isin(channels)) &
        (data['product'].isin(products))
    ]
    
    # Générer les prévisions (simulation)
    forecast_data = filtered_data.copy()
    forecast_data['forecast'] = forecast_data['viewing_minutes'] * 1.1  # Exemple de calcul
    output_file = config.get('forecast_output', 'forecast_results.xlsx')
    forecast_data.to_excel(output_file, index=False)
    return output_file
