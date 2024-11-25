import os
import json
import pandas as pd
from django.shortcuts import render, redirect
from django.contrib import messages
from .models import ReferenceFile, Channel, Product, Contract
from .forms import ReferenceFileForm, NewDealForm, AudienceForecastForm
from .utils import save_config, load_config, calculate_forecast

CONFIG_FILE = os.path.join(os.path.expanduser('~'), '.config', 'forecast_config.json')

CONFIG_FILE = "config.json"

COLUMN_MAPPING = {
    'product': 'PROD_EN_NAME',  # Colonne pour les produits
    'channel': 'CHANNEL_ALIAS',  # Colonne pour les chaînes
    }

def home(request):
    return render(request, 'home.html')


def save_channels(data):
    """Enregistre les chaînes à partir des données Excel."""
    channel_column = COLUMN_MAPPING.get('channel')
    if channel_column not in data.columns:
        raise ValueError(f"Le fichier ne contient pas la colonne '{channel_column}'.")
    for channel in data[channel_column].unique():
        Channel.objects.get_or_create(name=channel)


def save_products(data):
    """Enregistre les produits à partir des données Excel."""
    product_column = COLUMN_MAPPING.get('product')
    if product_column not in data.columns:
        raise ValueError(f"Le fichier ne contient pas la colonne '{product_column}'.")
    for product in data[product_column].unique():
        Product.objects.get_or_create(name=product)


def view_forecast(request):
    config_data = load_config(CONFIG_FILE)
    forecast_file = config_data.get('forecast_output')
    if forecast_file and os.path.exists(forecast_file):
        return redirect(f'/media/{forecast_file}')
    else:
        messages.error(request, "Le fichier de prévisions n'existe pas.")
        return redirect('configure_forecast')


def contract_list(request):
    contracts = Contract.objects.all()
    return render(request, 'contract_list.html', {'contracts': contracts})


def upload_file(request):
    """Gère le téléchargement des fichiers Excel."""
    if request.method == 'POST':
        form = ReferenceFileForm(request.POST, request.FILES)
        if form.is_valid():
            file_instance = form.save()
            config_data = load_config(CONFIG_FILE)
            config_data[file_instance.file_type] = file_instance.file.path
            save_config(CONFIG_FILE, config_data)

            try:
                data = pd.read_excel(file_instance.file.path)
                if file_instance.file_type == 'product':
                    save_products(data)
                elif file_instance.file_type == 'channel':
                    save_channels(data)
            except ValueError as e:
                messages.error(request, f"Erreur : {e}")
                return redirect('file_upload')
            except Exception as e:
                messages.error(request, f"Erreur inattendue : {e}")
                return redirect('file_upload')

            messages.success(request, f"Fichier {file_instance.file_type} chargé avec succès.")
            return redirect('file_upload')
    else:
        form = ReferenceFileForm()
    files = ReferenceFile.objects.all()
    return render(request, 'upload_file.html', {'form': form, 'files': files})

def configure_forecast(request):
    channels = Channel.objects.all()
    products = Product.objects.all()
    if request.method == 'POST':
        form = AudienceForecastForm(request.POST)
        if form.is_valid():
            selected_channels = request.POST.getlist('channels')
            selected_products = request.POST.getlist('products')
            forecast_file = calculate_forecast(
                form.cleaned_data['start_date'],
                form.cleaned_data['end_date'],
                selected_channels,
                selected_products
            )
            messages.success(request, f"Prévisions générées : {forecast_file}")
            return redirect('configure_forecast')
    else:
        form = AudienceForecastForm()
    return render(request, 'configure_forecast.html', {
        'form': form,
        'channels': channels,
        'products': products
    })


def new_deal(request):
    if request.method == 'POST':
        form = NewDealForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, "Nouveau contrat créé avec succès.")
            return redirect('contract_list')
    else:
        form = NewDealForm()
    return render(request, 'new_deal.html', {'form': form})

def preferences(request):
    config_data = load_config(CONFIG_FILE)
    if request.method == 'POST':
        # Sauvegarder les préférences modifiées par l'utilisateur
        updated_config = request.POST.dict()  # Récupérer les données du formulaire
        save_config(CONFIG_FILE, updated_config)
        messages.success(request, "Préférences mises à jour.")
        return redirect('preferences')

    return render(request, 'preferences.html', {'config_data': config_data})