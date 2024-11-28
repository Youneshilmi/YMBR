from django import forms
from .models import ReferenceFile, AudienceForecastConfig
from .models import Contract
from calendar import month_name

class ReferenceFileForm(forms.ModelForm):
    class Meta:
        model = ReferenceFile
        fields = ['file_type', 'file']

def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['file_type'].widget.attrs.update({'class': 'form-select'})
        self.fields['file'].widget.attrs.update({'class': 'form-control'})

from .models import AudienceForecastConfig, Channel, Product


class AudienceForecastForm(forms.ModelForm):
    reference_month = forms.ChoiceField(
        choices=[(str(month), month_name[month]) for month in range(1, 13)],  # Choices from 1 to 12 with full month names
        label="Mois de référence",
        widget=forms.Select,  # Dropdown for month selection
    )
    start_date = forms.ChoiceField(
        choices=[(str(year), str(year)) for year in range(2024, 2031)],  # Define a range of years
        label="Date de début",
        widget=forms.Select,  # Use a select widget for year selection
    )
    end_date = forms.ChoiceField(
        choices=[(str(year), str(year)) for year in range(2024, 2031)],  # Define a range of years
        label="Date de fin",
        widget=forms.Select,  # Use a select widget for year selection
    )
    channels = forms.ModelMultipleChoiceField(
        queryset=Channel.objects.all(),
        widget=forms.SelectMultiple(attrs={'size': 10}),  # Using SelectMultiple widget
        required=True,
        label="Chaînes"
    )
    products = forms.ModelMultipleChoiceField(
        queryset=Product.objects.all(),
        widget=forms.SelectMultiple(attrs={'size': 10}),  # Using SelectMultiple widget
        required=True,
        label="Produits"
    )
    class Meta:
        model = AudienceForecastConfig
        fields = ['reference_month','start_date', 'end_date', 'channels', 'products']

class NewDealForm(forms.ModelForm):
    class Meta:
        model = Contract
        fields = ['network_name', 'cnt_name_grp', 'business_model']