from django import forms
from .models import ReferenceFile, AudienceForecastConfig
from .models import Contract

class ReferenceFileForm(forms.ModelForm):
    class Meta:
        model = ReferenceFile
        fields = ['file_type', 'file']

def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['file_type'].widget.attrs.update({'class': 'form-select'})
        self.fields['file'].widget.attrs.update({'class': 'form-control'})

class AudienceForecastForm(forms.ModelForm):
    start_date = forms.DateField(widget=forms.SelectDateWidget, label="Date de début")
    end_date = forms.DateField(widget=forms.SelectDateWidget, label="Date de fin")
    
    class Meta:
        model = AudienceForecastConfig
        fields = ['start_date', 'end_date', 'channels', 'products']

class NewDealForm(forms.ModelForm):
    class Meta:
        model = Contract
        fields = ['network_name', 'cnt_name_grp', 'business_model']