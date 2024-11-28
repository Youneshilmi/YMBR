from django.db import models

def __str__(self):
        return f"{self.get_file_type_display()} uploaded on {self.uploaded_at}"

class ReferenceFile(models.Model):
    FILE_TYPES = [
        ('audience', 'Audience File'),
        ('cost', 'Cost File'),
        ('channel', 'Channel Grouping File'),
        ('product', 'Product Grouping File'),
    ]
    file_type = models.CharField(max_length=50, choices=FILE_TYPES)
    file = models.FileField(upload_to='reference_files/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

class AudienceForecastConfig(models.Model):
    start_date = models.DateField()
    end_date = models.DateField()
    channels = models.CharField(max_length=255)
    products = models.CharField(max_length=255)

class Contract(models.Model):
    network_name = models.CharField(max_length=255)
    cnt_name_grp = models.CharField(max_length=255)
    business_model = models.CharField(max_length=100)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.network_name} - {self.business_model}"
    
class Channel(models.Model):
    name = models.CharField(max_length=255, unique=True)

    def __str__(self):
        return self.name

class Product(models.Model):
    name = models.CharField(max_length=255, unique=True)

    def __str__(self):
        return self.name