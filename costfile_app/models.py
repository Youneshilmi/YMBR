from django.db import models
import datetime
def __str__(self):
        return f"{self.get_file_type_display()} uploaded on {self.uploaded_at}"

class ReferenceFile(models.Model):
    FILE_TYPES = [
        ('audience', 'Audience File'),
        ('cost', 'Cost File'),
        ('channel', 'Channel Grouping File'),
        ('product', 'Product Grouping File'),
        ('contract', 'Contracts File')
    ]
    file_type = models.CharField(max_length=50, choices=FILE_TYPES)
    file = models.FileField(upload_to='reference_files/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

class AudienceForecastConfig(models.Model):
    start_date = models.PositiveIntegerField()  # Store year as a number
    end_date = models.PositiveIntegerField()    # Store year as a number
    channels = models.CharField(max_length=255)
    products = models.CharField(max_length=255)

class Contract(models.Model):
    key = models.CharField(max_length=255)
    provider = models.CharField(max_length=255)
    business_model = models.CharField(max_length=255)
    varf = models.CharField(max_length=100, default='unspecified')
    checktype = models.CharField(max_length=255, default='unspecified')
    year = models.PositiveIntegerField(default=datetime.datetime.now().year)
    allocation = models.CharField(max_length=255, default='unspecified')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.key} - {self.business_model}"
    
class Channel(models.Model):
    name = models.CharField(max_length=255, unique=True)

    def __str__(self):
        return self.name

class Product(models.Model):
    name = models.CharField(max_length=255, unique=True)

    def __str__(self):
        return self.name