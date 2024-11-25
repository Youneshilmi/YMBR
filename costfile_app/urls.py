from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),  # Nom : 'home'
    path('file_upload/', views.upload_file, name='file_upload'),  # Nom : 'file_upload'
    path('preferences/', views.preferences, name='preferences'),  # Nom : 'preferences'
    path('configure_forecast/', views.configure_forecast, name='configure_forecast'),  # Nom : 'configure_forecast'
    path('view_forecast/', views.view_forecast, name='view_forecast'),  # Nom : 'view_forecast'
    path('contract_list/', views.contract_list, name='contract_list'),  # Assurez-vous que cette vue existe
    path('new_deal/', views.new_deal, name='new_deal'),
]
