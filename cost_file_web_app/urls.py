from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('costfile_app.urls')),  # Inclure les URLs de l'application costfile_app
]
