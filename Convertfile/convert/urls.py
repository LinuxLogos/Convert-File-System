from .views import *
from django.urls import path # type: ignore

urlpatterns = [
    path("", homepage, name="homepage"),
    path('pdftoexcel', pdftoexcel, name='pdftoexcel'),
    path ("pdftoword", pdftoword, name='pdftoword'),
]
