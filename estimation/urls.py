# estimation/urls.py

from django.urls import path
from . import views

urlpatterns = [

# Authentification
    path('login/', views.client_login, name='client_login'),
    path('register/', views.client_register, name='client_register'),
    path('logout/', views.client_logout, name='client_logout'),
    path('profile/', views.client_profile, name='client_profile'),
    path('change-password/', views.client_change_password, name='client_change_password'),
    path('', views.index, name='index'),
    path('projets/', views.project_selection, name='project_selection'),
    path('categories/', views.category_selection, name='category_selection'),
    path('elements/<int:categorie_id>/', views.item_selection, name='item_selection'),
    path('rapport/<int:projet_id>/', views.rapport_projet, name='rapport_projet'),
    path('ajax/update-quantity/', views.ajax_update_quantity, name='ajax_update_quantity'),

    # Nouvelles URLs pour les éléments personnalisés
    path('demandes-personnalisees/', views.demandes_personnalisees, name='demandes_personnalisees'),
    path('supprimer-demande/<int:demande_id>/', views.supprimer_demande, name='supprimer_demande'),
    path('export-pdf/<int:projet_id>/', views.export_pdf_reportlab, name='export_pdf'),
    path('export-excel/<int:projet_id>/', views.export_excel_advanced, name='export_excel'),

    path('sablage-tuyauterie/<int:categorie_id>/', views.sablage_tuyauterie, name='sablage_tuyauterie'),
    path('ajax/calculer-surface-sablage/', views.ajax_calculer_surface_sablage, name='ajax_calculer_surface_sablage'),

]