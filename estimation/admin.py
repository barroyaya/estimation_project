from django.contrib import admin
from django.utils import timezone
from django.utils.html import format_html
from django.urls import reverse
from django.http import HttpResponseRedirect
from .models import *


@admin.register(Projet)
class ProjetAdmin(admin.ModelAdmin):
    list_display = ['nom', 'client', 'date_creation', 'actif', 'actions_projet']
    list_filter = ['actif', 'date_creation']
    search_fields = ['nom', 'client']
    readonly_fields = ['date_creation']

    def actions_projet(self, obj):
        return format_html(
            '<a class="button" href="{}">Voir estimation</a>',
            reverse('rapport_projet', args=[obj.pk])
        )

    actions_projet.short_description = "Actions"


@admin.register(Categorie)
class CategorieAdmin(admin.ModelAdmin):
    list_display = ['nom', 'type_categorie', 'code']
    list_filter = ['type_categorie']
    search_fields = ['nom', 'code']


@admin.register(Discipline)
class DisciplineAdmin(admin.ModelAdmin):
    list_display = ['nom', 'code', 'couleur_preview']
    search_fields = ['nom', 'code']

    def couleur_preview(self, obj):
        return format_html(
            '<div style="width: 30px; height: 20px; background-color: {};"></div>',
            obj.couleur
        )

    couleur_preview.short_description = "Couleur"


@admin.register(Element)
class ElementAdmin(admin.ModelAdmin):
    list_display = ['numero', 'designation', 'prix_unitaire', 'unite', 'categorie', 'discipline', 'actif']
    list_filter = ['categorie', 'discipline', 'unite', 'actif']
    search_fields = ['designation', 'numero', 'caracteristiques']
    list_editable = ['prix_unitaire', 'actif']

    fieldsets = (
        ('Informations générales', {
            'fields': ('numero', 'designation', 'caracteristiques')
        }),
        ('Prix et unité', {
            'fields': ('prix_unitaire', 'unite')
        }),
        ('Classification', {
            'fields': ('categorie', 'discipline')
        }),
        ('Statut', {
            'fields': ('actif',)
        }),
    )


@admin.register(DemandeElement)
class DemandeElementAdmin(admin.ModelAdmin):
    list_display = ['designation', 'projet', 'categorie', 'statut', 'prix_unitaire_admin', 'date_demande',
                    'actions_demande']
    list_filter = ['statut', 'categorie', 'discipline', 'date_demande']
    search_fields = ['designation', 'caracteristiques', 'projet__nom']
    readonly_fields = ['date_demande', 'projet', 'categorie', 'discipline']
    list_editable = ['prix_unitaire_admin']

    fieldsets = (
        ('Informations de la demande', {
            'fields': ('projet', 'categorie', 'discipline', 'date_demande')
        }),
        ('Élément demandé', {
            'fields': ('designation', 'caracteristiques', 'unite', 'quantite')
        }),
        ('Validation administrateur', {
            'fields': ('statut', 'prix_unitaire_admin', 'commentaire_admin', 'date_validation')
        }),
    )

    def actions_demande(self, obj):
        if obj.statut == 'en_attente':
            return format_html(
                '<a class="button" href="javascript:void(0)" onclick="approuverDemande({})">Approuver</a> '
                '<a class="button" href="javascript:void(0)" onclick="rejeterDemande({})">Rejeter</a>',
                obj.id, obj.id
            )
        return obj.get_statut_display()

    actions_demande.short_description = "Actions"

    def save_model(self, request, obj, form, change):
        if change and obj.statut in ['approuve', 'rejete'] and not obj.date_validation:
            obj.date_validation = timezone.now()
        super().save_model(request, obj, form, change)

    actions = ['approuver_demandes', 'rejeter_demandes']

    def approuver_demandes(self, request, queryset):
        count = 0
        for demande in queryset:
            if demande.statut == 'en_attente':
                demande.statut = 'approuve'
                demande.date_validation = timezone.now()
                demande.save()
                count += 1
        self.message_user(request, f"{count} demande(s) approuvée(s).")

    approuver_demandes.short_description = "Approuver les demandes sélectionnées"

    def rejeter_demandes(self, request, queryset):
        count = 0
        for demande in queryset:
            if demande.statut == 'en_attente':
                demande.statut = 'rejete'
                demande.date_validation = timezone.now()
                demande.save()
                count += 1
        self.message_user(request, f"{count} demande(s) rejetée(s).")

    rejeter_demandes.short_description = "Rejeter les demandes sélectionnées"


class EstimationElementInline(admin.TabularInline):
    model = EstimationElement
    extra = 0
    readonly_fields = ['cout_total_display', 'type_element']

    def cout_total_display(self, obj):
        return f"{obj.cout_total:,.2f} CFA"

    cout_total_display.short_description = "Coût total"

    def type_element(self, obj):
        if obj.element:
            return "Standard"
        elif obj.demande_element:
            return f"Personnalisé ({obj.demande_element.get_statut_display()})"
        return "Inconnu"

    type_element.short_description = "Type"


@admin.register(EstimationElement)
class EstimationElementAdmin(admin.ModelAdmin):
    list_display = ['projet', 'element', 'quantite', 'prix_unitaire_utilise', 'cout_total_display']
    list_filter = ['projet', 'element__categorie', 'element__discipline']
    search_fields = ['projet__nom', 'element__designation']

    def cout_total_display(self, obj):
        return f"{obj.cout_total:,.2f} CFA"

    cout_total_display.short_description = "Coût total"


@admin.register(EstimationSummary)
class EstimationSummaryAdmin(admin.ModelAdmin):
    list_display = ['projet', 'cout_total_ttc_display', 'derniere_mise_a_jour']
    readonly_fields = [
        'cout_total_materiel', 'cout_total_main_oeuvre',
        'cout_total_transport', 'cout_total_etude',
        'cout_total_ht', 'tva_montant', 'cout_total_ttc',
        'derniere_mise_a_jour'
    ]

    def cout_total_ttc_display(self, obj):
        return f"{obj.cout_total_ttc:,.2f} CFA"

    cout_total_ttc_display.short_description = "Total TTC"

    actions = ['recalculer_totaux']

    def recalculer_totaux(self, request, queryset):
        for summary in queryset:
            summary.calculer_totaux()
        self.message_user(request, "Totaux recalculés avec succès.")

    recalculer_totaux.short_description = "Recalculer les totaux"


# Configuration du site admin
admin.site.site_header = "Administration - Système d'Estimation"
admin.site.site_title = "Estimation Admin"
admin.site.index_title = "Gestion des Projets d'Estimation"