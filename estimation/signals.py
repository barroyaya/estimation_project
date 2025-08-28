# estimation/signals.py
from decimal import Decimal
from django.db.models.signals import post_save, post_delete
from django.dispatch import receiver

from .models import EstimationSummary, DemandeElement, EstimationElement, Element


def _recalc_summary_for_project(projet):
    if not projet:
        return
    summary, _ = EstimationSummary.objects.get_or_create(projet=projet)
    summary.calculer_totaux()


# === Demandes personnalisées (validation, prix, quantité, suppression) ===
@receiver(post_save, sender=DemandeElement)
def demandeelement_saved(sender, instance: DemandeElement, **kwargs):
    # Quel que soit le statut, on recalcule : le template filtre ce qui est approuvé.
    _recalc_summary_for_project(instance.projet)


@receiver(post_delete, sender=DemandeElement)
def demandeelement_deleted(sender, instance: DemandeElement, **kwargs):
    _recalc_summary_for_project(instance.projet)


# === Éléments standards sélectionnés (quantité, ajout, suppression) ===
@receiver(post_save, sender=EstimationElement)
def estimationelement_saved(sender, instance: EstimationElement, **kwargs):
    _recalc_summary_for_project(instance.projet)


@receiver(post_delete, sender=EstimationElement)
def estimationelement_deleted(sender, instance: EstimationElement, **kwargs):
    _recalc_summary_for_project(instance.projet)


# === Changement de prix ou de catégorie d’un Element standard ===
# Si un prix unitaire d'Element change, tous les projets qui l'utilisent doivent être recalculés
@receiver(post_save, sender=Element)
def element_saved(sender, instance: Element, **kwargs):
    projets_ids = (EstimationElement.objects
                   .filter(element=instance)
                   .values_list('projet_id', flat=True)
                   .distinct())
    for pid in projets_ids:
        try:
            summary = EstimationSummary.objects.get(projet_id=pid)
        except EstimationSummary.DoesNotExist:
            continue
        summary.calculer_totaux()
