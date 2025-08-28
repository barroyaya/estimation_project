# estimation/apps.py
from django.apps import AppConfig

class EstimationConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "estimation"

    def ready(self):
        # Import side-effects pour enregistrer les receivers
        from . import signals  # noqa: F401
