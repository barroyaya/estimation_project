# estimation/management/commands/fix_units_fk.py
from django.core.management.base import BaseCommand
from django.db import connection, transaction
from django.db.utils import OperationalError

UNITS = [
    ("u",   "Unité",          "u"),
    ("ml",  "Mètre linéaire", "m"),
    ("m2",  "m²",             "m²"),
    ("m3",  "m³",             "m³"),
    ("kg",  "Kilogramme",     "kg"),
    ("h",   "Heure",          "h"),
    ("j",   "Jour",           "j"),
    ("ens", "Ensemble",       "ens"),
    ("ff",  "Forfait",        "ff"),
]

class Command(BaseCommand):
    help = "Crée les unités par défaut et remappe les valeurs texte (u, ml, m2, ...) vers les ids de estimation_unite dans les tables."

    def handle(self, *args, **kwargs):
        self.stdout.write(self.style.HTTP_INFO("==> Vérification/Création des unités"))
        with connection.cursor() as c, transaction.atomic():
            # Créer la table si le modèle a déjà été migré (sinon, ce script doit être lancé après vos migrations de schéma)
            try:
                # Crée les unités si elles n'existent pas
                for code, libelle, symbole in UNITS:
                    c.execute("""
                        INSERT INTO estimation_unite (code, libelle, symbole, is_active)
                        SELECT ?, ?, ?, 1
                        WHERE NOT EXISTS (SELECT 1 FROM estimation_unite WHERE code = ?)
                    """, [code, libelle, symbole, code])
                self.stdout.write(self.style.SUCCESS("   ✔ Unités présentes"))
            except OperationalError as e:
                self.stdout.write(self.style.ERROR(
                    "La table estimation_unite n'existe pas encore. Lance d'abord les migrations de schéma."
                ))
                raise e

            self.stdout.write(self.style.HTTP_INFO("==> Remappage des valeurs texte vers les ids"))

            # Pour SQLite, on peut faire une mise à jour avec sous-requête
            # Remappe estimation_element.unite_id
            c.execute("""
                UPDATE estimation_element
                SET unite_id = (
                    SELECT id FROM estimation_unite u WHERE u.code = estimation_element.unite_id
                )
                WHERE unite_id IN (SELECT code FROM estimation_unite)
            """)
            count_el = c.rowcount

            # Remappe estimation_demandeelement.unite_id
            try:
                c.execute("""
                    UPDATE estimation_demandeelement
                    SET unite_id = (
                        SELECT id FROM estimation_unite u WHERE u.code = estimation_demandeelement.unite_id
                    )
                    WHERE unite_id IN (SELECT code FROM estimation_unite)
                """)
                count_dem = c.rowcount
            except OperationalError:
                # La table peut ne pas exister selon votre état de projet
                count_dem = 0

            self.stdout.write(self.style.SUCCESS(f"   ✔ Éléments mis à jour : {count_el}"))
            self.stdout.write(self.style.SUCCESS(f"   ✔ Demandes mises à jour : {count_dem}"))

            # Optionnel : forcer les valeurs manquantes sur 'u'
            c.execute("""
                UPDATE estimation_element
                SET unite_id = (SELECT id FROM estimation_unite WHERE code='u')
                WHERE unite_id NOT IN (SELECT id FROM estimation_unite)
                   OR unite_id IS NULL
            """)
            c.execute("""
                UPDATE estimation_demandeelement
                SET unite_id = (SELECT id FROM estimation_unite WHERE code='u')
                WHERE unite_id NOT IN (SELECT id FROM estimation_unite)
                   OR unite_id IS NULL
            """)

        self.stdout.write(self.style.SUCCESS("==> Remappage terminé sans erreur."))
