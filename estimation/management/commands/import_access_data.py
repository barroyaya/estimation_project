import csv
import os
from decimal import Decimal
from django.core.management.base import BaseCommand
from django.db import transaction
from estimation.models import Categorie, Discipline, Element


class Command(BaseCommand):
    help = 'Import des données depuis Access (fichiers CSV)'

    def add_arguments(self, parser):
        parser.add_argument('csv_directory', help='Répertoire contenant les fichiers CSV')

    def handle(self, *args, **options):
        csv_directory = options['csv_directory']

        # Créer les disciplines de base
        self.create_disciplines()

        # Créer les catégories de base
        self.create_categories()

        # Mapper les fichiers CSV aux catégories
        csv_mappings = {
            'MATELEC.csv': ('MATELEC', 'ELEC'),
            'MATGC.csv': ('MATGC', 'GC'),
            'MATINST.csv': ('MATINST', 'INST'),
            'MATPROCES.csv': ('MATPROCES', 'PROC'),
            'MATUY.csv': ('MATUY', 'TUY'),
            'MOELEC.csv': ('MOELEC', 'ELEC'),
            'MOINST.csv': ('MOINST', 'INST'),
            'MOTUY.csv': ('MOTUY', 'TUY'),
            'TRANSPORT.csv': ('TRANSPORT', 'TRANS'),
        }

        total_imported = 0

        for csv_file, (categorie_code, discipline_code) in csv_mappings.items():
            csv_path = os.path.join(csv_directory, csv_file)

            if os.path.exists(csv_path):
                count = self.import_csv_file(csv_path, categorie_code, discipline_code)
                total_imported += count
                self.stdout.write(
                    self.style.SUCCESS(f'✓ {csv_file}: {count} éléments importés')
                )
            else:
                self.stdout.write(
                    self.style.WARNING(f'⚠ Fichier non trouvé: {csv_file}')
                )

        self.stdout.write(
            self.style.SUCCESS(f'\n🎉 Import terminé: {total_imported} éléments au total')
        )

    def create_disciplines(self):
        """Créer les disciplines de base"""
        disciplines = [
            ('ELEC', 'Électricité', '#007bff'),
            ('TUY', 'Tuyauterie', '#28a745'),
            ('INST', 'Instrumentation', '#ffc107'),
            ('GC', 'Génie Civil', '#6c757d'),
            ('PROC', 'Procédés', '#17a2b8'),
            ('TRANS', 'Transport', '#fd7e14'),
        ]

        for code, nom, couleur in disciplines:
            Discipline.objects.get_or_create(
                code=code,
                defaults={'nom': nom, 'couleur': couleur}
            )

    def create_categories(self):
        """Créer les catégories de base"""
        categories = [
            ('MATELEC', 'Matériel Électrique', 'materiel'),
            ('MATGC', 'Matériel Génie Civil', 'materiel'),
            ('MATINST', 'Matériel Instrumentation', 'materiel'),
            ('MATPROCES', 'Études et Procédés', 'etude'),
            ('MATUY', 'Matériel Tuyauterie', 'materiel'),
            ('MOELEC', 'Main d\'œuvre Électrique', 'main_oeuvre'),
            ('MOINST', 'Main d\'œuvre Instrumentation', 'main_oeuvre'),
            ('MOTUY', 'Main d\'œuvre Tuyauterie', 'main_oeuvre'),
            ('TRANSPORT', 'Transport', 'transport'),
        ]

        for code, nom, type_cat in categories:
            Categorie.objects.get_or_create(
                code=code,
                defaults={'nom': nom, 'type_categorie': type_cat}
            )

    def import_csv_file(self, csv_path, categorie_code, discipline_code):
        """Importer un fichier CSV spécifique selon la structure Access"""
        try:
            categorie = Categorie.objects.get(code=categorie_code)
            discipline = Discipline.objects.get(code=discipline_code)
        except (Categorie.DoesNotExist, Discipline.DoesNotExist) as e:
            self.stdout.write(
                self.style.ERROR(f'Erreur: {e}')
            )
            return 0

        count = 0

        with open(csv_path, 'r', encoding='utf-8-sig') as file:
            # Détecter automatiquement le délimiteur
            sample = file.read(1024)
            file.seek(0)

            delimiter = ','
            if ';' in sample:
                delimiter = ';'
            elif '\t' in sample:
                delimiter = '\t'

            reader = csv.DictReader(file, delimiter=delimiter)

            with transaction.atomic():
                for row_num, row in enumerate(reader, 1):
                    try:
                        # Nettoyer les données
                        clean_row = {k.strip(): v.strip() if v else '' for k, v in row.items()}

                        # Mapping spécifique selon vos structures Access
                        designation = self.get_designation(clean_row, categorie_code)
                        if not designation:
                            continue

                        caracteristiques = self.get_caracteristiques(clean_row, categorie_code)
                        prix_unitaire = self.get_prix_unitaire(clean_row)
                        unite = self.get_unite(clean_row)
                        numero = self.get_numero(clean_row, categorie_code)

                        # Créer l'élément
                        Element.objects.create(
                            numero=numero or '',
                            designation=designation,
                            caracteristiques=caracteristiques or '',
                            prix_unitaire=prix_unitaire,
                            unite=unite,
                            categorie=categorie,
                            discipline=discipline,
                            actif=True
                        )

                        count += 1

                    except Exception as e:
                        self.stdout.write(
                            self.style.ERROR(f'Erreur ligne {row_num}: {e}')
                        )
                        continue

        return count

    def get_designation(self, row, categorie_code):
        """Récupérer la désignation selon la catégorie"""
        possible_names = ['Désignation', 'Designation']
        return self.get_field_value(row, possible_names)

    def get_caracteristiques(self, row, categorie_code):
        """Récupérer les caractéristiques selon la catégorie"""
        # Selon vos données Access
        mappings = {
            'MATELEC': ['Caractéristiques', 'Caracteristiques'],
            'MATGC': ['Caractéristiques', 'Caracteristiques'],
            'MATINST': ['Tâches', 'Taches', 'Caractéristiques'],
            'MATPROCES': ['Tâches', 'Taches'],
            'MATUY': ['Diamètre', 'Débit/Epaiss', 'Schedule/Sé', 'Matière'],
            'MOELEC': ['Caractéristiques', 'Observation'],
            'MOTUY': ['Diamètre', 'Matière', 'Caractéristiques'],
            'TRANSPORT': ['Type_transp', 'Objet_transp'],
        }

        possible_names = mappings.get(categorie_code, ['Caractéristiques'])

        # Pour certaines tables, combiner plusieurs champs
        if categorie_code == 'MATUY':
            diametre = self.get_field_value(row, ['Diamètre'])
            matiere = self.get_field_value(row, ['Matière'])
            schedule = self.get_field_value(row, ['Schedule/Sé'])
            parts = [p for p in [diametre, matiere, schedule] if p]
            return ' - '.join(parts)
        elif categorie_code == 'TRANSPORT':
            type_transp = self.get_field_value(row, ['Type_transp'])
            objet_transp = self.get_field_value(row, ['Objet_transp'])
            parts = [p for p in [type_transp, objet_transp] if p]
            return ' - '.join(parts)
        else:
            return self.get_field_value(row, possible_names)

    def get_prix_unitaire(self, row):
        """Récupérer et nettoyer le prix unitaire"""
        prix_str = self.get_field_value(row, [
            'Prix Unitaire', 'Prix_Unitaire', 'PrixUnitaire',
            'Prix de Base', 'Prix', 'Coût_journa', 'Tarif_par_kr'
        ])

        if prix_str:
            # Nettoyer le prix
            prix_str = str(prix_str).replace('CFA', '').replace('€', '').replace(' ', '').replace(',', '.')
            # Enlever les caractères non numériques sauf le point
            prix_str = ''.join(c for c in prix_str if c.isdigit() or c == '.')
            try:
                return Decimal(prix_str) if prix_str else Decimal('0')
            except:
                return Decimal('0')
        return Decimal('0')

    def get_unite(self, row):
        """Récupérer l'unité"""
        unite = self.get_field_value(row, ['Unité', 'Unite', 'U'])
        return self.map_unite(unite)

    def get_numero(self, row, categorie_code):
        """Récupérer le numéro selon la catégorie"""
        mappings = {
            'MATELEC': ['NumElect'],
            'MATGC': ['NumGC'],
            'MATINST': ['NumInstr'],
            'MATPROCES': ['NumMATPRO'],
            'MATUY': ['NumTUY'],
            'MOELEC': ['NumMOELEI'],
            'MOINST': ['NumMOInst'],
            'MOTUY': ['NumTUY'],
            'TRANSPORT': ['Numtransp'],
        }

        possible_names = mappings.get(categorie_code, ['Numero', 'Code', 'Ref'])
        return self.get_field_value(row, possible_names)

    def get_field_value(self, row, possible_names):
        """Récupérer une valeur selon plusieurs noms de colonnes possibles"""
        for name in possible_names:
            if name in row and row[name]:
                return row[name]
        return ''

    def map_unite(self, unite_access):
        """Mapper les unités Access vers Django"""
        mapping = {
            'u': 'u', 'U': 'u', 'unité': 'u', 'piece': 'u',
            'ml': 'ml', 'mL': 'ml', 'm': 'ml', 'metre': 'ml',
            'ens': 'ens', 'Ens': 'ens', 'ensemble': 'ens',
            'h': 'h', 'H': 'h', 'heure': 'h', 'heures': 'h',
            'j': 'j', 'J': 'j', 'jour': 'j', 'jours': 'j',
            'ff': 'ff', 'forfait': 'ff', 'Forfait': 'ff',
            'kg': 'kg', 'Kg': 'kg', 'kilogramme': 'kg',
            'm2': 'm2', 'm²': 'm2', 'metre_carre': 'm2',
            'm3': 'm3', 'm³': 'm3', 'metre_cube': 'm3',
        }

        return mapping.get(unite_access, 'u')