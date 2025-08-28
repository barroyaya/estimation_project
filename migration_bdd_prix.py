# migration_bdd_prix.py - Script spécialisé pour votre fichier "BDD prix.xlsx"
import os
import django
import pandas as pd
from decimal import Decimal, InvalidOperation

# Configuration Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'estimation_project.settings')
django.setup()

from estimation.models import Categorie, Discipline, Element


class BddPrixMigrator:

    def __init__(self):
        self.create_base_data()
        print("✓ Données de base créées (disciplines et catégories)")

    def create_base_data(self):
        """Créer les disciplines et catégories correspondant à vos onglets"""

        # Disciplines
        disciplines = [
            ('ELEC', 'Électricité', '#007bff'),
            ('TUY', 'Tuyauterie', '#28a745'),
            ('INST', 'Instrumentation', '#ffc107'),
            ('GC', 'Génie Civil', '#6c757d'),
        ]

        for code, nom, couleur in disciplines:
            Discipline.objects.get_or_create(
                code=code,
                defaults={'nom': nom, 'couleur': couleur}
            )

        # Catégories correspondant à vos onglets
        categories = [
            ('MAT_TUY', 'Matériel Tuyauterie', 'materiel'),
            ('MAT_GC', 'Matériel Génie Civil', 'materiel'),
            ('MAT_ELEC', 'Matériel Électrique', 'materiel'),
            ('MAT_INST', 'Matériel Instrumentation', 'materiel'),
            ('MO_INST', 'Main d\'œuvre Instrumentation', 'main_oeuvre'),
            ('MO_ELEC', 'Main d\'œuvre Électrique', 'main_oeuvre'),
            ('MO_TUY', 'Main d\'œuvre Tuyauterie', 'main_oeuvre'),
        ]

        for code, nom, type_cat in categories:
            Categorie.objects.get_or_create(
                code=code,
                defaults={'nom': nom, 'type_categorie': type_cat}
            )

    def clean_price(self, prix_str):
        """Nettoyer et convertir les prix"""
        if pd.isna(prix_str) or prix_str == '' or prix_str is None:
            return Decimal('0')

        prix_str = str(prix_str)
        # Enlever espaces et autres caractères
        prix_str = prix_str.replace(' ', '').replace(',', '.')

        # Garder seulement chiffres et point
        prix_clean = ''.join(c for c in prix_str if c.isdigit() or c == '.')

        try:
            return Decimal(prix_clean) if prix_clean else Decimal('0')
        except InvalidOperation:
            return Decimal('0')

    def map_unite(self, unite):
        """Mapper les unités vers Django"""
        if pd.isna(unite) or unite == '':
            return 'u'

        unite = str(unite).strip().lower()

        mapping = {
            'u': 'u', 'unité': 'u',
            'ml': 'ml', 'm': 'ml', 'metre': 'ml', 'mètre': 'ml',
            'ens': 'ens', 'ensemble': 'ens',
            'h': 'h', 'heure': 'h', 'heures': 'h',
            'j': 'j', 'jour': 'j', 'jours': 'j',
            'ff': 'ff', 'forfait': 'ff',
            'kg': 'kg', 'kilogramme': 'kg',
            'm2': 'm2', 'm²': 'm2',
            'm3': 'm3', 'm³': 'm3',
        }

        return mapping.get(unite, 'u')

    def import_mat_tuy(self, df):
        """Importer l'onglet Mat TUY - Matériel Tuyauterie"""
        categorie = Categorie.objects.get(code='MAT_TUY')
        discipline = Discipline.objects.get(code='TUY')

        count = 0
        for _, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                # Construire les caractéristiques à partir de plusieurs colonnes
                carac_parts = []

                diametre = row.get('Diamètre', '')
                if pd.notna(diametre) and str(diametre) != 'nan' and diametre != '':
                    carac_parts.append(f"Diamètre: {diametre}")

                debit_epaiss = row.get('Débit/Epaisseur', '')
                if pd.notna(debit_epaiss) and str(debit_epaiss) != 'nan' and debit_epaiss != '':
                    carac_parts.append(f"Débit/Épaisseur: {debit_epaiss}")

                schedule = row.get('Schédule/Série', '')
                if pd.notna(schedule) and str(schedule) != 'nan' and schedule != '':
                    carac_parts.append(f"Schédule: {schedule}")

                matiere = row.get('Matière', '')
                if pd.notna(matiere) and str(matiere) != 'nan' and matiere != '':
                    carac_parts.append(f"Matière: {matiere}")

                caracteristiques = ' - '.join(carac_parts)

                # Prix unitaire ou prix de base
                prix_unitaire = self.clean_price(
                    row.get('Prix Unitaire', row.get('Prix de Base', 0))
                )

                unite = self.map_unite(row.get('Unité', 'u'))
                numero = str(row.get('NumMatrl', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=unite,
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1

            except Exception as e:
                print(f"Erreur Mat TUY ligne {_ + 2}: {e}")
                continue

        return count

    def import_gc(self, df):
        """Importer l'onglet GC - Génie Civil"""
        categorie = Categorie.objects.get(code='MAT_GC')
        discipline = Discipline.objects.get(code='GC')

        count = 0
        for _, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                # Caractéristiques avec poids si disponible
                caracteristiques = str(row.get('Caractéristiques', '')).strip()
                poids = row.get('Poids(enkg)', '')

                carac_parts = []
                if caracteristiques != 'nan' and caracteristiques:
                    carac_parts.append(caracteristiques)

                if pd.notna(poids) and str(poids) != 'nan' and poids != '':
                    carac_parts.append(f"Poids: {poids} kg")

                caracteristiques_final = ' - '.join(carac_parts)

                prix_unitaire = self.clean_price(row.get('Prix Unitaire', 0))
                unite = self.map_unite(row.get('Unité', 'u'))
                numero = str(row.get('NumGC', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques_final,
                    prix_unitaire=prix_unitaire,
                    unite=unite,
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1

            except Exception as e:
                print(f"Erreur GC ligne {_ + 2}: {e}")
                continue

        return count

    def import_mat_elec(self, df):
        """Importer l'onglet MAT ELEC - Matériel Électrique"""
        categorie = Categorie.objects.get(code='MAT_ELEC')
        discipline = Discipline.objects.get(code='ELEC')

        count = 0
        for _, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                caracteristiques = str(row.get('Caractéristiques', '')).strip()
                if caracteristiques == 'nan':
                    caracteristiques = ''

                prix_unitaire = self.clean_price(row.get('Prix Unitaire', 0))
                unite = self.map_unite(row.get('Unité', 'u'))
                numero = str(row.get('NumElect', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=unite,
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1

            except Exception as e:
                print(f"Erreur MAT ELEC ligne {_ + 2}: {e}")
                continue

        return count

    def import_mat_inst(self, df):
        """Importer l'onglet MAT INST - Matériel Instrumentation"""
        categorie = Categorie.objects.get(code='MAT_INST')
        discipline = Discipline.objects.get(code='INST')

        count = 0
        for _, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                caracteristiques = str(row.get('Caractéristiques', '')).strip()
                if caracteristiques == 'nan':
                    caracteristiques = ''

                prix_unitaire = self.clean_price(row.get('Prix Unitaire', 0))
                unite = self.map_unite(row.get('Unité', 'u'))
                numero = str(row.get('NumInstr', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=unite,
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1

            except Exception as e:
                print(f"Erreur MAT INST ligne {_ + 2}: {e}")
                continue

        return count

    def import_mo_inst(self, df):
        """Importer l'onglet MO INST - Main d'œuvre Instrumentation"""
        categorie = Categorie.objects.get(code='MO_INST')
        discipline = Discipline.objects.get(code='INST')

        count = 0
        for _, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                # Pour la main d'œuvre, utiliser OBSERVATIONS comme caractéristiques
                caracteristiques = str(row.get('OBSERVATIONS', '')).strip()
                if caracteristiques == 'nan':
                    caracteristiques = ''

                prix_unitaire = self.clean_price(row.get('Prix Unitaire', 0))
                unite = self.map_unite(row.get('Unité', 'h'))  # Main d'œuvre souvent en heures
                numero = str(row.get('NumMOInstr', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=unite,
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1

            except Exception as e:
                print(f"Erreur MO INST ligne {_ + 2}: {e}")
                continue

        return count

    def import_mo_elec(self, df):
        """Importer l'onglet MO ELEC - Main d'œuvre Électrique"""
        categorie = Categorie.objects.get(code='MO_ELEC')
        discipline = Discipline.objects.get(code='ELEC')

        count = 0
        for _, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                # Utiliser Observations comme caractéristiques
                caracteristiques = str(row.get('Observations', '')).strip()
                if caracteristiques == 'nan':
                    caracteristiques = ''

                # Le prix peut être dans 'Prix unitaire7'
                prix_unitaire = self.clean_price(
                    row.get('Prix unitaire7', row.get('Prix Unitaire', 0))
                )

                unite = self.map_unite(row.get('Unité', 'h'))  # Main d'œuvre souvent en heures
                numero = str(row.get('NumMOElect', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=unite,
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1

            except Exception as e:
                print(f"Erreur MO ELEC ligne {_ + 2}: {e}")
                continue

        return count

    def import_mo_tuy(self, df):
        """Importer l'onglet MO TUY - Main d'œuvre Tuyauterie"""
        categorie = Categorie.objects.get(code='MO_TUY')
        discipline = Discipline.objects.get(code='TUY')

        count = 0
        for _, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                # Construire caractéristiques avec les données disponibles
                carac_parts = []

                # Essayer différents champs pour les caractéristiques
                for field in ['Observation', 'OBSERVATIONS', 'Diamètre', 'Matière']:
                    if field in row and pd.notna(row[field]) and str(row[field]) != 'nan':
                        value = str(row[field]).strip()
                        if value:
                            carac_parts.append(f"{field}: {value}")

                caracteristiques = ' - '.join(carac_parts)

                # Essayer plusieurs champs pour le prix
                prix_unitaire = Decimal('0')
                for prix_field in ['Prix Unitaire', 'Prix unitaire', 'Prix de Base', 'prix_unitaire']:
                    if prix_field in row:
                        prix_unitaire = self.clean_price(row[prix_field])
                        break

                unite = self.map_unite(row.get('Unité', 'h'))

                # Essayer plusieurs champs pour le numéro
                numero = ''
                for num_field in ['NumMOTuy', 'NumMO', 'Numero']:
                    if num_field in row and pd.notna(row[num_field]):
                        numero = str(row[num_field]).strip()
                        break

                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=unite,
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1

            except Exception as e:
                print(f"Erreur MO TUY ligne {_ + 2}: {e}")
                continue

        return count

    def import_from_excel(self, excel_path):
        """Importer directement votre fichier BDD prix.xlsx"""

        print(f"📂 Lecture du fichier: {excel_path}")

        # Lire tous les onglets
        try:
            excel_data = pd.read_excel(excel_path, sheet_name=None)
        except Exception as e:
            print(f"❌ Erreur lors de la lecture du fichier Excel: {e}")
            return

        print(f"✓ {len(excel_data)} onglets détectés")

        # Mapping des onglets vers les méthodes d'import
        import_methods = {
            'Mat TUY': self.import_mat_tuy,
            'GC': self.import_gc,
            'MAT ELEC': self.import_mat_elec,
            'MAT INST': self.import_mat_inst,
            'MO INST': self.import_mo_inst,
            'MO ELEC': self.import_mo_elec,
            'MO TUY': self.import_mo_tuy,
        }

        total_imported = 0

        for sheet_name, df in excel_data.items():
            if sheet_name in import_methods:
                print(f"\n📋 Import de l'onglet: {sheet_name}")
                print(f"   Lignes détectées: {len(df)}")

                try:
                    count = import_methods[sheet_name](df)
                    total_imported += count
                    print(f"   ✅ {count} éléments importés avec succès")

                except Exception as e:
                    print(f"   ❌ Erreur lors de l'import: {e}")

            else:
                print(f"⚠️  Onglet ignoré: {sheet_name}")

        print(f"\n🎉 Import terminé!")
        print(f"📊 Total importé: {total_imported} éléments")

        # Afficher un résumé
        self.afficher_resume()

        return total_imported

    def afficher_resume(self):
        """Afficher un résumé des données importées"""
        print(f"\n📈 RÉSUMÉ DE L'IMPORT:")
        print(f"{'=' * 50}")

        for cat in Categorie.objects.all():
            count = Element.objects.filter(categorie=cat).count()
            if count > 0:
                print(f"{cat.nom:30} {count:4} éléments")

        total = Element.objects.count()
        print(f"{'=' * 50}")
        print(f"{'TOTAL GÉNÉRAL':30} {total:4} éléments")

        print(f"\n💰 Quelques exemples de prix:")
        for elem in Element.objects.all()[:5]:
            print(f"  • {elem.designation[:40]:40} {elem.prix_unitaire:>10} CFA")


def main():
    """Fonction principale"""
    print("=" * 60)
    print("🚀 MIGRATION DE LA BASE DE DONNÉES BDD PRIX.XLSX")
    print("=" * 60)

    migrator = BddPrixMigrator()

    # Chemin vers votre fichier Excel
    excel_path = input("Chemin vers le fichier 'BDD prix.xlsx': ").strip().strip('"')

    if not excel_path:
        excel_path = "BDD_prix.xlsx"  # Par défaut dans le répertoire courant

    if not os.path.exists(excel_path):
        print(f"❌ Fichier non trouvé: {excel_path}")
        return

    # Confirmer avant l'import
    response = input(f"\n⚠️  Ceci va importer les données dans Django. Continuer? (oui/non): ").lower()
    if response not in ['oui', 'o', 'yes', 'y']:
        print("Import annulé.")
        return

    # Lancer l'import
    migrator.import_from_excel(excel_path)

    print(f"\n✅ Migration terminée!")
    print(f"👉 Vous pouvez maintenant lancer votre serveur Django:")
    print(f"   python manage.py runserver")


if __name__ == '__main__':
    main()