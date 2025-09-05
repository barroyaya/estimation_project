# migration_bdd_prix_fk.py - Import "BDD prix.xlsx" compatible avec Unite (FK)
import os
import django
import pandas as pd
from decimal import Decimal, InvalidOperation

# ⚙️ Paramètres Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'estimation_project.settings')
django.setup()

from estimation.models import Categorie, Discipline, Element, Unite  # <-- FK Unite

# Unités supportées (code -> (libellé, symbole))
UNITS = {
    "u":   ("Unité", "u"),
    "ml":  ("Mètre linéaire", "m"),
    "m2":  ("m²", "m²"),
    "m3":  ("m³", "m³"),
    "kg":  ("Kilogramme", "kg"),
    "h":   ("Heure", "h"),
    "j":   ("Jour", "j"),
    "ens": ("Ensemble", "ens"),
    "ff":  ("Forfait", "ff"),
}

def ensure_units():
    """Crée (ou récupère) les unités référentielles et renvoie un dict code -> Unite instance."""
    code_to_obj = {}
    for code, (libelle, symbole) in UNITS.items():
        obj, _ = Unite.objects.get_or_create(code=code, defaults={
            "libelle": libelle, "symbole": symbole, "is_active": True
        })
        code_to_obj[code] = obj
    return code_to_obj


class BddPrixMigrator:

    def __init__(self):
        self.units = ensure_units()
        self.create_base_data()
        print("✓ Données de base créées (disciplines, catégories, unités)")

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
            Discipline.objects.get_or_create(code=code, defaults={'nom': nom, 'couleur': couleur})

        # Catégories
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
            Categorie.objects.get_or_create(code=code, defaults={'nom': nom, 'type_categorie': type_cat})

    def clean_price(self, prix_str):
        """Nettoyer et convertir les prix"""
        if pd.isna(prix_str) or prix_str == '' or prix_str is None:
            return Decimal('0')
        prix_str = str(prix_str).replace(' ', '').replace(',', '.')
        prix_clean = ''.join(c for c in prix_str if c.isdigit() or c == '.')
        try:
            return Decimal(prix_clean) if prix_clean else Decimal('0')
        except InvalidOperation:
            return Decimal('0')

    def map_unite_code(self, unite):
        """Convertit la saisie en code canonique ('ml', 'm2', ...)"""
        if pd.isna(unite) or unite == '':
            return 'u'
        u = str(unite).strip().lower()
        mapping = {
            'u': 'u', 'unité': 'u', 'unite': 'u',
            'ml': 'ml', 'm': 'ml', 'metre': 'ml', 'mètre': 'ml',
            'ens': 'ens', 'ensemble': 'ens',
            'h': 'h', 'heure': 'h', 'heures': 'h',
            'j': 'j', 'jour': 'j', 'jours': 'j',
            'ff': 'ff', 'forfait': 'ff',
            'kg': 'kg', 'kilogramme': 'kg',
            'm2': 'm2', 'm²': 'm2',
            'm3': 'm3', 'm³': 'm3',
        }
        return mapping.get(u, 'u')

    def unit_obj(self, raw_unite):
        """Renvoie l'objet Unite correspondant (fallback sur 'u')."""
        code = self.map_unite_code(raw_unite)
        return self.units.get(code) or self.units['u']

    # ————— IMPORTS DES FEUILLES —————

    def import_mat_tuy(self, df):
        categorie = Categorie.objects.get(code='MAT_TUY')
        discipline = Discipline.objects.get(code='TUY')
        count = 0
        for i, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                carac_parts = []
                diametre = row.get('Diamètre', '')
                if pd.notna(diametre) and str(diametre) not in ('nan', ''):
                    carac_parts.append(f"Diamètre: {diametre}")
                debit_epaiss = row.get('Débit/Epaisseur', '')
                if pd.notna(debit_epaiss) and str(debit_epaiss) not in ('nan', ''):
                    carac_parts.append(f"Débit/Épaisseur: {debit_epaiss}")
                schedule = row.get('Schédule/Série', '')
                if pd.notna(schedule) and str(schedule) not in ('nan', ''):
                    carac_parts.append(f"Schédule: {schedule}")
                matiere = row.get('Matière', '')
                if pd.notna(matiere) and str(matiere) not in ('nan', ''):
                    carac_parts.append(f"Matière: {matiere}")
                caracteristiques = ' - '.join(carac_parts)

                prix_unitaire = self.clean_price(row.get('Prix Unitaire', row.get('Prix de Base', 0)))
                numero = str(row.get('NumMatrl', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=self.unit_obj(row.get('Unité', 'u')),   # <-- FK Unite
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1
            except Exception as e:
                print(f"Erreur Mat TUY ligne {i + 2}: {e}")
        return count

    def import_gc(self, df):
        categorie = Categorie.objects.get(code='MAT_GC')
        discipline = Discipline.objects.get(code='GC')
        count = 0
        for i, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                carac_parts = []
                caracteristiques = str(row.get('Caractéristiques', '')).strip()
                if caracteristiques and caracteristiques != 'nan':
                    carac_parts.append(caracteristiques)
                poids = row.get('Poids(enkg)', '')
                if pd.notna(poids) and str(poids) not in ('nan', ''):
                    carac_parts.append(f"Poids: {poids} kg")
                caracteristiques_final = ' - '.join(carac_parts)

                prix_unitaire = self.clean_price(row.get('Prix Unitaire', 0))
                numero = str(row.get('NumGC', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques_final,
                    prix_unitaire=prix_unitaire,
                    unite=self.unit_obj(row.get('Unité', 'u')),   # <-- FK Unite
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1
            except Exception as e:
                print(f"Erreur GC ligne {i + 2}: {e}")
        return count

    def import_mat_elec(self, df):
        categorie = Categorie.objects.get(code='MAT_ELEC')
        discipline = Discipline.objects.get(code='ELEC')
        count = 0
        for i, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                caracteristiques = str(row.get('Caractéristiques', '')).strip()
                if caracteristiques == 'nan':
                    caracteristiques = ''

                prix_unitaire = self.clean_price(row.get('Prix Unitaire', 0))
                numero = str(row.get('NumElect', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=self.unit_obj(row.get('Unité', 'u')),   # <-- FK Unite
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1
            except Exception as e:
                print(f"Erreur MAT ELEC ligne {i + 2}: {e}")
        return count

    def import_mat_inst(self, df):
        categorie = Categorie.objects.get(code='MAT_INST')
        discipline = Discipline.objects.get(code='INST')
        count = 0
        for i, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                caracteristiques = str(row.get('Caractéristiques', '')).strip()
                if caracteristiques == 'nan':
                    caracteristiques = ''

                prix_unitaire = self.clean_price(row.get('Prix Unitaire', 0))
                numero = str(row.get('NumInstr', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=self.unit_obj(row.get('Unité', 'u')),   # <-- FK Unite
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1
            except Exception as e:
                print(f"Erreur MAT INST ligne {i + 2}: {e}")
        return count

    def import_mo_inst(self, df):
        categorie = Categorie.objects.get(code='MO_INST')
        discipline = Discipline.objects.get(code='INST')
        count = 0
        for i, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue
                caracteristiques = str(row.get('OBSERVATIONS', '')).strip()
                if caracteristiques == 'nan':
                    caracteristiques = ''
                prix_unitaire = self.clean_price(row.get('Prix Unitaire', 0))
                numero = str(row.get('NumMOInstr', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=self.unit_obj(row.get('Unité', 'h')),   # <-- FK Unite
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1
            except Exception as e:
                print(f"Erreur MO INST ligne {i + 2}: {e}")
        return count

    def import_mo_elec(self, df):
        categorie = Categorie.objects.get(code='MO_ELEC')
        discipline = Discipline.objects.get(code='ELEC')
        count = 0
        for i, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue
                caracteristiques = str(row.get('Observations', '')).strip()
                if caracteristiques == 'nan':
                    caracteristiques = ''
                prix_unitaire = self.clean_price(row.get('Prix unitaire7', row.get('Prix Unitaire', 0)))
                numero = str(row.get('NumMOElect', '')).strip()
                if numero == 'nan':
                    numero = ''

                Element.objects.create(
                    numero=numero,
                    designation=designation,
                    caracteristiques=caracteristiques,
                    prix_unitaire=prix_unitaire,
                    unite=self.unit_obj(row.get('Unité', 'h')),   # <-- FK Unite
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1
            except Exception as e:
                print(f"Erreur MO ELEC ligne {i + 2}: {e}")
        return count

    def import_mo_tuy(self, df):
        categorie = Categorie.objects.get(code='MO_TUY')
        discipline = Discipline.objects.get(code='TUY')
        count = 0
        for i, row in df.iterrows():
            try:
                designation = str(row.get('Désignation', '')).strip()
                if not designation or designation == 'nan':
                    continue

                carac_parts = []
                for field in ['Observation', 'OBSERVATIONS', 'Diamètre', 'Matière']:
                    if field in row and pd.notna(row[field]) and str(row[field]) != 'nan':
                        value = str(row[field]).strip()
                        if value:
                            carac_parts.append(f"{field}: {value}")
                caracteristiques = ' - '.join(carac_parts)

                prix_unitaire = Decimal('0')
                for prix_field in ['Prix Unitaire', 'Prix unitaire', 'Prix de Base', 'prix_unitaire']:
                    if prix_field in row:
                        prix_unitaire = self.clean_price(row[prix_field])
                        break

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
                    unite=self.unit_obj(row.get('Unité', 'h')),   # <-- FK Unite
                    categorie=categorie,
                    discipline=discipline,
                    actif=True
                )
                count += 1
            except Exception as e:
                print(f"Erreur MO TUY ligne {i + 2}: {e}")
        return count

    # ————— Driver —————

    def import_from_excel(self, excel_path):
        print(f"📂 Lecture du fichier: {excel_path}")
        try:
            excel_data = pd.read_excel(excel_path, sheet_name=None)
        except Exception as e:
            print(f"❌ Erreur lors de la lecture Excel: {e}")
            return 0
        print(f"✓ {len(excel_data)} onglets détectés")

        import_methods = {
            'Mat TUY': self.import_mat_tuy,
            'GC': self.import_gc,
            'MAT ELEC': self.import_mat_elec,
            'MAT INST': self.import_mat_inst,
            'MO INST': self.import_mo_inst,
            'MO ELEC': self.import_mo_elec,
            'MO TUY': self.import_mo_tuy,
        }

        total = 0
        for sheet_name, df in excel_data.items():
            if sheet_name in import_methods:
                print(f"\n📋 Onglet: {sheet_name} ({len(df)} lignes)")
                try:
                    cnt = import_methods[sheet_name](df)
                    total += cnt
                    print(f"   ✅ {cnt} éléments importés")
                except Exception as e:
                    print(f"   ❌ Erreur import {sheet_name}: {e}")
            else:
                print(f"⚠️ Onglet ignoré: {sheet_name}")

        print(f"\n🎉 Import terminé — total: {total} éléments")
        self.afficher_resume()
        return total

    def afficher_resume(self):
        print(f"\n📈 RÉSUMÉ DE L'IMPORT")
        print("=" * 50)
        for cat in Categorie.objects.all():
            count = Element.objects.filter(categorie=cat).count()
            if count > 0:
                print(f"{cat.nom:30} {count:4} éléments")
        total = Element.objects.count()
        print("=" * 50)
        print(f"{'TOTAL GÉNÉRAL':30} {total:4} éléments")
        print("\n💰 Échantillon prix:")
        for elem in Element.objects.all()[:5]:
            print(f"  • {elem.designation[:40]:40} {elem.prix_unitaire:>10} CFA")


def main():
    print("=" * 60)
    print("🚀 IMPORT BDD_prix.xlsx (avec Unite FK)")
    print("=" * 60)

    migrator = BddPrixMigrator()
    excel_path = input("Chemin vers 'BDD prix.xlsx': ").strip().strip('"') or "BDD_prix.xlsx"

    if not os.path.exists(excel_path):
        print(f"❌ Fichier non trouvé: {excel_path}")
        return

    go = input("⚠️ Importer maintenant ? (oui/non): ").lower()
    if go not in ("oui", "o", "yes", "y"):
        print("Import annulé.")
        return

    migrator.import_from_excel(excel_path)
    print("\n✅ Terminé.")

if __name__ == '__main__':
    main()
