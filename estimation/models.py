# estimation/models.py

import json
from django.db import models
from django.http import JsonResponse


# -----------------------------
# Clients / Projets / Référentiels
# -----------------------------

class Client(models.Model):
    """Modèle pour les clients de l'entreprise"""
    nom = models.CharField(max_length=100)
    email = models.EmailField(unique=True)
    telephone = models.CharField(max_length=20, blank=True)
    entreprise = models.CharField(max_length=200, blank=True)
    adresse = models.TextField(blank=True)
    date_creation = models.DateTimeField(auto_now_add=True)
    actif = models.BooleanField(default=True)

    # Champs d'authentification (hash attendu)
    password = models.CharField(max_length=128)
    last_login = models.DateTimeField(null=True, blank=True)

    def __str__(self):
        return f"{self.nom} ({self.email})"

    def check_password(self, raw_password):
        from django.contrib.auth.hashers import check_password
        return check_password(raw_password, self.password)

    def set_password(self, raw_password):
        from django.contrib.auth.hashers import make_password
        self.password = make_password(raw_password)

    class Meta:
        verbose_name = "Client"
        verbose_name_plural = "Clients"


class Projet(models.Model):
    nom = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    client = models.ForeignKey(Client, on_delete=models.CASCADE, related_name='projets')
    client_nom = models.CharField(max_length=200, blank=True)  # compatibilité
    date_creation = models.DateTimeField(auto_now_add=True)
    actif = models.BooleanField(default=True)

    def __str__(self):
        return self.nom

    class Meta:
        verbose_name = "Projet"
        verbose_name_plural = "Projets"


class Categorie(models.Model):
    TYPES_CHOICES = [
        ('materiel', 'Matériel'),
        ('main_oeuvre', 'Main d\'œuvre'),
        ('transport', 'Transport'),
        ('etude', 'Étude'),
    ]

    nom = models.CharField(max_length=100)
    type_categorie = models.CharField(max_length=20, choices=TYPES_CHOICES)
    code = models.CharField(max_length=20, unique=True)
    description = models.TextField(blank=True)

    def __str__(self):
        return self.nom

    class Meta:
        verbose_name = "Catégorie"
        verbose_name_plural = "Catégories"


class Discipline(models.Model):
    nom = models.CharField(max_length=100)
    code = models.CharField(max_length=20, unique=True)
    couleur = models.CharField(max_length=7, default='#007bff')  # hex

    def __str__(self):
        return self.nom

    class Meta:
        verbose_name = "Discipline"
        verbose_name_plural = "Disciplines"


# -----------------------------
# Unités (nouveau modèle pour avoir le bouton +)
# -----------------------------

class Unite(models.Model):
    code = models.CharField(max_length=10, unique=True)        # ex: 'm2', 'kg', 'u'
    libelle = models.CharField(max_length=100)                 # ex: 'm²', 'Kilogramme'
    symbole = models.CharField(max_length=20, blank=True)      # ex: 'm²', 'kg'
    is_active = models.BooleanField(default=True)

    class Meta:
        verbose_name = "Unité"
        verbose_name_plural = "Unités"
        ordering = ["libelle"]

    def __str__(self):
        return f"{self.libelle} ({self.code})" if self.code else self.libelle


# -----------------------------
# Eléments & Demandes
# -----------------------------

class Element(models.Model):
    numero = models.CharField(max_length=50, blank=True)
    designation = models.CharField(max_length=300)
    caracteristiques = models.TextField(blank=True)
    prix_unitaire = models.DecimalField(max_digits=15, decimal_places=2)
    # ⬇️ ancien CharField(choices) remplacé par FK pour afficher le bouton +
    unite = models.ForeignKey(Unite, on_delete=models.PROTECT, related_name='elements')
    categorie = models.ForeignKey(Categorie, on_delete=models.CASCADE)
    discipline = models.ForeignKey(Discipline, on_delete=models.CASCADE)
    actif = models.BooleanField(default=True)

    def __str__(self):
        return f"{self.designation} - {self.prix_unitaire} CFA"

    class Meta:
        verbose_name = "Élément"
        verbose_name_plural = "Éléments"
        ordering = ['numero', 'designation']


class DemandeElement(models.Model):
    STATUT_CHOICES = [
        ('en_attente', 'En attente'),
        ('approuve', 'Approuvé'),
        ('rejete', 'Rejeté'),
    ]

    projet = models.ForeignKey(Projet, on_delete=models.CASCADE)
    categorie = models.ForeignKey(Categorie, on_delete=models.CASCADE)
    discipline = models.ForeignKey(Discipline, on_delete=models.CASCADE)

    # Saisie client
    designation = models.CharField(max_length=300)
    caracteristiques = models.TextField(blank=True)
    # ⬇️ FK pour cohérence et bouton +
    unite = models.ForeignKey(Unite, on_delete=models.PROTECT, related_name='demandes')
    quantite = models.DecimalField(max_digits=10, decimal_places=2, default=1)

    # Validation admin
    statut = models.CharField(max_length=20, choices=STATUT_CHOICES, default='en_attente')
    prix_unitaire_admin = models.DecimalField(max_digits=15, decimal_places=2, null=True, blank=True)
    commentaire_admin = models.TextField(blank=True, help_text="Commentaire de l'administrateur")

    # Horodatage
    date_demande = models.DateTimeField(auto_now_add=True)
    date_validation = models.DateTimeField(null=True, blank=True)

    @property
    def cout_total(self):
        if self.statut == 'approuve' and self.prix_unitaire_admin:
            return self.quantite * self.prix_unitaire_admin
        return 0

    @property
    def est_integrable(self):
        return self.statut == 'approuve' and self.prix_unitaire_admin is not None

    def __str__(self):
        return f"{self.designation} - {self.get_statut_display()}"

    class Meta:
        verbose_name = "Demande d'élément"
        verbose_name_plural = "Demandes d'éléments"
        ordering = ['-date_demande']


class EstimationElement(models.Model):
    projet = models.ForeignKey(Projet, on_delete=models.CASCADE)
    element = models.ForeignKey(Element, on_delete=models.CASCADE, null=True, blank=True)
    demande_element = models.ForeignKey(DemandeElement, on_delete=models.CASCADE, null=True, blank=True)
    quantite = models.DecimalField(max_digits=10, decimal_places=2, default=1)
    prix_unitaire_fixe = models.DecimalField(max_digits=15, decimal_places=2, null=True, blank=True)
    date_ajout = models.DateTimeField(auto_now_add=True)

    @property
    def prix_unitaire_utilise(self):
        if self.prix_unitaire_fixe:
            return self.prix_unitaire_fixe
        elif self.demande_element and self.demande_element.prix_unitaire_admin:
            return self.demande_element.prix_unitaire_admin
        elif self.element:
            return self.element.prix_unitaire
        return 0

    @property
    def cout_total(self):
        return self.quantite * self.prix_unitaire_utilise

    @property
    def designation(self):
        if self.demande_element:
            return self.demande_element.designation
        elif self.element:
            return self.element.designation
        return "Élément inconnu"

    @property
    def caracteristiques(self):
        if self.demande_element:
            return self.demande_element.caracteristiques
        elif self.element:
            return self.element.caracteristiques
        return ""

    @property
    def unite_display(self):
        if self.demande_element and self.demande_element.unite:
            return self.demande_element.unite.libelle
        elif self.element and self.element.unite:
            return self.element.unite.libelle
        return "Unité"

    def __str__(self):
        return f"{self.projet.nom} - {self.designation}"

    class Meta:
        verbose_name = "Élément d'estimation"
        verbose_name_plural = "Éléments d'estimation"


class EstimationSummary(models.Model):
    projet = models.OneToOneField(Projet, on_delete=models.CASCADE)
    cout_total_materiel = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    cout_total_main_oeuvre = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    cout_total_transport = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    cout_total_etude = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    tva_taux = models.DecimalField(max_digits=5, decimal_places=2, default=18)
    cout_total_ht = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    tva_montant = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    cout_total_ttc = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    derniere_mise_a_jour = models.DateTimeField(auto_now=True)

    def calculer_totaux(self):
        """Calcule les totaux en incluant les éléments standards et les demandes personnalisées approuvées"""
        elements_standards = EstimationElement.objects.filter(projet=self.projet)
        demandes_approuvees = DemandeElement.objects.filter(
            projet=self.projet,
            statut='approuve',
            prix_unitaire_admin__isnull=False
        )

        self.cout_total_materiel = 0
        self.cout_total_main_oeuvre = 0
        self.cout_total_transport = 0
        self.cout_total_etude = 0

        for element in elements_standards:
            if element.element and element.element.categorie:
                type_cat = element.element.categorie.type_categorie
                cout = element.cout_total
                if type_cat == 'materiel':
                    self.cout_total_materiel += cout
                elif type_cat == 'main_oeuvre':
                    self.cout_total_main_oeuvre += cout
                elif type_cat == 'transport':
                    self.cout_total_transport += cout
                elif type_cat == 'etude':
                    self.cout_total_etude += cout

        for demande in demandes_approuvees:
            if demande.categorie:
                type_cat = demande.categorie.type_categorie
                cout = demande.cout_total
                if type_cat == 'materiel':
                    self.cout_total_materiel += cout
                elif type_cat == 'main_oeuvre':
                    self.cout_total_main_oeuvre += cout
                elif type_cat == 'transport':
                    self.cout_total_transport += cout
                elif type_cat == 'etude':
                    self.cout_total_etude += cout

        self.cout_total_ht = (
            self.cout_total_materiel + self.cout_total_main_oeuvre +
            self.cout_total_transport + self.cout_total_etude
        )
        self.tva_montant = self.cout_total_ht * (self.tva_taux / 100)
        self.cout_total_ttc = self.cout_total_ht + self.tva_montant

        # Sessions de sablage validées (considérées main d'œuvre)
        sessions_sablage = SessionSablage.objects.filter(projet=self.projet, valide=True)
        for session in sessions_sablage:
            self.cout_total_main_oeuvre += session.cout_total

        self.cout_total_ht = (
            self.cout_total_materiel + self.cout_total_main_oeuvre +
            self.cout_total_transport + self.cout_total_etude
        )
        self.tva_montant = self.cout_total_ht * (self.tva_taux / 100)
        self.cout_total_ttc = self.cout_total_ht + self.tva_montant
        self.save()

    def __str__(self):
        return f"Résumé - {self.projet.nom}"

    class Meta:
        verbose_name = "Résumé d'estimation"
        verbose_name_plural = "Résumés d'estimation"


# -----------------------------
# Sablage (optionnel)
# -----------------------------

def ajax_calculer_surface_sablage(request):
    """Calcul AJAX de la surface pour aperçu en temps réel (idéalement à placer dans views.py)"""
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            type_piece = data.get('type_piece')
            dn = int(data.get('dn', 0))
            quantite = float(data.get('quantite', 0))

            SURFACES_SABLAGE = {
                15: {'tube': 0.067, 'coude_90': 0.004, 'coude_45': 0.002, 'coude_90_r5d': 0.007,
                     'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                20: {'tube': 0.084, 'coude_90': 0.005, 'coude_45': 0.003, 'coude_90_r5d': 0.013,
                     'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                25: {'tube': 0.105, 'coude_90': 0.006, 'coude_45': 0.003, 'coude_90_r5d': 0.021,
                     'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                32: {'tube': 0.133, 'coude_90': 0.010, 'coude_45': 0.005, 'coude_90_r5d': 0.033,
                     'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                40: {'tube': 0.152, 'coude_90': 0.014, 'coude_45': 0.007, 'coude_90_r5d': 0.045,
                     'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                50: {'tube': 0.189, 'coude_90': 0.023, 'coude_45': 0.011, 'coude_90_r5d': 0.076,
                     'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                65: {'tube': 0.229, 'coude_90': 0.034, 'coude_45': 0.017, 'coude_90_r5d': 0.114,
                     'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                80: {'tube': 0.279, 'coude_90': 0.050, 'coude_45': 0.025, 'coude_90_r5d': 0.167,
                     'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                100: {'tube': 0.359, 'coude_90': 0.086, 'coude_45': 0.043, 'coude_90_r5d': 0.287,
                      'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                125: {'tube': 0.439, 'coude_90': 0.131, 'coude_45': 0.066, 'coude_90_r5d': 0.438,
                      'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                150: {'tube': 0.529, 'coude_90': 0.190, 'coude_45': 0.095, 'coude_90_r5d': 0.633,
                      'coude_secteur': 0.194, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                200: {'tube': 0.688, 'coude_90': 0.330, 'coude_45': 0.165, 'coude_90_r5d': 1.099,
                      'coude_secteur': 0.338, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                250: {'tube': 0.858, 'coude_90': 0.513, 'coude_45': 0.257, 'coude_90_r5d': 1.711,
                      'coude_secteur': 0.525, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                300: {'tube': 1.018, 'coude_90': 0.731, 'coude_45': 0.365, 'coude_90_r5d': 2.436,
                      'coude_secteur': 0.748, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                350: {'tube': 1.117, 'coude_90': 0.936, 'coude_45': 0.468, 'coude_90_r5d': 3.120,
                      'coude_secteur': 0.958, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                400: {'tube': 1.277, 'coude_90': 1.222, 'coude_45': 0.611, 'coude_90_r5d': 4.076,
                      'coude_secteur': 1.252, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                450: {'tube': 1.436, 'coude_90': 1.547, 'coude_45': 0.774, 'coude_90_r5d': 5.158,
                      'coude_secteur': 1.585, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                500: {'tube': 1.596, 'coude_90': 1.910, 'coude_45': 0.955, 'coude_90_r5d': 6.368,
                      'coude_secteur': 1.953, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                600: {'tube': 1.915, 'coude_90': 2.751, 'coude_45': 1.376, 'coude_90_r5d': 9.169,
                      'coude_secteur': 2.815, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                700: {'tube': 2.234, 'coude_90': 3.744, 'coude_45': 1.872, 'coude_90_r5d': 12.48,
                      'coude_secteur': 3.283, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                750: {'tube': 0, 'coude_90': 4.298, 'coude_45': 2.149, 'coude_90_r5d': 14.33,
                      'coude_secteur': 0, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                800: {'tube': 2.554, 'coude_90': 4.892, 'coude_45': 2.446, 'coude_90_r5d': 16.30,
                      'coude_secteur': 5.003, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                900: {'tube': 2.871, 'coude_90': 6.185, 'coude_45': 3.093, 'coude_90_r5d': 20.62,
                      'coude_secteur': 6.330, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
                1000: {'tube': 3.192, 'coude_90': 7.641, 'coude_45': 3.821, 'coude_90_r5d': 25.47,
                       'coude_secteur': 7.823, 'te': 0, 'bride': 0, 'reduction': 0, 'cap': 0},
            }

            if dn in SURFACES_SABLAGE and type_piece in SURFACES_SABLAGE[dn]:
                surface_unitaire = SURFACES_SABLAGE[dn][type_piece]
                surface_totale = quantite * surface_unitaire
                return JsonResponse({
                    'success': True,
                    'surface_unitaire': surface_unitaire,
                    'surface_totale': surface_totale,
                    'disponible': surface_unitaire > 0,
                    'dn_mm': {
                        15: 21.3, 20: 26.7, 25: 33.4, 32: 42.2, 40: 48.3,
                        50: 60.3, 65: 73, 80: 88.9, 100: 114.3, 125: 139.7,
                        150: 168.3, 200: 219.1, 250: 273, 300: 323.9, 350: 355.6,
                        400: 406.5, 450: 457.2, 500: 508, 600: 609.6, 700: 711,
                        750: 762, 800: 813, 900: 914, 1000: 1016
                    }.get(dn, dn)
                })
            else:
                return JsonResponse({'success': False, 'message': 'Combinaison non disponible', 'disponible': False})
        except (ValueError, TypeError, KeyError) as e:
            return JsonResponse({'success': False, 'message': f'Données invalides: {str(e)}', 'disponible': False})

    return JsonResponse({'success': False, 'message': 'Méthode non autorisée'})


class CalculSablage(models.Model):
    """Stocke les détails des calculs de sablage"""
    projet = models.ForeignKey(Projet, on_delete=models.CASCADE)
    type_piece = models.CharField(max_length=50)
    diametre_dn = models.IntegerField(help_text="DN en mm")
    quantite = models.DecimalField(max_digits=10, decimal_places=2)
    surface_unitaire = models.DecimalField(max_digits=10, decimal_places=6, help_text="m² / unité")
    surface_totale = models.DecimalField(max_digits=15, decimal_places=6, help_text="m² total")
    date_creation = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.type_piece} DN{self.diametre_dn} - {self.quantite} pcs"

    class Meta:
        verbose_name = "Calcul de sablage"
        verbose_name_plural = "Calculs de sablage"


class SessionSablage(models.Model):
    """Session de calcul de sablage pour un projet"""
    projet = models.ForeignKey(Projet, on_delete=models.CASCADE)
    surface_globale = models.DecimalField(max_digits=15, decimal_places=6, default=0)
    prix_unitaire_m2 = models.DecimalField(max_digits=10, decimal_places=2, default=5000)
    cout_total = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    valide = models.BooleanField(default=False)
    date_validation = models.DateTimeField(null=True, blank=True)
    calculs = models.ManyToManyField(CalculSablage, blank=True)

    def calculer_total(self):
        self.surface_globale = sum(calc.surface_totale for calc in self.calculs.all())
        self.cout_total = self.surface_globale * self.prix_unitaire_m2
        self.save()
        return self.cout_total

    def __str__(self):
        return f"Session sablage - {self.projet.nom} - {self.surface_globale} m²"

    class Meta:
        verbose_name = "Session de sablage"
        verbose_name_plural = "Sessions de sablage"
